const express = require('express');
const { google } = require('googleapis');
const xlsx = require('xlsx');
const path = require('path');

const app = express();

app.set('view engine', 'ejs');
app.use(express.static(path.join(__dirname, 'public')));

// ID Folder Utama Google Drive
const MAIN_FOLDER_ID = '1kyl4AOQCfu8r1c4qDMZZF28bxWU2dO17';

let drive;
try {
    drive = google.drive({
        version: 'v3',
        auth: new google.auth.GoogleAuth({
            keyFile: path.join(__dirname, 'service-account.json'),
            scopes: ['https://www.googleapis.com/auth/drive.readonly']
        })
    });
    console.log('✅ Konfigurasi Google Auth berhasil.');
} catch (e) {
    console.error('❌ Gagal memuat service-account.json:', e.message);
}

// Fungsi Parser Angka Mutlak
const parseStrictNumber = (val) => {
    if (val === undefined || val === null || val === '-' || val === '') return 0;
    if (typeof val === 'number') return Math.round(val);
    let str = String(val).trim();
    if (str.includes('.') && !str.includes(',')) {
        let parts = str.split('.');
        if (parts.length > 1 && parts[parts.length - 1].length === 3) str = parts.join('');
    }
    str = str.replace(/[^0-9]/g, '');
    return parseInt(str, 10) || 0;
};

// ==========================================
// MESIN SCANNER VERTIKAL (ANTI MERGED-CELLS)
// ==========================================
const getBulletproofColIndex = (rows, exactKeywords, partialKeywords = []) => {
    let maxCols = 0;
    const maxRowsToScan = Math.min(rows.length, 15);
    for (let i = 0; i < maxRowsToScan; i++) {
        if (rows[i] && rows[i].length > maxCols) maxCols = rows[i].length;
    }

    const scan = (keywords) => {
        for (let kw of keywords) {
            let cleanKw = kw.toLowerCase().replace(/[^a-z0-9+]/g, '');
            for (let col = 0; col < maxCols; col++) {
                let colText = '';
                for (let row = 0; row < maxRowsToScan; row++) {
                    if (rows[row] && rows[row][col] !== undefined) {
                        colText += String(rows[row][col]).toLowerCase().replace(/[^a-z0-9+]/g, '');
                    }
                }
                if (colText.includes(cleanKw)) return col;
            }
        }
        return -1;
    };

    let idx = scan(exactKeywords);
    if (idx === -1) idx = scan(partialKeywords);
    return idx;
};

// Pencari Nama Petugas yang Akurat
const getValidName = (r) => {
    if (r[1] !== undefined) {
        let val = String(r[1]).trim();
        if (val && !val.includes('@') && isNaN(val) && val.length > 2) return val;
    }
    for (let idx = 1; idx < Math.min(r.length, 5); idx++) {
        if (r[idx] !== undefined) {
            let val = String(r[idx]).trim();
            if (val && !val.includes('@') && isNaN(val) && val.length > 3) return val;
        }
    }
    return '';
};

// Fungsi penarik data Excel mentah (Rows)
const fetchExcelRows = async (folderId, keywords) => {
    if (!drive) return [];
    
    const searchInParent = async (parentId) => {
        try {
            const qQuery = parentId ? `'${parentId}' in parents and trashed = false` : `trashed = false`;
            const res = await drive.files.list({
                q: qQuery, pageSize: 100, fields: 'files(id, name, mimeType)', orderBy: 'name desc'
            });
            return res.data.files || [];
        } catch (e) { return []; }
    };

    try {
        let files = await searchInParent(folderId);
        if (files.length === 0 && folderId !== MAIN_FOLDER_ID) files = await searchInParent(MAIN_FOLDER_ID);
        if (files.length === 0) files = await searchInParent(null);

        files.sort((a, b) => (b.name || '').localeCompare(a.name || ''));
        let matchedFile = files.find(f => keywords.some(kw => f.name && f.name.toLowerCase().includes(kw.toLowerCase())));

        if (!matchedFile) {
            console.warn(`⚠️ File [${keywords.join(', ')}] tidak ditemukan.`);
            return [];
        }

        const fileId = matchedFile.id;
        const mimeType = matchedFile.mimeType;
        let buffer;

        if (mimeType === 'application/vnd.google-apps.spreadsheet') {
            const exportRes = await drive.files.export({
                fileId: fileId, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            }, { responseType: 'arraybuffer' });
            buffer = Buffer.from(exportRes.data);
        } else {
            const mediaRes = await drive.files.get({ fileId: fileId, alt: 'media' }, { responseType: 'arraybuffer' });
            buffer = Buffer.from(mediaRes.data);
        }

        const wb = xlsx.read(buffer, { type: 'buffer' });
        return xlsx.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 }) || [];
    } catch (err) {
        console.error(`⚠️ Gagal baca file [${keywords.join(', ')}]:`, err.message);
        return [];
    }
};

const parseDesaData = (rows) => {
    if (!rows || rows.length === 0) return [];
    let targetIdx = getBulletproofColIndex(rows, ['target', 'jumlahusaha']);
    let didataIdx = getBulletproofColIndex(rows, ['respondendidata', 'realisasi', 'selesai']);
    
    let dataStartIdx = rows.findIndex((r, idx) => idx > 1 && r && (r[0] === 1 || String(r[0]).trim() === '1'));
    if (dataStartIdx === -1) dataStartIdx = 3;

    return rows.slice(dataStartIdx).map(r => {
        let name = getValidName(r);
        if (!name) return null;
        let tVal = targetIdx !== -1 ? parseStrictNumber(r[targetIdx]) : 0;
        let dVal = didataIdx !== -1 ? parseStrictNumber(r[didataIdx]) : 0;
        return {
            name,
            target: tVal,
            didata: dVal,
            'Target Prelist Awal pada kolom index ke-5 (kolom F)': Math.max(0, tVal - dVal)
        };
    }).filter(item => item && item.name.toLowerCase() !== 'jumlah' && item.name.toLowerCase() !== 'total');
};

const parsePetugas = (rows, role) => {
    if (!rows || rows.length === 0) return [];

    let dataStartIdx = rows.findIndex((r, idx) => idx > 1 && r && (r[0] === 1 || String(r[0]).trim() === '1'));
    if (dataStartIdx === -1) dataStartIdx = 3;

    // Pemetaan posisi kolom mutlak sesuai file Excel rekap_petugas
    let targetIdx = 5; // Kolom F (Target Prelist)
    let didataIdx = 6; // Kolom G (Responden Didata)
    let harianIdx = 7; // Kolom H (+ Didata)
    let pctIdx = 8;    // Kolom I (% Didata)

    let compositeHeaders = [];
    let maxCols = 0;
    for (let r = 0; r < dataStartIdx; r++) {
        if (rows[r] && rows[r].length > maxCols) maxCols = rows[r].length;
    }

    for (let c = 0; c < maxCols; c++) {
        let label = '';
        for (let r = 0; r < dataStartIdx; r++) {
            if (rows[r] && rows[r][c] !== undefined && rows[r][c] !== null) {
                let txt = String(rows[r][c]).trim();
                if (txt && !label.includes(txt)) label += (label ? ' ' : '') + txt;
            }
        }
        if (c === targetIdx) label = 'Target Prelist';
        if (c === didataIdx) label = 'Responden Didata';
        if (c === harianIdx) label = '+ Didata';
        if (c === pctIdx) label = '% Didata';
        compositeHeaders[c] = label || `Kolom_${c}`;
    }

    return rows.slice(dataStartIdx).map(r => {
        let name = getValidName(r);
        if (!name) return null;

        let targetVal = parseStrictNumber(r[targetIdx]);
        let didataVal = parseStrictNumber(r[didataIdx]);
        let harianVal = parseStrictNumber(r[harianIdx]);

        // Perhitungan Sisa Target Murni: Target Prelist - Responden Didata
        let sisaTargetVal = Math.max(0, targetVal - didataVal);

        // Perhitungan / Pengambilan Persentase
        let rawPct = r[pctIdx];
        let formattedPct = '0.0%';
        let numericPct = 0;

        if (rawPct !== undefined && rawPct !== null && rawPct !== '' && rawPct !== '-') {
            let strPct = String(rawPct).trim();
            if (!strPct.includes('%')) {
                let num = parseFloat(strPct.replace(',', '.'));
                if (!isNaN(num)) {
                    numericPct = num <= 1 ? num * 100 : num;
                    formattedPct = numericPct.toFixed(1) + '%';
                }
            } else {
                formattedPct = strPct;
                numericPct = parseFloat(strPct.replace('%', '').trim()) || 0;
            }
        } else if (targetVal > 0) {
            numericPct = (didataVal / targetVal) * 100;
            formattedPct = numericPct.toFixed(1) + '%';
        }

        // KUMPULAN KUNCI LENGKAP (Kompatibel dengan semua template EJS)
        let rowObj = { 
            name, 
            target: targetVal, 
            targetPrelist: targetVal,
            didata: didataVal, 
            realisasi: didataVal,
            respondenDidata: didataVal,
            harian: harianVal, 
            progress: formattedPct,
            progressVal: numericPct,
            persentase: formattedPct,
            persentaseDidataDraft: formattedPct,
            'Target Prelist Awal dari spreadsheet index ke-5': targetVal,   // Target Prelist Awal dari spreadsheet index ke-5
            'Target Prelist Awal pada kolom index ke-5 (kolom F)': targetVal,   // Target Prelist Awal pada kolom index ke-5 (kolom F)
            rawCells: []
        };
        
        let rowCols = Math.max(compositeHeaders.length, r.length);
        for (let idx = 0; idx < rowCols; idx++) {
            let h = compositeHeaders[idx] || `Kolom_${idx}`;
            let val = r[idx] !== undefined ? r[idx] : '-';
            rowObj.rawCells.push({ header: h, value: val });
        }

        return rowObj;
    }).filter(item => item && item.name.toLowerCase() !== 'jumlah' && item.name.toLowerCase() !== 'total');
};

app.get('/', async (req, res) => {
    let activeDate = '2026-08-03';
    let persentaseKretek = '80.9';
    let totalTargetKretek = 14764;
    let totalDidataKretek = 11939;
    let topPcl = [], bottomPcl = [], parsedDesa = [], detailedPml = [], detailedPcl = [];

    try {
        if (drive) {
            let targetFolderId = MAIN_FOLDER_ID;
            try {
                let folderRes;
                try {
                    folderRes = await drive.files.list({
                        q: `'${MAIN_FOLDER_ID}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false`,
                        orderBy: 'name desc',
                        pageSize: 10
                    });
                } catch (err) {
                    folderRes = await drive.files.list({
                        q: `mimeType = 'application/vnd.google-apps.folder' and trashed = false`,
                        orderBy: 'name desc',
                        pageSize: 10
                    });
                }

                const folders = folderRes.data.files || [];
                if (folders.length > 0) {
                    const validFolder = folders.find(f => /^\d{4}-\d{2}-\d{2}$/.test(f.name.trim())) || folders[0];
                    activeDate = validFolder.name;
                    targetFolderId = validFolder.id;
                }
            } catch (e) {
                console.error("Gagal mengambil folder:", e.message);
            }

            const [kecRows, desaRows, pmlRows, pclRows] = await Promise.all([
                fetchExcelRows(targetFolderId, ['rekap_wilayah_kecamatan', 'kecamatan', 'kec']),
                fetchExcelRows(targetFolderId, ['rekap_wilayah_desa', 'desa', 'kalurahan']),
                fetchExcelRows(targetFolderId, ['rekap_petugas_pml', 'pml', 'petugas_pml']),
                fetchExcelRows(targetFolderId, ['rekap_petugas_pcl', 'pcl', 'petugas_pcl'])
            ]);

            if (kecRows && kecRows.length > 0) {
                let targetIdx = getBulletproofColIndex(kecRows, ['target', 'jumlahusaha']);
                let didataIdx = getBulletproofColIndex(kecRows, ['respondendidata', 'realisasi', 'selesai']);
                let kretekRow = kecRows.find(r => r.join(' ').toLowerCase().includes('kretek') || r.join(' ').toLowerCase().includes('3402030'));
                
                if (kretekRow) {
                    totalTargetKretek = targetIdx !== -1 ? parseStrictNumber(kretekRow[targetIdx]) : 14764;
                    totalDidataKretek = didataIdx !== -1 ? parseStrictNumber(kretekRow[didataIdx]) : 11939;
                    if (totalTargetKretek > 0) persentaseKretek = ((totalDidataKretek / totalTargetKretek) * 100).toFixed(1);
                }
            }

            parsedDesa = parseDesaData(desaRows);
            detailedPml = parsePetugas(pmlRows, 'pml');
            detailedPcl = parsePetugas(pclRows, 'pcl');

            if (detailedPcl.length > 0) {
                topPcl = [...detailedPcl].sort((a, b) => b.harian - a.harian).slice(0, 5);
                bottomPcl = [...detailedPcl].sort((a, b) => a.harian - b.harian).slice(0, 5);
            }
        }
    } catch (err) {
        console.error("⚠️ Terjadi kendala saat memuat data:", err.message);
    }

    res.render('index', {
        activeDate, persentaseKretek, totalTargetKretek, totalDidataKretek,
        topPcl, bottomPcl, parsedDesa, detailedPml, detailedPcl
    });
});

app.listen(3000, () => {
    console.log('🚀 Server berjalan sukses di http://localhost:3000');
});