const express = require('express');
const { google } = require('googleapis');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();

app.set('view engine', 'ejs');
app.use(express.static(path.join(__dirname, 'public')));

// ID Folder Utama Google Drive
const MAIN_FOLDER_ID = '1h3RtIf9YRpBlnIHxgI0WHSCfEyw4_0DT';

let drive;
try {
    let authConfig;
    if (process.env.GOOGLE_CREDENTIALS) {
        authConfig = {
            credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS),
            scopes: ['https://www.googleapis.com/auth/drive.readonly']
        };
    } else {
        authConfig = {
            keyFile: path.join(__dirname, 'service-account.json'),
            scopes: ['https://www.googleapis.com/auth/drive.readonly']
        };
    }

    drive = google.drive({
        version: 'v3',
        auth: new google.auth.GoogleAuth(authConfig)
    });
    console.log('✅ Konfigurasi Google Auth berhasil.');
} catch (e) {
    console.error('❌ Gagal memuat Google Auth:', e.message);
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

    const scan = (keywords, exactMatchOnly = false) => {
        for (let kw of keywords) {
            let cleanKw = kw.toLowerCase().replace(/[^a-z0-9+]/g, '');
            for (let col = 0; col < maxCols; col++) {
                let colText = '';
                for (let row = 0; row < maxRowsToScan; row++) {
                    if (rows[row] && rows[row][col] !== undefined) {
                        colText += String(rows[row][col]).toLowerCase().replace(/[^a-z0-9+]/g, '');
                    }
                }
                if (exactMatchOnly) {
                    if (colText === cleanKw) return col;
                } else {
                    if (colText.includes(cleanKw)) return col;
                }
            }
        }
        return -1;
    };

    let idx = scan(exactKeywords, true); // Cari exact match dulu
    if (idx === -1) idx = scan(exactKeywords, false);
    if (idx === -1) idx = scan(partialKeywords, false);
    return idx;
};

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

const fetchExcelRows = async (folderId, keywords) => {
    if (!drive) return [];
    
    const searchInParent = async (parentId) => {
        try {
            const qQuery = parentId ? `'${parentId}' in parents and trashed = false` : `trashed = false`;
            const res = await drive.files.list({
                q: qQuery, 
                pageSize: 100, 
                fields: 'files(id, name, mimeType, modifiedTime)', 
                orderBy: 'modifiedTime desc'
            });
            return res.data.files || [];
        } catch (e) { return []; }
    };

    try {
        let files = await searchInParent(folderId);
        if (files.length === 0 && folderId !== MAIN_FOLDER_ID) files = await searchInParent(MAIN_FOLDER_ID);
        if (files.length === 0) files = await searchInParent(null);

        files.sort((a, b) => new Date(b.modifiedTime) - new Date(a.modifiedTime));
        let matchedFile = files.find(f => keywords.some(kw => f.name && f.name.toLowerCase().includes(kw.toLowerCase())));

        if (!matchedFile) {
            console.warn(`⚠️ File [${keywords.join(', ')}] tidak ditemukan.`);
            return { rows: [], modifiedTime: null };
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
        const sheetRows = xlsx.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 }) || [];
        return { rows: sheetRows, modifiedTime: matchedFile.modifiedTime };
    } catch (err) {
        console.error(`⚠️ Gagal baca file [${keywords.join(', ')}]:`, err.message);
        return { rows: [], modifiedTime: null };
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
        return {
            name,
            target: targetIdx !== -1 ? parseStrictNumber(r[targetIdx]) : 0,
            didata: didataIdx !== -1 ? parseStrictNumber(r[didataIdx]) : 0
        };
    }).filter(item => item && item.name.toLowerCase() !== 'jumlah' && item.name.toLowerCase() !== 'total');
};

const parsePetugas = (rows, role) => {
    if (!rows || rows.length === 0) return [];

    let dataStartIdx = rows.findIndex((r, idx) => idx > 1 && r && (r[0] === 1 || String(r[0]).trim() === '1'));
    if (dataStartIdx === -1) dataStartIdx = 3;

    let targetIdx = getBulletproofColIndex(rows, ['targetprelist', 'jumlahusaha', 'slsditarget', 'prelist'], ['target']);
    
    // PENGECUALIAN MUTLAK: Hindari kolom yang mengandung kata 'draft' untuk Responden Didata PCL
    let didataIdx = -1;
    let maxCols = 0;
    for (let i = 0; i < Math.min(rows.length, 15); i++) {
        if (rows[i] && rows[i].length > maxCols) maxCols = rows[i].length;
    }

    for (let c = 0; c < maxCols; c++) {
        let colText = '';
        for (let r = 0; r < Math.min(rows.length, 15); r++) {
            if (rows[r] && rows[r][c] !== undefined) {
                colText += String(rows[r][c]).toLowerCase().replace(/[^a-z0-9]/g, '');
            }
        }
        if ((colText.includes('respondendidata') || colText === 'didata') && !colText.includes('draft')) {
            didataIdx = c;
            break;
        }
    }

    if (didataIdx === -1) {
        didataIdx = getBulletproofColIndex(rows, ['respondendidata', 'realisasi', 'selesai'], ['didata']);
    }

    let harianIdx = getBulletproofColIndex(rows, ['+didata', 'progressharian', 'harian'], ['+']);
    let pctDraftIdx = getBulletproofColIndex(rows, ['persentase', 'didatadraft', 'draft', 'capaian'], ['%']);

    if (role === 'pcl') {
        if (targetIdx === -1) targetIdx = 5;
        if (didataIdx === -1) didataIdx = 6;
        if (harianIdx === -1) harianIdx = 7;
    } else {
        if (targetIdx === -1) targetIdx = 4;
        if (didataIdx === -1) didataIdx = 5;
        if (harianIdx === -1) harianIdx = 6;
        if (pctDraftIdx === -1) pctDraftIdx = 7;
    }

    let compositeHeaders = [];
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
        compositeHeaders[c] = label || `Kolom_${c}`;
    }

    return rows.slice(dataStartIdx).map(r => {
        let name = getValidName(r);
        if (!name) return null;

        let targetVal = parseStrictNumber(r[targetIdx]);
        let didataVal = parseStrictNumber(r[didataIdx]);
        let harianVal = parseStrictNumber(r[harianIdx]);

        let rawPctDraft = pctDraftIdx !== -1 ? r[pctDraftIdx] : (targetVal > 0 ? (didataVal / targetVal) * 100 : 0);
        let formattedPctDraft = rawPctDraft;
        if (typeof rawPctDraft === 'number') {
            formattedPctDraft = (rawPctDraft >= 0 && rawPctDraft <= 1 ? rawPctDraft * 100 : rawPctDraft).toFixed(2) + '%';
        } else if (typeof rawPctDraft === 'string') {
            let num = parseFloat(rawPctDraft.replace('%', '').replace(',', '.'));
            if (!isNaN(num)) {
                formattedPctDraft = (num >= 0 && num <= 1 ? num * 100 : num).toFixed(2) + '%';
            }
        }

        let rowObj = { 
            name, 
            rawCells: [], 
            target: targetVal, 
            didata: didataVal, 
            realisasi: didataVal,
            respondenDidata: didataVal,
            harian: harianVal, 
            progress: harianVal,
            sisaTarget: Math.max(0, targetVal - didataVal),
            persentaseDidataDraft: formattedPctDraft,
            persentase: formattedPctDraft
        };
        
        let rowCols = Math.max(compositeHeaders.length, r.length);
        for (let idx = 0; idx < rowCols; idx++) {
            let h = compositeHeaders[idx] || `Kolom_${idx}`;
            let val = r[idx] !== undefined ? r[idx] : '-';
            
            if (typeof val === 'number' && val >= 0 && val <= 1) {
                let hLower = h.toLowerCase();
                if (hLower.includes('%') || hLower.includes('persen') || hLower.includes('progress') || hLower.includes('capaian')) {
                    val = (val * 100).toFixed(2) + '%';
                }
            } else if (typeof val === 'string' && (val.toLowerCase().includes('%') || val.toLowerCase().includes('persen'))) {
                let num = parseFloat(val.replace(',', '.'));
                if (!isNaN(num)) val = (num <= 1 ? num * 100 : num).toFixed(2) + '%';
            }
            rowObj.rawCells.push({ header: h, value: val });
        }

        return rowObj;
    }).filter(item => item && item.name.toLowerCase() !== 'jumlah' && item.name.toLowerCase() !== 'total');
};

app.get('/', async (req, res) => {
    let activeDate = new Date().toISOString().split('T')[0];
    let persentaseKretek = '0';
    let totalTargetKretek = 0;
    let totalDidataKretek = 0;
    let topPcl = [], bottomPcl = [], parsedDesa = [], detailedPml = [], detailedPcl = [];

    try {
        if (drive) {
            const [kecRes, desaRes, pmlRes, pclRes] = await Promise.all([
                fetchExcelRows(MAIN_FOLDER_ID, ['rekap_wilayah_kecamatan', 'kecamatan', 'kec']),
                fetchExcelRows(MAIN_FOLDER_ID, ['rekap_wilayah_desa', 'desa', 'kalurahan']),
                fetchExcelRows(MAIN_FOLDER_ID, ['rekap_petugas_pml', 'pml', 'petugas_pml']),
                fetchExcelRows(MAIN_FOLDER_ID, ['rekap_petugas_pcl', 'pcl', 'petugas_pcl'])
            ]);

            let dates = [kecRes.modifiedTime, desaRes.modifiedTime, pmlRes.modifiedTime, pclRes.modifiedTime].filter(Boolean);
            if (dates.length > 0) {
                dates.sort((a, b) => new Date(b) - new Date(a));
                activeDate = dates[0].split('T')[0];
            }

            const kecRows = kecRes.rows;
            if (kecRows && kecRows.length > 0) {
                let targetIdx = getBulletproofColIndex(kecRows, ['target', 'jumlahusaha']);
                let didataIdx = getBulletproofColIndex(kecRows, ['respondendidata', 'realisasi', 'selesai']);
                let kretekRow = kecRows.find(r => r.join(' ').toLowerCase().includes('kretek') || r.join(' ').toLowerCase().includes('3402030'));
                
                if (kretekRow) {
                    totalTargetKretek = targetIdx !== -1 ? parseStrictNumber(kretekRow[targetIdx]) : 0;
                    totalDidataKretek = didataIdx !== -1 ? parseStrictNumber(kretekRow[didataIdx]) : 0;
                    if (totalTargetKretek > 0) persentaseKretek = ((totalDidataKretek / totalTargetKretek) * 100).toFixed(1);
                }
            }

            parsedDesa = parseDesaData(desaRes.rows);
            detailedPml = parsePetugas(pmlRes.rows, 'pml');
            detailedPcl = parsePetugas(pclRes.rows, 'pcl');

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

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 Server berjalan sukses di port ${PORT}`);
});
