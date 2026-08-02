const express = require('express');
const { google } = require('googleapis');
const xlsx = require('xlsx');
const path = require('path');

const app = express();

app.set('view engine', 'ejs');
app.use(express.static(path.join(__dirname, 'public')));

const MAIN_FOLDER_ID = '1kyI4AOQCfu8r1c4qDMZZF28bxWU2dO17';

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

const getColIndex = (headers, keywords) => {
    for (let kw of keywords) {
        let idx = headers.findIndex(h => h && h.toLowerCase().replace(/\s+/g, '').includes(kw.toLowerCase().replace(/\s+/g, '')));
        if (idx !== -1) return idx;
    }
    return -1;
};

const fetchExcelFromDrive = async (folderId, fileKeyword) => {
    if (!drive) return { headers: [], data: [] };
    try {
        const fileRes = await drive.files.list({
            q: `'${folderId}' in parents and name contains '${fileKeyword}' and trashed = false`,
            pageSize: 1
        });
        const files = fileRes.data.files;
        if (files.length === 0) return { headers: [], data: [] };

        const file = files[0];
        const fileId = file.id;
        const mimeType = file.mimeType;

        let buffer;
        if (mimeType === 'application/vnd.google-apps.spreadsheet') {
            const exportRes = await drive.files.export({
                fileId: fileId,
                mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            }, { responseType: 'arraybuffer' });
            buffer = Buffer.from(exportRes.data);
        } else {
            const mediaRes = await drive.files.get({
                fileId: fileId,
                alt: 'media'
            }, { responseType: 'arraybuffer' });
            buffer = Buffer.from(mediaRes.data);
        }

        const wb = xlsx.read(buffer, { type: 'buffer' });
        const sheetName = wb.SheetNames[0];
        const rows = xlsx.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1 });

        if (rows.length < 2) return { headers: [], data: [] };

        const h1 = rows[0];
        const h2 = rows[1];
        const headers = h2.map((val, idx) => {
            let parent = h1[idx] !== undefined && h1[idx] !== '' ? String(h1[idx]).trim() : '';
            let child = val !== undefined && val !== '' ? String(val).trim() : '';
            
            if (parent.toLowerCase().includes('progres pendataan')) {
                return child || parent || `Col_${idx}`;
            }

            if (parent && parent !== child && !parent.includes('No') && !parent.includes('Nama') && !parent.includes('Email') && !parent.includes('Role')) {
                return `${parent} - ${child}`;
            }
            return child || parent || `Col_${idx}`;
        });

        const data = rows.slice(2).filter(r => r && r.length > 0 && r.join(' ').trim() !== '');
        return { headers, data };
    } catch (err) {
        console.error(`⚠️ Gagal baca file keyword "${fileKeyword}":`, err.message);
        return { headers: [], data: [] };
    }
};

app.get('/', async (req, res) => {
    let activeDate = '2026-07-31';
    let persentaseKretek = '80.9';
    let totalTargetKretek = 14764;
    let totalDidataKretek = 11939;
    let topPcl = [];
    let bottomPcl = [];
    let parsedDesa = [];
    let detailedPml = [];
    let detailedPcl = [];

    try {
        if (drive) {
            const folderRes = await drive.files.list({
                q: `'${MAIN_FOLDER_ID}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false`,
                orderBy: 'name desc',
                pageSize: 50
            });

            const folders = folderRes.data.files;
            if (folders && folders.length > 0) {
                let activeFolder = folders[0];
                activeDate = activeFolder.name;

                const [kecRes, desaRes, pmlRes, pclRes] = await Promise.all([
                    fetchExcelFromDrive(activeFolder.id, 'rekap_wilayah_kecamatan'),
                    fetchExcelFromDrive(activeFolder.id, 'rekap_wilayah_desa'),
                    fetchExcelFromDrive(activeFolder.id, 'rekap_petugas_pml'),
                    fetchExcelFromDrive(activeFolder.id, 'rekap_petugas_pcl')
                ]);

                if (kecRes.data.length > 0) {
                    let kretekRow = kecRes.data.find(r => {
                        let rowText = r.join(' ').toLowerCase();
                        return rowText.includes('kretek') || rowText.includes('3402030');
                    });
                    if (kretekRow) {
                        const targetIdx = getColIndex(kecRes.headers, ['target', 'jumlah usaha', 'usaha']);
                        const didataIdx = getColIndex(kecRes.headers, ['responden didata', 'didata', 'realisasi']);
                        totalTargetKretek = targetIdx !== -1 ? parseFloat(kretekRow[targetIdx]) || 14764 : 14764;
                        totalDidataKretek = didataIdx !== -1 ? parseFloat(kretekRow[didataIdx]) || 11939 : 11939;
                        if (totalTargetKretek > 0) {
                            persentaseKretek = ((totalDidataKretek / totalTargetKretek) * 100).toFixed(1);
                        }
                    }
                }

                const getValidName = (headers, r) => {
                    for (let idx of [2, 1, 3, 4, 0]) {
                        if (r[idx] !== undefined) {
                            let val = String(r[idx]).trim();
                            if (val && !val.includes('@') && isNaN(val) && val.length > 2 && !val.toLowerCase().includes('http') && !val.toLowerCase().includes('nik')) {
                                return val;
                            }
                        }
                    }
                    for (let i = 0; i < r.length; i++) {
                        let val = String(r[i] || '').trim();
                        if (val && !val.includes('@') && isNaN(val) && val.length > 2 && !val.toLowerCase().includes('http')) {
                            return val;
                        }
                    }
                    return '';
                };

                const parseTableData = (headers, rows) => {
                    const targetIdx = getColIndex(headers, ['target', 'jumlah usaha', 'usaha']);
                    const didataIdx = getColIndex(headers, ['responden didata', 'didata', 'realisasi']);

                    return rows.map(r => {
                        let name = getValidName(headers, r);
                        let target = targetIdx !== -1 ? parseFloat(r[targetIdx]) || 0 : 0;
                        let didata = didataIdx !== -1 ? parseFloat(r[didataIdx]) || 0 : 0;
                        return { name, target, didata };
                    }).filter(item => 
                        item.name !== '' && 
                        item.name.toLowerCase() !== 'jumlah' && 
                        item.name.toLowerCase() !== 'total' && 
                        !item.name.toLowerCase().includes('kapanewon')
                    );
                };

                parsedDesa = parseTableData(desaRes.headers, desaRes.data);

                const mapDetailedRows = (headers, rows) => {
                    return rows.map(r => {
                        let name = getValidName(headers, r);
                        let target = 0;
                        let didata = 0;
                        let harian = 0;
                        
                        let rowObj = { name, rawCells: [], target: 0, didata: 0, harian: 0 };
                        headers.forEach((h, idx) => {
                            let val = r[idx] !== undefined ? r[idx] : '-';
                            
                            let hLow = h.toLowerCase();
                            if ((hLow.includes('%') || hLow.includes('persen')) && val !== '-') {
                                let num = parseFloat(val);
                                if (!isNaN(num)) {
                                    val = num <= 1 ? (num * 100) + '%' : num + '%';
                                }
                            }

                            rowObj.rawCells.push({ header: h, value: val });
                            
                            if (hLow.includes('target') || hLow.includes('jumlah usaha')) target = parseFloat(r[idx]) || target;
                            if (hLow.includes('didata') || hLow.includes('realisasi')) didata = parseFloat(r[idx]) || didata;
                            
                            // Ambil langsung dari kolom +didata / + didata / harian di file XLS
                            let cleanH = hLow.replace(/\s+/g, '');
                            if (cleanH.includes('+didata') || cleanH.includes('harian')) {
                                harian = parseFloat(r[idx]) || harian;
                            }
                        });
                        rowObj.target = target;
                        rowObj.didata = didata;
                        rowObj.harian = harian;
                        return rowObj;
                    }).filter(item => item.name !== '' && item.name.toLowerCase() !== 'jumlah' && item.name.toLowerCase() !== 'total');
                };

                detailedPml = mapDetailedRows(pmlRes.headers, pmlRes.data);
                detailedPcl = mapDetailedRows(pclRes.headers, pclRes.data);

                if (detailedPcl.length > 0) {
                    const sortedAsc = [...detailedPcl].sort((a, b) => a.harian - b.harian);
                    const sortedDesc = [...detailedPcl].sort((a, b) => b.harian - a.harian);

                    topPcl = sortedDesc.slice(0, 5);
                    bottomPcl = sortedAsc.slice(0, 5);
                }
            }
        }
    } catch (err) {
        console.error("⚠️ Terjadi kendala saat memuat data dari Drive:", err.message);
    }

    res.render('index', {
        activeDate,
        persentaseKretek,
        totalTargetKretek,
        totalDidataKretek,
        topPcl,
        bottomPcl,
        parsedDesa,
        detailedPml,
        detailedPcl
    });
});

app.listen(3000, () => {
    console.log('🚀 Server berjalan sukses di http://localhost:3000');
});