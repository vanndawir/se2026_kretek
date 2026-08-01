const express = require('express');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
// Izinkan Express membaca folder 'public' untuk file statis (gambar, CSS, dll)
app.set('view engine', 'ejs');
app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
    let kecHeaders = [], kecData = [];
    let desaHeaders = [], desaData = [];
    let pmlHeaders = [], pmlData = [];
    let pclHeaders = [], pclData = [];
    let activeDate = '2026-07-31';
    let topPcl = [];
    let bottomPcl = [];
    let pclHarianMap = {};
    let pmlHarianMap = {};
    let desaHarianMap = {};

    try {
        const baseDir = path.join(__dirname, 'data');
        if (fs.existsSync(baseDir)) {
            const folders = fs.readdirSync(baseDir).filter(f => fs.statSync(path.join(baseDir, f)).isDirectory()).sort().reverse();
            if (folders.length > 0) {
                activeDate = folders[0];
                const latestDir = path.join(baseDir, activeDate);

                const readExcelNormalized = (filename) => {
                    const filePath = path.join(latestDir, filename);
                    if (fs.existsSync(filePath)) {
                        const wb = xlsx.readFile(filePath);
                        const rows = xlsx.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
                        
                        if (rows.length < 2) return { headers: [], data: [] };

                        const h1 = rows[0];
                        const h2 = rows[1];
                        const headers = h2.map((val, idx) => {
                            let parent = h1[idx] !== undefined && h1[idx] !== '' ? String(h1[idx]).trim() : '';
                            let child = val !== undefined && val !== '' ? String(val).trim() : '';
                            if (parent && parent !== child && !parent.includes('No') && !parent.includes('Nama') && !parent.includes('Email') && !parent.includes('Role')) {
                                return `${parent} - ${child}`;
                            }
                            return child || parent || `Col_${idx}`;
                        });

                        const data = rows.slice(2).filter(r => {
                            if (!r || r.length === 0) return false;
                            const rowString = r.join(' ').toLowerCase();
                            if (rowString.includes('tidak diketahui') || rowString.includes('nan') || rowString.trim() === '') return false;
                            return true;
                        });

                        return { headers, data };
                    }
                    return { headers: [], data: [] };
                };

                const kecRes = readExcelNormalized(`rekap_wilayah_kecamatan_${activeDate}.xls`);
                kecHeaders = kecRes.headers; kecData = kecRes.data;

                const desaRes = readExcelNormalized(`rekap_wilayah_desa_${activeDate}.xls`);
                desaHeaders = desaRes.headers; desaData = desaRes.data;

                const pmlRes = readExcelNormalized(`rekap_petugas_pml_${activeDate}.xls`);
                pmlHeaders = pmlRes.headers; pmlData = pmlRes.data;

                const pclRes = readExcelNormalized(`rekap_petugas_pcl_${activeDate}.xls`);
                pclHeaders = pclRes.headers; pclData = pclRes.data;

                const getColIndex = (headers, keywords) => {
                    for (let kw of keywords) {
                        let idx = headers.findIndex(h => h && h.toLowerCase().includes(kw));
                        if (idx !== -1) return idx;
                    }
                    return -1;
                };

                const pclDidataIdx = getColIndex(pclHeaders, ['responden didata', 'didata']);
                const pmlDidataIdx = getColIndex(pmlHeaders, ['responden didata', 'didata']);
                const desaDidataIdx = getColIndex(desaHeaders, ['responden didata', 'didata']);

                let prevPclMap = {}, prevPmlMap = {}, prevDesaMap = {};
                if (folders.length > 1) {
                    const prevDate = folders[1];
                    const prevDir = path.join(baseDir, prevDate);

                    const loadPrevNormalized = (fname, callback) => {
                        const fpath = path.join(prevDir, fname);
                        if (fs.existsSync(fpath)) {
                            const wbPrev = xlsx.readFile(fpath);
                            const rowsPrev = xlsx.utils.sheet_to_json(wbPrev.Sheets[wbPrev.SheetNames[0]], { header: 1 });
                            if (rowsPrev.length >= 2) {
                                const h1 = rowsPrev[0];
                                const h2 = rowsPrev[1];
                                const headers = h2.map((val, idx) => {
                                    let parent = h1[idx] !== undefined && h1[idx] !== '' ? String(h1[idx]).trim() : '';
                                    let child = val !== undefined && val !== '' ? String(val).trim() : '';
                                    return (parent && parent !== child) ? `${parent} - ${child}` : (child || parent);
                                });
                                callback(headers, rowsPrev.slice(2));
                            }
                        }
                    };

                    loadPrevNormalized(`rekap_petugas_pcl_${prevDate}.xls`, (headers, rows) => {
                        const pIdx = getColIndex(headers, ['responden didata', 'didata']);
                        if (pIdx !== -1) {
                            rows.forEach(r => { if (r[1]) prevPclMap[String(r[1]).trim()] = parseInt(r[pIdx]) || 0; });
                        }
                    });

                    loadPrevNormalized(`rekap_petugas_pml_${prevDate}.xls`, (headers, rows) => {
                        const pIdx = getColIndex(headers, ['responden didata', 'didata']);
                        if (pIdx !== -1) {
                            rows.forEach(r => { if (r[1]) prevPmlMap[String(r[1]).trim()] = parseInt(r[pIdx]) || 0; });
                        }
                    });

                    loadPrevNormalized(`rekap_wilayah_desa_${prevDate}.xls`, (headers, rows) => {
                        const dIdx = getColIndex(headers, ['responden didata', 'didata']);
                        if (dIdx !== -1) {
                            rows.forEach(r => {
                                const name = String(r[2] || r[1] || '').trim();
                                if (name) prevDesaMap[name] = parseInt(r[dIdx]) || 0;
                            });
                        }
                    });
                }

                const processedPclRows = pclData.map(r => {
                    const name = r[1] ? String(r[1]).trim() : '';
                    const currentVal = pclDidataIdx !== -1 ? (parseInt(r[pclDidataIdx]) || 0) : 0;
                    const harian = prevPclMap[name] !== undefined ? Math.max(0, currentVal - prevPclMap[name]) : currentVal;
                    return { row: r, name, currentVal, harian };
                });

                const sortedByHarianDesc = [...processedPclRows].sort((a, b) => b.harian - a.harian);
                const sortedByHarianAsc = [...processedPclRows].sort((a, b) => a.harian - b.harian);

                topPcl = sortedByHarianDesc.slice(0, 5);
                bottomPcl = sortedByHarianAsc.slice(0, 5);

                processedPclRows.forEach(p => { pclHarianMap[p.name] = p.harian; });
                pmlData.forEach(r => {
                    const name = r[1] ? String(r[1]).trim() : '';
                    const cur = pmlDidataIdx !== -1 ? (parseInt(r[pmlDidataIdx]) || 0) : 0;
                    pmlHarianMap[name] = prevPmlMap[name] !== undefined ? Math.max(0, cur - prevPmlMap[name]) : cur;
                });
                desaData.forEach(r => {
                    const name = String(r[2] || r[1] || '').trim();
                    const cur = desaDidataIdx !== -1 ? (parseInt(r[desaDidataIdx]) || 0) : 0;
                    desaHarianMap[name] = prevDesaMap[name] !== undefined ? Math.max(0, cur - prevDesaMap[name]) : cur;
                });
            }
        }
    } catch (err) {
        console.error("Kesalahan:", err.message);
    }

    res.render('index', {
        activeDate,
        kecHeaders, kecData,
        desaHeaders, desaData,
        pmlHeaders, pmlData,
        pclHeaders, pclData,
        topPcl, bottomPcl,
        pclHarianMap, pmlHarianMap, desaHarianMap
    });
});

app.listen(3000, () => {
    console.log('Server berjalan di http://localhost:3000');
});
