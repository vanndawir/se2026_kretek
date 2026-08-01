const express = require('express');
const fs = require('fs');
const path = require('path');
const cheerio = require('cheerio');

const app = express();
app.set('view engine', 'ejs');
app.use(express.static('public'));

// Fungsi helper untuk mencari file terbaru berdasarkan pola (prefix) dan tanggal di nama file (YYYY-MM-DD)
function getLatestFile(prefix) {
    const files = fs.readdirSync(__dirname);
    // Filter file yang diawali prefix dan berakhiran .xls
    const matchedFiles = files.filter(f => f.startsWith(prefix) && f.endsWith('.xls'));
    
    if (matchedFiles.length === 0) return null;

    // Urutkan berdasarkan nama file secara descending (karena format YYYY-MM-DD terurut secara leksikografis)
    matchedFiles.sort().reverse();
    return matchedFiles[0];
}

// Fungsi helper untuk memparsing file .xls yang berformat HTML tabel
function parseHtmlXls(filePath) {
    if (!filePath || !fs.existsSync(filePath)) return { headers: [], rows: [] };

    const htmlContent = fs.readFileSync(filePath, 'utf-8');
    const $ = cheerio.load(htmlContent);
    
    const rows = [];
    const headers = [];

    // Ambil baris header (biasanya di dalam <thead> tr)
    $('table tr').each((i, el) => {
        const rowData = [];
        $(el).find('th, td').each((j, cell) => {
            rowData.push($(cell).text().trim());
        });
        
        if (i < 2) {
            headers.push(rowData); // Jika ada multi-row header
        } else {
            rows.push(rowData);
        }
    });

    return { headers, rows };
}

app.get('/', (req, res) => {
    // Auto-detect file rekap terbaru
    const latestPcl = getLatestFile('rekap_petugas_pcl_');
    const latestPml = getLatestFile('rekap_petugas_pml_');
    const latestDesa = getLiveFile = getLatestFile('rekap_wilayah_desa_');
    const latestKec = getLatestFile('rekap_wilayah_kecamatan_');

    const pclData = parseHtmlXls(latestPcl);
    const pmlData = parseHtmlXls(latestPml);
    const desaData = parseHtmlXls(latestDesa);
    const kecData = parseHtmlXls(latestKec);

    res.render('index', {
        latestPcl,
        latestPml,
        latestDesa,
        latestKec,
        pclData,
        pmlData,
        desaData,
        kecData
    });
});

app.listen(3000, () => {
    console.log('Server berjalan di http://localhost:3000');
});
