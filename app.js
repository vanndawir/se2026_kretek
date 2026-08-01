const express = require('express');
const fs = require('fs');
const path = require('path');
const cheerio = require('cheerio');

const app = express();
const PORT = process.env.PORT || 3000;

app.set('view engine', 'ejs');
app.use(express.static('public'));

// Fungsi helper untuk mencari file terbaru berdasarkan prefix dan tanggal (YYYY-MM-DD)
function getLatestFile(prefix) {
    try {
        const files = fs.readdirSync(__dirname);
        // Filter file yang diawali prefix tertentu dan berakhiran .xls
        const matchedFiles = files.filter(f => f.startsWith(prefix) && f.endsWith('.xls'));
        
        if (matchedFiles.length === 0) return null;

        // Urutkan secara descending karena format tanggal YYYY-MM-DD otomatis terurut leksikografis
        matchedFiles.sort().reverse();
        return matchedFiles[0];
    } catch (error) {
        console.error(`Gagal membaca direktori untuk prefix ${prefix}:`, error);
        return null;
    }
}

// Fungsi helper untuk memparsing file .xls yang berformat HTML tabel
function parseHtmlXls(filename) {
    if (!filename) return { headers: [], rows: [] };

    const filePath = path.join(__dirname, filename);
    if (!fs.existsSync(filePath)) return { headers: [], rows: [] };

    const htmlContent = fs.readFileSync(filePath, 'utf-8');
    const $ = cheerio.load(htmlContent);
    
    const headers = [];
    const rows = [];

    // Mengambil baris dari tabel HTML
    $('table tr').each((i, el) => {
        const rowData = [];
        $(el).find('th, td').each((j, cell) => {
            rowData.push($(cell).text().trim());
        });
        
        // Memisahkan baris header dan baris data
        if (i < 2) {
            headers.push(rowData); // Mengakomodasi multi-row header jika ada
        } else {
            rows.push(rowData);
        }
    });

    return { headers, rows, filename };
}

// Route Utama
app.get('/', (req, res) => {
    // Auto-detect file rekap terbaru berdasarkan tanggal
    const latestPcl = getLatestFile('rekap_petugas_pcl_');
    const latestPml = getLatestFile('rekap_petugas_pml_');
    const latestDesa = getLatestFile('rekap_wilayah_desa_');
    const latestKec = getLatestFile('rekap_wilayah_kecamatan_');

    // Parse data dari masing-masing file terbaru
    const pclData = parseHtmlXls(latestPcl);
    const pmlData = parseHtmlXls(latestPml);
    const desaData = parseHtmlXls(latestDesa);
    const kecData = parseHtmlXls(latestKec);

    res.render('index', {
        pclData,
        pmlData,
        desaData,
        kecData
    });
});

app.listen(PORT, () => {
    console.log(`Server berjalan di http://localhost:${PORT}`);
});
