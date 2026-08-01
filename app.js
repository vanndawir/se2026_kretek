const express = require('express');
const fs = require('fs');
const path = require('path');
const cheerio = require('cheerio');

const app = express();
const PORT = process.env.PORT || 3000;

app.set('view engine', 'ejs');
app.use(express.static('public'));

// Arahkan ke folder 'data' tempat file .xls berada
const dataDir = path.join(__dirname, 'data');

// Fungsi helper untuk mencari file terbaru di dalam folder data
function getLatestFile(prefix) {
    try {
        if (!fs.existsSync(dataDir)) return null;
        const files = fs.readdirSync(dataDir);
        const matchedFiles = files.filter(f => f.startsWith(prefix) && f.endsWith('.xls'));
        
        if (matchedFiles.length === 0) return null;

        // Urutkan berdasarkan tanggal (YYYY-MM-DD) terbaru
        matchedFiles.sort().reverse();
        return matchedFiles[0];
    } catch (error) {
        console.error(`Gagal membaca folder data untuk ${prefix}:`, error);
        return null;
    }
}

// Fungsi helper untuk memparsing file .xls (HTML table)
function parseHtmlXls(filename) {
    if (!filename) return { headers: [], rows: [], filename: '' };

    const filePath = path.join(dataDir, filename);
    if (!fs.existsSync(filePath)) return { headers: [], rows: [], filename: '' };

    const htmlContent = fs.readFileSync(filePath, 'utf-8');
    const $ = cheerio.load(htmlContent);
    
    const headers = [];
    const rows = [];

    $('table tr').each((i, el) => {
        const rowData = [];
        $(el).find('th, td').each((j, cell) => {
            rowData.push($(cell).text().trim());
        });
        
        if (i < 2) {
            headers.push(rowData);
        } else {
            rows.push(rowData);
        }
    });

    return { headers, rows, filename };
}

// Route Utama
app.get('/', (req, res) => {
    const latestPcl = getLatestFile('rekap_petugas_pcl_');
    const latestPml = getLatestFile('rekap_petugas_pml_');
    const latestDesa = getLatestFile('rekap_wilayah_desa_');
    const latestKec = getLatestFile('rekap_wilayah_kecamatan_');

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
