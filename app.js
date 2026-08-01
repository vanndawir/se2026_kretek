const express = require('express');
const fs = require('fs');
const path = require('path');
const cheerio = require('cheerio');

const app = express();
const PORT = process.env.PORT || 3000;

app.set('view engine', 'ejs');
app.use(express.static('public'));

// Fungsi helper untuk mencari file terbaru di root ATAU di folder 'data'
app.get('/', (req, res) => {
    function getLatestFile(prefix) {
        try {
            // Cek di folder utama (__dirname)
            let searchDir = __dirname;
            let files = fs.readdirSync(searchDir);
            let matchedFiles = files.filter(f => f.startsWith(prefix) && f.endsWith('.xls'));

            // Jika tidak ketemu di root, cek di dalam folder 'data'
            if (matchedFiles.length === 0) {
                const dataDir = path.join(__dirname, 'data');
                if (fs.existsSync(dataDir)) {
                    files = fs.readdirSync(dataDir);
                    matchedFiles = files.filter(f => f.startsWith(prefix) && f.endsWith('.xls'));
                    if (matchedFiles.length > 0) {
                        searchDir = dataDir;
                    }
                }
            }

            if (matchedFiles.length === 0) return { filename: '', dir: '' };

            // Urutkan berdasarkan tanggal (YYYY-MM-DD) terbaru
            matchedFiles.sort().reverse();
            return { filename: matchedFiles[0], dir: searchDir };
        } catch (error) {
            console.error(`Gagal mencari ${prefix}:`, error);
            return { filename: '', dir: '' };
        }
    }

    function parseHtmlXls(fileInfo) {
        if (!fileInfo || !fileInfo.filename) return { headers: [], rows: [], filename: 'Tidak ditemukan' };

        const filePath = path.join(fileInfo.dir, fileInfo.filename);
        if (!fs.existsSync(filePath)) return { headers: [], rows: [], filename: 'Tidak ditemukan' };

        const htmlContent = fs.readFileSync(filePath, 'utf-8');
        const $ = cheerio.load(htmlContent);
        
        const headers = [];
        const rows = [];

        $('table tr').each((i, el) => {
            const rowData = [];
            $(el).find('th, td').each((j, cell) => {
                rowData.push($(cell).text().trim());
            });
            
            if (rowData.length > 0) {
                if (i < 2) {
                    headers.push(rowData);
                } else {
                    rows.push(rowData);
                }
            }
        });

        return { headers, rows, filename: fileInfo.filename };
    }

    const pclData = parseHtmlXls(getLatestFile('rekap_petugas_pcl_'));
    const pmlData = parseHtmlXls(getLatestFile('rekap_petugas_pml_'));
    const desaData = parseHtmlXls(getLatestFile('rekap_wilayah_desa_'));
    const kecData = parseHtmlXls(getLatestFile('rekap_wilayah_kecamatan_'));

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
