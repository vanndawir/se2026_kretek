const express = require('express');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Konfigurasi View Engine EJS
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// Middleware
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

// Fungsi pembaca file Excel/HTML rekap yang aman
function readExcelSafely(filename) {
    try {
        const filePath = path.join(__dirname, filename);
        if (!fs.existsSync(filePath)) {
            console.error("File tidak ditemukan di path:", filePath);
            return { headers: [], rows: [] };
        }
        
        // Baca file sebagai buffer (karena file .xls ini berformat HTML)
        const fileBuffer = fs.readFileSync(filePath);
        const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        // Baris 0 & 1 adalah header tabel, data riil mulai baris ke-2
        const headers = rawData[1] || rawData[0] || [];
        const rows = rawData.slice(2).filter(r => r && r.length > 0);
        return { headers, rows };
    } catch (e) {
        console.error(`Gagal membaca ${filename}:`, e.message);
        return { headers: [], rows: [] };
    }
}

// ROUTE UTAMA
app.get('/', (req, res) => {
    const pclStore = readExcelSafely('rekap_petugas_pcl_2026-07-31.xls');

    // Proses Top & Bottom PCL untuk progress harian
    let processedPcl = pclStore.rows.map(row => {
        return {
            name: row[1] || '-',
            harian: parseFloat(row[7]) || 0,
            persen: row[8] || '0%'
        };
    }).sort((a, b) => b.harian - a.harian);

    const topPcl = processedPcl.slice(0, 5);
    const bottomPcl = [...processedPcl].reverse().slice(0, 5);

    // Render ke file views/index.ejs
    res.render('index', {
        activeDate: '31 Juli 2026',
        topPcl: topPcl,
        bottomPcl: bottomPcl,
        kecHeaders: [],
        kecData: [],
        desaHeaders: [],
        desaData: [],
        pmlHeaders: [],
        pmlData: [],
        pclHeaders: pclStore.headers,
        pclData: pclStore.rows,
        desaHarianMap: {},
        pmlHarianMap: {},
        pclHarianMap: {}
    });
});

app.listen(PORT, () => {
    console.log(`Server Kretek running on port ${PORT}`);
});
