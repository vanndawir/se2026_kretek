const express = require('express');
const XLSX = require('xlsx');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Konfigurasi View Engine EJS
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// Middleware
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

// Fungsi pembantu untuk membaca file excel / html rekap
function readExcelSafely(filename) {
    try {
        const filePath = path.join(__dirname, filename);
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const rawData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
        
        // Baris 0 & 1 biasanya header, data mulai baris ke-2
        const headers = rawData[1] || rawData[0] || [];
        const rows = rawData.slice(2).filter(r => r && r.length > 0);
        return { headers, rows };
    } catch (e) {
        console.error(`Gagal membaca ${filename}:`, e.message);
        return { headers: [], rows: [] };
    }
}

// ROUTE UTAMA (Menampilkan Halaman Web & Mencegah Cannot GET /)
app.get('/', (req, res) => {
    // Ambil data PCL/PPL dari file terbaru
    const pclStore = readExcelSafely('rekap_petugas_pcl_2026-07-31.xls');

    // Proses data PCL untuk Top & Bottom Progress Harian jika diperlukan
    let processedPcl = pclStore.rows.map(row => {
        return {
            name: row[1] || '-',
            harian: parseFloat(row[7]) || 0 // Kolom harian
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
