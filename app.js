const express = require('express');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Konfigurasi View Engine
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

// Fungsi universal untuk membaca file .xls (format HTML tabel) dengan aman
function readRekapFile(filename) {
    try {
        const filePath = path.join(__dirname, filename);
        if (!fs.existsSync(filePath)) {
            console.warn(`File tidak ditemukan: ${filename}`);
            return { headers: [], rows: [] };
        }
        
        const fileBuffer = fs.readFileSync(filePath);
        const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        // Baris 0 & 1 adalah header ganda tabel, data riil mulai baris ke-2
        const headers = rawData[1] || rawData[0] || [];
        const rows = rawData.slice(2).filter(r => r && r.length > 0);
        return { headers, rows };
    } catch (e) {
        console.error(`Gagal membaca ${filename}:`, e.message);
        return { headers: [], rows: [] };
    }
}

app.get('/', (req, res) => {
    // 1. Baca semua file rekap tanggal 31 Juli 2026
    const kecStore = readRekapFile('rekap_wilayah_kecamatan_2026-07-31.xls');
    const desaStore = readRekapFile('rekap_wilayah_desa_2026-07-31.xls');
    const pmlStore = readRekapFile('rekap_petugas_pml_2026-07-31.xls');
    const pclStore = readRekapFile('rekap_petugas_pcl_2026-07-31.xls');

    // 2. Olah Top 5 & Bottom 5 PCL berdasarkan progress harian (+ Didata di indeks kolom ke-7)
    let processedPcl = pclStore.rows.map(row => {
        return {
            no: row[0],
            name: row[1] || '-',
            email: row[2] || '-',
            role: row[3] || 'PPL',
            subSlv: row[4] || 0,
            target: parseFloat(row[5]) || 0,
            didata: parseFloat(row[6]) || 0,
            harian: parseFloat(row[7]) || 0,
            persen: row[8] || '0%'
        };
    });

    // Urutkan untuk Top 5 (tertinggi) dan Bottom 5 (terendah)
    const sortedDesc = [...processedPcl].sort((a, b) => b.harian - a.harian);
    const topPcl = sortedDesc.slice(0, 5);
    const bottomPcl = [...processedPcl].sort((a, b) => a.harian - b.harian).slice(0, 5);

    // 3. Render ke views/index.ejs dengan data lengkap
    res.render('index', {
        activeDate: '31 Juli 2026',
        topPcl: topPcl,
        bottomPcl: bottomPcl,
        kecHeaders: kecStore.headers,
        kecData: kecStore.rows,
        desaHeaders: desaStore.headers,
        desaData: desaStore.rows,
        pmlHeaders: pmlStore.headers,
        pmlData: pmlStore.rows,
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
