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

// Fungsi pembaca file Excel/HTML rekap yang aman & akurat
function readRekapFile(filename) {
    try {
        const filePath = path.join(__dirname, filename);
        if (!fs.existsSync(filePath)) {
            console.warn(`File tidak ditemukan: ${filename}`);
            return { headers: [], rows: [], objects: [] };
        }
        
        const fileBuffer = fs.readFileSync(filePath);
        const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        const headers = rawData[1] || rawData[0] || [];
        const rawRows = rawData.slice(2).filter(r => r && r.length > 0);

        // Ubah baris array menjadi objek agar mudah diakses di EJS maupun API
        const objects = rawRows.map(row => ({
            no: row[0],
            nama: row[1] || row[2] || '-',
            email: row[2] || '',
            role: row[3] || 'PPL',
            subSlv: row[4] || 0,
            target: parseFloat(row[5]) || 0,
            didata: parseFloat(row[6]) || 0,
            harian: parseFloat(row[7]) || 0,
            persen: row[8] || '0%'
        }));

        return { headers, rows: rawRows, objects };
    } catch (e) {
        console.error(`Gagal membaca ${filename}:`, e.message);
        return { headers: [], rows: [], objects: [] };
    }
}

app.get('/', (req, res) => {
    // Baca data dari file tanggal 31 Juli 2026
    const kecStore = readRekapFile('rekap_wilayah_kecamatan_2026-07-31.xls');
    const desaStore = readRekapFile('rekap_wilayah_desa_2026-07-31.xls');
    const pmlStore = readRekapFile('rekap_petugas_pml_2026-07-31.xls');
    const pclStore = readRekapFile('rekap_petugas_pcl_2026-07-31.xls');

    // Urutkan PCL berdasarkan progress harian (+ Didata)
    const sortedPcl = [...pclStore.objects].sort((a, b) => b.harian - a.harian);
    const topPcl = sortedPcl.slice(0, 5);
    const bottomPcl = [...sortedPcl].sort((a, b) => a.harian - b.harian).slice(0, 5);

    // Kirim data lengkap ke file index.ejs
    res.render('index', {
        activeDate: '31 Juli 2026',
        topPcl: topPcl,
        bottomPcl: bottomPcl,
        kecHeaders: kecStore.headers,
        kecData: kecStore.objects,
        desaHeaders: desaStore.headers,
        desaData: desaStore.objects,
        pmlHeaders: pmlStore.headers,
        pmlData: pmlStore.objects,
        pclHeaders: pclStore.headers,
        pclData: pclStore.objects,
        desaHarianMap: {},
        pmlHarianMap: {},
        pclHarianMap: {}
    });
});

app.listen(PORT, () => {
    console.log(`Server Kretek running on port ${PORT}`);
});
