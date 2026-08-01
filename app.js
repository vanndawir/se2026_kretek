const express = require('express');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

// Fungsi universal pembaca file Excel (.xls format HTML) di folder data/2026-07-31/
function readRekapFile(filename) {
    try {
        const filePath = path.join(__dirname, 'data', '2026-07-31', filename);
        if (!fs.existsSync(filePath)) {
            console.warn(`File tidak ditemukan: ${filePath}`);
            return [];
        }
        
        const fileBuffer = fs.readFileSync(filePath);
        const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        // Baris 0 & 1 adalah header ganda tabel, data riil mulai baris ke-2
        const rows = rawData.slice(2).filter(r => r && r.length > 0);

        return rows.map(row => ({
            no: row[0],
            nama: row[1] || '-',         // Kolom Nama Petugas
            email: row[2] || '-',        // Kolom Email
            role: row[3] || 'PPL',       // Kolom Role (PPL/PML)
            subSlv: row[4] || 0,         // Sub-SLS Diampu
            target: parseFloat(row[5]) || 0,  // Target Prelist
            didata: parseFloat(row[6]) || 0,  // Responden Didata
            harian: parseFloat(row[7]) || 0,  // + Didata (Progress Harian)
            persen: row[8] || '0%'       // % Didata
        }));
    } catch (e) {
        console.error(`Gagal membaca ${filename}:`, e.message);
        return [];
    }
}

app.get('/', (req, res) => {
    // Baca data dari folder data/2026-07-31/
    const kecData = readRekapFile('rekap_wilayah_kecamatan_2026-07-31.xls');
    const desaData = readRekapFile('rekap_wilayah_desa_2026-07-31.xls');
    const pmlData = readRekapFile('rekap_petugas_pml_2026-07-31.xls');
    const pclData = readRekapFile('rekap_petugas_pcl_2026-07-31.xls');

    // Urutkan PCL berdasarkan progress harian tertinggi untuk Top 5
    const sortedPclDesc = [...pclData].sort((a, b) => b.harian - a.harian);
    const topPcl = sortedPclDesc.slice(0, 5);

    // Urutkan PCL berdasarkan progress harian terendah untuk 5 PCL Progress Terendah
    const sortedPclAsc = [...pclData].sort((a, b) => a.harian - b.harian);
    const bottomPcl = sortedPclAsc.slice(0, 5);

    res.render('index', {
        activeDate: '31 Juli 2026',
        topPcl: topPcl,
        bottomPcl: bottomPcl,
        kecData: kecData,
        desaData: desaData,
        pmlData: pmlData,
        pclData: pclData,
        desaHarianMap: {},
        pmlHarianMap: {},
        pclHarianMap: {}
    });
});

app.listen(PORT, () => {
    console.log(`Server Kretek running on port ${PORT}`);
});
