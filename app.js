const express = require('express');
const XLSX = require('xlsx');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

// Endpoint API untuk mengambil data petugas PCL / PPL
app.get('/api/petugas', (req, res) => {
    try {
        // Path file Excel rekap petugas
        const filePath = path.join(__dirname, 'rekap_petugas_pcl_2026-07-31.xls');
        
        // Membaca workbook (SheetJS otomatis mengenali format HTML di dalam file .xls)
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Gunakan header: 1 untuk membaca baris sebagai array (menghindari error multi-header tabel HTML)
        const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Data aktual mulai dari baris ke-3 (indeks 2), karena baris 0 dan 1 adalah header tabel
        const rows = rawData.slice(2);

        const petugasList = rows
            .filter(row => row && row.length > 1 && row[1]) // Pastikan baris tidak kosong & ada namanya
            .map((row) => {
                return {
                    no: row[0],
                    nama: row[1],
                    email: row[2],
                    role: row[3] ? String(row[3]).trim().toUpperCase() : '', // PCL / PPL
                    subSlv: row[4],
                    target: row[5],
                    didata: row[6],
                    persenDidata: row[8]
                };
            })
            .filter(petugas => {
                // KUNCI UTAMA: Melonggarkan filter agar 'PPL' (seperti Nanda) dan 'PCL' ikut lolos
                const r = petugas.role;
                return r === 'PCL' || r === 'PPL' || r.includes('PCL') || r.includes('PPL');
            });

        res.json({
            success: true,
            total: petugasList.length,
            data: petugasList
        });

    } catch (error) {
        console.error('Gagal membaca file rekap:', error);
        res.status(500).json({ success: false, message: error.message });
    }
});

app.listen(PORT, () => {
    console.log(`Server Sensus Ekonomi Kretek running on port ${PORT}`);
});
