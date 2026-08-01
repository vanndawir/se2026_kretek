const express = require('express');
const { google } = require('googleapis');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.static(path.join(__dirname, 'public')));

// Konfigurasi Google Sheets ID Anda
const SPREADSHEET_ID = process.env.SPREADSHEET_ID || 'MASUKKAN_SPREADSHEET_ID_ANDA';

// Fungsi helper untuk mengambil data dari Google Sheets
async function getSheetData(range) {
    try {
        const auth = new google.auth.GoogleAuth({
            keyFile: process.env.GOOGLE_APPLICATION_CREDENTIALS ? undefined : './credentials.json', // Sesuaikan jika pakai env
            scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
        });
        
        // Jika menggunakan API Key atau Service Account publik/standar
        const sheets = google.sheets({ version: 'v4', auth });
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: range,
        });
        return response.data.values || [];
    } catch (error) {
        console.error(`Error fetching range ${range}:`, error.message);
        return [];
    }
}

app.get('/', async (req, res) => {
    try {
        // AMBIL DATA TANPA BATASAN ANGKA DI BELAKANG HURUF (A2:Z) AGAR TIDAK ADA DATA YANG TERPOTONG
        const kecRows = await getSheetData('Kecamatan!A2:Z');
        const desaRows = await getSheetData('Desa!A2:Z');
        const pmlRows = await getSheetData('PML!A2:Z');
        const pclRows = await getSheetData('PCL!A2:Z'); // <--- Nanda Mardhiana ada di sini, dijamin terbaca sampai bawah!

        const kecHeaders = kecRows.length > 0 ? kecRows[0] : [];
        const kecData = kecRows.length > 1 ? kecRows.slice(1) : [];

        const desaHeaders = desaRows.length > 0 ? desaRows[0] : [];
        const desaData = desaRows.length > 1 ? desaRows.slice(1) : [];

        const pmlHeaders = pmlRows.length > 0 ? pmlRows[0] : [];
        const pmlData = pmlRows.length > 1 ? pmlRows.slice(1) : [];

        const pclHeaders = pclRows.length > 0 ? pclRows[0] : [];
        const pclData = pclRows.length > 1 ? pclRows.slice(1) : [];

        // Mapping progress harian (jika ada kolom harian)
        let desaHarianMap = {};
        let pmlHarianMap = {};
        let pclHarianMap = {};

        // Proses PCL untuk Top & Bottom Ranking
        let pclListFormatted = [];
        pclData.forEach((row) => {
            // Cari nama secara dinamis di baris
            let name = '';
            for (let c = 0; c < row.length; c++) {
                let val = String(row[c] || '').trim();
                if (val && isNaN(val) && val.length > 2) {
                    name = val;
                    break;
                }
            }
            if (!name) name = row[1] || 'Tanpa Nama';

            // Cari kolom harian (+ Didata)
            let harian = 0;
            for (let i = 0; i < pclHeaders.length; i++) {
                const h = pclHeaders[i] ? String(pclHeaders[i]).toLowerCase() : '';
                if (h.includes('+ didata') || h.includes('harian')) {
                    harian = parseInt(row[i]) || 0;
                    break;
                }
            }

            pclHarianMap[name] = harian;
            pclListFormatted.push({ name, harian });
        });

        // Urutkan untuk Top 5 dan Bottom 5 PCL
        pclListFormatted.sort((a, b) => b.harian - a.harian);
        const topPcl = pclListFormatted.slice(0, 5);
        const bottomPcl = [...pclListFormatted].reverse().slice(0, 5);

        res.render('index', {
            activeDate: new Date().toLocaleDateString('id-ID', { dateStyle: 'full' }),
            kecHeaders, kecData,
            desaHeaders, desaData, desaHarianMap,
            pmlHeaders, pmlData, pmlHarianMap,
            pclHeaders, pclData, pclHarianMap,
            topPcl, bottomPcl
        });
    } catch (err) {
        console.error(err);
        res.status(500).send("Terjadi kesalahan saat memuat data: " + err.message);
    }
});

app.listen(PORT, () => {
    console.log(`Server berjalan di port ${PORT}`);
});
