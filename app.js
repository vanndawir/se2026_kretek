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

function readRekapFile(filename) {
    try {
        let filePath = path.join(__dirname, 'data', '2026-07-31', filename);
        if (!fs.existsSync(filePath)) {
            filePath = path.join(__dirname, filename);
        }

        if (!fs.existsSync(filePath)) {
            console.warn(`File tidak ditemukan: ${filename}`);
            return [];
        }
        
        const fileBuffer = fs.readFileSync(filePath);
        const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        const rows = rawData.filter(r => {
            if (!r || r.length === 0) return false;
            const text = r.join(' ').toLowerCase();
            if (text.includes('rekap') || text.includes('nama petugas') || (text.includes('kecamatan') && text.includes('target'))) {
                return false;
            }
            return true;
        });

        return rows.map((row, index) => {
            let label = '-';
            for (let i = 0; i < row.length; i++) {
                const val = row[i];
                if (typeof val === 'string' && val.trim().length > 1 && !val.includes('@') && !val.includes('%') && isNaN(val)) {
                    label = val.trim();
                    break;
                }
            }
            if (label === '-' && row[1]) label = String(row[1]).trim();
            if (label === '-' && row[0]) label = String(row[0]).trim();

            return {
                no: row[0] || (index + 1),
                nama: label,
                target: parseFloat(row[5]) || 0,
                didata: parseFloat(row[6]) || 0,
                harian: parseFloat(row[7]) || 0,
                persen: row[8] || '0%'
            };
        });
    } catch (e) {
        console.error(`Gagal membaca ${filename}:`, e.message);
        return [];
    }
}

app.get('/', (req, res) => {
    try {
        const kecData = readRekapFile('rekap_wilayah_kecamatan_2026-07-31.xls');
        const desaData = readRekapFile('rekap_wilayah_desa_2026-07-31.xls');
        const pmlData = readRekapFile('rekap_petugas_pml_2026-07-31.xls');
        const pclData = readRekapFile('rekap_petugas_pcl_2026-07-31.xls');

        const sortedPclDesc = [...pclData].sort((a, b) => b.harian - a.harian);
        const topPcl = sortedPclDesc.slice(0, 5);
        const bottomPcl = [...pclData].sort((a, b) => a.harian - b.harian).slice(0, 5);

        res.render('index', {
            activeDate: '31 Juli 2026',
            topPcl: topPcl,
            bottomPcl: bottomPcl,
            kecamatan: kecData,
            kecData: kecData,
            desa: desaData,
            desaData: desaData,
            pml: pmlData,
            pmlData: pmlData,
            pcl: pclData,
            pclData: pclData
        });
    } catch (err) {
        console.error("Render Error:", err);
        res.status(500).send("Terjadi kesalahan pada server: " + err.message);
    }
});

app.listen(PORT, () => {
    console.log(`Server Kretek running on port ${PORT}`);
});
