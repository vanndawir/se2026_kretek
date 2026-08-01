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
        let filePath = path.join(__dirname, filename);
        if (!fs.existsSync(filePath)) {
            filePath = path.join(__dirname, 'data', '2026-07-31', filename);
        }

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
        const rows = rawData.slice(2).filter(r => r && r.length > 0);

        const objects = rows.map(row => {
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
                no: row[0] || '',
                nama: label,
                name: label,
                Nama: label,
                namaPetugas: label,
                wilayah: label,
                kecamatan: label,
                desa: label,
                email: row[2] || '-',
                role: row[3] || 'PPL',
                subSlv: row[4] || 0,
                target: parseFloat(row[5]) || 0,
                didata: parseFloat(row[6]) || 0,
                harian: parseFloat(row[7]) || 0,
                persen: row[8] || '0%'
            };
        });

        return { headers, rows, objects };
    } catch (e) {
        console.error(`Gagal membaca ${filename}:`, e.message);
        return { headers: [], rows: [], objects: [] };
    }
}

app.get('/', (req, res) => {
    try {
        const kecStore = readRekapFile('rekap_wilayah_kecamatan_2026-07-31.xls');
        const desaStore = readRekapFile('rekap_wilayah_desa_2026-07-31.xls');
        const pmlStore = readRekapFile('rekap_petugas_pml_2026-07-31.xls');
        const pclStore = readRekapFile('rekap_petugas_pcl_2026-07-31.xls');

        const sortedPclDesc = [...pclStore.objects].sort((a, b) => b.harian - a.harian);
        const topPcl = sortedPclDesc.slice(0, 5);
        const bottomPcl = [...pclStore.objects].sort((a, b) => a.harian - b.harian).slice(0, 5);

        res.render('index', {
            activeDate: '31 Juli 2026',
            topPcl: topPcl,
            bottomPcl: bottomPcl,
            // Kirim kedua varian nama variabel agar cocok dengan EJS apa pun
            kecamatan: kecStore.objects,
            kecData: kecStore.objects,
            kecHeaders: kecStore.headers,
            desa: desaStore.objects,
            desaData: desaStore.objects,
            desaHeaders: desaStore.headers,
            pml: pmlStore.objects,
            pmlData: pmlStore.objects,
            pmlHeaders: pmlStore.headers,
            pcl: pclStore.objects,
            pclData: pclStore.objects,
            pclHeaders: pclStore.headers,
            desaHarianMap: {},
            pmlHarianMap: {},
            pclHarianMap: {}
        });
    } catch (err) {
        console.error("Render Error:", err);
        res.status(500).send("Terjadi kesalahan pada server: " + err.message);
    }
});

app.listen(PORT, () => {
    console.log(`Server Kretek running on port ${PORT}`);
});
