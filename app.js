const express = require('express');
const { google } = require('googleapis');
const xlsx = require('xlsx');
const path = require('path');
const http = require('http');
const { Server } = require('socket.io');
const sqlite3 = require('sqlite3').verbose();
const { GoogleGenAI } = require('@google/genai');

const app = express();
const server = http.createServer(app);
const io = new Server(server);

// Inisialisasi Gemini AI (Menggunakan API Key Anda)
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY || });

app.set('view engine', 'ejs');
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Inisialisasi Database SQLite untuk Curhat Lapangan
const db = new sqlite3.Database(path.join(__dirname, 'chat.db'), (err) => {
    if (err) {
        console.error('❌ Gagal terhubung ke database:', err.message);
    } else {
        console.log('✅ Berhasil terhubung ke database SQLite (chat.db).');
    }
});

// Buat tabel chats jika belum ada
db.run(`CREATE TABLE IF NOT EXISTS chats (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT,
    message TEXT,
    sender_type TEXT,
    timestamp TEXT
)`);

// ID Folder Utama Google Drive
const MAIN_FOLDER_ID = '1kyl4AOQCfu8r1c4qDMZZF28bxWU2dO17';

let drive;
try {
    drive = google.drive({
        version: 'v3',
        auth: new google.auth.GoogleAuth({
            keyFile: path.join(__dirname, 'service-account.json'),
            scopes: ['https://www.googleapis.com/auth/drive.readonly']
        })
    });
    console.log('✅ Konfigurasi Google Auth berhasil.');
} catch (e) {
    console.error('❌ Gagal memuat service-account.json:', e.message);
}

// Fungsi Parser Angka Mutlak
const parseStrictNumber = (val) => {
    if (val === undefined || val === null || val === '-' || val === '') return 0;
    if (typeof val === 'number') return Math.round(val);
    let str = String(val).trim();
    if (str.includes('.') && !str.includes(',')) {
        let parts = str.split('.');
        if (parts.length > 1 && parts[parts.length - 1].length === 3) str = parts.join('');
    }
    str = str.replace(/[^0-9]/g, '');
    return parseInt(str, 10) || 0;
};

const getBulletproofColIndex = (rows, exactKeywords, partialKeywords = []) => {
    let maxCols = 0;
    const maxRowsToScan = Math.min(rows.length, 15);
    for (let i = 0; i < maxRowsToScan; i++) {
        if (rows[i] && rows[i].length > maxCols) maxCols = rows[i].length;
    }

    const allKeywords = [...exactKeywords, ...partialKeywords].sort((a, b) => b.length - a.length);

    for (let kw of allKeywords) {
        let cleanKw = kw.toLowerCase().replace(/[^a-z0-9+]/g, '');
        if (!cleanKw) continue;
        for (let col = 0; col < maxCols; col++) {
            let colText = '';
            for (let row = 0; row < maxRowsToScan; row++) {
                if (rows[row] && rows[row][col] !== undefined) {
                    colText += String(rows[row][col]).toLowerCase().replace(/[^a-z0-9+]/g, '');
                }
            }
            if (colText.includes(cleanKw)) return col;
        }
    }
    return -1;
};

const getValidName = (r) => {
    if (r[1] !== undefined) {
        let val = String(r[1]).trim();
        if (val && !val.includes('@') && isNaN(val) && val.length > 2) return val;
    }
    for (let idx = 1; idx < Math.min(r.length, 5); idx++) {
        if (r[idx] !== undefined) {
            let val = String(r[idx]).trim();
            if (val && !val.includes('@') && isNaN(val) && val.length > 3) return val;
        }
    }
    return '';
};

const fetchExcelRows = async (folderId, keywords) => {
    if (!drive) return [];
    
    const searchInParent = async (parentId) => {
        try {
            const qQuery = parentId ? `'${parentId}' in parents and trashed = false` : `trashed = false`;
            const res = await drive.files.list({
                q: qQuery, pageSize: 100, fields: 'files(id, name, mimeType)', orderBy: 'name desc'
            });
            return res.data.files || [];
        } catch (e) { return []; }
    };

    try {
        let files = await searchInParent(folderId);
        if (files.length === 0 && folderId !== MAIN_FOLDER_ID) files = await searchInParent(MAIN_FOLDER_ID);
        if (files.length === 0) files = await searchInParent(null);

        files.sort((a, b) => (b.name || '').localeCompare(a.name || ''));
        let matchedFile = files.find(f => keywords.some(kw => f.name && f.name.toLowerCase().includes(kw.toLowerCase())));

        if (!matchedFile) return [];

        const fileId = matchedFile.id;
        const mimeType = matchedFile.mimeType;
        let buffer;

        if (mimeType === 'application/vnd.google-apps.spreadsheet') {
            const exportRes = await drive.files.export({
                fileId: fileId, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            }, { responseType: 'arraybuffer' });
            buffer = Buffer.from(exportRes.data);
        } else {
            const mediaRes = await drive.files.get({ fileId: fileId, alt: 'media' }, { responseType: 'arraybuffer' });
            buffer = Buffer.from(mediaRes.data);
        }

        const wb = xlsx.read(buffer, { type: 'buffer' });
        return xlsx.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 }) || [];
    } catch (err) {
        return [];
    }
};

const parseDesaData = (rows) => {
    if (!rows || rows.length === 0) return [];
    let targetIdx = getBulletproofColIndex(rows, ['target', 'jumlahusaha']);
    let didataIdx = getBulletproofColIndex(rows, ['respondendidata', 'realisasi', 'selesai']);
    
    let dataStartIdx = rows.findIndex((r, idx) => idx > 1 && r && (r[0] === 1 || String(r[0]).trim() === '1'));
    if (dataStartIdx === -1) dataStartIdx = 3;

    return rows.slice(dataStartIdx).map(r => {
        let name = getValidName(r);
        if (!name) return null;
        let tVal = targetIdx !== -1 ? parseStrictNumber(r[targetIdx]) : 0;
        let dVal = didataIdx !== -1 ? parseStrictNumber(r[didataIdx]) : 0;
        return {
            name,
            target: tVal,
            didata: dVal
        };
    }).filter(item => item && item.name.toLowerCase() !== 'jumlah' && item.name.toLowerCase() !== 'total');
};

const parsePetugas = (rows, role) => {
    if (!rows || rows.length === 0) return [];

    let dataStartIdx = rows.findIndex((r, idx) => idx > 1 && r && (r[0] === 1 || String(r[0]).trim() === '1'));
    if (dataStartIdx === -1) dataStartIdx = 3;

    let targetIdx = getBulletproofColIndex(rows, ['target', 'prelist'], ['beban']);
    let didataIdx = getBulletproofColIndex(rows, ['respondendidata', 'realisasi'], ['didata']);
    let harianIdx = getBulletproofColIndex(rows, ['+didata', 'harian'], ['+']);
    let pctIdx = getBulletproofColIndex(rows, ['%didata', 'persentasedidata'], ['%']);
    let didataDraftIdx = getBulletproofColIndex(rows, ['didatadraft', 'didata+draft'], ['draft']);

    if (targetIdx === -1) targetIdx = 5;
    if (didataIdx === -1) didataIdx = 6;
    if (harianIdx === -1) harianIdx = 7;
    if (pctIdx === -1) pctIdx = 8;
    if (didataDraftIdx === -1) didataDraftIdx = 9;

    let compositeHeaders = [];
    let maxCols = 0;
    for (let r = 0; r < dataStartIdx; r++) {
        if (rows[r] && rows[r].length > maxCols) maxCols = rows[r].length;
    }

    for (let c = 0; c < maxCols; c++) {
        let parts = [];
        for (let r = 0; r < dataStartIdx; r++) {
            if (rows[r] && rows[r][c] !== undefined && rows[r][c] !== null) {
                let txt = String(rows[r][c]).trim().replace(/[\r\n]+/g, ' ');
                if (txt && !parts.includes(txt)) parts.push(txt);
            }
        }
        if (c === targetIdx) compositeHeaders[c] = 'Target Prelist';
        else if (c === didataIdx) compositeHeaders[c] = 'Responden Didata';
        else if (c === harianIdx) compositeHeaders[c] = '+ Didata';
        else if (c === pctIdx) compositeHeaders[c] = '% Didata';
        else if (c === didataDraftIdx) compositeHeaders[c] = 'Didata + Draft';
        else compositeHeaders[c] = parts.join(' ') || `Kolom_${c}`;
    }

    return rows.slice(dataStartIdx).map(r => {
        let name = getValidName(r);
        if (!name) return null;

        let targetVal = parseStrictNumber(r[targetIdx]);
        let didataVal = parseStrictNumber(r[didataIdx]);
        let harianVal = parseStrictNumber(r[harianIdx]);
        let didataDraftVal = parseStrictNumber(r[didataDraftIdx]);
        if (didataDraftVal < didataVal) didataDraftVal = didataVal;

        let rawPct = r[pctIdx];
        let formattedPct = '0.0%';
        let numericPct = 0;

        if (rawPct !== undefined && rawPct !== null && rawPct !== '' && rawPct !== '-') {
            let strPct = String(rawPct).trim().replace('%', '');
            let num = parseFloat(strPct.replace(',', '.'));
            if (!isNaN(num)) {
                numericPct = num <= 1 ? num * 100 : num;
                if (numericPct > 100) numericPct = 100;
                formattedPct = numericPct.toFixed(1) + '%';
            }
        } else if (targetVal > 0) {
            numericPct = Math.min(100, (didataVal / targetVal) * 100);
            formattedPct = numericPct.toFixed(1) + '%';
        }

        let rawPctDraft = r[10];
        let formattedPctDraft = '0.0%';
        let numericPctDraft = 0;

        if (rawPctDraft !== undefined && rawPctDraft !== null && rawPctDraft !== '' && rawPctDraft !== '-') {
            let strVal = String(rawPctDraft).trim();
            let num = parseFloat(strVal.replace('%', '').replace(',', '.'));
            if (!isNaN(num)) {
                if (typeof rawPctDraft === 'number' && !strVal.includes('%')) {
                    numericPctDraft = num * 100;
                } else {
                    numericPctDraft = num;
                }
                formattedPctDraft = numericPctDraft.toFixed(1) + '%';
            } else {
                formattedPctDraft = strVal;
            }
        }

        let rowObj = { 
            name, 
            target: targetVal, 
            didata: didataVal, 
            harian: harianVal, 
            progress: formattedPct,
            progressVal: numericPct,
            didataDraft: didataDraftVal,
            progressDraft: formattedPctDraft,
            progressDraftVal: numericPctDraft,
            rawCells: []
        };
        
        for (let idx = 11; idx <= 16; idx++) {
            let h = compositeHeaders[idx] || `Kolom_${idx}`;
            let rawVal = (idx < r.length && r[idx] !== undefined && r[idx] !== null) ? r[idx] : '-';
            let val = (rawVal !== '-' && !isNaN(rawVal)) ? parseStrictNumber(rawVal) : rawVal;
            rowObj.rawCells.push({ header: h, value: val });
        }

        return rowObj;
    }).filter(item => item && item.name.toLowerCase() !== 'jumlah' && item.name.toLowerCase() !== 'total');
};

io.on('connection', (socket) => {
    db.all(`SELECT * FROM (SELECT * FROM chats ORDER BY id DESC LIMIT 50) ORDER BY id ASC`, [], (err, rows) => {
        if (!err) {
            socket.emit('init_chat', rows);
        } else {
            socket.emit('init_chat', []);
        }
    });

    socket.on('send_message', (data) => {
        const sender = data.sender || 'Petugas Anonim';
        const message = data.message;
        const senderType = data.sender_type || 'user'; 
        const timestamp = new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });

        db.run(`INSERT INTO chats (name, message, sender_type, timestamp) VALUES (?, ?, ?, ?)`, [sender, message, senderType, timestamp], function(err) {
            if (!err) {
                const messageData = {
                    id: this.lastID,
                    name: sender,
                    message,
                    sender_type: senderType,
                    timestamp
                };
                io.emit('new_message', messageData);

                if (senderType === 'user') {
                    setTimeout(async () => {
                        try {
                            const response = await ai.models.generateContent({
                                model: 'gemini-2.5-flash',
                                contents: `Kamu adalah asisten virtual yang ramah, empatik, dan suportif untuk para petugas lapangan Sensus Ekonomi 2026 (SE2026) di Kapanewon Kretek, Bantul. Petugas menyampaikan kendala: "${message}". Berikan tanggapan yang menyemangati, solutif secara umum, dan gunakan bahasa Indonesia yang santai namun profesional.`
                            });

                            const aiReplyText = response.text || "Tetap semangat rekan data! Admin utama akan segera meninjau laporanmu.";
                            const aiTimestamp = new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });

                            db.run(
                                `INSERT INTO chats (name, message, sender_type, timestamp) VALUES (?, ?, ?, ?)`,
                                ['Asisten AI Kretek 🤖', aiReplyText, 'ai', aiTimestamp],
                                function (aiErr) {
                                    if (!aiErr) {
                                        io.emit('new_message', {
                                            id: this.lastID,
                                            name: 'Asisten AI Kretek 🤖',
                                            message: aiReplyText,
                                            sender_type: 'ai',
                                            timestamp: aiTimestamp
                                        });
                                    }
                                }
                            );
                        } catch (aiError) {
                            console.error("Gagal memproses Auto-Reply AI:", aiError);
                        }
                    }, 2000);
                }
            }
        });
    });
});

app.get('/', async (req, res) => {
    let activeDate = '2026-08-03';
    let persentaseKretek = '80.9';
    let totalTargetKretek = 14764;
    let totalDidataKretek = 11939;
    let topPcl = [], bottomPcl = [], parsedDesa = [], detailedPml = [], detailedPcl = [];

    try {
        if (drive) {
            let targetFolderId = MAIN_FOLDER_ID;
            try {
                let folderRes = await drive.files.list({
                    q: `'${MAIN_FOLDER_ID}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false`,
                    orderBy: 'name desc',
                    pageSize: 10
                });
                const folders = folderRes.data.files || [];
                if (folders.length > 0) {
                    const validFolder = folders.find(f => /^\d{4}-\d{2}-\d{2}$/.test(f.name.trim())) || folders[0];
                    activeDate = validFolder.name;
                    targetFolderId = validFolder.id;
                }
            } catch (e) {}

            const [kecRows, desaRows, pmlRows, pclRows] = await Promise.all([
                fetchExcelRows(targetFolderId, ['rekap_wilayah_kecamatan', 'kecamatan', 'kec']),
                fetchExcelRows(targetFolderId, ['rekap_wilayah_desa', 'desa', 'kalurahan']),
                fetchExcelRows(targetFolderId, ['rekap_petugas_pml', 'pml', 'petugas_pml']),
                fetchExcelRows(targetFolderId, ['rekap_petugas_pcl', 'pcl', 'petugas_pcl'])
            ]);

            if (kecRows && kecRows.length > 0) {
                let targetIdx = getBulletproofColIndex(kecRows, ['target', 'jumlahusaha']);
                let didataIdx = getBulletproofColIndex(kecRows, ['respondendidata', 'realisasi', 'selesai']);
                let kretekRow = kecRows.find(r => r.join(' ').toLowerCase().includes('kretek') || r.join(' ').toLowerCase().includes('3402030'));
                
                if (kretekRow) {
                    totalTargetKretek = targetIdx !== -1 ? parseStrictNumber(kretekRow[targetIdx]) : 14764;
                    totalDidataKretek = didataIdx !== -1 ? parseStrictNumber(kretekRow[didataIdx]) : 11939;
                    if (totalTargetKretek > 0) persentaseKretek = ((totalDidataKretek / totalTargetKretek) * 100).toFixed(1);
                }
            }

            parsedDesa = parseDesaData(desaRows);
            detailedPml = parsePetugas(pmlRows, 'pml');
            detailedPcl = parsePetugas(pclRows, 'pcl');
        }
    } catch (err) {}

    res.render('index', {
        activeDate, persentaseKretek, totalTargetKretek, totalDidataKretek,
        topPcl, bottomPcl, parsedDesa, detailedPml, detailedPcl
    });
});

server.listen(3000, () => {
    console.log('🚀 Server berjalan sukses di http://localhost:3000');
});