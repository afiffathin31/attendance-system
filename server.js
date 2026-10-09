const express = require('express');
const cors = require('cors');
const cookieParser = require('cookie-parser');
const path = require('path');
const fs = require('fs');
const QRCode = require('qrcode');

const db = require('./database/db');
const analyticsService = require('./services/analyticsService');
const excelService = require('./services/excelService');
const syncService = require('./services/syncService');

const CONFIG_PATH = path.join(__dirname, 'config.json');

function getConfig() {
  if (fs.existsSync(CONFIG_PATH)) {
    try {
      return JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf-8'));
    } catch (e) {
      console.error('Error reading config.json', e);
    }
  }
  return {
    port: 3000,
    adminPassword: 'admin123',
    officeStartTime: '08:00',
    toleranceMinutes: 15
  };
}

function saveConfig(cfg) {
  fs.writeFileSync(CONFIG_PATH, JSON.stringify(cfg, null, 2), 'utf-8');
}

const app = express();
const config = getConfig();

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(cookieParser());

// Serve static frontend files
app.use(express.static(path.join(__dirname, 'public')));

// Simple Auth Middleware
function requireAdmin(req, res, next) {
  const token = req.cookies.admin_session || req.headers['authorization'];
  if (token === 'admin-authenticated-token') {
    return next();
  }
  return res.status(401).json({ error: 'Akses ditolak. Silakan login sebagai Admin.' });
}

// ----------------------------------------------------------------------------
// AUTH ENDPOINTS
// ----------------------------------------------------------------------------
app.post('/api/admin/login', (req, res) => {
  const { password } = req.body;
  const currentConfig = getConfig();
  if (password === currentConfig.adminPassword) {
    res.cookie('admin_session', 'admin-authenticated-token', {
      httpOnly: true,
      maxAge: 24 * 60 * 60 * 1000 // 1 day
    });
    return res.json({ success: true, message: 'Login berhasil.' });
  }
  return res.status(401).json({ error: 'Kata sandi salah!' });
});

app.get('/api/admin/check', (req, res) => {
  const token = req.cookies.admin_session;
  if (token === 'admin-authenticated-token') {
    return res.json({ authenticated: true });
  }
  return res.json({ authenticated: false });
});

app.post('/api/admin/logout', (req, res) => {
  res.clearCookie('admin_session');
  res.json({ success: true, message: 'Logout berhasil.' });
});

// ----------------------------------------------------------------------------
// ATTENDANCE ENDPOINTS
// ----------------------------------------------------------------------------

// List available departments
app.get('/api/attendance/departments', (req, res) => {
  res.json({ departments: db.getDepartments() });
});

// Offline Attendance Submission
app.post('/api/attendance/offline', async (req, res) => {
  try {
    const { name, department } = req.body;
    if (!name || !department) {
      return res.status(400).json({ error: 'Nama dan Departemen wajib diisi!' });
    }

    const currentConfig = getConfig();
    const officeStartTime = currentConfig.officeStartTime || '08:00';
    const tolerance = parseInt(currentConfig.toleranceMinutes, 10) || 15;

    // Check on-time vs late
    const now = new Date();
    const [targetH, targetM] = officeStartTime.split(':').map(Number);
    const targetDate = new Date(now);
    targetDate.setHours(targetH || 0, targetM || 0, 0, 0);

    const diffMs = now.getTime() - targetDate.getTime();
    const diffMinutes = Math.floor(diffMs / (60 * 1000));

    let status = 'ON_TIME';
    let lateMinutes = 0;
    if (diffMinutes > tolerance) {
      status = 'LATE';
      lateMinutes = diffMinutes;
    }

    const todayStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;

    // Prevent double check-in on the same day for office
    if (db.isRecordedForSession(name, 'Kantor Harian', todayStr)) {
      return res.status(400).json({ error: `Peserta "${name}" sudah melakukan absensi kantor hari ini.` });
    }

    const record = db.addRecord({
      name,
      department,
      type: 'OFFLINE',
      sessionName: 'Kantor Harian',
      status,
      lateMinutes
    });

    // Auto sync Excel file in background
    excelService.syncExcelFile().catch(err => console.error(err));

    res.json({
      success: true,
      message: status === 'ON_TIME' 
        ? `Absensi berhasil dicatat! Anda hadir TEPAT WAKTU.` 
        : `Absensi berhasil dicatat! Anda tercatat TERLAMBAT (${lateMinutes} menit).`,
      record
    });
  } catch (err) {
    console.error('Error submitting offline attendance:', err);
    res.status(500).json({ error: 'Gagal mencatat absensi: ' + err.message });
  }
});

// Fetch Records with filters
app.get('/api/attendance/records', (req, res) => {
  const { date, department, type, status, search } = req.query;
  const records = db.getRecords({ date, department, type, status, search });
  res.json({ records, total: records.length });
});

// Fetch Analytics Summary & Charts data
app.get('/api/attendance/analytics', (req, res) => {
  const { date, department } = req.query;
  const analytics = analyticsService.calculateAnalytics({ date, department });
  res.json(analytics);
});

// Download Master Excel (.xlsx)
app.get('/api/attendance/export-excel', async (req, res) => {
  try {
    const buffer = await excelService.getExcelBuffer();
    const dateStr = new Date().toISOString().split('T')[0];
    const filename = `Master_Absensi_${dateStr}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);
  } catch (err) {
    console.error('Error exporting Excel:', err);
    res.status(500).json({ error: 'Gagal mengekspor file Excel: ' + err.message });
  }
});

// QR Code Generator
app.get('/api/qrcode', async (req, res) => {
  try {
    const host = req.get('host') || 'localhost:3000';
    const protocol = req.protocol || 'http';
    const offlineUrl = `${protocol}://${host}/index.html`;

    const qrDataUrl = await QRCode.toDataURL(offlineUrl, {
      width: 320,
      margin: 2,
      color: {
        dark: '#1E3A8A',
        light: '#FFFFFF'
      }
    });

    res.json({ qrDataUrl, targetUrl: offlineUrl });
  } catch (err) {
    res.status(500).json({ error: 'Gagal membuat QR Code' });
  }
});

// ----------------------------------------------------------------------------
// IN-MEETING ONLINE ATTENDANCE ENDPOINTS (Chrome Extension & Quick Sync)
// ----------------------------------------------------------------------------
app.post('/api/bot/batch-sync', async (req, res) => {
  try {
    let { participants, sessionName, scheduledStartTime, toleranceMinutes, platform } = req.body;
    if (!participants) return res.status(400).json({ error: 'Daftar nama peserta diperlukan' });

    let namesArray = [];
    if (Array.isArray(participants)) {
      namesArray = participants;
    } else if (typeof participants === 'string') {
      namesArray = participants.split(/[\r\n,]+/).map(s => s.trim()).filter(Boolean);
    }

    if (namesArray.length === 0) {
      return res.status(400).json({ error: 'Tidak ada nama peserta valid yang ditemukan' });
    }

    const recorded = await syncService.batchSyncParticipants(
      namesArray,
      sessionName || 'Online Meeting',
      scheduledStartTime || '09:00',
      parseInt(toleranceMinutes, 10) || 15,
      platform || 'GOOGLE_MEET'
    );

    res.json({
      success: true,
      message: `Berhasil mencatat ${recorded.length} peserta baru ke sistem absensi dan file Excel!`,
      totalSynced: recorded.length,
      records: recorded
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// ----------------------------------------------------------------------------
// SETTINGS & SEED DATA ENDPOINTS
// ----------------------------------------------------------------------------
app.get('/api/admin/settings', requireAdmin, (req, res) => {
  const current = getConfig();
  res.json({
    officeStartTime: current.officeStartTime,
    toleranceMinutes: current.toleranceMinutes,
    port: current.port
  });
});

app.post('/api/admin/settings', requireAdmin, (req, res) => {
  const { officeStartTime, toleranceMinutes, adminPassword } = req.body;
  const current = getConfig();
  if (officeStartTime) current.officeStartTime = officeStartTime;
  if (toleranceMinutes) current.toleranceMinutes = parseInt(toleranceMinutes, 10);
  if (adminPassword && adminPassword.trim().length >= 4) current.adminPassword = adminPassword.trim();

  saveConfig(current);
  res.json({ success: true, message: 'Pengaturan berhasil disimpan.', settings: current });
});

app.post('/api/attendance/seed-sample', requireAdmin, async (req, res) => {
  try {
    const sampleNames = [
      { name: 'Budi Santoso', dept: 'IT', type: 'ONLINE_ZOOM', status: 'ON_TIME', late: 0, dateOffset: 0 },
      { name: 'Siti Rahmawati', dept: 'Finance', type: 'ONLINE_ZOOM', status: 'ON_TIME', late: 0, dateOffset: 0 },
      { name: 'Ahmad Fauzi', dept: 'IT', type: 'ONLINE_MEET', status: 'LATE', late: 18, dateOffset: 0 },
      { name: 'Dewi Lestari', dept: 'HRD', type: 'OFFLINE', status: 'ON_TIME', late: 0, dateOffset: 0 },
      { name: 'Rian Pratama', dept: 'Marketing', type: 'OFFLINE', status: 'LATE', late: 25, dateOffset: 0 },
      { name: 'Eko Wahyudi', dept: 'Operasional', type: 'OFFLINE', status: 'ON_TIME', late: 0, dateOffset: 0 },
      { name: 'Nadia Putri', dept: 'Desain', type: 'ONLINE_MEET', status: 'ON_TIME', late: 0, dateOffset: 0 },
      { name: 'Hendra Wijaya', dept: 'Sales', type: 'OFFLINE', status: 'LATE', late: 35, dateOffset: 0 },
      { name: 'Budi Santoso', dept: 'IT', type: 'OFFLINE', status: 'ON_TIME', late: 0, dateOffset: -1 },
      { name: 'Siti Rahmawati', dept: 'Finance', type: 'OFFLINE', status: 'ON_TIME', late: 0, dateOffset: -1 },
      { name: 'Ahmad Fauzi', dept: 'IT', type: 'OFFLINE', status: 'ON_TIME', late: 0, dateOffset: -1 },
      { name: 'Dewi Lestari', dept: 'HRD', type: 'ONLINE_ZOOM', status: 'ON_TIME', late: 0, dateOffset: -1 },
      { name: 'Rian Pratama', dept: 'Marketing', type: 'ONLINE_ZOOM', status: 'LATE', late: 20, dateOffset: -1 },
      { name: 'Budi Santoso', dept: 'IT', type: 'ONLINE_ZOOM', status: 'ON_TIME', late: 0, dateOffset: -2 },
      { name: 'Siti Rahmawati', dept: 'Finance', type: 'ONLINE_ZOOM', status: 'LATE', late: 12, dateOffset: -2 }
    ];

    sampleNames.forEach(item => {
      const d = new Date();
      d.setDate(d.getDate() + item.dateOffset);
      const hours = item.status === 'ON_TIME' ? 8 : 8;
      const minutes = item.status === 'ON_TIME' ? Math.floor(Math.random() * 10) : (15 + item.late);
      d.setHours(hours, minutes, 0);

      db.addRecord({
        name: item.name,
        department: item.dept,
        type: item.type,
        sessionName: item.type === 'OFFLINE' ? 'Kantor Harian' : 'Morning Sync',
        timestamp: d.toISOString(),
        status: item.status,
        lateMinutes: item.late
      });
    });

    await excelService.syncExcelFile();
    res.json({ success: true, message: 'Sample data berhasil diinjeksikan dan Excel telah disinkronkan!' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/attendance/clear', requireAdmin, async (req, res) => {
  db.clearAllRecords();
  await excelService.syncExcelFile();
  res.json({ success: true, message: 'Semua data absensi berhasil dibersihkan.' });
});

// Start Server
const PORT = process.env.PORT || config.port || 3000;
app.listen(PORT, async () => {
  console.log(`=======================================================`);
  console.log(`🚀 Sistem Absensi Hybrid Berjalan di http://localhost:${PORT}`);
  console.log(`📱 Form Absensi Offline: http://localhost:${PORT}/index.html`);
  console.log(`📷 QR Code Display:      http://localhost:${PORT}/qr.html`);
  console.log(`📊 Dashboard Admin:       http://localhost:${PORT}/admin.html`);
  console.log(`=======================================================`);

  // Initial sync of excel file
  await excelService.syncExcelFile();
});
