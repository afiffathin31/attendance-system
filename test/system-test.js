const path = require('path');
const fs = require('fs');

async function runTests() {
  console.log('=== SYSTEM TEST SUITE: ATTENDANCE HYBRID SYSTEM ===\n');

  // 1. Test Participant Name Parser
  console.log('Test 1: Memverifikasi Parsing Display Name Meeting Online...');
  const syncService = require('../services/syncService');
  
  const testCases = [
    { input: '[IT] Budi Santoso', expectedDept: 'IT', expectedName: 'Budi Santoso' },
    { input: 'Finance - Siti Rahma', expectedDept: 'FAT', expectedName: 'Siti Rahma' },
    { input: 'FAT - Budi Santoso', expectedDept: 'FAT', expectedName: 'Budi Santoso' },
    { input: 'FAT – Ahmad Fauzi', expectedDept: 'FAT', expectedName: 'Ahmad Fauzi' },
    { input: 'FAT - Hendra Wijaya', expectedDept: 'FAT', expectedName: 'Hendra Wijaya' },
    { input: 'Solution Architect - Hendra Pratama', expectedDept: 'Solution Architect', expectedName: 'Hendra Pratama' },
    { input: 'Solution Architect - Muhammad Rizky Pratama', expectedDept: 'Solution Architect', expectedName: 'Muhammad Rizky Pratama' },
    { input: 'Ahmad Fauzi - HRD', expectedDept: 'HCMC', expectedName: 'Ahmad Fauzi' },
    { input: 'Rian Pratama (Marketing)', expectedDept: 'Sales', expectedName: 'Rian Pratama' },
    { input: '[EBS] Siti Rahmah', expectedDept: 'EBS', expectedName: 'Siti Rahmah' },
    { input: 'Dewi Lestari (Host)', expectedDept: 'Umum', expectedName: 'Dewi Lestari' }
  ];

  for (const tc of testCases) {
    const res = syncService.parseParticipantName(tc.input);
    console.log(`  Input: "${tc.input}" -> Dept: "${res.department}", Name: "${res.name}"`);
    if (res.name !== tc.expectedName || res.department.toLowerCase() !== tc.expectedDept.toLowerCase()) {
      throw new Error(`Gagal parsing: ${tc.input} (didapat: Dept=${res.department}, Name=${res.name})`);
    }
  }
  console.log('  ✅ Test 1 Lulus: Semua variasi nama berhasil diparsing dengan akurat.\n');

  // 1b. Test Same-Department Multi-Participants Batch Sync
  console.log('Test 1b: Memverifikasi Batch Sync Banyak Peserta dari Departemen yang Sama...');
  const db = require('../database/db');
  db.clearAllRecords();

  const sameDeptBatch = [
    'FAT - Budi Santoso',
    'FAT - Siti Rahma',
    'FAT – Ahmad Fauzi',
    'Solution Architect - Hendra Pratama',
    'Solution Architect - Muhammad Rizky Pratama'
  ];

  const firstSync = await syncService.batchSyncParticipants(sameDeptBatch, 'Kajian Rutin', '08:00', 15);
  console.log(`  Sync 1: Berhasil mencatat ${firstSync.length} dari ${sameDeptBatch.length} peserta.`);
  if (firstSync.length !== 5) {
    throw new Error(`Data departemen sama ter-skip! Harusnya 5, tetapi hanya ${firstSync.length} yang tercatat.`);
  }

  // Verifikasi apakah auto-sync 5 menit berikutnya menduplikasi atau tidak
  const secondSync = await syncService.batchSyncParticipants(sameDeptBatch, 'Kajian Rutin', '08:00', 15);
  console.log(`  Sync 2 (Auto-sync 5 menit): Tercatat ${secondSync.length} peserta baru (semua harus di-skip karena sudah terdata).`);
  if (secondSync.length !== 0) {
    throw new Error(`Deduplikasi gagal! Terjadi duplikasi sebanyak ${secondSync.length}`);
  }
  console.log('  ✅ Test 1b Lulus: Peserta dari departemen yang sama TIDAK ter-skip dan deduplikasi bekerja tepat.\n');

  // 2. Test Database & Record Addition
  console.log('Test 2: Menguji Database Persistence...');
  db.clearAllRecords();

  const rec1 = db.addRecord({
    name: 'Budi Santoso',
    department: 'IT',
    type: 'ONLINE_ZOOM',
    sessionName: 'Pagi Hari',
    status: 'ON_TIME',
    lateMinutes: 0
  });

  const rec2 = db.addRecord({
    name: 'Siti Rahma',
    department: 'Finance',
    type: 'OFFLINE',
    sessionName: 'Kantor Harian',
    status: 'LATE',
    lateMinutes: 20
  });

  const records = db.getRecords();
  console.log(`  Tercatat ${records.length} data di database.`);
  if (records.length !== 2) throw new Error('Data tidak tersimpan dengan benar di DB');
  console.log('  ✅ Test 2 Lulus: Database berfungsi dengan baik.\n');

  // 3. Test Analytics Service
  console.log('Test 3: Menguji Kalkulasi Analitik & Grafik...');
  const analyticsService = require('../services/analyticsService');
  const analytics = analyticsService.calculateAnalytics();
  
  console.log('  Ringkasan:', JSON.stringify(analytics.summary));
  console.log('  Departemen:', analytics.departments.map(d => `${d.department}: ${d.onTimeRate}% on-time`).join(', '));
  console.log('  Individu:', analytics.individuals.map(i => `${i.name}: ${i.onTimeRate}%`).join(', '));
  
  if (analytics.summary.total !== 2 || analytics.summary.totalOnTime !== 1 || analytics.summary.totalLate !== 1) {
    throw new Error('Kalkulasi analitik tidak sesuai!');
  }
  console.log('  ✅ Test 3 Lulus: Kalkulasi analitik presisi.\n');

  // 4. Test Excel Generation
  console.log('Test 4: Menguji Pembuatan File Master Excel (.xlsx)...');
  const excelService = require('../services/excelService');
  const syncSuccess = await excelService.syncExcelFile();
  if (!syncSuccess) throw new Error('Sinkronisasi Excel gagal!');

  const excelPath = excelService.EXCEL_PATH;
  if (!fs.existsSync(excelPath)) {
    throw new Error(`File Excel tidak ditemukan di: ${excelPath}`);
  }
  const stats = fs.statSync(excelPath);
  console.log(`  File Excel master terbentuk: ${excelPath} (${stats.size} bytes)`);
  if (stats.size < 1000) throw new Error('Ukuran file Excel terlalu kecil');
  console.log('  ✅ Test 4 Lulus: Master file Excel (.xlsx) berhasil dibuat dan disinkronkan.\n');

  // 5. Test QR Code Generation
  console.log('Test 5: Menguji Generator QR Code...');
  const QRCode = require('qrcode');
  const qrData = await QRCode.toDataURL('http://localhost:3000/index.html');
  if (!qrData.startsWith('data:image/png;base64,')) {
    throw new Error('Format QR Code tidak valid!');
  }
  console.log('  ✅ Test 5 Lulus: QR Code berhasil di-generate.\n');

  console.log('🎉 SEMUA PENGUJIAN SISTEM BERHASIL DENGAN SEMPURNA! 🎉');
}

runTests().catch(err => {
  console.error('❌ PENGUJIAN GAGAL:', err);
  process.exit(1);
});
