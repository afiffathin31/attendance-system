const http = require('http');

async function testHttp() {
  console.log('=== HTTP ROUTE INTEGRATION TEST ===\n');

  // Start server in background for testing
  const app = require('../server');
  // Wait 1.5s for server to bind
  await new Promise(r => setTimeout(r, 1500));

  const BASE = 'http://localhost:3000';

  // 1. Test QR Code endpoint
  console.log('1. Menguji GET /api/qrcode...');
  const qrRes = await fetch(`${BASE}/api/qrcode`);
  const qrData = await qrRes.json();
  if (!qrData.qrDataUrl) throw new Error('QR API failed');
  console.log('   ✅ OK: Target URL ->', qrData.targetUrl);

  // 2. Test Admin Login
  console.log('2. Menguji POST /api/admin/login...');
  const loginRes = await fetch(`${BASE}/api/admin/login`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ password: 'admin123' })
  });
  const cookie = loginRes.headers.get('set-cookie');
  const loginData = await loginRes.json();
  if (!loginData.success) throw new Error('Login failed');
  console.log('   ✅ OK: Admin logged in, cookie acquired');

  // 3. Test Offline Attendance
  console.log('3. Menguji POST /api/attendance/offline...');
  const attendRes = await fetch(`${BASE}/api/attendance/offline`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name: 'Dimas Anggara', department: 'Operasional' })
  });
  const attendData = await attendRes.json();
  if (!attendData.success) throw new Error('Offline attendance failed: ' + attendData.error);
  console.log(`   ✅ OK: Absensi berhasil dicatat untuk ${attendData.record.name} (${attendData.record.status})`);

  // 4. Test Seed Sample Data
  console.log('4. Menguji POST /api/attendance/seed-sample...');
  const seedRes = await fetch(`${BASE}/api/attendance/seed-sample`, {
    method: 'POST',
    headers: { 
      'Content-Type': 'application/json',
      'Cookie': cookie
    }
  });
  const seedData = await seedRes.json();
  if (!seedData.success) throw new Error('Seed sample failed');
  console.log('   ✅ OK: Sample data injected');

  // 5. Test Analytics API
  console.log('5. Menguji GET /api/attendance/analytics...');
  const analyticsRes = await fetch(`${BASE}/api/attendance/analytics`);
  const analyticsData = await analyticsRes.json();
  console.log(`   ✅ OK: Total ${analyticsData.summary.total} data, On-time rate: ${analyticsData.summary.onTimeRate}%, Departemen: ${analyticsData.departments.length}`);

  // 6. Test Excel Export Buffer
  console.log('6. Menguji GET /api/attendance/export-excel...');
  const excelRes = await fetch(`${BASE}/api/attendance/export-excel`);
  const excelBuffer = await excelRes.arrayBuffer();
  if (excelBuffer.byteLength < 5000) throw new Error('Excel export is too small');
  console.log(`   ✅ OK: File Excel valid diunduh (${excelBuffer.byteLength} bytes)`);

  console.log('\n🎉 SEMUA ENDPOINT HTTP DAN FLOW ABSENSI BERHASIL DIUJI! 🎉');
  process.exit(0);
}

testHttp().catch(err => {
  console.error('❌ HTTP TEST ERROR:', err);
  process.exit(1);
});
