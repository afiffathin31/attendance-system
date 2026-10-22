const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const db = require('../database/db');
const analyticsService = require('./analyticsService');

const EXCEL_PATH = path.join(__dirname, '..', 'data', 'master_absensi.xlsx');

async function buildWorkbook() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Sistem Absensi Otomatis';
  workbook.lastModifiedBy = 'Admin Absensi';
  workbook.created = new Date();
  workbook.modified = new Date();

  const records = db.getRecords();
  const analytics = analyticsService.calculateAnalytics();

  // --------------------------------------------------------------------------
  // SHEET 1: Log Absensi Lengkap
  // --------------------------------------------------------------------------
  const sheetLog = workbook.addWorksheet('Log Absensi', {
    views: [{ showGridLines: true }]
  });

  sheetLog.columns = [
    { header: 'No', key: 'no', width: 6 },
    { header: 'ID Absensi', key: 'id', width: 22 },
    { header: 'Tanggal', key: 'date', width: 14 },
    { header: 'Jam Masuk', key: 'time', width: 12 },
    { header: 'Nama Lengkap', key: 'name', width: 25 },
    { header: 'Departemen', key: 'department', width: 18 },
    { header: 'Metode', key: 'type', width: 16 },
    { header: 'Sesi / Rapat', key: 'sessionName', width: 22 },
    { header: 'Status', key: 'status', width: 16 },
    { header: 'Keterlambatan (menit)', key: 'lateMinutes', width: 22 }
  ];

  // Header Styling
  sheetLog.getRow(1).eachCell(cell => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '1E3A8A' } // Navy blue
    };
    cell.font = {
      color: { argb: 'FFFFFF' },
      bold: true,
      size: 11
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });
  sheetLog.getRow(1).height = 28;

  // Add Data Rows
  records.forEach((r, idx) => {
    const isLate = r.status === 'LATE';
    const row = sheetLog.addRow({
      no: idx + 1,
      id: r.id,
      date: r.date,
      time: r.time,
      name: r.name,
      department: r.department,
      type: r.type === 'ONLINE_ZOOM' ? 'Zoom Meeting' : (r.type === 'ONLINE_MEET' ? 'Google Meet' : 'Offline Kantor'),
      sessionName: r.sessionName,
      status: isLate ? 'TERLAMBAT' : 'TEPAT WAKTU',
      lateMinutes: isLate ? `${r.lateMinutes} menit` : '-'
    });

    row.height = 22;
    row.alignment = { vertical: 'middle' };

    // Format Status Cell
    const statusCell = row.getCell('status');
    statusCell.alignment = { vertical: 'middle', horizontal: 'center' };
    statusCell.font = { bold: true, color: { argb: isLate ? 'B91C1C' : '047857' } };
    statusCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: isLate ? 'FEE2E2' : 'D1FAE5' }
    };

    // Center alignment for specific columns
    row.getCell('no').alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell('date').alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell('time').alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell('type').alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell('lateMinutes').alignment = { vertical: 'middle', horizontal: 'center' };
  });

  // --------------------------------------------------------------------------
  // SHEET 2: Rekap Performa Individu
  // --------------------------------------------------------------------------
  const sheetInd = workbook.addWorksheet('Performa Individu', {
    views: [{ showGridLines: true }]
  });

  sheetInd.columns = [
    { header: 'No', key: 'no', width: 6 },
    { header: 'Nama Karyawan', key: 'name', width: 25 },
    { header: 'Departemen', key: 'department', width: 18 },
    { header: 'Total Hadir', key: 'total', width: 14 },
    { header: 'Tepat Waktu', key: 'onTime', width: 14 },
    { header: 'Terlambat', key: 'late', width: 14 },
    { header: 'Ketepatan Waktu (%)', key: 'onTimeRate', width: 20 },
    { header: 'Rata-rata Telat (menit)', key: 'avgLate', width: 22 }
  ];

  sheetInd.getRow(1).eachCell(cell => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '065F46' } // Forest Green
    };
    cell.font = { color: { argb: 'FFFFFF' }, bold: true, size: 11 };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });
  sheetInd.getRow(1).height = 28;

  analytics.individuals.forEach((ind, idx) => {
    const row = sheetInd.addRow({
      no: idx + 1,
      name: ind.name,
      department: ind.department,
      total: ind.total,
      onTime: ind.onTime,
      late: ind.late,
      onTimeRate: `${ind.onTimeRate}%`,
      avgLate: ind.avgLateMinutes > 0 ? `${ind.avgLateMinutes} menit` : '-'
    });

    row.height = 22;
    row.alignment = { vertical: 'middle' };
    row.getCell('no').alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell('total').alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell('onTime').alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell('late').alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell('onTimeRate').alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell('avgLate').alignment = { vertical: 'middle', horizontal: 'center' };

    const rateCell = row.getCell('onTimeRate');
    if (ind.onTimeRate >= 80) {
      rateCell.font = { bold: true, color: { argb: '047857' } };
    } else {
      rateCell.font = { bold: true, color: { argb: 'B91C1C' } };
    }
  });

  // --------------------------------------------------------------------------
  // SHEET 3: Rekap Departemen
  // --------------------------------------------------------------------------
  const sheetDept = workbook.addWorksheet('Rekap Departemen', {
    views: [{ showGridLines: true }]
  });

  sheetDept.columns = [
    { header: 'No', key: 'no', width: 6 },
    { header: 'Departemen', key: 'department', width: 22 },
    { header: 'Total Kehadiran', key: 'total', width: 16 },
    { header: 'Tepat Waktu', key: 'onTime', width: 14 },
    { header: 'Terlambat', key: 'late', width: 14 },
    { header: 'Ketepatan Waktu (%)', key: 'onTimeRate', width: 22 }
  ];

  sheetDept.getRow(1).eachCell(cell => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '4338CA' } // Indigo
    };
    cell.font = { color: { argb: 'FFFFFF' }, bold: true, size: 11 };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });
  sheetDept.getRow(1).height = 28;

  analytics.departments.forEach((dept, idx) => {
    const row = sheetDept.addRow({
      no: idx + 1,
      department: dept.department,
      total: dept.total,
      onTime: dept.onTime,
      late: dept.late,
      onTimeRate: `${dept.onTimeRate}%`
    });

    row.height = 22;
    row.alignment = { vertical: 'middle' };
    row.getCell('no').alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell('total').alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell('onTime').alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell('late').alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell('onTimeRate').alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell('onTimeRate').font = { bold: true };
  });

  // --------------------------------------------------------------------------
  // SHEET 4: Input Data Karyawan (Dropdown 14 Departemen)
  // --------------------------------------------------------------------------
  const sheetKaryawan = workbook.addWorksheet('Data Karyawan', {
    views: [{ showGridLines: true }]
  });

  sheetKaryawan.columns = [
    { header: 'No', key: 'no', width: 6 },
    { header: 'Nama Lengkap Karyawan', key: 'name', width: 28 },
    { header: 'Departemen / Divisi', key: 'department', width: 25 },
    { header: 'Total Hadir', key: 'total', width: 14 },
    { header: 'Tepat Waktu', key: 'onTime', width: 14 },
    { header: 'Terlambat', key: 'late', width: 14 },
    { header: '% On-Time', key: 'onTimeRate', width: 16 },
    { header: 'Status Performa', key: 'status', width: 22 }
  ];

  // Header Styling
  sheetKaryawan.getRow(1).eachCell(cell => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '0F172A' } // Dark Slate
    };
    cell.font = { color: { argb: 'FFFFFF' }, bold: true, size: 11 };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });
  sheetKaryawan.getRow(1).height = 28;

  // Kosongkan baris data karyawan agar user dapat menginputkan nama dan divisi sendiri secara mandiri
  // Data Validation Dropdown for Column C (Departemen) - 14 Departemen Resmi
  const deptListString = '"' + [
    'EBS', 'Solution Architect', 'Could Automation', 'Research and innovation',
    'Infrastructur Analytic', 'FAT', 'Sales', 'Cyber Cecurity', 'PMO', 'BOD',
    'Modern Application', 'Data Analytics', 'HCMC', 'PKL/Magang'
  ].join(',') + '"';

  sheetKaryawan.dataValidations.add('C2:C500', {
    type: 'list',
    allowBlank: true,
    formulae: [deptListString]
  });

  return workbook;
}

async function syncExcelFile() {
  try {
    const workbook = await buildWorkbook();
    const dataDir = path.dirname(EXCEL_PATH);
    if (!fs.existsSync(dataDir)) {
      fs.mkdirSync(dataDir, { recursive: true });
    }
    await workbook.xlsx.writeFile(EXCEL_PATH);
    console.log(`[ExcelService] Master Excel file synchronized at: ${EXCEL_PATH}`);
    return true;
  } catch (err) {
    console.error('[ExcelService] Error writing excel file:', err);
    return false;
  }
}

async function getExcelBuffer() {
  const workbook = await buildWorkbook();
  return await workbook.xlsx.writeBuffer();
}

module.exports = {
  syncExcelFile,
  getExcelBuffer,
  EXCEL_PATH
};
