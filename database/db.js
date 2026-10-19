const fs = require('fs');
const path = require('path');

const DATA_DIR = path.join(__dirname, '..', 'data');
const DB_FILE = path.join(DATA_DIR, 'attendance.json');

// Ensure directory exists
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

function loadDb() {
  if (!fs.existsSync(DB_FILE)) {
    const initialData = {
      records: [],
      meetings: [],
      departments: ['IT', 'HRD', 'Finance', 'Marketing', 'Operasional', 'Desain', 'Sales']
    };
    fs.writeFileSync(DB_FILE, JSON.stringify(initialData, null, 2), 'utf-8');
    return initialData;
  }
  try {
    const content = fs.readFileSync(DB_FILE, 'utf-8');
    return JSON.parse(content);
  } catch (err) {
    console.error('Error reading db file, resetting:', err);
    return { records: [], meetings: [], departments: [] };
  }
}

function saveDb(data) {
  try {
    fs.writeFileSync(DB_FILE, JSON.stringify(data, null, 2), 'utf-8');
  } catch (err) {
    console.error('Error saving db file:', err);
  }
}

function getRecords(filters = {}) {
  const db = loadDb();
  let records = db.records || [];

  if (filters.date) {
    records = records.filter(r => r.date === filters.date);
  }
  if (filters.department && filters.department !== 'ALL') {
    records = records.filter(r => r.department.toLowerCase() === filters.department.toLowerCase());
  }
  if (filters.type && filters.type !== 'ALL') {
    records = records.filter(r => r.type === filters.type);
  }
  if (filters.status && filters.status !== 'ALL') {
    records = records.filter(r => r.status === filters.status);
  }
  if (filters.search) {
    const q = filters.search.toLowerCase();
    records = records.filter(r => r.name.toLowerCase().includes(q) || r.department.toLowerCase().includes(q));
  }

  // Sort newest first
  return records.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
}

function addRecord(record) {
  const db = loadDb();
  if (!db.records) db.records = [];

  const now = new Date(record.timestamp || Date.now());
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const dateStr = `${year}-${month}-${day}`;
  const timeStr = now.toTimeString().split(' ')[0]; // HH:mm:ss

  const newRecord = {
    id: `ATT-${Date.now()}-${Math.floor(Math.random() * 1000)}`,
    name: record.name.trim(),
    department: record.department.trim(),
    type: record.type || 'OFFLINE', // OFFLINE, ONLINE_ZOOM, ONLINE_MEET
    sessionName: record.sessionName || 'Kantor Harian',
    timestamp: now.toISOString(),
    date: dateStr,
    time: timeStr,
    status: record.status || 'ON_TIME', // ON_TIME, LATE
    lateMinutes: record.lateMinutes || 0
  };

  db.records.push(newRecord);

  // Add department if not existing
  if (newRecord.department && !db.departments.includes(newRecord.department)) {
    db.departments.push(newRecord.department);
  }

  saveDb(db);
  return newRecord;
}

function isRecordedForSession(name, sessionName, dateStr, department) {
  const db = loadDb();
  const targetName = name.trim().toLowerCase();
  return (db.records || []).some(r => {
    const matchName = r.name.toLowerCase() === targetName;
    const matchSession = !sessionName || r.sessionName === sessionName;
    const matchDate = !dateStr || r.date === dateStr;
    if (department && r.department) {
      return matchName && matchSession && matchDate && r.department.toLowerCase() === department.trim().toLowerCase();
    }
    return matchName && matchSession && matchDate;
  });
}

function getDepartments() {
  const db = loadDb();
  return db.departments || ['IT', 'HRD', 'Finance', 'Marketing', 'Operasional', 'Desain', 'Sales'];
}

function clearAllRecords() {
  const db = loadDb();
  db.records = [];
  saveDb(db);
}

module.exports = {
  loadDb,
  saveDb,
  getRecords,
  addRecord,
  isRecordedForSession,
  getDepartments,
  clearAllRecords
};
