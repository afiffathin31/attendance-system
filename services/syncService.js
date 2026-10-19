const db = require('../database/db');
const excelService = require('./excelService');

const DEPARTMENTS_MAP = {
  // 14 Official Departments & Aliases
  'ebs': 'EBS',
  'solution architect': 'Solution Architect',
  'solution architecture': 'Solution Architect',
  'architect': 'Solution Architect',
  'sa': 'Solution Architect',
  'could automation': 'Could Automation',
  'cloud automation': 'Could Automation',
  'ca': 'Could Automation',
  'research and innovation': 'Research and innovation',
  'research & innovation': 'Research and innovation',
  'rni': 'Research and innovation',
  'r&i': 'Research and innovation',
  'research': 'Research and innovation',
  'infrastructur analytic': 'Infrastructur Analytic',
  'infrastructure analytic': 'Infrastructur Analytic',
  'infrastructure analytics': 'Infrastructur Analytic',
  'infrastruktur analitik': 'Infrastructur Analytic',
  'ia': 'Infrastructur Analytic',
  'fat': 'FAT',
  'finance, accounting & tax': 'FAT',
  'finance': 'FAT',
  'keuangan': 'FAT',
  'accounting': 'FAT',
  'tax': 'FAT',
  'pajak': 'FAT',
  'sales': 'Sales',
  'marketing': 'Sales',
  'pemasaran': 'Sales',
  'bd': 'Sales',
  'business development': 'Sales',
  'cyber cecurity': 'Cyber Cecurity',
  'cyber security': 'Cyber Cecurity',
  'security': 'Cyber Cecurity',
  'keamanan siber': 'Cyber Cecurity',
  'cs': 'Cyber Cecurity',
  'pmo': 'PMO',
  'project management': 'PMO',
  'bod': 'BOD',
  'board of directors': 'BOD',
  'direksi': 'BOD',
  'modern application': 'Modern Application',
  'modern applications': 'Modern Application',
  'modern app': 'Modern Application',
  'modern apps': 'Modern Application',
  'ma': 'Modern Application',
  'data analytics': 'Data Analytics',
  'data analytic': 'Data Analytics',
  'data analyst': 'Data Analytics',
  'data science': 'Data Analytics',
  'da': 'Data Analytics',
  'hcmc': 'HCMC',
  'hrd': 'HCMC',
  'hr': 'HCMC',
  'human capital': 'HCMC',
  'personalia': 'HCMC',
  'pkl/magang': 'PKL/Magang',
  'pkl': 'PKL/Magang',
  'magang': 'PKL/Magang',
  'intern': 'PKL/Magang',
  'internship': 'PKL/Magang',
  // Common corporate & legacy
  'it': 'IT',
  'ti': 'IT',
  'legal': 'Legal',
  'operasional': 'Operasional',
  'ops': 'Operasional',
  'admin': 'Admin',
  'umum': 'Umum'
};

function normalizeDepartment(str) {
  if (!str) return null;
  const rawKey = str.trim().toLowerCase();
  if (DEPARTMENTS_MAP[rawKey]) return DEPARTMENTS_MAP[rawKey];
  const cleanedKey = rawKey.replace(/[\s\-_]+/g, ' ');
  if (DEPARTMENTS_MAP[cleanedKey]) return DEPARTMENTS_MAP[cleanedKey];
  return null;
}

function isKnownDepartment(str) {
  return normalizeDepartment(str) !== null;
}

function isJunkText(text) {
  if (!text || text.length < 2 || text.length > 70) return true;
  if (text.includes('\n')) return true;

  const lower = text.toLowerCase();
  const junkSubstrings = [
    'more_vert', 'visual_effects', 'frame_person', 'lock_person', 'call_end',
    'back_hand', 'closed_caption', 'mood', 'computer_arrow', 'videocam',
    'mic', 'push_pin', 'volume_up', 'opsi lainnya', 'latar belakang',
    'bingkai ulang', 'orang lain', 'tambahkan orang', 'telusuri orang',
    'dalam rapat ini', 'kontributor', 'penyelenggara', 'turn on', 'turn off',
    'mute', 'audio', 'video'
  ];

  for (const j of junkSubstrings) {
    if (lower.includes(j)) return true;
  }

  // Reject snake_case words without spaces (icon identifiers)
  if (text.includes('_') && !text.includes(' ')) return true;

  // Real human names with department rarely exceed 8 words
  if (text.split(/\s+/).length > 8) return true;

  return false;
}

/**
 * Parse participant display name e.g. "[IT] Budi Santoso", "FAT - Siti Rahma", "Ahmad Fauzi - HRD"
 */
function parseParticipantName(rawName) {
  if (!rawName) return null;
  let clean = String(rawName).trim();
  if (isJunkText(clean)) return null;

  // 1. Ignore headers like "Participants (5)", "Peserta (10)"
  if (/^(participants|peserta)\s*\(/i.test(clean)) return null;

  // 2. Remove common bot/system suffixes and meeting badges
  clean = clean.replace(/\((You|Host|Co-host|Host,\s*me|Me|Anda|Penyelenggara|Penyelenggara rapat|Guest|Tamu)\)/gi, '').trim();

  // 3. Remove trailing action buttons
  clean = clean.replace(/\b(invite|mute all|more|mute)\b/gi, '').trim();

  if (!clean || clean.length < 2) return null;

  // Pattern 1: [Departemen] Nama Lengkap OR {Departemen} Nama Lengkap
  const bracketPrefix = clean.match(/^[\[\{]([^\]\}]+)[\]\}]\s*[-–—:]?\s*(.+)$/);
  if (bracketPrefix) {
    const rawDept = bracketPrefix[1].trim();
    return {
      department: normalizeDepartment(rawDept) || rawDept,
      name: bracketPrefix[2].trim()
    };
  }

  // Pattern 2: Nama Lengkap [Departemen]
  const bracketSuffix = clean.match(/^(.+?)\s*[-–—:]?\s*[\[\{]([^\]\}]+)[\]\}]$/);
  if (bracketSuffix) {
    const rawDept = bracketSuffix[2].trim();
    return {
      name: bracketSuffix[1].trim(),
      department: normalizeDepartment(rawDept) || rawDept
    };
  }

  // Pattern 3: (Departemen) Nama Lengkap
  const parenPrefix = clean.match(/^\(([^)]+)\)\s*[-–—:]?\s*(.+)$/);
  if (parenPrefix) {
    const rawDept = parenPrefix[1].trim();
    return {
      department: normalizeDepartment(rawDept) || rawDept,
      name: parenPrefix[2].trim()
    };
  }

  // Pattern 4: Nama Lengkap (Departemen)
  const parenMatch = clean.match(/^(.+?)\s*\(([^)]+)\)$/);
  if (parenMatch) {
    const p1 = parenMatch[1].trim();
    const p2 = parenMatch[2].trim();
    if (isKnownDepartment(p2)) {
      return { name: p1, department: normalizeDepartment(p2) };
    } else if (isKnownDepartment(p1)) {
      return { department: normalizeDepartment(p1), name: p2 };
    } else {
      return { name: p1, department: p2 };
    }
  }

  // Pattern 5: Separator patterns: -, –, —, |, /, :, _
  if (/[-–—|:_/]/.test(clean)) {
    const parts = clean.split(/[-–—|:_]/).map(s => s.trim()).filter(Boolean);
    if (parts.length >= 2) {
      const p1 = parts[0];
      const p2 = parts.slice(1).join(' ').trim();

      const p1Dept = normalizeDepartment(p1);
      const p2Dept = normalizeDepartment(p2);

      if (p1Dept && !p2Dept) {
        return { department: p1Dept, name: p2 };
      }
      if (p2Dept && !p1Dept) {
        return { name: p1, department: p2Dept };
      }

      // If neither or both in dictionary, use acronym & word heuristics
      const p1IsAcronym = /^[A-Za-z0-9&]{2,5}$/.test(p1) && (p1 === p1.toUpperCase() || p1.length <= 4);
      const p2IsAcronym = /^[A-Za-z0-9&]{2,5}$/.test(p2) && (p2 === p2.toUpperCase() || p2.length <= 4);

      if (p1IsAcronym && !p2IsAcronym) {
        return { department: p1, name: p2 };
      }
      if (p2IsAcronym && !p1IsAcronym) {
        return { name: p1, department: p2 };
      }

      // Default corporate convention is [Departemen] - [Nama Lengkap]
      return { department: p1, name: p2 };
    }
  }

  // Default fallback if no department format
  return {
    name: clean,
    department: 'Umum'
  };
}

/**
 * Evaluate on-time vs late
 */
function evaluateStatus(scheduledTimeStr, toleranceMinutes) {
  const now = new Date();
  const [targetH, targetM] = (scheduledTimeStr || '09:00').split(':').map(Number);
  
  const targetDate = new Date(now);
  targetDate.setHours(targetH || 0, targetM || 0, 0, 0);

  const diffMs = now.getTime() - targetDate.getTime();
  const diffMinutes = Math.floor(diffMs / (60 * 1000));

  if (diffMinutes <= (toleranceMinutes || 15)) {
    return { status: 'ON_TIME', lateMinutes: 0 };
  } else {
    return { status: 'LATE', lateMinutes: diffMinutes };
  }
}

/**
 * Synchronize batch of participants (from Chrome Extension)
 */
async function batchSyncParticipants(namesList, sessionName, scheduledStartTime, toleranceMinutes, platform = 'GOOGLE_MEET') {
  const recorded = [];
  const today = new Date().toISOString().split('T')[0];

  for (const rawName of namesList) {
    const parsed = parseParticipantName(rawName);
    if (!parsed || !parsed.name) continue;

    if (db.isRecordedForSession(parsed.name, sessionName || 'Online Meeting', today, parsed.department)) continue;

    const evalResult = evaluateStatus(scheduledStartTime || '09:00', toleranceMinutes || 15);
    const rec = db.addRecord({
      name: parsed.name,
      department: parsed.department,
      type: platform === 'ZOOM' ? 'ONLINE_ZOOM' : 'ONLINE_MEET',
      sessionName: sessionName || 'Online Meeting',
      status: evalResult.status,
      lateMinutes: evalResult.lateMinutes
    });
    recorded.push(rec);
  }

  if (recorded.length > 0) {
    await excelService.syncExcelFile();
  }
  return recorded;
}

module.exports = {
  parseParticipantName,
  batchSyncParticipants,
  isJunkText
};
