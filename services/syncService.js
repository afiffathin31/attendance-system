const db = require('../database/db');
const excelService = require('./excelService');

function isKnownDepartment(str) {
  if (!str) return false;
  const s = str.trim().toLowerCase();
  const known = [
    'it', 'hrd', 'hr', 'finance', 'keuangan', 'marketing', 'pemasaran',
    'operasional', 'ops', 'sales', 'desain', 'design', 'legal', 'admin',
    'umum', 'accounting', 'engineering', 'qa', 'support', 'business'
  ];
  return known.includes(s);
}

function isJunkText(text) {
  if (!text || text.length < 2 || text.length > 50) return true;
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

  // Real human names rarely exceed 5 words
  if (text.split(/\s+/).length > 5) return true;

  return false;
}

/**
 * Parse participant display name e.g. "[IT] Budi Santoso" or "Finance - Siti Rahma"
 */
function parseParticipantName(rawName) {
  if (!rawName) return null;
  let clean = rawName.trim();
  if (isJunkText(clean)) return null;

  // Remove common bot/system suffixes
  clean = clean.replace(/\((You|Host|Co-host|Me|Anda|Penyelenggara|Penyelenggara rapat)\)/gi, '').trim();

  // Pattern 1: [Departemen] Nama Lengkap
  const bracketMatch = clean.match(/^\[([^\]]+)\]\s*(.+)$/);
  if (bracketMatch) {
    return {
      department: bracketMatch[1].trim(),
      name: bracketMatch[2].trim()
    };
  }

  // Pattern 2: Dash Separator (e.g. Finance - Siti Rahma OR Ahmad Fauzi - HRD)
  if (clean.includes('-') || clean.includes('–') || clean.includes('—')) {
    const parts = clean.split(/[-–—]/);
    if (parts.length >= 2) {
      const p1 = parts[0].trim();
      const p2 = parts.slice(1).join('-').trim();

      if (isKnownDepartment(p1)) {
        return { department: p1, name: p2 };
      } else if (isKnownDepartment(p2)) {
        return { name: p1, department: p2 };
      } else {
        if (p1.split(/\s+/).length === 1 && p2.split(/\s+/).length > 1) {
          return { department: p1, name: p2 };
        }
        return { name: p1, department: p2 };
      }
    }
  }

  // Pattern 3: Nama Lengkap (Departemen)
  const parenMatch = clean.match(/^(.+?)\s*\(([^)]+)\)$/);
  if (parenMatch) {
    return {
      name: parenMatch[1].trim(),
      department: parenMatch[2].trim()
    };
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

    if (db.isRecordedForSession(parsed.name, sessionName || 'Online Meeting', today)) continue;

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
