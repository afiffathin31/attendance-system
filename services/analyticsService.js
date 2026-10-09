const db = require('../database/db');

function calculateAnalytics(filters = {}) {
  const records = db.getRecords(filters);

  const total = records.length;
  let totalOnTime = 0;
  let totalLate = 0;
  let sumLateMinutes = 0;
  let totalOnline = 0;
  let totalOffline = 0;

  const departmentMap = {};
  const individualMap = {};
  const dateMap = {};

  records.forEach(r => {
    // Basic counters
    if (r.status === 'ON_TIME') {
      totalOnTime++;
    } else {
      totalLate++;
      sumLateMinutes += (r.lateMinutes || 0);
    }

    if (r.type && r.type.startsWith('ONLINE')) {
      totalOnline++;
    } else {
      totalOffline++;
    }

    // Department aggregation
    const dept = r.department || 'Lainnya';
    if (!departmentMap[dept]) {
      departmentMap[dept] = { department: dept, total: 0, onTime: 0, late: 0 };
    }
    departmentMap[dept].total++;
    if (r.status === 'ON_TIME') departmentMap[dept].onTime++;
    else departmentMap[dept].late++;

    // Individual aggregation
    const key = r.name.toLowerCase();
    if (!individualMap[key]) {
      individualMap[key] = {
        name: r.name,
        department: r.department,
        total: 0,
        onTime: 0,
        late: 0,
        sumLateMinutes: 0,
        lastSeen: r.timestamp
      };
    }
    individualMap[key].total++;
    if (r.status === 'ON_TIME') {
      individualMap[key].onTime++;
    } else {
      individualMap[key].late++;
      individualMap[key].sumLateMinutes += (r.lateMinutes || 0);
    }
    if (new Date(r.timestamp) > new Date(individualMap[key].lastSeen)) {
      individualMap[key].lastSeen = r.timestamp;
      individualMap[key].department = r.department; // update to latest department
    }

    // Date trend aggregation
    const d = r.date;
    if (!dateMap[d]) {
      dateMap[d] = { date: d, onTime: 0, late: 0, total: 0 };
    }
    dateMap[d].total++;
    if (r.status === 'ON_TIME') dateMap[d].onTime++;
    else dateMap[d].late++;
  });

  const onTimeRate = total > 0 ? Math.round((totalOnTime / total) * 1000) / 10 : 0;
  const lateRate = total > 0 ? Math.round((totalLate / total) * 1000) / 10 : 0;
  const avgLateMinutes = totalLate > 0 ? Math.round((sumLateMinutes / totalLate) * 10) / 10 : 0;

  // Process department stats
  const departments = Object.values(departmentMap).map(d => ({
    ...d,
    onTimeRate: d.total > 0 ? Math.round((d.onTime / d.total) * 1000) / 10 : 0,
    lateRate: d.total > 0 ? Math.round((d.late / d.total) * 1000) / 10 : 0
  })).sort((a, b) => b.total - a.total);

  // Process individual leaderboard
  const individuals = Object.values(individualMap).map(ind => ({
    name: ind.name,
    department: ind.department,
    total: ind.total,
    onTime: ind.onTime,
    late: ind.late,
    onTimeRate: ind.total > 0 ? Math.round((ind.onTime / ind.total) * 1000) / 10 : 0,
    avgLateMinutes: ind.late > 0 ? Math.round((ind.sumLateMinutes / ind.late) * 10) / 10 : 0,
    lastSeen: ind.lastSeen
  })).sort((a, b) => {
    // Sort primarily by onTimeRate desc, then total desc
    if (b.onTimeRate !== a.onTimeRate) return b.onTimeRate - a.onTimeRate;
    return b.total - a.total;
  });

  // Process date trends (sorted chronologically)
  const trends = Object.values(dateMap).sort((a, b) => a.date.localeCompare(b.date));

  return {
    summary: {
      total,
      totalOnTime,
      totalLate,
      onTimeRate,
      lateRate,
      avgLateMinutes,
      totalOnline,
      totalOffline
    },
    departments,
    individuals,
    trends
  };
}

module.exports = {
  calculateAnalytics
};
