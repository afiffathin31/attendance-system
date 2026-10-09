function parseJamMulai(val) {
  if (!val) return "08:00";
  if (val instanceof Date) {
    return "DATE";
  }
  if (typeof val === "number") {
    let frac = val;
    if (frac > 1) {
      if (frac <= 24) return `${String(Math.floor(frac)).padStart(2, '0')}:00`;
      if (frac > 24 && frac <= 100) frac = frac / 100;
      else frac = frac - Math.floor(frac);
    }
    const totalMinutes = Math.round(frac * 1440);
    const h = Math.floor(totalMinutes / 60) % 24;
    const m = totalMinutes % 60;
    return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
  }
  const str = String(val).trim().replace(".", ":");
  const match = str.match(/(\d{1,2}):(\d{2})/);
  if (match) {
    return `${match[1].padStart(2, '0')}:${match[2]}`;
  }
  return "08:00";
}

function evaluateOnTime(timeStr, targetTimeStr, toleranceMinutes) {
  let [h, m] = (timeStr || "00:00").split(":").map(Number);
  if (isNaN(h)) h = 0;
  if (isNaN(m)) m = 0;

  const targetParsed = parseJamMulai(targetTimeStr);
  let [th, tm] = targetParsed.split(":").map(Number);
  if (isNaN(th)) th = 8;
  if (isNaN(tm)) tm = 0;

  const currentMin = h * 60 + m;
  const targetMin = th * 60 + tm;
  const diff = currentMin - targetMin;
  const tol = (toleranceMinutes !== undefined && toleranceMinutes !== null && !isNaN(toleranceMinutes)) ? Number(toleranceMinutes) : 15;

  if (diff <= tol) {
    return { status: "TEPAT WAKTU", lateMinutes: 0, keterlambatanText: "-" };
  } else {
    return { status: "TERLAMBAT", lateMinutes: diff, keterlambatanText: `${diff} Menit` };
  }
}

console.log("Test 10:21 vs 08:00 (tol 15):", evaluateOnTime("10:21:05", "08:00", 15));
console.log("Test 10:21 vs 10:00 (tol 15):", evaluateOnTime("10:21:05", "10:00", 15));
console.log("Test 10:21 vs 10:30 (tol 15):", evaluateOnTime("10:21:05", "10:30", 15));
console.log("Test 10:21 vs 10:15 (tol 15):", evaluateOnTime("10:21:05", "10:15", 15));
console.log("Test 10:21 vs 10.21 (tol 15):", evaluateOnTime("10:21:05", "10.21", 15));
console.log("Test 10:21 vs E4 default 09:00 (tol 5):", evaluateOnTime("10:21:05", "09:00", 5));