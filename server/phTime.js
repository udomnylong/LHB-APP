// Server-side Asia/Phnom_Penh (UTC+7) time helpers — mirrors Code.gs's phDateStr/phTimeStr
// so the ported attendance logic produces identical strings.
const PH_MS = 7 * 3600000;

function phNow() {
  return new Date(Date.now() + PH_MS);
}
function phDateStr(d) {
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
}
function phTimeStr(d) {
  return `${String(d.getUTCHours()).padStart(2, '0')}:${String(d.getUTCMinutes()).padStart(2, '0')}`;
}
function phTimeStrFull(d) {
  return `${phTimeStr(d)}:${String(d.getUTCSeconds()).padStart(2, '0')}`;
}

module.exports = { phNow, phDateStr, phTimeStr, phTimeStrFull };
