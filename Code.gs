// ============================================================
// LHB HR SYSTEM — Google Apps Script v5.5
// ============================================================

const SS_ID              = '16ryjqdieYbZAaG9phRMVInz_Yt6bP8KtWmEYXBcZRH0';
const TELEGRAM_TOKEN     = PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN') || '';
const TELEGRAM_CHAT      = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT')  || '549942306';
const TELEGRAM_GROUP     = PropertiesService.getScriptProperties().getProperty('TELEGRAM_Group') || '';
const WEBHOOK_URL        = 'https://script.google.com/macros/s/AKfycbwLvbmtgWiOMG-zssr8T_jgGPXEErvre4mekNlOYNKmzYSe4kjf5gzHgS5B5Pac0UHE/exec';
const FOOD_FOLDER_ID     = '1Ue7-K0QPDVwQcRszw5xF7b3SH25yGj5y';
const STAFF_PHOTO_FOLDER = '1BMeeqss2J_eoU-o8At7Wri-UNDzMO42DW7XzKeanz2vNgPrzJrICf5IL6OgAn6_ulWbS1B8X';
const FOLDER_ID          = FOOD_FOLDER_ID;

// ============================================================
// PERFORMANCE — Spreadsheet & Sheet Data Cache
// ============================================================

// Cache Spreadsheet object within one request execution (avoids repeated openById calls)
var _ss = null;
function getSS() {
  if (!_ss) _ss = SpreadsheetApp.openById(SS_ID);
  return _ss;
}

// Cache sheet data in CacheService (5 min TTL) — good for rarely-changing sheets like StaffInfo/User
var SHEET_CACHE_TTL = 300;
function getCachedSheetData(sheetName) {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('sd_' + sheetName);
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }
  var ws = getSS().getSheetByName(sheetName);
  if (!ws) return null;
  var data = ws.getDataRange().getValues();
  try { cache.put('sd_' + sheetName, JSON.stringify(data), SHEET_CACHE_TTL); } catch(e) {}
  return data;
}
function invalidateSheetCache(sheetName) {
  try {
    var c = CacheService.getScriptCache();
    c.remove('sd_'    + sheetName);
    c.remove('doget_' + sheetName);
  } catch(e) {}
}

// Per-sheet cache TTL (seconds)
function getCacheTTL(sheet, isFiltered) {
  if (isFiltered) return 30;           // filtered id+date → 30s (changes on each check-in)
  if (sheet === 'Project')   return 600;
  if (sheet === 'StaffInfo') return 300;
  if (sheet === 'CheckIn' || sheet === 'CheckOut') return 30;
  if (sheet === 'StaffLeave'|| sheet === 'StaffOT') return 120;
  return 60;
}

// ── Fast UTC+7 (Asia/Phnom_Penh) date helpers ──
// Avoids Utilities.formatDate() which costs ~50ms per call (GAS API overhead)
var PH_MS = 7 * 3600000;
function phDateStr(d) {
  var u = new Date(d.getTime() + PH_MS);
  return u.getUTCFullYear() + '-' +
    ('0'+(u.getUTCMonth()+1)).slice(-2) + '-' +
    ('0'+u.getUTCDate()).slice(-2);
}
function phTimeStr(d) {
  var u = new Date(d.getTime() + PH_MS);
  return ('0'+u.getUTCHours()).slice(-2) + ':' + ('0'+u.getUTCMinutes()).slice(-2);
}
function phTimeStrFull(d) {
  var u = new Date(d.getTime() + PH_MS);
  return ('0'+u.getUTCHours()).slice(-2) + ':' + ('0'+u.getUTCMinutes()).slice(-2) + ':' + ('0'+u.getUTCSeconds()).slice(-2);
}

// ============================================================
// WORK SCHEDULE — ប្ដូរតាមពេលវេលាការងាររបស់ LHB
// ============================================================
const TZ                = 'Asia/Phnom_Penh';   // UTC+7
const CHECKIN_LATE_H    = 8;   // Check In ក្រោយ 08:00 → Late
const CHECKIN_LATE_M    = 0;
const CHECKOUT_EARLY_H  = 17;  // Check Out មុន 17:00 → Early Leave
const CHECKOUT_EARLY_M  = 0;

// ============================================================
// SERVER-SIDE TIME ENFORCEMENT
// ⛔ Override client Time/Date/Timestamp with GAS server time.
//    Prevents phone-clock manipulation fraud.
// ============================================================

// Enforce server time for Attendance sheet CheckIn/CheckOut fields.
// Mutates p.data and p.keyDate so the subsequent upsert logic uses server date.
function enforceAttendanceTime_(p) {
  var now    = new Date();
  var sDate  = phDateStr(now);
  var sTime  = phTimeStr(now);
  var sH     = parseInt(sTime.split(':')[0]);
  var sMn    = parseInt(sTime.split(':')[1]);
  var total  = sH * 60 + sMn;

  var row    = p.data || {};
  row.Date   = sDate;
  p.keyDate  = sDate; // fix search key so upsert finds correct server-date row

  var clientCheckIn  = String(row.CheckIn  || '');
  var clientCheckOut = String(row.CheckOut || '');
  var clientLate     = String(row.Late     || '');
  var clientEarly    = String(row.Early    || '');
  Logger.log('[ATT-AUDIT] type:' + p.type
    + ' | ID:' + (row.ID||'?')
    + ' | clientCheckIn:'  + clientCheckIn
    + ' | clientCheckOut:' + clientCheckOut
    + ' | clientLate:'     + clientLate
    + ' | clientEarly:'    + clientEarly
    + ' | serverTime:'     + sTime);

  if (p.type === 'checkin') {
    row.CheckIn = sTime;
    var cutoff = CHECKIN_LATE_H * 60 + CHECKIN_LATE_M;
    if (total > cutoff) {
      var lateMin = total - cutoff;
      row.Late   = 'Late ' + lateMin + ' min';
      row.Status = 'Late';
    } else {
      row.Late   = '';
      row.Status = 'Present';
    }
  } else if (p.type === 'checkout') {
    row.CheckOut = sTime;
    var endCutoff = CHECKOUT_EARLY_H * 60 + CHECKOUT_EARLY_M;
    if (total < endCutoff) {
      row.Early = 'Early ' + (endCutoff - total) + ' min';
    } else {
      row.Early = '';
    }
  }

  p.data = row;
  return p;
}
function enforceServerTime_(row, sheet) {
  var now      = new Date();
  var sTime    = phTimeStrFull(now);
  var sDate    = phDateStr(now);
  var sTs      = String(now.getTime());

  // ── Audit log: compare client time vs server time ──
  var clientTime = String(row['Time'] || '?');
  var clientTs   = parseInt(row['Timestamp'] || 0);
  var driftSec   = clientTs > 0 ? Math.round((now.getTime() - clientTs) / 1000) : 9999;
  Logger.log('[TIME-AUDIT] ' + sheet
    + ' | ID:'   + (row['ID']   || '?')
    + ' | Name:' + (row['Name'] || '?')
    + ' | clientTime:'  + clientTime
    + ' | serverTime:'  + sTime
    + ' | driftSec:'    + driftSec);

  if (Math.abs(driftSec) > 300) {   // > 5 minutes difference
    Logger.log('[⚠️ SUSPICIOUS] Clock drift > 5min for ' + (row['ID']||'?')
               + ' | drift=' + driftSec + 's'
               + ' | clientTime=' + clientTime
               + ' | serverTime=' + sTime);
  }

  // ── Replace client values with verified server values ──
  row['Time']      = sTime;
  row['Date']      = sDate;
  row['Timestamp'] = sTs;

  // ── Recalculate LateEarly / Minutes from server time ──
  var hParts   = sTime.split(':');
  var sH       = parseInt(hParts[0]);
  var sMin     = parseInt(hParts[1]);
  var nowTotal = sH * 60 + sMin;

  if (sheet === 'CheckIn') {
    var cutoff = CHECKIN_LATE_H * 60 + CHECKIN_LATE_M;
    if (nowTotal > cutoff) {
      row['LateEarly'] = 'Late';
      row['Minutes']   = String(nowTotal - cutoff);
    } else {
      row['LateEarly'] = '';
      row['Minutes']   = '0';
    }
  } else if (sheet === 'CheckOut') {
    var endCutoff = CHECKOUT_EARLY_H * 60 + CHECKOUT_EARLY_M;
    if (nowTotal < endCutoff) {
      row['LateEarly'] = 'Early';
      row['Minutes']   = String(endCutoff - nowTotal);
    } else {
      row['LateEarly'] = '';
      row['Minutes']   = '0';
    }
  }

  return row;
}

// ============================================================
// doGet — Read Sheet / Photo
// ============================================================
function doGet(e) {
  try {
    var params = e.parameter || {};
    Logger.log('doGet: ' + JSON.stringify(params));

    if (params.action === 'getStaffPhoto') {
      try {
        var staffId = String(params.staffId || '').trim().toUpperCase();
        if (!staffId) return respond({ status:'error', msg:'staffId required' });
        var folder   = DriveApp.getFolderById(STAFF_PHOTO_FOLDER);
        var allFiles = folder.getFiles();
        while (allFiles.hasNext()) {
          var f     = allFiles.next();
          var fname = f.getName().toUpperCase().replace(/\.[^.]+$/, '');
          if (fname === staffId || fname.indexOf(staffId) >= 0 || staffId.indexOf(fname) >= 0) {
            f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            return respond({ status:'ok', url:'https://lh3.googleusercontent.com/d/' + f.getId(), fileId:f.getId(), name:f.getName() });
          }
        }
        return respond({ status:'notfound', msg:'No photo for staffId: ' + staffId });
      } catch(pe) { return respond({ status:'error', msg:pe.message }); }
    }

    // ── checkSession via GET ──
    if (params.action === 'checkSession') {
      var ws2=getSS().getSheetByName('User');
      if (!ws2) return respond({ status:'ok', active:false });
      var data2=getCachedSheetData('User'), hdrs2=data2[0].map(function(h){return String(h).trim();});
      var uIdx2=hdrs2.indexOf('Username'), tIdx2=hdrs2.indexOf('SessionToken'), tsIdx2=hdrs2.indexOf('SessionTime');
      if (tIdx2<0||tsIdx2<0) return respond({ status:'ok', active:false });
      var username2=String(params.username||'').trim().toLowerCase();
      var myToken2=String(params.myToken||'').trim();
      var SESSION_TIMEOUT2=10*60*1000;
      for (var i2=1;i2<data2.length;i2++) {
        if (String(data2[i2][uIdx2]||'').trim().toLowerCase()===username2) {
          var tok2=String(data2[i2][tIdx2]||'').trim();
          var ts2=Number(data2[i2][tsIdx2]||0);
          var now2=new Date().getTime();
          if (tok2 && (now2-ts2)<SESSION_TIMEOUT2 && tok2!==myToken2) {
            return respond({ status:'ok', active:true, since:ts2 });
          }
          return respond({ status:'ok', active:false });
        }
      }
      return respond({ status:'ok', active:false });
    }

    // ── setSession via GET ──
    if (params.action === 'setSession') {
      var ws3=getSS().getSheetByName('User');
      if (!ws3) return respond({ status:'error', msg:'User sheet not found' });
      var data3=getCachedSheetData('User'), hdrs3=data3[0].map(function(h){return String(h).trim();});
      var uIdx3=hdrs3.indexOf('Username'), tIdx3=hdrs3.indexOf('SessionToken'), tsIdx3=hdrs3.indexOf('SessionTime');
      if (tIdx3<0) { ws3.getRange(1,ws3.getLastColumn()+1).setValue('SessionToken'); tIdx3=ws3.getLastColumn()-1; }
      if (tsIdx3<0) { ws3.getRange(1,ws3.getLastColumn()+1).setValue('SessionTime'); tsIdx3=ws3.getLastColumn()-1; }
      var username3=String(params.username||'').trim().toLowerCase();
      var token3=String(params.token||'').trim();
      var now3=new Date().getTime();
      for (var i3=1;i3<data3.length;i3++) {
        if (String(data3[i3][uIdx3]||'').trim().toLowerCase()===username3) {
          ws3.getRange(i3+1,tIdx3+1).setValue(token3);
          ws3.getRange(i3+1,tsIdx3+1).setValue(now3);
          invalidateSheetCache('User');
          return respond({ status:'ok' });
        }
      }
      return respond({ status:'error', msg:'User not found' });
    }

    // ── clearSession via GET ──
    if (params.action === 'clearSession') {
      var ws4=getSS().getSheetByName('User');
      if (!ws4) return respond({ status:'ok' });
      var data4=getCachedSheetData('User'), hdrs4=data4[0].map(function(h){return String(h).trim();});
      var uIdx4=hdrs4.indexOf('Username'), tIdx4=hdrs4.indexOf('SessionToken'), tsIdx4=hdrs4.indexOf('SessionTime');
      if (tIdx4<0) return respond({ status:'ok' });
      var username4=String(params.username||'').trim().toLowerCase();
      var token4=String(params.token||'').trim();
      for (var i4=1;i4<data4.length;i4++) {
        if (String(data4[i4][uIdx4]||'').trim().toLowerCase()===username4) {
          var stored4=String(data4[i4][tIdx4]||'').trim();
          if (!token4 || stored4===token4) {
            ws4.getRange(i4+1,tIdx4+1).setValue('');
            ws4.getRange(i4+1,tsIdx4+1).setValue('');
            invalidateSheetCache('User');
          }
          return respond({ status:'ok' });
        }
      }
      return respond({ status:'ok' });
    }

    if (params.action === 'getPhotoUrl') {
      try {
        var fileName  = String(params.fileName || '').trim();
        var staffIter = DriveApp.getFolderById(STAFF_PHOTO_FOLDER).getFilesByName(fileName);
        var foodIter  = DriveApp.getFolderById(FOOD_FOLDER_ID).getFilesByName(fileName);
        var file = null;
        if (staffIter.hasNext()) file = staffIter.next();
        else if (foodIter.hasNext()) file = foodIter.next();
        if (file) {
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          return respond({ status:'ok', url:'https://lh3.googleusercontent.com/d/' + file.getId(), fileId:file.getId() });
        }
        return respond({ status:'notfound' });
      } catch(de) { return respond({ status:'error', msg:de.message }); }
    }

    var sheetRaw = String(params.sheet || '').trim();
    var sheet    = (sheetRaw && sheetRaw !== 'undefined' && sheetRaw !== 'null') ? sheetRaw : 'StaffInfo';

    // ── Server-side filter params (CheckIn/CheckOut fast path) ──
    var filterId   = String(params.id   || '').trim();
    var filterDate = String(params.date || '').trim();
    var isFiltered = !!(filterId && filterDate);

    // ── Output cache check ──
    var outKey = 'doget_' + sheet + (isFiltered ? '_' + filterId + '_' + filterDate : '');
    var sc = CacheService.getScriptCache();
    if (!params.nocache) {
      var hit = sc.get(outKey);
      if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
    }

    var ss = getSS();
    var ws = ss.getSheetByName(sheet);
    if (!ws) {
      var avail = ss.getSheets().map(function(s){ return s.getName(); }).join(', ');
      return respond({ status:'error', msg:'Sheet "' + sheet + '" not found. Available: ' + avail });
    }

    var vals = ws.getDataRange().getValues();
    if (vals.length < 2) return respond({ status:'ok', data:[] });

    var headers = vals[0].map(function(h){ return String(h).trim(); });

    // Pre-compute column indices for filter (avoid indexOf per row)
    var idColIdx   = isFiltered ? headers.indexOf('ID')   : -1;
    var dateColIdx = isFiltered ? headers.indexOf('Date') : -1;

    var rows = [];
    for (var i = 1; i < vals.length; i++) {
      var row = vals[i];
      if (!row.some(function(c){ return c !== '' && c !== null; })) continue;

      // ── Server-side filter: skip non-matching rows early ──
      if (isFiltered) {
        var rId   = idColIdx   >= 0 ? String(row[idColIdx]   || '').trim() : '';
        var rDate = dateColIdx >= 0 ? String(row[dateColIdx] || '').trim() : '';
        if (row[dateColIdx] instanceof Date) rDate = phDateStr(row[dateColIdx]);
        else rDate = rDate.slice(0, 10);
        if (rId !== filterId || rDate !== filterDate) continue;
      }

      var obj = {};
      headers.forEach(function(h, j) {
        var v = row[j];
        if (v instanceof Date) {
          // Time-only: 1899-12-30 → format as HH:mm (fast, no API call)
          if (v.getFullYear() === 1899 && v.getMonth() === 11 && v.getDate() === 30) {
            v = phTimeStr(v);
          } else {
            v = phDateStr(v);
          }
        }
        obj[h] = (v !== undefined && v !== null) ? String(v) : '';
      });
      rows.push(obj);
    }

    // ── Cache and return ──
    var result = JSON.stringify({ status:'ok', data:rows });
    try { sc.put(outKey, result, getCacheTTL(sheet, isFiltered)); } catch(e) {}
    Logger.log('Sheet ' + sheet + ': ' + rows.length + ' rows (cached ' + getCacheTTL(sheet, isFiltered) + 's)');
    return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    Logger.log('doGet ERROR: ' + err.message);
    return respond({ status:'error', msg:err.message });
  }
}

// ============================================================
// doPost — Write Data / Telegram
// ============================================================
function doPost(e) {
  try {
    var raw = e.postData ? e.postData.contents : '';
    if (!raw || raw.length === 0)
      return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);

    var parsed = null;
    try { parsed = JSON.parse(raw); } catch(ex) {}

    if (parsed && parsed.update_id) {
      handleTelegramUpdate(parsed);
      return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
    }

    var p = parsed;
    if (!p) return respond({ status:'error', msg:'Invalid JSON' });
    Logger.log('doPost action: ' + p.action);

    // ── sendOTP ──
    if (p.action === 'sendOTP') {
      var phone = String(p.phone||'').replace(/[^0-9]/g,'');
      if (!phone || phone.length < 8) return respond({ status:'error', msg:'Phone invalid' });
      var ws = getSS().getSheetByName('StaffInfo');
      var data = getCachedSheetData('StaffInfo');
      var hdrs = data[0].map(function(h){ return String(h).trim(); });
      var phoneIdx = hdrs.indexOf('Phone'), chatIdx = hdrs.indexOf('TelegramChatId');
      var otpIdx = hdrs.indexOf('OTP'), expIdx = hdrs.indexOf('OTPExpire');
      if (chatIdx<0){ws.getRange(1,ws.getLastColumn()+1).setValue('TelegramChatId');}
      if (otpIdx<0) {ws.getRange(1,ws.getLastColumn()+1).setValue('OTP');}
      if (expIdx<0) {ws.getRange(1,ws.getLastColumn()+1).setValue('OTPExpire');}
      var hdrs2=ws.getRange(1,1,1,ws.getLastColumn()).getValues()[0].map(function(h){return String(h).trim();});
      var chatIdx2=hdrs2.indexOf('TelegramChatId'), otpIdx2=hdrs2.indexOf('OTP'), expIdx2=hdrs2.indexOf('OTPExpire');
      var pn0=phone.replace(/^0+/,''), staffRow=null, rowNum=-1;
      for (var i=1;i<data.length;i++){
        var rp=String(data[i][phoneIdx]||'').replace(/[^0-9]/g,'');
        if(rp===phone||rp.replace(/^0+/,'')===pn0||rp.slice(-9)===phone.slice(-9)){staffRow=data[i];rowNum=i+1;break;}
      }
      if (!staffRow) return respond({ status:'error', msg:'Phone '+phone+' not found' });
      var chatId=String(staffRow[chatIdx>=0?chatIdx:chatIdx2]||'').trim();
      if (!chatId) return respond({ status:'error', msg:'Not registered. Send /start to @lhb_system_bot' });
      var otp=String(Math.floor(100000+Math.random()*900000)), expire=new Date().getTime()+5*60*1000;
      ws.getRange(rowNum,otpIdx2+1).setValue(otp);
      ws.getRange(rowNum,expIdx2+1).setValue(expire);
      invalidateSheetCache('StaffInfo');
      var name=String(data[rowNum-1][hdrs.indexOf('Name')]||'');
      UrlFetchApp.fetch('https://api.telegram.org/bot'+TELEGRAM_TOKEN+'/sendMessage',{
        method:'post',contentType:'application/json',muteHttpExceptions:true,
        payload:JSON.stringify({chat_id:chatId,text:'LHB HR OTP\n\n'+name+'\nCode: '+otp+'\n\nExpire 5min'})
      });
      return respond({ status:'ok' });
    }

    // ── verifyOTP ──
    if (p.action === 'verifyOTP') {
      var phone=String(p.phone||'').replace(/[^0-9]/g,''), code=String(p.code||'').trim();
      var ws=getSS().getSheetByName('StaffInfo');
      var data=getCachedSheetData('StaffInfo'), hdrs=data[0].map(function(h){return String(h).trim();});
      var phoneIdx=hdrs.indexOf('Phone'), otpIdx=hdrs.indexOf('OTP'), expIdx=hdrs.indexOf('OTPExpire');
      if (otpIdx<0) return respond({ status:'error', msg:'OTP column missing. Run setupHeaders().' });
      var pn0=phone.replace(/^0+/,''), staffRow=null, rowNum=-1;
      for (var i=1;i<data.length;i++){
        var rp=String(data[i][phoneIdx]||'').replace(/[^0-9]/g,'');
        if(rp===phone||rp.replace(/^0+/,'')===pn0||rp.slice(-9)===phone.slice(-9)){staffRow=data[i];rowNum=i+1;break;}
      }
      if (!staffRow) return respond({ status:'error', msg:'Phone not found' });
      var savedOtp=String(staffRow[otpIdx]||'').trim(), savedExp=Number(staffRow[expIdx]||0);
      if (!savedOtp)                     return respond({ status:'error', msg:'OTP not requested' });
      if (new Date().getTime()>savedExp) return respond({ status:'error', msg:'OTP Expired' });
      if (savedOtp!==code)               return respond({ status:'error', msg:'OTP incorrect' });
      var staffObj={};
      hdrs.forEach(function(h,j){staffObj[h]=staffRow[j]!==undefined?String(staffRow[j]):'';});
      ws.getRange(rowNum,otpIdx+1).setValue('');
      ws.getRange(rowNum,expIdx+1).setValue('');
      invalidateSheetCache('StaffInfo');
      return respond({ status:'ok', staff:staffObj });
    }

    // ── append ──
    if (p.action === 'append') {
      var ss=getSS(), ws=ss.getSheetByName(p.sheet);
      if (!ws) return respond({ status:'error', msg:'Sheet not found: '+p.sheet });
      var headers=ws.getRange(1,1,1,ws.getLastColumn()).getValues()[0].map(function(h){return String(h).trim();});
      var row=p.data||{};
      // ⛔ SECURITY: replace client time with verified server time
      if (p.sheet === 'CheckIn' || p.sheet === 'CheckOut') {
        row = enforceServerTime_(row, p.sheet);
      }
      ws.appendRow(headers.map(function(h){return row[h]!==undefined?row[h]:'';}));
      invalidateSheetCache(p.sheet); // clear cache so next read is fresh
      sendCheckNotification(p.sheet, row);
      return respond({ status:'ok' });
    }

    // ── upsert ──
    if (p.action === 'upsert') {
      var ss=getSS(), ws=ss.getSheetByName(p.sheet);
      if (!ws) return respond({ status:'error', msg:'Sheet not found: '+p.sheet });
      // ⛔ SECURITY: enforce server time for Attendance CheckIn/CheckOut
      if (p.sheet === 'Attendance' && (p.type === 'checkin' || p.type === 'checkout')) {
        p = enforceAttendanceTime_(p);
      }
      var data=ws.getDataRange().getValues(), headers=data[0].map(function(h){return String(h).trim();});
      var idIdx=headers.indexOf('ID'), dateIdx=headers.indexOf('Date');
      for (var i=1;i<data.length;i++){
        if (String(data[i][idIdx]).trim()===String(p.keyValue).trim()&&normDate(data[i][dateIdx])===String(p.keyDate).trim()){
          var row=p.data||{};
          // ⛔ SECURITY: replace client time with verified server time
          if (p.sheet === 'CheckIn' || p.sheet === 'CheckOut') {
            row = enforceServerTime_(row, p.sheet);
          }
          // Batch write: merge into existing row, then setValues once (faster than N setValue calls)
          var merged=headers.map(function(h,j){return (row[h]!==undefined&&row[h]!=='')?row[h]:data[i][j];});
          ws.getRange(i+1,1,1,headers.length).setValues([merged]);
          sendCheckNotification(p.sheet, row);
          return respond({ status:'ok', action:'updated' });
        }
      }
      var row=p.data||{};
      // ⛔ SECURITY: replace client time with verified server time
      if (p.sheet === 'CheckIn' || p.sheet === 'CheckOut') {
        row = enforceServerTime_(row, p.sheet);
      }
      ws.appendRow(headers.map(function(h){return row[h]!==undefined?row[h]:'';}));
      sendCheckNotification(p.sheet, row);
      return respond({ status:'ok', action:'appended' });
    }

    // ── update ──
    if (p.action === 'update') {
      var ss=getSS(), ws=ss.getSheetByName(p.sheet);
      if (!ws) return respond({ status:'error', msg:'Sheet not found: '+p.sheet });
      var data=ws.getDataRange().getValues(), headers=data[0].map(function(h){return String(h).trim();});
      var keyField=String(p.keyField||'ID').trim();
      var idIdx=headers.indexOf(keyField);
      if (idIdx<0) idIdx=headers.indexOf('ID');
      var keyVal=String(p.keyValue||p.id||'').trim();
      // Optional secondary date key — matches StartDate or Date column
      var keyDate=String(p.keyDate||'').trim();
      var dateColIdx=-1;
      if (keyDate) {
        dateColIdx=headers.indexOf('StartDate');
        if (dateColIdx<0) dateColIdx=headers.indexOf('Date');
      }
      for (var i=1;i<data.length;i++){
        if (String(data[i][idIdx]).trim()===keyVal){
          if (keyDate && dateColIdx>=0 && normDate(data[i][dateColIdx])!==keyDate) continue;
          var row=p.data||{};
          // Batch write: merge into existing row, then setValues once (faster than N setValue calls)
          var merged=headers.map(function(h,j){return row[h]!==undefined?row[h]:data[i][j];});
          ws.getRange(i+1,1,1,headers.length).setValues([merged]);
          return respond({ status:'ok' });
        }
      }
      return respond({ status:'error', msg:'Row not found: '+keyVal+(keyDate?' @ '+keyDate:'')+' in '+keyField });
    }

    // ── delete ──
    if (p.action === 'delete') {
      var ss=getSS(), ws=ss.getSheetByName(p.sheet);
      if (!ws) return respond({ status:'error', msg:'Sheet not found: '+p.sheet });
      var data=ws.getDataRange().getValues(), headers=data[0].map(function(h){return String(h).trim();});
      var keyField=String(p.keyField||'ID').trim();
      var idIdx=headers.indexOf(keyField);
      if (idIdx<0) idIdx=headers.indexOf('ID');
      var keyVal=String(p.keyValue||p.id||'').trim();
      for (var i=data.length-1;i>=1;i--){
        if (String(data[i][idIdx]).trim()===keyVal){ws.deleteRow(i+1);return respond({ status:'ok' });}
      }
      return respond({ status:'error', msg:'Row not found: '+keyVal+' in '+keyField });
    }

    // ── deleteByIdDate ──
    if (p.action === 'deleteByIdDate') {
      var ss=getSS(), ws=ss.getSheetByName(p.sheet);
      if (!ws) return respond({ status:'ok', msg:'Sheet not found' });
      var data=ws.getDataRange().getValues(), headers=data[0].map(function(h){return String(h).trim();});
      var idIdx=headers.indexOf('ID'), dateIdx=headers.indexOf('Date');
      var deleted=0;
      for (var i=data.length-1;i>=1;i--){
        if (String(data[i][idIdx]).trim()===String(p.keyId).trim()&&normDate(data[i][dateIdx])===String(p.keyDate).trim()){
          ws.deleteRow(i+1); deleted++;
        }
      }
      return respond({ status:'ok', deleted:deleted });
    }

    // ── uploadPhoto — upload file ថ្មីទៅ Drive Folder ──
    if (p.action === 'uploadPhoto') {
      var folderId = p.folderId || FOOD_FOLDER_ID;
      var decoded  = Utilities.base64Decode(p.base64);
      var blob     = Utilities.newBlob(decoded, p.mimeType||'image/jpeg', p.fileName||'photo.jpg');
      var folder   = DriveApp.getFolderById(folderId);
      var file     = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      var fid = file.getId();
      var url = 'https://drive.google.com/file/d/' + fid + '/view?usp=sharing';
      return respond({ status:'ok', url:url, fileId:fid });
    }

    // ── replacePhoto — replace content Drive file ដើម (fileId/URL មិនប្ដូរ) ──
    // ⚠️ ត្រូវការ: Services → Drive API → Enable
    if (p.action === 'replacePhoto') {
      var rfid  = String(p.fileId   || '').trim();
      var rb64  = String(p.base64   || '').trim();
      var rmime = String(p.mimeType || 'image/jpeg').trim();
      if (!rfid) return respond({ status:'error', msg:'replacePhoto: fileId required' });
      if (!rb64) return respond({ status:'error', msg:'replacePhoto: base64 required' });
      try {
        var rblob = Utilities.newBlob(Utilities.base64Decode(rb64), rmime);
        Drive.Files.update({ mimeType: rmime }, rfid, rblob);
        var rurl = 'https://drive.google.com/file/d/' + rfid + '/view?usp=sharing';
        return respond({ status:'ok', fileId:rfid, url:rurl });
      } catch(re) {
        Logger.log('replacePhoto error: ' + re.message);
        return respond({ status:'error', msg:'replacePhoto failed: ' + re.message });
      }
    }

    // ── updateStaffPhoto — update column Photo ក្នុង StaffInfo Sheet ──
    if (p.action === 'updateStaffPhoto') {
      var ugmail    = String(p.gmail    || '').toLowerCase().trim();
      var uid       = String(p.id       || '').trim();
      var uphotoUrl = String(p.photoUrl || '').trim();
      if (!uphotoUrl)      return respond({ status:'error', msg:'updateStaffPhoto: photoUrl required' });
      if (!ugmail && !uid) return respond({ status:'error', msg:'updateStaffPhoto: gmail or id required' });
      try {
        var uws        = getSS().getSheetByName('StaffInfo');
        if (!uws)      return respond({ status:'error', msg:'StaffInfo sheet not found' });
        var udata      = getCachedSheetData('StaffInfo');
        var uhdrs      = udata[0].map(function(h){ return String(h).trim(); });
        var uhdrsLow   = uhdrs.map(function(h){ return h.toLowerCase(); });

        // ── Find columns (case-insensitive + name variations) ──
        function findCol(names) {
          for (var ni = 0; ni < names.length; ni++) {
            var idx = uhdrs.indexOf(names[ni]);
            if (idx >= 0) return idx;
          }
          for (var ni2 = 0; ni2 < names.length; ni2++) {
            var idx2 = uhdrsLow.indexOf(names[ni2].toLowerCase());
            if (idx2 >= 0) return idx2;
          }
          return -1;
        }

        var ugmIdx = findCol(['Gmail','Email','email','gmail','G-Mail']);
        var uidIdx = findCol(['ID','id','StaffID','staffid']);
        var uphIdx = findCol(['Photo','photo','PhotoUrl','photoUrl']);

        Logger.log('updateStaffPhoto — gmail:'+ugmail+' id:'+uid+' gmailCol:'+ugmIdx+' idCol:'+uidIdx+' photoCol:'+uphIdx);
        Logger.log('Headers: ' + uhdrs.join(' | '));

        if (uphIdx < 0) return respond({ status:'error', msg:'Photo column not found. Headers: ' + uhdrs.join(', ') });

        var uupdated = false;
        for (var ui = 1; ui < udata.length; ui++) {
          var rGmail = ugmIdx >= 0 ? String(udata[ui][ugmIdx] || '').toLowerCase().trim() : '';
          var rId    = uidIdx >= 0 ? String(udata[ui][uidIdx]  || '').trim() : '';
          var matchG = ugmail && rGmail && (rGmail === ugmail);
          var matchI = uid    && rId    && (rId    === uid);
          if (matchG || matchI) {
            Logger.log('Match row '+(ui+1)+': rGmail='+rGmail+' rId='+rId);
            uws.getRange(ui + 1, uphIdx + 1).setValue(uphotoUrl);
            uupdated = true;
            break;
          }
        }
        SpreadsheetApp.flush();
        invalidateSheetCache('StaffInfo');

        if (!uupdated) {
          // Log first 3 rows to help debug
          var sample = [];
          for (var si = 1; si < Math.min(4, udata.length); si++) {
            sample.push('row'+si+': gmail='+(ugmIdx>=0?udata[si][ugmIdx]:'-')+' id='+(uidIdx>=0?udata[si][uidIdx]:'-'));
          }
          Logger.log('Not found. Sample: ' + sample.join(' | '));
          return respond({ status:'error', msg:'Staff not found — sent: gmail='+ugmail+' id='+uid });
        }

        Logger.log('updateStaffPhoto OK: '+(ugmail||uid)+' → '+uphotoUrl);
        return respond({ status:'ok', msg:'Photo updated' });
      } catch(ue) {
        Logger.log('updateStaffPhoto error: ' + ue.message);
        return respond({ status:'error', msg:ue.message });
      }
    }

    // ── appendWithPhotos (Food) ──
    if (p.action === 'appendWithPhotos') {
      var ss=getSS();
      var folder=DriveApp.getFolderById(p.folderId||FOOD_FOLDER_ID);
      var row=p.data||{};
      var photoFields={PhotoMorning:'photoMorning',PhotoLunch:'photoLunch',PhotoEvening:'photoEvening'};
      for (var key in photoFields){
        var b64=p.photos?p.photos[photoFields[key]]:null;
        if(b64&&b64.length>100){
          try{
            var blob=Utilities.newBlob(Utilities.base64Decode(b64),'image/jpeg',(p.fileNames&&p.fileNames[photoFields[key]])||key+'.jpg');
            var f=folder.createFile(blob);
            f.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
            row[key]='https://lh3.googleusercontent.com/d/'+f.getId();
          }catch(pe){Logger.log('Photo err '+key+': '+pe.message);}
        }
      }
      var ws=ss.getSheetByName('Food');
      if (!ws) return respond({ status:'error', msg:'Food sheet not found' });
      var headers=ws.getRange(1,1,1,ws.getLastColumn()).getValues()[0].map(function(h){return String(h).trim();});
      ws.appendRow(headers.map(function(h){return row[h]!==undefined?row[h]:'';}));
      return respond({ status:'ok', photoMorning:row.PhotoMorning||'', photoLunch:row.PhotoLunch||'', photoEvening:row.PhotoEvening||'' });
    }

    // ── checkSession (POST) ──
    if (p.action === 'checkSession') {
      var ws=getSS().getSheetByName('User');
      if (!ws) return respond({ status:'ok', active: false });
      var data=getCachedSheetData('User'), hdrs=data[0].map(function(h){return String(h).trim();});
      var uIdx=hdrs.indexOf('Username'), tIdx=hdrs.indexOf('SessionToken'), tsIdx=hdrs.indexOf('SessionTime');
      if (tIdx<0||tsIdx<0) return respond({ status:'ok', active: false });
      var username=String(p.username||'').trim().toLowerCase();
      var SESSION_TIMEOUT = 10 * 60 * 1000;
      for (var i=1;i<data.length;i++) {
        if (String(data[i][uIdx]||'').trim().toLowerCase()===username) {
          var tok=String(data[i][tIdx]||'').trim();
          var ts=Number(data[i][tsIdx]||0);
          var now=new Date().getTime();
          if (tok && (now-ts)<SESSION_TIMEOUT && tok!==String(p.myToken||'').trim()) {
            return respond({ status:'ok', active: true, since: ts });
          }
          return respond({ status:'ok', active: false });
        }
      }
      return respond({ status:'ok', active: false });
    }

    // ── setSession (POST) ──
    if (p.action === 'setSession') {
      var ws=getSS().getSheetByName('User');
      if (!ws) return respond({ status:'error', msg:'User sheet not found' });
      var data=getCachedSheetData('User'), hdrs=data[0].map(function(h){return String(h).trim();});
      var uIdx=hdrs.indexOf('Username'), tIdx=hdrs.indexOf('SessionToken'), tsIdx=hdrs.indexOf('SessionTime');
      if (tIdx<0) { ws.getRange(1,ws.getLastColumn()+1).setValue('SessionToken'); tIdx=ws.getLastColumn()-1; hdrs.push('SessionToken'); }
      if (tsIdx<0) { ws.getRange(1,ws.getLastColumn()+1).setValue('SessionTime'); tsIdx=ws.getLastColumn()-1; hdrs.push('SessionTime'); }
      var username=String(p.username||'').trim().toLowerCase();
      var token=String(p.token||'').trim();
      var now=new Date().getTime();
      for (var i=1;i<data.length;i++) {
        if (String(data[i][uIdx]||'').trim().toLowerCase()===username) {
          ws.getRange(i+1,tIdx+1).setValue(token);
          ws.getRange(i+1,tsIdx+1).setValue(now);
          invalidateSheetCache('User');
          return respond({ status:'ok' });
        }
      }
      return respond({ status:'error', msg:'User not found' });
    }

    // ── clearSession (POST) ──
    if (p.action === 'clearSession') {
      var ws=getSS().getSheetByName('User');
      if (!ws) return respond({ status:'ok' });
      var data=getCachedSheetData('User'), hdrs=data[0].map(function(h){return String(h).trim();});
      var uIdx=hdrs.indexOf('Username'), tIdx=hdrs.indexOf('SessionToken'), tsIdx=hdrs.indexOf('SessionTime');
      if (tIdx<0) return respond({ status:'ok' });
      var username=String(p.username||'').trim().toLowerCase();
      var token=String(p.token||'').trim();
      for (var i=1;i<data.length;i++) {
        if (String(data[i][uIdx]||'').trim().toLowerCase()===username) {
          var storedTok=String(data[i][tIdx]||'').trim();
          if (!token || storedTok===token) {
            ws.getRange(i+1,tIdx+1).setValue('');
            ws.getRange(i+1,tsIdx+1).setValue('');
            invalidateSheetCache('User');
          }
          return respond({ status:'ok' });
        }
      }
      return respond({ status:'ok' });
    }

    // ── setupHeaders ──
    if (p.action === 'setupHeaders') { setupHeaders(); return respond({ status:'ok', msg:'Done' }); }

    return respond({ status:'error', msg:'Unknown action: '+p.action });

  } catch(err) {
    Logger.log('doPost ERROR: '+err.message);
    return respond({ status:'error', msg:err.message });
  }
}

// ============================================================
// TELEGRAM BOT
// ============================================================
function handleTelegramUpdate(body) {
  try {
    var uid=String(body.update_id||'');
    if(uid){var cache=CacheService.getScriptCache();if(cache.get('tg_'+uid))return;cache.put('tg_'+uid,'1',21600);}
    var msg=body.message||body.edited_message;
    if(!msg)return;
    if(Math.floor(new Date().getTime()/1000)-(msg.date||0)>30)return;
    var chatId=String(msg.chat.id), text=(msg.text||'').trim();
    Logger.log('TG: '+chatId+' → '+text);
    if(text==='/start'){sendTelegramMsg(chatId,'LHB HR Bot!\n\nRegister:\n/register 0XXXXXXXXX');return;}
    if(text.startsWith('/register')){
      var phone=text.replace('/register','').trim().replace(/[^0-9]/g,'');
      if(!phone||phone.length<8){sendTelegramMsg(chatId,'Format: /register 0XXXXXXXXX');return;}
      sendTelegramMsg(chatId,registerStaff(phone,chatId));return;
    }
    if(text==='/status'){sendTelegramMsg(chatId,getRegisteredInfo(chatId)||'Not registered.');}
  } catch(e){Logger.log('TG error: '+e.message);}
}

function registerStaff(phone, chatId) {
  var ws=getSS().getSheetByName('StaffInfo');
  var data=getCachedSheetData('StaffInfo'), hdrs=data[0].map(function(h){return String(h).trim();});
  var phoneIdx=hdrs.indexOf('Phone'), nameIdx=hdrs.indexOf('Name'), idIdx=hdrs.indexOf('ID');
  var chatIdx=hdrs.indexOf('TelegramChatId');
  if(chatIdx<0){ws.getRange(1,ws.getLastColumn()+1).setValue('TelegramChatId');chatIdx=ws.getLastColumn()-1;}
  var pn0=phone.replace(/^0+/,'');
  for(var i=1;i<data.length;i++){
    var p=String(data[i][phoneIdx]||'').replace(/[^0-9]/g,'');
    if(p===phone||p.replace(/^0+/,'')===pn0||p.slice(-9)===phone.slice(-9)){
      ws.getRange(i+1,chatIdx+1).setValue(chatId);
      invalidateSheetCache('StaffInfo');
      return 'Register OK!\n'+String(data[i][nameIdx])+' ('+String(data[i][idIdx])+')\nPhone: 0'+pn0;
    }
  }
  return 'Phone 0'+pn0+' not found. Contact Admin.';
}

function getRegisteredInfo(chatId) {
  var ws=getSS().getSheetByName('StaffInfo');
  var data=getCachedSheetData('StaffInfo'), hdrs=data[0].map(function(h){return String(h).trim();});
  var chatIdx=hdrs.indexOf('TelegramChatId'), nameIdx=hdrs.indexOf('Name'), idIdx=hdrs.indexOf('ID');
  if(chatIdx<0)return null;
  for(var i=1;i<data.length;i++){if(String(data[i][chatIdx]).trim()===chatId)return 'Registered: '+data[i][nameIdx]+' ('+data[i][idIdx]+')';}
  return null;
}

function sendTelegramMsg(chatId, text) {
  if(!TELEGRAM_TOKEN)return;
  UrlFetchApp.fetch('https://api.telegram.org/bot'+TELEGRAM_TOKEN+'/sendMessage',{
    method:'post',contentType:'application/json',muteHttpExceptions:true,
    payload:JSON.stringify({chat_id:chatId,text:text})
  });
}

function sendCheckNotification(sheet, row) {
  try {
    if(sheet!=='CheckIn'&&sheet!=='CheckOut')return;
    // Use the enforced server time already stored in row.Time; fallback to now
    var serverTime = (row['Time'] || '').substring(0, 5) || phTimeStr(new Date());
    var emoji=sheet==='CheckIn'?'🟢':'🟡', type=sheet==='CheckIn'?'CHECK IN':'CHECK OUT';
    var msg=emoji+' '+type+' | '+serverTime+'\n\n'+(row.Name||'')+'  '+(row.ID||'')+'\n'+(row.Position||'')+' | '+(row.Department||'')+'\n'+(row.ProjectName||'');
    var target=TELEGRAM_GROUP||TELEGRAM_CHAT;
    if(target)sendTelegramMsg(target,msg);
  }catch(e){Logger.log('Notify: '+e.message);}
}

function normDate(v) {
  if(!v)return '';
  if(v instanceof Date) return phDateStr(v);
  return String(v).trim().slice(0,10);
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// SETUP HEADERS — v5.3
// ============================================================
function setupHeaders() {
  var ss = getSS();
  var HEADERS = {
    User:       ['Username','Password','Name','Role','Email','Department','Position','SessionToken','SessionTime'],
    StaffInfo:  ['ID','Name','NameLatin','Sex','LV','Position','Department','ProjectName',
                 'DateOfBirth','StartingDate','ResignDate','Salary','Gmail','BankName','BankNumber',
                 'Photo','Phone','EmploymentStatus','TelegramChatId','OTP','OTPExpire'],
    Attendance: ['ID','Name','Position','Department','ProjectName','Date','CheckIn','CheckOut','Late','Early','Status'],
    StaffLeave: ['ID','Name','TypeOfLeave','StartDate','EndDate','Days','Reason','Status'],
    Project:    ['ProjectID','ProjectName','Location','Latitude','Longitude','Radius','Status'],
    StaffOT:    ['ID','Name','Date','Hours','TimeFrom','TimeTo','TypeOfWork','Reason','Status'],
    CheckIn:    ['ID','Name','Gmail','ProjectName','Date','Time','Timestamp','Latitude','Longitude','Accuracy','LateEarly','Minutes','Position','Department'],
    CheckOut:   ['ID','Name','Gmail','ProjectName','Date','Time','Timestamp','Latitude','Longitude','Accuracy','LateEarly','Minutes','Position','Department'],
    Food:       ['Date','ID','Name','Sex','Position','ProjectName','Morning','Lunch','Evening','Total','UnitPrice','TotalPrice','PhotoMorning','PhotoLunch','PhotoEvening','Comment','Remark'],
    WorkPlace:  ['Date','Time','ID','Name','Department','ProjectName','Comment','Photo','Status'],
    Comment:    ['Date','Time','ID','Name','Department','ProjectName','Comment','Photo','Status'],
    EvaluateStaff: ['EvalNo','RequestBy','StaffName','StaffID','DateEvaluate','KPIScore','PreviousSalary','CurrentSalary','ApprovedBy','Remark'],
  };
  for (var name in HEADERS) {
    var headers  = HEADERS[name];
    var ws       = ss.getSheetByName(name) || ss.insertSheet(name);
    var lastCol  = ws.getLastColumn();
    var existing = lastCol > 0
      ? ws.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h).trim(); })
      : [];
    var added = 0;
    headers.forEach(function(h) {
      if (existing.indexOf(h) < 0) {
        ws.getRange(1, ws.getLastColumn() + 1).setValue(h);
        existing.push(h);
        added++;
      }
    });
    Logger.log(name + ': ' + (added === 0 ? 'OK ✅' : added + ' column(s) added ✅'));
  }
  Logger.log('setupHeaders() done! ✅');
}

// ============================================================
// ADD EmploymentStatus COLUMN
// ============================================================
function addEmploymentStatusColumn() {
  var ss = getSS();
  var ws = ss.getSheetByName('StaffInfo');
  if (!ws) { Logger.log('❌ StaffInfo not found!'); return; }
  var lastCol  = ws.getLastColumn();
  var lastRow  = ws.getLastRow();
  var headers  = ws.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h).trim(); });
  var colIndex = headers.indexOf('EmploymentStatus');
  if (colIndex < 0) {
    colIndex = lastCol;
    ws.getRange(1, lastCol + 1).setValue('EmploymentStatus');
    Logger.log('✅ Column "EmploymentStatus" added at column ' + (lastCol + 1));
  } else {
    Logger.log('ℹ️ Already exists at column ' + (colIndex + 1));
  }
  if (lastRow < 2) { Logger.log('ℹ️ No data rows.'); return; }
  var sheetColNum = colIndex + 1;
  var dataRange   = ws.getRange(2, sheetColNum, lastRow - 1, 1);
  var values      = dataRange.getValues();
  var filled      = 0;
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === '' || values[i][0] === null || values[i][0] === undefined) {
      values[i][0] = 'កំពុងធ្វើការ'; filled++;
    }
  }
  dataRange.setValues(values);
  Logger.log('✅ Filled ' + filled + ' rows with "កំពុងធ្វើការ"');
  try {
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['កំពុងធ្វើការ','បានឈប់','បានព្យួរ'], true)
      .setAllowInvalid(false).build();
    ws.getRange(2, sheetColNum, lastRow + 100, 1).setDataValidation(rule);
  } catch(ve) { Logger.log('⚠️ Validation: ' + ve.message); }
  try {
    ws.getRange(1, sheetColNum).setBackground('#d1fae5').setFontColor('#065f46').setFontWeight('bold');
  } catch(he) {}
  SpreadsheetApp.flush();
  Logger.log('🎉 Done!');
}

// ============================================================
// TEST & SETUP FUNCTIONS
// ============================================================
function setupWebhook() {
  UrlFetchApp.fetch('https://api.telegram.org/bot'+TELEGRAM_TOKEN+'/deleteWebhook?drop_pending_updates=true',{method:'post',muteHttpExceptions:true});
  Utilities.sleep(2000);
  var r=UrlFetchApp.fetch('https://api.telegram.org/bot'+TELEGRAM_TOKEN+'/setWebhook',{
    method:'post',contentType:'application/json',muteHttpExceptions:true,
    payload:JSON.stringify({url:WEBHOOK_URL,allowed_updates:['message'],drop_pending_updates:true})
  });
  Logger.log('setWebhook: '+r.getContentText());
}

function testStaffPhotoFolder() {
  Logger.log('=== Staff Photo Folder Test ===');
  try {
    var folder = DriveApp.getFolderById(STAFF_PHOTO_FOLDER);
    Logger.log('Folder: ' + folder.getName());
    var files = folder.getFiles(), count = 0;
    while (files.hasNext() && count < 5) {
      var f = files.next();
      Logger.log('File: ' + f.getName() + ' | ID: ' + f.getId());
      count++;
    }
  } catch(e) { Logger.log('ERROR: ' + e.message); }
}

function verifyDeployment() {
  Logger.log('=== DEPLOYMENT CHECK ===');
  Logger.log('SS_ID: ' + SS_ID);
  var r=UrlFetchApp.fetch('https://api.telegram.org/bot'+TELEGRAM_TOKEN+'/getWebhookInfo',{muteHttpExceptions:true});
  var cur=JSON.parse(r.getContentText()).result;
  Logger.log('Webhook: '+(cur?cur.url:'none'));
  Logger.log(cur&&cur.url===WEBHOOK_URL?'✅ CORRECT!':'❌ Run setupWebhook()');
}

function testSendMessage() {
  sendTelegramMsg(TELEGRAM_CHAT, 'LHB HR v5.3 Test OK!');
}

// ============================================================
// WARMUP — Run this via Time-based Trigger every 5 minutes
// to keep the GAS script warm and avoid cold-start delays (2-5s)
// Setup: Apps Script → Triggers → Add Trigger → warmup → Time-driven → Minutes timer → Every 5 minutes
// ============================================================
function warmup() {
  try {
    getSS(); // open spreadsheet to warm the connection
    // Pre-cache the most-read sheets
    var sheets = ['CheckIn', 'CheckOut', 'StaffInfo', 'Project'];
    sheets.forEach(function(s) {
      var outKey = 'doget_' + s;
      if (!CacheService.getScriptCache().get(outKey)) {
        var ws = getSS().getSheetByName(s);
        if (!ws) return;
        var vals = ws.getDataRange().getValues();
        if (vals.length < 2) return;
        var headers = vals[0].map(function(h){ return String(h).trim(); });
        var rows = [];
        for (var i = 1; i < vals.length; i++) {
          var row = vals[i];
          if (!row.some(function(c){ return c !== '' && c !== null; })) continue;
          var obj = {};
          headers.forEach(function(h, j) {
            var v = row[j];
            if (v instanceof Date) {
              v = (v.getFullYear()===1899 && v.getMonth()===11 && v.getDate()===30) ? phTimeStr(v) : phDateStr(v);
            }
            obj[h] = (v !== undefined && v !== null) ? String(v) : '';
          });
          rows.push(obj);
        }
        try { CacheService.getScriptCache().put(outKey, JSON.stringify({status:'ok',data:rows}), getCacheTTL(s, false)); } catch(e) {}
      }
    });
    Logger.log('[warmup] OK — ' + new Date().toISOString());
  } catch(e) { Logger.log('[warmup] Error: ' + e.message); }
}