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
  row.Date   = "'" + sDate;   // leading "'" forces plain text — see enforceServerTime_
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
    row.CheckIn = "'" + sTime;
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
    row.CheckOut = "'" + sTime;
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
  // Leading "'" forces Sheets to store these as plain text instead of
  // auto-converting to Date/Time cells — Apps Script reconstructs auto-converted
  // Date cells using the project's configured timezone, which can silently drift
  // from the fixed UTC+7 math above and make times read back wrong (see doGet).
  row['Time']      = "'" + sTime;
  row['Date']      = "'" + sDate;
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
    // (strip the leading "'" that forces the Sheets cell to stay plain text)
    var serverTime = String(row['Time'] || '').replace(/^'/, '').substring(0, 5) || phTimeStr(new Date());
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
    // NOTE: every sheet except StaffInfo is excluded — Attendance/CheckIn/CheckOut/
    // Project/StaffOT/StaffLeave/Food/Comment/Workplace/EvaluateStaff/User all read
    // and write through Cloud Run/Postgres now (see assets/api-client.js), so warming
    // their GAS doGet cache is dead work. StaffInfo is kept: it's still read via GAS
    // doGet directly by staff-portal.html, food-scan.html, and the embedded QR-scan
    // kiosk page hr-system.html generates (buildScanPageHTML) — all public/no-login
    // surfaces that fall back to (or primarily use) the Sheet for staff lookups.
    var sheets = ['StaffInfo'];
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