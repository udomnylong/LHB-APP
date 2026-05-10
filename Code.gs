// ============================================================
// LHB HR SYSTEM — Google Apps Script
// Version: 4.0 — StaffInfo-only OTP (no OTP_Sessions tab)
// ============================================================

const SS_ID         = '16ryjqdieYbZAaG9phRMVInz_Yt6bP8KtWmEYXBcZRH0';
const TELEGRAM_TOKEN  = PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN') || '';
const TELEGRAM_CHAT   = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT')  || '549942306';   // Admin personal
const TELEGRAM_GROUP  = PropertiesService.getScriptProperties().getProperty('TELEGRAM_Group') || '';           // Group chat
const WEBHOOK_URL    = 'https://script.google.com/macros/s/AKfycbwReBXqhqr1hXNbmtN7GjxeEBFW--RzgdatCiUQ2PVwbxV5F-20BMQ9cWAIB5W_Nkd2/exec';
const FOLDER_ID     = '1Ue7-K0QPDVwQcRszw5xF7b3SH25yGj5y';

// ============================================================
// doGet — Read Sheet Data
// ============================================================
function doGet(e) {
  try {
    const params = e.parameter || {};

    if (params.action === 'getPhotoUrl') {
      try {
        const folder = DriveApp.getFolderById(FOLDER_ID);
        const files  = folder.getFilesByName(params.fileName);
        if (files.hasNext()) {
          const file     = files.next();
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          const fileId   = file.getId();
          const thumbUrl = 'https://lh3.googleusercontent.com/d/' + fileId;
          if (params.staffId && params.date && params.mealType) {
            try {
              const ss2 = SpreadsheetApp.openById(SS_ID);
              const fw  = ss2.getSheetByName('Food');
              if (fw && fw.getLastRow() > 1) {
                const dat  = fw.getDataRange().getValues();
                const hdr  = dat[0].map(h => String(h).trim());
                const iIdx = hdr.indexOf('ID'), dIdx = hdr.indexOf('Date');
                const mn   = {m:'morning',l:'lunch',e:'evening',morning:'morning',lunch:'lunch',evening:'evening'};
                const cm   = {morning:'PhotoMorning',lunch:'PhotoLunch',evening:'PhotoEvening'};
                const cNm  = cm[mn[params.mealType]||params.mealType];
                const cIdx = cNm ? hdr.indexOf(cNm) : -1;
                if (cIdx >= 0) {
                  for (let i=1; i<dat.length; i++) {
                    if (String(dat[i][iIdx]).trim()===String(params.staffId).trim() &&
                        normDate(dat[i][dIdx])===String(params.date).trim()) {
                      fw.getRange(i+1,cIdx+1).setValue(thumbUrl); break;
                    }
                  }
                }
              }
            } catch(se){ Logger.log('Sheet err:'+se.message); }
          }
          return respond({ status:'ok', url:thumbUrl, fileId:fileId });
        }
        return respond({ status:'notfound' });
      } catch(de) { return respond({ status:'error', msg:de.message }); }
    }

    // Read sheet
    const sheetRaw = String(params.sheet || '').trim();
    const sheet = (sheetRaw && sheetRaw !== 'undefined' && sheetRaw !== 'null') ? sheetRaw : 'StaffInfo';
    Logger.log('doGet sheet=' + sheet);
    const ss = SpreadsheetApp.openById(SS_ID);
    const ws = ss.getSheetByName(sheet);
    if (!ws) return respond({ status:'error', msg:'Sheet "' + sheet + '" not found' });
    const vals = ws.getDataRange().getValues();
    if (vals.length < 2) return respond({ status:'ok', data:[] });
    const headers = vals[0].map(h => String(h).trim());
    const rows = vals.slice(1)
      .filter(r => r.some(c => c !== '' && c !== null))
      .map(r => {
        const obj = {};
        headers.forEach((h,i) => {
          let v = r[i];
          if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          obj[h] = v !== undefined && v !== null ? String(v) : '';
        });
        return obj;
      });
    return respond({ status:'ok', data:rows });
  } catch(err) { return respond({ status:'error', msg:err.message }); }
}

// ============================================================
// doPost — Write / Telegram Updates
// ============================================================
function doPost(e) {
  try {
    const raw = e.postData ? e.postData.contents : '';
    Logger.log('doPost: ' + (raw ? raw.slice(0,200) : 'EMPTY'));
    if (!raw || raw.length === 0) {
      return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
    }

    let parsed;
    try { parsed = JSON.parse(raw); } catch(ex) { parsed = null; }

    // Telegram update
    if (parsed && parsed.update_id) {
      handleTelegramUpdate(parsed);
      return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
    }

    let p = parsed;
    if (!p) {
      try { p = JSON.parse(decodeURIComponent(raw.replace(/^payload=/, ''))); } catch(ex2) {
        return respond({ status:'error', msg:'Invalid JSON' });
      }
    }
    Logger.log('action=' + p.action);

    // ── sendOTP ──
    if (p.action === 'sendOTP') {
      try {
        const phone = String(p.phone||'').replace(/[^0-9]/g,'');
        if (!phone || phone.length < 8) return respond({ status:'error', msg:'Phone invalid' });

        const ss   = SpreadsheetApp.openById(SS_ID);
        const ws   = ss.getSheetByName('StaffInfo');
        const data = ws.getDataRange().getValues();
        const hdrs = data[0].map(h => String(h).trim());

        const phoneIdx  = hdrs.indexOf('Phone');
        const chatIdx   = hdrs.indexOf('TelegramChatId');
        const otpIdx    = hdrs.indexOf('OTP');
        const expireIdx = hdrs.indexOf('OTPExpire');

        // Ensure columns exist
        let colsAdded = false;
        if (chatIdx < 0)   { ws.getRange(1, ws.getLastColumn()+1).setValue('TelegramChatId'); colsAdded = true; }
        if (otpIdx < 0)    { ws.getRange(1, ws.getLastColumn()+1).setValue('OTP');            colsAdded = true; }
        if (expireIdx < 0) { ws.getRange(1, ws.getLastColumn()+1).setValue('OTPExpire');       colsAdded = true; }

        // Re-read headers if columns were added
        const hdrs2      = colsAdded ? ws.getRange(1,1,1,ws.getLastColumn()).getValues()[0].map(h=>String(h).trim()) : hdrs;
        const chatIdx2   = hdrs2.indexOf('TelegramChatId');
        const otpIdx2    = hdrs2.indexOf('OTP');
        const expireIdx2 = hdrs2.indexOf('OTPExpire');

        const phoneNo0 = phone.replace(/^0+/,'');
        let staffRow = null, rowNum = -1;
        for (let i=1; i<data.length; i++) {
          const rp = String(data[i][phoneIdx]||'').replace(/[^0-9]/g,'');
          if (rp === phone || rp.replace(/^0+/,'') === phoneNo0 || rp.slice(-9) === phone.slice(-9)) {
            staffRow = data[i]; rowNum = i+1; break;
          }
        }
        if (!staffRow) return respond({ status:'error', msg:'Phone ' + phone + ' not in StaffInfo' });

        const chatId = String(staffRow[chatIdx >= 0 ? chatIdx : chatIdx2]||'').trim();
        if (!chatId) return respond({ status:'error', msg:'Phone not registered. Send /start to @lhb_system_bot' });

        const otp    = String(Math.floor(100000 + Math.random()*900000));
        const expire = new Date().getTime() + 5*60*1000;
        ws.getRange(rowNum, otpIdx2+1).setValue(otp);
        ws.getRange(rowNum, expireIdx2+1).setValue(expire);

        const msg = 'LHB HR OTP\n\n' + String(staffRow[hdrs2.indexOf('Name')]||staffRow[hdrs.indexOf('Name')]||'') + '\nCode: ' + otp + '\n\nExpire 5min';
        UrlFetchApp.fetch('https://api.telegram.org/bot'+TELEGRAM_TOKEN+'/sendMessage', {
          method:'post', contentType:'application/json', muteHttpExceptions:true,
          payload: JSON.stringify({ chat_id:chatId, text:msg })
        });
        return respond({ status:'ok' });
      } catch(err) { return respond({ status:'error', msg:'sendOTP: '+err.message }); }
    }

    // ── verifyOTP ──
    if (p.action === 'verifyOTP') {
      try {
        const phone = String(p.phone||'').replace(/[^0-9]/g,'');
        const code  = String(p.code||'').trim();
        const ss    = SpreadsheetApp.openById(SS_ID);
        const ws    = ss.getSheetByName('StaffInfo');
        const data  = ws.getDataRange().getValues();
        const hdrs  = data[0].map(h => String(h).trim());

        const phoneIdx  = hdrs.indexOf('Phone');
        const otpIdx    = hdrs.indexOf('OTP');
        const expireIdx = hdrs.indexOf('OTPExpire');
        if (otpIdx < 0) return respond({ status:'error', msg:'OTP column not found. Run /register first.' });

        const phoneNo0 = phone.replace(/^0+/,'');
        let staffRow = null, rowNum = -1;
        for (let i=1; i<data.length; i++) {
          const rp = String(data[i][phoneIdx]||'').replace(/[^0-9]/g,'');
          if (rp === phone || rp.replace(/^0+/,'') === phoneNo0 || rp.slice(-9) === phone.slice(-9)) {
            staffRow = data[i]; rowNum = i+1; break;
          }
        }
        if (!staffRow) return respond({ status:'error', msg:'Phone not found' });

        const savedOtp    = String(staffRow[otpIdx]   || '').trim();
        const savedExpire = Number(staffRow[expireIdx] || 0);
        Logger.log('verifyOTP: input='+code+' saved='+savedOtp+' expire='+savedExpire);

        if (!savedOtp)                              return respond({ status:'error', msg:'OTP not requested. Tap "ទទួល OTP" first.' });
        if (new Date().getTime() > savedExpire)     return respond({ status:'error', msg:'OTP Expired. Request again.' });
        if (savedOtp !== code)                      return respond({ status:'error', msg:'OTP incorrect.' });

        let staffObj = {};
        hdrs.forEach((h,j) => { staffObj[h] = staffRow[j] !== undefined ? String(staffRow[j]) : ''; });
        ws.getRange(rowNum, otpIdx+1).setValue('');
        ws.getRange(rowNum, expireIdx+1).setValue('');
        return respond({ status:'ok', staff:staffObj });
      } catch(err) { return respond({ status:'error', msg:'verifyOTP: '+err.message }); }
    }

    // ── uploadPhoto ──
    if (p.action === 'uploadPhoto') {
      try {
        const decoded  = Utilities.base64Decode(p.base64);
        const blob     = Utilities.newBlob(decoded, p.mimeType||'image/jpeg', p.fileName);
        const folder   = DriveApp.getFolderById(p.folderId||FOLDER_ID);
        const file     = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const fileId   = file.getId();
        const thumbUrl = 'https://lh3.googleusercontent.com/d/' + fileId;
        Logger.log('Photo uploaded: ' + p.fileName + ' → ' + thumbUrl);
        return respond({ status:'ok', url:thumbUrl, fileId:fileId });
      } catch(err) { return respond({ status:'error', msg:'Upload: '+err.message }); }
    }

    // ── appendWithPhotos (Food) ──
    if (p.action === 'appendWithPhotos') {
      try {
        const ss     = SpreadsheetApp.openById(SS_ID);
        const folder = DriveApp.getFolderById(p.folderId||FOLDER_ID);
        const row    = p.data || {};
        const photoFields = {PhotoMorning:'photoMorning', PhotoLunch:'photoLunch', PhotoEvening:'photoEvening'};
        for (const [sheetCol, dataKey] of Object.entries(photoFields)) {
          const b64 = p.photos ? p.photos[dataKey] : null;
          if (b64 && b64.length > 100) {
            try {
              const decoded  = Utilities.base64Decode(b64);
              const blob     = Utilities.newBlob(decoded, 'image/jpeg', p.fileNames[dataKey]||sheetCol+'.jpg');
              const file     = folder.createFile(blob);
              file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
              row[sheetCol]  = 'https://lh3.googleusercontent.com/d/' + file.getId();
            } catch(pe) { Logger.log('Photo err ('+sheetCol+'): '+pe.message); }
          }
        }
        const ws      = ss.getSheetByName('Food');
        if (!ws) return respond({ status:'error', msg:'Food sheet not found' });
        const headers = ws.getRange(1,1,1,ws.getLastColumn()).getValues()[0].map(h=>String(h).trim());
        ws.appendRow(headers.map(h => row[h] !== undefined ? row[h] : ''));
        return respond({ status:'ok', photoMorning:row.PhotoMorning||'', photoLunch:row.PhotoLunch||'', photoEvening:row.PhotoEvening||'' });
      } catch(err) { return respond({ status:'error', msg:'appendWithPhotos: '+err.message }); }
    }

    // ── append ──
    if (p.action === 'append') {
      const ss  = SpreadsheetApp.openById(SS_ID);
      const ws  = ss.getSheetByName(p.sheet);
      if (!ws) return respond({ status:'error', msg:'Sheet not found: '+p.sheet });
      const headers = ws.getRange(1,1,1,ws.getLastColumn()).getValues()[0].map(h=>String(h).trim());
      const row     = p.data || {};
      ws.appendRow(headers.map(h => row[h] !== undefined ? row[h] : ''));
      sendCheckNotification(p.sheet, row);
      return respond({ status:'ok' });
    }

    // ── upsert (Attendance) ──
    if (p.action === 'upsert') {
      const ss  = SpreadsheetApp.openById(SS_ID);
      const ws  = ss.getSheetByName(p.sheet);
      if (!ws) return respond({ status:'error', msg:'Sheet not found: '+p.sheet });
      const data    = ws.getDataRange().getValues();
      const headers = data[0].map(h=>String(h).trim());
      const idIdx   = headers.indexOf('ID');
      const dateIdx = headers.indexOf('Date');
      for (let i=1; i<data.length; i++) {
        if (String(data[i][idIdx]).trim()===String(p.keyValue).trim() && normDate(data[i][dateIdx])===String(p.keyDate).trim()) {
          const row = p.data || {};
          headers.forEach((h,j) => { if (row[h] !== undefined && row[h] !== '') ws.getRange(i+1,j+1).setValue(row[h]); });
          return respond({ status:'ok', action:'updated' });
        }
      }
      const row = p.data || {};
      ws.appendRow(headers.map(h => row[h] !== undefined ? row[h] : ''));
      return respond({ status:'ok', action:'appended' });
    }

    // ── update ──
    if (p.action === 'update') {
      const ss  = SpreadsheetApp.openById(SS_ID);
      const ws  = ss.getSheetByName(p.sheet);
      if (!ws) return respond({ status:'error', msg:'Sheet not found' });
      const data    = ws.getDataRange().getValues();
      const headers = data[0].map(h=>String(h).trim());
      const idIdx   = headers.indexOf('ID');
      for (let i=1; i<data.length; i++) {
        if (String(data[i][idIdx]).trim()===String(p.id).trim()) {
          const row = p.data || {};
          headers.forEach((h,j) => { if (row[h] !== undefined) ws.getRange(i+1,j+1).setValue(row[h]); });
          return respond({ status:'ok' });
        }
      }
      return respond({ status:'error', msg:'Row not found' });
    }

    // ── delete ──
    if (p.action === 'delete') {
      const ss  = SpreadsheetApp.openById(SS_ID);
      const ws  = ss.getSheetByName(p.sheet);
      if (!ws) return respond({ status:'error', msg:'Sheet not found' });
      const data    = ws.getDataRange().getValues();
      const headers = data[0].map(h=>String(h).trim());
      const idIdx   = headers.indexOf('ID');
      for (let i=data.length-1; i>=1; i--) {
        if (String(data[i][idIdx]).trim()===String(p.id).trim()) {
          ws.deleteRow(i+1);
          return respond({ status:'ok' });
        }
      }
      return respond({ status:'error', msg:'Row not found' });
    }

    // ── setupHeaders ──
    if (p.action === 'setupHeaders') {
      const ss = SpreadsheetApp.openById(SS_ID);
      const SHEET_HEADERS = {
        User:       ['Username','Password','Name','Role','Email','Department','Position'],
        StaffInfo:  ['ID','Name','Sex','LV','Position','Department','ProjectName','DateOfBirth','StartingDate','Salary','Gmail','BankName','BankNumber','Photo','Phone','TelegramChatId','OTP','OTPExpire'],
        Attendance: ['ID','Name','Position','Department','ProjectName','CheckIn','CheckOut','Late','Early','Status','Date'],
        StaffLeave: ['ID','Name','TypeOfLeave','StartDate','EndDate','Days','Reason','Status'],
        Project:    ['ProjectID','ProjectName','Location','Latitude','Longitude','Radius','Status'],
        StaffOT:    ['ID','Name','Date','Hours','TimeFrom','TimeTo','TypeOfWork','Reason','Status'],
        CheckIn:    ['ID','Name','Gmail','ProjectName','Date','Time','Timestamp','Latitude','Longitude','Accuracy','LateEarly','Minutes','Position','Department'],
        CheckOut:   ['ID','Name','Gmail','ProjectName','Date','Time','Timestamp','Latitude','Longitude','Accuracy','LateEarly','Minutes','Position','Department'],
        Food:       ['Date','ID','Name','Sex','Position','ProjectName','Morning','Lunch','Evening','Total','UnitPrice','TotalPrice','PhotoMorning','PhotoLunch','PhotoEvening','Comment','Remark'],
    WorkPlace:  ['Date','Time','ID','Name','Department','ProjectName','Comment','Photo','Status'],
    Comment:    ['Date','Time','ID','Name','Department','ProjectName','Comment','Photo','Status'],
      };
      for (const [name, headers] of Object.entries(SHEET_HEADERS)) {
        let ws = ss.getSheetByName(name);
        if (!ws) ws = ss.insertSheet(name);
        ws.getRange(1,1,1,headers.length).setValues([headers]);
      }
      return respond({ status:'ok', msg:'Headers set up' });
    }

    return respond({ status:'error', msg:'Unknown action: '+p.action });
  } catch(err) {
    Logger.log('doPost error: '+err.message);
    return respond({ status:'error', msg:err.message });
  }
}

// ============================================================
// TELEGRAM BOT — Webhook Handler
// ============================================================
function handleTelegramUpdate(body) {
  try {
    const updateId = String(body.update_id||'');

    // Dedup 1: CacheService (fast, in-memory)
    if (updateId) {
      const cache = CacheService.getScriptCache();
      if (cache.get('tg_upd_'+updateId)) {
        Logger.log('Dup (cache) update_id='+updateId+' — skipped');
        return;
      }
      cache.put('tg_upd_'+updateId, '1', 21600); // 6 hours
    }

    const msg = body.message || body.edited_message;
    if (!msg) return;

    // Dedup 2: Message date — ignore messages older than 30 seconds
    const msgDate = msg.date || 0; // Unix timestamp
    const nowSec  = Math.floor(new Date().getTime() / 1000);
    const ageSec  = nowSec - msgDate;
    Logger.log('Message age: ' + ageSec + 's (update_id='+updateId+')');
    if (ageSec > 30) {
      Logger.log('Old message ('+ageSec+'s ago) — skipped to prevent replay');
      return;
    }

    const chatId = String(msg.chat.id);
    const text   = (msg.text||'').trim();
    Logger.log('TG msg chatId='+chatId+': '+text);

    if (text === '/start') {
      sendTelegramReply(chatId, 'LHB HR Bot!\n\nRegister phone:\n/register 0XXXXXXXXX\n\nExample:\n/register 0968099996');
      return;
    }
    if (text.startsWith('/register')) {
      const phoneRaw = text.replace('/register','').trim();
      const phone    = phoneRaw.replace(/[^0-9]/g,'');
      if (!phone || phone.length < 8) {
        sendTelegramReply(chatId, 'Format: /register 0XXXXXXXXX\nExample: /register 0968099996');
        return;
      }
      sendTelegramReply(chatId, registerStaff(phone, chatId));
      return;
    }
    if (text === '/status') {
      const info = getRegisteredInfo(chatId);
      sendTelegramReply(chatId, info || 'Not registered. Send /register 0XXXXXXXXX');
    }
  } catch(err) { Logger.log('handleTelegramUpdate error: '+err.message); }
}

function registerStaff(phone, chatId) {
  // Save TelegramChatId to StaffInfo
  const ss      = SpreadsheetApp.openById(SS_ID);
  const ws      = ss.getSheetByName('StaffInfo');
  const data    = ws.getDataRange().getValues();
  const hdrs    = data[0].map(h => String(h).trim());
  const phoneIdx = hdrs.indexOf('Phone');
  const nameIdx  = hdrs.indexOf('Name');
  const idIdx    = hdrs.indexOf('ID');
  var chatColIdx = hdrs.indexOf('TelegramChatId');

  if (chatColIdx < 0) {
    ws.getRange(1, ws.getLastColumn()+1).setValue('TelegramChatId');
    chatColIdx = ws.getLastColumn() - 1;
  }

  const phoneNo0 = phone.replace(/^0+/,'');
  for (var i = 1; i < data.length; i++) {
    const raw  = String(data[i][phoneIdx]||'');
    const p    = raw.replace(/[^0-9]/g,'');
    const pNo0 = p.replace(/^0+/,'');
    if (p === phone || pNo0 === phoneNo0 || p.slice(-9) === phone.slice(-9)) {
      const name = String(data[i][nameIdx]||'');
      const id   = String(data[i][idIdx]||'');
      ws.getRange(i+1, chatColIdx+1).setValue(chatId);
      Logger.log('Registered: '+name+' ('+id+') chatId='+chatId);
      return 'Register OK!\n' + name + ' (' + id + ')\nPhone: 0' + phoneNo0 + '\nYou can Check In/Out now!';
    }
  }
  return 'Phone 0' + phoneNo0 + ' not found in HR System.\nContact Admin to add your Phone number.';
}

function getRegisteredInfo(chatId) {
  const ss   = SpreadsheetApp.openById(SS_ID);
  const ws   = ss.getSheetByName('StaffInfo');
  const data = ws.getDataRange().getValues();
  const hdrs = data[0].map(h=>String(h).trim());
  const chatIdx = hdrs.indexOf('TelegramChatId');
  const nameIdx = hdrs.indexOf('Name');
  const idIdx   = hdrs.indexOf('ID');
  if (chatIdx < 0) return null;
  for (var i=1; i<data.length; i++) {
    if (String(data[i][chatIdx]).trim() === chatId) {
      return 'Registered: ' + data[i][nameIdx] + ' (' + data[i][idIdx] + ')';
    }
  }
  return null;
}

function sendTelegramReply(chatId, text) {
  UrlFetchApp.fetch('https://api.telegram.org/bot'+TELEGRAM_TOKEN+'/sendMessage', {
    method:'post', contentType:'application/json', muteHttpExceptions:true,
    payload: JSON.stringify({ chat_id:chatId, text:text })
  });
}

// ============================================================
// TELEGRAM NOTIFICATIONS — Check In/Out
// ============================================================
function sendCheckNotification(sheet, row) {
  try {
    if (sheet !== 'CheckIn' && sheet !== 'CheckOut') return;
    const type  = sheet === 'CheckIn' ? 'CHECK IN' : 'CHECK OUT';
    const emoji = sheet === 'CheckIn' ? '🟢' : '🟡';
    const time  = row.Time || '';
    const late  = row.LateEarly || '';
    const name  = row.Name || '';
    const id    = row.ID   || '';
    const proj  = row.ProjectName || '';
    const dept  = row.Department  || '';
    const gmail = row.Gmail || '';
    const msg   = emoji + ' ' + type + ' | ' + time + '\n\n' +
      name + '  ' + id + '\n' +
      (row.Position||'') + ' | ' + dept + '\n' +
      proj + '\n' +
      (late ? late + '\n' : '') +
      (gmail ? gmail + '\n' : '');
    // Send to Group (Check In/Out notifications)
    const targetChat = TELEGRAM_GROUP || TELEGRAM_CHAT;
    UrlFetchApp.fetch('https://api.telegram.org/bot'+TELEGRAM_TOKEN+'/sendMessage', {
      method:'post', contentType:'application/json', muteHttpExceptions:true,
      payload: JSON.stringify({ chat_id:targetChat, text:msg })
    });
  } catch(e) { Logger.log('Notification error: '+e.message); }
}

function sendMorningSummary() { sendSummary('morning'); }
function sendEveningSummary() { sendSummary('evening'); }

function sendSummary(time) {
  try {
    const ss      = SpreadsheetApp.openById(SS_ID);
    const today   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const ciWs    = ss.getSheetByName('CheckIn');
    const staffWs = ss.getSheetByName('StaffInfo');
    if (!ciWs || !staffWs) return;

    const ciData  = ciWs.getDataRange().getValues();
    const ciHdrs  = ciData[0].map(h=>String(h).trim());
    const dateIdx = ciHdrs.indexOf('Date');
    const idIdx   = ciHdrs.indexOf('ID');

    const checkedIn = new Set();
    for (let i=1; i<ciData.length; i++) {
      if (normDate(ciData[i][dateIdx]) === today) checkedIn.add(String(ciData[i][idIdx]).trim());
    }

    const stData  = staffWs.getDataRange().getValues();
    const stHdrs  = stData[0].map(h=>String(h).trim());
    const totalStaff = stData.length - 1;
    const late = [];
    stData.slice(1).forEach(r => {
      const id = String(r[stHdrs.indexOf('ID')]||'').trim();
      if (!checkedIn.has(id)) late.push(String(r[stHdrs.indexOf('Name')]||''));
    });

    const emoji = time === 'morning' ? '🌅' : '🌙';
    const msg   = emoji + ' Summary ' + today + '\n\n' +
      'Total: ' + totalStaff + '\n' +
      'Check In: ' + checkedIn.size + '\n' +
      'Absent: ' + late.length + '\n\n' +
      (late.length > 0 ? 'Absent:\n' + late.slice(0,10).join('\n') : '');
    const groupChat = TELEGRAM_GROUP || TELEGRAM_CHAT;
    UrlFetchApp.fetch('https://api.telegram.org/bot'+TELEGRAM_TOKEN+'/sendMessage', {
      method:'post', contentType:'application/json', muteHttpExceptions:true,
      payload: JSON.stringify({ chat_id:groupChat, text:msg })
    });
  } catch(e) { Logger.log('Summary error: '+e.message); }
}

// ============================================================
// UTILITY FUNCTIONS
// ============================================================
function normDate(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const s = String(v).trim();
  if (s.match(/^\d{4}-\d{2}-\d{2}/)) return s.slice(0,10);
  return s;
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// WEBHOOK SETUP & TEST FUNCTIONS
// ============================================================
function setupWebhook() {
  // Step 1: Delete webhook + drop ALL pending
  const del = UrlFetchApp.fetch(
    'https://api.telegram.org/bot'+TELEGRAM_TOKEN+'/deleteWebhook?drop_pending_updates=true',
    {method:'post', muteHttpExceptions:true}
  );
  Logger.log('deleteWebhook: ' + del.getContentText());
  Utilities.sleep(2000);

  // Step 2: Drain any remaining updates via getUpdates
  const drain = UrlFetchApp.fetch(
    'https://api.telegram.org/bot'+TELEGRAM_TOKEN+'/getUpdates?offset=-1&limit=1',
    {muteHttpExceptions:true}
  );
  const dj = JSON.parse(drain.getContentText());
  if (dj.ok && dj.result && dj.result.length > 0) {
    const lastId = dj.result[dj.result.length-1].update_id;
    UrlFetchApp.fetch(
      'https://api.telegram.org/bot'+TELEGRAM_TOKEN+'/getUpdates?offset='+(lastId+1),
      {muteHttpExceptions:true}
    );
    Logger.log('Drained updates up to: ' + lastId);
  }
  Utilities.sleep(1000);

  // Step 3: Set new webhook
  const r = UrlFetchApp.fetch('https://api.telegram.org/bot'+TELEGRAM_TOKEN+'/setWebhook', {
    method:'post', contentType:'application/json', muteHttpExceptions:true,
    payload: JSON.stringify({
      url: WEBHOOK_URL,
      allowed_updates: ['message'],
      drop_pending_updates: true,
      max_connections: 1
    })
  });
  Logger.log('setWebhook: ' + r.getContentText());
}

function checkWebhook() {
  const r = UrlFetchApp.fetch('https://api.telegram.org/bot'+TELEGRAM_TOKEN+'/getWebhookInfo', {muteHttpExceptions:true});
  Logger.log(r.getContentText());
}

function testBotConnection() {
  const r = UrlFetchApp.fetch('https://api.telegram.org/bot'+TELEGRAM_TOKEN+'/getMe', {muteHttpExceptions:true});
  Logger.log('Bot: ' + r.getContentText());
}

function verifyCorrectDeployment() {
  // Verify this is the correct (new) deployment
  Logger.log('=== DEPLOYMENT VERIFICATION ===');
  Logger.log('WEBHOOK_URL: ' + WEBHOOK_URL);
  Logger.log('Code version: 4.0 (StaffInfo-only OTP)');
  Logger.log('OTP_Sessions: REMOVED');
  Logger.log('registerStaff() saves to: StaffInfo.TelegramChatId');

  // Check current webhook
  const r = UrlFetchApp.fetch('https://api.telegram.org/bot'+TELEGRAM_TOKEN+'/getWebhookInfo', {muteHttpExceptions:true});
  const info = JSON.parse(r.getContentText());
  const webhookUrl = info.result ? info.result.url : 'none';
  Logger.log('Active webhook URL: ' + webhookUrl);

  if (webhookUrl === WEBHOOK_URL) {
    Logger.log('✅ Webhook points to THIS deployment — CORRECT!');
  } else {
    Logger.log('❌ Webhook points to: ' + webhookUrl);
    Logger.log('   Expected:          ' + WEBHOOK_URL);
    Logger.log('   FIX: Run setupWebhook() to update!');
  }
  Logger.log('================================');
}

function testSendMessage() {
  sendTelegramReply(TELEGRAM_CHAT, 'LHB HR Bot Test OK! (Admin)');
  if (TELEGRAM_GROUP) sendTelegramReply(TELEGRAM_GROUP, 'LHB HR Bot Test OK! (Group)');
  Logger.log('Admin: '+TELEGRAM_CHAT+' | Group: '+TELEGRAM_GROUP);
}

function testPhoneMatch() {
  const phone = '0968099996';
  const ss    = SpreadsheetApp.openById(SS_ID);
  const ws    = ss.getSheetByName('StaffInfo');
  const data  = ws.getDataRange().getValues();
  const hdrs  = data[0].map(h=>String(h).trim());
  Logger.log('Headers: ' + hdrs.join(', '));
  Logger.log('TelegramChatId col: ' + hdrs.indexOf('TelegramChatId'));
  const phoneClean = phone.replace(/[^0-9]/g,'');
  const phoneNo0   = phoneClean.replace(/^0+/,'');
  for (var i=1; i<data.length; i++) {
    const p = String(data[i][hdrs.indexOf('Phone')]||'').replace(/[^0-9]/g,'');
    if (p && (p===phoneClean || p.replace(/^0+/,'')===phoneNo0)) {
      Logger.log('MATCH row '+(i+1)+': '+data[i][hdrs.indexOf('Name')]);
      return;
    }
  }
  Logger.log('No match for '+phone);
}

function doWebhook(e) {
  try { handleTelegramUpdate(JSON.parse(e.postData.contents)); } catch(err) { Logger.log('doWebhook:'+err.message); }
}
