// ============================================================
// LHB HR SYSTEM — Google Apps Script v5.2
// ============================================================

const SS_ID              = '16ryjqdieYbZAaG9phRMVInz_Yt6bP8KtWmEYXBcZRH0';
const TELEGRAM_TOKEN     = PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN') || '';
const TELEGRAM_CHAT      = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT')  || '549942306';
const TELEGRAM_GROUP     = PropertiesService.getScriptProperties().getProperty('TELEGRAM_Group') || '';
const WEBHOOK_URL        = 'https://script.google.com/macros/s/AKfycbwReBXqhqr1hXNbmtN7GjxeEBFW--RzgdatCiUQ2PVwbxV5F-20BMQ9cWAIB5W_Nkd2/exec';
const FOOD_FOLDER_ID     = '1Ue7-K0QPDVwQcRszw5xF7b3SH25yGj5y';
const STAFF_PHOTO_FOLDER = '1BMeeqss2J_eoU-o8At7Wri-UNDzMO42DW7XzKeanz2vNgPrzJrICf5IL6OgAn6_ulWbS1B8X';
const FOLDER_ID          = FOOD_FOLDER_ID;

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

    // ── checkSession via GET (no CORS issue) ──
    if (params.action === 'checkSession') {
      var ss2=SpreadsheetApp.openById(SS_ID), ws2=ss2.getSheetByName('User');
      if (!ws2) return respond({ status:'ok', active:false });
      var data2=ws2.getDataRange().getValues(), hdrs2=data2[0].map(function(h){return String(h).trim();});
      var uIdx2=hdrs2.indexOf('Username'), tIdx2=hdrs2.indexOf('SessionToken'), tsIdx2=hdrs2.indexOf('SessionTime');
      if (tIdx2<0||tsIdx2<0) return respond({ status:'ok', active:false });
      var username2=String(params.username||'').trim().toLowerCase();
      var myToken2=String(params.myToken||'').trim();
      var SESSION_TIMEOUT2=8*60*60*1000;
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
      var ss3=SpreadsheetApp.openById(SS_ID), ws3=ss3.getSheetByName('User');
      if (!ws3) return respond({ status:'error', msg:'User sheet not found' });
      var data3=ws3.getDataRange().getValues(), hdrs3=data3[0].map(function(h){return String(h).trim();});
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
          return respond({ status:'ok' });
        }
      }
      return respond({ status:'error', msg:'User not found' });
    }

    // ── clearSession via GET ──
    if (params.action === 'clearSession') {
      var ss4=SpreadsheetApp.openById(SS_ID), ws4=ss4.getSheetByName('User');
      if (!ws4) return respond({ status:'ok' });
      var data4=ws4.getDataRange().getValues(), hdrs4=data4[0].map(function(h){return String(h).trim();});
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
    Logger.log('Reading sheet: ' + sheet);

    var ss = SpreadsheetApp.openById(SS_ID);
    var ws = ss.getSheetByName(sheet);
    if (!ws) {
      var avail = ss.getSheets().map(function(s){ return s.getName(); }).join(', ');
      return respond({ status:'error', msg:'Sheet "' + sheet + '" not found. Available: ' + avail });
    }

    var vals = ws.getDataRange().getValues();
    if (vals.length < 2) return respond({ status:'ok', data:[] });

    var headers = vals[0].map(function(h){ return String(h).trim(); });
    var rows = [];
    for (var i = 1; i < vals.length; i++) {
      var row = vals[i];
      if (!row.some(function(c){ return c !== '' && c !== null; })) continue;
      var obj = {};
      headers.forEach(function(h, j) {
        var v = row[j];
        if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        obj[h] = (v !== undefined && v !== null) ? String(v) : '';
      });
      rows.push(obj);
    }
    Logger.log('Sheet ' + sheet + ': ' + rows.length + ' rows');
    return respond({ status:'ok', data:rows });

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
      var ss = SpreadsheetApp.openById(SS_ID);
      var ws = ss.getSheetByName('StaffInfo');
      var data = ws.getDataRange().getValues();
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
      var ss=SpreadsheetApp.openById(SS_ID), ws=ss.getSheetByName('StaffInfo');
      var data=ws.getDataRange().getValues(), hdrs=data[0].map(function(h){return String(h).trim();});
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
      return respond({ status:'ok', staff:staffObj });
    }

    // ── append ──
    if (p.action === 'append') {
      var ss=SpreadsheetApp.openById(SS_ID), ws=ss.getSheetByName(p.sheet);
      if (!ws) return respond({ status:'error', msg:'Sheet not found: '+p.sheet });
      var headers=ws.getRange(1,1,1,ws.getLastColumn()).getValues()[0].map(function(h){return String(h).trim();});
      var row=p.data||{};
      ws.appendRow(headers.map(function(h){return row[h]!==undefined?row[h]:'';}));
      sendCheckNotification(p.sheet, row);
      return respond({ status:'ok' });
    }

    // ── upsert ──
    if (p.action === 'upsert') {
      var ss=SpreadsheetApp.openById(SS_ID), ws=ss.getSheetByName(p.sheet);
      if (!ws) return respond({ status:'error', msg:'Sheet not found: '+p.sheet });
      var data=ws.getDataRange().getValues(), headers=data[0].map(function(h){return String(h).trim();});
      var idIdx=headers.indexOf('ID'), dateIdx=headers.indexOf('Date');
      for (var i=1;i<data.length;i++){
        if (String(data[i][idIdx]).trim()===String(p.keyValue).trim()&&normDate(data[i][dateIdx])===String(p.keyDate).trim()){
          var row=p.data||{};
          headers.forEach(function(h,j){if(row[h]!==undefined&&row[h]!=='')ws.getRange(i+1,j+1).setValue(row[h]);});
          return respond({ status:'ok', action:'updated' });
        }
      }
      var row=p.data||{};
      ws.appendRow(headers.map(function(h){return row[h]!==undefined?row[h]:'';}));
      return respond({ status:'ok', action:'appended' });
    }

    // ── update ──
    if (p.action === 'update') {
      var ss=SpreadsheetApp.openById(SS_ID), ws=ss.getSheetByName(p.sheet);
      if (!ws) return respond({ status:'error', msg:'Sheet not found: '+p.sheet });
      var data=ws.getDataRange().getValues(), headers=data[0].map(function(h){return String(h).trim();});
      var idIdx=headers.indexOf('ID'), keyVal=String(p.keyValue||p.id||'').trim();
      for (var i=1;i<data.length;i++){
        if (String(data[i][idIdx]).trim()===keyVal){
          var row=p.data||{};
          headers.forEach(function(h,j){if(row[h]!==undefined)ws.getRange(i+1,j+1).setValue(row[h]);});
          return respond({ status:'ok' });
        }
      }
      return respond({ status:'error', msg:'Row not found' });
    }

    // ── delete ──
    if (p.action === 'delete') {
      var ss=SpreadsheetApp.openById(SS_ID), ws=ss.getSheetByName(p.sheet);
      if (!ws) return respond({ status:'error', msg:'Sheet not found: '+p.sheet });
      var data=ws.getDataRange().getValues(), headers=data[0].map(function(h){return String(h).trim();});
      var idIdx=headers.indexOf('ID'), keyVal=String(p.keyValue||p.id||'').trim();
      for (var i=data.length-1;i>=1;i--){
        if (String(data[i][idIdx]).trim()===keyVal){ws.deleteRow(i+1);return respond({ status:'ok' });}
      }
      return respond({ status:'error', msg:'Row not found: '+keyVal });
    }

    // ── deleteByIdDate ──
    if (p.action === 'deleteByIdDate') {
      var ss=SpreadsheetApp.openById(SS_ID), ws=ss.getSheetByName(p.sheet);
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

    // ── uploadPhoto ──
    if (p.action === 'uploadPhoto') {
      var folderId = p.folderId || FOOD_FOLDER_ID;
      var decoded  = Utilities.base64Decode(p.base64);
      var blob     = Utilities.newBlob(decoded, p.mimeType||'image/jpeg', p.fileName||'photo.jpg');
      var folder   = DriveApp.getFolderById(folderId);
      var file     = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      var url = 'https://lh3.googleusercontent.com/d/' + file.getId();
      return respond({ status:'ok', url:url, fileId:file.getId() });
    }

    // ── appendWithPhotos (Food) ──
    if (p.action === 'appendWithPhotos') {
      var ss=SpreadsheetApp.openById(SS_ID);
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

    // ── checkSession: is user already logged in elsewhere? ──
    if (p.action === 'checkSession') {
      var ss=SpreadsheetApp.openById(SS_ID), ws=ss.getSheetByName('User');
      if (!ws) return respond({ status:'ok', active: false });
      var data=ws.getDataRange().getValues(), hdrs=data[0].map(function(h){return String(h).trim();});
      var uIdx=hdrs.indexOf('Username'), tIdx=hdrs.indexOf('SessionToken'), tsIdx=hdrs.indexOf('SessionTime');
      if (tIdx<0||tsIdx<0) return respond({ status:'ok', active: false });
      var username=String(p.username||'').trim().toLowerCase();
      var SESSION_TIMEOUT = 8 * 60 * 60 * 1000; // 8 hours
      for (var i=1;i<data.length;i++) {
        if (String(data[i][uIdx]||'').trim().toLowerCase()===username) {
          var tok=String(data[i][tIdx]||'').trim();
          var ts=Number(data[i][tsIdx]||0);
          var now=new Date().getTime();
          // Active session = has token AND not expired AND not same token as requester
          if (tok && (now-ts)<SESSION_TIMEOUT && tok!==String(p.myToken||'').trim()) {
            return respond({ status:'ok', active: true, since: ts });
          }
          return respond({ status:'ok', active: false });
        }
      }
      return respond({ status:'ok', active: false });
    }

    // ── setSession: register login token ──
    if (p.action === 'setSession') {
      var ss=SpreadsheetApp.openById(SS_ID), ws=ss.getSheetByName('User');
      if (!ws) return respond({ status:'error', msg:'User sheet not found' });
      var data=ws.getDataRange().getValues(), hdrs=data[0].map(function(h){return String(h).trim();});
      var uIdx=hdrs.indexOf('Username'), tIdx=hdrs.indexOf('SessionToken'), tsIdx=hdrs.indexOf('SessionTime');
      // Add columns if missing
      if (tIdx<0) { ws.getRange(1,ws.getLastColumn()+1).setValue('SessionToken'); tIdx=ws.getLastColumn()-1; hdrs.push('SessionToken'); }
      if (tsIdx<0) { ws.getRange(1,ws.getLastColumn()+1).setValue('SessionTime'); tsIdx=ws.getLastColumn()-1; hdrs.push('SessionTime'); }
      var username=String(p.username||'').trim().toLowerCase();
      var token=String(p.token||'').trim();
      var now=new Date().getTime();
      for (var i=1;i<data.length;i++) {
        if (String(data[i][uIdx]||'').trim().toLowerCase()===username) {
          ws.getRange(i+1,tIdx+1).setValue(token);
          ws.getRange(i+1,tsIdx+1).setValue(now);
          return respond({ status:'ok' });
        }
      }
      return respond({ status:'error', msg:'User not found' });
    }

    // ── clearSession: logout, clear token ──
    if (p.action === 'clearSession') {
      var ss=SpreadsheetApp.openById(SS_ID), ws=ss.getSheetByName('User');
      if (!ws) return respond({ status:'ok' });
      var data=ws.getDataRange().getValues(), hdrs=data[0].map(function(h){return String(h).trim();});
      var uIdx=hdrs.indexOf('Username'), tIdx=hdrs.indexOf('SessionToken'), tsIdx=hdrs.indexOf('SessionTime');
      if (tIdx<0) return respond({ status:'ok' });
      var username=String(p.username||'').trim().toLowerCase();
      var token=String(p.token||'').trim();
      for (var i=1;i<data.length;i++) {
        if (String(data[i][uIdx]||'').trim().toLowerCase()===username) {
          // Only clear if token matches (prevent clearing someone else's session)
          var storedTok=String(data[i][tIdx]||'').trim();
          if (!token || storedTok===token) {
            ws.getRange(i+1,tIdx+1).setValue('');
            ws.getRange(i+1,tsIdx+1).setValue('');
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
  var ss=SpreadsheetApp.openById(SS_ID), ws=ss.getSheetByName('StaffInfo');
  var data=ws.getDataRange().getValues(), hdrs=data[0].map(function(h){return String(h).trim();});
  var phoneIdx=hdrs.indexOf('Phone'), nameIdx=hdrs.indexOf('Name'), idIdx=hdrs.indexOf('ID');
  var chatIdx=hdrs.indexOf('TelegramChatId');
  if(chatIdx<0){ws.getRange(1,ws.getLastColumn()+1).setValue('TelegramChatId');chatIdx=ws.getLastColumn()-1;}
  var pn0=phone.replace(/^0+/,'');
  for(var i=1;i<data.length;i++){
    var p=String(data[i][phoneIdx]||'').replace(/[^0-9]/g,'');
    if(p===phone||p.replace(/^0+/,'')===pn0||p.slice(-9)===phone.slice(-9)){
      ws.getRange(i+1,chatIdx+1).setValue(chatId);
      return 'Register OK!\n'+String(data[i][nameIdx])+' ('+String(data[i][idIdx])+')\nPhone: 0'+pn0;
    }
  }
  return 'Phone 0'+pn0+' not found. Contact Admin.';
}

function getRegisteredInfo(chatId) {
  var ss=SpreadsheetApp.openById(SS_ID), ws=ss.getSheetByName('StaffInfo');
  var data=ws.getDataRange().getValues(), hdrs=data[0].map(function(h){return String(h).trim();});
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
    var emoji=sheet==='CheckIn'?'🟢':'🟡', type=sheet==='CheckIn'?'CHECK IN':'CHECK OUT';
    var msg=emoji+' '+type+' | '+(row.Time||'')+'\n\n'+(row.Name||'')+'  '+(row.ID||'')+'\n'+(row.Position||'')+' | '+(row.Department||'')+'\n'+(row.ProjectName||'');
    var target=TELEGRAM_GROUP||TELEGRAM_CHAT;
    if(target)sendTelegramMsg(target,msg);
  }catch(e){Logger.log('Notify: '+e.message);}
}

function normDate(v) {
  if(!v)return '';
  if(v instanceof Date)return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');
  return String(v).trim().slice(0,10);
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// SETUP HEADERS — v5.2 (EmploymentStatus added)
// ============================================================
function setupHeaders() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var HEADERS = {
    User:       ['Username','Password','Name','Role','Email','Department','Position','SessionToken','SessionTime'],
    StaffInfo:  ['ID','Name','NameLatin','Sex','LV','Position','Department','ProjectName',
                 'DateOfBirth','StartingDate','Salary','Gmail','BankName','BankNumber',
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
// ADD EmploymentStatus COLUMN + FILL EXISTING STAFF
// ============================================================
/**
 * addEmploymentStatusColumn()
 *
 * 1. បន្ថែម column "EmploymentStatus" ទៅ StaffInfo (ប្រសិនបើមិនទាន់មាន)
 * 2. Fill existing staff rows ទាំងអស់ → "កំពុងធ្វើការ" (ប្រសិនបើ cell ទទេ)
 *
 * HOW TO RUN:
 *   Apps Script Editor → ជ្រើស function "addEmploymentStatusColumn" → ▶ Run
 */
function addEmploymentStatusColumn() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var ws = ss.getSheetByName('StaffInfo');

  if (!ws) {
    Logger.log('❌ Sheet "StaffInfo" not found!');
    return;
  }

  var lastCol  = ws.getLastColumn();
  var lastRow  = ws.getLastRow();
  var headers  = ws.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h).trim(); });
  var colIndex = headers.indexOf('EmploymentStatus'); // 0-based

  // ── Step 1: Add column header if missing ──
  if (colIndex < 0) {
    colIndex = lastCol; // 0-based index of new column
    ws.getRange(1, lastCol + 1).setValue('EmploymentStatus');
    Logger.log('✅ Column "EmploymentStatus" added at column ' + (lastCol + 1));
  } else {
    Logger.log('ℹ️ Column "EmploymentStatus" already exists at column ' + (colIndex + 1));
  }

  if (lastRow < 2) {
    Logger.log('ℹ️ No data rows to fill.');
    return;
  }

  // ── Step 2: Fill empty cells → "កំពុងធ្វើការ" ──
  var sheetColNum = colIndex + 1; // 1-based for getRange
  var dataRange   = ws.getRange(2, sheetColNum, lastRow - 1, 1);
  var values      = dataRange.getValues();
  var filled      = 0;

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === '' || values[i][0] === null || values[i][0] === undefined) {
      values[i][0] = 'កំពុងធ្វើការ';
      filled++;
    }
  }

  dataRange.setValues(values);
  Logger.log('✅ Filled ' + filled + ' rows with "កំពុងធ្វើការ"');
  Logger.log('✅ addEmploymentStatusColumn() DONE! Total staff rows: ' + (lastRow - 1));

  // ── Step 3: Apply dropdown validation ──
  try {
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['កំពុងធ្វើការ', 'បានឈប់', 'បានព្យួរ'], true)
      .setAllowInvalid(false)
      .build();
    ws.getRange(2, sheetColNum, lastRow + 100, 1).setDataValidation(rule);
    Logger.log('✅ Dropdown validation applied (rows 2 to ' + (lastRow + 100) + ')');
  } catch(ve) {
    Logger.log('⚠️ Validation skipped: ' + ve.message);
  }

  // ── Step 4: Highlight the column header ──
  try {
    ws.getRange(1, sheetColNum).setBackground('#d1fae5').setFontColor('#065f46').setFontWeight('bold');
    Logger.log('✅ Header highlighted green');
  } catch(he) {
    Logger.log('⚠️ Highlight skipped: ' + he.message);
  }

  SpreadsheetApp.flush();
  Logger.log('🎉 All done! ចុច Refresh hr-system.html ដើម្បីឃើញ EmploymentStatus');
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
  Logger.log('Folder ID: ' + STAFF_PHOTO_FOLDER);
  try {
    var folder = DriveApp.getFolderById(STAFF_PHOTO_FOLDER);
    Logger.log('Folder Name: ' + folder.getName());
    var files = folder.getFiles(), count = 0;
    while (files.hasNext() && count < 5) {
      var f = files.next();
      Logger.log('File: ' + f.getName() + ' | lh3: https://lh3.googleusercontent.com/d/' + f.getId());
      count++;
    }
    Logger.log('Total shown: ' + count + ' files');
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
  sendTelegramMsg(TELEGRAM_CHAT, 'LHB HR v5.2 Test OK!');
}
