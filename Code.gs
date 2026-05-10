// ============================================================
// LHB HR SYSTEM — Google Apps Script v5.1
// ============================================================

const SS_ID              = '16ryjqdieYbZAaG9phRMVInz_Yt6bP8KtWmEYXBcZRH0';
const TELEGRAM_TOKEN     = PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN') || '';
const TELEGRAM_CHAT      = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT')  || '549942306';
const TELEGRAM_GROUP     = PropertiesService.getScriptProperties().getProperty('TELEGRAM_Group') || '';
const WEBHOOK_URL        = 'https://script.google.com/macros/s/AKfycbwReBXqhqr1hXNbmtN7GjxeEBFW--RzgdatCiUQ2PVwbxV5F-20BMQ9cWAIB5W_Nkd2/exec';
const FOOD_FOLDER_ID     = '1Ue7-K0QPDVwQcRszw5xF7b3SH25yGj5y';           // Food/Report photos
const STAFF_PHOTO_FOLDER = '1BMeeqss2J_eoU-o8At7Wri-UNDzMO42DW7XzKeanz2vNgPrzJrICf5IL6OgAn6_ulWbS1B8X'; // Staff ID photos
const FOLDER_ID          = FOOD_FOLDER_ID; // default alias

// ============================================================
// doGet — Read Sheet / Photo
// ============================================================
function doGet(e) {
  try {
    var params = e.parameter || {};
    Logger.log('doGet: ' + JSON.stringify(params));

    // ── Get Staff Photo by Staff ID (search in Staff Photo Folder) ──
    if (params.action === 'getStaffPhoto') {
      try {
        var staffId = String(params.staffId || '').trim().toUpperCase();
        if (!staffId) return respond({ status:'error', msg:'staffId required' });
        var folder   = DriveApp.getFolderById(STAFF_PHOTO_FOLDER);
        var allFiles = folder.getFiles();
        while (allFiles.hasNext()) {
          var f     = allFiles.next();
          var fname = f.getName().toUpperCase().replace(/\.[^.]+$/, ''); // strip extension
          if (fname === staffId || fname.indexOf(staffId) >= 0 || staffId.indexOf(fname) >= 0) {
            f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            return respond({ status:'ok', url:'https://lh3.googleusercontent.com/d/' + f.getId(), fileId:f.getId(), name:f.getName() });
          }
        }
        return respond({ status:'notfound', msg:'No photo for staffId: ' + staffId });
      } catch(pe) { return respond({ status:'error', msg:pe.message }); }
    }

    // ── Get Photo URL by filename ──
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

    // ── Read Sheet ──
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
// SETUP HEADERS
// ============================================================
function setupHeaders() {
  var ss=SpreadsheetApp.openById(SS_ID);
  var HEADERS={
    User:       ['Username','Password','Name','Role','Email','Department','Position'],
    StaffInfo:  ['ID','Name','NameLatin','Sex','LV','Position','Department','ProjectName','DateOfBirth','StartingDate','Salary','Gmail','BankName','BankNumber','Photo','Phone','TelegramChatId','OTP','OTPExpire'],
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
  for(var name in HEADERS){
    var headers=HEADERS[name];
    var ws=ss.getSheetByName(name)||ss.insertSheet(name);
    var existing=ws.getLastColumn()>0?ws.getRange(1,1,1,ws.getLastColumn()).getValues()[0].map(function(h){return String(h).trim();})  :[];
    var added=0;
    headers.forEach(function(h){if(existing.indexOf(h)<0){ws.getRange(1,ws.getLastColumn()+1).setValue(h);existing.push(h);added++;}});
    Logger.log(name+': '+(added===0?'OK ✅':added+' added ✅'));
  }
  Logger.log('setupHeaders() done! ✅');
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
    Logger.log('Folder Access: ' + folder.getSharingAccess());
    var files = folder.getFiles(), count = 0;
    while (files.hasNext() && count < 5) {
      var f = files.next();
      Logger.log('File: ' + f.getName() + ' | ID: ' + f.getId() + ' | lh3: https://lh3.googleusercontent.com/d/' + f.getId());
      count++;
    }
    Logger.log('Total shown: ' + count + ' files');
  } catch(e) { Logger.log('ERROR: ' + e.message); }
}

function verifyDeployment() {
  Logger.log('=== DEPLOYMENT CHECK ===');
  Logger.log('WEBHOOK_URL: ' + WEBHOOK_URL);
  Logger.log('SS_ID: ' + SS_ID);
  Logger.log('FOOD_FOLDER_ID: ' + FOOD_FOLDER_ID);
  Logger.log('STAFF_PHOTO_FOLDER: ' + STAFF_PHOTO_FOLDER);
  var r=UrlFetchApp.fetch('https://api.telegram.org/bot'+TELEGRAM_TOKEN+'/getWebhookInfo',{muteHttpExceptions:true});
  var cur=JSON.parse(r.getContentText()).result;
  Logger.log('Webhook: '+(cur?cur.url:'none'));
  Logger.log(cur&&cur.url===WEBHOOK_URL?'✅ CORRECT!':'❌ Run setupWebhook()');
}

function testSendMessage() {
  sendTelegramMsg(TELEGRAM_CHAT, 'LHB HR v5.1 Test OK!');
}
