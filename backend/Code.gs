// 在 ryan25620280@gmail.com 的 Google Apps Script 專案部署。
const OWNER_EMAIL = 'ryan25620280@gmail.com';
const TARGET_SHEET = '1GTKYhBMuO0JK9khsKR-hMzDMF_fgjkslr1c-lB2E8IU';
const LINE_CHANNEL_ID = '2011216294';
const RESPONSE_HEADERS = ['收件時間','收件ID','內容雜湊','姓名','民國生日','手機','預約日期','時段','拿藥需求','LINE稱呼','備註','處理狀態','LINE ID'];
const USER_HEADERS = ['LINE ID','LINE稱呼','首次登入時間','最後登入時間'];
const CLOSED_DATES = []; // 例如 ['2026-12-25']，請與前端同步。
const CLOSED_WEEKDAYS = [6];

function setup() {
  if (Session.getEffectiveUser().getEmail() !== OWNER_EMAIL) throw new Error('請使用指定的管理帳號執行。');
  const props = PropertiesService.getScriptProperties();
  const book = SpreadsheetApp.openById(TARGET_SHEET);
  [['使用者資料',USER_HEADERS],['表單回覆 1',RESPONSE_HEADERS]].forEach(([name,headers])=>{
    const sheet=book.getSheetByName(name);
    if(!sheet)throw new Error('找不到工作表：'+name);
    if(sheet.getLastRow()===0)sheet.getRange(1,1,1,headers.length).setValues([headers]);
    else if(JSON.stringify(sheet.getRange(1,1,1,headers.length).getValues()[0])!==JSON.stringify(headers))throw new Error('欄位與程式不符，未覆寫：'+name);
    sheet.setFrozenRows(1);
  });
  props.setProperty('SPREADSHEET_ID',book.getId());
}
function verifyLine_(token) {
  if(typeof token!=='string'||token.length<20||token.length>6000)throw new Error('請重新登入 LINE');
  const response=UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify',{method:'post',payload:{id_token:token,client_id:LINE_CHANNEL_ID},muteHttpExceptions:true});
  if(response.getResponseCode()!==200)throw new Error('LINE 驗證失敗，請重新登入');
  const user=JSON.parse(response.getContentText());
  if(user.aud!==LINE_CHANNEL_ID||typeof user.sub!=='string'||!user.sub||typeof user.name!=='string'||!user.name||user.exp*1000<=Date.now())throw new Error('需要有效的 openid 與 profile 授權');
  return {id:user.sub,name:user.name};
}
function upsertUser_(book,user) {
  const sheet=book.getSheetByName('使用者資料'),now=Utilities.formatDate(new Date(),'Asia/Taipei','yyyy-MM-dd HH:mm:ss');
  const match=sheet.getLastRow()>1?sheet.getRange(2,1,sheet.getLastRow()-1,1).createTextFinder(user.id).matchEntireCell(true).findNext():null;
  const row=match?match.getRow():sheet.getLastRow()+1;
  const first=match?sheet.getRange(row,3).getDisplayValue():now;
  sheet.getRange(row,1,1,4).setNumberFormat('@').setValues([[user.id,user.name,first,now].map(s=>"'"+s)]);
}
function json_(value) { return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON); }
function doGet() { return json_({service:'埋線預約收件',ready:!!PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')}); }
function validate_(p) {
  if(!p || typeof p !== 'object')throw new Error('資料格式錯誤');
  const limits={name:60,birth:7,phone:10,date:10,session:30,medicine:40,lineName:100,notes:1000,requestId:36,consent:3};
  for(const k in limits)if(typeof p[k]!=='string'||p[k].length>limits[k])throw new Error('欄位格式錯誤');
  if(!p.name.trim()||!/^09\d{8}$/.test(p.phone)||p.consent!=='yes'||!/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(p.requestId))throw new Error('必要資料錯誤');
  const today=Utilities.formatDate(new Date(),'Asia/Taipei','yyyy-MM-dd');
  const m=p.birth.match(/^(\d{2,3})(\d{2})(\d{2})$/);
  if(!m)throw new Error('生日格式錯誤');
  const y=Number(m[1])+1911,mo=Number(m[2]),day=Number(m[3]),b=new Date(Date.UTC(y,mo-1,day));
  if(Number(m[1])<1||b.getUTCFullYear()!==y||b.getUTCMonth()!==mo-1||b.getUTCDate()!==day||b.toISOString().slice(0,10)>today)throw new Error('生日錯誤');
  if(!/^\d{4}-\d{2}-\d{2}$/.test(p.date))throw new Error('日期格式錯誤');
  const d=new Date(p.date+'T00:00:00Z');
  if(!Number.isFinite(d.getTime())||d.toISOString().slice(0,10)!==p.date||p.date<=today||CLOSED_WEEKDAYS.indexOf(d.getUTCDay())>=0||CLOSED_DATES.indexOf(p.date)>=0)throw new Error('不開放的預約日期');
  if(['上午','下午','晚上(余彬醫師)'].indexOf(p.session)<0||['埋線+門診拿藥','埋線+不用拿藥','不確定，再與醫師討論'].indexOf(p.medicine)<0)throw new Error('選項錯誤');
}
function doPost(e) {
  const lock=LockService.getScriptLock();
  try {
    if(!e||!e.postData||e.postData.contents.length>12000)throw new Error('請求格式錯誤');
    const p=JSON.parse(e.postData.contents);
    if(!['login','booking'].includes(p.action))throw new Error('不支援的操作');
    const user=verifyLine_(p.idToken);
    p.lineName=user.name;
    if(p.action==='booking')validate_(p);
    const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if(!id)throw new Error('尚未設定收件');
    const content=JSON.stringify([user.id,...['name','birth','phone','date','session','medicine','lineName','notes'].map(k=>p[k])]);
    const hash=Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,content));
    lock.waitLock(20000);
    const book=SpreadsheetApp.openById(id);
    upsertUser_(book,user);
    if(p.action==='login'){SpreadsheetApp.flush();return json_({ok:true,user});}
    const sheet=book.getSheetByName('表單回覆 1');
    if(sheet.getLastRow()>1){
      const match=sheet.getRange(2,2,sheet.getLastRow()-1,1).createTextFinder(p.requestId).matchEntireCell(true).findNext();
      if(match){
        if(sheet.getRange(match.getRow(),3).getValue()!==hash)throw new Error('此收件編號已存在，請聯繫診所確認先前資料。');
        return json_({ok:true,requestId:p.requestId});
      }
    }
    // 文字格式與前置單引號避免試算表公式注入，並保留手機及生日的前導零。
    const safe=s=>"'"+String(s);
    const row=[Utilities.formatDate(new Date(),'Asia/Taipei','yyyy-MM-dd HH:mm:ss'),p.requestId,hash,p.name,p.birth,p.phone,p.date,p.session,p.medicine,user.name,p.notes,'待確認',user.id].map(safe);
    sheet.getRange(sheet.getLastRow()+1,1,1,row.length).setNumberFormat('@').setValues([row]);
    SpreadsheetApp.flush();
    return json_({ok:true,requestId:p.requestId});
  }catch(error){return json_({ok:false,error:String(error.message||'收件失敗')});}
  finally {if(lock.hasLock())lock.releaseLock();}
}

