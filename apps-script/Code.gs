/** TRT v2. Ledger writes are one locked, append-only event per transaction. */
var CATALOG_ = [
 {id:'K1',name:'TRT-K1',unit:'桶',kg:25,initial:86,source:'自泡'},
 {id:'L1',name:'TRT-L1',unit:'桶',kg:25,initial:43,source:'自泡'},
 {id:'MC2',name:'TRT-MC2',unit:'桶',kg:30,initial:35,source:'採購',reminder:false},
 {id:'S1',name:'TRT-S1',unit:'桶',kg:25,initial:22,source:'自泡'},
 {id:'brownCarbon',name:'牛皮紙袋活性碳',unit:'包',kg:20,initial:28,source:'採購',threshold:10},
 {id:'blackCarbon',name:'黑色袋活性碳',unit:'包',kg:20,initial:47,source:'採購',leadDays:14},
 {id:'FB',name:'TRT-FB',unit:'包',kg:25,initial:40,source:'採購',batch:50,leadDays:3},
 {id:'polymer',name:'高分子',unit:'包',kg:25,initial:8,source:'採購',threshold:1}
];
var COMPANIES_=['弘大','永達','宏悅'];
function doGet() {
 return HtmlService.createHtmlOutputFromFile('App').setTitle('TRT 藥劑配送').addMetaTag('viewport','width=device-width, initial-scale=1');
}
// Deliberately reject legacy, unauthenticated POSTs after coordinated cutover.
function doPost(){return ContentService.createTextOutput(JSON.stringify({ok:false,message:'網站已更新，請使用專屬連結重新開啟'})).setMimeType(ContentService.MimeType.JSON);}
function props_(){return PropertiesService.getScriptProperties();}
function db_(){var id=props_().getProperty('SPREADSHEET_ID');if(!id)throw Error('尚未設定 SPREADSHEET_ID');return SpreadsheetApp.openById(id);}
function hash_(s){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s).map(function(b){return ('0'+((b+256)%256).toString(16)).slice(-2);}).join('');}
function token_(){return Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,'');}
function onOpen(){SpreadsheetApp.getUi().createMenu('TRT 系統').addItem('初次建立 v2（不覆寫舊表）','setupV2_').addItem('更新專屬連結網址','refreshLinks_').addToUi();}
// These owner-only functions MUST end with underscore in any remotely callable path.
// setupV2/refreshLinks are additionally gated by effective and active owner identity.
function owner_(){var a=Session.getActiveUser().getEmail(),e=Session.getEffectiveUser().getEmail();if(!a||a!==e)throw Error('請由擁有者在試算表或編輯器執行');}
function setupV2_(){owner_();var lock=LockService.getScriptLock();lock.waitLock(30000);try{
 var p=props_();if(!p.getProperty('SPREADSHEET_ID')){var ss=SpreadsheetApp.getActiveSpreadsheet();if(!ss)throw Error('請先設定 SPREADSHEET_ID');p.setProperty('SPREADSHEET_ID',ss.getId());}
 var db=db_(),ledger=db.getSheetByName('TRT_事件紀錄');if(!ledger){ledger=db.insertSheet('TRT_事件紀錄');ledger.appendRow(['事件 JSON（勿手動修改）']);ledger.setFrozenRows(1);}
 if(ledger.getLastRow()===1)append_(ledger,{id:'opening-v2',type:'opening',at:new Date().toISOString(),actor:'system',stock:CATALOG_.reduce(function(a,p){a[p.id]=p.initial;return a;},{})});
 var keys=db.getSheetByName('TRT_專屬連結')||db.insertSheet('TRT_專屬連結');if(keys.getLastRow()===0)keys.appendRow(['身分','專屬憑證（勿公開）','GitHub 專屬入口']);
 var existing=keys.getDataRange().getValues();['管理員'].concat(COMPANIES_).forEach(function(name){if(existing.some(function(r){return r[0]===name;}))return;var t=token_();p.setProperty('AUTH_'+hash_(t),JSON.stringify({role:name==='管理員'?'admin':'customer',company:name}));keys.appendRow([name,t,'待設定入口網址']);});SpreadsheetApp.flush();
 }finally{lock.releaseLock();}refreshLinks_();}
function refreshLinks_(){owner_();var url=props_().getProperty('FRONTEND_URL');if(!url)return;var sh=db_().getSheetByName('TRT_專屬連結');if(!sh)return;var rows=sh.getDataRange().getValues();rows.slice(1).forEach(function(r,i){sh.getRange(i+2,3).setValue(url+'#token='+encodeURIComponent(r[1]));});}
function auth_(token){if(typeof token!=='string'||!/^[a-f0-9]{64}$/.test(token))throw Error('請使用有效的專屬連結');var s=props_().getProperty('AUTH_'+hash_(token));if(!s)throw Error('連結已失效，請聯絡配送窗口');return JSON.parse(s);}
function events_(sheet){if(sheet.getLastRow()<2)return [];return sheet.getRange(2,1,sheet.getLastRow()-1,1).getValues().map(function(r){return JSON.parse(r[0]);});}
function append_(sheet,e){sheet.appendRow([JSON.stringify(e)]);SpreadsheetApp.flush();}
function state_(events){var s={stock:{},orders:{},movements:[]};events.forEach(function(e){
 if(e.type==='opening')s.stock=Object.assign({},e.stock);
 if(e.type==='order')s.orders[e.order.id]=JSON.parse(JSON.stringify(e.order));
 var o=s.orders[e.orderId];
 if(e.type==='receive'){o.receivedAt=e.at;o.status='已收到';o.reserved=e.reserved;}
 if(e.type==='reserve')o.reserved=true;
 if(e.type==='cancel'){o.status='已取消';o.reserved=false;}
 if(e.type==='ship'){Object.keys(o.items).forEach(function(k){s.stock[k]-=o.items[k];});o.status='已出貨';o.reserved=false;o.shippedAt=e.at;s.movements.push(e);}
 if(e.type==='adjust'){s.stock[e.product]+=e.delta;s.movements.push(e);}
 });return s;}
function reserved_(s,k){return Object.keys(s.orders).reduce(function(n,id){var o=s.orders[id];return n+(o.reserved?(o.items[k]||0):0);},0);}
function enough_(s,o){return Object.keys(o.items).every(function(k){return o.items[k]<=s.stock[k]-reserved_(s,k);});}
function items_(input){if(!input||typeof input!=='object'||Array.isArray(input))throw Error('藥劑清單格式錯誤');var out={};Object.keys(input).forEach(function(k){if(!CATALOG_.some(function(p){return p.id===k;}))throw Error('未知品項');var q=input[k];if(typeof q!=='number'||!Number.isSafeInteger(q)||q<0||q>9999)throw Error('數量需為 0～9999 的整數');if(q)out[k]=q;});if(!Object.keys(out).length)throw Error('請至少選一種藥劑');return out;}
function view_(s,user,version){var orders=Object.keys(s.orders).map(function(k){return s.orders[k];}).filter(function(o){return user.role==='admin'||o.company===user.company;}).sort(function(a,b){return b.at.localeCompare(a.at);});
 var result={ok:true,version:version,user:user,products:CATALOG_.map(function(p){var c=Object.assign({},p);delete c.initial;return c;}),orders:orders};
 if(user.role==='admin'){result.inventory=CATALOG_.map(function(p){var r=reserved_(s,p.id),available=s.stock[p.id]-r;return {id:p.id,onHand:s.stock[p.id],reserved:r,available:available,low:p.threshold!==undefined&&available<=p.threshold};});result.movements=s.movements.slice(-100).reverse();}return result;}
function api(request){var lock=LockService.getScriptLock(),writeStarted=false;try{
 if(!request||typeof request!=='object')throw Error('請求格式錯誤');var user=auth_(request.token);lock.waitLock(30000);var sheet=db_().getSheetByName('TRT_事件紀錄');if(!sheet)throw Error('尚未初始化');var events=events_(sheet),s=state_(events),action=request.action;
 if(action==='read')return view_(s,user,events.length);
 if(action==='legacy')return {ok:true,legacy:legacy_(user)};
 if(typeof request.requestId!=='string'||!/^[-a-zA-Z0-9_]{16,100}$/.test(request.requestId))throw Error('缺少有效的交易識別碼');
 var fingerprint=hash_(JSON.stringify({action:action,orderId:request.orderId||'',items:request.items||{},product:request.product||'',delta:request.delta||0,note:request.note||''}));
 var actor=user.role==='admin'?'admin':user.company;var duplicate=events.filter(function(e){return e.id===request.requestId&&e.actor===actor;})[0];if(duplicate){if(duplicate.fingerprint!==fingerprint)throw Error('同一交易識別碼不能重用於不同內容');return view_(s,user,events.length);}
 var event={id:request.requestId,actor:actor,at:new Date().toISOString(),fingerprint:fingerprint};
 if(action==='order'){
 if(user.role!=='customer')throw Error('請使用客戶入口叫藥');var items=items_(request.items);event.type='order';event.order={id:Utilities.getUuid(),company:user.company,at:event.at,items:items,status:'待確認',reserved:false};
 }else{
 if(user.role!=='admin')throw Error('您沒有管理權限');
 if(action==='adjust'){var product=CATALOG_.filter(function(p){return p.id===request.product;})[0],d=request.delta,note=String(request.note||'').trim();if(!product||!Number.isSafeInteger(d)||!d||Math.abs(d)>9999||!note||note.length>300)throw Error('請選品項、輸入有效數量並填原因');if(s.stock[product.id]+d<reserved_(s,product.id))throw Error('不能扣除已預留的庫存');event.type='adjust';event.product=product.id;event.delta=d;event.note=note;
 }else{var o=s.orders[request.orderId];if(!o)throw Error('找不到訂單');event.orderId=o.id;
 if(action==='receive'){if(o.status!=='待確認')return view_(s,user,events.length);event.type='receive';event.reserved=enough_(s,o);}
 else if(action==='reserve'){if(o.status!=='已收到')throw Error('請先按收到');if(o.reserved)return view_(s,user,events.length);if(!enough_(s,o))throw Error('可用庫存不足，請先入庫');event.type='reserve';}
 else if(action==='ship'){if(o.status==='已出貨')return view_(s,user,events.length);if(o.status!=='已收到'||!o.reserved)throw Error('請先收到並預留庫存');event.type='ship';}
 else if(action==='cancel'){if(o.status==='已取消')return view_(s,user,events.length);if(o.status==='已出貨')throw Error('已出貨訂單不能取消，請另做退貨入庫');event.type='cancel';}
 else throw Error('不支援的操作');}
 }
 writeStarted=true;append_(sheet,event);events.push(event);return view_(state_(events),user,events.length);
 }catch(err){return {ok:false,message:err.message||String(err),retrySameRequest:writeStarted};}finally{if(lock.hasLock())lock.releaseLock();}}

function legacy_(user){var db=db_(),names=user.role==='admin'?COMPANIES_:[user.company],out=[],keys=['K1','L1','MC2','S1','brownCarbon','blackCarbon','FB','polymer'];names.forEach(function(name){var sh=db.getSheetByName(name);if(!sh||sh.getLastRow()<2)return;var values=sh.getDataRange().getValues();if(String(values[0][1]).indexOf('公司')<0)return;values.slice(1).forEach(function(r){var items={};keys.forEach(function(k,i){var q=Number(r[i+2]);if(Number.isFinite(q)&&q>0)items[k]=q;});if(!Object.keys(items).length)return;var d=r[0];out.push({company:name,at:d instanceof Date?Utilities.formatDate(d,'Asia/Taipei','yyyy/MM/dd HH:mm'):String(d||''),items:items});});});return out.reverse();}
