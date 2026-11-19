/** MTTQ An Giang - backend V3.4 – Một Admin. Chạy setup() một lần trong Apps Script gắn với Sheet riêng. */
const SHEETS={TICKETS:'YeuCau',LOG:'LichSu',SETTINGS:'CauHinh',DASH:'TongQuan'};
const HEADERS=['MaPhieu','MaXacThucHash','NgayGui','HoTen','DonVi','Email','DienThoai','NhomHoTro','MucDo','TieuDe','MoTa','TrangThai','KetQua','NgayCapNhat','DaThongBao','AnhMinhChung','NguoiPhuTrach','GhiChuNoiBo','HoanThanh','AnhKetQua'];
const STATUS=['Chờ tiếp nhận','Đang xử lý','Hoàn thành'];
function setup(){
 const ss=SpreadsheetApp.getActiveSpreadsheet();if(!ss)throw Error('Hãy mở Apps Script từ Google Sheets mới.');
 let sh=ss.getSheetByName(SHEETS.TICKETS)||ss.insertSheet(SHEETS.TICKETS);
 if(sh.getLastRow()===0){sh.appendRow(HEADERS);sh.setFrozenRows(1);sh.getRange('A:A').setNumberFormat('@');sh.getRange('B:B').setNumberFormat('@');}else{const existing=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];HEADERS.forEach((h,i)=>{if(!existing[i])sh.getRange(1,i+1).setValue(h)})}
 let log=ss.getSheetByName(SHEETS.LOG)||ss.insertSheet(SHEETS.LOG);if(!log.getLastRow())log.appendRow(['ThoiGian','MaPhieu','HanhDong','ChiTiet']);
 let cfg=ss.getSheetByName(SHEETS.SETTINGS)||ss.insertSheet(SHEETS.SETTINGS);if(!cfg.getLastRow())cfg.appendRow(['ThietLap','GiaTri']);
 const validation=SpreadsheetApp.newDataValidation().requireValueInList(STATUS,true).setAllowInvalid(false).build();sh.getRange('L2:L').setDataValidation(validation);
 sh.getRange('S2:S').insertCheckboxes();
 sh.getRange('A1:T1').setFontWeight('bold').setBackground('#b91c1c').setFontColor('#ffffff');
 sh.setFrozenRows(1);if(!sh.getFilter())sh.getRange(1,1,Math.max(2,sh.getLastRow()),HEADERS.length).createFilter();
 createDashboard_(ss);
 PropertiesService.getScriptProperties().setProperty('SHEET_ID',ss.getId());
 const props=PropertiesService.getScriptProperties();if(!props.getProperty('EVIDENCE_FOLDER_ID')){const folder=DriveApp.createFolder('MTTQ An Giang - Minh chung ho so');props.setProperty('EVIDENCE_FOLDER_ID',folder.getId())}
 ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='onTicketEdit').forEach(t=>ScriptApp.deleteTrigger(t));
 ScriptApp.newTrigger('onTicketEdit').forSpreadsheet(ss).onEdit().create();
 return 'Đã thiết lập YeuCau, LichSu, CauHinh, TongQuan và checkbox HoanThanh.';
}
function db(){const id=PropertiesService.getScriptProperties().getProperty('SHEET_ID');if(!id)throw Error('Chưa chạy setup().');return SpreadsheetApp.openById(id)}
function digest(s){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s,Utilities.Charset.UTF_8).map(b=>('0'+(b&255).toString(16)).slice(-2)).join('')}
function clean(v,n){return String(v==null?'':v).trim().slice(0,n).replace(/[<>]/g,'')}
function safeCell(v){return /^[=+\-@]/.test(v)?"'"+v:v}
function logTicket(id,action,detail){db().getSheetByName(SHEETS.LOG).appendRow([new Date(),id,action,detail||''])}
function submitTicket(p){
 if(!p||typeof p!=='object')throw Error('Dữ liệu không hợp lệ');
 const name=clean(p.name,100),unit=clean(p.unit,150),email=clean(p.email,160).toLowerCase(),phone=clean(p.phone,20),category=clean(p.category,80),priority=clean(p.priority,20),subject=clean(p.subject,180),detail=clean(p.detail,5000);
 if(!name||!unit||!email||!category||!subject||!detail||!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))throw Error('Vui lòng điền đầy đủ thông tin hợp lệ.');
 if(!['Phần mềm tác nghiệp','Chữ ký số','Chuyển đổi số','Khác'].includes(category))throw Error('Nhóm hỗ trợ không hợp lệ');
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 let id,code;try{
  const sh=db().getSheetByName(SHEETS.TICKETS);id='AG-'+Utilities.formatDate(new Date(),'Asia/Ho_Chi_Minh','yyyyMMdd')+'-'+Utilities.getUuid().slice(0,8).toUpperCase();
  code=Utilities.getUuid().replace(/-/g,'').slice(0,12).toUpperCase();
  const attachment=saveEvidence_(p.image,id);
  try{sh.appendRow([id,digest(id+':'+code),new Date(),safeCell(name),safeCell(unit),email,safeCell(phone),category,priority==='Khẩn'?'Khẩn':'Bình thường',safeCell(subject),safeCell(detail),STATUS[0],'',new Date(),'',attachment.url,'','',false,'']);}catch(err){if(attachment.fileId)DriveApp.getFileById(attachment.fileId).setTrashed(true);throw err}
 }finally{lock.releaseLock()}
 logTicket(id,'Tiếp nhận','Yêu cầu mới');refreshDashboard_(db());
 let mailed=false;
 try{MailApp.sendEmail({to:email,subject:'[MTTQ An Giang] Đã tiếp nhận yêu cầu '+id,body:'Kính gửi '+name+',\n\nYêu cầu hỗ trợ của anh/chị đã được tiếp nhận.\nMã phiếu: '+id+'\nMã xác thực: '+code+'\n\nVui lòng giữ kín mã xác thực. Tra cứu tại https://mattransoag.github.io/hotrokythuat/\n\nỦy ban MTTQ Việt Nam tỉnh An Giang.'});mailed=true;}catch(e){logTicket(id,'Lỗi gửi Gmail',String(e).slice(0,250))}
 return {ok:true,id:id,verificationCode:code,emailSent:mailed};
}
/** Ảnh chỉ lưu riêng tư trong Drive; không cấp quyền 'anyone with link'. */
function saveEvidence_(image,id){
 if(!image)return {url:'',fileId:''};
 if(typeof image!=='object'||typeof image.data!=='string')throw Error('Ảnh minh chứng không hợp lệ');
 const type=String(image.type||'');
 const ext={'image/jpeg':'jpg','image/png':'png','image/webp':'webp'}[type];
 if(!ext)throw Error('Chỉ nhận ảnh JPG, PNG hoặc WebP');
 if(image.data.length>4200000)throw Error('Ảnh vượt quá 3 MB');
 let bytes;try{bytes=Utilities.base64Decode(image.data)}catch(e){throw Error('Dữ liệu ảnh không hợp lệ')}
 if(bytes.length>3*1024*1024)throw Error('Ảnh vượt quá 3 MB');
 const folderId=PropertiesService.getScriptProperties().getProperty('EVIDENCE_FOLDER_ID');
 if(!folderId)throw Error('Chưa chạy setup() để tạo thư mục minh chứng');
 const blob=Utilities.newBlob(bytes,type,id+'.'+ext);
 const file=DriveApp.getFolderById(folderId).createFile(blob);
 return {url:file.getUrl(),fileId:file.getId()};
}
function lookupTicket(id,code){
 id=clean(id,50).toUpperCase();code=clean(code,80).toUpperCase();
 if(!id||!code)throw Error('Cần nhập mã phiếu và mã xác thực');
 const sh=db().getSheetByName(SHEETS.TICKETS),last=sh.getLastRow();if(last<2)throw Error('Không tìm thấy phiếu hoặc mã xác thực không đúng');
 const ids=sh.getRange(2,1,last-1,1).getDisplayValues();
 let index=ids.findIndex(r=>r[0]===id);if(index<0)throw Error('Không tìm thấy phiếu hoặc mã xác thực không đúng');
 const r=sh.getRange(index+2,1,1,HEADERS.length).getValues()[0];
 if(r[1]!==digest(id+':'+code))throw Error('Không tìm thấy phiếu hoặc mã xác thực không đúng');
 return {id:r[0],date:String(r[2]),name:r[3],unit:r[4],category:r[7],subject:r[9],detail:r[10],status:r[11],result:r[12]||'',updated:String(r[13])};
}
/** Bảng tổng quan riêng tư: chỉ ai được chia sẻ Sheet mới xem được. */
function createDashboard_(ss){
 let sh=ss.getSheetByName(SHEETS.DASH)||ss.insertSheet(SHEETS.DASH);
 sh.clear();
 sh.getRange('A1:D1').merge().setValue('ỦY BAN MTTQ VIỆT NAM TỈNH AN GIANG – DASHBOARD ADMIN').setBackground('#b91c1c').setFontColor('#ffffff').setFontWeight('bold').setFontSize(15);
 const stats=[['TỔNG HỒ SƠ',0],['CHỜ TIẾP NHẬN',0],['ĐANG XỬ LÝ',0],['HOÀN THÀNH',0]];
 sh.getRange(3,1,stats.length,2).setValues(stats);refreshDashboard_(ss);sh.getRange('A3:A6').setFontWeight('bold');sh.getRange('B3:B6').setFontSize(18).setFontWeight('bold');
 sh.getRange('A8:D8').merge().setValue('QUẢN LÝ: Mở tab YeuCau → lọc trạng thái → nhập kết quả cột M → tích ô HoanThanh cột S.').setWrap(true);
 sh.getRange('A9:D9').merge().setValue('Ảnh minh chứng và ảnh kết quả lưu trong Google Drive riêng; không chia sẻ Sheet hoặc thư mục công khai.').setWrap(true);
 sh.setColumnWidths(1,4,210);sh.setRowHeight(8,55);sh.setRowHeight(9,55);
}
/** Ghi số liệu trực tiếp, không phụ thuộc dấu phân cách công thức theo ngôn ngữ Sheets. */
function refreshDashboard_(ss){
 const tickets=ss.getSheetByName(SHEETS.TICKETS),dash=ss.getSheetByName(SHEETS.DASH);
 if(!tickets||!dash)return;
 const last=tickets.getLastRow(),counts=[0,0,0,0];
 if(last>=2){
  const rows=tickets.getRange(2,1,last-1,12).getDisplayValues();
  rows.forEach(r=>{if(!validTicketId_(r[0]))return;counts[0]++;if(r[11]==='Chờ tiếp nhận')counts[1]++;else if(r[11]==='Đang xử lý')counts[2]++;else if(r[11]==='Hoàn thành')counts[3]++});
 }
 dash.getRange(3,2,4,1).setValues(counts.map(x=>[x]));
}
function onOpen(){
 SpreadsheetApp.getUi().createMenu('MTTQ - Quản lý hồ sơ')
 .addItem('Mở Dashboard','openDashboard_')
 .addItem('Trả lời nhanh: Đã tiếp nhận','quickAcknowledge_')
 .addItem('Trả lời nhanh: Đang kiểm tra','quickChecking_')
 .addItem('Đánh dấu hoàn thành hồ sơ đang chọn','completeSelected_')
 .addToUi();
}
function openDashboard_(){SpreadsheetApp.getActive().setActiveSheet(SpreadsheetApp.getActive().getSheetByName(SHEETS.DASH))}
function selectedTicketRow_(){const sh=SpreadsheetApp.getActiveSheet(),row=sh.getActiveRange().getRow();if(sh.getName()!==SHEETS.TICKETS||row<2)throw Error('Hãy chọn một dòng hồ sơ trong tab YeuCau.');return {sh,row}}
function quickAcknowledge_(){setQuickResult_('Đã tiếp nhận yêu cầu. Đang kiểm tra và sẽ phản hồi khi có kết quả.')}
function quickChecking_(){setQuickResult_('Đang kiểm tra sự cố. Vui lòng chờ cập nhật tiếp theo.')}
function setQuickResult_(message){const o=selectedTicketRow_();o.sh.getRange(o.row,13).setValue(message);o.sh.getRange(o.row,12).setValue('Đang xử lý');notifyTicketRow_(o.sh,o.row)}
function completeSelected_(){const o=selectedTicketRow_();o.sh.getRange(o.row,19).setValue(true);notifyTicketRow_(o.sh,o.row)}
function onTicketEdit(e){
 if(!e||!e.range||e.range.getSheet().getName()!==SHEETS.TICKETS||e.range.getRow()<2)return;
 const col=e.range.getColumn();if(![12,13,19].includes(col))return;
 const sh=e.range.getSheet(),row=e.range.getRow();
 if(col===19&&e.range.getValue()!==true)return;
 if(col===13)return; // ghi kết quả trước, chỉ gửi khi thay đổi trạng thái hoặc hoàn thành
 notifyTicketRow_(sh,row);
}
function notifyTicketRow_(sh,row){
 const r=sh.getRange(row,1,1,HEADERS.length).getValues()[0];
 if(r[18]===true){if(!String(r[12]||'').trim()){sh.getRange(row,19).setValue(false);throw Error('Vui lòng nhập Kết quả xử lý (cột M) trước khi hoàn thành.')}sh.getRange(row,12).setValue('Hoàn thành')}
 if(!STATUS.includes(sh.getRange(row,12).getValue()))return;
 const status=sh.getRange(row,12).getValue(),result=String(sh.getRange(row,13).getValue()||'');
 sh.getRange(row,14).setValue(new Date());logTicket(r[0],'Cập nhật trạng thái',status);refreshDashboard_(db());
 try{MailApp.sendEmail({to:r[5],subject:'[MTTQ An Giang] Cập nhật phiếu '+r[0],body:'Mã phiếu: '+r[0]+'\nTrạng thái: '+status+'\nKết quả: '+(result||'Chưa có')+'\n\nVui lòng tra cứu bằng mã phiếu và mã xác thực đã nhận.'});sh.getRange(row,15).setValue(new Date());return {mailSent:true}}catch(err){logTicket(r[0],'Lỗi gửi Gmail',String(err).slice(0,250));return {mailSent:false,mailError:String(err).slice(0,250)}}
}

/** Admin duy nhất: cấu hình ADMIN_PASSWORD trong Script Properties, tuyệt đối không lưu vào GitHub. */
function adminLogin(username,password){
 if(String(username)!=='admin')throw Error('Sai tên đăng nhập hoặc mật khẩu');
 const configured=PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
 if(!configured)throw Error('Admin chưa được cấu hình. Vào Project Settings → Script properties.');
 const attempts=CacheService.getScriptCache(),key='admin_failed_count',failed=Number(attempts.get(key)||0);
 if(failed>=10)throw Error('Đăng nhập tạm khóa 15 phút. Vui lòng thử lại sau.');
 if(String(password)!==configured){attempts.put(key,String(failed+1),900);throw Error('Sai tên đăng nhập hoặc mật khẩu')}
 attempts.remove(key);
 const token=Utilities.getUuid()+Utilities.getUuid();
 attempts.put('admin_session_'+token,'1',1200);
 return {token,expiresMinutes:20};
}
function requireAdmin_(token){
 if(typeof token!=='string'||!/^[-a-f0-9]{72}$/.test(token)||CacheService.getScriptCache().get('admin_session_'+token)!=='1')throw Error('Phiên quản trị đã hết hạn. Vui lòng đăng nhập lại.');
}
function adminLogout(token){if(typeof token==='string')CacheService.getScriptCache().remove('admin_session_'+token);return true}
function validTicketId_(id){return /^AG-\d{8}-[A-F0-9]{8}$/.test(String(id||'').trim())}
function adminListTickets(token){
 requireAdmin_(token);
 const sh=db().getSheetByName(SHEETS.TICKETS),last=sh.getLastRow();
 // Chỉ đếm dòng có mã phiếu thật, bỏ qua hàng trống, checkbox và công thức.
 if(last<2)return {items:[],stats:{total:0,pending:0,working:0,done:0}};
 const ids=sh.getRange(2,1,last-1,1).getDisplayValues();
 const validRows=[];
 for(let i=0;i<ids.length;i++)if(validTicketId_(ids[i][0]))validRows.push(i+2);
 const counts={total:validRows.length,pending:0,working:0,done:0};
 const statuses=sh.getRange(2,12,last-1,1).getDisplayValues();
 validRows.forEach(row=>{const status=statuses[row-2][0];if(status==='Chờ tiếp nhận')counts.pending++;else if(status==='Đang xử lý')counts.working++;else if(status==='Hoàn thành')counts.done++});
 const recent=validRows.slice(-200).reverse();
 const items=recent.map(row=>{const r=sh.getRange(row,1,1,HEADERS.length).getDisplayValues()[0];return {id:r[0],date:r[2],name:r[3],unit:r[4],email:r[5],category:r[7],subject:r[9],detail:r[10],status:r[11],result:r[12]||'',image:r[15]||'',imageResult:r[19]||''}});
 refreshDashboard_(db());return {items,stats:counts};
}
function adminUpdateTicket(token,id,status,result){
 requireAdmin_(token);id=clean(id,50);result=clean(result,5000);
 if(!STATUS.includes(status))throw Error('Trạng thái không hợp lệ');
 if(status==='Hoàn thành'&&!result)throw Error('Cần nhập kết quả trước khi hoàn thành');
 const sh=db().getSheetByName(SHEETS.TICKETS),last=sh.getLastRow();
 if(last<2)throw Error('Không tìm thấy hồ sơ');
 const ids=sh.getRange(2,1,last-1,1).getDisplayValues(),i=ids.findIndex(r=>r[0]===id);
 if(i<0||!validTicketId_(id))throw Error('Không tìm thấy hồ sơ');
 const row=i+2;const previous=sh.getRange(row,12,1,2).getDisplayValues()[0];if(previous[0]===status&&previous[1]===result)return {ok:true,unchanged:true};sh.getRange(row,13).setValue(safeCell(result));sh.getRange(row,12).setValue(status);
 sh.getRange(row,19).setValue(status==='Hoàn thành');
 const notification=notifyTicketRow_(sh,row);
 return {ok:true,mailSent:!!(notification&&notification.mailSent),mailError:notification&&notification.mailError||''};
}

/** Đếm phiên truy cập công khai, không lưu IP, email hoặc dữ liệu cá nhân. */
function getPublicVisitStats(visitorId){
 const props=PropertiesService.getScriptProperties();
 const tz=Session.getScriptTimeZone()||'Asia/Ho_Chi_Minh';
 const day=Utilities.formatDate(new Date(),tz,'yyyyMMdd');
 const dailyKey='VISITS_DAY_'+day;
 const totalKey='VISITS_TOTAL';
 const key=String(visitorId||'').trim();
 const valid=/^[a-zA-Z0-9_-]{16,80}$/.test(key);
 const cache=CacheService.getScriptCache();
 const lock=LockService.getScriptLock();
 lock.waitLock(10000);
 try{
  const cacheKey=valid?'visitor_'+Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,key)).slice(0,45):'';
  if(!valid||!cache.get(cacheKey)){
   props.setProperty(dailyKey,String(Number(props.getProperty(dailyKey)||0)+1));
   props.setProperty(totalKey,String(Number(props.getProperty(totalKey)||0)+1));
   if(valid)cache.put(cacheKey,'1',21600);
  }
  return {today:Number(props.getProperty(dailyKey)||0),total:Number(props.getProperty(totalKey)||0)};
 }finally{lock.releaseLock()}
}
function doGet(e){const t=HtmlService.createTemplateFromFile('App');t.entryView=e&&e.parameter&&['send','lookup','admin'].includes(e.parameter.view)?e.parameter.view:'send';return t.evaluate().setTitle('Hỗ trợ kỹ thuật MTTQ An Giang').addMetaTag('viewport','width=device-width,initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)}
