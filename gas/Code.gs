/** MTTQ An Giang - backend V3.1. Chạy setup() một lần trong Apps Script gắn với Sheet riêng. */
const SHEETS={TICKETS:'YeuCau',LOG:'LichSu',SETTINGS:'CauHinh'};
const HEADERS=['MaPhieu','MaXacThucHash','NgayGui','HoTen','DonVi','Email','DienThoai','NhomHoTro','MucDo','TieuDe','MoTa','TrangThai','KetQua','NgayCapNhat','DaThongBao'];
const STATUS=['Chờ tiếp nhận','Đang xử lý','Hoàn thành'];
function setup(){
 const ss=SpreadsheetApp.getActiveSpreadsheet();if(!ss)throw Error('Hãy mở Apps Script từ Google Sheets mới.');
 let sh=ss.getSheetByName(SHEETS.TICKETS)||ss.insertSheet(SHEETS.TICKETS);
 if(sh.getLastRow()===0){sh.appendRow(HEADERS);sh.setFrozenRows(1);sh.getRange('A:A').setNumberFormat('@');sh.getRange('B:B').setNumberFormat('@');}
 let log=ss.getSheetByName(SHEETS.LOG)||ss.insertSheet(SHEETS.LOG);if(!log.getLastRow())log.appendRow(['ThoiGian','MaPhieu','HanhDong','ChiTiet']);
 let cfg=ss.getSheetByName(SHEETS.SETTINGS)||ss.insertSheet(SHEETS.SETTINGS);if(!cfg.getLastRow())cfg.appendRow(['ThietLap','GiaTri']);
 const validation=SpreadsheetApp.newDataValidation().requireValueInList(STATUS,true).setAllowInvalid(false).build();sh.getRange('L2:L').setDataValidation(validation);
 PropertiesService.getScriptProperties().setProperty('SHEET_ID',ss.getId());
 ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='onTicketEdit').forEach(t=>ScriptApp.deleteTrigger(t));
 ScriptApp.newTrigger('onTicketEdit').forSpreadsheet(ss).onEdit().create();
 return 'Đã tạo các trang YeuCau, LichSu, CauHinh và trình kích hoạt thông báo.';
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
  sh.appendRow([id,digest(id+':'+code),new Date(),safeCell(name),safeCell(unit),email,safeCell(phone),category,priority==='Khẩn'?'Khẩn':'Bình thường',safeCell(subject),safeCell(detail),STATUS[0],'',new Date(),'']);
 }finally{lock.releaseLock()}
 logTicket(id,'Tiếp nhận','Yêu cầu mới');
 let mailed=false;
 try{MailApp.sendEmail({to:email,subject:'[MTTQ An Giang] Đã tiếp nhận yêu cầu '+id,body:'Kính gửi '+name+',\n\nYêu cầu hỗ trợ của anh/chị đã được tiếp nhận.\nMã phiếu: '+id+'\nMã xác thực: '+code+'\n\nVui lòng giữ kín mã xác thực. Tra cứu tại https://mattransoag.github.io/hotrokythuat/\n\nỦy ban MTTQ Việt Nam tỉnh An Giang.'});mailed=true;}catch(e){logTicket(id,'Lỗi gửi Gmail',String(e).slice(0,250))}
 return {ok:true,id:id,verificationCode:code,emailSent:mailed};
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
function onTicketEdit(e){
 if(!e||!e.range||e.range.getSheet().getName()!==SHEETS.TICKETS||e.range.getColumn()!==12||e.range.getRow()<2)return;
 const sh=e.range.getSheet(),row=e.range.getRow(),r=sh.getRange(row,1,1,HEADERS.length).getValues()[0];
 if(!STATUS.includes(r[11]))return;
 sh.getRange(row,14).setValue(new Date());
 logTicket(r[0],'Cập nhật trạng thái',r[11]);
 try{MailApp.sendEmail({to:r[5],subject:'[MTTQ An Giang] Cập nhật phiếu '+r[0],body:'Yêu cầu '+r[0]+' đã được cập nhật trạng thái: '+r[11]+'.\nKết quả: '+(r[12]||'Chưa có')+'\n\nTra cứu bằng mã phiếu và mã xác thực đã gửi trước đó.'});sh.getRange(row,15).setValue(new Date());}catch(err){logTicket(r[0],'Lỗi gửi Gmail',String(err).slice(0,250))}
}
function doGet(){return HtmlService.createHtmlOutputFromFile('App').setTitle('Hỗ trợ kỹ thuật MTTQ An Giang').addMetaTag('viewport','width=device-width,initial-scale=1')}
