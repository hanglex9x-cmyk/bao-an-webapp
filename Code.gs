/*************************************************
 * PHẦN MỀM BÁN HÀNG SƠN NƯỚC BẢO ÂN - Apps Script Web App
 * Dùng chung file Sheet: Phan-mem-ban-hang-GoogleSheets
 * ID: 1hKbNBsxkSST-pzb4ciegTDQtHe2o_KAJhvb7ICgTkNk
 * Tab có sẵn: KhachHang, SanPham, NhatKyBanHang, ThuTien, TongHopCongNo, Dashboard, InHoaDon
 * Tab tự tạo khi chạy caiDatHeThong(): CauHinh, DonHang, DiemTichLuy, ThuChi, NhapKho, NguoiDung
 * QUAN TRỌNG: NhatKyBanHang + ThuTien đã có công thức sẵn ở các cột đen tới dòng 1003/2003.
 * Code này CHỈ ghi vào cột nhập tay, không ghi đè công thức.
 *************************************************/
var TZ = 'Asia/Ho_Chi_Minh';
var BRANDS = ['Nippon', 'Jotun', 'Dulux', 'TOA', 'Phụ kiện'];
var UNITS = ['5L', '17L', '4kg', '18kg', '20kg', 'cái', 'lít', 'kg'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Sơn nước Bảo Ân - Phần mềm bán hàng')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sh_(name) { return ss_().getSheetByName(name); }
function fmtDate_(d) { if (!d) return ''; if (Object.prototype.toString.call(d) === '[object Date]') return Utilities.formatDate(d, TZ, 'dd/MM/yyyy'); return String(d); }
function parseDate_(s) { if (!s) return new Date(); if (Object.prototype.toString.call(s) === '[object Date]') return s; var p = String(s).split(/[-/]/); if (p.length === 3 && p[0].length === 4) return new Date(+p[0], +p[1]-1, +p[2]); if (p.length === 3) return new Date(+p[2], +p[1]-1, +p[0]); return new Date(s); }
function money_(v) { if (v === '' || v === null || v === undefined || v === '-') return 0; if (typeof v === 'number') return v; var n = parseFloat(String(v).replace(/[^\d,-]/g, '').replace(/\./g, '').replace(',', '.')); return isNaN(n) ? 0 : n; }
function pct_(v) { if (typeof v === 'number') return v <= 1 ? v : v/100; var n = parseFloat(String(v||'0').replace('%','').replace(',','.')); return isNaN(n) ? 0 : n/100; }
function todayStr_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

/* ============ CẤU HÌNH ============ */
function getConfig() {
  var s = sh_('CauHinh'); var cfg = { TenCuaHang: 'Cửa hàng sơn nước Bảo Ân', DiaChi:'', SDT:'', MST:'', BankCode:'', TenNganHang:'', SoTaiKhoan:'', ChuTaiKhoan:'', TienCho1Diem: 10000, TonCanhBao: 5, DiemBatTat: 1, LogoUrl:'', Slogan:'Nơi đặt trọn niềm tin', HD_LoiCamOn:'Cảm ơn quý khách đã tin tưởng Bảo Ân!', HD_HienLogo:1, HD_HienDiaChi:1, HD_HienSDT:1, HD_HienMST:1, HD_HienSDTKH:1, HD_HienHanTT:1, HD_HienQR:1, HD_HienDiem:1, HD_HienChietKhau:1, HD_HienDonGia:1, HD_HienHinhSP:1, HD_HienDaThuConNo:1, HD_HienGhiChu:1, HD_HienChuKy:1 };
  if (!s) return cfg;
  var v = s.getRange(2, 1, Math.max(1, s.getLastRow()-1), 2).getValues();
  v.forEach(function(r){ if (r[0]) cfg[String(r[0]).trim()] = r[1]; });
  cfg.TienCho1Diem = money_(cfg.TienCho1Diem) || 10000; cfg.TonCanhBao = money_(cfg.TonCanhBao) || 5;
  return cfg;
}
function saveConfig(c) { requireRole_(['Admin']); var s = sh_('CauHinh'); if (!s) { caiDatHeThong(); s = sh_('CauHinh'); } var keys = Object.keys(c); var rows = keys.map(function(k){ return [k, c[k]]; }); s.getRange(2, 1, Math.max(rows.length, s.getLastRow()), 2).clearContent(); s.getRange(2, 1, rows.length, 2).setValues(rows); return getConfig(); }

/* ============ PHÂN QUYỀN ============ */
function getCurrentUser() {
  var email = ''; try { email = Session.getActiveUser().getEmail() || ''; } catch(e) {}
  var s = sh_('NguoiDung'); var role = 'Admin', name = email || 'Chủ cửa hàng', found = false, active = true;
  if (s && s.getLastRow() >= 2) {
    var v = s.getRange(2, 1, s.getLastRow()-1, 5).getValues();
    for (var i=0;i<v.length;i++){ if (email && String(v[i][0]).toLowerCase() === email.toLowerCase()) { found = true; name = v[i][1] || email; role = v[i][2] || 'Xem'; active = String(v[i][3]||'Hoạt động') === 'Hoạt động'; break; } }
    // Có danh sách user nhưng email hiện tại không có trong danh sách và có lấy được email -> chỉ Xem
    if (!found && email) { var hasAny = v.some(function(r){ return r[0]; }); if (hasAny) role = 'Xem'; }
  }
  var perms = { Admin:['*'], BanHang:['dashboard','taodon','donhang','lichsu','khachhang','congno','diem','loinhuan'], Kho:['dashboard','donhang','lichsu','tonkho','hanghoa'], KeToan:['dashboard','donhang','lichsu','khachhang','congno','thuchi','diem','loinhuan'], Xem:['dashboard','donhang','lichsu','tonkho','khachhang','congno','diem','loinhuan'] };
  return { email: email, name: name, role: role, active: active, perms: perms[role] || [] };
}
function requireRole_(roles) { var u = getCurrentUser(); if (!u.active) throw new Error('Tài khoản đã bị khóa.'); if (u.role === 'Admin') return u; if (roles.indexOf(u.role) < 0) throw new Error('Tài khoản vai trò ' + u.role + ' không có quyền làm việc này.'); return u; }
function getUsers() { var s = sh_('NguoiDung'); if (!s || s.getLastRow() < 2) return []; return s.getRange(2,1,s.getLastRow()-1,5).getValues().filter(function(r){return r[0];}).map(function(r){ return { email:r[0], name:r[1], role:r[2], status:r[3]||'Hoạt động', note:r[4]||'' }; }); }
function saveUser(u) { requireRole_(['Admin']); var s = sh_('NguoiDung'); if (!s) { caiDatHeThong(); s = sh_('NguoiDung'); } var v = s.getLastRow()>=2 ? s.getRange(2,1,s.getLastRow()-1,1).getValues() : []; for (var i=0;i<v.length;i++){ if (String(v[i][0]).toLowerCase() === String(u.email).toLowerCase()) { s.getRange(i+2,1,1,5).setValues([[u.email,u.name,u.role,u.status||'Hoạt động',u.note||'']]); return getUsers(); } } s.appendRow([u.email,u.name,u.role,u.status||'Hoạt động',u.note||'']); return getUsers(); }
function deleteUser(email) { requireRole_(['Admin']); var s = sh_('NguoiDung'); var v = s.getRange(2,1,s.getLastRow()-1,1).getValues(); for (var i=0;i<v.length;i++){ if (String(v[i][0]).toLowerCase()===String(email).toLowerCase()) { s.deleteRow(i+2); break; } } return getUsers(); }

/* ============ DỮ LIỆU GỐC ============ */
function readCustomers_() { var s = sh_('KhachHang'); if (!s || s.getLastRow()<4) return []; return s.getRange(4,1,s.getLastRow()-3,9).getValues().filter(function(r){return r[0];}).map(function(r){ return { ma:r[0], ten:r[1], loai:r[2]||'', sdt:String(r[3]||''), tinh:r[4]||'', hanMuc:money_(r[5]), kyHan:money_(r[6]), noDauKy:money_(r[7]), ghiChu:r[8]||'' }; }); }
function readProducts_() { var s = sh_('SanPham'); if (!s || s.getLastRow()<4) return []; var last = s.getLastRow(); var vals = s.getRange(4,1,last-3,9).getValues(); return vals.filter(function(r){return r[0];}).map(function(r){ return { ma:r[0], nhom:r[1]||'', ten:r[2], dvt:r[3]||'', gia:money_(r[4]), tonDau:money_(r[5]), tonMin:money_(r[6]), hinh:String(r[7]||''), giaVon:money_(r[8]) }; }); }
function stockMap_() { var nhap = {}, xuat = {}; var n = sh_('NhapKho'); if (n && n.getLastRow()>=2) n.getRange(2,2,n.getLastRow()-1,4).getValues().forEach(function(r){ if(r[0]) nhap[r[0]] = (nhap[r[0]]||0) + money_(r[3]); }); var s = sh_('NhatKyBanHang'); if (s && s.getLastRow()>=4) s.getRange(4,5,s.getLastRow()-3,4).getValues().forEach(function(r){ if(r[0]) xuat[r[0]] = (xuat[r[0]]||0) + money_(r[3]); }); var m = {}; readProducts_().forEach(function(p){ var ton = p.tonDau + (nhap[p.ma]||0) - (xuat[p.ma]||0); m[p.ma] = { ton: ton, nhap: nhap[p.ma]||0, xuat: xuat[p.ma]||0 }; }); return m; }
function getProducts() { var st = stockMap_(); return readProducts_().map(function(p){ p.ton = st[p.ma] ? st[p.ma].ton : p.tonDau; p.nhap = st[p.ma] ? st[p.ma].nhap : 0; p.xuat = st[p.ma] ? st[p.ma].xuat : 0; return p; }); }
function getCustomers() { var pts = pointsMap_(); var debt = debtMap_(); var mua = {}; var dh = sh_('DonHang'); if (dh && dh.getLastRow()>=2) dh.getRange(2,1,dh.getLastRow()-1,5).getValues().forEach(function(r){ if (r[2]) mua[r[2]] = (mua[r[2]]||0) + money_(r[4]); }); return readCustomers_().map(function(c){ c.diem = pts[c.ma]||0; c.noHienTai = debt[c.ma]!==undefined ? debt[c.ma] : c.noDauKy; c.tongMua = mua[c.ma]||0; return c; }); }
function saveCustomer(c) { requireRole_(['Admin','BanHang','KeToan']); var s = sh_('KhachHang'); var v = s.getLastRow()>=4 ? s.getRange(4,1,s.getLastRow()-3,1).getValues() : []; c.ma = String(c.ma||'').trim(); if (!String(c.ten||'').trim()) throw new Error('Chưa nhập tên khách hàng.'); if (!c.ma) { var mx=0; v.forEach(function(r){ var mm=String(r[0]||'').match(/^KH(\d+)$/i); if (mm) { var nn=parseInt(mm[1],10); if (nn>mx) mx=nn; } }); var num=String(mx+1); while(num.length<3) num='0'+num; c.ma='KH'+num; } for (var i=0;i<v.length;i++){ if (v[i][0]===c.ma) { s.getRange(i+4,1,1,9).setValues([[c.ma,c.ten,c.loai||'',c.sdt||'',c.tinh||'',money_(c.hanMuc),money_(c.kyHan)||30,money_(c.noDauKy),c.ghiChu||'']]); return getCustomers(); } } var row = 4 + v.filter(function(r){return r[0];}).length; if (row>103) throw new Error('Đã đủ 100 khách trong khung file gốc. Hãy liên hệ để mở rộng.'); s.getRange(row,1,1,9).setValues([[c.ma,c.ten,c.loai||'',c.sdt||'',c.tinh||'',money_(c.hanMuc),money_(c.kyHan)||30,money_(c.noDauKy),c.ghiChu||'']]); return getCustomers(); }
function saveProduct(p) { requireRole_(['Admin','Kho']);
  if (!p.ma) { // để trống mã -> tự cấp theo tiền tố hãng/nhóm
    var preMap = { 'Nippon':'NIP', 'Jotun':'JOT', 'Dulux':'DUL', 'TOA':'TOA', 'Phụ kiện':'PK', 'Vật tư':'VT', 'Thiết bị':'VT' };
    var pre = preMap[p.nhom] || 'SP'; var ss0 = sh_('SanPham'); var vv = ss0.getLastRow()>=4 ? ss0.getRange(4,1,ss0.getLastRow()-3,1).getValues() : []; var mx = 0;
    vv.forEach(function(r){ var mm = String(r[0]||'').match(new RegExp('^' + pre + '(\\d+)$','i')); if (mm) { var nn = parseInt(mm[1],10); if (nn>mx) mx=nn; } });
    var num = String(mx+1); while (num.length<3) num='0'+num; p.ma = pre + num;
  }
  var s = sh_('SanPham'); var v = s.getLastRow()>=4 ? s.getRange(4,1,s.getLastRow()-3,1).getValues() : []; var row = [p.ma,p.nhom,p.ten,p.dvt,money_(p.gia),money_(p.tonDau),money_(p.tonMin)||5,p.hinh||'',money_(p.giaVon)]; for (var i=0;i<v.length;i++){ if (v[i][0]===p.ma) { s.getRange(i+4,1,1,9).setValues([row]); return getProducts(); } } var rowN = 4 + v.filter(function(r){return r[0];}).length; if (rowN>103) throw new Error('Đã đủ 100 sản phẩm trong khung file gốc.'); s.getRange(rowN,1,1,9).setValues([row]); return getProducts(); }

function deleteProduct(ma) { requireRole_(['Admin','Kho']); var s = sh_('SanPham'); var v = s.getLastRow()>=4 ? s.getRange(4,1,s.getLastRow()-3,1).getValues() : []; for (var i=0;i<v.length;i++){ if (String(v[i][0])===String(ma)) { s.deleteRow(i+4); return getProducts(); } } throw new Error('Không tìm thấy sản phẩm ' + ma); }

/* ============ CÔNG NỢ + ĐIỂM ============ */
function debtMap_() { var s = sh_('TongHopCongNo'); var m = {}; if (s && s.getLastRow()>=4) s.getRange(4,1,s.getLastRow()-3,8).getValues().forEach(function(r){ if(r[0]) m[r[0]] = money_(r[7]); }); return m; }
function getDebt() { var s = sh_('TongHopCongNo'); if (!s || s.getLastRow()<4) return []; return s.getRange(4,1,s.getLastRow()-3,16).getValues().filter(function(r){return r[0];}).map(function(r){ return { ma:r[0], ten:r[1], loai:r[2], hanMuc:money_(r[3]), noDau:money_(r[4]), ban:money_(r[5]), thu:money_(r[6]), noCuoi:money_(r[7]), trongHan:money_(r[8]), qh1_30:money_(r[9]), qh31_60:money_(r[10]), qh61_90:money_(r[11]), qh90:money_(r[12]), tongQH:money_(r[13]), canhBao:r[15]||'' }; }); }
function pointsMap_() { var s = sh_('DiemTichLuy'); var m = {}; if (s && s.getLastRow()>=2) s.getRange(2,2,s.getLastRow()-1,6).getValues().forEach(function(r){ if(r[0]) m[r[0]] = (m[r[0]]||0) + money_(r[4]) - money_(r[5]); }); return m; }
function getPoints() { var s = sh_('DiemTichLuy'); var hist = []; if (s && s.getLastRow()>=2) hist = s.getRange(2,1,s.getLastRow()-1,9).getValues().filter(function(r){return r[0];}).map(function(r){ return { ngay:fmtDate_(r[0]), maKH:r[1], tenKH:r[2], soHD:r[3], tienHD:money_(r[4]), diemCong:money_(r[5]), diemTru:money_(r[6]), ghiChu:r[8]||'' }; }).reverse(); var pts = pointsMap_(); var byKH = readCustomers_().map(function(c){ return { ma:c.ma, ten:c.ten, sdt:c.sdt, diem: pts[c.ma]||0 }; }).sort(function(a,b){return b.diem-a.diem;}); return { byKH: byKH, history: hist.slice(0,200) }; }

/* ============ SỐ CHỨNG TỪ TIẾP THEO ============ */
function nextNo_(sheetName, col, prefix) { var s = sh_(sheetName); var yy = Utilities.formatDate(new Date(), TZ, 'yy'); var pre = prefix + yy + '-'; var max = 0; if (s && s.getLastRow()>=4) { s.getRange(4,col,Math.max(1,s.getLastRow()-3),1).getValues().forEach(function(r){ var v = String(r[0]||''); if (v.indexOf(pre)===0) { var n = parseInt(v.substring(pre.length),10); if (!isNaN(n) && n>max) max=n; } }); } else if (s && sheetName==='ThuTien') {} var num = String(max+1); while(num.length<4) num='0'+num; return pre+num; }
function nextThuNo_() { var s = sh_('ThuTien'); var yy = Utilities.formatDate(new Date(), TZ, 'yy'); var pre = 'PT'+yy+'-'; var max=0; if (s && s.getLastRow()>=4) s.getRange(4,2,s.getLastRow()-3,1).getValues().forEach(function(r){ var v=String(r[0]||''); if(v.indexOf(pre)===0){var n=parseInt(v.substring(pre.length),10); if(n>max)max=n;} }); var num=String(max+1); while(num.length<4)num='0'+num; return pre+num; }
function firstEmptyRow_(sheetName, colLetter, startRow, maxRow) { var s = sh_(sheetName); var v = s.getRange(startRow, colLetter==='B'?2:1, maxRow-startRow+1, 1).getValues(); for (var i=0;i<v.length;i++){ if (!v[i][0]) return startRow+i; } throw new Error(sheetName + ' đã hết dòng trống trong khung sẵn.'); }


/* ============ TẠO ĐƠN BÁN HÀNG ============ */
function getBootstrap() { return { config: getConfig(), user: getCurrentUser(), customers: getCustomers(), products: getProducts(), brands: BRANDS, units: UNITS, today: todayStr_(), nextInvoice: nextNo_('NhatKyBanHang', 2, 'HD') }; }

/* Tìm khách theo mã/tên; chưa có thì tự tạo với mã KH tự động (dùng khi lập đơn gõ tên khách) */
function resolveCustomer_(p) {
  var cus = readCustomers_();
  var ma = String(p.maKH||'').trim();
  if (ma && cus.some(function(c){return c.ma===ma;})) { p.maKH = ma; return ''; }
  var ten = String(p.tenKH||'').trim(); if (!ten) throw new Error('Chưa nhập tên khách hàng.');
  var found = null; cus.forEach(function(c){ if (String(c.ten||'').trim().toLowerCase() === ten.toLowerCase()) found = c; });
  if (found) { p.maKH = found.ma; return ''; }
  var s = sh_('KhachHang'); var v = s.getLastRow()>=4 ? s.getRange(4,1,s.getLastRow()-3,1).getValues() : [];
  var mx = 0; v.forEach(function(r){ var mm = String(r[0]||'').match(/^KH(\d+)$/i); if (mm) { var nn = parseInt(mm[1],10); if (nn>mx) mx=nn; } });
  var num = String(mx+1); while (num.length<3) num='0'+num; var maMoi = 'KH'+num;
  var row = 4 + v.filter(function(r){return r[0];}).length; if (row>103) throw new Error('Đã đủ 100 khách trong khung file gốc.');
  s.getRange(row,1,1,9).setValues([[maMoi, ten, String(p.loaiKH||'Khách lẻ'), String(p.sdtKH||''), '', 0, 30, 0, 'Tạo tự động khi lập đơn']]);
  p.maKH = maMoi; return maMoi;
}
function createInvoice(p) {
  requireRole_(['Admin','BanHang']);
  if (!p || (!p.maKH && !String(p.tenKH||'').trim())) throw new Error('Chưa nhập tên khách hàng.');
  if (!p.items || !p.items.length) throw new Error('Đơn chưa có mặt hàng nào.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var khMoi = resolveCustomer_(p); var kh = null; readCustomers_().forEach(function(c){ if(c.ma===p.maKH) kh=c; }); if (!kh) throw new Error('Không tìm thấy khách ' + p.maKH);
    var prods = {}; readProducts_().forEach(function(x){ prods[x.ma]=x; }); var st = stockMap_();
    var total = 0; p.items.forEach(function(it){ if (!prods[it.maSP]) throw new Error('Không có sản phẩm ' + it.maSP); var sl = money_(it.sl); if (sl<=0) throw new Error('Số lượng phải lớn hơn 0: ' + it.maSP); var gia = it.gia!=='' && it.gia!==undefined && it.gia!==null ? money_(it.gia) : prods[it.maSP].gia; var ck = pct_(it.ck); total += sl*gia*(1-ck); var ton = st[it.maSP] ? st[it.maSP].ton : prods[it.maSP].tonDau; if (!p.choVuotTon && ton < sl) throw new Error('Không đủ tồn: ' + prods[it.maSP].ten + ' còn ' + ton + ' ' + prods[it.maSP].dvt + ', đơn cần ' + sl); });
    var soHD = nextNo_('NhatKyBanHang', 2, 'HD'); var ngay = parseDate_(p.ngay);
    var s = sh_('NhatKyBanHang'); var row = firstEmptyRow_('NhatKyBanHang','B',4,1003);
    p.items.forEach(function(it, idx){ var r = row+idx; var giaDefault = prods[it.maSP].gia; var gia = it.gia!=='' && it.gia!==undefined && it.gia!==null ? money_(it.gia) : giaDefault;
      s.getRange(r,1).setValue(ngay); s.getRange(r,2).setValue(soHD); s.getRange(r,3).setValue(p.maKH); s.getRange(r,5).setValue(it.maSP); s.getRange(r,8).setValue(money_(it.sl)); s.getRange(r,10).setValue(pct_(it.ck)); s.getRange(r,13).setValue(p.ghiChu||'');
      if (gia !== giaDefault) s.getRange(r,9).setValue(gia); // chỉ ghi đè giá khi giá bán khác giá niêm yết, còn lại để công thức tự lấy
    });
    SpreadsheetApp.flush();
    // Đơn hàng theo dõi
    var traNgay = Math.min(money_(p.traNgay), total); var dh = sh_('DonHang'); if (dh) dh.appendRow([soHD, ngay, p.maKH, kh.ten, total, traNgay, total-traNgay, p.trangThaiDon||'Mới', traNgay>=total && total>0 ? 'Đã thu đủ' : (traNgay>0 ? 'Thu 1 phần' : 'Chưa thu'), traNgay>0?1:0, new Date(), getCurrentUser().name, p.ghiChu||'']);
    // Trả ngay đợt 1 -> phiếu thu
    var soPT = ''; if (traNgay>0) { soPT = ghiThuTien_(ngay, p.maKH, traNgay, p.hinhThuc||'Tiền mặt', soHD, 'Thu đợt 1 - trả ngay khi xuất đơn ' + soHD); }
    // Tích điểm theo hoá đơn (có công tắc DiemBatTat trong CauHinh)
    var cfg = getConfig(); var diemBat = !(cfg.DiemBatTat===0 || cfg.DiemBatTat==='0' || cfg.DiemBatTat===false || cfg.DiemBatTat==='false'); var diem = diemBat ? Math.floor(total / (cfg.TienCho1Diem || 10000)) : 0; var ds = sh_('DiemTichLuy'); if (ds && diem>0) ds.appendRow([ngay, p.maKH, kh.ten, soHD, total, diem, 0, '', 'Tích điểm theo hoá đơn: ' + Math.round(cfg.TienCho1Diem).toLocaleString('vi-VN') + 'đ = 1 điểm']);
    return { soHD: soHD, soPT: soPT, tongTien: total, traNgay: traNgay, conNo: total-traNgay, diemCong: diem, tenKH: kh.ten, khMoi: khMoi };
  } finally { lock.releaseLock(); }
}
function ghiThuTien_(ngay, maKH, soTien, hinhThuc, soHD, ghiChu) { var s = sh_('ThuTien'); var row = firstEmptyRow_('ThuTien','B',4,2003); var soPT = nextThuNo_(); s.getRange(row,1).setValue(ngay); s.getRange(row,2).setValue(soPT); s.getRange(row,3).setValue(maKH); s.getRange(row,5).setValue(soTien); s.getRange(row,6).setValue(hinhThuc||'Tiền mặt'); s.getRange(row,7).setValue(soHD||''); s.getRange(row,8).setValue(ghiChu||''); SpreadsheetApp.flush(); return soPT; }

/* ============ THU TIỀN THEO ĐỢT 2-3 ĐỢT + CÔNG NỢ ============ */
var _payAllCache = null;
function allPayments_() { if (_payAllCache) return _payAllCache; var map = {}; var s = sh_('ThuTien'); if (s && s.getLastRow()>=4) s.getRange(4,1,s.getLastRow()-3,8).getValues().forEach(function(r){ if (!r[6]) return; var k = String(r[6]); if (!map[k]) map[k] = []; map[k].push({ ngay:fmtDate_(r[0]), soPT:r[1], maKH:r[2], tenKH:r[3], soTien:money_(r[4]), hinhThuc:r[5], soHD:r[6], ghiChu:r[7]||'', dot: map[k].length+1 }); }); _payAllCache = map; return map; }
function paymentsOfInvoice_(soHD) { var m = allPayments_(); return m[String(soHD)] || []; }
function addPayment(p) { // {soHD, maKH, soTien, hinhThuc, ngay, ghiChu}
  requireRole_(['Admin','BanHang','KeToan']); if (!p.soHD || !p.maKH) throw new Error('Thiếu số hoá đơn / khách hàng.'); var soTien = money_(p.soTien); if (soTien<=0) throw new Error('Số tiền thu phải lớn hơn 0.');
  var det = getInvoiceDetail(p.soHD); if (soTien > det.conNo + 1) throw new Error('Thu vượt số còn nợ của đơn: còn nợ ' + det.conNo.toLocaleString('vi-VN') + 'đ');
  var pays = paymentsOfInvoice_(p.soHD); var dot = pays.length + 1; if (dot > 5) throw new Error('Đơn này đã có 5 đợt thu, kiểm tra lại.');
  var soPT = ghiThuTien_(parseDate_(p.ngay), p.maKH, soTien, p.hinhThuc||'Tiền mặt', p.soHD, (p.ghiChu||'') + ' [Thu đợt ' + dot + ' của ' + p.soHD + ']');
  // cập nhật DonHang
  var dh = sh_('DonHang'); if (dh && dh.getLastRow()>=2) { var v = dh.getRange(2,1,dh.getLastRow()-1,10).getValues(); for (var i=0;i<v.length;i++){ if (String(v[i][0])===String(p.soHD)) { var daThu = money_(v[i][5]) + soTien; var tong = money_(v[i][4]); dh.getRange(i+2,6).setValue(daThu); dh.getRange(i+2,7).setValue(tong-daThu); dh.getRange(i+2,9).setValue(daThu>=tong ? 'Đã thu đủ' : 'Thu 1 phần'); dh.getRange(i+2,10).setValue(dot); break; } } }
  return { soPT: soPT, dot: dot, detail: getInvoiceDetail(p.soHD) };
}

/* ============ HOÀN ĐƠN HÀNG ============ */
function hoanDon(soHD, lyDo) {
  requireRole_(['Admin','BanHang']);
  soHD = String(soHD||'').trim(); if (!soHD) throw new Error('Thiếu số hoá đơn.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var dh = sh_('DonHang'); var dhRow = -1, dhStatus = '';
    if (dh && dh.getLastRow()>=2) { var vv = dh.getRange(2,1,dh.getLastRow()-1,13).getValues();
      for (var i=0;i<vv.length;i++){ if (String(vv[i][0])===soHD) { dhRow=i+2; dhStatus=String(vv[i][7]||''); break; } } }
    if (dhStatus==='Đã hoàn') throw new Error('Đơn ' + soHD + ' đã hoàn trước đó, không hoàn 2 lần.');
    var s = sh_('NhatKyBanHang'); var lines=[];
    if (s && s.getLastRow()>=4) s.getRange(4,1,s.getLastRow()-3,13).getValues().forEach(function(r){
      if (String(r[1])===soHD) lines.push({ ngay:r[0], maKH:r[2], tenKH:r[3], maSP:r[4], tenSP:r[5], dvt:r[6], sl:money_(r[7]), ck:pct_(r[9]), tt:money_(r[10]) });
    });
    if (!lines.length) throw new Error('Không tìm thấy dòng nào của đơn ' + soHD + ' trong NhatKyBanHang.');
    var goc = lines.filter(function(l){return l.sl>0;});
    if (!goc.length) throw new Error('Đơn ' + soHD + ' đã có dòng hoàn, không hoàn 2 lần.');
    var pays = paymentsOfInvoice_(soHD); var daThu = pays.reduce(function(a,b){return a+b.soTien;},0);
    var tongGoc = goc.reduce(function(a,l){return a+l.tt;},0);
    var ngayHoan = new Date();
    var note = 'Hoàn đơn ' + soHD + (lyDo ? ' - Lý do: ' + lyDo : '');
    if (daThu>0) note += '. Đã thu ' + Math.round(daThu).toLocaleString('vi-VN') + 'đ trước khi hoàn - kế toán xử lý hoàn tiền cho khách, KHÔNG tự xoá phiếu thu.';
    var r0 = firstEmptyRow_('NhatKyBanHang','B',4,1003);
    if (r0 + goc.length - 1 > 1003) throw new Error('Không đủ dòng trống trong NhatKyBanHang để ghi dòng hoàn.');
    goc.forEach(function(l, idx){
      var rr = r0+idx;
      s.getRange(rr,1).setValue(ngayHoan);   // A Ngày: chỉ ghi cột nhập tay
      s.getRange(rr,2).setValue(soHD);       // B Số HĐ
      s.getRange(rr,3).setValue(l.maKH);     // C Mã KH
      s.getRange(rr,5).setValue(l.maSP);     // E Mã SP
      s.getRange(rr,8).setValue(-l.sl);      // H Số lượng âm -> hoàn tồn, công thức cột K tự tính thành tiền âm
      s.getRange(rr,10).setValue(l.ck);      // J Chiết khấu như dòng gốc
      s.getRange(rr,13).setValue(note);      // M Ghi chú hoàn đơn
    });
    SpreadsheetApp.flush();
    // hoàn điểm: trừ lại số điểm đã cộng cho đơn này
    var diemDaCong = 0; var ds = sh_('DiemTichLuy');
    if (ds && ds.getLastRow()>=2) ds.getRange(2,4,ds.getLastRow()-1,3).getValues().forEach(function(r){ if (String(r[0])===soHD) diemDaCong += money_(r[2]); });
    if (ds && diemDaCong>0) ds.appendRow([ngayHoan, goc[0].maKH, goc[0].tenKH, soHD, 0, 0, Math.round(diemDaCong), '', 'Hoàn đơn ' + soHD + ' - trừ điểm đã cộng']);
    // cập nhật trạng thái DonHang
    var ghiDon = 'Hoàn ngày ' + fmtDate_(ngayHoan) + (lyDo ? ': ' + lyDo : '') + '. Đã thu ' + Math.round(daThu).toLocaleString('vi-VN') + 'đ trước đó.';
    if (dh) {
      if (dhRow>0) { dh.getRange(dhRow,8).setValue('Đã hoàn'); var gc = dh.getRange(dhRow,13).getValue(); dh.getRange(dhRow,13).setValue((gc?gc+' | ':'')+ghiDon); }
      else dh.appendRow([soHD, ngayHoan, goc[0].maKH, goc[0].tenKH, tongGoc, daThu, Math.max(0,tongGoc-daThu), 'Đã hoàn', daThu>0?'Thu 1 phần':'Chưa thu', pays.length, ngayHoan, getCurrentUser().name, ghiDon]);
    }
    return { soHD: soHD, daThu: daThu, diemHoan: Math.round(diemDaCong), soDongHoan: goc.length };
  } finally { lock.releaseLock(); }
}

/* ============ GIÁ VỐN BÌNH QUÂN + LỢI NHUẬN ============ */
function avgCostMap_() { // giá vốn bình quân gia quyền theo từng lần nhập trong NhapKho
  var m = {}; var n = sh_('NhapKho');
  if (n && n.getLastRow()>=2) n.getRange(2,2,n.getLastRow()-1,5).getValues().forEach(function(r){
    if (!r[0]) return; var ma=String(r[0]); var sl=money_(r[3]); var gia=money_(r[4]);
    if (!m[ma]) m[ma]={tien:0,sl:0}; m[ma].tien += sl*gia; m[ma].sl += sl;
  });
  var out={}; Object.keys(m).forEach(function(k){ out[k] = m[k].sl>0 ? m[k].tien/m[k].sl : 0; }); return out;
}
var _linesAllCache = null;
function allLines_() { if (_linesAllCache) return _linesAllCache; var map = {}; var s = sh_('NhatKyBanHang'); if (s && s.getLastRow()>=4) s.getRange(4,1,s.getLastRow()-3,11).getValues().forEach(function(r){ if (!r[1]) return; var k = String(r[1]); if (!map[k]) map[k] = []; map[k].push({ maSP:r[4], sl:money_(r[7]), tt:money_(r[10]) }); }); _linesAllCache = map; return map; }
function linesOfInvoice_(soHD) { var m = allLines_(); return m[String(soHD)] || []; }
function getProfit(from, to) {
  requireRole_(['Admin','BanHang','KeToan','Kho','Xem']);
  var costs = avgCostMap_(); var prodCost = {};
  readProducts_().forEach(function(p){ prodCost[p.ma] = p.giaVon||0; });
  var rows = []; var tD=0, tV=0, tL=0;
  groupInvoices_().forEach(function(o){
    if (o.trangThaiDon==='Đã hoàn') return; // đơn hoàn không tính vào lợi nhuận
    var d = parseDate_(o.ngay);
    if (from && d < parseDate_(from)) return;
    if (to) { var dd = parseDate_(to); dd.setHours(23,59,59,999); if (d > dd) return; }
    var doanhThu = 0, giaVon = 0;
    linesOfInvoice_(o.soHD).forEach(function(l){ if (l.sl<=0) return; var c = costs[l.maSP] || prodCost[l.maSP] || 0; doanhThu += l.tt; giaVon += l.sl*c; });
    var lai = doanhThu - giaVon;
    rows.push({ soHD:o.soHD, ngay:o.ngay, tenKH:o.tenKH, soMatHang:o.soMatHang, doanhThu:doanhThu, giaVon:giaVon, lai:lai });
    tD += doanhThu; tV += giaVon; tL += lai;
  });
  rows.sort(function(a,b){return String(b.soHD).localeCompare(String(a.soHD));});
  return { rows: rows, tongDoanhThu: tD, tongGiaVon: tV, tongLai: tL };
}

/* ============ ĐƠN HÀNG / LỊCH SỬ XUẤT / CHI TIẾT HOÁ ĐƠN ============ */
function groupInvoices_() { var s = sh_('NhatKyBanHang'); var map = {}; if (!s || s.getLastRow()<4) return []; s.getRange(4,1,s.getLastRow()-3,13).getValues().forEach(function(r){ if(!r[1]) return; var k = String(r[1]); if(!map[k]) map[k] = { soHD:k, ngay:fmtDate_(r[0]), ngayObj:r[0], maKH:r[2], tenKH:r[3], soMatHang:0, tongTien:0, hanTT:fmtDate_(r[11]), ghiChu:r[12]||'' }; map[k].soMatHang++; map[k].tongTien += money_(r[10]); }); var dhMap = {}; var dh = sh_('DonHang'); if (dh && dh.getLastRow()>=2) dh.getRange(2,1,dh.getLastRow()-1,10).getValues().forEach(function(r){ dhMap[String(r[0])] = { trangThaiDon:r[7]||'Mới', trangThaiTT:r[8]||'', daThu:money_(r[5]) }; }); return Object.keys(map).map(function(k){ var o = map[k]; var pays = paymentsOfInvoice_(k); var daThu2 = pays.reduce(function(a,b){return a+b.soTien;},0); o.daThu = dhMap[k] && dhMap[k].daThu ? dhMap[k].daThu : daThu2; o.conNo = Math.max(0, o.tongTien - o.daThu); o.soDotThu = pays.length; o.trangThaiDon = dhMap[k] ? dhMap[k].trangThaiDon : 'Mới'; o.trangThaiTT = o.conNo<=0 && o.tongTien>0 ? 'Đã thu đủ' : (o.daThu>0 ? 'Thu 1 phần' : 'Chưa thu'); return o; }).sort(function(a,b){ return String(b.soHD).localeCompare(String(a.soHD)); }); }
function getInvoices(f) { f = f || {}; var list = groupInvoices_(); if (f.tuNgay) list = list.filter(function(o){ return parseDate_(o.ngay) >= parseDate_(f.tuNgay); }); if (f.denNgay) list = list.filter(function(o){ return parseDate_(o.ngay) <= parseDate_(f.denNgay); }); if (f.maKH) list = list.filter(function(o){ return o.maKH===f.maKH; }); if (f.trangThai) list = list.filter(function(o){ return o.trangThaiDon===f.trangThai || o.trangThaiTT===f.trangThai; }); if (f.tim) { var q = String(f.tim).toLowerCase(); list = list.filter(function(o){ return o.soHD.toLowerCase().indexOf(q)>=0 || String(o.tenKH).toLowerCase().indexOf(q)>=0; }); } return list.slice(0, 300); }
function getOrders() { return getInvoices({}); }
function updateOrderStatus(soHD, status) { requireRole_(['Admin','BanHang','Kho']); var dh = sh_('DonHang'); if (!dh) throw new Error('Chưa có tab DonHang, hãy chạy cài đặt hệ thống.'); if (dh.getLastRow()>=2) { var v = dh.getRange(2,1,dh.getLastRow()-1,1).getValues(); for (var i=0;i<v.length;i++){ if (String(v[i][0])===String(soHD)) { dh.getRange(i+2,8).setValue(status); return true; } } } // đơn cũ chưa có trong DonHang -> tạo dòng theo dõi
  var inv = groupInvoices_().filter(function(o){return o.soHD===soHD;})[0]; if (inv) dh.appendRow([soHD, parseDate_(inv.ngay), inv.maKH, inv.tenKH, inv.tongTien, inv.daThu, inv.conNo, status, inv.trangThaiTT, inv.soDotThu, new Date(), getCurrentUser().name, '']); return true; }
function getInvoiceDetail(soHD) { var s = sh_('NhatKyBanHang'); var items = []; var head = null; var imgMap = {}; readProducts_().forEach(function(p){ imgMap[p.ma]=p.hinh; }); if (s && s.getLastRow()>=4) s.getRange(4,1,s.getLastRow()-3,13).getValues().forEach(function(r){ if (String(r[1])===String(soHD)) { if(!head) head = { soHD:soHD, ngay:fmtDate_(r[0]), maKH:r[2], tenKH:r[3], hanTT:fmtDate_(r[11]), ghiChu:r[12]||'' }; items.push({ maSP:r[4], tenSP:r[5], dvt:r[6], sl:money_(r[7]), gia:money_(r[8]), ck:pct_(r[9]), thanhTien:money_(r[10]), hinh: imgMap[r[4]]||'' }); } }); if (!head) throw new Error('Không tìm thấy hoá đơn ' + soHD); var tong = items.reduce(function(a,b){return a+b.thanhTien;},0); var pays = paymentsOfInvoice_(soHD); var daThu = pays.reduce(function(a,b){return a+b.soTien;},0); var kh = null; readCustomers_().forEach(function(c){ if(c.ma===head.maKH) kh=c; }); var cfg = getConfig(); var qr = ''; if (cfg.BankCode && cfg.SoTaiKhoan) qr = 'https://img.vietqr.io/image/' + encodeURIComponent(cfg.BankCode) + '-' + encodeURIComponent(cfg.SoTaiKhoan) + '-compact2.png?amount=' + Math.round(Math.max(0,tong-daThu)) + '&addInfo=' + encodeURIComponent('Thanh toan ' + soHD) + '&accountName=' + encodeURIComponent(cfg.ChuTaiKhoan||''); var diemBat = !(cfg.DiemBatTat===0 || cfg.DiemBatTat==='0' || cfg.DiemBatTat===false || cfg.DiemBatTat==='false'); return { head: head, items: items, tongTien: tong, payments: pays, daThu: daThu, conNo: Math.max(0,tong-daThu), customer: kh, config: cfg, qrUrl: qr, diemHD: diemBat ? Math.floor(tong / (cfg.TienCho1Diem||10000)) : 0 }; }

/* ============ TỒN KHO + NHẬP HÀNG ============ */
function getStock() { var cfg = getConfig(); return getProducts().map(function(p){ p.canNhap = p.ton <= (p.tonMin || cfg.TonCanhBao); return p; }); }
function addStockImport(p) { requireRole_(['Admin','Kho']); var s = sh_('NhapKho'); if (!s) { caiDatHeThong(); s = sh_('NhapKho'); } s.appendRow([parseDate_(p.ngay), p.maSP, p.tenSP||'', p.dvt||'', money_(p.sl), money_(p.giaNhap), money_(p.sl)*money_(p.giaNhap), p.ncc||'', p.ghiChu||'', getCurrentUser().name]); return getStock(); }
function getImportHistory() { requireRole_(['Admin','BanHang','Kho','KeToan','Xem']); var s = sh_('NhapKho'); var out = []; if (s && s.getLastRow()>=2) { var v = s.getRange(2,1,s.getLastRow()-1,10).getValues(); for (var i=v.length-1;i>=0 && out.length<100;i--) { var r=v[i]; if (!r[1]) continue; out.push({ ngay: fmtDate_(r[0]), maSP: r[1], tenSP: r[2]||'', dvt: r[3]||'', sl: money_(r[4]), giaNhap: money_(r[5]), thanhTien: money_(r[6]), ncc: r[7]||'', ghiChu: r[8]||'', nguoi: r[9]||'' }); } } return out; }
function importProducts(rows) { requireRole_(['Admin','Kho']); var ok=0, skip=[]; (rows||[]).forEach(function(r){ try { if (!r.ma || !r.ten) { skip.push(r.ma||'?'); return; } saveProduct(r); if (r.tonDau && money_(r.tonDau)>0 && r.khongNhapKho!=='co') {} ok++; } catch(e) { skip.push((r.ma||'?')+': '+e.message); } }); return { ok: ok, skip: skip, products: getProducts() }; }
function templateCsv() { var lines = ['MaSP,NhomHang,TenSanPham,DVT,DonGia,TonDauKy,TonToiThieu,HinhAnh,GiaVon']; ['NIP001,Nippon,Sơn nội thất Nippon Odour-less,5L,850000,20,5,https://cdn.horme.com.sg/images/product/202512091411431KVYGB15019587_700.jpg,680000','JOT001,Jotun,Sơn nội thất Jotun Essence,5L,780000,20,5,https://www.jotun.com/globalassetsjot03/211316-desktop.png,620000','DUL001,Dulux,Sơn nội thất Dulux Ambiance 5in1,5L,920000,20,5,https://sieuthison.thietkewebbanson.com/image/cache/catalog/san-pham/son-dulux/noi-that/dul032-600x600.png,740000','TOA001,TOA,Sơn nội thất TOA 4 Seasons,5L,650000,20,5,https://down-th.img.susercontent.com/file/th-11134207-7qul5-lg0akjqtnrwvc6,520000','PK001,Phụ kiện,Cọ sơn / rulo / khay sơn,cái,45000,100,20,https://www.homestratosphere.com/wp-content/uploads/2018/08/roller-tray-082218-870x870.jpg,30000'].forEach(function(x){lines.push(x);}); return lines.join('\n'); }

/* ============ THU CHI ============ */
function getCashbook() { var s = sh_('ThuChi'); if (!s || s.getLastRow()<2) return { rows: [], tongThu:0, tongChi:0 }; var rows = s.getRange(2,1,s.getLastRow()-1,8).getValues().filter(function(r){return r[0];}).map(function(r){ return { ngay:fmtDate_(r[0]), loai:r[1], hangMuc:r[2], soTien:money_(r[3]), phuongThuc:r[4]||'', lienQuan:r[5]||'', ghiChu:r[6]||'', nguoi:r[7]||'' }; }).reverse(); return { rows: rows.slice(0,300), tongThu: rows.filter(function(r){return r.loai==='Thu';}).reduce(function(a,b){return a+b.soTien;},0), tongChi: rows.filter(function(r){return r.loai==='Chi';}).reduce(function(a,b){return a+b.soTien;},0) }; }
function addCashEntry(p) { requireRole_(['Admin','KeToan']); var s = sh_('ThuChi'); if (!s) { caiDatHeThong(); s = sh_('ThuChi'); } s.appendRow([parseDate_(p.ngay), p.loai, p.hangMuc, money_(p.soTien), p.phuongThuc||'Tiền mặt', p.lienQuan||'', p.ghiChu||'', getCurrentUser().name]); return getCashbook(); }

/* ============ DASHBOARD ============ */
function mondayOf_(d){ var x=new Date(d.getFullYear(), d.getMonth(), d.getDate()); var day=(x.getDay()+6)%7; x.setDate(x.getDate()-day); return x; }
function getDashboard() {
  var inv = groupInvoices_();
  var now = new Date(); var today = Utilities.formatDate(now, TZ, 'dd/MM/yyyy');
  var mm = Utilities.formatDate(now, TZ, 'MM'); var yyyy = Utilities.formatDate(now, TZ, 'yyyy');
  function dOf(o){ var d = (o.ngayObj instanceof Date) ? o.ngayObj : parseDate_(o.ngay); return (d && !isNaN(d.getTime())) ? d : null; }
  function sameDay(d){ return d && Utilities.formatDate(d, TZ, 'dd/MM/yyyy')===today; }
  function sameMonth(d){ return d && Utilities.formatDate(d, TZ, 'MM/yyyy')===(mm+'/'+yyyy); }
  var invToday = inv.filter(function(o){ return sameDay(dOf(o)); });
  var invMonth = inv.filter(function(o){ return sameMonth(dOf(o)); });
  var doanhHomNay = invToday.reduce(function(a,b){return a+b.tongTien;},0);
  var doanhThang = invMonth.reduce(function(a,b){return a+b.tongTien;},0);
  var daThuHomNay = 0; var ts = sh_('ThuTien'); if (ts && ts.getLastRow()>=4) ts.getRange(4,1,ts.getLastRow()-3,5).getValues().forEach(function(r){ if (fmtDate_(r[0])===today) daThuHomNay += money_(r[4]); });
  var laiGop = 0; try { laiGop = getProfit(yyyy+'-'+mm+'-01', Utilities.formatDate(now, TZ, 'yyyy-MM-dd')).tongLai; } catch(e) {}
  var chiPhiThang = 0; getCashbook().rows.forEach(function(r){ var p = String(r.ngay||'').split('/'); if (p[1]===mm && p[2]===yyyy && r.loai==='Chi') chiPhiThang += r.soTien; });
  var debt = getDebt(); var tongNo = debt.reduce(function(a,b){return a+b.noCuoi;},0); var tongQH = debt.reduce(function(a,b){return a+b.tongQH;},0);
  var stock = getStock(); var canhBaoTon = stock.filter(function(p){return p.canNhap;});
  var cash = getCashbook(); var paysAll = inv.reduce(function(a,b){return a+b.daThu;},0);
  function key(d){ return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
  var days=[]; var dmap={}; for (var i=13;i>=0;i--){ var dd=new Date(now.getFullYear(), now.getMonth(), now.getDate()-i); var k=key(dd); dmap[k]=0; days.push({label: Utilities.formatDate(dd, TZ, 'dd/MM'), val: 0, k:k}); }
  var wmap={}; var weeks=[]; var mon = mondayOf_(now); for (var w=7;w>=0;w--){ var wd=new Date(mon.getFullYear(), mon.getMonth(), mon.getDate()-7*w); var wk=key(wd); wmap[wk]=0; weeks.push({label: Utilities.formatDate(wd, TZ, 'dd/MM'), val:0, k:wk}); }
  var mmap={}; var months=[]; for (var m2=1;m2<=12;m2++){ mmap[m2]=0; months.push({label:'T'+m2, val:0, m:m2}); }
  inv.forEach(function(o){ var d=dOf(o); if(!d) return; var k=key(d); if (dmap[k]!==undefined) dmap[k]+=o.tongTien; var mk=key(mondayOf_(d)); if (wmap[mk]!==undefined) wmap[mk]+=o.tongTien; if (d.getFullYear()===now.getFullYear()) mmap[d.getMonth()+1]+=o.tongTien; });
  days.forEach(function(x){ x.val=dmap[x.k]; delete x.k; }); weeks.forEach(function(x){ x.val=wmap[x.k]; delete x.k; }); months.forEach(function(x){ x.val=mmap[x.m]; delete x.m; });
  var bcMap={}; var prodMap={}; readProducts_().forEach(function(p){ prodMap[p.ma]=p; });
  invMonth.forEach(function(o){ linesOfInvoice_(o.soHD).forEach(function(l){ if (l.sl<=0) return; if(!bcMap[l.maSP]) bcMap[l.maSP]={ma:l.maSP, ten: prodMap[l.maSP]?prodMap[l.maSP].ten:l.maSP, dvt: prodMap[l.maSP]?prodMap[l.maSP].dvt:'', sl:0, doanh:0}; bcMap[l.maSP].sl+=l.sl; bcMap[l.maSP].doanh+=l.tt; }); });
  var banChay = Object.keys(bcMap).map(function(k){return bcMap[k];}).sort(function(a,b){ return b.sl-a.sl || b.doanh-a.doanh; }).slice(0,5);
  var donConNo = inv.filter(function(o){ return o.conNo>0 && o.trangThaiDon!=='Đã hoàn'; });
  var xuLy = { donConNoN: donConNo.length, donConNoTong: donConNo.reduce(function(a,b){return a+b.conNo;},0), khachNoN: debt.filter(function(d){return d.noCuoi>0;}).length, quaHanTong: tongQH, quaHanKH: debt.filter(function(d){return d.tongQH>0;}).length, donChuaGiao: inv.filter(function(o){ return o.trangThaiDon==='Mới'||o.trangThaiDon==='Đang giao'; }).length, spCanNhap: canhBaoTon.length, spAmKho: stock.filter(function(p){return p.ton<0;}).length };
  return { soDonHomNay: invToday.length, doanhHomNay: doanhHomNay, doanhThang: doanhThang, tongSoDon: inv.length, tongNo: tongNo, tongQuaHan: tongQH, soKHNoQuaHan: xuLy.quaHanKH, daThuTheoDon: paysAll, conNoTheoDon: inv.reduce(function(a,b){return a+b.conNo;},0), tongThuSach: cash.tongThu, tongChiSach: cash.tongChi, canhBaoTon: canhBaoTon, topNo: debt.slice().sort(function(a,b){return b.noCuoi-a.noCuoi;}).slice(0,10), donMoiNhat: inv.slice(0,8), tongSP: stock.length, tongTon: stock.reduce(function(a,b){return a+b.ton;},0), todayStr: today, monthStr: mm+'/'+yyyy, homNay: { soDon: invToday.length, doanh: doanhHomNay, daThu: daThuHomNay, conNo: invToday.reduce(function(a,b){return a+b.conNo;},0) }, thang: { doanh: doanhThang, laiGop: laiGop, chiPhi: chiPhiThang, loiNhuan: laiGop-chiPhiThang, soDon: invMonth.length }, xuLy: xuLy, rev: { days: days, weeks: weeks, months: months }, banChayThang: banChay, canDatThem: canhBaoTon.slice().sort(function(a,b){return a.ton-b.ton;}).slice(0,8) };
}

/* ============ BÁO CÁO XUẤT NHẬP TỒN THEO KỲ ============ */
function getXNT(tu, den) {
  requireRole_(['Admin','BanHang','Kho','KeToan','Xem']);
  var dTu = tu ? parseDate_(tu) : null; var dDen = den ? parseDate_(den) : new Date(); if (dDen) dDen.setHours(23,59,59,999);
  function inR(d){ return d && (!dTu || d>=dTu) && (!dDen || d<=dDen); }
  function bef(d){ return d && dTu && d<dTu; }
  function dOf(v){ var d = (v instanceof Date) ? v : parseDate_(fmtDate_(v)); return (d && !isNaN(d.getTime())) ? d : null; }
  var nT={}, nK={}, xT={}, xK={};
  var ns = sh_('NhapKho'); if (ns && ns.getLastRow()>=2) ns.getRange(2,1,ns.getLastRow()-1,5).getValues().forEach(function(r){ if(!r[1]) return; var d=dOf(r[0]); if(!d) return; var m=String(r[1]); if (bef(d)) nT[m]=(nT[m]||0)+money_(r[4]); else if (inR(d)) nK[m]=(nK[m]||0)+money_(r[4]); });
  var bs = sh_('NhatKyBanHang'); if (bs && bs.getLastRow()>=4) bs.getRange(4,1,bs.getLastRow()-3,8).getValues().forEach(function(r){ if(!r[4]) return; var d=dOf(r[0]); if(!d) return; var m=String(r[4]); var sl=money_(r[7]); if (bef(d)) xT[m]=(xT[m]||0)+sl; else if (inR(d)) xK[m]=(xK[m]||0)+sl; });
  var rows = readProducts_().map(function(p){ var td = p.tonDau + (nT[p.ma]||0) - (xT[p.ma]||0); var nk = nK[p.ma]||0; var xk = xK[p.ma]||0; return { ma:p.ma, nhom:p.nhom, ten:p.ten, dvt:p.dvt, tonDauKy:td, nhap:nk, xuat:xk, tonCuoi:td+nk-xk }; });
  return { rows: rows, tu: tu||'', den: den||'' };
}

/* ============ LỊCH SỬ MUA HÀNG THEO KHÁCH ============ */
function getCustomerHistory(maKH) { requireRole_(['Admin','BanHang','Kho','KeToan','Xem']); var kh=null; getCustomers().forEach(function(c){ if(c.ma===maKH) kh=c; }); var orders = groupInvoices_().filter(function(o){ return o.maKH===maKH; }); var tongSP = 0; var bs = sh_('NhatKyBanHang'); if (bs && bs.getLastRow()>=4) bs.getRange(4,3,bs.getLastRow()-3,6).getValues().forEach(function(r){ if (String(r[0])===String(maKH)) tongSP += money_(r[5]); }); return { kh: kh, orders: orders, tongSP: tongSP }; }
function deleteCustomer(ma) { requireRole_(['Admin']); if (String(ma)==='KH001') throw new Error('Không xoá được khách mặc định KH001 (Khách lẻ).'); var coDon = groupInvoices_().some(function(o){ return o.maKH===ma; }); if (coDon) throw new Error('Khách đã có đơn mua hàng — không xoá để tránh mất lịch sử.'); var kh=null; getCustomers().forEach(function(c){ if(c.ma===ma) kh=c; }); if (kh && (kh.diem>0 || kh.noHienTai>0)) throw new Error('Khách còn điểm hoặc còn nợ — không xoá được.'); var s = sh_('KhachHang'); var v = s.getLastRow()>=4 ? s.getRange(4,1,s.getLastRow()-3,1).getValues() : []; for (var i=0;i<v.length;i++){ if (String(v[i][0])===String(ma)) { s.getRange(i+4,1,1,9).clearContent(); return getCustomers(); } } throw new Error('Không tìm thấy khách ' + ma); }

/* ============ CÀI ĐẶT HỆ THỐNG + MENU ============ */
function caiDatHeThong() { var book = ss_();
  function ensure(name, headers, widths) { var s = book.getSheetByName(name); if (!s) s = book.insertSheet(name); if (s.getLastRow()===0) { s.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold'); s.setFrozenRows(1); if (widths) widths.forEach(function(w,i){ s.setColumnWidth(i+1,w); }); } return s; }
  ensure('CauHinh', ['Khoá','Giá trị'], [200,500]); var cfg = sh_('CauHinh'); if (cfg.getLastRow()<2) cfg.getRange(2,1,12,2).setValues([['TenCuaHang','Cửa hàng sơn nước Bảo Ân'],['DiaChi',''],['SDT',''],['MST',''],['BankCode',''],['TenNganHang',''],['SoTaiKhoan',''],['ChuTaiKhoan',''],['TienCho1Diem',10000],['TonCanhBao',5],['DiemBatTat',1],['GhiChu','BankCode là mã ngân hàng của VietQR, ví dụ: VCB, MB, ACB, TCB...']]);
  // nếu CauHinh đã có sẵn mà thiếu khoá DiemBatTat thì bổ sung
  var k2 = cfg.getRange(2,1,Math.max(1,cfg.getLastRow()-1),1).getValues().map(function(r){return String(r[0]);}); if (k2.indexOf('DiemBatTat')<0) cfg.appendRow(['DiemBatTat',1]);
  ensure('DonHang', ['Số HĐ','Ngày bán','Mã KH','Tên KH','Tổng tiền','Đã thu','Còn nợ','Trạng thái đơn','Trạng thái thanh toán','Số đợt thu','Ngày tạo','Người tạo','Ghi chú']);
  ensure('DiemTichLuy', ['Ngày','Mã KH','Tên KH','Số HĐ','Số tiền HĐ','Điểm cộng','Điểm trừ','Điểm sau','Ghi chú']);
  ensure('ThuChi', ['Ngày','Loại','Hạng mục','Số tiền','Phương thức','Liên quan','Ghi chú','Người thực hiện']);
  ensure('NhapKho', ['Ngày nhập','Mã SP','Tên SP','ĐVT','Số lượng','Đơn giá nhập','Thành tiền','Nhà cung cấp','Ghi chú','Người nhập']);
  ensure('NguoiDung', ['Email','Họ tên','Vai trò','Trạng thái','Ghi chú']);
  var sp = sh_('SanPham'); if (sp) { if (!sp.getRange(3,6).getValue()) sp.getRange(3,6).setValue('Tồn đầu kỳ'); if (!sp.getRange(3,7).getValue()) sp.getRange(3,7).setValue('Tồn tối thiểu'); if (!sp.getRange(3,9).getValue()) sp.getRange(3,9).setValue('Giá vốn'); }
  return 'Đã cài xong các tab phụ: CauHinh (kèm công tắc DiemBatTat), DonHang, DiemTichLuy, ThuChi, NhapKho, NguoiDung + cột Giá vốn (I) trong SanPham.';
}
function onOpen() { SpreadsheetApp.getUi().createMenu('🎨 Sơn Bảo Ân').addItem('⚙ Cài đặt hệ thống (tạo tab phụ)', 'caiDatHeThong').addItem('🌐 Mở hướng dẫn Web App', 'moHuongDan_').addToUi(); }
function moHuongDan_() { SpreadsheetApp.getUi().alert('Triển khai Web App: Tiện ích mở rộng → Apps Script → Dán Code.gs + Index.html → Deploy → Web App → Execute as Me → Who has access: Anyone with Google account (để lấy email phân quyền) → Lấy link Web App để dùng hằng ngày.'); }
