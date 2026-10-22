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

/* ===== API cho giao dien tinh (GitHub Pages / app dien thoai) =====
   Client tinh gui POST {fn, args} -> chay dung ham server cua ban cu -> tra JSON.
   Ban cu qua doGet van chay song song, khong doi. Chi cac ham trong API_MAP_ moi goi duoc qua API. */
var API_MAP_ = { dangNhap: dangNhap, dangXuat: dangXuat, doiMatKhau: doiMatKhau, getConfig: getConfig, saveConfig: saveConfig, getCurrentUser: getCurrentUser, getUsers: getUsers, saveUser: saveUser, deleteUser: deleteUser, getCustomers: getCustomers, saveCustomer: saveCustomer, deleteCustomer: deleteCustomer, getCustomerHistory: getCustomerHistory, getProducts: getProducts, saveProduct: saveProduct, deleteProduct: deleteProduct, getCatalog: getCatalog, saveCatalogBrand: saveCatalogBrand, deleteCatalogBrand: deleteCatalogBrand, saveCatalogMain: saveCatalogMain, deleteCatalogMain: deleteCatalogMain, getBootstrap: getBootstrap, createInvoice: createInvoice, updateInvoice: updateInvoice, hoanDon: hoanDon, addPayment: addPayment, getInvoices: getInvoices, getOrders: getOrders, getInvoiceDetail: getInvoiceDetail, getStock: getStock, addStockImport: addStockImport, getImportHistory: getImportHistory, importProducts: importProducts, templateCsv: templateCsv, getCashbook: getCashbook, addCashEntry: addCashEntry, getDashboard: getDashboard, getDebt: getDebt, getPoints: getPoints, getProfit: getProfit, getXNT: getXNT, updateOrderStatus: updateOrderStatus, inlineInvoiceImages: inlineInvoiceImages, getSuppliers: getSuppliers, saveSupplier: saveSupplier, deleteSupplier: deleteSupplier, getCareLogs: getCareLogs, saveCareLog: saveCareLog, deleteCareLog: deleteCareLog };
function doPost(e) {
  var out;
  try {
    var body = (e && e.postData && e.postData.contents) || '{}';
    var req = JSON.parse(body);
    var fn = String(req.fn || '');
    var f = API_MAP_[fn];
    if (!f) throw new Error('Hàm không được phép qua API: ' + fn);
    var data = f.apply(null, req.args || []);
    out = { ok: true, data: (data === undefined ? null : data) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sh_(name) { return ss_().getSheetByName(name); }
function cellVal_(v) { return (v && Object.prototype.toString.call(v) === '[object Date]') ? fmtDate_(v) : v; }
function strCell_(v) { var x = cellVal_(v); return (x === null || x === undefined) ? '' : String(x); }
function fmtDate_(d) { if (!d) return ''; if (Object.prototype.toString.call(d) === '[object Date]') return Utilities.formatDate(d, TZ, 'dd/MM/yyyy'); return String(d); }
function parseDate_(s) { if (!s) return new Date(); if (Object.prototype.toString.call(s) === '[object Date]') return s; var p = String(s).split(/[-/]/); if (p.length === 3 && p[0].length === 4) return new Date(+p[0], +p[1]-1, +p[2]); if (p.length === 3) return new Date(+p[2], +p[1]-1, +p[0]); return new Date(s); }
function money_(v) { if (v === '' || v === null || v === undefined || v === '-') return 0; if (typeof v === 'number') return v; var n = parseFloat(String(v).replace(/[^\d,-]/g, '').replace(/\./g, '').replace(',', '.')); return isNaN(n) ? 0 : n; }
function pct_(v) { if (typeof v === 'number') return v <= 1 ? v : v/100; var n = parseFloat(String(v||'0').replace('%','').replace(',','.')); return isNaN(n) ? 0 : n/100; }
function todayStr_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

/* ============ CẤU HÌNH ============ */

/* ============ DANG NHAP PHAN MEM (tai khoan + mat khau, phien luu Cache 6 gio) ============ */
var _PHIEN = null;
function _auth(tk){ if (_PHIEN) return _PHIEN; if (!tk || typeof tk !== 'string') throw new Error('Bạn chưa đăng nhập (hoặc phiên đã hết hạn). Hãy đăng nhập lại.'); var c = CacheService.getScriptCache().get('sess_' + tk); if (!c) throw new Error('Phiên đăng nhập đã hết hạn. Hãy đăng nhập lại.'); _PHIEN = JSON.parse(c); return _PHIEN; }
function bamMatKhau_(u, p){ var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'BAOAN#' + u + '#' + p); return raw.map(function(b){ return ('0' + (b & 0xff).toString(16)).slice(-2); }).join(''); }
function taiKhoanSheet_(){ var s = sh_('TaiKhoan'); if (!s) { s = ss_().insertSheet('TaiKhoan'); s.getRange(1,1,1,5).setValues([['TenDangNhap','MatKhauBam','TenHienThi','VaiTro','TrangThai']]); try { s.hideSheet(); } catch(e) {} } if (s.getLastRow() < 2) { s.appendRow(['baoan', bamMatKhau_('baoan','1234'), 'Admin Bảo Ân', 'Admin', 'Hoạt động']); } return s; }
function dangNhap(u, p){ u = String(u||'').trim().toLowerCase(); p = String(p||''); if (!u || !p) throw new Error('Nhập tài khoản và mật khẩu.'); var s = taiKhoanSheet_(); var v = s.getRange(2,1,Math.max(1,s.getLastRow()-1),5).getValues(); for (var i=0;i<v.length;i++){ if (String(v[i][0]).toLowerCase() === u) { if (String(v[i][4]||'Hoạt động') !== 'Hoạt động') throw new Error('Tài khoản đã bị khoá.'); if (String(v[i][1]) !== bamMatKhau_(u, p)) throw new Error('Sai mật khẩu.'); var token = Utilities.getUuid(); var phien = { u:u, ten:v[i][2]||u, role:v[i][3]||'Xem' }; CacheService.getScriptCache().put('sess_' + token, JSON.stringify(phien), 21600); return { token:token, ten:phien.ten, role:phien.role }; } } throw new Error('Tài khoản không tồn tại.'); }
function dangXuat(tk){ try { if (tk) CacheService.getScriptCache().remove('sess_' + tk); } catch(e) {} return true; }
function doiMatKhau(tk, cu, moi){ var ph = _auth(tk); moi = String(moi||''); if (moi.length < 4) throw new Error('Mật khẩu mới phải có ít nhất 4 ký tự.'); var s = taiKhoanSheet_(); var v = s.getRange(2,1,Math.max(1,s.getLastRow()-1),5).getValues(); for (var i=0;i<v.length;i++){ if (String(v[i][0]).toLowerCase() === ph.u) { if (String(v[i][1]) !== bamMatKhau_(ph.u, String(cu||''))) throw new Error('Mật khẩu cũ chưa đúng.'); s.getRange(i+2,2).setValue(bamMatKhau_(ph.u, moi)); return true; } } throw new Error('Không tìm thấy tài khoản.'); }

function getConfig(tk) { _auth(tk);
  var s = sh_('CauHinh'); var cfg = { TenCuaHang: 'Cửa hàng sơn nước Bảo Ân', DiaChi:'', SDT:'', MST:'', BankCode:'', TenNganHang:'', SoTaiKhoan:'', ChuTaiKhoan:'', TienCho1Diem: 10000, TonCanhBao: 5, DiemBatTat: 1, LogoUrl:'', Slogan:'Nơi đặt trọn niềm tin', HD_LoiCamOn:'Cảm ơn quý khách đã tin tưởng Bảo Ân!', HD_HienTenCH:1, HD_HienLogo:1, HD_HienSlogan:1, HD_HienDiaChi:1, HD_HienSDT:1, HD_HienMST:1, HD_HienSDTKH:1, HD_HienHanTT:1, HD_HienQR:1, HD_HienDiem:1, HD_HienChietKhau:1, HD_HienDonGia:1, HD_HienHinhSP:1, HD_HienDaThuConNo:1, HD_HienGhiChu:1, HD_HienChuKy:1, HD_HienTenKyNB:1, HD_HienTenKyKH:1, ThemeMau:'cam' };
  if (!s) return cfg;
  var v = s.getRange(2, 1, Math.max(1, s.getLastRow()-1), 2).getValues();
  v.forEach(function(r){ if (r[0]) cfg[String(r[0]).trim()] = cellVal_(r[1]); });
  cfg.TienCho1Diem = money_(cfg.TienCho1Diem) || 10000; cfg.TonCanhBao = money_(cfg.TonCanhBao) || 5;
  return cfg;
}
function saveConfig(tk, c) { _auth(tk); requireRole_(['Admin']); var s = sh_('CauHinh'); if (!s) { caiDatHeThong(); s = sh_('CauHinh'); } var merged = {}; if (s.getLastRow()>=2) s.getRange(2,1,s.getLastRow()-1,2).getValues().forEach(function(r){ if (r[0]) merged[String(r[0])] = r[1]; }); Object.keys(c||{}).forEach(function(k){ merged[k] = c[k]; }); var keys = Object.keys(merged); var rows = keys.map(function(k){ return [k, merged[k]]; }); s.getRange(2, 1, Math.max(rows.length, s.getLastRow()), 2).clearContent(); s.getRange(2, 1, rows.length, 2).setValues(rows); return getConfig(); }

/* ============ PHÂN QUYỀN ============ */
function getCurrentUser(tk) { if (tk) _auth(tk); if (_PHIEN) { var pm = { Admin:['*'], BanHang:['dashboard','taodon','donhang','lichsu','khachhang','congno','diem','loinhuan'], Kho:['dashboard','donhang','lichsu','tonkho','hanghoa'], KeToan:['dashboard','donhang','lichsu','khachhang','congno','thuchi','diem','loinhuan'], Xem:['dashboard','donhang','lichsu','tonkho','khachhang','congno','diem','loinhuan'] }; return { email:_PHIEN.u, name:_PHIEN.ten, role:_PHIEN.role, active:true, perms: pm[_PHIEN.role] || [] }; }
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
function getUsers(tk) { _auth(tk); requireRole_(['Admin']); var s = sh_('NguoiDung'); if (!s || s.getLastRow() < 2) return []; return s.getRange(2,1,s.getLastRow()-1,5).getValues().filter(function(r){return r[0];}).map(function(r){ return { email:r[0], name:r[1], role:r[2], status:r[3]||'Hoạt động', note:r[4]||'' }; }); }
function saveUser(tk, u) { _auth(tk); requireRole_(['Admin']); var s = sh_('NguoiDung'); if (!s) { caiDatHeThong(); s = sh_('NguoiDung'); } var v = s.getLastRow()>=2 ? s.getRange(2,1,s.getLastRow()-1,1).getValues() : []; for (var i=0;i<v.length;i++){ if (String(v[i][0]).toLowerCase() === String(u.email).toLowerCase()) { s.getRange(i+2,1,1,5).setValues([[u.email,u.name,u.role,u.status||'Hoạt động',u.note||'']]); return getUsers(); } } s.appendRow([u.email,u.name,u.role,u.status||'Hoạt động',u.note||'']); return getUsers(); }
function deleteUser(tk, email) { _auth(tk); requireRole_(['Admin']); var s = sh_('NguoiDung'); var v = s.getRange(2,1,s.getLastRow()-1,1).getValues(); for (var i=0;i<v.length;i++){ if (String(v[i][0]).toLowerCase()===String(email).toLowerCase()) { s.deleteRow(i+2); break; } } return getUsers(); }

/* ============ DỮ LIỆU GỐC ============ */
function readCustomers_() { var s = sh_('KhachHang'); if (!s || s.getLastRow()<4) return []; return s.getRange(4,1,s.getLastRow()-3,9).getValues().filter(function(r){return r[0];}).map(function(r){ return { ma:strCell_(r[0]), ten:strCell_(r[1]), loai:strCell_(r[2]), sdt:String(r[3]||''), tinh:strCell_(r[4]), hanMuc:money_(r[5]), kyHan:money_(r[6]), noDauKy:money_(r[7]), ghiChu:strCell_(r[8]) }; }); }
function readProducts_() { var s = sh_('SanPham'); if (!s || s.getLastRow()<4) return []; var last = s.getLastRow(); var vals = s.getRange(4,1,last-3,10).getValues(); return vals.filter(function(r){return r[0];}).map(function(r){ var bt=[]; try{ bt=JSON.parse(r[9]||'[]'); if(!Array.isArray(bt))bt=[]; }catch(e){} return { ma:strCell_(r[0]), nhom:strCell_(r[1]), ten:strCell_(r[2]), dvt:strCell_(r[3]), gia:money_(r[4]), tonDau:money_(r[5]), tonMin:money_(r[6]), hinh:String(r[7]||''), giaVon:money_(r[8]), bienThe:bt }; }); }
function stockMap_(prodList) { var nhap = {}, xuat = {}; var n = sh_('NhapKho'); if (n && n.getLastRow()>=2) n.getRange(2,2,n.getLastRow()-1,4).getValues().forEach(function(r){ if(r[0]) nhap[r[0]] = (nhap[r[0]]||0) + money_(r[3]); }); var s = sh_('NhatKyBanHang'); if (s && s.getLastRow()>=4) s.getRange(4,5,s.getLastRow()-3,4).getValues().forEach(function(r){ if(r[0]) xuat[r[0]] = (xuat[r[0]]||0) + money_(r[3]); }); var m = {}; (prodList || readProducts_()).forEach(function(p){ var ton = p.tonDau + (nhap[p.ma]||0) - (xuat[p.ma]||0); m[p.ma] = { ton: ton, nhap: nhap[p.ma]||0, xuat: xuat[p.ma]||0 }; }); return m; }
function getProducts(tk) { _auth(tk); var st = stockMap_(); return readProducts_().map(function(p){ p.ton = st[p.ma] ? st[p.ma].ton : p.tonDau; p.nhap = st[p.ma] ? st[p.ma].nhap : 0; p.xuat = st[p.ma] ? st[p.ma].xuat : 0; return p; }); }
function getCustomers(tk) { _auth(tk); var pts = pointsMap_(); var debt = debtMap_(); var mua = {}; var dh = sh_('DonHang'); if (dh && dh.getLastRow()>=2) dh.getRange(2,1,dh.getLastRow()-1,5).getValues().forEach(function(r){ if (r[2]) mua[String(r[2])] = (mua[String(r[2])]||0) + money_(r[4]); }); return readCustomers_().map(function(c){ c.diem = pts[c.ma]||0; c.noHienTai = debt[c.ma]!==undefined ? debt[c.ma] : c.noDauKy; c.tongMua = mua[c.ma]||0; return c; }); }
function saveCustomer_goc(tk, c) { _auth(tk); requireRole_(['Admin','BanHang','KeToan']); var s = sh_('KhachHang'); var v = s.getLastRow()>=4 ? s.getRange(4,1,s.getLastRow()-3,1).getValues() : []; c.ma = String(c.ma||'').trim(); if (!String(c.ten||'').trim()) throw new Error('Chưa nhập tên khách hàng.'); if (!c.ma) { var mx=0; v.forEach(function(r){ var mm=String(r[0]||'').match(/^KH(\d+)$/i); if (mm) { var nn=parseInt(mm[1],10); if (nn>mx) mx=nn; } }); var num=String(mx+1); while(num.length<3) num='0'+num; c.ma='KH'+num; } for (var i=0;i<v.length;i++){ if (v[i][0]===c.ma) { s.getRange(i+4,1,1,9).setValues([[c.ma,c.ten,c.loai||'',c.sdt||'',c.tinh||'',money_(c.hanMuc),money_(c.kyHan)||30,money_(c.noDauKy),c.ghiChu||'']]); return getCustomers(); } } var row = 0; for (var qi=0; qi<v.length; qi++){ if (!v[qi][0]) { row = 4+qi; break; } } if (!row) row = 4 + v.length; if (row>103) throw new Error('Đã đủ 100 khách trong khung file gốc. Hãy liên hệ để mở rộng.'); s.getRange(row,1,1,9).setValues([[c.ma,c.ten,c.loai||'',c.sdt||'',c.tinh||'',money_(c.hanMuc),money_(c.kyHan)||30,money_(c.noDauKy),c.ghiChu||'']]); return getCustomers(); }
function saveProduct_goc(tk, p) { _auth(tk); requireRole_(['Admin','Kho']);
  if (!p.ma) { // để trống mã -> tự cấp theo tiền tố hãng/nhóm
    var preMap = { 'Nippon':'NIP', 'Jotun':'JOT', 'Dulux':'DUL', 'TOA':'TOA', 'Phụ kiện':'PK', 'Vật tư':'VT', 'Thiết bị':'VT' };
    var pre = preMap[String(p.nhom||'').split(' - ')[0]] || 'SP'; var ss0 = sh_('SanPham'); var vv = ss0.getLastRow()>=4 ? ss0.getRange(4,1,ss0.getLastRow()-3,1).getValues() : []; var mx = 0;
    vv.forEach(function(r){ var mm = String(r[0]||'').match(new RegExp('^' + pre + '(\\d+)$','i')); if (mm) { var nn = parseInt(mm[1],10); if (nn>mx) mx=nn; } });
    var num = String(mx+1); while (num.length<3) num='0'+num; p.ma = pre + num;
  }
  var s = sh_('SanPham'); var v = s.getLastRow()>=4 ? s.getRange(4,1,s.getLastRow()-3,1).getValues() : []; var btJson=''; try{ btJson=JSON.stringify(p.bienThe||[]); }catch(e){} var row = [p.ma,p.nhom,p.ten,p.dvt,money_(p.gia),money_(p.tonDau),money_(p.tonMin)||5,p.hinh||'',money_(p.giaVon),btJson]; for (var i=0;i<v.length;i++){ if (v[i][0]===p.ma) { s.getRange(i+4,1,1,10).setValues([row]); return getProducts(); } } var rowN = 0; for (var qi=0; qi<v.length; qi++){ if (!v[qi][0]) { rowN = 4+qi; break; } } if (!rowN) rowN = 4 + v.length; if (rowN>103) throw new Error('Đã đủ 100 sản phẩm trong khung file gốc.'); s.getRange(rowN,1,1,10).setValues([row]); return getProducts(); }

function deleteProduct_goc(tk, ma) { _auth(tk); requireRole_(['Admin','Kho']); var nGD = 0; var nkS = sh_('NhatKyBanHang'); if (nkS && nkS.getLastRow()>=4) nkS.getRange(4,5,nkS.getLastRow()-3,1).getValues().forEach(function(r){ if (String(r[0])===String(ma)) nGD++; }); var nkN = sh_('NhapKho'); if (nkN && nkN.getLastRow()>=2) nkN.getRange(2,2,nkN.getLastRow()-1,1).getValues().forEach(function(r){ if (String(r[0])===String(ma)) nGD++; }); if (nGD>0) throw new Error('Không xoá được: sản phẩm ' + ma + ' đã có ' + nGD + ' dòng bán/nhập kho. Xoá sẽ làm hỏng hoá đơn cũ — hãy sửa tên/giá thay vì xoá.'); var s = sh_('SanPham'); var v = s.getLastRow()>=4 ? s.getRange(4,1,s.getLastRow()-3,1).getValues() : []; for (var i=0;i<v.length;i++){ if (String(v[i][0])===String(ma)) { s.deleteRow(i+4); return getProducts(); } } throw new Error('Không tìm thấy sản phẩm ' + ma); }

/* ============ CÔNG NỢ + ĐIỂM ============ */
function debtMap_() { var m = {}; readCustomers_().forEach(function(c){ m[c.ma] = c.noDauKy || 0; }); var dh = sh_('DonHang'); if (dh && dh.getLastRow()>=2) dh.getRange(2,1,dh.getLastRow()-1,7).getValues().forEach(function(r){ if (r[2]) { var k = strCell_(r[2]); if (m[k]===undefined) m[k] = 0; m[k] += money_(r[4]) - money_(r[5]); } }); return m; }
function getDebt(tk) { _auth(tk); var s = sh_('TongHopCongNo'); var base = []; if (s && s.getLastRow()>=4) base = s.getRange(4,1,s.getLastRow()-3,16).getValues().filter(function(r){return r[0];}).map(function(r){ return { ma:strCell_(r[0]), ten:strCell_(r[1]), loai:r[2], hanMuc:money_(r[3]), noDau:money_(r[4]), ban:money_(r[5]), thu:money_(r[6]), noCuoi:money_(r[7]), trongHan:money_(r[8]), qh1_30:money_(r[9]), qh31_60:money_(r[10]), qh61_90:money_(r[11]), qh90:money_(r[12]), tongQH:money_(r[13]), canhBao:r[15]||'' }; });
  var agg = {}; var dh = sh_('DonHang'); if (dh && dh.getLastRow()>=2) dh.getRange(2,1,dh.getLastRow()-1,7).getValues().forEach(function(r){ if (r[2]) { var k = strCell_(r[2]); if (!agg[k]) agg[k] = { ban:0, thu:0 }; agg[k].ban += money_(r[4]); agg[k].thu += money_(r[5]); } });
  var seen = {}; base.forEach(function(d){ seen[d.ma] = 1; var a = agg[d.ma]; if (a) { d.ban = a.ban; d.thu = a.thu; d.noCuoi = d.noDau + a.ban - a.thu; } });
  readCustomers_().forEach(function(c){ if (!seen[c.ma] && agg[c.ma]) base.push({ ma:c.ma, ten:c.ten, loai:c.loai, hanMuc:c.hanMuc, noDau:c.noDauKy, ban:agg[c.ma].ban, thu:agg[c.ma].thu, noCuoi:c.noDauKy + agg[c.ma].ban - agg[c.ma].thu, trongHan:0, qh1_30:0, qh31_60:0, qh61_90:0, qh90:0, tongQH:0, canhBao:'' }); });
  return base; }
function pointsMap_() { var s = sh_('DiemTichLuy'); var m = {}; if (s && s.getLastRow()>=2) s.getRange(2,2,s.getLastRow()-1,6).getValues().forEach(function(r){ if(r[0]) m[strCell_(r[0])] = (m[strCell_(r[0])]||0) + money_(r[4]) - money_(r[5]); }); return m; }
function getPoints(tk) { _auth(tk); var s = sh_('DiemTichLuy'); var hist = []; if (s && s.getLastRow()>=2) hist = s.getRange(2,1,s.getLastRow()-1,9).getValues().filter(function(r){return r[0];}).map(function(r){ return { ngay:fmtDate_(r[0]), maKH:strCell_(r[1]), tenKH:strCell_(r[2]), soHD:r[3], tienHD:money_(r[4]), diemCong:money_(r[5]), diemTru:money_(r[6]), ghiChu:r[8]||'' }; }).reverse(); var pts = pointsMap_(); var byKH = readCustomers_().map(function(c){ return { ma:c.ma, ten:c.ten, sdt:c.sdt, diem: pts[c.ma]||0 }; }).sort(function(a,b){return b.diem-a.diem;}); return { byKH: byKH, history: hist.slice(0,200) }; }

/* ============ SỐ CHỨNG TỪ TIẾP THEO ============ */
function nextNo_(sheetName, col, prefix) { var s = sh_(sheetName); var yy = Utilities.formatDate(new Date(), TZ, 'yy'); var pre = prefix + yy + '-'; var max = 0; if (s && s.getLastRow()>=4) { s.getRange(4,col,Math.max(1,s.getLastRow()-3),1).getValues().forEach(function(r){ var v = String(r[0]||''); if (v.indexOf(pre)===0) { var n = parseInt(v.substring(pre.length),10); if (!isNaN(n) && n>max) max=n; } }); } else if (s && sheetName==='ThuTien') {} var num = String(max+1); while(num.length<4) num='0'+num; return pre+num; }
function nextThuNo_() { var s = sh_('ThuTien'); var yy = Utilities.formatDate(new Date(), TZ, 'yy'); var pre = 'PT'+yy+'-'; var max=0; if (s && s.getLastRow()>=4) s.getRange(4,2,s.getLastRow()-3,1).getValues().forEach(function(r){ var v=String(r[0]||''); if(v.indexOf(pre)===0){var n=parseInt(v.substring(pre.length),10); if(n>max)max=n;} }); var num=String(max+1); while(num.length<4)num='0'+num; return pre+num; }
function firstEmptyRow_(sheetName, colLetter, startRow, maxRow) { var s = sh_(sheetName); var cap = Math.min(maxRow, Math.max(startRow, s.getLastRow()+1)); var v = s.getRange(startRow, colLetter==='B'?2:1, cap-startRow+1, 1).getValues(); for (var i=0;i<v.length;i++){ if (!v[i][0]) return startRow+i; } throw new Error(sheetName + ' đã hết dòng trống trong khung sẵn.'); }


/* ============ TẠO ĐƠN BÁN HÀNG ============ */
function catalogSheet_() { var ss = SpreadsheetApp.getActiveSpreadsheet(); var s = ss.getSheetByName('DanhMuc'); if (!s) { s = ss.insertSheet('DanhMuc'); s.getRange(1,1,1,2).setValues([['Hang','SPChinh']]); } return s; }
function catalogRead_() { var s = catalogSheet_(); var out = { brands: [], mains: {} }; if (s.getLastRow() < 2) return out; s.getRange(2,1,s.getLastRow()-1,2).getValues().forEach(function(r){ var b = String(r[0]||'').trim(), m = String(r[1]||'').trim(); if (!b) return; if (out.brands.indexOf(b)<0) out.brands.push(b); if (m) { out.mains[b] = out.mains[b]||[]; if (out.mains[b].indexOf(m)<0) out.mains[b].push(m); } }); return out; }
function getCatalog(tk) { _auth(tk); return { catalog: catalogRead_() }; }
function saveCatalogBrand(tk, ten) { _auth(tk); requireRole_(['Admin','Kho']); ten = String(ten||'').trim(); if (!ten) throw new Error('Chưa nhập tên hãng.'); var c = catalogRead_(); var all = BRANDS.concat(c.brands); readProducts_().forEach(function(p){ var b = String(p.nhom||'').split(' - ')[0]; if (b) all.push(b); }); if (all.some(function(x){ return x.toLowerCase()===ten.toLowerCase(); })) throw new Error('Hãng "' + ten + '" đã có rồi.'); catalogSheet_().appendRow([ten, '']); return { catalog: catalogRead_() }; }
function deleteCatalogBrand(tk, ten) { _auth(tk); requireRole_(['Admin','Kho']); ten = String(ten||'').trim(); var n = 0; readProducts_().forEach(function(p){ if (String(p.nhom||'').split(' - ')[0]===ten) n++; }); if (n>0) throw new Error('Chưa xoá được: đang có ' + n + ' sản phẩm thuộc hãng ' + ten + '. Hãy sửa/xoá các sản phẩm đó trước.'); if (BRANDS.indexOf(ten)>=0) throw new Error('Hãng "' + ten + '" là hãng mặc định của hệ thống, không xoá được.'); var s = catalogSheet_(); var vals = s.getLastRow()>=2 ? s.getRange(2,1,s.getLastRow()-1,2).getValues() : []; var kept = vals.filter(function(r){ return String(r[0]||'').trim()!==ten; }); if (s.getLastRow()>=2) s.getRange(2,1,s.getLastRow()-1,2).clearContent(); if (kept.length) s.getRange(2,1,kept.length,2).setValues(kept); return { catalog: catalogRead_() }; }
function saveCatalogMain(tk, hang, chinh) { _auth(tk); requireRole_(['Admin','Kho']); hang = String(hang||'').trim(); chinh = String(chinh||'').trim(); if (!hang || !chinh) throw new Error('Thiếu hãng hoặc tên SP chính.'); var c = catalogRead_(); var exists = (c.mains[hang]||[]).some(function(x){ return x.toLowerCase()===chinh.toLowerCase(); }); if (!exists) readProducts_().forEach(function(p){ if (String(p.nhom||'')===hang+' - '+chinh) exists = true; }); if (exists) throw new Error('SP chính "' + chinh + '" của hãng ' + hang + ' đã có rồi.'); catalogSheet_().appendRow([hang, chinh]); return { catalog: catalogRead_() }; }
function deleteCatalogMain(tk, hang, chinh) { _auth(tk); requireRole_(['Admin','Kho']); hang = String(hang||'').trim(); chinh = String(chinh||'').trim(); var n = 0; readProducts_().forEach(function(p){ if (String(p.nhom||'')===hang+' - '+chinh) n++; }); if (n>0) throw new Error('Chưa xoá được: đang có ' + n + ' sản phẩm thuộc nhóm ' + hang + ' - ' + chinh + '.'); var s = catalogSheet_(); var vals = s.getLastRow()>=2 ? s.getRange(2,1,s.getLastRow()-1,2).getValues() : []; var kept = vals.filter(function(r){ return !(String(r[0]||'').trim()===hang && String(r[1]||'').trim()===chinh); }); if (s.getLastRow()>=2) s.getRange(2,1,s.getLastRow()-1,2).clearContent(); if (kept.length) s.getRange(2,1,kept.length,2).setValues(kept); return { catalog: catalogRead_() }; }
function saveCustomer(tk, c){ var lock = LockService.getScriptLock(); lock.waitLock(20000); try { return saveCustomer_goc(tk, c); } finally { lock.releaseLock(); } }
function saveProduct(tk, p){ var lock = LockService.getScriptLock(); lock.waitLock(20000); try { return saveProduct_goc(tk, p); } finally { lock.releaseLock(); } }
function deleteProduct(tk, ma){ var lock = LockService.getScriptLock(); lock.waitLock(20000); try { return deleteProduct_goc(tk, ma); } finally { lock.releaseLock(); } }
function addStockImport(tk, p){ var lock = LockService.getScriptLock(); lock.waitLock(20000); try { return addStockImport_goc(tk, p); } finally { lock.releaseLock(); } }
function addCashEntry(tk, p){ var lock = LockService.getScriptLock(); lock.waitLock(20000); try { return addCashEntry_goc(tk, p); } finally { lock.releaseLock(); } }
function getBootstrap(tk) { _auth(tk); return { config: getConfig(), user: getCurrentUser(), customers: getCustomers(), products: getProducts(), catalog: catalogRead_(), brands: BRANDS, units: UNITS, today: todayStr_(), nextInvoice: nextNo_('NhatKyBanHang', 2, 'HD') }; }

/* Tìm khách theo mã/tên; chưa có thì tự tạo với mã KH tự động (dùng khi lập đơn gõ tên khách) */
function resolveCustomer_(p, cusIn) {
  var cus = cusIn || readCustomers_();
  var ma = String(p.maKH||'').trim();
  if (ma && cus.some(function(c){return c.ma===ma;})) { p.maKH = ma; return ''; }
  var ten = String(p.tenKH||'').trim(); if (!ten) throw new Error('Chưa nhập tên khách hàng.');
  var found = null; cus.forEach(function(c){ if (String(c.ten||'').trim().toLowerCase() === ten.toLowerCase()) found = c; });
  if (found) { p.maKH = found.ma; return ''; }
  var s = sh_('KhachHang'); var v = s.getLastRow()>=4 ? s.getRange(4,1,s.getLastRow()-3,1).getValues() : [];
  var mx = 0; v.forEach(function(r){ var mm = String(r[0]||'').match(/^KH(\d+)$/i); if (mm) { var nn = parseInt(mm[1],10); if (nn>mx) mx=nn; } });
  var num = String(mx+1); while (num.length<3) num='0'+num; var maMoi = 'KH'+num;
  var row = 0; for (var qi=0; qi<v.length; qi++){ if (!v[qi][0]) { row = 4+qi; break; } } if (!row) row = 4 + v.length; if (row>103) throw new Error('Đã đủ 100 khách trong khung file gốc.');
  s.getRange(row,1,1,9).setValues([[maMoi, ten, String(p.loaiKH||'Khách lẻ'), String(p.sdtKH||''), '', 0, 30, 0, 'Tạo tự động khi lập đơn']]);
  cus.push({ ma:maMoi, ten:ten, loai:String(p.loaiKH||'Khách lẻ'), sdt:String(p.sdtKH||'') });
  p.maKH = maMoi; return maMoi;
}
/* Cong thuc chuan cua 1 dong NhatKyBanHang — app TU GHI cho moi dong no tao ra, khong phu thuoc khung cong thuc co san trong Sheet (bai hoc su co dong 204 ngay 21/10/2026: khung dung o dong 203 lam don moi mat gia/thanh tien) */
var NK_F_ = {
  D: '=IF(C{R}="";"";IFERROR(INDEX(KhachHang!$B$4:$B$203;MATCH(C{R};KhachHang!$A$4:$A$99;0));"Sai mã KH"))',
  F: '=IF(E{R}="";"";IFERROR(INDEX(SanPham!$C$4:$C$101;MATCH(E{R};SanPham!$A$4:$A$101;0));"Sai mã SP"))',
  G: '=IF(E{R}="";"";IFERROR(INDEX(SanPham!$D$4:$D$101;MATCH(E{R};SanPham!$A$4:$A$101;0));""))',
  I: '=IF(E{R}="";"";IFERROR(INDEX(SanPham!$E$4:$E$101;MATCH(E{R};SanPham!$A$4:$A$101;0));0))',
  K: '=IF(OR(C{R}="";E{R}="");"";H{R}*I{R}*(1-J{R}))',
  L: '=IF(OR(A{R}="";C{R}="");"";A{R}+IFERROR(INDEX(KhachHang!$G$4:$G$99;MATCH(C{R};KhachHang!$A$4:$A$99;0));0))'
};
function nkF_(k, row){ return NK_F_[k].split('{R}').join(row); }
function healNkRow_(s, row, isLinhdong, giaLaCongThuc){
  s.getRange(row,4).setFormula(nkF_('D', row));
  s.getRange(row,11).setFormula(nkF_('K', row));
  s.getRange(row,12).setFormula(nkF_('L', row));
  if (!isLinhdong) {
    s.getRange(row,6).setFormula(nkF_('F', row));
    s.getRange(row,7).setFormula(nkF_('G', row));
    // Cot I (don gia) KHONG ghi cong thuc nua: gia duoc dong bang bang gia tri tai thoi diem ban (ham ghi dong tu ghi setValue) — sua gia SP sau nay khong lam troi hoa don cu.
  }
}
function createInvoice(tk, p) { _auth(tk);
  requireRole_(['Admin','BanHang']);
  if (!p || (!p.maKH && !String(p.tenKH||'').trim())) throw new Error('Chưa nhập tên khách hàng.');
  if (!p.items || !p.items.length) throw new Error('Đơn chưa có mặt hàng nào.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var cus = readCustomers_(); var khMoi = resolveCustomer_(p, cus); var kh = null; cus.forEach(function(c){ if(c.ma===p.maKH) kh=c; }); if (!kh) throw new Error('Không tìm thấy khách ' + p.maKH);
    var prodList = readProducts_(); var prods = {}; prodList.forEach(function(x){ prods[x.ma]=x; }); var st = stockMap_(prodList);
    var total = 0; p.items.forEach(function(it){ if (it.maSP==='LINHDONG') { var slL = money_(it.sl); if (slL<=0) throw new Error('Số lượng phải lớn hơn 0: hàng linh động'); total += slL * money_(it.gia) * (1 - pct_(it.ck)); return; } if (!prods[it.maSP]) throw new Error('Không có sản phẩm ' + it.maSP); var sl = money_(it.sl); if (sl<=0) throw new Error('Số lượng phải lớn hơn 0: ' + it.maSP); var gia = it.gia!=='' && it.gia!==undefined && it.gia!==null ? money_(it.gia) : prods[it.maSP].gia; var ck = pct_(it.ck); total += sl*gia*(1-ck); var ton = st[it.maSP] ? st[it.maSP].ton : prods[it.maSP].tonDau; if (!p.choVuotTon && ton < sl) throw new Error('Không đủ tồn: ' + prods[it.maSP].ten + ' còn ' + ton + ' ' + prods[it.maSP].dvt + ', đơn cần ' + sl); });
    var giamDon = Math.min(Math.max(0, money_(p.giamDon)), total); total = Math.max(0, total - giamDon); var ngay = parseDate_(p.ngay); if (isNaN(ngay.getTime())) throw new Error('Ngày bán không hợp lệ.'); var cfg = getConfig();
    // Đọc cột B một lần: vừa lấy số HĐ tiếp theo vừa tìm dòng trống (trước đây đọc 2 lần)
    var s = sh_('NhatKyBanHang'); var colB = s.getRange(4, 2, Math.min(1000, Math.max(10, s.getLastRow()-3+20)), 1).getValues(); var preHD = 'HD' + Utilities.formatDate(new Date(), TZ, 'yy') + '-'; var maxNo = 0, row = 0;
    colB.forEach(function(r, i){ var v = String(r[0]||''); if (v.indexOf(preHD)===0) { var nn = parseInt(v.substring(preHD.length),10); if (!isNaN(nn) && nn>maxNo) maxNo = nn; } });
    // Tim DOAN TRONG LIEN TUC du cho tat ca mat hang cua don: chi lay dong trong dau tien roi ghi ca block se GHI DE len dong cua don khac khi giua so co dong da bi xoa trang (sua don giam so dong de lai dong trong giua so).
    var needRows = p.items.length; var runLen = 0;
    for (var ci = 0; ci < colB.length; ci++) { if (!String(colB[ci][0]||'')) { runLen++; if (runLen >= needRows) { row = 4 + ci - needRows + 1; break; } } else runLen = 0; }
    var dhScan = sh_('DonHang'); if (dhScan && dhScan.getLastRow()>=2) dhScan.getRange(2,1,dhScan.getLastRow()-1,1).getValues().forEach(function(r){ var v = String(r[0]||''); if (v.indexOf(preHD)===0) { var nn2 = parseInt(v.substring(preHD.length),10); if (!isNaN(nn2) && nn2>maxNo) maxNo = nn2; } });
    if (!row && colB.length<1000) { var tailRow = 4 + colB.length; if (tailRow + needRows - 1 <= 1003) row = tailRow; } if (!row) throw new Error('NhatKyBanHang không còn đoạn trống liên tục đủ ' + needRows + ' dòng cho đơn này.');
    var soHD = String(maxNo+1); while (soHD.length<4) soHD = '0'+soHD; soHD = preHD + soHD;
    var n = p.items.length; var blkAC = [], blkE = [], blkH = [], blkJ = [], blkMN = [];
    p.items.forEach(function(it){ blkAC.push([ngay, soHD, p.maKH]); blkE.push([it.maSP]); blkH.push([money_(it.sl)]); blkJ.push([pct_(it.ck)]); var dt=it.dungTich||''; var mauFull=dt+(dt&&(it.mau||'')?' | ':'')+(it.mau||''); blkMN.push([p.ghiChu||'', mauFull]); });
    s.getRange(row,1,n,3).setValues(blkAC); s.getRange(row,5,n,1).setValues(blkE); s.getRange(row,8,n,1).setValues(blkH); s.getRange(row,10,n,1).setValues(blkJ); s.getRange(row,13,n,2).setValues(blkMN);
    p.items.forEach(function(it, idx){ if (it.maSP==='LINHDONG') { s.getRange(row+idx,6).setValue(String(it.tenSP||'Hàng linh động').slice(0,120)); s.getRange(row+idx,7).setValue(String(it.dvt||'cái').slice(0,20)); s.getRange(row+idx,9).setValue(money_(it.gia)); healNkRow_(s, row+idx, true, false); return; } var giaDefault = prods[it.maSP].gia; var gia = it.gia!=='' && it.gia!==undefined && it.gia!==null ? money_(it.gia) : giaDefault; s.getRange(row+idx,9).setValue(gia); healNkRow_(s, row+idx, false, false); });
    // Đơn hàng theo dõi
    var traNgay = Math.min(money_(p.traNgay), total); var dh = sh_('DonHang'); if (dh) dh.appendRow([soHD, ngay, p.maKH, kh.ten, total, traNgay, total-traNgay, p.trangThaiDon||'Mới', traNgay>=total && total>0 ? 'Đã thu đủ' : (traNgay>0 ? 'Thu 1 phần' : 'Chưa thu'), traNgay>0?1:0, new Date(), getCurrentUser().name, p.ghiChu||'']);
    // Trả ngay đợt 1 -> phiếu thu
    var soPT = ''; if (traNgay>0) { soPT = ghiThuTien_(ngay, p.maKH, traNgay, p.hinhThuc||'Tiền mặt', soHD, 'Thu đợt 1 - trả ngay khi xuất đơn ' + soHD); }
    // Tích điểm theo hoá đơn (có công tắc DiemBatTat trong CauHinh)
    var diemBat = !(cfg.DiemBatTat===0 || cfg.DiemBatTat==='0' || cfg.DiemBatTat===false || cfg.DiemBatTat==='false'); var diem = diemBat ? Math.floor(total / (cfg.TienCho1Diem || 10000)) : 0; var ds = sh_('DiemTichLuy'); if (ds && diem>0) ds.appendRow([ngay, p.maKH, kh.ten, soHD, total, diem, 0, '', 'Tích điểm theo hoá đơn: ' + Math.round(cfg.TienCho1Diem).toLocaleString('vi-VN') + 'đ = 1 điểm']);
    return { soHD: soHD, soPT: soPT, tongTien: Math.round(total), giamDon: Math.round(giamDon), traNgay: Math.round(traNgay), conNo: Math.round(total-traNgay), diemCong: diem, tenKH: kh.ten, khMoi: khMoi };
  } finally { lock.releaseLock(); }
}
function ghiThuTien_(ngay, maKH, soTien, hinhThuc, soHD, ghiChu) { var s = sh_('ThuTien'); var row = firstEmptyRow_('ThuTien','B',4,2003); var soPT = nextThuNo_(); s.getRange(row,1,1,3).setValues([[ngay, soPT, maKH]]); s.getRange(row,5,1,4).setValues([[soTien, hinhThuc||'Tiền mặt', soHD||'', ghiChu||'']]); s.getRange(row,4).setFormula('=IF(C'+row+'="";"";IFERROR(INDEX(KhachHang!$B$4:$B$99;MATCH(C'+row+';KhachHang!$A$4:$A$99;0));"Sai mã KH"))'); s.getRange(row,9).setFormula('=IF(AND(C'+row+'<>"";C'+row+'=SelKH;A'+row+'>=TuNgay;A'+row+'<=DenNgay);A'+row+'+0,5+ROW()/100000;"")'); _payAllCache = null; return soPT; }

/* ============ THU TIỀN THEO ĐỢT 2-3 ĐỢT + CÔNG NỢ ============ */
var _payAllCache = null;
function allPayments_() { if (_payAllCache) return _payAllCache; var map = {}; var s = sh_('ThuTien'); if (s && s.getLastRow()>=4) s.getRange(4,1,s.getLastRow()-3,8).getValues().forEach(function(r){ if (!r[6]) return; var k = String(r[6]); if (!map[k]) map[k] = []; map[k].push({ ngay:fmtDate_(r[0]), soPT:r[1], maKH:r[2], tenKH:r[3], soTien:money_(r[4]), hinhThuc:r[5], soHD:r[6], ghiChu:r[7]||'', dot: map[k].length+1 }); }); _payAllCache = map; return map; }
function paymentsOfInvoice_(soHD) { var m = allPayments_(); return m[String(soHD)] || []; }
function addPayment(tk, p) { _auth(tk); // {soHD, maKH, soTien, hinhThuc, ngay, ghiChu}
  requireRole_(['Admin','BanHang','KeToan']); if (!p.soHD || !p.maKH) throw new Error('Thiếu số hoá đơn / khách hàng.'); var soTien = money_(p.soTien); if (soTien<=0) throw new Error('Số tiền thu phải lớn hơn 0.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var det = invoiceDetail_(p.soHD);
    if (String(det.head.maKH) !== String(p.maKH)) throw new Error('Khách hàng không khớp với khách của đơn ' + p.soHD + ' (' + det.head.maKH + ').');
    if (soTien > det.conNo + 1) throw new Error('Thu vượt số còn nợ của đơn: còn nợ ' + det.conNo.toLocaleString('vi-VN') + 'đ');
    var pays = paymentsOfInvoice_(p.soHD); var dot = pays.length + 1; if (dot > 5) throw new Error('Đơn này đã có 5 đợt thu, kiểm tra lại.');
    var soPT = ghiThuTien_(parseDate_(p.ngay), p.maKH, soTien, p.hinhThuc||'Tiền mặt', p.soHD, (p.ghiChu||'') + ' [Thu đợt ' + dot + ' của ' + p.soHD + ']');
    // cập nhật DonHang (uu tien dong khop ca so HD lan ma KH — phong so HD trung voi dong mau)
    var dh = sh_('DonHang'); if (dh && dh.getLastRow()>=2) { var v = dh.getRange(2,1,dh.getLastRow()-1,10).getValues(); var pick = -1; for (var i=0;i<v.length;i++){ if (String(v[i][0])===String(p.soHD)) { if (String(v[i][2])===String(p.maKH)) { pick = i; break; } if (pick<0) pick = i; } } if (pick>=0) { var daThu = money_(v[pick][5]) + soTien; var tong = money_(v[pick][4]); dh.getRange(pick+2,6).setValue(daThu); dh.getRange(pick+2,7).setValue(tong-daThu); dh.getRange(pick+2,9).setValue(daThu>=tong ? 'Đã thu đủ' : 'Thu 1 phần'); dh.getRange(pick+2,10).setValue(dot); } }
    return { soPT: soPT, dot: dot, detail: invoiceDetail_(p.soHD) };
  } finally { lock.releaseLock(); }
}

/* ============ HOÀN ĐƠN HÀNG ============ */
function hoanDon(tk, soHD, lyDo) { _auth(tk);
  requireRole_(['Admin','BanHang']);
  soHD = String(soHD||'').trim(); if (!soHD) throw new Error('Thiếu số hoá đơn.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var dh = sh_('DonHang'); var dhRow = -1, dhStatus = ''; var dhCands = [];
    if (dh && dh.getLastRow()>=2) { var vv = dh.getRange(2,1,dh.getLastRow()-1,13).getValues();
      for (var i=0;i<vv.length;i++){ if (String(vv[i][0])===soHD) dhCands.push({ row:i+2, status:String(vv[i][7]||''), maKH:String(vv[i][2]||'') }); } }
    var s = sh_('NhatKyBanHang'); var lines=[];
    if (s && s.getLastRow()>=4) s.getRange(4,1,s.getLastRow()-3,14).getValues().forEach(function(r){
      if (String(r[1])===soHD) lines.push({ ngay:r[0], maKH:r[2], tenKH:r[3], maSP:r[4], tenSP:r[5], dvt:r[6], gia:money_(r[8]), sl:money_(r[7]), ck:pct_(r[9]), tt:money_(r[10]), mau:r[13]||'' });
    });
    if (!lines.length) throw new Error('Không tìm thấy dòng nào của đơn ' + soHD + ' trong NhatKyBanHang.');
    var goc = lines.filter(function(l){return l.sl>0;});
    if (!goc.length) throw new Error('Đơn ' + soHD + ' đã có dòng hoàn, không hoàn 2 lần.');
    if (dhCands.length) { var pickH = dhCands.filter(function(c){ return c.maKH===String(goc[0].maKH); })[0] || dhCands[0]; dhRow = pickH.row; dhStatus = pickH.status; }
    if (dhStatus==='Đã hoàn') throw new Error('Đơn ' + soHD + ' đã hoàn trước đó, không hoàn 2 lần.');
    var pays = paymentsOfInvoice_(soHD); var daThu = pays.reduce(function(a,b){return a+b.soTien;},0);
    var tongGoc = goc.reduce(function(a,l){return a+l.tt;},0);
    var ngayHoan = new Date();
    var note = 'Hoàn đơn ' + soHD + (lyDo ? ' - Lý do: ' + lyDo : '');
    if (daThu>0) note += '. Đã thu ' + Math.round(daThu).toLocaleString('vi-VN') + 'đ trước khi hoàn - kế toán xử lý hoàn tiền cho khách, KHÔNG tự xoá phiếu thu.';
    // Tim DOAN TRONG LIEN TUC du cho tat ca dong hoan (nhu createInvoice) — firstEmptyRow_ chi cho dong trong dau tien, ghi ca block se de len dong cua don khac khi giua so co dong da bi xoa trang.
    var colBH = s.getRange(4, 2, Math.max(1, s.getLastRow()-3), 1).getValues(); var r0 = 0; var runH = 0;
    for (var hi = 0; hi < colBH.length; hi++) { if (!String(colBH[hi][0]||'')) { runH++; if (runH >= goc.length) { r0 = 4 + hi - goc.length + 1; break; } } else runH = 0; }
    if (!r0) { var tailH = Math.max(4, s.getLastRow() + 1); if (tailH + goc.length - 1 <= 1003) r0 = tailH; }
    if (!r0 || r0 + goc.length - 1 > 1003) throw new Error('Không đủ dòng trống liên tục trong NhatKyBanHang để ghi dòng hoàn.');
    goc.forEach(function(l, idx){
      var rr = r0+idx;
      s.getRange(rr,1).setValue(ngayHoan);   // A Ngày: chỉ ghi cột nhập tay
      s.getRange(rr,2).setValue(soHD);       // B Số HĐ
      s.getRange(rr,3).setValue(l.maKH);     // C Mã KH
      s.getRange(rr,5).setValue(l.maSP);     // E Mã SP
      s.getRange(rr,8).setValue(-l.sl);      // H Số lượng âm -> hoàn tồn, công thức cột K tự tính thành tiền âm
      s.getRange(rr,10).setValue(l.ck);      // J Chiết khấu như dòng gốc
      s.getRange(rr,13).setValue(note); s.getRange(rr,14).setValue(l.mau||''); if (l.maSP==='LINHDONG') { s.getRange(rr,6).setValue(l.tenSP||'Hàng linh động'); s.getRange(rr,7).setValue(l.dvt||'cái'); s.getRange(rr,9).setValue(l.gia||0); } else s.getRange(rr,9).setValue(l.gia||0); healNkRow_(s, rr, l.maSP==='LINHDONG', false);      // M Ghi chú hoàn đơn
    });
    SpreadsheetApp.flush();
    // hoàn điểm: trừ lại số điểm đã cộng cho đơn này
    var diemDaCong = 0; var ds = sh_('DiemTichLuy');
    if (ds && ds.getLastRow()>=2) ds.getRange(2,4,ds.getLastRow()-1,4).getValues().forEach(function(r){ if (String(r[0])===soHD) diemDaCong += money_(r[2]) - money_(r[3]); });
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

/* ============ SỬA ĐƠN HÀNG (giữ nguyên số HĐ + các đợt thu; ghi đè dòng hàng tại chỗ, khớp lại kho/điểm/công nợ) ============ */
function updateInvoice(tk, p) { _auth(tk);
  requireRole_(['Admin','BanHang']);
  var soHD = String((p && p.soHD) || '').trim(); if (!soHD) throw new Error('Thiếu số hoá đơn cần sửa.');
  if (!p.items || !p.items.length) throw new Error('Đơn chưa có mặt hàng nào.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var dh = sh_('DonHang'); var dhRow = -1, dhStatus = '', dhTong = 0; var dhCands = [];
    if (dh && dh.getLastRow()>=2) { var vv = dh.getRange(2,1,dh.getLastRow()-1,13).getValues();
      for (var i=0;i<vv.length;i++){ if (String(vv[i][0])===soHD) dhCands.push({ row:i+2, status:String(vv[i][7]||''), tong:money_(vv[i][4]), maKH:String(vv[i][2]||'') }); } }
    var det = invoiceDetail_(soHD);
    if (dhCands.length) { var pickU = dhCands.filter(function(c){ return c.maKH===String(det.head.maKH); })[0] || dhCands[0]; dhRow = pickU.row; dhStatus = pickU.status; dhTong = pickU.tong; }
    if (dhStatus==='Đã hoàn') throw new Error('Đơn ' + soHD + ' đã hoàn — không sửa được.');
    var lineSum0 = (det.items||[]).reduce(function(a,b){ return a + (Number(b.thanhTien)||0); }, 0);
    if (dhTong>0 && lineSum0<=0) throw new Error('Đơn ' + soHD + ' đang lỗi công thức trong Sheet (chi tiết đọc ra 0đ nhưng sổ đơn ghi ' + dhTong.toLocaleString('vi-VN') + 'đ). Hãy khôi phục công thức trong Sheet trước khi sửa đơn.');
    var daThu = Math.round(det.daThu||0);
    var s = sh_('NhatKyBanHang'); var oldRows = [];
    if (s && s.getLastRow()>=4) s.getRange(4,1,s.getLastRow()-3,14).getValues().forEach(function(r, ix){
      if (String(r[1])===soHD) oldRows.push({ row: ix+4, maSP: String(r[4]||''), sl: money_(r[7]) });
    });
    if (!oldRows.length) throw new Error('Không tìm thấy dòng nào của đơn ' + soHD + ' trong NhatKyBanHang.');
    var ownQty = {}; oldRows.forEach(function(o){ if (o.maSP && o.maSP!=='LINHDONG') ownQty[o.maSP] = (ownQty[o.maSP]||0) + o.sl; });
    var cus = readCustomers_(); var khMoi = resolveCustomer_(p, cus); var kh = null; cus.forEach(function(c){ if(c.ma===p.maKH) kh=c; }); if (!kh) throw new Error('Không tìm thấy khách ' + p.maKH);
    var prodList = readProducts_(); var prods = {}; prodList.forEach(function(x){ prods[x.ma]=x; }); var st = stockMap_(prodList);
    var subTotal = 0;
    p.items.forEach(function(it){
      if (it.maSP==='LINHDONG') { var slL = money_(it.sl); if (slL<=0) throw new Error('Số lượng phải lớn hơn 0: hàng linh động'); subTotal += slL * money_(it.gia) * (1 - pct_(it.ck)); return; }
      if (!prods[it.maSP]) throw new Error('Không có sản phẩm ' + it.maSP);
      var sl = money_(it.sl); if (sl<=0) throw new Error('Số lượng phải lớn hơn 0: ' + it.maSP);
      var gia = it.gia!=='' && it.gia!==undefined && it.gia!==null ? money_(it.gia) : prods[it.maSP].gia;
      subTotal += sl*gia*(1-pct_(it.ck));
      var ton = (st[it.maSP] ? st[it.maSP].ton : prods[it.maSP].tonDau) + (ownQty[it.maSP]||0);
      if (!p.choVuotTon && ton < sl) throw new Error('Không đủ tồn: ' + prods[it.maSP].ten + ' còn ' + ton + ' ' + prods[it.maSP].dvt + ' (đã tính cả số đang nằm trong đơn này), đơn cần ' + sl);
    });
    var giamDon = Math.min(Math.max(0, money_(p.giamDon)), subTotal); var newTotal = Math.round(Math.max(0, subTotal - giamDon)); subTotal = Math.round(subTotal);
    if (newTotal < daThu) throw new Error('Tổng mới của đơn (' + newTotal.toLocaleString('vi-VN') + 'đ) nhỏ hơn số khách đã trả (' + daThu.toLocaleString('vi-VN') + 'đ). Hãy xử lý phần tiền đã thu trước rồi mới sửa đơn.');
    var ngayCu = parseDate_(det.head.ngay); var ghiChuMoi = String(p.ghiChu||''); var homNay = new Date();
    function vietDong_(row, it, oldLine){
      var dt = it.dungTich||''; var mauFull = dt + (dt && it.mau ? ' | ' : '') + (it.mau||'');
      s.getRange(row,1,1,3).setValues([[ngayCu, soHD, p.maKH]]);
      s.getRange(row,5).setValue(it.maSP);
      s.getRange(row,8).setValue(money_(it.sl));
      s.getRange(row,10).setValue(pct_(it.ck));
      s.getRange(row,13,1,2).setValues([[ghiChuMoi, mauFull]]);
      if (it.maSP==='LINHDONG') { s.getRange(row,6).setValue(String(it.tenSP||'Hàng linh động').slice(0,120)); s.getRange(row,7).setValue(String(it.dvt||'cái').slice(0,20)); s.getRange(row,9).setValue(money_(it.gia)); healNkRow_(s, row, true, false); return; }
      var prod = prods[it.maSP]; var gia = it.gia!=='' && it.gia!==undefined && it.gia!==null ? money_(it.gia) : prod.gia;
      s.getRange(row,9).setValue(gia);
      healNkRow_(s, row, false, false);
    }
    p.items.forEach(function(it, idx){
      var row, oldLine = null;
      if (idx < oldRows.length) { row = oldRows[idx].row; oldLine = oldRows[idx]; }
      else row = firstEmptyRow_('NhatKyBanHang','B',4,1003);
      vietDong_(row, it, oldLine);
    });
    for (var k = p.items.length; k < oldRows.length; k++) {
      var rr = oldRows[k].row;
      s.getRange(rr,1,1,3).clearContent();
      s.getRange(rr,5).clearContent(); s.getRange(rr,8).clearContent(); s.getRange(rr,10).clearContent();
      s.getRange(rr,13,1,2).clearContent();
      s.getRange(rr,9).clearContent();
      if (oldRows[k].maSP==='LINHDONG') { s.getRange(rr,6,1,2).clearContent(); }
    }
    SpreadsheetApp.flush();
    var cfg = getConfig();
    var diemBat = !(cfg.DiemBatTat===0 || cfg.DiemBatTat==='0' || cfg.DiemBatTat===false || cfg.DiemBatTat==='false');
    var ds = sh_('DiemTichLuy'); var netOld = 0;
    if (ds && ds.getLastRow()>=2) ds.getRange(2,1,ds.getLastRow()-1,9).getValues().forEach(function(r){ if (String(r[3])===soHD) netOld += money_(r[5]) - money_(r[6]); });
    netOld = Math.round(netOld);
    if (ds && netOld!==0) { if (netOld>0) ds.appendRow([homNay, det.head.maKH, det.head.tenKH, soHD, 0, 0, netOld, '', 'Sửa đơn ' + soHD + ' — trừ điểm theo tổng cũ']); else ds.appendRow([homNay, det.head.maKH, det.head.tenKH, soHD, 0, -netOld, 0, '', 'Sửa đơn ' + soHD + ' — bù điểm theo tổng cũ']); }
    var diemMoi = diemBat ? Math.floor(newTotal / (cfg.TienCho1Diem || 10000)) : 0;
    if (ds && diemMoi>0) ds.appendRow([homNay, p.maKH, kh.ten, soHD, newTotal, diemMoi, 0, '', 'Sửa đơn ' + soHD + ' — cộng điểm theo tổng mới']);
    var pays = paymentsOfInvoice_(soHD);
    var ttTT = (newTotal>0 && daThu>=newTotal) ? 'Đã thu đủ' : (daThu>0 ? 'Thu 1 phần' : 'Chưa thu');
    var noteSua = 'Sửa đơn ' + fmtDate_(homNay) + ' bởi ' + getCurrentUser().name + ': tổng ' + Math.round(det.tongTien).toLocaleString('vi-VN') + 'đ → ' + newTotal.toLocaleString('vi-VN') + 'đ';
    if (dh) {
      if (dhRow>0) {
        dh.getRange(dhRow,3,1,2).setValues([[p.maKH, kh.ten]]);
        dh.getRange(dhRow,5,1,3).setValues([[newTotal, daThu, newTotal-daThu]]);
        dh.getRange(dhRow,9).setValue(ttTT); dh.getRange(dhRow,10).setValue(pays.length);
        var gcCu = dh.getRange(dhRow,13).getValue(); dh.getRange(dhRow,13).setValue((gcCu?gcCu+' | ':'')+noteSua);
      } else dh.appendRow([soHD, ngayCu, p.maKH, kh.ten, newTotal, daThu, newTotal-daThu, 'Mới', ttTT, pays.length, homNay, getCurrentUser().name, noteSua]);
    }
    return { soHD: soHD, tongCu: Math.round(det.tongTien), tongTien: newTotal, giamDon: Math.round(giamDon), daThu: daThu, conNo: newTotal-daThu, diemCu: netOld, diemMoi: diemMoi, tenKH: kh.ten, khMoi: khMoi };
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
function getProfit(tk, from, to) { _auth(tk);
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
    linesOfInvoice_(o.soHD).forEach(function(l){ if (l.sl<=0) return; var c = costs[l.maSP] || prodCost[l.maSP] || 0; giaVon += l.sl*c; }); doanhThu = o.tongTien; // doanh thu theo so DonHang (da tru giam gia don) — khong cong tong dong tho nua
    var lai = doanhThu - giaVon;
    rows.push({ soHD:o.soHD, ngay:o.ngay, tenKH:o.tenKH, soMatHang:o.soMatHang, doanhThu:doanhThu, giaVon:giaVon, lai:lai });
    tD += doanhThu; tV += giaVon; tL += lai;
  });
  rows.sort(function(a,b){return String(b.soHD).localeCompare(String(a.soHD));});
  return { rows: rows, tongDoanhThu: tD, tongGiaVon: tV, tongLai: tL };
}

/* ============ ĐƠN HÀNG / LỊCH SỬ XUẤT / CHI TIẾT HOÁ ĐƠN ============ */
function groupInvoices_() { var s = sh_('NhatKyBanHang'); var map = {}; if (!s || s.getLastRow()<4) return []; s.getRange(4,1,s.getLastRow()-3,13).getValues().forEach(function(r){ if(!r[1]) return; var k = String(r[1]); if(!map[k]) map[k] = { soHD:k, ngay:fmtDate_(r[0]), maKH:r[2], tenKH:r[3], soMatHang:0, tongTien:0, hanTT:fmtDate_(r[11]), ghiChu:r[12]||'' }; map[k].soMatHang++; map[k].tongTien += money_(r[10]); }); var dhMap = {}; var dh = sh_('DonHang'); if (dh && dh.getLastRow()>=2) dh.getRange(2,1,dh.getLastRow()-1,10).getValues().forEach(function(r){ dhMap[String(r[0])] = { trangThaiDon:r[7]||'Mới', trangThaiTT:r[8]||'', daThu:money_(r[5]), tong:money_(r[4]) }; }); return Object.keys(map).map(function(k){ var o = map[k]; var pays = paymentsOfInvoice_(k); var daThu2 = pays.reduce(function(a,b){return a+b.soTien;},0); o.daThu = dhMap[k] && dhMap[k].daThu ? dhMap[k].daThu : daThu2; o.tongTien = Math.round(o.tongTien); if (dhMap[k] && dhMap[k].tong>0 && dhMap[k].trangThaiDon!=='Đã hoàn') o.tongTien = Math.round(dhMap[k].tong); o.daThu = Math.round(o.daThu); o.conNo = Math.max(0, o.tongTien - o.daThu); o.soDotThu = pays.length; o.trangThaiDon = dhMap[k] ? dhMap[k].trangThaiDon : 'Mới'; o.trangThaiTT = o.conNo<=0 && o.tongTien>0 ? 'Đã thu đủ' : (o.daThu>0 ? 'Thu 1 phần' : 'Chưa thu'); return o; }).sort(function(a,b){ var d = parseDate_(b.ngay) - parseDate_(a.ngay); if (d) return d; return String(b.soHD).localeCompare(String(a.soHD)); }); }
function getInvoices(tk, f) { _auth(tk); f = f || {}; var list = groupInvoices_(); if (f.tuNgay) list = list.filter(function(o){ return parseDate_(o.ngay) >= parseDate_(f.tuNgay); }); if (f.denNgay) list = list.filter(function(o){ return parseDate_(o.ngay) <= parseDate_(f.denNgay); }); if (f.maKH) list = list.filter(function(o){ return o.maKH===f.maKH; }); if (f.trangThai) list = list.filter(function(o){ return o.trangThaiDon===f.trangThai || o.trangThaiTT===f.trangThai; }); if (f.tim) { var q = String(f.tim).toLowerCase(); list = list.filter(function(o){ return o.soHD.toLowerCase().indexOf(q)>=0 || String(o.tenKH).toLowerCase().indexOf(q)>=0; }); } return list.slice(0, 300); }
function getOrders(tk) { _auth(tk); return getInvoices(null, {}); }
function updateOrderStatus(tk, soHD, status) { _auth(tk); requireRole_(['Admin','BanHang','Kho']); if (status==='Đã hoàn') throw new Error('Không đặt trạng thái Đã hoàn ở đây — hãy dùng nút Hoàn đơn để hoàn kho và trừ điểm đúng quy trình.'); if (['Mới','Đang giao','Hoàn thành','Đã hủy'].indexOf(status)<0) throw new Error('Trạng thái không hợp lệ: ' + status); var dh = sh_('DonHang'); if (!dh) throw new Error('Chưa có tab DonHang, hãy chạy cài đặt hệ thống.'); var inv0 = groupInvoices_().filter(function(o){return o.soHD===soHD;})[0]; if (dh.getLastRow()>=2) { var v = dh.getRange(2,1,dh.getLastRow()-1,3).getValues(); var pickS = -1; for (var i=0;i<v.length;i++){ if (String(v[i][0])===String(soHD)) { if (pickS<0) pickS = i; if (inv0 && String(v[i][2])===String(inv0.maKH)) { pickS = i; break; } } } if (pickS>=0) { dh.getRange(pickS+2,8).setValue(status); return true; } } // đơn cũ chưa có trong DonHang -> tạo dòng theo dõi
  var inv = groupInvoices_().filter(function(o){return o.soHD===soHD;})[0]; if (inv) dh.appendRow([soHD, parseDate_(inv.ngay), inv.maKH, inv.tenKH, inv.tongTien, inv.daThu, inv.conNo, status, inv.trangThaiTT, inv.soDotThu, new Date(), getCurrentUser().name, '']); return true; }
/* ===== NHUNG ANH (URL -> data URI) DE XUAT PDF KHONG MAT ANH SAN PHAM / MA QR ===== */
function inlineImgMap_(urls) {
  var out = {}; var need = [];
  var cache = null; try { cache = CacheService.getScriptCache(); } catch(e) {}
  (urls || []).forEach(function(u) {
    if (!u || typeof u !== 'string' || out[u]) return;
    if (u.indexOf('data:') === 0 || !/^https?:\/\//i.test(u)) { out[u] = u; return; }
    out[u] = u;
    var key = 'img_' + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, u).map(function(b){ return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
    var hit = null; if (cache) { try { hit = cache.get(key); } catch(e) {} }
    if (hit) { out[u] = hit; return; }
    need.push({ u: u, key: key });
  });
  if (need.length) {
    var resps = [];
    try { resps = UrlFetchApp.fetchAll(need.map(function(n){ return { url: n.u, muteHttpExceptions: true, followRedirects: true }; })); } catch(e) { resps = []; }
    var total = 0;
    need.forEach(function(n, i) {
      var r = resps && resps[i]; if (!r) return;
      try {
        if (r.getResponseCode() !== 200) return;
        var blob = r.getBlob(); var ct = blob.getContentType() || '';
        if (ct.indexOf('image/') !== 0) return;
        var bytes = blob.getBytes(); if (!bytes || !bytes.length || bytes.length > 1200000) return;
        var duri = 'data:' + ct + ';base64,' + Utilities.base64Encode(bytes);
        total += duri.length; if (total > 6000000) return;
        out[n.u] = duri;
        if (cache && duri.length < 90000) { try { cache.put(n.key, duri, 21600); } catch(e) {} }
      } catch(e) {}
    });
  }
  return out;
}
function inlineDetailImages_(d) {
  if (!d) return d;
  var urls = [];
  if (d.qrUrl) urls.push(d.qrUrl);
  (d.items || []).forEach(function(it){ if (it.hinh) urls.push(it.hinh); });
  if (!urls.length) return d;
  var map = inlineImgMap_(urls);
  if (d.qrUrl && map[d.qrUrl]) d.qrUrl = map[d.qrUrl];
  (d.items || []).forEach(function(it){ if (it.hinh && map[it.hinh]) it.hinh = map[it.hinh]; });
  return d;
}
function inlineInvoiceImages(tk, urls) { _auth(tk); if (!urls || !urls.length) return {}; return inlineImgMap_(urls.slice(0, 30)); }
function kiemTraQuyenAnh() {
  var url = 'https://img.vietqr.io/image/MB-1234567890-compact2.svg?amount=1000&addInfo=Kiem%20tra%20quyen%20anh';
  try {
    var r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    var body = r.getContentText() || '';
    var coPng = body.indexOf('data:image/png') >= 0;
    Logger.log('kiemTraQuyenAnh: HTTP ' + r.getResponseCode() + ' | SVG co data:image/png: ' + coPng + ' | KET QUA: ' + ((r.getResponseCode() === 200 && coPng) ? 'DA CO QUYEN tai anh ngoai — UrlFetchApp chay tot' : 'Tai duoc nhung noi dung khac du kien, kiem tra lai'));
    return (r.getResponseCode() === 200 && coPng);
  } catch (e) {
    Logger.log('kiemTraQuyenAnh: LOI ' + e.message + ' — neu la loi quyen (authorization) thi can cap quyen ket noi dich vu ngoai (UrlFetchApp) cho project.');
    return false;
  }
}
function getInvoiceDetail(tk, soHD) { _auth(tk); var d = invoiceDetail_(soHD); inlineDetailImages_(d); return d; }
function invoiceDetail_(soHD) { var s = sh_('NhatKyBanHang'); var items = []; var head = null; var imgMap = {}; readProducts_().forEach(function(p){ imgMap[p.ma]=p.hinh; }); if (s && s.getLastRow()>=4) s.getRange(4,1,s.getLastRow()-3,14).getValues().forEach(function(r){ if (String(r[1])===String(soHD)) { if(!head) head = { soHD:soHD, ngay:fmtDate_(r[0]), maKH:r[2], tenKH:r[3], hanTT:fmtDate_(r[11]), ghiChu:r[12]||'' }; items.push({ maSP:r[4], tenSP:r[5], dvt:r[6], sl:money_(r[7]), gia:money_(r[8]), ck:pct_(r[9]), thanhTien:Math.round(money_(r[10])), mau:r[13]||'', hinh: imgMap[r[4]]||'' }); } }); if (!head) throw new Error('Không tìm thấy hoá đơn ' + soHD); var tong = items.reduce(function(a,b){return a+b.thanhTien;},0); var pays = paymentsOfInvoice_(soHD); var daThu = pays.reduce(function(a,b){return a+b.soTien;},0); tong = Math.round(tong); daThu = Math.round(daThu); var giamDon = 0; var dhS = sh_('DonHang'); if (dhS && dhS.getLastRow()>=2) { var dhRows = dhS.getRange(2,1,dhS.getLastRow()-1,8).getValues(); var dhPick = null; for (var di=0; di<dhRows.length; di++) { if (String(dhRows[di][0])===String(soHD)) { if (!dhPick) dhPick = dhRows[di]; if (String(dhRows[di][2])===String(head.maKH)) { dhPick = dhRows[di]; break; } } } if (dhPick) { head.trangThaiDon = String(dhPick[7]||''); var dhTong = money_(dhPick[4]); if (dhTong < tong) { giamDon = tong - dhTong; tong = dhTong; } } } var kh = null; readCustomers_().forEach(function(c){ if(c.ma===head.maKH) kh=c; }); var cfg = getConfig(); var qr = ''; if (cfg.BankCode && cfg.SoTaiKhoan) qr = 'https://img.vietqr.io/image/' + encodeURIComponent(cfg.BankCode) + '-' + encodeURIComponent(cfg.SoTaiKhoan) + '-compact2.svg?amount=' + Math.round(Math.max(0,tong-daThu)) + '&addInfo=' + encodeURIComponent('Thanh toan ' + soHD) + '&accountName=' + encodeURIComponent(cfg.ChuTaiKhoan||''); var diemBat = !(cfg.DiemBatTat===0 || cfg.DiemBatTat==='0' || cfg.DiemBatTat===false || cfg.DiemBatTat==='false'); return { head: head, items: items, tongTien: tong, giamDon: giamDon, payments: pays, daThu: daThu, conNo: Math.max(0,tong-daThu), customer: kh, config: cfg, qrUrl: qr, diemHD: diemBat ? Math.floor(tong / (cfg.TienCho1Diem||10000)) : 0 }; }

/* ============ TỒN KHO + NHẬP HÀNG ============ */
function getStock(tk) { _auth(tk); var cfg = getConfig(); return getProducts().map(function(p){ p.canNhap = p.ton <= (p.tonMin || cfg.TonCanhBao); return p; }); }
function addStockImport_goc(tk, p) { _auth(tk); requireRole_(['Admin','Kho']); var s = sh_('NhapKho'); if (!s) { caiDatHeThong(); s = sh_('NhapKho'); } s.appendRow([parseDate_(p.ngay), p.maSP, p.tenSP||'', p.dvt||'', money_(p.sl), money_(p.giaNhap), money_(p.sl)*money_(p.giaNhap), p.ncc||'', p.ghiChu||'', getCurrentUser().name]); return getStock(); }
function getImportHistory(tk) { _auth(tk); requireRole_(['Admin','BanHang','Kho','KeToan','Xem']); var s = sh_('NhapKho'); var out = []; if (s && s.getLastRow()>=2) { var v = s.getRange(2,1,s.getLastRow()-1,10).getValues(); for (var i=v.length-1;i>=0 && out.length<100;i--) { var r=v[i]; if (!r[1]) continue; out.push({ ngay: fmtDate_(r[0]), maSP: r[1], tenSP: r[2]||'', dvt: r[3]||'', sl: money_(r[4]), giaNhap: money_(r[5]), thanhTien: money_(r[6]), ncc: r[7]||'', ghiChu: r[8]||'', nguoi: r[9]||'' }); } } return out; }
function importProducts(tk, rows) { _auth(tk); requireRole_(['Admin','Kho']); var ok=0, skip=[]; (rows||[]).forEach(function(r){ try { if (!r.ma || !r.ten) { skip.push(r.ma||'?'); return; } saveProduct(null, r); if (r.tonDau && money_(r.tonDau)>0 && r.khongNhapKho!=='co') {} ok++; } catch(e) { skip.push((r.ma||'?')+': '+e.message); } }); return { ok: ok, skip: skip, products: getProducts() }; }
function templateCsv(tk) { _auth(tk); var lines = ['MaSP,NhomHang,TenSanPham,DVT,DonGia,TonDauKy,TonToiThieu,HinhAnh,GiaVon']; ['NIP001,Nippon,Sơn nội thất Nippon Odour-less,5L,850000,20,5,https://cdn.horme.com.sg/images/product/202512091411431KVYGB15019587_700.jpg,680000','JOT001,Jotun,Sơn nội thất Jotun Essence,5L,780000,20,5,https://www.jotun.com/globalassetsjot03/211316-desktop.png,620000','DUL001,Dulux,Sơn nội thất Dulux Ambiance 5in1,5L,920000,20,5,https://sieuthison.thietkewebbanson.com/image/cache/catalog/san-pham/son-dulux/noi-that/dul032-600x600.png,740000','TOA001,TOA,Sơn nội thất TOA 4 Seasons,5L,650000,20,5,https://down-th.img.susercontent.com/file/th-11134207-7qul5-lg0akjqtnrwvc6,520000','PK001,Phụ kiện,Cọ sơn / rulo / khay sơn,cái,45000,100,20,https://www.homestratosphere.com/wp-content/uploads/2018/08/roller-tray-082218-870x870.jpg,30000'].forEach(function(x){lines.push(x);}); return lines.join('\n'); }

/* ============ THU CHI ============ */
function getCashbook(tk) { _auth(tk); var s = sh_('ThuChi'); if (!s || s.getLastRow()<2) return { rows: [], tongThu:0, tongChi:0 }; var rows = s.getRange(2,1,s.getLastRow()-1,8).getValues().filter(function(r){return r[0];}).map(function(r){ return { ngay:fmtDate_(r[0]), loai:r[1], hangMuc:r[2], soTien:money_(r[3]), phuongThuc:r[4]||'', lienQuan:r[5]||'', ghiChu:r[6]||'', nguoi:r[7]||'' }; }).reverse(); return { rows: rows.slice(0,300), tongThu: rows.filter(function(r){return r.loai==='Thu';}).reduce(function(a,b){return a+b.soTien;},0), tongChi: rows.filter(function(r){return r.loai==='Chi';}).reduce(function(a,b){return a+b.soTien;},0) }; }
function addCashEntry_goc(tk, p) { _auth(tk); requireRole_(['Admin','KeToan']); var s = sh_('ThuChi'); if (!s) { caiDatHeThong(); s = sh_('ThuChi'); } s.appendRow([parseDate_(p.ngay), p.loai, p.hangMuc, money_(p.soTien), p.phuongThuc||'Tiền mặt', p.lienQuan||'', p.ghiChu||'', getCurrentUser().name]); return getCashbook(); }

/* ============ DASHBOARD ============ */
function mondayOf_(d){ var x=new Date(d.getFullYear(), d.getMonth(), d.getDate()); var day=(x.getDay()+6)%7; x.setDate(x.getDate()-day); return x; }
function getDashboard(tk) { _auth(tk);
  var inv = groupInvoices_();
  var now = new Date(); var today = Utilities.formatDate(now, TZ, 'dd/MM/yyyy');
  var mm = Utilities.formatDate(now, TZ, 'MM'); var yyyy = Utilities.formatDate(now, TZ, 'yyyy');
  function dOf(o){ var d = parseDate_(o.ngay); return (d && !isNaN(d.getTime())) ? d : null; }
  function sameDay(d){ return d && Utilities.formatDate(d, TZ, 'dd/MM/yyyy')===today; }
  function sameMonth(d){ return d && Utilities.formatDate(d, TZ, 'MM/yyyy')===(mm+'/'+yyyy); }
  var invToday = inv.filter(function(o){ return sameDay(dOf(o)); });
  var invMonth = inv.filter(function(o){ return sameMonth(dOf(o)); });
  var doanhHomNay = invToday.reduce(function(a,b){return a+b.tongTien;},0);
  var doanhThang = invMonth.reduce(function(a,b){return a+b.tongTien;},0);
  var daThuHomNay = 0; var ts = sh_('ThuTien'); if (ts && ts.getLastRow()>=4) ts.getRange(4,1,ts.getLastRow()-3,5).getValues().forEach(function(r){ if (fmtDate_(r[0])===today) daThuHomNay += money_(r[4]); });
  var laiGop = 0; try { laiGop = getProfit(tk, yyyy+'-'+mm+'-01', Utilities.formatDate(now, TZ, 'yyyy-MM-dd')).tongLai; } catch(e) {}
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
function getXNT(tk, tu, den) { _auth(tk);
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
function getCustomerHistory(tk, maKH) { _auth(tk); requireRole_(['Admin','BanHang','Kho','KeToan','Xem']); var kh=null; getCustomers().forEach(function(c){ if(c.ma===maKH) kh=c; }); var orders = groupInvoices_().filter(function(o){ return o.maKH===maKH; }); var tongSP = 0; var bs = sh_('NhatKyBanHang'); if (bs && bs.getLastRow()>=4) bs.getRange(4,3,bs.getLastRow()-3,6).getValues().forEach(function(r){ if (String(r[0])===String(maKH)) tongSP += money_(r[5]); }); return { kh: kh, orders: orders, tongSP: tongSP }; }
function deleteCustomer(tk, ma) { _auth(tk); requireRole_(['Admin']); if (String(ma)==='KH001') throw new Error('Không xoá được khách mặc định KH001 (Khách lẻ).'); var coDon = groupInvoices_().some(function(o){ return o.maKH===ma; }); if (coDon) throw new Error('Khách đã có đơn mua hàng — không xoá để tránh mất lịch sử.'); var kh=null; getCustomers().forEach(function(c){ if(c.ma===ma) kh=c; }); if (kh && (kh.diem>0 || kh.noHienTai>0)) throw new Error('Khách còn điểm hoặc còn nợ — không xoá được.'); var s = sh_('KhachHang'); var v = s.getLastRow()>=4 ? s.getRange(4,1,s.getLastRow()-3,1).getValues() : []; for (var i=0;i<v.length;i++){ if (String(v[i][0])===String(ma)) { s.getRange(i+4,1,1,9).clearContent(); return getCustomers(); } } throw new Error('Không tìm thấy khách ' + ma); }

/* ============ NHÀ CUNG CẤP + CHĂM SÓC KHÁCH HÀNG (V197) ============ */
function ensureCareTabs_() { var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName('NhaCungCap')) { var a = ss.insertSheet('NhaCungCap'); a.getRange(1,1,1,7).setValues([['Mã NCC','Tên nhà cung cấp','Người liên hệ','SĐT','Địa chỉ','Mã số thuế','Ghi chú']]).setFontWeight('bold'); a.setFrozenRows(1); }
  if (!ss.getSheetByName('ChamSocKH')) { var b = ss.insertSheet('ChamSocKH'); b.getRange(1,1,1,7).setValues([['Mã CS','Ngày','Mã KH','Loại','Nội dung','Ngày hẹn tiếp','Người thực hiện']]).setFontWeight('bold'); b.setFrozenRows(1); }
}
function nextPrefCode_(vals, prefix) { var mx = 0; vals.forEach(function(r){ var m = new RegExp('^' + prefix + '(\\d+)$').exec(String(r[0]||'')); if (m) mx = Math.max(mx, +m[1]); }); var n = mx + 1; return prefix + (n > 999 ? String(n) : ('00' + n).slice(-3)); }
function firstEmptyIdx_(vals) { for (var i = 0; i < vals.length; i++) { if (!vals[i][0]) return i + 2; } return vals.length + 2; }
function readSuppliers_() { var s = sh_('NhaCungCap'); if (!s || s.getLastRow() < 2) return []; return s.getRange(2,1,s.getLastRow()-1,7).getValues().filter(function(r){ return r[0]; }).map(function(r){ return { ma: strCell_(r[0]), ten: strCell_(r[1]), lienHe: strCell_(r[2]), sdt: String(r[3]||''), diaChi: strCell_(r[4]), mst: String(r[5]||''), ghiChu: strCell_(r[6]) }; }); }
function getSuppliers(tk) { _auth(tk); ensureCareTabs_(); return readSuppliers_(); }
function saveSupplier(tk, s, isNew) { var lock = LockService.getScriptLock(); lock.waitLock(20000); try { return saveSupplier_goc(tk, s, isNew); } finally { lock.releaseLock(); } }
function saveSupplier_goc(tk, s, isNew) { _auth(tk); requireRole_(['Admin','Kho']); ensureCareTabs_(); var ten = String((s && s.ten) || '').trim(); if (!ten) throw new Error('Chưa nhập tên nhà cung cấp.');
  var sh = sh_('NhaCungCap'); var vals = sh.getLastRow() >= 2 ? sh.getRange(2,1,sh.getLastRow()-1,7).getValues() : [];
  for (var i = 0; i < vals.length; i++) { if (vals[i][1] && String(vals[i][1]).trim().toLowerCase() === ten.toLowerCase() && String(vals[i][0]) !== String(s.ma || '')) throw new Error('Đã có nhà cung cấp tên "' + ten + '" (mã ' + vals[i][0] + ').'); }
  var rowData = [ten, (s.lienHe || ''), String(s.sdt || ''), (s.diaChi || ''), String(s.mst || ''), (s.ghiChu || '')];
  function ghiNCC_(row, maGhi) { sh.getRange(row,4).setNumberFormat('@'); sh.getRange(row,6).setNumberFormat('@'); sh.getRange(row,1,1,7).setValues([[maGhi].concat(rowData)]); }
  if (!isNew && s.ma) { for (var j = 0; j < vals.length; j++) { if (String(vals[j][0]) === String(s.ma)) { ghiNCC_(j+2, s.ma); return getSuppliers(); } } throw new Error('Không tìm thấy nhà cung cấp ' + s.ma); }
  var ma = nextPrefCode_(vals, 'NCC'); ghiNCC_(firstEmptyIdx_(vals), ma); return getSuppliers(); }
function deleteSupplier(tk, ma) { var lock = LockService.getScriptLock(); lock.waitLock(20000); try { return deleteSupplier_goc(tk, ma); } finally { lock.releaseLock(); } }
function deleteSupplier_goc(tk, ma) { _auth(tk); requireRole_(['Admin','Kho']); ensureCareTabs_(); var sup = null; readSuppliers_().forEach(function(x){ if (x.ma === String(ma)) sup = x; }); if (!sup) throw new Error('Không tìm thấy nhà cung cấp ' + ma);
  var nk = sh_('NhapKho'); var n = 0; if (nk && nk.getLastRow() >= 2) nk.getRange(2,8,nk.getLastRow()-1,1).getValues().forEach(function(r){ if (String(r[0] || '').trim().toLowerCase() === sup.ten.toLowerCase()) n++; });
  if (n > 0) throw new Error('Không xoá được: nhà cung cấp "' + sup.ten + '" đã có ' + n + ' dòng nhập kho. Hãy sửa tên thay vì xoá để giữ lịch sử.');
  var sh = sh_('NhaCungCap'); var vals = sh.getRange(2,1,sh.getLastRow()-1,1).getValues(); for (var i = 0; i < vals.length; i++) { if (String(vals[i][0]) === String(ma)) { sh.getRange(i+2,1,1,7).clearContent(); return getSuppliers(); } } throw new Error('Không tìm thấy nhà cung cấp ' + ma); }
function readCareLogs_(maKH) { var s = sh_('ChamSocKH'); if (!s || s.getLastRow() < 2) return []; var cus = {}; readCustomers_().forEach(function(c){ cus[c.ma] = c; }); var v = s.getRange(2,1,s.getLastRow()-1,7).getValues(); var out = [];
  for (var i = v.length - 1; i >= 0; i--) { var r = v[i]; if (!r[0]) continue; if (maKH && String(r[2]) !== String(maKH)) continue; var c = cus[String(r[2])] || {}; out.push({ ma: strCell_(r[0]), ngay: fmtDate_(r[1]), maKH: strCell_(r[2]), tenKH: c.ten || '', sdt: c.sdt || '', loai: strCell_(r[3]), noiDung: strCell_(r[4]), ngayHen: (r[5] ? fmtDate_(r[5]) : ''), nguoi: strCell_(r[6]) }); if (out.length >= 300) break; }
  return out; }
function getCareLogs(tk, maKH) { _auth(tk); ensureCareTabs_(); return readCareLogs_(maKH || ''); }
function saveCareLog(tk, log, isNew) { var lock = LockService.getScriptLock(); lock.waitLock(20000); try { return saveCareLog_goc(tk, log, isNew); } finally { lock.releaseLock(); } }
function saveCareLog_goc(tk, log, isNew) { _auth(tk); requireRole_(['Admin','BanHang','KeToan']); ensureCareTabs_(); var maKH = String((log && log.maKH) || ''); var okKH = false; readCustomers_().forEach(function(c){ if (c.ma === maKH) okKH = true; }); if (!okKH) throw new Error('Chưa chọn khách hàng hợp lệ.');
  if (!String((log && log.noiDung) || '').trim()) throw new Error('Chưa nhập nội dung chăm sóc.');
  var ngay = parseDate_(log.ngay); if (!ngay) throw new Error('Ngày chăm sóc không hợp lệ.'); var hen = log.ngayHen ? parseDate_(log.ngayHen) : ''; if (log.ngayHen && !hen) throw new Error('Ngày hẹn tiếp không hợp lệ.');
  var sh = sh_('ChamSocKH'); var vals = sh.getLastRow() >= 2 ? sh.getRange(2,1,sh.getLastRow()-1,7).getValues() : [];
  var rowData = [ngay, maKH, String(log.loai || 'Gọi điện'), String(log.noiDung || ''), hen, getCurrentUser().name];
  if (!isNew && log.ma) { for (var j = 0; j < vals.length; j++) { if (String(vals[j][0]) === String(log.ma)) { sh.getRange(j+2,1,1,7).setValues([[log.ma].concat(rowData)]); return readCareLogs_(''); } } throw new Error('Không tìm thấy bản ghi chăm sóc ' + log.ma); }
  var ma = nextPrefCode_(vals, 'CS'); sh.getRange(firstEmptyIdx_(vals),1,1,7).setValues([[ma].concat(rowData)]); return readCareLogs_(''); }
function deleteCareLog(tk, ma) { var lock = LockService.getScriptLock(); lock.waitLock(20000); try { return deleteCareLog_goc(tk, ma); } finally { lock.releaseLock(); } }
function deleteCareLog_goc(tk, ma) { _auth(tk); requireRole_(['Admin','BanHang','KeToan']); ensureCareTabs_(); var sh = sh_('ChamSocKH'); if (!sh || sh.getLastRow() < 2) throw new Error('Chưa có bản ghi chăm sóc nào.'); var vals = sh.getRange(2,1,sh.getLastRow()-1,1).getValues(); for (var i = 0; i < vals.length; i++) { if (String(vals[i][0]) === String(ma)) { sh.getRange(i+2,1,1,7).clearContent(); return readCareLogs_(''); } } throw new Error('Không tìm thấy bản ghi chăm sóc ' + ma); }

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
  ensure('NhaCungCap', ['Mã NCC','Tên nhà cung cấp','Người liên hệ','SĐT','Địa chỉ','Mã số thuế','Ghi chú']);
  ensure('ChamSocKH', ['Mã CS','Ngày','Mã KH','Loại','Nội dung','Ngày hẹn tiếp','Người thực hiện']);
  var sp = sh_('SanPham'); if (sp) { if (!sp.getRange(3,6).getValue()) sp.getRange(3,6).setValue('Tồn đầu kỳ'); if (!sp.getRange(3,7).getValue()) sp.getRange(3,7).setValue('Tồn tối thiểu'); if (!sp.getRange(3,9).getValue()) sp.getRange(3,9).setValue('Giá vốn'); }
  return 'Đã cài xong các tab phụ: CauHinh (kèm công tắc DiemBatTat), DonHang, DiemTichLuy, ThuChi, NhapKho, NguoiDung + cột Giá vốn (I) trong SanPham.';
}
function onOpen() { SpreadsheetApp.getUi().createMenu('🎨 Sơn Bảo Ân').addItem('⚙ Cài đặt hệ thống (tạo tab phụ)', 'caiDatHeThong').addItem('🌐 Mở hướng dẫn Web App', 'moHuongDan_').addToUi(); }
function moHuongDan_() { SpreadsheetApp.getUi().alert('Triển khai Web App: Tiện ích mở rộng → Apps Script → Dán Code.gs + Index.html → Deploy → Web App → Execute as Me → Who has access: Anyone with Google account (để lấy email phân quyền) → Lấy link Web App để dùng hằng ngày.'); }