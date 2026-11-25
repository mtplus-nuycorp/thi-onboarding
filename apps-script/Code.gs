/**
 * TOOL THI ONBOARDING - Backend Google Apps Script
 * Dán toàn bộ file này vào: Google Sheet > Tiện ích mở rộng > Apps Script > Code.gs
 * Đáp án và hướng dẫn chấm chỉ nằm trong Sheet, web chỉ nhận đề đã trộn (không kèm đáp án).
 */

var SHEET = {
  CAU_HOI: 'CAU_HOI', PHAN_THI: 'PHAN_THI', NHAN_VIEN: 'NHAN_VIEN',
  KET_QUA: 'KET_QUA', CHI_TIET: 'CHI_TIET', PHIEN: 'PHIEN'
};
var GRACE_MIN = 3;               // cho phép nộp trễ tối đa 3 phút (mạng chậm)

// ================= API =================

function doGet() {
  return json_({ ok: true, msg: 'API thi dang chay' });
}

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    var handlers = { config: apiConfig, login: apiLogin, start: apiStart, submit: apiSubmit, admin: apiAdmin,
      dsTuLuan: apiDsTuLuan, chamTuLuan: apiChamTuLuan,
      // tương thích trang web bản cũ (đang chờ cập nhật): không trả điểm
      chamTL: function () { return { tl: [], diemTL: null, tong: null, ketQua: 'ĐÃ NỘP BÀI', choCham: false }; } };
    var fn = handlers[req.action];
    if (!fn) throw new Error('Yêu cầu không hợp lệ');
    var out = fn(req);
    out.ok = true;
    return json_(out);
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// ================= ĐỌC DỮ LIỆU =================

function sh_(name) {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) throw new Error('Thiếu tab ' + name);
  return s;
}
function rows_(name) {
  var v = sh_(name).getDataRange().getValues();
  v.shift();
  return v;
}
function norm_(s) { return String(s == null ? '' : s).trim(); }
function key_(s) { return norm_(s).toUpperCase(); }
function yes_(s) { var k = key_(s); return k === 'CÓ' || k === 'CO' || k === 'X' || k === 'TRUE' || k === 'YES'; }
function num_(v, d) { var n = Number(v); return isNaN(n) || v === '' ? d : n; }
function round2_(n) { return Math.round(n * 100) / 100; }

function getPhans_() {
  var map = {};
  rows_(SHEET.PHAN_THI).forEach(function (r) {
    var ma = key_(r[0]);
    if (!ma) return;
    map[ma] = {
      ma: ma, ten: norm_(r[1]), phut: num_(r[2], 30), soTN: num_(r[3], 0), soTL: num_(r[4], 0),
      tyTrongTN: num_(r[5], 50), diemDat: num_(r[6], 10), mo: yes_(r[7]),
      boiCanh: String(r[8] == null ? '' : r[8]).split(/\n+/).map(norm_).filter(function (x) { return x; })
    };
  });
  return map;
}

function getNV_(maNV) {
  var k = key_(maNV);
  if (!k) throw new Error('Vui lòng nhập Mã NV');
  var byMa = null, byTen = null;
  rows_(SHEET.NHAN_VIEN).forEach(function (r) {
    if (!yes_(r[4])) return;
    var nv = nvOf_(r);
    if (key_(r[0]) && key_(r[0]) === k) byMa = byMa || nv;
    else if (key_(r[1]) === k) byTen = byTen || nv;
  });
  var found = byMa || byTen;
  if (!found) throw new Error('Không tìm thấy "' + norm_(maNV) + '". Nhập Mã NV hoặc đúng tên trong danh sách (vd AnhVTN), hoặc liên hệ quản lý.');
  return found;
}

/** Người thi: mã = Mã NV, nếu trống thì dùng Tên làm mã */
function nvOf_(r) {
  var ten = norm_(r[1]);
  return { ma: norm_(r[0]) || ten, ten: ten, nhom: norm_(r[2]), duAn: norm_(r[3]) };
}

/** Đọc ngân hàng câu hỏi. Chỉ lấy câu đang dùng, đủ dữ liệu để chấm. */
function getQuestions_() {
  var list = [];
  rows_(SHEET.CAU_HOI).forEach(function (r) {
    var id = norm_(r[0]);
    if (!id || !yes_(r[10])) return;
    var q = { id: id, phan: key_(r[1]), loai: key_(r[2]), text: norm_(r[3]), opts: [], dapAn: [], hd: norm_(r[9]), bc: yes_(r[12]) };
    for (var i = 0; i < 4; i++) q.opts.push(norm_(r[4 + i]));
    q.dapAn = parseAnswer_(r[8]);
    if (q.loai === 'TN') {
      var nOpt = q.opts.filter(function (o) { return o; }).length;
      if (nOpt < 2 || q.dapAn.length === 0) return;
    } else if (q.loai !== 'TL') return;   // tự luận chấm chéo bằng tay, không cần hướng dẫn chấm mới được rút
    list.push(q);
  });
  return list;
}

/** "B, D" -> [1,3] */
function parseAnswer_(v) {
  var out = [];
  String(v == null ? '' : v).toUpperCase().replace(/[^A-D]/g, ' ').split(/\s+/).forEach(function (ch) {
    if (!ch) return;
    var idx = 'ABCD'.indexOf(ch);
    if (idx >= 0 && out.indexOf(idx) < 0) out.push(idx);
  });
  return out.sort();
}

function shuffle_(arr) {
  var a = arr.slice();
  for (var i = a.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var t = a[i]; a[i] = a[j]; a[j] = t;
  }
  return a;
}

// ================= API: CẤU HÌNH, ĐĂNG NHẬP =================

function apiConfig() {
  var phans = getPhans_();
  var qs = getQuestions_();
  var list = [];
  Object.keys(phans).forEach(function (k) {
    var p = phans[k];
    if (!p.mo) return;
    var nTN = qs.filter(function (q) { return q.phan === k && q.loai === 'TN'; }).length;
    var nTL = qs.filter(function (q) { return q.phan === k && q.loai === 'TL'; }).length;
    list.push({
      ma: k, ten: p.ten, phut: p.phut, diemDat: p.diemDat,
      soTN: p.soTN > 0 ? Math.min(p.soTN, nTN) : nTN,
      soTL: Math.min(p.soTL, nTL)
    });
  });
  return { phans: list };
}

function apiLogin(req) {
  var nv = getNV_(req.maNV);
  var hist = {};
  rows_(SHEET.KET_QUA).forEach(function (r) {
    if (key_(r[1]) !== key_(nv.ma)) return;
    var p = key_(r[4]);
    var h = hist[p] || (hist[p] = { lanThi: 0, diemCao: null, dat: false });
    h.lanThi++;
    if (r[10] !== '' && r[10] != null) h.diemCao = h.diemCao == null ? Number(r[10]) : Math.max(h.diemCao, Number(r[10]));
    if (norm_(r[11]).indexOf('ĐẠT') === 0) h.dat = true;
  });
  var cfg = apiConfig();
  // Không công khai điểm cho thí sinh: chỉ trả số lần đã thi
  cfg.phans.forEach(function (p) { p.lichSu = { lanThi: hist[p.ma] ? hist[p.ma].lanThi : 0, diemCao: null, dat: false }; });
  return { nv: nv, phans: cfg.phans };
}

// ================= API: BẮT ĐẦU THI =================

function apiStart(req) {
  var nv = getNV_(req.maNV);
  var phans = getPhans_();
  var p = phans[key_(req.phan)];
  if (!p || !p.mo) throw new Error('Phần thi này đang đóng');

  var qs = getQuestions_().filter(function (q) { return q.phan === p.ma; });
  var tn = shuffle_(qs.filter(function (q) { return q.loai === 'TN'; }));
  var tl = shuffle_(qs.filter(function (q) { return q.loai === 'TL'; }));
  if (p.soTN > 0) tn = tn.slice(0, p.soTN);
  tl = tl.slice(0, Math.max(0, p.soTL));
  var picked = tn.concat(tl);
  if (!picked.length) throw new Error('Phần thi chưa có câu hỏi');

  // Bối cảnh doanh nghiệp: bốc ngẫu nhiên 1 bối cảnh cho cả đề, dùng cho các câu tự luận có đánh dấu
  var boiCanh = '';
  if (p.boiCanh.length && tl.some(function (q) { return q.bc; })) boiCanh = p.boiCanh[Math.floor(Math.random() * p.boiCanh.length)];

  var items = [];
  var out = [];
  picked.forEach(function (q, i) {
    if (q.loai === 'TN') {
      var idx = [];
      q.opts.forEach(function (o, k) { if (o) idx.push(k); });
      var order = shuffle_(idx);
      items.push({ id: q.id, t: 'TN', o: order });
      out.push({ stt: i + 1, loai: 'TN', text: q.text, options: order.map(function (k) { return q.opts[k]; }) });
    } else {
      var bc = q.bc ? boiCanh : '';
      items.push({ id: q.id, t: 'TL', bc: bc });
      out.push({ stt: i + 1, loai: 'TL', text: q.text + (bc ? '\n\nBối cảnh doanh nghiệp của bạn: ' + bc : '') });
    }
  });

  var attemptId = Utilities.getUuid();
  var maDe = 'MĐ-' + p.ma + '-' + String(Math.floor(10000 + Math.random() * 90000));
  var now = new Date();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    sh_(SHEET.PHIEN).appendRow([attemptId, nv.ma, p.ma, maDe, now, JSON.stringify(items), '']);
  } finally { lock.releaseLock(); }

  return {
    attemptId: attemptId, maDe: maDe, phan: { ma: p.ma, ten: p.ten, phut: p.phut },
    batDau: now.getTime(), hetGio: now.getTime() + p.phut * 60000, boiCanh: boiCanh, questions: out
  };
}

// ================= API: NỘP BÀI & CHẤM =================

function apiSubmit(req) {
  var phienSh = sh_(SHEET.PHIEN);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var row, rowIdx;
  try {
    var data = phienSh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === req.attemptId) { row = data[i]; rowIdx = i + 1; break; }
    }
    if (!row) throw new Error('Không tìm thấy bài thi. Vui lòng thi lại.');
    if (row[6]) throw new Error('Bài này đã nộp rồi.');
    phienSh.getRange(rowIdx, 7).setValue(new Date());
  } finally { lock.releaseLock(); }

  var maNV = row[1], phanMa = row[2], maDe = row[3], batDau = new Date(row[4]);
  var items = JSON.parse(row[5]);
  var nv = getNV_(maNV);
  var p = getPhans_()[key_(phanMa)];
  var bank = {};
  // đọc cả câu đã tắt "Đang dùng" để vẫn chấm được bài đang làm dở
  rows_(SHEET.CAU_HOI).forEach(function (r) {
    var q = { id: norm_(r[0]), loai: key_(r[2]), text: norm_(r[3]), opts: [], dapAn: parseAnswer_(r[8]), hd: norm_(r[9]) };
    for (var k = 0; k < 4; k++) q.opts.push(norm_(r[4 + k]));
    bank[q.id] = q;
  });

  var answers = req.answers || [];
  var now = new Date();
  var phutLam = round2_((now - batDau) / 60000);
  var tre = phutLam > p.phut + GRACE_MIN;

  // --- Chấm trắc nghiệm ---
  var review = [], chiTiet = [], tlJobs = [];
  var nTN = 0, dungTN = 0;
  items.forEach(function (it, i) {
    var q = bank[it.id];
    if (!q) return;
    if (it.t === 'TN') {
      nTN++;
      var chon = (Array.isArray(answers[i]) ? answers[i] : []).map(Number)
        .filter(function (d) { return d >= 0 && d < it.o.length; });
      var chonGoc = chon.map(function (d) { return it.o[d]; }).sort();
      var dung = chonGoc.length === q.dapAn.length && chonGoc.every(function (v, k) { return v === q.dapAn[k]; });
      if (dung) dungTN++;
      var dapAnHienThi = q.dapAn.map(function (g) { return it.o.indexOf(g); }).filter(function (d) { return d >= 0; }).sort();
      review.push({ stt: i + 1, loai: 'TN', text: q.text, options: it.o.map(function (g) { return q.opts[g]; }), chon: chon.sort(), dapAn: dapAnHienThi, dung: dung });
      chiTiet.push([it.id, 'TN', q.text, lettersOf_(chonGoc), lettersOf_(q.dapAn), dung ? 'ĐÚNG' : 'SAI', dung ? 10 : 0, '']);
    } else {
      var tra = norm_(answers[i]);
      var rv = { stt: i + 1, id: it.id, loai: 'TL', text: q.text, traLoi: tra, diem: null, nhanXet: '' };
      review.push(rv);
      var ct = [it.id, 'TL', q.text + (it.bc ? '\n[Bối cảnh: ' + it.bc + ']' : ''), tra, '', '', '', ''];
      chiTiet.push(ct);
      tlJobs.push({ q: q, tra: tra, rv: rv, ct: ct });
    }
  });

  // --- Tự luận: chấm chéo bằng tay (điền cột "Điểm chấm chéo" trong CHI_TIET rồi menu Tính lại điểm) ---
  tlJobs.forEach(function (j) {
    if (!j.tra) setTL_(j, 0, 'Bỏ trống, không trả lời.');
    else setTL_(j, null, '');
  });

  var diemTN = nTN ? round2_(dungTN / nTN * 10) : null;
  var diemTL = null, choCham = false;
  if (tlJobs.length) {
    var s = 0;
    tlJobs.forEach(function (j) { if (j.rv.diem == null) choCham = true; else s += j.rv.diem; });
    if (!choCham) diemTL = round2_(s / tlJobs.length);
  }
  var tong = tinhTong_(diemTN, diemTL, tlJobs.length > 0, choCham, p.tyTrongTN);
  var ketQua = choCham ? 'CHỜ CHẤM TỰ LUẬN' : (tong >= p.diemDat - 1e-9 ? 'ĐẠT' : 'CHƯA ĐẠT');
  if (tre) ketQua += ' (nộp trễ)';

  // --- Ghi kết quả ---
  lock.waitLock(20000);
  var lanThi = 1;
  try {
    rows_(SHEET.KET_QUA).forEach(function (r) { if (key_(r[1]) === key_(nv.ma) && key_(r[4]) === p.ma) lanThi++; });
    var kq = sh_(SHEET.KET_QUA);
    kq.getRange(kq.getLastRow() + 1, 1, 1, 14).setValues([[
      now, nv.ma, nv.ten, nv.nhom, p.ma, p.ten, lanThi, maDe,
      diemTN == null ? '' : diemTN, diemTL == null ? '' : diemTL, tong == null ? '' : tong, ketQua, phutLam, req.attemptId
    ]]);
    var ctSh = sh_(SHEET.CHI_TIET);
    var ctRows = chiTiet.map(function (c) {
      return [req.attemptId, now, nv.ma, nv.ten, p.ma, c[0], c[1], c[2], c[3], c[4], c[5], c[6], c[7], ''];
    });
    if (ctRows.length) ctSh.getRange(ctSh.getLastRow() + 1, 1, ctRows.length, 14).setValues(ctRows);
  } finally { lock.releaseLock(); }

  // Không công khai điểm, đáp án hay đúng/sai cho thí sinh. Điểm chỉ lưu trong Sheet và trang quản lý.
  return {
    daNop: true, maDe: maDe, phan: { ma: p.ma, ten: p.ten }, lanThi: lanThi, phutLam: phutLam, nTL: tlJobs.length,
    // các trường dưới đây chỉ để trang web bản cũ không bị lỗi, không chứa điểm
    ketQua: 'ĐÃ NỘP BÀI', tong: null, diemTN: null, dungTN: 0, nTN: 0, diemTL: null, diemDat: null, choCham: false,
    review: review.filter(function (x) { return x.loai === 'TL'; }).map(function (x) { return { stt: x.stt, id: x.id, loai: 'TL', text: x.text, traLoi: x.traLoi, diem: null, nhanXet: '' }; })
  };
}

function tinhTong_(diemTN, diemTL, coTL, choCham, tyTrongTN) {
  if (choCham) return null;
  if (diemTN == null) return diemTL;
  if (!coTL) return diemTN;
  var w = tyTrongTN / 100;
  return round2_(diemTN * w + diemTL * (1 - w));
}

function lettersOf_(arr) { return arr.map(function (g) { return 'ABCD'.charAt(g); }).join(','); }

function setTL_(j, diem, nhanXet) {
  j.rv.diem = diem; j.rv.nhanXet = nhanXet;
  j.ct[5] = diem == null ? 'CHỜ CHẤM' : (diem >= 5 ? 'ĐẠT Ý' : 'THIẾU Ý');
  j.ct[6] = diem == null ? '' : diem;
  j.ct[7] = nhanXet;
}

// ================= API: QUẢN LÝ =================

function apiAdmin(req) {
  var pass = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!pass) throw new Error('Chưa đặt mật khẩu quản lý (menu Tool thi > Cài đặt trong Sheet)');
  if (norm_(req.password) !== pass) throw new Error('Sai mật khẩu');

  var phans = getPhans_();
  var nvs = rows_(SHEET.NHAN_VIEN).filter(function (r) { return (norm_(r[0]) || norm_(r[1])) && yes_(r[4]); }).map(nvOf_);
  var ketQua = rows_(SHEET.KET_QUA).filter(function (r) { return norm_(r[1]); }).map(function (r) {
    return {
      tg: r[0] instanceof Date ? r[0].getTime() : null, ma: norm_(r[1]), ten: norm_(r[2]), nhom: norm_(r[3]),
      phan: key_(r[4]), lan: num_(r[6], 1), maDe: norm_(r[7]),
      tn: r[8] === '' ? null : Number(r[8]), tl: r[9] === '' ? null : Number(r[9]),
      tong: r[10] === '' ? null : Number(r[10]), kq: norm_(r[11]), phut: num_(r[12], null), maBai: norm_(r[13])
    };
  });
  // Thống kê từng câu + ai sai câu gì
  var cau = {}, sai = [];
  rows_(SHEET.CHI_TIET).forEach(function (r) {
    if (!norm_(r[0])) return;
    var id = norm_(r[5]), loai = key_(r[6]);
    var diem = r[13] !== '' && r[13] != null ? Number(r[13]) : (r[11] === '' ? null : Number(r[11]));
    var isSai = loai === 'TN' ? key_(r[10]) === 'SAI' : (diem != null && diem < 5);
    var c = cau[id] || (cau[id] = { id: id, phan: key_(r[4]), loai: loai, text: norm_(r[7]), luot: 0, sai: 0 });
    c.luot++;
    if (isSai) {
      c.sai++;
      sai.push({
        maBai: norm_(r[0]), tg: r[1] instanceof Date ? r[1].getTime() : null, ma: norm_(r[2]), ten: norm_(r[3]),
        phan: key_(r[4]), id: id, loai: loai, traLoi: norm_(r[8]), dapAn: norm_(r[9]), diem: diem, nhanXet: norm_(r[12])
      });
    }
  });
  return {
    phans: Object.keys(phans).map(function (k) { return phans[k]; }),
    nhanVien: nvs, ketQua: ketQua,
    cau: Object.keys(cau).map(function (k) { return cau[k]; }), sai: sai,
    sheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl()
  };
}

// ================= API: CHẤM TỰ LUẬN TRÊN WEB QUẢN LÝ =================

function checkPass_(req) {
  var pass = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!pass || norm_(req.password) !== pass) throw new Error('Sai mật khẩu');
}

/** Danh sách câu tự luận để chấm: kèm câu hỏi, bối cảnh, hướng dẫn chấm, bài làm, điểm hiện tại. */
function apiDsTuLuan(req) {
  checkPass_(req);
  var hd = {};
  rows_(SHEET.CAU_HOI).forEach(function (r) { hd[norm_(r[0])] = norm_(r[9]); });
  var maDe = {};
  rows_(SHEET.KET_QUA).forEach(function (r) { maDe[norm_(r[13])] = norm_(r[7]); });
  var list = [];
  var data = sh_(SHEET.CHI_TIET).getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (key_(r[6]) !== 'TL' || !norm_(r[0])) continue;
    var diem = norm_(r[13]) !== '' ? Number(r[13]) : (r[11] === '' ? null : Number(r[11]));
    list.push({
      row: i + 1, maBai: norm_(r[0]), tg: r[1] instanceof Date ? r[1].getTime() : null, ma: norm_(r[2]), ten: norm_(r[3]),
      phan: key_(r[4]), id: norm_(r[5]), cauHoi: norm_(r[7]), traLoi: norm_(r[8]), huongDan: hd[norm_(r[5])] || '',
      diem: diem, nhanXet: norm_(r[12]), maDe: maDe[norm_(r[0])] || ''
    });
  }
  return { list: list };
}

/** Lưu điểm chấm tự luận (0-10) + nhận xét, rồi tính lại tổng điểm. */
function apiChamTuLuan(req) {
  checkPass_(req);
  var items = req.items || [];
  var sh = sh_(SHEET.CHI_TIET);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  var saved = 0;
  try {
    items.forEach(function (it) {
      var row = Number(it.row);
      if (!row || row < 2) return;
      var cur = sh.getRange(row, 1, 1, 14).getValues()[0];
      if (norm_(cur[0]) !== norm_(it.maBai) || norm_(cur[5]) !== norm_(it.id)) throw new Error('Dữ liệu đã thay đổi, hãy tải lại trang.');
      var d = Number(it.diem);
      if (it.diem === '' || it.diem == null || isNaN(d) || d < 0 || d > 10) throw new Error('Điểm phải từ 0 đến 10');
      sh.getRange(row, 11, 1, 4).setValues([[d >= 5 ? 'ĐẠT Ý' : 'THIẾU Ý', cur[11], norm_(it.nhanXet), d]]);
      saved++;
    });
  } finally { lock.releaseLock(); }
  tinhLaiTatCa_();
  return { saved: saved };
}

// ================= MENU TRONG SHEET =================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Tool thi')
    .addItem('Đổi mật khẩu trang quản lý', 'menuCaiDat')
    .addItem('Tính lại điểm sau khi chấm chéo', 'menuTinhLai')
    .addToUi();
}

function menuCaiDat() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Mật khẩu trang quản lý', 'Nhập mật khẩu mới (bỏ trống để giữ nguyên):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK || !norm_(r.getResponseText())) return;
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', norm_(r.getResponseText()));
  ui.alert('Đã đổi mật khẩu.');
}

function menuTinhLai() {
  var n = tinhLaiTatCa_();
  SpreadsheetApp.getUi().alert('Đã tính lại điểm cho ' + n + ' bài.');
}

/** Tính lại điểm TL + tổng cho mọi bài, ưu tiên cột "Điểm chấm chéo" nếu có. */
function tinhLaiTatCa_() {
  var phans = getPhans_();
  var tl = {};
  rows_(SHEET.CHI_TIET).forEach(function (r) {
    if (key_(r[6]) !== 'TL') return;
    var m = norm_(r[0]);
    var t = tl[m] || (tl[m] = { n: 0, s: 0, cho: false });
    t.n++;
    var d = norm_(r[13]) !== '' ? Number(r[13]) : (r[11] === '' ? null : Number(r[11]));
    if (d == null || isNaN(d)) t.cho = true; else t.s += d;
  });
  var kq = sh_(SHEET.KET_QUA);
  var data = kq.getDataRange().getValues();
  var changed = 0;
  for (var i = 1; i < data.length; i++) {
    var r = data[i], m = norm_(r[13]), p = phans[key_(r[4])];
    if (!m || !p || !tl[m]) continue;
    var t = tl[m];
    var diemTN = r[8] === '' ? null : Number(r[8]);
    var diemTL = t.cho ? null : round2_(t.s / t.n);
    var tong = tinhTong_(diemTN, diemTL, true, t.cho, p.tyTrongTN);
    var tre = norm_(r[11]).indexOf('nộp trễ') >= 0;
    var ketQua = t.cho ? 'CHỜ CHẤM TỰ LUẬN' : (tong >= p.diemDat - 1e-9 ? 'ĐẠT' : 'CHƯA ĐẠT');
    if (tre) ketQua += ' (nộp trễ)';
    kq.getRange(i + 1, 10, 1, 3).setValues([[diemTL == null ? '' : diemTL, tong == null ? '' : tong, ketQua]]);
    changed++;
  }
  return changed;
}
