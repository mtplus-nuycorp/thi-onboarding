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
var DEFAULT_MODEL = 'gemini-flash-lite-latest';
var FALLBACK_MODELS = ['gemini-flash-lite-latest', 'gemini-3.1-flash-lite', 'gemini-flash-latest'];

// ================= API =================

function doGet() {
  return json_({ ok: true, msg: 'API thi dang chay' });
}

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    var handlers = { config: apiConfig, login: apiLogin, start: apiStart, submit: apiSubmit, chamTL: apiChamTL, admin: apiAdmin };
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
      tyTrongTN: num_(r[5], 50), diemDat: num_(r[6], 10), mo: yes_(r[7])
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
    var q = { id: id, phan: key_(r[1]), loai: key_(r[2]), text: norm_(r[3]), opts: [], dapAn: [], hd: norm_(r[9]) };
    for (var i = 0; i < 4; i++) q.opts.push(norm_(r[4 + i]));
    q.dapAn = parseAnswer_(r[8]);
    if (q.loai === 'TN') {
      var nOpt = q.opts.filter(function (o) { return o; }).length;
      if (nOpt < 2 || q.dapAn.length === 0) return;
    } else if (q.loai === 'TL') {
      if (!q.hd) return;   // chưa có hướng dẫn chấm -> chưa rút vào đề
    } else return;
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
  cfg.phans.forEach(function (p) { p.lichSu = hist[p.ma] || { lanThi: 0, diemCao: null, dat: false }; });
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
      items.push({ id: q.id, t: 'TL' });
      out.push({ stt: i + 1, loai: 'TL', text: q.text });
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
    batDau: now.getTime(), hetGio: now.getTime() + p.phut * 60000, questions: out
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
      var ct = [it.id, 'TL', q.text, tra, '', '', '', ''];
      chiTiet.push(ct);
      tlJobs.push({ q: q, tra: tra, rv: rv, ct: ct });
    }
  });

  // --- Tự luận: AI chấm ngầm ở lệnh chamTL ngay sau khi nộp, để người thi có điểm trắc nghiệm ngay ---
  tlJobs.forEach(function (j) {
    if (!j.tra) setTL_(j, 0, 'Bỏ trống, không trả lời.');
    else setTL_(j, null, 'AI đang chấm...');
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

  return {
    maDe: maDe, phan: { ma: p.ma, ten: p.ten }, lanThi: lanThi, phutLam: phutLam,
    diemTN: diemTN, dungTN: dungTN, nTN: nTN, diemTL: diemTL, tong: tong,
    diemDat: p.diemDat, ketQua: ketQua, choCham: choCham, review: review
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

/** Chấm nhiều bài tự luận song song. Ghi điểm vào job.rv và job.ct. */
function gradeEssays_(jobs) {
  if (!jobs.length) return;
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('GEMINI_API_KEY');
  var model = props.getProperty('GEMINI_MODEL') || DEFAULT_MODEL;
  var pending = [];
  jobs.forEach(function (j) {
    if (!j.tra) { setTL_(j, 0, 'Bỏ trống, không trả lời.'); return; }
    if (!key) { setTL_(j, null, 'Chưa cài API key chấm tự động, chờ chấm chéo.'); return; }
    pending.push(j);
  });
  // Gói miễn phí hay báo "quá tải" (503): thử lần lượt nhiều model cho tới khi chấm được
  var models = [model].concat(FALLBACK_MODELS.filter(function (m) { return m !== model; }));
  models = models.concat(models);   // chạy 2 vòng, vòng 2 nghỉ 3 giây trước khi thử lại
  var lastErr = '';
  for (var mi = 0; mi < models.length && pending.length; mi++) {
    if (mi === models.length / 2) Utilities.sleep(3000);
    var res;
    try { res = UrlFetchApp.fetchAll(pending.map(function (j) { return geminiRequest_(key, models[mi], j.q, j.tra); })); }
    catch (e) { res = []; lastErr = e.message; }
    var still = [];
    pending.forEach(function (j, k) {
      try {
        var r = res[k];
        if (!r || r.getResponseCode() !== 200) throw new Error('HTTP ' + (r ? r.getResponseCode() : 'loi'));
        var body = JSON.parse(r.getContentText());
        var txt = body.candidates[0].content.parts.map(function (pt) { return pt.text || ''; }).join('');
        var g = JSON.parse(txt);
        var d = Math.max(0, Math.min(10, Number(g.diem)));
        if (isNaN(d)) throw new Error('Không đọc được điểm');
        setTL_(j, round2_(d), norm_(g.nhanXet));
      } catch (e) { lastErr = e.message; still.push(j); }
    });
    pending = still;
  }
  pending.forEach(function (j) { setTL_(j, null, 'Lỗi chấm tự động (' + lastErr + '), chờ chấm chéo.'); });
}

function setTL_(j, diem, nhanXet) {
  j.rv.diem = diem; j.rv.nhanXet = nhanXet;
  j.ct[5] = diem == null ? 'CHỜ CHẤM' : (diem >= 5 ? 'ĐẠT Ý' : 'THIẾU Ý');
  j.ct[6] = diem == null ? '' : diem;
  j.ct[7] = nhanXet;
}

function geminiRequest_(key, model, q, traLoi) {
  var prompt =
    'Bạn là giám khảo chấm bài tự luận nội bộ cho Trưởng phòng/Trưởng nhóm Marketing.\n' +
    'Chấm câu trả lời theo HƯỚNG DẪN CHẤM. Mỗi ý trong hướng dẫn có trọng số ngang nhau. ' +
    'Cho điểm theo mức độ đầy đủ và đúng ý, chấp nhận cách diễn đạt khác nghĩa tương đương. ' +
    'Ví dụ cụ thể, sát thực tế được cộng điểm. Trả lời lan man, sai ý hoặc chép lại đề thì không có điểm.\n\n' +
    'CÂU HỎI:\n' + q.text + '\n\nHƯỚNG DẪN CHẤM:\n' + q.hd + '\n\nBÀI LÀM:\n' + traLoi + '\n\n' +
    'Trả về JSON: {"diem": số từ 0 đến 10 (có thể lẻ 0.5), "nhanXet": "1-3 câu tiếng Việt: đạt ý nào, thiếu ý nào"}';
  return {
    url: 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent',
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-goog-api-key': key },
    payload: JSON.stringify({
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.2, responseMimeType: 'application/json' }
    })
  };
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

// ================= MENU TRONG SHEET =================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Tool thi')
    .addItem('Cài đặt (API key chấm tự luận, mật khẩu quản lý)', 'menuCaiDat')
    .addItem('Chấm lại các câu tự luận đang chờ', 'menuChamLai')
    .addItem('Tính lại điểm sau khi chấm chéo', 'menuTinhLai')
    .addToUi();
}

function menuCaiDat() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var r1 = ui.prompt('API key Gemini', 'Dán API key lấy ở aistudio.google.com (bỏ trống để giữ nguyên):', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  if (norm_(r1.getResponseText())) props.setProperty('GEMINI_API_KEY', norm_(r1.getResponseText()));
  var r2 = ui.prompt('Mật khẩu trang quản lý', 'Đặt mật khẩu để xem trang quản lý (bỏ trống để giữ nguyên):', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  if (norm_(r2.getResponseText())) props.setProperty('ADMIN_PASSWORD', norm_(r2.getResponseText()));
  var r3 = ui.prompt('Model AI (không bắt buộc)', 'Mặc định: ' + DEFAULT_MODEL + '. Bỏ trống để giữ nguyên:', ui.ButtonSet.OK_CANCEL);
  if (r3.getSelectedButton() === ui.Button.OK && norm_(r3.getResponseText())) props.setProperty('GEMINI_MODEL', norm_(r3.getResponseText()));
  var test = testGemini_();
  ui.alert('Đã lưu cài đặt.\n\nKiểm tra AI chấm bài: ' + test);
}

function testGemini_() {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('GEMINI_API_KEY');
  if (!key) return 'chưa có API key (tự luận sẽ chờ chấm chéo)';
  var job = { q: { text: 'Khoe CDN mấy lần một tuần?', hd: 'Khoe 3 lần/tuần' }, tra: '3 lần mỗi tuần', rv: {}, ct: [] };
  gradeEssays_([job]);
  return job.rv.diem == null ? 'LỖI - ' + job.rv.nhanXet : 'OK (bài mẫu được ' + job.rv.diem + ' điểm)';
}

/** Web gọi ngay sau khi nộp: AI chấm các câu tự luận của bài này, trả điểm mới. */
function apiChamTL(req) {
  var id = norm_(req.attemptId);
  if (!id) throw new Error('Thiếu mã bài');
  chamLaiCore_(id, 6);   // chấm bài này + tranh thủ chấm lại tối đa 6 câu đang chờ của bài khác
  var out = { tl: [], diemTL: null, tong: null, ketQua: '' };
  rows_(SHEET.CHI_TIET).forEach(function (r) {
    if (norm_(r[0]) !== id || key_(r[6]) !== 'TL') return;
    var d = norm_(r[13]) !== '' ? Number(r[13]) : (r[11] === '' ? null : Number(r[11]));
    out.tl.push({ id: norm_(r[5]), diem: d, nhanXet: norm_(r[12]) });
  });
  rows_(SHEET.KET_QUA).forEach(function (r) {
    if (norm_(r[13]) !== id) return;
    out.diemTL = r[9] === '' ? null : Number(r[9]);
    out.tong = r[10] === '' ? null : Number(r[10]);
    out.ketQua = norm_(r[11]);
  });
  out.choCham = out.tong == null;
  return out;
}

/** Chấm lại câu TL chưa có điểm (do lỗi mạng/AI), sau đó tính lại tổng. */
function menuChamLai() {
  var n = chamLaiCore_(null);
  SpreadsheetApp.getUi().alert('Đã chấm lại ' + n.cau + ' câu, cập nhật ' + n.bai + ' bài.');
}

function chamLaiCore_(onlyAttempt, extra) {
  var ct = sh_(SHEET.CHI_TIET);
  var data = ct.getDataRange().getValues();
  var bank = {};
  rows_(SHEET.CAU_HOI).forEach(function (r) { bank[norm_(r[0])] = { text: norm_(r[3]), hd: norm_(r[9]) }; });
  var jobs = [], others = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (key_(r[6]) !== 'TL' || r[11] !== '' || norm_(r[13])) continue;
    var q = bank[norm_(r[5])];
    if (!q || !q.hd) continue;
    var job = { q: q, tra: norm_(r[8]), rv: {}, ct: [], row: i + 1 };
    if (!onlyAttempt || norm_(r[0]) === onlyAttempt) jobs.push(job); else others.push(job);
  }
  if (onlyAttempt && extra) jobs = jobs.concat(others.slice(0, extra));
  for (var s = 0; s < jobs.length; s += 10) {
    var batch = jobs.slice(s, s + 10);
    gradeEssays_(batch);
    batch.forEach(function (j) {
      ct.getRange(j.row, 11, 1, 3).setValues([[j.ct[5], j.ct[6], j.ct[7]]]);
    });
  }
  var n = tinhLaiTatCa_();
  return { cau: jobs.length, bai: n };
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
