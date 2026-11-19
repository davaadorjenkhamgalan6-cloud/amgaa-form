/**
 * Үзлэгийн хүсэлт — зөвшөөрлийн урсгал (Google Apps Script)
 *
 * 1. Менежер формыг илгээнэ → doPost → Sheet-д "Хүлээгдэж буй" төлөвтэй бичигдэнэ
 * 2. Захирал болон танд "Зөвшөөрөх / Татгалзах" товчтой имэйл очно
 * 3. Аль нэг нь шийдвэр гаргамагц → менежерт имэйл, нөгөө хүнд мэдэгдэл очно
 *
 * СУУЛГАХ: README-ийн "Apps Script суулгах" хэсгийг үзнэ үү.
 */

// ====== ЗӨВХӨН ЭНЭ ХЭСГИЙГ ӨӨРЧИЛНӨ ======
const CONFIG = {
  APPROVERS: [
    { key: 'director', name: 'Захирал',     email: 'balsan.b@bodi-insurance.com' },
    { key: 'me',       name: 'Энх-Амгалан', email: 'enkh-amgalan.d@bodi-insurance.com' },
  ],
  TIMEZONE: 'Asia/Ulaanbaatar',
  // Хоосон орхивол автоматаар олно. Имэйлийн товч ажиллахгүй бол
  // Deploy хийсний дараах ".../exec" холбоосыг энд тавина.
  WEB_APP_URL: 'https://script.google.com/macros/s/AKfycby2BlODHGqv6pkayWV74rQgsOlFgEZzMyb8a3ME4F0Bus0QcTlhldrWwlDQz7FrhcTCLQ/exec',
};
// ==========================================

const HEADERS = [
  'id', 'Илгээсэн огноо', 'Төлөв', 'Шийдвэр гаргасан', 'Шийдвэрийн огноо', 'Татгалзсан шалтгаан',
  'Объект', 'Хаяг', 'Үзлэгийн төрөл', 'Яаралтай байдал', 'Тайлбар', 'Хүссэн огноо',
  'Хүсэлт гаргагч', 'Хүсэлт гаргагчийн имэйл',
];
const COL = Object.fromEntries(HEADERS.map((h, i) => [h, i + 1]));
const STATUS = { pending: 'Хүлээгдэж буй', approved: 'Зөвшөөрсөн', rejected: 'Татгалзсан' };
const TYPE_LABEL = { electrical: 'Цахилгаан', fire: 'Гал', both: 'Цахилгаан ба гал' };
const PRIORITY_LABEL = { high: 'Өндөр (ажлын 3 өдөр)', medium: 'Дунд (ажлын 7 өдөр)', low: 'Бага (ажлын 14 өдөр)' };

/** Нэг удаа гараар ажиллуулна: Sheet үүсгэж, нууц түлхүүр хадгалж, имэйлийн эрх авна. */
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SECRET')) props.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  let ss;
  const id = props.getProperty('SHEET_ID');
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create('Үзлэгийн хүсэлтүүд');
    props.setProperty('SHEET_ID', ss.getId());
  }
  const sh = ss.getSheets()[0];
  // Огноо, дугаарыг Sheets өөрөө хөрвүүлэхгүйн тулд бүх нүдийг текст болгоно
  sh.getRange(1, 1, sh.getMaxRows(), HEADERS.length).setNumberFormat('@');
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  MailApp.getRemainingDailyQuota(); // имэйл илгээх эрх асуулгах
  Logger.log('Бэлэн. Sheet: ' + ss.getUrl());
}

// ---------- 1. Форм илгээх ----------
function doPost(e) {
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const req = {
      object_name: clean(d.object_name, 200),
      address: clean(d.address, 300),
      request_type: TYPE_LABEL[d.request_type] ? d.request_type : '',
      priority: PRIORITY_LABEL[d.priority] ? d.priority : '',
      description: clean(d.description, 3000),
      preferred_date: /^\d{4}-\d{2}-\d{2}$/.test(d.preferred_date || '') ? d.preferred_date : '',
      requested_by: clean(d.requested_by, 120),
      requester_email: clean(d.requester_email, 200),
    };
    const missing = ['object_name', 'address', 'request_type', 'priority', 'description', 'requested_by', 'requester_email']
      .filter(k => !req[k]);
    if (missing.length) return json({ ok: false, error: 'Заавал бөглөх талбар дутуу: ' + missing.join(', ') });
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(req.requester_email)) return json({ ok: false, error: 'Имэйл хаяг буруу байна' });

    const id = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyMMdd-HHmmss') + '-' + Utilities.getUuid().slice(0, 4);
    sheet().appendRow([
      id, now(), STATUS.pending, '', '', '',
      req.object_name, req.address, TYPE_LABEL[req.request_type], PRIORITY_LABEL[req.priority],
      req.description, req.preferred_date, req.requested_by, req.requester_email,
    ]);

    CONFIG.APPROVERS.forEach(a => {
      MailApp.sendEmail({
        to: a.email,
        subject: 'Шинэ үзлэгийн хүсэлт: ' + req.object_name + ' — зөвшөөрөх үү?',
        htmlBody: approverEmailHtml(id, a, req),
        replyTo: req.requester_email,
        name: 'Үзлэгийн хүсэлт',
      });
    });
    return json({ ok: true, id: id });
  } catch (err) {
    console.error(err);
    return json({ ok: false, error: 'Серверийн алдаа: ' + err.message });
  }
}

// ---------- 2. Имэйлийн товч дарахад нээгдэх хуудас ----------
function doGet(e) {
  const p = (e && e.parameter) || {};
  const approver = CONFIG.APPROVERS.find(a => a.key === p.who);
  if (!p.id || !approver || !validSig(p.id, p.who, p.sig)) return page('Холбоос буруу байна', '<p>Энэ холбоос хүчингүй байна.</p>');

  const row = findRow(p.id);
  if (!row) return page('Хүсэлт олдсонгүй', '<p>Хүсэлт олдсонгүй.</p>');
  const r = readRow(row);
  if (r.status !== STATUS.pending) return page('Шийдвэр гарсан', alreadyHtml(r));

  const action = p.action === 'reject' ? 'reject' : 'approve';
  const form = action === 'approve'
    ? '<button id="b" class="ok" onclick="go()">Зөвшөөрөх</button>'
    : '<label>Татгалзсан шалтгаан <span style="color:#dc2626">*</span></label>' +
      '<textarea id="reason" placeholder="Менежерт энэ шалтгаан имэйлээр очно"></textarea>' +
      '<button id="b" class="no" onclick="go()">Татгалзах</button>';
  const script =
    '<script>function go(){var r=document.getElementById("reason");var reason=r?r.value.trim():"";' +
    'if(r&&!reason){alert("Шалтгаанаа бичнэ үү");return;}' +
    'var b=document.getElementById("b");b.disabled=true;b.textContent="Түр хүлээнэ үү…";' +
    'google.script.run.withSuccessHandler(function(m){document.getElementById("act").innerHTML="<p class=\\"done\\">"+m+"</p>";})' +
    '.withFailureHandler(function(err){b.disabled=false;b.textContent="Дахин оролдох";alert(err.message);})' +
    '.decide(' + JSON.stringify(p.id) + ',' + JSON.stringify(p.who) + ',' + JSON.stringify(p.sig) + ',' + JSON.stringify(action) + ',reason);}</script>';
  return page(action === 'approve' ? 'Хүсэлт зөвшөөрөх' : 'Хүсэлт татгалзах',
    detailsHtml(r) + '<div id="act">' + form + '</div>' + script);
}

// ---------- 3. Шийдвэр (хуудасны товчноос дуудагдана) ----------
function decide(id, who, sig, action, reason) {
  const approver = CONFIG.APPROVERS.find(a => a.key === who);
  if (!approver || !validSig(id, who, sig)) throw new Error('Холбоос хүчингүй');
  if (action !== 'approve' && action !== 'reject') throw new Error('Буруу үйлдэл');
  reason = clean(reason, 1000);
  if (action === 'reject' && !reason) throw new Error('Татгалзсан шалтгаан бичнэ үү');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let r;
  try {
    const row = findRow(id);
    if (!row) throw new Error('Хүсэлт олдсонгүй');
    r = readRow(row);
    if (r.status !== STATUS.pending) return 'Энэ хүсэлтэд ' + r.decidedBy + ' аль хэдийн "' + r.status + '" гэж шийдвэр гаргасан байна.';
    r.status = action === 'approve' ? STATUS.approved : STATUS.rejected;
    r.decidedBy = approver.name;
    r.decidedAt = now();
    r.reason = reason;
    const sh = sheet();
    sh.getRange(row, COL['Төлөв'], 1, 4).setValues([[r.status, r.decidedBy, r.decidedAt, r.reason]]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  // Менежерт
  MailApp.sendEmail({
    to: r.requesterEmail,
    subject: (action === 'approve' ? 'Үзлэгийн хүсэлт зөвшөөрөгдлөө: ' : 'Үзлэгийн хүсэлт татгалзагдлаа: ') + r.object,
    htmlBody: requesterEmailHtml(r, action),
    name: 'Үзлэгийн хүсэлт',
  });
  // Нөгөө хянагчид мэдэгдэл
  CONFIG.APPROVERS.filter(a => a.key !== who).forEach(a => {
    MailApp.sendEmail({
      to: a.email,
      subject: 'Шийдвэр гарлаа (' + r.status + '): ' + r.object,
      htmlBody: '<p>' + esc(approver.name) + ' энэ хүсэлтийг <b>' + esc(r.status) + '</b>.' +
        (reason ? '<br>Шалтгаан: ' + esc(reason) : '') + '</p>' + detailsHtml(r) +
        '<p style="color:#6b7280">Танаас нэмэлт үйлдэл шаардахгүй.</p>',
      name: 'Үзлэгийн хүсэлт',
    });
  });
  return action === 'approve'
    ? 'Зөвшөөрлөө. ' + r.requestedBy + ' (' + r.requesterEmail + ') руу мэдэгдэл явлаа.'
    : 'Татгалзлаа. ' + r.requestedBy + ' (' + r.requesterEmail + ') руу шалтгаантай мэдэгдэл явлаа.';
}

// ---------- Туслах функцууд ----------
function sheet() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('setup() функцийг эхлээд ажиллуулна уу');
  return SpreadsheetApp.openById(id).getSheets()[0];
}
function findRow(id) {
  const sh = sheet();
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}
function readRow(row) {
  const v = sheet().getRange(row, 1, 1, HEADERS.length).getValues()[0];
  const g = h => {
    const x = v[COL[h] - 1];
    return x instanceof Date ? Utilities.formatDate(x, CONFIG.TIMEZONE, h === 'Хүссэн огноо' ? 'yyyy-MM-dd' : 'yyyy-MM-dd HH:mm') : x;
  };
  return {
    id: g('id'), createdAt: g('Илгээсэн огноо'), status: g('Төлөв'), decidedBy: g('Шийдвэр гаргасан'),
    decidedAt: g('Шийдвэрийн огноо'), reason: g('Татгалзсан шалтгаан'), object: g('Объект'), address: g('Хаяг'),
    type: g('Үзлэгийн төрөл'), priority: g('Яаралтай байдал'), description: g('Тайлбар'),
    preferredDate: g('Хүссэн огноо'), requestedBy: g('Хүсэлт гаргагч'), requesterEmail: g('Хүсэлт гаргагчийн имэйл'),
  };
}
function sign(id, who) {
  const secret = PropertiesService.getScriptProperties().getProperty('SECRET');
  if (!secret) throw new Error('setup() функцийг эхлээд ажиллуулна уу');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(id + '|' + who, secret)).replace(/=+$/, '');
}
function validSig(id, who, sig) {
  return !!sig && sign(id, who) === sig;
}
function webAppUrl() {
  return CONFIG.WEB_APP_URL || ScriptApp.getService().getUrl();
}
function link(id, who, action) {
  return webAppUrl() + '?action=' + action + '&id=' + encodeURIComponent(id) + '&who=' + who + '&sig=' + sign(id, who);
}
function now() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm');
}
function clean(s, max) {
  return String(s == null ? '' : s).trim().slice(0, max);
}
function esc(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function detailsHtml(r) {
  const rows = [
    ['Объект', r.object], ['Хаяг', r.address], ['Үзлэгийн төрөл', r.type], ['Яаралтай байдал', r.priority],
    ['Тайлбар', r.description], ['Хүссэн огноо', r.preferredDate || '—'],
    ['Хүсэлт гаргагч', r.requestedBy + ' (' + r.requesterEmail + ')'], ['Илгээсэн', r.createdAt],
  ];
  return '<table style="border-collapse:collapse;font-size:14px;width:100%;max-width:560px">' +
    rows.map(x => '<tr><td style="padding:6px 10px;color:#6b7280;vertical-align:top;white-space:nowrap">' + esc(x[0]) +
      '</td><td style="padding:6px 10px;white-space:pre-wrap">' + esc(x[1]) + '</td></tr>').join('') + '</table>';
}
function approverEmailHtml(id, approver, req) {
  const r = {
    object: req.object_name, address: req.address, type: TYPE_LABEL[req.request_type], priority: PRIORITY_LABEL[req.priority],
    description: req.description, preferredDate: req.preferred_date, requestedBy: req.requested_by,
    requesterEmail: req.requester_email, createdAt: now(),
  };
  const btn = (href, text, bg) => '<a href="' + esc(href) + '" style="display:inline-block;padding:12px 22px;margin:4px 8px 4px 0;' +
    'background:' + bg + ';color:#fff;text-decoration:none;border-radius:8px;font-weight:600">' + text + '</a>';
  return '<div style="font-family:Segoe UI,Roboto,Arial,sans-serif;color:#111827">' +
    '<p>Сайн байна уу, ' + esc(approver.name) + '.</p><p>Шинэ үзлэгийн хүсэлт ирлээ. Зөвшөөрөх үү?</p>' +
    detailsHtml(r) + '<p style="margin-top:20px">' +
    btn(link(id, approver.key, 'approve'), '✔ Зөвшөөрөх', '#16a34a') +
    btn(link(id, approver.key, 'reject'), '✖ Татгалзах', '#dc2626') + '</p>' +
    '<p style="color:#6b7280;font-size:12px">Аль нэг хянагч шийдвэр гаргахад хангалттай. Дугаар: ' + esc(id) + '</p></div>';
}
function requesterEmailHtml(r, action) {
  const head = action === 'approve'
    ? '<p>Таны үзлэгийн хүсэлтийг <b style="color:#16a34a">зөвшөөрлөө</b>. Үзлэгийн хугацааг тохирохоор удахгүй тантай холбогдоно.</p>'
    : '<p>Таны үзлэгийн хүсэлтийг <b style="color:#dc2626">татгалзлаа</b>.</p><p><b>Шалтгаан:</b> ' + esc(r.reason) + '</p>';
  return '<div style="font-family:Segoe UI,Roboto,Arial,sans-serif;color:#111827"><p>Сайн байна уу, ' + esc(r.requestedBy) + '.</p>' +
    head + detailsHtml(r) + '</div>';
}
function alreadyHtml(r) {
  return '<p>Энэ хүсэлтэд <b>' + esc(r.decidedBy) + '</b> ' + esc(r.decidedAt) + '-д <b>"' + esc(r.status) + '"</b> гэж шийдвэр гаргасан байна.</p>' +
    (r.reason ? '<p>Шалтгаан: ' + esc(r.reason) + '</p>' : '') + detailsHtml(r);
}
function page(title, body) {
  const html = '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>body{margin:0;padding:24px 16px;font-family:system-ui,Segoe UI,Roboto,sans-serif;background:#f3f4f6;color:#111827}' +
    '.card{max-width:640px;margin:0 auto;background:#fff;border-radius:12px;padding:28px;box-shadow:0 1px 3px rgba(0,0,0,.08)}' +
    'h1{margin:0 0 16px;font-size:22px}label{display:block;font-weight:600;font-size:14px;margin:16px 0 6px}' +
    'textarea{width:100%;box-sizing:border-box;min-height:90px;padding:10px;font:inherit;border:1px solid #d1d5db;border-radius:8px}' +
    'button{margin-top:20px;width:100%;padding:12px;font-size:16px;font-weight:600;color:#fff;border:0;border-radius:8px;cursor:pointer}' +
    'button.ok{background:#16a34a}button.no{background:#dc2626}button:disabled{opacity:.6;cursor:wait}' +
    '.done{margin-top:20px;padding:12px;border-radius:8px;background:#ecfdf5;color:#065f46}</style></head>' +
    '<body><div class="card"><h1>' + esc(title) + '</h1>' + body + '</div></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle(title);
}
