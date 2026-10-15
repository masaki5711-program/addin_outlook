// テンプレートの変数を展開する（Office.js に依存しない純粋関数）
//
//   {宛名} {自分} {件名}            … メールの情報
//   {今日} {明日} {明後日} {今日+7}  … 日付
//   {3営業日後} {来週月曜}          … 土日を除いた日付 / 来週の曜日
//   {?項目名} {?項目名:既定値}      … 挿入前に入力欄を出す

const WEEKDAYS = '日月火水木金土';
const INPUT_RE = /\{\?([^{}:]+)(?::([^{}]*))?\}/g;

export function formatDate(date) {
  return `${date.getMonth() + 1}月${date.getDate()}日(${WEEKDAYS[date.getDay()]})`;
}

export function addDays(date, n) {
  const d = new Date(date);
  d.setDate(d.getDate() + n);
  return d;
}

export function addBusinessDays(date, n) {
  const d = new Date(date);
  let left = n;
  while (left > 0) {
    d.setDate(d.getDate() + 1);
    if (d.getDay() !== 0 && d.getDay() !== 6) left--;
  }
  return d;
}

// 月曜始まりで「来週」の指定曜日（weekday は 0=日 … 6=土）
export function nextWeekday(date, weekday) {
  const d = new Date(date);
  const fromMonday = (d.getDay() + 6) % 7;
  d.setDate(d.getDate() - fromMonday + 7 + ((weekday + 6) % 7));
  return d;
}

export function extractInputs(text) {
  const seen = new Map();
  for (const m of text.matchAll(INPUT_RE)) {
    const label = m[1].trim();
    if (!seen.has(label)) seen.set(label, m[2] ?? '');
  }
  return [...seen].map(([label, defaultValue]) => ({ label, defaultValue }));
}

export function render(text, ctx = {}, inputs = {}) {
  const now = ctx.now ?? new Date();
  const fixed = {
    宛名: ctx.salutation ?? '',
    自分: ctx.myName ?? '',
    件名: ctx.subject ?? '',
    今日: formatDate(now),
    明日: formatDate(addDays(now, 1)),
    明後日: formatDate(addDays(now, 2)),
  };
  return text
    .replace(/\{今日([+-]\d+)\}/g, (_, n) => formatDate(addDays(now, Number(n))))
    .replace(/\{(\d+)営業日後\}/g, (_, n) => formatDate(addBusinessDays(now, Number(n))))
    .replace(/\{来週([月火水木金土日])曜?\}/g, (_, w) => formatDate(nextWeekday(now, WEEKDAYS.indexOf(w))))
    .replace(/\{(宛名|自分|件名|今日|明日|明後日)\}/g, (_, key) => fixed[key])
    // 入力値は最後に差し込む（入力された文字列の中の {…} は展開しない）
    .replace(INPUT_RE, (_, label, def) => inputs[label.trim()] || def || `［${label.trim()}］`);
}

export function escapeHtml(s) {
  return s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
}

// Outlook の HTML エディタと同じく 1 行 1 <div> にする
export function textToHtml(text) {
  return text
    .split(/\r?\n/)
    .map((line) => `<div>${escapeHtml(line).replace(/ {2}/g, ' &nbsp;') || '<br>'}</div>`)
    .join('');
}
