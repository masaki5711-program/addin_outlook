// 宛先の表示名から「山田様」「関係者各位」などの宛名を作る（Office.js に依存しない純粋関数）

const JAPANESE = /[぀-ヿ㐀-鿿豈-﫿]/;
const COMPANY = /株式会社|有限会社|合同会社|[（(]株[）)]|[（(]有[）)]|㈱|㈲/g;

export const DEFAULT_SALUTATION = {
  style: 'family', // 'family' = 姓のみ / 'full' = フルネーム
  honorific: '様',
  maxNames: 3,
  groupText: '関係者各位',
  unknownName: 'ご担当者',
};

// 「山田 太郎 / Taro Yamada」「Taro Yamada (山田 太郎)」のような併記は日本語の方を使う
export function cleanName(displayName) {
  const segments = displayName
    .replace(/【[^】]*】|\[[^\]]*\]/g, ' ')
    .replace(COMPANY, ' ')
    .split(/[/／|｜（()）]/)
    .map((s) => s.replace(/[\s　]+/g, ' ').trim())
    .filter(Boolean);
  return segments.find((s) => JAPANESE.test(s)) ?? segments[0] ?? '';
}

export function personName(recipient, style) {
  const raw = recipient?.displayName ?? '';
  if (!raw || raw.includes('@')) return null;
  const name = cleanName(raw);
  if (!name) return null;
  if (style === 'family' && JAPANESE.test(name)) return name.split(' ')[0];
  return name;
}

export function buildSalutation(recipients, options = {}) {
  const { style, honorific, maxNames, groupText, unknownName } = { ...DEFAULT_SALUTATION, ...options };
  if (!recipients?.length) return '';
  if (recipients.length > maxNames) return groupText;
  const names = [...new Set(recipients.map((r) => personName(r, style) ?? unknownName))];
  return names.map((n) => (n.endsWith(honorific) ? n : n + honorific)).join('、');
}
