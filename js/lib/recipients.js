// テンプレートの To / CC 欄（「山田 太郎 <yamada@example.com>; sato@example.com」）を扱う純粋関数

const EMAIL = /^[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+$/;

export function parseRecipients(text) {
  return (text ?? '')
    .split(/[;；,、\n]/)
    .map((s) => s.trim())
    .filter(Boolean)
    .map((s) => {
      const m = s.match(/^(.*?)\s*<([^<>]+)>$/);
      return {
        displayName: m ? m[1].replace(/^["']|["']$/g, '').trim() : '',
        emailAddress: (m ? m[2] : s).trim(),
      };
    });
}

export function invalidAddresses(recipients) {
  return recipients.filter((r) => !EMAIL.test(r.emailAddress)).map((r) => r.emailAddress);
}

// existing に入っていないものだけ返す（大文字小文字は区別しない・重複も除く）
export function missingRecipients(existing, wanted) {
  const have = new Set(existing.map((r) => r.emailAddress.toLowerCase()));
  return wanted.filter((r) => {
    const key = r.emailAddress.toLowerCase();
    if (have.has(key)) return false;
    have.add(key);
    return true;
  });
}

export function formatRecipients(recipients) {
  return recipients.map((r) => (r.displayName ? `${r.displayName} <${r.emailAddress}>` : r.emailAddress)).join('; ');
}

// Office.js の addAsync に渡す形
export function toOfficeRecipients(recipients) {
  return recipients.map((r) => (r.displayName ? { displayName: r.displayName, emailAddress: r.emailAddress } : r.emailAddress));
}
