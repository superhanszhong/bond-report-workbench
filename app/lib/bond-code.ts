// 上清所政金债代码可能被 Excel 当作数字，丢失开头的 0。
// 仅补齐已知的 09 + 年份 + 政金行别结构，不对其他代码补位或改写 X/Z。
export function normalizeBondCode(value: unknown) {
  const code = String(value ?? "").trim().toUpperCase().replace(/\.(?:IB|SH|SZ)$/, "");
  return /^9\d{2}0[234]\d{2,3}(?:[XZ]\d*)?$/.test(code) ? `0${code}` : code;
}

// Identity is separate from the original display code. Never collapse different dates or issue numbers.
export function bondIssueKey(value: unknown) {
  return normalizeBondCode(value).replace(/[XZ](\d*)$/, (_, sequence) => `X${Number(sequence || 1)}`);
}

export function baseBondCode(value: unknown) {
  return normalizeBondCode(value).replace(/[XZ]\d*$/, "");
}
