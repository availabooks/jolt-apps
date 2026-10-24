// @ts-check

/**
 * A small Markdown → HTML converter for the AI's explanations. Installed apps can't import code from other
 * sites, so this covers what AI replies use: headings, paragraphs, lists (nested), block quotes, fenced code,
 * tables, rules, and inline bold, italic, strikethrough, code and links.
 *
 * Safe for innerHTML: all text is escaped, and only http(s) and mailto links are kept.
 */

const LIST = /^(\s*)([-*+]|\d+[.)])\s+(.*)$/;
const FENCE = /^\s*(```+|~~~+)/;
const HEADING = /^\s{0,3}(#{1,6})\s+(.*?)\s*#*\s*$/;
const RULE = /^\s{0,3}([-*_])(\s*\1){2,}\s*$/;
const QUOTE = /^\s{0,3}>/;
const TABLE_RULE = /^\s*\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$/;

/** @param {string} s */
const escape = (s) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");

/** @param {string} md @returns {string} */
export function markdownToHtml(md) {
  const lines = md.replace(/\r\n?/g, "\n").split("\n");
  const out = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (!line.trim()) {
      i++;
      continue;
    }
    const fence = line.match(FENCE);
    if (fence) {
      const body = [];
      for (i++; i < lines.length && !lines[i].trim().startsWith(fence[1]); i++) body.push(lines[i]);
      i++;
      out.push(`<pre><code>${escape(body.join("\n"))}</code></pre>`);
      continue;
    }
    const h = line.match(HEADING);
    if (h) {
      out.push(`<h${h[1].length}>${inline(h[2])}</h${h[1].length}>`);
      i++;
      continue;
    }
    if (RULE.test(line)) {
      out.push("<hr>");
      i++;
      continue;
    }
    if (QUOTE.test(line)) {
      const body = [];
      while (i < lines.length && QUOTE.test(lines[i])) body.push(lines[i++].replace(/^\s{0,3}>\s?/, ""));
      out.push(`<blockquote>${markdownToHtml(body.join("\n"))}</blockquote>`);
      continue;
    }
    if (isTable(lines, i)) {
      const [html, next] = table(lines, i);
      out.push(html);
      i = next;
      continue;
    }
    if (LIST.test(line)) {
      const [html, next] = list(lines, i);
      out.push(html);
      i = next;
      continue;
    }
    const para = [];
    while (i < lines.length && lines[i].trim() && (para.length === 0 || !startsBlock(lines, i))) para.push(lines[i++].trim());
    out.push(`<p>${inline(para.join(" "))}</p>`);
  }
  return out.join("\n");
}

/** @param {string[]} lines @param {number} i */
function startsBlock(lines, i) {
  const l = lines[i];
  return FENCE.test(l) || HEADING.test(l) || RULE.test(l) || QUOTE.test(l) || LIST.test(l) || isTable(lines, i);
}

/** @param {string[]} lines @param {number} i */
const isTable = (lines, i) => lines[i].includes("|") && i + 1 < lines.length && TABLE_RULE.test(lines[i + 1]) && lines[i + 1].includes("-");

/** @param {string} row */
function cells(row) {
  return row
    .trim()
    .replace(/^\|/, "")
    .replace(/(?<!\\)\|$/, "")
    .split(/(?<!\\)\|/)
    .map((c) => c.trim().replace(/\\\|/g, "|"));
}

/** @param {string[]} lines @param {number} i @returns {[string, number]} */
function table(lines, i) {
  const head = cells(lines[i]);
  const align = cells(lines[i + 1]).map((c) => (c.startsWith(":") && c.endsWith(":") ? "center" : c.endsWith(":") ? "right" : ""));
  /** @param {string} tag @param {string[]} row */
  const tr = (tag, row) =>
    `<tr>${head.map((_, c) => `<${tag}${align[c] ? ` style="text-align:${align[c]}"` : ""}>${inline(row[c] ?? "")}</${tag}>`).join("")}</tr>`;
  const body = [];
  for (i += 2; i < lines.length && lines[i].trim() && lines[i].includes("|"); i++) body.push(tr("td", cells(lines[i])));
  return [`<table><thead>${tr("th", head)}</thead><tbody>${body.join("")}</tbody></table>`, i];
}

/**
 * A list starting at line i, with deeper-indented items as nested lists.
 * @param {string[]} lines @param {number} i @returns {[string, number]}
 */
function list(lines, i) {
  const first = /** @type {RegExpMatchArray} */ (lines[i].match(LIST));
  const indent = first[1].length;
  const ordered = /\d/.test(first[2]);
  /** @type {{ text: string[], children: string[] }[]} */
  const items = [];
  while (i < lines.length) {
    const line = lines[i];
    const m = line.match(LIST);
    if (m && m[1].length === indent && /\d/.test(m[2]) === ordered) {
      items.push({ text: [m[3]], children: [] });
      i++;
    } else if (m && m[1].length > indent) {
      const [html, next] = list(lines, i);
      items[items.length - 1].children.push(html);
      i = next;
    } else if (!line.trim()) {
      // A blank line ends the list unless another item of this list or a nested one follows.
      let j = i;
      while (j < lines.length && !lines[j].trim()) j++;
      const n = j < lines.length ? lines[j].match(LIST) : null;
      if (!n || n[1].length < indent) break;
      i = j;
    } else if (!m && !startsBlock(lines, i)) {
      items[items.length - 1].text.push(line.trim());
      i++;
    } else break;
  }
  const tag = ordered ? "ol" : "ul";
  const start = ordered ? parseInt(first[2], 10) : 1;
  const open = start !== 1 ? `<ol start="${start}">` : `<${tag}>`;
  return [`${open}${items.map((it) => `<li>${inline(it.text.join(" "))}${it.children.join("")}</li>`).join("")}</${tag}>`, i];
}

/**
 * Inline formatting. Code spans and link addresses are set aside first, so nothing inside them is
 * formatted (an _ in a URL isn't italics).
 * @param {string} text
 */
export function inline(text) {
  /** @type {string[]} */
  const kept = [];
  /** @param {string} html */
  const keep = (html) => `${kept.push(html) - 1}`;
  let s = text.replace(/`([^`]+)`/g, (_, c) => keep(`<code>${escape(c)}</code>`));
  s = s.replace(/\[([^\]]+)\]\(([^)\s]+)\)/g, (_, label, url) =>
    /^(https?:|mailto:)/i.test(url) ? `${keep(`<a href="${escape(url)}" target="_blank" rel="noopener noreferrer">`)}${label}${keep("</a>")}` : label,
  );
  s = escape(s);
  s = s.replace(/(\*\*|__)(?=\S)([\s\S]*?\S)\1/g, "<strong>$2</strong>");
  s = s.replace(/(^|[^\w*])\*(?=\S)([^*]*?\S)\*(?!\*)/g, "$1<em>$2</em>");
  s = s.replace(/(^|\W)_(?=\S)([^_]*?\S)_(?!\w)/g, "$1<em>$2</em>");
  s = s.replace(/~~(?=\S)([\s\S]*?\S)~~/g, "<del>$1</del>");
  return s.replace(/(\d+)/g, (_, n) => kept[Number(n)]);
}
