import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";
import { plainComment } from "../utils/developerClaude.js";

const readWeb = (name) => readFileSync(new URL(`../../web/${name}`, import.meta.url), "utf8");
const between = (source, start, end) => {
  const a = source.indexOf(start);
  const b = source.indexOf(end, a);
  assert.ok(a >= 0 && b > a, `${start} is present`);
  return source.slice(a, b);
};

// The Developer tab's note formatter, run the way the browser runs it.
const context = vm.createContext({});
vm.runInContext(between(readWeb("app/admin.js"), "function escapeHtml(", "function applyAdminArtisanPreset("), context);
vm.runInContext(between(readWeb("app/developer.js"), "function devInline(", "function renderDeveloperDetail("), context);
const md = (text) => context.devMarkdown(text);

test("Claude's note shows bold, headings, lists and ticked boxes", () => {
  const html = md("### Summary\nAll **done** now.\n\n- [x] Read the code\n- [ ] Push\n- plain point\n\n1. first\n2. second");
  assert.match(html, /<h5>Summary<\/h5>/);
  assert.match(html, /<p>All <strong>done<\/strong> now\.<\/p>/);
  assert.match(html, /<li class="dev-md-task"><span class="dev-md-box done" role="img" aria-label="Done">✓<\/span><span>Read the code<\/span><\/li>/);
  assert.match(html, /<span class="dev-md-box" role="img" aria-label="Not done"><\/span><span>Push<\/span>/);
  assert.match(html, /<li>plain point<\/li><\/ul>/);
  assert.match(html, /<ol><li>first<\/li><li>second<\/li><\/ol>/);
  assert.doesNotMatch(html, /\*\*|###|\[x\]/);
});

test("links in a note are clickable and open in a new tab", () => {
  const html = md("[View job run](https://github.com/a/b/actions/runs/1?x=1&y=2) and https://example.com/page.");
  assert.match(html, /<a href="https:\/\/github\.com\/a\/b\/actions\/runs\/1\?x=1&amp;y=2" target="_blank" rel="noopener noreferrer">View job run<\/a>/);
  assert.match(html, /<a href="https:\/\/example\.com\/page" [^>]+>https:\/\/example\.com\/page<\/a>\./);
});

test("HTML and unsafe links in a note are shown as text, never run", () => {
  const html = md('<img src=x onerror="alert(1)">\n**<script>alert(2)</script>**\n[click](javascript:alert(3))\n`<b>code</b>`\n```\n<iframe src="x"></iframe>\n```');
  assert.doesNotMatch(html, /<img|<script|<iframe|<b>|href="javascript/i);
  assert.match(html, /&lt;img src=x onerror=&quot;alert\(1\)&quot;&gt;/);
  assert.match(html, /<strong>&lt;script&gt;alert\(2\)&lt;\/script&gt;<\/strong>/);
  assert.match(html, /\[click\]\(javascript:alert\(3\)\)/);
  assert.match(html, /<code>&lt;b&gt;code&lt;\/b&gt;<\/code>/);
  assert.match(html, /<pre><code>&lt;iframe src=&quot;x&quot;&gt;&lt;\/iframe&gt;<\/code><\/pre>/);
});

test("a link cannot break out of its href", () => {
  const html = md('[x](https://a.com/"onmouseover="alert(1)) https://b.com/"><script>');
  assert.doesNotMatch(html, /<script|"onmouseover/);
  for (const href of html.match(/href="[^"]*"/g) || []) assert.doesNotMatch(href, /[<>]/);
});

test("the server keeps markdown links in Claude's note but drops HTML", () => {
  const note = plainComment('**Claude finished** <img src="spinner.gif">\n\n[View job run](https://github.com/a/b/actions/runs/1)');
  assert.equal(note, "**Claude finished** \n\n[View job run](https://github.com/a/b/actions/runs/1)");
});
