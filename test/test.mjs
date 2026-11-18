// Runs the page's own code in Node on the test workbooks.   node test/test.mjs   (Node 22 or later)
// fixture.xlsx keeps its text inline (as openpyxl writes it); fixture-sst.xlsx is the same workbook saved by
// LibreOffice, with a shared strings list and chart caches, as Excel writes them.
import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";
import vm from "node:vm";

const html = readFileSync(new URL("../index.html", import.meta.url), "utf8");
vm.runInThisContext(html.match(/<script>([\s\S]*?)<\/script>/)[1]); // there's no document in Node, so the page itself doesn't start
const { scrambleWorkbook, parseSelection, readZip, Package, UserError } = globalThis.SpreadsheetScrambler;

const SHEETS = ["Sales", "Lookup", "Notes"];
const PRIVATE = ["Acme Trading LLC", "Globex Holdings", "Initech Partners", "Bilal", "Haddad", "شركة الاتصالات", "North", "Overdue",
  "Disputed", "Jane Example", "intranet.example.org", "Very secret note", "Board approved", "Quarterly customers", "Inline secret value"];
const options = (o = {}) => ({ text: true, numbers: true, dates: true, bool: true, personal: true, headerRows: 0,
  selection: parseSelection(o.ranges || "", SHEETS, o.sheets || SHEETS), ...o });

async function parts(buffer) { // every XML part of a workbook as text, with XML escapes undone
  const pkg = new Package(await readZip(buffer)), out = new Map();
  for (const e of pkg.entries) if (/\.(xml|rels|vml)$/.test(e.name)) out.set(e.name, (await pkg.get(e.name)).replace(/&amp;/g, "&").replace(/&quot;/g, '"'));
  return out;
}
const all = p => [...p.values()].join("\n");
const count = (s, re) => (s.match(re) || []).length;
const texts = s => [...s.matchAll(/<t[^>]*>([^<]*)<\/t>/g)].map(m => m[1]);
function cellText(p, sheet, ref) {
  const m = p.get(sheet).match(new RegExp(`<c r="${ref}"([^>]*)>([\\s\\S]*?)</c>`));
  if (!/t="s"/.test(m[1])) return texts(m[2]).join("");
  const si = [...p.get("xl/sharedStrings.xml").matchAll(/<si>([\s\S]*?)<\/si>/g)][Number(m[2].match(/<v>(\d+)<\/v>/)[1])];
  return texts(si[1]).join("");
}
async function withInlineCell(file) { // adds an inline string cell, which Excel itself rarely writes
  const pkg = new Package(await readZip(file));
  const sheet = await pkg.get("xl/worksheets/sheet1.xml");
  pkg.set("xl/worksheets/sheet1.xml", sheet.replace(/(<row r="3"[^>]*>)/, '$1<c r="J3" t="inlineStr"><is><t>Inline secret value</t></is></c>'));
  return (await pkg.build()).arrayBuffer();
}

let failures = 0;
async function test(name, fn) {
  try { await fn(); console.log("ok  ", name); } catch (err) { failures++; console.log("FAIL", name, "\n    ", err.message); }
}

for (const fixtureName of ["fixture.xlsx", "fixture-sst.xlsx"]) {
  const url = new URL(`./${fixtureName}`, import.meta.url);
  if (!existsSync(url)) { console.log("skip", fixtureName, "(not built)"); continue; }
  const raw = readFileSync(url);
  const workbook = await withInlineCell(raw.buffer.slice(raw.byteOffset, raw.byteOffset + raw.byteLength));
  const before = await parts(workbook);
  const sst = before.has("xl/sharedStrings.xml");
  const t = name => `${fixtureName}: ${name}`;

  await test(t("everything: no private value is left anywhere in the file"), async () => {
    const r = await scrambleWorkbook(workbook, options());
    const text = all(await parts(await r.blob.arrayBuffer()));
    for (const s of PRIVATE) assert.ok(!text.includes(s), `still in the file: ${s}`);
    for (const n of ["1234567", "98765", "4711", "0.153"]) assert.ok(!text.includes(`>${n}<`), `number still in the file: ${n}`);
    assert.equal(r.check.left, 0);
    assert.ok(r.check.searched >= 8, `searched only ${r.check.searched} texts`);
  });

  await test(t("everything: structure is kept"), async () => {
    const r = await scrambleWorkbook(workbook, options());
    const after = await parts(await r.blob.arrayBuffer());
    assert.deepEqual([...after.keys()].sort(), [...before.keys()].filter(n => n !== "docProps/custom.xml").sort());
    for (const name of [...before.keys()].filter(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))) {
      assert.equal(count(after.get(name), /<c r=/g), count(before.get(name), /<c r=/g), `cells in ${name}`);
      assert.equal(count(after.get(name), /<f[ >]/g), count(before.get(name), /<f[ >]/g), `formulas in ${name}`);
      assert.equal(count(after.get(name), /<\/f><v>/g), 0, `saved formula results in ${name}`);
    }
    const wb = after.get("xl/workbook.xml");
    for (const s of SHEETS) assert.ok(wb.includes(`name="${s}"`), `sheet name ${s}`);
    assert.ok(/fullCalcOnLoad="1"/.test(wb));
    assert.ok(/name="Name"/.test(after.get("xl/tables/table1.xml")), "table heading kept");
    assert.equal(cellText(after, "xl/worksheets/sheet2.xml", "A1"), "Name", "table heading cell kept");
    const sales = r.stats.get("Sales");
    assert.ok(sales.text >= 10 && sales.number >= 3 && sales.date === 4 && sales.zero === 1, JSON.stringify(sales));
    if (sst) assert.ok(after.has("xl/sharedStrings.xml") && sales.text >= 10, "shared strings used");
  });

  await test(t("dropdown, filter, formulas and named value follow the scrambled values"), async () => {
    const r = await scrambleWorkbook(workbook, options());
    const after = await parts(await r.blob.arrayBuffer());
    const sheet = after.get("xl/worksheets/sheet1.xml"), wb = after.get("xl/workbook.xml");
    const strings = texts(sheet + (after.get("xl/sharedStrings.xml") || ""));
    const list = sheet.match(/<formula1>"([^"]*)"<\/formula1>/)[1].split(",");
    for (const item of list.slice(0, 3)) assert.ok(strings.includes(item), `dropdown item ${item} is a scrambled cell value`);
    const region = cellText(after, "xl/worksheets/sheet1.xml", "B2");
    assert.equal(sheet.match(/<filter val="([^"]*)"/)[1], region, "filter value matches the scrambled region");
    assert.equal(sheet.match(/COUNTIF\(B2:B5,"([^"]*)"\)/)[1], region, "formula text matches the scrambled region");
    assert.equal(wb.match(/<definedName[^>]*name="Flagship"[^>]*>"([^"]*)"<\/definedName>/)[1],
      cellText(after, "xl/worksheets/sheet1.xml", "A2"), "named value matches the scrambled customer");
    assert.equal(cellText(after, "xl/worksheets/sheet2.xml", "A2"), cellText(after, "xl/worksheets/sheet1.xml", "A2"), "same text, same fake");
  });

  await test(t("only text, only A2:A3 and I2 on Sales, headings kept"), async () => {
    const r = await scrambleWorkbook(workbook, options({ numbers: false, dates: false, bool: false, personal: false, headerRows: 1,
      sheets: ["Sales"], selection: parseSelection("A2:A3, I2", SHEETS, ["Sales"]) }));
    const after = await parts(await r.blob.arrayBuffer()), text = all(after);
    assert.notEqual(cellText(after, "xl/worksheets/sheet1.xml", "A3"), "Globex Holdings", "A3 scrambled");
    assert.equal(cellText(after, "xl/worksheets/sheet1.xml", "A5"), "Acme Trading LLC", "A5 kept");
    assert.equal(cellText(after, "xl/worksheets/sheet2.xml", "A3"), "Globex Holdings", "other sheets kept");
    assert.ok(!text.includes("Haddad"), "I2 scrambled");
    assert.equal(cellText(after, "xl/worksheets/sheet1.xml", "A1"), "Customer", "heading kept");
    assert.ok(text.includes(">1234567<"), "numbers kept");
    assert.ok(text.includes("Jane Example"), "personal details kept when not asked");
    assert.equal(r.check.left, 0);
  });

  await test(t("only numbers on the hidden sheet"), async () => {
    const r = await scrambleWorkbook(workbook, options({ text: false, sheets: ["Notes"], selection: parseSelection("", SHEETS, ["Notes"]) }));
    const after = await parts(await r.blob.arrayBuffer());
    assert.ok(!/<v>42<\/v>/.test(after.get("xl/worksheets/sheet3.xml")), "number on Notes scrambled");
    assert.ok(all(after).includes("Board approved"), "text kept");
    assert.ok(after.get("xl/worksheets/sheet1.xml").includes("<v>1234567</v>"), "other sheets untouched");
  });
}

await test("ranges are read and checked", () => {
  const sel = parseSelection("B2:F9, H:H; 'Sales'!3:10", SHEETS, ["Sales", "Lookup"]);
  assert.deepEqual(sel.get("Sales"), [{ c1: 2, c2: 6, r1: 2, r2: 9 }, { c1: 8, c2: 8, r1: 1, r2: Infinity }, { c1: 1, c2: Infinity, r1: 3, r2: 10 }]);
  assert.equal(sel.get("Lookup").length, 2);
  assert.equal(parseSelection("", SHEETS, ["Sales"]).get("Sales"), null);
  assert.throws(() => parseSelection("B2:", SHEETS, SHEETS), UserError);
  assert.throws(() => parseSelection("Nope!A1", SHEETS, SHEETS), UserError);
  assert.throws(() => parseSelection("Notes!A1", SHEETS, ["Sales"]), UserError);
});

await test("a file that isn't a workbook is refused politely", async () => {
  await assert.rejects(scrambleWorkbook(new TextEncoder().encode("hello").buffer, options()), UserError);
});

console.log(failures ? `\n${failures} failed` : "\nall passed");
process.exit(failures ? 1 : 0);
