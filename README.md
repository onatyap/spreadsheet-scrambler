# Spreadsheet Scrambler

Make a copy of an Excel workbook with made-up values and the same structure, so you can share it, test with it or hand it to other tools without exposing the real data.

It is one HTML file. Open it in a browser, choose a workbook, pick what to scramble and download the copy.

**Use it here: https://onatyap.github.io/spreadsheet-scrambler/**

## What it's for

- Sending a realistic sample to a vendor, consultant or support team without sending real customer or financial data.
- Giving a spreadsheet to an AI assistant or another online tool so it can build formulas, reports or code, without it ever seeing the real values.
- Test data for developers and data teams that has the exact shape of production files.
- Demos, training and bug reports with files that look real but aren't.

## Private by design

- **Nothing leaves your computer.** The page reads the file into memory, rewrites it there and hands you the new file. There is no upload and no server.
- **No AI, LLM or agent of any kind**, and no call to any provider, API or service.
- **No analytics, cookies, fonts, CDNs or other outside code.** There are no dependencies: the page unzips and zips the workbook with the browser's own compression.
- **The page can't send anything even by mistake.** Its Content-Security-Policy is `default-src 'none'; connect-src 'none'; form-action 'none'`, so the browser blocks every network request it might make.
- **It works offline.** Save `index.html` and open it from your disk.
- **Search and preview files are inert.** The search and link-preview tags in the page, `sitemap.xml`, `llms.txt` and `og-image.png` are only read by search engines and chat apps. The page itself never loads them.

To check for yourself, read `index.html` (about 1,000 lines with the page's HTML and CSS, no minified code) or open your browser's developer tools on the Network tab while you scramble a file. You'll see no requests. The test in `test/test.mjs` also fails if the page ever gains a way to reach the network: an outside script, style, font or image, a weaker security policy, or code that calls `fetch`, `XMLHttpRequest`, `WebSocket`, `sendBeacon` and the like.

## What you can choose

- **Sheets:** any or all of them, hidden sheets included.
- **What to scramble:** text, numbers, dates and times, TRUE/FALSE, in any combination.
- **Which cells:** every cell, or ranges such as `B2:F500, H:H, 3:10`. A sheet name in front limits a range to that sheet: `Sales!A2:C40`.
- **Headings:** leave the first N rows of each sheet as they are.
- **Personal details:** remove author names, comment authors and file properties.

## How values change

- **Text:** every letter becomes a random letter and every digit a random digit. Spaces, punctuation and length stay, and Arabic, Cyrillic, Greek, Hebrew, Chinese, Japanese, Korean, Thai and Devanagari text stays in its own script. The same text always becomes the same made-up text within a file, so lookups, filters and dropdown lists keep working.
- **Numbers:** same sign and the same number of digits before and after the decimal point, with new digits. Zero stays zero.
- **Dates and times:** moved by up to a year, earlier or later. A date is any number in a date or time format.
- **TRUE/FALSE:** picked at random for each cell.
- **Copies of values** kept inside the file are changed to match: comments on scrambled cells, chart caches, pivot table caches (rebuilt when Excel opens the file), filter settings, dropdown lists and values copied from linked workbooks.

Random choices come from the browser's secure random number generator (`crypto.getRandomValues`) and are never stored, so the scramble can't be reversed or repeated. Keep your original.

## What stays the same

- Sheet names, layout, formatting, column widths, merged cells, table headings and every part of the file the page doesn't need to change, byte for byte.
- Formulas. Their saved results are cleared, because they could give the real values away, and the file is marked so Excel works them out again on open. Tools that read the file without calculating it will see empty formula cells.
- Text and numbers typed inside formulas, unless the text matches a value that was scrambled.
- Pictures, embedded files, macros, data connections, change history, slicers, and header and footer text. The page warns you when the file has any of these.

## Check after scrambling

Before it offers the download, the page reads the new file back and looks for every original text it scrambled. If any is still there, it stops and saves nothing.

## Use it

- **Online:** https://onatyap.github.io/spreadsheet-scrambler/
- **Offline:** download `index.html` and open it in Chrome, Edge, Firefox or Safari (2023 or later).
- **Your own copy:** fork the repository, then in Settings → Pages choose "Deploy from a branch", `main`, `/ (root)`.

## Limits

- Works with `.xlsx` and `.xlsm`. Not with `.xls`, `.csv`, `.ods`, password-protected workbooks or files over 4 GB.
- Large workbooks are processed in memory, so a file of a few hundred MB needs a few GB of free memory.

## Develop

Everything is in `index.html`. The test runs the same code in Node and needs Node 22 or later:

```bash
node test/test.mjs
```

`test/fixture.xlsx` is a small invented workbook. To rebuild it you need Python with `openpyxl`:

```bash
python3 test/make-fixture.py
```

## Licence

MIT. See `LICENSE`.
