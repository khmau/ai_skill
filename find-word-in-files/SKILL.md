---
name: find-word-in-files
description: Finds a given word or phrase (for example "Intel") in PDF, Word (DOCX/DOC), PowerPoint (PPTX/PPT) or HTML files, or standalone image files (PNG, JPG, TIFF, SVG, WebP, ...) - in the diagrams, images, logos, screenshots and vector art (text layer plus OCR), in the body text (paragraphs, headings, tables, captions, headers, footers, speaker notes), and in hyperlink targets (URLs, mailto addresses, linked file paths). Reports found/not found per category with page or slide numbers and the matched text, and can write an HTML report with marked screenshots. Use when the user asks whether a document, its figures, pictures, logos, text or links mention a term, or when auditing documents or images for a brand name or trademark (e.g. Intel-to-Altera rebranding).
---

# Find a word in files: graphics, text and hyperlinks

Reports whether a word or phrase appears, and where, in three separately reported places:

- **Graphics**: diagrams, pictures, logos, screenshots and outlined vector art, read from the text layer and with OCR.
- **Body text**: the ordinary text of the document, i.e. paragraphs, headings, lists, tables, captions, page headers and footers, text boxes and speaker notes.
- **Hyperlink targets**: the web address, e-mail address or file path behind a link. The visible link text is body text, and is also shown as context for each link match.

## Inputs

| Kind | Extensions |
|------|------------|
| Documents | `.pdf`, `.docx` `.docm` `.doc` `.rtf`, `.pptx` `.pptm` `.ppt`, `.html` `.htm` |
| Vector images | `.svg`, `.svgz` |
| Raster images | `.png` `.jpg` `.jpeg` `.gif` `.bmp` `.tif` `.tiff` `.webp` `.avif` `.jp2` `.ico` `.psd` `.tga` `.ppm`, and `.emf` `.wmf` (Windows); other extensions are detected from the file content |

Multi-page TIFFs and animated GIFs are scanned frame by frame and reported as pages. EXIF rotation is applied before OCR.

## Setup (once)

```bash
pip install -r requirements.txt
```

- Needs Python 3.10+. On Windows `python` may be Python 2; use `py -3`.
- DOCX/DOC/PPTX/PPT are converted to PDF with Microsoft Word/PowerPoint (Windows) or LibreOffice (`soffice`) so page and slide numbers are exact. Without either, DOCX/PPTX are read directly and DOCX page numbers become estimates (`page ~N`). DOC/PPT require a converter.

## Run

Paths are relative to this skill's folder.

```bash
python scripts/find_word_in_graphics.py <file> "<word or phrase>" [more terms ...] --html <report.html> --crops <tmp-dir>
```

| Option | Purpose |
|--------|---------|
| `--html [FILE]` | Write a self-contained HTML report with marked screenshots (default `<input>_report.html` next to the input) |
| `--crops DIR` | Save a PNG of every graphic with a match, each match boxed in its term's colour |
| `--json`, `-o FILE` | JSON output; write the text/JSON report to a UTF-8 file (use `-o` instead of `>` in Windows PowerShell 5) |
| `--pages 1-5,9` | Limit pages/slides (or frames of a multi-page image) |
| `--case-sensitive`, `--substring` | Default is case-insensitive whole word: "Intel" matches "Intel®", "IntelFPGA", "intel_fpga", not "Intelligent" |
| `--no-text` | Skip body text; check graphics (and hyperlinks) only |
| `--no-links` | Skip hyperlink targets |
| `--no-ocr` | Text layer only: fast, but misses words in pictures and logos |
| `--ocr-all` | Also OCR vector diagrams whose labels are already selectable text (slower; for suspected outlined text) |
| `--native` | Read DOCX/PPTX directly instead of converting |
| `--min-score 0.5` | OCR confidence threshold |

Exit code: `0` found, `1` not found, `2` error. Expect roughly 20-30 s for a 60-page PDF with `--html`; Office conversion adds about 10 s.

## Workflow

1. Run the script with `--html` (next to the input, or where the user asks) and `--crops` pointing at a temp directory.
2. For OCR matches with confidence below 0.8, or anything surprising, open the crop PNG and confirm visually before reporting it. The crops are the same marked screenshots the HTML report embeds.
3. Answer in this format, then link the HTML report (`HTML report:` line, or `report` in JSON):

```markdown
**"Intel": FOUND** in graphics on 62 of 62 pages (ug-683290-746921.pdf)

| Page | Matched text | Where | Found by |
|------|--------------|-------|----------|
| 1 | intel | cover logo | OCR of vector art |
| 3 | Intel FPGA IP; Mailbox Client Intel FPGA IP | Figure 1. Mailbox Client Intel FPGA IP System Block Diagram | diagram text |
| 6 | Mailbox Client Intel FPGA IP | Figure 2. Mailbox Client Intel FPGA IP Interfaces | diagram text |
| 2-62 | intel. | page header logo (repeated) | OCR of vector art |

**In body text**: 364 matches on 62 pages (list the distinctive ones; summarise running headers and footers)

| Page | Matched text | Where |
|------|--------------|-------|
| 1 | Updated for Intel® Quartus® Prime Design Suite: 22.3 | |
| 2, 4-57, 59-62 | Mailbox Client Intel® FPGA IP User Guide | page footer (repeated) |

**In hyperlink targets**

| Page | Link target | Link text |
|------|-------------|-----------|
| 1-62 | mailto:FPGAtechdocfeedback@intel.com?subject=... | Send Feedback (repeated) |
| 4 | https://www.intel.com/content/www/us/en/docs/programmable/683091/... | Avalon Interface Specifications |
```

When the user asked only about graphics, or the body-text list is long, lead with the graphics and hyperlink findings and give the body-text count with a pointer to the HTML report. If nothing matched: `**"Altera": NOT FOUND** in any diagram, image, text or hyperlink (62 pages checked)`. Always pass on `note:` lines (images that could not be loaded, OCR disabled, conversion fallback).

## Reading the report

Each match line shows page/slide, how it was found, the full label or line containing the term, then `[figure caption | page header/footer | image name]`. Body-text matches follow under `in body text:`, with the line or paragraph and where it is (`page header`, `page footer`, `figure caption`, `table`, `text box`, `slide title`, `speaker notes`, or `section "heading"` for HTML). Hyperlink matches follow under `in hyperlink targets:`, each with the decoded link target and `[link text "..."]` (or `linked picture or shape` for a linked logo or text box).

| Label | Meaning |
|-------|---------|
| `diagram text` | Selectable text inside a vector diagram (block diagram, flowchart, SmartArt, chart, grouped shapes) |
| `vector art (OCR)` | Word drawn as outlines, e.g. a logo; invisible to Ctrl+F |
| `image (OCR)` | Word inside a raster picture, screenshot or scanned page |
| `text on image` | Selectable text placed over a picture |
| `body text` | Ordinary document text outside graphics |
| `hyperlink` | The term is in the link target (URL, mailto, file path), not necessarily in the visible text |
| `repeated on N pages` | Same graphic or text line at the same spot on 3 or more pages, usually a header/footer logo or running title; for links, the same target on 3 or more pages, such as a feedback link in the footer |
| `xN` | The line contains the term N times |
| `-` (no page) | Single image file, HTML (`<img> #n src`, `inline <svg> #n`), or DOCX header/PPTX master images in native mode |

The HTML report needs no other files. It shows a FOUND / NOT FOUND card per term, then one card per graphic, with:

- the screenshot, each match boxed in its term's colour (click to enlarge);
- a thumbnail of the page with the graphic outlined in amber;
- a table of matched text with the term highlighted, how it was found, OCR confidence (below 0.8 in amber) and count.

A graphic repeated on many pages, such as a header logo, is one card listing its pages. With several terms, buttons filter the cards by term. Text read from file structure in `--native` mode has no picture, and its card says so.

Body-text matches are in an **In text** section after the graphics: one card per page with the whole page and every match boxed, and a table of the matched lines. Text repeated on many pages (running headers and footers) is one card with a strip of its first occurrence, and is not boxed again on each page. HTML and `--native` text cards have no screenshot. Page screenshots make the report larger, about 120 KB per page with a text match; use `--no-text` when only graphics matter.

Hyperlink matches are in a separate **In hyperlinks** section, one card per link target: a strip of the page with the link area boxed, and a table of the target with the term highlighted, the link text and the count. A link used on many pages is one card listing its pages. HTML, SVG and `--native` link cards have no screenshot.

## How body text is found

- PDF and converted Office files: every visible line of the text layer that is not already reported as diagram text, including captions, tables, code blocks, headers and footers. Invisible OCR layers of scanned pages are skipped, since the scan itself is OCR'd as a graphic. A phrase wrapped onto the next line is still found.
- DOCX with `--native`: paragraphs of the body, tables, text boxes, headers, footers, footnotes and endnotes, with estimated pages. A word split over formatting runs ("**Int**el") is joined first.
- PPTX: slide text boxes, titles and tables (natively), and speaker notes in both modes; notes are not part of the PDF export, so they are read from the file and have no screenshot.
- HTML: visible text by block (paragraph, heading, list item, table cell, page title). `<script>`, `<style>`, `<noscript>` and inline SVG are skipped; inline SVG text is reported as diagram text.
- Matching is the same as for graphics: whole words, case-insensitive by default.

## How hyperlinks are found

- PDF and converted Office files: web, mailto and file links on each page. Jumps within the document (table of contents, cross-references) have no target text and are skipped.
- DOCX/PPTX with `--native`: hyperlinks on text, pictures and shapes, and `HYPERLINK` fields, including headers, footers, footnotes and slide masters.
- HTML: `<a href>` and image-map `<area href>`, including in-page anchors (`#intel-history`). SVG: `<a href>` / `xlink:href`.
- The target is URL-decoded (`%20` becomes a space) and punctuation counts as a word break, so "Intel" matches `www.intel.com`, `intel-fpga`, `/IntelFPGA/` and `?subject=Intel%20FPGA`. A phrase may run together or be split by punctuation: "Intel FPGA" matches `intelfpga` and `intel-fpga`. A word joined into a longer lowercase name, such as "intel" in `intelrealsense.com`, needs `--substring`.

## How graphics are found (PDF and converted Office files)

- Raster images are always OCR'd. A full-page image (a scan) is OCR'd entirely, so every occurrence on a scanned page is reported.
- Vector drawings are clustered into regions. A region is dropped when it is a table (ruled grid, or a "Table N" caption above it), a plain frame (border, shaded code block), or a page-wide background.
- Short labels up to 20 pt beside a drawing count as part of it.
- Selectable text inside a region is matched directly. OCR runs on regions with glyph-like paths (logos, outlined text). Use `--ocr-all` to OCR every region.
- OCR tolerates accent artefacts and look-alike characters (l/I/1, O/0) and missing spaces ("IntelFPGA").

## Limitations

- Region detection is heuristic. A text box with an icon may be counted as a diagram, and a table without rules or caption is not recognised as a table. Check the `where` column and the crops.
- OCR (RapidOCR, PP-OCR models, Latin and Chinese) reads horizontal text. Vertical or rotated text and text smaller than about 6 px inside pictures may be missed.
- HTML: `<img>`, `<object>`, `<embed>` and inline `<svg>` are checked; remote images are downloaded. CSS background images and `<canvas>` are not.
- Image files are OCR'd only, so `--no-ocr` checks nothing except the `<text>` and links of an SVG. EMF/WMF need Windows to render.
- Body text: text in PowerPoint slide layouts and masters is not read natively (the converted PDF includes what appears on the slides). HTML `alt` and `title` attributes are not searched (an image's `alt` inside a link is shown as link context), and text hidden with CSS is searched like visible text.
- Hyperlinks: only the target is searched; the link text is body text. A URL typed as plain text without a link is body text. In `--native` DOCX mode, page numbers are estimates.
- JSON output adds `text_checked` and `links_checked` and, per term in `summary`, `in_text` and `in_links`; body-text matches have `source` `body`, link matches `source` `link`.
