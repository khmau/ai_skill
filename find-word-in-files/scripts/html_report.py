"""Self-contained HTML report for find_word_in_graphics.py (screenshots are embedded as data URIs)."""
from __future__ import annotations

import base64
import datetime
import html
import json
import os

LEGEND = [
    ("diagram text", "Selectable text inside a vector diagram."),
    ("vector art (OCR)", "Word drawn as shapes, such as a logo or outlined label. Plain text search cannot see it."),
    ("text on image", "Selectable text placed on top of a picture."),
    ("image (OCR)", "Word inside a raster picture, screenshot or scan."),
]
TEXT_LEGEND = ("body text", "Ordinary text of the document: paragraphs, headings, lists, tables, captions, "
                            "headers, footers, text boxes and speaker notes.")
LINK_LEGEND = ("hyperlink", "Target of a link (web address, e-mail address or file path). Readers see only the "
                            "link text, so the target is easy to miss.")

CSS = """
:root { --bg:#f5f6f8; --card:#fff; --text:#1f2937; --muted:#6b7280; --line:#e5e7eb; --soft:#f9fafb;
        --ok:#15803d; --ok-bg:#dcfce7; --no:#b91c1c; --no-bg:#fee2e2; --warn:#b45309; --warn-bg:#fef3c7; }
* { box-sizing:border-box; }
body { margin:0; background:var(--bg); color:var(--text);
       font:14px/1.5 "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif; }
.wrap { max-width:1280px; margin:0 auto; padding:0 32px; }
header.top { background:#111827; color:#f9fafb; padding:28px 0 24px; }
.eyebrow { font-size:12px; letter-spacing:.08em; text-transform:uppercase; color:#9ca3af; }
h1 { margin:4px 0 2px; font-size:24px; font-weight:600; overflow-wrap:anywhere; }
.path { color:#9ca3af; font-size:12px; overflow-wrap:anywhere; }
.meta { display:flex; flex-wrap:wrap; gap:6px 22px; margin:16px 0 0; padding:0; }
.meta div { display:flex; gap:6px; font-size:13px; }
.meta dt { color:#9ca3af; } .meta dd { margin:0; color:#e5e7eb; }
main.wrap { padding-top:24px; padding-bottom:48px; }
h2 { font-size:16px; margin:28px 0 12px; display:flex; align-items:baseline; gap:10px; }
h2 small { color:var(--muted); font-weight:400; font-size:13px; }
.verdict { display:flex; align-items:center; gap:12px; font-size:15px; margin:0 0 16px; }
.summary { display:grid; grid-template-columns:repeat(auto-fill, minmax(260px, 1fr)); gap:14px; }
.term-card { background:var(--card); border:1px solid var(--line); border-left:6px solid var(--c);
             border-radius:10px; padding:14px 16px; }
.term-card .name { font-size:17px; font-weight:600; overflow-wrap:anywhere; }
.term-card .line { display:flex; justify-content:space-between; align-items:center; gap:10px; }
.term-card .sub { color:var(--muted); margin-top:4px; }
.badge { font-size:11px; font-weight:700; letter-spacing:.05em; padding:2px 9px; border-radius:999px; white-space:nowrap; }
.badge.found { background:var(--ok-bg); color:var(--ok); } .badge.missing { background:var(--no-bg); color:var(--no); }
.notes { background:var(--warn-bg); border:1px solid #fde68a; color:#78350f; border-radius:10px;
         padding:10px 16px; margin-top:16px; }
.notes ul { margin:4px 0; padding-left:20px; }
.filters { display:flex; flex-wrap:wrap; gap:8px; margin:22px 0 0; }
.filters button { font:inherit; border:1px solid var(--line); background:var(--card); border-radius:999px;
                  padding:4px 12px; cursor:pointer; display:flex; align-items:center; gap:6px; }
.filters button.active { border-color:#111827; background:#111827; color:#fff; }
.swatch { width:10px; height:10px; border-radius:3px; background:var(--c); display:inline-block; flex:none; }
.card { background:var(--card); border:1px solid var(--line); border-radius:12px; margin:0 0 18px; overflow:hidden;
        box-shadow:0 1px 2px rgba(0,0,0,.04); }
.card > header { display:flex; flex-wrap:wrap; align-items:center; gap:6px 12px; padding:12px 18px;
                 border-bottom:1px solid var(--line); }
.card .num { color:var(--muted); font-variant-numeric:tabular-nums; }
.card .page { font-weight:700; }
.card .where { color:var(--muted); overflow-wrap:anywhere; }
.card .count { margin-left:auto; color:var(--muted); white-space:nowrap; }
.chip { font-size:12px; padding:1px 8px; border-radius:6px; background:#f3f4f6; color:#374151; white-space:nowrap; }
.chip.rep { background:#e0e7ff; color:#3730a3; }
.visual { display:grid; grid-template-columns:minmax(0, 1fr); background:var(--soft); border-bottom:1px solid var(--line); }
.visual.has-locator { grid-template-columns:minmax(0, 1fr) 250px; }
.shot { margin:0; padding:16px; display:flex; justify-content:center; align-items:flex-start; }
.shot img { width:auto; height:auto; max-width:100%; max-height:640px; background:#fff; border:1px solid var(--line);
            cursor:zoom-in; }
.locator { padding:16px 16px 16px 0; }
.locator img { width:100%; height:auto; background:#fff; border:1px solid var(--line); cursor:zoom-in; display:block; }
.locator .cap { color:var(--muted); font-size:12px; margin-top:6px; text-align:center; }
.noshot { margin:16px; padding:18px; border:1px dashed #d1d5db; border-radius:8px; color:var(--muted); background:#fff; }
table.hits { width:100%; border-collapse:collapse; font-size:13px; }
table.hits th { text-align:left; font-weight:500; color:var(--muted); padding:8px 18px 6px; border-bottom:1px solid var(--line); }
table.hits td { padding:7px 18px; border-bottom:1px solid #f3f4f6; vertical-align:top; }
table.hits tr:last-child td { border-bottom:0; }
table.hits td.term { white-space:nowrap; font-weight:600; }
table.hits td.text { overflow-wrap:anywhere; }
table.hits td.num { font-variant-numeric:tabular-nums; white-space:nowrap; }
table.hits td.url { font-family:Consolas, "SFMono-Regular", Menlo, monospace; font-size:12px; width:60%; }
.card.link .shot, .card.strip .shot { padding:12px 16px; }
.card.page-shot .shot img { max-height:900px; }
.low { color:var(--warn); font-weight:600; }
mark { background:color-mix(in srgb, var(--c) 20%, transparent); box-shadow:inset 0 -2px 0 var(--c); color:inherit;
       padding:0 1px; border-radius:2px; }
.empty { background:var(--card); border:1px solid var(--line); border-radius:12px; padding:28px; text-align:center;
         color:var(--muted); }
.legend { display:grid; grid-template-columns:repeat(auto-fill, minmax(280px, 1fr)); gap:8px 24px; color:var(--muted); }
.legend b { color:var(--text); font-weight:600; }
footer { color:var(--muted); font-size:12px; margin-top:28px; border-top:1px solid var(--line); padding-top:12px; }
#lightbox { position:fixed; inset:0; display:none; align-items:center; justify-content:center; padding:28px;
            background:rgba(17,24,39,.88); cursor:zoom-out; z-index:10; }
#lightbox.open { display:flex; }
#lightbox img { max-width:100%; max-height:100%; background:#fff; box-shadow:0 12px 48px rgba(0,0,0,.5); }
@media (max-width:860px) {
  .wrap { padding:0 16px; }
  .visual.has-locator { grid-template-columns:minmax(0, 1fr); }
  .locator { padding:0 16px 16px; max-width:300px; }
}
@media print {
  body { background:#fff; } header.top { background:#fff; color:#000; border-bottom:2px solid #000; }
  .filters, #lightbox { display:none !important; } .card { break-inside:avoid; box-shadow:none; }
  .shot img { max-height:none; }
}
"""

JS = """
const box = document.getElementById('lightbox'), big = box.querySelector('img');
document.querySelectorAll('img.zoomable').forEach(img => img.addEventListener('click', () => {
  big.src = img.src; big.alt = img.alt; box.classList.add('open');
}));
box.addEventListener('click', () => box.classList.remove('open'));
document.addEventListener('keydown', e => { if (e.key === 'Escape') box.classList.remove('open'); });
document.querySelectorAll('.filters button').forEach(btn => btn.addEventListener('click', () => {
  const term = btn.dataset.term;
  document.querySelectorAll('.filters button').forEach(b => b.classList.toggle('active', b === btn));
  document.querySelectorAll('article.card').forEach(card => {
    card.hidden = term !== undefined && !JSON.parse(card.dataset.terms).includes(term);
  });
}));
"""


def _esc(value):
    return html.escape(str(value), quote=True)


def _plural(n, word):
    return f"{n} {word}" if n == 1 else f"{n} {word}{'es' if word.endswith('ch') else 's'}"


def _data_uri(data, mime):
    return f"data:{mime};base64,{base64.b64encode(data).decode('ascii')}"


def _marked(text, spans, color):
    out, pos = [], 0
    for s, e in sorted(spans):
        if s < pos:
            continue
        out += [_esc(text[pos:s]), f'<mark style="--c:{color}">{_esc(text[s:e])}</mark>']
        pos = e
    out.append(_esc(text[pos:]))
    return "".join(out)


def _scope(stats):
    kinds = ["diagram", "image"] + (["text"] if stats["text"] is not None else []) \
        + (["hyperlink"] if stats["links"] is not None else [])
    return ", ".join(kinds[:-1]) + " or " + kinds[-1]


def write_html_report(path, *, info, unit, terms, hits, notes, options, stats, colors, graphics, hit_graphic,
                      labels, page_ranges, highlight):
    """Write the report to path. hits must already be sorted; graphics/hit_graphic come from record_graphic."""
    Unit = unit.capitalize()
    ranges = lambda pages: page_ranges(pages).replace("-", "\u2013")
    scope = _scope(stats)
    order = {id(h): i for i, h in enumerate(hits)}
    groups = {}
    for h in hits:
        key = hit_graphic.get(id(h))
        if h.source == "link":  # one card per link target; without screenshots, one card per page
            key = ("link", h.text) if key else ("links", h.page)
        elif h.source == "body":  # one card per page; text repeated on many pages gets its own card
            key = key or (("text repeated", h.text.lower(), h.where) if h.repeated else ("text", h.page))
        elif key is None:
            key = ("no screenshot", h.page, h.where)
        groups.setdefault(key, []).append(h)

    def position(item):
        hs = item[1]
        pages = [h.page for h in hs if h.page is not None]
        ys = [h.bbox[1] for h in hs if h.bbox]
        return (min(pages) if pages else 0, min(ys) if pages and ys else 0, min(order[id(h)] for h in hs))

    ordered = sorted(groups.items(), key=position)
    graphic_groups = [g for g in ordered if g[1][0].source not in ("link", "body")]
    text_groups = [g for g in ordered if g[1][0].source == "body"]
    link_groups = [g for g in ordered if g[1][0].source == "link"]
    cards = [_card(n, key, hs, graphics.get(hit_graphic.get(id(hs[0]))), unit, terms, colors, labels, ranges,
                   highlight)
             for n, (key, hs) in enumerate(graphic_groups + text_groups + link_groups, 1)]
    n_graphic, n_text = len(graphic_groups), len(text_groups)
    graphic_cards, text_cards = cards[:n_graphic], cards[n_graphic:n_graphic + n_text]
    link_cards = cards[n_graphic + n_text:]

    found_terms = [t for t in terms if any(h.term == t for h in hits)]
    term_cards = []
    for t in terms:
        th = [h for h in hits if h.term == t]
        pages = sorted({h.page for h in th if h.page is not None})
        in_text = sum(h.source == "body" for h in th)
        in_links = sum(h.source == "link" for h in th)
        subs = []
        if th:
            sub = f"{len(th)} match{'es' if len(th) != 1 else ''}"
            if pages:
                sub += f" on {len(pages)} {unit}{'s' if len(pages) != 1 else ''}: {ranges(pages)}"
            subs.append(sub)
            split = [f"{len(th) - in_text - in_links} in graphics"]
            if stats["text"] is not None:
                split.append(f"{in_text} in text")
            if stats["links"] is not None:
                split.append(f"{in_links} in hyperlinks")
            if len(split) > 1:
                subs.append(", ".join(split))
            badge = '<span class="badge found">FOUND</span>'
        else:
            subs.append(f"No {scope} contains this term")
            badge = '<span class="badge missing">NOT FOUND</span>'
        term_cards.append(f'<div class="term-card" style="--c:{colors[t]}"><div class="line">'
                          f'<span class="name">\u201c{_esc(t)}\u201d</span>{badge}</div>'
                          + "".join(f'<div class="sub">{_esc(s)}</div>' for s in subs) + "</div>")

    meta = [("Type", info["type"].upper())]
    if "size" in info:
        meta.append(("Size", info["size"]))
    if "pages" in info:
        scanned = info.get("scanned_pages", info["pages"])
        meta.append((f"{Unit}s", f"{info['pages']}" + (f" ({scanned} scanned)" if scanned != info["pages"] else "")))
    if info.get("converted_with"):
        meta.append(("Converted with", info["converted_with"]))
    if info.get("method"):
        meta.append(("Method", info["method"]))
    if "images" in info:
        meta.append(("Graphics", f"{info['images']} images, {info['inline_svgs']} inline SVGs"))
    meta += list(options.items())
    meta.append(("Generated", datetime.datetime.now().strftime("%Y-%m-%d %H:%M")))

    if len(terms) == 1:
        verdict = "found in" if hits else f"not found in any {scope} of"
        verdict = f"\u201c{_esc(terms[0])}\u201d {verdict} this file"
    else:
        verdict = f"{len(found_terms)} of {len(terms)} terms found"
    verdict_badge = '<span class="badge found">FOUND</span>' if hits else '<span class="badge missing">NOT FOUND</span>'

    parts = [
        "<!DOCTYPE html>", '<html lang="en"><head><meta charset="utf-8">',
        '<meta name="viewport" content="width=device-width, initial-scale=1">',
        f"<title>{_esc(', '.join(terms))} in {_esc(os.path.basename(info['file']))}</title>",
        f"<style>{CSS}</style></head><body>",
        '<header class="top"><div class="wrap"><div class="eyebrow">Find word in files</div>',
        f"<h1>{_esc(os.path.basename(info['file']))}</h1><div class=\"path\">{_esc(info['file'])}</div>",
        '<dl class="meta">' + "".join(f"<div><dt>{_esc(k)}</dt><dd>{_esc(v)}</dd></div>" for k, v in meta) + "</dl>",
        '</div></header><main class="wrap">',
        f'<div class="verdict">{verdict_badge}<span>{verdict}</span></div>',
        '<section class="summary">' + "".join(term_cards) + "</section>",
    ]
    if notes:
        parts.append('<section class="notes"><b>Notes</b><ul>' + "".join(f"<li>{_esc(n)}</li>" for n in notes)
                     + "</ul></section>")
    if len(found_terms) > 1:
        parts.append('<nav class="filters"><button class="active">All</button>' + "".join(
            f'<button data-term="{_esc(t)}"><span class="swatch" style="--c:{colors[t]}"></span>{_esc(t)}</button>'
            for t in found_terms) + "</nav>")
    if graphic_cards:
        n = sum(len(hs) for _, hs in graphic_groups)
        parts.append(f"<h2 id=\"graphics\">In diagrams and images <small>{_plural(n, 'match')} in "
                     f"{_plural(len(graphic_cards), 'graphic')}</small></h2>")
        parts += graphic_cards
    if text_cards:
        text_hits = [h for _, hs in text_groups for h in hs]
        text_pages = {h.page for h in text_hits if h.page is not None}
        parts.append(f"<h2 id=\"text\">In text <small>{_plural(len(text_hits), 'match')}"
                     + (f" on {len(text_pages)} {unit}{'s' if len(text_pages) != 1 else ''}" if text_pages else "")
                     + "</small></h2>")
        parts += text_cards
    if link_cards:
        link_hits = [h for _, hs in link_groups for h in hs]
        parts.append(f"<h2 id=\"links\">In hyperlinks <small>{_plural(len(link_hits), 'match')} in "
                     f"{_plural(len({h.text for h in link_hits}), 'link target')}</small></h2>")
        parts += link_cards
    checked = [_plural(stats["checked"], "graphic")]
    if stats["text"] is not None:
        n = stats["text"]
        checked.append(f"{n} {'line or paragraph' if n == 1 else 'lines or paragraphs'} of text")
    if stats["links"] is not None:
        checked.append(_plural(stats["links"], "hyperlink"))
    checked = ", ".join(checked[:-1]) + " and " + checked[-1] if len(checked) > 1 else checked[0]
    if not cards:
        parts.append(f'<h2>Matches</h2><div class="empty">None of the {checked} checked contain '
                     f'{"this term" if len(terms) == 1 else "these terms"}.</div>')
    legend = LEGEND + ([TEXT_LEGEND] if stats["text"] is not None else []) \
        + ([LINK_LEGEND] if stats["links"] is not None else [])
    parts.append('<h2>How each match was found</h2><div class="legend">' + "".join(
        f"<div><b>{_esc(name)}</b>: {_esc(text)}</div>" for name, text in legend) + "</div>")
    footer = (f"Checked {checked}; OCR ran on {stats['ocr_runs']} distinct graphics; {stats['elapsed']:.1f} s. "
              "Boxes in the screenshots use each term's colour.")
    if any(g.thumb for g in graphics.values()):
        footer += f" The amber outline on a {unit} thumbnail shows where the graphic sits."
    parts.append(f"<footer>{_esc(footer)}</footer>")
    parts += ['</main><div id="lightbox"><img alt=""></div>', f"<script>{JS}</script></body></html>"]

    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(parts))


def _card(n, key, hs, graphic, unit, terms, colors, labels, ranges, highlight):
    Unit = unit.capitalize()
    link = hs[0].source == "link"
    text = hs[0].source == "body"
    pages = sorted({h.page for h in hs if h.page is not None})
    approx = "~" if any(h.page_approx for h in hs) else ""
    wheres = list(dict.fromkeys(h.where for h in hs if h.where))
    if not pages:
        title = "Links" if link else "Text" if text else "Image"
    elif len(pages) == 1:
        title = f"{Unit} {approx}{pages[0]}"
    else:
        title = f"{Unit}s {approx}{ranges(pages)}"
    kinds = list(dict.fromkeys(labels[h.source, h.region] for h in hs))
    head = [f'<span class="num">#{n}</span>', f'<span class="page">{_esc(title)}</span>']
    if len(wheres) == 1:
        head.append(f'<span class="where">{_esc(wheres[0])}</span>')
    head += [f'<span class="chip">{_esc(k)}</span>' for k in kinds]
    if len(pages) > 1:
        same = "link" if link else "text" if text else "graphic"
        head.append(f'<span class="chip rep">same {same} on {len(pages)} {unit}s</span>')
    head.append(f'<span class="count">{_plural(len(hs), "match")}</span>')

    if graphic and graphic.image:
        w, h = graphic.size
        shot = (f'<figure class="shot"><img class="zoomable" src="{_data_uri(graphic.image, graphic.mime)}" '
                f'width="{w}" height="{h}" alt="{_esc(title)}: matches boxed"></figure>')
    elif link:
        shot = ('<div class="noshot">No screenshot: these links were read from the file\'s markup, '
                'which has no page picture to mark.</div>')
    elif text:
        shot = ('<div class="noshot">No screenshot: this text was read from the file\'s markup or XML, '
                'which has no page picture to mark.</div>')
    else:
        shot = ('<div class="noshot">No screenshot: this text was read from the file structure (for example SmartArt, '
                'a chart, grouped shapes or an SVG that could not be drawn), so there is no picture to mark.</div>')
    locator = ""
    if graphic and graphic.thumb:
        locator = (f'<aside class="locator"><img class="zoomable" src="{_data_uri(graphic.thumb, "image/jpeg")}" '
                   f'alt="{_esc(Unit)} {pages[0] if pages else ""} with the graphic outlined">'
                   f'<div class="cap">Position on {unit} {approx}{pages[0] if pages else ""}</div></aside>')
    visual = f'<div class="visual{" has-locator" if locator else ""}">{shot}{locator}</div>'

    if link:
        table, kind = _link_table(hs, terms, colors, highlight), " link"
    elif text:
        table = _text_table(hs, terms, colors, highlight)
        kind = " strip" if len(pages) > 1 else " page-shot"
    else:
        table, kind = _hit_table(hs, terms, colors, labels, highlight), ""
    card_terms = _esc(json.dumps(list(dict.fromkeys(h.term for h in hs)), ensure_ascii=False))
    return (f'<article class="card{kind}" data-terms="{card_terms}">'
            f'<header>{"".join(head)}</header>{visual}{table}</article>')


def _term_cell(term, colors):
    return f'<td class="term"><span class="swatch" style="--c:{colors[term]}"></span> {_esc(term)}</td>'


def _hit_table(hs, terms, colors, labels, highlight):
    rows = {}
    for h in sorted(hs, key=lambda h: terms.index(h.term)):
        rows.setdefault((h.term, h.text, h.source, h.region), []).append(h)
    trs = []
    for (term, text, source, region), rhs in rows.items():
        scores = [x.score for x in rhs if x.score is not None]
        if scores:
            low = ' class="low" title="Low OCR confidence: check the screenshot"' if min(scores) < 0.8 else ""
            conf = f"<span{low}>{min(scores):.2f}</span>"
        else:
            conf = "exact text"
        trs.append(f'<tr>{_term_cell(term, colors)}'
                   f'<td class="text">{_marked(text, highlight(text, term, source), colors[term])}</td>'
                   f'<td>{_esc(labels[source, region])}</td><td class="num">{conf}</td>'
                   f'<td class="num">{len(rhs)}</td></tr>')
    return ('<table class="hits"><thead><tr><th>Term</th><th>Matched text</th><th>Found by</th>'
            '<th>Confidence</th><th>Count</th></tr></thead><tbody>' + "".join(trs) + "</tbody></table>")


def _window(text, spans, n=260):
    """A long paragraph cut to about n characters around its first match, with the spans moved to match."""
    if len(text) <= n or not spans:
        return text, spans
    first = min(s for s, _ in spans)
    a = max(0, min(first - 80, len(text) - n))
    lead = "\u2026" if a else ""
    cut = lead + text[a:a + n] + ("\u2026" if a + n < len(text) else "")
    return cut, [(s - a + len(lead), e - a + len(lead)) for s, e in spans if s >= a and e <= a + n]


def _text_table(hs, terms, colors, highlight):
    rows = {}
    for h in sorted(hs, key=lambda h: terms.index(h.term)):
        rows.setdefault((h.term, h.text, h.where), []).append(h)
    trs = []
    for (term, text, where), rhs in rows.items():
        shown, spans = _window(text, highlight(text, term, "body"))
        trs.append(f'<tr>{_term_cell(term, colors)}<td class="text">{_marked(shown, spans, colors[term])}</td>'
                   f'<td class="text">{_esc(where)}</td><td class="num">{len(rhs)}</td></tr>')
    return ('<table class="hits"><thead><tr><th>Term</th><th>Matched text</th><th>Where</th>'
            '<th>Count</th></tr></thead><tbody>' + "".join(trs) + "</tbody></table>")


def _link_table(hs, terms, colors, highlight):
    rows = {}
    for h in sorted(hs, key=lambda h: terms.index(h.term)):
        rows.setdefault((h.term, h.text, h.where), []).append(h)
    trs = [f'<tr>{_term_cell(term, colors)}'
           f'<td class="text url">{_marked(text, highlight(text, term, "link"), colors[term])}</td>'
           f'<td class="text">{_esc(where)}</td><td class="num">{len(rhs)}</td></tr>'
           for (term, text, where), rhs in rows.items()]
    return ('<table class="hits"><thead><tr><th>Term</th><th>Link target</th><th>Where</th>'
            '<th>Count</th></tr></thead><tbody>' + "".join(trs) + "</tbody></table>")
