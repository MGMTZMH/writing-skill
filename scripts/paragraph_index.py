"""Index the paragraphs of a .docx so reviewers can cite them by ID.

Usage:
    python paragraph_index.py <input.docx> [--out <paragraph_index.json>]

Writes two files:
  <out>                       JSON list of {id, text, style, is_heading, heading_hint, in_table, notes}
  <out stem>_map.txt          one line per paragraph: id | style | first 60 chars (ASCII-safe),
                              then [notes: ...] when the paragraph carries footnote or endnote marks

IDs (P000, P001, ...) run over paragraphs with non-empty text, in document
order: top-level body paragraphs, paragraphs inside tables (in_table =
true) and inside content controls. Paragraphs inside text boxes are NOT
indexed: their text is folded into the host paragraph's text, which is
also where the comment injector looks for a quote, so anchors must be taken
from body or table paragraphs. Text is the concatenation of all w:t
elements, the same extraction the comment injector uses to find the
paragraph, so an anchor found here is an anchor it can find; the comment
then highlights the sentence holding the quote. A quote taken from a text
box is refused as ambiguous by add_comments.py (the box's paragraph and its
host paragraph both contain it). The texts of the footnotes and endnotes a
paragraph's marks point to are listed in its notes field
([{kind, id, text}], in mark order), apart from its text, so an anchor still
matches the body text while a reader of the index can see what the notes
say. Nothing is printed except a one-line ASCII summary.
"""
import argparse
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.dont_write_bytecode = True  # never leave a __pycache__ folder inside the skill

from docx import Document  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.text.paragraph import Paragraph  # noqa: E402

from docx_utils import long_path as fs  # noqa: E402  (paths over 260 characters on Windows)
from docx_utils import note_texts, paragraph_notes  # noqa: E402


W_P = qn("w:p")
W_TBL = qn("w:tbl")
W_SDT = qn("w:sdt")
W_T = qn("w:t")
W_R = qn("w:r")
W_RPR = qn("w:rPr")
W_B = qn("w:b")


def para_text(p_el):
    return "".join(t.text or "" for t in p_el.iter(W_T))


def iter_paragraphs(doc):
    """Yield (element, in_table) for body paragraphs in document order."""
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == W_P:
            yield child, False
        elif child.tag == W_TBL:
            for p in child.iter(W_P):
                yield p, True
        elif child.tag == W_SDT:
            for p in child.iter(W_P):
                yield p, False


def style_name(doc, p_el):
    try:
        return Paragraph(p_el, doc._body).style.name or ""
    except Exception:  # noqa: BLE001
        return ""


def all_runs_bold(p_el):
    runs = [r for r in p_el.iter(W_R) if para_text(r).strip()]
    if not runs:
        return False
    for r in runs:
        rpr = r.find(W_RPR)
        if rpr is None or rpr.find(W_B) is None:
            return False
    return True


def ascii_safe(s):
    return s.encode("ascii", "replace").decode("ascii")


def build_index(docx_path):
    doc = Document(str(fs(docx_path)))
    texts = note_texts(doc)
    items = []
    n = 0
    for p_el, in_table in iter_paragraphs(doc):
        text = para_text(p_el)
        if not text.strip():
            continue
        style = style_name(doc, p_el)
        is_heading = style.lower().startswith(("heading", "title", "subtitle"))
        words = text.strip().split()
        heading_hint = (
            not is_heading
            and len(words) <= 12
            and not text.strip().endswith((".", "?", "!", ":", ";", ","))
            and all_runs_bold(p_el)
        )
        items.append(
            {
                "id": "P%03d" % n,
                "text": text,
                "style": style,
                "is_heading": is_heading,
                "heading_hint": heading_hint,
                "in_table": in_table,
                "notes": paragraph_notes(p_el, texts),
            }
        )
        n += 1
    return items


def main(argv=None):
    ap = argparse.ArgumentParser(description="Index the paragraphs of a .docx by ID.")
    ap.add_argument("docx")
    ap.add_argument("--out", default=None)
    args = ap.parse_args(argv)

    src = Path(args.docx)
    if not fs(src).exists():
        print("ERROR: input not found: %s" % ascii_safe(str(src)))
        return 2
    out = Path(args.out) if args.out else src.with_name(src.stem + "_paragraph_index.json")
    fs(out.parent).mkdir(parents=True, exist_ok=True)

    items = build_index(src)
    fs(out).write_text(json.dumps(items, ensure_ascii=False, indent=1), encoding="utf-8")

    map_path = out.with_name(out.stem + "_map.txt")
    lines = []
    for it in items:
        flag = "H" if it["is_heading"] else ("h" if it["heading_hint"] else " ")
        tbl = "T" if it["in_table"] else " "
        line = "%s %s%s %-22s | %s" % (it["id"], flag, tbl, it["style"][:22], ascii_safe(it["text"][:60]))
        if it["notes"]:
            line += "  [notes: %s]" % ", ".join("%s %s" % (n["kind"], n["id"]) for n in it["notes"])
        lines.append(line)
    fs(map_path).write_text("\n".join(lines) + "\n", encoding="utf-8")

    print("OK: indexed %d non-empty paragraphs -> %s (map: %s)" % (len(items), ascii_safe(out.name), ascii_safe(map_path.name)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
