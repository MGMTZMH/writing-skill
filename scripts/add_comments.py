"""Inject Word comments into a COPY of a .docx, with safety checks.

Usage:
    python add_comments.py <original.docx> <comments.json> --out <output.docx>
                           [--author "Structure Reviewer"] [--date ISO8601]

comments.json is a list of objects with at least:
    "anchor_quote": verbatim phrase that occurs in exactly one paragraph, and
                    only once inside it
    "comment":      the comment text  (alias: "comment_text")
Other keys are ignored. An empty list copies the original byte-for-byte and
reports zero injected comments.

Each comment highlights the sentence that holds its quote (both sentences
when the quote runs across a sentence end), not the whole paragraph. A quote
taken from a text box is refused as ambiguous (the box's own paragraph and its
host paragraph both contain it), so quotes come from body or table
paragraphs; if a quote cannot be mapped to a sentence all the same, the
comment covers the whole paragraph and the OK line says so.

Checks, in the order they run (the script stops at the first failing group):
  1. the original exists; the output path differs from it and does not
     already exist unless --force is given
  2. the comments file reads as a list, and every comment has text
  3. every anchor occurs in exactly one paragraph of the document body
     (tables included), after normalizing curly quotes, apostrophes, dashes
     and non-breaking spaces on both sides, and only once inside that
     paragraph, counted as the placement counts it (white space collapsed,
     overlapping occurrences included); the exact substring from the
     document is what gets injected
  4. the injection reports the same count as the number of comments
  5. every body paragraph's text in the output equals the original's
Exit codes: 0 ok, 2 anchor problems, 3 injection count mismatch,
4 body text changed, 5 other error, 6 output exists (use --force).
Output is ASCII only.
"""
import argparse
import datetime as _dt
import json
import os
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.dont_write_bytecode = True  # never leave a __pycache__ folder inside the skill

from docx import Document  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

from docx_utils import inject_comments_ex, phrase_count  # noqa: E402
from docx_utils import long_path as fs  # noqa: E402  (paths over 260 characters on Windows)

W_P = qn("w:p")
W_T = qn("w:t")

_NORMALIZE = {
    0x2018: "'", 0x2019: "'", 0x201A: "'", 0x201B: "'",
    0x201C: '"', 0x201D: '"', 0x201E: '"', 0x201F: '"',
    0x2013: "-", 0x2014: "-", 0x2012: "-", 0x2010: "-", 0x2011: "-",
    0x00A0: " ", 0x2009: " ", 0x202F: " ",
}


def norm(s):
    return s.translate(_NORMALIZE)


def ascii_safe(s):
    return s.encode("ascii", "replace").decode("ascii")


def body_texts(docx_path):
    """All body paragraph texts, same iteration inject_comments uses."""
    doc = Document(str(fs(docx_path)))
    return ["".join(t.text or "" for t in p.iter(W_T)) for p in doc.element.body.iter(W_P)]


def resolve_anchors(texts, comments):
    """Return (resolved, errors). resolved = list of exact substrings."""
    normed = [norm(t) for t in texts]
    resolved, errors = [], []
    for i, c in enumerate(comments):
        anchor = (c.get("anchor_quote") or "").strip()
        if not anchor:
            errors.append("comment %d: empty anchor_quote" % (i + 1))
            resolved.append(None)
            continue
        na = norm(anchor)
        hits = [k for k, t in enumerate(normed) if na in t]
        if len(hits) == 0:
            errors.append("comment %d: anchor not found: %s" % (i + 1, ascii_safe(anchor[:70])))
            resolved.append(None)
        elif len(hits) > 1:
            previews = "; ".join(ascii_safe(texts[k][:50]) for k in hits[:4])
            errors.append("comment %d: anchor found in %d paragraphs (%s): %s" % (i + 1, len(hits), previews, ascii_safe(anchor[:70])))
            resolved.append(None)
        elif phrase_count(texts[hits[0]], anchor) > 1:
            # counted as the placement counts it, so a repeat that differs only in spacing is caught too
            errors.append("comment %d: anchor found %d times in one paragraph (quote more words so it occurs once): %s"
                          % (i + 1, phrase_count(texts[hits[0]], anchor), ascii_safe(anchor[:70])))
            resolved.append(None)
        else:
            k = hits[0]
            start = normed[k].find(na)
            exact = texts[k][start:start + len(na)]
            if exact not in texts[k]:
                errors.append("comment %d: could not recover exact substring" % (i + 1))
                resolved.append(None)
            else:
                resolved.append(exact)
    return resolved, errors


def main(argv=None):
    ap = argparse.ArgumentParser(description="Inject Word comments into a copy of a .docx with checks.")
    ap.add_argument("original")
    ap.add_argument("comments_json")
    ap.add_argument("--out", required=True)
    ap.add_argument("--author", default="Structure Reviewer")
    ap.add_argument("--date", default=None, help="ISO 8601 timestamp; default now (UTC)")
    ap.add_argument("--force", action="store_true", help="overwrite an existing --out file")
    args = ap.parse_args(argv)

    src, out = Path(args.original), Path(args.out)
    if not fs(src).exists():
        print("ERROR: original not found")
        return 5
    if os.path.normcase(str(fs(src))) == os.path.normcase(str(fs(out))):
        print("ERROR: --out must differ from the original (the original is never modified)")
        return 5
    if fs(out).exists() and not args.force:
        print("ERROR: output already exists: %s (pass --force to overwrite, or choose a new name such as _v2)" % ascii_safe(out.name))
        return 6

    try:
        comments = json.loads(fs(args.comments_json).read_text(encoding="utf-8"))
    except Exception as exc:  # noqa: BLE001
        print("ERROR: cannot read comments JSON: %s" % ascii_safe(str(exc)))
        return 5
    if not isinstance(comments, list):
        print("ERROR: comments JSON must be a list")
        return 5
    for i, c in enumerate(comments):
        if "comment" not in c and "comment_text" in c:
            c["comment"] = c["comment_text"]
        if not (c.get("comment") or "").strip():
            print("ERROR: comment %d has empty text" % (i + 1))
            return 5

    texts_before = body_texts(src)
    resolved, errors = resolve_anchors(texts_before, comments)
    if errors:
        print("ANCHOR CHECK FAILED (%d problem(s)); nothing written:" % len(errors))
        for e in errors:
            print("  - " + e)
        return 2

    fs(out.parent).mkdir(parents=True, exist_ok=True)
    shutil.copyfile(str(fs(src)), str(fs(out)))

    date = args.date or _dt.datetime.now(_dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    items = [{"search": anchor, "text": c["comment"], "author": args.author, "highlight": anchor}
             for anchor, c in zip(resolved, comments)]
    try:
        added, notes = inject_comments_ex(fs(out), items, date=date)
    except Exception as exc:  # noqa: BLE001 -- report with a documented exit code
        print("ERROR: comment injection failed (%s); the copy is kept for inspection: %s"
              % (ascii_safe(str(exc)), ascii_safe(out.name)))
        return 5
    if added != len(comments):
        print("INJECTION COUNT MISMATCH: %d injected, %d expected; output kept for inspection: %s" % (added, len(comments), ascii_safe(out.name)))
        return 3

    texts_after = body_texts(out)
    if texts_after != texts_before:
        n_diff = sum(1 for a, b in zip(texts_before, texts_after) if a != b) + abs(len(texts_after) - len(texts_before))
        print("BODY TEXT CHANGED: %d paragraph(s) differ; output kept for inspection: %s" % (n_diff, ascii_safe(out.name)))
        return 4

    print("OK: injected %d comment(s) by '%s' into %s; body text unchanged (%d paragraphs); "
          "%d on a sentence, %d on a whole paragraph"
          % (added, ascii_safe(args.author), ascii_safe(out.name), len(texts_after), added - len(notes), len(notes)))
    for n, message in notes:
        print("  NOTE: comment %d: %s" % (n + 1, ascii_safe(message)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
