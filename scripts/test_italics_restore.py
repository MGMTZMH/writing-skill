"""Regression test: statistical symbols keep their italics through a replacement.

`replace_paragraph_text` used to drop italic formatting on every run shorter
than two characters, which silently de-italicized the p, n, F and t of
statistical notation on any paragraph an editor touched. Restoring them is
boundary-checked and case-sensitive, so the test asserts BOTH directions:
the standalone symbols come back italic, and nothing else is italic (a
letter inside an ordinary word never is).

Run it after any change to _restore_inline_formatting or to
SINGLE_CHAR_ITALIC_TOKENS:

    python test_italics_restore.py

Exits 0 when every assertion holds, 1 otherwise, and prints one line per case.
"""
import sys

sys.dont_write_bytecode = True  # never leave a __pycache__ folder inside the skill

import tempfile  # noqa: E402
from pathlib import Path  # noqa: E402

from editor_io import load_docx_utils  # noqa: E402

docx_utils = load_docx_utils(verbose=True)
from docx import Document  # noqa: E402

CASES = [
    # (description, [(text, italic), ...], the symbols that must be italic)
    ("standalone p and n beside the words paper and not",
     [("The coefficient is positive (", False), ("p", True), (" < .05) with ", False),
      ("n", True), (" = 402; this paper does not report panel a.", False)],
     "pn"),
    ("F and t in a test report",
     [("The restriction is rejected, ", False), ("F", True), ("(1, 385) = 6.19, and ", False),
      ("t", True), (" = 2.49 for the first contrast.", False)],
     "Ft"),
    ("Greek beta",
     [("The estimate is ", False), ("\u03b2", True), (" = 0.104 in the full model.", False)],
     "\u03b2"),
]


def italic_problems(para, expect_italic):
    """What is wrong with the italics of one paragraph after a replacement.

    Every expected symbol must be italic, and every italic run must be a single
    expected symbol standing alone, never part of a longer word.
    """
    problems, italic_tokens, text, pos = [], [], para.text, 0
    for r in para.runs:
        start, pos = pos, pos + len(r.text)
        token = r.text.strip()
        if not r.italic or not token:
            continue
        s = start + r.text.index(token)
        e = s + len(token)
        alone = (s == 0 or not text[s - 1].isalnum()) and (e >= len(text) or not text[e].isalnum())
        if len(token) != 1 or token not in expect_italic or not alone:
            problems.append("unexpected italic text %a" % token)
        italic_tokens.append(token)
    for ch in expect_italic:
        if ch not in italic_tokens:
            problems.append("symbol %a lost its italics" % ch)
    return problems


def run_case(tmp, i, desc, runs, expect_italic):
    doc = Document()
    para = doc.add_paragraph()
    for text, ital in runs:
        r = para.add_run(text)
        if ital:
            r.italic = True
    src = tmp / ("case%d_src.docx" % i)
    doc.save(str(src))

    full_text = "".join(t for t, _ in runs)
    doc2 = Document(str(src))
    docx_utils.replace_paragraph_text(doc2.paragraphs[0], full_text)
    out = tmp / ("case%d_out.docx" % i)
    doc2.save(str(out))

    para3 = Document(str(out)).paragraphs[0]
    italic_text = "".join(r.text for r in para3.runs if r.italic)
    problems = italic_problems(para3, expect_italic)
    if para3.text != full_text:
        problems.append("paragraph text changed: %a" % para3.text)

    status = "PASS" if not problems else "FAIL"
    print("%s  %s" % (status, desc))
    for p in problems:
        print("        %s" % p)
    if not problems:
        print("        italic runs: %a" % italic_text)
    return not problems


def main():
    with tempfile.TemporaryDirectory(prefix="editor_italics_") as td:
        tmp = Path(td)
        results = [run_case(tmp, i, desc, runs, exp)
                   for i, (desc, runs, exp) in enumerate(CASES, start=1)]
    passed = sum(results)
    print("%d of %d cases passed" % (passed, len(results)))
    return 0 if passed == len(results) else 1


if __name__ == "__main__":
    raise SystemExit(main())
