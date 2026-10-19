"""Shared helpers for the language editor's scripts.

Nothing in here touches a .docx directly: docx work goes through
docx_utils (loaded by load_docx_utils below from one exact file path).
These are the parsing and bookkeeping primitives the scripts share.
"""
import importlib.util
import json
import os
import re
import sys
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parent.parent
BUNDLED_SCRIPTS = SKILL_DIR / "scripts"
DOCX_UTILS_ENV = "LANGUAGE_EDITOR_DOCX_UTILS"

# ANCHOR: a phrase copied from the sentence a COMMENT is about (the comment then
# highlights that sentence; see apply_language_edits.py).
EDIT_FIELDS = ("ORIGINAL", "ACTION", "IMPORTANCE", "REVISED", "COMMENT", "ANCHOR", "PASSES", "ISSUE_LABELS")
OVERRIDE_CLASSES = ("reversal", "partial", "drift")
IMPORTANCE_VALUES = ("MAJOR", "MINOR")
ACTIONS = ("KEEP", "REVISE")


# ---------------------------------------------------------------------------
# docx_utils resolution (workflows/proofread.md, File resolution)
# ---------------------------------------------------------------------------
def load_docx_utils(verbose=True):
    """Import docx_utils from one exact file, so a run never depends on the working folder.

    Default: the copy bundled with the skill. To use another copy, set the
    environment variable LANGUAGE_EDITOR_DOCX_UTILS to its path; the tests do
    this to make a step fail on purpose. When it is set, a notice is printed
    even with verbose=False, so a run never uses another copy unannounced.
    """
    override = os.environ.get(DOCX_UTILS_ENV)
    path = Path(override).resolve() if override else BUNDLED_SCRIPTS / "docx_utils.py"
    if not path.is_file():
        raise RuntimeError("docx_utils.py not found at %s" % path)
    spec = importlib.util.spec_from_file_location("docx_utils", str(path))
    module = importlib.util.module_from_spec(spec)
    sys.modules["docx_utils"] = module
    spec.loader.exec_module(module)
    if override:
        print("NOTICE: docx_utils loaded from %s, not the skill's own copy (%s is set)"
              % (str(path).encode("ascii", "replace").decode(), DOCX_UTILS_ENV))
    elif verbose:
        print("docx_utils: %s" % str(path).encode("ascii", "replace").decode())
    return module


# ---------------------------------------------------------------------------
# Text normalization
# ---------------------------------------------------------------------------
def norm(s):
    """Quote-normalized, whitespace-collapsed text for matching."""
    s = (s.replace("\u2018", "'").replace("\u2019", "'")
          .replace("\u201c", '"').replace("\u201d", '"')
          .replace("\u2013", "-").replace("\u2014", "-"))
    return re.sub(r"\s+", " ", s).strip()


def tidy_comment(s):
    """Comment text as delivered: runs of spaces and tabs collapsed, ends trimmed.

    Dashes and every other character stay exactly as the editor wrote them.
    """
    return re.sub(r"[ \t]+", " ", s).strip()


# ---------------------------------------------------------------------------
# Edit files
# ---------------------------------------------------------------------------
def parse_edit_file(path, fields=EDIT_FIELDS):
    """Parse one agent edit file into a list of entry dicts.

    Reads the block after an 'Edit File' heading ('## Edit File' or
    '### 1. Edit File'; the whole file if there is none). Entries start at
    '### P###' (three or more digits) and end at a '---' line or the next entry.
    """
    path = Path(path)
    text = path.read_text(encoding="utf-8")
    m = re.search(r"^#+\s*(?:\d+\.\s*)?Edit File\s*$", text, re.M | re.I)
    body = text[m.end():] if m else text

    field_re = re.compile(r"^(%s):\s?(.*)$" % "|".join(fields))
    entries, cur, field = [], None, None
    for line in body.splitlines():
        head = re.match(r"^#+\s*(P\d{3,})\b", line)
        if head:
            if cur:
                entries.append(cur)
            cur, field = {"id": head.group(1), "source": path.name}, None
            continue
        if cur is None:
            continue
        if line.strip() == "---":
            entries.append(cur)
            cur, field = None, None
            continue
        f = field_re.match(line)
        if f:
            field = f.group(1)
            cur[field] = f.group(2)
        elif field:
            cur[field] = (cur[field] + "\n" + line).strip("\n")
    if cur:
        entries.append(cur)

    for e in entries:
        for k in fields:
            if k in e:
                e[k] = e[k].strip("\r\n") if k == "REVISED" else e[k].strip()
        e["ACTION"] = (e.get("ACTION", "KEEP").split() or ["KEEP"])[0].upper()
        if e.get("IMPORTANCE"):
            e["IMPORTANCE"] = e["IMPORTANCE"].split()[0].strip("[]").upper()
    return entries


def pid_number(pid):
    """Paragraph position from an id: 'P007' -> 7, 'P1000' -> 1000."""
    return int(pid[1:])


def find_rule_set(references=None):
    """Name of the one rules file shipped in references/, recorded in the manifest."""
    folder = Path(references) if references else SKILL_DIR / "references"
    found = sorted(folder.glob("rules_language_v*.txt"))
    if len(found) != 1:
        raise SystemExit("expected exactly one rules_language_v*.txt in %s, found %d; pass --rule-set"
                         % (folder, len(found)))
    return found[0].stem


def has_comment(entry):
    c = entry.get("COMMENT", "")
    return bool(c) and c.strip().lower() not in ("null", "none", "n/a")


def validate_importance(entries):
    """Every entry carrying a comment must declare IMPORTANCE (workflows/proofread.md, Required Outputs).

    Returns (ok, problems).
    """
    problems = []
    for e in entries:
        if not has_comment(e):
            continue
        imp = e.get("IMPORTANCE", "")
        if imp == "CRITICAL" or re.match(r"^\[CRITICAL\]", e["COMMENT"], re.I):
            problems.append("%s (%s): CRITICAL is not a level; use MAJOR or MINOR" % (e["id"], e.get("source", "?")))
            continue
        if imp in IMPORTANCE_VALUES:
            continue
        if re.match(r"^\[(MAJOR|MINOR)\]", e["COMMENT"]):
            e["IMPORTANCE"] = re.match(r"^\[(\w+)\]", e["COMMENT"]).group(1).upper()
            continue
        problems.append("%s (%s): IMPORTANCE missing or invalid (%r)" % (e["id"], e.get("source", "?"), imp))
    return not problems, problems


def load_para_index(path):
    """Read a para_index.txt written by index_manuscript.py into {P###: text}."""
    out = {}
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        m = re.match(r"^(P\d{3,})(?: \[[^\]]*\])? ?(.*)$", line)
        if m:
            out[m.group(1)] = m.group(2)
    return out


def write_manifest(path, manifest):
    Path(path).write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")
