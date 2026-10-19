"""Create source-based tracked Word revisions after independent meaning review.

--out is the internal accepted/clean reference; --tracked-out is the author's
Word document; --plan-out stores COMPLETE approved text and source identities.
Use fresh, distinct paths. With omitted new paths, compatible defaults are
<out-stem>_TRACKED.docx and approved_edit_plan.json beside --out.

Exit 0: ready, validated output (possibly with recorded unsupported-edit skips).
Exit 1: dependency, preflight or input error. Exit 2: command-line usage error.
Exit 3: unresolved incoming revisions. Exit 4: comment placement failure.
Exit 5: plan/preservation/validation failure. Exit 6: publication failure.
Any nonzero result is NOT deliverable. The manifest is written ready LAST.

Example (one command line in any shell, quote paths):
  python apply_language_edits.py --input original.docx --out internal/edited_CLEAN.docx --tracked-out edited_TRACKED.docx --plan-out internal/approved_edit_plan.json --edits edits.md --para-index para_index.txt --overrides verifier_overrides.json --manifest language_edit_manifest.json

Overrides: {} when none; each P### maps to class reversal, partial (full
approved paragraph in revised), or drift (notes containing anchor and note).
"""
import sys
sys.dont_write_bytecode = True
import argparse
import fnmatch
import glob
import hashlib
import importlib.metadata
import json
import os
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from io import BytesIO

from editor_io import (ACTIONS, OVERRIDE_CLASSES, find_rule_set, has_comment,
                       load_docx_utils, load_para_index, norm, parse_edit_file,
                       pid_number, tidy_comment, validate_importance)

DRIFT_PREFIX = "Please check this sentence: the edit may have shifted the emphasis slightly while keeping your claim unchanged. What changed: "

def dependency_check():
    if sys.version_info < (3, 12):
        raise ValueError("Python 3.12 or newer is required; tested minimum is Python 3.12")
    required = {"python-docx": "1.2.0", "lxml": "6.1.1"}
    actual = {}
    for name, version in required.items():
        try:
            actual[name] = importlib.metadata.version(name)
        except importlib.metadata.PackageNotFoundError as exc:
            raise ValueError("missing dependency %s; install this skill's requirements.txt" % name) from exc
        if actual[name] != version:
            raise ValueError("%s==%s required; installed %s" % (name, version, actual[name]))
    return {"python": sys.version, "executable": sys.executable, "packages": actual}

def expand_edit_paths(patterns, long_path):
    paths = []
    for pat in patterns:
        folder, name = os.path.split(pat)
        if any(ch in name for ch in "*?["):
            if not any(ch in folder for ch in "*?"):
                where = long_path(folder or ".")
                hits = sorted([p for p in where.iterdir() if fnmatch.fnmatch(p.name, name)], key=lambda p:p.name) if where.is_dir() else []
            else:
                hits = [long_path(p) for p in sorted(glob.glob(pat))]
        else:
            hits = [long_path(pat)] if long_path(pat).is_file() else []
        if not hits:
            raise ValueError("no edit file matches " + pat)
        paths.extend(long_path(hit) for hit in hits)
    if len({str(p.resolve()) for p in paths}) != len(paths):
        raise ValueError("an edit file was supplied more than once")
    return paths

def find_paragraph(doc, entry, paras, view):
    """Exact indexed location first; one unique fallback, never a guess."""
    pos = pid_number(entry["id"])
    texts = [norm(view(p._p)) for p in doc.paragraphs]
    tests = []
    if entry["id"] in paras:
        target = norm(paras[entry["id"]])
        tests.append(("index-text", lambda t:t == target))
    prefix = norm(entry.get("ORIGINAL", "")).strip('"').strip("[]")
    if len(prefix) >= 20:
        tests.append(("prefix", lambda t:t.startswith(prefix)))
    for how, test in tests:
        if pos < len(texts) and test(texts[pos]):
            return pos, doc.paragraphs[pos], how
    for how, test in tests:
        hits = [i for i,t in enumerate(texts) if test(t)]
        if len(hits) == 1:
            return hits[0], doc.paragraphs[hits[0]], how
        if len(hits) > 1:
            return None, None, "ambiguous: %s matches paragraphs %s" % (how, hits)
    return None, None, "no paragraph matched"

def json_bytes(value):
    return json.dumps(value, ensure_ascii=False, indent=2).encode("utf-8")

def path_key(path):
    return os.path.normcase(str(path.resolve()))

def publish_file(stage, dest):
    """Exclusive create; no check-then-overwrite race, even across filesystems."""
    created = False
    try:
        with dest.open("xb") as out:
            created = True
            with stage.open("rb") as src:
                while True:
                    block = src.read(1024 * 1024)
                    if not block:
                        break
                    out.write(block)
            out.flush()
            os.fsync(out.fileno())
    except Exception:
        if created:
            dest.unlink(missing_ok=True)
        raise

def run(args):
    runtime = dependency_check()
    du = load_docx_utils()
    import docx_revisions as dr
    # Share explicitly selected utility module, including test-injected failures.
    dr.du = du
    from docx import Document
    lp = du.long_path
    src, out = lp(args.input), lp(args.out)
    tracked = lp(args.tracked_out or out.with_name(out.stem.replace("_CLEAN", "") + "_TRACKED.docx"))
    planpath = lp(args.plan_out or out.parent / "approved_edit_plan.json")
    manifestpath = lp(args.manifest or out.parent / "language_edit_manifest.json")
    outputs = [out, tracked, planpath, manifestpath]
    editpaths = expand_edit_paths(args.edits, lp)
    inputs = [src, lp(args.overrides)] + editpaths + ([lp(args.para_index)] if args.para_index else [])
    if len({path_key(p) for p in outputs}) != len(outputs) or {path_key(p) for p in outputs} & {path_key(p) for p in inputs}:
        raise ValueError("input/output paths collide; supply distinct fresh paths")
    for p in outputs:
        if p.exists():
            raise ValueError("refusing to overwrite existing output: " + str(p))
    if not lp(args.overrides).is_file():
        raise ValueError("overrides file not found: " + args.overrides)
    for p in inputs:
        if not p.is_file():
            raise ValueError("input file not found: " + str(p))
    source_bytes = src.read_bytes()
    source_hash = dr.sha(source_bytes)
    source = dr.read_package_bytes(source_bytes)
    incoming = dr.unresolved_revisions(source)
    if incoming:
        print("STOP: unresolved tracked changes in %s; accept or reject them and restart from indexing. Nothing published." % sorted({r['part'] for r in incoming}))
        return 3
    doc = Document(BytesIO(source_bytes))
    paras = load_para_index(lp(args.para_index)) if args.para_index else {}
    entries = []
    for path in editpaths:
        parsed = parse_edit_file(path)
        print("parsed %d entries from %s" % (len(parsed), path.name))
        entries.extend(parsed)
    if not entries:
        raise ValueError("no edit entries parsed")
    ok, issues = validate_importance(entries)
    if not ok:
        raise ValueError("every COMMENT needs IMPORTANCE MAJOR|MINOR: " + "; ".join(issues))
    seen = set()
    for e in entries:
        if e["ACTION"] not in ACTIONS:
            raise ValueError("unknown ACTION for " + e["id"])
        if e["ACTION"] == "REVISE":
            if e["id"] in seen:
                raise ValueError(e["id"] + " has two REVISE entries; merge them")
            seen.add(e["id"])
            if not e.get("REVISED"):
                raise ValueError(e["id"] + " has no complete REVISED text")
    overrides = json.loads(lp(args.overrides).read_text(encoding="utf-8"))
    if not isinstance(overrides, dict):
        raise ValueError("overrides must be a JSON object")
    for pid, ov in overrides.items():
        if not isinstance(ov, dict) or ov.get("class", "reversal") not in OVERRIDE_CLASSES:
            raise ValueError("invalid override class for " + pid)
        cls, notes = ov.get("class", "reversal"), ov.get("notes", [])
        if cls == "partial" and not (ov.get("revised") or "").strip():
            raise ValueError("partial override needs the reconstructed text for " + pid)
        if cls == "drift" and not notes:
            raise ValueError("drift override needs notes for " + pid)
        if not isinstance(notes, list) or any(not isinstance(n, dict) or not (n.get("anchor") or "").strip() or not (n.get("note") or "").strip() for n in notes):
            raise ValueError("every drift note needs anchor and note for " + pid)
    timestamp = datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")
    rules = Path(__file__).resolve().parent.parent / "references"
    plan = {"schema_version": 1, "source_sha256": source_hash, "source_path": str(src),
            "rule_file_sha256": {p.name: dr.sha(p.read_bytes()) for p in sorted(rules.glob("rules_*.txt"))},
            "revision_author": args.author, "revision_date_utc": timestamp,
            "paragraphs": [], "comments": [], "runtime": runtime}
    manifest = {"schema_version": 1, "status": "staging", "input": src.name,
                "source_sha256": source_hash, "output": tracked.name, "internal_clean": out.name,
                "date": timestamp, "rule_set": args.rule_set or find_rule_set(), "author": args.author,
                "revisions": [], "comments": [], "refusals": [], "skipped": [], "notes": [], "drift": [], "runtime": runtime}
    root = dr.xml(source[dr.DOCUMENT])
    raw_paras = dr.body_paragraphs(root)
    rows, entry_locs = {}, {}
    proposed = sum(e["ACTION"] == "REVISE" for e in entries)
    partials = 0
    for n, e in enumerate(entries):
        i, p, how = find_paragraph(doc, e, paras, du.paragraph_view_text)
        if p is None:
            # Ambiguous/missing manuscript identity is not an unsupported-feature skip.
            raise dr.RevisionError(e["id"] + ": " + how)
        entry_locs[n] = i
        original = du.paragraph_view_text(raw_paras[i])
        if i not in rows:
            rows[i] = {"pid": e["id"], "part": dr.DOCUMENT, "body_index": i,
                       "source_paragraph_sha256": dr.fingerprint(raw_paras[i]), "original_text": original,
                       "approved_text": original, "status": "kept", "matched_by": how,
                       "index_text": paras.get(e["id"]), "original_preview": e.get("ORIGINAL", "")}
        row = rows[i]
        if e["ACTION"] != "REVISE":
            continue
        if row["status"] != "kept":
            raise dr.RevisionError("multiple edit IDs resolve to the same source paragraph")
        ov = overrides.get(e["id"], {})
        cls = ov.get("class", "reversal") if ov else None
        proposed_text = e["REVISED"]
        if cls == "reversal":
            row["status"] = "reverted"
            row["proposed_text"] = proposed_text
            manifest["refusals"].append({"class": "reversal", "pid": e["id"], "test": ov.get("test", "blind meaning check"),
                                         "reason": ov.get("reason", ""), "original_span": original,
                                         "proposed_replacement": proposed_text, "verdict": "REVISE -> KEEP"})
            continue
        approved = ov["revised"] if cls == "partial" else proposed_text
        # Existing typography convention; protected text must still match exactly.
        approved = du.smartify_quotes(approved)
        row["proposed_text"] = proposed_text
        if cls == "partial" and approved != du.smartify_quotes(proposed_text):
            partials += 1
            manifest["refusals"].append({"class": "partial", "pid": e["id"], "test": ov.get("test", "blind meaning check"),
                                         "reason": ov.get("reason", ""), "original_span": original,
                                         "proposed_replacement": proposed_text, "applied_replacement": approved,
                                         "verdict": "REVISE -> PARTIAL"})
            print("PARTIAL applied: " + e["id"])
        elif cls == "partial":
            manifest["notes"].append({"pid": e["id"], "note": "partial override matches the proposed revision; nothing reversed"})
            print("WARN: partial override %s matches the proposed revision; nothing reversed" % e["id"])
        if approved.count("--") > original.count("--"):
            manifest["notes"].append({"pid": e["id"], "note": "REVISED adds '--', which is not converted to a dash"})
            print("WARN: %s REVISED contains '--'; it is not converted" % e["id"])
        try:
            dr.rewrite_paragraph(raw_paras[i], approved, dr.IDs(source), args.author, timestamp)
        except dr.RevisionError:
            raise
        except ValueError as exc:
            row["status"], row["skip_reason"] = "skipped", str(exc)
            manifest["skipped"].append({"pid": e["id"], "why": str(exc)})
            continue
        row["approved_text"] = approved
        row["status"] = "no_op" if approved == original else ("partial" if cls == "partial" else "applied")
        if approved != original:
            manifest["revisions"].append({"pid": e["id"], "doc_index": i, "matched_by": how,
                                          "labels": e.get("ISSUE_LABELS", ""), "original": original, "revised": approved})
            print("REVISE applied: " + e["id"])
    # Full plan precedes all Word writes; every resolved paragraph carries exact text.
    plan["paragraphs"] = [rows[i] for i in sorted(rows)]
    for row in plan["paragraphs"]:
        row["text_differences"] = [{"source_start": a, "source_end": b, "inserted_text": text}
                                   for a, b, text in dr.differences(row["original_text"], row["approved_text"])]
    for pid in sorted(set(overrides) - seen):
        manifest["notes"].append({"pid": pid, "note": "override for a paragraph with no REVISE entry; nothing reversed"})
        print("WARN: override %s matches no REVISE entry; nothing reversed" % pid)
    q = 1
    seen_comments = set()
    def comment(e, i, text, anchor, kind, importance):
        nonlocal q
        key = (i, text)
        if key in seen_comments:
            raise dr.AnchorError("duplicate retained comment; explicitly consolidate it in edit entries")
        seen_comments.add(key)
        final = rows[i]["approved_text"]
        start, end = dr.resolve_anchor(final, anchor)
        item = {"id": "Q%03d" % q, "pid": e["id"], "body_index": i, "type": kind,
                "text": text, "anchor_evidence": anchor, "anchor_text": final[start:end], "importance": importance}
        q += 1
        plan["comments"].append(item)
        manifest["comments"].append({"pid": e["id"], "action": e["ACTION"], "importance": importance, "text": text, "id": item["id"]})
    for n, e in enumerate(entries):
        i = entry_locs[n]
        if has_comment(e):
            text = re.sub(r"^\s*\[(?:MAJOR|MINOR|FLAG)\]\s*", "", e["COMMENT"])
            text = re.sub(r"\s*Author decides\.?\s*$", "", text)
            comment(e, i, tidy_comment(text), e.get("ANCHOR"), "author_query", e["IMPORTANCE"])
        ov = overrides.get(e["id"], {})
        if e["ACTION"] == "REVISE" and rows[i]["status"] in {"applied", "partial"} and rows[i]["approved_text"] != rows[i]["original_text"]:
            for note in ov.get("notes", []):
                comment(e, i, tidy_comment(DRIFT_PREFIX + note["note"].strip()), note["anchor"], "drift", "MINOR")
                manifest["drift"].append({"pid": e["id"], "anchor": note["anchor"], "note": note["note"]})
    for p in outputs:
        p.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".tracked-stage-", dir=str(tracked.parent)) as temp:
        stage = lp(temp)
        stage_plan = stage / "plan.json"
        stage_plan.write_bytes(json_bytes(plan))
        generated, clean, checks = dr.build_output(source, plan, source_hash)
        stage_tracked, stage_clean = stage / "tracked.docx", stage / "clean.docx"
        dr.write_package(stage_tracked, generated)
        dr.write_package(stage_clean, clean)
        # Reopen the actual serialized package for independent projection checks.
        dr.validate_result(source, dr.read_package(stage_tracked), plan)
        v = du.verify_docx_formatting(stage_clean, src)
        if v.get("errors"):
            raise dr.RevisionError("format/quotation verification failed: " + "; ".join(v["errors"]))
        checks["warnings"] = v.get("warnings", [])
        print("verify: errors=0 warnings=%d" % len(checks["warnings"]))
        checks["word_ui_review"] = "not performed by this program; see tracked_output.md"
        manifest["verify"] = checks
        full = sum(r["status"] == "reverted" for r in rows.values())
        no_op = sum(r["status"] == "no_op" for r in rows.values())
        manifest["counts"] = {"proposed": proposed, "applied": len(manifest["revisions"]), "reverted": full, "no_op": no_op,
                              "partial": partials, "revert_rate_pct": round(100 * (full + partials) / proposed, 1) if proposed else 0.0,
                              "skipped": len(manifest["skipped"]), "drift_notes": len(manifest["drift"]),
                              "comments_listed": len(plan["comments"]), "comments_injected": len(plan["comments"])}
        stage_plan.write_bytes(json_bytes(plan))
        if dr.sha(src.read_bytes()) != source_hash:
            raise dr.RevisionError("source hash changed during run; nothing published")
        manifest["outputs"] = {"tracked": {"path": str(tracked), "sha256": dr.sha(stage_tracked.read_bytes())},
                               "internal_clean": {"path": str(out), "sha256": dr.sha(stage_clean.read_bytes())},
                               "approved_plan": {"path": str(planpath), "sha256": dr.sha(stage_plan.read_bytes())}}
        manifest["status"] = "ready"
        published = []
        try:
            for staged, dest in ((stage_clean, out), (stage_tracked, tracked), (stage_plan, planpath)):
                # Track destination before exclusive creation, so a partial failed copy is removed.
                if dest.exists():
                    raise FileExistsError("output appeared during run: " + str(dest))
                publish_file(staged, dest)
                published.append(dest)
            if dr.sha(src.read_bytes()) != source_hash:
                raise dr.RevisionError("source hash changed during publication")
            for output in manifest["outputs"].values():
                if dr.sha(lp(output["path"]).read_bytes()) != output["sha256"]:
                    raise dr.RevisionError("published output hash mismatch: " + output["path"])
            stage_manifest = stage / "ready.json"
            stage_manifest.write_bytes(json_bytes(manifest))
            publish_file(stage_manifest, manifestpath)
        except Exception as exc:
            for dest in published:
                dest.unlink(missing_ok=True)
            # Failure record only; never a ready manifest for an incomplete set.
            failure = {"schema_version": 1, "status": "incomplete", "diagnostics": [str(exc)], "outputs_published": False}
            if not manifestpath.exists():
                with manifestpath.open("xb") as stream:
                    stream.write(json_bytes(failure))
            print("STOP: publication failed; output set incomplete: " + str(exc))
            return 6
    print("READY: %s; applied=%d skipped=%d comments=%d" % (tracked, len(manifest["revisions"]), len(manifest["skipped"]), len(plan["comments"])))
    print("summary: %.1f%% challenged" % manifest["counts"]["revert_rate_pct"])
    return 0

def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    for name in ("input", "out", "overrides"):
        ap.add_argument("--" + name, required=True)
    ap.add_argument("--tracked-out")
    ap.add_argument("--plan-out")
    ap.add_argument("--edits", required=True, nargs="+")
    ap.add_argument("--para-index")
    ap.add_argument("--author", default="Language Editor")
    ap.add_argument("--manifest")
    ap.add_argument("--rule-set")
    args = ap.parse_args(argv)
    try:
        return run(args)
    except Exception as exc:
        # Imports may fail before docx_revisions can be imported.
        print(json.dumps({"status": "failed", "diagnostics": [str(exc)]}, ensure_ascii=False), file=sys.stderr)
        if type(exc).__name__ == "AnchorError":
            return 4
        if type(exc).__name__ == "RevisionError":
            return 5
        return 1

if __name__ == "__main__":
    raise SystemExit(main())
