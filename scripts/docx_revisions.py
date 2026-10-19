"""Portable, source-based Word text revisions and fail-closed validation.

No Word automation or agent-specific API is used. See references/tracked_output.md.
"""
import sys
sys.dont_write_bytecode = True
import copy
import hashlib
import json
import posixpath
import re
import zipfile
from difflib import SequenceMatcher
from pathlib import Path
from urllib.parse import unquote
from io import BytesIO
from lxml import etree
import docx_utils as du

W = du.WML
NS = {"w": W}
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
DOCUMENT = "word/document.xml"
REVISION_NAMES = {"ins", "del", "moveFrom", "moveTo", "moveFromRangeStart",
                  "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd",
                  "rPrChange", "pPrChange", "sectPrChange", "tblPrChange",
                  "tblPrExChange", "trPrChange", "tcPrChange", "tblGridChange",
                  "cellIns", "cellDel", "cellMerge", "numberingChange"}
REVISION_NAMES |= {"customXml" + kind + "Range" + edge
                   for kind in ("Ins", "Del", "MoveFrom", "MoveTo", "ConflictIns", "ConflictDel")
                   for edge in ("Start", "End")}
REVISION_NAMES |= {"conflictIns", "conflictDel"}

class RevisionError(ValueError):
    """A fatal package, plan or preservation error."""

class AnchorError(RevisionError):
    """A retained comment cannot be placed safely."""

def w(name):
    return "{%s}%s" % (W, name)

def sha(data):
    return hashlib.sha256(data).hexdigest()

def xml(data):
    return etree.fromstring(data, etree.XMLParser(resolve_entities=False, no_network=True))

def serialize(el):
    return etree.tostring(el, encoding="UTF-8", xml_declaration=True, standalone=True)

def semantic(el):
    """Namespace-prefix-independent structural value; includes significant text."""
    if el is None:
        return None
    return [el.tag, sorted(el.attrib.items()), el.text or "", [semantic(c) for c in el]]

def fingerprint(el):
    return sha(json.dumps(semantic(el), ensure_ascii=False).encode("utf-8"))

def read_package(path):
    return read_package_bytes(Path(path).read_bytes())

def read_package_bytes(data):
    """Parse exactly the byte snapshot used for the source document hash."""
    with zipfile.ZipFile(BytesIO(data)) as z:
        if len(z.namelist()) != len(set(z.namelist())):
            raise RevisionError("duplicate ZIP member names")
        parts = {i.filename: z.read(i) for i in z.infolist()}
    validate_package(parts)
    return parts

def write_package(path, parts):
    with zipfile.ZipFile(str(path), "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)

def validate_package(parts):
    for required in (DOCUMENT, "[Content_Types].xml", "_rels/.rels"):
        if required not in parts:
            raise RevisionError("missing package part: " + required)
    roots = {}
    for name, data in parts.items():
        if name.endswith((".xml", ".rels")):
            try:
                roots[name] = xml(data)
            except etree.XMLSyntaxError as exc:
                raise RevisionError("invalid XML in %s: %s" % (name, exc)) from exc
    for name, root in roots.items():
        if not name.endswith(".rels"):
            continue
        ids = [c.get("Id") for c in root]
        if None in ids or len(ids) != len(set(ids)):
            raise RevisionError("duplicate or missing relationship ID in " + name)
        base = "" if name == "_rels/.rels" else posixpath.dirname(posixpath.dirname(name))
        for rel in root:
            if rel.get("TargetMode") == "External":
                continue
            target = unquote((rel.get("Target") or "").split("#", 1)[0])
            resolved = posixpath.normpath(posixpath.join(base, target)).lstrip("/")
            if resolved not in parts:
                raise RevisionError("missing relationship target %s in %s" % (resolved, name))
    reference_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    for name, root in roots.items():
        if not name.endswith(".xml"):
            continue
        rel_name = posixpath.join(posixpath.dirname(name), "_rels", posixpath.basename(name) + ".rels")
        rel_ids = {r.get("Id") for r in roots.get(rel_name, [])}
        for element in root.iter():
            for key, value in element.attrib.items():
                if key in {"{%s}%s" % (reference_ns, attr) for attr in ("id", "embed", "link")} and value not in rel_ids:
                    raise RevisionError("unresolved relationship %s in %s" % (value, name))
    ct = roots["[Content_Types].xml"]
    defaults = {c.get("Extension") for c in ct if c.tag == "{%s}Default" % CT}
    overrides = {c.get("PartName", "").lstrip("/") for c in ct if c.tag == "{%s}Override" % CT}
    if len(defaults) != sum(c.tag == "{%s}Default" % CT for c in ct) or len(overrides) != sum(c.tag == "{%s}Override" % CT for c in ct):
        raise RevisionError("duplicate content-type declarations")
    for name in parts:
        if name.endswith("/") or name == "[Content_Types].xml":
            continue
        if name not in overrides and name.rsplit(".", 1)[-1] not in defaults:
            raise RevisionError("no content type for " + name)
    return roots

def unresolved_revisions(parts):
    found = []
    for name, data in parts.items():
        if not name.startswith("word/") or not name.endswith(".xml"):
            continue
        root = xml(data)
        for el in root.iter():
            if isinstance(el.tag, str) and (el.tag.startswith("{" + W + "}")
                                           or "schemas.microsoft.com/office/word/" in el.tag):
                local = etree.QName(el).localname
                if local in REVISION_NAMES or local.endswith("PrChange"):
                    found.append({"part": name, "tag": local, "id": el.get(w("id"))})
    return found

def body_paragraphs(root):
    body = root.find(w("body"))
    if body is None:
        raise RevisionError("Word body not found")
    return body.findall(w("p"))

def project_element(el, accept):
    """Materialize a Word text-revision view, independently of the writer."""
    result = copy.deepcopy(el)
    for change in list(result.iter(w("ins"), w("del")))[::-1]:
        parent = change.getparent()
        if parent is None:
            continue
        keep = (change.tag == w("ins")) == accept
        i = parent.index(change)
        if keep:
            for child in list(change):
                for node in child.iter(w("delText")):
                    node.tag = w("t")
                parent.insert(i, child)
                i += 1
        parent.remove(change)
    return result

def paragraph_text(el, accept=True):
    return du.paragraph_view_text(project_element(el, accept))

class IDs:
    def __init__(self, parts):
        ids = []
        for name, data in parts.items():
            if name.startswith("word/") and name.endswith(".xml"):
                ids.extend(int(el.get(w("id"))) for el in xml(data).iter()
                           if (el.get(w("id")) or "").isdigit())
        self.value = max(ids, default=-1) + 1

    def next(self):
        result = str(self.value)
        self.value += 1
        return result

def run_like(source, text, deleted=False):
    result = etree.Element(w("r"), attrib=dict(source.attrib))
    props = source.find(w("rPr"))
    if props is not None:
        result.append(copy.deepcopy(props))
    t = etree.SubElement(result, w("delText") if deleted else w("t"))
    t.text = text
    t.set(XML_SPACE, "preserve")
    return result

def run_style(run):
    return json.dumps([sorted(run.attrib.items()), semantic(run.find(w("rPr")))], sort_keys=True)

def text_parts(p):
    """Ordered source objects, with distinct identities even for repeated fields."""
    children = list(p)
    fields = {begin: end for begin, end in du._field_spans(p)}
    result, i, pos = [], 0, 0
    while i < len(children):
        child = children[i]
        if i in fields:
            group = children[i:fields[i] + 1]
            text = du._field_result_text(group)
            editable = False
            i = fields[i] + 1
        else:
            group = [child]
            if child.tag == w("r"):
                text = du._run_view_text(child)
            elif child.tag == w("hyperlink"):
                text = "".join(du._run_view_text(r) for r in child.findall(w("r")))
            else:
                text = ""
            editable = (child.tag == w("r") and bool(text)
                        and all(c.tag in {w("rPr"), w("t")} for c in child))
            i += 1
        result.append({"start": pos, "end": pos + len(text), "text": text,
                       "editable": editable, "children": group,
                       "object_id": "object-%d" % len(result)})
        pos += len(text)
    return result

def differences(old, new):
    pattern = re.compile(r"\s+|\w+|[^\w\s]", re.UNICODE)
    a, b = list(pattern.finditer(old)), list(pattern.finditer(new))
    ao = [m.start() for m in a] + [len(old)]
    bo = [m.start() for m in b] + [len(new)]
    return [(ao[i], ao[j], new[bo[k]:bo[l]])
            for tag, i, j, k, l in SequenceMatcher(None, [m.group() for m in a],
                                                   [m.group() for m in b], autojunk=False).get_opcodes()
            if tag != "equal"]

def rewrite_paragraph(p, approved, ids, author, timestamp):
    """All-or-nothing source-run editing. ValueError is a recorded per-edit skip."""
    old = du.paragraph_view_text(p)
    if old == approved:
        return copy.deepcopy(p)
    reason = du.paragraph_lock_reason(p)
    if reason:
        raise ValueError("paragraph not rewritten: it holds " + reason)
    if re.findall(r"\[\^[^\]\s]*\]", old) != re.findall(r"\[\^[^\]\s]*\]", approved):
        raise ValueError("footnote or endnote marks were added, removed, moved or renumbered")
    parts = text_parts(p)
    if "".join(part["text"] for part in parts) != old:
        raise ValueError("unsupported paragraph text layout")
    edits = differences(old, approved)
    insertion_templates = {}
    for a, b, new in edits:
        for part in parts:
            if part["editable"] or part["children"][0].tag == w("pPr"):
                continue
            s, e = part["start"], part["end"]
            if (s < b and e > a) or (s == e and a <= s <= b):
                raise ValueError("edit crosses protected %s boundary" % part["object_id"])
        if not new:
            continue
        runs = [part["children"][0] for part in parts if part["editable"]
                and ((part["start"] < b and part["end"] > a) if a < b
                     else part["start"] <= a <= part["end"])]
        if not runs or len({run_style(r) for r in runs}) != 1:
            raise ValueError("inserted wording has ambiguous mixed formatting or no source run")
        insertion_templates[a] = runs[0]
    result = copy.deepcopy(p)
    for child in list(result):
        result.remove(child)
    placed = set()
    def emit_insert(a, new, template):
        if a in placed or not new:
            return
        ins = etree.SubElement(result, w("ins"), {w("id"): ids.next(), w("author"): author, w("date"): timestamp})
        ins.append(run_like(insertion_templates[a], new))
        placed.add(a)
    for part in parts:
        if not part["editable"]:
            for child in part["children"]:
                result.append(copy.deepcopy(child))
            continue
        s, e, run = part["start"], part["end"], part["children"][0]
        relevant = [(a, b, new) for a, b, new in edits if (a < e and b > s) or s <= a <= e]
        if not relevant:
            result.append(copy.deepcopy(run))
            continue
        cuts = sorted({s, e} | {max(s, min(e, x)) for a, b, new in relevant for x in (a, b)})
        for x, y in zip(cuts, cuts[1:]):
            for a, b, new in relevant:
                if a == x:
                    emit_insert(a, new, run)
            deleted = any(a <= x and y <= b for a, b, new in edits if a < b)
            fragment = run_like(run, old[x:y], deleted)
            if deleted:
                wrapper = etree.SubElement(result, w("del"), {w("id"): ids.next(), w("author"): author, w("date"): timestamp})
                wrapper.append(fragment)
            else:
                result.append(fragment)
        for a, b, new in relevant:
            if a == e:
                emit_insert(a, new, run)
    if paragraph_text(result, True) != approved or paragraph_text(result, False) != old:
        raise RevisionError("writer text projection disagrees with the complete edit plan")
    return result

def resolve_anchor(text, anchor):
    if not anchor or not anchor.strip():
        raise AnchorError("every retained comment needs an ANCHOR in final accepted text")
    hit, count = du._find_phrase(text, anchor)
    if hit is None or count != 1:
        raise AnchorError("ANCHOR is missing or repeated in final accepted text: " + anchor)
    spans = [(s, e) for s, e in du.sentence_spans(text) if s < hit[1] and e > hit[0]]
    if len(spans) != 1:
        raise AnchorError("ANCHOR must identify one sentence: " + anchor)
    return spans[0]

def child_text(child):
    p = etree.Element(w("p"))
    p.append(copy.deepcopy(child))
    return paragraph_text(p)

def split_visible_child(child, offset, ids):
    """Split plain text or insertion wrappers only; never enter protected XML."""
    if child.tag == w("r") and all(c.tag in {w("rPr"), w("t")} for c in child):
        text = du.get_run_text(child)
        return run_like(child, text[:offset]), run_like(child, text[offset:])
    if child.tag == w("ins"):
        left, right = copy.deepcopy(child), copy.deepcopy(child)
        left[:] = []
        right[:] = []
        left.set(w("id"), ids.next())
        right.set(w("id"), ids.next())
        pos = 0
        for run in child:
            n = len(du.get_run_text(run))
            if pos + n <= offset:
                left.append(copy.deepcopy(run))
            elif pos >= offset:
                right.append(copy.deepcopy(run))
            else:
                a, b = split_visible_child(run, offset - pos, ids)
                left.append(a)
                right.append(b)
            pos += n
        return left, right
    raise AnchorError("sentence boundary falls inside a protected field, link or object")

def boundary_index(p, offset, ids):
    pos = 0
    field_depth = 0
    for i, child in enumerate(list(p)):
        if child.tag == w("pPr"):
            continue
        if pos == offset and field_depth == 0:
            return i
        n = len(child_text(child))
        if pos < offset < pos + n:
            if field_depth:
                raise AnchorError("sentence boundary falls inside a complex field")
            left, right = split_visible_child(child, offset - pos, ids)
            p.remove(child)
            p.insert(i, right)
            p.insert(i, left)
            return i + 1
        for fc in child.iter(w("fldChar")):
            if fc.get(w("fldCharType")) == "begin":
                field_depth += 1
            elif fc.get(w("fldCharType")) == "end":
                field_depth -= 1
        pos += n
    if offset == pos and field_depth == 0:
        return len(p)
    raise AnchorError("sentence boundary cannot be mapped to safe XML")

def place_comments(parts, root, comments, ids, author, timestamp):
    if not comments:
        return
    comments_path = "word/comments.xml"
    croot = xml(parts[comments_path]) if comments_path in parts else etree.Element(w("comments"), nsmap={"w": W})
    existing = [int(c.get(w("id"))) for c in croot if (c.get(w("id")) or "").isdigit()]
    next_comment = max(existing, default=-1) + 1
    paras = body_paragraphs(root)
    for item in comments:
        p = paras[item["body_index"]]
        text = paragraph_text(p)
        start, end = resolve_anchor(text, item["anchor_text"])
        cid = str(next_comment)
        next_comment += 1
        finish = boundary_index(p, end, ids)
        marker = etree.Element(w("commentRangeEnd"), {w("id"): cid})
        ref = etree.Element(w("r"))
        etree.SubElement(ref, w("commentReference"), {w("id"): cid})
        p.insert(finish, marker)
        p.insert(finish + 1, ref)
        begin = boundary_index(p, start, ids)
        p.insert(begin, etree.Element(w("commentRangeStart"), {w("id"): cid}))
        c = etree.SubElement(croot, w("comment"), {w("id"): cid, w("author"): author, w("date"): timestamp})
        t = etree.SubElement(etree.SubElement(etree.SubElement(c, w("p")), w("r")), w("t"))
        t.text = item["text"]
        t.set(XML_SPACE, "preserve")
        item["word_comment_id"] = cid
        item["sentence_start"], item["sentence_end"] = start, end
        item["anchor_text"] = text[start:end]
    parts[comments_path] = serialize(croot)
    rel_path = "word/_rels/document.xml.rels"
    relroot = xml(parts[rel_path]) if rel_path in parts else etree.Element("{%s}Relationships" % REL, nsmap={None: REL})
    comment_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    links = [r for r in relroot if r.get("Type") == comment_rel]
    if links and any(r.get("Target") != "comments.xml" for r in links):
        raise RevisionError("unsupported nonstandard comments relationship")
    if not links:
        used = {r.get("Id") for r in relroot}
        n = 1
        while "rId%d" % n in used:
            n += 1
        etree.SubElement(relroot, "{%s}Relationship" % REL,
                         Id="rId%d" % n, Type=comment_rel, Target="comments.xml")
    parts[rel_path] = serialize(relroot)
    ct = xml(parts["[Content_Types].xml"])
    if not any(c.get("PartName") == "/word/comments.xml" for c in ct):
        etree.SubElement(ct, "{%s}Override" % CT, PartName="/word/comments.xml",
                         ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")
    parts["[Content_Types].xml"] = serialize(ct)

def strip_new_comments(root, new_ids):
    result = copy.deepcopy(root)
    for el in list(result.iter(w("commentRangeStart"), w("commentRangeEnd"), w("commentReference"))):
        if el.get(w("id")) not in new_ids:
            continue
        parent = el.getparent()
        parent.remove(el)
        if parent.tag == w("r") and all(c.tag == w("rPr") for c in parent):
            parent.getparent().remove(parent)
    return result

def paragraph_signature(p):
    """Semantic XML preserving every source character's complete run properties.

    Adjacent equal-property text runs may be split or coalesced. Everything else
    (objects, fields, anchors, attributes, paragraph properties) must be equal.
    """
    events = []
    for child in p:
        if child.tag == w("r") and all(c.tag in {w("rPr"), w("t")} for c in child):
            style = run_style(child)
            text = du.get_run_text(child)
            if not text:
                events.append(["xml", semantic(child)])
            elif events and events[-1][0:2] == ["text", style]:
                events[-1][2] += text
            else:
                events.append(["text", style, text])
        else:
            events.append(["xml", semantic(child)])
    return [sorted(p.attrib.items()), events]

def validate_comments(parts, plan):
    root = xml(parts[DOCUMENT])
    croot = xml(parts["word/comments.xml"]) if "word/comments.xml" in parts else etree.Element(w("comments"))
    bodies = {c.get(w("id")): c for c in croot}
    if len(bodies) != len(croot):
        raise RevisionError("duplicate comment IDs")
    for tag in ("commentRangeStart", "commentRangeEnd", "commentReference"):
        seen = [e.get(w("id")) for e in root.iter(w(tag))]
        if len(seen) != len(set(seen)) or any(cid not in bodies for cid in seen):
            raise RevisionError("invalid " + tag + " coverage")
    for item in plan["comments"]:
        cid = item["word_comment_id"]
        p = project_element(body_paragraphs(root)[item["body_index"]], True)
        active, text, starts, ends, refs = False, "", 0, 0, 0
        for child in p:
            if child.tag == w("commentRangeStart") and child.get(w("id")) == cid:
                active, starts = True, starts + 1
            elif child.tag == w("commentRangeEnd") and child.get(w("id")) == cid:
                active, ends = False, ends + 1
            elif active:
                text += child_text(child)
            refs += len([e for e in child.iter(w("commentReference")) if e.get(w("id")) == cid])
        body_text = "".join(t.text or "" for t in bodies[cid].iter(w("t")))
        if (starts, ends, refs) != (1, 1, 1) or text != item["anchor_text"] or body_text != item["text"]:
            raise RevisionError("comment placement or text mismatch for " + item["id"])

def validate_inserted_styles(original, revised):
    """Check donor formatting from source XML without using writer differences."""
    source_runs = []
    pos = 0
    for part in text_parts(original):
        if part["editable"]:
            source_runs.append((part["start"], part["end"], run_style(part["children"][0])))
    children = list(revised)
    pos = 0
    for n, child in enumerate(children):
        if child.tag == w("ins"):
            # A replacement is insertion fragment(s) followed by deleted source
            # fragments. Annotation-induced insertion splits consume no source.
            end = n + 1
            while end < len(children) and children[end].tag == w("ins"):
                end += 1
            removed = 0
            while end < len(children) and children[end].tag == w("del"):
                removed += sum(len(t.text or "") for t in children[end].iter(w("delText")))
                end += 1
            if removed:
                donors = {style for s, e, style in source_runs if s < pos + removed and e > pos}
            else:
                donors = {style for s, e, style in source_runs if s <= pos <= e}
            styles = {run_style(run) for run in child}
            if len(donors) != 1 or styles != donors:
                raise RevisionError("inserted formatting does not match its unambiguous source span")
        elif child.tag == w("del"):
            pos += sum(len(t.text or "") for t in child.iter(w("delText")))
        else:
            pos += len(child_text(child))

def validate_result(source, generated, plan):
    validate_package(generated)
    original_root, final_root = xml(source[DOCUMENT]), xml(generated[DOCUMENT])
    revs = list(final_root.iter(w("ins"), w("del")))
    rev_ids = [e.get(w("id")) for e in revs]
    if None in rev_ids or len(rev_ids) != len(set(rev_ids)):
        raise RevisionError("revision IDs must be unique")
    for rev in revs:
        if rev.get(w("author")) != plan["revision_author"] or rev.get(w("date")) != plan["revision_date_utc"]:
            raise RevisionError("revision metadata differs from plan")
        if any(c.tag != w("r") for c in rev):
            raise RevisionError("unsupported revision content")
    new_ids = {c["word_comment_id"] for c in plan["comments"]}
    original_comment_ids = ({c.get(w("id")) for c in xml(source["word/comments.xml"])}
                            if "word/comments.xml" in source else set())
    generated_comment_ids = ({c.get(w("id")) for c in xml(generated["word/comments.xml"])}
                             if "word/comments.xml" in generated else set())
    if len(new_ids) != len(plan["comments"]) or new_ids & original_comment_ids:
        raise RevisionError("new comment IDs must be distinct from one another and the source")
    if generated_comment_ids != original_comment_ids | new_ids:
        raise RevisionError("unplanned or missing comment body")
    comparable = strip_new_comments(final_root, new_ids)
    accepted, rejected = project_element(comparable, True), project_element(comparable, False)
    old_paras, acc_paras, rej_paras = map(body_paragraphs, (original_root, accepted, rejected))
    if not len(old_paras) == len(acc_paras) == len(rej_paras):
        raise RevisionError("body paragraph count changed")
    records = {r["body_index"]: r for r in plan["paragraphs"]}
    if len(records) != len(plan["paragraphs"]):
        raise RevisionError("overlapping plan targets")
    for i, (old, acc, rej) in enumerate(zip(old_paras, acc_paras, rej_paras)):
        validate_inserted_styles(old, body_paragraphs(comparable)[i])
        if paragraph_signature(old) != paragraph_signature(rej):
            raise RevisionError("rejected view fails original text/format/object recovery at P%03d" % i)
        if i in records:
            row = records[i]
            if row["source_paragraph_sha256"] != fingerprint(old) or row["original_text"] != du.paragraph_view_text(old):
                raise RevisionError("source paragraph fingerprint mismatch")
            if du.paragraph_view_text(acc) != row["approved_text"]:
                raise RevisionError("accepted text differs from complete plan at " + row["pid"])
        elif paragraph_signature(old) != paragraph_signature(acc):
            raise RevisionError("unedited paragraph changed")
    # Check all non-paragraph body objects, tables, attributes and section layout.
    for root in (original_root, accepted, rejected):
        body = root.find(w("body"))
        for i, p in enumerate(body.findall(w("p"))):
            # Keep each paragraph's position relative to tables and other body
            # objects; its text and properties were compared independently above.
            body.replace(p, etree.Element("paragraphPosition", index=str(i)))
    if semantic(original_root) != semantic(accepted) or semantic(original_root) != semantic(rejected):
        raise RevisionError("non-body-paragraph content or section structure changed")
    allowed = {DOCUMENT}
    if plan["comments"]:
        allowed |= {"word/comments.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"}
    for name, data in source.items():
        if name not in allowed and generated.get(name) != data:
            raise RevisionError("unchanged package part differs: " + name)
    if set(generated) - set(source) - allowed:
        raise RevisionError("unexpected new package part")
    for name in ("word/_rels/document.xml.rels", "[Content_Types].xml"):
        if name in allowed:
            old = xml(source[name]) if name in source else None
            new = xml(generated[name])
            before = [semantic(c) for c in old] if old is not None else []
            after = [semantic(c) for c in new]
            if any(item not in after for item in before):
                raise RevisionError("existing relationship or content type was changed")
    if "word/comments.xml" in source:
        orig = xml(source["word/comments.xml"])
        current = xml(generated["word/comments.xml"])
        current_by_id = {c.get(w("id")): c for c in current}
        if any(semantic(c) != semantic(current_by_id.get(c.get(w("id")))) for c in orig):
            raise RevisionError("existing comment body changed")
    validate_comments(generated, plan)
    return {"errors": [], "warnings": [], "accepted_text": "passed", "rejected_text_and_format": "passed",
            "protected_parts": "passed", "comments": "passed", "revision_count": len(revs),
            "unchanged_part_sha256": {n: sha(d) for n, d in source.items() if n not in allowed}}

def build_output(source, plan, source_sha256):
    if (plan.get("schema_version") != 1 or not isinstance(plan.get("paragraphs"), list)
            or not isinstance(plan.get("comments"), list)
            or not {"source_sha256", "revision_author", "revision_date_utc"} <= set(plan)):
        raise RevisionError("invalid or incomplete edit plan")
    if plan["source_sha256"] != source_sha256:
        raise RevisionError("source hash differs from approved edit plan")
    result = dict(source)
    root = xml(source[DOCUMENT])
    paras = body_paragraphs(root)
    ids = IDs(source)
    seen = set()
    for row in plan["paragraphs"]:
        required = {"pid", "part", "body_index", "original_text", "approved_text", "status", "source_paragraph_sha256"}
        if not required <= set(row) or row["part"] != DOCUMENT:
            raise RevisionError("incomplete or unsupported paragraph plan")
        i = row["body_index"]
        if not isinstance(i, int) or not 0 <= i < len(paras) or i in seen:
            raise RevisionError("invalid or overlapping paragraph location")
        seen.add(i)
        if fingerprint(paras[i]) != row["source_paragraph_sha256"] or du.paragraph_view_text(paras[i]) != row["original_text"]:
            raise RevisionError("plan does not identify the original paragraph")
        revised = rewrite_paragraph(paras[i], row["approved_text"], ids, plan["revision_author"], plan["revision_date_utc"])
        paras[i].getparent().replace(paras[i], revised)
    place_comments(result, root, plan["comments"], ids, plan["revision_author"], plan["revision_date_utc"])
    result[DOCUMENT] = serialize(root)
    verify = validate_result(source, result, plan)
    clean = dict(result)
    clean[DOCUMENT] = serialize(project_element(root, True))
    validate_package(clean)
    return result, clean, verify
