# Tracked Word output

The proofreading coordinator runs `scripts/apply_language_edits.py` after the
independent meaning review. The program uses ordinary Python and edits the DOCX
package directly. It needs no assistant-specific document API or Word automation.
Read `execution.md` for the complete agent workflow and isolated role sessions.

## Runtime and command

Use Python 3.12 or newer, `python-docx==1.2.0` and `lxml==6.1.1`, as declared in
`requirements.txt`. Python 3.12 is the minimum release checked by the program;
the implementation was tested with Python 3.12.14. The manifest records the
actual interpreter and package versions. A different runtime is not silently
accepted in place of the pinned packages.

The following is one command line. Substitute resolved paths, preserving quotes.
It works in PowerShell and POSIX shells with a suitable Python command:

```text
python "<skill>/scripts/apply_language_edits.py" --input "<run>/original.docx" --out "<run>/internal/edited_CLEAN.docx" --tracked-out "<run>/edited_TRACKED.docx" --plan-out "<run>/internal/approved_edit_plan.json" --edits "<run>/language_agent_methods.md" "<run>/language_agent_results.md" --para-index "<run>/para_index.txt" --overrides "<run>/verifier_overrides.json" --author "Language Editor" --manifest "<run>/language_edit_manifest.json"
```

| Argument | Meaning |
| --- | --- |
| `--input` | Untouched original DOCX with incoming revisions resolved |
| `--out` | Internal accepted/clean reference, not the author's primary deliverable |
| `--tracked-out` | Word document with actual insertions, deletions and retained comments |
| `--plan-out` | Complete source-identified approved edit plan |
| `--edits` | One or more edit files; quoted glob patterns are expanded by the program |
| `--para-index` | Source paragraph index; strongly recommended |
| `--overrides` | Meaning-check JSON; `{}` when no finding requires an override |
| `--author` | Revision and new-comment author |
| `--manifest` | JSON status, checks, output hashes, skips and reversal records |
| `--rule-set` | Optional audit name; default is the bundled language rules filename |

Use fresh, distinct paths for every run. No output may overwrite an input,
another output, or an existing deliverable. For older invocations that omit the
two new path arguments, the program derives `<out-stem>_TRACKED.docx` (removing
`_CLEAN` from the stem) and `approved_edit_plan.json` beside `--out`. New hosts
should supply all paths explicitly. The default manifest is
`language_edit_manifest.json` beside `--out`.

When review finds no edits or comments, supply one `ACTION: KEEP` entry with an
unchanged source paragraph and `COMMENT: null`, plus `{}` overrides. Empty parsed
edit files are treated as missing workflow output, not as a completed review.

If invoking from Python, use an argument list with `subprocess.run`, not a shell
command assembled from manuscript text. Wait for completion and read both the
exit status and JSON manifest. The program supplies file processing; it does
not perform the AI meaning review or provide independent model sessions.

## Approved plan

`schema_version: 1` contains the untouched source SHA-256, source path, rule-file
hashes, revision author, UTC date and runtime. Each resolved paragraph has its
paragraph ID, `word/document.xml` part, zero-based body-paragraph index, source
XML fingerprint, full `original_text`, full `approved_text` and a separate
status: `kept`, `no_op`, `applied`, `partial`, `reverted` or `skipped`. Index text
and agent previews remain separate from actual source text. No reconstruction
or comparison relies on shortened audit excerpts.
`text_differences` records exact source character offsets and inserted text for
the approved word-level differences; the writer computes its own differences
from the complete text rather than accepting agent-supplied XML.

Full reversals keep source text and generate no revision. Partial reversals use
the complete reconstructed paragraph supplied by the meaning checker. Existing
straight-to-curly quote conversion is applied before approved text is frozen.
Significant source spaces are preserved in XML; REVISED values retain leading
and trailing spaces. Tabs and manual line breaks retain the existing edit lock.

Each retained issue has a stable `Q###` ID, paragraph location, full comment,
type, importance, supplied anchor evidence and final accepted sentence. After
placement, its Word comment ID and sentence offsets are recorded. Duplicate
comments at one location are rejected; consolidate related questions in the
edit entries before invoking the program. It never guesses that one issue
replaces another.

## Revisions and preservation

The writer starts with the original package. Word-level differences use exact
character offsets. Equal text keeps its source run properties. Deleted text
becomes `w:del/w:r/w:delText`; inserted text becomes `w:ins/w:r/w:t`.
Every revision has a unique ID, author and UTC date. Significant whitespace uses
`xml:space="preserve"`. The clean reference is produced by accepting these
generated revisions, so both files share the same approved plan.

New wording inherits the replaced span's complete run properties when they are
unambiguous. For a pure insertion, adjacent eligible source runs must agree.
Unchanged bold, italic, font, size, colour and other run properties remain in
place. A replacement crossing different formatting is skipped when its new
wording cannot be assigned one unambiguous source style. The original paragraph
remains intact and the manifest explains the skip.

The supported editing scope remains top-level body paragraphs. Tables, headers,
footers, notes, images and other package parts are preserved. Text safely outside
hyperlinks, complex citation fields and note references can be edited. Protected
objects retain distinct identities and order even when their displayed text or
numeric note IDs repeat. An edit that changes or crosses protected content,
bookmarks or existing comment boundaries is skipped. Existing comment text,
IDs, relationships and range boundaries remain intact; this implementation does
not silently widen an existing comment to make an edit fit.

Existing locks remain: content controls, simple fields, equations, symbol-font
characters, superscript/subscript text, tabs/manual breaks, inline pictures and
unsupported or incompatible visible formatting. Supported floating objects
remain protected in place. The implementation may conservatively skip a change
near a protected boundary instead of attempting an uncertain mapping. Skips are
per-edit outcomes, not evidence that the whole run failed.

Incoming text, move, formatting, table, section and numbering revisions are
checked in all Word XML parts, including headers, footers and notes. Resolve
them in Word and restart from indexing before creating a new revision set.

## Comments

Supply `COMMENT`, `IMPORTANCE: MAJOR|MINOR` and a sentence `ANCHOR` for every
retained manuscript question, including `ACTION: KEEP` entries. Related questions
may share one comment when all distinct issues remain explicit.

The program locates the anchor in the final accepted text, including insertions
and excluding deletions. It maps the sentence to safe XML boundaries, splitting
plain runs or insertion wrappers where needed and assigning new IDs to split
revision fragments. Markers are never placed inside a hyperlink or field.

A missing, repeated, removed or unsafe anchor is a fatal placement error. Read
the final passage, choose a unique relevant sentence anchor and rerun with fresh
paths. The program does not turn a failed sentence anchor into a whole-paragraph
comment. A meaning-drift comment is added only when its edit remains applied.
New comments remain annotations in both the accepted and rejected views.

## Validation and publication

The program saves the complete plan in staging before writing Word output. It
then reopens the generated package and checks:

- XML parsing, relationship targets, content types, revision metadata and IDs.
- Accepted paragraph text against the full approved plan.
- Rejected text, source run properties, original anchors and protected objects
  against source XML, allowing equivalent splitting of text runs.
- Inserted formatting against its unambiguous source donor, independently of
  the writer's difference calculation.
- Unedited paragraphs, tables and section content; hashes of unchanged parts.
- Original comments and relationships, and exact once-only placement of every
  retained new comment on its final sentence.
- Existing formatting and quotation checks on the clean reference.

The source hash is rechecked before and during publication. Output files are
created exclusively at fresh paths, their hashes are checked, and the
`status: "ready"` manifest is written last. If publication fails, the program
removes files it created and records `status: "incomplete"` when possible.
Failure before publication leaves no deliverable. An existing output collision
does not overwrite the existing file. A missing ready manifest means the output
set is not ready to deliver.

| Exit | Meaning and next action |
| --- | --- |
| `0` | Validated outputs and ready manifest; inspect recorded skips and warnings |
| `1` | Dependency, preflight or input error; correct the named cause |
| `2` | Invalid command arguments |
| `3` | Incoming tracked revisions; resolve them and restart from indexing |
| `4` | Retained comment cannot be placed; repair its final-text anchor |
| `5` | Invalid plan or failed content/format/package preservation; do not deliver |
| `6` | Publication failed; treat the attempted output set as incomplete |

Fatal diagnostics are printed as JSON on stderr (incoming-revision stops print
the affected parts on stdout). A preflight or staging failure may have no
manifest because no output was published. Never infer readiness from a filename
or successful file creation alone.

## Tests and Word review

Run `python "<skill>/scripts/test_tracked_output.py"` for the complete revision
path and the other bundled `test_*.py` files for regression checks. Fixtures are
created in temporary directories. The tracked tests use a separate XML traversal
to check accepted and rejected text, plus deliberate text/format/resource
corruption, bad anchors, invalid plans, dependency failures and output collisions.

Before release, open representative output in Word, inspect visible revisions
and comments, and exercise Accept All and Reject All on copies. Automated XML
checks do not prove interoperability with every Word version. Fresh-agent trials
must separately exercise the whole structure and proofreading workflows with
isolated reviews; the file writer cannot establish model suitability by itself.
