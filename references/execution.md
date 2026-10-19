# Running the skill in an AI environment

Read `SKILL.md` first, resolve this folder as `[SKILL_DIR]`, and choose one
function. This file supplies the shared execution contract. The workflow
files contain the full role instructions and required output formats.

## Capabilities and startup

The AI assistant is the session the author works in. It runs programs,
starts agents and coordinates the final output. An agent is a separate AI
session receiving one task, its necessary files and its own instructions.

Before starting, check these capabilities with the tools actually available:

| Operation | Required capability | Check |
|---|---|---|
| Read instructions and manuscript | Local file access | Read `SKILL.md`, the chosen workflow, and the supplied DOCX |
| Save run records and outputs | File writing | Create the run folder next to the manuscript; preserve earlier runs |
| Run bundled programs | Python 3.12 or later, `python-docx==1.2.0`, `lxml==6.1.1` | Run the dependency command below and each relevant CLI's `--help` |
| Independent review | Separate sessions or isolated model calls | Give each role only its assigned inputs; do not inherit the editing conversation |
| Review the delivered document | Microsoft Word for the author | Open tracked output to review comments and accept or reject revisions |

Use the host's equivalents of file reading, file writing, program execution
and separate-session dispatch. Native skill discovery may load `SKILL.md`;
reading it directly is also a complete entry route. No assistant-specific
Word tool is required by the Python programs.

If a required capability is unavailable, identify it and ask for that setup.
Consecutive role instructions in the same continuing conversation do not
provide an independent meaning check. Do not report a complete run while a
required review or program check is missing.

Use the same selected capable model across the judgment roles. Schedule
independent roles in successive batches if concurrency is limited, then
wait for all required output files before a dependent stage. Record the
actual model identifier, tool environment, date and chosen function in the
run's `run_context.json`. The program manifest records runtime versions and
document validation; model availability alone does not establish suitability.

Dependency check, with the host's resolved Python interpreter:

```sh
python -c "import sys, docx, lxml; assert sys.version_info >= (3,12); print(sys.version); print(docx.__version__, lxml.__version__); assert docx.__version__ == '1.2.0' and lxml.__version__ == '6.1.1'"
python -m pip install -r "<skill>/requirements.txt"
```

Run the installation command if the declared packages are missing or their versions differ in
the selected environment; alternatively select an environment with those versions. `defusedxml` is optional and is used only by
`scripts/office/unpack.py`; the normal workflows do not need that tool.

## Role inputs, instructions and completion

| Role | Assigned inputs | Complete instructions | Required output |
|---|---|---|---|
| Structure coordinator | Manuscript index, section map, all candidate and advocate files | `workflows/structure_comments.md` Steps 0–6 | Audit table, `comments_final.json`, commented copy and report |
| Section commenters | Assigned section units, cross-check units, theme and rule packet | Step 2 prompt template and assignment table; `references/comment_contract.md` | One candidate file for each assigned role, including an explicit no-candidates result |
| Whole-paper commenter | Full index, theme, Parts 0–2 of the structural guideline | Step 2 STORY modification of the same template | STORY candidate file |
| Author's advocate | All candidates, index including notes, relevant rules and contract | Step 3 prompt template | Rebuttal with manuscript/rule evidence and DEFEATS, WEAKENS or FAILS for each candidate |
| Section editor | Assigned chunk, all language rules, relevant tables/notes, authorship, locked paragraphs | Editor prompt below; proofreading per-chunk workflow and Step 7 | Section edit file with all required records saved before returning |
| Meaning checker | Original and proposed paragraphs, authorship and verification standard only | Proofreading Step 8 and the prompt below | `meaning_verification.md` covering every edited paragraph |

For an editing session, fill in this prompt and give it only the listed files
and instructions. The coordinator resolves input paths and section boundaries.

```text
Edit the wording of [CHUNK NAME, P### RANGE] in this manuscript.
Read [SKILL_DIR]/references/rules_language_v13.txt in full and the
per-chunk workflow, Step 7 tests, and Required Outputs in
[SKILL_DIR]/workflows/proofread.md.
Inputs: [RUN]/para_index.txt, [RUN]/locked_paragraphs.txt,
[relevant tables.txt and notes.txt], authorship [sole/co-authored].
Use only the assigned paragraphs, with the whole-paper context needed for
the stated cross-checks. Preserve the author's claims and all protected
content. Perform all four passes; an applicable problem justifies an edit,
and a strong paragraph may remain unchanged. Run all three Step 7 tests on
each proposed change, with the prescribed span-level reversal records.
Write every retained question as COMMENT, IMPORTANCE and a unique ANCHOR.
Save [RUN]/language_agent_[CHUNK].md with Quick Diagnosis, edit entries,
Step 7 QC records, Edit Log, Verb Upgrade Log and Quick Checklist/Metrics.
If no paragraph needs an edit or comment, include one ACTION: KEEP entry
with COMMENT: null to record the completed review.
Do not write the Word output. Return the path and completion status only.
```

For the meaning checker, start a fresh session without the editing agent's
history or reasons. Prepare full original/revised pairs, not the editing log
or a 50-character ORIGINAL preview. Supply the Step 8 instructions separately
if the host cannot restrict access to a section of the workflow file.

```text
Independently check every ORIGINAL/REVISED paragraph pair in
[RUN]/meaning_check_input.json under the complete Step 8 verification
standard supplied here. Authorship: [sole/co-authored]. You have not been
given the editing reasons. Compare the words and citations, and check
scope, strength, attribution, mechanism, technical terms and stakes.
For each paragraph report PASS, MINOR DRIFT or FLAG FOR REVERSAL, with
the exact changed spans and reason. Follow Step 8's same-value-number and
minor-nuance rules. Save the complete report to
[RUN]/meaning_verification.md, then return its path.
```

The coordinator applies the verifier's decisions through
`verifier_reverts.json`, preserving full and partial reversal records and
drift anchors. It does not let the editor review its own edits in place of
this separate session.

## Commands and quoting

Use explicit paths. In Python, pass an argument list to `subprocess.run`
with `check=True`; do not build a shell command from manuscript text.
The examples below use placeholders to be replaced with actual paths.
The program expands the quoted edit-file pattern itself.

POSIX shell:

```sh
python "<skill>/scripts/index_manuscript.py" --input "<draft>.docx" --out-dir "<run>"
python "<skill>/scripts/apply_language_edits.py" \
  --input "<draft>.docx" \
  --out "<run>/internal/edited_CLEAN.docx" \
  --tracked-out "<run>/edited_TRACKED.docx" \
  --plan-out "<run>/internal/approved_edit_plan.json" \
  --edits "<run>/language_agent_*.md" \
  --para-index "<run>/para_index.txt" \
  --overrides "<run>/verifier_reverts.json" \
  --author "Language Editor" \
  --manifest "<run>/language_edit_manifest.json"
```

PowerShell uses an argument array so spaces and non-English paths stay intact:

```powershell
$skillPython = 'C:/path/to/python.exe'
$skillFolder = 'C:/path/to/技能文件夹'
$runFolder = 'C:/path/to/run folder'
$draftPath = 'C:/path/to/original draft.docx'
& $skillPython "$skillFolder/scripts/index_manuscript.py" --input $draftPath --out-dir $runFolder
if ($LASTEXITCODE -ne 0) { throw 'Indexing failed; read the diagnostic.' }
$applyArgs = @(
  "$skillFolder/scripts/apply_language_edits.py",
  '--input', $draftPath,
  '--out', "$runFolder/internal/edited_CLEAN.docx",
  '--tracked-out', "$runFolder/edited_TRACKED.docx",
  '--plan-out', "$runFolder/internal/approved_edit_plan.json",
  '--edits', "$runFolder/language_agent_*.md",
  '--para-index', "$runFolder/para_index.txt",
  '--overrides', "$runFolder/verifier_reverts.json",
  '--author', 'Language Editor',
  '--manifest', "$runFolder/language_edit_manifest.json"
)
& $skillPython @applyArgs
if ($LASTEXITCODE -ne 0) { throw 'Output is not ready; read the diagnostic.' }
Get-Content -LiteralPath "$runFolder/language_edit_manifest.json"
```

For structure comments, use the same interpreter with
`scripts/paragraph_index.py` and `scripts/add_comments.py`; their complete
commands and file formats are in the structure workflow. All paths resolve
from the supplied skill and run folders, not a developer's working directory.

## Completion and testing

A structure run completes only after all assigned commenting roles, the
applicable advocate review, the final audit and comment injection checks.
Zero retained comments still requires the verified unchanged copy and report.

A proofreading run completes only after the section files, three self-checks,
independent meaning check, overrides, approved plan and document validation.
Require exit status 0, a ready manifest, complete retained-comment coverage
and recorded output hashes. Read `tracked_output.md` for error meanings and
rerun rules. Deliver tracked output and existing audit records; the accepted
clean reference remains internal.

Run each `scripts/test_*.py` file with the selected interpreter. These tests
exercise document mechanics; they do not establish model judgment quality.
For a new environment, give a fresh agent the folder and a representative
test manuscript, run each function separately, and inspect the outputs and
records. Record the actual environment tested rather than assuming that
all agents or Word versions behave alike.
