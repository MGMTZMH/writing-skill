# Structure comments (workflow)

The skill's `SKILL.md` sends you here when the user chooses structure comments. Follow
this file exactly. `[SKILL_DIR]` is the skill's root folder, the one that holds `SKILL.md`.
Read `references/execution.md` for capability checks, tool mapping and role isolation.

You read a manuscript draft and leave Word comments about its STRUCTURE
and STORYTELLING: what is said where, in what order, what is missing,
what is opened and never closed, whether the story poses a puzzle and
resolves it. You do not touch the text. You do not write sentences for the
authors. You do not comment on wording.

Inputs, from the opening questions in `SKILL.md`: the draft (.docx; ask
for a .docx if given a PDF), optional `--max N` (a cap on the comments in
the run; by default there is none), and optional `--section X` to review only one
part of the paper: INTRODUCTION, THEORY, METHODS, RESULTS, DISCUSSION or
STORY (a request about one section maps to it).

## Scope

Quantitative empirical papers: archival, survey, experimental, single-
study or multi-study. If the draft is qualitative or a pure theory paper,
say so and stop; the rules were built for quantitative papers and would
be confidently wrong on the others. Wording-level diagnostics (bloated
sentences, buried subjects, stuffed sentences, jargon, cadence) are out
of scope and belong to a proofreading pass; the structural omissions
they point at (undefined constructs, unstated assumptions, unexplained
sample sizes, uninterpreted figures, unexplained analysis techniques) are
covered by B2, B3, C2.1, C2.3, C2.4, C3.3, S5 and D7. The draft's language is English. The draft is a .docx:
python-docx reads body and table paragraphs; text boxes get no ID of
their own (their words are folded into the host paragraph's text), so
anchors must come from body or table paragraphs.

## File resolution

`[SKILL_DIR]` is the skill's root folder (the folder that holds
`SKILL.md`). Everything the skill needs is bundled.

| Need | Path |
|---|---|
| Rules and eligibility classes (the authority) | `[SKILL_DIR]/references/guideline_structure_v4_5.txt` |
| Comment shape, schema, audit table, worked pairs | `[SKILL_DIR]/references/comment_contract.md` |
| Paragraph index | `python "[SKILL_DIR]/scripts/paragraph_index.py"` |
| Comment injection with checks | `python "[SKILL_DIR]/scripts/add_comments.py"` |
| docx utilities used by the scripts | `[SKILL_DIR]/scripts/docx_utils.py` (`inject_comments_ex`, `verify_docx_formatting`) |
| Optional XML inspection | `[SKILL_DIR]/scripts/office/unpack.py` (needs `defusedxml`) |

Outputs:
- Commented copy: same folder as the input, named `<stem>_structure_comments.docx`. Never overwrite the input. If a file of that name exists from an EARLIER run, use `_v2`, `_v3`; within the current run, if a comment is revised after injection, rerun the script with `--force` on the file this run created.
- Working files (index, section map, candidates, audit, final JSON) and the report: `paper_structure_output/` created next to the input. Report name: `structure_comments_<stem>.md`.

Use the Python version and pinned packages in `references/execution.md` and `requirements.txt`; run the shared dependency check before starting.

## Roles and model

| Step | Role | Execution |
|---|---|---|
| Section map, theme, audit, final prose, report | Coordinator | The selected capable session model |
| Five section commenters and one whole-paper commenter | Separate commenting sessions | The same selected model, assigned only their inputs below |
| Author’s advocate | Separate advocate session | The same selected model, candidates and manuscript/rule evidence |
| Indexing and comment placement | Bundled Python programs | The environment’s program-execution tool |

Reading, steelmanning and final decisions require the same judgment quality.
Do not substitute a less capable model for a commenting or advocate role.
Record the actual model identifier and tools used. Schedule roles within
available concurrency limits, using successive batches where necessary.

## Workflow

```
0  Resolve input -> paragraph index (script)
1  Section map by FUNCTION (orchestrator) + one-sentence theme
2  Six separate commenters, scheduled within capacity, each saving candidates
3  Author's advocate: strongest rebuttal per candidate
4  Audit table (orchestrator): keep / merge / drop; rank; write final prose
5  Inject into a copy (script); verify count and body text unchanged
6  Report + delivery message
```

Persist after every step (write-ahead, so a crash loses nothing): the index, the
section map, each agent's output file, the audit table, and the final
JSON are all on disk before the next step starts. On crash recovery,
check which files exist before re-dispatching anything.

### Step 0: Resolve input and index paragraphs

1. Take the .docx from the opening questions (else the most recent
   .docx mentioned in the conversation; else ask). Refuse PDFs and ask
   for the .docx.
2. Decide the output folder (File resolution above) and create it.
3. Run:
   ```bash
   python "[SKILL_DIR]/scripts/paragraph_index.py" "<input.docx>" --out "<outdir>/paragraph_index.json"
   ```
   This writes `paragraph_index.json` (list of `{id, text, style, is_heading, heading_hint, in_table, notes}`; `notes` holds the texts of the footnotes and endnotes the paragraph's marks point to) and a companion `paragraph_index_map.txt` (id, style, first 60 characters, ASCII-safe, and `[notes: ...]` on a paragraph that carries notes). Read the complete map with the environment’s file-reading tool.
   If python-docx raises `PackageNotFoundError` on a file under OneDrive, the file is probably an online-only placeholder: it reports its full size but has no content on disk, `zipfile.is_zipfile` returns False, and `attrib` shows the O flag. Ask the user to make it available offline (pin it), then rerun; do not try to repair the file.
   If the index has fewer than about 40 paragraphs or no hypothesis, method or result vocabulary, stop and ask whether this is the paper (a coding prompt or a memo indexes as a few dozen paragraphs and has none of that vocabulary).

### Step 1: Section map by function

Read `paragraph_index_map.txt` and, where a heading is ambiguous, the
full text of the paragraphs around it from the JSON. Assign every
paragraph to exactly one unit:

| Function | How to recognize it |
|---|---|
| FRONT | Title (first non-empty, often centered or Title style), abstract block (before the first body heading; may be labeled), keywords |
| INTRODUCTION | From the first body paragraph to the first heading that starts hypothesis development or literature review; paragraphs under that heading that still state the problem or the contribution stay here; heading-less introductions are common |
| THEORY | Hypothesis development and its literature review, wherever it sits; includes theory paragraphs that open a later study (multi-study) |
| METHODS(k) | "Method(s)", "Data", "Sample", "Measures", "Participants and design", "Procedure", "Estimation" for study k |
| RESULTS(k) | "Results", "Findings", "Results and discussion" for study k; in a multi-study paper, also the study's own "Discussion" (the one that closes each study, before the general discussion) |
| DISCUSSION | "(General) Discussion", "Implications", "Limitations", "Conclusion", anything after the last results unit |

Multi-study detection: repeated Method/Results headings, or headings
containing "Study 1", "Experiment 2", "Sample 1". Number the units k = 1,
2, ... in document order. A theory paragraph that opens study k belongs
to THEORY (it argues hypotheses), not to METHODS(k).

Then write two things to the report skeleton (`structure_comments_<stem>.md`):
- The section map table: function, unit, paragraph range, first words.
- The theme, one sentence: the paper's research question as you read it.
  Every later comment should serve this theme.

If `--section X` was given, dispatch only that commenter in Step 2
(INTRODUCTION = the FRONT+INTRODUCTION commenter, which reads the title, the
abstract and the introduction; STORY = the storytelling commenter).

### Step 2: Six separate commenters

Start the six required roles as separate AI sessions using the selected
model. Run them concurrently where capacity permits, otherwise in batches.
Each receives only its assigned task and inputs, reads them from disk
(do not paste the paper into the prompt), and saves its own output.
Use the template below, filling the bracketed fields.

Commenters and what they receive:

| Agent | Function units it receives | Rule packet (from the guideline) |
|---|---|---|
| FRONT+INTRODUCTION | FRONT, INTRODUCTION, plus the RESULTS units for A3.6 | Part 3 A and A-T, plus S1, S3(a), and B2.1 for near rivals declared in the introduction; frames G1-G4, G6 (G7 is left to the STORY commenter, which reads the passages where objections first arise) |
| THEORY | all THEORY units | Part 3 B, plus S3(b), S3(c), S4, S5, S7, S10 |
| METHODS | all METHODS(k) units, plus the RESULTS and DISCUSSION units for C3's verb scan and the RESULTS units for C3.3 (a study's own discussion counts as discussion: guideline 0.3) | Part 3 C |
| RESULTS | all RESULTS(k) units | Part 3 D, plus S2 (D7 and D6.10 search the whole paper for the sentence they need) |
| DISCUSSION | all DISCUSSION units, plus the INTRODUCTION for the contribution-match and tie-back checks, the THEORY units for the E2 rehash check on restated hypothesis arguments, the METHODS units for E2.1, and the RESULTS units for the E2 overreach, E2.1, E3.2b, E3.2c and E3.17 checks (including each study's own discussion, which counts as discussion for E2.1, E3.2b and E3.2c: guideline 0.3) | Part 3 E, plus S12 |
| STORY (whole paper) | the whole paper | Part 2 (S1-S18) and Part 1 (G1-G7) |

Every agent also reads Part 0 of the guideline (protocol) and the whole
of `comment_contract.md`.

Agent prompt template:

```
You are one of six reviewers commenting on the STRUCTURE and STORYTELLING
of a quantitative empirical paper. You leave candidate comments; you never
rewrite, never propose wording, never comment on grammar or word choice.
Silence is the default: a well-built section gets "No candidates".

Your function: [FUNCTION NAME]. Your units: [list of unit names with
paragraph ranges, e.g., METHODS(1) P041-P078, METHODS(2) P112-P130].
The paper's theme (research question), as read by the orchestrator:
[one sentence].

Read, in this order, using your file-reading tool:
1. [SKILL_DIR]/references/guideline_structure_v4_5.txt : Part 0,
   and your rule packet: [Part 3 X ... plus S...]. A rule whose class
   is FIRE produces a regular candidate; a rule whose class is OPTIONAL
   produces a candidate marked optional: yes. A covered rule has no
   check of its own, and an item a rule marks "Not checked" produces
   nothing.
2. [SKILL_DIR]/references/comment_contract.md : all of it. Your
   candidates must fit its schema (section 3) and its prohibitions
   (section 2).
3. The paper: [outdir]/paragraph_index.json (read your units; for the
   cross-checks named in your rules, e.g. S4 or E3.2, read the other
   units you need). Paragraph IDs are the "id" field. A paragraph's
   footnote and endnote texts are in its "notes" field: read them
   before you say the paper does not state something, since the
   guideline counts footnotes. Take anchor_quote from "text" only.

Procedure:
a. For each FIRE and OPTIONAL rule in your packet, run its Detect
   condition on your units. Write down the evidence (paragraph ID and a verbatim phrase).
b. For each detection, draft a candidate in the schema of the contract,
   section 3. anchor_quote must be verbatim, at most 15 words, and occur
   in exactly one paragraph of the whole index, and only once inside it
   (check by searching the "text" fields of the JSON; a match inside
   "notes" does not count). observation states what the text does; why_it_matters states
   the reader's experience; option names one move, not wording. Write
   the candidate text as you would say it to a colleague in one breath:
   plain words, one point per sentence, no sentence over 40 words, no
   opening shared with another of your candidates, no closing line that
   restates the point, and nothing about rules, agents or this
   procedure. The authors will read it alone in the margin, so it must
   stand on its own.
c. Set optional: yes for a candidate from an OPTIONAL rule.
d. Set confidence from the manuscript evidence supporting the detection,
   not from whether it needs judgment. Use medium to mark evidence or an
   inference needing review and state what needs checking in observation.
   Confidence does not decide retention; the audit checks evidence and
   the authors' strongest reply. Exact anchors are required for both labels.
e. Rank by importance and return every candidate you found; there is no
   limit on their number.

Save your complete output with your file-writing tool to
[outdir]/structure_comments_[stem]_agent_[FUNCTION].md
(first line: "model: [the model you are running on]; function: [FUNCTION];
date: [today]") and then return the same text. If you have no candidates, save and
return "No candidates for [FUNCTION]" plus one sentence saying why the
units pass. Plain ASCII punctuation.
```

For the STORY agent, replace the units line with "the whole paper" and
the packet with Parts 1 and 2; add: "S8 proportions: in a single-study
paper whose body runs about 12,000 to 14,000 words, count the words of
each part of the body from the section map (FRONT, headings, table rows
and figure captions excluded) and report them; a part outside its S8
range is a candidate, one S8 comment per paper, on the part furthest outside its range in words, anchored on that part's
heading or first sentence. State the theme
you read in one sentence at the top of your file."

Wait for all assigned commenters (six for a whole-paper run). Confirm each output file exists and is non-empty
before continuing (re-dispatch only the missing one).

### Step 3: Author's advocate

Before the author sees a proposed comment, another agent acts as the author’s advocate and develops the strongest response to each proposed comment using evidence from the manuscript and the relevant rule. The response may show that the manuscript already addresses the concern, or that a rule permits the author’s choice.

Collect every candidate (all agents, confidence high or medium) into
`[outdir]/candidates_all.md`, numbered C1... If there are no candidates,
record that result in the file and report, skip the advocate dispatch,
and continue with an empty audit table. Otherwise dispatch one
separate author’s-advocate session using the same selected model:

```
You are the authors' advocate. For each candidate comment below, read
the anchored paragraph and its neighbours, with their notes, in [outdir]/paragraph_index.json
and write the STRONGEST rebuttal the authors could give: the passage
already does what the comment asks (cite the paragraph ID), the
rules themselves allow the move (cite the rule's own text), the comment
misreads the design, or the comment is about wording rather than
structure. That other papers make the move is not a rebuttal: the
rule's class already decides whether the comment is optional. Then give one verdict
per candidate: DEFEATS (the rebuttal removes the problem), WEAKENS (the
problem stands but is smaller or partly addressed), FAILS (no good
rebuttal). Two sentences of rebuttal and one verdict line per candidate.
Do not soften verdicts to be polite. The orchestrator drops every
DEFEATS. For WEAKENS, it narrows the comment to the concern still
supported by the manuscript, accounting for what your reply resolves,
or drops it if none remains. FAILS does not ensure retention: the
comment still needs manuscript evidence and must pass the other checks.

Candidates: [paste candidates_all.md]
Save to [outdir]/advocate_[stem].md and return the same text.
```

### Step 4: Audit, rank, write the final prose (orchestrator)

The AI assistant coordinating these agents then makes the final decision about each comment. The AI assistant drops a comment when this evidence removes the alleged problem. If it answers only part of the concern, the assistant narrows the comment to the problem that remains. Every retained comment must be supported by the manuscript and understandable when read on its own.

Build the audit table from `comment_contract.md` section 4, one row per
candidate, and fill every column yourself from the files on disk:

1. Anchor: search the "text" fields of `paragraph_index.json` for the
   anchor_quote (a match inside "notes" does not count); exactly one hit,
   in one paragraph and once inside it, else DROP (or fix the anchor if the
   intended paragraph is obvious and a unique 15-word phrase exists in it).
   The comment will highlight the sentence holding the quote, so the quote
   must come from the sentence the comment is about.
2. Rule is FIRE or OPTIONAL in the guideline; else DROP. An OPTIONAL
   rule's comment is marked optional.
3. Move not allowed by the rules: if the rule's own text allows the
   move the candidate objects to, DROP. Never drop a candidate because
   other papers make the move.
4. Evidence and rebuttal: verify that the manuscript supports the concern
   after the authors' strongest reply. Recheck the evidence or inference
   identified by medium confidence and record the outcome; neither the
   confidence label nor a binary/judgment label decides retention.
   DEFEATS drops. For WEAKENS, narrow the comment to the supported concern
   that remains and account for what the reply resolves; drop it if none
   remains. FAILS does not substitute for evidence or the other checks.
5. Not wording; self-contained when read with its paragraph alone;
   duplicate or merge (two candidates on the same paragraph and issue
   become one); option is a move.
6. Decision and reason.

Each comment explains what the passage does and why it matters to a reader at that point, then suggests one action the authors could take. It does not write text for the authors. Authors can judge the comment by its reasoning without reading the rules.

Then rank the KEEPs by importance (high first, then by how early a
reader would stumble; within each unit, optional comments rank after
the regular ones). Every keeper goes into the document: there is no cap
on the number of comments, per unit or per paper, unless the user gave
`--max N`, which keeps the N highest-ranked in the run, with every
optional comment ranked after every regular one. Then
write the FINAL comment text for each keeper in the contract's shape
(section 1): quoted anchor, what it does, why it matters, one move,
importance carried by wording, no tags, no rule IDs, no proposed wording.
Write each one as a colleague would say it: plain words, one point per
sentence, no sentence over 40 words, no opening or closing shared across
comments, no emphasis frames, and nothing that names a rule, an agent or
this procedure; the authors read each comment alone in the margin, so it
must stand on its own.
Length bands: about 100-120 words for the one or two most consequential
comments, 60-90 for ordinary ones, 30-50 for minor ones; an optional
comment opens with "Optional:" and runs 30-60 words. Count the words
of every final comment and record the count in the audit table before
injecting; trim anything over its band. Then check every factual claim
in the FINAL prose against the paragraph index (counts, degrees of
freedom, which table or figure exists, what a passage or its notes say); the
advocate saw the candidate, not your rewrite, and a comment whose harm
clause ("a reader cannot check the cells") is false on the paper's own
arithmetic is worse than no comment. A
claim you cannot verify is cut, not softened. Keep each comment inside the rule
it was detected by (contract section 4). Write them to
`[outdir]/comments_final.json`:

```json
[
  {"id": "K1", "anchor_quote": "exact phrase from the paper", "comment": "final prose ...", "principle_id": "B4", "function": "THEORY", "importance": "high", "optional": false},
  ...
]
```

Save the audit table to the report as you go. A dropped candidate is
listed in the report with its one-line reason; it never reaches the
.docx. If there were no candidates, or none survives the audit, write
`[]` to `comments_final.json` and continue to Steps 5 and 6. Zero comments
is a normal completed result.

### Step 5: Inject into a copy and verify

Run this step even when `comments_final.json` is `[]`; the script then
produces an unchanged copy of the input and verifies zero added comments
and unchanged body text.

```bash
python "[SKILL_DIR]/scripts/add_comments.py" "<input.docx>" "<outdir>/comments_final.json" --out "<same folder as input>/<stem>_structure_comments.docx" --author "Structure Reviewer"
```

The script refuses to run if any anchor is missing, found in more than one
paragraph, or found more than once inside its paragraph (exit 2), refuses
to overwrite an existing `--out` file unless `--force` is passed, copies
the original, calls `inject_comments_ex`, which highlights the sentence
holding each quote (both sentences when a quote runs across a sentence
end), checks that the injected count equals the number of comments,
re-opens both files and checks that every body paragraph's text is
identical to the original, and prints an ASCII summary that says how
many comments sit on a sentence and how many fell back to a whole
paragraph. A quote taken from a text box is refused as ambiguous (the
box's paragraph and its host paragraph both contain it); take quotes from
body or table paragraphs (Scope). Read the summary. If it reports a failure, fix the JSON
and rerun; never edit the .docx by hand and never write new XML code. If
the summary says a comment covers a whole paragraph (a NOTE line names
it), its quote could not be placed on one sentence: replace the quote in
`comments_final.json` with up to 15 words from inside one sentence of that
paragraph that occur only once in it, change the quotation that opens the
comment's text to match, and rerun with `--force` on the file this run
created.
Optionally also run:

```bash
python -c "import sys; sys.path.insert(0, r'[SKILL_DIR]/scripts'); from docx_utils import verify_docx_formatting; from pathlib import Path; print(verify_docx_formatting(Path(r'<output.docx>'), Path(r'<input.docx>')))"
```

### Step 6: Report and delivery

Report `structure_comments_<stem>.md` contains, in order:
1. Header: date, input file, output file, theme sentence, scope note
   (quantitative; single or multi-study), `--max` if the user gave one.
2. Section map table.
3. Final comments table: id | function | paragraph | anchor | comment
   text | principle id | importance | optional. Regular comments first,
   then optional ones.
4. Audit table with every candidate and its decision.
5. Dropped candidates with one-line reasons (transparency; they never
   reach the .docx).
6. Verification block: injected count and the body-text-unchanged check.

For zero comments, retain this report structure, mark the final-comments
table empty, and explain whether no candidates were found or all were
dropped. Mark the audit and dropped-candidates tables empty when there
were no candidates; in that case, record that no advocate dispatch was
needed. Deliver the unchanged copy and the report.

Delivery message to the user: where the commented copy is; that the body
text is unchanged; how many regular and how many optional comments
went in and how many candidates were dropped; if any remain, the two or three most consequential comments in one line each;
and the recommendation to run the skill's proofread function separately for wording, once the structure is settled.

## What NOT to do

- Do not change, reorder, or delete any text. The output .docx differs
  from the input only by comment marks. The script enforces this; if it
  fails, stop and fix the cause.
- Do not write suggested sentences, paragraphs, titles, or abstracts.
  An option names a move ("state the direction", "give the counts").
- Do not put severity tags, rule IDs, or the name of any source in a comment.
- Do not comment on wording or voice, or on a brief recap that opens the
  discussion, which the rules accept.
- Do not run structure comments on a qualitative or theory paper.
- Use the same selected capable model for the commenters and advocate; keep each role’s assigned inputs separate.
- Do not paste the whole paper into agent prompts; agents read the index
  file.
- Review every candidate through the author's advocate and the audit
  table. When there are no candidates, use the zero-comment path above.
- Do not fabricate citations; "work on X" is the most specific pointer
  allowed.

## Crash recovery

Working files, in the order they are created: `paragraph_index.json`,
`paragraph_index_map.txt`, the report skeleton with section map and
theme, six `structure_comments_<stem>_agent_*.md`, `candidates_all.md`,
`advocate_<stem>.md`, the audit table in the report, `comments_final.json`,
the output .docx, the finished report. Resume at the first missing file.
If `candidates_all.md` and the report record no candidates, no advocate
file is expected; continue with the empty audit and `comments_final.json`
containing `[]`. An empty final list still requires the output copy and
finished report.
