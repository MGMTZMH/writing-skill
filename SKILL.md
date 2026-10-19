---
name: management-academic-writing-skill
description: "One writing skill with two functions for academic manuscripts in management and organization science; the user picks one at the start, and it runs on a Word draft (.docx). Proofread edits wording at the sentence level in four passes under 23 house style rules, tests every edit, undoes an edit that changes what a sentence claims, keeps a record of every edit it undid, and returns a Word copy with tracked changes and sentence-anchored judgment comments for the author to review. Structure comments reviews what is said where, in what order, and what is missing in a quantitative empirical paper, and leaves Word comments on the sentences concerned in an unchanged copy. Use when the user asks to proofread, edit, or polish the language of a paper draft, or to review its structure, organization, flow, or storytelling, or invokes /management-academic-writing-skill. Do NOT use for content or contribution review, for writing new text, or, for structure comments, on qualitative or pure theory papers."
---

# Management academic writing skill

Read [references/execution.md](references/execution.md) to check the required capabilities and map them to the available tools. Ask the user the opening questions below, then follow the workflow file for the chosen function. The skill has two functions for a manuscript draft in Word; the user chooses one at the start, and one run does one function.

| Function | What it does | What the user receives |
|---|---|---|
| Proofread | Edits the wording under our rules while trying not to change the author’s intended meaning, in four passes under 23 house style rules. Every edit is tested: an edit that changes what a sentence claims is undone and recorded, and a slight shift in emphasis that leaves the claim unchanged is kept and marked for the author to check. Missing logic and reasons come back as comments, never as new text. | A Word document with tracked changes (`language_editor_output/<stem>/<stem>_edited_TRACKED.docx`, next to the draft), sentence-anchored judgment comments, edit files and a manifest recording edits undone, skipped or marked. An accepted clean reference and complete edit plan are kept under the run’s `internal/` folder. |
| Structure comments | Checks a quantitative empirical paper against 134 structure rules and leaves comments about missing content and suggested improvements. Comments only; it never changes the text. | A copy of the draft with Word comments on the sentences concerned (`<stem>_structure_comments.docx`, next to the draft) and a report. |

Proofreading puts each edit through three tests: protect technical terms, construct labels, variable names and statistics; preserve the author’s claims; and require a rule-based reason, grammar correction, removal of ambiguity or shorter wording without lost meaning. An edit that fails a test is undone. A separate agent then checks meaning without seeing the editing reasons.

## Opening questions

Before any work, ask what is still unknown, in one message, and confirm
the answers in one line. Skip a question the user already answered (for
example, `proofread` or `structure` as the first argument, or a path).

1. Which function: proofread, or structure comments? Proofreading polishes
   wording that a later reorganization may cut, so when the structure may
   still change a lot, suggest structure comments first.
2. The draft: the path to the .docx. Given no file, or a PDF or another
   format, ask for the .docx.
3. The whole paper, or one section?
   - Structure comments take one `--section` value per run: INTRODUCTION
     (title, abstract, introduction), THEORY, METHODS, RESULTS, DISCUSSION
     (discussion and conclusion), or STORY (the paper as a whole). If the
     section matches no single value (for example, one study of several,
     which spans METHODS and RESULTS), say so and offer the values it
     spans, one per run, or the whole paper. There is no limit on the
     number of comments unless the user sets a total for the run with
     the optional `--max N`.
   - Proofread takes the section's name.
4. Proofread only: is the paper sole-authored or co-authored? The answer
   decides "I" or "we" in every edit.

## Then

- Proofread: read `workflows/proofread.md` in full and follow it exactly.
- Structure comments: read `workflows/structure_comments.md` in full and
  follow it exactly. It serves quantitative empirical papers only; for a
  qualitative or pure theory paper, say so and offer proofreading instead.
- A request for content or contribution review, or for new text: say that
  this skill does not do that.
- A request to move paragraphs or sections: neither function moves text;
  offer structure comments on what to move.

In both workflow files, `[SKILL_DIR]` means this folder, the one that holds
this file: rules are in `references/`, programs in `scripts/`, samples in
`examples/`.

## What holds in both functions

Neither function changes the original draft; each writes to a copy, and
the author decides on every comment and edit. Proofreading edits reach the
author as actual tracked changes in the output Word document; the author
accepts or rejects each edit. Retained manuscript clarification requests
are sentence-anchored Word comments, with related questions consolidated
without losing distinct issues.

A comment highlights the sentence it is about (the sentence holding its
quote or anchor phrase), not the whole paragraph, and the programs keep
comment marks outside links and citation fields. For structure comments, fix any reported whole-paragraph fallback and
rerun as the workflow says. For proofreading, a missing or ambiguous
sentence anchor blocks publication: fix the anchor in the accepted text
and rerun. Original comments and their connection to the original passage
are preserved; an unsafe edit is skipped and recorded. See
[references/tracked_output.md](references/tracked_output.md) for the file
contract, validation and supported document features.

Proofreading stops before any work when any relevant Word XML part has unresolved tracked
changes, including formatting or move revisions, and a paragraph the program cannot rewrite safely keeps its
original text and is listed with the reason.

The rule files are `references/guideline_structure_v4_5.txt`
(with `references/comment_contract.md`) and `references/rules_language_v13.txt`.
