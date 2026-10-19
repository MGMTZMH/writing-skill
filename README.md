# Writing help for management manuscripts

This folder is a skill: written instructions, rule files and small programs that an AI assistant follows when it works on a manuscript in management and organization science saved as a Word file. It has two functions, and the user picks one each time.

- The structure comments function checks a paper against 134 structure rules and leaves comments about missing content and suggested improvements.
- The proofreading function edits the wording under our rules while trying not to change the author’s intended meaning. Structure comments serve quantitative empirical papers and never change the text.

The original file is never changed: each function works on a copy.

## What the author gets back

- Structure comments: a copy of the draft with comments on the sentences they concern (`<draft>_structure_comments.docx`, next to the draft), and a report that lists every comment and every proposed comment the checks dropped, with the reason. A comment that opens with "Optional:" gives advice, such as a usual practice, or applies a stricter level of another check; it comes after the other comments on the same section.
- Proofreading: a Word document with tracked changes (`language_editor_output/<draft>/<draft>_edited_TRACKED.docx`) for the author to accept or reject, with sentence-anchored comments on issues requiring judgment. These include missing logical steps, unexplained choices and retained clarification questions. Related questions are combined without losing distinct issues. Edit files and a manifest record edits undone, skipped or marked; an accepted clean reference and the complete approved edit plan are kept under `internal/`.

## How it works

The AI assistant is the AI session the author works in. It runs the programs, starts the agents, and writes the final wording for each structure comment. An agent is a separate AI session that receives one task and only the necessary files and its own instruction prompts. All agents use the same selected AI model with different role instructions. The programs index paragraphs, place comments and check unchanged text for structure runs. For proofreading, they create real Word insertions and deletions, check that accepting them gives the approved text and rejecting them restores the source text and formatting, and verify protected objects and comment coverage before publishing.

- Structure comments: six separate agents read the paper, in parallel or in successive batches within the available concurrency limit, one each for the introduction, the theory, the methods, the results and the discussion, and one for the paper as a whole. They propose comments under the rules in `references/guideline_structure_v4_5.txt`. Before the author sees a proposed comment, another agent acts as the author’s advocate and develops the strongest response to each proposed comment using evidence from the manuscript and the relevant rule. The response may show that the manuscript already addresses the concern, or that a rule permits the author’s choice. The AI assistant coordinating these agents then makes the final decision about each comment. The AI assistant drops a comment when this evidence removes the alleged problem. If it answers only part of the concern, the assistant narrows the comment to the problem that remains. Every retained comment must be supported by the manuscript and understandable when read on its own. Each comment explains what the passage does and why it matters to a reader at that point, then suggests one action the authors could take. It does not write text for the authors. Authors can judge the comment by its reasoning without reading the rules.
- Proofreading: one agent per section edits in four passes under the rules in `references/rules_language_v13.txt`. The editing agent puts each edit through three tests, any of which can undo it. The first test protects technical terms, construct labels, variable names, and statistics. The second test protects the author’s claims: their scope, strength, source, the mechanism they describe, and the stakes the author set. For example, changing “is associated with” to “causes” alters the strength of the claim. The third test asks the agent to name the edit’s reason. The reason can be a rule in our list, such as cutting filler. Otherwise, the edit must fix a grammar error, remove an ambiguity, or cut words without losing meaning. An edit without such a reason is undone. A separate agent then checks the meaning. It sees only the original and revised paragraphs, but not the reasons for the edits. It asks whether the revision makes the same claim, with every hypothesis, construct label, number, and citation intact. Wording and minor nuance may change. If only emphasis shifts slightly while the claim remains unchanged, the edit remains, but the sentence will be flagged by a comment for the author to verify. If the claim changes, either the entire edit or the part responsible for the change is undone. Every undone edit is listed with the check that stopped it.

## Using it

To use the skill an author needs Microsoft Word and an AI assistant that can read files, run Python programs, and start agents. Ask the assistant to read `SKILL.md` in this folder, name the function and provide the draft. Native skill discovery is another entry route where available. The assistant follows `references/execution.md` to map the workflow to its tools and asks only for missing information: the file, scope, and, for proofreading, whether the paper has one author or several.

Requirements and commands are in `references/execution.md`; the pinned Python packages are in `requirements.txt`. The revision writer runs directly on the DOCX package without an assistant-specific Word tool. Microsoft Word is used to review comments and accept or reject revisions. Read `references/tracked_output.md` for supported document features and error meanings.

## How it was tested

The programs include regression tests in `scripts/test_*.py`, including tracked output tests. Run them with the declared Python interpreter as described in `references/execution.md`. They check original-file protection, edit and comment placement, accepted/rejected text and formatting, and blocked publication on failure. Automated checks do not establish interoperability with every Word version or the judgment quality of every AI model. The rules themselves have not yet been tested on real manuscripts against the judgment of human editors or reviewers.

## Known limits

- Structure comments serve quantitative empirical papers, including mixed-method papers when at least one study states and tests hypotheses. Whole-paper rules check the full paper; methods and results rules apply only to the quantitative studies.
- Proofreading edits body paragraphs only: text in tables, footnotes, headers and text boxes is not edited. If the draft contains unresolved tracked changes, the run stops and asks the author to resolve them.
- Some reference-manager citations stored as Word fields can be preserved while the surrounding text is edited. Paragraphs containing unsupported objects, such as equations, inline pictures or content controls (including some reference-manager citations), are left as written and listed with the reason. Whether a paragraph can be edited depends on how its contents are stored.
- Tracked edits preserve source runs and their formatting. Inserted text uses the tested formatting policy in `references/tracked_output.md`; an edit whose formatting or protected objects cannot be mapped safely is skipped and recorded.
- The agents' judgments can differ from one run to the next, and the checks can miss a change in meaning. The author decides on every comment and edit.
- A full run needs separate AI sessions with the assigned inputs, including an independent meaning checker that receives no editing history or reasons. Tool names and concurrency limits vary by environment; use the capability and role contracts in `references/execution.md`. Compatibility must be verified in the environment actually used.

## Folder contents

```
SKILL.md                 what the assistant asks the user, and which workflow it follows
README.md                this file
requirements.txt         Python packages the programs need
workflows/               the two step-by-step workflows the assistant follows
references/              rules, comment format, portable execution and tracked-output contracts
scripts/                 the programs and their tests
examples/                a sample edit file and a sample file of undone edits
```
