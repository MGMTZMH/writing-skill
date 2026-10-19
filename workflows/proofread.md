# Proofread (workflow)

The skill's `SKILL.md` sends you here when the user chooses proofreading. Follow this
file exactly. `[SKILL_DIR]` is the skill's root folder, the one that holds `SKILL.md`.
Read `references/execution.md` and `references/tracked_output.md` for the
capability, file and validation contracts.

This editor decides HOW a draft is worded. It works at the sentence level
and below: clarity, flow, verb strength, cadence, house style. It does not
decide WHAT a paper says or WHERE: it does not reorganize sections, add
missing content, write new paragraphs, or question an argument.

The contract: the editor proposes, the author disposes. Sentence-level
changes are written as tracked insertions and deletions in a copy of the
original Word document for the author to accept or reject; every edit faces
quality control whose default is reversal; every stopped edit is recorded
rather than forgotten. Missing logic and reasons are flagged
for the author, never written by the editor.

Rule set: `rules_language_v13.txt`.

## Before you start

The opening questions in `SKILL.md` give you two answers; confirm them in one line:

1. **Scope:** the whole manuscript, or one named section?
2. **Authorship:** is the paper sole-authored or co-authored? The answer
   decides first-person wording: "I" for a sole author, "we" for co-authors
   (rules_language rule 19 and Pass 1). Do not start editing without it, and
   pass it to every chunk agent.

The draft must have no unresolved tracked changes: the indexing script stops
on one (Step 1). The editor does not reorganize or rewrite content; structural
concerns and retained manuscript questions go to anchored Word comments, and the skill's
structure comments function is the place for them.

## Rule set

The rules are in `[SKILL_DIR]/references/rules_language_v13.txt`, in five
parts:

- Part 1: quick diagnosis labels
- Part 2: the four editing passes
- Part 3: intervention intensity (what is always edited, always flagged, flagged as optional, sometimes edited, never edited)
- Part 4: the 23 house style rules
- Part 5: detection procedures and flag-only checks

## File resolution

`[SKILL_DIR]` is the skill's root folder. Everything needed is bundled:

| Path | What it is |
|---|---|
| `references/rules_language_v13.txt` | the rule set |
| `README.md` | overview for users |
| `requirements.txt` | the Python packages the scripts were tested with |
| `scripts/index_manuscript.py` | stops on unresolved tracked changes; writes `para_index.txt`, `locked_paragraphs.txt`, `tables.txt`, `sections.txt`, `notes.txt` |
| `scripts/apply_language_edits.py` | applies the edits, verifies, writes the manifest |
| `scripts/editor_io.py` | shared helpers, including the docx_utils loader |
| `scripts/docx_utils.py` | shared document and comment helpers |
| `scripts/docx_revisions.py` | source-preserving tracked revisions and accepted/rejected views |
| `references/tracked_output.md` | edit-plan schema, output validation and errors |
| `references/execution.md` | capability checks, role contracts and shell commands |
| `examples/` | a sample edit file and a sample overrides file |

The scripts give the same result from any working folder.

All outputs for a draft go to `language_editor_output/[draft name]/` next to
the input document ([draft name] is the file name without extension); the
commands below call it `[output folder]`. If that folder already holds files,
never delete them: ask the author for a new folder name, because Step 9
would apply edit files left from an earlier run.

## Shared conventions

- Use the shipped DOCX programs; do not write replacement document code.
  Step 9 validates the approved plan, accepted and rejected views, source
  formatting, protected features and all retained comments before publication.
- Preserve curly quotes; never flatten typography.
- Every dispatched agent writes its output to its own file as the LAST step
  before returning, so a crash or a context reset does not lose the work.
- The program generates real tracked changes; the author reviews and
  accepts or rejects them in Word.
- Refusal records: when the Step 8 blind meaning check stops an edit (listed
  in the overrides file), the record keeps the original span, the proposed
  replacement, the test that stopped it, and the verdict; a partial reversal
  record also keeps the reconstructed text that was applied. The apply
  script writes these into `language_edit_manifest.json`, so the author can
  see every edit that was undone and why. Step 7 self-QC reversions are
  listed in the edit files. Edits the blind check kept as a slight shift in
  emphasis that leaves the claim unchanged are listed under "drift" and
  reach the author as margin comments that ask them to check the sentence.
- A paragraph the program cannot rewrite safely keeps its original text and
  is listed under "skipped" with the reason; nothing is ever half-rewritten.

---

# The workflow

## Steps

1. **Locate and index the manuscript:**
   - Use the draft from the opening questions; if none was given, ask for it
   - **MUST be .docx** -- if PDF or markdown, ask user for .docx version
   - Index it with the bundled script:
     ```bash
     python "[SKILL_DIR]/scripts/index_manuscript.py" --input "[input.docx]" --out-dir "[output folder]"
     ```
     `para_index.txt` carries the P### numbering that the edit files and the
     apply script both use. Use these indices; do not renumber.
   - **Exit code 3 means unresolved tracked changes.** Stop and tell the
     author to accept or reject all tracked changes in Word (Review, then
     Accept or Reject, then All Changes), save the draft, and start again.
     Never edit around them: the index shows the draft without pending
     insertions, so an edit would silently drop them.
   - **Footnote and endnote marks** appear in `para_index.txt` as `[^N]` where
     they sit. Keep every mark exactly where it is, with its number, in REVISED
     text; an edit that moves, drops, adds or renumbers one is skipped.
   - **`notes.txt`** gives the text of each footnote and endnote, by paragraph
     and mark. Notes are never edited, but read a paragraph's notes before
     flagging something in it as missing (a reason, a source, a definition):
     a note may supply it.
   - **Links and reference-manager citation fields** (for example Zotero or
     EndNote citations) appear as plain text in the index. Their words must
     stay unchanged and in the same order; an edit that changes them is
     skipped, and the paragraph keeps its original text.
   - **`locked_paragraphs.txt`** lists paragraphs the program will not rewrite,
     with the reason (a content control such as a Mendeley citation, a simple
     field such as a cross-reference, an equation, a symbol-font character,
     superscript or subscript text such as R squared, a tab or manual line
     break, a picture placed between words, or hidden, struck-through,
     capitalized or raised text, or text in a character style such as
     Emphasis, set by formatting that differs within the paragraph). Give
     every chunk agent this list: propose no REVISE for these paragraphs; a
     flag on them is fine. A problem an editor would otherwise fix in a
     locked paragraph becomes a comment on it, including a question for the
     author when it needs facts the draft does not hold.
     In a locked paragraph the index leaves out the words of content controls
     and simple fields, so a Mendeley citation shows as a gap: do not flag
     such a gap as a missing reference.

2. **Read the language rule set** from `[SKILL_DIR]/references/rules_language_v13.txt` (all five parts).

3. **If a section name is given** (e.g., "introduction", "methods"), find its paragraph range in the index and edit only those paragraphs. Otherwise, edit the full paper.

4. **Chunk the manuscript** by section from the paragraph index and `sections.txt` (see Chunk-by-Chunk Editing below). Record each chunk's P### range so every edit entry can be addressed back to the document.

5. **For each chunk:** read `tables.txt` if the chunk reports or interprets tables, then run the Quick Diagnosis (rules_language Part 1), then the Four Editing Passes (Part 2), governed by the four levels of editing (Part 3), the house style rules (Part 4), and the detection procedures and flag-only checks (Part 5). Use the authorship answer for every first-person change.

6. **After all chunks are edited:** run a cross-chunk consistency pass (verify construct labels, tense, tone, and acronym first use, and raise the causal-verb flag when the design calls for it; a flag from this pass goes in a KEEP entry of its own when the paragraph already has one). Write any fix from this pass into that paragraph's existing entry in its section's edit file (revise its REVISED text, or turn an existing KEEP entry into REVISE and keep its comment; add an entry only if the paragraph has none), so no paragraph has two REVISE entries: the apply script stops on that. Step 9 reads the per-section edit files directly; do not also write a compiled copy that matches the same file pattern.

7. **Quality control pass (adversarial review):**

   The editing agent puts each edit through three tests, any of which can undo it. The first test protects technical terms, construct labels, variable names, and statistics. The second test protects the author’s claims: their scope, strength, source, the mechanism they describe, and the stakes the author set. For example, changing “is associated with” to “causes” alters the strength of the claim. The third test asks the agent to name the edit’s reason. The reason can be a rule in our list, such as cutting filler. Otherwise, the edit must fix a grammar error, remove an ambiguity, or cut words without losing meaning. An edit without such a reason is undone. Each chunk agent runs these tests in order before writing its edit file; keep the span-level records below.

   **TEST 1: Technical Term Preservation (BLOCKING)**

   Did this edit change ANY of the following?
   - Variable names (e.g., "team cohesion" -> "group cohesion")
   - Construct labels (e.g., "relational improvisation" -> "team adaptation")
   - Measure names (e.g., "patent citation count" -> "innovation output")
   - Statistical notation (e.g., "r = 0.42" -> "correlation of 0.42")
   - Domain-specific terms that precision requires (jargon replaced by an exact common term under rules_language Pass 1 is not a change)

   **If YES to ANY:** AUTOMATIC REVERT. No exceptions. Technical precision is non-negotiable.

   **Example:**
   - REVERT: "R&D intensity" -> "research spending" (variable name changed)

   **TEST 2: Claim Change (REVERT ON FAIL)**

   Did the edit change what the author is claiming? Check for FIVE kinds of claim change:

   **a) Scope changes:**
   - Added quantifiers: "all", "never", "always", "every", "no" when original was bounded
   - Removed qualifiers: cut "some", "often", "typically", "in certain contexts"
   - Generalized beyond study scope: "managers" when original said "store managers in our sample"

   **b) Strength changes:**
   - Causation vs correlation: "causes" when original said "is associated with"
   - Evidence strength: "demonstrates" when original said "suggests"; "establishes" when original said "indicates"
   - Removed hedges without justification: cut "may", "could", "appears to"

   **c) Attribution changes:**
   - Removed citations: "Research shows X" when original said "Smith (2020) found X"
   - Removed source attribution: "X argued" -> stated as fact
   - Changed whose claim: "we find" when original said "prior work suggests"

   **d) Mechanism changes:**
   - Changed HOW something works: "through increased visibility" -> "by attracting attention" (different causal pathway)
   - Changed the process described: "Firms respond to supplier failures by adding backup suppliers" -> "Firms diversify their supplier base after supplier failures" (same outcome, different mechanism)
   - Added a relation the author did not write: joining two sentences with "because," "so," or "but" when the original stated no such link

   **e) Stakes removal (argumentative framing):**
   - Cut an evaluative word that carries the claim: "striking" in "a striking reversal"
   - ONLY when the word establishes why the gap or finding matters
   - OK to remove an intensifier that only adds emphasis ("fundamentally," "very")

   **If YES to ANY kind:** REVERT. No exceptions: an edit that changes what the author claims is never kept, however much better it reads.

   **Examples:**
   - REVERT: "team cohesion is associated with delivery speed" -> "team cohesion increases delivery speed" (strength: correlation -> causation)
   - REVERT: "some managers" -> "managers" (scope: removed qualifier)
   - KEEP: "the reason for that is" -> "because" (same meaning, fewer words)
   - KEEP: "a wealth of research" -> "substantial research" (argumentative framing preserved, word choice improved)

   **TEST 3: Value Creation (STRICT)**

   Did the edit make the prose MEASURABLY better? To pass, the edit must meet at least ONE of these four reasons:

   - **Rule-backed edit:** The edit carries out an edit rule, that is, any rule in rules_language that tells the editor to change text: the rules of the Four Editing Passes (Part 2) that change text, every item on the ALWAYS edit and SOMETIMES edit lists (Part 3), the corrections that NEVER edit item 1 still allows, the house style rules (Part 4), and the Part 5 procedures that end in an edit (for example the accessible-label rule). Three conditions: name the specific rule, not only its pass; the original must meet the rule's condition (for a house style rule, break the rule); and the edit must be the change the rule prescribes, with nothing added. A rule marked FLAG ONLY, an ALWAYS flag or OPTIONAL flag (Part 3), and a case where a rule says to flag instead of edit (a drifting construct label, rule 11; a style problem with a verb in a hypothesis sentence, NEVER edit item 7) never back an edit: they call for a comment.
   - **Grammar error correction:** Fixed objective grammar or punctuation error (subject-verb agreement, tense error, pronoun reference, article usage, sentence fragment, misplaced comma)
   - **Clarity gain:** Removed actual ambiguity or confusion (not just "sounds better")
   - **Concision gain:** Reduced the word count of the edited span without losing meaning

   **REVERT IS THE DEFAULT.** Edits must prove measurable value to survive.

   **AI Language Check:** Does the revised text sound MORE AI-generated than the original?

   Check for increased use of:
   - Over-polished/over-formal tone that loses natural voice
   - Generic academic phrases replacing specific language ("teams that missed deadlines" -> "teams that engaged in deadline noncompliance activities")
   - Unnecessarily complex phrasing replacing simple directness ("try to determine" -> "endeavor to ascertain")

   This check does not revert an edit by itself. When it finds any of these, go back to the reason you named and check it again: no edit rule prescribes a more polished, generic, or complex wording, so such an edit usually has no reason and is reverted as no value. An edit that meets a reason with concrete evidence is kept.

   **Meta-Guidance for Uncertain Cases:**

   When unsure if an edit adds value, apply this 3-step verification:

   **Step 1: Name the reason**
   - Which of the four reasons does this edit meet? For a rule-backed edit, name the rule (for example "Pass 1, filler", "Pass 2 A, buried lead", "house style rule 8").
   - If you can't name ONE -> REVERT immediately
   - If you can name one -> proceed to Step 2

   **Step 2: Demand concrete evidence**
   - "Rule-backed edit" -> Quote the words of the original that meet the rule's condition, and show that the edit makes the change the rule prescribes and nothing more. For common rules the evidence is:
     - Filler, redundant modifiers, prepositional phrases, verb-plus-noun phrases, possessives, positive form (Pass 1) -> Quote what was cut or replaced and show that the meaning is unchanged.
     - Plain words: pompous phrases, shorter words, jargon, foreign words (Pass 1) -> Name the plain word or phrase and show that the meaning is identical.
     - Third-person self-reference (Pass 1) -> Quote the phrase and give the authorship answer that sets "we" or "I".
     - Buried lead, stuffed sentence, abstract or chained sentences, parallel lists and list order (Pass 2 A to B4); sentence types, sentence length, pauses, combining short sentences (Pass 4) -> Name the problem and quote where it occurs.
     - Acronym at first use (Pass 2 C2) -> Quote the first use and the manuscript's own expansion.
     - Active voice (Pass 3) -> Quote the passive, name the actor, and say why none of the four passive exceptions applies.
     - Verb upgrade (Pass 3) -> What mechanism/event/relationship info did the new verb add, and where does the author already state it?
     - Synonym for a repeated action verb (Pass 3) -> Quote the repeated verb in both sentences and show that the synonym keeps the meaning.
     - House style fix (Part 4) -> Which house style rule number did the original violate?
   - "Grammar error" -> What specific rule was violated? Quote the error.
   - "Clarity gain" -> What specific confusion existed? Quote the ambiguous phrase.
   - "Concision gain" -> Count the words of the edited span before and after, and show the reduction.

   If you cannot provide concrete, quotable evidence -> REVERT

   **Step 3: Test for no-gain swap**
   Is this just a synonym swap with no functional difference?

   Common patterns:
   - Abbreviation expansions that no house style rule requires: etc.->and so forth; i.e.->that is or e.g.->for example inside parentheses, where rule 17 allows the abbreviation
   - A verb that states how strong the evidence is for another such verb, or plain verb to longer verb: indicates->shows, suggests->implies, uses->employs (evidence swaps also fail Test 2b)
   - Register shifts without precision gain: important -> significant, study -> research
   - Not a no-gain swap: a change an edit rule prescribes is a rule-backed edit, for example a move to a plainer word with identical meaning ("utilize" -> "use", "inter alia" -> "among other things"; Pass 1) or a change a house style rule requires ("i.e." or "e.g." in running text written out as "that is" or "for example", rule 17). This list never overrides an edit rule.

   If YES (no-gain swap) -> REVERT regardless of Steps 1-2

   **Decision tree:**
   ```
   Uncertain edit
     |
   Step 1: Can you name a reason?
     NO  -> REVERT
     YES -> Step 2
   Step 2: Can you provide concrete evidence?
     NO  -> REVERT
     YES -> Step 3
   Step 3: Is this a no-gain swap?
     YES -> REVERT
     NO  -> KEEP
   ```

   **Examples:**
   - REVERT: "research indicates" -> "research shows" (Step 3 fail: no-gain swap; also Test 2b, a verb that states how strong the evidence is, swapped)
   - KEEP: "in order to examine" -> "to examine" (Step 2: rule-backed edit, Pass 1 filler; "in order to" is on the filler list, 4 words->2 words, same meaning)

   **COMBINED DECISION LOGIC:**

   | Test 1 (Technical) | Test 2 (Claim) | Test 3 (Value) | ACTION |
   |-------------------|------------------|-------------------|---------|
   | FAIL (changed technical term) | - | - | **REVERT** (automatic, no exceptions) |
   | PASS | FAIL (claim changed) | - | **REVERT** (regardless of value) |
   | PASS | PASS | FAIL (no value) | **REVERT** |
   | PASS | PASS | PASS (a named reason with concrete evidence) | **KEEP** |

   **Granular QC for mixed edits:** When a sentence contains multiple distinct edits: (1) align original and revised text word-by-word and mark each distinct edit span, (2) run Tests 1-3 on each span separately and mark it KEEP or REVERT, (3) reconstruct the sentence starting from the original and applying only KEEP-marked changes. If unsure how to reconstruct, full-revert is always safe.

   **Record the reverts.** A proposal is any change you drafted and put through the three tests, whether or not it reached the edit file; a change you set aside before drafting it is not counted. In each edit file, list every reverted proposal with the test that stopped it, and give the totals (edited, reverted, revert rate).

### Step 8: Independent meaning check (MANDATORY)

A separate agent then checks the meaning. It sees only the original and revised paragraphs, but not the reasons for the edits. It asks whether the revision makes the same claim, with every hypothesis, construct label, number, and citation intact. Wording and minor nuance may change. If only emphasis shifts slightly while the claim remains unchanged, the edit remains, but the sentence will be flagged by a comment for the author to verify. If the claim changes, either the entire edit or the part responsible for the change is undone. Every undone edit is listed with the check that stopped it. Start this separate session after Step 7 and before Step 9; do not pass editing history or reasons.

**This step is mandatory for every run.** The built-in QC tests
(Tests 1-3) are applied by the same agents that made the edits. An independent agent
catches changed claims that self-QC misses.

**Agent setup:**
- A fresh, isolated AI session using the same selected capable model, with only the inputs listed here.
- Provide BOTH the ORIGINAL and REVISED text for every edited paragraph (from all
  section edit files)
- The agent must NOT receive the editing rationale or edit log -- it evaluates the
  revision blind, without knowing why changes were made

**Verification standard: THE SAME CLAIM.**

The standard is: **does the revised text make the same claim as the original?**
The claim is the same when its scope, strength, attribution and mechanism
are unchanged, every hypothesis and construct label stays as written, and
every citation is intact. Numbers keep their values, units and statistical
meaning; a same-value spelling change required by a house style rule
passes the number check. Wording and minor nuance may change: an edit that cuts a long
sentence to its point, as good editing does, keeps the claim.

**Agent procedure for each paragraph:**

1. **Word-level diff:** Compare ORIGINAL and REVISED word by word. List every
   word or phrase that was REMOVED from the original. List every word or phrase
   that was ADDED in the revised version.

2. **For each removal, answer:** Does removing this word/phrase change what the
   sentence claims? Specifically:
   - "consequences of whether" vs "whether" -- these mean different things
     (studying consequences vs. studying the decision). FLAG.
   - "some managers" vs "managers" -- removing "some" changes scope. FLAG.
   - "in order to" vs "to" -- removing filler, meaning identical. PASS.

3. **Citation count:** Count parenthetical citations in ORIGINAL. Count in REVISED.
   If counts differ, FLAG FOR REVERSAL immediately. List the specific missing
   citation(s).

4. **Verdict:** For each difference found, the agent says whether the claim
   is unchanged. A removal or addition that changes scope, strength,
   attribution or mechanism, or touches a hypothesis (the words of a
   hypothesis statement or of a sentence that restates one, or whether a
   hypothesis is supported), a construct label or a citation, or changes a
   number's value, unit or statistical meaning, is FLAG FOR REVERSAL. A change that only drops
   wording or minor nuance passes. A slight shift in emphasis that leaves the
   claim standing is MINOR DRIFT.

**Agent checks (applied via the word-level diff above):**
1. **Scope preservation:** Were any qualifiers removed (some, often, typically,
   in certain contexts, consequences of)? If yes, does removal change the claim?
2. **Strength preservation:** Were any hedges removed (may, suggests, is associated
   with)? If yes, does removal escalate the claim? For "research suggests
   that," check whether the phrase only repeats source attribution or also
   qualifies the evidence. Keeping a citation preserves attribution; it does
   not by itself show that the strength is unchanged. Pass the cut only when
   the citation condition in rules_language Pass 1 is met and the phrase has
   no uncertainty function in context. If the cut strengthens the claim, or
   its effect is unclear, flag it for reversal.
3. **Attribution preservation:** Are all citations and source attributions intact?
   Count them explicitly -- do not state "all citations preserved" without counting.
   The orchestrator tells the agent whether the paper is sole-authored or
   co-authored: in a sole-authored paper, changing a "we" or "our" that refers
   to the author alone into "I" or "my" is not an attribution change. A "we"
   that means the author and the reader together must stay.
4. **Mechanism preservation:** Is the described causal pathway unchanged? Was a
   relation word (because, so, but) added that the original did not state?
5. **Technical term preservation:** Are all construct labels, variable names, and
   domain-specific terms that precision requires identical? (Jargon replaced by an exact common
   term under rules_language Pass 1 passes.)
6. **Stakes preservation:** Are argumentative framing phrases retained when they
   serve a rhetorical purpose?

**Verdict per paragraph:** PASS | MINOR DRIFT (keep; a one-line note in `meaning_verification.md` that quotes the words that drifted) | FLAG FOR REVERSAL

**Output:** The agent must save its complete report to
`[output folder]/meaning_verification.md` (self-saving pattern).

**After the agent completes:**
1. Read the meaning verification report
2. Write `[output folder]/verifier_reverts.json` in every run: list each
   paragraph flagged for REVERSAL (format in Step 9), or write `{}` when none
   was flagged; the apply script stops if the file is missing. Do not change
   ACTION: REVISE to ACTION: KEEP in the edit files: the apply script counts
   reversals only from the overrides file, so a hand-edited KEEP is never
   counted. A partial reversal goes in the same file with class "partial":
   reconstruct the paragraph from the original with only the spans the
   verifier did not challenge, and put that text in "revised". The apply
   script applies the reconstruction and logs the challenged spans as a
   refusal of class "partial", so one challenged span never costs the
   paragraph its other edits. Each MINOR DRIFT verdict goes in the same
   file with class "drift" and a "notes" list: one {"anchor": a phrase of at
   most 15 words copied from the revised sentence that drifted, occurring
   only once in the revised paragraph, "note": the verifier's one-line note}
   per drifted spot. A "partial" entry may carry "notes" too. A note without
   an anchor stops the apply script before anything is written. The apply script keeps these edits and turns each note into
   a comment on that sentence that asks the author to check it, so the
   author sees each slight shift in emphasis that leaves the claim unchanged;
   drift notes never count as reversals. Any other class stops the apply script.
3. The apply script writes every reversal into `language_edit_manifest.json`
   as a refusal record; no separate report is compiled (Step 10)
4. Proceed to Step 9 only after the overrides file is written

9. **Create and validate tracked Word output.** Run the shipped program after
   the independent meaning check. It resolves overrides into the full edit plan,
   writes revisions on the original package, derives an internal clean reference,
   and publishes only after validation succeeds.
   ```bash
   python "[SKILL_DIR]/scripts/apply_language_edits.py" \
       --input "[input.docx]" \
       --out "[output folder]/internal/[draft name]_edited_CLEAN.docx" \
       --tracked-out "[output folder]/[draft name]_edited_TRACKED.docx" \
       --plan-out "[output folder]/internal/approved_edit_plan.json" \
       --edits "[output folder]/language_agent_*.md" \
       --para-index "[output folder]/para_index.txt" \
       --overrides "[output folder]/verifier_reverts.json" \
       --author "Language Editor" \
       --manifest "[output folder]/language_edit_manifest.json"
   ```

   This block uses POSIX shell continuation; equivalent PowerShell syntax is
   in `references/execution.md`. Keep the quotes around the `language_agent_*.md` pattern: the script expands
   it itself, so the command behaves the same in PowerShell, cmd, and bash.

   **Reverts go in the overrides file, not in prose.** Every edit the Step 8
   blind verifier stops is recorded as a refusal record of class "reversal";
   a span-level salvage is class "partial" and carries the reconstructed text:
   ```json
   {"P031": {"class": "reversal", "test": "blind meaning check",
             "reason": "the revision narrowed the claim from all firms to large firms"},
    "P047": {"class": "partial", "reason": "span 2 escalated a hedge into an assertion; the rest kept",
             "revised": "the paragraph rebuilt from the original, passing spans only"},
    "P052": {"class": "drift", "notes": [{"anchor": "a phrase from the revised sentence",
             "note": "a slight shift in emphasis; the claim is unchanged"}]}}
   ```
   With no reversals and no drift, the file holds `{}`. `examples/` has a sample.

   The script finds each paragraph by its text as well as its P### number.
   A missing or ambiguous match blocks publication; correct the index or
   edit entry rather than guessing the intended paragraph. Straight
   quotes in an edit file still match curly quotes in the document.

   **Check six things every time:**
   1. `applied` plus `reverted` plus `skipped` plus `no_op` equals the
      REVISE entries proposed. `no_op` records an approved result identical
      to the source. Partial reversals are also reported as a review outcome;
      their surviving text is counted under applied, no_op or skipped, so
      do not add `partial` again to this total. A comment that cannot be
      placed blocks publication.
   2. Every entry under `skipped` names its reason (a locked paragraph,
      ambiguous inserted formatting, a changed link or citation field, a moved or dropped
      `[^N]` mark) and left the paragraph exactly as it was. Report skips in
      the delivery.
   3. The process exit status is 0 and the manifest says `ready`. Read the
      validation result and output hashes, not just the printed file path.
      Every retained comment ID must be accounted for exactly once (or in
      an explicitly consolidated entry). A failed integrity or placement
      check blocks publication; see `references/tracked_output.md` for codes.
   4. Every new flag, author query and drift note is anchored to its relevant
      sentence in the accepted-text view. If an anchor is missing, removed
      by an edit or ambiguous, read the accepted passage and supply a unique
      phrase from the relevant sentence in ANCHOR or the drift note's anchor.
      Rerun from the original with fresh output paths. Do not guess relevance
      or publish a whole-paragraph fallback. Original comments retain their
      passage connection; unsafe edits are skipped and recorded.
   5. `verify: errors` is 0. `verify` compares the number of straight quotes
      (double or single) in body paragraphs with the original: a net
      increase is an error, and straight quotes that were already there are
      reported as a warning.
   6. Report the blind check's reversal rate, counting full reversals and
      partial reversals and excluding minor-drift notes. State the denominator
      and counting unit, and keep the supporting records available.

   **Statistical notation and formatting.** The tracked writer retains source
   run properties. Check that statistical symbols, citations and other
   protected content retain their source formatting. An unsupported edit is
   skipped, with its reason; do not flatten the paragraph to force the edit.

10. **Deliver the tracked Word document and editing records:**
    - The per-section edit files (`[output folder]/language_agent_*.md`) and `language_edit_manifest.json` form the complete editing record.
    - Provide the user with:
      - `[output folder]/[draft name]_edited_TRACKED.docx` (reviewable Word revisions and comments)
      - Edit count summary (N edits per section, total), flag count, and the number of comments that ask the author to check a slight shift in emphasis that leaves the claim unchanged
      - Skipped edits and locked paragraphs, with their reasons; paragraphs holding content-control citations (for example from Mendeley) were not edited, and running on a copy with citations converted to plain text includes them
      - Confirmation that all retained manuscript questions are in the Word comments
      - Instructions to open the tracked document in Word, review its comments, and accept or reject each revision

---

## Chunk-by-Chunk Editing

**Never edit the entire paper in one pass.** Split the manuscript into chunks by major section and edit each chunk independently so every section receives equal depth of attention.

### Chunking rules

- **Default chunks:** Title and Abstract, Introduction, Theory/Hypotheses, each Study (Method + Results + Discussion together), General Discussion (contribution + practical implications + limitations + conclusion).
- **If a single chunk exceeds ~2,000 words,** split it further at natural sub-section boundaries.
- **If a section name was given as input,** that section is a single chunk; skip chunking.

### Per-chunk workflow

For each chunk, in order:
1. Run the Quick Diagnosis (rules_language Part 1) on that chunk.
2. Apply Pass 1 (Clarity) to the full chunk (rules_language Part 2).
3. Apply Pass 2 (Structure) to the full chunk.
4. Apply Pass 3 (Verbs and voice) to the full chunk.
5. Apply Pass 4 (Rhythm) to the full chunk.
6. Run the Part 5 checks that apply to the chunk (flags, the title, citations).
7. Run Step 7's three tests on your own edits and record the results.
8. Record edit-log entries and verb upgrades as you go.
9. **Depth consistency check:** After editing each chunk, ask yourself: "Did I check for applicable Pass 2 to 4 problems as thoroughly in this chunk as in the first chunk?" A chunk with only Pass 1 edits is complete when the other passes find no applicable problem.

### Cross-chunk consistency pass (lightweight)

After all chunks are edited:
- Verify ALL expected output files exist on disk (file listing)
- Search all section files for key construct labels from the paper (e.g., the paper's main constructs, hypothesis labels)
- Flag any terminology drift (construct used differently across sections)
- Run the acronym first-use check across the whole manuscript (rules_language Part 5): first use is a whole-paper fact, not a chunk fact
- Decide from the methods whether the design is observational with no instrument, natural experiment or other argument for cause and effect; if so, raise the causal-verb flag (rules_language Pass 2 E) on each sentence that reports the paper's own findings with a strong causal verb, as a comment that never changes the verb
- Do NOT re-read all agent outputs in full

### Parallel agent dispatch (for full-paper edits)

Assign each chunk to a separate editing session using the same selected capable model, for whole-paper and single-section runs alike. Schedule sessions concurrently where capacity permits or in successive batches. Editing requires the same judgment quality as the coordinating review. Each agent receives the chunk text, `tables.txt` when the chunk reports or interprets tables, `notes.txt` when the chunk's paragraphs carry note marks, the authorship answer, `locked_paragraphs.txt`, the language rule set (rules_language_v13.txt, all five parts), the per-chunk workflow above, and Step 7's three tests.

**Required: self-saving pattern.** Each agent must save its output to a uniquely named file on disk as the LAST step of its execution, before returning to the orchestrator. This prevents data loss if context compacts between agent completion and orchestrator processing.

**File naming:** `[output folder]/language_agent_[section].md` (the pattern Step 9 reads)
- Introduction edits -> `[output folder]/language_agent_intro.md`
- Theory edits -> `[output folder]/language_agent_theory.md`
- Discussion edits -> `[output folder]/language_agent_discussion.md`
- (add more files for additional sections as needed)

**Each agent's output file must contain:**
- Quick Diagnosis for that chunk
- Edit File (all P### entries with ORIGINAL/ACTION/REVISED/COMMENT, and IMPORTANCE and ANCHOR on every comment)
- Step 7 QC results (every reverted proposal with the test that stopped it; totals)
- Edit Log table (every edit, plus KEPT AS-IS entries for strong passages)
- Verb Upgrade log (every upgrade made)
- Quick Checklist and Metrics
- Retained manuscript questions expressed in COMMENT entries with IMPORTANCE and ANCHOR; consolidate related questions without losing distinct issues

**After all agents complete:** Verify ALL expected output files exist on disk (use file listing or reading) before proceeding to the cross-chunk consistency pass.


## Required Outputs

### 1. Edit File

For each paragraph that needs revision, produce an edit entry in this format:
```
### P###
ORIGINAL: [first 50 chars of original paragraph]
ACTION: KEEP | REVISE
REVISED: [full revised paragraph text if ACTION: REVISE]
COMMENT: null
ANCHOR: [only with a comment: at most 15 words copied from the sentence the comment is about]

---
```
- Use the paragraph indices exactly as in `para_index.txt` (P000, P001, ... P999, P1000)
- ACTION: KEEP for unchanged paragraphs (usually optional). If a chunk needs no edits or comments, include one KEEP entry identifying an unchanged paragraph with COMMENT: null, so the saved file records a completed review. Any other ACTION stops the apply script.
- ACTION: REVISE for edited paragraphs -- provide FULL revised text. Keep every
  `[^N]` footnote mark where it sits and the words of links and citations unchanged;
  propose no REVISE for a paragraph listed in `locked_paragraphs.txt`
- Straight quotes in revised text are converted to curly quotes. Dashes are not
  converted: write the character you mean (em dash, U+2014; en dash, U+2013).
  A `--` stays as two hyphens, and the apply script warns when an edit adds one.
- COMMENT: use "null" by default. The exception is a flag-only check that
  points at one paragraph (rules_language Parts 2-5; for example a missing
  logical step, a missing reason for a measure, a relative clause that
  might narrow its noun, a coined acronym, a construct used before its
  definition, a footnote candidate, a title problem, a citation pile-up, or a problem in a locked paragraph
  that would otherwise be edited): write the comment and add
  `IMPORTANCE: MAJOR|MINOR`, which the apply script requires on any entry
  that carries one. MAJOR marks a problem that touches a hypothesis, a main
  construct, or a main finding; everything else is MINOR, including a
  problem in a locked paragraph that touches none of these. A rule that sets
  importance itself overrides this definition (rules_language Part 3, ALWAYS
  flag). Write the
  comment itself in plain words (rules_language Part 3, ALWAYS flag): no
  bracketed tags such as [MAJOR] or [MINOR] and no closing formula such as
  "Author decides." An optional flag (rules_language Part 3, OPTIONAL
  flags) opens with "Optional:". Add `ANCHOR:` with a phrase of at most 15 words
  copied from the sentence the comment is about and occurring only once in
  the paragraph (from the REVISED text when the paragraph is revised; where
  you can, choose words the edit leaves unchanged, so the flag still finds
  its sentence if the blind check reverses the edit): the comment highlights
  that sentence. Without a unique ANCHOR in the accepted paragraph, publication stops.
  Read that passage, supply a relevant unique sentence anchor and rerun. A flag can ride on a KEEP entry; write
  that entry into the file even though KEEP entries are otherwise optional.
  A paragraph that needs a second flag gets a second entry with ACTION: KEEP
  and its own COMMENT, IMPORTANCE and ANCHOR (the ANCHOR taken from the
  REVISED text when the paragraph is also revised); never a second REVISE
  entry, which stops the apply script.
  Questions that need facts absent from the manuscript, such as an
  acronym’s expansion, also belong in anchored COMMENT entries. Do not
  invent the missing fact. Ask necessary setup questions, including
  authorship, before editing.
- ISSUE_LABELS and PASSES are optional but useful: they let the manifest say
  which diagnosis and which pass produced each edit.

`examples/language_agent_example.md` shows a complete edit file.

### 2. Edit Log Table

| Issue Label | Original Snippet | Revised Snippet | Why This Improves It |
|-------------|-----------------|-----------------|---------------------|
| FILLER | "..." | "..." | One sentence |
| BURIED LEAD | "..." | "..." | One sentence |
| ... | ... | ... | ... |

Log every edit, and add KEPT AS-IS entries for passages that were already strong. Never edit to reach a count. For a full paper, entries should come from every section.

### 3. Verb Upgrade Log

| Original Verb/Phrase | Replacement | Rationale (precision, event, mechanism, or cadence) |
|---------------------|-------------|------------------------------------------------------|
| "is related to" | "correlates with" | Precision -- names the statistical relationship without adding a direction |
| "has an effect on" | "increases" | Precision: states the sign the reported result gives; the original already claimed an effect; never in an observational study (rules_language Pass 3) |
| ... | ... | ... |

Log every upgrade you make. An upgrade that adds what the author did not already state changes the claim (Test 2; rules_language Pass 3, limit 2).

### 4. Quick Checklist and Metrics

- Passive constructions changed: [count]
- Acronyms defined on first use: [list]
- Filler cuts: [count of phrases removed or replaced]
- House style violations corrected: [count by rule number]
- Flags raised: [count by check]

### Manuscript questions and comment coverage

Comments on issues that require the author’s judgment belong in the output Word document. Record each retained manuscript clarification request as a COMMENT with an ANCHOR. Combine duplicate or closely related questions when they address the same issue, preserve each distinct question, and identify any other relevant locations. Do not deliver a separate manuscript-question list. Use ACTION: KEEP for a question on unchanged text, or add it to the existing edit entry. Include IMPORTANCE and a unique sentence ANCHOR. Check every retained question against the output manifest and Word comments. Keep setup questions before the run; never invent missing facts to complete an edit.

---

## Final Verification

Before finalizing, confirm in 3-6 bullets that you:

- Produced the four required edit-record components above, the tracked Word document and its validated manifest; accounted for every retained manuscript question in the Word comments.
- Enforced all House Style rules in rules_language_v13.txt Part 4 (list any rule you could not apply and why).
- Logged every edit and every verb upgrade.
- Did not introduce new substantive claims, reasons, or evidence; flagged gaps instead.
- Preserved all technical terms and construct labels consistently.
- Verified that accepting revisions matches the approved plan, rejecting them restores the original text and formatting, and all protected features and comments pass validation.

### Editorial Balance Checkpoint

Review whether each section received all applicable checks and whether its edits match the problems found. **Signs of balanced editing:**
- **Intervention matches need:** Revise paragraphs that meet a rule's conditions and preserve already-strong prose under the strong-prose rule.
- **Pass coverage:** Check all four passes in every section; a pass may produce no edits when the section has no applicable problem.
- **Verb upgrades, where made:** Each adds information the author already states
- **Preserved strengths:** Edit log includes entries marked "KEPT AS-IS" for already-strong prose

**Signs of under-editing:**
- Applicable Pass 2 to 4 problems remain after Pass 1 edits.
- Later sections contain applicable Pass 2 to 4 problems that the edit log leaves unaddressed.
- Verb upgrade log shows only swaps that add no information
- Identified buried leads or stuffed sentences remain unaddressed.

**Signs of over-editing:**
- Every paragraph revised regardless of quality
- Evaluative words that carry the claim systematically removed ("a striking reversal" -> "a reversal")
- Original phrasing replaced with synonyms that add no clarity
- Technical precision sacrificed for brevity ("correlation" -> "link", "moderates" -> "affects")
- Missing logic or reasons written into the text instead of flagged

---

## Reviewing the tracked Word document

Deliver the validated `[draft name]_edited_TRACKED.docx`. Tell the author
to open it in Microsoft Word, review each insertion, deletion and comment,
and accept or reject revisions. The accepted clean reference and full edit
plan remain in the run's `internal/` folder for verification. Keep the
original draft untouched and retain undone/skipped edit records.

---

## Scope: Language, Not Content

This editor focuses on **language**: clarity, concision, verb strength, sentence structure, cadence, and house style compliance. It does NOT reorganize sections, add missing content, write new paragraphs, or restructure arguments.

Do not remove paragraphs or sections, reorganize, or question the argument's structure; the sentence-level cuts in the rule set (filler and asides that are off the paper’s topic and do nothing for their sentences) still apply. Repeated definitions receive a comment asking which wording to keep. Missing logic and missing reasons are flagged, never written. Retained structural concerns are delivered as anchored Word comments for the author.

---

## What NOT to Do

- **Never modify the original .docx.** Edits go into the edit files; Step 9's script writes them into a copy.
- **Do NOT add new citations, theories, reasons, or substantive claims.** You are editing, not co-authoring.
- **Do NOT reorganize paragraphs, add missing sections, or restructure arguments.** Those are structural decisions for the author. If you notice a structural gap, leave an anchored Word comment for the author without supplying the missing content.
- **Do NOT over-edit.** If the original phrasing is already strong, leave it: well-crafted academic prose does not need rewriting.
- **Do NOT change technical terms or construct labels.** If the author uses "relational improvisation" throughout, keep it -- don't alternate with "team adaptation" for variety.
- **Do NOT alter typographic conventions.** Preserve curly (smart) double and single quotes from the source; never flatten them to straight quotes.
- **Use the shipped DOCX programs.** `apply_language_edits.py` writes tracked edits from the original through `docx_revisions.py`, preserves protected content and safe source formatting, leaves locked paragraphs unchanged, and places sentence-anchored comments in the accepted view. Integrity or comment-placement failures block publication. Fix the input records and rerun from the original as the command contract requires.
- **Do NOT edit the whole paper in one shot.** Use the chunk-by-chunk process above so every section gets equal editorial attention.
- **Do NOT fabricate details.** If you can't read a section clearly, say so.
