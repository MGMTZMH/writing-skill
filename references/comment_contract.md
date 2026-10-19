# Comment contract for structure comments

Every Word comment this skill writes follows this contract. The commenter agents draft candidates in the schema of section 3; the orchestrator (the main thread) rewrites the survivors into the final form of section 1 and checks each against the audit table in section 4.

## 1. The shape of a comment

Each comment explains what the passage does and why it matters to a reader at that point, then suggests one action the authors could take. It does not write text for the authors. Authors can judge the comment by its reasoning without reading the rules.

A comment is 3 to 5 sentences of plain prose addressed to the authors as "you". It contains, in this order:

1. The anchor, quoted. At most 15 words of the passage, in quotation marks, exactly as written. The author must be able to find the trigger without the highlight.
2. What the passage or section does now, in plain words. Describe, do not judge.
3. Why that matters to a reader at this point in the paper. Name what the reader lacks, expects, or will misread. This is the sentence that makes the comment self-contained.
4. One concrete option, phrased as a move ("state which direction the relationship runs", "give the counts behind the changing degrees of freedom", "return to H3 in the discussion and say what its failure implies"). Never propose wording. Never write the sentence for them.

Importance is carried by the wording of sentence 2 or 3, never by a tag:
- highest: a sentence that says what else in the paper depends on the point, for example "The results and the contribution claim both rest on this step."
- ordinary: no qualifier.
- minor: "A smaller point:" or "Less consequential:"
- optional: a comment from an OPTIONAL rule opens with the word "Optional:" and runs 30 to 60 words. It is the one label a comment carries: it tells the authors that the point is optional advice. Within a section, an optional comment ranks after the regular ones.

Depth scales with importance: roughly 100-120 words for the one or two most consequential comments, 60-90 for ordinary ones, 30-50 for minor ones. The orchestrator counts words before injecting; a comment over its band is trimmed, not injected.

Cadence. Comments read in sequence in a Word margin, so a shared stamp is the first thing an author notices.

**Sentence length, address, and information.**

- Address the authors at least once per comment ("you", "your").
- Keep sentences under about 40 words.
- First person is allowed where it is the plain way to say it ("I could not tell how many plants contributed to each model").
- One observation gets one sentence, not three.

**Direct expression.**

- Use at most one "X, not Y" contrast per comment.
- Do not open a sentence with an emphasis frame ("What makes this the one to attend to is that ..."); state the fact.
- Do not end with a reassurance rider that answers an objection nobody raised ("... without changing the design"), or with a closing maxim.
- Do not frame a point as unique ("the only place", "the one to attend to").

**Variation across comments.**

- Vary the shape of the closing sentence across comments (not five "Doing X ... would let ..." closers in a row).
- Do not reuse an opening template across comments (two comments that both begin "The sentence ... is where ..." read as a stamp).
- Do not repeat a distinctive phrase across comments (a phrase such as "three planned contrasts" turning up in two of them).

**Terms and figure or table references.**

- Keep one analysis's vocabulary out of another passage (no mediation words in a non-mediation comment).
- When two figures or tables form one set, call them by one label.

Word comment author: "Structure Reviewer". Anchor: the sentence that contains the quoted phrase (both sentences when the quote runs across a sentence end), highlighted by `add_comments.py`; so take the quote from the sentence the comment is about. The quotation inside the comment still names the exact words.

## 2. What a comment never contains

- A severity tag ("[MAJOR]"), a rule ID ("(B4)"), or the words "guideline", "rule" or "storytelling principle": a comment stands on the reason it gives, which the authors can weigh, not on a list they cannot see. Nor the name of any author or book.
- Proposed text, a rewritten sentence, or a paragraph for the authors to paste.
- A cross-reference to another comment ("see my comment on H2").
- Praise as filler ("Nice setup here, but ..."). If a passage is strong, say nothing.
- Preemptive defense of the authors ("I realise space is tight, but ..."; "This may be intentional, but ...").
- A question that only asks the authors to explain ("Why did you log this?"). State what the reader cannot tell and what stating it would settle. The one exception: when the text supports two readings, say so and say what each reading would imply.
- A fabricated citation. "Work on X" is the most specific a pointer to literature gets.
- Anything about wording, grammar, verb choice, sentence length, voice, or typography, except the phrasing that a structure rule itself checks (the form of a hypothesis sentence, causal verbs in an observational design, one name per construct).
- An argument that belongs to a rule other than the one the comment was detected by, or that the audit dropped under its own rule.
- An objection to a move the rules themselves allow (for example a brief recap that opens the discussion). That other papers make a move is never, on its own, a reason to leave a comment out; the rule's class decides whether the comment is regular or optional.

## 3. Candidate schema (what commenter agents return)

One block per candidate, in this order, plain text:

```
CANDIDATE n
anchor_quote: <verbatim, at most 15 words, must occur in exactly one paragraph and only once inside it>
paragraph: P###
function: FRONT | INTRODUCTION | THEORY | METHODS(k) | RESULTS(k) | DISCUSSION
principle_id: <rule ID from guideline_structure_v4_5.txt; must be class FIRE or OPTIONAL>
optional: yes | no   (yes exactly when the rule's class is OPTIONAL)
observation: <what the text does now, one or two sentences>
why_it_matters: <the reader's experience, one or two sentences>
option: <one move, not wording>
importance: high | medium | low
confidence: high | medium   (a cue for review, not a keep/drop decision; medium marks evidence or an inference that needs checking)
```

Confidence describes support for the detection, not its importance. A judgment-based detection can have high confidence; a binary check can need further review. For medium confidence, state in the observation what needs checking.

Each agent returns every candidate it finds, with no limit on their number, ranked by importance. Silence is acceptable and expected for a well-structured section; write "No candidates for <function>" and give one sentence on why.

## 4. Audit table (orchestrator, one row per candidate)

Before the author sees a proposed comment, another agent acts as the author’s advocate and develops the strongest response to each proposed comment using evidence from the manuscript and the relevant rule. The response may show that the manuscript already addresses the concern, or that a rule permits the author’s choice.

The AI assistant coordinating these agents then makes the final decision about each comment. The AI assistant drops a comment when this evidence removes the alleged problem. If it answers only part of the concern, the assistant narrows the comment to the problem that remains. Every retained comment must be supported by the manuscript and understandable when read on its own.

| # | Anchor found once, in exactly one paragraph | Rule is FIRE or OPTIONAL | Move not allowed by the rules | Manuscript evidence and confidence review | Rebuttal verdict (author's advocate) | Not wording | Self-contained when read with its paragraph alone | Duplicate or merge | Option is a move | Words in band | Decision and reason |
|---|---|---|---|---|---|---|---|---|---|---|---|

Decision values: KEEP, MERGE(with #), DROP. KEEP requires manuscript evidence for the concern that remains after the authors' strongest reply, plus all other applicable checks in the table. Confidence guides what to recheck; neither label decides retention. Record the evidence supporting the detection and the outcome of any check prompted by medium confidence. DEFEATS is a DROP. For WEAKENS, narrow the comment to what the manuscript still supports and account for the part the reply resolves; if no supported concern remains, DROP. FAILS does not itself justify KEEP: the concern still needs manuscript evidence and must pass the other checks. The final comment must stay under the rule it was detected by: if the surviving argument after the rebuttal belongs to a different rule, or to a candidate that was dropped, the comment is trimmed back to its own rule's point or dropped. After the table, rank the KEEPs by importance; within each unit, optional comments rank after the regular ones. Every KEEP goes into the document: there is no cap on the number of comments, per unit or per paper, unless the user sets a total for the run with `--max N`, which keeps the N highest-ranked, with every optional comment ranked after every regular one. The report lists every DROP with its one-line reason.

## 5. Register

Follow the register of a developmental reviewer writing to colleagues: direct, specific, economical. State the problem; do not stack hedges. Do not explain textbook statistics to authors who know them. Quote the paper's exact words rather than paraphrasing them. Order matters within a section: the most consequential comment comes first in reading order only if the anchors happen that way; otherwise importance is carried by wording, not position.

Aim for a register that is graded rather than absolute, specific about the behaviour, free of commands, and that credits what was done. Written about the invented paper used below, it sounds like this: "The closing paragraph restates the two findings and then stops; it does end on what the plant panel now lets a reader claim, but the restatement takes more of the paragraph than the claim does."

## 6. Worked pairs

Each pair shows a candidate that fails the contract and the same substance written to pass it. All seven use one invented manuscript: a two-study paper on how decision autonomy on plant teams relates to product recall rates, built on an archival panel of 412 plants observed from 2011 through 2019 and a 244-participant laboratory follow-up that tests felt ownership as the mechanism. Its hypotheses are H1 (autonomy lowers the recall rate), H2 (autonomy is related to how long a plant takes to contain a defect) and H3 (the H2 relationship is stronger in plants with frequent line changeovers). The passages are invented.

### 6.1 Hypothesis without direction (B4)

Draft passage: "Hypothesis 2: team decision autonomy is associated with the time a plant takes to contain a defect."

Fails: "[MAJOR] (B4) Hypothesis lacks direction. Rewrite as: 'Team decision autonomy is negatively related to the time a plant takes to contain a defect.'"
Why it fails: tag, rule ID, proposed wording.

Passes: "'Team decision autonomy is associated with the time a plant takes to contain a defect' states a relationship without a direction. Your two preceding paragraphs argue that teams holding line-stop authority act before a supervisor is found, which points one way. A reader therefore reaches H2 expecting a sign and does not get one, and the test in Table 3 has no prediction to confirm or reject. You could state in H2 whether autonomy shortens or lengthens the time to contain a defect."

### 6.2 Storyline opened and never closed (S4)

Draft passage (theory): a full paragraph builds "escalation delay" as the mechanism behind H2; the construct never appears again after the hypothesis.

Fails: "Escalation delay is introduced here but I could not find it later. Did you drop it? See also my comment on the discussion."
Why it fails: a bare question; a cross-reference; no reader consequence; no move.

Passes: "The paragraph beginning 'Escalation delay grows when' builds a mechanism that carries H2, and the term does not appear again after the hypothesis. A reader asked to hold this mechanism in mind reaches your discussion without learning whether it was tested, could not be tested, or was set aside. One sentence in the discussion saying which of these is the case closes the thread you opened."

### 6.3 Causal verb in a correlational design (C3)

Draft passage (results, plant panel with fixed effects): "These results show that team decision autonomy drives the drop in recall rates."

Fails: "Causal language. Change 'drives' to 'is associated with' (C3)."
Why it fails: rule ID; this is also the wording of a single verb, which alone is a proofreader's territory. It becomes structural only when the claim outruns the design.

Passes: "'Team decision autonomy drives the drop in recall rates' makes a causal claim in a section that reports fixed-effects estimates on nine years of observational plant records. Your methods describe the design as correlational and offer no identification strategy. A reader who has followed them will read this sentence as overreach and may carry that doubt into the discussion. The claim in the results can match what the design supports, with the causal reading kept for the discussion, where you weigh it against alternative explanations."

### 6.4 Unsupported hypothesis never explained (E3.2b)

Draft passage: results state "Hypothesis 3 was not supported"; the discussion never mentions H3.

Fails: "H3 is unsupported and not discussed. Please discuss it. Consider citing the usual methods piece on interpreting null results."
Why it fails: imperative without reader consequence; a fabricated or unverifiable citation.

Passes: "The results report that 'Hypothesis 3 was not supported', and your discussion does not return to H3. Because H3 extends the mechanism behind H2, a reader will want to know whether its failure marks a boundary of that mechanism or reflects the design, such as the coarse changeover count for 2011 to 2014. A sentence or two in the discussion saying which, and what it means for the theory, turns the null into a finding."

### 6.5 Slow start to the results (S2)

Draft passage: results open with six paragraphs of descriptive statistics, correlations, and collinearity diagnostics before the first hypothesis test.

Fails: "Results section is too slow. Cut the descriptives to one paragraph."
Why it fails: a length judgment with no reader consequence; an instruction to cut.

Passes: "The results run from 'Table 2 reports descriptive statistics for the 412 plants' through six paragraphs of descriptive statistics, correlations, and collinearity diagnostics before H1 is tested. Your theory section has built expectations for three hypotheses, and a reader reaching the results is waiting for them. Diagnostics that a reader does not need in order to read the hypotheses can move to a note under the table, so the first test arrives after a paragraph or two."

### 6.6 Abstract missing a component (A7.2)

Draft abstract: names the question, the theory and the findings, but carries no data or sample sentence and never says what kind of study it is, so the study type cannot be inferred from it.

Fails: "Abstract missing component (3) type of study. [MINOR]"
Why it fails: tag; the numbered component list is internal vocabulary; no quotation.

Passes: "A smaller point: the abstract reports that 'team autonomy lowers recall rates' but never says whether the study is archival, experimental or both. Readers deciding what to read cannot tell what evidence stands behind the claim. You could name the design and the sample in one sentence."

### 6.7 A comment that should never be written

Draft passage (discussion): "In this study we examined how decision autonomy on plant teams relates to product recalls ..."

Do not write: "The discussion opens with a summary of the study. Lead with the theoretical contribution instead."
Why: the rules accept a brief recap of the question and the main findings as the opening of the discussion (rule E3.1). A comment against a move the rules accept is never written.
