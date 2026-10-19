# Example edit file: Methods chunk (P012-P014)

Authorship: co-authored. Rule set: rules_language_v13.txt.
A chunk agent's full file also holds the Quick Diagnosis, the Edit Log and
the Verb Upgrade Log (workflows/proofread.md, Required Outputs); this example shows the
parts the apply script and the blind meaning check use.

## Edit File

### P012
ORIGINAL: In order to examine how project teams respond to d
ACTION: REVISE
REVISED: To examine how project teams respond to deadline changes, we distributed surveys to 1,204 team members.
ISSUE_LABELS: FILLER; PASSIVE VOICE
PASSES: Pass 1; Pass 3
COMMENT: null

---

### P013
ORIGINAL: Teams that face deadline changes communicate with
ACTION: KEEP
COMMENT: The step from communicating with clients to missing fewer milestones is not stated, and readers need it to follow the argument. What connects the two?
IMPORTANCE: MINOR
ANCHOR: Teams that face deadline changes communicate with

---

### P013
ORIGINAL: Teams that face deadline changes communicate with
ACTION: KEEP
COMMENT: Is "deadline changes" short for the construct "client-initiated deadline changes"? Clarifying this would help readers know whether the two terms refer to the same construct.
IMPORTANCE: MINOR
ANCHOR: Teams that face deadline changes communicate with

---

### P014
ORIGINAL: We measure team autonomy as the share of decisions
ACTION: KEEP
COMMENT: The reason for measuring autonomy as the share of decisions made without manager approval is not stated; readers will want to know why this measure fits the construct.
IMPORTANCE: MINOR
ANCHOR: We measure team autonomy as the share of decisions

---

## Step 7 QC Results

| Para | Span (original -> proposed) | Test 1 | Test 2 | Test 3 | Decision |
|---|---|---|---|---|---|
| P012 | "In order to examine" -> "To examine" | PASS | PASS | Rule-backed edit, Pass 1 filler: "In order to" is on the filler list; same meaning, two words shorter | KEEP |
| P012 | "surveys were distributed" -> "we distributed surveys" | PASS | PASS | Rule-backed edit, Pass 3 active voice: names the actor; no passive exception applies | KEEP |

Totals: 2 spans proposed, 0 reverted.

