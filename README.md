# Management academic writing skill

This skill applies expert writing rules to help scholars convey their intended meaning clearly. The current prototype is ready to be applied to quantitative empirical papers in management and organization science. It offers two functions, used separately: structural feedback through comments and language editing through Word tracked changes. Both functions work on copies of the draft and leave substantive decisions to the author.

### Using the prototype

An author needs Microsoft Word and an AI assistant that can read files, run Python programs, and assign tasks to separate agents. An agent is a separate AI session given one task. The folder contains the rules and programs for the two functions.

### Structural feedback

Agents examine the paper's sections and the paper as a whole using 134 rules. Each rule specifies a condition that warrants a comment about a common mistake or an option the author could consider.

Another agent acts as the author's advocate and develops the strongest response to each proposed comment using evidence from the manuscript. The response may show that the paper already addresses the concern or that the comment misinterprets the study. The AI assistant drops unsupported comments. If the manuscript addresses only part of a concern, the assistant narrows the comment to the problem that remains.

The author receives a copy of the draft with comments attached to the sentences they concern. Each retained comment explains the problem, why it matters to a reader at that point, and an action the author could take. The author decides how to revise the argument.

### Language editing

The editing function applies sentence-level rules to wording, sentence construction, verbs, and rhythm. For example, it breaks sentences that combine several distinct claims and changes "this paper" to "we" in a co-authored paper.

The editing agent checks whether each revision preserves technical terms, construct labels, statistics, and the author's claims. A separate agent compares the original and revised paragraphs and undoes revisions that change the original meaning. If a passage needs a missing reason or a substantive choice, the skill leaves a comment instead of supplying the content.

The author receives a Word document with tracked changes to accept or reject. The technical appendix describes the rules and demonstrations in more detail.

The checks may miss edits that change an author's intended meaning. A revision can therefore reach the author without a warning that its meaning has changed. The skill is designed to preserve meaning, and the author decides on every comment and edit.
