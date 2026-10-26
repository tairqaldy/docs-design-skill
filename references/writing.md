# Writing guide: sound like a person, not a model

Readers can spot generated text, and so can graders. Text reads as AI-written less because of any
one word than because of habits: padding, false balance, symmetrical lists, a summary of what was
just said, and confident filler. The goal is the voice of a capable student or colleague who did
the work, knows the numbers, and has somewhere to be.

## How to write it

1. **Answer first.** The first sentence of an answer is the answer: "24 of the 38 BLOCKED
   reports are real, so P(O | BLOCKED) = 0.632." The reasoning comes after it.
2. **Use the specifics.** Write numbers, names, units and the actual values from the task. Cut any
   sentence that could appear unchanged in someone else's report on another topic.
3. **Explain with the mechanism, not with adjectives.** Write "because 14 of the 38 alarms come
   from clear paths", not "due to the significant impact of false positives".
4. **Prefer prose; save lists for real lists.** Reasoning belongs in paragraphs with
   connecting words (*so*, *because*, *but*, *which means*). Use bullets only for steps,
   checklists or genuinely parallel items, and keep them few.
5. **Vary the rhythm.** Mix a short sentence with longer ones. Don't make every paragraph three
   sentences, and don't make every list three items.
6. **Use plain verbs and a few hedges.** Write "shows", "drops", "uses" and "is". Hedge only where
   there is real uncertainty, and say what the uncertainty is.
7. **Keep the person consistent.** "I" or "we" is fine in assignments and reflections. A report can
   be impersonal. Pick one and keep it.
8. **Stop when the point is made.** No closing paragraph that restates the sections and no
   "Overall, this demonstrates…". If a conclusion is required, make it say something new: a
   limitation, a recommendation, or what would change the result.
9. **Interpret, don't decorate.** After a number, say what it means for the scenario: the robot,
   the patient, the budget. That is the part graders reward and models skip.
10. **Paraphrase the task.** Don't copy the question back as a heading sentence.

## Words and patterns to delete on sight

- **Openers:** "In today's world", "In the realm of", "When it comes to", "It is important to note
  that", "It's worth noting", "Let's dive in", "This report will explore".
- **Inflated words:** delve, crucial, pivotal, vital, robust (outside statistics), seamless,
  leverage, utilize, facilitate, showcase, underscore, highlight (as a verb, repeatedly),
  multifaceted, nuanced, landscape, realm, tapestry, testament, journey, game-changer, holistic,
  intricate, comprehensive.
- **Empty connectors:** furthermore and moreover in chains, "additionally", "in addition to
  this", "on the other hand" with nothing contrasting.
- **Stock structures:** "not only X but also Y"; "X is not just A, it's B"; triplets for
  rhythm ("clear, concise, and compelling"); rhetorical questions ("So what does this
  mean?"); colon reveals ("The result? A 20% drop.").
- **Closers:** "In conclusion", "Overall", "In summary", "Ultimately", "This highlights the
  importance of…", "plays a crucial role in…".
- **Typography tells:** an em dash in every paragraph (allow at most one or two per page and use
  commas, colons or parentheses instead), a **Bold label:** at the start of every bullet,
  Title Case Headings Everywhere (use sentence case), emoji, and exclamation marks.

## Before and after

> **Generated:** It is important to note that the posterior probability plays a crucial role in
> understanding the sensor's reliability. Not only does it account for the prior, but it also
> incorporates the likelihood, providing a comprehensive view of the situation.

> **Human:** A BLOCKED reading is right less often than the sensor's 80% suggests. Obstacles
> are rare (30%), so the 14 false alarms from clear paths make up more than a third of all
> BLOCKED readings, and P(O | BLOCKED) ends up at 0.632.

> **Generated:** In conclusion, this assignment has demonstrated the power of Hidden Markov
> Models in navigating uncertainty.

> **Human:** (no closing paragraph, or:) The filter would have been fooled by t2's CLEAR
> reading if it had trusted the sensor alone. The transition model is what pulled the belief
> back up at t3.

## Final pass

Read the draft once more and, for each sentence, ask: *would a tired, competent student actually
type this?* Delete or rewrite the ones that fail. Then check that every number in the text
matches the computed results; run the code and copy the values, never estimate them.
