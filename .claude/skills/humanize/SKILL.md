---
name: humanize
description: >-
  Draft or rewrite prose so it does not read as machine-generated. Enforces the
  hard no-em-dash rule and strips the vocabulary, punctuation and sentence
  patterns that give LLM writing away. Use when writing or editing anything a
  human will read (emails, Teams or Slack replies, letters, reports, README and
  docs prose, commit messages, PR descriptions), when asked to rephrase, soften
  or reword a message, when asked to make text sound natural, human, less
  robotic or less AI, and when asked to check text for em dashes or AI tells.
---

# Humanize

Applies to every word a person will read: chat replies in this session, email
drafts, files on disk, commit messages, PR bodies. Not only files.

## Rule 0: no em dashes. Ever.

The em dash (`—`, U+2014) is the single strongest signal that text was machine
written. It is banned outright. An en dash (`–`) used as spaced sentence
punctuation is banned for the same reason.

| Where an em dash wants to go | Use instead |
| --- | --- |
| Parenthetical aside | commas, or parentheses |
| Before an explanation, cause or list | colon |
| Joining two independent clauses | semicolon, or two sentences |
| Dramatic pause or afterthought | full stop, then a new sentence |
| Numeric range in prose | `to` ("5 to 10 minutes") |

Still allowed: hyphens in compound words, identifiers, CLI flags and paths
(`well-formed`, `--json-response`, `safe_bulk_delete`); en dashes inside
numeric ranges in tables (`5–10`).

Worked example:

- Bad: `The job failed — the bucket had no logging enabled.`
- Good: `The job failed because the bucket had no logging enabled.`
- Good: `The job failed. The bucket had no logging enabled.`
- Good: `The job failed for one reason: the bucket had no logging enabled.`

## The other tells

**Vocabulary to avoid.** delve, leverage (as a verb), robust, seamless,
seamlessly, holistic, myriad, plethora, landscape, realm, tapestry, testament,
underscore, pivotal, crucial, vital, foster, harness, unlock, elevate, empower,
streamline, embark, journey, ever-evolving, cutting-edge, game-changer,
"in today's fast-paced world", "at the end of the day". Use the plain word a
colleague would say out loud.

**Sentence shapes to avoid.**

- "It's not just X, it's Y" and "not only X but also Y"
- The reflexive rule of three: "clear, concise and compelling"
- Opening flattery: "Great question!", "Absolutely!", "Certainly!"
- "I hope this email finds you well"
- Throat-clearing hedges: "It's worth noting that", "It's important to
  remember that", "Please note that"
- Signposting every paragraph: "Firstly... Secondly... Lastly"
- "Let's dive in", "Let me break this down"
- A closing paragraph that restates what was just said: "In conclusion",
  "Overall", "To summarise"

**Formatting tells.**

- Bullet lists where two sentences of prose would do, especially in email
- Every bullet the same length with a bolded lead-in phrase
- Bold scattered mid-sentence for emphasis
- Section headings in a short message
- Emoji in professional correspondence
- Uniform sentence length. Real writing varies; some sentences are short.

## Procedure

1. **Match the register of the thread.** Read what the other person wrote
   first. Mirror their greeting, formality, sign-off, and whether they use
   bullets at all. A reply to a terse email should not be three headed
   sections.
2. **State the point in the first two sentences.** No warm-up paragraph.
3. **Keep the user's own words.** Job names, system names, bucket paths,
   ticket refs, product terms and numbers stay exactly as they wrote them.
   Do not upgrade their vocabulary.
4. **Vary sentence length on purpose.** Include at least one short sentence.
5. **Cut every hedge that does not change the meaning.**
6. **Prefer concrete nouns and numbers** over abstractions.
7. **One closing line.** Not a summary.

## Self-check before delivering

Run this on any file you wrote or edited:

```bash
grep -n -e '—' -e ' – ' <file>     # must return nothing
```

For chat and email text, check by eye:

- Zero em dashes, including inside quoted material you rewrote.
- Does the first sentence carry the point?
- Any word from the vocabulary list above? Replace it.
- Three consecutive sentences of the same length? Break one up.
- Would a busy colleague read it to the end? If not, cut a third.

## What humanizing is not

Do not fake humanity with deliberate typos, slang, filler, false enthusiasm or
forced informality. The goal is prose that reads as though a competent person
wrote it quickly and meant it. Clear and plain still beats casual.
