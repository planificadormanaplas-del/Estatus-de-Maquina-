# AI tells to avoid in personal-voice emails

Distilled from Wikipedia's *Signs of AI writing* guide and adapted to the email context specifically. Many of the patterns in that guide come up in long-form/encyclopedic writing and don't apply to short emails — this file keeps only the ones that actually show up in email drafting.

When drafting, scan for these and remove them. Most of the time, the fix is to delete the sentence rather than rewrite it.

---

## Vocabulary tells

LLMs over-use a specific set of words far more than humans do. None of these are wrong on their own, but together they form an unmistakable pattern. In an email, even one or two of these can read as AI.

**Verbs:** delve, leverage, navigate, ensure, foster, cultivate, streamline, underscore, showcase, highlight (as a verb), align with, garner, encompass, bolster

**Adjectives:** robust, seamless, holistic, comprehensive, vibrant, intricate, pivotal, crucial, valuable, key, enduring, meticulous

**Phrases:** "valuable insights," "key takeaways," "play a vital role," "in today's landscape," "in the realm of," "stands as," "serves as," "is a testament to"

**Email-specific filler that signals AI** (almost always wrong in personal voice):
- "I hope this email finds you well"
- "I hope you're doing well"
- "I trust this message finds you well"
- "Just wanted to follow up"
- "Just wanted to circle back"
- "I am writing to..."
- "Please feel free to..."
- "Please don't hesitate to..."
- "I'd be more than happy to..." (most personal voices say "happy to..." without the "I'd be")
- "I look forward to your prompt response"
- "Your prompt attention is appreciated"
- "Thank you for your patience and understanding"
- "Thank you for your time and consideration"

Cut these on sight.

**Phrases that LOOK like AI tells but might ACTUALLY be in the user's voice — don't strip these without checking:**

These are warm, in-voice phrases for Register B (sales inquiries, post-meeting follow-ups, business-development context). They might appear on naive AI-tell lists, but real people often use them, and the skill should preserve them when the user actually writes this way.

- `Thanks for reaching out!` — for cold inquiries
- `Great meeting you both!` / `Great meeting you!` — post-meeting opener
- `Let me know if you have any other questions!` — open-door close
- `happy to answer any additional questions you have` — open-door close
- `at your convenience` — courteous Register B touch
- `would work well for [your team / your situation]` — specifics-flavored warmth

The reason these aren't tells: AI generates them as part of a *cluster* of formulaic phrases. Real people use them singly, embedded in otherwise specific and direct prose. The presence of one of these is fine. The presence of three of them stacked together (e.g., "Thanks for reaching out! I hope this finds you well. I'd be more than happy to assist with your inquiry.") is the AI tell.

Check the user's `samples.md` to see which of these (if any) they actually use.

---

## Structural tells

**Negative parallelism.** "It's not just X — it's Y." "This isn't about X, it's about Y." This is one of the most reliable AI signatures. Almost never appears in personal-voice writing.

**Rule of three.** Lists of three parallel adjectives or noun phrases ("clear, concise, and actionable" / "innovative, scalable, and impactful"). LLMs reach for triplets constantly. Real personal writing tends toward specific singular things.

**Em dash for punched-up emphasis.** Em dashes are fine as ordinary punctuation, but the *formulaic* use — clause break, dramatic restatement, summing up — is a strong AI tell. Compare:
- AI-shaped: "We need to act now — not tomorrow, not next week, but today."
- Personal-shaped: just say "We need to act now."

**Bulleted lists with bold inline headers.** "**Background:** ... **Next steps:** ... **Timeline:** ..." This is a slide-deck format pretending to be an email. If three points genuinely need separation, use three short paragraphs instead.

**Title Case Headings inside emails.** Almost always wrong in email body text. If a heading is needed at all (rare), sentence case.

**Excessive boldface for emphasis.** Especially on the same word every time it appears, or on opening words of bullets. Plain text is almost always better.

**"In summary" / "In conclusion" / "Overall" closings.** A short email doesn't need a wrap-up sentence. The last substantive sentence IS the closing. Just stop writing.

**Closing with an unnecessary question — register-dependent.** In Register A (transactional), tacking on "Does that work for you?" when the next move is obvious is filler. In Register B (sales/business-dev), an open-door warm closer like "Let me know if you have any other questions!" is exactly right and IS in voice. Don't strip a Register B closer just because it looks like the rule says to.

---

## Tone tells

**Excess enthusiasm.** "Awesome!" "Fantastic!" "Wonderful news!" Most personal voices can be warm without performing excitement.

**Performative gratitude.** "Thank you SO much for..." "I really, really appreciate..." A single "Thanks!" usually carries the same weight without the theater.

**Restating what the other person said.** "So just to confirm, you're saying X, and you'd like Y by Z..." LLMs do this to demonstrate comprehension. Real writers just answer.

**Apologizing reflexively.** "Apologies for the delay" / "Sorry for any confusion" when neither is real. Skip it unless there's an actual thing to apologize for.

**Reassurance the recipient didn't ask for.** "Please rest assured that..." "I want to assure you..." If they didn't raise a worry, don't manufacture one to soothe.

---

## Formatting tells

**Curly quotation marks (" " ' ').** ChatGPT outputs these by default; Claude generally doesn't, but worth checking. Use straight quotes (" " ' ') in email drafts unless the user specifically wants otherwise.

**Markdown formatting that wasn't asked for.** No `**bold**`, no `# headers`, no bullet points unless the content genuinely calls for them. Plain prose by default.

---

## The single best test

Read the draft out loud. If any sentence feels like it could've been written by a moderately competent corporate communications consultant for any executive at any company, that sentence is wrong. Personal voice lives in specificity. The draft should sound like the user *talking*, not like a polished generic professional.
