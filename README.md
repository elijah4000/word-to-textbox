# Word → Text Box

Converts a Word `.docx` into rich text that keeps its formatting when pasted
into a Brightspace text-submission box.

- **Runs entirely in the browser.** No backend, no database, no analytics, no
  uploads. The production build sets a Content-Security-Policy with
  `connect-src 'none'`, so the page cannot send the file anywhere.
- **No AI.** Conversion is done by [mammoth.js](https://github.com/mwilliamson/mammoth.js)
  plus deterministic post-processing.
- **Never changes wording.** No spellcheck, no smart-quote changes, no
  whitespace cleanup. The only text the tool changes is text mammoth itself
  adds: footnote markers (`[1]` → superscript `1`) and footnote back-links (`↑`).

## Integrity features

**Text check.** The original text comes from `mammoth.extractRawText`. The final text
is extracted from the exact HTML that gets copied. Both are normalized the same
way (runs of whitespace collapse to one space, nothing else), then compared word by
word. Footnotes and endnotes are compared one by one. The check always runs;
when the text matches nothing is shown, and when it differs a panel opens with
removed and added words highlighted.

**Content warnings.** The raw Word XML is scanned for things that can't be
converted cleanly, and each one is listed with a count:

| Item | What happens |
| --- | --- |
| Equations, charts, SmartArt, embedded objects | **Not converted.** Missing from the output |
| Tracked changes | **Not converted.** Shown as if all changes were accepted |
| Text boxes | Text placed inline, box layout lost |
| Comments, headers, footers | Not included |
| Images | Embedded as data URIs, or left out if embedding is off |
| Footnotes, endnotes | Listed at the end, numbered in reference order |

If the text differs, or anything is "Not converted", the Copy button stays
disabled until the user ticks "I have reviewed…".

Note: the text check can only compare what mammoth reads. Content mammoth
can't read at all (e.g. equations) is missing from *both* sides. That is why the
warning panel scans the XML separately.

## Formatting options

All options are applied as inline styles, because Brightspace's editor drops
stylesheets and classes:

- Double spacing (`line-height: 2`)
- Hanging indent (0.5 in). It applies to paragraphs that have a hanging indent in
  Word, and to every paragraph after a "Works Cited" / "References" /
  "Bibliography" heading
- Font family and size
- Image embedding, with a warning when images exceed 1 MB

Tables always get inline borders so they stay visible after pasting.

## Copying

Uses `navigator.clipboard.write` with a `ClipboardItem` holding both `text/html`
and `text/plain`. If that is unavailable or blocked, the tool selects the rendered
preview and uses `document.execCommand("copy")` instead.

## Development

```bash
npm install
npm run dev          # http://localhost:5173
npm test             # unit tests (text check, detection, full pipeline on the samples)
npm run build        # static site in dist/
npm run make-sample  # regenerate samples/*.docx
```

`samples/sample.docx` has headings, bullet and numbered lists, a table,
two footnotes, a hyperlink, a tab, a line break, smart quotes and dashes, an image
and a hanging-indent Works Cited list. `samples/problems.docx` has an
equation, tracked changes, a comment, a header, a footer and a text box.
Results should always be previewed in the text box before submitting.
students should follow their instructor's submission rules.
