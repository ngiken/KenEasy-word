# CLAUDE.md — KenEasy Word Converter

Browser-only document to Word tool: drop ODT, FODT, RTF, HTML, Markdown, or text → reorder → choose a layout & preset → export to Microsoft Word (`.docx`) or plain text (`.txt`) separately or merged. No backend, upload, install, or account.

## Layered architecture

```text
web/
  index.html       # semantic UI shell only
  styles.css       # adaptive Apple-inspired light/dark design system
  config.js        # data layer: limits, input rules, page sizes, typography, presets, defaults
  i18n.js          # content layer: bilingual copy + translation service
  doc-engine.js    # domain layer: ODT/FODT/RTF/HTML/MD/TXT parsers, OpenXML DOCX generator
  app.js           # application layer: state, persistence, queue, DOM orchestration
  vendor/          # pinned offline JSZip + SortableJS + NOTICE.txt
docs/screenshots/  # README screenshots and animated tour
scratch/e2e/       # real-browser regression and screenshot tooling
```

Dependency direction is one-way: `config → i18n/engine → app → DOM`.

Do not hard-code a new input rule, page size, typography, preset, limit, default, or storage key inside an event handler. Put rules in `config.js`, user-facing copy in `i18n.js`, document parsing and OpenXML generation in `doc-engine.js`, and workflow coordination in `app.js`.

## Product goals

1. Open a link and it works — static hosting, no install or account.
2. Local only — files never leave the browser tab (`connect-src 'none'`).
3. Universal document flow — import ODT, FODT, RTF, HTML, Markdown, and text, reorder, choose recipe, generate Word (.docx) or Text.
4. Offline-capable — all runtime dependencies are vendored under `web/vendor/`.
5. Stable and extensible — rules are data-driven and layers remain decoupled.
6. Accessible and adaptive — keyboard sorting, semantic labels, reduced motion, mobile layout, and equivalent light/dark experiences.

## Hard rules

1. Never commit secrets: `*.pem`, `*.p12`, `*.key`, `.env`, credentials.
2. Never copy signing keys or raw recordings from the parent workspace.
3. This directory is its own repository (`https://github.com/ngiken/KenEasy-word.git`); never commit it through the parent `D:\code` repository.
4. No backend, file-upload API, runtime CDN, analytics, or tracking.
5. Preserve the dependency direction and the Layering / Decoupling / Rule-based / Data-driven design philosophy.
6. Vendor pins are intentional; bump through `scratch/fetch-vendor.ps1`, update `NOTICE.txt`, and rerun E2E.
7. Preserve the adaptive design system: neutral canvas, blue primary action, restrained KenEasy pink/cyan accents, and equivalent light/dark behavior.
8. README screenshots must come from the real app through `npm run capture:readme`, not hand-made mockups.

## Conversion rules

| Concern | Rule |
| --- | --- |
| ODT / FODT | XML parser: headings, paragraphs, styled runs (bold/italic/underline/color), lists, tables, embedded images |
| RTF | State machine tokenizer: control words, Unicode escapes, bold/italic/underline, alignments |
| HTML | DOMParser sanitization: h1-h6, p, blockquote, ul/ol, table, links, code blocks |
| Markdown | Custom line-oriented lexer: headers, bold/italic, unordered/ordered lists, tables, blockquotes |
| Plain Text | Preserves line breaks, detects headings and section breaks |
| DOCX Generation | Standards-compliant ECMA-376 OpenXML structure (`[Content_Types].xml`, `_rels`, `word/document.xml`, `word/styles.xml`, `word/media/`), packaged via JSZip |
| Delivery | Separate downloads, or merged single document with clean page breaks |
| Safety | 50 MB/file, 100 files max |
| Offline | Strict CSP `connect-src 'none'`; no external network requests |

## Testing

```powershell
cd scratch/e2e
npm test                 # expect: 32/32 checks passed + ALL GREEN
npm run capture:readme   # rebuild docs/screenshots/*.png and quick-tour.gif
```

The suite uses real Chromium and verifies actual downloaded DOCX/TXT files, OpenXML structure validation, embedded image preservation, merge and separate delivery, presets, persistence, keyboard ordering, i18n, appearance, responsive layout, and offline behavior.

## Safe change checklist

```text
[ ] git status contains no secrets or unexpected binaries
[ ] new product rules live in config.js, not UI handlers
[ ] new copy exists in both zh and en
[ ] vendor files and NOTICE are complete
[ ] index.html versioned assets match the release version
[ ] node --check passes for every first-party JS file
[ ] npm test reports ALL GREEN
[ ] README visuals regenerated and visually inspected when UI changes
[ ] README badges/changelog match the release version
```

## Local preview

```powershell
python -m http.server 5173 --directory web
```

Open <http://localhost:5173/>. Ship or host the `web/` directory only.
