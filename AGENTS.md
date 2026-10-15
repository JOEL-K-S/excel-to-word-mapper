# AGENTS.md — Development Guidelines & System Architecture

Single source of truth for AI agents and human contributors on **Excel to Word Mapper**. Read this fully before changing code. If a request conflicts with a rule here, stop and flag it instead of working around it.

**Rule levels:** **MUST / MUST NOT** = hard invariant (never break without explicit approval). **SHOULD** = default; deviate only with a stated reason. **MAY** = optional.

---

## 1. Project Overview

Client-side web app (Next.js 16 App Router + TypeScript). Users upload an Excel/CSV file and a Word `.docx` template containing `{placeholders}`, map columns to fields visually, then generate a single preview document or a batch bundled as `.docx` / `.zip`.

### Tech Stack

| Concern | Choice |
|---|---|
| Framework | Next.js 16 (Turbopack, App Router, TypeScript) |
| Styling | Vanilla CSS Modules (`page.module.css`) + design tokens (`globals.css`) |
| Excel parsing | SheetJS (`xlsx`) |
| Document generation | `docxtemplater` + `pizzip` |
| ZIP bundling | `jszip` |
| Icons | `lucide-react` |
| Deploy target | Vercel / static web server |

### Privacy Guarantee (Non-Negotiable)

- All parsing, placeholder extraction and templating MUST run in the browser (`ArrayBuffer` / `Blob`).
- MUST NOT add backend endpoints, third-party APIs, analytics, or logging that transmit or persist user document contents or spreadsheet data.
- MUST NOT log file contents or cell values to the console in committed code.
- New dependencies MUST NOT make network requests with user data. Justify every new dependency in the PR/commit message.

---

## 2. Directory Map

```text
excel-to-word-mapper/
├── AGENTS.md                 # This file
├── README.md                 # User-facing docs & feature overview
├── next.config.ts            # Next.js config: security & attachment headers
├── package.json              # Dependencies and scripts
├── public/examples/          # sample_candidates.xlsx, sample_offer_letter.docx
└── src/
    ├── app/
    │   ├── globals.css       # Design system tokens, mobile tap rules
    │   ├── layout.tsx        # Root layout, SEO metadata, viewport
    │   ├── page.module.css   # Component styles, grid, media queries
    │   └── page.tsx          # 3-step wizard SPA + state machine
    └── lib/
        ├── excelParser.ts    # SheetJS reader, multi-sheet handling, sample generator
        ├── matcher.ts        # Fuzzy auto-match (Levenshtein + heuristics)
        └── wordProcessor.ts  # XML normalizer, placeholder extractor, docxtemplater
                              # engine, triggerBrowserDownload, JSZip batcher
```

**Layering rule:** `src/lib/*` holds pure logic (no React, no DOM except where noted in `wordProcessor.ts` download helper). `src/app/*` holds UI and state only. Do not move business logic into components.

---

## 3. Core Architectural Rules

### A. Browser Downloads — `triggerBrowserDownload` (`src/lib/wordProcessor.ts`)

MUST NOT replace it with `window.location.href` or a bare `<a>` click. Required behaviour:

1. **Extension normalization** — strip any existing extension (`/\.[a-zA-Z0-9]+$/i`) from the base name *before* sanitizing, then enforce the exact final extension (`.docx`, `.xlsx`, `.zip`). Prevents `Document.docs.docx`.
2. **Explicit MIME types**
   - Word: `application/vnd.openxmlformats-officedocument.wordprocessingml.document`
   - Excel: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
   - ZIP: `application/zip`
3. **`File` binding** — build `new File([typedBlob], filename, { type: targetMime })` so Chromium, WebKit and Gecko register the exact filename.
4. **Mobile Web Share** — on touch devices (Android Chrome, iOS Safari) use `navigator.share({ files: [file] })` when supported, for native "Save to Downloads" / "Open with Word". MUST fall back to the standard anchor download when unsupported or when the user cancels.
5. **Memory safety** — delay `URL.revokeObjectURL(url)` by **≥ 60 s**.

### B. Word Template XML Handling

Word fractures text across runs, e.g. `<w:t>{</w:t><w:t>Candidate</w:t><w:t>_Name}</w:t>`.

- `extractPlaceholdersFromDocx` strips inner XML tags to rebuild continuous text before regex matching.
- `normalizeDocxZip` converts mail-merge guillemets (`«name»`, `&laquo;name&raquo;`) and padded placeholders (`{ name }`) to `{name}` before docxtemplater runs.
- Any change to these functions MUST be verified against: split-run placeholders, guillemet placeholders, padded placeholders, and placeholders in headers/footers/tables.

### C. Responsive UI Standards

Zero horizontal scroll and touch-ready at three breakpoints:

| Breakpoint | Requirements |
|---|---|
| `max-width: 900px` | Wizard nav shows step badges only; 2-col upload grids → 1 col; mapping grid stacks into cards |
| `max-width: 640px` | Buttons ≥ **44px** tap target; inputs **16px** font (prevents iOS auto-zoom) |
| `max-width: 480px` | Container padding 10–12px; modals/toasts use `inset: auto 12px 12px` and `clamp()` padding; respect `env(safe-area-inset-bottom)` |

### D. Security & Resource Limits

- **Upload limits:** Excel (`.xlsx`, `.xls`, `.csv`) ≤ **50 MB**; Word (`.docx`) ≤ **25 MB**. Validate before parsing.
- **Filename sanitization:** remove null bytes (`\0`), `..`, and `/ \ ? % * : | " < >`; truncate base name to **180 chars**.
- **Security headers** in `next.config.ts` MUST remain: `X-Frame-Options: DENY`, `X-Content-Type-Options: nosniff`, `Referrer-Policy`.
- Treat all uploaded content as untrusted: never `eval`, never inject spreadsheet/template values via `dangerouslySetInnerHTML`.

---

## 4. Coding Conventions

- **TypeScript:** strict mode; no `any` (use `unknown` + narrowing); no `@ts-ignore` without a comment explaining why; export explicit types for public `lib` functions.
- **Styling:** CSS Modules and tokens from `globals.css` only. No new colors/spacing values inline — add or reuse a token. No CSS frameworks.
- **State:** keep the wizard as an explicit state machine in `page.tsx`; avoid scattered booleans that can represent impossible states.
- **Errors:** every parse/generate path MUST surface a user-readable message (toast/modal) and never fail silently or crash the wizard.
- **Dependencies:** SHOULD prefer the existing stack. Do not add a library for something achievable in ~20 lines. Pin versions via lockfile; do not upgrade major versions as a side effect of another task.
- **Naming:** `camelCase` functions/vars, `PascalCase` components/types, `kebab-case` for non-component files only if introduced (existing files keep current names).
- **Comments:** explain *why*, not *what*. Document any non-obvious browser workaround (especially download/share quirks).

---

## 5. AI-Agent Working Protocol ("Stable Vibe Coding")

1. **Read before writing.** Open the files you will touch and their direct callers. Do not edit from memory of the file map.
2. **Smallest viable diff.** Change only what the task needs. No drive-by refactors, renames, reformatting, or dependency bumps.
3. **One concern per change.** Feature, fix, and refactor go in separate commits.
4. **Protect invariants.** Before editing `wordProcessor.ts` or `next.config.ts`, re-read Section 3 A, B, D.
5. **No hallucinated APIs.** If unsure about a Next.js 16, SheetJS, or docxtemplater API, check the installed package's types/docs in `node_modules` rather than guessing.
6. **Don't delete to make it pass.** Never remove validation, error handling, or a rule in this file to resolve a build error or test failure; fix the cause.
7. **State assumptions.** If the request is ambiguous, list assumptions up front (or ask) instead of silently choosing.
8. **Update docs with code.** Behaviour, limits, or architecture changes MUST update this file and `README.md` in the same change.
9. **Finish with a report:** what changed, which files, how it was verified, and anything left unverified.

### Do / Don't

| Do | Don't |
|---|---|
| Keep all processing client-side | Add API routes, server actions, or uploads for documents |
| Reuse `triggerBrowserDownload` | Add ad-hoc `<a download>` code elsewhere |
| Use design tokens | Hard-code colors/sizes |
| Test on a 375px viewport | Ship layout changes verified only on desktop |
| Keep `lib/` pure and testable | Import React/DOM into `lib/` logic |

### Prompt Template for Requests to an Agent

```text
Goal: <one sentence outcome>
Scope: <files/areas allowed to change>
Constraints: follow AGENTS.md; no new dependencies unless justified
Acceptance: <observable checks, e.g. "batch .zip downloads with full filenames">
```

---

## 6. Development Workflow

### Commands

```bash
npm run dev      # Local dev server (Turbopack)
npm run build    # Production build + TypeScript check (must exit 0)
npm run start    # Serve production build
```

If linting or tests are configured in `package.json` (e.g. `npm run lint`, `npm test`), they are part of the checklist below. When adding tests, SHOULD start with `src/lib/matcher.ts` and `wordProcessor.ts` (pure logic, highest regression risk).

### Definition of Done (Pre-Commit Checklist)

- [ ] `npm run build` exits `0` with zero TypeScript errors (and lint/tests pass, if configured)
- [ ] **Download integrity:** single `.docx`, batch `.zip`, and sample `.xlsx` save with full names and valid extensions (no double extensions)
- [ ] **Template handling:** split-run, guillemet, and padded placeholders still resolve
- [ ] **Responsive:** 375px viewport, no horizontal overflow; 44px tap targets and 16px inputs at ≤ 640px
- [ ] **Limits & errors:** oversized, wrong-type, and corrupt files show a clear message; app stays usable
- [ ] **Privacy:** no new network calls carrying user data (check the browser Network tab)
- [ ] **Git state:** `.gitignore` excludes `node_modules/`, `.next/`, `.env*`, build artifacts; no secrets or sample real data committed
- [ ] This file and `README.md` updated if behaviour or architecture changed

### Commit Conventions

Use Conventional Commits: `feat:`, `fix:`, `refactor:`, `docs:`, `style:`, `chore:`, `test:`. Subject ≤ 72 chars, imperative mood. Body explains *why* and notes verification performed.

---

## 7. Known Pitfalls

- **iOS Safari:** inputs below 16px trigger auto-zoom; `navigator.share` requires a user gesture and may reject on cancel — handle both.
- **Revoking object URLs early** truncates or cancels large downloads.
- **Word placeholders** are routinely split across runs; never regex raw XML without normalizing first.
- **Large batches** can exhaust browser memory; keep generation incremental and release references (Blob URLs, buffers) when done.
- **Multi-sheet workbooks:** always route through `excelParser.ts`; do not assume the first sheet is the intended one.

---

## 8. Changelog of This Document

Keep a one-line entry per meaningful rule change (date — change — reason) below.

- _Initial standardised version._