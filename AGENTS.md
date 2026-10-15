# AGENTS.md — Development Guidelines & System Architecture

This document serves as the single source of truth for AI agents and human contributors working on the **Excel to Word Mapper** codebase. Follow these conventions and architectural principles to ensure stable, maintainable, and regression-free development.

---

## 1. Project Overview & Architecture

**Excel to Word Mapper** is a client-side web application built with **Next.js 16 (App Router)** and **TypeScript**. It enables users to upload Excel/CSV spreadsheets and Word (`.docx`) templates containing `{placeholders}`, visually map columns to fields, and batch-generate customized documents (or single preview records) bundled as `.docx` or `.zip`.

### Core Technical Stack
- **Framework**: Next.js 16 (Turbopack, App Router, TypeScript)
- **Styling**: Vanilla CSS Modules (`page.module.css`) + Curated Design Tokens (`globals.css`)
- **Excel Parsing**: SheetJS (`xlsx`)
- **Document Generation**: `docxtemplater` + `pizzip`
- **ZIP Bundling**: `jszip`
- **Icons**: `lucide-react`
- **Deployment Target**: Vercel / Static Web Server

### Strict Privacy Guarantee
- **100% Client-Side Processing**: All parsing, placeholder extraction, and document templating execute entirely in the user's browser memory via JavaScript `ArrayBuffer` and `Blob` APIs.
- **No Remote Document Transmission**: Never introduce backend endpoints or third-party APIs that upload or persist user document contents.

---

## 2. Directory & File Map

```text
excel-to-word-mapper/
├── AGENTS.md                    # This architecture and coding guidelines guide
├── README.md                    # User-facing documentation & feature overview
├── next.config.ts               # Next.js config with security & attachment headers
├── package.json                 # Dependencies and npm scripts
├── public/
│   └── examples/                # Static example candidate Excel & offer letter template
│       ├── sample_candidates.xlsx
│       └── sample_offer_letter.docx
├── src/
│   ├── app/
│   │   ├── globals.css          # Executive Dark cyber design system & mobile tap rules
│   │   ├── layout.tsx           # Root HTML layout, SEO metadata, viewport configs
│   │   ├── page.module.css      # Component styles, CSS Grid, mobile media queries
│   │   └── page.tsx             # 3-Step Wizard SPA component & reactive state machine
│   └── lib/
│       ├── excelParser.ts       # SheetJS reader, multi-sheet handler, sample generator
│       ├── matcher.ts           # Fuzzy auto-matching algorithm (Levenshtein + heuristics)
│       └── wordProcessor.ts     # XML normalizer, placeholder extractor, docxtemplater engine,
│                                # bulletproof triggerBrowserDownload, and JSZip batcher
```

---

## 3. Core Architectural Rules & Implementation Details

### A. Bulletproof Browser Downloads (`triggerBrowserDownload`)
Never replace `triggerBrowserDownload` in [`src/lib/wordProcessor.ts`](file:///c:/Users/PC/Downloads/git/excel-to-word-mapper/src/lib/wordProcessor.ts) with a naive `window.location.href` or simple `<a>` click.

**Download Standards**:
1. **Extension Normalization**:
   - Always strip existing extensions (`/\.[a-zA-Z0-9]+$/i`) from base filenames before sanitizing to prevent double-extensions like `Document.docs.docx`.
   - Enforce exact final extensions (`.docx`, `.xlsx`, `.zip`).
2. **Explicit MIME Types**:
   - Word: `application/vnd.openxmlformats-officedocument.wordprocessingml.document`
   - Excel: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
   - ZIP: `application/zip`
3. **`File` Object Binding**: Construct a `new File([typedBlob], filename, { type: targetMime })` so Chromium, WebKit, and Gecko internally register the exact filename.
4. **Mobile Web Share API**: On touch devices (Android Chrome, iOS Safari), check `navigator.share({ files: [file] })` to enable native OS "Save to Downloads" or "Open with Word".
5. **Memory Safety**: Delay `URL.revokeObjectURL(url)` by at least 60 seconds so browser download managers finish writing to disk before memory release.

### B. Word Template XML Split Handling (`normalizeDocxZip`)
Microsoft Word frequently fractures text across separate XML tags:
```xml
<!-- Example of MS Word split tag issue -->
<w:t>{</w:t><w:t>Candidate</w:t><w:t>_Name}</w:t>
```
- [`extractPlaceholdersFromDocx`](file:///c:/Users/PC/Downloads/git/excel-to-word-mapper/src/lib/wordProcessor.ts) strips inner XML tags to reconstruct continuous text streams before regex matching.
- [`normalizeDocxZip`](file:///c:/Users/PC/Downloads/git/excel-to-word-mapper/src/lib/wordProcessor.ts) normalizes Mail Merge guillemets (`«placeholder»` or `&laquo;placeholder&raquo;`) and spaces (`{ name }`) into standard `{name}` syntax before docxtemplater processes the document.

### C. Responsive UI & Mobile Standards
The application must maintain zero horizontal scroll and fluid touch-readiness across 3 explicit breakpoints:
1. **Tablet (`@media (max-width: 900px)`)**:
   - Wizard nav hides step titles, showing only step badge numbers.
   - 2-column upload grids collapse to a single column.
   - Mapping grid columns stack into vertical cards.
2. **Small Tablet / Large Phone (`@media (max-width: 640px)`)**:
   - Interactive buttons have minimum 44px tap targets (Apple HIG standard).
   - Form inputs maintain `16px` font size to prevent iOS Safari auto-zoom.
3. **Smartphone (`@media (max-width: 480px)`)**:
   - Container padding narrows to `10px-12px`.
   - Modals and toast alerts use `inset: auto 12px 12px` and `clamp()` padding.
   - Safe-area insets (`env(safe-area-inset-bottom)`) are respected.

### D. Security & Resource Limits
- **Upload Quotas**:
  - Excel files (`.xlsx`, `.xls`, `.csv`): **50 MB** maximum.
  - Word templates (`.docx`): **25 MB** maximum.
- **Filename Sanitization**:
  - Filter out null bytes (`\0`), path traversal (`..`), and forbidden OS characters (`/ \ ? % * : | " < >`).
  - Truncate base filenames to 180 characters max to prevent filesystem buffer overflows.
- **Security Headers**:
  - Maintain `X-Frame-Options: DENY`, `X-Content-Type-Options: nosniff`, and `Referrer-Policy` in `next.config.ts`.

---

## 4. Development Workflow

### Commands
```bash
# Run local development server (Turbopack)
npm run dev

# Run production build & verify TypeScript/Linting
npm run build

# Start production server
npm run start
```

### Pre-Commit Checklist for Future Changes
Before committing any changes to the codebase:
1. **Compile**: Run `npm run build` and ensure exit code `0` with zero TypeScript errors.
2. **Download Integrity**: Verify that single `.docx`, batch `.zip`, and sample `.xlsx` downloads save with full names and valid file extensions.
3. **Responsive Verification**: Ensure all modifications stack cleanly on 375px mobile viewports without horizontal overflow.
4. **Git State**: Confirm `.gitignore` excludes `node_modules/`, `.next/`, `.env*`, and temporary build artifacts.
