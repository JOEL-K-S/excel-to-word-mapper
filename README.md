# Excel to Word Mapper

> **Automated In-Browser Document Generation & Bulk Mail Merge Alternative**  
> Map Excel spreadsheet columns to Word templates and export personalized documents as **Word (.docx)** or **PDF (.pdf)** in batch ZIP archives. 100% client-side with zero cloud server uploads.

[![Next.js](https://img.shields.io/badge/Next.js-16.4-black?logo=next.js)](https://nextjs.org)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.x-blue?logo=typescript)](https://www.typescriptlang.org)
[![CI](https://github.com/JOEL-K-S/excel-to-word-mapper/actions/workflows/ci.yml/badge.svg)](https://github.com/JOEL-K-S/excel-to-word-mapper/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-emerald.svg)](LICENSE)
[![Privacy: 100% Client-Side](https://img.shields.io/badge/Privacy-100%25%20Client--Side-indigo.svg)](#privacy--security-guarantee)

---

## 🌟 Why Excel to Word Mapper?

Traditional Microsoft Word Mail Merge is slow, requires a desktop Microsoft 365 license, cannot easily split records into individual files with custom names without complex VBA macros, and lacks direct in-browser PDF conversion.

**Excel to Word Mapper solves this completely in your browser:**

| Feature | Microsoft Word Mail Merge | Cloud Services (Zapier / APIs) | **Excel to Word Mapper** |
| :--- | :---: | :---: | :---: |
| **Privacy / Zero Data Upload** | ⚠️ Desktop only | ❌ Files sent to 3rd party servers | ✅ **100% Local In-Browser** |
| **Split into Separate Files** | ❌ Complex VBA required | ✅ (Paid paywalls) | ✅ **Built-in 1-Click ZIP** |
| **Word (.docx) & PDF Export** | ⚠️ Manual save per file | ✅ (Credits/Monthly fee) | ✅ **Instant Dual Export** |
| **Dynamic Custom Filenames** | ❌ Not native | ⚠️ Configuration heavy | ✅ **`{Column}` Tag Interpolation** |
| **Microsoft Office Required** | ❌ Yes (desktop license) | ❌ N/A | ✅ **No Office Required** |
| **Pricing** | Paid Office License | Recurring Subscription | ✅ **100% Free & Open Source** |

---

## 🚀 Key Features

- 🔒 **100% Client-Side Privacy**: Runs purely in WebAssembly and browser memory. No spreadsheets, customer records, or templates are ever sent to an external server. Compliant with GDPR, HIPAA, and internal confidentiality requirements.
- 📄 **Dual Output Formats**: Choose between fully editable Microsoft Word documents (`.docx`) and print-ready Adobe PDF (`.pdf`) files.
- ⚡ **Dynamic Filename Tagging**: Define output file names like `Invoice_{Client_Name}_{Invoice_No}.pdf` using clickable field chips.
- 🎯 **Intelligent Field Auto-Matching**: Fuzzy string matching connects Word `{tags}` with Excel column headers automatically.
- 📦 **Batch ZIP Export with Progress**: Generate hundreds or thousands of documents with real-time percentage progress and instant ZIP download.
- 🔍 **Interactive Single Record Preview**: Test document generation on any individual row before initiating a full batch export.
- 🎨 **Format & Style Fidelity**: Retains all Word document formatting, tables, typography, bullet points, headers, footers, and logos.
- 🤖 **AI & Search Engine Optimized**: Includes structured Schema.org JSON-LD (`WebApplication`, `HowTo`, `FAQPage`), `sitemap.xml`, `robots.txt`, and [`llms.txt`](public/llms.txt) for modern AI models (ChatGPT Search, Perplexity, Gemini, Claude).

---

## 📋 Common Use Cases

- 🎓 **Certificates & Diplomas**: Batch generate course completion certificates, awards, and credentials from attendee rosters.
- 🧾 **Invoices & Receipts**: Automate billing statements, purchase orders, and quotes from spreadsheet line items.
- 💼 **HR & Employment Contracts**: Produce employee offer letters, NDAs, compensation statements, and performance appraisals.
- 📊 **Academic Report Cards**: Distribute personalized student evaluations and transcripts without student data leaks.
- 🏠 **Real Estate Agreements**: Bulk generate leases, tenancy notices, and purchase addendums.
- 🏷️ **Event Badges & Tickets**: Produce personalized conference attendee passes, name tags, and registration confirmations.

---

## 🛠️ Step-by-Step Workflow

1. **Step 1: Upload Files**
   - Upload your Excel sheet (`.xlsx`, `.xls`, or `.csv`) and Microsoft Word template (`.docx` containing `{tag}` placeholders).
   - Alternatively, click **"Try with Sample Files"** to load example files instantly.
2. **Step 2: Map Fields**
   - Click **"Auto-Match"** to automatically pair placeholders with column headers, or select mappings and fallback defaults manually.
3. **Step 3: Choose Format & Export**
   - Select **Microsoft Word (.docx)** or **Adobe PDF (.pdf)**.
   - Configure your dynamic filename pattern.
   - Download a single test record or click **Generate & Download ZIP** to package all documents.

---

## 📚 Documentation

Quick links to guides, privacy specifications, and technical standards:

| I want to | Read |
|---|---|
| Learn how to format `{tags}` & spreadsheets | [Template & Workflow Guide](#-step-by-step-workflow) |
| Review the 100% zero-server privacy guarantee | [Privacy Policy](src/app/privacy/page.tsx) / [/privacy](https://excel-to-word-mapper.vercel.app/privacy) |
| Inspect security architecture & report vulnerabilities | [Security Policy](SECURITY.md) |
| Read the AI search & LLM ingestion manifest | [AI Manifest (`llms.txt`)](public/llms.txt) |
| Review agent rules & architectural invariants | [Engineering Standards](AGENTS.md) |
| Understand open-source license & permissions | [MIT License](LICENSE) |

---

## 💻 Local Development

```bash
# Clone the repository
git clone https://github.com/JOEL-K-S/excel-to-word-mapper.git
cd excel-to-word-mapper

# Install dependencies
npm install

# Start development server
npm run dev
```

Open [http://localhost:3000](http://localhost:3000) in your browser.

### Production Build

```bash
npm run build
npm run start
```

---

## 🔒 Privacy & Security Guarantee

Excel to Word Mapper is built with a zero-telemetry architecture:
- Spreadsheets are parsed locally via SheetJS (`xlsx`).
- Word templates are populated in-memory via `docxtemplater` & `pizzip`.
- PDF conversion is performed in-browser using `docx-preview`, `jspdf`, and `html2canvas`.
- Archives are zipped in-browser via `jszip` and downloaded directly via Blob URLs.

No analytics, cookies, tracking scripts, or backend servers are used.

---

## 📄 License

Distributed under the MIT License. See [LICENSE](LICENSE) for more information.

Author: [JOEL-K-S](https://github.com/JOEL-K-S)
