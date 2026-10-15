# Security Policy

At **Excel to Word Mapper**, the privacy and security of your data is our foundational design principle. This policy outlines our architectural security guarantees, supported versions, and procedures for reporting potential security vulnerabilities.

---

## 🔒 Architectural Security Guarantees

Excel to Word Mapper is built with a strict **zero-telemetry, zero-server-upload architecture**:

1. **100% Client-Side In-Memory Execution**:
   - All parsing of Excel workbooks (`.xlsx`, `.xls`, `.csv`), Word document templates (`.docx`), PDF rendering, and ZIP archive packaging run strictly inside your local web browser's memory (RAM).
   - **Zero Document Data Transmission**: At no point are your spreadsheets, confidential records, document templates, or generated files sent to any external server, database, API, or third-party cloud service.

2. **Zero Telemetry & Tracking**:
   - The application does not incorporate analytics trackers, telemetry SDKs, session replay tools, or tracking cookies.
   - No spreadsheet headers, document text, user identities, or file metadata are logged.

3. **Template & Input Sanitization**:
   - **Placeholder Delimiter Isolation**: Variables are processed strictly through structured token interpolation (Docxtemplater) without running arbitrary code or `eval()`.
   - **Filename Sanitization**: Output filenames generated from dynamic pattern tags are sanitized against illegal filesystem characters (`/`, `\`, `?`, `%`, `*`, `:`, `|`, `"`, `<`, `>`) and directory traversal sequences (`../`).
   - **Extension Normalization**: File extensions are explicitly normalized before download to prevent double-extension or MIME-confusion vulnerabilities.

4. **Safe Memory & Blob Handling**:
   - Generated documents are handled as ephemeral in-memory Blobs. Object URLs generated for downloads are cleaned up via `URL.revokeObjectURL()` to avoid memory retention leaks.
   - Closing or refreshing the browser tab purges all processed document data immediately.

5. **Offline Operability**:
   - Once loaded, the core mapping and generation features require no active internet connection and can be operated fully offline.

---

## 📦 Supported Versions

Security patches and dependency updates are actively maintained on the latest release and the `main` branch:

| Version | Status | Supported |
| :--- | :--- | :---: |
| **0.1.x (`main`)** | **Current Release** | :white_check_mark: |
| < 0.1.0 | Deprecated | :x: |

---

## 🛡️ Reporting a Vulnerability

We take all potential security issues and vulnerability reports seriously.

If you believe you have discovered a security vulnerability in Excel to Word Mapper (such as a client-side execution vulnerability, improper input sanitization, or dependency vulnerability):

1. **Do not disclose the issue publicly** in public GitHub issues, discussions, or social media.
2. **Submit a Private Vulnerability Report**:
   - Use GitHub's [Private Vulnerability Reporting](https://github.com/JOEL-K-S/excel-to-word-mapper/security/advisories/new) feature on this repository.
   - Alternatively, open an issue requesting a private security contact channel.

### What to Include in Your Report
To help us triage and resolve the issue quickly, please provide:
- A clear description of the vulnerability and its potential impact.
- Step-by-step reproduction instructions or a minimal Proof of Concept (PoC) file (sanitized of sensitive data).
- The browser name, version, and operating system used.

### Response Timeline
- **Initial Acknowledgement**: Within 48 hours of receipt.
- **Triage & Assessment**: Within 5 business days.
- **Remediation & Patch**: Promptly released to `main` upon verification, followed by a security advisory credit if requested.
