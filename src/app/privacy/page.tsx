import type { Metadata } from "next";
import Link from "next/link";
import {
  ShieldCheck,
  Lock,
  WifiOff,
  Code,
  FileSpreadsheet,
  ArrowLeft,
  CheckCircle2,
  ExternalLink,
  ServerOff,
} from "lucide-react";
import styles from "./privacy.module.css";

export const metadata: Metadata = {
  title: "Privacy Policy | 100% Client-Side Privacy Guarantee",
  description:
    "Learn how Excel to Word Mapper protects your confidential data with 100% in-browser processing, zero server uploads, and no telemetry.",
  alternates: {
    canonical: "https://excel-to-word-mapper.vercel.app/privacy",
  },
};

export default function PrivacyPage() {
  return (
    <div className={styles.container}>
      {/* Navigation Header */}
      <header className={styles.header}>
        <Link href="/" className={styles.brand}>
          <div className={styles.brandIcon}>
            <FileSpreadsheet size={22} />
          </div>
          <div className={styles.brandText}>
            <h1>Excel to Word Mapper</h1>
          </div>
        </Link>

        <Link href="/" className="btn btn-secondary" style={{ fontSize: "0.85rem", padding: "8px 16px" }}>
          <ArrowLeft size={16} /> Back to Mapper
        </Link>
      </header>

      {/* Hero Banner */}
      <div className={styles.hero}>
        <div className={styles.heroBadge}>
          <ShieldCheck size={16} /> Zero Cloud Storage • 100% Private
        </div>
        <h1 className={styles.heroTitle}>Privacy Policy &amp; Data Security</h1>
        <p className={styles.heroSubtitle}>
          Excel to Word Mapper was architected with a strict <strong>zero-telemetry, client-side only</strong> philosophy. 
          Your confidential documents, spreadsheet rows, and generated files never touch an external server.
        </p>
      </div>

      {/* Trust Pillars Grid */}
      <div className={styles.pillarsGrid}>
        <div className={styles.pillarCard}>
          <div className={styles.pillarIcon} style={{ background: "rgba(16, 185, 129, 0.15)", color: "var(--accent-emerald)" }}>
            <ServerOff size={20} />
          </div>
          <h2 className={styles.pillarTitle}>Zero Server Uploads</h2>
          <p className={styles.pillarDesc}>
            All document generation, parsing, and ZIP archiving occur strictly inside your local browser memory.
          </p>
        </div>

        <div className={styles.pillarCard}>
          <div className={styles.pillarIcon} style={{ background: "rgba(99, 102, 241, 0.15)", color: "var(--primary-light)" }}>
            <Lock size={20} />
          </div>
          <h2 className={styles.pillarTitle}>No Data Collection</h2>
          <p className={styles.pillarDesc}>
            We do not log spreadsheet contents, column names, document text, customer identities, or telemetry.
          </p>
        </div>

        <div className={styles.pillarCard}>
          <div className={styles.pillarIcon} style={{ background: "rgba(6, 182, 212, 0.15)", color: "var(--accent-cyan)" }}>
            <WifiOff size={20} />
          </div>
          <h2 className={styles.pillarTitle}>Works Fully Offline</h2>
          <p className={styles.pillarDesc}>
            Once the page loads, you can disconnect your internet completely. The application remains 100% operational.
          </p>
        </div>

        <div className={styles.pillarCard}>
          <div className={styles.pillarIcon} style={{ background: "rgba(245, 158, 11, 0.15)", color: "var(--accent-amber)" }}>
            <Code size={20} />
          </div>
          <h2 className={styles.pillarTitle}>Open Source Verified</h2>
          <p className={styles.pillarDesc}>
            Every line of source code is public and auditable on GitHub. You can verify network traffic yourself.
          </p>
        </div>
      </div>

      {/* Detailed Policy Content */}
      <article className={styles.content}>
        <section className={styles.section}>
          <h2 className={styles.sectionTitle}>
            <CheckCircle2 size={20} style={{ color: "var(--accent-emerald)" }} />
            1. Core Architectural Principle: Client-Side Processing
          </h2>
          <p className={styles.sectionText}>
            Traditional document generation and mail merge SaaS platforms require users to upload Excel sheets and Word templates to remote cloud servers for rendering. 
            This presents major compliance and security concerns when working with sensitive client, patient, or employee data.
          </p>
          <p className={styles.sectionText}>
            <strong>Excel to Word Mapper operates differently:</strong> It is built using client-side JavaScript, WebAssembly, and standard HTML5 File APIs. 
            When you select or drop a file:
          </p>
          <ul className={styles.sectionList}>
            <li>The file is read directly into your device's local memory (RAM) via the browser's <code>FileReader</code> API.</li>
            <li>Excel workbooks are parsed entirely in memory using the open-source SheetJS library.</li>
            <li>Word document templates are interpolated directly in memory using Docxtemplater and PizZip.</li>
            <li>PDF conversion is rendered client-side using Canvas and jsPDF directly on your GPU/CPU.</li>
            <li>ZIP files are bundled in memory via JSZip and triggered as a direct browser Blob download.</li>
          </ul>
        </section>

        <section className={styles.section}>
          <h2 className={styles.sectionTitle}>
            <CheckCircle2 size={20} style={{ color: "var(--accent-emerald)" }} />
            2. What Data Is Collected?
          </h2>
          <p className={styles.sectionText}>
            <strong>None.</strong> We do not collect, store, sell, or transmit any of the following:
          </p>
          <ul className={styles.sectionList}>
            <li>Spreadsheet records, customer names, financial values, or column headers.</li>
            <li>Word document templates, logos, styling, or output documents.</li>
            <li>User account information, emails, or personal identifiers.</li>
            <li>Tracking cookies, advertising identifiers, or session recordings.</li>
          </ul>
        </section>

        <section className={styles.section}>
          <h2 className={styles.sectionTitle}>
            <CheckCircle2 size={20} style={{ color: "var(--accent-emerald)" }} />
            3. Regulatory Compliance (GDPR, HIPAA, FERPA)
          </h2>
          <p className={styles.sectionText}>
            Because no personal, health, or student records are ever transmitted over the network or stored on our infrastructure:
          </p>
          <ul className={styles.sectionList}>
            <li><strong>GDPR (Europe):</strong> No international data transfer occurs. You retain exclusive data controllership on your local machine.</li>
            <li><strong>HIPAA (United States):</strong> Protected Health Information (PHI) is never received by any third-party server, eliminating cloud transmission risk.</li>
            <li><strong>FERPA (Education):</strong> Student grades, transcripts, and diplomas remain strictly confidential on the educator's device.</li>
          </ul>
        </section>

        <section className={styles.section}>
          <h2 className={styles.sectionTitle}>
            <CheckCircle2 size={20} style={{ color: "var(--accent-emerald)" }} />
            4. Web Hosting &amp; Infrastructure
          </h2>
          <p className={styles.sectionText}>
            The web application's static assets (HTML, CSS, JavaScript bundles) are delivered via Vercel's global Content Delivery Network (CDN). 
            Like all web hosting providers, Vercel may collect standard, non-identifying server access logs (such as requesting IP address, user-agent, and HTTP status codes) solely for DDoS mitigation and edge caching. 
            Vercel never has access to the files processed inside your browser session.
          </p>
        </section>

        <section className={styles.section}>
          <h2 className={styles.sectionTitle}>
            <CheckCircle2 size={20} style={{ color: "var(--accent-emerald)" }} />
            5. How to Verify Our Privacy Guarantee
          </h2>
          <p className={styles.sectionText}>
            You do not need to take our word for it—you can independently verify our claims:
          </p>
          <div className={styles.callout}>
            <strong>Test in Browser DevTools:</strong> Open your browser's Developer Tools (<kbd>F12</kbd> or <kbd>Ctrl+Shift+I</kbd>), switch to the <strong>Network</strong> tab, and generate a batch of documents. 
            You will observe that <strong>zero HTTP POST or PUT requests</strong> are dispatched during the entire file mapping and generation process.
          </div>
        </section>

        <section className={styles.section}>
          <h2 className={styles.sectionTitle}>
            <CheckCircle2 size={20} style={{ color: "var(--accent-emerald)" }} />
            6. Source Code &amp; Contact
          </h2>
          <p className={styles.sectionText}>
            Excel to Word Mapper is open-source under the MIT License. If you have questions or wish to audit the code:
          </p>
          <p className={styles.sectionText}>
            <a
              href="https://github.com/JOEL-K-S/excel-to-word-mapper"
              target="_blank"
              rel="noopener noreferrer"
              style={{ color: "var(--primary-light)", display: "inline-flex", alignItems: "center", gap: "6px" }}
            >
              Inspect Source on GitHub <ExternalLink size={14} />
            </a>
          </p>
        </section>
      </article>

      {/* Footer */}
      <footer className={styles.footer}>
        <div className={styles.footerLinks}>
          <Link href="/">Excel to Word Mapper</Link>
          <a href="https://github.com/JOEL-K-S/excel-to-word-mapper" target="_blank" rel="noopener noreferrer">
            GitHub Repository
          </a>
          <a href="/llms.txt" target="_blank" rel="noopener noreferrer">
            AI Manifest (llms.txt)
          </a>
        </div>
        <div>
          Last updated: October 2026 • Free &amp; Open Source under the MIT License • Created by{" "}
          <a
            href="https://github.com/JOEL-K-S"
            target="_blank"
            rel="noopener noreferrer"
            style={{ color: "var(--primary-light)", textDecoration: "none", fontWeight: 600 }}
          >
            JOEL-K-S
          </a>
        </div>
      </footer>
    </div>
  );
}
