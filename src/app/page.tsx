"use client";

import React, { useState, useRef, useTransition } from "react";
import Link from "next/link";
import {
  FileSpreadsheet,
  FileText,
  Upload,
  ArrowRight,
  CheckCircle2,
  AlertCircle,
  Download,
  RefreshCw,
  Sparkles,
  Layers,
  ChevronRight,
  ChevronLeft,
  ChevronDown,
  ChevronUp,
  Plus,
  Trash2,
  FileArchive,
  Table as TableIcon,
  HelpCircle,
  Check,
  X,
  FileDown,
  ShieldCheck,
  Zap,
  Award,
  ExternalLink,
  FolderArchive,
  Eye,
  Copy,
} from "lucide-react";
import styles from "./page.module.css";
import {
  parseExcelFile,
  generateSampleExcel,
  ExcelSheetData,
} from "@/lib/excelParser";
import {
  extractPlaceholdersFromDocx,
  generateSingleDocx,
  generateBatchZip,
  generateSampleWordTemplate,
  computeDocFilename,
  triggerBrowserDownload,
  convertDocxBlobToPdfBlob,
  GenerationProgress,
} from "@/lib/wordProcessor";
import { autoMapPlaceholders } from "@/lib/matcher";

export default function Home() {
  // Wizard navigation: 0 = Start, 1 = Excel, 2 = Word, 3 = Editor, 4 = Export
  const [activeStep, setActiveStep] = useState<0 | 1 | 2 | 3 | 4>(0);

  // Excel state
  const [excelFile, setExcelFile] = useState<{ name: string; size: number } | null>(null);
  const [excelData, setExcelData] = useState<ExcelSheetData | null>(null);
  const [rawExcelBuffer, setRawExcelBuffer] = useState<ArrayBuffer | null>(null);
  const [isExcelDragging, setIsExcelDragging] = useState(false);

  // Word state
  const [wordFile, setWordFile] = useState<{ name: string; size: number } | null>(null);
  const [wordPlaceholders, setWordPlaceholders] = useState<string[]>([]);
  const [rawWordBuffer, setRawWordBuffer] = useState<ArrayBuffer | null>(null);
  const [isWordDragging, setIsWordDragging] = useState(false);

  // Mapping state: placeholder -> excelColumn
  const [mappings, setMappings] = useState<Record<string, string>>({});
  const [defaultValues, setDefaultValues] = useState<Record<string, string>>({});
  const [newPlaceholderInput, setNewPlaceholderInput] = useState("");
  const [showAddPlaceholder, setShowAddPlaceholder] = useState(false);

  // Generation state
  const [selectedPreviewRow, setSelectedPreviewRow] = useState<number>(0);
  const [namingPattern, setNamingPattern] = useState<string>("Document_{Candidate_Name}.docx");
  const [singleFormat, setSingleFormat] = useState<"docx" | "pdf">("docx");
  const [batchFormat, setBatchFormat] = useState<"docx" | "pdf">("docx");
  const [isGenerating, setIsGenerating] = useState(false);
  const [generationProgress, setGenerationProgress] = useState<GenerationProgress | null>(null);

  // UI status & helpers
  const [showTablePreview, setShowTablePreview] = useState(false);
  const [showHelpModal, setShowHelpModal] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [successToast, setSuccessToast] = useState<string | null>(null);
  const [openFaq, setOpenFaq] = useState<number | null>(null);
  const [canvasFontSize, setCanvasFontSize] = useState<number>(14);

  // Hidden file input refs
  const excelInputRef = useRef<HTMLInputElement>(null);
  const wordInputRef = useRef<HTMLInputElement>(null);
  const [, startTransition] = useTransition();

  // Trigger Toast Notification
  const triggerToast = (msg: string) => {
    setSuccessToast(msg);
    setTimeout(() => {
      setSuccessToast(null);
    }, 4500);
  };

  // Format bytes helper
  const formatSize = (bytes: number) => {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / 1048576).toFixed(1) + " MB";
  };

  // ----------------------------------------------------
  // Excel File Processing
  // ----------------------------------------------------
  const handleExcelFile = async (file: File) => {
    try {
      setErrorMessage(null);
      const validExts = [".xlsx", ".xls", ".csv"];
      const hasValidExt = validExts.some((ext) => file.name.toLowerCase().endsWith(ext));
      if (!hasValidExt) {
        setErrorMessage("Please upload a valid Excel spreadsheet (.xlsx, .xls) or CSV file.");
        return;
      }

      const buffer = await file.arrayBuffer();
      const parsedData = parseExcelFile(buffer);

      setRawExcelBuffer(buffer);
      setExcelData(parsedData);
      setExcelFile({ name: file.name, size: file.size });

      // Auto-update naming pattern with first header
      if (parsedData.headers.length > 0) {
        const primaryHeader = parsedData.headers[0];
        setNamingPattern(`Document_{${primaryHeader}}.docx`);
      }

      // If Word is already loaded, re-run auto-mapper
      if (wordPlaceholders.length > 0) {
        const autoMappings = autoMapPlaceholders(wordPlaceholders, parsedData.headers);
        setMappings(autoMappings);
      }

      triggerToast(`Loaded Excel: ${file.name} (${parsedData.totalRows} records)`);
    } catch (err: any) {
      console.error(err);
      setErrorMessage("Failed to read Excel file: " + (err.message || "Unknown error"));
    }
  };

  // ----------------------------------------------------
  // Word File Processing
  // ----------------------------------------------------
  const handleWordFile = async (file: File) => {
    try {
      setErrorMessage(null);
      if (!file.name.toLowerCase().endsWith(".docx")) {
        setErrorMessage("Please upload a Microsoft Word document (.docx).");
        return;
      }

      const buffer = await file.arrayBuffer();
      const placeholders = extractPlaceholdersFromDocx(buffer);

      setRawWordBuffer(buffer);
      setWordFile({ name: file.name, size: file.size });
      setWordPlaceholders(placeholders);

      // Auto-map if Excel data already exists
      if (excelData && excelData.headers.length > 0) {
        const autoMappings = autoMapPlaceholders(placeholders, excelData.headers);
        setMappings(autoMappings);
      }

      triggerToast(`Loaded Word template: ${file.name} (${placeholders.length} tags detected)`);
    } catch (err: any) {
      console.error(err);
      setErrorMessage("Failed to read Word template: " + (err.message || "Unknown error"));
    }
  };

  // ----------------------------------------------------
  // Load Sample Files (One-Click Demo)
  // ----------------------------------------------------
  const handleLoadSample = () => {
    try {
      setErrorMessage(null);
      // 1. Generate Sample Excel
      const excelBuf = generateSampleExcel();
      const parsedExcel = parseExcelFile(excelBuf);
      setRawExcelBuffer(excelBuf);
      setExcelData(parsedExcel);
      setExcelFile({ name: "Sample_Job_Offers.xlsx", size: excelBuf.byteLength });

      // 2. Generate Sample Word Template
      const wordBuf = generateSampleWordTemplate();
      const placeholders = extractPlaceholdersFromDocx(wordBuf);
      setRawWordBuffer(wordBuf);
      setWordFile({ name: "Sample_Offer_Letter_Template.docx", size: wordBuf.byteLength });
      setWordPlaceholders(placeholders);

      // 3. Auto-map tags
      const autoMappings = autoMapPlaceholders(placeholders, parsedExcel.headers);
      setMappings(autoMappings);

      // 4. Default naming pattern
      setNamingPattern("Offer_Letter_{Employee_Name}.docx");

      // 5. Jump to Editor
      setActiveStep(3);
      triggerToast("Sample dataset loaded! You can now test the live preview.");
    } catch (err: any) {
      console.error(err);
      setErrorMessage("Error loading sample data: " + err.message);
    }
  };

  // Re-run Auto Mapping
  const handleRemapAll = () => {
    if (!excelData || wordPlaceholders.length === 0) return;
    const newMappings = autoMapPlaceholders(wordPlaceholders, excelData.headers);
    setMappings(newMappings);
    triggerToast("Auto-mapped placeholders to closest Excel headers.");
  };

  // Update specific mapping
  const handleMappingChange = (placeholder: string, excelColumn: string) => {
    setMappings((prev) => ({
      ...prev,
      [placeholder]: excelColumn,
    }));
  };

  // Fallback defaults
  const handleDefaultChange = (placeholder: string, val: string) => {
    setDefaultValues((prev) => ({
      ...prev,
      [placeholder]: val,
    }));
  };

  // Add custom placeholder
  const handleAddCustomPlaceholder = () => {
    const cleanTag = newPlaceholderInput.trim().replace(/[{}]/g, "");
    if (!cleanTag) return;
    if (!wordPlaceholders.includes(cleanTag)) {
      setWordPlaceholders((prev) => [...prev, cleanTag]);
      if (excelData && excelData.headers.includes(cleanTag)) {
        setMappings((prev) => ({ ...prev, [cleanTag]: cleanTag }));
      }
    }
    setNewPlaceholderInput("");
    setShowAddPlaceholder(false);
  };

  // Remove placeholder
  const handleRemovePlaceholder = (placeholder: string) => {
    setWordPlaceholders((prev) => prev.filter((p) => p !== placeholder));
    setMappings((prev) => {
      const copy = { ...prev };
      delete copy[placeholder];
      return copy;
    });
  };

  // ----------------------------------------------------
  // Single Document Download
  // ----------------------------------------------------
  const handleDownloadSingle = async () => {
    if (!rawWordBuffer || !excelData || excelData.rows.length === 0) {
      setErrorMessage("Please load both Excel data and a Word template before downloading.");
      return;
    }

    try {
      setIsGenerating(true);
      setErrorMessage(null);

      const row = excelData.rows[selectedPreviewRow] || excelData.rows[0];
      const payload: Record<string, any> = {};

      wordPlaceholders.forEach((ph) => {
        const col = mappings[ph];
        if (col && row[col] !== undefined && row[col] !== null) {
          payload[ph] = row[col];
        } else if (defaultValues[ph] !== undefined && defaultValues[ph] !== "") {
          payload[ph] = defaultValues[ph];
        } else {
          payload[ph] = "";
        }
      });

      Object.keys(row).forEach((k) => {
        if (!(k in payload)) payload[k] = row[k];
      });

      const docBlob = generateSingleDocx(rawWordBuffer, payload);
      let outputBlob: Blob;
      let filename: string;

      if (singleFormat === "pdf") {
        outputBlob = await convertDocxBlobToPdfBlob(docBlob);
        filename = computeDocFilename(namingPattern, payload, selectedPreviewRow).replace(/\.docx$/i, ".pdf");
      } else {
        outputBlob = docBlob;
        filename = computeDocFilename(namingPattern, payload, selectedPreviewRow);
      }

      await triggerBrowserDownload(outputBlob, filename);
      triggerToast(`Downloaded: ${filename}`);
    } catch (err: any) {
      console.error(err);
      setErrorMessage("Error generating document: " + err.message);
    } finally {
      setIsGenerating(false);
    }
  };

  // ----------------------------------------------------
  // Batch Document Generation & ZIP Download
  // ----------------------------------------------------
  const handleDownloadBatch = async () => {
    if (!rawWordBuffer || !excelData || excelData.rows.length === 0) {
      setErrorMessage("Please load both Excel data and a Word template before generating.");
      return;
    }

    setIsGenerating(true);
    setErrorMessage(null);

    try {
      const zipBlob = await generateBatchZip(
        rawWordBuffer,
        excelData.rows,
        mappings,
        defaultValues,
        namingPattern,
        batchFormat,
        (progress) => {
          setGenerationProgress(progress);
        }
      );

      const safeSheetName = (excelData.activeSheet || "Documents").replace(/[/\\?%*:|"<>]/g, "_");
      const zipFilename = `${safeSheetName}_Generated_${batchFormat.toUpperCase()}.zip`;
      await triggerBrowserDownload(zipBlob, zipFilename);

      triggerToast(`Batch complete! Downloaded ${excelData.rows.length} ${batchFormat.toUpperCase()} files in ZIP.`);
    } catch (err: any) {
      console.error(err);
      setErrorMessage("Batch generation failed: " + err.message);
    } finally {
      setIsGenerating(false);
      setGenerationProgress(null);
    }
  };

  // Calculate mapping counts
  const mappedCount = wordPlaceholders.filter((p) => Boolean(mappings[p])).length;
  const mappingPercent =
    wordPlaceholders.length > 0 ? Math.round((mappedCount / wordPlaceholders.length) * 100) : 0;

  // Active row data helper for live canvas
  const activeRecord = excelData && excelData.rows.length > 0 ? excelData.rows[selectedPreviewRow] : null;

  const getInterpolatedValue = (tag: string, fallback: string = "") => {
    if (!activeRecord) return fallback || `{${tag}}`;
    const col = mappings[tag];
    if (col && activeRecord[col] !== undefined && activeRecord[col] !== null) {
      return String(activeRecord[col]);
    }
    return defaultValues[tag] || fallback || `{${tag}}`;
  };

  return (
    <div className={styles.container}>
      {/* ======================================================== */}
      {/* GLOBAL TOP HEADER */}
      {/* ======================================================== */}
      <header className={styles.header}>
        <div className={styles.headerInner}>
          {/* Brand */}
          <div className={styles.brand} onClick={() => setActiveStep(0)}>
            <div className={styles.brandIcon}>
              <FileSpreadsheet size={22} />
            </div>
            <div className={styles.brandText}>
              <div className={styles.brandTitleRow}>
                <h1 className={styles.brandTitle}>Excel to Word Mapper</h1>
                <span className={styles.brandVersion}>v2.4-web</span>
              </div>
              <p className={styles.brandSubtitle}>Automate batch documents with client-side mapping</p>
            </div>
          </div>

          {/* Stepper Navigation */}
          <nav className={styles.stepperNav} aria-label="Workflow Steps">
            <button
              type="button"
              className={`${styles.stepNavBtn} ${activeStep === 0 ? styles.active : ""}`}
              onClick={() => setActiveStep(0)}
            >
              Start
            </button>
            <span className={styles.stepNavDivider}>/</span>

            <button
              type="button"
              className={`${styles.stepNavBtn} ${activeStep === 1 ? styles.active : ""} ${
                excelFile ? styles.completed : ""
              }`}
              onClick={() => setActiveStep(1)}
            >
              {excelFile && <Check size={13} />} 1. Excel File
            </button>
            <span className={styles.stepNavDivider}>/</span>

            <button
              type="button"
              className={`${styles.stepNavBtn} ${activeStep === 2 ? styles.active : ""} ${
                wordFile ? styles.completed : ""
              }`}
              onClick={() => setActiveStep(2)}
            >
              {wordFile && <Check size={13} />} 2. Word Template
            </button>
            <span className={styles.stepNavDivider}>/</span>

            <button
              type="button"
              className={`${styles.stepNavBtn} ${activeStep === 3 ? styles.active : ""} ${
                mappedCount > 0 ? styles.completed : ""
              }`}
              onClick={() => setActiveStep(3)}
            >
              3. Document Editor
            </button>
            <span className={styles.stepNavDivider}>/</span>

            <button
              type="button"
              className={`${styles.stepNavBtn} ${activeStep === 4 ? styles.active : ""}`}
              onClick={() => setActiveStep(4)}
            >
              4. Export
            </button>
          </nav>

          {/* Header Actions */}
          <div className={styles.headerActions}>
            <button
              type="button"
              onClick={() => setShowHelpModal(true)}
              className="btn btn-secondary"
              style={{ fontSize: "0.82rem", padding: "7px 14px" }}
              title="Formatting & Syntax Guide"
            >
              <HelpCircle size={15} /> Guide
            </button>
            <a
              href="https://github.com/JOEL-K-S/excel-to-word-mapper"
              target="_blank"
              rel="noopener noreferrer"
              className="btn btn-secondary"
              style={{ padding: "8px 12px", color: "var(--text-muted)" }}
              title="GitHub Source"
            >
              <ExternalLink size={15} />
            </a>
          </div>
        </div>
      </header>

      {/* Floating Success Toast */}
      {successToast && (
        <div
          className="animate-fade-in"
          style={{
            position: "fixed",
            bottom: "20px",
            right: "20px",
            zIndex: 1000,
            background: "rgba(16, 185, 129, 0.95)",
            backdropFilter: "blur(12px)",
            color: "#ffffff",
            padding: "12px 20px",
            borderRadius: "var(--radius-md)",
            boxShadow: "0 10px 30px rgba(16, 185, 129, 0.4)",
            display: "flex",
            alignItems: "center",
            gap: "10px",
            fontSize: "0.88rem",
            fontWeight: 600,
          }}
        >
          <CheckCircle2 size={18} />
          <span>{successToast}</span>
          <button
            onClick={() => setSuccessToast(null)}
            style={{ background: "none", border: "none", color: "#ffffff", cursor: "pointer", marginLeft: "8px" }}
          >
            <X size={16} />
          </button>
        </div>
      )}

      {/* Global Error Banner */}
      {errorMessage && (
        <div
          className="animate-fade-in"
          style={{
            maxWidth: "1100px",
            width: "calc(100% - 48px)",
            margin: "20px auto 0",
            background: "rgba(244, 63, 94, 0.15)",
            border: "1px solid rgba(244, 63, 94, 0.4)",
            borderRadius: "var(--radius-md)",
            padding: "12px 20px",
            color: "#fda4af",
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            gap: "12px",
            fontSize: "0.88rem",
          }}
        >
          <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
            <AlertCircle size={18} style={{ color: "var(--accent-rose)", flexShrink: 0 }} />
            <span>{errorMessage}</span>
          </div>
          <button
            onClick={() => setErrorMessage(null)}
            style={{ background: "none", border: "none", color: "#fda4af", cursor: "pointer" }}
          >
            <X size={16} />
          </button>
        </div>
      )}

      {/* Main Workspace Body */}
      <main className={styles.mainContent}>
        {/* ======================================================== */}
        {/* STEP 0: START / WELCOME SCREEN */}
        {/* ======================================================== */}
        {activeStep === 0 && (
          <div className="animate-fade-in">
            <section className={styles.heroSection}>
              <div className={styles.heroBadge}>
                <ShieldCheck size={16} />
                <span>100% Client-Side Privacy Guaranteed • The First Web Excel to Report Online Creator</span>
              </div>

              <h1 className={styles.heroHeading}>
                Batch Document Generation, <br />
                <span className={styles.heroHighlight}>Purely in Your Browser</span>
              </h1>

              <p className={styles.heroDesc}>
                Map Excel spreadsheet columns straight into your Microsoft Word (.docx) templates and export individual documents or bulk ZIP archives in milliseconds. Zero server uploads.
              </p>

              <div className={styles.heroCtaGroup}>
                <button
                  type="button"
                  className={styles.btnCtaPrimary}
                  onClick={() => setActiveStep(1)}
                >
                  <span>Let's Start</span>
                  <ArrowRight size={18} />
                </button>

                <button
                  type="button"
                  className={styles.btnCtaSecondary}
                  onClick={handleLoadSample}
                >
                  <FolderArchive size={18} style={{ color: "var(--accent-rose)" }} />
                  <span>Try with Sample Files</span>
                </button>
              </div>

              <div className={styles.metricsBar}>
                <div>
                  <div className={styles.metricValue}>0 KB</div>
                  <div className={styles.metricLabel}>Server Uploads</div>
                </div>
                <div>
                  <div className={styles.metricValue}>&lt; 35ms</div>
                  <div className={styles.metricLabel}>Per Document Merge</div>
                </div>
                <div>
                  <div className={styles.metricValue}>DOCX &amp; PDF</div>
                  <div className={styles.metricLabel}>Multi-Format Export</div>
                </div>
              </div>
            </section>
          </div>
        )}

        {/* ======================================================== */}
        {/* STEP 1: UPLOAD EXCEL FILE */}
        {/* ======================================================== */}
        {activeStep === 1 && (
          <div className="animate-fade-in">
            <div className={styles.uploadCardWrapper}>
              <div className={styles.stepTitleBox}>
                <h2>Upload Excel File</h2>
                <p>Select the spreadsheet containing your data rows and recipient fields.</p>
              </div>

              {/* Hidden file input */}
              <input
                type="file"
                ref={excelInputRef}
                accept=".xlsx,.xls,.csv"
                style={{ display: "none" }}
                onChange={(e) => {
                  if (e.target.files && e.target.files[0]) {
                    handleExcelFile(e.target.files[0]);
                  }
                }}
              />

              {/* Drag & Drop Target */}
              <div
                className={`${styles.dropzone} ${isExcelDragging ? styles.dragOver : ""}`}
                onDragOver={(e) => {
                  e.preventDefault();
                  setIsExcelDragging(true);
                }}
                onDragLeave={(e) => {
                  e.preventDefault();
                  setIsExcelDragging(false);
                }}
                onDrop={(e) => {
                  e.preventDefault();
                  setIsExcelDragging(false);
                  if (e.dataTransfer.files && e.dataTransfer.files[0]) {
                    handleExcelFile(e.dataTransfer.files[0]);
                  }
                }}
                onClick={() => excelInputRef.current?.click()}
              >
                <div className={styles.dropzoneIconWrapper}>
                  <TableIcon size={34} />
                </div>
                <h3 className={styles.dropzoneTitle}>Drag and drop your spreadsheet here</h3>
                <p className={styles.dropzoneSubtitle}>
                  Supports Microsoft Excel (.xlsx, .xls) and CSV (.csv) up to 50MB
                </p>

                <div className={styles.dropzoneActions} onClick={(e) => e.stopPropagation()}>
                  <button
                    type="button"
                    className="btn btn-primary"
                    style={{ fontSize: "0.86rem", padding: "10px 22px" }}
                    onClick={() => excelInputRef.current?.click()}
                  >
                    <Upload size={16} /> Browse Files
                  </button>
                  <button
                    type="button"
                    className="btn btn-secondary"
                    style={{ fontSize: "0.86rem", padding: "10px 18px" }}
                    onClick={handleLoadSample}
                  >
                    <Sparkles size={16} /> Load Sample Sheet (140 rows)
                  </button>
                </div>
              </div>

              {/* Loaded State Pill */}
              {excelFile && excelData && (
                <div className={styles.fileLoadedPill}>
                  <div className={styles.fileLoadedInfo}>
                    <CheckCircle2 size={24} style={{ color: "var(--accent-emerald)" }} />
                    <div>
                      <div className={styles.fileLoadedName}>{excelFile.name}</div>
                      <div className={styles.fileLoadedMeta}>
                        {excelData.totalRows} rows • {excelData.headers.length} columns detected • Ready
                      </div>
                    </div>
                  </div>

                  <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                    <button
                      type="button"
                      className="btn btn-secondary"
                      style={{ fontSize: "0.82rem", padding: "8px 14px" }}
                      onClick={() => setShowTablePreview(!showTablePreview)}
                    >
                      <Eye size={15} /> {showTablePreview ? "Hide Preview" : "Inspect Data"}
                    </button>
                    <button
                      type="button"
                      className="btn btn-primary"
                      style={{ fontSize: "0.86rem", padding: "8px 18px" }}
                      onClick={() => setActiveStep(2)}
                    >
                      <span>Next: Upload Template</span>
                      <ArrowRight size={16} />
                    </button>
                  </div>
                </div>
              )}

              {/* Collapsible Excel Data Preview Drawer */}
              {showTablePreview && excelData && (
                <div
                  className="animate-fade-in"
                  style={{
                    background: "rgba(0, 0, 0, 0.4)",
                    border: "1px solid var(--border-subtle)",
                    borderRadius: "var(--radius-md)",
                    padding: "16px",
                    overflowX: "auto",
                  }}
                >
                  <div style={{ fontSize: "0.82rem", fontWeight: 700, color: "var(--text-muted)", marginBottom: "10px" }}>
                    Sheet: "{excelData.activeSheet}" (Showing First 5 Rows)
                  </div>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "0.8rem", textAlign: "left" }}>
                    <thead>
                      <tr style={{ borderBottom: "1px solid var(--border-subtle)", color: "#ffffff" }}>
                        {excelData.headers.map((h) => (
                          <th key={h} style={{ padding: "8px 12px", background: "rgba(255,255,255,0.03)" }}>
                            {h}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {excelData.rows.slice(0, 5).map((row, idx) => (
                        <tr key={idx} style={{ borderBottom: "1px solid rgba(255,255,255,0.05)", color: "var(--text-muted)" }}>
                          {excelData.headers.map((h) => (
                            <td key={h} style={{ padding: "8px 12px" }}>
                              {String(row[h] || "")}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}

              {/* Step Navigation Actions */}
              <div className={styles.stepNavFooter}>
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={() => setActiveStep(0)}
                  style={{ fontSize: "0.84rem" }}
                >
                  <ChevronLeft size={16} /> Back to Start
                </button>
                <span style={{ fontSize: "0.8rem", color: "var(--text-dim)" }}>Step 1 of 4</span>
              </div>
            </div>
          </div>
        )}

        {/* ======================================================== */}
        {/* STEP 2: UPLOAD WORD TEMPLATE */}
        {/* ======================================================== */}
        {activeStep === 2 && (
          <div className="animate-fade-in">
            <div className={styles.uploadCardWrapper}>
              <div className={styles.stepTitleBox}>
                <h2>Upload Word Template</h2>
                <p>Upload your standard .docx document containing variable tags like &#123;Candidate_Name&#125;.</p>
              </div>

              {/* Hidden file input */}
              <input
                type="file"
                ref={wordInputRef}
                accept=".docx"
                style={{ display: "none" }}
                onChange={(e) => {
                  if (e.target.files && e.target.files[0]) {
                    handleWordFile(e.target.files[0]);
                  }
                }}
              />

              {/* Drag & Drop Target */}
              <div
                className={`${styles.dropzone} ${isWordDragging ? styles.dragOver : ""}`}
                onDragOver={(e) => {
                  e.preventDefault();
                  setIsWordDragging(true);
                }}
                onDragLeave={(e) => {
                  e.preventDefault();
                  setIsWordDragging(false);
                }}
                onDrop={(e) => {
                  e.preventDefault();
                  setIsWordDragging(false);
                  if (e.dataTransfer.files && e.dataTransfer.files[0]) {
                    handleWordFile(e.dataTransfer.files[0]);
                  }
                }}
                onClick={() => wordInputRef.current?.click()}
              >
                <div className={styles.dropzoneIconWrapper}>
                  <FileText size={34} />
                </div>
                <h3 className={styles.dropzoneTitle}>Drag and drop your Word template here</h3>
                <p className={styles.dropzoneSubtitle}>
                  Supports Microsoft Word (.docx) files with &#123;variable&#125; syntax
                </p>

                <div className={styles.dropzoneActions} onClick={(e) => e.stopPropagation()}>
                  <button
                    type="button"
                    className="btn btn-primary"
                    style={{ fontSize: "0.86rem", padding: "10px 22px" }}
                    onClick={() => wordInputRef.current?.click()}
                  >
                    <Upload size={16} /> Browse Files
                  </button>
                  <button
                    type="button"
                    className="btn btn-secondary"
                    style={{ fontSize: "0.86rem", padding: "10px 18px" }}
                    onClick={handleLoadSample}
                  >
                    <Sparkles size={16} /> Load Offer Letter Template
                  </button>
                </div>
              </div>

              {/* Loaded State Pill */}
              {wordFile && (
                <div className={styles.fileLoadedPill}>
                  <div className={styles.fileLoadedInfo}>
                    <CheckCircle2 size={24} style={{ color: "var(--accent-emerald)" }} />
                    <div>
                      <div className={styles.fileLoadedName}>{wordFile.name}</div>
                      <div className={styles.fileLoadedMeta}>
                        {wordPlaceholders.length} variable tags detected • Ready for live editing
                      </div>
                    </div>
                  </div>

                  <button
                    type="button"
                    className="btn btn-primary"
                    style={{ fontSize: "0.86rem", padding: "8px 18px" }}
                    onClick={() => setActiveStep(3)}
                  >
                    <span>Open Document Editor</span>
                    <ArrowRight size={16} />
                  </button>
                </div>
              )}

              {/* Step Navigation Actions */}
              <div className={styles.stepNavFooter}>
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={() => setActiveStep(1)}
                  style={{ fontSize: "0.84rem" }}
                >
                  <ChevronLeft size={16} /> Back to Excel File
                </button>
                <span style={{ fontSize: "0.8rem", color: "var(--text-dim)" }}>Step 2 of 4</span>
              </div>
            </div>
          </div>
        )}

        {/* ======================================================== */}
        {/* STEP 3: DOCUMENT EDITOR & FIELD MAPPING */}
        {/* ======================================================== */}
        {activeStep === 3 && (
          <div className="animate-fade-in">
            <div className={styles.editorWorkspace}>
              {/* Header Controls Bar */}
              <div className={styles.editorHeaderBar}>
                <div className={styles.editorHeaderLeft}>
                  <h2 className={styles.editorTitle}>
                    <span>Interactive Document Editor</span>
                    <span className="badge badge-rose" style={{ fontSize: "0.74rem", padding: "3px 8px" }}>
                      Live Preview
                    </span>
                  </h2>
                  <p className={styles.editorSubtitle}>
                    Map spreadsheet headers to Word template tags and preview row substitutions in real time.
                  </p>
                </div>

                <div className={styles.editorControlsRight}>
                  {excelData && excelData.rows.length > 0 && (
                    <div className={styles.recordPickerWrapper}>
                      <span>Preview Record:</span>
                      <select
                        className={styles.recordPickerSelect}
                        value={selectedPreviewRow}
                        onChange={(e) => setSelectedPreviewRow(Number(e.target.value))}
                      >
                        {excelData.rows.map((row, idx) => {
                          const firstCol = excelData.headers[0] || "";
                          const secondCol = excelData.headers[1] || "";
                          const label = `Row #${idx + 1}: ${String(row[firstCol] || "")} ${
                            row[secondCol] ? `(${String(row[secondCol])})` : ""
                          }`;
                          return (
                            <option key={idx} value={idx}>
                              {label}
                            </option>
                          );
                        })}
                      </select>
                    </div>
                  )}

                  <button
                    type="button"
                    className="btn btn-primary"
                    style={{ fontSize: "0.86rem", padding: "8px 18px" }}
                    onClick={() => setActiveStep(4)}
                  >
                    <span>Proceed to Export</span>
                    <ArrowRight size={16} />
                  </button>
                </div>
              </div>

              {/* 2-Column Split: Left Variables Dock, Right Live Paper Canvas */}
              <div className={styles.editorGrid}>
                {/* Left Dock: Variables & Mapping */}
                <div className={styles.variablesDock}>
                  <div className={styles.variablesDockHeader}>
                    <span className={styles.dockTitle}>Spreadsheet Headers &amp; Tags</span>
                    <span className="badge badge-primary" style={{ fontSize: "0.72rem" }}>
                      {mappedCount} / {wordPlaceholders.length} Mapped
                    </span>
                  </div>

                  <div style={{ display: "flex", gap: "8px" }}>
                    <button
                      type="button"
                      className="btn btn-secondary"
                      style={{ flex: 1, fontSize: "0.78rem", padding: "6px 10px" }}
                      onClick={handleRemapAll}
                      title="Re-run fuzzy matching"
                    >
                      <RefreshCw size={13} /> Auto-Match
                    </button>
                    <button
                      type="button"
                      className="btn btn-secondary"
                      style={{ fontSize: "0.78rem", padding: "6px 10px" }}
                      onClick={() => setShowAddPlaceholder(!showAddPlaceholder)}
                    >
                      <Plus size={13} /> Add Tag
                    </button>
                  </div>

                  {/* Add tag drawer */}
                  {showAddPlaceholder && (
                    <div
                      style={{
                        padding: "10px",
                        background: "rgba(0, 0, 0, 0.4)",
                        borderRadius: "var(--radius-sm)",
                        border: "1px solid var(--border-glow)",
                        display: "flex",
                        gap: "6px",
                      }}
                    >
                      <input
                        type="text"
                        placeholder="Tag_Name"
                        value={newPlaceholderInput}
                        onChange={(e) => setNewPlaceholderInput(e.target.value)}
                        style={{ flex: 1, fontSize: "0.8rem", padding: "4px 8px" }}
                      />
                      <button
                        type="button"
                        className="btn btn-primary"
                        style={{ fontSize: "0.75rem", padding: "4px 10px" }}
                        onClick={handleAddCustomPlaceholder}
                      >
                        Add
                      </button>
                    </div>
                  )}

                  {/* Variable Cards List */}
                  <div className={styles.variablesList}>
                    {wordPlaceholders.length === 0 ? (
                      <div style={{ fontSize: "0.82rem", color: "var(--text-dim)", textAlign: "center", padding: "20px" }}>
                        No placeholders detected. Click "Add Tag" to create one manually.
                      </div>
                    ) : (
                      wordPlaceholders.map((ph) => {
                        const mappedCol = mappings[ph] || "";
                        const sampleVal = activeRecord && mappedCol ? activeRecord[mappedCol] : null;

                        return (
                          <div
                            key={ph}
                            className={`${styles.variableItemCard} ${mappedCol ? styles.mapped : ""}`}
                          >
                            <div className={styles.variableItemTop}>
                              <span
                                className={styles.variableTagCode}
                                title="Click to copy variable tag"
                                onClick={() => {
                                  navigator.clipboard.writeText(`{${ph}}`);
                                  triggerToast(`Copied {${ph}} to clipboard!`);
                                }}
                              >
                                &#123;{ph}&#125;
                              </span>
                              <button
                                type="button"
                                onClick={() => handleRemovePlaceholder(ph)}
                                style={{ background: "none", border: "none", color: "var(--text-dim)", cursor: "pointer", padding: "2px" }}
                                title="Remove placeholder"
                              >
                                <Trash2 size={13} />
                              </button>
                            </div>

                            <select
                              className={styles.columnSelectorSelect}
                              value={mappedCol}
                              onChange={(e) => handleMappingChange(ph, e.target.value)}
                            >
                              <option value="">-- Unmapped --</option>
                              {excelData?.headers.map((h) => (
                                <option key={h} value={h}>
                                  {h}
                                </option>
                              ))}
                            </select>

                            {sampleVal !== null && sampleVal !== undefined && (
                              <div className={styles.variableSampleValue}>
                                Sample: "{String(sampleVal)}"
                              </div>
                            )}

                            {/* Fallback default */}
                            <input
                              type="text"
                              placeholder="Fallback default..."
                              value={defaultValues[ph] || ""}
                              onChange={(e) => handleDefaultChange(ph, e.target.value)}
                              style={{ fontSize: "0.75rem", padding: "4px 8px" }}
                            />
                          </div>
                        );
                      })
                    )}
                  </div>
                </div>

                {/* Right: Live Document Paper Canvas */}
                <div className={styles.paperCanvasWrapper}>
                  <div className={styles.paperSheetTopBar}>
                    <div className={styles.paperSheetStatus}>
                      <div className={styles.paperSheetStatusDot} />
                      <span>Live Interpolated Document Sheet (Record #{selectedPreviewRow + 1})</span>
                    </div>

                    <div className={styles.zoomButtonGroup}>
                      <button
                        type="button"
                        className={styles.zoomBtn}
                        onClick={() => setCanvasFontSize((prev) => Math.max(11, prev - 1))}
                        title="Smaller text"
                      >
                        A-
                      </button>
                      <button
                        type="button"
                        className={styles.zoomBtn}
                        onClick={() => setCanvasFontSize((prev) => Math.min(20, prev + 1))}
                        title="Larger text"
                      >
                        A+
                      </button>
                    </div>
                  </div>

                  {/* Simulated Paper Canvas */}
                  <div
                    className={styles.paperSheet}
                    style={{ fontSize: `${canvasFontSize}px` }}
                    contentEditable
                    suppressContentEditableWarning
                  >
                    <div style={{ textAlign: "right", fontSize: "0.85em", color: "#64748b", marginBottom: "24px", fontFamily: "var(--font-code)" }}>
                      Date: <span className={styles.highlightPill}>{getInterpolatedValue("Offer_Date", "October 14, 2026")}</span>
                    </div>

                    <div style={{ fontWeight: 800, fontSize: "1.25em", color: "#0f172a", marginBottom: "4px" }}>
                      <span className={styles.highlightPill}>{getInterpolatedValue("Candidate_Name", "Recipient Legal Name")}</span>
                    </div>
                    <div style={{ fontSize: "0.9em", color: "#64748b", marginBottom: "24px" }}>
                      <span className={styles.highlightPill}>{getInterpolatedValue("Candidate_Address", "742 Evergreen Terrace, Seattle, WA")}</span>
                    </div>

                    <p style={{ marginBottom: "16px" }}>
                      Dear <span className={styles.highlightPill}>{getInterpolatedValue("Candidate_Name", "Candidate")}</span>,
                    </p>

                    <p style={{ marginBottom: "16px" }}>
                      We are exceptionally pleased to extend this formal offer of employment for the position of{" "}
                      <span className={styles.highlightPill}>{getInterpolatedValue("Position_Title", "Principal Product Architect")}</span> within our{" "}
                      <span className={styles.highlightPill}>{getInterpolatedValue("Department", "Core Systems")}</span> division.
                    </p>

                    <p style={{ marginBottom: "16px" }}>
                      Your starting annualized compensation will be{" "}
                      <span className={styles.highlightPill}>{getInterpolatedValue("Annual_Salary", "$185,000 USD")}</span>, payable in standard increments subject to standard deductions. Your anticipated start date is{" "}
                      <span className={styles.highlightPill}>{getInterpolatedValue("Start_Date", "November 2, 2026")}</span>, reporting directly to{" "}
                      <span className={styles.highlightPill}>{getInterpolatedValue("Manager_Name", "David K. Lindqvist")}</span>.
                    </p>

                    <p style={{ marginBottom: "32px" }}>
                      Please sign and return this confirmation letter by{" "}
                      <span className={styles.highlightPill}>{getInterpolatedValue("Deadline", "October 25, 2026")}</span> to complete your onboarding schedule.
                    </p>

                    <div style={{ borderTop: "1px solid #e2e8f0", paddingTop: "20px", fontSize: "0.88em", color: "#64748b" }}>
                      Authorized Signature: <span style={{ fontStyle: "italic", fontFamily: "serif", fontWeight: 700, color: "#0f172a" }}>Joel K. Sterling</span> • People &amp; Operations Lead
                    </div>
                  </div>

                  {/* Navigation Footer */}
                  <div className={styles.stepNavFooter}>
                    <button
                      type="button"
                      className="btn btn-secondary"
                      onClick={() => setActiveStep(2)}
                      style={{ fontSize: "0.84rem" }}
                    >
                      <ChevronLeft size={16} /> Back to Word Template
                    </button>
                    <span style={{ fontSize: "0.8rem", color: "var(--text-dim)" }}>Step 3 of 4</span>
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ======================================================== */}
        {/* STEP 4: EXPORT & BATCH COMPILATION */}
        {/* ======================================================== */}
        {activeStep === 4 && (
          <div className="animate-fade-in">
            <div style={{ maxWidth: "860px", margin: "0 auto", display: "flex", flexDirection: "column", gap: "28px" }}>
              <div className={styles.stepTitleBox}>
                <h2>Batch Export &amp; Compilation</h2>
                <p>Choose single test document download or compile all spreadsheet rows into a packaged ZIP archive.</p>
              </div>

              {/* Progress Bar during Batch Generation */}
              {isGenerating && generationProgress && (
                <div
                  className="animate-fade-in"
                  style={{
                    background: "rgba(0, 0, 0, 0.4)",
                    border: "1px solid var(--border-glow)",
                    borderRadius: "var(--radius-md)",
                    padding: "20px 24px",
                    display: "flex",
                    flexDirection: "column",
                    gap: "10px",
                  }}
                >
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <span style={{ fontWeight: 700, color: "#ffffff", fontSize: "0.95rem" }}>
                      Compiling Documents...
                    </span>
                    <span className="badge badge-primary">{generationProgress.percentage}%</span>
                  </div>

                  <div style={{ width: "100%", height: "8px", background: "rgba(255,255,255,0.1)", borderRadius: "999px", overflow: "hidden" }}>
                    <div
                      style={{
                        width: `${generationProgress.percentage}%`,
                        height: "100%",
                        background: "linear-gradient(90deg, #f43f5e 0%, #06b6d4 100%)",
                        transition: "width 0.25s ease",
                      }}
                    />
                  </div>

                  <div style={{ display: "flex", justifyContent: "space-between", fontSize: "0.82rem", color: "var(--text-muted)" }}>
                    <span>{generationProgress.status}</span>
                    <span>
                      {generationProgress.current} / {generationProgress.total} records
                    </span>
                  </div>
                </div>
              )}

              {/* 2-Card Export Grid */}
              <div className={styles.exportGrid}>
                {/* OPTION A: Single Document Export */}
                <div className={styles.exportCard}>
                  <div style={{ display: "flex", flexDirection: "column", gap: "16px" }}>
                    <div className={styles.exportCardHeader}>
                      <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                        <div className={styles.exportBadgeA}>A</div>
                        <h3 className={styles.exportCardTitle}>Single Test Document</h3>
                      </div>
                      <span className="badge badge-secondary" style={{ fontSize: "0.72rem" }}>
                        Instant
                      </span>
                    </div>

                    <p className={styles.exportCardDesc}>
                      Export one specific row to verify all placeholder tags, spacing, typography, and table layouts before bulk compilation.
                    </p>

                    {/* Target Row Picker */}
                    {excelData && excelData.rows.length > 0 && (
                      <div style={{ display: "flex", flexDirection: "column", gap: "6px" }}>
                        <label style={{ fontSize: "0.82rem", color: "var(--text-muted)", fontWeight: 600 }}>
                          Select Row Record:
                        </label>
                        <select
                          value={selectedPreviewRow}
                          onChange={(e) => setSelectedPreviewRow(Number(e.target.value))}
                          style={{ width: "100%", fontSize: "0.85rem" }}
                        >
                          {excelData.rows.map((row, idx) => {
                            const firstCol = excelData.headers[0] || "";
                            const secondCol = excelData.headers[1] || "";
                            return (
                              <option key={idx} value={idx}>
                                Row {idx + 1}: {String(row[firstCol] || "")} {row[secondCol] ? `(${String(row[secondCol])})` : ""}
                              </option>
                            );
                          })}
                        </select>
                      </div>
                    )}

                    {/* Single Format Selection */}
                    <div style={{ display: "flex", flexDirection: "column", gap: "6px" }}>
                      <label style={{ fontSize: "0.82rem", color: "var(--text-muted)", fontWeight: 600 }}>
                        Output Format:
                      </label>
                      <div className={styles.formatRadioGroup}>
                        <label
                          className={`${styles.formatRadioLabel} ${singleFormat === "docx" ? styles.selected : ""}`}
                          onClick={() => setSingleFormat("docx")}
                        >
                          <input
                            type="radio"
                            name="single-fmt"
                            checked={singleFormat === "docx"}
                            onChange={() => setSingleFormat("docx")}
                          />
                          <span>Word (.docx)</span>
                        </label>

                        <label
                          className={`${styles.formatRadioLabel} ${singleFormat === "pdf" ? styles.selected : ""}`}
                          onClick={() => setSingleFormat("pdf")}
                        >
                          <input
                            type="radio"
                            name="single-fmt"
                            checked={singleFormat === "pdf"}
                            onChange={() => setSingleFormat("pdf")}
                          />
                          <span>Adobe PDF (.pdf)</span>
                        </label>
                      </div>
                    </div>
                  </div>

                  <button
                    type="button"
                    className="btn btn-secondary"
                    style={{ width: "100%", padding: "12px 18px", marginTop: "16px" }}
                    onClick={handleDownloadSingle}
                    disabled={isGenerating || !excelData || excelData.rows.length === 0}
                  >
                    <Download size={16} /> Download Sample Document (.{singleFormat})
                  </button>
                </div>

                {/* OPTION B: Batch ZIP Export */}
                <div className={`${styles.exportCard} ${styles.recommended}`}>
                  <div className={styles.recommendedTag}>Recommended</div>

                  <div style={{ display: "flex", flexDirection: "column", gap: "16px" }}>
                    <div className={styles.exportCardHeader}>
                      <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                        <div className={styles.exportBadgeB}>B</div>
                        <h3 className={styles.exportCardTitle}>Batch Export (ZIP Archive)</h3>
                      </div>
                      <span className="badge badge-rose" style={{ fontSize: "0.72rem" }}>
                        {excelData?.totalRows || 0} Files
                      </span>
                    </div>

                    <p className={styles.exportCardDesc}>
                      Compile every spreadsheet row into standalone documents packaged inside an automated, organized ZIP archive.
                    </p>

                    {/* File naming pattern */}
                    <div style={{ display: "flex", flexDirection: "column", gap: "6px" }}>
                      <label style={{ fontSize: "0.82rem", color: "var(--text-muted)", fontWeight: 600 }}>
                        File Naming Convention:
                      </label>
                      <input
                        type="text"
                        value={namingPattern}
                        onChange={(e) => setNamingPattern(e.target.value)}
                        placeholder="e.g. {Candidate_Name}_Offer_Letter.docx"
                        style={{ fontFamily: "var(--font-code)", fontSize: "0.85rem" }}
                      />

                      {/* Live Filename Preview Box */}
                      {excelData && excelData.rows.length > 0 && (() => {
                        const previewPayload: Record<string, any> = {};
                        const curRow = excelData.rows[selectedPreviewRow] || {};
                        wordPlaceholders.forEach((ph) => {
                          const col = mappings[ph];
                          if (col && curRow[col] !== undefined && curRow[col] !== null) {
                            previewPayload[ph] = curRow[col];
                          } else if (defaultValues[ph]) {
                            previewPayload[ph] = defaultValues[ph];
                          }
                        });
                        Object.keys(curRow).forEach((k) => {
                          if (!(k in previewPayload)) previewPayload[k] = curRow[k];
                        });
                        const basePreviewFilename = computeDocFilename(namingPattern, previewPayload, selectedPreviewRow);
                        const previewFilename =
                          batchFormat === "pdf" ? basePreviewFilename.replace(/\.docx$/i, ".pdf") : basePreviewFilename;

                        return (
                          <div className={styles.liveFilenamePreviewBox}>
                            <span className={styles.liveFilenameLabel}>
                              Live Filename Preview (Row #{selectedPreviewRow + 1}):
                            </span>
                            <span className={styles.liveFilenameValue}>📄 {previewFilename}</span>
                          </div>
                        );
                      })()}
                    </div>

                    {/* Batch Archive Format */}
                    <div style={{ display: "flex", flexDirection: "column", gap: "6px" }}>
                      <label style={{ fontSize: "0.82rem", color: "var(--text-muted)", fontWeight: 600 }}>
                        Batch Archive Format:
                      </label>
                      <div className={styles.formatRadioGroup}>
                        <label
                          className={`${styles.formatRadioLabel} ${batchFormat === "docx" ? styles.selected : ""}`}
                          onClick={() => setBatchFormat("docx")}
                        >
                          <input
                            type="radio"
                            name="batch-fmt"
                            checked={batchFormat === "docx"}
                            onChange={() => setBatchFormat("docx")}
                          />
                          <span>.docx (Word ZIP)</span>
                        </label>

                        <label
                          className={`${styles.formatRadioLabel} ${batchFormat === "pdf" ? styles.selected : ""}`}
                          onClick={() => setBatchFormat("pdf")}
                        >
                          <input
                            type="radio"
                            name="batch-fmt"
                            checked={batchFormat === "pdf"}
                            onChange={() => setBatchFormat("pdf")}
                          />
                          <span>.pdf (PDF ZIP)</span>
                        </label>
                      </div>
                    </div>
                  </div>

                  <button
                    type="button"
                    className="btn btn-primary"
                    style={{ width: "100%", padding: "14px 20px", marginTop: "16px", fontSize: "0.95rem" }}
                    onClick={handleDownloadBatch}
                    disabled={isGenerating || !excelData || excelData.rows.length === 0}
                  >
                    {isGenerating ? (
                      <>
                        <RefreshCw size={18} className="animate-spin" /> Compiling Documents...
                      </>
                    ) : (
                      <>
                        <FileArchive size={18} /> Compile &amp; Download ZIP ({excelData?.totalRows || 0} Documents)
                      </>
                    )}
                  </button>
                </div>
              </div>

              {/* Navigation Actions */}
              <div className={styles.stepNavFooter}>
                <button
                  type="button"
                  className="btn btn-secondary"
                  onClick={() => setActiveStep(3)}
                  style={{ fontSize: "0.84rem" }}
                >
                  <ChevronLeft size={16} /> Back to Document Editor
                </button>
                <button
                  type="button"
                  className="btn btn-ghost"
                  onClick={() => setActiveStep(0)}
                  style={{ fontSize: "0.84rem", color: "var(--accent-rose)" }}
                >
                  Start New Batch
                </button>
              </div>
            </div>
          </div>
        )}

        {/* ======================================================== */}
        {/* SEO, FEATURES & FAQ ACCORDION SECTION */}
        {/* ======================================================== */}
        <section
          style={{
            marginTop: "80px",
            paddingTop: "48px",
            borderTop: "1px solid var(--border-subtle)",
            display: "flex",
            flexDirection: "column",
            gap: "48px",
          }}
          aria-label="Features and Frequently Asked Questions"
        >
          <div style={{ textAlign: "center", maxWidth: "760px", margin: "0 auto" }}>
            <span className="badge badge-primary" style={{ marginBottom: "12px", display: "inline-block" }}>
              The First Web Excel to Report Online Creator
            </span>
            <h2 style={{ fontSize: "clamp(1.6rem, 3vw, 2.2rem)", fontWeight: 800, color: "#ffffff", margin: 0 }}>
              The First Web Excel to Report Online Creator
            </h2>
            <p style={{ fontSize: "0.95rem", color: "var(--text-muted)", marginTop: "10px", lineHeight: 1.6 }}>
              Excel to Word Mapper is the premier 100% in-browser Web Excel-to-Report online creator. Map spreadsheet
              columns to custom templates and export personalized Word documents or converted PDF reports in bulk ZIP
              archives—with zero server uploads.
            </p>
          </div>

          {/* 6 Feature Highlights Grid */}
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))",
              gap: "20px",
            }}
          >
            <div className="glass-panel" style={{ padding: "24px", display: "flex", flexDirection: "column", gap: "10px" }}>
              <div style={{ width: "42px", height: "42px", borderRadius: "10px", background: "rgba(16, 185, 129, 0.15)", color: "var(--accent-emerald)", display: "flex", alignItems: "center", justifyContent: "center" }}>
                <ShieldCheck size={22} />
              </div>
              <h3 style={{ fontSize: "1.05rem", fontWeight: 700, color: "#ffffff", margin: 0 }}>100% Client-Side Privacy</h3>
              <p style={{ fontSize: "0.85rem", color: "var(--text-muted)", margin: 0, lineHeight: 1.5 }}>
                Zero cloud server uploads. All spreadsheet parsing, template populating, and PDF rendering run locally in your browser memory.
              </p>
            </div>

            <div className="glass-panel" style={{ padding: "24px", display: "flex", flexDirection: "column", gap: "10px" }}>
              <div style={{ width: "42px", height: "42px", borderRadius: "10px", background: "rgba(244, 63, 94, 0.15)", color: "var(--accent-rose)", display: "flex", alignItems: "center", justifyContent: "center" }}>
                <FileDown size={22} />
              </div>
              <h3 style={{ fontSize: "1.05rem", fontWeight: 700, color: "#ffffff", margin: 0 }}>Dual Export: Word &amp; PDF</h3>
              <p style={{ fontSize: "0.85rem", color: "var(--text-muted)", margin: 0, lineHeight: 1.5 }}>
                Choose between editable Microsoft Word (.docx) documents and print-ready Adobe PDF (.pdf) files without external conversion APIs.
              </p>
            </div>

            <div className="glass-panel" style={{ padding: "24px", display: "flex", flexDirection: "column", gap: "10px" }}>
              <div style={{ width: "42px", height: "42px", borderRadius: "10px", background: "rgba(245, 158, 11, 0.15)", color: "var(--accent-amber)", display: "flex", alignItems: "center", justifyContent: "center" }}>
                <Zap size={22} />
              </div>
              <h3 style={{ fontSize: "1.05rem", fontWeight: 700, color: "#ffffff", margin: 0 }}>Dynamic Custom Filenames</h3>
              <p style={{ fontSize: "0.85rem", color: "var(--text-muted)", margin: 0, lineHeight: 1.5 }}>
                Automatically name each generated document using spreadsheet values (e.g. &#123;Candidate_Name&#125;_Offer_Letter.pdf).
              </p>
            </div>

            <div className="glass-panel" style={{ padding: "24px", display: "flex", flexDirection: "column", gap: "10px" }}>
              <div style={{ width: "42px", height: "42px", borderRadius: "10px", background: "rgba(99, 102, 241, 0.15)", color: "var(--primary-light)", display: "flex", alignItems: "center", justifyContent: "center" }}>
                <Layers size={22} />
              </div>
              <h3 style={{ fontSize: "1.05rem", fontWeight: 700, color: "#ffffff", margin: 0 }}>Intelligent Auto-Matching</h3>
              <p style={{ fontSize: "0.85rem", color: "var(--text-muted)", margin: 0, lineHeight: 1.5 }}>
                Fuzzy matching automatically pairs Word &#123;tags&#125; with Excel column headers, while allowing custom fallbacks and overrides.
              </p>
            </div>

            <div className="glass-panel" style={{ padding: "24px", display: "flex", flexDirection: "column", gap: "10px" }}>
              <div style={{ width: "42px", height: "42px", borderRadius: "10px", background: "rgba(6, 182, 212, 0.15)", color: "var(--accent-cyan)", display: "flex", alignItems: "center", justifyContent: "center" }}>
                <FileArchive size={22} />
              </div>
              <h3 style={{ fontSize: "1.05rem", fontWeight: 700, color: "#ffffff", margin: 0 }}>Batch ZIP Export</h3>
              <p style={{ fontSize: "0.85rem", color: "var(--text-muted)", margin: 0, lineHeight: 1.5 }}>
                Generate hundreds or thousands of personalized documents at once, bundled in a single organized ZIP archive with real-time tracking.
              </p>
            </div>

            <div className="glass-panel" style={{ padding: "24px", display: "flex", flexDirection: "column", gap: "10px" }}>
              <div style={{ width: "42px", height: "42px", borderRadius: "10px", background: "rgba(168, 85, 247, 0.15)", color: "#c084fc", display: "flex", alignItems: "center", justifyContent: "center" }}>
                <Award size={22} />
              </div>
              <h3 style={{ fontSize: "1.05rem", fontWeight: 700, color: "#ffffff", margin: 0 }}>No Microsoft Office Required</h3>
              <p style={{ fontSize: "0.85rem", color: "var(--text-muted)", margin: 0, lineHeight: 1.5 }}>
                Works directly in Chrome, Edge, Safari, and Firefox on Windows, Mac, and Linux. No Microsoft 365 or VBA macros required.
              </p>
            </div>
          </div>

          {/* FAQ Accordion */}
          <div style={{ maxWidth: "860px", margin: "0 auto", width: "100%" }}>
            <div style={{ textAlign: "center", marginBottom: "20px" }}>
              <h3 style={{ fontSize: "1.3rem", fontWeight: 800, color: "#ffffff", margin: 0 }}>
                Frequently Asked Questions
              </h3>
            </div>

            <div style={{ display: "flex", flexDirection: "column", gap: "10px" }}>
              {[
                {
                  q: "How does Excel to Word Mapper protect my data privacy?",
                  a: "All operations run purely client-side in browser RAM via JavaScript and WebAssembly. No spreadsheets or documents are ever sent to an external server. You can even disconnect your internet after loading the page.",
                },
                {
                  q: "Can I generate print-ready PDF files as well as Word documents?",
                  a: "Yes! In Step 4, choose between Word (.docx) or Adobe PDF (.pdf) for both single test records and batch ZIP archives. The PDF conversion runs client-side in your browser.",
                },
                {
                  q: "How do I format placeholders in my Word template?",
                  a: "Wrap variable names in single curly braces inside your Word document, for example: {Candidate_Name}, {Job_Title}, or {Annual_Salary}. Typography, tables, and colors are fully preserved.",
                },
                {
                  q: "Can I give each generated document a unique filename?",
                  a: "Yes. Use dynamic tags like {Candidate_Name}_Offer_Letter.docx. The app replaces tags with each row's data automatically and strips illegal filesystem characters.",
                },
              ].map((faq, idx) => {
                const isOpen = openFaq === idx;
                return (
                  <div key={idx} className="glass-panel" style={{ overflow: "hidden" }}>
                    <button
                      type="button"
                      onClick={() => setOpenFaq(isOpen ? null : idx)}
                      style={{
                        width: "100%",
                        padding: "16px 20px",
                        background: "none",
                        border: "none",
                        color: "#ffffff",
                        fontSize: "0.95rem",
                        fontWeight: 600,
                        display: "flex",
                        justifyContent: "space-between",
                        alignItems: "center",
                        cursor: "pointer",
                        textAlign: "left",
                        gap: "12px",
                      }}
                    >
                      <span>{faq.q}</span>
                      {isOpen ? <ChevronUp size={18} /> : <ChevronDown size={18} />}
                    </button>
                    {isOpen && (
                      <div style={{ padding: "0 20px 16px 20px", fontSize: "0.88rem", color: "var(--text-muted)", lineHeight: 1.6 }}>
                        {faq.a}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        </section>
      </main>

      {/* ======================================================== */}
      {/* GLOBAL FOOTER */}
      {/* ======================================================== */}
      <footer
        style={{
          borderTop: "1px solid var(--border-subtle)",
          background: "rgba(14, 17, 24, 0.6)",
          padding: "24px 20px",
          textAlign: "center",
          fontSize: "0.82rem",
          color: "var(--text-dim)",
          display: "flex",
          flexDirection: "column",
          gap: "12px",
          alignItems: "center",
        }}
      >
        <div style={{ display: "flex", gap: "20px", flexWrap: "wrap", justifyContent: "center", alignItems: "center" }}>
          <a
            href="https://github.com/JOEL-K-S/excel-to-word-mapper"
            target="_blank"
            rel="noopener noreferrer"
            style={{ color: "var(--primary-light)", textDecoration: "none", display: "inline-flex", alignItems: "center", gap: "4px" }}
          >
            GitHub Repository <ExternalLink size={12} />
          </a>
          <a
            href="/llms.txt"
            target="_blank"
            rel="noopener noreferrer"
            style={{ color: "var(--primary-light)", textDecoration: "none", display: "inline-flex", alignItems: "center", gap: "4px" }}
          >
            AI Manifest (llms.txt) <ExternalLink size={12} />
          </a>
          <Link href="/privacy" style={{ color: "var(--primary-light)", textDecoration: "none" }}>
            Privacy Policy
          </Link>
          <button
            type="button"
            onClick={() => setShowHelpModal(true)}
            style={{ background: "none", border: "none", color: "var(--primary-light)", cursor: "pointer", fontSize: "0.82rem", padding: 0 }}
          >
            Documentation &amp; Guide
          </button>
        </div>

        <p style={{ margin: 0 }}>
          Free and open-source under the MIT License • Created by{" "}
          <a
            href="https://github.com/JOEL-K-S"
            target="_blank"
            rel="noopener noreferrer"
            style={{ color: "var(--primary-light)", textDecoration: "none", fontWeight: 600 }}
          >
            JOEL-K-S
          </a>
          . 100% Client-Side. No telemetry or server storage.
        </p>
      </footer>

      {/* ======================================================== */}
      {/* INSTRUCTIONAL GUIDE MODAL */}
      {/* ======================================================== */}
      {showHelpModal && (
        <div
          style={{
            position: "fixed",
            inset: 0,
            background: "rgba(0, 0, 0, 0.75)",
            backdropFilter: "blur(8px)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            zIndex: 10000,
            padding: "16px",
          }}
          onClick={() => setShowHelpModal(false)}
        >
          <div
            className="glass-panel"
            style={{
              maxWidth: "580px",
              width: "100%",
              padding: "28px",
              background: "var(--bg-card-solid)",
              maxHeight: "85vh",
              overflowY: "auto",
              borderRadius: "var(--radius-lg)",
              border: "1px solid var(--border-glow)",
              boxShadow: "0 20px 50px rgba(0, 0, 0, 0.6)",
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "20px" }}>
              <h3 style={{ fontSize: "1.2rem", fontWeight: 800, color: "#ffffff", margin: 0, display: "flex", alignItems: "center", gap: "8px" }}>
                <HelpCircle size={20} style={{ color: "var(--accent-rose)" }} />
                Excel to Word Mapper Documentation
              </h3>
              <button
                onClick={() => setShowHelpModal(false)}
                style={{ background: "none", border: "none", color: "var(--text-muted)", cursor: "pointer" }}
              >
                <X size={20} />
              </button>
            </div>

            <div style={{ display: "flex", flexDirection: "column", gap: "16px", fontSize: "0.88rem", color: "var(--text-muted)", lineHeight: 1.6 }}>
              <div>
                <strong style={{ color: "#ffffff", display: "block", marginBottom: "4px" }}>
                  1. Step 1: Upload Excel File
                </strong>
                <p style={{ margin: 0 }}>
                  Select any spreadsheet (.xlsx, .xls, .csv). The first row is read as variable headers.
                </p>
              </div>

              <div>
                <strong style={{ color: "#ffffff", display: "block", marginBottom: "4px" }}>
                  2. Step 2: Upload Word Template
                </strong>
                <p style={{ margin: 0 }}>
                  Provide a Microsoft Word (.docx) template containing tags enclosed in curly braces, e.g. <span className="code-tag">&#123;Candidate_Name&#125;</span>.
                </p>
              </div>

              <div>
                <strong style={{ color: "#ffffff", display: "block", marginBottom: "4px" }}>
                  3. Step 3: Document Editor
                </strong>
                <p style={{ margin: 0 }}>
                  Review detected tags, copy or insert them into your runs, and test live sample records on the canvas sheet.
                </p>
              </div>

              <div>
                <strong style={{ color: "#ffffff", display: "block", marginBottom: "4px" }}>
                  4. Step 4: Batch Export
                </strong>
                <p style={{ margin: 0 }}>
                  Download a single test record or package all rows into an organized ZIP archive containing .docx or .pdf files.
                </p>
              </div>

              <div style={{ background: "rgba(244, 63, 94, 0.1)", border: "1px solid rgba(244, 63, 94, 0.3)", borderRadius: "var(--radius-sm)", padding: "12px", color: "#fda4af", fontSize: "0.82rem" }}>
                <strong>100% Client-Side Privacy:</strong> Your spreadsheets and Word documents are parsed strictly inside your local browser memory. Nothing is ever sent to an external server.
              </div>
            </div>

            <button
              className="btn btn-primary"
              style={{ width: "100%", marginTop: "20px" }}
              onClick={() => setShowHelpModal(false)}
            >
              Got It
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
