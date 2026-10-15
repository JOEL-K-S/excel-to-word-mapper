"use client";

import React, { useState, useRef, useTransition } from "react";
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
  Eye,
  Plus,
  Trash2,
  FileArchive,
  Table as TableIcon,
  HelpCircle,
  Check,
  X,
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
  GenerationProgress,
} from "@/lib/wordProcessor";
import { autoMapPlaceholders } from "@/lib/matcher";

export default function Home() {
  // Wizard navigation
  const [activeStep, setActiveStep] = useState<1 | 2 | 3>(1);

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
  const [namingPattern, setNamingPattern] = useState<string>("Document_{Employee_Name}.docx");
  const [isGenerating, setIsGenerating] = useState(false);
  const [generationProgress, setGenerationProgress] = useState<GenerationProgress | null>(null);

  // UI status & helpers
  const [showTablePreview, setShowTablePreview] = useState(false);
  const [showHelpModal, setShowHelpModal] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [successToast, setSuccessToast] = useState<string | null>(null);

  // Hidden file input refs
  const excelInputRef = useRef<HTMLInputElement>(null);
  const wordInputRef = useRef<HTMLInputElement>(null);

  const [, startTransition] = useTransition();

  const triggerToast = (msg: string) => {
    setSuccessToast(msg);
    setTimeout(() => setSuccessToast(null), 4000);
  };

  // ----------------------------------------------------
  // Excel File Handler
  // ----------------------------------------------------
  const processExcelBuffer = (buffer: ArrayBuffer, fileName: string, fileSize: number) => {
    try {
      const parsed = parseExcelFile(buffer);
      setRawExcelBuffer(buffer);
      setExcelData(parsed);
      setExcelFile({ name: fileName, size: fileSize });
      setErrorMessage(null);

      // Auto-map if Word placeholders already exist
      if (wordPlaceholders.length > 0) {
        const auto = autoMapPlaceholders(wordPlaceholders, parsed.headers);
        setMappings(auto);
      }

      // Smart naming pattern selection based on Excel headers
      const candidateNameCols = ["file_name", "filename", "name", "candidate_name", "employee_name", "student_name"];
      const matchedNameCol = parsed.headers.find((h) =>
        candidateNameCols.includes(h.toLowerCase().replace(/[^a-z0-9_]/g, ""))
      );
      if (matchedNameCol) {
        setNamingPattern(
          matchedNameCol.toLowerCase().includes("file")
            ? `{${matchedNameCol}}.docx`
            : `Report_{${matchedNameCol}}.docx`
        );
      }

      triggerToast(`Successfully loaded Excel sheet "${parsed.activeSheet}" with ${parsed.totalRows} rows.`);
    } catch (err: any) {
      console.error(err);
      setErrorMessage(err.message || "Failed to parse Excel file.");
    }
  };

// Security & resource limit constants
const MAX_EXCEL_SIZE = 50 * 1024 * 1024; // 50MB
const MAX_WORD_SIZE = 25 * 1024 * 1024; // 25MB

function validateUploadedFile(
  file: File,
  allowedExts: string[],
  maxSize: number,
  typeLabel: string
): string | null {
  const dotIndex = file.name.lastIndexOf(".");
  const ext = dotIndex !== -1 ? file.name.slice(dotIndex).toLowerCase() : "";
  if (!allowedExts.includes(ext)) {
    return `Security validation: Invalid ${typeLabel} file extension (${ext || "none"}). Only ${allowedExts.join(", ")} allowed.`;
  }
  if (file.size > maxSize) {
    const maxMb = Math.round(maxSize / (1024 * 1024));
    return `Security validation: File size (${(file.size / (1024 * 1024)).toFixed(1)}MB) exceeds maximum safe limit of ${maxMb}MB.`;
  }
  return null;
}

  const onExcelFileSelected = (file: File) => {
    const validationError = validateUploadedFile(
      file,
      [".xlsx", ".xls", ".csv"],
      MAX_EXCEL_SIZE,
      "Excel"
    );
    if (validationError) {
      setErrorMessage(validationError);
      return;
    }
    const reader = new FileReader();
    reader.onload = (event) => {
      const buffer = event.target?.result as ArrayBuffer;
      processExcelBuffer(buffer, file.name, file.size);
    };
    reader.readAsArrayBuffer(file);
  };

  const handleExcelFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    onExcelFileSelected(file);
    e.target.value = "";
  };

  const handleSheetChange = (sheetName: string) => {
    if (!rawExcelBuffer) return;
    try {
      const parsed = parseExcelFile(rawExcelBuffer, sheetName);
      setExcelData(parsed);
      if (wordPlaceholders.length > 0) {
        setMappings(autoMapPlaceholders(wordPlaceholders, parsed.headers));
      }
    } catch (err: any) {
      setErrorMessage(err.message);
    }
  };

  // ----------------------------------------------------
  // Word File Handler
  // ----------------------------------------------------
  const processWordBuffer = (buffer: ArrayBuffer, fileName: string, fileSize: number) => {
    try {
      const placeholders = extractPlaceholdersFromDocx(buffer);
      setRawWordBuffer(buffer);
      setWordFile({ name: fileName, size: fileSize });
      setWordPlaceholders(placeholders);
      setErrorMessage(null);

      // Set intelligent default naming pattern
      if (placeholders.length > 0) {
        const preferredName = placeholders.find((p) =>
          ["name", "student_name", "candidate_name", "employee_name", "file_name"].includes(p.toLowerCase())
        ) || placeholders[0];
        setNamingPattern(`Report_{${preferredName}}.docx`);
      }

      // Auto-map if Excel data already exists
      if (excelData && excelData.headers.length > 0) {
        const auto = autoMapPlaceholders(placeholders, excelData.headers);
        setMappings(auto);
      }

      triggerToast(`Loaded Word template. Detected ${placeholders.length} placeholders.`);
    } catch (err: any) {
      console.error(err);
      setErrorMessage(err.message || "Failed to parse Word document.");
    }
  };

  const onWordFileSelected = (file: File) => {
    const validationError = validateUploadedFile(
      file,
      [".docx"],
      MAX_WORD_SIZE,
      "Word"
    );
    if (validationError) {
      setErrorMessage(validationError);
      return;
    }
    const reader = new FileReader();
    reader.onload = (event) => {
      const buffer = event.target?.result as ArrayBuffer;
      processWordBuffer(buffer, file.name, file.size);
    };
    reader.readAsArrayBuffer(file);
  };

  const handleWordFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    onWordFileSelected(file);
    e.target.value = "";
  };

  // ----------------------------------------------------
  // 1-Click Sample Demo Loader
  // ----------------------------------------------------
  const handleLoadSample = () => {
    try {
      // 1. Generate & process sample Excel
      const excelBuffer = generateSampleExcel();
      const parsedExcel = parseExcelFile(excelBuffer);
      setRawExcelBuffer(excelBuffer);
      setExcelData(parsedExcel);
      setExcelFile({ name: "Acme_Candidates_2026.xlsx", size: excelBuffer.byteLength });

      // 2. Generate & process sample Word Template
      const wordBuffer = generateSampleWordTemplate();
      const placeholders = extractPlaceholdersFromDocx(wordBuffer);
      setRawWordBuffer(wordBuffer);
      setWordFile({ name: "Offer_Letter_Template.docx", size: wordBuffer.byteLength });
      setWordPlaceholders(placeholders);

      // 3. Auto-map
      const auto = autoMapPlaceholders(placeholders, parsedExcel.headers);
      setMappings(auto);
      setNamingPattern("Offer_Letter_{Employee_Name}.docx");

      setErrorMessage(null);
      triggerToast("Sample Candidate Excel & Offer Letter Template loaded!");
    } catch (err: any) {
      console.error(err);
      setErrorMessage("Error initializing sample files: " + err.message);
    }
  };

  // ----------------------------------------------------
  // Robust Sample Download Handlers (Guaranteed Extensions)
  // ----------------------------------------------------
  const handleDownloadSampleExcel = async (e?: React.MouseEvent) => {
    if (e) {
      e.preventDefault();
      e.stopPropagation();
    }
    try {
      const buffer = generateSampleExcel();
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      await triggerBrowserDownload(blob, "sample_candidates.xlsx");
      triggerToast("Downloaded: sample_candidates.xlsx");
    } catch (err: any) {
      console.error(err);
      setErrorMessage("Failed to download sample Excel: " + err.message);
    }
  };

  const handleDownloadSampleWord = async (e?: React.MouseEvent) => {
    if (e) {
      e.preventDefault();
      e.stopPropagation();
    }
    try {
      const buffer = generateSampleWordTemplate();
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      await triggerBrowserDownload(blob, "sample_offer_letter.docx");
      triggerToast("Downloaded: sample_offer_letter.docx");
    } catch (err: any) {
      console.error(err);
      setErrorMessage("Failed to download sample Word template: " + err.message);
    }
  };

  // ----------------------------------------------------
  // Placeholder & Mapping Helpers
  // ----------------------------------------------------
  const handleMappingChange = (placeholder: string, column: string) => {
    setMappings((prev) => ({
      ...prev,
      [placeholder]: column,
    }));
  };

  const handleDefaultChange = (placeholder: string, val: string) => {
    setDefaultValues((prev) => ({
      ...prev,
      [placeholder]: val,
    }));
  };

  const handleAddCustomPlaceholder = () => {
    const trimmed = newPlaceholderInput.trim().replace(/[{}]/g, "");
    if (!trimmed) return;
    if (!wordPlaceholders.includes(trimmed)) {
      setWordPlaceholders((prev) => [...prev, trimmed]);
      if (excelData && excelData.headers.length > 0) {
        const auto = autoMapPlaceholders([trimmed], excelData.headers);
        setMappings((prev) => ({ ...prev, ...auto }));
      }
    }
    setNewPlaceholderInput("");
    setShowAddPlaceholder(false);
  };

  const handleRemovePlaceholder = (placeholder: string) => {
    setWordPlaceholders((prev) => prev.filter((p) => p !== placeholder));
    setMappings((prev) => {
      const copy = { ...prev };
      delete copy[placeholder];
      return copy;
    });
  };

  const handleRemapAll = () => {
    if (!excelData) return;
    const auto = autoMapPlaceholders(wordPlaceholders, excelData.headers);
    setMappings(auto);
    triggerToast("Auto-mapping refreshed across all fields.");
  };

  // ----------------------------------------------------
  // Single Document Generation & Download
  // ----------------------------------------------------
  const handleDownloadSingle = async () => {
    if (!rawWordBuffer || !excelData || excelData.rows.length === 0) {
      setErrorMessage("Please ensure both Excel data and Word template are loaded.");
      return;
    }

    try {
      const row = excelData.rows[selectedPreviewRow] || {};
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

      // Include all raw fields
      Object.keys(row).forEach((k) => {
        if (!(k in payload)) payload[k] = row[k];
      });

      const blob = generateSingleDocx(rawWordBuffer, payload);

      // Compute safe single file name with guaranteed .docx extension
      const filename = computeDocFilename(namingPattern, payload, selectedPreviewRow);

      // Trigger safe browser download with persistent URL
      await triggerBrowserDownload(blob, filename);

      triggerToast(`Downloaded: ${filename}`);
    } catch (err: any) {
      console.error(err);
      setErrorMessage("Error generating document: " + err.message);
    }
  };

  // ----------------------------------------------------
  // Batch Document Generation & ZIP Download
  // ----------------------------------------------------
  const handleDownloadBatch = async () => {
    if (!rawWordBuffer || !excelData || excelData.rows.length === 0) {
      setErrorMessage("Please load both Excel and Word files before generating.");
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
        (progress) => {
          setGenerationProgress(progress);
        }
      );

      // Download ZIP with guaranteed safe filename & extension
      const safeSheetName = (excelData.activeSheet || "Documents").replace(/[/\\?%*:|"<>]/g, "_");
      const zipFilename = `${safeSheetName}_Generated_Documents.zip`;
      await triggerBrowserDownload(zipBlob, zipFilename);

      triggerToast(`Batch complete! Downloaded ${excelData.rows.length} documents in ZIP.`);
    } catch (err: any) {
      console.error(err);
      setErrorMessage("Batch generation failed: " + err.message);
    } finally {
      setIsGenerating(false);
      setGenerationProgress(null);
    }
  };

  // Calculate mapping completeness
  const mappedCount = wordPlaceholders.filter((p) => Boolean(mappings[p])).length;
  const mappingPercent = wordPlaceholders.length > 0
    ? Math.round((mappedCount / wordPlaceholders.length) * 100)
    : 0;

  // Format bytes
  const formatSize = (bytes: number) => {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / 1048576).toFixed(1) + " MB";
  };

  return (
    <div className={styles.container}>
      {/* Header Bar */}
      <header className={styles.header}>
        <div className={styles.brand}>
          <div className={styles.brandIcon}>
            <FileSpreadsheet size={26} />
          </div>
          <div className={styles.brandText}>
            <h1>Excel to Word Mapper</h1>
            <p>Automate personalized documents with intelligent field mapping &amp; batch ZIP export</p>
          </div>
        </div>

        <div className={styles.headerActions}>
          <button
            className="btn btn-secondary"
            onClick={() => setShowHelpModal(!showHelpModal)}
            title="Formatting &amp; Placeholder Guide"
          >
            <HelpCircle size={16} />
            Guide
          </button>
          <button
            className="btn btn-primary"
            onClick={handleLoadSample}
            title="Load ready-made demo files"
          >
            <Sparkles size={16} />
            Try with Sample Files
          </button>
        </div>
      </header>

      {/* Toast Notification */}
      {successToast && (
        <div
          className="animate-fade-in"
          style={{
            position: "fixed",
            bottom: "16px",
            left: "16px",
            right: "16px",
            maxWidth: "460px",
            marginLeft: "auto",
            background: "rgba(16, 185, 129, 0.95)",
            backdropFilter: "blur(12px)",
            color: "#ffffff",
            padding: "12px 18px",
            borderRadius: "var(--radius-md)",
            boxShadow: "0 10px 30px rgba(16, 185, 129, 0.4)",
            display: "flex",
            alignItems: "center",
            gap: "10px",
            zIndex: 9999,
            fontWeight: 600,
            fontSize: "0.88rem",
            wordBreak: "break-word" as const,
          }}
        >
          <CheckCircle2 size={18} style={{ flexShrink: 0 }} />
          {successToast}
        </div>
      )}

      {/* Error Alert */}
      {errorMessage && (
        <div
          className="animate-fade-in"
          style={{
            background: "rgba(244, 63, 94, 0.15)",
            border: "1px solid rgba(244, 63, 94, 0.4)",
            borderRadius: "var(--radius-md)",
            padding: "14px 20px",
            color: "#fda4af",
            marginBottom: "24px",
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            gap: "12px",
          }}
        >
          <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
            <AlertCircle size={20} style={{ color: "var(--accent-rose)" }} />
            <span>{errorMessage}</span>
          </div>
          <button
            onClick={() => setErrorMessage(null)}
            style={{ background: "none", border: "none", color: "#fda4af", cursor: "pointer" }}
          >
            <X size={18} />
          </button>
        </div>
      )}

      {/* Wizard Step Navigation */}
      <nav className={styles.wizardNav} aria-label="Workflow Steps">
        <button
          className={`${styles.wizardStep} ${activeStep === 1 ? styles.active : ""} ${
            excelFile && wordFile ? styles.completed : ""
          }`}
          onClick={() => setActiveStep(1)}
        >
          <span className={styles.stepNumber}>
            {excelFile && wordFile ? <Check size={14} /> : "1"}
          </span>
          <span>1. Upload Files</span>
        </button>

        <button
          className={`${styles.wizardStep} ${activeStep === 2 ? styles.active : ""} ${
            mappedCount > 0 ? styles.completed : ""
          }`}
          onClick={() => setActiveStep(2)}
          disabled={!excelData || wordPlaceholders.length === 0}
        >
          <span className={styles.stepNumber}>
            {mappedCount === wordPlaceholders.length && wordPlaceholders.length > 0 ? (
              <Check size={14} />
            ) : (
              "2"
            )}
          </span>
          <span>2. Map Fields ({mappedCount}/{wordPlaceholders.length})</span>
        </button>

        <button
          className={`${styles.wizardStep} ${activeStep === 3 ? styles.active : ""}`}
          onClick={() => setActiveStep(3)}
          disabled={!excelData || wordPlaceholders.length === 0}
        >
          <span className={styles.stepNumber}>3</span>
          <span>3. Generate &amp; Export</span>
        </button>
      </nav>

      {/* ======================================================== */}
      {/* STEP 1: UPLOAD FILES */}
      {/* ======================================================== */}
      {activeStep === 1 && (
        <div className="animate-fade-in">
          {/* Quick-start sample banner if nothing is loaded */}
          {!excelFile && !wordFile && (
            <div className={styles.sampleBanner}>
              <div className={styles.sampleInfo}>
                <Sparkles className={styles.sampleIcon} />
                <div>
                  <div className={styles.sampleTitle}>Need Example Template Files?</div>
                  <div className={styles.sampleDesc}>
                    Load sample files directly into the mapper, or download the raw files to inspect in Excel &amp; Word.
                  </div>
                </div>
              </div>
              <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", alignItems: "center" }}>
                <button
                  type="button"
                  onClick={handleDownloadSampleExcel}
                  className="btn btn-secondary"
                  style={{ fontSize: "0.84rem", padding: "8px 14px" }}
                  title="Download raw Excel file with guaranteed .xlsx extension"
                >
                  <Download size={14} /> Download Sample .xlsx
                </button>
                <button
                  type="button"
                  onClick={handleDownloadSampleWord}
                  className="btn btn-secondary"
                  style={{ fontSize: "0.84rem", padding: "8px 14px" }}
                  title="Download raw Word template with guaranteed .docx extension"
                >
                  <Download size={14} /> Download Sample .docx
                </button>
                <button className="btn btn-primary" onClick={handleLoadSample}>
                  <Sparkles size={16} /> Load Demo In App
                </button>
              </div>
            </div>
          )}

          <div className={styles.uploadGrid}>
            {/* Left: Excel Upload */}
            <div className={`glass-panel ${styles.dropzoneCard}`}>
              <div className={styles.cardHeader}>
                <div className={styles.cardTitle}>
                  <FileSpreadsheet style={{ color: "var(--accent-emerald)" }} size={22} />
                  <span>Excel Data Source</span>
                </div>
                {excelFile && (
                  <span className="badge badge-success">
                    <CheckCircle2 size={13} /> Loaded
                  </span>
                )}
              </div>

              <input
                ref={excelInputRef}
                type="file"
                accept=".xlsx, .xls, .csv"
                style={{ display: "none" }}
                onChange={handleExcelFileChange}
              />

              {!excelFile ? (
                <div
                  className={`${styles.dropArea} ${isExcelDragging ? styles.dragActive : ""}`}
                  onClick={() => excelInputRef.current?.click()}
                  onDragOver={(e) => {
                    e.preventDefault();
                    setIsExcelDragging(true);
                  }}
                  onDragLeave={() => setIsExcelDragging(false)}
                  onDrop={(e) => {
                    e.preventDefault();
                    setIsExcelDragging(false);
                    const file = e.dataTransfer.files?.[0];
                    if (file) {
                      onExcelFileSelected(file);
                    }
                  }}
                >
                  <div className={styles.dropIcon}>
                    <Upload size={28} />
                  </div>
                  <div className={styles.dropTextPrimary}>Click or Drag Excel File Here</div>
                  <div className={styles.dropTextSecondary}>
                    Supports .xlsx, .xls, and .csv with row data
                  </div>
                  <button
                    type="button"
                    onClick={handleDownloadSampleExcel}
                    style={{ background: "transparent", border: "none", cursor: "pointer", fontSize: "0.82rem", color: "var(--accent-cyan)", textDecoration: "underline", display: "inline-flex", alignItems: "center", gap: "4px" }}
                  >
                    <Download size={13} /> Or download sample .xlsx
                  </button>
                </div>
              ) : (
                <div className={styles.fileActiveCard}>
                  <div className={styles.fileSummary}>
                    <div className={styles.fileMeta}>
                      <FileSpreadsheet size={28} style={{ color: "var(--accent-emerald)" }} />
                      <div>
                        <div className={styles.fileName}>{excelFile.name}</div>
                        <div style={{ fontSize: "0.8rem", color: "var(--text-dim)" }}>
                          {formatSize(excelFile.size)}
                        </div>
                      </div>
                    </div>
                    <button
                      className="btn btn-secondary"
                      style={{ padding: "6px 14px", fontSize: "0.82rem" }}
                      onClick={() => excelInputRef.current?.click()}
                    >
                      <RefreshCw size={14} /> Replace File
                    </button>
                  </div>

                  {excelData && (
                    <>
                      <div className={styles.fileStats}>
                        <span className="badge badge-primary">
                          {excelData.totalRows} Data Rows
                        </span>
                        <span className="badge badge-cyan">
                          {excelData.headers.length} Columns
                        </span>
                      </div>

                      {excelData.sheetNames.length > 1 && (
                        <div className={styles.sheetSelector}>
                          <label>Worksheet:</label>
                          <select
                            value={excelData.activeSheet}
                            onChange={(e) => handleSheetChange(e.target.value)}
                          >
                            {excelData.sheetNames.map((sheet) => (
                              <option key={sheet} value={sheet}>
                                {sheet}
                              </option>
                            ))}
                          </select>
                        </div>
                      )}

                      <div style={{ display: "flex", gap: "8px", marginTop: "8px" }}>
                        <button
                          className="btn btn-secondary"
                          style={{ width: "100%", fontSize: "0.84rem", padding: "8px" }}
                          onClick={() => setShowTablePreview(!showTablePreview)}
                        >
                          <TableIcon size={15} />
                          {showTablePreview ? "Hide Data Table" : "Preview Rows & Columns"}
                        </button>
                      </div>
                    </>
                  )}
                </div>
              )}
            </div>

            {/* Right: Word Template Upload */}
            <div className={`glass-panel ${styles.dropzoneCard}`}>
              <div className={styles.cardHeader}>
                <div className={styles.cardTitle}>
                  <FileText style={{ color: "var(--primary-light)" }} size={22} />
                  <span>Word Template (.docx)</span>
                </div>
                {wordFile && (
                  <span className="badge badge-primary">
                    <CheckCircle2 size={13} /> Loaded
                  </span>
                )}
              </div>

              <input
                ref={wordInputRef}
                type="file"
                accept=".docx"
                style={{ display: "none" }}
                onChange={handleWordFileChange}
              />

              {!wordFile ? (
                <div
                  className={`${styles.dropArea} ${isWordDragging ? styles.dragActive : ""}`}
                  onClick={() => wordInputRef.current?.click()}
                  onDragOver={(e) => {
                    e.preventDefault();
                    setIsWordDragging(true);
                  }}
                  onDragLeave={() => setIsWordDragging(false)}
                  onDrop={(e) => {
                    e.preventDefault();
                    setIsWordDragging(false);
                    const file = e.dataTransfer.files?.[0];
                    if (file) {
                      onWordFileSelected(file);
                    }
                  }}
                >
                  <div className={styles.dropIcon}>
                    <Upload size={28} />
                  </div>
                  <div className={styles.dropTextPrimary}>Click or Drag Word Template Here</div>
                  <div className={styles.dropTextSecondary}>
                    Must be a .docx containing &#123;placeholders&#125;
                  </div>
                  <button
                    type="button"
                    onClick={handleDownloadSampleWord}
                    style={{ background: "transparent", border: "none", cursor: "pointer", fontSize: "0.82rem", color: "var(--primary-light)", textDecoration: "underline", display: "inline-flex", alignItems: "center", gap: "4px" }}
                  >
                    <Download size={13} /> Or download sample .docx template
                  </button>
                </div>
              ) : (
                <div className={styles.fileActiveCard}>
                  <div className={styles.fileSummary}>
                    <div className={styles.fileMeta}>
                      <FileText size={28} style={{ color: "var(--primary-light)" }} />
                      <div>
                        <div className={styles.fileName}>{wordFile.name}</div>
                        <div style={{ fontSize: "0.8rem", color: "var(--text-dim)" }}>
                          {formatSize(wordFile.size)}
                        </div>
                      </div>
                    </div>
                    <button
                      className="btn btn-secondary"
                      style={{ padding: "6px 14px", fontSize: "0.82rem" }}
                      onClick={() => wordInputRef.current?.click()}
                    >
                      <RefreshCw size={14} /> Replace File
                    </button>
                  </div>

                  <div>
                    <div
                      style={{
                        display: "flex",
                        justifyContent: "space-between",
                        alignItems: "center",
                        marginBottom: "8px",
                      }}
                    >
                      <span style={{ fontSize: "0.85rem", fontWeight: 600 }}>
                        Detected Placeholders:
                      </span>
                      <span className="badge badge-cyan">{wordPlaceholders.length} Found</span>
                    </div>

                    <div className={styles.tagList}>
                      {wordPlaceholders.length > 0 ? (
                        wordPlaceholders.map((tag) => (
                          <span key={tag} className="code-tag">
                            &#123;{tag}&#125;
                          </span>
                        ))
                      ) : (
                        <span style={{ color: "var(--accent-amber)", fontSize: "0.82rem" }}>
                          No placeholders detected. You can add placeholders manually in step 2!
                        </span>
                      )}
                    </div>
                  </div>
                </div>
              )}
            </div>
          </div>

          {/* Collapsible Excel Data Preview */}
          {excelData && showTablePreview && (
            <div className={`glass-panel ${styles.dropzoneCard}`} style={{ marginBottom: "32px" }}>
              <div
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                }}
              >
                <div style={{ fontWeight: 700, fontSize: "1rem", color: "#ffffff" }}>
                  Worksheet Preview: {excelData.activeSheet} (Showing first 10 rows)
                </div>
                <span className="badge badge-primary">{excelData.totalRows} Total Records</span>
              </div>

              <div className={styles.tableWrapper}>
                <table className={styles.dataTable}>
                  <thead>
                    <tr>
                      <th style={{ width: "50px" }}>#</th>
                      {excelData.headers.map((h) => (
                        <th key={h}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {excelData.rows.slice(0, 10).map((row, i) => (
                      <tr key={i}>
                        <td style={{ color: "var(--text-dim)", fontWeight: 600 }}>{i + 1}</td>
                        {excelData.headers.map((h) => (
                          <td key={h}>{String(row[h] ?? "")}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* Bottom Next Step Action */}
          <div className={styles.navFooter}>
            <div>
              {excelFile && wordFile ? (
                <span style={{ color: "var(--accent-emerald)", fontSize: "0.9rem", fontWeight: 600 }}>
                  ✓ Both files ready for column mapping
                </span>
              ) : (
                <span style={{ color: "var(--text-muted)", fontSize: "0.88rem" }}>
                  Please upload both files or click "Try with Sample Files" to continue
                </span>
              )}
            </div>
            <button
              className="btn btn-primary"
              disabled={!excelData || !wordFile}
              onClick={() => setActiveStep(2)}
            >
              Proceed to Mapping <ArrowRight size={16} />
            </button>
          </div>
        </div>
      )}

      {/* ======================================================== */}
      {/* STEP 2: MAP FIELDS */}
      {/* ======================================================== */}
      {activeStep === 2 && (
        <div className="animate-fade-in">
          <div className={`glass-panel ${styles.mappingCard}`}>
            {/* Mapping Header & Status */}
            <div className={styles.mappingHeader}>
              <div>
                <h2 style={{ fontSize: "1.3rem", fontWeight: 800, color: "#ffffff" }}>
                  Map Template Placeholders to Excel Columns
                </h2>
                <p style={{ color: "var(--text-muted)", fontSize: "0.88rem", marginTop: "4px" }}>
                  Link each Word placeholder &#123;tag&#125; to its corresponding Excel column header.
                </p>
              </div>

              <div className={styles.mappingStats}>
                <span
                  className={`badge ${
                    mappedCount === wordPlaceholders.length
                      ? "badge-success"
                      : "badge-amber"
                  }`}
                  style={{ padding: "6px 14px", fontSize: "0.85rem" }}
                >
                  {mappedCount} of {wordPlaceholders.length} Mapped ({mappingPercent}%)
                </span>

                <button className="btn btn-secondary" onClick={handleRemapAll} title="Re-run fuzzy auto-mapping">
                  <RefreshCw size={14} /> Auto-Match
                </button>

                <button
                  className="btn btn-secondary"
                  onClick={() => setShowAddPlaceholder(!showAddPlaceholder)}
                >
                  <Plus size={14} /> Add Placeholder
                </button>
              </div>
            </div>

            {/* Custom Placeholder Input Drawer */}
            {showAddPlaceholder && (
              <div
                style={{
                  background: "rgba(0, 0, 0, 0.35)",
                  padding: "16px",
                  borderRadius: "var(--radius-md)",
                  border: "1px solid var(--border-glow)",
                  marginBottom: "20px",
                  display: "flex",
                  alignItems: "center",
                  gap: "12px",
                  flexWrap: "wrap",
                }}
              >
                <span style={{ fontSize: "0.88rem", fontWeight: 600 }}>Placeholder Name:</span>
                <input
                  type="text"
                  placeholder="e.g. Signatory_Title"
                  value={newPlaceholderInput}
                  onChange={(e) => setNewPlaceholderInput(e.target.value)}
                  style={{ flex: 1, minWidth: "200px" }}
                />
                <button className="btn btn-primary" onClick={handleAddCustomPlaceholder}>
                  Add Tag
                </button>
                <button
                  className="btn btn-ghost"
                  onClick={() => setShowAddPlaceholder(false)}
                >
                  Cancel
                </button>
              </div>
            )}

            {/* Mapping Rows List */}
            <div className={styles.mappingList}>
              {wordPlaceholders.length === 0 ? (
                <div style={{ textAlign: "center", padding: "40px 20px", color: "var(--text-muted)" }}>
                  No placeholders currently listed. Click "Add Placeholder" to create one manually.
                </div>
              ) : (
                wordPlaceholders.map((placeholder) => {
                  const mappedCol = mappings[placeholder] || "";
                  const isMapped = Boolean(mappedCol);
                  const firstRowSample =
                    excelData && excelData.rows.length > 0 && mappedCol
                      ? excelData.rows[0][mappedCol]
                      : null;

                  return (
                    <div
                      key={placeholder}
                      className={`${styles.mappingRow} ${
                        isMapped ? styles.mapped : styles.unmapped
                      }`}
                    >
                      {/* Left: Placeholder info */}
                      <div className={styles.placeholderCol}>
                        <FileText
                          size={18}
                          style={{
                            color: isMapped ? "var(--accent-emerald)" : "var(--accent-amber)",
                            flexShrink: 0,
                          }}
                        />
                        <span className="code-tag" style={{ fontSize: "0.9rem" }}>
                          &#123;{placeholder}&#125;
                        </span>
                      </div>

                      {/* Center: Connection Arrow */}
                      <div className={`${styles.arrowCol} ${isMapped ? styles.active : ""}`}>
                        <ArrowRight size={18} />
                      </div>

                      {/* Right: Excel Column Selector */}
                      <div className={styles.selectCol}>
                        <select
                          value={mappedCol}
                          onChange={(e) => handleMappingChange(placeholder, e.target.value)}
                        >
                          <option value="">-- Choose Excel Column (or Unmapped) --</option>
                          {excelData?.headers.map((header) => (
                            <option key={header} value={header}>
                              {header}
                            </option>
                          ))}
                        </select>
                        {isMapped && firstRowSample !== null && firstRowSample !== undefined && (
                          <div className={styles.previewValue}>
                            Sample: <strong>"{String(firstRowSample)}"</strong>
                          </div>
                        )}
                      </div>

                      {/* Fallback default value */}
                      <div className={styles.defaultCol}>
                        <label>Default / Fallback</label>
                        <input
                          type="text"
                          placeholder="Leave blank or static text"
                          value={defaultValues[placeholder] || ""}
                          onChange={(e) => handleDefaultChange(placeholder, e.target.value)}
                        />
                      </div>

                      {/* Delete placeholder button */}
                      <div>
                        <button
                          className="btn btn-ghost"
                          style={{ padding: "8px", color: "var(--text-dim)" }}
                          title="Remove placeholder"
                          onClick={() => handleRemovePlaceholder(placeholder)}
                        >
                          <Trash2 size={16} />
                        </button>
                      </div>
                    </div>
                  );
                })
              )}
            </div>

            {/* Excel preview accordion in step 2 */}
            {excelData && (
              <div className={styles.previewTableSection}>
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                  }}
                >
                  <span style={{ fontWeight: 700, fontSize: "0.95rem" }}>
                    Source Data Columns ({excelData.headers.length} Columns in "{excelData.activeSheet}")
                  </span>
                  <div style={{ display: "flex", gap: "6px", flexWrap: "wrap" }}>
                    {excelData.headers.slice(0, 8).map((h) => (
                      <span key={h} className="badge badge-primary">
                        {h}
                      </span>
                    ))}
                    {excelData.headers.length > 8 && (
                      <span className="badge badge-secondary">
                        +{excelData.headers.length - 8} more
                      </span>
                    )}
                  </div>
                </div>
              </div>
            )}
          </div>

          {/* Navigation Footer */}
          <div className={styles.navFooter}>
            <button className="btn btn-secondary" onClick={() => setActiveStep(1)}>
              <ChevronLeft size={16} /> Back to Files
            </button>
            <button
              className="btn btn-primary"
              disabled={mappedCount === 0}
              onClick={() => setActiveStep(3)}
            >
              Continue to Generate ({mappedCount} Mapped) <ChevronRight size={16} />
            </button>
          </div>
        </div>
      )}

      {/* ======================================================== */}
      {/* STEP 3: GENERATE & EXPORT */}
      {/* ======================================================== */}
      {activeStep === 3 && (
        <div className="animate-fade-in">
          {/* Progress Bar during Batch Generation */}
          {isGenerating && generationProgress && (
            <div className={`${styles.progressContainer} animate-fade-in`} style={{ marginBottom: "28px" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span style={{ fontWeight: 700, color: "#ffffff" }}>
                  Generating Documents...
                </span>
                <span className="badge badge-primary">
                  {generationProgress.percentage}%
                </span>
              </div>

              <div className={styles.progressBarTrack}>
                <div
                  className={styles.progressBarFill}
                  style={{ width: `${generationProgress.percentage}%` }}
                />
              </div>

              <div className={styles.progressStatus}>
                <span>{generationProgress.status}</span>
                <span>
                  {generationProgress.current} / {generationProgress.total} records
                </span>
              </div>
            </div>
          )}

          <div className={styles.generateGrid}>
            {/* Option A: Single Record Preview & Test Download */}
            <div className={`glass-panel ${styles.generateCard}`}>
              <div>
                <div style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: "8px" }}>
                  <FileText size={24} style={{ color: "var(--primary-light)" }} />
                  <h3 style={{ fontSize: "1.2rem", fontWeight: 700, color: "#ffffff" }}>
                    Single Document Test
                  </h3>
                </div>
                <p className={styles.cardDesc}>
                  Select any row from your Excel sheet to preview the interpolated values and download a single filled Word document.
                </p>
              </div>

              {excelData && excelData.rows.length > 0 && (
                <div>
                  <label style={{ fontSize: "0.85rem", color: "var(--text-muted)", fontWeight: 600, display: "block", marginBottom: "6px" }}>
                    Select Row Record:
                  </label>
                  <select
                    value={selectedPreviewRow}
                    onChange={(e) => setSelectedPreviewRow(Number(e.target.value))}
                    style={{ width: "100%" }}
                  >
                    {excelData.rows.map((row, idx) => {
                      const firstColKey = excelData.headers[0] || "";
                      const secondColKey = excelData.headers[1] || "";
                      const label = `Row ${idx + 1}: ${String(row[firstColKey] || "")} ${
                        row[secondColKey] ? `— ${String(row[secondColKey])}` : ""
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

              {/* Data Preview Box */}
              <div>
                <div style={{ fontSize: "0.82rem", fontWeight: 700, color: "var(--text-muted)", marginBottom: "6px" }}>
                  Mapped Values for Row #{selectedPreviewRow + 1}:
                </div>
                <div className={styles.previewBox}>
                  {wordPlaceholders.map((ph) => {
                    const col = mappings[ph];
                    const val =
                      col && excelData?.rows[selectedPreviewRow]?.[col] !== undefined
                        ? String(excelData.rows[selectedPreviewRow][col])
                        : defaultValues[ph] || "(empty)";

                    return (
                      <div key={ph} className={styles.previewItem}>
                        <span className={styles.previewKey}>&#123;{ph}&#125;</span>
                        <span className={styles.previewVal}>{val}</span>
                      </div>
                    );
                  })}
                </div>
              </div>

              <button
                className="btn btn-secondary"
                style={{ width: "100%", marginTop: "auto" }}
                onClick={handleDownloadSingle}
                disabled={isGenerating}
              >
                <Download size={16} /> Download Single Document (.docx)
              </button>
            </div>

            {/* Option B: Batch Generation & ZIP Download */}
            <div className={`glass-panel ${styles.generateCard}`} style={{ border: "1px solid var(--border-glow)" }}>
              <div>
                <div style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: "8px" }}>
                  <FileArchive size={24} style={{ color: "var(--accent-cyan)" }} />
                  <h3 style={{ fontSize: "1.2rem", fontWeight: 700, color: "#ffffff" }}>
                    Batch Export All Records
                  </h3>
                </div>
                <p className={styles.cardDesc}>
                  Generate separate customized Word documents for all <strong>{excelData?.totalRows || 0}</strong> rows in the sheet, packaged neatly in a single ZIP download.
                </p>
              </div>

              {/* Dynamic File Naming Configuration */}
              <div className={styles.namingPatternSection}>
                <label style={{ fontSize: "0.85rem", color: "var(--text-muted)", fontWeight: 600 }}>
                  Output File Naming Pattern:
                </label>
                <input
                  type="text"
                  value={namingPattern}
                  onChange={(e) => setNamingPattern(e.target.value)}
                  placeholder="e.g. {File_Name}.docx or Report_{name}.docx"
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
                  const previewFilename = computeDocFilename(namingPattern, previewPayload, selectedPreviewRow);
                  return (
                    <div style={{
                      marginTop: "8px",
                      padding: "10px 14px",
                      background: "rgba(0, 0, 0, 0.4)",
                      borderRadius: "var(--radius-sm)",
                      border: "1px solid var(--border-glow)",
                      display: "flex",
                      flexDirection: "column",
                      gap: "3px"
                    }}>
                      <div style={{ fontSize: "0.75rem", color: "var(--text-muted)", textTransform: "uppercase", letterSpacing: "0.04em", fontWeight: 700 }}>
                        Live Filename Preview (Row #{selectedPreviewRow + 1}):
                      </div>
                      <div style={{ fontSize: "0.92rem", color: "var(--accent-cyan)", fontFamily: "var(--font-code)", fontWeight: 700, wordBreak: "break-all" }}>
                        📄 {previewFilename}
                      </div>
                    </div>
                  );
                })()}

                <div style={{ fontSize: "0.78rem", color: "var(--text-dim)", marginTop: "6px" }}>
                  Click to insert field into filename pattern:
                </div>
                <div className={styles.helperChips}>
                  {/* Excel Column Chips */}
                  {excelData?.headers.map((h) => (
                    <button
                      key={`col-${h}`}
                      type="button"
                      className={styles.chipBtn}
                      style={{ borderColor: "rgba(16, 185, 129, 0.4)", color: "var(--accent-emerald)" }}
                      title={`Insert Excel Column "${h}"`}
                      onClick={() => setNamingPattern(`{${h}}.docx`)}
                    >
                      +&#123;{h}&#125;
                    </button>
                  ))}
                  {/* Word Placeholder Chips */}
                  {wordPlaceholders
                    .filter((ph) => !excelData?.headers.includes(ph))
                    .map((ph) => (
                      <button
                        key={`ph-${ph}`}
                        type="button"
                        className={styles.chipBtn}
                        title={`Insert Placeholder "${ph}"`}
                        onClick={() => setNamingPattern((prev) => `${prev.replace(".docx", "")}_{${ph}}.docx`)}
                      >
                        +&#123;{ph}&#125;
                      </button>
                    ))}
                </div>
              </div>

              {/* Batch Generation Summary Card */}
              <div
                style={{
                  background: "rgba(0, 0, 0, 0.25)",
                  padding: "16px",
                  borderRadius: "var(--radius-md)",
                  border: "1px solid var(--border-subtle)",
                  display: "flex",
                  flexDirection: "column",
                  gap: "8px",
                  fontSize: "0.86rem",
                }}
              >
                <div style={{ display: "flex", justifyContent: "space-between" }}>
                  <span style={{ color: "var(--text-muted)" }}>Total Documents:</span>
                  <strong style={{ color: "#ffffff" }}>{excelData?.totalRows || 0} Files</strong>
                </div>
                <div style={{ display: "flex", justifyContent: "space-between" }}>
                  <span style={{ color: "var(--text-muted)" }}>Mapped Fields:</span>
                  <strong style={{ color: "var(--accent-emerald)" }}>
                    {mappedCount} / {wordPlaceholders.length}
                  </strong>
                </div>
                <div style={{ display: "flex", justifyContent: "space-between" }}>
                  <span style={{ color: "var(--text-muted)" }}>Archive Format:</span>
                  <strong style={{ color: "var(--accent-cyan)" }}>Standard .ZIP</strong>
                </div>
              </div>

              <button
                className="btn btn-primary"
                style={{
                  width: "100%",
                  padding: "14px 20px",
                  fontSize: "1rem",
                  marginTop: "auto",
                }}
                onClick={handleDownloadBatch}
                disabled={isGenerating || !excelData || excelData.rows.length === 0}
              >
                {isGenerating ? (
                  <>
                    <RefreshCw size={18} className="animate-spin" /> Generating Documents...
                  </>
                ) : (
                  <>
                    <Download size={18} /> Generate &amp; Download ZIP ({excelData?.totalRows || 0} Files)
                  </>
                )}
              </button>
            </div>
          </div>

          {/* Navigation Footer */}
          <div className={styles.navFooter}>
            <button className="btn btn-secondary" onClick={() => setActiveStep(2)}>
              <ChevronLeft size={16} /> Back to Field Mapping
            </button>
            <button className="btn btn-secondary" onClick={() => setActiveStep(1)}>
              Upload New Files
            </button>
          </div>
        </div>
      )}

      {/* Guide & Placeholder Documentation Modal */}
      {showHelpModal && (
        <div
          style={{
            position: "fixed",
            inset: 0,
            background: "rgba(0, 0, 0, 0.75)",
            backdropFilter: "blur(8px)",
            display: "flex",
            alignItems: "flex-end",
            justifyContent: "center",
            zIndex: 10000,
            padding: "12px",
          }}
          onClick={() => setShowHelpModal(false)}
        >
          <div
            className="glass-panel"
            style={{
              maxWidth: "600px",
              width: "100%",
              padding: "clamp(18px, 4vw, 32px)",
              background: "var(--bg-card-solid)",
              maxHeight: "85vh",
              overflowY: "auto",
              borderRadius: "var(--radius-lg)",
              marginBottom: "env(safe-area-inset-bottom, 0px)",
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                marginBottom: "20px",
              }}
            >
              <h3 style={{ fontSize: "1.25rem", fontWeight: 800, color: "#ffffff" }}>
                How to Format Your Word &amp; Excel Files
              </h3>
              <button
                onClick={() => setShowHelpModal(false)}
                style={{ background: "none", border: "none", color: "var(--text-muted)", cursor: "pointer" }}
              >
                <X size={20} />
              </button>
            </div>

            <div style={{ display: "flex", flexDirection: "column", gap: "18px", fontSize: "0.9rem", color: "var(--text-muted)" }}>
              <div>
                <strong style={{ color: "#ffffff", display: "block", marginBottom: "6px" }}>
                  1. Word Template Placeholders
                </strong>
                <p>
                  In your Microsoft Word document, write your variables inside curly braces, for example:
                  <span className="code-tag" style={{ margin: "0 6px" }}>&#123;Candidate_Name&#125;</span>,
                  <span className="code-tag" style={{ margin: "0 6px" }}>&#123;Job_Title&#125;</span>, or
                  <span className="code-tag" style={{ margin: "0 6px" }}>&#123;Annual_Salary&#125;</span>.
                </p>
              </div>

              <div>
                <strong style={{ color: "#ffffff", display: "block", marginBottom: "6px" }}>
                  2. Excel Column Headers
                </strong>
                <p>
                  Ensure the first row of your Excel sheet contains clear headers (e.g., "Candidate_Name", "Job_Title"). DocuBridge will automatically match them to Word placeholders with identical or similar names.
                </p>
              </div>

              <div>
                <strong style={{ color: "#ffffff", display: "block", marginBottom: "6px" }}>
                  3. Batch Generation &amp; Privacy
                </strong>
                <p>
                  All document generation and parsing runs 100% locally in your browser. Your confidential spreadsheet records and Word templates never leave your computer or get uploaded to third-party cloud servers.
                </p>
              </div>
            </div>

            <button
              className="btn btn-primary"
              style={{ width: "100%", marginTop: "24px" }}
              onClick={() => setShowHelpModal(false)}
            >
              Got it!
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
