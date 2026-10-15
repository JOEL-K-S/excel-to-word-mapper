import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import JSZip from "jszip";

export interface PlaceholderInfo {
  tag: string;
  count: number;
}

export interface GenerationProgress {
  current: number;
  total: number;
  percentage: number;
  status: string;
}

/**
 * Normalizes XML content in the docx zip:
 * - Trims whitespace inside placeholders, e.g. {  name  } -> {name}
 * - Converts Word Mail Merge guillemets «name» -> {name}
 */
export function normalizeDocxZip(zip: PizZip): PizZip {
  const xmlFileNames = Object.keys(zip.files).filter((fileName) => {
    return (
      fileName.startsWith("word/") &&
      (fileName.includes("document.xml") ||
        fileName.includes("header") ||
        fileName.includes("footer")) &&
      fileName.endsWith(".xml")
    );
  });

  xmlFileNames.forEach((fileName) => {
    const file = zip.files[fileName];
    if (!file) return;

    let xmlContent = file.asText();

    // Convert Word Mail Merge guillemets «field» or &laquo;field&raquo; into {field}
    xmlContent = xmlContent.replace(/«\s*([a-zA-Z0-9_\-\.\s#^/]+)\s*»/g, "{$1}");
    xmlContent = xmlContent.replace(/&laquo;\s*([a-zA-Z0-9_\-\.\s#^/]+)\s*&raquo;/g, "{$1}");

    // Normalize spacing inside braces in XML: { field } -> {field}
    xmlContent = xmlContent.replace(/\{\s+([a-zA-Z0-9_\-\.]+)\s+\}/g, "{$1}");

    zip.file(fileName, xmlContent);
  });

  return zip;
}

/**
 * Robustly inspects a DOCX file and extracts all template placeholders.
 * Supports {placeholder}, {{placeholder}}, and «placeholder» syntax across document body, headers, and footers.
 */
export function extractPlaceholdersFromDocx(arrayBuffer: ArrayBuffer): string[] {
  try {
    const zip = new PizZip(arrayBuffer);
    const placeholdersSet = new Set<string>();

    const xmlFileNames = Object.keys(zip.files).filter((fileName) => {
      return (
        fileName.startsWith("word/") &&
        (fileName.includes("document.xml") ||
          fileName.includes("header") ||
          fileName.includes("footer")) &&
        fileName.endsWith(".xml")
      );
    });

    xmlFileNames.forEach((fileName) => {
      const file = zip.files[fileName];
      if (!file) return;

      const xmlContent = file.asText();

      // Strip all XML tags to reconstruct continuous text (solves MS Word split w:t nodes)
      const plainText = xmlContent.replace(/<[^>]+>/g, "");

      // Match {placeholder}, {{placeholder}}, and «placeholder»
      const regex = /(\{+|«+)([a-zA-Z0-9_\-\.\s#^/]+)(\}+|»+)/g;
      let match: RegExpExecArray | null;

      while ((match = regex.exec(plainText)) !== null) {
        let tag = match[2].trim();
        // Remove loop / condition prefixes like #, ^, / if present
        if (tag.startsWith("#") || tag.startsWith("^") || tag.startsWith("/")) {
          tag = tag.substring(1).trim();
        }
        if (tag && !tag.includes("<") && !tag.includes(">") && tag.length <= 80) {
          placeholdersSet.add(tag);
        }
      }
    });

    return Array.from(placeholdersSet);
  } catch (error) {
    console.error("Error inspecting docx placeholders:", error);
    throw new Error(
      "Unable to read Word document. Please ensure it is a valid .docx file."
    );
  }
}

/**
 * Fortifies and sanitizes file names:
 * - Removes null bytes (\0)
 * - Replaces path traversal indicators (..)
 * - Replaces forbidden OS characters (/ \ ? % * : | " < >)
 * - Strips leading dots/dashes to prevent hidden files
 * - Enforces safe character length limit (180 chars)
 */
export function sanitizeFilename(raw: string): string {
  if (!raw) return "";
  let clean = raw
    .replace(/\0/g, "")
    .replace(/[/\\?%*:|"<>]/g, "_")
    .replace(/\.{2,}/g, "_")
    .replace(/^[\s._-]+/, "")
    .trim();
  return clean.slice(0, 180);
}

/**
 * Intelligently computes a safe, non-empty filename for a generated document.
 * Guarantees a clean valid .docx extension.
 */
export function computeDocFilename(
  pattern: string,
  rowData: Record<string, any>,
  rowIndex: number
): string {
  let name = (pattern || "").trim();

  // If pattern has placeholders, replace them (case-insensitive)
  Object.keys(rowData).forEach((key) => {
    const rawVal = rowData[key];
    if (rawVal !== undefined && rawVal !== null && String(rawVal).trim() !== "") {
      const cleanVal = sanitizeFilename(String(rawVal));
      name = name.replace(new RegExp(`\\{${key}\\}`, "gi"), cleanVal);
    }
  });

  // Strip any remaining unreplaced {placeholder} tags
  name = name.replace(/\{[^}]+\}/g, "").trim();

  // Strip any existing extension (.docx, .docs, .doc, etc.) before cleaning
  name = name.replace(/\.[a-zA-Z0-9]+$/i, "").trim();

  // Strip trailing punctuation like "Document_", "Document -", etc.
  name = name.replace(/[_\-\s]+$/, "").trim();

  // Sanitize the assembled name
  name = sanitizeFilename(name);

  // If filename became empty or left as generic "Document"
  if (!name || name.toLowerCase() === "document") {
    // Check if row data has a file name or person name column
    const candidateKeys = [
      "File_Name",
      "filename",
      "name",
      "student_name",
      "candidate_name",
      "employee_name",
      "employee_id",
      "id",
    ];

    let foundVal = "";
    for (const key of Object.keys(rowData)) {
      if (candidateKeys.includes(key.toLowerCase()) && rowData[key]) {
        foundVal = sanitizeFilename(String(rowData[key])).replace(/\.[a-zA-Z0-9]+$/i, "");
        break;
      }
    }

    if (foundVal) {
      name = foundVal;
    } else {
      name = `Document_Row_${rowIndex + 1}`;
    }
  }

  // Ensure exactly one clean .docx extension
  return `${name}.docx`;
}

/**
 * Generates a single filled DOCX document using docxtemplater.
 */
export function generateSingleDocx(
  templateBuffer: ArrayBuffer,
  data: Record<string, any>
): Blob {
  let zip = new PizZip(templateBuffer);
  zip = normalizeDocxZip(zip);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    nullGetter: () => "", // Don't crash on null/undefined, output blank
  });

  // Render the document with mapped values
  doc.render(data);

  const out = doc.getZip().generate({
    type: "blob",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });

  return out;
}

/**
 * Converts a generated DOCX Blob into a crisp, vector-rendered PDF Blob entirely in-browser.
 * Dynamically loads docx-preview, jspdf, and html2canvas on demand for optimal initial bundle size.
 */
export async function convertDocxBlobToPdfBlob(docxBlob: Blob): Promise<Blob> {
  const [docx, { jsPDF }, html2canvasModule] = await Promise.all([
    import("docx-preview"),
    import("jspdf"),
    import("html2canvas"),
  ]);
  const html2canvas = html2canvasModule.default || html2canvasModule;

  const container = document.createElement("div");
  container.style.position = "fixed";
  container.style.top = "-99999px";
  container.style.left = "-99999px";
  container.style.width = "794px"; // Standard A4 width at 96 DPI
  container.style.background = "#ffffff";
  container.style.color = "#000000";
  container.style.zIndex = "-1000";
  document.body.appendChild(container);

  try {
    const arrayBuffer = await docxBlob.arrayBuffer();

    await docx.renderAsync(arrayBuffer, container, undefined, {
      inWrapper: true,
      ignoreWidth: false,
      ignoreHeight: false,
      ignoreFonts: false,
      breakPages: true,
      useBase64URL: true,
    });

    // Let the DOM compute styles
    await new Promise((resolve) => setTimeout(resolve, 300));

    // Find all docx page wrapper sections
    const pageWrappers = container.querySelectorAll(".docx-wrapper > section.docx, section.docx");

    const pdf = new jsPDF({
      orientation: "portrait",
      unit: "pt",
      format: "a4",
      compress: true,
    });

    const pdfWidth = pdf.internal.pageSize.getWidth();
    const pdfHeight = pdf.internal.pageSize.getHeight();

    if (pageWrappers.length > 0) {
      for (let p = 0; p < pageWrappers.length; p++) {
        const pageEl = pageWrappers[p] as HTMLElement;
        const canvas = await html2canvas(pageEl, {
          scale: 2,
          useCORS: true,
          logging: false,
          backgroundColor: "#ffffff",
        });

        if (p > 0) {
          pdf.addPage("a4", "portrait");
        }

        const imgData = canvas.toDataURL("image/jpeg", 0.95);
        pdf.addImage(imgData, "JPEG", 0, 0, pdfWidth, pdfHeight, undefined, "FAST");
      }
    } else {
      const canvas = await html2canvas(container, {
        scale: 2,
        useCORS: true,
        logging: false,
        backgroundColor: "#ffffff",
      });
      const imgData = canvas.toDataURL("image/jpeg", 0.95);
      const imgHeight = (canvas.height * pdfWidth) / canvas.width;
      pdf.addImage(imgData, "JPEG", 0, 0, pdfWidth, Math.min(imgHeight, pdfHeight), undefined, "FAST");
    }

    const pdfOutput = pdf.output("blob");
    return pdfOutput;
  } finally {
    if (container.parentNode) {
      container.parentNode.removeChild(container);
    }
  }
}

/**
 * Triggers a bulletproof browser download that guarantees the file name and extension
 * are preserved and never lost or aborted by the browser across all operating systems.
 */
export async function triggerBrowserDownload(blob: Blob, rawFilename: string) {
  let ext = ".docx";
  const lower = (rawFilename || "").toLowerCase().trim();
  if (lower.endsWith(".zip")) {
    ext = ".zip";
  } else if (lower.endsWith(".xlsx")) {
    ext = ".xlsx";
  } else if (lower.endsWith(".pdf")) {
    ext = ".pdf";
  } else {
    ext = ".docx";
  }

  // Strip ANY trailing extension from rawFilename so we never get double extensions
  let baseName = (rawFilename || "").replace(/\.[a-zA-Z0-9]+$/i, "").trim();
  baseName = sanitizeFilename(baseName) || "Generated_Document";

  // The true, guaranteed final filename
  const filename = `${baseName}${ext}`;

  // Map explicit MIME types for guaranteed OS file association
  const MIME_TYPES: Record<string, string> = {
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".zip": "application/zip",
    ".pdf": "application/pdf",
  };

  const targetMime = MIME_TYPES[ext] || blob.type || "application/octet-stream";

  // Re-wrap Blob with explicit MIME type
  const typedBlob = new Blob([blob], { type: targetMime });

  // Create a File object when available so the browser natively binds the filename and MIME
  let fileObj: Blob = typedBlob;
  try {
    if (typeof File !== "undefined") {
      fileObj = new File([typedBlob], filename, {
        type: targetMime,
        lastModified: Date.now(),
      });
    }
  } catch (e) {
    fileObj = typedBlob;
  }

  // Mobile Web Share API support: on mobile touch devices (Android Chrome, iOS Safari),
  // sharing the file directly to the system allows instant "Save to Downloads" or "Open in Word"
  if (
    typeof navigator !== "undefined" &&
    typeof navigator.share === "function" &&
    typeof navigator.canShare === "function"
  ) {
    try {
      const shareData = {
        files: [fileObj as File],
        title: filename,
      };
      if (navigator.canShare(shareData)) {
        const isTouchMobile = /Android|iPhone|iPad|iPod/i.test(navigator.userAgent || "");
        if (isTouchMobile) {
          await navigator.share(shareData);
          return;
        }
      }
    } catch (e: any) {
      if (e && e.name === "AbortError") {
        return; // User cancelled share sheet
      }
      // fallback to regular download link below
    }
  }

  const url = URL.createObjectURL(fileObj);
  const link = document.createElement("a");
  link.style.display = "none";
  link.href = url;
  link.download = filename;
  link.setAttribute("download", filename);
  link.setAttribute("data-downloadurl", `${targetMime}:${filename}:${url}`);
  document.body.appendChild(link);

  try {
    link.dispatchEvent(
      new MouseEvent("click", {
        bubbles: true,
        cancelable: true,
        view: window,
      })
    );
  } catch (e) {
    link.click();
  }

  // Wait 60 seconds before revoking the URL so the browser download manager has finished saving
  setTimeout(() => {
    try {
      if (link.parentNode) {
        link.parentNode.removeChild(link);
      }
      URL.revokeObjectURL(url);
    } catch (e) {
      // ignore
    }
  }, 60000);
}

/**
 * Generates batch documents for all rows and bundles them into a ZIP archive.
 * Supports exporting individual documents as either .docx or .pdf.
 */
export async function generateBatchZip(
  templateBuffer: ArrayBuffer,
  rows: Record<string, any>[],
  mapping: Record<string, string>, // placeholder -> excelColumn
  defaultValues: Record<string, string>,
  fileNamePattern: string,
  exportFormat: "docx" | "pdf" = "docx",
  onProgress?: (progress: GenerationProgress) => void
): Promise<Blob> {
  const zip = new JSZip();
  const total = rows.length;

  for (let i = 0; i < total; i++) {
    const row = rows[i];

    // Construct data payload for this row based on mappings
    const rowData: Record<string, any> = {};
    Object.keys(mapping).forEach((placeholder) => {
      const colName = mapping[placeholder];
      if (colName && row[colName] !== undefined && row[colName] !== null) {
        rowData[placeholder] = row[colName];
      } else if (defaultValues[placeholder] !== undefined) {
        rowData[placeholder] = defaultValues[placeholder];
      } else {
        rowData[placeholder] = "";
      }
    });

    // Also include raw row fields just in case placeholder matches directly
    Object.keys(row).forEach((k) => {
      if (!(k in rowData)) {
        rowData[k] = row[k];
      }
    });

    // Generate individual docx
    const docBlob = generateSingleDocx(templateBuffer, rowData);
    let outputBlob: Blob;
    let filename: string;

    if (exportFormat === "pdf") {
      outputBlob = await convertDocxBlobToPdfBlob(docBlob);
      filename = computeDocFilename(fileNamePattern, rowData, i).replace(/\.docx$/i, ".pdf");
    } else {
      outputBlob = docBlob;
      filename = computeDocFilename(fileNamePattern, rowData, i);
    }

    // Add to ZIP
    zip.file(filename, outputBlob);

    // Report progress
    if (onProgress) {
      const current = i + 1;
      const percentage = Math.round((current / total) * 100);
      onProgress({
        current,
        total,
        percentage,
        status: `Generating ${exportFormat.toUpperCase()} document ${current} of ${total} (${filename})...`,
      });
    }

    // Small yield to keep UI thread responsive on large datasets
    if (i % 5 === 0) {
      await new Promise((r) => setTimeout(r, 0));
    }
  }

  if (onProgress) {
    onProgress({
      current: total,
      total,
      percentage: 100,
      status: "Compressing documents into ZIP archive...",
    });
  }

  const zipBlob = await zip.generateAsync({
    type: "blob",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
    mimeType: "application/zip",
  });

  return zipBlob;
}

/**
 * Creates a high quality pre-formatted sample Word (.docx) template with placeholders.
 */
export function generateSampleWordTemplate(): ArrayBuffer {
  const zip = new PizZip();

  // [Content_Types].xml
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
  );

  // _rels/.rels
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
);

  // word/document.xml
  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Title"/>
        <w:jc w:val="center"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:sz w:val="48"/>
          <w:color w:val="1E293B"/>
        </w:rPr>
        <w:t>OFFER OF EMPLOYMENT</w:t>
      </w:r>
    </w:p>

    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r>
        <w:rPr><w:i/><w:color w:val="64748B"/><w:sz w:val="22"/></w:rPr>
        <w:t>CONFIDENTIAL &amp; PROPRIETARY</w:t>
      </w:r>
    </w:p>

    <w:p><w:r><w:t></w:t></w:r></w:p>

    <w:p>
      <w:r>
        <w:rPr><w:b/><w:sz w:val="24"/></w:rPr>
        <w:t>Dear {Employee_Name},</w:t>
      </w:r>
    </w:p>

    <w:p>
      <w:r>
        <w:rPr><w:sz w:val="24"/></w:rPr>
        <w:t>We are delighted to extend an offer for the position of </w:t>
      </w:r>
      <w:r>
        <w:rPr><w:b/><w:color w:val="4F46E5"/><w:sz w:val="24"/></w:rPr>
        <w:t>{Job_Title}</w:t>
      </w:r>
      <w:r>
        <w:rPr><w:sz w:val="24"/></w:rPr>
        <w:t> within the </w:t>
      </w:r>
      <w:r>
        <w:rPr><w:b/><w:sz w:val="24"/></w:rPr>
        <w:t>{Department}</w:t>
      </w:r>
      <w:r>
        <w:rPr><w:sz w:val="24"/></w:rPr>
        <w:t> division at Acme Technologies.</w:t>
      </w:r>
    </w:p>

    <w:p><w:r><w:t></w:t></w:r></w:p>

    <w:p>
      <w:r>
        <w:rPr><w:b/><w:sz w:val="26"/><w:color w:val="0F172A"/></w:rPr>
        <w:t>Key Terms of Employment:</w:t>
      </w:r>
    </w:p>

    <w:p>
      <w:r><w:rPr><w:b/><w:sz w:val="22"/></w:rPr><w:t>• Candidate Name: </w:t></w:r>
      <w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>{Employee_Name}</w:t></w:r>
    </w:p>

    <w:p>
      <w:r><w:rPr><w:b/><w:sz w:val="22"/></w:rPr><w:t>• Official Title: </w:t></w:r>
      <w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>{Job_Title}</w:t></w:r>
    </w:p>

    <w:p>
      <w:r><w:rPr><w:b/><w:sz w:val="22"/></w:rPr><w:t>• Department: </w:t></w:r>
      <w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>{Department}</w:t></w:r>
    </w:p>

    <w:p>
      <w:r><w:rPr><w:b/><w:sz w:val="22"/></w:rPr><w:t>• Effective Start Date: </w:t></w:r>
      <w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>{Start_Date}</w:t></w:r>
    </w:p>

    <w:p>
      <w:r><w:rPr><w:b/><w:sz w:val="22"/></w:rPr><w:t>• Annual Base Compensation: </w:t></w:r>
      <w:r><w:rPr><w:b/><w:color w:val="059669"/><w:sz w:val="22"/></w:rPr><w:t>{Annual_Salary}</w:t></w:r>
    </w:p>

    <w:p>
      <w:r><w:rPr><w:b/><w:sz w:val="22"/></w:rPr><w:t>• Office Location: </w:t></w:r>
      <w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>{Office_Location}</w:t></w:r>
    </w:p>

    <w:p>
      <w:r><w:rPr><w:b/><w:sz w:val="22"/></w:rPr><w:t>• Reporting Manager: </w:t></w:r>
      <w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>{Manager_Name}</w:t></w:r>
    </w:p>

    <w:p><w:r><w:t></w:t></w:r></w:p>

    <w:p>
      <w:r>
        <w:rPr><w:sz w:val="24"/></w:rPr>
        <w:t>We look forward to welcoming you aboard and achieving great milestones together.</w:t>
      </w:r>
    </w:p>

    <w:p><w:r><w:t></w:t></w:r></w:p>
    <w:p><w:r><w:t></w:t></w:r></w:p>

    <w:p>
      <w:r>
        <w:rPr><w:b/><w:sz w:val="24"/></w:rPr>
        <w:t>Sincerely,</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:rPr><w:i/><w:sz w:val="24"/></w:rPr>
        <w:t>{Manager_Name}</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:r>
        <w:rPr><w:color w:val="64748B"/><w:sz w:val="22"/></w:rPr>
        <w:t>Acme Corporation Executive Team</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

  zip.file("word/document.xml", docXml);

  return zip.generate({ type: "arraybuffer" });
}
