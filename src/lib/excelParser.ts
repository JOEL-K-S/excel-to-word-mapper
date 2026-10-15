import * as XLSX from "xlsx";

export interface ExcelSheetData {
  sheetNames: string[];
  activeSheet: string;
  headers: string[];
  rows: Record<string, any>[];
  totalRows: number;
}

/**
 * Parses an Excel or CSV file buffer and returns structured sheet data.
 */
export function parseExcelFile(
  data: ArrayBuffer,
  sheetName?: string
): ExcelSheetData {
  const workbook = XLSX.read(data, {
    type: "array",
    cellDates: true,
    cellNF: false,
    cellText: false,
  });

  const sheetNames = workbook.SheetNames;
  if (!sheetNames || sheetNames.length === 0) {
    throw new Error("The uploaded Excel workbook contains no sheets.");
  }

  const targetSheet = sheetName && sheetNames.includes(sheetName)
    ? sheetName
    : sheetNames[0];

  const worksheet = workbook.Sheets[targetSheet];

  // Parse rows as raw JSON array of objects
  const rawRows: Record<string, any>[] = XLSX.utils.sheet_to_json(worksheet, {
    raw: false,
    dateNF: "yyyy-mm-dd",
    defval: "",
  });

  // Extract all distinct headers across rows or from worksheet range
  let headers: string[] = [];
  if (rawRows.length > 0) {
    const headerSet = new Set<string>();
    rawRows.forEach((row) => {
      Object.keys(row).forEach((k) => headerSet.add(k.trim()));
    });
    headers = Array.from(headerSet);
  } else {
    // If empty rows, try reading header row from range
    const range = XLSX.utils.decode_range(worksheet["!ref"] || "A1:A1");
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cell = worksheet[XLSX.utils.encode_cell({ r: range.s.r, c: C })];
      if (cell && cell.v) headers.push(String(cell.v).trim());
    }
  }

  return {
    sheetNames,
    activeSheet: targetSheet,
    headers,
    rows: rawRows,
    totalRows: rawRows.length,
  };
}

/**
 * Creates a sample Excel workbook buffer for demonstration & instant testing.
 */
export function generateSampleExcel(): ArrayBuffer {
  const sampleData = [
    {
      "Employee_ID": "EMP-1001",
      "Employee_Name": "Sophia Vance",
      "Job_Title": "Senior Frontend Architect",
      "Department": "Engineering",
      "Start_Date": "2026-11-01",
      "Annual_Salary": "$145,000",
      "Manager_Name": "David Chen",
      "Office_Location": "San Francisco, CA",
      "Email": "sophia.vance@acmecorp.io",
    },
    {
      "Employee_ID": "EMP-1002",
      "Employee_Name": "Marcus Sterling",
      "Job_Title": "Lead Product Designer",
      "Department": "User Experience",
      "Start_Date": "2026-11-15",
      "Annual_Salary": "$138,000",
      "Manager_Name": "Elena Rostova",
      "Office_Location": "New York, NY",
      "Email": "marcus.sterling@acmecorp.io",
    },
    {
      "Employee_ID": "EMP-1003",
      "Employee_Name": "Aaliyah Patel",
      "Job_Title": "Cloud Infrastructure Lead",
      "Department": "DevOps & Cloud",
      "Start_Date": "2026-12-01",
      "Annual_Salary": "$152,000",
      "Manager_Name": "David Chen",
      "Office_Location": "Austin, TX",
      "Email": "aaliyah.patel@acmecorp.io",
    },
    {
      "Employee_ID": "EMP-1004",
      "Employee_Name": "Lucas Silva",
      "Job_Title": "AI Solutions Engineer",
      "Department": "Applied Intelligence",
      "Start_Date": "2026-12-05",
      "Annual_Salary": "$160,000",
      "Manager_Name": "Rachel Scott",
      "Office_Location": "Seattle, WA",
      "Email": "lucas.silva@acmecorp.io",
    },
    {
      "Employee_ID": "EMP-1005",
      "Employee_Name": "Emily Zhao",
      "Job_Title": "Principal Security Analyst",
      "Department": "Cybersecurity",
      "Start_Date": "2026-12-10",
      "Annual_Salary": "$148,000",
      "Manager_Name": "Rachel Scott",
      "Office_Location": "Boston, MA",
      "Email": "emily.zhao@acmecorp.io",
    },
  ];

  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(sampleData);

  // Set column widths
  worksheet["!cols"] = [
    { wch: 14 },
    { wch: 20 },
    { wch: 28 },
    { wch: 20 },
    { wch: 14 },
    { wch: 16 },
    { wch: 18 },
    { wch: 20 },
    { wch: 28 },
  ];

  XLSX.utils.book_append_sheet(workbook, worksheet, "Candidates");
  return XLSX.write(workbook, { bookType: "xlsx", type: "array" });
}
