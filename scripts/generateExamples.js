const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const PizZip = require('pizzip');

const projectRoot = path.resolve(__dirname, '..');
const publicExamplesDir = path.join(projectRoot, 'public', 'examples');
const examplesDir = path.join(projectRoot, 'examples');

[publicExamplesDir, examplesDir].forEach((dir) => {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
});

// 1. Generate sample Excel file
console.log('Generating sample_candidates.xlsx...');
const sampleCandidates = [
  {
    "Employee_ID": "EMP-1001",
    "Employee_Name": "Sophia Vance",
    "Job_Title": "Senior Frontend Architect",
    "Department": "Engineering",
    "Start_Date": "2026-11-01",
    "Annual_Salary": "$145,000",
    "Manager_Name": "David Chen",
    "Office_Location": "San Francisco, CA",
    "Email": "sophia.vance@acmecorp.io"
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
    "Email": "marcus.sterling@acmecorp.io"
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
    "Email": "aaliyah.patel@acmecorp.io"
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
    "Email": "lucas.silva@acmecorp.io"
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
    "Email": "emily.zhao@acmecorp.io"
  }
];

const wb = XLSX.utils.book_new();
const ws = XLSX.utils.json_to_sheet(sampleCandidates);
ws['!cols'] = [
  { wch: 14 },
  { wch: 20 },
  { wch: 28 },
  { wch: 20 },
  { wch: 14 },
  { wch: 16 },
  { wch: 18 },
  { wch: 20 },
  { wch: 28 }
];
XLSX.utils.book_append_sheet(wb, ws, "Candidates");

const excelBuffer = XLSX.write(wb, { bookType: "xlsx", type: "buffer" });
fs.writeFileSync(path.join(publicExamplesDir, 'sample_candidates.xlsx'), excelBuffer);
fs.writeFileSync(path.join(examplesDir, 'sample_candidates.xlsx'), excelBuffer);
console.log('✓ Created sample_candidates.xlsx in public/examples/ and examples/');

// 2. Generate sample Word (.docx) template
console.log('Generating sample_offer_letter.docx...');
const zip = new PizZip();

zip.file(
  "[Content_Types].xml",
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
);

zip.file(
  "_rels/.rels",
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
);

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
const wordBuffer = zip.generate({ type: "nodebuffer" });

fs.writeFileSync(path.join(publicExamplesDir, 'sample_offer_letter.docx'), wordBuffer);
fs.writeFileSync(path.join(examplesDir, 'sample_offer_letter.docx'), wordBuffer);
console.log('✓ Created sample_offer_letter.docx in public/examples/ and examples/');
console.log('All example files generated successfully!');
