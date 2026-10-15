import type { Metadata, Viewport } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import "./globals.css";

const geistSans = Geist({
  variable: "--font-geist-sans",
  subsets: ["latin"],
});

const geistMono = Geist_Mono({
  variable: "--font-geist-mono",
  subsets: ["latin"],
});

const siteUrl = "https://excel-to-word-mapper.vercel.app";

export const viewport: Viewport = {
  themeColor: "#4f46e5",
  width: "device-width",
  initialScale: 1,
};

export const metadata: Metadata = {
  metadataBase: new URL(siteUrl),
  title: {
    default: "Excel to Word Mapper | The First Web Excel to Report Online Creator",
    template: "%s | Excel to Word Mapper",
  },
  description:
    "The first 100% in-browser Web Excel to Report online creator. Transform spreadsheets (.xlsx, .csv) into personalized Word documents & PDF reports with zero server uploads.",
  keywords: [
    "first web excel to report online creator",
    "web excel to report online creator",
    "excel to report creator online",
    "online excel report maker",
    "automated excel to report generator",
    "excel to word mapper",
    "browser based report creator",
    "mail merge alternative",
    "bulk docx generator",
    "excel to pdf generator",
    "automated document generation",
    "batch word merge",
    "excel spreadsheet to word template",
    "free mail merge without word",
    "generate pdf from excel data",
    "client side mail merge",
    "privacy first document generator",
    "bulk certificate generator",
    "batch invoice generator excel",
    "docxtemplater online tool",
    "open source mail merge",
  ],
  authors: [{ name: "JOEL-K-S", url: "https://github.com/JOEL-K-S/excel-to-word-mapper" }],
  creator: "JOEL-K-S",
  publisher: "JOEL-K-S",
  applicationName: "Excel to Word Mapper",
  category: "Productivity & Business Utility",
  alternates: {
    canonical: siteUrl,
  },
  robots: {
    index: true,
    follow: true,
    googleBot: {
      index: true,
      follow: true,
      "max-video-preview": -1,
      "max-image-preview": "large",
      "max-snippet": -1,
    },
  },
  openGraph: {
    title: "Excel to Word Mapper | The First Web Excel to Report Online Creator",
    description:
      "The first 100% in-browser Web Excel to Report online creator. Transform spreadsheets into personalized Word documents & PDF reports with zero server uploads.",
    url: siteUrl,
    siteName: "Excel to Word Mapper",
    locale: "en_US",
    type: "website",
  },
  twitter: {
    card: "summary_large_image",
    title: "Excel to Word Mapper | The First Web Excel to Report Online Creator",
    description:
      "The first 100% in-browser Web Excel to Report online creator. Generate custom Word & PDF reports from spreadsheet data with zero cloud uploads.",
  },
};

const jsonLd = {
  "@context": "https://schema.org",
  "@graph": [
    {
      "@type": "WebApplication",
      "@id": `${siteUrl}/#webapp`,
      "name": "Excel to Word Mapper - The First Web Excel to Report Online Creator",
      "url": siteUrl,
      "description":
        "The first 100% in-browser Web Excel to Report online creator. Free, private tool that maps spreadsheet data into Word (.docx) templates and bulk exports personalized Word files or PDF reports in ZIP.",
      "applicationCategory": "BusinessApplication",
      "operatingSystem": "All modern browsers (Chrome, Edge, Firefox, Safari)",
      "offers": {
        "@type": "Offer",
        "price": "0",
        "priceCurrency": "USD",
      },
      "featureList": [
        "100% Client-Side Privacy (Zero Server Uploads)",
        "Intelligent Field Auto-Matching",
        "Dual Format Export: Word (.docx) and Adobe PDF (.pdf)",
        "Dynamic Custom Filename Formatting with {Column} tags",
        "Batch ZIP Export with Real-time Progress Tracking",
        "Interactive Single Record Test Mode",
      ],
      "browserRequirements": "Requires JavaScript. Requires HTML5.",
    },
    {
      "@type": "HowTo",
      "@id": `${siteUrl}/#howto`,
      "name": "How to Bulk Generate Word Documents and PDFs from Excel",
      "description":
        "Step-by-step tutorial to mail merge Excel spreadsheets into Word and PDF documents without Microsoft Office.",
      "step": [
        {
          "@type": "HowToStep",
          "name": "Upload Excel Data and Word Template",
          "text": "Upload an Excel workbook (.xlsx, .csv) and a Word document (.docx) containing {placeholder} tags.",
        },
        {
          "@type": "HowToStep",
          "name": "Match Placeholders to Columns",
          "text": "Review the automated column matching or customize pairings and fallback values for each field.",
        },
        {
          "@type": "HowToStep",
          "name": "Select Format and Download",
          "text": "Choose between Word (.docx) or PDF (.pdf) format, configure file naming patterns, and download single tests or a complete batch ZIP archive.",
        },
      ],
    },
    {
      "@type": "FAQPage",
      "@id": `${siteUrl}/#faq`,
      "mainEntity": [
        {
          "@type": "Question",
          "name": "Is my Excel and Word document data secure?",
          "acceptedAnswer": {
            "@type": "Answer",
            "text":
              "Yes, 100%. Excel to Word Mapper runs entirely inside your browser. No spreadsheets, customer records, or templates are ever sent to an external server or stored in the cloud.",
          },
        },
        {
          "@type": "Question",
          "name": "Can I export directly to PDF without Microsoft Word?",
          "acceptedAnswer": {
            "@type": "Answer",
            "text":
              "Yes. You can choose either Word (.docx) or PDF (.pdf) format before downloading. PDF conversion runs completely on the client side inside your browser.",
          },
        },
        {
          "@type": "Question",
          "name": "How do I format placeholders in my Word template?",
          "acceptedAnswer": {
            "@type": "Answer",
            "text":
              "Simply wrap your tag names in single curly braces inside your Word document, for example: {Customer_Name}, {Invoice_Date}, or {Total_Amount}.",
          },
        },
        {
          "@type": "Question",
          "name": "Can I give each generated document a unique filename?",
          "acceptedAnswer": {
            "@type": "Answer",
            "text":
              "Yes. You can use dynamic filename patterns such as {Customer_Name}_Invoice_{Invoice_No}.docx. The app will replace tags with each row's data automatically.",
          },
        },
        {
          "@type": "Question",
          "name": "Why should I use this instead of Microsoft Word Mail Merge?",
          "acceptedAnswer": {
            "@type": "Answer",
            "text":
              "Traditional mail merge requires desktop Microsoft Office, is cumbersome to split into individual files, and doesn't natively batch export into PDF or custom named files without complex VBA macros. Excel to Word Mapper does all this in seconds right in your web browser.",
          },
        },
      ],
    },
  ],
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en" className={`${geistSans.variable} ${geistMono.variable}`}>
      <head>
        <script
          type="application/ld+json"
          dangerouslySetInnerHTML={{ __html: JSON.stringify(jsonLd) }}
        />
      </head>
      <body>{children}</body>
    </html>
  );
}
