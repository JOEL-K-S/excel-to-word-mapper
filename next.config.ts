import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  /* config options here */
  cacheComponents: true,
  partialPrefetching: true,
  async headers() {
    return [
      {
        source: "/examples/sample_candidates.xlsx",
        headers: [
          {
            key: "Content-Disposition",
            value: 'attachment; filename="sample_candidates.xlsx"',
          },
          {
            key: "Content-Type",
            value: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          },
        ],
      },
      {
        source: "/examples/sample_offer_letter.docx",
        headers: [
          {
            key: "Content-Disposition",
            value: 'attachment; filename="sample_offer_letter.docx"',
          },
          {
            key: "Content-Type",
            value: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          },
        ],
      },
      {
        source: "/(.*)",
        headers: [
          {
            key: "X-Frame-Options",
            value: "DENY",
          },
          {
            key: "X-Content-Type-Options",
            value: "nosniff",
          },
          {
            key: "Referrer-Policy",
            value: "strict-origin-when-cross-origin",
          },
          {
            key: "Permissions-Policy",
            value: "camera=(), microphone=(), geolocation=(), payment=()",
          },
        ],
      },
    ];
  },
};

export default nextConfig;
