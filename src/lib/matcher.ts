/**
 * Normalizes a key string for flexible matching (e.g., 'Employee_Name' -> 'employeename').
 */
export function normalizeKey(str: string): string {
  return str.toLowerCase().replace(/[^a-z0-9]/g, "");
}

/**
 * Tokenizes a string into distinct lowercase words:
 * e.g. "Employee_Name" or "candidateName" -> ["employee", "name"]
 */
export function tokenizeKey(str: string): string[] {
  return str
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    .replace(/[^a-zA-Z0-9]/g, " ")
    .toLowerCase()
    .split(/\s+/)
    .filter((w) => w.length > 0);
}

// Strictly separated semantic synonym groups
const SYNONYM_GROUPS: string[][] = [
  ["candidate", "employee", "applicant", "person", "client", "worker", "recipient"],
  ["name", "fullname", "recipientname", "candidatename", "employeename"],
  ["first", "firstname", "givenname"],
  ["last", "lastname", "surname", "familyname"],
  ["job", "position", "role", "title", "designation", "jobtitle", "positiontitle"],
  ["salary", "compensation", "pay", "wage", "ctc", "annualsalary", "remuneration", "stipend"],
  ["address", "location", "officelocation", "street", "city", "place", "state", "candidateaddress"],
  ["start", "startdate", "joining", "dateofjoining", "hire", "hiredate", "effective", "effectivedate"],
  ["manager", "supervisor", "lead", "boss", "managername", "reportingmanager"],
  ["date", "offerdate", "issuedate", "createddate"],
  ["deadline", "expiry", "due", "duedate", "validuntil", "validto", "expiration"],
  ["department", "division", "team", "dept", "unit", "group", "costcenter"],
  ["email", "mail", "emailaddress"],
  ["phone", "mobile", "contact", "cell", "phonenumber", "tel"],
  ["id", "code", "number", "num", "badge", "identifier"],
];

function getTokenMatchScore(wordA: string, wordB: string): number {
  if (wordA === wordB) return 1.0;
  for (const group of SYNONYM_GROUPS) {
    if (group.includes(wordA) && group.includes(wordB)) {
      return 0.8;
    }
  }
  // Stemming / prefix match
  if (wordA.length >= 4 && wordB.length >= 4) {
    if (wordA.startsWith(wordB) || wordB.startsWith(wordA)) {
      return 0.7;
    }
  }
  return 0;
}

/**
 * Computes a fine-grained similarity score between a template placeholder and an Excel header (0 to 1).
 */
export function calculateMatchScore(placeholder: string, header: string): number {
  const normPh = normalizeKey(placeholder);
  const normH = normalizeKey(header);

  // 1. Exact match (case-insensitive)
  if (normPh === normH) return 1.0;

  const tokensPh = tokenizeKey(placeholder);
  const tokensH = tokenizeKey(header);

  // Disqualify if header is an ID/code and placeholder is NOT looking for an ID/code
  const phIsId = tokensPh.includes("id") || tokensPh.includes("code") || tokensPh.includes("number");
  const hIsId = tokensH.includes("id") || tokensH.includes("code") || tokensH.includes("number");
  if (hIsId && !phIsId) {
    return 0;
  }

  // 2. Direct substring match
  if (normH.includes(normPh) || normPh.includes(normH)) {
    const minLen = Math.min(normPh.length, normH.length);
    const maxLen = Math.max(normPh.length, normH.length);
    const ratio = minLen / maxLen;
    // e.g. "salary" inside "annual_salary"
    return 0.85 + 0.1 * ratio;
  }

  // 3. Token-by-token comparison
  let matchedTokensScore = 0;
  const usedHTokens = new Set<number>();

  for (const tPh of tokensPh) {
    let bestTokenScore = 0;
    let bestHIndex = -1;

    for (let j = 0; j < tokensH.length; j++) {
      if (usedHTokens.has(j)) continue;
      const score = getTokenMatchScore(tPh, tokensH[j]);
      if (score > bestTokenScore) {
        bestTokenScore = score;
        bestHIndex = j;
      }
    }

    if (bestTokenScore > 0 && bestHIndex >= 0) {
      matchedTokensScore += bestTokenScore;
      usedHTokens.add(bestHIndex);
    }
  }

  if (matchedTokensScore > 0) {
    const avgLen = (tokensPh.length + tokensH.length) / 2;
    const ratio = matchedTokensScore / avgLen;
    // Strong base score for matching semantic concepts
    const score = 0.45 + 0.5 * ratio;
    return Math.min(0.95, score);
  }

  return 0;
}

/**
 * Automatically creates an optimal mapping between Word placeholders and Excel column headers.
 */
export function autoMapPlaceholders(
  placeholders: string[],
  headers: string[]
): Record<string, string> {
  const mapping: Record<string, string> = {};
  if (!headers || headers.length === 0) {
    placeholders.forEach((p) => {
      mapping[p] = "";
    });
    return mapping;
  }

  placeholders.forEach((placeholder) => {
    let bestHeader = "";
    let bestScore = 0;

    for (const header of headers) {
      const score = calculateMatchScore(placeholder, header);
      if (score > bestScore) {
        bestScore = score;
        bestHeader = header;
      }
    }

    // Require at least 0.5 similarity to auto-map
    if (bestScore >= 0.5) {
      mapping[placeholder] = bestHeader;
    } else {
      mapping[placeholder] = "";
    }
  });

  return mapping;
}
