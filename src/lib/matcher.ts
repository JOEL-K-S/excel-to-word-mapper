/**
 * Normalizes a key string for flexible matching (e.g., 'Employee_Name' -> 'employeename').
 */
export function normalizeKey(str: string): string {
  return str.toLowerCase().replace(/[^a-z0-9]/g, "");
}

/**
 * Automatically creates a mapping between Word placeholders and Excel column headers.
 */
export function autoMapPlaceholders(
  placeholders: string[],
  headers: string[]
): Record<string, string> {
  const mapping: Record<string, string> = {};
  const normalizedHeaders = headers.map((h) => ({
    original: h,
    normalized: normalizeKey(h),
  }));

  placeholders.forEach((placeholder) => {
    const normPlaceholder = normalizeKey(placeholder);

    // 1. Exact match (case-sensitive)
    if (headers.includes(placeholder)) {
      mapping[placeholder] = placeholder;
      return;
    }

    // 2. Exact match normalized (ignoring spaces, underscores, casing)
    const exactNorm = normalizedHeaders.find(
      (h) => h.normalized === normPlaceholder
    );
    if (exactNorm) {
      mapping[placeholder] = exactNorm.original;
      return;
    }

    // 3. Substring containment match (e.g. 'salary' matches 'annual_salary' or vice versa)
    const containsMatch = normalizedHeaders.find(
      (h) =>
        h.normalized.includes(normPlaceholder) ||
        normPlaceholder.includes(h.normalized)
    );
    if (containsMatch) {
      mapping[placeholder] = containsMatch.original;
      return;
    }

    // Default to unmapped
    mapping[placeholder] = "";
  });

  return mapping;
}
