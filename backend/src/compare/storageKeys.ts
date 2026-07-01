/**
 * R2 object-key helpers for the Document Compare feature.
 *
 * Defines the key layout under the `comparisons/` prefix. Each comparison
 * stores two artifacts in R2:
 *   - the redline `.docx` (native Word tracked changes)
 *   - the diff `.json`   (ordered hunks for in-app rendering)
 *
 * These are pure string formatters with no I/O and no imports. The returned
 * keys are passed verbatim to `uploadFile` / `downloadFile` in
 * `backend/src/lib/storage.ts` and are persisted in the
 * `document_comparisons.redline_storage_path` / `diff_storage_path` columns.
 */

// ---------------------------------------------------------------------------
// Comparison storage key helpers
// ---------------------------------------------------------------------------

/**
 * Base R2 prefix for a single comparison's artifacts.
 * @returns `comparisons/{userId}/{comparisonId}`
 */
export function comparisonPrefix(userId: string, comparisonId: string): string {
  return `comparisons/${userId}/${comparisonId}`;
}

/**
 * R2 key for the redline `.docx` artifact of a comparison.
 * @returns `comparisons/{userId}/{comparisonId}/redline.docx`
 */
export function comparisonRedlineKey(
  userId: string,
  comparisonId: string,
): string {
  return `${comparisonPrefix(userId, comparisonId)}/redline.docx`;
}

/**
 * R2 key for the diff `.json` artifact of a comparison.
 * @returns `comparisons/{userId}/{comparisonId}/diff.json`
 */
export function comparisonDiffKey(
  userId: string,
  comparisonId: string,
): string {
  return `${comparisonPrefix(userId, comparisonId)}/diff.json`;
}
