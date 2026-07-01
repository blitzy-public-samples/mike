/**
 * Post-emission validation for the Document Compare / Redline engine.
 *
 * After {@link module:compare/emitTrackedChanges emitTrackedChanges} rezips the
 * merged `word/document.xml` back into a `.docx`, this module sanity-checks the
 * emitted bytes BEFORE they are persisted to R2 or returned to the caller
 * (AAP §0.1.1 step 7, §0.6.2). It performs two independent checks:
 *
 *   1. Reparse (FATAL). The emitted bytes are reloaded as a zip and their
 *      `word/document.xml` is reparsed into a well-formed OOXML tree containing
 *      `<w:document><w:body>`. A genuine failure here means we produced an
 *      invalid `.docx`, so the error is allowed to propagate and the comparison
 *      fails loudly.
 *
 *   2. LibreOffice render smoke test (SOFT / best-effort). The emitted bytes are
 *      optionally handed to {@link docxToPdf} as a render probe. `docxToPdf`
 *      throws when the LibreOffice/soffice binary is unavailable (and on other
 *      conversion errors); because this container/CI may legitimately lack that
 *      binary, ANY error from the render probe is swallowed and treated as a
 *      *skipped* check. The render probe NEVER fails validation.
 *
 * Isolation: this module imports only `jszip`, the shared read-side OOXML
 * helpers from `./parseDocx` (which own the canonical preserve-order parser
 * config), and `docxToPdf` from `../lib/convert`. It reaches into no route,
 * persistence, assistant/chat, tabular-review, or workflow code path, and it
 * never mutates its input bytes.
 *
 * @module compare/validate
 */

import JSZip from "jszip";
// Reuse convert.docxToPdf so validation goes through the SAME LibreOffice path
// as the rest of the backend (no second runtime, no direct soffice spawn).
import { docxToPdf } from "../lib/convert";
// Reuse the engine's canonical preserve-order parser + body locator so the
// reparse check can never diverge from how the document was originally parsed.
import { createParser, findBody, type XNode } from "./parseDocx";

/** Canonical OOXML main-document part path (forward-slash form). */
const DOCUMENT_XML_PATH = "word/document.xml";

/**
 * Outcome of {@link validate}.
 *
 * `validate` only ever RESOLVES with this object when the fatal reparse check
 * has already succeeded; a reparse failure rejects the promise instead of
 * returning. The render field is advisory and environment-dependent.
 */
export interface ValidateResult {
  /**
   * Always `true` when {@link validate} resolves: the emitted redline was
   * successfully reloaded as a zip and its `word/document.xml` reparsed into a
   * tree containing `<w:document><w:body>`.
   */
  reparsed: true;
  /**
   * `true` only when the optional LibreOffice render probe ran AND produced a
   * non-empty PDF. `false` when the probe was skipped — e.g. the
   * LibreOffice/soffice binary is unavailable, or the conversion errored. The
   * render probe is best-effort and never fails validation, so callers and
   * tests must treat `false` as "not checked", NOT as "invalid".
   */
  renderChecked: boolean;
}

/**
 * Resolve the main-document part (`word/document.xml`) from an opened archive.
 *
 * Some Windows/Word archives store entries with backslash separators
 * (e.g. `word\document.xml`) even though the zip spec mandates forward slashes,
 * and JSZip matches entries by exact string. This mirrors the read-side
 * fallback used by `parseDocx` / `convert` so validation stays robust for the
 * same inputs the rest of the engine accepts.
 *
 * @param zip - the archive loaded from the emitted redline bytes
 * @returns the `word/document.xml` entry, or `null` if it is absent entirely
 */
function readDocumentXmlEntry(zip: JSZip): JSZip.JSZipObject | null {
  const direct = zip.file(DOCUMENT_XML_PATH);
  if (direct) return direct;
  return zip.file(DOCUMENT_XML_PATH.replace(/\//g, "\\"));
}

/**
 * Validate the emitted redline `.docx` before it is persisted or returned.
 *
 * The reparse check is the single fatal gate: if the emitted bytes cannot be
 * reopened as a zip, are missing `word/document.xml`, cannot be parsed, or lack
 * a `<w:document><w:body>`, this function throws (rejects). The LibreOffice
 * render probe is best-effort and is skipped — not failed — whenever the
 * soffice binary is unavailable or the conversion errors.
 *
 * The function performs no mutation of its input and is deterministic in its
 * fatal outcome (reparse either works or it does not). Only `renderChecked` is
 * environment-dependent; it flips based on binary availability and never
 * affects the emitted bytes.
 *
 * @param redlineBytes - the emitted redline `.docx` bytes to validate
 * @returns a {@link ValidateResult} once the fatal reparse check has passed
 * @throws Error if the redline cannot be reparsed into a valid OOXML document
 */
export async function validate(redlineBytes: Buffer): Promise<ValidateResult> {
  // --- Check 1: reparse (FATAL) --------------------------------------------
  // Reload the emitted bytes as a zip. An invalid archive is fatal, so let the
  // error propagate rather than swallowing it.
  const zip = await JSZip.loadAsync(redlineBytes);

  const entry = readDocumentXmlEntry(zip);
  if (!entry) {
    throw new Error("validate: emitted redline missing word/document.xml");
  }

  const xml = await entry.async("string");

  // Parse with the engine's shared preserve-order parser config (reused from
  // parseDocx so it can never diverge). A parse failure is fatal and propagates.
  const tree = createParser().parse(xml) as XNode[];

  // Defensive structural assertion: a valid main document must contain
  // <w:document><w:body>. A missing body signals a malformed emission.
  const body = findBody(tree);
  if (!body) {
    throw new Error(
      "validate: emitted redline has no <w:document><w:body> element",
    );
  }

  // --- Check 2: LibreOffice render smoke test (SOFT / best-effort) ----------
  // `docxToPdf` throws when the LibreOffice/soffice binary is unavailable (and
  // on other conversion errors). The render probe must never fail the
  // comparison, so any error is treated as a skip; only the reparse check above
  // is a fatal gate.
  let renderChecked = false;
  try {
    const pdf = await docxToPdf(redlineBytes);
    renderChecked = pdf.length > 0;
  } catch {
    renderChecked = false;
  }

  return { reparsed: true, renderChecked };
}
