/**
 * Structured diff-JSON contract for the Document Compare / Redline engine.
 *
 * This module defines the canonical, hunk-based diff types shared across the
 * feature — the backend comparison engine, the comparisons route payload, the
 * R2-persisted `diff.json` artifact, and the frontend inline + side-by-side
 * redline renderers — together with {@link buildDiffJson}, a deterministic
 * builder that converts an ordered list of typed text segments into the final
 * hunk list, assigning the per-column character offsets as it goes.
 *
 * The module is dependency-free: it imports nothing and performs no side
 * effects, so identical inputs always produce byte-identical output under
 * `JSON.stringify`.
 *
 * @module compare/diffJson
 */

/**
 * Classifies a hunk (or an input segment) by which column(s) contain its text.
 *
 * - `"equal"` — text present in BOTH the base and revised columns.
 * - `"del"`   — text present only in the BASE column (a deletion).
 * - `"ins"`   — text present only in the REVISED column (an insertion).
 */
export type DiffHunkType = "ins" | "del" | "equal";

/**
 * A half-open character range `[start, end)` into a single column's plaintext.
 *
 * Offsets are UTF-16 code-unit positions (i.e. `String#length` units), so they
 * agree exactly with the lengths {@link buildDiffJson} accumulates.
 */
export interface DiffRange {
  /** Inclusive character offset at which the range begins. */
  start: number;
  /** Exclusive character offset at which the range ends. */
  end: number;
}

/**
 * A single ordered unit of the structured diff.
 *
 * Range nullability is fully determined by {@link DiffHunk.type}:
 * - `equal` → both {@link DiffHunk.baseRange} and {@link DiffHunk.revisedRange}
 *   are non-null (the text occupies both columns).
 * - `del`   → {@link DiffHunk.baseRange} is non-null and
 *   {@link DiffHunk.revisedRange} is `null` (base-only text).
 * - `ins`   → {@link DiffHunk.revisedRange} is non-null and
 *   {@link DiffHunk.baseRange} is `null` (revised-only text).
 *
 * Hunks are emitted in document order, so a consumer can reconstruct both
 * columns with a single walk of the array (base column = `equal` + `del` text;
 * revised column = `equal` + `ins` text).
 */
export interface DiffHunk {
  /** Which column(s) contain this hunk's text. */
  type: DiffHunkType;
  /** The verbatim text of this hunk. */
  text: string;
  /** Offsets into the BASE plaintext; `null` for pure insertions (`ins`). */
  baseRange: DiffRange | null;
  /** Offsets into the REVISED plaintext; `null` for pure deletions (`del`). */
  revisedRange: DiffRange | null;
}

/**
 * The complete structured diff: an ordered list of {@link DiffHunk}s.
 *
 * This is the payload returned by the comparisons route inside the comparison
 * object and the shape persisted to object storage as `diff.json`.
 */
export interface DiffJson {
  /** Hunks in document order. */
  hunks: DiffHunk[];
}

/**
 * An ordered, typed text segment consumed by {@link buildDiffJson}.
 *
 * Reuses {@link DiffHunkType} as the segment classifier so callers need not
 * import a separate segment type. Segments carry no offsets; the builder
 * assigns the per-column character offsets.
 */
export interface DiffSegment {
  /** Which column(s) the segment's text belongs to. */
  type: DiffHunkType;
  /** The verbatim text of the segment. */
  text: string;
}

/**
 * Builds the final {@link DiffJson} from an ordered list of typed text segments
 * in a single deterministic left-to-right pass.
 *
 * The builder maintains two running character cursors — one for the BASE
 * plaintext and one for the REVISED plaintext — both starting at `0`. These
 * index into the same concatenated plaintexts the engine diffs against
 * (paragraphs joined by `"\n"`). For each non-empty segment it emits exactly
 * one {@link DiffHunk} and advances the cursors:
 *
 * - `equal` → `baseRange` and `revisedRange` both span the segment; BOTH cursors
 *   advance by `text.length`.
 * - `del`   → `baseRange` spans the segment, `revisedRange` is `null`; ONLY the
 *   base cursor advances by `text.length`.
 * - `ins`   → `revisedRange` spans the segment, `baseRange` is `null`; ONLY the
 *   revised cursor advances by `text.length`.
 *
 * Behavioral guarantees:
 * - Zero-length segments (`text === ""`) are skipped, so no empty hunk is ever
 *   emitted and the cursors are unaffected by them.
 * - Each remaining segment yields exactly one hunk in input order; segments are
 *   not coalesced.
 * - Lengths and offsets use `String#length` (UTF-16 code units) throughout, so
 *   ranges and cursors always agree.
 * - The function is pure and deterministic: identical input always yields a
 *   deeply-equal result and an identical `JSON.stringify` output.
 *
 * @param segments Ordered typed text segments to convert into hunks.
 * @returns The structured diff with hunks in document order.
 *
 * @example
 * // [equal "Hello ", del "cruel ", ins "kind ", equal "world"] yields:
 * //   equal { baseRange: 0..6,   revisedRange: 0..6  }
 * //   del   { baseRange: 6..12,  revisedRange: null  }
 * //   ins   { baseRange: null,   revisedRange: 6..11 }
 * //   equal { baseRange: 12..17, revisedRange: 11..16 }
 */
export function buildDiffJson(segments: DiffSegment[]): DiffJson {
  const hunks: DiffHunk[] = [];

  // Running cursors into the BASE and REVISED plaintexts (UTF-16 code units).
  let baseOffset = 0;
  let revisedOffset = 0;

  for (const segment of segments) {
    const { type, text } = segment;

    // Skip zero-length segments so empty hunks never appear in the output and
    // the cursors remain unaffected.
    if (text.length === 0) {
      continue;
    }

    const length = text.length;

    switch (type) {
      case "equal": {
        const baseRange: DiffRange = { start: baseOffset, end: baseOffset + length };
        const revisedRange: DiffRange = {
          start: revisedOffset,
          end: revisedOffset + length,
        };
        hunks.push({ type: "equal", text, baseRange, revisedRange });
        baseOffset += length;
        revisedOffset += length;
        break;
      }
      case "del": {
        const baseRange: DiffRange = { start: baseOffset, end: baseOffset + length };
        hunks.push({ type: "del", text, baseRange, revisedRange: null });
        baseOffset += length;
        break;
      }
      case "ins": {
        const revisedRange: DiffRange = {
          start: revisedOffset,
          end: revisedOffset + length,
        };
        hunks.push({ type: "ins", text, baseRange: null, revisedRange });
        revisedOffset += length;
        break;
      }
      default: {
        // Exhaustiveness guard: unreachable for valid DiffHunkType values. If a
        // new member is added to DiffHunkType, this assignment fails to compile
        // under `strict` mode, forcing a matching case to be handled here.
        const exhaustive: never = type;
        throw new Error(`Unsupported diff segment type: ${String(exhaustive)}`);
      }
    }
  }

  return { hunks };
}
