/**
 * Word-level text diff for the Document Compare / Redline engine.
 *
 * {@link wordDiff} computes the differences between the plaintext of a matched
 * base/revised paragraph pair and returns them as an ordered list of typed
 * {@link DiffSegment}s (`ins` / `del` / `equal`). It is built on `fast-diff`,
 * the dependency-free Myers-diff library already used elsewhere in the backend
 * (`../lib/docxTrackedChanges`), so it introduces no new package.
 *
 * Granularity is word-level: the two plaintexts are split into word/whitespace
 * tokens which are diffed as atomic units, so an edited word reads as a single
 * `del` + `ins` pair rather than a run of mid-word character edits. If a
 * paragraph carries more distinct tokens than the token-encoding space can
 * address, the character-level `fast-diff` output is used as a deterministic
 * fallback.
 *
 * The function is pure and deterministic: it uses no clock, randomness, or
 * cursor hint, so identical inputs always yield identical segment arrays. It
 * only classifies characters into segments — it never adds, drops, or mutates
 * a character — which preserves the reconstruction invariant relied upon by
 * {@link buildDiffJson}:
 *
 *   concat(text of segments whose type is `equal` or `del`) === baseText
 *   concat(text of segments whose type is `equal` or `ins`) === revisedText
 *
 * @module compare/wordDiff
 */

// Default import of fast-diff (the package uses `export = diff`); this resolves
// under `esModuleInterop: true`, exactly as `../lib/docxTrackedChanges` imports it.
import fastDiff from "fast-diff";

// Reuse the segment contract owned by ./diffJson so the engine has a single
// source of truth for the segment shape and its type union.
import type { DiffSegment, DiffHunkType } from "./diffJson";

/**
 * Maps a raw `fast-diff` operation code to a {@link DiffHunkType}.
 *
 * Uses the library's named constants (`fastDiff.DELETE === -1`,
 * `fastDiff.INSERT === 1`, `fastDiff.EQUAL === 0`) rather than magic numbers.
 * The `never` default is an exhaustiveness guard: if `fast-diff`'s operation
 * set ever changes, this fails to compile under `strict` mode.
 */
function opToType(op: -1 | 0 | 1): DiffHunkType {
  switch (op) {
    case fastDiff.DELETE:
      return "del";
    case fastDiff.INSERT:
      return "ins";
    case fastDiff.EQUAL:
      return "equal";
    default: {
      const exhaustive: never = op;
      throw new Error(`Unsupported fast-diff operation: ${String(exhaustive)}`);
    }
  }
}

/**
 * Appends `text` to `segments`, coalescing it into the trailing segment when
 * that segment shares the same `type`. Zero-length text is ignored.
 *
 * `fast-diff` already merges adjacent same-op runs internally, so this is a
 * deterministic safety net that also keeps the token- and character-level code
 * paths uniform.
 */
function pushSegment(
  segments: DiffSegment[],
  type: DiffHunkType,
  text: string,
): void {
  if (text.length === 0) {
    return;
  }
  const lastIndex = segments.length - 1;
  if (lastIndex >= 0 && segments[lastIndex].type === type) {
    segments[lastIndex].text += text;
    return;
  }
  segments.push({ type, text });
}

/**
 * Splits a paragraph plaintext into atomic word/whitespace tokens.
 *
 * The capturing split on whitespace runs keeps the separators as their own
 * tokens, and empty fragments (produced at the string boundaries) are removed.
 * Concatenating the returned tokens in order reproduces `text` exactly, which
 * is what keeps the reconstruction invariant intact after the tokens are
 * diffed and decoded.
 */
function tokenize(text: string): string[] {
  if (text.length === 0) {
    return [];
  }
  return text.split(/(\s+)/).filter((token) => token.length > 0);
}

/**
 * Maximum number of distinct tokens the encoding can represent. Each token is
 * encoded as a single UTF-16 code unit, so the addressable space is the Basic
 * Multilingual Plane (`0x0000`–`0xFFFF`). Index `0` is reserved as an unused
 * sentinel so no token is ever encoded as `NUL`, leaving `0xFFFF` usable slots.
 */
const MAX_DISTINCT_TOKENS = 0xffff;

/** The result of encoding two token streams for character-based diffing. */
interface TokenEncoding {
  /** Base token stream encoded as one code unit per token. */
  encodedBase: string;
  /** Revised token stream encoded as one code unit per token. */
  encodedRevised: string;
  /** Decode table: `tokens[codeUnit]` is the original token string. */
  tokens: string[];
}

/**
 * Encodes two token streams into single-code-unit strings so `fast-diff` can
 * diff them as atomic units. Identical token strings map to the same code unit
 * (first-occurrence order, scanning base then revised), so equal words compare
 * equal during the diff.
 *
 * Returns `null` when the number of distinct tokens exceeds
 * {@link MAX_DISTINCT_TOKENS}, signalling the caller to fall back to a
 * character-level diff. (This is unreachable for realistic paragraph text.)
 */
function encodeTokens(
  baseTokens: string[],
  revisedTokens: string[],
): TokenEncoding | null {
  // Index 0 is a reserved, never-referenced sentinel (avoids encoding NUL).
  const tokens: string[] = [""];
  const tokenToIndex = new Map<string, number>();

  const encode = (list: string[]): string | null => {
    let encoded = "";
    for (const token of list) {
      let index = tokenToIndex.get(token);
      if (index === undefined) {
        if (tokens.length > MAX_DISTINCT_TOKENS) {
          // No addressable code unit remains for a new token.
          return null;
        }
        index = tokens.length;
        tokens.push(token);
        tokenToIndex.set(token, index);
      }
      encoded += String.fromCharCode(index);
    }
    return encoded;
  };

  const encodedBase = encode(baseTokens);
  if (encodedBase === null) {
    return null;
  }
  const encodedRevised = encode(revisedTokens);
  if (encodedRevised === null) {
    return null;
  }
  return { encodedBase, encodedRevised, tokens };
}

/**
 * Character-level diff fallback: maps `fast-diff`'s raw character output
 * directly to segments. Used only when {@link encodeTokens} cannot address the
 * full token set. Preserves the reconstruction invariant (characters are only
 * classified, never altered).
 */
function charLevelDiff(baseText: string, revisedText: string): DiffSegment[] {
  const raw = fastDiff(baseText, revisedText);
  const segments: DiffSegment[] = [];
  for (const [op, text] of raw) {
    pushSegment(segments, opToType(op), text);
  }
  return segments;
}

/**
 * Computes the word-level diff between two paragraph plaintexts and returns the
 * ordered typed segments.
 *
 * Contract (as consumed by the compare orchestrator, `index.ts`):
 * - Identical inputs yield a single `equal` segment carrying the whole text
 *   (or an empty array when both inputs are empty).
 * - A modified pair yields interleaved `del` / `ins` / `equal` segments where
 *   each change is expressed at word granularity.
 * - Whole-paragraph insertions/deletions for unmatched paragraphs, and the
 *   `equal "\n"` separators between paragraphs, are emitted by the orchestrator
 *   — not here. `wordDiff` operates on a single paragraph's text and expects no
 *   embedded `"\n"`.
 *
 * @param baseText The base paragraph's plaintext.
 * @param revisedText The revised paragraph's plaintext.
 * @returns Ordered {@link DiffSegment}s satisfying the reconstruction invariant.
 */
export function wordDiff(baseText: string, revisedText: string): DiffSegment[] {
  // Fast path: identical paragraph text needs no diff.
  if (baseText === revisedText) {
    return baseText.length === 0 ? [] : [{ type: "equal", text: baseText }];
  }

  const encoded = encodeTokens(tokenize(baseText), tokenize(revisedText));
  if (encoded === null) {
    return charLevelDiff(baseText, revisedText);
  }

  const raw = fastDiff(encoded.encodedBase, encoded.encodedRevised);
  const segments: DiffSegment[] = [];
  for (const [op, chunk] of raw) {
    if (chunk.length === 0) {
      continue;
    }
    const type = opToType(op);
    // Decode each code unit back to its original token and concatenate.
    let text = "";
    for (let i = 0; i < chunk.length; i += 1) {
      text += encoded.tokens[chunk.charCodeAt(i)];
    }
    pushSegment(segments, type, text);
  }
  return segments;
}
