/**
 * Word Document Operations for Armenian Name Indexer
 * Handles interaction with Word document via Office.js API
 */

/* global Word */

import { IndexerSettings, IndexResult, NameMatch, ProgressCallback, NormalizationMode, NormalizationRule } from "../types";
import { parseArmenianName, normalizeArmenianSurname, parseExceptionsList, isExcluded } from "./armenian";

const clampPercent = (value: number): number => Math.max(0, Math.min(100, Math.floor(value)));

const removeFirstMatchingSuffix = (value: string, suffixes: string[]): string => {
  const suffix = suffixes.find((s) => s && value.endsWith(s));
  return suffix ? value.substring(0, value.length - suffix.length) : value;
};

const applyCustomNormalizationRules = (value: string, rules: NormalizationRule[]): string =>
  rules.reduce((current, rule) => {
    const flags = rule.flags && rule.flags.length > 0 ? rule.flags : "g";
    const regex = new RegExp(rule.pattern, flags);
    return current.replace(regex, rule.replacement);
  }, value);

const normalizeSurname = (surname: string, settings: IndexerSettings): string => {
  const normalization = settings.normalization;
  if (!normalization?.enabled) return surname;

  const strategies: Record<NormalizationMode, (value: string) => string> = {
    none: (value) => value,
    suffix: (value) => removeFirstMatchingSuffix(value, settings.suffixes),
    armenian: (value) => normalizeArmenianSurname(value, settings.suffixes),
    custom: (value) => applyCustomNormalizationRules(value, normalization.customRules),
  };

  return (strategies[normalization.mode] ?? ((value: string) => value))(surname);
};

function getNormalizedIndexEntryOrNull(
  fullName: string,
  settings: IndexerSettings,
  exceptions: Set<string>
): string | null {
  if (isExcluded(fullName, exceptions)) {
    return null;
  }

  const parsed = parseArmenianName(fullName);
  if (!parsed.firstName || !parsed.lastName) {
    return null;
  }

  const normalizedSurname = normalizeSurname(parsed.lastName, settings);
  return `${parsed.firstName} ${normalizedSurname}`;
}

async function indexMatches(
  context: Word.RequestContext,
  matches: NameMatch[],
  settings: IndexerSettings,
  exceptions: Set<string>,
  result: IndexResult,
  onProgress?: ProgressCallback,
  cancelToken?: { cancelled: boolean }
): Promise<void> {
  // Insert-field operations queue up on the JS side until a sync.
  // Sync too often and you pay round-trip latency; sync too rarely and
  // the pending command queue can grow very large on big documents.
  const syncEvery = 25;

  for (let i = matches.length - 1; i >= 0; i--) {
    if (cancelToken?.cancelled) {
      if (onProgress) onProgress(100, `Cancelled — ${result.indexed} indexed so far`);
      return;
    }

    const match = matches[i];
    const fullName = match.text;

    try {
      const normalizedName = getNormalizedIndexEntryOrNull(fullName, settings, exceptions);
      if (!normalizedName) {
        result.skipped++;
        continue;
      }

      markIndexEntry(match.range, normalizedName);
      result.indexed++;

      if (result.indexed % syncEvery === 0) {
        // eslint-disable-next-line office-addins/no-context-sync-in-loop
        await context.sync();

        if (onProgress) {
          const processed = matches.length - i;
          const percent = clampPercent(10 + (processed / matches.length) * 85);
          onProgress(percent, `Indexing… ${result.indexed} indexed, ${result.skipped} skipped`);
        }
      }
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      result.errors.push(`Error processing "${fullName}": ${msg}`);
      result.skipped++;
    }
  }
}

/**
 * Find all Armenian names in the document matching the pattern.
 *
 * Searches paragraph-by-paragraph (Office.js has no native regex search).
 * Paragraph texts are loaded in a single batch to minimize sync round-trips.
 * When the same name text appears multiple times in a paragraph, each
 * occurrence is matched to the correct search-result range by occurrence index.
 *
 * @param context - Word request context
 * @param pattern - Regex pattern to match (must have the `g` flag)
 * @param cancelToken - Optional object; set `cancelled = true` to abort mid-scan
 * @returns Promise resolving to array of name matches
 */
export async function findArmenianNamesInDocument(
  context: Word.RequestContext,
  pattern: RegExp,
  cancelToken?: { cancelled: boolean },
  onProgress?: ProgressCallback,
  progressRange: { start: number; end: number } = { start: 0, end: 10 }
): Promise<NameMatch[]> {
  const matches: NameMatch[] = [];

  // In Office.js, true parallelism against the Word object model is not supported.
  // Performance comes from batching: queue many operations, then sync.
  // This limit caps how many paragraph.search() calls we queue before syncing.
  const maxQueuedSearchesPerSync = 60;
  const paragraphChunkSize = 60;

  // --- Load all paragraph texts in a single sync ---
  const paragraphs = context.document.body.paragraphs;
  paragraphs.load("items/text");
  await context.sync();

  let globalOffset = 0;

  const totalParagraphs = paragraphs.items.length;
  const start = Math.max(0, Math.min(100, progressRange.start));
  const end = Math.max(0, Math.min(100, progressRange.end));
  const span = Math.max(0, end - start);

  type ParagraphOccurrence = {
    text: string;
    index: number;
    length: number;
    occurrenceIndex: number;
  };

  type ScanResult = {
    occurrences: ParagraphOccurrence[];
    uniqueTexts: string[];
  };

  type ScanWorkerResponse =
    | {
        type: "scanResult";
        requestId: string;
        results: Array<{ id: number; occurrences: ParagraphOccurrence[]; uniqueTexts: string[] }>;
      }
    | { type: "error"; requestId: string; message: string };

  const canUseWorker = typeof Worker !== "undefined";
  const worker = canUseWorker ? new Worker(new URL("./nameScan.worker.ts", import.meta.url)) : null;
  let nextRequestId = 0;

  const terminateWorker = () => {
    try {
      worker?.terminate();
    } catch {
      // ignore
    }
  };

  const scanChunkInWorker = (paragraphsToScan: Array<{ id: number; text: string }>): Promise<Map<number, ScanResult>> => {
    if (!worker) {
      return Promise.reject(new Error("Worker not available"));
    }

    const requestId = `${Date.now()}_${nextRequestId++}`;

    return new Promise((resolve, reject) => {
      const handler = (ev: MessageEvent) => {
        const data = ev.data as ScanWorkerResponse;
        if (!data || data.requestId !== requestId) return;

        worker.removeEventListener("message", handler);

        if (data.type === "error") {
          reject(new Error(data.message));
          return;
        }

        const map = new Map<number, ScanResult>();
        for (const r of data.results) {
          map.set(r.id, { occurrences: r.occurrences, uniqueTexts: r.uniqueTexts });
        }

        resolve(map);
      };

      worker.addEventListener("message", handler);

      worker.postMessage({
        type: "scan",
        requestId,
        patternSource: pattern.source,
        patternFlags: pattern.flags,
        paragraphs: paragraphsToScan,
      });
    });
  };

  const awaitWithCancellation = async <T>(promise: Promise<T>): Promise<T> => {
    if (!cancelToken) return promise;

    // Poll cancellation while we wait for the worker.
    // (We could also use a MessageChannel-based cancel signal, but polling keeps this minimal.)
    while (true) {
      if (cancelToken.cancelled) {
        terminateWorker();
        throw new Error("Cancelled");
      }

      // Wait a short interval or until the promise resolves.
      const outcome = await Promise.race([
        promise.then((value) => ({ type: "resolved" as const, value })),
        new Promise<{ type: "tick" }>((resolve) => setTimeout(() => resolve({ type: "tick" }), 50)),
      ]);

      if (outcome.type === "resolved") {
        return outcome.value;
      }
    }
  };

  const scanChunkOnMainThread = (paragraphsToScan: Array<{ id: number; text: string }>): Map<number, ScanResult> => {
    const map = new Map<number, ScanResult>();

    for (const { id, text } of paragraphsToScan) {
      // Reset regex state for each paragraph
      pattern.lastIndex = 0;

      const occurrencesSeen = new Map<string, number>();
      const occurrences: ParagraphOccurrence[] = [];

      let m: RegExpExecArray | null;
      while ((m = pattern.exec(text)) !== null) {
        const matchText = m[0];
        const seenCount = occurrencesSeen.get(matchText) ?? 0;
        occurrencesSeen.set(matchText, seenCount + 1);

        occurrences.push({
          text: matchText,
          index: m.index,
          length: matchText.length,
          occurrenceIndex: seenCount,
        });

        // Guard against zero-length matches causing infinite loops
        if (m[0].length === 0) {
          pattern.lastIndex++;
        }
      }

      map.set(id, { occurrences, uniqueTexts: Array.from(occurrencesSeen.keys()) });
    }

    return map;
  };

  type PendingParagraph = {
    globalOffset: number;
    occurrences: ParagraphOccurrence[];
    searchResultsByText: Record<string, Word.RangeCollection>;
  };

  let pending: PendingParagraph[] = [];
  let queuedSearchCount = 0;

  const flushPending = async () => {
    if (pending.length === 0) return;

    // Resolve all queued paragraph.search() calls at once.
    // eslint-disable-next-line office-addins/no-context-sync-in-loop
    await context.sync();

    for (const item of pending) {
      for (const occ of item.occurrences) {
        const results = item.searchResultsByText[occ.text];
        const range = results?.items?.[occ.occurrenceIndex];
        if (!range) continue;

        matches.push({
          text: occ.text,
          range,
          startIndex: item.globalOffset + occ.index,
          length: occ.length,
        });
      }
    }

    pending = [];
    queuedSearchCount = 0;
  };

  try {
    for (let chunkStart = 0; chunkStart < totalParagraphs; chunkStart += paragraphChunkSize) {
      if (cancelToken?.cancelled) break;

      const chunkEndExclusive = Math.min(totalParagraphs, chunkStart + paragraphChunkSize);

      if (onProgress && totalParagraphs > 0) {
        const percent = Math.floor(start + (chunkStart / totalParagraphs) * span);
        onProgress(percent, `Searching… ${chunkStart + 1}/${totalParagraphs}`);
      }

      // Collect texts for this chunk.
      // Keep offsets on the main thread; worker only does CPU scanning.
      const chunkTexts: Array<{ id: number; text: string; length: number }> = [];
      for (let i = chunkStart; i < chunkEndExclusive; i++) {
        const t = paragraphs.items[i].text;
        chunkTexts.push({ id: i, text: t, length: t.length });
      }

      // Scan (worker preferred, fallback to main thread).
      let scanMap: Map<number, ScanResult>;
      try {
        if (worker) {
          scanMap = await awaitWithCancellation(
            scanChunkInWorker(chunkTexts.map((p) => ({ id: p.id, text: p.text })))
          );
        } else {
          scanMap = scanChunkOnMainThread(chunkTexts.map((p) => ({ id: p.id, text: p.text })));
        }
      } catch {
        // If worker fails (blocked/older host), fall back to main-thread scanning.
        scanMap = scanChunkOnMainThread(chunkTexts.map((p) => ({ id: p.id, text: p.text })));
      }

      for (let i = chunkStart; i < chunkEndExclusive; i++) {
        if (cancelToken?.cancelled) break;

        const paragraph = paragraphs.items[i];
        const textLen = chunkTexts[i - chunkStart].length;
        const scan = scanMap.get(i);

        if (!scan || scan.occurrences.length === 0) {
          globalOffset += textLen;
          continue;
        }

        const searchResultsByText: Record<string, Word.RangeCollection> = Object.create(null) as Record<
          string,
          Word.RangeCollection
        >;

        for (const uniqueText of scan.uniqueTexts) {
          const searchResults = paragraph.search(uniqueText, {
            matchCase: true,
            matchWholeWord: false,
          });
          searchResults.load("items");
          searchResultsByText[uniqueText] = searchResults;
          queuedSearchCount++;
        }

        pending.push({
          globalOffset,
          occurrences: scan.occurrences,
          searchResultsByText,
        });

        if (queuedSearchCount >= maxQueuedSearchesPerSync) {
          await flushPending();
        }

        globalOffset += textLen;
      }

      if (cancelToken?.cancelled) {
        break;
      }
    }
  } finally {
    // Always clean up the worker (important for long-lived task panes).
    terminateWorker();
  }

  // Resolve any remaining queued searches.
  if (!cancelToken?.cancelled) {
    await flushPending();
  }

  if (onProgress) {
    onProgress(end, `Searching… ${totalParagraphs}/${totalParagraphs}`);
  }

  return matches;
}

/**
 * Mark an index entry at the specified range.
 * Inserts an XE (index entry) field immediately before the matched text.
 *
 * @param range - Range where to insert the index entry
 * @param entry - Index entry text (the name as it should appear in the index)
 */
export function markIndexEntry(range: Word.Range, entry: string): void {
  // Word.FieldType.xe = "XE" — the correct enum value for index-entry fields.
  // The `text` argument provides the field data: the quoted entry string.
  range.insertField(Word.InsertLocation.before, Word.FieldType.xe, `"${entry}"`, false);
}

/**
 * Clear all XE index entries from the document body.
 * Ported from VBA ClearAllIndexEntries.
 *
 * Field codes are loaded in a single batch before the deletion loop to
 * avoid one sync per field.
 *
 * @param context - Word request context
 * @param onProgress - Optional progress callback
 * @returns Promise resolving to number of entries deleted
 */
export async function clearAllIndexEntries(
  context: Word.RequestContext,
  onProgress?: ProgressCallback
): Promise<number> {
  const fields = context.document.body.fields;
  fields.load("items");
  await context.sync();

  if (fields.items.length === 0) {
    return 0;
  }

  // --- Batch: load all field codes in a single sync ---
  fields.items.forEach((f) => f.load("code"));
  await context.sync();

  const total = fields.items.length;
  let deleted = 0;

  // Iterate backwards so pending deletions don't shift remaining indices
  for (let i = total - 1; i >= 0; i--) {
    const field = fields.items[i];
    const fieldCode = (field.code ?? "").trim().toUpperCase();

    if (fieldCode.startsWith("XE")) {
      field.delete();
      deleted++;

      // Sync every 10 deletions to keep the operation queue manageable
      if (deleted % 10 === 0) {
        // eslint-disable-next-line office-addins/no-context-sync-in-loop
        await context.sync();

        if (onProgress) {
          const percent = Math.floor(((total - i) / total) * 100);
          onProgress(percent, `Clearing index entries… ${deleted} removed`);
        }
      }
    }
  }

  await context.sync();

  if (onProgress) {
    onProgress(100, `Cleared ${deleted} index entries`);
  }

  return deleted;
}

/**
 * Index all Armenian names in the document.
 * Main orchestration function — ported from VBA AutoIndexArmenianNames.
 *
 * @param context - Word request context
 * @param settings - Indexer settings
 * @param onProgress - Optional progress callback
 * @param cancelToken - Optional object; set `cancelled = true` to abort
 * @returns Promise resolving to index result
 */
export async function indexArmenianNames(
  context: Word.RequestContext,
  settings: IndexerSettings,
  onProgress?: ProgressCallback,
  cancelToken?: { cancelled: boolean }
): Promise<IndexResult> {
  const result: IndexResult = {
    indexed: 0,
    skipped: 0,
    errors: [],
  };

  try {
    const exceptions = parseExceptionsList(settings.exceptions.join("\n"));

    const pattern = new RegExp(settings.pattern, "g");
    const matches = await findArmenianNamesInDocument(context, pattern, cancelToken, onProgress, { start: 0, end: 10 });

    if (cancelToken?.cancelled) {
      if (onProgress) onProgress(100, "Cancelled");
      return result;
    }

    if (matches.length === 0) {
      if (onProgress) onProgress(100, "No names found");
      return result;
    }

    if (onProgress) {
      onProgress(10, `Found ${matches.length} potential names`);
    }

    // Process in reverse order (like VBA) to avoid field-insertion position drift
    await indexMatches(context, matches, settings, exceptions, result, onProgress, cancelToken);

    await context.sync();

    if (onProgress && !cancelToken?.cancelled) {
      onProgress(100, `Complete: ${result.indexed} indexed, ${result.skipped} skipped`);
    }
  } catch (error) {
    const msg = error instanceof Error ? error.message : String(error);
    result.errors.push(`Indexing error: ${msg}`);
    throw error;
  }

  return result;
}

/**
 * Preview which names would be indexed without writing any XE fields.
 * Returns the sorted, deduplicated list of normalized index strings.
 *
 * @param context - Word request context
 * @param settings - Indexer settings
 * @param onProgress - Optional progress callback
 * @returns Promise resolving to sorted array of index entry strings
 */
export async function previewArmenianNames(
  context: Word.RequestContext,
  settings: IndexerSettings,
  onProgress?: ProgressCallback
): Promise<string[]> {
  const exceptions = parseExceptionsList(settings.exceptions.join("\n"));
  const pattern = new RegExp(settings.pattern, "g");
  const matches = await findArmenianNamesInDocument(context, pattern, undefined, onProgress, {
    start: 0,
    end: 85,
  });

  const entries: string[] = [];

  for (const match of matches) {
    if (isExcluded(match.text, exceptions)) continue;

    const parsed = parseArmenianName(match.text);
    if (!parsed.firstName || !parsed.lastName) continue;

    const normalizedSurname = normalizeSurname(parsed.lastName, settings);
    entries.push(`${parsed.firstName} ${normalizedSurname}`);
  }

  if (onProgress) onProgress(100, `Preview complete: ${entries.length} names found`);

  return Array.from(new Set(entries)).sort();
}
