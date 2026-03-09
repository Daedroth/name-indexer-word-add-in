/* global self */

// CPU-only scanning worker: finds regex matches inside paragraph text.
// NOTE: Workers cannot access Office.js; they only process plain data.

export type ScanRequest = {
  type: "scan";
  requestId: string;
  patternSource: string;
  patternFlags: string;
  paragraphs: Array<{ id: number; text: string }>;
};

export type ParagraphOccurrence = {
  text: string;
  index: number;
  length: number;
  occurrenceIndex: number;
};

export type ScanResponse =
  | {
      type: "scanResult";
      requestId: string;
      results: Array<{ id: number; occurrences: ParagraphOccurrence[]; uniqueTexts: string[] }>;
    }
  | {
      type: "error";
      requestId: string;
      message: string;
    };

const ctx: DedicatedWorkerGlobalScope = self as unknown as DedicatedWorkerGlobalScope;

function ensureGlobalFlag(flags: string): string {
  return flags.includes("g") ? flags : flags + "g";
}

ctx.onmessage = (ev: MessageEvent) => {
  const data = ev.data as ScanRequest;

  if (!data || data.type !== "scan" || !data.requestId) {
    return;
  }

  try {
    const flags = ensureGlobalFlag(data.patternFlags ?? "g");
    const pattern = new RegExp(data.patternSource, flags);

    const results = data.paragraphs.map(({ id, text }) => {
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

      return {
        id,
        occurrences,
        uniqueTexts: Array.from(occurrencesSeen.keys()),
      };
    });

    const response: ScanResponse = {
      type: "scanResult",
      requestId: data.requestId,
      results,
    };

    ctx.postMessage(response);
  } catch (error) {
    const response: ScanResponse = {
      type: "error",
      requestId: data.requestId,
      message: error instanceof Error ? error.message : String(error),
    };

    ctx.postMessage(response);
  }
};
