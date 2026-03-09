/**
 * TypeScript interfaces and types for Name Indexer
 */

/* global Word */

export type NormalizationMode = "none" | "suffix" | "armenian" | "custom";

export interface NormalizationRule {
  /** Regex source pattern (without surrounding / /) */
  pattern: string;

  /** Regex flags, e.g. "g", "gi". If omitted, defaults to "g". */
  flags?: string;

  /** Replacement string (JavaScript-style regex replacement) */
  replacement: string;
}

/**
 * Settings for the name indexer
 */
export interface IndexerSettings {
  /** List of words to exclude from indexing (e.g., place names, common words) */
  exceptions: string[];

  /** Regex pattern to match names (as string for serialization) */
  pattern: string;

  /** Suffixes to remove from surnames (Unicode escaped strings supported) */
  suffixes: string[];

  /** Minimum and maximum number of capitalized words to match */
  wordCount: {
    min: number;
    max: number;
  };

  /** Optional normalization settings applied to the surname when creating index entries */
  normalization: {
    /** Master toggle to enable/disable normalization */
    enabled: boolean;

    /** Which normalization strategy to apply */
    mode: NormalizationMode;

    /** Custom replacement rules (applied in order) when mode="custom" */
    customRules: NormalizationRule[];
  };
}

/**
 * Result of name parsing
 */
export interface ParsedName {
  /** First name */
  firstName: string;

  /** Last name (surname) */
  lastName: string;

  /** Full original name */
  fullName: string;
}

/**
 * A matched name in the document
 */
export interface NameMatch {
  /** Full matched text */
  text: string;

  /** Word Range object */
  range: Word.Range;

  /** Start position in the document text */
  startIndex: number;

  /** Length of the match */
  length: number;
}

/**
 * Result of indexing operation
 */
export interface IndexResult {
  /** Number of names successfully indexed */
  indexed: number;

  /** Number of names skipped (due to exceptions or errors) */
  skipped: number;

  /** Any errors encountered */
  errors: string[];
}

/**
 * Progress callback function type
 */
export type ProgressCallback = (percent: number, status: string) => void;
