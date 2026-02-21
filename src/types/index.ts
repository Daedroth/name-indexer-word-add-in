/**
 * TypeScript interfaces and types for Armenian Name Indexer
 */

/* global Word */

/**
 * Settings for the Armenian name indexer
 */
export interface IndexerSettings {
  /** List of words to exclude from indexing (e.g., place names, common words) */
  exceptions: string[];
  
  /** Regex pattern to match Armenian names (as string for serialization) */
  pattern: string;
  
  /** List of Armenian suffixes to remove from surnames (Unicode escaped strings) */
  suffixes: string[];
  
  /** Minimum and maximum number of capitalized words to match */
  wordCount: {
    min: number;
    max: number;
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
 * A matched Armenian name in the document
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
