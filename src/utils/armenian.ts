/**
 * Armenian Name Processing Utilities
 * Ported from VBA FullScript.vb
 */

import { ParsedName } from "../types";

/**
 * Normalize Armenian surname by removing common suffixes
 * Ported from VBA NormalizeArmenianSurnameUnicode function
 *
 * @param surname - The surname to normalize
 * @param suffixes - Array of suffix strings to remove (Unicode escaped)
 * @returns Normalized surname
 */
export function normalizeArmenianSurname(surname: string, suffixes: string[]): string {
  let base = surname;

  // Armenian roots
  const yan = "\u0575\u0561\u0576"; // յան
  const yants = "\u0575\u0561\u0576\u0581"; // յանց
  const nChar = "\u0576"; // ն

  // Handle յանն → յան (remove definite article ն from յան)
  if (base.endsWith(yan + nChar)) {
    return base.substring(0, base.length - nChar.length);
  }

  // Handle յանցն → յանց (remove definite article ն from յանց)
  if (base.endsWith(yants + nChar)) {
    return base.substring(0, base.length - nChar.length);
  }

  // Process suffixes in order
  for (const suffix of suffixes) {
    // Special rule: do NOT remove "ն" if base ends with յան or յանց
    if (suffix === nChar) {
      if (base.endsWith(yan) || base.endsWith(yants)) {
        continue; // Skip this suffix
      }
    }

    if (base.endsWith(suffix)) {
      base = base.substring(0, base.length - suffix.length);
      break; // Only remove first matching suffix
    }
  }

  return base;
}

/**
 * Parse Armenian full name into first name and last name
 * Removes patronymic (middle name) if present
 *
 * @param fullName - Full name string
 * @returns Parsed name object with firstName and lastName
 */
export function parseArmenianName(fullName: string): ParsedName {
  const parts = fullName.trim().split(/\s+/);

  if (parts.length === 0) {
    return {
      firstName: "",
      lastName: "",
      fullName: fullName,
    };
  }

  const firstName = parts[0];
  const lastName = parts[parts.length - 1]; // Always take last word as surname

  return {
    firstName,
    lastName,
    fullName,
  };
}

/**
 * Create regex pattern to match Armenian names
 * Pattern matches sequences of capitalized Armenian words
 *
 * Armenian capital letters: \u0531–\u0556
 * Armenian lowercase letters: \u0561–\u0586
 *
 * @param wordCount - Min and max number of words to match
 * @returns RegExp pattern for matching Armenian names
 */
export function createArmenianNamePattern(wordCount: { min: number; max: number }): RegExp {
  // Pattern: Capital Armenian letter followed by lowercase Armenian letters
  // Repeated for min to max words separated by spaces
  const singleWord = "[\\u0531-\\u0556][\\u0561-\\u0586]+";
  const additionalWords = `(?:\\s+${singleWord})`;

  // Build pattern: first word + {min-1, max-1} additional words
  const minAdditional = Math.max(0, wordCount.min - 1);
  const maxAdditional = Math.max(0, wordCount.max - 1);

  const pattern = `${singleWord}${additionalWords}{${minAdditional},${maxAdditional}}`;

  return new RegExp(pattern, "g");
}

/**
 * Parse exceptions list from text
 * Handles comma-separated or newline-separated values
 *
 * Ported from VBA LoadExceptionsFromFile function
 *
 * @param text - Raw text containing exceptions
 * @returns Set of exception words (trimmed, non-empty)
 */
export function parseExceptionsList(text: string): Set<string> {
  const exceptions = new Set<string>();

  if (!text || text.trim().length === 0) {
    return exceptions;
  }

  // Normalize line endings then split by both newlines and commas simultaneously
  const normalized = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const items = normalized.split(/[\n,]+/);

  // Trim and add non-empty items
  items.forEach((item) => {
    const trimmed = item.trim();
    if (trimmed.length > 0) {
      exceptions.add(trimmed);
    }
  });

  return exceptions;
}

/**
 * Check if any word in the name is in the exceptions list
 *
 * @param fullName - Full name to check
 * @param exceptions - Set of exception words
 * @returns True if any word is in exceptions list
 */
export function isExcluded(fullName: string, exceptions: Set<string>): boolean {
  const words = fullName.trim().split(/\s+/).filter(Boolean);

  if (words.length === 0 || exceptions.size === 0) {
    return false;
  }

  // Exception words are treated as PREFIX patterns (equivalent to like("exception%"))
  // across each word-part of the name.
  for (const word of words) {
    for (const exception of exceptions) {
      const trimmed = exception.trim();
      if (!trimmed) continue;

      if (word.startsWith(trimmed)) {
        return true;
      }
    }
  }

  return false;
}

/**
 * Default suffix list for Armenian surnames
 * These are the suffixes from the VBA script
 */
export const DEFAULT_SUFFIXES = [
  "\u056B\u0581", //ից
  "\u056B\u0576", // ին
  "\u0578\u057E", // ով
  "\u0578\u0582\u0574", // ում
  "\u0568", // ը
  "\u056B", // ի
  "\u0576", // ն
];

/**
 * Default word count for matching names
 * Matches 2-3 capitalized words (min: 1 additional, max: 2 additional)
 */
export const DEFAULT_WORD_COUNT = {
  min: 2,
  max: 3,
};
