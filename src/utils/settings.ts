/**
 * Settings Manager for Armenian Name Indexer
 * Handles loading and saving settings to document properties
 */

/* global Word */

import { IndexerSettings } from "../types";
import { DEFAULT_SUFFIXES, DEFAULT_WORD_COUNT, createArmenianNamePattern } from "./armenian";

/**
 * Settings key prefix for document settings
 */
const SETTINGS_PREFIX = "ArmenianIndexer_";

/**
 * Settings keys
 */
const SETTINGS_KEYS = {
  EXCEPTIONS: SETTINGS_PREFIX + "exceptions",
  PATTERN: SETTINGS_PREFIX + "pattern",
  SUFFIXES: SETTINGS_PREFIX + "suffixes",
  WORD_COUNT_MIN: SETTINGS_PREFIX + "wordCountMin",
  WORD_COUNT_MAX: SETTINGS_PREFIX + "wordCountMax"
};

/**
 * Get default settings
 * 
 * @returns Default IndexerSettings object
 */
export function getDefaultSettings(): IndexerSettings {
  const defaultPattern = createArmenianNamePattern(DEFAULT_WORD_COUNT);
  
  return {
    exceptions: [],
    pattern: defaultPattern.source,
    suffixes: DEFAULT_SUFFIXES,
    wordCount: DEFAULT_WORD_COUNT
  };
}

/**
 * Load settings from document properties
 * Falls back to defaults if not found
 * 
 * @returns Promise resolving to IndexerSettings
 */
export async function loadSettings(): Promise<IndexerSettings> {
  return Word.run(async (context) => {
    const settings = context.document.settings;
    settings.load("items");
    await context.sync();

    const defaults = getDefaultSettings();

    // Load exceptions
    const exceptionsJson = settings.items[SETTINGS_KEYS.EXCEPTIONS];
    const exceptions = exceptionsJson 
      ? JSON.parse(exceptionsJson) 
      : defaults.exceptions;

    // Load pattern
    const pattern = settings.items[SETTINGS_KEYS.PATTERN] || defaults.pattern;

    // Load suffixes
    const suffixesJson = settings.items[SETTINGS_KEYS.SUFFIXES];
    const suffixes = suffixesJson 
      ? JSON.parse(suffixesJson) 
      : defaults.suffixes;

    // Load word count
    const wordCountMin = settings.items[SETTINGS_KEYS.WORD_COUNT_MIN];
    const wordCountMax = settings.items[SETTINGS_KEYS.WORD_COUNT_MAX];
    const wordCount = {
      min: wordCountMin ? parseInt(wordCountMin, 10) : defaults.wordCount.min,
      max: wordCountMax ? parseInt(wordCountMax, 10) : defaults.wordCount.max
    };

    return {
      exceptions,
      pattern,
      suffixes,
      wordCount
    };
  });
}

/**
 * Save settings to document properties
 * 
 * @param settings - Settings to save
 * @returns Promise that resolves when settings are saved
 */
export async function saveSettings(settings: IndexerSettings): Promise<void> {
  return Word.run(async (context) => {
    const docSettings = context.document.settings;

    // Save exceptions
    docSettings.add(SETTINGS_KEYS.EXCEPTIONS, JSON.stringify(settings.exceptions));

    // Save pattern
    docSettings.add(SETTINGS_KEYS.PATTERN, settings.pattern);

    // Save suffixes
    docSettings.add(SETTINGS_KEYS.SUFFIXES, JSON.stringify(settings.suffixes));

    // Save word count
    docSettings.add(SETTINGS_KEYS.WORD_COUNT_MIN, settings.wordCount.min.toString());
    docSettings.add(SETTINGS_KEYS.WORD_COUNT_MAX, settings.wordCount.max.toString());

    await context.sync();
  });
}

/**
 * Clear all settings from document properties
 * Note: Office.js doesn't provide a direct way to enumerate and delete settings
 * This function overwrites with empty values instead
 * 
 * @returns Promise that resolves when settings are cleared
 */
export async function clearSettings(): Promise<void> {
  return Word.run(async (context) => {
    const settings = context.document.settings;

    // Overwrite with empty/default values instead of deleting
    Object.values(SETTINGS_KEYS).forEach(key => {
      settings.add(key, "");
    });

    await context.sync();
  });
}
