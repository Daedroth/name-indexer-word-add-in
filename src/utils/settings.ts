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
 * Load settings from an existing Word request context.
 * Use this variant when already inside a Word.run() call.
 *
 * @param context - Word request context
 * @returns Promise resolving to IndexerSettings
 */
export async function loadSettingsInContext(context: Word.RequestContext): Promise<IndexerSettings> {
  const settings = context.document.settings;

  // Queue getItemOrNullObject calls for all keys (no sync yet)
  const exceptionsItem = settings.getItemOrNullObject(SETTINGS_KEYS.EXCEPTIONS);
  const patternItem = settings.getItemOrNullObject(SETTINGS_KEYS.PATTERN);
  const suffixesItem = settings.getItemOrNullObject(SETTINGS_KEYS.SUFFIXES);
  const wordCountMinItem = settings.getItemOrNullObject(SETTINGS_KEYS.WORD_COUNT_MIN);
  const wordCountMaxItem = settings.getItemOrNullObject(SETTINGS_KEYS.WORD_COUNT_MAX);

  // Load value and isNullObject for all at once (single sync)
  exceptionsItem.load("value");
  patternItem.load("value");
  suffixesItem.load("value");
  wordCountMinItem.load("value");
  wordCountMaxItem.load("value");

  await context.sync();

  const defaults = getDefaultSettings();

  const exceptions = exceptionsItem.isNullObject
    ? defaults.exceptions
    : JSON.parse(exceptionsItem.value as string);

  const pattern = patternItem.isNullObject ? defaults.pattern : (patternItem.value as string);

  const suffixes = suffixesItem.isNullObject
    ? defaults.suffixes
    : JSON.parse(suffixesItem.value as string);

  const wordCount = {
    min: wordCountMinItem.isNullObject
      ? defaults.wordCount.min
      : parseInt(wordCountMinItem.value as string, 10),
    max: wordCountMaxItem.isNullObject
      ? defaults.wordCount.max
      : parseInt(wordCountMaxItem.value as string, 10)
  };

  return { exceptions, pattern, suffixes, wordCount };
}

/**
 * Load settings from document properties.
 * Falls back to defaults if not found.
 *
 * @returns Promise resolving to IndexerSettings
 */
export async function loadSettings(): Promise<IndexerSettings> {
  return Word.run((context) => loadSettingsInContext(context));
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

    docSettings.add(SETTINGS_KEYS.EXCEPTIONS, JSON.stringify(settings.exceptions));
    docSettings.add(SETTINGS_KEYS.PATTERN, settings.pattern);
    docSettings.add(SETTINGS_KEYS.SUFFIXES, JSON.stringify(settings.suffixes));
    docSettings.add(SETTINGS_KEYS.WORD_COUNT_MIN, settings.wordCount.min.toString());
    docSettings.add(SETTINGS_KEYS.WORD_COUNT_MAX, settings.wordCount.max.toString());

    await context.sync();
  });
}

/**
 * Delete all add-in settings from document properties
 *
 * @returns Promise that resolves when settings are deleted
 */
export async function clearSettings(): Promise<void> {
  return Word.run(async (context) => {
    const settings = context.document.settings;

    // Use getItemOrNullObject + delete() for each key
    for (const key of Object.values(SETTINGS_KEYS)) {
      const item = settings.getItemOrNullObject(key);
      item.load("isNullObject");
      await context.sync();
      if (!item.isNullObject) {
        item.delete();
      }
    }

    await context.sync();
  });
}
