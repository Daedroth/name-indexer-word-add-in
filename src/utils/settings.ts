/**
 * Settings Manager for Name Indexer
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
  WORD_COUNT_MAX: SETTINGS_PREFIX + "wordCountMax",
  NORMALIZATION_ENABLED: SETTINGS_PREFIX + "normalizationEnabled",
  NORMALIZATION_MODE: SETTINGS_PREFIX + "normalizationMode",
  NORMALIZATION_CUSTOM_RULES: SETTINGS_PREFIX + "normalizationCustomRules",
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
    wordCount: DEFAULT_WORD_COUNT,
    // Preserve current behavior: Armenian-optimized normalization on by default.
    normalization: {
      enabled: true,
      mode: "armenian",
      customRules: [],
    },
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
  const normalizationEnabledItem = settings.getItemOrNullObject(SETTINGS_KEYS.NORMALIZATION_ENABLED);
  const normalizationModeItem = settings.getItemOrNullObject(SETTINGS_KEYS.NORMALIZATION_MODE);
  const normalizationCustomRulesItem = settings.getItemOrNullObject(SETTINGS_KEYS.NORMALIZATION_CUSTOM_RULES);

  // Load value and isNullObject for all at once (single sync)
  exceptionsItem.load("value");
  patternItem.load("value");
  suffixesItem.load("value");
  wordCountMinItem.load("value");
  wordCountMaxItem.load("value");
  normalizationEnabledItem.load("value");
  normalizationModeItem.load("value");
  normalizationCustomRulesItem.load("value");

  await context.sync();

  const defaults = getDefaultSettings();

  const exceptions = exceptionsItem.isNullObject ? defaults.exceptions : JSON.parse(exceptionsItem.value as string);

  const pattern = patternItem.isNullObject ? defaults.pattern : (patternItem.value as string);

  const suffixes = suffixesItem.isNullObject ? defaults.suffixes : JSON.parse(suffixesItem.value as string);

  const wordCount = {
    min: wordCountMinItem.isNullObject ? defaults.wordCount.min : parseInt(wordCountMinItem.value as string, 10),
    max: wordCountMaxItem.isNullObject ? defaults.wordCount.max : parseInt(wordCountMaxItem.value as string, 10),
  };

  const normalizationEnabled = normalizationEnabledItem.isNullObject
    ? defaults.normalization.enabled
    : (normalizationEnabledItem.value as string) === "true";

  const normalizationMode = normalizationModeItem.isNullObject
    ? defaults.normalization.mode
    : (normalizationModeItem.value as IndexerSettings["normalization"]["mode"]);

  const normalizationCustomRules = normalizationCustomRulesItem.isNullObject
    ? defaults.normalization.customRules
    : JSON.parse(normalizationCustomRulesItem.value as string);

  return {
    exceptions,
    pattern,
    suffixes,
    wordCount,
    normalization: {
      enabled: normalizationEnabled,
      mode: normalizationMode,
      customRules: normalizationCustomRules,
    },
  };
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

    docSettings.add(SETTINGS_KEYS.NORMALIZATION_ENABLED, settings.normalization.enabled ? "true" : "false");
    docSettings.add(SETTINGS_KEYS.NORMALIZATION_MODE, settings.normalization.mode);
    docSettings.add(SETTINGS_KEYS.NORMALIZATION_CUSTOM_RULES, JSON.stringify(settings.normalization.customRules));

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

    const items = Object.values(SETTINGS_KEYS).map((key) => settings.getItemOrNullObject(key));
    await context.sync();

    items.forEach((item) => {
      if (!item.isNullObject) {
        item.delete();
      }
    });

    await context.sync();
  });
}
