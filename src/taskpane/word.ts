/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import { IndexerSettings } from "../types";
import { loadSettings, saveSettings, getDefaultSettings } from "../utils/settings";
import { indexArmenianNames, clearAllIndexEntries, previewArmenianNames } from "../utils/wordOps";
import { createArmenianNamePattern } from "../utils/armenian";

/** Shared cancellation token — set `cancelled = true` to stop an in-progress operation */
let cancelToken = { cancelled: false };

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // Apply Office theme on load
    applyOfficeTheme(Office.context.officeTheme);

    // Poll for theme changes every 2 seconds (Office.js doesn't provide a reliable event)
    setInterval(() => {
      applyOfficeTheme(Office.context.officeTheme);
    }, 2000);

    document.getElementById("index-btn").onclick = handleIndexNames;
    document.getElementById("preview-btn").onclick = handlePreviewNames;
    document.getElementById("clear-btn").onclick = handleClearEntries;
    document.getElementById("cancel-btn").onclick = handleCancel;
    document.getElementById("save-btn").onclick = handleSaveSettings;
    document.getElementById("load-btn").onclick = handleLoadSettings;
    document.getElementById("reset-btn").onclick = handleResetSettings;
    document.getElementById("exceptions-file").onchange = handleFileUpload;

    // Auto-regenerate the regex pattern whenever word count inputs change
    document.getElementById("wordcount-min").oninput = regeneratePattern;
    document.getElementById("wordcount-max").oninput = regeneratePattern;

    handleLoadSettings();
  }
});

/**
 * Detect whether the Office theme is dark by measuring the brightness
 * of the body background color, then set data-theme on <body>.
 */
function applyOfficeTheme(theme: Office.OfficeTheme) {
  try {
    const bg = theme && theme.bodyBackgroundColor;
    if (!bg || bg.length < 7) {
      document.body.dataset.theme = "light";
      return;
    }
    const r = parseInt(bg.slice(1, 3), 16);
    const g = parseInt(bg.slice(3, 5), 16);
    const b = parseInt(bg.slice(5, 7), 16);
    // Perceived brightness (ITU-R BT.601 luma)
    const brightness = (r * 299 + g * 587 + b * 114) / 1000;
    document.body.dataset.theme = brightness < 140 ? "dark" : "light";
  } catch {
    document.body.dataset.theme = "light";
  }
}

// ---------------------------------------------------------------------------
// Settings handlers
// ---------------------------------------------------------------------------

async function handleLoadSettings() {
  try {
    const settings = await loadSettings();
    populateUIFromSettings(settings);
    showMessage("Settings loaded successfully", "success");
  } catch (error) {
    console.error("Error loading settings:", error);
    showMessage("Error loading settings: " + toErrorMessage(error), "error");
    populateUIFromSettings(getDefaultSettings());
  }
}

async function handleSaveSettings() {
  try {
    const settings = getSettingsFromUI();

    try {
      new RegExp(settings.pattern);
    } catch (error) {
      showMessage("Invalid regex pattern: " + toErrorMessage(error), "error");
      return;
    }

    await saveSettings(settings);
    showMessage("Settings saved successfully", "success");
  } catch (error) {
    console.error("Error saving settings:", error);
    showMessage("Error saving settings: " + toErrorMessage(error), "error");
  }
}

function handleResetSettings() {
  const defaults = getDefaultSettings();
  populateUIFromSettings(defaults);
  showMessage("Settings reset to defaults", "info");
}

// ---------------------------------------------------------------------------
// File upload
// ---------------------------------------------------------------------------

function handleFileUpload(event: Event) {
  const input = event.target as HTMLInputElement;
  const file = input.files?.[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const text = e.target?.result as string;
    (document.getElementById("exceptions-textarea") as HTMLTextAreaElement).value = text;
    document.getElementById("file-name").textContent = file.name;
    showMessage(`Loaded ${file.name}`, "success");
  };
  reader.onerror = () => showMessage("Error reading file", "error");
  reader.readAsText(file, "UTF-8");
}

// ---------------------------------------------------------------------------
// Core action handlers
// ---------------------------------------------------------------------------

async function handleIndexNames() {
  const settings = getSettingsFromUI();
  if (!validatePattern(settings.pattern)) return;

  const confirmed = confirm(
    "This will insert XE index fields throughout the document.\n\n" +
      "Tip: Save a copy of your document before proceeding — " +
      "the operation cannot be easily undone in a single step.\n\nProceed?"
  );
  if (!confirmed) return;

  cancelToken = { cancelled: false };

  showProgress(true, true /* show cancel */);
  hideResult();
  hidePreview();
  setButtonsEnabled(false);

  try {
    await Word.run(async (context) => {
      const result = await indexArmenianNames(context, settings, updateProgress, cancelToken);

      showProgress(false);

      if (cancelToken.cancelled) {
        showMessage(`Cancelled — ${result.indexed} names indexed before stopping.`, "info");
      } else if (result.errors.length > 0) {
        showMessage(
          `Indexed ${result.indexed} names, skipped ${result.skipped}.\n\nErrors:\n${result.errors.join("\n")}`,
          "warning"
        );
      } else if (result.indexed === 0) {
        showMessage("No names were indexed. Check your pattern and exceptions list.", "warning");
      } else {
        showMessage(
          `Successfully indexed ${result.indexed} names (${result.skipped} skipped)`,
          "success"
        );
      }
    });
  } catch (error) {
    console.error("Error indexing names:", error);
    showProgress(false);
    showMessage("Error indexing names: " + toErrorMessage(error), "error");
  } finally {
    setButtonsEnabled(true);
  }
}

async function handlePreviewNames() {
  const settings = getSettingsFromUI();
  if (!validatePattern(settings.pattern)) return;

  showProgress(true, false);
  hideResult();
  hidePreview();
  setButtonsEnabled(false);

  try {
    await Word.run(async (context) => {
      const entries = await previewArmenianNames(context, settings, updateProgress);

      showProgress(false);

      if (entries.length === 0) {
        showMessage("No names found matching current settings.", "info");
      } else {
        showPreview(entries);
      }
    });
  } catch (error) {
    console.error("Error previewing names:", error);
    showProgress(false);
    showMessage("Error during preview: " + toErrorMessage(error), "error");
  } finally {
    setButtonsEnabled(true);
  }
}

async function handleClearEntries() {
  const confirmed = confirm(
    "Are you sure you want to clear all index entries? This cannot be undone."
  );
  if (!confirmed) return;

  showProgress(true, false);
  hideResult();
  hidePreview();
  setButtonsEnabled(false);

  try {
    await Word.run(async (context) => {
      const count = await clearAllIndexEntries(context, updateProgress);
      showProgress(false);

      if (count === 0) {
        showMessage("No index entries found", "info");
      } else {
        showMessage(`Successfully removed ${count} index entries`, "success");
      }
    });
  } catch (error) {
    console.error("Error clearing entries:", error);
    showProgress(false);
    showMessage("Error clearing entries: " + toErrorMessage(error), "error");
  } finally {
    setButtonsEnabled(true);
  }
}

function handleCancel() {
  cancelToken.cancelled = true;
  document.getElementById("status-text").textContent = "Cancelling…";
}

// ---------------------------------------------------------------------------
// Pattern auto-regeneration
// ---------------------------------------------------------------------------

function regeneratePattern() {
  const min = parseInt((document.getElementById("wordcount-min") as HTMLInputElement).value, 10);
  const max = parseInt((document.getElementById("wordcount-max") as HTMLInputElement).value, 10);

  if (isNaN(min) || isNaN(max) || min < 1 || max < min) return;

  const pattern = createArmenianNamePattern({ min, max });
  (document.getElementById("pattern-input") as HTMLInputElement).value = pattern.source;
}

// ---------------------------------------------------------------------------
// UI helpers
// ---------------------------------------------------------------------------

function validatePattern(patternStr: string): boolean {
  try {
    new RegExp(patternStr, "g");
    return true;
  } catch (error) {
    showMessage("Invalid regex pattern: " + toErrorMessage(error), "error");
    return false;
  }
}

function updateProgress(percent: number, status: string) {
  const bar = document.getElementById("progress-bar");
  const text = document.getElementById("status-text");
  bar.style.width = percent + "%";
  text.textContent = status;
}

function showProgress(show: boolean, showCancel = false) {
  document.getElementById("progress-section").style.display = show ? "block" : "none";
  document.getElementById("cancel-btn").style.display =
    show && showCancel ? "inline-block" : "none";
  if (show) updateProgress(0, "Starting…");
}

function hideResult() {
  document.getElementById("result-section").style.display = "none";
}

function showMessage(message: string, type: "success" | "error" | "warning" | "info") {
  const section = document.getElementById("result-section");
  const messageDiv = document.getElementById("result-message");

  messageDiv.textContent = message;
  messageDiv.className = "message " + type;
  section.style.display = "block";

  if (type === "success" || type === "info") {
    setTimeout(() => {
      section.style.display = "none";
    }, 4000);
  }
}

function showPreview(entries: string[]) {
  const section = document.getElementById("preview-section");
  const count = document.getElementById("preview-count");
  const list = document.getElementById("preview-list");

  count.textContent = `${entries.length} unique name${entries.length === 1 ? "" : "s"} found`;

  list.innerHTML = "";
  entries.forEach((entry) => {
    const li = document.createElement("li");
    li.textContent = entry;
    list.appendChild(li);
  });

  section.style.display = "block";
}

function hidePreview() {
  document.getElementById("preview-section").style.display = "none";
}

function setButtonsEnabled(enabled: boolean) {
  const ids = ["index-btn", "preview-btn", "clear-btn", "save-btn", "load-btn", "reset-btn"];
  ids.forEach((id) => {
    (document.getElementById(id) as HTMLButtonElement).disabled = !enabled;
  });
}

function getSettingsFromUI(): IndexerSettings {
  const exceptionsText = (
    document.getElementById("exceptions-textarea") as HTMLTextAreaElement
  ).value;
  const pattern = (document.getElementById("pattern-input") as HTMLInputElement).value;
  const suffixesText = (document.getElementById("suffixes-textarea") as HTMLTextAreaElement).value;
  const wordCountMin = parseInt(
    (document.getElementById("wordcount-min") as HTMLInputElement).value,
    10
  );
  const wordCountMax = parseInt(
    (document.getElementById("wordcount-max") as HTMLInputElement).value,
    10
  );

  const exceptions: string[] = [];
  if (exceptionsText.trim().length > 0) {
    exceptionsText.split(/[\n,]+/).forEach((line) => {
      const trimmed = line.trim();
      if (trimmed.length > 0) exceptions.push(trimmed);
    });
  }

  const suffixes: string[] = [];
  if (suffixesText.trim().length > 0) {
    suffixesText.split(",").forEach((item) => {
      const trimmed = item.trim();
      if (trimmed.length > 0) suffixes.push(trimmed);
    });
  }

  return {
    exceptions,
    pattern,
    suffixes,
    wordCount: { min: wordCountMin, max: wordCountMax }
  };
}

function populateUIFromSettings(settings: IndexerSettings) {
  (document.getElementById("exceptions-textarea") as HTMLTextAreaElement).value =
    settings.exceptions.join("\n");
  (document.getElementById("pattern-input") as HTMLInputElement).value = settings.pattern;
  (document.getElementById("suffixes-textarea") as HTMLTextAreaElement).value =
    settings.suffixes.join(", ");
  (document.getElementById("wordcount-min") as HTMLInputElement).value =
    settings.wordCount.min.toString();
  (document.getElementById("wordcount-max") as HTMLInputElement).value =
    settings.wordCount.max.toString();
}

function toErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
