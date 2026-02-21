/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import { IndexerSettings } from "../types";
import { loadSettings, saveSettings, getDefaultSettings } from "../utils/settings";
import { indexArmenianNames, clearAllIndexEntries } from "../utils/wordOps";
import { createArmenianNamePattern } from "../utils/armenian";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    // Wire up event handlers
    document.getElementById("index-btn").onclick = handleIndexNames;
    document.getElementById("clear-btn").onclick = handleClearEntries;
    document.getElementById("save-btn").onclick = handleSaveSettings;
    document.getElementById("load-btn").onclick = handleLoadSettings;
    document.getElementById("reset-btn").onclick = handleResetSettings;
    document.getElementById("exceptions-file").onchange = handleFileUpload;
    
    // Load settings on startup
    handleLoadSettings();
  }
});

/**
 * Load settings from document and populate UI
 */
async function handleLoadSettings() {
  try {
    const settings = await loadSettings();
    populateUIFromSettings(settings);
    showMessage("Settings loaded successfully", "success");
  } catch (error) {
    console.error("Error loading settings:", error);
    showMessage("Error loading settings: " + error.message, "error");
    // Load defaults on error
    populateUIFromSettings(getDefaultSettings());
  }
}

/**
 * Save settings from UI to document
 */
async function handleSaveSettings() {
  try {
    const settings = getSettingsFromUI();
    
    // Validate pattern is valid regex
    try {
      new RegExp(settings.pattern);
    } catch (error) {
      showMessage("Invalid regex pattern: " + error.message, "error");
      return;
    }
    
    await saveSettings(settings);
    showMessage("Settings saved successfully", "success");
  } catch (error) {
    console.error("Error saving settings:", error);
    showMessage("Error saving settings: " + error.message, "error");
  }
}

/**
 * Reset settings to defaults
 */
function handleResetSettings() {
  const defaults = getDefaultSettings();
  populateUIFromSettings(defaults);
  showMessage("Settings reset to defaults", "info");
}

/**
 * Handle file upload for exceptions
 */
function handleFileUpload(event: Event) {
  const input = event.target as HTMLInputElement;
  const file = input.files?.[0];
  
  if (!file) {
    return;
  }
  
  const reader = new FileReader();
  reader.onload = (e) => {
    const text = e.target?.result as string;
    const textarea = document.getElementById("exceptions-textarea") as HTMLTextAreaElement;
    textarea.value = text;
    
    const fileNameSpan = document.getElementById("file-name");
    fileNameSpan.textContent = file.name;
    
    showMessage(`Loaded ${file.name}`, "success");
  };
  
  reader.onerror = () => {
    showMessage("Error reading file", "error");
  };
  
  reader.readAsText(file, "UTF-8");
}

/**
 * Handle Index Names button click
 */
async function handleIndexNames() {
  try {
    // Get settings from UI
    const settings = getSettingsFromUI();
    
    // Validate pattern
    let pattern: RegExp;
    try {
      pattern = new RegExp(settings.pattern, "g");
    } catch (error) {
      showMessage("Invalid regex pattern: " + error.message, "error");
      return;
    }
    
    // Show progress section
    showProgress(true);
    hideResult();
    setButtonsEnabled(false);
    
    // Run indexing
    await Word.run(async (context) => {
      const result = await indexArmenianNames(
        context,
        settings,
        updateProgress
      );
      
      // Show result
      showProgress(false);
      
      if (result.errors.length > 0) {
        const errorMsg = `Indexed ${result.indexed} names, skipped ${result.skipped}.\n\nErrors:\n${result.errors.join("\n")}`;
        showMessage(errorMsg, "warning");
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
    showMessage("Error indexing names: " + error.message, "error");
  } finally {
    setButtonsEnabled(true);
  }
}

/**
 * Handle Clear All Entries button click
 */
async function handleClearEntries() {
  // Confirm with user
  const confirmed = confirm("Are you sure you want to clear all index entries? This cannot be undone.");
  
  if (!confirmed) {
    return;
  }
  
  try {
    showProgress(true);
    hideResult();
    setButtonsEnabled(false);
    
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
    showMessage("Error clearing entries: " + error.message, "error");
  } finally {
    setButtonsEnabled(true);
  }
}

/**
 * Update progress UI
 */
function updateProgress(percent: number, status: string) {
  const progressBar = document.getElementById("progress-bar");
  const statusText = document.getElementById("status-text");
  
  progressBar.style.width = percent + "%";
  statusText.textContent = status;
}

/**
 * Show/hide progress section
 */
function showProgress(show: boolean) {
  const section = document.getElementById("progress-section");
  section.style.display = show ? "block" : "none";
  
  if (show) {
    updateProgress(0, "Starting...");
  }
}

/**
 * Hide result section
 */
function hideResult() {
  const section = document.getElementById("result-section");
  section.style.display = "none";
}

/**
 * Show message in result section
 */
function showMessage(message: string, type: "success" | "error" | "warning" | "info") {
  const section = document.getElementById("result-section");
  const messageDiv = document.getElementById("result-message");
  
  // Set message
  messageDiv.textContent = message;
  
  // Set CSS class based on type
  messageDiv.className = "message " + type;
  
  // Show section
  section.style.display = "block";
  
  // Auto-hide success and info messages after 4 seconds
  if (type === "success" || type === "info") {
    setTimeout(() => {
      section.style.display = "none";
    }, 4000);
  }
}

/**
 * Enable/disable buttons
 */
function setButtonsEnabled(enabled: boolean) {
  const buttons = [
    "index-btn",
    "clear-btn",
    "save-btn",
    "load-btn",
    "reset-btn"
  ];
  
  buttons.forEach(id => {
    const button = document.getElementById(id) as HTMLButtonElement;
    button.disabled = !enabled;
  });
}

/**
 * Get settings from UI form
 */
function getSettingsFromUI(): IndexerSettings {
  const exceptionsText = (document.getElementById("exceptions-textarea") as HTMLTextAreaElement).value;
  const pattern = (document.getElementById("pattern-input") as HTMLInputElement).value;
  const suffixesText = (document.getElementById("suffixes-textarea") as HTMLTextAreaElement).value;
  const wordCountMin = parseInt((document.getElementById("wordcount-min") as HTMLInputElement).value, 10);
  const wordCountMax = parseInt((document.getElementById("wordcount-max") as HTMLInputElement).value, 10);
  
  // Parse exceptions
  const exceptions: string[] = [];
  if (exceptionsText.trim().length > 0) {
    const lines = exceptionsText.split(/[\n,]+/);
    lines.forEach(line => {
      const trimmed = line.trim();
      if (trimmed.length > 0) {
        exceptions.push(trimmed);
      }
    });
  }
  
  // Parse suffixes
  const suffixes: string[] = [];
  if (suffixesText.trim().length > 0) {
    const items = suffixesText.split(",");
    items.forEach(item => {
      const trimmed = item.trim();
      if (trimmed.length > 0) {
        suffixes.push(trimmed);
      }
    });
  }
  
  return {
    exceptions,
    pattern,
    suffixes,
    wordCount: {
      min: wordCountMin,
      max: wordCountMax
    }
  };
}

/**
 * Populate UI form from settings
 */
function populateUIFromSettings(settings: IndexerSettings) {
  // Exceptions
  const exceptionsTextarea = document.getElementById("exceptions-textarea") as HTMLTextAreaElement;
  exceptionsTextarea.value = settings.exceptions.join("\n");
  
  // Pattern
  const patternInput = document.getElementById("pattern-input") as HTMLInputElement;
  patternInput.value = settings.pattern;
  
  // Suffixes
  const suffixesTextarea = document.getElementById("suffixes-textarea") as HTMLTextAreaElement;
  suffixesTextarea.value = settings.suffixes.join(", ");
  
  // Word count
  const wordCountMin = document.getElementById("wordcount-min") as HTMLInputElement;
  const wordCountMax = document.getElementById("wordcount-max") as HTMLInputElement;
  wordCountMin.value = settings.wordCount.min.toString();
  wordCountMax.value = settings.wordCount.max.toString();
}
