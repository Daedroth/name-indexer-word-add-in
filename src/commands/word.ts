/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office Word console */

import { loadSettingsInContext } from "../utils/settings";
import { indexArmenianNames } from "../utils/wordOps";

type ProgressDialogMessage =
  | { type: "progress"; percent: number; status: string }
  | { type: "warning"; message: string }
  | { type: "complete"; message: string; severity?: "success" | "info" | "warning" | "error" }
  | { type: "error"; message: string };

type ProgressDialogCommand = { type: "cancel" };

function toErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

async function openProgressDialog(onCommand: (cmd: ProgressDialogCommand) => void): Promise<Office.Dialog | null> {
  // Use a relative URL so it works in dev (localhost) and prod (GitHub Pages).
  const dialogUrl = new URL("progress.html", window.location.href).toString();

  return new Promise((resolve) => {
    Office.context.ui.displayDialogAsync(
      dialogUrl,
      {
        height: 35,
        width: 30,
        displayInIframe: true,
      },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.warn("Failed to open progress dialog:", result.error);
          resolve(null);
          return;
        }

        const dialog = result.value;

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          try {
            const data = JSON.parse(arg.message) as ProgressDialogCommand;
            onCommand(data);
          } catch (error) {
            console.error("Failed to parse dialog message:", toErrorMessage(error));
          }
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          // If user closes the dialog, treat it like a cancel request.
          onCommand({ type: "cancel" });
        });

        resolve(dialog);
      }
    );
  });
}

function trySendDialogMessage(dialog: Office.Dialog | null, message: ProgressDialogMessage) {
  if (!dialog) return;
  try {
    dialog.messageChild(JSON.stringify(message));
  } catch (error) {
    console.warn("Failed to send dialog message:", toErrorMessage(error));
  }
}

/**
 * Index Armenian names in the document using saved settings.
 * Triggered from the ribbon without opening the task pane.
 */
export async function indexArmenianNamesCommand(event: Office.AddinCommands.Event) {
  let isCompleted = false;
  const safeComplete = () => {
    try {
      if (isCompleted) return;
      isCompleted = true;
      // In some edge cases (sideload/runtime issues), the event argument can be missing.
      // Guard so we never throw here (otherwise Office will treat it as a timeout).
      if (event && typeof event.completed === "function") {
        event.completed();
      }
    } catch (error) {
      console.error("Failed to call event.completed():", error instanceof Error ? error.message : String(error));
    }
  };

  let dialog: Office.Dialog | null = null;

  try {
    // Add-in Commands are time-limited and must call event.completed() quickly.
    // We complete immediately, then continue the long-running indexing flow.
    safeComplete();

    const cancelToken = { cancelled: false };

    // Dialog is best-effort. If it fails (popup blocked, older host), we still proceed.
    dialog = await openProgressDialog((cmd) => {
      if (cmd.type === "cancel") {
        cancelToken.cancelled = true;
      }
    });

    // Throttle progress messages to reduce UI chatter.
    let lastProgressSentAt = 0;
    let lastProgressPercent = -1;
    const onProgress = (percent: number, status: string) => {
      const now = Date.now();
      const rounded = Math.max(0, Math.min(100, Math.floor(percent)));

      // Send if percent changed, but no more than ~8 times/sec.
      if (rounded === lastProgressPercent && now - lastProgressSentAt < 250) return;

      lastProgressSentAt = now;
      lastProgressPercent = rounded;
      trySendDialogMessage(dialog, { type: "progress", percent: rounded, status });
    };

    trySendDialogMessage(dialog, { type: "progress", percent: 0, status: "Loading settings…" });

    await Word.run(async (context) => {
      // Load settings within the same context — avoids nested Word.run()
      const settings = await loadSettingsInContext(context);

      if (!settings.exceptions || settings.exceptions.length === 0) {
        trySendDialogMessage(dialog, {
          type: "warning",
          message:
            "Exceptions list is empty. If you intended to exclude words (e.g., place names), open Settings to add exceptions. You can Cancel if this was accidental.",
        });
      }

      const result = await indexArmenianNames(context, settings, onProgress, cancelToken);

      console.log(
        result.indexed > 0
          ? `Indexed ${result.indexed} names (${result.skipped} skipped).`
          : "No Armenian names found. Check pattern and exceptions."
      );

      if (result.errors.length > 0) {
        console.warn("Indexing errors:", result.errors.join("\n"));
      }

      if (cancelToken.cancelled) {
        trySendDialogMessage(dialog, {
          type: "complete",
          severity: "info",
          message: `Cancelled — ${result.indexed} names indexed before stopping.`,
        });
      } else if (result.errors.length > 0) {
        trySendDialogMessage(dialog, {
          type: "complete",
          severity: "warning",
          message: `Indexed ${result.indexed} names, skipped ${result.skipped}.\n\nErrors:\n${result.errors.join("\n")}`,
        });
      } else if (result.indexed === 0) {
        trySendDialogMessage(dialog, {
          type: "complete",
          severity: "warning",
          message: "No names were indexed. Check your pattern and exceptions list.",
        });
      } else {
        trySendDialogMessage(dialog, {
          type: "complete",
          severity: "success",
          message: `Successfully indexed ${result.indexed} names (${result.skipped} skipped).`,
        });
      }
    });
  } catch (error) {
    const message = toErrorMessage(error);
    console.error("Error indexing names:", message);
    trySendDialogMessage(dialog, { type: "error", message: "Error indexing names: " + message });
  } finally {
    // Always signal completion so Word unblocks the ribbon button
    safeComplete();
  }
}
