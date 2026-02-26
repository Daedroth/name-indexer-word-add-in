/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office Word console */

import { loadSettingsInContext } from "../utils/settings";
import { indexArmenianNames } from "../utils/wordOps";

/**
 * Index Armenian names in the document using saved settings.
 * Triggered from the ribbon without opening the task pane.
 */
export async function indexArmenianNamesCommand(event: Office.AddinCommands.Event) {
  try {
    await Word.run(async (context) => {
      // Load settings within the same context — avoids nested Word.run()
      const settings = await loadSettingsInContext(context);

      const result = await indexArmenianNames(context, settings);

      console.log(
        result.indexed > 0
          ? `Indexed ${result.indexed} names (${result.skipped} skipped).`
          : "No Armenian names found. Check pattern and exceptions."
      );

      if (result.errors.length > 0) {
        console.warn("Indexing errors:", result.errors.join("\n"));
      }
    });
  } catch (error) {
    console.error("Error indexing names:", error instanceof Error ? error.message : String(error));
  }

  // Always signal completion so Word unblocks the ribbon button
  event.completed();
}
