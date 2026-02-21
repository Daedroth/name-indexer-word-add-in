/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office Word console */

import { loadSettings } from "../utils/settings";
import { indexArmenianNames } from "../utils/wordOps";

/**
 * Index Armenian names in the document using saved settings
 * This command can be triggered from the ribbon without opening the task pane
 * @param event
 */
export async function indexArmenianNamesCommand(event: Office.AddinCommands.Event) {
  try {
    await Word.run(async (context) => {
      // Load settings from document
      const settings = await loadSettings();
      
      // Run indexing with default settings
      const result = await indexArmenianNames(context, settings);
      
      // Show completion message
      if (result.indexed === 0) {
        Office.context.ui.displayDialogAsync(
          "about:blank",  
          { height: 30, width: 40 },
          (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              const dialog = asyncResult.value;
              dialog.addEventHandler(Office.EventType.DialogMessageReceived, () => {
                dialog.close();
              });
            }
          }
        );
        console.log("No names were indexed");
      } else {
        console.log(`Successfully indexed ${result.indexed} names (${result.skipped} skipped)`);
      }
    });
  } catch (error) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    console.error("Error indexing names:", error);
  }

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

// Legacy export for backwards compatibility
export async function insertBlueParagraphInWord(event: Office.AddinCommands.Event) {
  await indexArmenianNamesCommand(event);
}
