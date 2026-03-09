import { indexArmenianNamesCommand } from "./word";

/* global Office, console */

// Register the add-in commands with the Office host application.
Office.onReady(() => {
  try {
    Office.actions.associate("indexArmenianNamesCommand", indexArmenianNamesCommand);
  } catch (error) {
    // If association fails, Office will time out when the user clicks the command.
    // Logging here makes the failure diagnosable via DevTools.
    // eslint-disable-next-line no-console
    console.error("Failed to associate add-in command:", error);
  }
});
