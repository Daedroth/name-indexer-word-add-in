import { indexArmenianNamesCommand } from "./word";

/* global Office */

// Register the add-in commands with the Office host application.
Office.onReady(async () => {
  Office.actions.associate("indexArmenianNamesCommand", indexArmenianNamesCommand);
});
