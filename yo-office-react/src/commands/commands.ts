/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

import { tryCatch } from "../taskpane/lib/utils";

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
let _count = 0;
let visible = false;
async function action(event: Office.AddinCommands.Event) {
  _count++;
  if (visible === false) {
    await Office.addin.showAsTaskpane();
    visible = true;
  } else {
    await Office.addin.hide();
    visible = false;
  }
  document.getElementById("runtimeTest").textContent = "Go" + _count;

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

async function toggleProtection(args) {
  try {
    await Excel.run(async function (context) {
      // TODO1: Queue commands to reverse the protection status of the current worksheet.
      var sheet = context.workbook.worksheets.getActiveWorksheet();

      // TODO2: Queue command to load the sheet's "protection.protected" property from the document and re-synchronize the document and task pane. These steps must be completed whenever your code needs to read information from the Office document.
      sheet.load("protection/protected");
      await context.sync();
      console.log("protected status: ", sheet.protection.protected);

      // TODO3: Move the queued toggle logic here.
      if (sheet.protection.protected) {
        sheet.protection.unprotect();
      } else {
        sheet.protection.protect();
      }

      // TODO4: Move the final call of `context.sync` here and ensure that it does not run until the toggle logic has been queued.
      await context.sync();
    });
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  }
  args.completed();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal() as any;

// The add-in command functions need to be available in global scope
g.action = action;
g.toggleProtection = toggleProtection;
