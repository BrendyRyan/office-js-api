/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

import { tryCatch } from "../taskpane/taskpane";

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
let _count = 0;
function action(event: Office.AddinCommands.Event) {
  _count++;
  Office.addin.showAsTaskpane();
  document.getElementById("random").textContent = "Go" + _count;

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

async function toggleProtection(args) {
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
g.toggleProtection = tryCatch(toggleProtection);
