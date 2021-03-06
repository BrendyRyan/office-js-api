/* global document, Office */

// convert to async-await
(async function () {
  "use strict";

  await Office.onReady();
  // TODO1: Assign handler to the OK button.
  document.getElementById("ok-button").onclick = sendStringToParentPage;

  // TODO2: Create the OK button handler
  function sendStringToParentPage() {
    var userName = (document.getElementById("name-box") as any).value;
    Office.context.ui.messageParent(userName);
  }
})();
