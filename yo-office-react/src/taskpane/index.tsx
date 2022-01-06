import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";

/* global document, Office, module, require */

initializeIcons();

let isOfficeInitialized = false;

const title = "Contoso Task Pane Add-in";

const render = (Component) => {
  ReactDOM.render(
    <AppContainer>
      <ThemeProvider>
        <Component title={title} isOfficeInitialized={isOfficeInitialized} />
      </ThemeProvider>
    </AppContainer>,
    document.getElementById("container")
  );
};

/* Render application after Office initializes */
Office.onReady(function(info){
  if (info.host === Office.HostType.Excel) {
    // Do Excel-specific initialization (for example, make add-in task pane's
    // appearance compatible with Excel "green").
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
  console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
}
}
if (info.platform === Office.PlatformType.PC) {
  // Make minor layout changes in the task pane.
}
console.log(`Office.js is now ready in ${info.host} on ${info.platform}`);
  isOfficeInitialized = true
  render(App)
})


if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
