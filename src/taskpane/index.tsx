import { hot } from "react-hot-loader/root";
import App from "./components/App";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";

/* global document, Office, module, require */

initializeIcons();

let isOfficeInitialized = false;

const title = "Task Pane Add-in";

const HotApp = hot(App);

const render = (Component) => {
  ReactDOM.render(
    <ThemeProvider>
      <Component title={title} isOfficeInitialized={isOfficeInitialized} />
    </ThemeProvider>,
    document.getElementById("container")
  );
};

/* Render application after Office initializes */
Office.onReady(() => {
  isOfficeInitialized = true;
  render(HotApp);
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(hot(NextApp));
  });
}
