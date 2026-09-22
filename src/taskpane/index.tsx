import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import * as React from "react";
import { createRoot } from "react-dom/client";

/* global document, Office, module, require */

let isOfficeInitialized = false;

const title = "Task Pane Add-in";

const render = (Component) => {
  const rootElement: HTMLElement | null = document.getElementById("container");
  const root = rootElement ? createRoot(rootElement) : undefined;

  root?.render(
    <FluentProvider theme={webLightTheme}>
      <Component title={title} isOfficeInitialized={isOfficeInitialized} />
    </FluentProvider>
  );
};

/* Render application after Office initializes */
Office.onReady(() => {
  isOfficeInitialized = true;
  render(App);
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
