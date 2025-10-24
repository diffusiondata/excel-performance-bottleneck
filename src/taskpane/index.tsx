import React from "react";
import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { createRoot } from "react-dom/client";

/* global document, Office */

let isOfficeInitialized = false;

const title = "Performance regression reproducer.";

const render = (Component: typeof App): void => {
  createRoot(document.getElementById("container") as HTMLElement).render(
    <React.StrictMode>
      <FluentProvider theme={webLightTheme}>
        <Component title={title} isOfficeInitialized={isOfficeInitialized} />
      </FluentProvider>
    </React.StrictMode>
  );
};

/* Render application after Office initializes */
void Office.onReady(() => {
  isOfficeInitialized = true;
  render(App);
});
