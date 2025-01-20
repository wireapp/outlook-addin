/* global window */

export const config = {
  ...window.config,
  apiVersion: window.config?.apiVersion || "v5",
};
