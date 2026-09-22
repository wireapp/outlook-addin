import officeAddins from "eslint-plugin-office-addins";
import { fixupConfigRules } from "@eslint/compat";
import tsParser from "@typescript-eslint/parser";
import globals from "globals";

const sourceFiles = ["src/**/*.{js,jsx,ts,tsx}"];

export default [
  { ignores: ["dist/**", "node_modules/**"] },
  // The bundled React rules still use rule APIs removed in ESLint 10.
  ...fixupConfigRules(officeAddins.configs.react).map((config) => ({
    ...config,
    files: sourceFiles,
  })),
  {
    files: sourceFiles,
    languageOptions: {
      parser: tsParser,
      globals: globals.browser,
    },
  },
  {
    files: ["src/**/*.{ts,tsx}"],
    rules: {
      // TypeScript checks declarations and names, including type-only references.
      "no-undef": "off",
      "no-redeclare": "off",
    },
  },
];
