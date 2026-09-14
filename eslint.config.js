const officeAddins = require("eslint-plugin-office-addins");
const tsParser = require("@typescript-eslint/parser");
const tsEsLint = require("typescript-eslint");

export default [
  ...tsEsLint.configs.recommended,
  ...officeAddins.configs.recommended,
  {
    plugins: {
      "office-addins": officeAddins,
    },
    languageOptions: {
      parser: tsParser,
    },
  },
];
