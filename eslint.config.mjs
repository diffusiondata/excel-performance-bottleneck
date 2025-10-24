import typescriptEslint from "@typescript-eslint/eslint-plugin";
import officeAddins from "eslint-plugin-office-addins";
import react from "eslint-plugin-react";
import globals from "globals";
import tsParser from "@typescript-eslint/parser";
import path from "node:path";
import { fileURLToPath } from "node:url";
import jsdoc from "eslint-plugin-jsdoc";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

export default [
    {
        ignores: [
            "**/dist/",
            "**/node_modules/",
            "**/*.test.ts",
            "**/*.test.tsx",
            "**/src/__tests__/*",
            "**/vitest.config.ts",
            "**/vite.config.ts",
            "bin/**",
            "vitest.workspace.ts",
            "prettier.config.js"
        ]
    },
    {
        files: ["**/*.{js,jsx,ts,tsx}"],
        languageOptions: {
            globals: {
                ...globals.browser,
                ...globals.node,
                console: "readonly",
                setTimeout: "readonly",
                Excel: "readonly",
                OfficeRuntime: "readonly",
                OfficeExtension: "readonly",
                global: "writable",
                CustomFunctions: "writable",
            },
            parser: tsParser,
            ecmaVersion: 2018,
            sourceType: "module",
            parserOptions: {
                projectService: true,
            },
        },
        settings: {
            react: {
                version: "detect",
            },
            jsdoc: {
                mode: "typescript"
            }
        },
        plugins: {
            "@typescript-eslint": typescriptEslint,
            "office-addins": officeAddins,
            "react": react,
            "jsdoc": jsdoc,
        },
        rules: {
            // Base TypeScript rules
            ...typescriptEslint.configs.recommended.rules,

            "office-addins/no-context-sync-in-loop": "error",
            "office-addins/load-object-before-read": "error",

            ...react.configs.recommended.rules,

            "@typescript-eslint/no-misused-promises": ["warn", {
                checksVoidReturn: false,
                checksConditionals: true,
            }],
            "@typescript-eslint/await-thenable": "error",
            "@typescript-eslint/no-floating-promises": "error",
            "@typescript-eslint/require-await": "warn",

            "@typescript-eslint/no-explicit-any": "warn",
            "@typescript-eslint/no-empty-function": "off",
            "@typescript-eslint/no-non-null-assertion": "error",
            "react/react-in-jsx-scope": "off",
            "react/prop-types": "off", // We use TypeScript

            "no-duplicate-imports": "warn",
            "@typescript-eslint/no-unused-vars": ["warn", {
                "argsIgnorePattern": "^_",
                "caughtErrorsIgnorePattern": "^_"
            }],

            ...jsdoc.configs.recommended.rules,

            "jsdoc/require-jsdoc": ["off"],

            // Because we're using TypeScript
            "jsdoc/require-param-type": "off",
            "jsdoc/require-returns-type": "warn",
            "jsdoc/require-yields-type": "off",
            "jsdoc/require-throws-type": "off"
        },
    }
];