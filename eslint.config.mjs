// ESLint v10 flat config (eslintrc format was removed in v10). ESM because
// @eslint-react/eslint-plugin ships as an ES module only.
import { dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import tsParser from '@typescript-eslint/parser';
import tsPlugin from '@typescript-eslint/eslint-plugin';
import security from 'eslint-plugin-security';
import eslintReact from '@eslint-react/eslint-plugin';
import reactHooks from 'eslint-plugin-react-hooks';
import globals from 'globals';

const tsconfigRootDir = dirname(fileURLToPath(import.meta.url));
const reactConfig = eslintReact.configs['recommended-type-checked'];
const hooksRules = reactHooks.rules;

/** @type {import('eslint').Linter.Config[]} */
export default [
  { ignores: ['dist/**', 'node_modules/**'] },

  // Source files — full TypeScript type-aware linting
  {
    files: ['src/**/*.ts'],
    ignores: ['src/__tests__/**', 'src/admin/**'],
    languageOptions: {
      parser: tsParser,
      parserOptions: {
        project: './tsconfig.json',
        tsconfigRootDir,
      },
    },
    plugins: {
      '@typescript-eslint': tsPlugin,
      security,
    },
    rules: {
      // TypeScript
      'no-unused-vars': 'off',
      '@typescript-eslint/no-unused-vars': [
        'error',
        { argsIgnorePattern: '^_', varsIgnorePattern: '^_' },
      ],
      '@typescript-eslint/no-explicit-any': 'warn',
      '@typescript-eslint/no-floating-promises': 'error',

      // Security
      'security/detect-object-injection': 'warn',
      'security/detect-non-literal-regexp': 'warn',
      'security/detect-non-literal-fs-filename': 'warn',
      'security/detect-possible-timing-attacks': 'warn',
      'security/detect-unsafe-regex': 'error',
    },
  },

  // Test files — parse without project reference (not in tsconfig.json includes)
  {
    files: ['src/__tests__/**/*.ts'],
    languageOptions: {
      parser: tsParser,
    },
    plugins: {
      '@typescript-eslint': tsPlugin,
    },
    rules: {
      'no-unused-vars': 'off',
      '@typescript-eslint/no-unused-vars': [
        'error',
        { argsIgnorePattern: '^_', varsIgnorePattern: '^_' },
      ],
    },
  },

  // Admin SPA (React, built by Vite) and its jsdom tests: type-aware against
  // tsconfig.admin.json. React rules come from @eslint-react; the hooks rules
  // come from eslint-plugin-react-hooks, so its overlapping @eslint-react
  // rules are switched off rather than reporting everything twice.
  {
    files: ['src/admin/**/*.{ts,tsx}'],
    languageOptions: {
      parser: tsParser,
      parserOptions: {
        project: './tsconfig.admin.json',
        tsconfigRootDir,
        ecmaFeatures: { jsx: true },
      },
      globals: globals.browser,
    },
    settings: reactConfig.settings,
    plugins: {
      ...reactConfig.plugins,
      '@typescript-eslint': tsPlugin,
      'react-hooks': reactHooks,
      security,
    },
    rules: {
      ...Object.fromEntries(
        Object.entries(reactConfig.rules).map(([rule, level]) => [
          rule,
          rule.replace('@eslint-react/', '') in hooksRules ? 'off' : level,
        ]),
      ),
      // The classic hooks pair. The React Compiler rules in react-hooks v7
      // (set-state-in-effect, use-memo, purity, ...) assume the compiler;
      // this SPA is not compiled by it and loads its data from effects.
      'react-hooks/rules-of-hooks': 'error',
      'react-hooks/exhaustive-deps': 'warn',
      // Wants an AbortController on every fetch in an effect. The components
      // here live for the page, a resolved fetch after unmount is a no-op
      // under React 18, and aborting would route an AbortError into the
      // error state of components that are still mounted (StrictMode).
      '@eslint-react/web-api-no-leaked-fetch': 'off',

      'no-unused-vars': 'off',
      '@typescript-eslint/no-unused-vars': [
        'error',
        { argsIgnorePattern: '^_', varsIgnorePattern: '^_' },
      ],
      '@typescript-eslint/no-explicit-any': 'warn',
      '@typescript-eslint/no-floating-promises': 'error',
      // detect-object-injection is left off here: it is a Node heuristic for
      // request-controlled keys, and every hit in the SPA indexes an array by
      // number or a Record by a typed union.
      'security/detect-unsafe-regex': 'error',
    },
  },
];
