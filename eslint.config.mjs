import js from '@eslint/js';
import globals from 'globals';

export default [
  js.configs.recommended,
  {
    files: ['app.js'],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'script',
      globals: {
        ...globals.browser,
        pdfjsLib: 'readonly',
        html2canvas: 'readonly',
        Chart: 'readonly',
        window: 'readonly',
        confirm: 'readonly',
        localStorage: 'readonly',
        navigator: 'readonly'
      }
    },
    rules: {
      'no-unused-vars': ['error', { varsIgnorePattern: '^[A-Z_]+$', caughtErrors: 'none' }],
      'no-alert': 'off',
      'no-confusing-arrow': 'off'
    }
  }
];
