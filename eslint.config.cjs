const js = require('@eslint/js');
const eslintConfigPrettier = require('eslint-config-prettier/flat');
const eslintPluginJsonc = require('eslint-plugin-jsonc');
const eslintPluginPrettierRecommended = require('eslint-plugin-prettier/recommended');
//const googleAppsScript = require('eslint-plugin-googleappsscript');

module.exports = [
  js.configs.recommended,
  eslintConfigPrettier,
  ...eslintPluginJsonc.configs['flat/recommended-with-jsonc'],
  ...eslintPluginJsonc.configs['flat/prettier'],
  eslintPluginPrettierRecommended,
  //googleAppsScript,
  {
    languageOptions: {
      ecmaVersion: 'latest',
      sourceType: 'script'
      //globals: {
      //  ...googleAppsScript.environments.googleappsscript.globals,
      //}
    },
    rules: {
      'prefer-template': 'warn',
      'no-undef': 'off',
      'no-unused-vars': ['warn', { vars: 'local' }],
      'no-var': 'error',
      'prefer-const': 'warn',
      'prefer-arrow-callback': 'warn',
      'no-continue': 'off',
      'no-prototype-builtins': 'off',
      'prettier/prettier': [
        'error',
        {
          trailingComma: 'none',
          singleQuote: true,
          printWidth: 120,
          // below line only for windows users facing CLRF and eslint/prettier error
          // non windows users feel free to delete it
          endOfLine: 'auto'
        }
      ]
    }
  }
];
