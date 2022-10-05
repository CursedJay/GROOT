module.exports = {
  env: {
    es6: true
  },
  parserOptions: {
    ecmaVersion: 2021
  },
  extends: ['eslint-config-prettier'],
  plugins: ['eslint-plugin-prettier'],
  rules: {
    'no-continue': 'off',
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
};
