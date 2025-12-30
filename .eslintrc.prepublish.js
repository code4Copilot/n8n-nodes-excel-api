module.exports = {
  root: true,
  env: {
    es2021: true,
    node: true,
    jest: true,
  },
  parser: '@typescript-eslint/parser',
  parserOptions: {
    ecmaVersion: 'latest',
    sourceType: 'module',
  },
  ignorePatterns: ['dist/', 'node_modules/'],
  extends: ['eslint:recommended'],
  rules: {
    // Allow unused vars if they are named 'this' (TypeScript type annotations)
    'no-unused-vars': ['error', { 
      varsIgnorePattern: '^this$',
      argsIgnorePattern: '^this$'
    }],
  },
};