import eslint from 'eslint';
import neostandard from 'neostandard';

const { FlatESLint } = eslint;

export default [
  ...neostandard({
    ts: true,
  }),
  {
    rules: {
      'no-unused-vars': 'off',
      '@typescript-eslint/no-unused-vars': ['error', {
        argsIgnorePattern: '^_',
        varsIgnorePattern: '^_',
      }],
    },
  },
];
