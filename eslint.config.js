const spfxProfile = require('@microsoft/eslint-config-spfx/lib/flat-profiles/react');

module.exports = [
  // The local preview build and the browser tests are not part of the web part.
  { ignores: ['preview/**', 'preview-dist/**', 'tests/**'] },
  ...spfxProfile,
  {
    files: ['**/*.ts', '**/*.tsx'],
    languageOptions: {
      parserOptions: {
        tsconfigRootDir: __dirname,
        project: './tsconfig.json'
      }
    }
  }
];
