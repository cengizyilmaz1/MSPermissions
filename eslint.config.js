'use strict';

const js = require('@eslint/js');
const globals = require('globals');

module.exports = [
    {
        ignores: [
            'node_modules/**',
            'docs/**',
            '.generated/**',
            '.cache/**',
            'data/**',
            'fixtures/**',
            'customdata/**'
        ]
    },
    js.configs.recommended,
    {
        // Node.js CommonJS sources (build pipeline, tests, config)
        files: ['**/*.js'],
        ignores: ['src/templates/js/**'],
        languageOptions: {
            ecmaVersion: 2023,
            sourceType: 'commonjs',
            globals: {
                ...globals.node
            }
        },
        rules: {
            'no-unused-vars': ['warn', { argsIgnorePattern: '^_', varsIgnorePattern: '^_' }],
            'no-console': 'off',
            'prefer-const': 'error',
            'no-var': 'error',
            eqeqeq: ['error', 'smart'],
            'no-implicit-coercion': 'off',
            curly: ['error', 'all']
        }
    },
    {
        // Browser runtime bundle served to the site
        files: ['src/templates/js/**/*.js'],
        languageOptions: {
            ecmaVersion: 2023,
            sourceType: 'script',
            globals: {
                ...globals.browser
            }
        },
        rules: {
            'no-unused-vars': ['warn', { argsIgnorePattern: '^_', varsIgnorePattern: '^_' }],
            'prefer-const': 'error',
            'no-var': 'error',
            eqeqeq: ['error', 'smart']
        }
    }
];
