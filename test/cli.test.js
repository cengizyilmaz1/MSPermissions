const assert = require('assert');
const test = require('node:test');

const { parseArgs, validateOptions, formatHelp } = require('../scripts/node/lib/cli');

const DEMO_COMMAND = {
    name: 'demo',
    summary: 'Demo command.',
    usage: 'demo [options]',
    options: [
        { name: 'input', type: 'string', required: true, description: 'Input path.' },
        { name: 'flag', type: 'boolean', description: 'A boolean flag.' }
    ],
    run() {}
};

const OPTIONAL_COMMAND = {
    name: 'opt',
    options: [{ name: 'out', type: 'string', description: 'Output path.' }],
    run() {}
};

test('parseArgs reads --key value and boolean --flag', () => {
    assert.deepEqual(parseArgs(['--input', 'x', '--flag']), { input: 'x', flag: true });
});

test('parseArgs treats a flag followed by another flag as boolean', () => {
    assert.deepEqual(parseArgs(['--a', '--b', 'val']), { a: true, b: 'val' });
});

test('validateOptions accepts a valid argument set', () => {
    assert.equal(validateOptions(DEMO_COMMAND, { input: 'x', flag: true }), null);
});

test('validateOptions rejects unknown options', () => {
    assert.equal(
        validateOptions(DEMO_COMMAND, { input: 'x', bogus: 1 }),
        'Unknown option: --bogus'
    );
});

test('validateOptions reports a missing required option', () => {
    assert.equal(validateOptions(DEMO_COMMAND, { flag: true }), 'Missing required option: --input');
});

test('validateOptions requires a value for string options', () => {
    assert.equal(validateOptions(OPTIONAL_COMMAND, { out: true }), 'Option --out requires a value');
    assert.equal(validateOptions(OPTIONAL_COMMAND, { out: 'y' }), null);
});

test('validateOptions allows global options', () => {
    assert.equal(validateOptions(OPTIONAL_COMMAND, { help: true, verbose: true }), null);
});

test('formatHelp renders usage and option flags', () => {
    const help = formatHelp(DEMO_COMMAND, 'node scripts/cli.js');
    assert.match(help, /Usage: node scripts\/cli\.js demo \[options\]/);
    assert.match(help, /--input <value>/);
    assert.match(help, /--flag\b/);
    assert.match(help, /--help\b/);
});
