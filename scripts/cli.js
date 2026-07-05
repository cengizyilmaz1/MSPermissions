#!/usr/bin/env node
'use strict';

/**
 * Unified entry point for the data pipeline.
 *
 *   node scripts/cli.js <command> [options]
 *
 * Each subcommand is a thin dispatch to the matching module under
 * `scripts/node/`. The individual modules remain runnable on their own for
 * backward compatibility, but this single entry point gives one discoverable
 * surface with a shared `--help`.
 */

const { command: refresh } = require('./node/refresh-data');
const { command: refreshDocs } = require('./node/refresh-permission-docs');
const { command: normalize } = require('./node/normalize-data');
const { command: validate } = require('./node/validate-data');
const { command: build } = require('./node/build-site');
const { runCommand } = require('./node/lib/cli');

const COMMANDS = [refresh, refreshDocs, normalize, validate, build];
const REGISTRY = new Map(COMMANDS.map((command) => [command.name, command]));

function printRootHelp() {
    const width = Math.max(...COMMANDS.map((command) => command.name.length));
    const lines = [
        'Graph Permissions Explorer — data pipeline CLI',
        '',
        'Usage: node scripts/cli.js <command> [options]',
        '',
        'Commands:'
    ];
    COMMANDS.forEach((command) => {
        lines.push(`  ${command.name.padEnd(width)}  ${command.summary || ''}`);
    });
    lines.push(
        '',
        'Run "node scripts/cli.js <command> --help" for command-specific options.',
        '',
        'Global options:',
        '  --help     Show help for the CLI or a command.',
        '  --verbose  Enable debug-level logging.',
        '  --quiet    Only log warnings and errors.'
    );
    console.log(lines.join('\n'));
}

function main() {
    const argv = process.argv.slice(2);
    const [commandName, ...rest] = argv;

    if (!commandName || commandName === '--help' || commandName === '-h') {
        printRootHelp();
        return;
    }

    const command = REGISTRY.get(commandName);
    if (!command) {
        console.error(`error: Unknown command "${commandName}"`);
        console.error('');
        printRootHelp();
        process.exitCode = 1;
        return;
    }

    runCommand(command, rest, { invocation: 'node scripts/cli.js' });
}

main();
