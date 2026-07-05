'use strict';

/**
 * Shared CLI toolkit for the Node pipeline entry points.
 *
 * Provides:
 * - `parseArgs`  : lenient `--key value` / `--flag` parser.
 * - `runMain`    : consistent error handling and exit codes.
 * - `runCommand` : option validation + auto-generated `--help` for a command
 *                  descriptor, then delegates to the command's `run(args)`.
 *
 * A command descriptor has the shape:
 *   {
 *     name: 'normalize',
 *     summary: 'Normalize raw inputs into a snapshot.',
 *     usage: 'normalize [options]',
 *     options: [
 *       { name: 'raw-dir', type: 'string', description: '...', default: '...' },
 *       { name: 'fixture', type: 'boolean', description: '...' }
 *     ],
 *     run(args) { ... }
 *   }
 */

const GLOBAL_OPTIONS = [
    { name: 'help', type: 'boolean', description: 'Show this help message and exit.' },
    { name: 'verbose', type: 'boolean', description: 'Enable debug-level logging.' },
    { name: 'quiet', type: 'boolean', description: 'Only log warnings and errors.' }
];

/**
 * Parse `--key value` and boolean `--flag` style arguments into an object.
 *
 * @param {string[]} argv - Arguments, typically `process.argv.slice(2)`.
 * @returns {Record<string, string|boolean>} Parsed arguments.
 */
function parseArgs(argv) {
    const args = {};
    for (let index = 0; index < argv.length; index += 1) {
        const part = argv[index];
        if (!part.startsWith('--')) {
            continue;
        }

        const key = part.slice(2);
        const next = argv[index + 1];
        if (!next || next.startsWith('--')) {
            args[key] = true;
        } else {
            args[key] = next;
            index += 1;
        }
    }
    return args;
}

/**
 * Run a CLI entry function with consistent error handling. Any thrown error or
 * rejected promise is reported as a single clean line and sets a non-zero exit
 * code, instead of leaking a raw stack trace to the user.
 *
 * @param {() => void | Promise<void>} main - The entry function to execute.
 */
function runMain(main) {
    Promise.resolve()
        .then(main)
        .catch((error) => {
            const message = error && error.message ? error.message : String(error);
            console.error(`error: ${message}`);
            process.exitCode = 1;
        });
}

/**
 * Render an aligned help screen for a command descriptor.
 *
 * @param {object} command - The command descriptor.
 * @param {string} invocation - How the command is invoked, shown in the usage line.
 * @returns {string}
 */
function formatHelp(command, invocation) {
    const options = [...(command.options || []), ...GLOBAL_OPTIONS];
    const usage = command.usage || `${command.name} [options]`;
    const lines = [];

    if (command.summary) {
        lines.push(command.summary, '');
    }
    lines.push(`Usage: ${invocation} ${usage}`, '');

    if (options.length > 0) {
        lines.push('Options:');
        const rendered = options.map((option) => {
            const valueHint = option.type === 'boolean' ? '' : ' <value>';
            return {
                flag: `  --${option.name}${valueHint}`,
                description: option.description || ''
            };
        });
        const width = Math.max(...rendered.map((item) => item.flag.length));
        rendered.forEach((item) => {
            const suffix = item.description ? `  ${item.description}` : '';
            lines.push(`${item.flag.padEnd(width)}${suffix}`);
        });
    }

    return lines.join('\n');
}

/**
 * Validate parsed args against a command's option spec.
 *
 * @param {object} command - The command descriptor.
 * @param {Record<string, string|boolean>} args - Parsed arguments.
 * @returns {string|null} An error message, or null when valid.
 */
function validateOptions(command, args) {
    const spec = [...(command.options || []), ...GLOBAL_OPTIONS];
    const known = new Set(spec.map((option) => option.name));

    for (const key of Object.keys(args)) {
        if (!known.has(key)) {
            return `Unknown option: --${key}`;
        }
    }

    for (const option of command.options || []) {
        if (option.required && (args[option.name] === undefined || args[option.name] === true)) {
            return `Missing required option: --${option.name}`;
        }
        if (option.type === 'string' && args[option.name] === true) {
            return `Option --${option.name} requires a value`;
        }
    }

    return null;
}

/**
 * Parse, validate and dispatch a command descriptor.
 *
 * @param {object} command - The command descriptor.
 * @param {string[]} [argv] - Raw arguments (defaults to process argv).
 * @param {object} [context] - Extra context, e.g. `{ invocation }` for help text.
 */
function runCommand(command, argv = process.argv.slice(2), context = {}) {
    const invocation = context.invocation || 'node scripts/cli.js';
    const args = parseArgs(argv);

    if (args.help || args.h) {
        console.log(formatHelp(command, invocation));
        return;
    }

    const validationError = validateOptions(command, args);
    if (validationError) {
        console.error(`error: ${validationError}`);
        console.error('');
        console.error(formatHelp(command, invocation));
        process.exitCode = 1;
        return;
    }

    runMain(() => command.run(args));
}

module.exports = {
    parseArgs,
    runMain,
    runCommand,
    formatHelp,
    validateOptions,
    GLOBAL_OPTIONS
};
