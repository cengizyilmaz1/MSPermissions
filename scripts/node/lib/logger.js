'use strict';

/**
 * Minimal, dependency-free structured logger for the Node pipeline.
 *
 * Features:
 * - Ordered levels: silent < error < warn < info < debug.
 * - ISO timestamp and level tag on every line for readable CI logs.
 * - Level resolved from (in priority order): explicit `createLogger` option,
 *   `--quiet` / `--verbose` CLI flags, `LOG_LEVEL` env var, default `info`.
 * - Optional ANSI colors, enabled only on a TTY and disabled with `NO_COLOR`.
 * - Optional scope label so each entry point can tag its output.
 */

const LEVELS = { silent: 0, error: 1, warn: 2, info: 3, debug: 4 };

const COLORS = {
    error: '\u001b[31m', // red
    warn: '\u001b[33m', // yellow
    info: '\u001b[36m', // cyan
    debug: '\u001b[90m', // gray
    success: '\u001b[32m', // green
    reset: '\u001b[0m'
};

function colorsEnabled() {
    if (process.env.NO_COLOR) {
        return false;
    }
    return Boolean(process.stdout && process.stdout.isTTY);
}

/**
 * Resolve the active log level from environment and process arguments.
 *
 * @param {string} [explicit] - An explicit level that overrides everything else.
 * @returns {keyof typeof LEVELS}
 */
function resolveLevel(explicit) {
    if (explicit && explicit in LEVELS) {
        return explicit;
    }

    const argv = process.argv.slice(2);
    if (argv.includes('--quiet')) {
        return 'warn';
    }
    if (argv.includes('--verbose') || argv.includes('--debug')) {
        return 'debug';
    }

    const fromEnv = (process.env.LOG_LEVEL || '').toLowerCase();
    if (fromEnv in LEVELS) {
        return fromEnv;
    }

    return 'info';
}

function paint(level, text) {
    if (!colorsEnabled()) {
        return text;
    }
    const color = COLORS[level] || '';
    return `${color}${text}${COLORS.reset}`;
}

/**
 * Create a logger instance.
 *
 * @param {object} [options]
 * @param {string} [options.scope] - Label prepended to each message.
 * @param {keyof typeof LEVELS} [options.level] - Explicit level override.
 * @returns {{
 *   error: (...args: unknown[]) => void,
 *   warn: (...args: unknown[]) => void,
 *   info: (...args: unknown[]) => void,
 *   success: (...args: unknown[]) => void,
 *   debug: (...args: unknown[]) => void,
 *   level: keyof typeof LEVELS
 * }}
 */
function createLogger(options = {}) {
    const activeLevel = resolveLevel(options.level);
    const threshold = LEVELS[activeLevel];
    const scope = options.scope ? `[${options.scope}] ` : '';

    const emit = (level, stream, args) => {
        const gate = level === 'success' ? LEVELS.info : LEVELS[level];
        if (gate > threshold) {
            return;
        }

        const timestamp = new Date().toISOString();
        const tag = paint(level, (level === 'success' ? 'INFO' : level.toUpperCase()).padEnd(5));
        const message = args
            .map((value) => (typeof value === 'string' ? value : JSON.stringify(value)))
            .join(' ');
        stream(`${timestamp} ${tag} ${scope}${message}`);
    };

    return {
        level: activeLevel,
        error: (...args) => emit('error', console.error, args),
        warn: (...args) => emit('warn', console.warn, args),
        info: (...args) => emit('info', console.log, args),
        success: (...args) => emit('success', console.log, args),
        debug: (...args) => emit('debug', console.log, args)
    };
}

module.exports = { createLogger, LEVELS };
