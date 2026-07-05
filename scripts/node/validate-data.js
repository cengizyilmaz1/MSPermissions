const path = require('path');

const thresholds = require('../../src/config/validation-thresholds');
const { validateNormalizedData } = require('../../src/lib/site-data');
const { loadJson, writeJson } = require('../../src/lib/utils');
const { runCommand } = require('./lib/cli');
const { createLogger } = require('./lib/logger');

const ROOT_DIR = path.join(__dirname, '..', '..');
const DEFAULT_INPUT = path.join(ROOT_DIR, '.generated', 'normalized', 'site-data.json');

const command = {
    name: 'validate',
    summary: 'Validate a normalized snapshot against freshness and integrity thresholds.',
    usage: 'validate [options]',
    options: [
        {
            name: 'input',
            type: 'string',
            description: 'Path to the normalized site-data.json.',
            default: '.generated/normalized/site-data.json'
        },
        {
            name: 'summary',
            type: 'string',
            description: 'Optional path to write a JSON validation summary.'
        },
        {
            name: 'fixture',
            type: 'boolean',
            description: 'Relax thresholds for the small fixture dataset.'
        }
    ],
    run(args) {
        const log = createLogger({ scope: 'validate' });
        const input = path.resolve(args.input || DEFAULT_INPUT);
        const summaryPath = args.summary ? path.resolve(args.summary) : null;
        const normalized = loadJson(input);

        if (!normalized) {
            throw new Error(`Normalized snapshot not found: ${input}`);
        }

        const result = validateNormalizedData(normalized, thresholds, {
            fixtureMode: Boolean(args.fixture)
        });

        if (summaryPath) {
            writeJson(summaryPath, result);
        }

        result.warnings.forEach((warning) => log.warn(warning));
        result.errors.forEach((error) => log.error(error));
        log.info(
            `validated: permissions=${result.metrics.permissions} apps=${result.metrics.apps} categories=${result.metrics.categories}`
        );

        if (!result.valid) {
            process.exitCode = 1;
        }
        return result;
    }
};

if (require.main === module) {
    runCommand(command);
}

module.exports = { command, runCli: command.run };
