const path = require('path');

const thresholds = require('../../src/config/validation-thresholds');
const { validateNormalizedData } = require('../../src/lib/site-data');
const { loadJson, writeJson } = require('../../src/lib/utils');

const ROOT_DIR = path.join(__dirname, '..', '..');
const DEFAULT_INPUT = path.join(ROOT_DIR, '.generated', 'normalized', 'site-data.json');

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

function runCli() {
    const args = parseArgs(process.argv.slice(2));
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

    result.warnings.forEach((warning) => console.warn(`warning: ${warning}`));
    result.errors.forEach((error) => console.error(`error: ${error}`));
    console.log(`validated: permissions=${result.metrics.permissions} apps=${result.metrics.apps} categories=${result.metrics.categories}`);

    if (!result.valid) {
        process.exitCode = 1;
    }
}

if (require.main === module) {
    runCli();
}

module.exports = { runCli };
