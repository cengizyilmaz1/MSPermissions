const path = require('path');

const { normalizeRawData, writeNormalizedData } = require('../../src/lib/site-data');

const ROOT_DIR = path.join(__dirname, '..', '..');
const DEFAULT_RAW_DIR = path.join(ROOT_DIR, '.generated', 'raw');
const DEFAULT_OUTPUT_DIR = path.join(ROOT_DIR, '.generated', 'normalized');

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
    const rawDir = path.resolve(args['raw-dir'] || DEFAULT_RAW_DIR);
    const outputDir = path.resolve(args.output || DEFAULT_OUTPUT_DIR);
    const ingestedAt = args['ingested-at'] || new Date().toISOString();

    const normalized = normalizeRawData(rawDir, { ingestedAt });
    writeNormalizedData(normalized, outputDir);

    console.log(`Normalized raw data from ${rawDir}`);
    console.log(`Snapshot ID: ${normalized.snapshotId}`);
    console.log(`Permissions: ${normalized.stats.permissions}`);
    console.log(`Apps: ${normalized.stats.apps}`);
}

if (require.main === module) {
    runCli();
}

module.exports = { runCli };
