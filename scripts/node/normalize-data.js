const path = require('path');

const { normalizeRawData, writeNormalizedData } = require('../../src/lib/site-data');
const { runCommand } = require('./lib/cli');
const { createLogger } = require('./lib/logger');

const ROOT_DIR = path.join(__dirname, '..', '..');
const DEFAULT_RAW_DIR = path.join(ROOT_DIR, '.generated', 'raw');
const DEFAULT_OUTPUT_DIR = path.join(ROOT_DIR, '.generated', 'normalized');

const command = {
    name: 'normalize',
    summary: 'Normalize raw inputs into a deterministic site-data snapshot.',
    usage: 'normalize [options]',
    options: [
        {
            name: 'raw-dir',
            type: 'string',
            description: 'Directory containing raw inputs.',
            default: '.generated/raw'
        },
        {
            name: 'output',
            type: 'string',
            description: 'Output directory for the normalized snapshot.',
            default: '.generated/normalized'
        },
        {
            name: 'ingested-at',
            type: 'string',
            description: 'Override the ingestion timestamp (ISO 8601).'
        }
    ],
    run(args) {
        const log = createLogger({ scope: 'normalize' });
        const rawDir = path.resolve(args['raw-dir'] || DEFAULT_RAW_DIR);
        const outputDir = path.resolve(args.output || DEFAULT_OUTPUT_DIR);
        const ingestedAt = args['ingested-at'] || new Date().toISOString();

        const normalized = normalizeRawData(rawDir, { ingestedAt });
        writeNormalizedData(normalized, outputDir);

        log.info(`Normalized raw data from ${rawDir}`);
        log.info(`Snapshot ID: ${normalized.snapshotId}`);
        log.info(`Permissions: ${normalized.stats.permissions}`);
        log.info(`Apps: ${normalized.stats.apps}`);
        return normalized;
    }
};

if (require.main === module) {
    runCommand(command);
}

module.exports = { command, runCli: command.run };
