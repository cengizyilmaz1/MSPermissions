const path = require('path');
const { spawnSync } = require('child_process');

const { generatePermissionDocsFromGraphDocs } = require('./lib/graph-docs-parser');
const { buildPermissionCsv } = require('./lib/openapi-permissions');
const { normalizeRawData, writeNormalizedData } = require('../../src/lib/site-data');
const { cleanDir } = require('../../src/lib/utils');
const { runCommand } = require('./lib/cli');
const { createLogger } = require('./lib/logger');

const ROOT_DIR = path.join(__dirname, '..', '..');
const DEFAULT_RAW_DIR = path.join(ROOT_DIR, '.generated', 'raw');
const DEFAULT_NORMALIZED_DIR = path.join(ROOT_DIR, '.generated', 'normalized');
const DEFAULT_CUSTOM_APP_DATA = path.join(ROOT_DIR, 'customdata', 'OtherMicrosoftApps.csv');

function findPowerShell() {
    const candidates = process.platform === 'win32' ? ['pwsh', 'powershell'] : ['pwsh'];
    for (const candidate of candidates) {
        const result = spawnSync(
            candidate,
            ['-NoLogo', '-NoProfile', '-Command', '$PSVersionTable.PSVersion.ToString()'],
            {
                stdio: 'pipe',
                encoding: 'utf8'
            }
        );

        if (result.status === 0) {
            return candidate;
        }
    }

    throw new Error('PowerShell was not found. Install pwsh or Windows PowerShell.');
}

function tryGetGraphAccessToken() {
    const command = process.platform === 'win32' ? 'az.cmd' : 'az';
    const result = spawnSync(
        command,
        ['account', 'get-access-token', '--resource-type', 'ms-graph', '--output', 'json'],
        {
            stdio: 'pipe',
            encoding: 'utf8'
        }
    );

    if (result.status !== 0) {
        return null;
    }

    try {
        const parsed = JSON.parse(result.stdout);
        return parsed.accessToken || null;
    } catch {
        return null;
    }
}

function runPowerShellFile(shell, filePath, args) {
    const result = spawnSync(shell, ['-NoLogo', '-NoProfile', '-File', filePath, ...args], {
        cwd: ROOT_DIR,
        stdio: 'inherit',
        encoding: 'utf8'
    });

    if (result.status !== 0) {
        throw new Error(`PowerShell script failed: ${path.basename(filePath)}`);
    }
}

const command = {
    name: 'refresh',
    summary: 'Fetch upstream data, parse docs, and write a normalized snapshot.',
    usage: 'refresh [options]',
    options: [
        {
            name: 'raw-dir',
            type: 'string',
            description: 'Directory to write fetched raw inputs.',
            default: '.generated/raw'
        },
        {
            name: 'normalized-dir',
            type: 'string',
            description: 'Directory to write the normalized snapshot.',
            default: '.generated/normalized'
        },
        {
            name: 'custom-app-data',
            type: 'string',
            description: 'CSV of community-contributed app IDs.',
            default: 'customdata/OtherMicrosoftApps.csv'
        },
        {
            name: 'skip-fetch',
            type: 'boolean',
            description: 'Skip upstream fetch and only normalize existing raw inputs.'
        }
    ],
    async run(args) {
        const log = createLogger({ scope: 'refresh' });
        const rawDir = path.resolve(args['raw-dir'] || DEFAULT_RAW_DIR);
        const normalizedDir = path.resolve(args['normalized-dir'] || DEFAULT_NORMALIZED_DIR);
        const customAppData = path.resolve(args['custom-app-data'] || DEFAULT_CUSTOM_APP_DATA);
        const skipFetch = Boolean(args['skip-fetch']);
        const ingestedAt = new Date().toISOString();

        if (!skipFetch) {
            cleanDir(rawDir);
            const shell = findPowerShell();
            const accessToken = tryGetGraphAccessToken();
            const tokenArgs = accessToken ? ['-AccessToken', accessToken] : [];

            runPowerShellFile(
                shell,
                path.join(ROOT_DIR, 'scripts', 'powershell', 'Export-GraphPermissions.ps1'),
                ['-OutputPath', rawDir, ...tokenArgs]
            );
            runPowerShellFile(
                shell,
                path.join(ROOT_DIR, 'scripts', 'powershell', 'Export-MicrosoftApps.ps1'),
                ['-OutputPath', rawDir, '-CustomAppDataPath', customAppData, ...tokenArgs]
            );
            const openApiResult = await buildPermissionCsv(rawDir, { logger: log });
            log.info(
                `generated openapi permission.csv: records=${openApiResult.records} permissions=${openApiResult.permissions} v1=${openApiResult.v1} beta=${openApiResult.beta}`
            );
            runPowerShellFile(
                shell,
                path.join(ROOT_DIR, 'scripts', 'powershell', 'Parse-GraphOpenAPIProperties.ps1'),
                ['-OutputPath', rawDir]
            );
            const permissionDocsResult = generatePermissionDocsFromGraphDocs(rawDir);
            log.info(
                `generated learn api methods: permissions=${permissionDocsResult.api.permissions} mappings=${permissionDocsResult.api.mappings}`
            );
            log.info(
                `generated learn powershell methods: permissions=${permissionDocsResult.powershell.permissions} mappings=${permissionDocsResult.powershell.mappings}`
            );
            log.info(
                `generated learn code examples: permissions=${permissionDocsResult.codeExamples.permissions} snippets=${permissionDocsResult.codeExamples.snippets}`
            );
            log.info(
                `generated learn resource docs: resources=${permissionDocsResult.resources.resources} properties=${permissionDocsResult.resources.propertyTables} relationships=${permissionDocsResult.resources.relationshipTables} json=${permissionDocsResult.resources.jsonRepresentations}`
            );
        }

        const normalized = normalizeRawData(rawDir, { ingestedAt });
        writeNormalizedData(normalized, normalizedDir);

        log.info(
            `refreshed snapshot=${normalized.snapshotId} permissions=${normalized.stats.permissions} apps=${normalized.stats.apps}`
        );
        return normalized;
    }
};

if (require.main === module) {
    runCommand(command);
}

module.exports = { command, runCli: command.run };
