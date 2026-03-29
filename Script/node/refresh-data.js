const path = require('path');
const { spawnSync } = require('child_process');

const { generatePermissionDocsFromGraphDocs } = require('./lib/graph-docs-parser');
const { normalizeRawData, writeNormalizedData } = require('../../src/lib/site-data');
const { cleanDir } = require('../../src/lib/utils');

const ROOT_DIR = path.join(__dirname, '..', '..');
const DEFAULT_RAW_DIR = path.join(ROOT_DIR, '.generated', 'raw');
const DEFAULT_NORMALIZED_DIR = path.join(ROOT_DIR, '.generated', 'normalized');
const DEFAULT_CUSTOM_APP_DATA = path.join(ROOT_DIR, 'customdata', 'OtherMicrosoftApps.csv');

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

function findPowerShell() {
    const candidates = process.platform === 'win32' ? ['pwsh', 'powershell'] : ['pwsh'];
    for (const candidate of candidates) {
        const result = spawnSync(candidate, ['-NoLogo', '-NoProfile', '-Command', '$PSVersionTable.PSVersion.ToString()'], {
            stdio: 'pipe',
            encoding: 'utf8'
        });

        if (result.status === 0) {
            return candidate;
        }
    }

    throw new Error('PowerShell was not found. Install pwsh or Windows PowerShell.');
}

function tryGetGraphAccessToken() {
    const command = process.platform === 'win32' ? 'az.cmd' : 'az';
    const result = spawnSync(command, ['account', 'get-access-token', '--resource-type', 'ms-graph', '--output', 'json'], {
        stdio: 'pipe',
        encoding: 'utf8'
    });

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

function runCli() {
    const args = parseArgs(process.argv.slice(2));
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

        runPowerShellFile(shell, path.join(ROOT_DIR, 'Script', 'powershell', 'Export-GraphPermissions.ps1'), [
            '-OutputPath',
            rawDir,
            ...tokenArgs
        ]);
        runPowerShellFile(shell, path.join(ROOT_DIR, 'Script', 'powershell', 'Export-MicrosoftApps.ps1'), [
            '-OutputPath',
            rawDir,
            '-CustomAppDataPath',
            customAppData,
            ...tokenArgs
        ]);
        runPowerShellFile(shell, path.join(ROOT_DIR, 'Script', 'powershell', 'Parse-GraphOpenAPI.ps1'), [
            '-OutputPath',
            rawDir
        ]);
        runPowerShellFile(shell, path.join(ROOT_DIR, 'Script', 'powershell', 'Parse-GraphOpenAPIProperties.ps1'), [
            '-OutputPath',
            rawDir
        ]);
        const permissionDocsResult = generatePermissionDocsFromGraphDocs(rawDir);
        console.log(`generated learn api methods: permissions=${permissionDocsResult.api.permissions} mappings=${permissionDocsResult.api.mappings}`);
        console.log(`generated learn powershell methods: permissions=${permissionDocsResult.powershell.permissions} mappings=${permissionDocsResult.powershell.mappings}`);
        console.log(`generated learn code examples: permissions=${permissionDocsResult.codeExamples.permissions} snippets=${permissionDocsResult.codeExamples.snippets}`);
        console.log(`generated learn resource docs: resources=${permissionDocsResult.resources.resources} properties=${permissionDocsResult.resources.propertyTables} relationships=${permissionDocsResult.resources.relationshipTables} json=${permissionDocsResult.resources.jsonRepresentations}`);
    }

    const normalized = normalizeRawData(rawDir, { ingestedAt });
    writeNormalizedData(normalized, normalizedDir);

    console.log(`refreshed snapshot=${normalized.snapshotId} permissions=${normalized.stats.permissions} apps=${normalized.stats.apps}`);
}

if (require.main === module) {
    runCli();
}

module.exports = { runCli };
