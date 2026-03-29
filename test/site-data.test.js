const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const test = require('node:test');

const { buildSite } = require('../Script/node/build-site');
const { normalizeRawData, validateNormalizedData, writeNormalizedData } = require('../src/lib/site-data');
const thresholds = require('../src/config/validation-thresholds');

const fixturesDir = path.join(__dirname, '..', 'fixtures', 'raw');

test('normalizeRawData builds deterministic permission and app data', () => {
    const normalized = normalizeRawData(fixturesDir, {
        ingestedAt: '2026-01-01T00:00:00Z'
    });

    assert.equal(normalized.permissions.length, 2);
    assert.equal(normalized.apps.length, 4);
    assert.equal(normalized.snapshotId.length, 16);

    const userReadAll = normalized.permissions.find((item) => item.value === 'User.Read.All');
    assert.ok(userReadAll);
    assert.equal(userReadAll.methods.confidence, 'exact');
    assert.equal(userReadAll.methods.powershell.confidence, 'exact');
    assert.equal(userReadAll.methods.powershell.v1[0].command, 'Get-MgUser');
    assert.match(userReadAll.codeExamples.csharp, /View official example on Microsoft Learn/);
    assert.match(userReadAll.codeExamples.javascript, /graphClient\.users/);
    assert.match(userReadAll.codeExamples.python, /graph_client\.users/);
    assert.equal(userReadAll.properties.resourceName.toLowerCase(), 'user');
    assert.equal(userReadAll.properties.version, 'v1');
    assert.equal(userReadAll.properties.source, 'learn+openapi');
    assert.equal(userReadAll.jsonRepresentation.source, 'learn');
    assert.equal(userReadAll.jsonRepresentation.value.displayName, 'Adele Vance');
    assert.equal(userReadAll.relationships.items[0].name, 'manager');

    const communityApp = normalized.apps.find((item) => item.isCommunity);
    assert.ok(communityApp);
    assert.equal(communityApp.title, 'Mystery Microsoft App');
    assert.equal(communityApp.filterGroup, 'community');
});

test('fixture validation passes and build writes public JSON contracts', () => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'permissions-fixture-'));
    const normalizedDir = path.join(tempRoot, 'normalized');
    const outputDir = path.join(tempRoot, 'site');
    const normalized = normalizeRawData(fixturesDir, {
        ingestedAt: '2026-01-01T00:00:00Z'
    });

    writeNormalizedData(normalized, normalizedDir);

    const validation = validateNormalizedData(normalized, thresholds, { fixtureMode: true });
    assert.equal(validation.valid, true);

    buildSite(path.join(normalizedDir, 'site-data.json'), outputDir);

    assert.ok(fs.existsSync(path.join(outputDir, 'index.html')));
    assert.ok(fs.existsSync(path.join(outputDir, 'microsoft-apps.html')));
    assert.ok(fs.existsSync(path.join(outputDir, 'llms.txt')));
    assert.ok(fs.existsSync(path.join(outputDir, 'llms-full.txt')));
    assert.ok(fs.existsSync(path.join(outputDir, 'data', 'catalog', 'permissions.json')));
    assert.ok(fs.existsSync(path.join(outputDir, 'data', 'catalog', 'apps-manifest.json')));
    assert.ok(fs.existsSync(path.join(outputDir, 'data', 'permissions', 'user-read-all.json')));
    assert.ok(!fs.existsSync(path.join(outputDir, 'js', 'search-index.js')));

    const appsPage = fs.readFileSync(path.join(outputDir, 'microsoft-apps.html'), 'utf8');
    assert.match(appsPage, /skeleton-row/);
    assert.match(appsPage, /build-info\.json/);
    assert.match(appsPage, /Catalog structure/);

    const appsManifest = JSON.parse(fs.readFileSync(path.join(outputDir, 'data', 'catalog', 'apps-manifest.json'), 'utf8'));
    assert.ok(!Object.hasOwn(appsManifest, 'pages'));
    assert.equal(appsManifest.searchIndex[0].length, 3);

    const mysteryApp = normalized.apps.find((item) => item.title === 'Mystery Microsoft App');
    assert.ok(mysteryApp);
    const appDetailPage = fs.readFileSync(path.join(outputDir, 'apps', `${mysteryApp.anchor}.html`), 'utf8');
    assert.match(appDetailPage, /Mystery Microsoft App/);
    assert.match(appDetailPage, /Community maintained/);

    const permissionPage = fs.readFileSync(path.join(outputDir, 'permissions', 'user-read-all.html'), 'utf8');
    assert.match(permissionPage, /Exact Microsoft Learn match/);
    assert.match(permissionPage, /Get-MgUser/);
    assert.match(permissionPage, /View official example on Microsoft Learn/);
    assert.match(permissionPage, /Adele Vance/);
    assert.match(permissionPage, /manager/);
});

test('normalizeRawData accepts sharded code example files', () => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'permissions-sharded-'));
    const rawDir = path.join(tempRoot, 'raw');
    fs.mkdirSync(rawDir, { recursive: true });

    for (const entry of fs.readdirSync(fixturesDir)) {
        if (entry === 'GraphPermissionCodeExamples.json') {
            continue;
        }

        fs.copyFileSync(path.join(fixturesDir, entry), path.join(rawDir, entry));
    }

    const codeExamples = JSON.parse(fs.readFileSync(path.join(fixturesDir, 'GraphPermissionCodeExamples.json'), 'utf8'));
    const entries = Object.entries(codeExamples);
    fs.writeFileSync(
        path.join(rawDir, 'GraphPermissionCodeExamples.part-001.json'),
        `${JSON.stringify(Object.fromEntries(entries.slice(0, 1)), null, 2)}\n`
    );
    fs.writeFileSync(
        path.join(rawDir, 'GraphPermissionCodeExamples.part-002.json'),
        `${JSON.stringify(Object.fromEntries(entries.slice(1)), null, 2)}\n`
    );

    const normalized = normalizeRawData(rawDir, {
        ingestedAt: '2026-01-01T00:00:00Z'
    });

    const userReadAll = normalized.permissions.find((item) => item.value === 'User.Read.All');
    assert.ok(userReadAll);
    assert.match(userReadAll.codeExamples.csharp, /View official example on Microsoft Learn/);
});
