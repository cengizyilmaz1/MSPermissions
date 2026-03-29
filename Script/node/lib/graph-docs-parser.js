const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawnSync } = require('child_process');

const { ensureDir, readUtf8, writeJson, writeJsonSharded } = require('../../../src/lib/utils');

const GRAPH_DOCS_REPO_URL = 'https://github.com/microsoftgraph/microsoft-graph-docs-contrib.git';
const VERSION_SPECS = [
    {
        key: 'v1',
        docsDir: path.join('api-reference', 'v1.0', 'api'),
        resourceDir: path.join('api-reference', 'v1.0', 'resources'),
        includeDir: path.join('api-reference', 'v1.0', 'includes', 'permissions'),
        includeRoot: path.join('api-reference', 'v1.0', 'includes'),
        learnView: 'graph-rest-1.0'
    },
    {
        key: 'beta',
        docsDir: path.join('api-reference', 'beta', 'api'),
        resourceDir: path.join('api-reference', 'beta', 'resources'),
        includeDir: path.join('api-reference', 'beta', 'includes', 'permissions'),
        includeRoot: path.join('api-reference', 'beta', 'includes'),
        learnView: 'graph-rest-beta'
    }
];
const CODE_LANGUAGES = ['csharp', 'javascript', 'powershell', 'python'];

const METHOD_ORDER = {
    GET: 1,
    POST: 2,
    PATCH: 3,
    PUT: 4,
    DELETE: 5,
    HEAD: 6,
    OPTIONS: 7
};

function runGit(args, options = {}) {
    const result = spawnSync('git', args, {
        cwd: options.cwd,
        stdio: 'pipe',
        encoding: 'utf8'
    });

    if (result.status !== 0) {
        const message = (result.stderr || result.stdout || '').trim();
        throw new Error(`git ${args.join(' ')} failed${message ? `: ${message}` : '.'}`);
    }

    return result.stdout.trim();
}

function cloneGraphDocs(repoDir) {
    ensureDir(path.dirname(repoDir));

    if (fs.existsSync(path.join(repoDir, '.git'))) {
        runGit(['fetch', '--depth', '1', 'origin', 'main'], { cwd: repoDir });
        runGit(['reset', '--hard', 'origin/main'], { cwd: repoDir });
        runGit([
            'sparse-checkout',
            'set',
            ...VERSION_SPECS.flatMap((spec) => [spec.docsDir, spec.resourceDir, spec.includeRoot])
        ], { cwd: repoDir });
        return;
    }

    runGit([
        '-c',
        'core.longpaths=true',
        'clone',
        '--depth',
        '1',
        '--filter=blob:none',
        '--sparse',
        '--branch',
        'main',
        GRAPH_DOCS_REPO_URL,
        repoDir
    ]);
    runGit([
        'sparse-checkout',
        'set',
        ...VERSION_SPECS.flatMap((spec) => [spec.docsDir, spec.resourceDir, spec.includeRoot])
    ], { cwd: repoDir });
}

function createShortTempRoot() {
    if (process.platform === 'win32') {
        const driveRoot = path.parse(os.tmpdir()).root || 'C:\\';
        return fs.mkdtempSync(path.join(driveRoot, 'gd-'));
    }

    return fs.mkdtempSync(path.join(os.tmpdir(), 'graph-docs-'));
}

function listMarkdownFiles(dirPath) {
    const results = [];

    if (!fs.existsSync(dirPath)) {
        return results;
    }

    for (const entry of fs.readdirSync(dirPath, { withFileTypes: true })) {
        const fullPath = path.join(dirPath, entry.name);
        if (entry.isDirectory()) {
            results.push(...listMarkdownFiles(fullPath));
            continue;
        }

        if (entry.isFile() && fullPath.toLowerCase().endsWith('.md')) {
            results.push(fullPath);
        }
    }

    return results.sort((left, right) => left.localeCompare(right));
}

function stripMarkdown(value) {
    return String(value || '')
        .replace(/<!--[\s\S]*?-->/g, ' ')
        .replace(/<br\s*\/?>/gi, ', ')
        .replace(/\[([^\]]+)\]\([^)]+\)/g, '$1')
        .replace(/`([^`]+)`/g, '$1')
        .replace(/[*_>#]/g, '')
        .replace(/&nbsp;/gi, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}

function normalizeMarkdownStructure(value) {
    if (!value) {
        return '';
    }

    let normalized = String(value)
        .replace(/\r\n/g, '\n')
        .replace(/\r/g, '\n');

    normalized = normalized
        .replace(/^\s*---\s*([\s\S]*?)\s*---\s*/m, (match) => {
            const compact = match.replace(/\n+/g, ' ').trim();
            return `${compact}\n`;
        })
        .replace(/(\S)\s+(#{1,6}\s+)/g, '$1\n$2')
        .replace(/(\S)\s+(\[!INCLUDE\s+\[[^\]]*\]\([^)]+\)\])/g, '$1\n$2')
        .replace(/(```[a-z0-9#+.-]*)(\s+)/gi, '$1\n')
        .replace(/```\s+(#{1,6}\s+)/g, '```\n$1')
        .replace(/\|\s+(?=\|[:A-Za-z0-9`<])/g, '|\n')
        .replace(/\n{3,}/g, '\n\n');

    return normalized;
}

function normalizeHeadingText(value) {
    return stripMarkdown(value)
        .toLowerCase()
        .replace(/[:`]/g, '')
        .replace(/\s+/g, ' ')
        .trim();
}

function splitTableRow(line) {
    return line
        .trim()
        .replace(/^\|/, '')
        .replace(/\|$/, '')
        .split('|')
        .map((cell) => cell.trim());
}

function extractSection(markdown, headingText) {
    const lines = normalizeMarkdownStructure(markdown).split(/\r?\n/);
    const target = normalizeHeadingText(headingText);
    let startIndex = -1;
    let startLevel = 0;

    for (let index = 0; index < lines.length; index += 1) {
        const headingMatch = lines[index].match(/^(#{2,6})\s+(.+)$/);
        if (!headingMatch) {
            continue;
        }

        const [, hashes, rawHeading] = headingMatch;
        const normalizedHeading = normalizeHeadingText(rawHeading);
        const matchesTarget = normalizedHeading === target
            || normalizedHeading.startsWith(`${target} `)
            || normalizedHeading.startsWith(`${target} -`)
            || normalizedHeading.startsWith(`${target}:`);

        if (matchesTarget) {
            startIndex = index + 1;
            startLevel = hashes.length;
            break;
        }
    }

    if (startIndex === -1) {
        return '';
    }

    const collected = [];
    for (let index = startIndex; index < lines.length; index += 1) {
        const headingMatch = lines[index].match(/^(#{1,6})\s+(.+)$/);
        if (headingMatch && headingMatch[1].length <= startLevel) {
            break;
        }

        collected.push(lines[index]);
    }

    return collected.join('\n').trim();
}

function expandPermissionIncludes(sectionText, markdownFile) {
    let expanded = normalizeMarkdownStructure(sectionText);
    const includePattern = /\[!INCLUDE\s+\[[^\]]*\]\(([^)]+)\)\]/g;

    for (let depth = 0; depth < 5; depth += 1) {
        let replaced = false;

        expanded = expanded.replace(includePattern, (_, relativePath) => {
            replaced = true;
            const includePath = path.resolve(path.dirname(markdownFile), relativePath);
            if (!fs.existsSync(includePath)) {
                return '';
            }
            return `\n${normalizeMarkdownStructure(readUtf8(includePath))}\n`;
        });

        if (!replaced) {
            break;
        }
    }

    return expanded;
}

function expandMarkdownIncludes(sectionText, markdownFile) {
    return expandPermissionIncludes(sectionText, markdownFile);
}

function extractMarkdownTables(sectionText) {
    const tables = [];
    const lines = normalizeMarkdownStructure(sectionText).split(/\r?\n/);
    let current = [];

    for (const line of lines) {
        if (line.trim().startsWith('|')) {
            current.push(line);
            continue;
        }

        if (current.length > 0) {
            tables.push(current);
            current = [];
        }
    }

    if (current.length > 0) {
        tables.push(current);
    }

    return tables;
}

function extractPermissionTokens(text) {
    const matches = stripMarkdown(text).match(/\b[A-Za-z][A-Za-z0-9-]*(?:\.[A-Za-z0-9-]+)+\b/g) || [];
    const seen = new Set();
    const results = [];

    for (const token of matches) {
        if (!/[A-Z]/.test(token) || seen.has(token)) {
            continue;
        }

        seen.add(token);
        results.push(token);
    }

    return results;
}

function parsePermissionTable(tableLines) {
    if (tableLines.length < 3) {
        return [];
    }

    const headers = splitTableRow(tableLines[0]).map((cell) => stripMarkdown(cell).toLowerCase());
    const permissionTypeIndex = headers.findIndex((cell) => cell.includes('permission type'));
    const orderedPermissionsIndex = headers.findIndex((cell) => cell.includes('permissions') && cell.includes('least to most privileged'));
    const leastPrivilegeIndex = headers.findIndex((cell) => cell.includes('least privileged'));
    const higherPrivilegeIndex = headers.findIndex((cell) => cell.includes('higher privileged'));

    if (permissionTypeIndex === -1 || (orderedPermissionsIndex === -1 && leastPrivilegeIndex === -1)) {
        return [];
    }

    const rows = [];

    for (let index = 2; index < tableLines.length; index += 1) {
        const cells = splitTableRow(tableLines[index]);
        if (cells.length <= permissionTypeIndex) {
            continue;
        }

        const permissionType = stripMarkdown(cells[permissionTypeIndex]);
        const lowerType = permissionType.toLowerCase();
        const supportsDelegated = lowerType.includes('delegated');
        const supportsApplication = lowerType.includes('application');

        if (!supportsDelegated && !supportsApplication) {
            continue;
        }

        if (orderedPermissionsIndex !== -1) {
            const ordered = extractPermissionTokens(cells[orderedPermissionsIndex] || '');
            ordered.forEach((permission, permissionIndex) => {
                rows.push({
                    permission,
                    supportsDelegated,
                    supportsApplication,
                    isLeastPrivilege: permissionIndex === 0
                });
            });
            continue;
        }

        const leastPermissions = extractPermissionTokens(cells[leastPrivilegeIndex] || '');
        const higherPermissions = higherPrivilegeIndex === -1 ? [] : extractPermissionTokens(cells[higherPrivilegeIndex] || '');

        leastPermissions.forEach((permission) => {
            rows.push({
                permission,
                supportsDelegated,
                supportsApplication,
                isLeastPrivilege: true
            });
        });

        higherPermissions.forEach((permission) => {
            rows.push({
                permission,
                supportsDelegated,
                supportsApplication,
                isLeastPrivilege: false
            });
        });
    }

    return rows;
}

function extractPermissionsFromMarkdown(markdown, markdownFile) {
    const permissionsSection = extractSection(markdown, 'Permissions');
    const sources = [];

    if (permissionsSection) {
        sources.push(expandPermissionIncludes(permissionsSection, markdownFile));
    }

    sources.push(expandPermissionIncludes(markdown, markdownFile));

    const seen = new Set();
    const results = [];

    sources.forEach((sourceText) => {
        extractMarkdownTables(sourceText).forEach((tableLines) => {
            parsePermissionTable(tableLines).forEach((item) => {
                const key = [
                    item.permission,
                    item.supportsDelegated ? 'd' : '-',
                    item.supportsApplication ? 'a' : '-',
                    item.isLeastPrivilege ? 'l' : 'h'
                ].join('|');

                if (!seen.has(key)) {
                    seen.add(key);
                    results.push(item);
                }
            });
        });
    });

    return results;
}

function normalizeEndpoint(rawPath) {
    let value = stripMarkdown(rawPath)
        .replace(/^https?:\/\/graph\.microsoft\.com\/(?:v1\.0|beta)/i, '')
        .replace(/^\/(?:v1\.0|beta)(?=\/)/i, '')
        .trim();

    if (!value) {
        return null;
    }

    if (!value.startsWith('/')) {
        value = `/${value}`;
    }

    return value;
}

function extractHttpRequests(markdown) {
    const httpSection = extractSection(markdown, 'HTTP request');
    const searchSpace = normalizeMarkdownStructure(httpSection || markdown);

    const results = [];
    const seen = new Set();
    const httpBlocks = extractCodeFences(searchSpace)
        .filter((entry) => entry.language === 'http')
        .map((entry) => entry.code);

    httpBlocks.forEach((block) => {
        for (const line of block.split(/\r?\n/)) {
            const methodMatch = line.trim().match(/^(GET|POST|PUT|PATCH|DELETE|HEAD|OPTIONS)\s+(.+)$/i);
            if (!methodMatch) {
                continue;
            }

            const method = methodMatch[1].toUpperCase();
            const endpoint = normalizeEndpoint(methodMatch[2]);
            if (!endpoint) {
                continue;
            }

            const key = `${method}|${endpoint}`;
            if (seen.has(key)) {
                continue;
            }

            seen.add(key);
            results.push({ method, path: endpoint });
        }
    });

    return results;
}

function extractPageTitle(markdown, markdownFile) {
    const frontMatterTitle = markdown.match(/^\s*---[\s\S]*?\btitle:\s*"([^"]+)"[\s\S]*?---/m);
    if (frontMatterTitle) {
        return stripMarkdown(frontMatterTitle[1]);
    }

    const heading = markdown.match(/^#\s+(.+)$/m);
    if (heading) {
        return stripMarkdown(heading[1]);
    }

    return path.basename(markdownFile, '.md');
}

function buildDocLink(markdownFile, versionSpec) {
    const slug = path.basename(markdownFile, '.md');
    return `https://learn.microsoft.com/en-us/graph/api/${slug}?view=${versionSpec.learnView}`;
}

function buildResourceDocLink(markdownFile, versionSpec) {
    const slug = path.basename(markdownFile, '.md');
    return `https://learn.microsoft.com/en-us/graph/api/resources/${slug}?view=${versionSpec.learnView}`;
}

function extractFrontMatterValue(markdown, key) {
    const frontMatterMatch = markdown.match(/^\s*---\r?\n([\s\S]*?)\r?\n---/);
    if (!frontMatterMatch) {
        return '';
    }

    const pattern = new RegExp(`^${key}:\\s*(.+)$`, 'im');
    const match = frontMatterMatch[1].match(pattern);
    if (!match) {
        return '';
    }

    return String(match[1] || '').replace(/^["']|["']$/g, '').trim();
}

function removeFrontMatter(markdown) {
    return markdown.replace(/^\s*---\r?\n[\s\S]*?\r?\n---\s*/m, '');
}

function extractIntroText(markdown) {
    const description = extractFrontMatterValue(markdown, 'description');
    if (description) {
        return stripMarkdown(description);
    }

    const content = removeFrontMatter(markdown);
    const lines = content.split(/\r?\n/);
    const collected = [];

    for (const line of lines) {
        if (/^#\s+/.test(line)) {
            continue;
        }

        if (/^#{2,6}\s+/.test(line)) {
            break;
        }

        const cleaned = stripMarkdown(line);
        if (!cleaned) {
            if (collected.length > 0) {
                break;
            }
            continue;
        }

        collected.push(cleaned);
    }

    return collected.join(' ').trim();
}

function extractSnippetIncludePaths(markdown, markdownFile, language) {
    const pattern = new RegExp(`\\[!INCLUDE\\s+\\[[^\\]]*\\]\\(([^)\\r\\n]*snippets[\\\\/]+${language}[\\\\/][^)\\r\\n]+\\.md)\\)\\]`, 'gi');
    const seen = new Set();
    const results = [];
    const normalized = normalizeMarkdownStructure(markdown);
    let match;

    while ((match = pattern.exec(normalized)) !== null) {
        const includePath = path.resolve(path.dirname(markdownFile), match[1]);
        if (!seen.has(includePath) && fs.existsSync(includePath)) {
            seen.add(includePath);
            results.push(includePath);
        }
    }

    return results;
}

function normalizeCodeFenceLanguage(language) {
    const value = String(language || '').toLowerCase().trim();

    if (['c#', 'csharp', 'cs', 'dotnet'].includes(value)) {
        return 'csharp';
    }

    if (['javascript', 'js', 'node', 'nodejs', 'typescript', 'ts'].includes(value)) {
        return 'javascript';
    }

    if (['powershell', 'pwsh', 'ps1'].includes(value)) {
        return 'powershell';
    }

    if (['python', 'py'].includes(value)) {
        return 'python';
    }

    if (['http', 'https'].includes(value)) {
        return 'http';
    }

    return value;
}

function extractCodeFences(markdown) {
    const results = [];
    const fencePattern = /```\s*([a-z0-9#+.-]*)\s*([\s\S]*?)```/gi;
    const normalized = normalizeMarkdownStructure(markdown);
    let match;

    while ((match = fencePattern.exec(normalized)) !== null) {
        results.push({
            language: normalizeCodeFenceLanguage(match[1]),
            code: match[2].trim()
        });
    }

    return results.filter((entry) => entry.code);
}

function extractFirstCodeFence(markdown, preferredLanguage = null) {
    const fences = extractCodeFences(markdown);
    if (preferredLanguage) {
        const preferred = fences.find((entry) => entry.language === normalizeCodeFenceLanguage(preferredLanguage));
        if (preferred) {
            return preferred.code;
        }
    }

    return fences[0]?.code || '';
}

function extractInlineCodeEntries(markdown, language) {
    const normalizedLanguage = normalizeCodeFenceLanguage(language);
    return extractCodeFences(markdown)
        .filter((entry) => entry.language === normalizedLanguage)
        .map((entry) => entry.code);
}

function normalizeDocType(value) {
    const type = stripMarkdown(value)
        .replace(/^microsoft\.graph\./i, '')
        .trim();

    const collectionMatch = type.match(/^collection\((.+)\)$/i);
    if (collectionMatch) {
        return `${collectionMatch[1].replace(/^microsoft\.graph\./i, '')} collection`;
    }

    return type;
}

function parseNamedTable(sectionText, firstColumnMatchers) {
    const tables = extractMarkdownTables(sectionText);
    const items = [];
    const seen = new Set();

    tables.forEach((tableLines) => {
        if (tableLines.length < 3) {
            return;
        }

        const headers = splitTableRow(tableLines[0]).map((cell) => stripMarkdown(cell).toLowerCase());
        const firstIndex = headers.findIndex((cell) => firstColumnMatchers.some((matcher) => cell.includes(matcher)));
        const typeIndex = headers.findIndex((cell) => cell.includes('type'));
        const descriptionIndex = headers.findIndex((cell) => cell.includes('description'));

        if (firstIndex === -1 || typeIndex === -1) {
            return;
        }

        for (let index = 2; index < tableLines.length; index += 1) {
            const cells = splitTableRow(tableLines[index]);
            const name = stripMarkdown(cells[firstIndex] || '');
            const type = normalizeDocType(cells[typeIndex] || '');
            const description = stripMarkdown(cells[descriptionIndex] || '');

            if (!name || !type) {
                continue;
            }

            const key = `${name}|${type}`;
            if (seen.has(key)) {
                continue;
            }

            seen.add(key);
            items.push({
                name,
                type,
                description
            });
        }
    });

    return items;
}

function extractJsonRepresentation(markdown, markdownFile) {
    const candidates = [
        extractSection(markdown, 'JSON representation'),
        extractSection(markdown, 'JSON')
    ].filter(Boolean);

    for (const section of candidates) {
        const expanded = expandMarkdownIncludes(section, markdownFile);
        const jsonFence = extractFirstCodeFence(expanded, 'json');
        if (jsonFence) {
            return jsonFence;
        }
    }

    return '';
}

function getResourceVersionBucket(registry, resourceName) {
    if (!registry.has(resourceName)) {
        registry.set(resourceName, {
            v1: null,
            beta: null
        });
    }

    return registry.get(resourceName);
}

function buildResourceDocsFromDocsTree(graphDocsRoot) {
    const registry = new Map();
    let filesParsed = 0;
    let filesWithMappings = 0;
    let propertyTables = 0;
    let relationshipTables = 0;
    let jsonRepresentations = 0;

    for (const versionSpec of VERSION_SPECS) {
        const resourceDir = path.join(graphDocsRoot, versionSpec.resourceDir);
        for (const markdownFile of listMarkdownFiles(resourceDir)) {
            filesParsed += 1;

            const markdown = readUtf8(markdownFile);
            const propertiesSection = extractSection(markdown, 'Properties');
            const relationshipsSection = extractSection(markdown, 'Relationships');
            const properties = propertiesSection
                ? parseNamedTable(expandMarkdownIncludes(propertiesSection, markdownFile), ['property', 'name'])
                : [];
            const relationships = relationshipsSection
                ? parseNamedTable(expandMarkdownIncludes(relationshipsSection, markdownFile), ['relationship', 'name'])
                : [];
            const jsonRepresentation = extractJsonRepresentation(markdown, markdownFile);
            const description = extractIntroText(markdown);

            if (properties.length === 0 && relationships.length === 0 && !jsonRepresentation && !description) {
                continue;
            }

            if (properties.length > 0) {
                propertyTables += 1;
            }

            if (relationships.length > 0) {
                relationshipTables += 1;
            }

            if (jsonRepresentation) {
                jsonRepresentations += 1;
            }

            filesWithMappings += 1;

            const resourceName = path.basename(markdownFile, '.md');
            const bucket = getResourceVersionBucket(registry, resourceName);
            bucket[versionSpec.key] = {
                name: resourceName,
                title: extractPageTitle(markdown, markdownFile),
                docLink: buildResourceDocLink(markdownFile, versionSpec),
                description,
                properties,
                relationships,
                jsonRepresentation
            };
        }
    }

    const data = {};
    Array.from(registry.keys())
        .sort((left, right) => left.localeCompare(right))
        .forEach((resourceName) => {
            data[resourceName] = registry.get(resourceName);
        });

    return {
        data,
        metrics: {
            filesParsed,
            filesWithMappings,
            resources: Object.keys(data).length,
            propertyTables,
            relationshipTables,
            jsonRepresentations
        }
    };
}

function createCodeRegistry() {
    return new Map();
}

function getCodeVersionBucket(registry, permission, version) {
    if (!registry.has(permission)) {
        registry.set(permission, {
            v1: Object.fromEntries(CODE_LANGUAGES.map((language) => [language, new Map()])),
            beta: Object.fromEntries(CODE_LANGUAGES.map((language) => [language, new Map()]))
        });
    }

    return registry.get(permission)[version];
}

function addCodeExampleMapping(registry, permission, version, language, entry) {
    const bucket = getCodeVersionBucket(registry, permission, version)[language];
    const key = `${entry.docLink}|${entry.endpoint || ''}|${entry.code}`;
    const existing = bucket.get(key);

    if (!existing) {
        bucket.set(key, { ...entry });
        return;
    }

    existing.supportsDelegated = existing.supportsDelegated || entry.supportsDelegated;
    existing.supportsApplication = existing.supportsApplication || entry.supportsApplication;
    existing.isLeastPrivilege = existing.isLeastPrivilege || entry.isLeastPrivilege;
    if (!existing.title && entry.title) {
        existing.title = entry.title;
    }
}

function compareCodeExamples(left, right) {
    return Number(Boolean(right.isLeastPrivilege)) - Number(Boolean(left.isLeastPrivilege))
        || String(left.title || '').localeCompare(String(right.title || ''))
        || String(left.endpoint || '').localeCompare(String(right.endpoint || ''));
}

function finalizeCodeRegistry(registry) {
    const results = {};

    for (const permission of Array.from(registry.keys()).sort((left, right) => left.localeCompare(right))) {
        const versionBuckets = registry.get(permission);
        results[permission] = {
            v1: Object.fromEntries(CODE_LANGUAGES.map((language) => [
                language,
                Array.from(versionBuckets.v1[language].values()).sort(compareCodeExamples)
            ])),
            beta: Object.fromEntries(CODE_LANGUAGES.map((language) => [
                language,
                Array.from(versionBuckets.beta[language].values()).sort(compareCodeExamples)
            ]))
        };
    }

    return results;
}

function countCodeExamples(data) {
    return Object.values(data).reduce((total, permissionEntry) =>
        total
        + CODE_LANGUAGES.reduce((languageTotal, language) =>
            languageTotal
            + (permissionEntry.v1?.[language] || []).length
            + (permissionEntry.beta?.[language] || []).length, 0), 0);
}

function extractPowerShellCommands(code) {
    const matches = code.match(/\b(?:Get|New|Set|Update|Remove|Invoke|Find|Test|Connect|Add)-Mg[A-Za-z0-9]+\b/g) || [];
    const seen = new Set();
    const results = [];

    matches.forEach((command) => {
        if (!seen.has(command)) {
            seen.add(command);
            results.push(command);
        }
    });

    return results;
}

function createRegistry() {
    return new Map();
}

function getVersionBucket(registry, permission, version) {
    if (!registry.has(permission)) {
        registry.set(permission, {
            v1: new Map(),
            beta: new Map()
        });
    }

    return registry.get(permission)[version];
}

function addApiMapping(registry, permission, version, entry) {
    const bucket = getVersionBucket(registry, permission, version);
    const key = `${entry.method}|${entry.path}`;
    const existing = bucket.get(key);

    if (!existing) {
        bucket.set(key, { ...entry });
        return;
    }

    existing.supportsDelegated = existing.supportsDelegated || entry.supportsDelegated;
    existing.supportsApplication = existing.supportsApplication || entry.supportsApplication;
    existing.isLeastPrivilege = existing.isLeastPrivilege || entry.isLeastPrivilege;
    if (!existing.docLink && entry.docLink) {
        existing.docLink = entry.docLink;
    }
}

function addPowerShellMapping(registry, permission, version, entry) {
    const bucket = getVersionBucket(registry, permission, version);
    const key = `${entry.command}|${entry.endpoint || ''}|${entry.docLink || ''}`;
    const existing = bucket.get(key);

    if (!existing) {
        bucket.set(key, { ...entry });
        return;
    }

    existing.supportsDelegated = existing.supportsDelegated || entry.supportsDelegated;
    existing.supportsApplication = existing.supportsApplication || entry.supportsApplication;
    existing.isLeastPrivilege = existing.isLeastPrivilege || entry.isLeastPrivilege;
    if (!existing.title && entry.title) {
        existing.title = entry.title;
    }
    if (!existing.code && entry.code) {
        existing.code = entry.code;
    }
}

function compareApiMappings(left, right) {
    const leftOrder = METHOD_ORDER[left.method] || 99;
    const rightOrder = METHOD_ORDER[right.method] || 99;
    return leftOrder - rightOrder || left.path.localeCompare(right.path);
}

function comparePowerShellMappings(left, right) {
    return left.command.localeCompare(right.command)
        || String(left.endpoint || '').localeCompare(String(right.endpoint || ''))
        || String(left.title || '').localeCompare(String(right.title || ''));
}

function finalizeRegistry(registry, comparator) {
    const results = {};

    for (const permission of Array.from(registry.keys()).sort((left, right) => left.localeCompare(right))) {
        const versionBuckets = registry.get(permission);
        results[permission] = {
            v1: Array.from(versionBuckets.v1.values()).sort(comparator),
            beta: Array.from(versionBuckets.beta.values()).sort(comparator)
        };
    }

    return results;
}

function countMappings(data) {
    return Object.values(data).reduce((total, entry) =>
        total + (entry.v1 || []).length + (entry.beta || []).length, 0);
}

function buildPermissionMethodsFromDocsTree(graphDocsRoot) {
    const registry = createRegistry();
    let filesParsed = 0;
    let filesWithMappings = 0;

    for (const versionSpec of VERSION_SPECS) {
        const apiDir = path.join(graphDocsRoot, versionSpec.docsDir);
        for (const markdownFile of listMarkdownFiles(apiDir)) {
            filesParsed += 1;

            const markdown = readUtf8(markdownFile);
            const permissions = extractPermissionsFromMarkdown(markdown, markdownFile);
            const requests = extractHttpRequests(markdown);

            if (permissions.length === 0 || requests.length === 0) {
                continue;
            }

            filesWithMappings += 1;
            const docLink = buildDocLink(markdownFile, versionSpec);

            permissions.forEach((permissionDescriptor) => {
                requests.forEach((requestDescriptor) => {
                    addApiMapping(registry, permissionDescriptor.permission, versionSpec.key, {
                        method: requestDescriptor.method,
                        path: requestDescriptor.path,
                        docLink,
                        supportsDelegated: permissionDescriptor.supportsDelegated,
                        supportsApplication: permissionDescriptor.supportsApplication,
                        isLeastPrivilege: permissionDescriptor.isLeastPrivilege
                    });
                });
            });
        }
    }

    const data = finalizeRegistry(registry, compareApiMappings);
    return {
        data,
        metrics: {
            filesParsed,
            filesWithMappings,
            permissions: Object.keys(data).length,
            mappings: countMappings(data)
        }
    };
}

function buildPermissionPowerShellFromDocsTree(graphDocsRoot) {
    const registry = createRegistry();
    let filesParsed = 0;
    let filesWithMappings = 0;

    for (const versionSpec of VERSION_SPECS) {
        const apiDir = path.join(graphDocsRoot, versionSpec.docsDir);
        for (const markdownFile of listMarkdownFiles(apiDir)) {
            filesParsed += 1;

            const markdown = readUtf8(markdownFile);
            const permissions = extractPermissionsFromMarkdown(markdown, markdownFile);
            const requests = extractHttpRequests(markdown);
            const snippetPaths = extractSnippetIncludePaths(markdown, markdownFile, 'powershell');

            if (permissions.length === 0) {
                continue;
            }

            const docLink = buildDocLink(markdownFile, versionSpec);
            const title = extractPageTitle(markdown, markdownFile);
            const endpoint = requests[0]?.path || null;
            const inlineCodes = extractInlineCodeEntries(markdown, 'powershell');

            const snippetEntries = snippetPaths.flatMap((snippetPath) => {
                const snippetMarkdown = readUtf8(snippetPath);
                const code = extractFirstCodeFence(snippetMarkdown, 'powershell');
                const commands = extractPowerShellCommands(code);

                if (commands.length === 0) {
                    return [];
                }

                return commands.map((command) => ({
                    command,
                    endpoint,
                    title,
                    docLink,
                    code
                }));
            });

            inlineCodes.forEach((code) => {
                extractPowerShellCommands(code).forEach((command) => {
                    snippetEntries.push({
                        command,
                        endpoint,
                        title,
                        docLink,
                        code
                    });
                });
            });

            if (snippetEntries.length === 0) {
                continue;
            }

            filesWithMappings += 1;

            permissions.forEach((permissionDescriptor) => {
                snippetEntries.forEach((entry) => {
                    addPowerShellMapping(registry, permissionDescriptor.permission, versionSpec.key, {
                        ...entry,
                        supportsDelegated: permissionDescriptor.supportsDelegated,
                        supportsApplication: permissionDescriptor.supportsApplication,
                        isLeastPrivilege: permissionDescriptor.isLeastPrivilege
                    });
                });
            });
        }
    }

    const data = finalizeRegistry(registry, comparePowerShellMappings);
    return {
        data,
        metrics: {
            filesParsed,
            filesWithMappings,
            permissions: Object.keys(data).length,
            mappings: countMappings(data)
        }
    };
}

function buildPermissionCodeExamplesFromDocsTree(graphDocsRoot) {
    const registry = createCodeRegistry();
    let filesParsed = 0;
    let filesWithMappings = 0;

    for (const versionSpec of VERSION_SPECS) {
        const apiDir = path.join(graphDocsRoot, versionSpec.docsDir);
        for (const markdownFile of listMarkdownFiles(apiDir)) {
            filesParsed += 1;

            const markdown = readUtf8(markdownFile);
            const permissions = extractPermissionsFromMarkdown(markdown, markdownFile);
            const requests = extractHttpRequests(markdown);

            if (permissions.length === 0) {
                continue;
            }

            const endpoint = requests[0]?.path || null;
            const docLink = buildDocLink(markdownFile, versionSpec);
            const title = extractPageTitle(markdown, markdownFile);
            const snippetEntries = [];

            CODE_LANGUAGES.forEach((language) => {
                const seenCode = new Set();
                extractSnippetIncludePaths(markdown, markdownFile, language).forEach((snippetPath) => {
                    const code = extractFirstCodeFence(readUtf8(snippetPath), language);
                    if (!code) {
                        return;
                    }

                    seenCode.add(code);
                    snippetEntries.push({
                        language,
                        endpoint,
                        title,
                        docLink,
                        code
                    });
                });

                extractInlineCodeEntries(markdown, language).forEach((code) => {
                    if (seenCode.has(code)) {
                        return;
                    }

                    snippetEntries.push({
                        language,
                        endpoint,
                        title,
                        docLink,
                        code
                    });
                });
            });

            if (snippetEntries.length === 0) {
                continue;
            }

            filesWithMappings += 1;

            permissions.forEach((permissionDescriptor) => {
                snippetEntries.forEach((entry) => {
                    addCodeExampleMapping(registry, permissionDescriptor.permission, versionSpec.key, entry.language, {
                        endpoint: entry.endpoint,
                        title: entry.title,
                        docLink: entry.docLink,
                        code: entry.code,
                        supportsDelegated: permissionDescriptor.supportsDelegated,
                        supportsApplication: permissionDescriptor.supportsApplication,
                        isLeastPrivilege: permissionDescriptor.isLeastPrivilege
                    });
                });
            });
        }
    }

    const data = finalizeCodeRegistry(registry);
    return {
        data,
        metrics: {
            filesParsed,
            filesWithMappings,
            permissions: Object.keys(data).length,
            snippets: countCodeExamples(data)
        }
    };
}

function generatePermissionDocsFromGraphDocs(outputDir, options = {}) {
    const resolvedOutputDir = path.resolve(outputDir);
    const configuredRepoDir = options.repoDir || process.env.GRAPH_DOCS_CACHE_DIR || null;
    const tempRoot = configuredRepoDir ? null : createShortTempRoot();
    const repoDir = configuredRepoDir ? path.resolve(configuredRepoDir) : path.join(tempRoot, 'repo');

    try {
        if (!options.repoDir) {
            cloneGraphDocs(repoDir);
        }

        const apiResult = buildPermissionMethodsFromDocsTree(repoDir);
        const powerShellResult = buildPermissionPowerShellFromDocsTree(repoDir);
        const codeExamplesResult = buildPermissionCodeExamplesFromDocsTree(repoDir);
        const resourceDocsResult = buildResourceDocsFromDocsTree(repoDir);

        writeJson(path.join(resolvedOutputDir, 'GraphPermissionMethods.json'), apiResult.data);
        writeJson(path.join(resolvedOutputDir, 'GraphPermissionPowerShell.json'), powerShellResult.data);
        writeJsonSharded(path.join(resolvedOutputDir, 'GraphPermissionCodeExamples.json'), codeExamplesResult.data, {
            maxEntriesPerPart: 120
        });
        writeJson(path.join(resolvedOutputDir, 'GraphResourceDocumentation.json'), resourceDocsResult.data);

        return {
            outputDir: resolvedOutputDir,
            api: apiResult.metrics,
            powershell: powerShellResult.metrics,
            codeExamples: codeExamplesResult.metrics,
            resources: resourceDocsResult.metrics
        };
    } finally {
        if (tempRoot) {
            fs.rmSync(tempRoot, { recursive: true, force: true });
        }
    }
}

function generatePermissionMethodsFromGraphDocs(outputFile, options = {}) {
    const outputPath = path.resolve(outputFile);
    const result = generatePermissionDocsFromGraphDocs(path.dirname(outputPath), options);
    return {
        outputFile: outputPath,
        ...result.api
    };
}

module.exports = {
    buildPermissionMethodsFromDocsTree,
    buildPermissionCodeExamplesFromDocsTree,
    buildPermissionPowerShellFromDocsTree,
    buildResourceDocsFromDocsTree,
    extractHttpRequests,
    extractPermissionsFromMarkdown,
    generatePermissionDocsFromGraphDocs,
    generatePermissionMethodsFromGraphDocs
};
