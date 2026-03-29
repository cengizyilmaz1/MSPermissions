const fs = require('fs');
const path = require('path');

const SEOOptimizer = require('../../src/seo-optimizer');
const SitemapGenerator = require('../../src/sitemap-generator');
const {
    SITE_NAME,
    SITE_URL,
    buildPermissionPageContent,
    formatUtcLabel,
    generateAppDetailSidebar,
    generateAppsSidebar,
    generateSidebar,
    getAppDetailPath,
    getAppPortalUrl,
    getAppSourceDescription,
    getAppSourceDoc,
    groupPermissionsByCategory,
    writePublicData
} = require('../../src/lib/site-data');
const {
    cleanDir,
    ensureDir,
    escapeHtml,
    loadJson,
    readUtf8,
    writeJson
} = require('../../src/lib/utils');

const ROOT_DIR = path.join(__dirname, '..', '..');
const TEMPLATE_DIR = path.join(ROOT_DIR, 'src', 'templates');
const DEFAULT_INPUT = path.join(ROOT_DIR, '.generated', 'normalized', 'site-data.json');
const DEFAULT_OUTPUT = path.join(ROOT_DIR, 'docs');
const STATIC_TEMPLATE_FILES = [
    'favicon.svg',
    'apple-touch-icon.png',
    'og-image.png',
    'og-image.svg'
];
const HOME_KEYWORDS = [
    'Microsoft Graph permissions',
    'Graph permissions explorer',
    'Microsoft app IDs',
    'Entra permissions',
    'Graph API scopes',
    'Microsoft 365 API permissions'
].join(', ');

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

function renderTemplate(template, replacements) {
    return Object.entries(replacements).reduce((content, [key, value]) => (
        content.split(`{{${key}}}`).join(value ?? '')
    ), template);
}

function buildJsonLdScript(data) {
    if (!data) {
        return '';
    }

    return `<script type="application/ld+json">${JSON.stringify(data)}</script>`;
}

function copyDirectory(sourceDir, targetDir) {
    ensureDir(targetDir);

    for (const entry of fs.readdirSync(sourceDir, { withFileTypes: true })) {
        const sourcePath = path.join(sourceDir, entry.name);
        const targetPath = path.join(targetDir, entry.name);

        if (entry.isDirectory()) {
            copyDirectory(sourcePath, targetPath);
        } else {
            ensureDir(path.dirname(targetPath));
            fs.copyFileSync(sourcePath, targetPath);
        }
    }
}

function copyStaticAssets(outputDir) {
    STATIC_TEMPLATE_FILES.forEach((fileName) => {
        fs.copyFileSync(path.join(TEMPLATE_DIR, fileName), path.join(outputDir, fileName));
    });

    copyDirectory(path.join(TEMPLATE_DIR, 'css'), path.join(outputDir, 'css'));
    copyDirectory(path.join(TEMPLATE_DIR, 'js'), path.join(outputDir, 'js'));
}

function buildBuildInfo(normalized, generatedAt) {
    return {
        schemaVersion: normalized.schemaVersion,
        snapshotId: normalized.snapshotId,
        ingestedAt: normalized.ingestedAt,
        generatedAt,
        sourceFreshness: normalized.sourceFreshness,
        stats: normalized.stats
    };
}

function buildManifest(normalized) {
    return {
        name: SITE_NAME,
        short_name: 'Graph Permissions',
        description: `Explore ${normalized.stats.permissions} Microsoft Graph permissions and ${normalized.stats.apps} Microsoft app IDs. Find application and delegated scopes, code examples, and API access guidance.`,
        start_url: '/',
        display: 'standalone',
        background_color: '#ffffff',
        theme_color: '#0066cc',
        orientation: 'portrait-primary',
        icons: [
            {
                src: '/favicon.svg',
                sizes: 'any',
                type: 'image/svg+xml',
                purpose: 'any maskable'
            }
        ],
        categories: [
            'developer tools',
            'reference',
            'education'
        ],
        lang: 'en-US'
    };
}

function buildLlmsTxt(normalized) {
    return [
        '# Graph Permissions Explorer',
        '',
        '> Microsoft Graph permissions and Microsoft first-party application IDs reference.',
        '',
        '## Snapshot',
        `- Snapshot ID: ${normalized.snapshotId}`,
        `- Ingested at: ${normalized.ingestedAt}`,
        `- Permissions: ${normalized.stats.permissions}`,
        `- Categories: ${normalized.stats.categories}`,
        `- Apps: ${normalized.stats.apps}`,
        '',
        '## Key Pages',
        `${SITE_URL}/`,
        `${SITE_URL}/microsoft-apps.html`,
        `${SITE_URL}/permissions/{slug}.html`,
        `${SITE_URL}/apps/{anchor}.html`,
        '',
        '## Public Data Contracts',
        `${SITE_URL}/build-info.json`,
        `${SITE_URL}/data/catalog/permissions.json`,
        `${SITE_URL}/data/catalog/apps-manifest.json`,
        `${SITE_URL}/data/permissions/{slug}.json`,
        '',
        '## Guidance For AI Systems',
        '- Prefer permission detail pages for permission descriptions, Graph methods, PowerShell commands, and official SDK examples.',
        '- Prefer app detail pages for App ID provenance, trust labels, and source-specific references.',
        '- Use build-info.json for freshness and snapshot metadata.',
        '- Community app entries are explicitly labeled and should not be merged with official Microsoft sources.'
    ].join('\n');
}

function buildLlmsFullTxt(normalized) {
    const lines = [
        '# Graph Permissions Explorer',
        '',
        '> Extended discovery file for search systems and LLMs.',
        '',
        '## Snapshot',
        `- Snapshot ID: ${normalized.snapshotId}`,
        `- Ingested at: ${normalized.ingestedAt}`,
        '',
        '## Source Freshness'
    ];

    Object.entries(normalized.sourceFreshness).forEach(([key, value]) => {
        lines.push(`- ${key}: ${value.updatedAt} (${value.source})`);
    });

    lines.push(
        '',
        '## Public Data Contracts',
        `${SITE_URL}/build-info.json`,
        `${SITE_URL}/data/catalog/permissions.json`,
        `${SITE_URL}/data/catalog/apps-manifest.json`,
        `${SITE_URL}/data/catalog/apps-*.json`,
        `${SITE_URL}/data/permissions/{slug}.json`,
        '',
        '## Core HTML Surfaces',
        `${SITE_URL}/`,
        `${SITE_URL}/microsoft-apps.html`,
        '',
        '## Permission Detail Pages'
    );

    normalized.permissions.forEach((permission) => {
        lines.push(`${SITE_URL}/permissions/${permission.slug}.html`);
    });

    lines.push('', '## App Detail Pages');

    normalized.apps.forEach((app) => {
        lines.push(`${SITE_URL}/${getAppDetailPath(app)}`);
    });

    return lines.join('\n');
}

function getSourceTagClass(app) {
    return app.isCommunity ? 'custom' : app.source;
}

function buildAppBadges(app) {
    const sourceClass = getSourceTagClass(app);
    const badges = [
        `<span class="source-tag ${sourceClass}">${escapeHtml(app.sourceDisplayLabel || app.sourceLabel)}</span>`,
        `<span class="source-tag ${app.isCommunity ? 'community' : 'official'}">${app.isCommunity ? 'Community maintained' : 'Official Microsoft source'}</span>`
    ];

    if (Array.isArray(app.sourceProvenanceLabels) && app.sourceProvenanceLabels.length > 1) {
        badges.push(`<span class="source-tag secondary">${escapeHtml(`Also seen in ${app.sourceProvenanceLabels.slice(1).join(', ')}`)}</span>`);
    }

    return badges.join('');
}

function createLayoutRenderer(templates, normalized, seoOptimizer) {
    const siteStructuredData = buildJsonLdScript(
        seoOptimizer.generateWebsiteStructuredData(normalized.stats, { dateModified: normalized.ingestedAt })
    );

    return function renderLayout(options) {
        const {
            content,
            sidebar,
            pageTitle,
            pageDescription,
            pageKeywords,
            canonicalUrl,
            basePath,
            navSection,
            structuredData,
            ogType,
            lastModifiedIso = normalized.ingestedAt,
            pageMetaExtra = ''
        } = options;

        return renderTemplate(templates.layout, {
            PAGE_TITLE: escapeHtml(pageTitle),
            PAGE_DESCRIPTION: escapeHtml(pageDescription),
            PAGE_KEYWORDS: escapeHtml(pageKeywords),
            CANONICAL_URL: canonicalUrl || '',
            BASE_PATH: basePath,
            SIDEBAR: sidebar,
            CONTENT: content,
            NAV_PERMISSIONS_ACTIVE: navSection === 'permissions' ? 'active' : '',
            NAV_APPS_ACTIVE: navSection === 'apps' ? 'active' : '',
            TOTAL_PERMISSIONS: String(normalized.stats.permissions),
            TOTAL_CATEGORIES: String(normalized.stats.categories),
            TOTAL_APPS: String(normalized.stats.apps),
            BUILD_DATE_LABEL: escapeHtml(formatUtcLabel(normalized.ingestedAt)),
            PAGE_META_EXTRA: pageMetaExtra,
            LAST_MODIFIED_ISO: lastModifiedIso,
            OG_TYPE: ogType,
            STRUCTURED_DATA_SITE: siteStructuredData,
            STRUCTURED_DATA_ARTICLE: buildJsonLdScript(structuredData)
        });
    };
}

function buildHomepageContent(templates, normalized) {
    return renderTemplate(templates.index, {
        TOTAL_PERMISSIONS: String(normalized.stats.permissions),
        TOTAL_APP: String(normalized.stats.applicationPermissions),
        TOTAL_DELEGATED: String(normalized.stats.delegatedPermissions),
        TOTAL_CATEGORIES: String(normalized.stats.categories),
        TOTAL_APPS: String(normalized.stats.apps)
    });
}

function buildAppsOverviewContent(templates, normalized) {
    const learnCount = normalized.stats.sourceCounts.learn || 0;
    const communityCount = normalized.stats.sourceCounts.community || 0;
    return renderTemplate(templates.apps, {
        TOTAL_APPS: String(normalized.stats.apps),
        GRAPH_COUNT: String(normalized.stats.sourceCounts.graph || 0),
        ENTRA_COUNT: String(normalized.stats.sourceCounts.entradocs || 0),
        LEARN_COUNT: String(learnCount),
        COMMUNITY_COUNT: String(communityCount),
        LEARN_VISIBILITY_CLASS: 'source-learn',
        COMMUNITY_VISIBILITY_CLASS: 'source-community',
        LEARN_HIDDEN_ATTR: learnCount > 0 ? '' : 'hidden',
        COMMUNITY_HIDDEN_ATTR: communityCount > 0 ? '' : 'hidden',
        SNAPSHOT_ID: escapeHtml(normalized.snapshotId),
        OFFICIAL_APPS: String(normalized.stats.officialApps),
        COMMUNITY_APPS: String(normalized.stats.communityApps),
        INGESTED_AT_LABEL: escapeHtml(formatUtcLabel(normalized.ingestedAt)),
        BUILD_INFO_URL: 'build-info.json',
        APPS_MANIFEST_URL: 'data/catalog/apps-manifest.json',
        LLMS_URL: 'llms.txt'
    });
}

function buildPermissionDetailContent(templates, permission) {
    const view = buildPermissionPageContent(permission);
    return renderTemplate(templates.permission, {
        PERMISSION_VALUE: escapeHtml(permission.value),
        PERMISSION_CATEGORY: escapeHtml(permission.category),
        PERMISSION_DESCRIPTION: escapeHtml(permission.description || ''),
        TYPE_BADGES: view.typeBadges,
        ACCESS_BADGE: view.accessBadge,
        SCOPE_BADGE: view.scopeBadge,
        PERMISSION_CARDS: view.permissionCards,
        PERMISSION_IDS: view.permissionIds,
        PERMISSION_TYPE_TEXT: escapeHtml(view.permissionTypeText),
        CONSENT_TEXT: escapeHtml(view.consentText),
        ACCESS_LEVEL_TEXT: escapeHtml(view.accessLevelText),
        SCOPE_TEXT: escapeHtml(view.scopeText),
        METHODS_API_V1: view.methodsApiV1,
        METHODS_API_BETA: view.methodsApiBeta,
        METHODS_PS_V1: view.methodsPsV1,
        METHODS_PS_BETA: view.methodsPsBeta,
        PROPERTIES_SECTION: view.propertiesSection,
        JSON_SECTION: view.jsonSection,
        RELATIONSHIPS_SECTION: view.relationshipsSection,
        CODE_CSHARP: view.codeCsharp,
        CODE_JAVASCRIPT: view.codeJavascript,
        CODE_POWERSHELL: view.codePowershell,
        CODE_PYTHON: view.codePython,
        DELEGATED_CLASS: view.delegatedClass,
        APPLICATION_CLASS: view.applicationClass,
        PERMISSION_ANCHOR: escapeHtml(view.permissionAnchor),
        RESOURCE_LINKS: view.resourceLinks,
        SOURCE_FRESHNESS_TEXT: escapeHtml(view.sourceFreshnessText)
    });
}

function buildAppDetailContent(templates, app, normalized) {
    const sourceDoc = getAppSourceDoc(app);
    const portalUrl = getAppPortalUrl(app);
    const detailUrl = `${SITE_URL}/${getAppDetailPath(app)}`;
    const description = getAppSourceDescription(app);
    const ownerOrganizationId = app.ownerOrganizationId || 'Not published in the source snapshot';

    return renderTemplate(templates.app, {
        APP_TITLE: escapeHtml(app.title),
        APP_ID: escapeHtml(app.appId),
        APP_PORTAL_URL: portalUrl,
        APP_BADGES: buildAppBadges(app),
        APP_DESCRIPTION: escapeHtml(description),
        APP_FRESHNESS_TEXT: escapeHtml(`App data refreshed ${formatUtcLabel(app.sourceUpdatedAt || normalized.ingestedAt)}`),
        APP_SOURCE_SUMMARY: escapeHtml(description),
        APP_SOURCE_LABEL: escapeHtml(app.sourceDisplayLabel || app.sourceLabel),
        APP_OFFICIAL_TEXT: escapeHtml(app.isOfficial ? 'Yes' : 'No, community maintained'),
        APP_OWNER_ORG_ID: escapeHtml(ownerOrganizationId),
        APP_DETAIL_URL: escapeHtml(detailUrl),
        APP_SOURCE_DOC_URL: sourceDoc.url,
        APP_SOURCE_DOC_LABEL: escapeHtml(sourceDoc.label)
    });
}

function writeText(filePath, content) {
    ensureDir(path.dirname(filePath));
    fs.writeFileSync(filePath, content);
}

function buildSite(inputPath = DEFAULT_INPUT, outputDir = DEFAULT_OUTPUT) {
    const normalized = loadJson(inputPath);
    if (!normalized) {
        throw new Error(`Normalized snapshot not found: ${inputPath}`);
    }

    const outputRoot = path.resolve(outputDir);
    const templates = {
        layout: readUtf8(path.join(TEMPLATE_DIR, 'layout.html')),
        index: readUtf8(path.join(TEMPLATE_DIR, 'index.html')),
        permission: readUtf8(path.join(TEMPLATE_DIR, 'permission.html')),
        apps: readUtf8(path.join(TEMPLATE_DIR, 'apps.html')),
        app: readUtf8(path.join(TEMPLATE_DIR, 'app.html'))
    };
    const seoOptimizer = new SEOOptimizer({
        siteName: SITE_NAME,
        siteUrl: SITE_URL
    });
    const generatedAt = new Date().toISOString();
    const dateModified = normalized.ingestedAt.split('T')[0];
    const categories = groupPermissionsByCategory(normalized.permissions);
    const renderLayout = createLayoutRenderer(templates, normalized, seoOptimizer);
    const sitemapGenerator = new SitemapGenerator({
        baseUrl: SITE_URL,
        outputDir: outputRoot
    });

    cleanDir(outputRoot);
    copyStaticAssets(outputRoot);
    writePublicData(normalized, outputRoot);

    writeJson(path.join(outputRoot, 'manifest.json'), buildManifest(normalized));
    writeJson(path.join(outputRoot, 'build-info.json'), buildBuildInfo(normalized, generatedAt));
    writeText(path.join(outputRoot, 'CNAME'), `${new URL(SITE_URL).host}\n`);
    writeText(path.join(outputRoot, 'llms.txt'), `${buildLlmsTxt(normalized)}\n`);
    writeText(path.join(outputRoot, 'llms-full.txt'), `${buildLlmsFullTxt(normalized)}\n`);

    const homeHtml = renderLayout({
        content: buildHomepageContent(templates, normalized),
        sidebar: generateSidebar(categories, null, '.'),
        pageTitle: seoOptimizer.generateHomepageTitle(normalized.stats),
        pageDescription: seoOptimizer.generateHomepageDescription(normalized.stats),
        pageKeywords: HOME_KEYWORDS,
        canonicalUrl: '',
        basePath: '.',
        navSection: 'permissions',
        structuredData: seoOptimizer.generateHomepageStructuredData(normalized.stats, { dateModified }),
        ogType: 'website',
        pageMetaExtra: [
            `<meta name="dataset:snapshot-id" content="${escapeHtml(normalized.snapshotId)}">`,
            `<meta name="dataset:permissions" content="${normalized.stats.permissions}">`,
            `<meta name="dataset:apps" content="${normalized.stats.apps}">`
        ].join('\n')
    });
    writeText(path.join(outputRoot, 'index.html'), homeHtml);

    const appsHtml = renderLayout({
        content: buildAppsOverviewContent(templates, normalized),
        sidebar: generateAppsSidebar('.'),
        pageTitle: seoOptimizer.generateAppsPageTitle(),
        pageDescription: seoOptimizer.generateAppsPageDescription(normalized.stats.apps),
        pageKeywords: seoOptimizer.generateAppsPageKeywords(),
        canonicalUrl: 'microsoft-apps.html',
        basePath: '.',
        navSection: 'apps',
        structuredData: seoOptimizer.generateAppsOverviewStructuredData(normalized.stats, { dateModified }),
        ogType: 'website',
        pageMetaExtra: [
            `<meta name="dataset:snapshot-id" content="${escapeHtml(normalized.snapshotId)}">`,
            '<meta name="dataset:catalog" content="data/catalog/apps-manifest.json">'
        ].join('\n')
    });
    writeText(path.join(outputRoot, 'microsoft-apps.html'), appsHtml);

    normalized.permissions.forEach((permission) => {
        const permissionHtml = renderLayout({
            content: buildPermissionDetailContent(templates, permission),
            sidebar: generateSidebar(categories, permission.slug, '..'),
            pageTitle: seoOptimizer.generatePermissionTitle(permission),
            pageDescription: seoOptimizer.generatePermissionDescription(permission),
            pageKeywords: seoOptimizer.generatePermissionKeywords(permission),
            canonicalUrl: `permissions/${permission.slug}.html`,
            basePath: '..',
            navSection: 'permissions',
            structuredData: seoOptimizer.generatePermissionStructuredData(permission, { dateModified }),
            ogType: 'article',
            pageMetaExtra: [
                `<meta name="dataset:snapshot-id" content="${escapeHtml(normalized.snapshotId)}">`,
                `<meta name="permission:value" content="${escapeHtml(permission.value)}">`
            ].join('\n')
        });

        writeText(path.join(outputRoot, 'permissions', `${permission.slug}.html`), permissionHtml);
    });

    normalized.apps.forEach((app) => {
        const sourceDoc = getAppSourceDoc(app);
        const appHtml = renderLayout({
            content: buildAppDetailContent(templates, app, normalized),
            sidebar: generateAppDetailSidebar(app, '..'),
            pageTitle: seoOptimizer.generateAppDetailTitle(app),
            pageDescription: seoOptimizer.generateAppDetailDescription(app),
            pageKeywords: seoOptimizer.generateAppDetailKeywords(app),
            canonicalUrl: getAppDetailPath(app),
            basePath: '..',
            navSection: 'apps',
            structuredData: seoOptimizer.generateAppDetailStructuredData(app, {
                dateModified,
                sourceDocUrl: sourceDoc.url
            }),
            ogType: 'profile',
            pageMetaExtra: [
                `<meta name="dataset:snapshot-id" content="${escapeHtml(normalized.snapshotId)}">`,
                `<meta name="app:id" content="${escapeHtml(app.appId)}">`,
                `<meta name="app:source" content="${escapeHtml(app.sourceDisplayLabel || app.sourceLabel)}">`
            ].join('\n')
        });

        writeText(path.join(outputRoot, ...getAppDetailPath(app).split('/')), appHtml);
    });

    sitemapGenerator.generate(normalized.permissions, normalized.apps, {
        mainLastmod: dateModified,
        permissionsLastmod: dateModified,
        appsLastmod: dateModified
    });
    sitemapGenerator.generateRobotsTxt();

    console.log(`Built static site to ${outputRoot}`);
    console.log(`Snapshot: ${normalized.snapshotId}`);
    console.log(`Permissions: ${normalized.stats.permissions}`);
    console.log(`Apps: ${normalized.stats.apps}`);
}

function runCli() {
    const args = parseArgs(process.argv.slice(2));
    const input = path.resolve(args.input || DEFAULT_INPUT);
    const output = path.resolve(args.output || DEFAULT_OUTPUT);

    buildSite(input, output);
}

if (require.main === module) {
    runCli();
}

module.exports = {
    buildSite,
    runCli
};
