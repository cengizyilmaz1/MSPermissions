const fs = require('fs');
const path = require('path');

const SEOOptimizer = require('../../src/seo-optimizer');
const SitemapGenerator = require('../../src/sitemap-generator');
const { writeLlmsFiles } = require('../../src/llms-generator');
const {
    SITE_NAME,
    SITE_URL,
    buildPermissionPageContent,
    formatUtcLabel,
    generateAppDetailSidebar,
    generateAppsSidebar,
    generateFaqHtml,
    generateFaqItemsHtml,
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
const { runCommand } = require('./lib/cli');
const { createLogger } = require('./lib/logger');

const log = createLogger({ scope: 'build' });

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

function renderTemplate(template, replacements) {
    return Object.entries(replacements).reduce(
        (content, [key, value]) => content.split(`{{${key}}}`).join(value ?? ''),
        template
    );
}

function buildJsonLdScript(data) {
    const payload = Array.isArray(data) ? data.filter(Boolean) : data;

    if (!payload || (Array.isArray(payload) && payload.length === 0)) {
        return '';
    }

    return `<script type="application/ld+json">${JSON.stringify(payload)}</script>`;
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
        categories: ['developer tools', 'reference', 'education'],
        lang: 'en-US'
    };
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
        badges.push(
            `<span class="source-tag secondary">${escapeHtml(`Also seen in ${app.sourceProvenanceLabels.slice(1).join(', ')}`)}</span>`
        );
    }

    return badges.join('');
}

function createLayoutRenderer(templates, normalized, seoOptimizer) {
    const siteStructuredData = buildJsonLdScript(
        seoOptimizer.generateWebsiteStructuredData(normalized.stats, {
            dateModified: normalized.ingestedAt
        })
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
            breadcrumb = null,
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
            STRUCTURED_DATA_BREADCRUMB: buildJsonLdScript(
                seoOptimizer.generateBreadcrumbStructuredData(breadcrumb)
            ),
            STRUCTURED_DATA_ARTICLE: buildJsonLdScript(structuredData)
        });
    };
}

function buildHomepageContent(templates, normalized, faqEntries) {
    return renderTemplate(templates.index, {
        TOTAL_PERMISSIONS: String(normalized.stats.permissions),
        TOTAL_APP: String(normalized.stats.applicationPermissions),
        TOTAL_DELEGATED: String(normalized.stats.delegatedPermissions),
        TOTAL_CATEGORIES: String(normalized.stats.categories),
        TOTAL_APPS: String(normalized.stats.apps),
        HOME_FAQ_ITEMS: generateFaqItemsHtml(faqEntries)
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

function buildPermissionDetailContent(templates, permission, view, faqEntries) {
    return renderTemplate(templates.permission, {
        ANSWER_BLOCK: view.answerBlock,
        FAQ_SECTION: generateFaqHtml(faqEntries),
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

/**
 * Names the app and its App ID in one sentence so the claim survives being
 * quoted on its own by an answer engine.
 */
function buildAppSummary(app) {
    const trust = app.isCommunity
        ? 'a community-maintained entry in this catalog'
        : `an official Microsoft first-party application recorded in ${app.sourceDisplayLabel || app.sourceLabel}`;

    return `${app.title} is ${trust}. Its Microsoft application (client) ID is ${app.appId}, which is the value you will see for this application in Microsoft Entra ID sign-in logs, audit logs, and service principal listings.`;
}

function buildAppDetailContent(templates, app, normalized, faqEntries) {
    const sourceDoc = getAppSourceDoc(app);
    const portalUrl = getAppPortalUrl(app);
    const detailUrl = `${SITE_URL}/${getAppDetailPath(app)}`;
    const description = getAppSourceDescription(app);
    const ownerOrganizationId = app.ownerOrganizationId || 'Not published in the source snapshot';

    return renderTemplate(templates.app, {
        APP_SUMMARY_TEXT: escapeHtml(buildAppSummary(app)),
        APP_FAQ_SECTION: generateFaqHtml(faqEntries),
        APP_TITLE: escapeHtml(app.title),
        APP_ID: escapeHtml(app.appId),
        APP_PORTAL_URL: portalUrl,
        APP_BADGES: buildAppBadges(app),
        APP_DESCRIPTION: escapeHtml(description),
        APP_FRESHNESS_TEXT: escapeHtml(
            `App data refreshed ${formatUtcLabel(app.sourceUpdatedAt || normalized.ingestedAt)}`
        ),
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

    const llmsSizes = writeLlmsFiles(normalized, outputRoot, { siteUrl: SITE_URL });

    const homeFaqEntries = seoOptimizer.buildHomepageFaqEntries(normalized.stats);
    const homeHtml = renderLayout({
        content: buildHomepageContent(templates, normalized, homeFaqEntries),
        sidebar: generateSidebar(categories, null, '.'),
        pageTitle: seoOptimizer.generateHomepageTitle(normalized.stats),
        pageDescription: seoOptimizer.generateHomepageDescription(normalized.stats),
        pageKeywords: HOME_KEYWORDS,
        canonicalUrl: '',
        basePath: '.',
        navSection: 'permissions',
        structuredData: [
            ...seoOptimizer.generateHomepageStructuredData(normalized.stats, {
                dateModified
            }),
            seoOptimizer.generateFaqStructuredData(homeFaqEntries, { url: `${SITE_URL}/` })
        ],
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
        structuredData: seoOptimizer.generateAppsOverviewStructuredData(normalized.stats, {
            dateModified
        }),
        breadcrumb: [
            { name: 'Home', url: `${SITE_URL}/` },
            { name: 'Microsoft app IDs', url: `${SITE_URL}/microsoft-apps.html` }
        ],
        ogType: 'website',
        pageMetaExtra: [
            `<meta name="dataset:snapshot-id" content="${escapeHtml(normalized.snapshotId)}">`,
            '<meta name="dataset:catalog" content="data/catalog/apps-manifest.json">'
        ].join('\n')
    });
    writeText(path.join(outputRoot, 'microsoft-apps.html'), appsHtml);

    normalized.permissions.forEach((permission) => {
        const view = buildPermissionPageContent(permission);
        const permissionUrl = `${SITE_URL}/permissions/${permission.slug}.html`;
        const faqEntries = seoOptimizer.buildPermissionFaqEntries(permission, view);
        const permissionHtml = renderLayout({
            content: buildPermissionDetailContent(templates, permission, view, faqEntries),
            sidebar: generateSidebar(categories, permission.slug, '..'),
            pageTitle: seoOptimizer.generatePermissionTitle(permission),
            pageDescription: seoOptimizer.generatePermissionDescription(permission),
            pageKeywords: seoOptimizer.generatePermissionKeywords(permission),
            canonicalUrl: `permissions/${permission.slug}.html`,
            basePath: '..',
            navSection: 'permissions',
            structuredData: [
                seoOptimizer.generatePermissionStructuredData(permission, {
                    dateModified
                }),
                seoOptimizer.generateFaqStructuredData(faqEntries, { url: permissionUrl })
            ],
            breadcrumb: [
                { name: 'Home', url: `${SITE_URL}/` },
                {
                    name: `${permission.category} permissions`,
                    url: `${SITE_URL}/#${encodeURIComponent(permission.category)}`
                },
                { name: permission.value, url: permissionUrl }
            ],
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
        const appUrl = `${SITE_URL}/${getAppDetailPath(app)}`;
        const faqEntries = seoOptimizer.buildAppFaqEntries(app, sourceDoc);
        const appHtml = renderLayout({
            content: buildAppDetailContent(templates, app, normalized, faqEntries),
            sidebar: generateAppDetailSidebar(app, '..'),
            pageTitle: seoOptimizer.generateAppDetailTitle(app),
            pageDescription: seoOptimizer.generateAppDetailDescription(app),
            pageKeywords: seoOptimizer.generateAppDetailKeywords(app),
            canonicalUrl: getAppDetailPath(app),
            basePath: '..',
            navSection: 'apps',
            structuredData: [
                ...seoOptimizer.generateAppDetailStructuredData(app, {
                    dateModified,
                    sourceDocUrl: sourceDoc.url
                }),
                seoOptimizer.generateFaqStructuredData(faqEntries, { url: appUrl })
            ],
            breadcrumb: [
                { name: 'Home', url: `${SITE_URL}/` },
                { name: 'Microsoft app IDs', url: `${SITE_URL}/microsoft-apps.html` },
                { name: app.title, url: appUrl }
            ],
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

    log.info(`Built static site to ${outputRoot}`);
    log.info(`Snapshot: ${normalized.snapshotId}`);
    log.info(`Permissions: ${normalized.stats.permissions}`);
    log.info(`Apps: ${normalized.stats.apps}`);
    log.info(
        `AI discovery: llms.txt ${Math.round(llmsSizes.llmsBytes / 1024)} KB, llms-full.txt ${Math.round(llmsSizes.llmsFullBytes / 1024)} KB`
    );
}

const command = {
    name: 'build',
    summary: 'Render the static site and public JSON contracts from a snapshot.',
    usage: 'build [options]',
    options: [
        {
            name: 'input',
            type: 'string',
            description: 'Path to the normalized site-data.json.',
            default: '.generated/normalized/site-data.json'
        },
        {
            name: 'output',
            type: 'string',
            description: 'Output directory for the generated site.',
            default: 'docs'
        }
    ],
    run(args) {
        const input = path.resolve(args.input || DEFAULT_INPUT);
        const output = path.resolve(args.output || DEFAULT_OUTPUT);
        return buildSite(input, output);
    }
};

if (require.main === module) {
    runCommand(command);
}

module.exports = {
    buildSite,
    command,
    runCli: command.run
};
