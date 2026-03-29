const fs = require('fs');
const path = require('path');

const {
    escapeHtml,
    fileDate,
    formatUtcLabel,
    hashContent,
    loadJson,
    loadJsonWithShards,
    latestJsonDate,
    readUtf8,
    slugify,
    writeJson
} = require('./utils');
const { getRelationships, resolveResourceName } = require('../config/resource-mapping');

const SITE_URL = 'https://permissions.cengizyilmaz.net';
const SITE_NAME = 'Graph Permissions Explorer';
const DEFAULT_SCHEMA_VERSION = 2;
const DEFAULT_APP_CHUNK_SIZE = 200;

const RAW_FILES = {
    appRoles: 'GraphAppRoles.json',
    delegateRoles: 'GraphDelegateRoles.json',
    microsoftApps: 'MicrosoftApps.json',
    permissionMethods: 'GraphPermissionMethods.json',
    permissionPowerShell: 'GraphPermissionPowerShell.json',
    permissionCodeExamples: 'GraphPermissionCodeExamples.json',
    permissionCsv: 'permission.csv',
    resourceSchemas: 'GraphResourceSchemas.json',
    resourceDocs: 'GraphResourceDocumentation.json'
};

const ADMIN_CONSENT_CATEGORIES = [
    'Application', 'AuditLog', 'Directory', 'Domain', 'IdentityRisk', 'Policy',
    'RoleManagement', 'SecurityEvents', 'ThreatIndicators', 'Organization',
    'DeviceManagement', 'PrivilegedAccess', 'CrossTenantAccess'
];

const ACCESS_LEVEL_LABELS = {
    read: 'Read',
    write: 'Write',
    readwrite: 'Read/Write',
    full: 'Full Control'
};

const ACCESS_LEVEL_TEXT = {
    read: 'Read-only access to resources',
    write: 'Write access to resources',
    readwrite: 'Read and write access to resources',
    full: 'Full control over resources'
};

const SCOPE_LABELS = {
    all: 'All Resources',
    owned: 'Owned Only',
    shared: 'Shared Resources',
    user: 'User Scope'
};

const SCOPE_TEXT = {
    all: 'Access to all resources in the organization',
    owned: 'Access only to resources owned by the app',
    shared: 'Access to resources shared with the user',
    user: 'Access to current user resources'
};

const SOURCE_LABELS = {
    graph: 'Microsoft Graph',
    entradocs: 'Entra Docs',
    learn: 'Microsoft Learn',
    custom: 'Community'
};

const SOURCE_DOCS = {
    graph: {
        label: 'Microsoft Graph',
        url: 'https://graph.microsoft.com'
    },
    entradocs: {
        label: 'Entra Docs known GUID catalog',
        url: 'https://github.com/MicrosoftDocs/entra-docs'
    },
    learn: {
        label: 'Microsoft Learn first-party app article',
        url: 'https://learn.microsoft.com/troubleshoot/azure/active-directory/verify-first-party-apps-sign-in'
    },
    custom: {
        label: 'Community maintained app list',
        url: 'https://github.com/cengizyilmaz1/Permissions'
    }
};

const schemaIndexCache = new WeakMap();
const resourceDocsIndexCache = new WeakMap();
const primitivePropertyTypes = new Set([
    'string',
    'boolean',
    'int16',
    'int32',
    'int64',
    'integer',
    'double',
    'single',
    'float',
    'decimal',
    'guid',
    'date',
    'date-time',
    'time',
    'duration',
    'binary',
    'stream',
    'base64url'
]);

function ensureRequiredRawFiles(rawDir) {
    const missing = Object.values(RAW_FILES)
        .filter((filename) => {
            const filePath = path.join(rawDir, filename);
            return filename === RAW_FILES.permissionCodeExamples
                ? !loadJsonWithShards(filePath, null)
                : !fs.existsSync(filePath);
        })
        .map((filename) => path.join(rawDir, filename));

    if (missing.length > 0) {
        throw new Error(`Missing required raw input files: ${missing.join(', ')}`);
    }
}

function parsePermissionValue(value) {
    const parts = value.split('.');
    const resource = parts[0] || 'Other';
    const action = parts.slice(1).join('.') || value;
    const category = resource;
    const lowerValue = value.toLowerCase();

    let accessLevel = 'read';
    if (lowerValue.includes('readwrite')) {
        accessLevel = 'readwrite';
    } else if (lowerValue.includes('write') && !lowerValue.includes('read')) {
        accessLevel = 'write';
    } else if (lowerValue.includes('manage') || lowerValue.includes('full') || lowerValue.includes('control')) {
        accessLevel = 'full';
    }

    let scope = 'user';
    if (lowerValue.includes('.all')) {
        scope = 'all';
    } else if (lowerValue.includes('ownedby') || lowerValue.includes('owned') || lowerValue.includes('appowned')) {
        scope = 'owned';
    } else if (lowerValue.includes('shared')) {
        scope = 'shared';
    }

    const requiresAdmin = ADMIN_CONSENT_CATEGORIES.some((item) =>
        resource.toLowerCase().includes(item.toLowerCase())
    ) || scope === 'all';

    return { resource, action, category, accessLevel, scope, requiresAdmin };
}

function normalizeArray(value, label) {
    if (!Array.isArray(value)) {
        throw new Error(`${label} must be a JSON array.`);
    }

    return value;
}

function loadPermissionCsv(filePath) {
    const content = readUtf8(filePath);
    const lines = content.split(/\r?\n/);
    const byPermission = {};
    const byCategory = {};

    for (let index = 1; index < lines.length; index += 1) {
        const line = lines[index].trim();
        if (!line) {
            continue;
        }

        const match = line.match(/^"([^"]+)","([^"]+)","([^"]+)","([^"]+)"$/);
        if (!match) {
            continue;
        }

        const [, permission, method, endpoint, apiVersion] = match;
        const permissionKey = permission.toLowerCase();
        const categoryKey = permission.split('.')[0].toLowerCase();
        const entry = { method, endpoint };

        if (!byPermission[permissionKey]) {
            byPermission[permissionKey] = { 'v1.0': [], beta: [] };
        }

        if (!byCategory[categoryKey]) {
            byCategory[categoryKey] = { 'v1.0': [], beta: [] };
        }

        if (!byPermission[permissionKey][apiVersion].some((item) => item.method === method && item.endpoint === endpoint)) {
            byPermission[permissionKey][apiVersion].push(entry);
        }

        if (!byCategory[categoryKey][apiVersion].some((item) => item.method === method && item.endpoint === endpoint)) {
            byCategory[categoryKey][apiVersion].push(entry);
        }
    }

    return {
        byPermission,
        byCategory,
        permissionCount: Object.keys(byPermission).length,
        categoryCount: Object.keys(byCategory).length
    };
}

function getAppAnchor(app) {
    return `${slugify(app.title || app.appDisplayName || 'microsoft-app')}-${(app.appId || 'unknown-app').toLowerCase()}`;
}

function getAppDetailPath(app) {
    return `apps/${app.anchor}.html`;
}

function getAppPortalUrl(app) {
    return `https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/${app.appId}`;
}

function getAppSourceDescription(app) {
    if (app.isCommunity) {
        return 'This entry comes from the community-maintained list and is shown separately from official Microsoft sources.';
    }

    if (Array.isArray(app.sourceProvenanceLabels) && app.sourceProvenanceLabels.length > 1) {
        return `This app ID is corroborated by ${app.sourceProvenanceLabels.join(', ')} and is published with the highest-priority source shown first.`;
    }

    switch (app.source) {
    case 'graph':
        return 'This app ID was collected directly from Microsoft Graph service principal data.';
    case 'entradocs':
        return 'This app ID was collected from the Entra documentation known GUID catalog.';
    case 'learn':
        return 'This app ID was collected from a Microsoft Learn article about verifying first-party applications.';
    default:
        return 'This app ID is published as part of the Graph Permissions Explorer reference dataset.';
    }
}

function getAppSourceDoc(app) {
    const preferredSources = Array.isArray(app.sourceProvenance) && app.sourceProvenance.length > 0
        ? app.sourceProvenance
        : [app.source];

    for (const source of preferredSources) {
        if (SOURCE_DOCS[source]) {
            return SOURCE_DOCS[source];
        }
    }

    return SOURCE_DOCS.custom;
}

function normalizeApps(apps, sourceUpdatedAt) {
    return apps
        .map((item) => {
            const rawSource = String(item.Source || 'Unknown').trim();
            const source = rawSource.toLowerCase();
            const rawSources = Array.isArray(item.Sources) && item.Sources.length > 0 ? item.Sources : [rawSource];
            const sourceProvenance = Array.from(new Set(rawSources
                .map((entry) => String(entry || '').trim().toLowerCase())
                .filter(Boolean)));
            const hasOfficialSource = sourceProvenance.some((entry) => entry !== 'custom');
            const hasCommunitySource = sourceProvenance.includes('custom');
            const isCommunity = !hasOfficialSource && hasCommunitySource;
            const sourceLabel = SOURCE_LABELS[source] || rawSource;
            const sourceProvenanceLabels = sourceProvenance.map((entry) => SOURCE_LABELS[entry] || entry);
            const sourceDisplayLabel = sourceProvenanceLabels.join(' + ') || sourceLabel;
            const rawTitle = String(item.AppDisplayName || 'Unknown Microsoft App').trim();
            const title = rawTitle.replace(/\s*\[Community Contributed\]\s*$/i, '').trim() || rawTitle;
            const appId = String(item.AppId || '').toLowerCase();
            const filterGroups = Array.from(new Set(sourceProvenance.map((entry) => entry === 'custom' ? 'community' : entry)));

            return {
                appId,
                title,
                rawTitle,
                source,
                sourceLabel,
                sourceDisplayLabel,
                sourceProvenance,
                sourceProvenanceLabels,
                sourceGroup: isCommunity ? 'community' : 'official',
                filterGroup: isCommunity ? 'community' : source,
                filterGroups,
                isOfficial: hasOfficialSource,
                isCommunity,
                ownerOrganizationId: item.AppOwnerOrganizationId || '',
                sourceUpdatedAt,
                anchor: getAppAnchor({ title, appId })
            };
        })
        .filter((item) => item.appId)
        .sort((left, right) => left.title.localeCompare(right.title));
}

function normalizePermissions(appRoles, delegateRoles) {
    const permissionsMap = new Map();

    for (const item of appRoles) {
        const value = item.Value || '';
        if (!value) {
            continue;
        }

        const existing = permissionsMap.get(value) || {
            value,
            slug: slugify(value),
            description: '',
            ...parsePermissionValue(value),
            application: null,
            delegated: null
        };

        existing.application = {
            id: item.Id || '',
            displayName: item.DisplayName || '',
            description: item.Description || '',
            isEnabled: item.IsEnabled !== false
        };

        existing.description = existing.description || item.Description || '';
        permissionsMap.set(value, existing);
    }

    for (const item of delegateRoles) {
        const value = item.Value || '';
        if (!value) {
            continue;
        }

        const existing = permissionsMap.get(value) || {
            value,
            slug: slugify(value),
            description: '',
            ...parsePermissionValue(value),
            application: null,
            delegated: null
        };

        existing.delegated = {
            id: item.Id || '',
            displayName: item.AdminConsentDisplayName || item.UserConsentDisplayName || '',
            description: item.AdminConsentDescription || item.UserConsentDescription || '',
            userConsentDisplayName: item.UserConsentDisplayName || '',
            userConsentDescription: item.UserConsentDescription || '',
            type: item.Type || 'Admin',
            isEnabled: item.IsEnabled !== false
        };

        existing.description = existing.description || item.AdminConsentDescription || item.UserConsentDescription || '';
        permissionsMap.set(value, existing);
    }

    return Array.from(permissionsMap.values())
        .map((item) => ({
            ...item,
            description: item.description || item.application?.description || item.delegated?.description || ''
        }))
        .sort((left, right) => left.value.localeCompare(right.value));
}

function buildSourceFreshness(rawDir) {
    const rawPath = (filename) => path.join(rawDir, filename);

    return {
        graphPermissions: {
            source: 'Microsoft Graph service principal',
            updatedAt: new Date(Math.max(
                fs.statSync(rawPath(RAW_FILES.appRoles)).mtimeMs,
                fs.statSync(rawPath(RAW_FILES.delegateRoles)).mtimeMs
            )).toISOString()
        },
        microsoftApps: {
            source: 'Microsoft Graph + Entra Docs + Microsoft Learn + Community',
            updatedAt: fileDate(rawPath(RAW_FILES.microsoftApps))
        },
        learnPermissionMethods: {
            source: 'Microsoft Learn API permissions tables',
            updatedAt: fileDate(rawPath(RAW_FILES.permissionMethods))
        },
        learnPowerShell: {
            source: 'Microsoft Learn PowerShell snippets',
            updatedAt: fileDate(rawPath(RAW_FILES.permissionPowerShell))
        },
        learnCodeExamples: {
            source: 'Microsoft Learn SDK code snippets',
            updatedAt: latestJsonDate(rawPath(RAW_FILES.permissionCodeExamples)) || fileDate(rawPath(RAW_FILES.permissionCodeExamples))
        },
        learnResourceDocs: {
            source: 'Microsoft Learn resource type documentation',
            updatedAt: fileDate(rawPath(RAW_FILES.resourceDocs))
        },
        openApiMethods: {
            source: 'Microsoft Graph OpenAPI metadata (resource-family fallback)',
            updatedAt: fileDate(rawPath(RAW_FILES.permissionCsv))
        },
        openApiSchemas: {
            source: 'Microsoft Graph OpenAPI metadata',
            updatedAt: fileDate(rawPath(RAW_FILES.resourceSchemas))
        }
    };
}

function getMethodContext(permission, permissionMethods, permissionCsv) {
    const exact = permissionMethods.byPermission[permission.value.toLowerCase()];
    if (exact) {
        return {
            versionedMethods: {
                'v1.0': (exact.v1 || []).map((item) => ({
                    method: item.method,
                    endpoint: item.path,
                    docLink: item.docLink || '',
                    supportsDelegated: Boolean(item.supportsDelegated),
                    supportsApplication: Boolean(item.supportsApplication),
                    isLeastPrivilege: Boolean(item.isLeastPrivilege)
                })),
                beta: (exact.beta || []).map((item) => ({
                    method: item.method,
                    endpoint: item.path,
                    docLink: item.docLink || '',
                    supportsDelegated: Boolean(item.supportsDelegated),
                    supportsApplication: Boolean(item.supportsApplication),
                    isLeastPrivilege: Boolean(item.isLeastPrivilege)
                }))
            },
            confidence: 'exact',
            label: 'Exact Microsoft Learn match',
            sourceKey: 'learnPermissionMethods',
            sourceKeys: ['learnPermissionMethods']
        };
    }

    const family = permissionCsv.byCategory[permission.category.toLowerCase()];
    if (family) {
        return {
            versionedMethods: family,
            confidence: 'category-derived',
            label: `${permission.category} resource-family fallback`,
            sourceKey: 'openApiMethods',
            sourceKeys: ['learnPermissionMethods', 'openApiMethods']
        };
    }

    return {
        versionedMethods: { 'v1.0': [], beta: [] },
        confidence: 'unavailable',
        label: 'No Learn or OpenAPI mapping available',
        sourceKey: 'learnPermissionMethods',
        sourceKeys: ['learnPermissionMethods', 'openApiMethods']
    };
}

function getPowerShellContext(permission, permissionPowerShell) {
    const exact = permissionPowerShell.byPermission[permission.value.toLowerCase()];
    if (exact) {
        return {
            versionedCommands: {
                v1: (exact.v1 || []).map((item) => ({
                    command: item.command,
                    endpoint: item.endpoint || '',
                    title: item.title || '',
                    docLink: item.docLink || '',
                    supportsDelegated: Boolean(item.supportsDelegated),
                    supportsApplication: Boolean(item.supportsApplication),
                    isLeastPrivilege: Boolean(item.isLeastPrivilege),
                    code: item.code || ''
                })),
                beta: (exact.beta || []).map((item) => ({
                    command: item.command,
                    endpoint: item.endpoint || '',
                    title: item.title || '',
                    docLink: item.docLink || '',
                    supportsDelegated: Boolean(item.supportsDelegated),
                    supportsApplication: Boolean(item.supportsApplication),
                    isLeastPrivilege: Boolean(item.isLeastPrivilege),
                    code: item.code || ''
                }))
            },
            confidence: 'exact',
            label: 'Exact Microsoft Learn PowerShell match',
            sourceKey: 'learnPowerShell',
            sourceKeys: ['learnPowerShell']
        };
    }

    return {
        versionedCommands: { v1: [], beta: [] },
        confidence: 'unavailable',
        label: 'No Microsoft Learn PowerShell mapping available',
        sourceKey: 'learnPowerShell',
        sourceKeys: ['learnPowerShell']
    };
}

function getCodeExampleContext(permission, permissionCodeExamples) {
    const exact = permissionCodeExamples.byPermission[permission.value.toLowerCase()];
    if (exact) {
        return {
            versions: {
                v1: exact.v1 || { csharp: [], javascript: [], powershell: [], python: [] },
                beta: exact.beta || { csharp: [], javascript: [], powershell: [], python: [] }
            },
            confidence: 'exact',
            label: 'Exact Microsoft Learn code snippet match',
            sourceKey: 'learnCodeExamples',
            sourceKeys: ['learnCodeExamples']
        };
    }

    return {
        versions: {
            v1: { csharp: [], javascript: [], powershell: [], python: [] },
            beta: { csharp: [], javascript: [], powershell: [], python: [] }
        },
        confidence: 'unavailable',
        label: 'No Microsoft Learn code snippets available',
        sourceKey: 'learnCodeExamples',
        sourceKeys: ['learnCodeExamples']
    };
}

function normalizeSchemaLookupKey(value) {
    return String(value || '')
        .replace(/^microsoft\.graph\./i, '')
        .replace(/collection$/i, '')
        .replace(/[^a-z0-9]/gi, '')
        .toLowerCase();
}

function getSchemaIndex(resourceSchemas) {
    if (schemaIndexCache.has(resourceSchemas)) {
        return schemaIndexCache.get(resourceSchemas);
    }

    const index = new Map();
    Object.entries(resourceSchemas).forEach(([key, value]) => {
        index.set(normalizeSchemaLookupKey(key), { key, value });
    });
    schemaIndexCache.set(resourceSchemas, index);
    return index;
}

function getResourceDocsIndex(resourceDocs) {
    if (resourceDocsIndexCache.has(resourceDocs)) {
        return resourceDocsIndexCache.get(resourceDocs);
    }

    const index = new Map();
    Object.entries(resourceDocs || {}).forEach(([key, value]) => {
        index.set(normalizeSchemaLookupKey(key), { key, value });
    });
    resourceDocsIndexCache.set(resourceDocs, index);
    return index;
}

function singularizeResourceSegment(segment) {
    const value = String(segment || '')
        .replace(/^microsoft\.graph\./i, '')
        .replace(/\(.*\)$/g, '')
        .trim();

    if (!value) {
        return '';
    }

    if (/ies$/i.test(value)) {
        return value.replace(/ies$/i, 'y');
    }

    if (/sses$/i.test(value)) {
        return value.replace(/es$/i, '');
    }

    if (/ses$/i.test(value) && !/sses$/i.test(value)) {
        return value.replace(/es$/i, '');
    }

    if (/s$/i.test(value) && !/(ss|us|is)$/i.test(value)) {
        return value.slice(0, -1);
    }

    return value;
}

function toPascalCaseToken(value) {
    return String(value || '')
        .split(/[^a-z0-9]+/i)
        .filter(Boolean)
        .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
        .join('');
}

function toCamelCaseToken(value) {
    const pascal = toPascalCaseToken(value);
    return pascal ? `${pascal.charAt(0).toLowerCase()}${pascal.slice(1)}` : '';
}

function getEndpointResourceCandidates(methodContext) {
    const candidates = [];
    const seen = new Set();
    const endpoints = [
        ...(methodContext.versionedMethods['v1.0'] || []),
        ...(methodContext.versionedMethods.beta || [])
    ]
        .map((item) => item.endpoint)
        .filter(Boolean);

    const addCandidate = (value) => {
        if (!value) {
            return;
        }

        const normalized = String(value).trim();
        const key = normalizeSchemaLookupKey(normalized);
        if (!key || seen.has(key)) {
            return;
        }

        seen.add(key);
        candidates.push(normalized);
    };

    endpoints.forEach((endpoint) => {
        const cleanPath = endpoint.split('?')[0].replace(/^\/+/, '');
        const segments = cleanPath
            .split('/')
            .map((segment) => segment.replace(/\(.*?\)/g, '').trim())
            .filter((segment) => segment && !segment.startsWith('{') && !segment.startsWith('$') && segment.toLowerCase() !== 'me');

        const singularSegments = segments.map((segment) => singularizeResourceSegment(segment));

        singularSegments.forEach(addCandidate);

        for (let index = 1; index < singularSegments.length; index += 1) {
            addCandidate(`${singularSegments[index - 1]}${toPascalCaseToken(singularSegments[index])}`);
        }

        const lastSegment = singularSegments[singularSegments.length - 1];
        const penultimateSegment = singularSegments[singularSegments.length - 2];

        addCandidate(lastSegment);
        if (penultimateSegment && lastSegment) {
            addCandidate(`${penultimateSegment}${toPascalCaseToken(lastSegment)}`);
        }
    });

    return candidates;
}

function findSchemaMatch(resourceSchemas, candidates) {
    const schemaIndex = getSchemaIndex(resourceSchemas);

    for (const candidate of candidates) {
        const match = schemaIndex.get(normalizeSchemaLookupKey(candidate));
        if (match) {
            return match;
        }
    }

    return null;
}

function findResourceDocMatch(resourceDocs, candidates) {
    const docsIndex = getResourceDocsIndex(resourceDocs);

    for (const candidate of candidates) {
        const match = docsIndex.get(normalizeSchemaLookupKey(candidate));
        if (match) {
            return match;
        }
    }

    return null;
}

function selectResourceMetadata(permission, resourceSchemas, resourceDocs, methodContext) {
    const categoryKey = permission.category.toLowerCase();
    const resourceName = resolveResourceName(categoryKey);
    const categoryCandidates = [permission.category, resourceName, categoryKey];
    const categorySchemaMatch = findSchemaMatch(resourceSchemas, categoryCandidates);
    const categoryDocMatch = findResourceDocMatch(resourceDocs, categoryCandidates);

    if (categorySchemaMatch || categoryDocMatch) {
        const selected = categoryDocMatch || categorySchemaMatch;
        return {
            resourceName: selected.key,
            schema: categorySchemaMatch?.value || null,
            docs: categoryDocMatch?.value || null,
            confidence: normalizeSchemaLookupKey(permission.category) === normalizeSchemaLookupKey(selected.key)
                ? 'exact-category'
                : 'mapped'
        };
    }

    const endpointCandidates = getEndpointResourceCandidates(methodContext);
    const endpointSchemaMatch = findSchemaMatch(resourceSchemas, endpointCandidates);
    const endpointDocMatch = findResourceDocMatch(resourceDocs, endpointCandidates);

    if (endpointSchemaMatch || endpointDocMatch) {
        const selected = endpointDocMatch || endpointSchemaMatch;
        return {
            resourceName: selected.key,
            schema: endpointSchemaMatch?.value || null,
            docs: endpointDocMatch?.value || null,
            confidence: 'endpoint-derived'
        };
    }

    return {
        resourceName,
        schema: null,
        docs: null,
        confidence: 'unavailable'
    };
}

function cleanProperties(properties) {
    if (!Array.isArray(properties)) {
        return [];
    }

    const seen = new Set();
    const result = [];

    for (const property of properties) {
        if (!property || typeof property.name !== 'string' || property.name.startsWith('@odata')) {
            continue;
        }

        if (seen.has(property.name)) {
            continue;
        }

        seen.add(property.name);
        result.push({
            name: property.name,
            type: String(property.type || 'string').replace(/microsoft\.graph\./gi, ''),
            description: property.description || '',
            readOnly: Boolean(property.readOnly),
            nullable: Boolean(property.nullable)
        });
    }

    return result;
}

function cleanDocItems(items) {
    if (!Array.isArray(items)) {
        return [];
    }

    const seen = new Set();
    const result = [];

    items.forEach((item) => {
        if (!item || typeof item.name !== 'string') {
            return;
        }

        const name = item.name.trim();
        if (!name || seen.has(name)) {
            return;
        }

        seen.add(name);
        result.push({
            name,
            type: String(item.type || 'string').replace(/microsoft\.graph\./gi, ''),
            description: item.description || '',
            readOnly: Boolean(item.readOnly),
            nullable: Boolean(item.nullable)
        });
    });

    return result;
}

function mergeNamedItems(primaryItems, secondaryItems) {
    const merged = new Map();

    [...(primaryItems || []), ...(secondaryItems || [])].forEach((item) => {
        if (!item?.name) {
            return;
        }

        const existing = merged.get(item.name);
        if (!existing) {
            merged.set(item.name, { ...item });
            return;
        }

        merged.set(item.name, {
            ...existing,
            type: existing.type || item.type,
            description: existing.description || item.description,
            readOnly: existing.readOnly || item.readOnly,
            nullable: existing.nullable || item.nullable
        });
    });

    return Array.from(merged.values());
}

function getDocsVersionBundle(resourceSelection) {
    if (!resourceSelection.docs) {
        return { version: null, bundle: null };
    }

    if (resourceSelection.docs.v1 && (
        resourceSelection.docs.v1.properties?.length
        || resourceSelection.docs.v1.relationships?.length
        || resourceSelection.docs.v1.jsonRepresentation
    )) {
        return { version: 'v1', bundle: resourceSelection.docs.v1 };
    }

    if (resourceSelection.docs.beta && (
        resourceSelection.docs.beta.properties?.length
        || resourceSelection.docs.beta.relationships?.length
        || resourceSelection.docs.beta.jsonRepresentation
    )) {
        return { version: 'beta', bundle: resourceSelection.docs.beta };
    }

    return { version: null, bundle: null };
}

function buildPropertiesData(resourceSelection) {
    const schema = resourceSelection.schema;
    const docsVersion = getDocsVersionBundle(resourceSelection);
    const docsItems = cleanDocItems(docsVersion.bundle?.properties || []);

    if ((!schema || !schema.properties) && docsItems.length === 0) {
        return {
            resourceName: resourceSelection.resourceName,
            confidence: resourceSelection.confidence,
            version: null,
            source: 'unavailable',
            docLink: docsVersion.bundle?.docLink || '',
            items: []
        };
    }

    const v1Items = cleanProperties(schema?.properties?.['v1.0']);
    const betaItems = cleanProperties(schema?.properties?.beta);
    const schemaVersion = v1Items.length > 0 ? 'v1' : (betaItems.length > 0 ? 'beta' : null);
    const schemaItems = schemaVersion === 'v1' ? v1Items : betaItems;
    const version = docsVersion.version || schemaVersion;
    const items = version === docsVersion.version
        ? mergeNamedItems(docsItems, schemaItems)
        : schemaItems;

    return {
        resourceName: resourceSelection.resourceName,
        confidence: docsVersion.version ? `${resourceSelection.confidence}-docs` : resourceSelection.confidence,
        version,
        source: docsVersion.version ? (schemaItems.length > 0 ? 'learn+openapi' : 'learn') : 'openapi',
        docLink: docsVersion.bundle?.docLink || '',
        items
    };
}

function parseJsonRepresentationValue(rawValue) {
    if (!rawValue) {
        return null;
    }

    try {
        return JSON.parse(rawValue);
    } catch {
        return null;
    }
}

function buildJsonRepresentation(propertiesData, resourceSelection) {
    const docsVersion = getDocsVersionBundle(resourceSelection);
    const docsJson = docsVersion.bundle?.jsonRepresentation || '';
    const parsedDocsJson = parseJsonRepresentationValue(docsJson);

    if (docsJson) {
        return {
            resourceName: propertiesData.resourceName,
            confidence: `${resourceSelection.confidence}-docs`,
            version: docsVersion.version,
            source: 'learn',
            docLink: docsVersion.bundle?.docLink || '',
            raw: docsJson,
            value: parsedDocsJson
        };
    }

    if (!propertiesData.version || propertiesData.items.length === 0) {
        return {
            resourceName: propertiesData.resourceName,
            confidence: propertiesData.confidence,
            version: null,
            source: 'unavailable',
            docLink: docsVersion.bundle?.docLink || '',
            raw: null,
            value: null
        };
    }

    const jsonValue = {};
    for (const property of propertiesData.items.slice(0, 20)) {
        jsonValue[property.name] = buildJsonSampleValue(property.type);
    }

    return {
        resourceName: propertiesData.resourceName,
        confidence: propertiesData.confidence,
        version: propertiesData.version,
        source: propertiesData.source === 'learn+openapi' ? 'learn+generated' : 'generated',
        docLink: propertiesData.docLink || '',
        raw: null,
        value: jsonValue
    };
}

function buildJsonSampleValue(typeName) {
    const type = String(typeName || 'string').replace(/microsoft\.graph\./gi, '').trim();
    const lowerType = type.toLowerCase();

    if (lowerType.includes('collection')) {
        const collectionType = type.replace(/\s+collection$/i, '');
        if (!collectionType || collectionType.toLowerCase() === lowerType) {
            return ['...'];
        }
        return [buildJsonSampleValue(collectionType)];
    }

    if (primitivePropertyTypes.has(lowerType)) {
        if (lowerType === 'boolean') {
            return true;
        }

        if (['int16', 'int32', 'int64', 'integer', 'double', 'single', 'float', 'decimal'].includes(lowerType)) {
            return 0;
        }

        if (lowerType === 'guid') {
            return '00000000-0000-0000-0000-000000000000';
        }

        if (lowerType === 'date') {
            return '2026-01-01';
        }

        if (lowerType === 'date-time') {
            return '2026-01-01T00:00:00Z';
        }

        return 'String';
    }

    if (lowerType.includes('identityset')) {
        return {
            user: {
                id: '00000000-0000-0000-0000-000000000000',
                displayName: 'Adele Vance'
            }
        };
    }

    if (lowerType.includes('emailaddress')) {
        return {
            name: 'Adele Vance',
            address: 'adele.vance@contoso.com'
        };
    }

    if (lowerType.includes('recipient')) {
        return {
            emailAddress: buildJsonSampleValue('emailAddress')
        };
    }

    if (lowerType.includes('itembody')) {
        return {
            contentType: 'text',
            content: 'Sample content'
        };
    }

    if (lowerType.includes('datetimetimezone')) {
        return {
            dateTime: '2026-01-01T00:00:00',
            timeZone: 'UTC'
        };
    }

    if (lowerType.includes('location')) {
        return {
            displayName: 'Conference Room',
            locationType: 'conferenceRoom'
        };
    }

    if (lowerType === 'object') {
        return {
            sample: 'value'
        };
    }

    return {
        '@type': type,
        id: '00000000-0000-0000-0000-000000000000'
    };
}

function deriveRelationshipsFromSchema(schemaSelection) {
    const schema = schemaSelection.schema;
    if (!schema?.properties) {
        return [];
    }

    const sourceItems = [
        ...cleanProperties(schema.properties['v1.0']),
        ...cleanProperties(schema.properties.beta)
    ];
    const seen = new Set();
    const items = [];

    sourceItems.forEach((property) => {
        const type = String(property.type || '').replace(/microsoft\.graph\./gi, '');
        const lowerType = type.toLowerCase();
        const isCollection = lowerType.includes('collection');
        const isComplex = !primitivePropertyTypes.has(lowerType) && lowerType !== 'object';
        const looksRelational = isCollection || isComplex || /(^|[^a-z])(members?|owners?|children|messages|events|drives|lists|tabs|channels|appointments|instances|decisions|users|groups|permissions|assignments|photos?)$/i.test(property.name);

        if (!looksRelational || seen.has(property.name)) {
            return;
        }

        seen.add(property.name);
        items.push({
            name: property.name,
            type,
            description: property.description || `Related ${property.name} data exposed by this resource.`
        });
    });

    return items.slice(0, 20);
}

function buildRelationshipsData(permission, resourceSelection) {
    const docsVersion = getDocsVersionBundle(resourceSelection);
    const docsItems = cleanDocItems(docsVersion.bundle?.relationships || []);
    const configuredItems = getRelationships(permission.category.toLowerCase());
    const derivedItems = deriveRelationshipsFromSchema(resourceSelection);
    const merged = [];
    const seen = new Set();

    [...docsItems, ...configuredItems, ...derivedItems].forEach((item) => {
        if (!item || !item.name || seen.has(item.name)) {
            return;
        }

        seen.add(item.name);
        merged.push(item);
    });

    return {
        resourceName: resourceSelection.resourceName,
        confidence: merged.length > 0
            ? (docsItems.length > 0
                ? `${resourceSelection.confidence}-docs`
                : (configuredItems.length > 0 ? resourceSelection.confidence : 'schema-derived'))
            : 'unavailable',
        source: docsItems.length > 0
            ? (configuredItems.length > 0 || derivedItems.length > 0 ? 'learn+openapi' : 'learn')
            : (configuredItems.length > 0 || derivedItems.length > 0 ? 'openapi' : 'unavailable'),
        docLink: docsVersion.bundle?.docLink || '',
        items: merged
    };
}

function dedupeMethods(methods) {
    const merged = new Map();

    for (const item of methods || []) {
        const key = `${item.method}|${item.endpoint}`;
        const existing = merged.get(key);
        if (!existing) {
            merged.set(key, {
                method: item.method,
                endpoint: item.endpoint,
                docLink: item.docLink || '',
                supportsDelegated: Boolean(item.supportsDelegated),
                supportsApplication: Boolean(item.supportsApplication),
                isLeastPrivilege: Boolean(item.isLeastPrivilege)
            });
            continue;
        }

        existing.supportsDelegated = existing.supportsDelegated || Boolean(item.supportsDelegated);
        existing.supportsApplication = existing.supportsApplication || Boolean(item.supportsApplication);
        existing.isLeastPrivilege = existing.isLeastPrivilege || Boolean(item.isLeastPrivilege);
        if (!existing.docLink && item.docLink) {
            existing.docLink = item.docLink;
        }
    }

    return Array.from(merged.values());
}

function dedupePowerShellCommands(commands) {
    const merged = new Map();

    for (const item of commands || []) {
        const key = `${item.command}|${item.endpoint}|${item.docLink}`;
        const existing = merged.get(key);
        if (!existing) {
            merged.set(key, {
                command: item.command,
                endpoint: item.endpoint || '',
                title: item.title || '',
                docLink: item.docLink || '',
                supportsDelegated: Boolean(item.supportsDelegated),
                supportsApplication: Boolean(item.supportsApplication),
                isLeastPrivilege: Boolean(item.isLeastPrivilege),
                code: item.code || ''
            });
            continue;
        }

        existing.supportsDelegated = existing.supportsDelegated || Boolean(item.supportsDelegated);
        existing.supportsApplication = existing.supportsApplication || Boolean(item.supportsApplication);
        existing.isLeastPrivilege = existing.isLeastPrivilege || Boolean(item.isLeastPrivilege);
        if (!existing.title && item.title) {
            existing.title = item.title;
        }
        if (!existing.code && item.code) {
            existing.code = item.code;
        }
    }

    return Array.from(merged.values());
}

function buildSampleEndpoint(permission, methodContext) {
    const allMethods = [
        ...(methodContext.versionedMethods['v1.0'] || []),
        ...(methodContext.versionedMethods.beta || [])
    ];

    const preferred = allMethods.find((item) => item.method === 'GET') || allMethods[0];
    if (preferred) {
        return preferred.endpoint.replace(/{[^}]+}/g, '{id}');
    }

    const resource = resolveResourceName(permission.category.toLowerCase());
    if (resource === 'user') {
        return '/me';
    }

    return `/${resource}`;
}

function buildMethodSourceMetadata(methodContext, sourceFreshness) {
    const primarySource = sourceFreshness[methodContext.sourceKey];
    const secondarySource = methodContext.sourceKeys
        .filter((key) => key !== methodContext.sourceKey)
        .map((key) => sourceFreshness[key])
        .find(Boolean);

    return {
        source: primarySource?.source || '',
        sourceUpdatedAt: primarySource?.updatedAt || null,
        secondarySource: secondarySource?.source || null,
        secondarySourceUpdatedAt: secondarySource?.updatedAt || null,
        confidence: methodContext.confidence,
        label: methodContext.label
    };
}

function buildPowerShellSourceMetadata(powerShellContext, sourceFreshness) {
    const primarySource = sourceFreshness[powerShellContext.sourceKey];

    return {
        source: primarySource?.source || '',
        sourceUpdatedAt: primarySource?.updatedAt || null,
        confidence: powerShellContext.confidence,
        label: powerShellContext.label
    };
}

function buildCodeExampleSourceMetadata(codeExampleContext, sourceFreshness) {
    const primarySource = sourceFreshness[codeExampleContext.sourceKey];

    return {
        source: primarySource?.source || '',
        sourceUpdatedAt: primarySource?.updatedAt || null,
        confidence: codeExampleContext.confidence,
        label: codeExampleContext.label
    };
}

function buildSchemaSourceMetadata(resourceSelection, sourceFreshness) {
    const docsSource = sourceFreshness.learnResourceDocs;
    return {
        source: resourceSelection.docs ? docsSource.source : sourceFreshness.openApiSchemas.source,
        sourceUpdatedAt: resourceSelection.docs ? docsSource.updatedAt : sourceFreshness.openApiSchemas.updatedAt,
        secondarySource: resourceSelection.docs && resourceSelection.schema ? sourceFreshness.openApiSchemas.source : null,
        secondarySourceUpdatedAt: resourceSelection.docs && resourceSelection.schema ? sourceFreshness.openApiSchemas.updatedAt : null,
        confidence: resourceSelection.confidence,
        resourceName: resourceSelection.resourceName,
        hasLearnDocs: Boolean(resourceSelection.docs),
        hasOpenApiSchema: Boolean(resourceSelection.schema)
    };
}

function pickOfficialCodeExample(codeExampleContext, language) {
    const versions = [codeExampleContext.versions.v1, codeExampleContext.versions.beta];
    for (const version of versions) {
        const entry = (version?.[language] || [])[0];
        if (entry) {
            return entry;
        }
    }
    return null;
}

function buildOfficialCodeBlock(label, languageClass, entry) {
    const title = entry.title ? `<div class="method-meta">${escapeHtml(entry.title)}</div>` : '';
    const link = entry.docLink
        ? `<div class="method-meta"><a href="${entry.docLink}" target="_blank" rel="noopener">View official example on Microsoft Learn</a></div>`
        : '';

    return `<div class="code-block">
    <div class="code-header">
        <span>${label}</span>
        <button class="copy-btn" onclick="copyCode(this)">Copy</button>
    </div>
    ${title}${link}
    <pre><code class="language-${languageClass}">${escapeHtml(entry.code)}</code></pre>
</div>`;
}

function buildCodeExamples(permission, sampleEndpoint, powerShellContext, codeExampleContext) {
    const cleanEndpoint = sampleEndpoint.startsWith('/') ? sampleEndpoint : `/${sampleEndpoint}`;
    const escapedScope = escapeHtml(permission.value);
    const endpoint = escapeHtml(cleanEndpoint);
    const csharpExample = pickOfficialCodeExample(codeExampleContext, 'csharp');
    const javascriptExample = pickOfficialCodeExample(codeExampleContext, 'javascript');
    const pythonExample = pickOfficialCodeExample(codeExampleContext, 'python');
    const powershellExample = pickOfficialCodeExample(codeExampleContext, 'powershell')
        || [...(powerShellContext.versionedCommands.v1 || []), ...(powerShellContext.versionedCommands.beta || [])]
            .find((item) => item.code) || null;

    return {
        csharp: csharpExample
            ? buildOfficialCodeBlock('C# / .NET SDK', 'csharp', csharpExample)
            : `<div class="code-block">
    <div class="code-header">
        <span>C# / .NET SDK</span>
        <button class="copy-btn" onclick="copyCode(this)">Copy</button>
    </div>
    <pre><code class="language-csharp">using Azure.Identity;
using Microsoft.Graph;

var scopes = new[] { "${escapedScope}" };
var credential = new InteractiveBrowserCredential(
    new InteractiveBrowserCredentialOptions
    {
        ClientId = "YOUR_CLIENT_ID",
        TenantId = "YOUR_TENANT_ID",
        RedirectUri = new Uri("http://localhost")
    });

var graphClient = new GraphServiceClient(credential, scopes);
var response = await graphClient
    .WithUrl("https://graph.microsoft.com/v1.0${endpoint}")
    .GetAsync();</code></pre>
</div>`,
        javascript: javascriptExample
            ? buildOfficialCodeBlock('JavaScript', 'javascript', javascriptExample)
            : `<div class="code-block">
    <div class="code-header">
        <span>JavaScript</span>
        <button class="copy-btn" onclick="copyCode(this)">Copy</button>
    </div>
    <pre><code class="language-javascript">import { Client } from "@microsoft/microsoft-graph-client";
import { InteractiveBrowserCredential } from "@azure/identity";

const credential = new InteractiveBrowserCredential({
  clientId: "YOUR_CLIENT_ID",
  tenantId: "YOUR_TENANT_ID",
  redirectUri: "http://localhost"
});

const token = await credential.getToken(["${escapedScope}"]);
const client = Client.init({
  authProvider: (done) => done(null, token.token)
});

const response = await client.api("${endpoint}").get();</code></pre>
</div>`,
        powershell: powershellExample
            ? buildOfficialCodeBlock('PowerShell', 'powershell', powershellExample)
            : `<div class="code-block">
    <div class="code-header">
        <span>PowerShell</span>
        <button class="copy-btn" onclick="copyCode(this)">Copy</button>
    </div>
    <pre><code class="language-powershell">Connect-MgGraph -Scopes "${escapedScope}"
Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0${endpoint}"</code></pre>
</div>`,
        python: pythonExample
            ? buildOfficialCodeBlock('Python', 'python', pythonExample)
            : `<div class="code-block">
    <div class="code-header">
        <span>Python</span>
        <button class="copy-btn" onclick="copyCode(this)">Copy</button>
    </div>
    <pre><code class="language-python">from azure.identity import InteractiveBrowserCredential
import requests

credential = InteractiveBrowserCredential(
    client_id="YOUR_CLIENT_ID",
    tenant_id="YOUR_TENANT_ID"
)

token = credential.get_token("${escapedScope}")
response = requests.get(
    "https://graph.microsoft.com/v1.0${endpoint}",
    headers={"Authorization": f"Bearer {token.token}"}
)

print(response.json())</code></pre>
</div>`
    };
}

function buildPermissionSourceMetadata(permission, sourceFreshness, methodContext, powerShellContext, codeExampleContext, resourceSelection) {
    const ingestedFrom = new Set([
        sourceFreshness.graphPermissions.source
    ]);

    [...methodContext.sourceKeys, ...powerShellContext.sourceKeys, ...codeExampleContext.sourceKeys].forEach((key) => {
        if (sourceFreshness[key]?.source) {
            ingestedFrom.add(sourceFreshness[key].source);
        }
    });

    if (resourceSelection.docs && sourceFreshness.learnResourceDocs?.source) {
        ingestedFrom.add(sourceFreshness.learnResourceDocs.source);
    }

    if (resourceSelection.schema && sourceFreshness.openApiSchemas?.source) {
        ingestedFrom.add(sourceFreshness.openApiSchemas.source);
    }

    return {
        permission: {
            source: sourceFreshness.graphPermissions.source,
            sourceUpdatedAt: sourceFreshness.graphPermissions.updatedAt,
            isOfficial: true
        },
        methods: buildMethodSourceMetadata(methodContext, sourceFreshness),
        powershell: buildPowerShellSourceMetadata(powerShellContext, sourceFreshness),
        codeExamples: buildCodeExampleSourceMetadata(codeExampleContext, sourceFreshness),
        schemas: buildSchemaSourceMetadata(resourceSelection, sourceFreshness),
        page: {
            sourceUpdatedAt: sourceFreshness.graphPermissions.updatedAt,
            ingestedFrom: Array.from(ingestedFrom)
        }
    };
}

function buildPermissionDetail(permission, sourceFreshness, permissionMethods, permissionPowerShell, permissionCodeExamples, permissionCsv, resourceSchemas, resourceDocs) {
    const methodContext = getMethodContext(permission, permissionMethods, permissionCsv);
    const powerShellContext = getPowerShellContext(permission, permissionPowerShell);
    const codeExampleContext = getCodeExampleContext(permission, permissionCodeExamples);
    const resourceSelection = selectResourceMetadata(permission, resourceSchemas, resourceDocs, methodContext);
    const properties = buildPropertiesData(resourceSelection);
    const jsonRepresentation = buildJsonRepresentation(properties, resourceSelection);
    const relationships = buildRelationshipsData(permission, resourceSelection);
    const sampleEndpoint = buildSampleEndpoint(permission, methodContext);
    const sources = buildPermissionSourceMetadata(permission, sourceFreshness, methodContext, powerShellContext, codeExampleContext, resourceSelection);

    return {
        ...permission,
        hasApplication: Boolean(permission.application),
        hasDelegated: Boolean(permission.delegated),
        methods: {
            confidence: methodContext.confidence,
            label: methodContext.label,
            sourceKey: methodContext.sourceKey,
            api: {
                v1: dedupeMethods(methodContext.versionedMethods['v1.0']),
                beta: dedupeMethods(methodContext.versionedMethods.beta)
            },
            powershell: {
                confidence: powerShellContext.confidence,
                label: powerShellContext.label,
                sourceKey: powerShellContext.sourceKey,
                v1: dedupePowerShellCommands(powerShellContext.versionedCommands.v1),
                beta: dedupePowerShellCommands(powerShellContext.versionedCommands.beta)
            }
        },
        properties,
        jsonRepresentation,
        relationships,
        resource: {
            name: resourceSelection.resourceName,
            confidence: resourceSelection.confidence,
            docLink: properties.docLink || relationships.docLink || ''
        },
        sampleEndpoint,
        codeExamples: buildCodeExamples(permission, sampleEndpoint, powerShellContext, codeExampleContext),
        sources
    };
}

function groupPermissionsByCategory(permissions) {
    const grouped = {};

    for (const permission of permissions) {
        if (!grouped[permission.category]) {
            grouped[permission.category] = [];
        }

        grouped[permission.category].push(permission);
    }

    return Object.keys(grouped)
        .sort((left, right) => left.localeCompare(right))
        .reduce((accumulator, key) => {
            accumulator[key] = grouped[key].sort((left, right) => left.value.localeCompare(right.value));
            return accumulator;
        }, {});
}

function getAppSourceCounts(apps) {
    const counts = {
        graph: 0,
        entradocs: 0,
        learn: 0,
        custom: 0,
        official: 0,
        community: 0
    };

    for (const app of apps) {
        const sources = Array.isArray(app.sourceProvenance) && app.sourceProvenance.length > 0
            ? app.sourceProvenance
            : [app.source];
        sources.forEach((source) => {
            if (source === 'custom') {
                counts.custom += 1;
                return;
            }

            counts[source] = (counts[source] || 0) + 1;
        });
        if (app.isCommunity) {
            counts.community += 1;
        } else {
            counts.official += 1;
        }
    }

    return counts;
}

function getStats(permissions, apps, sourceFreshness) {
    const categories = groupPermissionsByCategory(permissions);
    const appCounts = getAppSourceCounts(apps);

    return {
        permissions: permissions.length,
        categories: Object.keys(categories).length,
        apps: apps.length,
        applicationPermissions: permissions.filter((item) => item.application).length,
        delegatedPermissions: permissions.filter((item) => item.delegated).length,
        officialApps: appCounts.official,
        communityApps: appCounts.community,
        sourceCounts: appCounts,
        freshestSourceUpdatedAt: Object.values(sourceFreshness)
            .map((item) => item.updatedAt)
            .sort()
            .slice(-1)[0]
    };
}

function loadRawInputs(rawDir) {
    ensureRequiredRawFiles(rawDir);

    const appRoles = normalizeArray(loadJson(path.join(rawDir, RAW_FILES.appRoles), null), RAW_FILES.appRoles);
    const delegateRoles = normalizeArray(loadJson(path.join(rawDir, RAW_FILES.delegateRoles), null), RAW_FILES.delegateRoles);
    const microsoftApps = normalizeArray(loadJson(path.join(rawDir, RAW_FILES.microsoftApps), null), RAW_FILES.microsoftApps);
    const permissionMethodsRaw = loadJson(path.join(rawDir, RAW_FILES.permissionMethods), null);
    const permissionPowerShellRaw = loadJson(path.join(rawDir, RAW_FILES.permissionPowerShell), null);
    const permissionCodeExamplesRaw = loadJsonWithShards(path.join(rawDir, RAW_FILES.permissionCodeExamples), null);
    const resourceSchemas = loadJson(path.join(rawDir, RAW_FILES.resourceSchemas), null);
    const resourceDocs = loadJson(path.join(rawDir, RAW_FILES.resourceDocs), null);

    if (!permissionMethodsRaw || typeof permissionMethodsRaw !== 'object' || Array.isArray(permissionMethodsRaw)) {
        throw new Error(`${RAW_FILES.permissionMethods} must be a JSON object.`);
    }

    if (!permissionPowerShellRaw || typeof permissionPowerShellRaw !== 'object' || Array.isArray(permissionPowerShellRaw)) {
        throw new Error(`${RAW_FILES.permissionPowerShell} must be a JSON object.`);
    }

    if (!permissionCodeExamplesRaw || typeof permissionCodeExamplesRaw !== 'object' || Array.isArray(permissionCodeExamplesRaw)) {
        throw new Error(`${RAW_FILES.permissionCodeExamples} must be a JSON object.`);
    }

    if (!resourceSchemas || typeof resourceSchemas !== 'object' || Array.isArray(resourceSchemas)) {
        throw new Error(`${RAW_FILES.resourceSchemas} must be a JSON object.`);
    }

    if (!resourceDocs || typeof resourceDocs !== 'object' || Array.isArray(resourceDocs)) {
        throw new Error(`${RAW_FILES.resourceDocs} must be a JSON object.`);
    }

    const permissionMethods = {
        byPermission: Object.entries(permissionMethodsRaw).reduce((accumulator, [permission, versions]) => {
            accumulator[permission.toLowerCase()] = versions;
            return accumulator;
        }, {})
    };
    const permissionPowerShell = {
        byPermission: Object.entries(permissionPowerShellRaw).reduce((accumulator, [permission, versions]) => {
            accumulator[permission.toLowerCase()] = versions;
            return accumulator;
        }, {})
    };
    const permissionCodeExamples = {
        byPermission: Object.entries(permissionCodeExamplesRaw).reduce((accumulator, [permission, versions]) => {
            accumulator[permission.toLowerCase()] = versions;
            return accumulator;
        }, {})
    };
    const permissionCsv = loadPermissionCsv(path.join(rawDir, RAW_FILES.permissionCsv));
    const sourceFreshness = buildSourceFreshness(rawDir);

    return {
        appRoles,
        delegateRoles,
        microsoftApps,
        permissionMethods,
        permissionPowerShell,
        permissionCodeExamples,
        permissionCsv,
        resourceSchemas,
        resourceDocs,
        sourceFreshness
    };
}

function normalizeRawData(rawDir, options = {}) {
    const loaded = loadRawInputs(rawDir);
    const ingestedAt = options.ingestedAt || new Date().toISOString();
    const normalizedPermissions = normalizePermissions(loaded.appRoles, loaded.delegateRoles)
        .map((permission) => buildPermissionDetail(permission, loaded.sourceFreshness, loaded.permissionMethods, loaded.permissionPowerShell, loaded.permissionCodeExamples, loaded.permissionCsv, loaded.resourceSchemas, loaded.resourceDocs));
    const normalizedApps = normalizeApps(loaded.microsoftApps, loaded.sourceFreshness.microsoftApps.updatedAt);
    const stats = getStats(normalizedPermissions, normalizedApps, loaded.sourceFreshness);

    const normalized = {
        schemaVersion: DEFAULT_SCHEMA_VERSION,
        ingestedAt,
        snapshotId: '',
        sourceFreshness: loaded.sourceFreshness,
        stats,
        categories: Object.entries(groupPermissionsByCategory(normalizedPermissions)).map(([name, items]) => ({
            name,
            count: items.length
        })),
        permissions: normalizedPermissions,
        apps: normalizedApps
    };

    normalized.snapshotId = hashContent(JSON.stringify({
        ingestedAt: normalized.ingestedAt,
        sourceFreshness: normalized.sourceFreshness,
        stats: normalized.stats
    })).slice(0, 16);

    return normalized;
}

function writeNormalizedData(normalized, outputDir) {
    writeJson(path.join(outputDir, 'site-data.json'), normalized);
}

function buildPermissionCatalog(normalized) {
    return {
        schemaVersion: normalized.schemaVersion,
        snapshotId: normalized.snapshotId,
        ingestedAt: normalized.ingestedAt,
        totalPermissions: normalized.stats.permissions,
        categories: normalized.categories,
        items: normalized.permissions.map((item) => [item.value, item.slug, item.category])
    };
}

function buildAppsManifest(normalized, chunkSize = DEFAULT_APP_CHUNK_SIZE) {
    const chunks = [];

    for (let index = 0; index < normalized.apps.length; index += chunkSize) {
        const count = Math.min(chunkSize, normalized.apps.length - index);
        chunks.push({
            file: `apps-${String(chunks.length + 1).padStart(3, '0')}.json`,
            count,
            start: index,
            end: index + count - 1
        });
    }

    return {
        schemaVersion: normalized.schemaVersion,
        snapshotId: normalized.snapshotId,
        ingestedAt: normalized.ingestedAt,
        totalApps: normalized.stats.apps,
        counts: normalized.stats.sourceCounts,
        sources: [
            { key: 'graph', label: SOURCE_LABELS.graph, count: normalized.stats.sourceCounts.graph || 0 },
            { key: 'entradocs', label: SOURCE_LABELS.entradocs, count: normalized.stats.sourceCounts.entradocs || 0 },
            { key: 'learn', label: SOURCE_LABELS.learn, count: normalized.stats.sourceCounts.learn || 0 },
            { key: 'community', label: SOURCE_LABELS.custom, count: normalized.stats.sourceCounts.community || 0 }
        ],
        detailBasePath: 'apps/',
        searchIndex: normalized.apps.map((app) => [app.title, app.appId, app.anchor]),
        chunks
    };
}

function buildPermissionPublicData(permission, normalized) {
    return {
        schemaVersion: normalized.schemaVersion,
        snapshotId: normalized.snapshotId,
        ingestedAt: normalized.ingestedAt,
        permission: {
            value: permission.value,
            slug: permission.slug,
            category: permission.category,
            description: permission.description,
            accessLevel: permission.accessLevel,
            scope: permission.scope,
            hasApplication: permission.hasApplication,
            hasDelegated: permission.hasDelegated
        },
        resource: permission.resource,
        sources: permission.sources,
        tabs: {
            apiV1Html: generateApiMethodsTable(permission, 'v1'),
            apiBetaHtml: generateApiMethodsTable(permission, 'beta'),
            psV1Html: generatePowerShellMethodsHtml(permission, 'v1'),
            psBetaHtml: generatePowerShellMethodsHtml(permission, 'beta'),
            propertiesHtml: generatePropertiesHtml(permission),
            jsonHtml: generateJsonRepresentationHtml(permission),
            relationshipsHtml: generateRelationshipsHtml(permission)
        }
    };
}

function writePublicData(normalized, outputDir, options = {}) {
    const appChunkSize = options.appChunkSize || DEFAULT_APP_CHUNK_SIZE;
    const permissionsCatalog = buildPermissionCatalog(normalized);
    const appsManifest = buildAppsManifest(normalized, appChunkSize);

    writeJson(path.join(outputDir, 'data', 'catalog', 'permissions.json'), permissionsCatalog);
    writeJson(path.join(outputDir, 'data', 'catalog', 'apps-manifest.json'), appsManifest);

    appsManifest.chunks.forEach((chunk) => {
        const items = normalized.apps.slice(chunk.start, chunk.end + 1);
        writeJson(path.join(outputDir, 'data', 'catalog', chunk.file), {
            schemaVersion: normalized.schemaVersion,
            snapshotId: normalized.snapshotId,
            ingestedAt: normalized.ingestedAt,
            items
        });
    });

    normalized.permissions.forEach((permission) => {
        writeJson(
            path.join(outputDir, 'data', 'permissions', `${permission.slug}.json`),
            buildPermissionPublicData(permission, normalized)
        );
    });
}

function validateNormalizedData(normalized, thresholds, options = {}) {
    const fixtureMode = Boolean(options.fixtureMode);
    const errors = [];
    const warnings = [];
    const now = Date.now();
    const maxStalenessMs = Number(thresholds.maxStalenessHours || 6) * 60 * 60 * 1000;
    const minPermissions = fixtureMode ? 1 : Number(thresholds.minPermissions || 1);
    const minOfficialApps = fixtureMode ? 1 : Number(thresholds.minOfficialApps || 1);
    const minCategories = fixtureMode ? 1 : Number(thresholds.minCategories || 1);

    if (normalized.permissions.length < minPermissions) {
        errors.push(`Expected at least ${minPermissions} permissions but found ${normalized.permissions.length}.`);
    }

    if (normalized.stats.officialApps < minOfficialApps) {
        errors.push(`Expected at least ${minOfficialApps} official apps but found ${normalized.stats.officialApps}.`);
    }

    if (normalized.categories.length < minCategories) {
        errors.push(`Expected at least ${minCategories} categories but found ${normalized.categories.length}.`);
    }

    for (const sourceKey of thresholds.requiredSources || []) {
        const freshness = normalized.sourceFreshness[sourceKey];
        if (!freshness?.updatedAt) {
            errors.push(`Missing freshness metadata for required source "${sourceKey}".`);
            continue;
        }

        const ageMs = now - new Date(freshness.updatedAt).getTime();
        if (!fixtureMode && ageMs > maxStalenessMs) {
            errors.push(`Source "${sourceKey}" is stale: ${freshness.updatedAt}.`);
        }
    }

    const appIds = new Set();
    const duplicateApps = new Set();
    normalized.apps.forEach((app) => {
        if (appIds.has(app.appId)) {
            duplicateApps.add(app.appId);
        }
        appIds.add(app.appId);
    });

    if (duplicateApps.size > 0) {
        errors.push(`Duplicate App IDs detected: ${Array.from(duplicateApps).slice(0, 10).join(', ')}.`);
    }

    const slugs = new Set();
    const duplicateSlugs = new Set();
    normalized.permissions.forEach((permission) => {
        if (slugs.has(permission.slug)) {
            duplicateSlugs.add(permission.slug);
        }
        slugs.add(permission.slug);
    });

    if (duplicateSlugs.size > 0) {
        errors.push(`Duplicate permission slugs detected: ${Array.from(duplicateSlugs).slice(0, 10).join(', ')}.`);
    }

    const unavailableMethods = normalized.permissions.filter((item) => item.methods.confidence === 'unavailable').length;
    const unavailablePowerShell = normalized.permissions.filter((item) => item.methods.powershell.confidence === 'unavailable').length;
    const unavailableCodeExamples = normalized.permissions.filter((item) =>
        !item.codeExamples.csharp.includes('View official example on Microsoft Learn')
        && !item.codeExamples.javascript.includes('View official example on Microsoft Learn')
        && !item.codeExamples.python.includes('View official example on Microsoft Learn')
        && !item.codeExamples.powershell.includes('View official example on Microsoft Learn')).length;
    const unavailableSchemas = normalized.permissions.filter((item) => item.properties.version === null).length;

    if (unavailableMethods > 0) {
        warnings.push(`${unavailableMethods} permissions do not have Microsoft Learn or OpenAPI endpoint metadata.`);
    }

    if (unavailablePowerShell > 0) {
        warnings.push(`${unavailablePowerShell} permissions do not have Microsoft Learn PowerShell command metadata.`);
    }

    if (unavailableCodeExamples > 0) {
        warnings.push(`${unavailableCodeExamples} permissions do not have official Microsoft Learn code examples for the displayed languages.`);
    }

    if (unavailableSchemas > 0) {
        warnings.push(`${unavailableSchemas} permissions do not have Learn or OpenAPI resource metadata.`);
    }

    return {
        valid: errors.length === 0,
        errors,
        warnings,
        metrics: {
            permissions: normalized.permissions.length,
            categories: normalized.categories.length,
            apps: normalized.apps.length,
            officialApps: normalized.stats.officialApps,
            communityApps: normalized.stats.communityApps,
            unavailableMethods,
            unavailablePowerShell,
            unavailableCodeExamples,
            unavailableSchemas
        }
    };
}

function generateSidebar(categories, currentSlug = null, basePath = '.') {
    return Object.entries(categories).sort((left, right) => left[0].localeCompare(right[0])).map(([category, permissions]) => {
        const expanded = currentSlug ? permissions.some((item) => item.slug === currentSlug) : false;

        return `
        <div class="nav-category ${expanded ? 'expanded' : ''}">
            <button class="nav-category-header" onclick="toggleCategory(this)">
                <svg class="nav-icon" viewBox="0 0 24 24"><path d="M9 18l6-6-6-6"/></svg>
                <span class="nav-category-name">${escapeHtml(category)}</span>
                <span class="nav-count">${permissions.length}</span>
            </button>
            <ul class="nav-items">
                ${permissions.map((permission) => `
                    <li><a href="${basePath}/permissions/${permission.slug}.html"
                           class="nav-item ${permission.slug === currentSlug ? 'active' : ''}"
                           data-permission="${permission.value.toLowerCase()}">${escapeHtml(permission.value)}</a></li>`).join('')}
            </ul>
        </div>`;
    }).join('');
}

function generateAppsSidebar(basePath = '.') {
    return `
        <div class="nav-category expanded">
            <button class="nav-category-header" onclick="toggleCategory(this)">
                <svg class="nav-icon" viewBox="0 0 24 24"><path d="M9 18l6-6-6-6"/></svg>
                <span class="nav-category-name">Apps Reference</span>
                <span class="nav-count">4</span>
            </button>
            <ul class="nav-items">
                <li><a href="${basePath}/microsoft-apps.html#apps-overview" class="nav-item">Overview</a></li>
                <li><a href="${basePath}/microsoft-apps.html#apps-catalog" class="nav-item">Catalog</a></li>
                <li><a href="${basePath}/microsoft-apps.html#apps-sources" class="nav-item">Sources</a></li>
                <li><a href="${basePath}/microsoft-apps.html#apps-data-contracts" class="nav-item">Data Contracts</a></li>
            </ul>
        </div>`;
}

function generateAppDetailSidebar(app, basePath = '..') {
    const sourceDoc = getAppSourceDoc(app);
    return `
        <div class="nav-category expanded">
            <button class="nav-category-header" onclick="toggleCategory(this)">
                <svg class="nav-icon" viewBox="0 0 24 24"><path d="M9 18l6-6-6-6"/></svg>
                <span class="nav-category-name">App Details</span>
                <span class="nav-count">4</span>
            </button>
            <ul class="nav-items">
                <li><a href="${basePath}/microsoft-apps.html" class="nav-item">Back to catalog</a></li>
                <li><a href="${basePath}/build-info.json" class="nav-item">Build info</a></li>
                <li><a href="${basePath}/data/catalog/apps-manifest.json" class="nav-item">Apps manifest</a></li>
                <li><a href="${sourceDoc.url}" class="nav-item" target="_blank" rel="noopener">${escapeHtml(sourceDoc.label)}</a></li>
            </ul>
        </div>`;
}

function generateAppsTableRows(apps, basePath = '.') {
    return apps.map((app) => {
        const portalUrl = getAppPortalUrl(app);
        const sourceClass = app.source === 'custom' ? 'custom' : app.source;
        const detailPath = `${basePath}/${getAppDetailPath(app)}`;
        const searchableSource = [
            app.source,
            app.sourceLabel,
            app.sourceDisplayLabel,
            ...(app.sourceProvenance || []),
            ...(app.sourceProvenanceLabels || [])
        ].filter(Boolean).join(' ').toLowerCase();
        return `
            <tr id="${app.anchor}" data-appid="${app.appId}" data-name="${escapeHtml(app.title.toLowerCase())}" data-source="${escapeHtml(searchableSource)}" data-filter-groups="${escapeHtml((app.filterGroups || [app.filterGroup]).join('|'))}">
                <td class="app-name">
                    <a href="${detailPath}" class="app-name-link">
                        <svg viewBox="0 0 24 24"><rect x="2" y="3" width="20" height="14" rx="2"/><path d="M8 21h8m-4-4v4"/></svg>
                        ${escapeHtml(app.title)}
                    </a>
                </td>
                <td><code class="id-code" onclick="copyToClipboard('${app.appId}')" title="Click to copy">${escapeHtml(app.appId)}</code></td>
                <td><span class="source-tag ${sourceClass}">${escapeHtml(app.sourceDisplayLabel || app.sourceLabel)}</span></td>
                <td class="action-col">
                    <div class="action-btns">
                        <a href="${detailPath}" title="Open detail page">
                            <svg viewBox="0 0 24 24"><path d="M5 12h14"/><path d="m12 5 7 7-7 7"/></svg>
                        </a>
                        <button onclick="copyToClipboard('${app.appId}')" title="Copy App ID">
                            <svg viewBox="0 0 24 24"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
                        </button>
                        <a href="${portalUrl}" target="_blank" rel="noopener" title="View in Azure Portal">
                            <svg viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                        </a>
                    </div>
                </td>
            </tr>
        `;
    }).join('');
}

function buildTypeBadges(permission) {
    let badges = '';
    if (permission.hasApplication) {
        badges += '<span class="badge badge-app"><svg viewBox="0 0 24 24"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>Application</span>';
    }
    if (permission.hasDelegated) {
        badges += '<span class="badge badge-delegated"><svg viewBox="0 0 24 24"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>Delegated</span>';
    }
    return badges;
}

function generatePermissionCards(permission) {
    let cards = '';

    if (permission.application) {
        cards += `
        <div class="detail-card detail-card-app">
            <div class="card-header">
                <span class="card-type app">
                    <svg viewBox="0 0 24 24"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>
                    Application Permission
                </span>
                ${permission.application.isEnabled === false ? '<span class="badge badge-disabled">Disabled</span>' : ''}
            </div>
            <div class="card-body">
                <h4>${escapeHtml(permission.application.displayName)}</h4>
                <p>${escapeHtml(permission.application.description)}</p>
            </div>
            <div class="card-footer">
                <span class="label">Permission ID:</span>
                <code class="id-code" onclick="copyToClipboard('${escapeHtml(permission.application.id)}')">${escapeHtml(permission.application.id)}</code>
            </div>
        </div>`;
    }

    if (permission.delegated) {
        const consentType = permission.delegated.type === 'User' ? 'User consent allowed' : 'Admin consent required';
        cards += `
        <div class="detail-card detail-card-delegated">
            <div class="card-header">
                <span class="card-type delegated">
                    <svg viewBox="0 0 24 24"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>
                    Delegated Permission
                </span>
                <span class="consent-type ${permission.delegated.type === 'User' ? 'user' : 'admin'}">${consentType}</span>
            </div>
            <div class="card-body">
                <h4>${escapeHtml(permission.delegated.displayName)}</h4>
                <p>${escapeHtml(permission.delegated.description)}</p>
                ${permission.delegated.userConsentDescription ? `
                <div class="user-consent-info">
                    <strong>User sees:</strong> ${escapeHtml(permission.delegated.userConsentDescription)}
                </div>` : ''}
            </div>
            <div class="card-footer">
                <span class="label">Permission ID:</span>
                <code class="id-code" onclick="copyToClipboard('${escapeHtml(permission.delegated.id)}')">${escapeHtml(permission.delegated.id)}</code>
            </div>
        </div>`;
    }

    return cards;
}

function generatePermissionIds(permission) {
    const rows = [];
    if (permission.application?.id) {
        rows.push(`
            <div class="id-row">
                <span>Application</span>
                <code class="id-code" onclick="copyToClipboard('${escapeHtml(permission.application.id)}')">${escapeHtml(permission.application.id)}</code>
            </div>`);
    }

    if (permission.delegated?.id) {
        rows.push(`
            <div class="id-row">
                <span>Delegated</span>
                <code class="id-code" onclick="copyToClipboard('${escapeHtml(permission.delegated.id)}')">${escapeHtml(permission.delegated.id)}</code>
            </div>`);
    }

    return rows.join('') || '<p class="no-data-message">No permission IDs available.</p>';
}

function buildLearnResourceLink(permission, version = 'v1.0') {
    if (permission.resource?.docLink) {
        return permission.resource.docLink;
    }

    const resource = resolveResourceName(permission.category.toLowerCase());
    const view = version === 'beta' ? 'graph-rest-beta' : 'graph-rest-1.0';
    return `https://learn.microsoft.com/en-us/graph/api/resources/${resource}?view=${view}`;
}

function buildMethodNoteHtml(label, description, className) {
    return `
        <div class="data-version-note ${className}">
            <span class="version-chip">${escapeHtml(label)}</span>
            <p>${escapeHtml(description)}</p>
        </div>`;
}

function generateResourceLinksHtml(permission) {
    const permissionReferenceUrl = `https://learn.microsoft.com/graph/permissions-reference#${permission.value.toLowerCase().replace(/\./g, '')}`;
    const resourceDocUrl = permission.resource?.docLink || buildLearnResourceLink(permission, permission.properties.version === 'beta' ? 'beta' : 'v1.0');
    const graphExplorerUrl = `https://developer.microsoft.com/graph/graph-explorer?request=${encodeURIComponent(permission.sampleEndpoint || '/me')}&method=GET&version=v1.0`;

    return `
        <ul class="resource-links">
            <li>
                <a href="${permissionReferenceUrl}" target="_blank" rel="noopener">
                    <svg viewBox="0 0 24 24"><path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/></svg>
                    Permission reference
                </a>
            </li>
            <li>
                <a href="${resourceDocUrl}" target="_blank" rel="noopener">
                    <svg viewBox="0 0 24 24"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"/></svg>
                    Resource type docs
                </a>
            </li>
            <li>
                <a href="${graphExplorerUrl}" target="_blank" rel="noopener">
                    <svg viewBox="0 0 24 24"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/></svg>
                    Graph Explorer sample
                </a>
            </li>
        </ul>`;
}

function generateApiMethodsTable(permission, version) {
    const versionKey = version === 'v1' ? 'v1.0' : 'beta';
    const endpoints = dedupeMethods(permission.methods.api[version] || []);
    const label = version === 'v1' ? 'Microsoft Graph v1.0' : 'Microsoft Graph beta';
    const noteClass = permission.methods.confidence === 'exact' ? 'stable' : 'beta';
    const noteDescription = permission.methods.confidence === 'exact'
        ? `${label} endpoints are mapped directly from refreshed Microsoft Learn permissions tables.`
        : permission.methods.confidence === 'category-derived'
            ? `${label} endpoints are shown from Microsoft Graph OpenAPI resource-family metadata because Microsoft Learn does not publish a direct mapping for this permission.`
            : `${label} endpoints are not available from refreshed Microsoft Learn or Microsoft Graph OpenAPI metadata for this permission.`;
    const note = buildMethodNoteHtml(permission.methods.label, noteDescription, noteClass);

    if (endpoints.length === 0) {
        return `${note}<p class="no-methods">No API methods available for this version.</p>`;
    }

    const sorted = [...endpoints].sort((left, right) => {
        const order = { GET: 1, POST: 2, PATCH: 3, PUT: 4, DELETE: 5 };
        const leftOrder = order[left.method] || 99;
        const rightOrder = order[right.method] || 99;
        return leftOrder - rightOrder || left.endpoint.localeCompare(right.endpoint);
    });

    let html = `${note}<table class="methods-table"><thead><tr><th>Methods</th></tr></thead><tbody>`;
    sorted.forEach((item) => {
        const graphExplorerUrl = `https://developer.microsoft.com/graph/graph-explorer?request=${encodeURIComponent(item.endpoint.split('?')[0])}&method=${item.method}&version=${versionKey}`;
        const learnLink = item.docLink
            ? `<a href="${item.docLink}" target="_blank" rel="noopener" class="try-btn" title="Open on Microsoft Learn">
                    <svg viewBox="0 0 24 24"><path d="M14 3h7v7"/><path d="M10 14 21 3"/><path d="M20 14v5a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h5"/></svg>
                </a>`
            : '';
        html += `<tr>
            <td>
                <span class="http-method ${item.method.toLowerCase()}">${escapeHtml(item.method)}</span>
                <code class="endpoint-path">${escapeHtml(item.endpoint)}</code>
                ${learnLink}
                <a href="${graphExplorerUrl}" target="_blank" rel="noopener" class="try-btn" title="Try in Graph Explorer">
                    <svg viewBox="0 0 24 24"><polygon points="5 3 19 12 5 21 5 3"/></svg>
                </a>
            </td>
        </tr>`;
    });
    html += '</tbody></table>';
    return html;
}

function generatePowerShellMethodsHtml(permission, version) {
    const label = version === 'v1' ? 'Microsoft Graph PowerShell v1.0' : 'Microsoft Graph PowerShell beta';
    const powerShellData = permission.methods?.powershell || {
        confidence: 'unavailable',
        label: 'No Microsoft Learn PowerShell mapping available',
        v1: [],
        beta: []
    };
    const commands = dedupePowerShellCommands(powerShellData[version] || []);
    const learnLink = version === 'v1'
        ? 'https://learn.microsoft.com/en-us/powershell/microsoftgraph/'
        : 'https://learn.microsoft.com/en-us/powershell/microsoftgraph/overview?view=graph-powershell-beta';
    const noteClass = powerShellData.confidence === 'exact' ? 'stable' : 'beta';
    const noteDescription = powerShellData.confidence === 'exact'
        ? `${label} commands are mapped directly from refreshed Microsoft Learn PowerShell snippets.`
        : `${label} commands are not available from refreshed Microsoft Learn PowerShell snippets for this permission.`;
    const note = buildMethodNoteHtml(powerShellData.label, noteDescription, noteClass);

    if (commands.length === 0) {
        return `
        ${note}
        <div class="no-data-card">
            <p class="no-data-message">No deterministic PowerShell command map is available for this permission.</p>
            <a href="${learnLink}" target="_blank" rel="noopener" class="learn-link">
                <svg viewBox="0 0 24 24" width="16" height="16"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                Browse PowerShell docs
            </a>
        </div>`;
    }

    const sorted = [...commands].sort((left, right) =>
        left.command.localeCompare(right.command)
        || String(left.endpoint || '').localeCompare(String(right.endpoint || ''))
    );

    let html = `${note}<table class="methods-table"><thead><tr><th>Commands</th></tr></thead><tbody>`;
    sorted.forEach((item) => {
        const learnButton = item.docLink
            ? `<a href="${item.docLink}" target="_blank" rel="noopener" class="try-btn" title="Open on Microsoft Learn">
                    <svg viewBox="0 0 24 24"><path d="M14 3h7v7"/><path d="M10 14 21 3"/><path d="M20 14v5a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h5"/></svg>
                </a>`
            : '';
        html += `<tr>
            <td>
                <code class="endpoint-path">${escapeHtml(item.command)}</code>
                ${item.endpoint ? `<span class="method-meta">${escapeHtml(item.endpoint)}</span>` : ''}
                ${item.title ? `<div class="method-meta">${escapeHtml(item.title)}</div>` : ''}
                ${learnButton}
            </td>
        </tr>`;
    });
    html += '</tbody></table>';
    return html;
}

function prependDataVersionNote(content, version, confidenceLabel, label) {
    if (!content) {
        return content;
    }

    const isBeta = version === 'beta';
    const versionLabel = isBeta ? 'Microsoft Graph beta' : 'Microsoft Graph v1.0';
    const description = isBeta
        ? `${label} is shown from beta metadata because a stable v1.0 schema is not available for this resource mapping.`
        : `${label} is shown from stable Microsoft Graph v1.0 metadata.`;

    return `
        <div class="data-version-note ${isBeta ? 'beta' : 'stable'}">
            <span class="version-chip">${versionLabel}</span>
            <span class="version-chip">${escapeHtml(confidenceLabel)}</span>
            <p>${escapeHtml(description)}</p>
        </div>
        ${content}`;
}

function generatePropertiesHtml(permission) {
    if (!permission.properties.version || permission.properties.items.length === 0) {
        return `<p class="no-properties">Properties metadata is not available for this permission mapping. <a href="${buildLearnResourceLink(permission, 'v1.0')}" target="_blank" rel="noopener">View on Microsoft Learn</a></p>`;
    }

    let html = `<div class="table-container">
        <table class="modern-table properties-table">
            <thead>
                <tr>
                    <th>Property</th>
                    <th>Type</th>
                    <th>Description</th>
                </tr>
            </thead>
            <tbody>`;

    permission.properties.items.slice(0, 15).forEach((property) => {
        const badges = [];
        if (property.readOnly) {
            badges.push('<span class="prop-badge readonly">Read-only</span>');
        }
        if (property.nullable) {
            badges.push('<span class="prop-badge nullable">Nullable</span>');
        }

        html += `
            <tr>
                <td class="property-name"><code>${escapeHtml(property.name)}</code></td>
                <td class="type-col"><code>${escapeHtml(property.type)}</code>${badges.join('')}</td>
                <td class="desc-col">${escapeHtml(property.description)}</td>
            </tr>`;
    });

    html += '</tbody></table></div>';

    if (permission.properties.items.length > 15) {
        html += `<p class="more-props">Showing 15 of ${permission.properties.items.length} properties.</p>`;
    }

    return prependDataVersionNote(html, permission.properties.version === 'v1' ? 'v1' : 'beta', permission.properties.confidence, 'Properties');
}

function generateJsonRepresentationHtml(permission) {
    if (!permission.jsonRepresentation.version || (!permission.jsonRepresentation.value && !permission.jsonRepresentation.raw)) {
        return `<p class="no-data-message">JSON representation is not available for this permission mapping. <a href="${buildLearnResourceLink(permission, 'v1.0')}#json-representation" target="_blank" rel="noopener">View on Microsoft Learn</a></p>`;
    }

    const jsonString = permission.jsonRepresentation.value
        ? JSON.stringify(permission.jsonRepresentation.value, null, 2)
        : permission.jsonRepresentation.raw;
    return prependDataVersionNote(`
        <div class="json-block">
            <div class="json-header">
                <span class="json-title">JSON representation</span>
                <button class="copy-btn" onclick="copyCode(this)">Copy</button>
            </div>
            <pre><code class="language-json">${escapeHtml(jsonString)}</code></pre>
        </div>`,
    permission.jsonRepresentation.version === 'v1' ? 'v1' : 'beta',
    permission.jsonRepresentation.confidence,
    'JSON representation');
}

function generateRelationshipsHtml(permission) {
    if (!permission.relationships.items.length) {
        return `<div class="no-data-card">
            <p class="no-data-message">Relationships metadata is not available for this permission mapping.</p>
            <a href="${buildLearnResourceLink(permission, 'v1.0')}" target="_blank" rel="noopener" class="learn-link">
                <svg viewBox="0 0 24 24" width="16" height="16"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                View resource documentation
            </a>
        </div>`;
    }

    let html = `<div class="table-container">
        <table class="modern-table properties-table">
            <thead>
                <tr>
                    <th>Relationship</th>
                    <th>Type</th>
                    <th>Description</th>
                </tr>
            </thead>
            <tbody>`;

    permission.relationships.items.forEach((item) => {
        html += `
            <tr>
                <td class="property-name"><code>${escapeHtml(item.name)}</code></td>
                <td class="type-col"><code>${escapeHtml(item.type)}</code></td>
                <td class="desc-col">${escapeHtml(item.description)}</td>
            </tr>`;
    });

    html += '</tbody></table></div>';

    return prependDataVersionNote(
        html,
        permission.properties.version === 'v1' ? 'v1' : 'beta',
        permission.relationships.confidence,
        'Relationships'
    );
}

function buildPermissionPageContent(permission) {
    const permissionTypeText = permission.hasApplication && permission.hasDelegated
        ? 'Application permissions or delegated permissions'
        : permission.hasApplication
            ? 'Application permissions'
            : 'Delegated permissions';
    const consentText = permission.hasApplication
        ? 'Application permissions always require admin consent.'
        : permission.delegated?.type === 'User'
            ? 'Users can consent to this permission during sign-in.'
            : 'This delegated permission requires admin consent.';

    return {
        typeBadges: buildTypeBadges(permission),
        accessBadge: `<span class="badge badge-${permission.accessLevel}">${ACCESS_LEVEL_LABELS[permission.accessLevel] || 'Read'}</span>`,
        scopeBadge: `<span class="badge badge-scope">${SCOPE_LABELS[permission.scope] || 'User Scope'}</span>`,
        permissionCards: generatePermissionCards(permission),
        permissionIds: generatePermissionIds(permission),
        permissionTypeText,
        consentText,
        accessLevelText: ACCESS_LEVEL_TEXT[permission.accessLevel] || ACCESS_LEVEL_TEXT.read,
        scopeText: SCOPE_TEXT[permission.scope] || SCOPE_TEXT.user,
        methodsApiV1: generateApiMethodsTable(permission, 'v1'),
        methodsApiBeta: generateApiMethodsTable(permission, 'beta'),
        methodsPsV1: generatePowerShellMethodsHtml(permission, 'v1'),
        methodsPsBeta: generatePowerShellMethodsHtml(permission, 'beta'),
        propertiesSection: `
            <section class="permission-section" id="properties">
                <h2>
                    <svg viewBox="0 0 24 24"><path d="M12 3v18m9-9H3"/></svg>
                    Properties
                </h2>
                ${generatePropertiesHtml(permission)}
            </section>`,
        jsonSection: `
            <section class="permission-section" id="json-representation">
                <h2>
                    <svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
                    JSON Representation
                </h2>
                ${generateJsonRepresentationHtml(permission)}
            </section>`,
        relationshipsSection: `
            <section class="permission-section" id="relationships">
                <h2>
                    <svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="2"/><path d="M16.24 7.76a6 6 0 0 1 0 8.49m-8.48-.01a6 6 0 0 1 0-8.49m11.31-2.82a10 10 0 0 1 0 14.14m-14.14 0a10 10 0 0 1 0-14.14"/></svg>
                    Relationships
                </h2>
                ${generateRelationshipsHtml(permission)}
            </section>`,
        codeCsharp: permission.codeExamples.csharp,
        codeJavascript: permission.codeExamples.javascript,
        codePowershell: permission.codeExamples.powershell,
        codePython: permission.codeExamples.python,
        delegatedClass: permission.hasDelegated ? 'supported' : 'not-supported',
        applicationClass: permission.hasApplication ? 'supported' : 'not-supported',
        permissionAnchor: permission.value.toLowerCase().replace(/\./g, ''),
        resourceLinks: generateResourceLinksHtml(permission),
        sourceFreshnessText: `Permission data: ${formatUtcLabel(permission.sources.permission.sourceUpdatedAt)}`
    };
}

module.exports = {
    DEFAULT_APP_CHUNK_SIZE,
    DEFAULT_SCHEMA_VERSION,
    SITE_NAME,
    SITE_URL,
    buildAppsManifest,
    buildPermissionCatalog,
    buildPermissionPageContent,
    formatUtcLabel,
    generateAppDetailSidebar,
    generateAppsSidebar,
    generateAppsTableRows,
    generateSidebar,
    getAppAnchor,
    getAppDetailPath,
    getAppPortalUrl,
    getAppSourceDescription,
    getAppSourceDoc,
    getAppSourceCounts,
    groupPermissionsByCategory,
    loadRawInputs,
    normalizeRawData,
    validateNormalizedData,
    writeNormalizedData,
    writePublicData
};
