/**
 * llms.txt / llms-full.txt generator for Graph Permissions Explorer.
 *
 * `llms.txt` follows the llmstxt.org layout: a single H1, a blockquote summary,
 * short prose, then H2 sections that contain nothing but Markdown link lists.
 * It is a navigational index, so it links to categories and contracts rather
 * than enumerating every permission.
 *
 * `llms-full.txt` inlines the whole reference so a single fetch gives a model
 * the complete permission and application catalog.
 */

const fs = require('fs');
const path = require('path');

const { getAppDetailPath } = require('./lib/site-data');

/** Endpoint lists are unbounded upstream; the detail page stays authoritative. */
const MAX_METHODS_PER_BUCKET = 20;
const MAX_COMMANDS_PER_BUCKET = 12;
const MAX_PROPERTIES = 25;
/** llms.txt is a navigational index, so it lists the largest categories only. */
const MAX_INDEXED_CATEGORIES = 30;

function escapeCell(value) {
    return String(value ?? '')
        .replace(/\|/g, '\\|')
        .replace(/\s+/g, ' ')
        .trim();
}

function collapse(value) {
    return String(value ?? '')
        .replace(/\s+/g, ' ')
        .trim();
}

function permissionTypeLabel(permission) {
    if (permission.hasApplication && permission.hasDelegated) {
        return 'Application and Delegated';
    }

    return permission.hasApplication ? 'Application' : 'Delegated';
}

function consentLabel(permission) {
    if (permission.hasApplication) {
        return 'Admin consent required (application permissions always require admin consent)';
    }

    if (permission.delegated?.type === 'User') {
        return 'User consent allowed at sign-in';
    }

    return 'Admin consent required';
}

function buildLlmsTxt(normalized, options = {}) {
    const { siteUrl } = options;
    const permissionsByCategory = new Map();

    normalized.permissions.forEach((permission) => {
        permissionsByCategory.set(
            permission.category,
            (permissionsByCategory.get(permission.category) || 0) + 1
        );
    });

    const rankedCategories = [...permissionsByCategory.entries()].sort(
        (left, right) => right[1] - left[1] || left[0].localeCompare(right[0])
    );
    const indexedCategories = rankedCategories
        .slice(0, MAX_INDEXED_CATEGORIES)
        .sort((left, right) => left[0].localeCompare(right[0]));
    const remainingCategories = rankedCategories.length - indexedCategories.length;

    const lines = [
        '# Graph Permissions Explorer',
        '',
        `> An open reference for all ${normalized.stats.permissions} Microsoft Graph API permissions and ${normalized.stats.apps} Microsoft first-party application IDs, rebuilt from official Microsoft sources.`,
        '',
        'Every permission page publishes the real Microsoft Graph identifiers: the app role GUID for the application permission and the OAuth2 permission scope GUID for the delegated permission. Those GUIDs can be pasted straight into an app registration manifest. Pages also carry the consent requirement, the Graph REST methods and Microsoft Graph PowerShell commands that use the permission, and official SDK code examples.',
        '',
        `Data is refreshed from Microsoft Graph service principals, Microsoft Learn, the Microsoft Entra known-GUID catalog, and the Microsoft Graph OpenAPI metadata. Community-contributed application entries are labeled as community and are never presented as official Microsoft sources. The current snapshot is ${normalized.snapshotId}, ingested ${normalized.ingestedAt}.`,
        '',
        `Permissions are grouped by the resource prefix of the permission name, giving ${rankedCategories.length} categories in this snapshot. The largest categories are indexed below; llms-full.txt expands every category and every permission inline.`,
        '',
        'When citing this site, link to the canonical permission or application detail URL rather than to the homepage.',
        '',
        '## Core pages',
        '',
        `- [Graph Permissions Explorer home](${siteUrl}/): Overview of the permission catalog, grouped by category, with search across permissions and apps.`,
        `- [Microsoft first-party application IDs](${siteUrl}/microsoft-apps.html): Searchable catalog of ${normalized.stats.apps} Microsoft application (client) IDs with the source each one came from.`,
        `- [Full reference for LLMs](${siteUrl}/llms-full.txt): The entire permission and application reference inlined as Markdown in one file.`,
        '',
        '## Machine-readable data',
        '',
        `- [Permission catalog](${siteUrl}/data/catalog/permissions.json): Every permission as a compact tuple \`[value, slug, category, applicationId, delegatedId, requiresAdmin]\`. An empty ID string means that permission type does not exist; \`requiresAdmin\` is 1 when admin consent is required.`,
        `- [Permission detail records](${siteUrl}/data/permissions/{slug}.json): Full per-permission record including Graph methods, PowerShell commands, official SDK code examples, and resource schema.`,
        `- [Applications manifest](${siteUrl}/data/catalog/apps-manifest.json): Chunk list plus a search index of \`[title, appId, anchor]\`; the full records live in the \`apps-*.json\` chunks it points to.`,
        `- [Build info](${siteUrl}/build-info.json): Snapshot ID, ingestion time, per-source freshness, and catalog counts. Use this to check how current the data is.`,
        '',
        '## Permission categories',
        ''
    ];

    indexedCategories.forEach(([category, count]) => {
        lines.push(
            `- [${category} permissions](${siteUrl}/#${encodeURIComponent(category)}): ${count} Microsoft Graph ${count === 1 ? 'permission that governs' : 'permissions that govern'} ${category} data and operations.`
        );
    });

    if (remainingCategories > 0) {
        lines.push(
            `- [All remaining categories](${siteUrl}/llms-full.txt): The other ${remainingCategories} permission categories, with every permission expanded inline.`
        );
    }

    lines.push(
        '',
        '## URL patterns',
        '',
        `- [Permission detail page](${siteUrl}/permissions/{slug}.html): Canonical page for one permission. The slug is the permission name lowercased with dots replaced by hyphens, for example \`User.Read\` becomes \`user-read\`.`,
        `- [Application detail page](${siteUrl}/apps/{anchor}.html): Canonical page for one Microsoft application ID. The anchor is the slugified app name followed by its App ID.`,
        '',
        '## Optional',
        '',
        `- [Sitemap index](${siteUrl}/sitemap.xml): Every published URL, split into per-section sitemaps.`,
        `- [Source repository](https://github.com/cengizyilmaz1/MSPermissions): Build pipeline, data normalization, and issue tracker. MIT licensed.`,
        `- [Microsoft Graph permissions reference](https://learn.microsoft.com/graph/permissions-reference): The upstream Microsoft documentation this site is derived from and should be verified against.`,
        `- [Author](https://cengizyilmaz.net/): Cengiz Yilmaz, who maintains this reference.`,
        `- [IndieTools listing](https://www.indietools.app/products/microsoft-graph-permissions): Directory listing for this project.`,
        ''
    );

    return lines.join('\n');
}

function appendPermissionMethods(lines, permission) {
    const apiV1 = permission.methods?.api?.v1 || [];
    const apiBeta = permission.methods?.api?.beta || [];
    const psV1 = permission.methods?.powershell?.v1 || [];
    const psBeta = permission.methods?.powershell?.beta || [];

    const renderMethods = (label, entries) => {
        if (entries.length === 0) {
            return;
        }

        lines.push(`- ${label}:`);
        entries.slice(0, MAX_METHODS_PER_BUCKET).forEach((entry) => {
            lines.push(`  - \`${entry.method} ${entry.endpoint}\``);
        });

        if (entries.length > MAX_METHODS_PER_BUCKET) {
            lines.push(
                `  - ...and ${entries.length - MAX_METHODS_PER_BUCKET} more; see the permission detail page for the full list.`
            );
        }
    };

    const renderCommands = (label, entries) => {
        if (entries.length === 0) {
            return;
        }

        const commands = [...new Set(entries.map((entry) => entry.command).filter(Boolean))];
        if (commands.length === 0) {
            return;
        }

        const shown = commands.slice(0, MAX_COMMANDS_PER_BUCKET);
        const suffix =
            commands.length > MAX_COMMANDS_PER_BUCKET
                ? `, ...and ${commands.length - MAX_COMMANDS_PER_BUCKET} more`
                : '';
        lines.push(`- ${label}: ${shown.map((command) => `\`${command}\``).join(', ')}${suffix}`);
    };

    renderMethods('Graph REST methods (v1.0)', apiV1);
    renderMethods('Graph REST methods (beta)', apiBeta);
    renderCommands('Microsoft Graph PowerShell (v1.0)', psV1);
    renderCommands('Microsoft Graph PowerShell (beta)', psBeta);
}

function appendPermission(lines, permission, siteUrl) {
    const url = `${siteUrl}/permissions/${permission.slug}.html`;

    lines.push('', `### ${permission.value}`, '');
    lines.push(
        `${permission.value} is a Microsoft Graph ${permissionTypeLabel(permission).toLowerCase()} permission in the ${permission.category} category. ${collapse(permission.description)}`.trim()
    );
    lines.push('');
    lines.push(`- Canonical URL: ${url}`);
    lines.push(`- Permission types: ${permissionTypeLabel(permission)}`);
    lines.push(`- Category: ${permission.category}`);
    lines.push(`- Access level: ${permission.accessLevel}`);
    lines.push(`- Scope: ${permission.scope}`);
    lines.push(`- Consent: ${consentLabel(permission)}`);

    if (permission.application?.id) {
        lines.push(
            `- Application permission (app role) ID: \`${permission.application.id}\` - ${collapse(permission.application.displayName)}`
        );
        if (permission.application.description) {
            lines.push(
                `  - Application permission description: ${collapse(permission.application.description)}`
            );
        }
    } else {
        lines.push('- Application permission: not available for this permission.');
    }

    if (permission.delegated?.id) {
        lines.push(
            `- Delegated permission (OAuth2 scope) ID: \`${permission.delegated.id}\` - ${collapse(permission.delegated.displayName)}`
        );
        if (permission.delegated.description) {
            lines.push(
                `  - Delegated permission description: ${collapse(permission.delegated.description)}`
            );
        }
        if (permission.delegated.type) {
            lines.push(`  - Delegated consent type: ${permission.delegated.type}`);
        }
    } else {
        lines.push('- Delegated permission: not available for this permission.');
    }

    if (permission.resource?.name) {
        const docLink = permission.resource.docLink ? ` (${permission.resource.docLink})` : '';
        lines.push(`- Primary Graph resource: \`${permission.resource.name}\`${docLink}`);
    }

    appendPermissionMethods(lines, permission);

    const properties = permission.properties?.items || [];
    if (properties.length > 0) {
        const shown = properties.slice(0, MAX_PROPERTIES).map((item) => item.name);
        const suffix =
            properties.length > MAX_PROPERTIES
                ? `, ...and ${properties.length - MAX_PROPERTIES} more`
                : '';
        lines.push(`- Resource properties: ${shown.join(', ')}${suffix}`);
    }

    lines.push(`- Machine-readable record: ${siteUrl}/data/permissions/${permission.slug}.json`);
}

function buildLlmsFullTxt(normalized, options = {}) {
    const { siteUrl } = options;
    const lines = [
        '# Graph Permissions Explorer - Full Reference',
        '',
        `> The complete Microsoft Graph permission and Microsoft first-party application reference from ${siteUrl}, inlined as Markdown for language models.`,
        '',
        `Snapshot ${normalized.snapshotId}, ingested ${normalized.ingestedAt}. This file contains ${normalized.stats.permissions} permissions across ${normalized.stats.categories} categories and ${normalized.stats.apps} Microsoft application IDs.`,
        '',
        'All permission IDs below are the real Microsoft Graph app role and OAuth2 permission scope GUIDs and can be used directly in an app registration manifest. Application permissions always require admin consent. Delegated permissions with a User consent type can be consented to by the signed-in user.',
        '',
        'Long endpoint and command lists are truncated here; the linked canonical page and JSON record hold the complete data. Official SDK code examples are intentionally excluded from this file for size and are available on each permission page.',
        '',
        '## How to cite this reference',
        '',
        '- Cite the canonical permission or application detail URL, not this file.',
        '- Attribute the reference to Graph Permissions Explorer by Cengiz Yilmaz.',
        '- Application entries labeled "Community" are community-maintained and must not be presented as official Microsoft sources.',
        '- Verify security-sensitive answers against the Microsoft Graph permissions reference at https://learn.microsoft.com/graph/permissions-reference.',
        '',
        '## Snapshot and source freshness',
        ''
    ];

    Object.entries(normalized.sourceFreshness).forEach(([key, value]) => {
        lines.push(`- ${key}: updated ${value.updatedAt} from ${value.source}`);
    });

    lines.push(
        '',
        '## Key definitions',
        '',
        '- Microsoft Graph permission: a named authorization string, such as `User.Read`, that an application requests to access Microsoft 365 or Microsoft Entra ID data through the Microsoft Graph API.',
        '- Application permission (app role): grants app-only access with no signed-in user. It applies tenant-wide and always requires administrator consent.',
        '- Delegated permission (OAuth2 permission scope): grants access on behalf of a signed-in user. Effective access is the intersection of the permission and the user\u2019s own rights.',
        '- Permission ID: the GUID that identifies the permission inside an app registration manifest. Application and delegated variants of the same permission name have different GUIDs.',
        '- Microsoft first-party application ID: the client ID of an application published by Microsoft, used to recognize Microsoft services in Entra ID sign-in and audit logs.',
        '',
        '## Permissions',
        '',
        `${normalized.stats.permissions} Microsoft Graph permissions, grouped by category.`
    );

    const byCategory = new Map();
    normalized.permissions.forEach((permission) => {
        if (!byCategory.has(permission.category)) {
            byCategory.set(permission.category, []);
        }
        byCategory.get(permission.category).push(permission);
    });

    [...byCategory.keys()]
        .sort((left, right) => String(left).localeCompare(String(right)))
        .forEach((category) => {
            const permissions = byCategory.get(category);
            lines.push('', `## ${category} permissions`, '');
            lines.push(
                `${permissions.length} Microsoft Graph ${permissions.length === 1 ? 'permission governs' : 'permissions govern'} ${category} data. Category index: ${siteUrl}/#${encodeURIComponent(category)}`
            );

            permissions.forEach((permission) => appendPermission(lines, permission, siteUrl));
        });

    lines.push(
        '',
        '## Microsoft first-party application IDs',
        '',
        `${normalized.stats.apps} application (client) IDs published by Microsoft or contributed by the community. Overview page: ${siteUrl}/microsoft-apps.html`,
        '',
        '| Application | App ID | Source | Detail page |',
        '| --- | --- | --- | --- |'
    );

    normalized.apps.forEach((app) => {
        const trust = app.isCommunity ? 'Community' : app.sourceDisplayLabel || app.sourceLabel;
        lines.push(
            `| ${escapeCell(app.title)} | \`${escapeCell(app.appId)}\` | ${escapeCell(trust)} | ${siteUrl}/${getAppDetailPath(app)} |`
        );
    });

    lines.push('');

    return lines.join('\n');
}

function writeLlmsFiles(normalized, outputDir, options = {}) {
    const llmsTxt = `${buildLlmsTxt(normalized, options)}\n`;
    const llmsFullTxt = `${buildLlmsFullTxt(normalized, options)}\n`;

    fs.writeFileSync(path.join(outputDir, 'llms.txt'), llmsTxt);
    fs.writeFileSync(path.join(outputDir, 'llms-full.txt'), llmsFullTxt);

    return {
        llmsBytes: Buffer.byteLength(llmsTxt),
        llmsFullBytes: Buffer.byteLength(llmsFullTxt)
    };
}

module.exports = {
    buildLlmsFullTxt,
    buildLlmsTxt,
    writeLlmsFiles
};
