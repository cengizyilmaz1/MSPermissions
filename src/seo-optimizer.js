class SEOOptimizer {
    constructor(options = {}) {
        this.siteName = options.siteName || 'Graph Permissions Explorer';
        this.siteUrl = options.siteUrl || 'https://permissions.cengizyilmaz.net';
        this.author = options.author || 'Cengiz Yilmaz';
        this.twitterHandle = options.twitterHandle || '@cengizyilmaz_';
    }

    generatePermissionTitle(permission) {
        const primaryTitle = `${permission.value} | Graph Permissions`;
        if (primaryTitle.length <= 60) {
            return primaryTitle;
        }

        const compactTitle = `${permission.value} | MS Graph`;
        if (compactTitle.length <= 60) {
            return compactTitle;
        }

        return permission.value;
    }

    generatePermissionDescription(permission) {
        const value = permission.value;
        const category = permission.category;
        const hasApp = Boolean(permission.application);
        const hasDelegated = Boolean(permission.delegated);

        let permissionType = 'Delegated';
        if (hasApp && hasDelegated) {
            permissionType = 'Application and Delegated';
        } else if (hasApp) {
            permissionType = 'Application';
        }

        const baseDescription = permission.application?.description
            || permission.delegated?.description
            || '';

        let description = `${value} is a ${permissionType} permission in Microsoft Graph API. `;

        if (baseDescription) {
            const cleanDescription = baseDescription.replace(/Allows the app to /gi, '').trim();
            description += cleanDescription.slice(0, 100);
        } else {
            description += `Use it to access ${category} resources in Azure AD and Microsoft 365.`;
        }

        return description.slice(0, 155) + (description.length > 155 ? '...' : '');
    }

    generatePermissionKeywords(permission) {
        const value = permission.value;
        const category = permission.category;
        const parts = value.split('.');

        const keywords = [
            value,
            `${value} permission`,
            `Microsoft Graph ${value}`,
            `${category} permission`,
            `Graph API ${category}`,
            'Microsoft Graph API',
            'Azure AD permissions',
            'Entra ID permissions',
            'OAuth scopes',
            'Microsoft 365 API',
            ...parts.map((part) => `${part} permission`),
            permission.application ? 'Application permission' : null,
            permission.delegated ? 'Delegated permission' : null,
            'Graph API scopes',
            'Azure permissions'
        ].filter(Boolean);

        return [...new Set(keywords)].join(', ');
    }

    generatePermissionStructuredData(permission, options = {}) {
        const description = this.generatePermissionDescription(permission);
        const { dateModified = null } = options;
        const additionalProperty = [
            {
                '@type': 'PropertyValue',
                name: 'Category',
                value: permission.category
            },
            {
                '@type': 'PropertyValue',
                name: 'Access Level',
                value: permission.accessLevel
            },
            {
                '@type': 'PropertyValue',
                name: 'Scope',
                value: permission.scope
            },
            {
                '@type': 'PropertyValue',
                name: 'Supports Application',
                value: String(Boolean(permission.application))
            },
            {
                '@type': 'PropertyValue',
                name: 'Supports Delegated',
                value: String(Boolean(permission.delegated))
            },
            {
                '@type': 'PropertyValue',
                name: 'REST v1 Methods',
                value: String(permission.methods?.api?.v1?.length || 0)
            },
            {
                '@type': 'PropertyValue',
                name: 'REST Beta Methods',
                value: String(permission.methods?.api?.beta?.length || 0)
            },
            {
                '@type': 'PropertyValue',
                name: 'PowerShell v1 Commands',
                value: String(permission.methods?.powershell?.v1?.length || 0)
            },
            {
                '@type': 'PropertyValue',
                name: 'PowerShell Beta Commands',
                value: String(permission.methods?.powershell?.beta?.length || 0)
            }
        ];

        if (permission.application?.id) {
            additionalProperty.push({
                '@type': 'PropertyValue',
                name: 'Application Permission ID',
                value: permission.application.id
            });
        }

        if (permission.delegated?.id) {
            additionalProperty.push({
                '@type': 'PropertyValue',
                name: 'Delegated Permission ID',
                value: permission.delegated.id
            });
        }

        const structuredData = {
            '@context': 'https://schema.org',
            '@type': 'TechArticle',
            headline: `${permission.value} - Microsoft Graph Permission`,
            description,
            author: {
                '@type': 'Person',
                name: this.author,
                url: 'https://cengizyilmaz.net'
            },
            publisher: {
                '@type': 'Organization',
                name: this.siteName,
                url: this.siteUrl,
                logo: {
                    '@type': 'ImageObject',
                    url: `${this.siteUrl}/favicon.svg`
                }
            },
            mainEntityOfPage: {
                '@type': 'WebPage',
                '@id': `${this.siteUrl}/permissions/${permission.slug}.html`
            },
            keywords: this.generatePermissionKeywords(permission),
            articleSection: permission.category,
            about: {
                '@type': 'SoftwareApplication',
                name: 'Microsoft Graph API',
                applicationCategory: 'DeveloperApplication',
                operatingSystem: 'Cross-platform'
            },
            isAccessibleForFree: true,
            inLanguage: 'en-US',
            additionalProperty
        };

        if (dateModified) {
            structuredData.dateModified = dateModified;
        }

        return structuredData;
    }

    generateAppsPageTitle() {
        return 'Microsoft App IDs Catalog | Graph Permissions';
    }

    generateAppsPageDescription(appCount) {
        return `Browse ${appCount} Microsoft first-party application IDs with a lightweight searchable catalog and dedicated detail pages for each app. Sources include Microsoft Graph, Entra Docs, Microsoft Learn, and clearly labeled community data.`;
    }

    generateAppsPageKeywords() {
        return [
            'Microsoft app IDs',
            'first-party application IDs',
            'Microsoft Graph app IDs',
            'Entra application IDs',
            'Microsoft 365 client IDs',
            'service principal app IDs'
        ].join(', ');
    }

    generateAppDetailTitle(app) {
        return `${app.title} App ID | Graph Permissions`;
    }

    generateAppDetailDescription(app) {
        const trustText = app.isCommunity
            ? 'community-maintained'
            : 'official Microsoft';

        return `${app.title} is a ${trustText} application identifier in the Graph Permissions Explorer catalog. Use App ID ${app.appId} for sign-in log analysis, service principal investigations, and reference sharing.`;
    }

    generateAppDetailKeywords(app) {
        return [
            `${app.title} app id`,
            app.appId,
            `${app.title} client id`,
            `${app.sourceLabel} app id`,
            'Microsoft first-party app id',
            'service principal app id'
        ].join(', ');
    }

    generateAppsOverviewStructuredData(stats, options = {}) {
        const { dateModified = null } = options;

        const entries = [
            {
                '@context': 'https://schema.org',
                '@type': 'CollectionPage',
                name: this.generateAppsPageTitle(stats.apps),
                description: this.generateAppsPageDescription(stats.apps),
                url: `${this.siteUrl}/microsoft-apps.html`,
                inLanguage: 'en-US',
                isPartOf: {
                    '@type': 'WebSite',
                    name: this.siteName,
                    url: this.siteUrl
                },
                mainEntity: {
                    '@type': 'Dataset',
                    name: 'Microsoft first-party application ID catalog',
                    description: `Catalog of ${stats.apps} Microsoft first-party application IDs.`
                }
            },
            {
                '@context': 'https://schema.org',
                '@type': 'Dataset',
                name: 'Microsoft first-party application ID catalog',
                description: `Searchable catalog of ${stats.apps} Microsoft first-party application IDs with official and community source labels.`,
                url: `${this.siteUrl}/data/catalog/apps-manifest.json`,
                license: 'https://opensource.org/licenses/MIT',
                creator: {
                    '@type': 'Person',
                    name: this.author
                },
                distribution: [
                    {
                        '@type': 'DataDownload',
                        encodingFormat: 'application/json',
                        contentUrl: `${this.siteUrl}/data/catalog/apps-manifest.json`
                    }
                ]
            }
        ];

        if (dateModified) {
            entries.forEach((entry) => {
                entry.dateModified = dateModified;
            });
        }

        return entries;
    }

    generateHomepageTitle(stats) {
        return 'Microsoft Graph Permissions Reference | Graph Permissions';
    }

    generateHomepageDescription(stats) {
        return `Explore ${stats.permissions} Microsoft Graph permissions and ${stats.apps} Microsoft app IDs. Find application and delegated scopes, code examples, and API access guidance.`;
    }

    generateWebsiteStructuredData(stats, options = {}) {
        const { dateModified = null } = options;
        const data = {
            '@context': 'https://schema.org',
            '@type': 'WebSite',
            name: this.siteName,
            alternateName: [
                'Microsoft Graph API Permissions Reference',
                'Graph API Permissions',
                'MS Graph Permissions'
            ],
            description: this.generateHomepageDescription(stats),
            url: this.siteUrl,
            inLanguage: 'en-US',
            isAccessibleForFree: true,
            publisher: {
                '@type': 'Person',
                name: this.author,
                url: 'https://cengizyilmaz.net',
                sameAs: [
                    'https://github.com/cengizyilmaz1',
                    'https://x.com/cengizyilmaz_',
                    'https://linkedin.com/in/cengizyilmazz'
                ]
            }
        };

        if (dateModified) {
            data.dateModified = dateModified;
        }

        return data;
    }

    generateHomepageStructuredData(stats, options = {}) {
        const { dateModified = null } = options;

        return [
            {
                '@context': 'https://schema.org',
                '@type': 'CollectionPage',
                name: this.generateHomepageTitle(stats),
                description: this.generateHomepageDescription(stats),
                url: `${this.siteUrl}/`,
                inLanguage: 'en-US',
                isPartOf: {
                    '@type': 'WebSite',
                    name: this.siteName,
                    url: this.siteUrl
                },
                mainEntity: {
                    '@type': 'Dataset',
                    name: 'Graph Permissions Explorer catalog',
                    description: 'Structured catalog of Microsoft Graph permissions, methods, code examples, and Microsoft first-party application IDs.',
                    url: `${this.siteUrl}/build-info.json`
                }
            },
            {
                '@context': 'https://schema.org',
                '@type': 'Dataset',
                name: 'Graph Permissions Explorer public catalog',
                description: 'Public JSON catalogs for Microsoft Graph permissions and Microsoft first-party application IDs.',
                url: `${this.siteUrl}/build-info.json`,
                license: 'https://opensource.org/licenses/MIT',
                creator: {
                    '@type': 'Person',
                    name: this.author
                },
                distribution: [
                    {
                        '@type': 'DataDownload',
                        encodingFormat: 'application/json',
                        contentUrl: `${this.siteUrl}/data/catalog/permissions.json`
                    },
                    {
                        '@type': 'DataDownload',
                        encodingFormat: 'application/json',
                        contentUrl: `${this.siteUrl}/data/catalog/apps-manifest.json`
                    }
                ]
            }
        ].map((entry) => {
            if (dateModified) {
                entry.dateModified = dateModified;
            }
            return entry;
        });
    }

    generateAppDetailStructuredData(app, options = {}) {
        const { dateModified = null, sourceDocUrl = null } = options;
        const pageUrl = `${this.siteUrl}/apps/${app.anchor}.html`;
        const entries = [
            {
                '@context': 'https://schema.org',
                '@type': 'ProfilePage',
                name: this.generateAppDetailTitle(app),
                description: this.generateAppDetailDescription(app),
                url: pageUrl,
                mainEntity: {
                    '@type': 'SoftwareApplication',
                    name: app.title,
                    identifier: app.appId,
                    applicationCategory: 'BusinessApplication',
                    operatingSystem: 'Cross-platform',
                    provider: {
                        '@type': 'Organization',
                        name: app.isCommunity ? 'Community maintained' : 'Microsoft'
                    },
                    additionalProperty: [
                        {
                            '@type': 'PropertyValue',
                            name: 'Source',
                            value: app.sourceLabel
                        },
                        {
                            '@type': 'PropertyValue',
                            name: 'Official',
                            value: String(Boolean(app.isOfficial))
                        }
                    ]
                }
            }
        ];

        if (sourceDocUrl) {
            entries[0].mainEntity.sameAs = [sourceDocUrl];
        }

        if (dateModified) {
            entries.forEach((entry) => {
                entry.dateModified = dateModified;
            });
        }

        return entries;
    }

    getCanonicalUrl(path) {
        if (!path || path === '/') {
            return `${this.siteUrl}/`;
        }

        return `${this.siteUrl}/${path.replace(/^\//, '')}`;
    }
}

module.exports = SEOOptimizer;
