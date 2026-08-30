class SEOOptimizer {
    constructor(options = {}) {
        this.siteName = options.siteName || 'Graph Permissions Explorer';
        this.siteUrl = options.siteUrl || 'https://permissions.cengizyilmaz.net';
        this.author = options.author || 'Cengiz Yilmaz';
        this.authorUrl = options.authorUrl || 'https://cengizyilmaz.net';
        this.twitterHandle = options.twitterHandle || '@cengizyilmaz_';
    }

    /**
     * Stable author entity reused across every page so search engines and LLMs
     * can resolve one identity instead of many look-alike Person nodes.
     */
    getAuthorEntity() {
        return {
            '@type': 'Person',
            '@id': `${this.siteUrl}/#author`,
            name: this.author,
            url: this.authorUrl,
            sameAs: [
                'https://cengizyilmaz.net/',
                'https://github.com/cengizyilmaz1',
                'https://x.com/cengizyilmaz_',
                'https://linkedin.com/in/cengizyilmazz'
            ]
        };
    }

    getPublisherEntity() {
        return {
            '@type': 'Organization',
            '@id': `${this.siteUrl}/#publisher`,
            name: this.siteName,
            url: `${this.siteUrl}/`,
            logo: {
                '@type': 'ImageObject',
                url: `${this.siteUrl}/favicon.svg`
            }
        };
    }

    /**
     * @param {Array<{ name: string, url: string }>} items Ordered trail, root first.
     */
    generateBreadcrumbStructuredData(items) {
        if (!Array.isArray(items) || items.length < 2) {
            return null;
        }

        return {
            '@context': 'https://schema.org',
            '@type': 'BreadcrumbList',
            itemListElement: items.map((item, index) => ({
                '@type': 'ListItem',
                position: index + 1,
                name: item.name,
                item: item.url
            }))
        };
    }

    generateFaqStructuredData(entries, options = {}) {
        if (!Array.isArray(entries) || entries.length === 0) {
            return null;
        }

        const { url = null } = options;
        const data = {
            '@context': 'https://schema.org',
            '@type': 'FAQPage',
            inLanguage: 'en-US',
            mainEntity: entries.map((entry) => ({
                '@type': 'Question',
                name: entry.question,
                acceptedAnswer: {
                    '@type': 'Answer',
                    text: entry.answer
                }
            }))
        };

        if (url) {
            data['@id'] = `${url}#faq`;
        }

        return data;
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

        const baseDescription =
            permission.application?.description || permission.delegated?.description || '';

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
            author: this.getAuthorEntity(),
            publisher: this.getPublisherEntity(),
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
            structuredData.datePublished = options.datePublished || dateModified;
        }

        return structuredData;
    }

    /**
     * Answer-style questions rendered on the page and mirrored as FAQPage
     * structured data, so generative engines can quote the permission facts
     * without re-deriving them from tables.
     */
    buildPermissionFaqEntries(permission, view) {
        const value = permission.value;
        const entries = [
            {
                question: `What is the ${value} Microsoft Graph permission?`,
                answer: `${view.summaryText} ${this.describePermissionUsage(permission)}`.trim()
            },
            {
                question: `Is ${value} an application or a delegated permission?`,
                answer: this.describePermissionSupport(permission)
            },
            {
                question: `Does ${value} require admin consent?`,
                answer: view.consentText
            }
        ];

        const ids = [];
        if (permission.application?.id) {
            ids.push(`the application (app role) ID is ${permission.application.id}`);
        }
        if (permission.delegated?.id) {
            ids.push(`the delegated (OAuth2 scope) ID is ${permission.delegated.id}`);
        }

        if (ids.length > 0) {
            entries.push({
                question: `What is the permission ID (GUID) for ${value}?`,
                answer: `For ${value}, ${ids.join(', and ')}. These are the real Microsoft Graph identifiers and can be used directly in an app registration manifest.`
            });
        }

        return entries;
    }

    describePermissionUsage(permission) {
        const restCount =
            (permission.methods?.api?.v1?.length || 0) +
            (permission.methods?.api?.beta?.length || 0);
        const psCount =
            (permission.methods?.powershell?.v1?.length || 0) +
            (permission.methods?.powershell?.beta?.length || 0);
        const parts = [];

        if (permission.resource?.name) {
            parts.push(
                `It is mapped to the Microsoft Graph ${permission.resource.name} resource type.`
            );
        }

        if (restCount > 0) {
            parts.push(
                `${restCount} documented Microsoft Graph REST ${restCount === 1 ? 'method requires' : 'methods require'} it.`
            );
        }

        if (psCount > 0) {
            parts.push(
                `${psCount} Microsoft Graph PowerShell ${psCount === 1 ? 'command is' : 'commands are'} documented for it.`
            );
        }

        return parts.join(' ');
    }

    describePermissionSupport(permission) {
        const value = permission.value;

        if (permission.application && permission.delegated) {
            return `${value} is available as both an application permission (app-only access, no signed-in user) and a delegated permission (access on behalf of a signed-in user).`;
        }

        if (permission.application) {
            return `${value} is an application permission only. It grants app-only access without a signed-in user and always requires admin consent.`;
        }

        return `${value} is a delegated permission only. It grants access on behalf of a signed-in user and cannot be used for app-only access.`;
    }

    buildAppFaqEntries(app, sourceDoc) {
        const trustAnswer = app.isCommunity
            ? `${app.title} comes from the community-maintained list in this project. It is labeled separately from official Microsoft sources and should be verified before you rely on it.`
            : `${app.title} comes from an official Microsoft source (${app.sourceDisplayLabel || app.sourceLabel}), so the App ID is published by Microsoft rather than inferred.`;

        return [
            {
                question: `What is the application ID for ${app.title}?`,
                answer: `The Microsoft application (client) ID for ${app.title} is ${app.appId}.`
            },
            {
                question: `Is ${app.title} an official Microsoft first-party application?`,
                answer: trustAnswer
            },
            {
                question: `Where can I use the ${app.title} App ID?`,
                answer: `Use ${app.appId} to identify ${app.title} in Microsoft Entra ID sign-in logs, audit logs, service principal and enterprise application reviews, and conditional access policy targeting. Reference documentation: ${sourceDoc?.url || `${this.siteUrl}/microsoft-apps.html`}`
            }
        ];
    }

    buildHomepageFaqEntries(stats) {
        return [
            {
                question: 'What is a Microsoft Graph permission?',
                answer: 'A Microsoft Graph permission (also called a scope or app role) is a named authorization string such as User.Read or Mail.ReadWrite that an application requests in order to access a specific set of Microsoft 365 or Microsoft Entra ID data through the Microsoft Graph API. Each permission has a stable GUID that identifies it inside an app registration manifest.'
            },
            {
                question:
                    'What is the difference between application and delegated permissions in Microsoft Graph?',
                answer: 'A delegated permission lets an application act on behalf of a signed-in user, so the effective access is the intersection of the permission and what that user is already allowed to do. An application permission (app role) lets the application act on its own with no signed-in user, so it applies tenant-wide and always requires administrator consent.'
            },
            {
                question: 'Which Microsoft Graph permissions require admin consent?',
                answer: 'All application permissions require administrator consent. Delegated permissions require administrator consent when the consent type is Admin; delegated permissions with a User consent type can be granted by the user during sign-in. Each permission page on this site states the consent requirement explicitly.'
            },
            {
                question: 'How do I find the GUID of a Microsoft Graph permission?',
                answer: `Every permission page on this site publishes the real Microsoft Graph identifiers: the app role ID for the application permission and the OAuth2 permission scope ID for the delegated permission. All ${stats.permissions} permissions are also available as machine-readable JSON at ${this.siteUrl}/data/catalog/permissions.json.`
            },
            {
                question: 'What is a Microsoft first-party application ID?',
                answer: `A Microsoft first-party application ID is the client ID of an application published by Microsoft itself, such as Microsoft Graph PowerShell or Office 365 Exchange Online. Recognizing these IDs makes Entra ID sign-in logs and audit logs readable. This site catalogs ${stats.apps} of them with the source that each one came from.`
            },
            {
                question: 'Where does the data on this site come from?',
                answer: 'Permission data is refreshed from Microsoft Graph service principals, Microsoft Learn permission and PowerShell documentation, Microsoft Graph OpenAPI metadata, and the Microsoft Entra documentation known GUID catalog. Community-contributed application entries are labeled separately and never merged into the official sources. Snapshot freshness is published at /build-info.json.'
            }
        ];
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
        const trustText = app.isCommunity ? 'community-maintained' : 'official Microsoft';

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
                    '@id': `${this.siteUrl}/#website`,
                    name: this.siteName,
                    url: `${this.siteUrl}/`
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
                creator: this.getAuthorEntity(),
                publisher: this.getPublisherEntity(),
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

    generateHomepageTitle(_stats) {
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
            '@id': `${this.siteUrl}/#website`,
            name: this.siteName,
            alternateName: [
                'Microsoft Graph API Permissions Reference',
                'Graph API Permissions',
                'MS Graph Permissions'
            ],
            description: this.generateHomepageDescription(stats),
            url: `${this.siteUrl}/`,
            inLanguage: 'en-US',
            isAccessibleForFree: true,
            license: 'https://opensource.org/licenses/MIT',
            author: this.getAuthorEntity(),
            publisher: this.getAuthorEntity(),
            potentialAction: {
                '@type': 'SearchAction',
                target: {
                    '@type': 'EntryPoint',
                    urlTemplate: `${this.siteUrl}/?q={search_term_string}`
                },
                'query-input': 'required name=search_term_string'
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
                    '@id': `${this.siteUrl}/#website`,
                    name: this.siteName,
                    url: `${this.siteUrl}/`
                },
                mainEntity: {
                    '@type': 'Dataset',
                    name: 'Graph Permissions Explorer catalog',
                    description:
                        'Structured catalog of Microsoft Graph permissions, methods, code examples, and Microsoft first-party application IDs.',
                    url: `${this.siteUrl}/build-info.json`
                }
            },
            {
                '@context': 'https://schema.org',
                '@type': 'Dataset',
                name: 'Graph Permissions Explorer public catalog',
                description:
                    'Public JSON catalogs for Microsoft Graph permissions and Microsoft first-party application IDs.',
                url: `${this.siteUrl}/build-info.json`,
                license: 'https://opensource.org/licenses/MIT',
                creator: this.getAuthorEntity(),
                publisher: this.getPublisherEntity(),
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
                inLanguage: 'en-US',
                isPartOf: {
                    '@id': `${this.siteUrl}/#website`
                },
                author: this.getAuthorEntity(),
                publisher: this.getPublisherEntity(),
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
