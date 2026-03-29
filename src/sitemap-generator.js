/**
 * Sitemap Generator for Graph Permissions Explorer
 */

const fs = require('fs');
const path = require('path');

class SitemapGenerator {
    constructor(options = {}) {
        this.baseUrl = options.baseUrl || 'https://permissions.cengizyilmaz.net';
        this.outputDir = options.outputDir || './docs';
        this.maxUrlsPerSitemap = options.maxUrlsPerSitemap || 200;
    }

    getISODate(date = new Date()) {
        return date.toISOString().split('T')[0];
    }

    escapeXml(str) {
        if (!str) return '';
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&apos;');
    }

    generateUrlEntry(url, options = {}) {
        const {
            lastmod = this.getISODate(),
            changefreq = 'weekly',
            priority = 0.5,
            images = []
        } = options;

        let entry = '  <url>\n';
        entry += `    <loc>${this.escapeXml(url)}</loc>\n`;
        entry += `    <lastmod>${lastmod}</lastmod>\n`;
        entry += `    <changefreq>${changefreq}</changefreq>\n`;
        entry += `    <priority>${priority}</priority>\n`;

        images.forEach(img => {
            entry += '    <image:image>\n';
            entry += `      <image:loc>${this.escapeXml(img.url)}</image:loc>\n`;
            if (img.title) entry += `      <image:title>${this.escapeXml(img.title)}</image:title>\n`;
            if (img.caption) entry += `      <image:caption>${this.escapeXml(img.caption)}</image:caption>\n`;
            entry += '    </image:image>\n';
        });

        entry += '  </url>\n';
        return entry;
    }

    generatePermissionsSitemap(permissions, lastmod = this.getISODate()) {
        let urls = '';

        permissions.forEach(perm => {
            let priority = 0.7;
            if (perm.application && perm.delegated) priority = 0.8;
            if (perm.category.match(/User|Group|Mail|Files|Calendar/i)) priority = 0.85;

            urls += this.generateUrlEntry(
                `${this.baseUrl}/permissions/${perm.slug}.html`,
                {
                    lastmod,
                    changefreq: 'weekly',
                    priority: priority.toFixed(1)
                }
            );
        });

        return this.wrapSitemap(urls);
    }

    generateAppsOverviewSitemap(lastmod = this.getISODate()) {
        const urls = this.generateUrlEntry(
            `${this.baseUrl}/microsoft-apps.html`,
            {
                lastmod,
                changefreq: 'daily',
                priority: 0.9
            }
        );

        return this.wrapSitemap(urls);
    }

    generateAppDetailsSitemap(apps, lastmod = this.getISODate()) {
        let urls = '';

        apps.forEach((app) => {
            urls += this.generateUrlEntry(
                `${this.baseUrl}/apps/${app.anchor}.html`,
                {
                    lastmod,
                    changefreq: 'weekly',
                    priority: app.isCommunity ? 0.6 : 0.7
                }
            );
        });

        return this.wrapSitemap(urls);
    }

    generateMainSitemap(lastmod = this.getISODate()) {
        const urls = this.generateUrlEntry(
            `${this.baseUrl}/`,
            {
                lastmod,
                changefreq: 'daily',
                priority: 1.0,
                images: [
                    {
                        url: `${this.baseUrl}/og-image.png`,
                        title: 'Graph Permissions Explorer',
                        caption: 'Microsoft Graph API permissions reference'
                    }
                ]
            }
        );

        return this.wrapSitemap(urls);
    }

    generateSitemapIndex(sitemapEntries) {
        let content = `<?xml version="1.0" encoding="UTF-8"?>
<sitemapindex xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
`;

        sitemapEntries.forEach(entry => {
            content += `  <sitemap>
    <loc>${this.baseUrl}/${entry.file}</loc>
    <lastmod>${entry.lastmod}</lastmod>
  </sitemap>
`;
        });

        content += '</sitemapindex>';
        return content;
    }

    wrapSitemap(urls) {
        return `<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9"
        xmlns:image="http://www.google.com/schemas/sitemap-image/1.1">
${urls}</urlset>`;
    }

    chunkArray(array, chunkSize) {
        const chunks = [];
        for (let i = 0; i < array.length; i += chunkSize) {
            chunks.push(array.slice(i, i + chunkSize));
        }
        return chunks;
    }

    generate(permissions, apps, options = {}) {
        const sitemapEntries = [];
        const mainLastmod = options.mainLastmod || this.getISODate();
        const permissionsLastmod = options.permissionsLastmod || this.getISODate();
        const appsLastmod = options.appsLastmod || this.getISODate();

        fs.writeFileSync(
            path.join(this.outputDir, 'sitemap-main.xml'),
            this.generateMainSitemap(mainLastmod)
        );
        sitemapEntries.push({ file: 'sitemap-main.xml', lastmod: mainLastmod });
        console.log('  Generated sitemap-main.xml');

        const permChunks = this.chunkArray(permissions, this.maxUrlsPerSitemap);
        permChunks.forEach((chunk, index) => {
            const filename = permChunks.length === 1
                ? 'sitemap-permissions.xml'
                : `sitemap-permissions-${index + 1}.xml`;

            fs.writeFileSync(
                path.join(this.outputDir, filename),
                this.generatePermissionsSitemap(chunk, permissionsLastmod)
            );
            sitemapEntries.push({ file: filename, lastmod: permissionsLastmod });
        });
        console.log(`  Generated ${permChunks.length} permissions sitemap(s)`);

        fs.writeFileSync(
            path.join(this.outputDir, 'sitemap-apps.xml'),
            this.generateAppsOverviewSitemap(appsLastmod)
        );
        sitemapEntries.push({ file: 'sitemap-apps.xml', lastmod: appsLastmod });

        const appChunks = this.chunkArray(apps, this.maxUrlsPerSitemap);
        appChunks.forEach((chunk, index) => {
            const filename = appChunks.length === 1
                ? 'sitemap-app-details.xml'
                : `sitemap-app-details-${index + 1}.xml`;

            fs.writeFileSync(
                path.join(this.outputDir, filename),
                this.generateAppDetailsSitemap(chunk, appsLastmod)
            );
            sitemapEntries.push({ file: filename, lastmod: appsLastmod });
        });
        console.log(`  Generated apps sitemap (${appChunks.length} app detail sitemap(s) + overview)`);

        fs.writeFileSync(
            path.join(this.outputDir, 'sitemap.xml'),
            this.generateSitemapIndex(sitemapEntries)
        );
        console.log(`  Generated sitemap.xml (index with ${sitemapEntries.length} sitemaps)`);

        return sitemapEntries;
    }

    generateRobotsTxt() {
        const content = `# Graph Permissions Explorer - Robots.txt
# https://permissions.cengizyilmaz.net

User-agent: *
Allow: /
Allow: /llms.txt
Allow: /llms-full.txt

# Block development and data directories
Disallow: /data/
Disallow: /Script/
Disallow: /src/
Disallow: /customdata/
Disallow: /node_modules/

# Sitemap
Sitemap: ${this.baseUrl}/sitemap.xml
`;
        fs.writeFileSync(path.join(this.outputDir, 'robots.txt'), content);
        console.log('  Generated robots.txt');
    }
}

module.exports = SitemapGenerator;
