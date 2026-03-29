let selectedIndex = -1;
let searchResults = [];
let currentSearchQuery = '';
let currentFilter = 'all';
let currentSearch = '';
let searchIndexPromise = null;
let appsCatalogPromise = null;
let toastContainer;

const FAVORITES_KEY = 'graph_permissions_favorites';

function getBasePath() {
    return document.body?.dataset.basePath || '.';
}

function joinBasePath(relativePath) {
    const base = getBasePath().replace(/\/$/, '');
    return `${base}/${relativePath.replace(/^\.\//, '').replace(/^\//, '')}`;
}

async function fetchJson(relativePath) {
    const response = await fetch(joinBasePath(relativePath), { credentials: 'same-origin' });
    if (!response.ok) {
        throw new Error(`Failed to fetch ${relativePath}`);
    }
    return response.json();
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text || '';
    return div.innerHTML;
}

function initTheme() {
    const savedTheme = localStorage.getItem('theme') || 'system';
    setTheme(savedTheme);

    document.querySelectorAll('.theme-btn').forEach((button) => {
        button.addEventListener('click', () => setTheme(button.dataset.theme));
    });
}

function setTheme(theme) {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('theme', theme);
    document.querySelectorAll('.theme-btn').forEach((button) => {
        button.classList.toggle('active', button.dataset.theme === theme);
    });
}

function initSidebarToggle() {
    const toggleButton = document.querySelector('.sidebar-toggle');
    if (!toggleButton) {
        return;
    }

    const isCollapsed = localStorage.getItem('sidebarCollapsed') === 'true';
    document.body.classList.toggle('sidebar-collapsed', isCollapsed);

    toggleButton.addEventListener('click', () => {
        const collapsed = document.body.classList.toggle('sidebar-collapsed');
        localStorage.setItem('sidebarCollapsed', String(collapsed));
    });
}

async function ensureSearchIndexLoaded() {
    if (window.SEARCH_INDEX) {
        return window.SEARCH_INDEX;
    }

    if (!searchIndexPromise) {
        searchIndexPromise = Promise.all([
            fetchJson('data/catalog/permissions.json'),
            fetchJson('data/catalog/apps-manifest.json')
        ]).then(([permissionsCatalog, appsManifest]) => {
            const detailBasePath = appsManifest.detailBasePath || 'apps/';
            window.SEARCH_INDEX = [
                ...((permissionsCatalog.items || []).map((item) => ({
                    type: 'permission',
                    title: item[0],
                    url: `permissions/${item[1]}.html`,
                    category: item[2],
                    description: item[2]
                }))),
                ...((appsManifest.searchIndex || []).map((item) => {
                    return {
                        type: 'app',
                        title: item[0],
                        appId: item[1],
                        url: `${detailBasePath}${item[2]}.html`,
                        description: `App ID: ${item[1]}`
                    };
                }))
            ];
            return window.SEARCH_INDEX;
        });
    }

    return searchIndexPromise;
}

async function openSearch() {
    const modal = document.getElementById('search-modal');
    const input = document.getElementById('search-input');
    const results = document.getElementById('search-results');
    if (!modal || !input || !results) {
        return;
    }

    modal.classList.add('open');
    input.focus();
    input.value = '';
    currentSearchQuery = '';
    renderSearchResults([]);
    results.innerHTML = '<div class="search-empty">Loading refreshed search catalog...</div>';

    try {
        await ensureSearchIndexLoaded();
        renderSearchResults([]);
    } catch {
        results.innerHTML = '<div class="search-empty no-results">Search is temporarily unavailable.</div>';
    }
}

function closeSearch() {
    const modal = document.getElementById('search-modal');
    if (modal) {
        modal.classList.remove('open');
    }
    selectedIndex = -1;
}

function scoreSearchResult(item, query) {
    const title = (item.title || '').toLowerCase();
    const description = (item.description || '').toLowerCase();
    const appId = (item.appId || '').toLowerCase();
    const category = (item.category || '').toLowerCase();
    let score = 0;

    if (title === query || appId === query) score += 120;
    if (title.startsWith(query) || appId.startsWith(query)) score += 90;
    if (title.includes(query)) score += 60;
    if (appId.includes(query)) score += 50;
    if (description.includes(query)) score += 20;
    if (category.includes(query)) score += 10;

    return score;
}

function renderSearchResults(results) {
    searchResults = results;
    const container = document.getElementById('search-results');
    if (!container) {
        return;
    }

    if (results.length === 0) {
        container.innerHTML = currentSearchQuery
            ? `<div class="search-empty no-results">No results found for <strong>${escapeHtml(currentSearchQuery)}</strong>.</div>`
            : '<div class="search-empty">Type to search...</div>';
        return;
    }

    container.innerHTML = results.slice(0, 20).map((item, index) => `
        <div class="search-result ${index === selectedIndex ? 'selected' : ''}"
             onclick="navigateTo('${item.url}')"
             data-index="${index}">
            <span class="search-result-title">${escapeHtml(item.title)}</span>
            <span class="search-result-desc">${escapeHtml(item.description || '').substring(0, 90)}</span>
            <span class="search-result-type">${escapeHtml(item.type)}</span>
        </div>
    `).join('');
}

function performSearch(query) {
    currentSearchQuery = (query || '').trim();
    if (!currentSearchQuery || !window.SEARCH_INDEX) {
        renderSearchResults([]);
        return;
    }

    const lowered = currentSearchQuery.toLowerCase();
    const results = window.SEARCH_INDEX
        .map((item) => ({ item, score: scoreSearchResult(item, lowered) }))
        .filter((entry) => entry.score > 0)
        .sort((left, right) => right.score - left.score || left.item.title.localeCompare(right.item.title))
        .slice(0, 20)
        .map((entry) => entry.item);

    selectedIndex = results.length > 0 ? 0 : -1;
    renderSearchResults(results);
}

function navigateTo(url) {
    window.location.href = joinBasePath(url);
}

function updateSelection() {
    document.querySelectorAll('.search-result').forEach((element, index) => {
        element.classList.toggle('selected', index === selectedIndex);
    });

    const selected = document.querySelector('.search-result.selected');
    if (selected) {
        selected.scrollIntoView({ block: 'nearest' });
    }
}

function updateCanonicalUrl() {
    const canonical = document.querySelector('link[rel="canonical"]');
    if (canonical) {
        canonical.href = canonical.href.split('#')[0];
    }
}

function highlightHashTarget() {
    if (!window.location.hash) {
        return;
    }

    const target = document.getElementById(window.location.hash.substring(1));
    if (!target) {
        return;
    }

    setTimeout(() => {
        target.scrollIntoView({ behavior: 'smooth', block: 'start' });
        target.classList.add('highlight');
        setTimeout(() => target.classList.remove('highlight'), 2000);
    }, 200);
}

function filterSidebar(query) {
    const lowered = (query || '').toLowerCase();
    document.querySelectorAll('.nav-category').forEach((category) => {
        const items = category.querySelectorAll('.nav-item');
        let hasVisibleItem = false;

        items.forEach((item) => {
            const haystack = [
                item.dataset.permission,
                item.dataset.appid,
                item.dataset.appname,
                item.textContent
            ].filter(Boolean).join(' ').toLowerCase();
            const matches = !lowered || haystack.includes(lowered);
            item.closest('li').style.display = matches ? '' : 'none';
            if (matches) {
                hasVisibleItem = true;
            }
        });

        category.style.display = hasVisibleItem || !lowered ? '' : 'none';
        if (lowered && hasVisibleItem) {
            category.classList.add('expanded');
        }
    });
}

function toggleCategory(button) {
    button.closest('.nav-category')?.classList.toggle('expanded');
}

function copyToClipboard(text) {
    navigator.clipboard.writeText(text).then(() => {
        showToast('Copied to clipboard!', 'success');
    }).catch(() => {
        showToast('Copy failed', 'error');
    });
}

function copyAppIdOnClick(event, appId) {
    event.preventDefault();
    copyToClipboard(appId);

    const link = event.currentTarget;
    const href = link.getAttribute('href');
    if (!href) {
        return;
    }

    window.history.pushState({}, '', `${window.location.pathname}${href}`);
    updateCanonicalUrl();

    const target = document.querySelector(href);
    if (target) {
        target.scrollIntoView({ behavior: 'smooth', block: 'start' });
        target.classList.add('highlight');
        setTimeout(() => target.classList.remove('highlight'), 2000);
    }

    document.querySelectorAll('.nav-item').forEach((item) => item.classList.remove('active'));
    link.classList.add('active');
}

function copyCode(button) {
    const container = button.closest('.code-block, .json-block');
    const code = container?.querySelector('code')?.textContent;
    if (!code) {
        return;
    }

    navigator.clipboard.writeText(code).then(() => {
        button.textContent = 'Copied!';
        button.classList.add('copied');
        setTimeout(() => {
            button.textContent = 'Copy';
            button.classList.remove('copied');
        }, 1500);
    });
}

function initTabs() {
    document.querySelectorAll('.tab-btn').forEach((button) => {
        button.addEventListener('click', () => {
            const container = button.closest('.code-tabs');
            if (!container) {
                return;
            }

            container.querySelectorAll('.tab-btn').forEach((item) => item.classList.remove('active'));
            container.querySelectorAll('.tab-content').forEach((item) => item.classList.remove('active'));
            button.classList.add('active');
            container.querySelector(`#tab-${button.dataset.tab}`)?.classList.add('active');
        });
    });

    document.querySelectorAll('.method-tab-btn').forEach((button) => {
        button.addEventListener('click', () => {
            const container = button.closest('.method-tabs');
            if (!container) {
                return;
            }

            container.querySelectorAll('.method-tab-btn').forEach((item) => item.classList.remove('active'));
            container.querySelectorAll('.method-tab-content').forEach((item) => item.classList.remove('active'));
            button.classList.add('active');
            container.querySelector(`#method-tab-${button.dataset.methodTab}`)?.classList.add('active');
        });
    });
}

function updateAppsCount(visibleCount) {
    const title = document.querySelector('.apps-table-section h2');
    if (!title) {
        return;
    }

    let badge = title.querySelector('.count-badge');
    if (!badge) {
        badge = document.createElement('span');
        badge.className = 'count-badge';
        title.appendChild(badge);
    }

    badge.textContent = `${visibleCount} shown`;
}

function updateAppsEmptyState(visibleCount) {
    const section = document.querySelector('.apps-table-section');
    const container = section?.querySelector('.table-container');
    if (!section || !container) {
        return;
    }

    let emptyState = section.querySelector('.apps-empty-state');
    if (!emptyState) {
        emptyState = document.createElement('div');
        emptyState.className = 'apps-empty-state';
        emptyState.innerHTML = `
            <svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"></circle><path d="m21 21-4.3-4.3"></path></svg>
            <div>
                <strong>No applications match the current filters.</strong>
                <p>Adjust the source filter or clear the search query to see more results.</p>
            </div>
        `;
        section.appendChild(emptyState);
    }

    emptyState.hidden = visibleCount > 0;
    container.hidden = visibleCount === 0;
}

function buildAppPortalUrl(appId) {
    return `https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/${appId}`;
}

function renderAppsRows(apps, detailBasePath = 'apps/') {
    const tbody = document.getElementById('apps-tbody');
    if (!tbody) {
        return;
    }

    tbody.innerHTML = apps.map((app) => {
        const detailUrl = joinBasePath(`${detailBasePath}${app.anchor}.html`);
        const portalUrl = buildAppPortalUrl(app.appId);
        const sourceClass = app.isCommunity ? 'custom' : app.source;
        const searchableSource = [
            app.source || '',
            app.sourceLabel || '',
            app.sourceDisplayLabel || '',
            ...((app.sourceProvenance || []).map((item) => String(item || ''))),
            ...((app.sourceProvenanceLabels || []).map((item) => String(item || '')))
        ].join(' ').toLowerCase();

        return `
            <tr id="${escapeHtml(app.anchor)}"
                data-appid="${escapeHtml(app.appId)}"
                data-name="${escapeHtml((app.title || '').toLowerCase())}"
                data-source="${escapeHtml(searchableSource)}"
                data-filter-groups="${escapeHtml(((app.filterGroups && app.filterGroups.length > 0) ? app.filterGroups : [app.filterGroup || 'all']).join('|'))}">
                <td class="app-name">
                    <a href="${detailUrl}" class="app-name-link">
                        <svg viewBox="0 0 24 24"><rect x="2" y="3" width="20" height="14" rx="2"/><path d="M8 21h8m-4-4v4"/></svg>
                        ${escapeHtml(app.title)}
                    </a>
                </td>
                <td><code class="id-code" onclick="copyToClipboard('${escapeHtml(app.appId)}')" title="Click to copy">${escapeHtml(app.appId)}</code></td>
                <td><span class="source-tag ${escapeHtml(sourceClass)}">${escapeHtml(app.sourceDisplayLabel || app.sourceLabel)}</span></td>
                <td class="action-col">
                    <div class="action-btns">
                        <a href="${detailUrl}" title="Open detail page">
                            <svg viewBox="0 0 24 24"><path d="M5 12h14"/><path d="m12 5 7 7-7 7"/></svg>
                        </a>
                        <button onclick="copyToClipboard('${escapeHtml(app.appId)}')" title="Copy App ID">
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

function applyAppsSourceVisibility(manifest) {
    const counts = manifest?.counts || {};
    const learnVisible = Number(counts.learn || 0) > 0;
    const communityVisible = Number(counts.community || 0) > 0;

    document.querySelectorAll('.source-learn').forEach((element) => {
        element.hidden = !learnVisible;
    });

    document.querySelectorAll('.source-community').forEach((element) => {
        element.hidden = !communityVisible;
    });
}

async function loadAppsCatalog() {
    const appsPage = document.querySelector('[data-apps-manifest]');
    if (!appsPage) {
        return [];
    }

    if (!appsCatalogPromise) {
        appsCatalogPromise = (async () => {
            const manifestPath = appsPage.dataset.appsManifest;
            const manifest = await fetchJson(manifestPath);
            applyAppsSourceVisibility(manifest);
            const manifestDir = manifestPath.split('/').slice(0, -1).join('/');
            const chunks = await Promise.all(
                (manifest.chunks || []).map((chunk) => fetchJson(`${manifestDir}/${chunk.file}`))
            );
            const apps = chunks.flatMap((chunk) => chunk.items || []);
            renderAppsRows(apps, manifest.detailBasePath || 'apps/');
            applyFilters();
            return apps;
        })().catch((error) => {
            const tbody = document.getElementById('apps-tbody');
            if (tbody) {
                tbody.innerHTML = `
                    <tr class="apps-loading-row">
                        <td colspan="4">
                            <div class="apps-loading-state apps-loading-error">
                                <span>Failed to load the refreshed apps catalog.</span>
                            </div>
                        </td>
                    </tr>
                `;
            }
            throw error;
        });
    }

    return appsCatalogPromise;
}

function applyFilters() {
    const rows = Array.from(document.querySelectorAll('#apps-tbody tr')).filter((row) => row.dataset.appid);
    if (!rows.length) {
        return;
    }

    let visibleCount = 0;
    rows.forEach((row) => {
        const filterGroups = (row.dataset.filterGroups || '').split('|').filter(Boolean);
        const matchesFilter = currentFilter === 'all' || filterGroups.includes(currentFilter);
        const name = row.dataset.name || '';
        const appId = row.dataset.appid || '';
        const source = row.dataset.source || '';
        const matchesSearch = !currentSearch || name.includes(currentSearch) || appId.includes(currentSearch) || source.includes(currentSearch);
        const visible = matchesFilter && matchesSearch;
        row.style.display = visible ? '' : 'none';
        if (visible) {
            visibleCount += 1;
        }
    });

    updateAppsCount(visibleCount);
    updateAppsEmptyState(visibleCount);
}

function filterApps(query) {
    currentSearch = (query || '').toLowerCase().trim();
    applyFilters();
}

function filterAppsBySource(source) {
    currentFilter = source;
    applyFilters();
}

function getFavorites() {
    try {
        return JSON.parse(localStorage.getItem(FAVORITES_KEY) || '[]');
    } catch {
        return [];
    }
}

function saveFavorites(favorites) {
    localStorage.setItem(FAVORITES_KEY, JSON.stringify(favorites));
}

function isFavorite(permissionValue) {
    return getFavorites().includes(permissionValue);
}

function toggleFavorite(permissionValue) {
    const favorites = getFavorites();
    const index = favorites.indexOf(permissionValue);

    if (index >= 0) {
        favorites.splice(index, 1);
        showToast('Removed from favorites', 'info');
    } else {
        favorites.push(permissionValue);
        showToast('Added to favorites', 'success');
    }

    saveFavorites(favorites);
    updateFavoriteButtons();
}

function updateFavoriteButtons() {
    document.querySelectorAll('.favorite-btn').forEach((button) => {
        const active = isFavorite(button.dataset.permission);
        button.classList.toggle('active', active);
        button.setAttribute('aria-pressed', String(active));
    });
}

function initFavorites() {
    document.querySelectorAll('.favorite-btn').forEach((button) => {
        button.addEventListener('click', (event) => {
            event.preventDefault();
            event.stopPropagation();
            toggleFavorite(button.dataset.permission);
        });
    });
    updateFavoriteButtons();
}

function exportCurrentPage(format) {
    const permissionValue = document.querySelector('.permission-title-row h1 code')?.textContent;
    if (!permissionValue) {
        showToast('No permission data to export', 'error');
        return;
    }

    const payload = [{
        permission: permissionValue,
        category: document.querySelector('.breadcrumb span')?.textContent || '',
        description: document.querySelector('.permission-desc')?.textContent || '',
        applicationId: document.querySelector('.detail-card-app .id-code')?.textContent || '',
        delegatedId: document.querySelector('.detail-card-delegated .id-code')?.textContent || '',
        exportedAt: new Date().toISOString()
    }];

    if (format === 'json') {
        downloadBlob(new Blob([JSON.stringify(payload[0], null, 2)], { type: 'application/json' }), `${permissionValue.toLowerCase().replace(/\./g, '-')}.json`);
    } else {
        const csv = [
            Object.keys(payload[0]).join(','),
            Object.values(payload[0]).map((value) => `"${String(value).replace(/"/g, '""')}"`).join(',')
        ].join('\n');
        downloadBlob(new Blob([csv], { type: 'text/csv;charset=utf-8;' }), `${permissionValue.toLowerCase().replace(/\./g, '-')}.csv`);
    }

    showToast(`Exported as ${format.toUpperCase()}`, 'success');
}

function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

function initExportMenu() {
    document.querySelectorAll('.export-menu-item').forEach((item) => {
        item.addEventListener('click', () => {
            const action = item.dataset.action;
            if (action === 'json') exportCurrentPage('json');
            if (action === 'csv') exportCurrentPage('csv');
            if (action === 'url') copyToClipboard(window.location.href);
            if (action === 'print') window.print();
        });
    });
}

function initToastNotifications() {
    toastContainer = document.getElementById('toast-container');
    if (toastContainer) {
        return;
    }

    toastContainer = document.createElement('div');
    toastContainer.id = 'toast-container';
    toastContainer.className = 'toast-container';
    document.body.appendChild(toastContainer);
}

function showToast(message, type = 'info') {
    if (!toastContainer) {
        initToastNotifications();
    }

    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.innerHTML = `
        <span class="toast-message">${escapeHtml(message)}</span>
        <button class="toast-close" onclick="this.parentElement.remove()">&times;</button>
    `;

    toastContainer.appendChild(toast);
    setTimeout(() => {
        toast.style.animation = 'fadeOut 0.3s ease forwards';
        setTimeout(() => toast.remove(), 300);
    }, 2500);
}

document.addEventListener('keydown', async (event) => {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'k') {
        event.preventDefault();
        await openSearch();
        return;
    }

    const modal = document.getElementById('search-modal');
    if (modal?.classList.contains('open')) {
        if (event.key === 'Escape') {
            closeSearch();
        } else if (event.key === 'ArrowDown') {
            event.preventDefault();
            if (selectedIndex < searchResults.length - 1) {
                selectedIndex += 1;
                updateSelection();
            }
        } else if (event.key === 'ArrowUp') {
            event.preventDefault();
            if (selectedIndex > 0) {
                selectedIndex -= 1;
                updateSelection();
            }
        } else if (event.key === 'Enter' && searchResults[selectedIndex]) {
            event.preventDefault();
            navigateTo(searchResults[selectedIndex].url);
        }
    }
});

document.addEventListener('DOMContentLoaded', async () => {
    initTheme();
    initSidebarToggle();
    initTabs();
    initFavorites();
    initExportMenu();
    initToastNotifications();

    const searchInput = document.getElementById('search-input');
    if (searchInput) {
        searchInput.addEventListener('input', async (event) => {
            if (event.target.value.trim() && !window.SEARCH_INDEX) {
                await ensureSearchIndexLoaded();
            }
            performSearch(event.target.value);
        });
    }

    const sidebarFilter = document.getElementById('sidebar-filter');
    if (sidebarFilter) {
        sidebarFilter.addEventListener('input', (event) => filterSidebar(event.target.value));
    }

    const appsSearch = document.getElementById('apps-search');
    if (appsSearch) {
        let timer;
        appsSearch.addEventListener('input', (event) => {
            clearTimeout(timer);
            timer = setTimeout(() => filterApps(event.target.value), 100);
        });
    }

    document.querySelectorAll('.filter-btn').forEach((button) => {
        button.addEventListener('click', () => {
            document.querySelectorAll('.filter-btn').forEach((item) => item.classList.remove('active'));
            button.classList.add('active');
            filterAppsBySource(button.dataset.filter);
        });
    });

    try {
        await loadAppsCatalog();
    } catch {
        showToast('Apps catalog could not be loaded.', 'error');
    }

    applyFilters();
    highlightHashTarget();
    updateCanonicalUrl();
    window.addEventListener('hashchange', () => {
        updateCanonicalUrl();
        highlightHashTarget();
    });

    initScrollToTop();
    initTableOverflowHints();

    const footerYear = document.getElementById('footer-year');
    if (footerYear) {
        footerYear.textContent = new Date().getFullYear();
    }
});

function initScrollToTop() {
    const btn = document.getElementById('scroll-top-btn');
    if (!btn) return;

    let ticking = false;
    window.addEventListener('scroll', () => {
        if (!ticking) {
            window.requestAnimationFrame(() => {
                btn.classList.toggle('visible', window.scrollY > 400);
                ticking = false;
            });
            ticking = true;
        }
    }, { passive: true });

    btn.addEventListener('click', () => {
        window.scrollTo({ top: 0, behavior: 'smooth' });
    });
}

function initTableOverflowHints() {
    document.querySelectorAll('.table-container').forEach((container) => {
        const checkOverflow = () => {
            container.classList.toggle('has-overflow', container.scrollWidth > container.clientWidth);
        };
        checkOverflow();
        window.addEventListener('resize', checkOverflow, { passive: true });
    });
}

window.openSearch = openSearch;
window.closeSearch = closeSearch;
window.navigateTo = navigateTo;
window.toggleCategory = toggleCategory;
window.copyToClipboard = copyToClipboard;
window.copyAppIdOnClick = copyAppIdOnClick;
window.copyCode = copyCode;
