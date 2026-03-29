const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

function ensureDir(dirPath) {
    if (!fs.existsSync(dirPath)) {
        fs.mkdirSync(dirPath, { recursive: true });
    }
}

function cleanDir(dirPath) {
    fs.rmSync(dirPath, { recursive: true, force: true });
    fs.mkdirSync(dirPath, { recursive: true });
}

function readUtf8(filePath) {
    let content = fs.readFileSync(filePath, 'utf8');
    if (content.charCodeAt(0) === 0xFEFF) {
        content = content.slice(1);
    }
    return content;
}

function loadJson(filePath, fallback = null) {
    if (!fs.existsSync(filePath)) {
        return fallback;
    }

    const content = readUtf8(filePath).trim();
    if (!content) {
        return fallback;
    }

    return JSON.parse(content);
}

function getJsonShardPaths(filePath) {
    const directory = path.dirname(filePath);
    const extension = path.extname(filePath);
    const baseName = path.basename(filePath, extension);

    if (!fs.existsSync(directory)) {
        return [];
    }

    const shardPattern = new RegExp(`^${baseName}\\.part-\\d{3}${extension.replace('.', '\\.')}$`);

    return fs.readdirSync(directory)
        .filter((entry) => shardPattern.test(entry))
        .sort((left, right) => left.localeCompare(right))
        .map((entry) => path.join(directory, entry));
}

function mergeJsonParts(parts) {
    if (parts.length === 0) {
        return null;
    }

    if (parts.every(Array.isArray)) {
        return parts.flat();
    }

    if (parts.every((part) => part && typeof part === 'object' && !Array.isArray(part))) {
        return Object.assign({}, ...parts);
    }

    throw new Error('JSON shard files must all be arrays or all be objects.');
}

function loadJsonWithShards(filePath, fallback = null) {
    if (fs.existsSync(filePath)) {
        return loadJson(filePath, fallback);
    }

    const shardPaths = getJsonShardPaths(filePath);
    if (shardPaths.length === 0) {
        return fallback;
    }

    const parts = shardPaths.map((shardPath) => loadJson(shardPath, null));
    return mergeJsonParts(parts);
}

function writeJson(filePath, value) {
    ensureDir(path.dirname(filePath));
    fs.writeFileSync(filePath, `${JSON.stringify(value, null, 2)}\n`);
}

function writeJsonSharded(filePath, value, options = {}) {
    const maxEntriesPerPart = options.maxEntriesPerPart || 150;
    const directory = path.dirname(filePath);
    const shardPaths = getJsonShardPaths(filePath);

    ensureDir(directory);

    shardPaths.forEach((shardPath) => {
        fs.rmSync(shardPath, { force: true });
    });

    fs.rmSync(filePath, { force: true });

    if (Array.isArray(value)) {
        if (value.length <= maxEntriesPerPart) {
            writeJson(filePath, value);
            return { mode: 'single', parts: 1 };
        }

        let partIndex = 0;
        for (let index = 0; index < value.length; index += maxEntriesPerPart) {
            partIndex += 1;
            const shardPath = filePath.replace(/\.json$/i, `.part-${String(partIndex).padStart(3, '0')}.json`);
            writeJson(shardPath, value.slice(index, index + maxEntriesPerPart));
        }

        return { mode: 'sharded', parts: partIndex };
    }

    if (value && typeof value === 'object') {
        const entries = Object.entries(value);
        if (entries.length <= maxEntriesPerPart) {
            writeJson(filePath, value);
            return { mode: 'single', parts: 1 };
        }

        let partIndex = 0;
        for (let index = 0; index < entries.length; index += maxEntriesPerPart) {
            partIndex += 1;
            const shardPath = filePath.replace(/\.json$/i, `.part-${String(partIndex).padStart(3, '0')}.json`);
            writeJson(shardPath, Object.fromEntries(entries.slice(index, index + maxEntriesPerPart)));
        }

        return { mode: 'sharded', parts: partIndex };
    }

    writeJson(filePath, value);
    return { mode: 'single', parts: 1 };
}

function escapeHtml(text) {
    if (text === null || text === undefined) {
        return '';
    }

    return String(text)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}

function slugify(text) {
    if (!text) {
        return '';
    }

    return String(text)
        .toLowerCase()
        .trim()
        .replace(/\s+/g, '-')
        .replace(/\./g, '-')
        .replace(/_/g, '-')
        .replace(/[^a-z0-9-]/g, '')
        .replace(/-+/g, '-')
        .replace(/^-|-$/g, '');
}

function hashContent(value) {
    return crypto.createHash('sha256').update(value).digest('hex');
}

function formatUtcLabel(isoString) {
    const value = isoString || new Date().toISOString();

    return `${new Intl.DateTimeFormat('en-US', {
        dateStyle: 'long',
        timeStyle: 'short',
        timeZone: 'UTC'
    }).format(new Date(value))} UTC`;
}

function fileDate(filePath) {
    return fs.statSync(filePath).mtime.toISOString();
}

function latestJsonDate(filePath) {
    if (fs.existsSync(filePath)) {
        return fileDate(filePath);
    }

    const shardPaths = getJsonShardPaths(filePath);
    if (shardPaths.length === 0) {
        return null;
    }

    return shardPaths
        .map((shardPath) => fs.statSync(shardPath).mtime.toISOString())
        .sort()
        .slice(-1)[0];
}

function fileDateOnly(filePath) {
    return fileDate(filePath).split('T')[0];
}

module.exports = {
    cleanDir,
    ensureDir,
    escapeHtml,
    fileDate,
    fileDateOnly,
    formatUtcLabel,
    getJsonShardPaths,
    hashContent,
    loadJson,
    loadJsonWithShards,
    latestJsonDate,
    readUtf8,
    slugify,
    writeJson,
    writeJsonSharded
};
