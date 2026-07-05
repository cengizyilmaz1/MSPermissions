'use strict';

/**
 * Node port of the OpenAPI permission miner (formerly Parse-GraphOpenAPI.ps1).
 *
 * The Microsoft Graph OpenAPI specifications tag every operation with the
 * resource entity it belongs to (e.g. `accessReviewDecisions.accessReviewDecision`).
 * This module scans the v1.0 and beta specifications line by line and records
 * `(Permission, Method, Endpoint, ApiVersion)` tuples, writing the exact same
 * `permission.csv` contract the PowerShell script produced — but with no external
 * runtime dependency.
 *
 * The scan intentionally mirrors the original text-mining rules so the output
 * stays byte-for-byte compatible with the existing dataset and the normalizer in
 * `src/lib/site-data.js`.
 */

const fs = require('fs');
const os = require('os');
const path = require('path');
const https = require('https');
const readline = require('readline');

const SOURCES = [
    {
        name: 'v1.0',
        url: 'https://raw.githubusercontent.com/microsoftgraph/msgraph-metadata/master/openapi/v1.0/openapi.yaml',
        file: 'openapi_graph_v1.yaml',
        meta: 'openapi_graph_v1.meta.json'
    },
    {
        name: 'beta',
        url: 'https://raw.githubusercontent.com/microsoftgraph/msgraph-metadata/master/openapi/beta/openapi.yaml',
        file: 'openapi_graph_beta.yaml',
        meta: 'openapi_graph_beta.meta.json'
    }
];

const PATH_RE = /^['"]?(\/[a-zA-Z0-9/{}_.-]+)['"]?:$/;
const METHOD_RE = /^(get|post|put|patch|delete):$/i;
const TAG_RE = /^-\s+([a-zA-Z][a-zA-Z0-9]+(?:\.[a-zA-Z0-9]+)+)$/;
const CACHE_MAX_AGE_MS = 24 * 60 * 60 * 1000;

function resolveCacheDir() {
    const base =
        process.env.GRAPH_OPENAPI_CACHE_DIR || path.join(os.tmpdir(), 'graph-openapi-cache');
    fs.mkdirSync(base, { recursive: true });
    return base;
}

function downloadFile(url, destination) {
    return new Promise((resolve, reject) => {
        const request = (target, redirects) => {
            if (redirects > 5) {
                reject(new Error(`Too many redirects for ${url}`));
                return;
            }

            https
                .get(target, (response) => {
                    const status = response.statusCode || 0;
                    if (status >= 300 && status < 400 && response.headers.location) {
                        response.resume();
                        request(
                            new URL(response.headers.location, target).toString(),
                            redirects + 1
                        );
                        return;
                    }
                    if (status !== 200) {
                        response.resume();
                        reject(new Error(`HTTP ${status} while downloading ${target}`));
                        return;
                    }

                    const stream = fs.createWriteStream(destination);
                    response.pipe(stream);
                    stream.on('finish', () => stream.close(() => resolve(destination)));
                    stream.on('error', reject);
                })
                .on('error', reject);
        };

        request(url, 0);
    });
}

async function ensureSpecFile(source, cacheDir, log) {
    const specPath = path.join(cacheDir, source.file);
    const metaPath = path.join(cacheDir, source.meta);

    if (fs.existsSync(specPath) && fs.existsSync(metaPath)) {
        try {
            const meta = JSON.parse(fs.readFileSync(metaPath, 'utf8'));
            const age = Date.now() - new Date(meta.downloaded).getTime();
            if (age < CACHE_MAX_AGE_MS) {
                log.info(`Using cached ${source.name} spec (age ${Math.round(age / 3600000)}h).`);
                return specPath;
            }
        } catch {
            // fall through to re-download
        }
    }

    log.info(`Downloading ${source.name} OpenAPI specification...`);
    try {
        await downloadFile(source.url, specPath);
        fs.writeFileSync(
            metaPath,
            JSON.stringify({ downloaded: new Date().toISOString(), url: source.url })
        );
    } catch (error) {
        if (fs.existsSync(specPath)) {
            log.warn(`Download failed (${error.message}); using cached ${source.name} spec.`);
            return specPath;
        }
        throw error;
    }

    return specPath;
}

/**
 * Create a stateful line processor that mines `(Permission, Method, Endpoint,
 * ApiVersion)` tuples and adds them (as JSON strings) to `results`.
 *
 * Extracted as a pure factory so the mining rules can be unit tested without
 * touching the network or the file system.
 *
 * @param {string} apiName - API version label, e.g. `v1.0` or `beta`.
 * @param {Set<string>} results - Sink of JSON-encoded tuples.
 * @returns {(line: string) => void}
 */
function createLineMiner(apiName, results) {
    let currentPath = 'Unknown';
    let currentMethod = 'Unknown';

    return (rawLine) => {
        const line = rawLine.trim();

        const pathMatch = line.match(PATH_RE);
        if (pathMatch) {
            currentPath = pathMatch[1];
            currentMethod = 'Unknown';
            return;
        }

        const methodMatch = line.match(METHOD_RE);
        if (methodMatch) {
            currentMethod = methodMatch[1].toUpperCase();
            return;
        }

        const tagMatch = line.match(TAG_RE);
        if (tagMatch && currentPath !== 'Unknown' && currentMethod !== 'Unknown') {
            const permission = tagMatch[1];
            if (!/microsoft\.graph/i.test(permission) && permission.length < 60) {
                results.add(JSON.stringify([permission, currentMethod, currentPath, apiName]));
            }
        }
    };
}

/**
 * Mine tuples from an in-memory iterable of lines. Synchronous and pure.
 *
 * @param {Iterable<string>} lines - Specification lines.
 * @param {string} apiName - API version label.
 * @param {Set<string>} [results] - Optional existing sink to append to.
 * @returns {Set<string>} The sink of JSON-encoded tuples.
 */
function mineLines(lines, apiName, results = new Set()) {
    const processLine = createLineMiner(apiName, results);
    for (const line of lines) {
        processLine(line);
    }
    return results;
}

async function mineSpec(specPath, apiName, results) {
    const rl = readline.createInterface({
        input: fs.createReadStream(specPath),
        crlfDelay: Infinity
    });

    const processLine = createLineMiner(apiName, results);
    for await (const rawLine of rl) {
        processLine(rawLine);
    }
}

function compareTuples(left, right) {
    for (let index = 0; index < left.length; index += 1) {
        const a = left[index].toLowerCase();
        const b = right[index].toLowerCase();
        if (a < b) {
            return -1;
        }
        if (a > b) {
            return 1;
        }
    }
    return 0;
}

function toCsv(tuples) {
    const header = '"Permission","Method","Endpoint","ApiVersion"';
    const rows = tuples.map((tuple) => tuple.map((value) => `"${value}"`).join(','));
    return [header, ...rows].join('\n') + '\n';
}

/**
 * Convert a sink of JSON-encoded tuples into a sorted array of tuples.
 *
 * @param {Set<string>} results
 * @returns {string[][]}
 */
function sortTuples(results) {
    return Array.from(results, (value) => JSON.parse(value)).sort(compareTuples);
}

/**
 * Build `permission.csv` from the live Graph OpenAPI specifications.
 *
 * @param {string} outputDir - Directory to write `permission.csv` into.
 * @param {object} [options]
 * @param {object} [options.logger] - Logger with info/warn methods.
 * @returns {Promise<{ path: string, records: number, permissions: number, v1: number, beta: number }>}
 */
async function buildPermissionCsv(outputDir, options = {}) {
    const log = options.logger || { info() {}, warn() {} };
    const cacheDir = resolveCacheDir();
    const results = new Set();

    for (const source of SOURCES) {
        const specPath = await ensureSpecFile(source, cacheDir, log);
        await mineSpec(specPath, source.name, results);
    }

    const tuples = sortTuples(results);

    fs.mkdirSync(outputDir, { recursive: true });
    const outputPath = path.join(outputDir, 'permission.csv');
    fs.writeFileSync(outputPath, toCsv(tuples));

    const permissions = new Set(tuples.map((tuple) => tuple[0]));
    return {
        path: outputPath,
        records: tuples.length,
        permissions: permissions.size,
        v1: tuples.filter((tuple) => tuple[3] === 'v1.0').length,
        beta: tuples.filter((tuple) => tuple[3] === 'beta').length
    };
}

module.exports = {
    buildPermissionCsv,
    SOURCES,
    mineLines,
    sortTuples,
    toCsv,
    compareTuples
};
