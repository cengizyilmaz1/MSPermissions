const assert = require('assert');
const test = require('node:test');

const {
    mineLines,
    sortTuples,
    toCsv,
    compareTuples
} = require('../scripts/node/lib/openapi-permissions');

const SAMPLE_LINES = [
    'paths:',
    '  /users:',
    '    get:',
    '      tags:',
    '        - users.user',
    '    post:',
    '      tags:',
    '        - users.user',
    '  /groups/{group-id}:',
    '    delete:',
    '      tags:',
    '        - groups.group',
    '      requestBody:',
    '        - microsoft.graph.group',
    "    '/quoted':",
    '      get:',
    '        - quoted.entity'
];

test('mineLines extracts permission tuples for known path + method', () => {
    const tuples = sortTuples(mineLines(SAMPLE_LINES, 'v1.0'));

    assert.deepEqual(tuples, [
        ['groups.group', 'DELETE', '/groups/{group-id}', 'v1.0'],
        ['quoted.entity', 'GET', '/quoted', 'v1.0'],
        ['users.user', 'GET', '/users', 'v1.0'],
        ['users.user', 'POST', '/users', 'v1.0']
    ]);
});

test('mineLines excludes microsoft.graph odata references', () => {
    const tuples = sortTuples(mineLines(SAMPLE_LINES, 'v1.0'));
    assert.ok(!tuples.some((tuple) => tuple[0].includes('microsoft.graph')));
});

test('mineLines ignores tags before any method is seen', () => {
    const lines = ['  /orphan:', '      tags:', '        - orphan.entity'];
    const tuples = sortTuples(mineLines(lines, 'beta'));
    assert.deepEqual(tuples, []);
});

test('mineLines resets the method when a new path begins', () => {
    const lines = [
        '  /a:',
        '    get:',
        '  /b:',
        '        - bee.entity' // no method active for /b yet
    ];
    const tuples = sortTuples(mineLines(lines, 'v1.0'));
    assert.deepEqual(tuples, []);
});

test('mineLines de-duplicates identical tuples', () => {
    const lines = ['  /x:', '    get:', '        - item.entity', '        - item.entity'];
    const results = mineLines(lines, 'v1.0');
    assert.equal(results.size, 1);
});

test('mineLines drops tags that are 60 characters or longer', () => {
    const longTag = `items.${'b'.repeat(60)}`;
    const lines = ['  /x:', '    get:', `        - ${longTag}`];
    assert.equal(mineLines(lines, 'v1.0').size, 0);
});

test('compareTuples orders case-insensitively across all columns', () => {
    assert.equal(compareTuples(['a', 'GET', '/x', 'v1.0'], ['a', 'POST', '/x', 'v1.0']), -1);
    assert.equal(compareTuples(['B', 'GET', '/x', 'v1.0'], ['a', 'GET', '/x', 'v1.0']), 1);
    assert.equal(compareTuples(['a', 'GET', '/x', 'v1.0'], ['a', 'GET', '/x', 'v1.0']), 0);
});

test('toCsv emits the exact quoted contract with a trailing newline', () => {
    const csv = toCsv([['users.user', 'GET', '/users', 'v1.0']]);
    assert.equal(
        csv,
        '"Permission","Method","Endpoint","ApiVersion"\n"users.user","GET","/users","v1.0"\n'
    );
});
