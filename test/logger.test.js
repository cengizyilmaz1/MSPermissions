const assert = require('assert');
const test = require('node:test');

// Force deterministic, color-free output for assertions.
process.env.NO_COLOR = '1';

const { createLogger } = require('../scripts/node/lib/logger');

function capture(run) {
    const calls = { log: [], warn: [], error: [] };
    const original = { log: console.log, warn: console.warn, error: console.error };
    console.log = (message) => calls.log.push(message);
    console.warn = (message) => calls.warn.push(message);
    console.error = (message) => calls.error.push(message);
    try {
        run();
    } finally {
        console.log = original.log;
        console.warn = original.warn;
        console.error = original.error;
    }
    return calls;
}

test('logger suppresses messages below the active level', () => {
    const calls = capture(() => {
        const log = createLogger({ level: 'warn', scope: 'test' });
        log.debug('d');
        log.info('i');
        log.warn('w');
        log.error('e');
    });

    assert.equal(calls.log.length, 0);
    assert.equal(calls.warn.length, 1);
    assert.equal(calls.error.length, 1);
    assert.match(calls.warn[0], /WARN/);
    assert.match(calls.warn[0], /\[test\]/);
    assert.match(calls.warn[0], /w$/);
});

test('logger emits info and success at info level but hides debug', () => {
    const calls = capture(() => {
        const log = createLogger({ level: 'info', scope: 'x' });
        log.info('hello');
        log.success('done');
        log.debug('nope');
    });

    assert.equal(calls.log.length, 2);
    assert.match(calls.log[0], /INFO/);
    assert.match(calls.log[0], /hello$/);
    assert.match(calls.log[1], /done$/);
});

test('logger prefixes every line with an ISO timestamp', () => {
    const calls = capture(() => {
        createLogger({ level: 'info' }).info('ping');
    });

    assert.match(calls.log[0], /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z /);
});
