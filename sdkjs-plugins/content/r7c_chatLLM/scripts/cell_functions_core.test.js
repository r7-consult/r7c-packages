const assert = require('assert');
const core = require('./cell_functions_core');

function testNormalizeScalar() {
    assert.strictEqual(core.normalizeScalar('  hello   world  '), 'hello world');
    assert.strictEqual(core.normalizeScalar([['a', 'b'], ['c', 'd']]), 'a\tb\nc\td');
    assert.strictEqual(core.normalizeScalar(null), '');
}

function testNormalizeEnumAndLabels() {
    assert.strictEqual(core.normalizeEnum('  SHORT '), 'short');
    assert.deepStrictEqual(core.normalizeLabels('lead, spam, lead, client'), ['lead', 'spam', 'client']);
}

function testCacheKeyStability() {
    const key1 = core.buildCacheKey({
        fn: 'R7_TRANSLATE',
        args: ['hello', 'RU'],
        model: 'openrouter/auto',
        promptTemplateVersion: 'v1',
        locale: 'ru-RU',
        cacheEpoch: 1
    });
    const key2 = core.buildCacheKey({
        fn: 'R7_TRANSLATE',
        args: ['hello', 'RU'],
        model: 'openrouter/auto',
        promptTemplateVersion: 'v1',
        locale: 'ru-RU',
        cacheEpoch: 1
    });
    assert.strictEqual(key1, key2);
}

function testErrorMapping() {
    assert.strictEqual(core.mapRuntimeErrorToCellCode({ code: 'INVALID_ARGS' }), '#R7.INVALID_ARGS');
    assert.strictEqual(core.mapRuntimeErrorToCellCode({ status: 429 }), '#R7.RATE_LIMIT');
    assert.strictEqual(core.mapRuntimeErrorToCellCode({ status: 408 }), '#R7.TIMEOUT');
    assert.strictEqual(core.mapRuntimeErrorToCellCode({ status: 500 }), '#R7.FAILED');
}

function run() {
    testNormalizeScalar();
    testNormalizeEnumAndLabels();
    testCacheKeyStability();
    testErrorMapping();
    console.log('cell_functions_core.test.js: ok');
}

run();
