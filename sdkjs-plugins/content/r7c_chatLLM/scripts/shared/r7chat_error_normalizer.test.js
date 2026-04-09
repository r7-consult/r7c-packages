const assert = require('assert');
const errors = require('./r7chat_error_normalizer');

function testQuotaNormalization() {
    const normalized = errors.normalizeProviderError({
        status: 429,
        message: 'You exceeded your current quota.',
        apiErrorCode: 'insufficient_quota'
    });

    assert.strictEqual(normalized.quotaExceeded, true);
    assert.strictEqual(normalized.status, 429);
}

function testBuildUserMessage() {
    const message = errors.buildProviderErrorMessage({
        status: 500,
        message: 'Upstream failure'
    });

    assert.strictEqual(message, 'Request failed (HTTP 500): Upstream failure');
}

function run() {
    testQuotaNormalization();
    testBuildUserMessage();
    console.log('r7chat_error_normalizer.test.js: ok');
}

run();
