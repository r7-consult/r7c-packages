const assert = require('assert');
const httpClient = require('./r7chat_http_client');

async function testAscPromiseRejectsCleanly() {
    const originalAsc = global.AscSimpleRequest;
    try {
        global.AscSimpleRequest = {
            request() {
                return Promise.reject(new Error('asc rejected'));
            }
        };

        await assert.rejects(
            () => httpClient.request({ url: 'https://example.com', preferAscOnly: true }),
            /asc rejected/
        );
    } finally {
        global.AscSimpleRequest = originalAsc;
    }
}

async function testAscCreateRequestConfigStyleWorks() {
    const originalAsc = global.AscSimpleRequest;
    try {
        global.AscSimpleRequest = {
            createRequest(options) {
                setTimeout(() => {
                    options.complete({
                        status: 200,
                        headers: { 'content-type': 'application/json' },
                        responseText: '{"ok":true}'
                    });
                }, 0);
                return undefined;
            }
        };

        const response = await httpClient.request({
            url: 'https://api.exa.ai/search',
            method: 'POST',
            body: '{}',
            preferAscOnly: true
        });

        assert.strictEqual(response.ok, true);
        assert.deepStrictEqual(response.json(), { ok: true });
    } finally {
        global.AscSimpleRequest = originalAsc;
    }
}

async function testRequestStreamViaFetchReturnsReader() {
    const originalFetch = global.fetch;
    const originalTextDecoderStream = global.TextDecoderStream;
    try {
        const reader = {
            async read() {
                return { done: true, value: '' };
            }
        };

        global.TextDecoderStream = function TextDecoderStream() {};
        global.fetch = async function () {
            return {
                ok: true,
                status: 200,
                statusText: 'OK',
                headers: {
                    forEach(callback) {
                        callback('text/event-stream', 'content-type');
                    }
                },
                body: {
                    pipeThrough() {
                        return {
                            getReader() {
                                return reader;
                            }
                        };
                    }
                }
            };
        };

        const response = await httpClient.requestStream({
            url: 'https://example.com/stream',
            method: 'POST',
            body: '{}'
        });

        assert.strictEqual(response.ok, true);
        assert.strictEqual(typeof response.reader.read, 'function');
    } finally {
        global.fetch = originalFetch;
        global.TextDecoderStream = originalTextDecoderStream;
    }
}

async function testPreferFetchBypassesAsc() {
    const originalAsc = global.AscSimpleRequest;
    const originalFetch = global.fetch;
    try {
        global.AscSimpleRequest = {
            request() {
                throw new Error('asc should not be used');
            }
        };
        global.fetch = async function () {
            return {
                status: 200,
                statusText: 'OK',
                headers: {
                    forEach(callback) {
                        callback('application/json', 'content-type');
                    }
                },
                async text() {
                    return '{"via":"fetch"}';
                }
            };
        };

        const response = await httpClient.request({
            url: 'https://example.com',
            preferFetch: true
        });

        assert.strictEqual(response.ok, true);
        assert.deepStrictEqual(response.json(), { via: 'fetch' });
    } finally {
        global.AscSimpleRequest = originalAsc;
        global.fetch = originalFetch;
    }
}

async function testJsonParsingFailureFallsBackToFetch() {
    const originalAsc = global.AscSimpleRequest;
    const originalFetch = global.fetch;
    try {
        global.AscSimpleRequest = {
            async request() {
                return {
                    status: 502,
                    statusText: 'Bad Gateway',
                    headers: { 'content-type': 'text/plain' },
                    responseText: 'JSON parsing failed'
                };
            }
        };
        global.fetch = async function () {
            return {
                status: 200,
                statusText: 'OK',
                headers: {
                    forEach(callback) {
                        callback('application/json', 'content-type');
                    }
                },
                async text() {
                    return '{"ok":true,"fallback":"fetch"}';
                }
            };
        };

        const response = await httpClient.request({
            url: 'https://example.com/image',
            retryViaFetchOnJsonParseFailure: true
        });

        assert.strictEqual(response.ok, true);
        assert.deepStrictEqual(response.json(), { ok: true, fallback: 'fetch' });
    } finally {
        global.AscSimpleRequest = originalAsc;
        global.fetch = originalFetch;
    }
}

async function run() {
    await testAscPromiseRejectsCleanly();
    await testAscCreateRequestConfigStyleWorks();
    await testRequestStreamViaFetchReturnsReader();
    await testPreferFetchBypassesAsc();
    await testJsonParsingFailureFallsBackToFetch();
    console.log('r7chat_http_client.test.js: ok');
}

run().catch((error) => {
    console.error(error);
    process.exit(1);
});
