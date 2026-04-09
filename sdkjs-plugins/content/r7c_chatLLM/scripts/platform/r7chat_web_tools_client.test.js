const assert = require('assert');
const webText = require('../shared/r7chat_web_text');
const webTools = require('./r7chat_web_tools_client');

function createHttpStub(handler) {
    return {
        async request(options) {
            return handler(options);
        }
    };
}

function createJsonResponse(body) {
    const bodyText = JSON.stringify(body);
    return {
        ok: true,
        status: 200,
        statusText: 'OK',
        headers: { 'content-type': 'application/json' },
        contentType: 'application/json',
        bodyText,
        json() {
            return JSON.parse(bodyText);
        }
    };
}

function createTextResponse(body, contentType) {
    return {
        ok: true,
        status: 200,
        statusText: 'OK',
        headers: { 'content-type': contentType || 'text/html' },
        contentType: contentType || 'text/html',
        bodyText: body,
        json() {
            return JSON.parse(body);
        }
    };
}

function createRawResponse(status, body, contentType) {
    const bodyText = typeof body === 'string' ? body : JSON.stringify(body || {});
    return {
        ok: Number(status || 0) >= 200 && Number(status || 0) < 300,
        status: Number(status || 0),
        statusText: '',
        headers: { 'content-type': contentType || 'application/json' },
        contentType: contentType || 'application/json',
        bodyText,
        json() {
            return JSON.parse(bodyText);
        }
    };
}

async function testExaSearchNormalization() {
    let captured = null;
    const client = webTools.createClient({
        http: createHttpStub((options) => {
            captured = options;
            return createJsonResponse({
                results: [{
                    title: 'Result title',
                    url: 'https://example.com/article',
                    text: 'Result snippet'
                }]
            });
        }),
        webText,
        loadSettings() {
            return {
                webTools: {
                    provider: 'exa',
                    providers: {
                        exa: { apiKey: 'exa-key' },
                        brave: { apiKey: '' }
                    }
                }
            };
        }
    });

    const result = await client.executeWebSearch({ query: 'latest llm news' });
    const payload = JSON.parse(captured.body);

    assert.strictEqual(captured.url, 'https://api.exa.ai/search');
    assert.strictEqual(captured.headers['x-api-key'], 'exa-key');
    assert.strictEqual(payload.query, 'latest llm news');
    assert.strictEqual(payload.type, 'auto');
    assert.strictEqual(payload.numResults, 5);
    assert.strictEqual(payload.contents.text.maxCharacters, 10000);
    assert.strictEqual(result.provider, 'exa');
    assert.strictEqual(result.results[0].url, 'https://example.com/article');
    assert.strictEqual(Array.isArray(result.data), true);
    assert.strictEqual(result.data[0].url, 'https://example.com/article');
}

async function testBraveSearchPayload() {
    let captured = null;
    const client = webTools.createClient({
        http: createHttpStub((options) => {
            captured = options;
            return createJsonResponse({
                web: {
                    results: [{
                        title: 'Brave title',
                        url: 'https://example.com/page',
                        description: 'Brave snippet'
                    }]
                }
            });
        }),
        webText,
        loadSettings() {
            return {
                webTools: {
                    provider: 'brave',
                    providers: {
                        exa: { apiKey: '' },
                        brave: { apiKey: 'brave-key' }
                    }
                }
            };
        }
    });

    const result = await client.executeWebSearch({ query: 'brave query' });
    assert.ok(captured.url.indexOf('https://api.search.brave.com/res/v1/web/search?') === 0);
    assert.strictEqual(captured.headers['X-Subscription-Token'], 'brave-key');
    assert.strictEqual(result.provider, 'brave');
    assert.strictEqual(result.results[0].snippet, 'Brave snippet');
}

async function testBraveSearchAcceptsValidBodyWhenStatusIsZero() {
    const client = webTools.createClient({
        http: createHttpStub(() => createRawResponse(0, {
            type: 'search',
            query: { original: 'ivan iii' },
            web: {
                results: [{
                    title: 'Ivan III - Britannica',
                    url: 'https://example.com/ivan-iii',
                    description: 'Grand Prince of Moscow'
                }]
            }
        })),
        webText,
        loadSettings() {
            return {
                webTools: {
                    provider: 'brave',
                    providers: {
                        exa: { apiKey: '' },
                        brave: { apiKey: 'brave-key' }
                    }
                }
            };
        }
    });

    const result = await client.executeWebSearch({ query: 'ivan iii' });
    assert.strictEqual(result.provider, 'brave');
    assert.strictEqual(result.errors.length, 0);
    assert.strictEqual(result.results.length, 1);
    assert.strictEqual(result.results[0].url, 'https://example.com/ivan-iii');
}

async function testBraveCrawlUrlValidation() {
    const client = webTools.createClient({
        http: createHttpStub(() => {
            throw new Error('should not fetch invalid urls');
        }),
        webText,
        loadSettings() {
            return {
                webTools: {
                    provider: 'brave',
                    providers: {
                        exa: { apiKey: '' },
                        brave: { apiKey: 'brave-key' }
                    }
                }
            };
        }
    });

    const result = await client.executeWebCrawling({ urls: ['http://127.0.0.1/test'] });
    assert.strictEqual(result.provider, 'brave');
    assert.strictEqual(result.results.length, 0);
    assert.strictEqual(result.errors[0].code, 'private_host_blocked');
}

async function testBraveCrawlExtractionAndTruncation() {
    const client = webTools.createClient({
        http: createHttpStub(() => createTextResponse(
            '<html><head><title>Doc</title></head><body><p>' + 'A'.repeat(7000) + '</p></body></html>',
            'text/html'
        )),
        webText,
        loadSettings() {
            return {
                webTools: {
                    provider: 'brave',
                    providers: {
                        exa: { apiKey: '' },
                        brave: { apiKey: 'brave-key' }
                    }
                }
            };
        }
    });

    const result = await client.executeWebCrawling({ urls: ['https://example.com/doc'] });
    assert.strictEqual(result.results.length, 1);
    assert.strictEqual(result.results[0].title, 'Doc');
    assert.strictEqual(result.results[0].truncated, true);
    assert.strictEqual(result.results[0].text.length <= 6000, true);
}

async function testExaCrawlPayload() {
    let captured = null;
    const client = webTools.createClient({
        http: createHttpStub((options) => {
            captured = options;
            return createJsonResponse({
                results: [{
                    url: 'https://example.com/doc',
                    title: 'Exa doc',
                    text: 'Extracted text'
                }],
                statuses: [{
                    id: 'https://example.com/doc',
                    status: 'success'
                }]
            });
        }),
        webText,
        loadSettings() {
            return {
                webTools: {
                    provider: 'exa',
                    providers: {
                        exa: { apiKey: 'exa-key' },
                        brave: { apiKey: '' }
                    }
                }
            };
        }
    });

    const result = await client.executeWebCrawling({ urls: ['https://example.com/doc'] });
    const payload = JSON.parse(captured.body);

    assert.strictEqual(captured.url, 'https://api.exa.ai/contents');
    assert.strictEqual(captured.headers['x-api-key'], 'exa-key');
    assert.deepStrictEqual(payload.urls, ['https://example.com/doc']);
    assert.strictEqual(payload.text.maxCharacters, 6000);
    assert.strictEqual(result.provider, 'exa');
    assert.strictEqual(result.results[0].url, 'https://example.com/doc');
}

async function testSearchShowsCorsDiagnosticInOnlyOfficeOrigin() {
    const originalLocation = global.location;
    global.location = { protocol: 'onlyoffice:' };
    try {
        const client = webTools.createClient({
            http: {
                getAvailableTransport() {
                    return 'fetch';
                },
                async request() {
                    throw new Error('Failed to fetch');
                }
            },
            webText,
            loadSettings() {
                return {
                    webTools: {
                        provider: 'brave',
                        providers: {
                            exa: { apiKey: '' },
                            brave: { apiKey: 'brave-key' }
                        }
                    }
                };
            }
        });

        const result = await client.executeWebSearch({ query: 'latest news' });
        assert.strictEqual(result.provider, 'brave');
        assert.strictEqual(result.errors[0].code, 'network_error');
        assert.ok(/onlyoffice visual plugins/i.test(result.errors[0].message));
    } finally {
        global.location = originalLocation;
    }
}

async function run() {
    await testExaSearchNormalization();
    await testBraveSearchPayload();
    await testBraveSearchAcceptsValidBodyWhenStatusIsZero();
    await testBraveCrawlUrlValidation();
    await testBraveCrawlExtractionAndTruncation();
    await testExaCrawlPayload();
    await testSearchShowsCorsDiagnosticInOnlyOfficeOrigin();
    console.log('r7chat_web_tools_client.test.js: ok');
}

run().catch((error) => {
    console.error(error);
    process.exit(1);
});
