const assert = require('assert');
const webText = require('./r7chat_web_text');

function testValidateWebUrl() {
    assert.strictEqual(webText.validateWebUrl('https://example.com/docs').isValid, true);
    assert.strictEqual(webText.validateWebUrl('http://127.0.0.1/test').isValid, false);
    assert.strictEqual(webText.validateWebUrl('https://localhost/app').reason, 'localhost_blocked');
    assert.strictEqual(webText.validateWebUrl('ftp://example.com/file').reason, 'unsupported_scheme');
}

function testValidateWebUrls() {
    const result = webText.validateWebUrls([
        'https://example.com/one',
        'https://example.com/one',
        'http://192.168.1.50/private',
        'https://example.com/two'
    ], { maxUrls: 2 });

    assert.deepStrictEqual(result.accepted, ['https://example.com/one', 'https://example.com/two']);
    assert.strictEqual(result.errors.length, 1);
    assert.strictEqual(result.errors[0].reason, 'private_host_blocked');
}

function testExtractHtmlContent() {
    const extracted = webText.extractPageContentFromHtml(
        '<html><head><title>Example &amp; Test</title><style>.x{}</style></head><body><h1>Hello</h1><p>World</p><script>alert(1)</script></body></html>',
        'https://example.com',
        { maxChars: 50, excerptChars: 20 }
    );

    assert.strictEqual(extracted.title, 'Example & Test');
    assert.strictEqual(extracted.text.includes('Hello'), true);
    assert.strictEqual(extracted.text.includes('World'), true);
    assert.strictEqual(extracted.text.includes('alert'), false);
    assert.strictEqual(extracted.excerpt.length <= 20, true);
}

function testExtractFetchedContentTruncatesText() {
    const text = 'A'.repeat(120);
    const extracted = webText.extractFetchedContent(text, 'text/plain', 'https://example.com/plain', {
        maxChars: 40,
        excerptChars: 10
    });

    assert.strictEqual(extracted.url, 'https://example.com/plain');
    assert.strictEqual(extracted.truncated, true);
    assert.strictEqual(extracted.text.length <= 40, true);
    assert.strictEqual(extracted.excerpt.length <= 10, true);
}

function run() {
    testValidateWebUrl();
    testValidateWebUrls();
    testExtractHtmlContent();
    testExtractFetchedContentTruncatesText();
    console.log('r7chat_web_text.test.js: ok');
}

run();
