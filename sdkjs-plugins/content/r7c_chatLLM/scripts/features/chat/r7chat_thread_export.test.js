const assert = require('assert');

function loadFresh(modulePath) {
    delete require.cache[require.resolve(modulePath)];
    return require(modulePath);
}

function setupExport() {
    global.R7Chat = {
        features: {},
        services: {}
    };
    loadFresh('./r7chat_thread_title');
    return loadFresh('./r7chat_thread_export');
}

function testSanitizeFilename() {
    const exporter = setupExport();
    assert.strictEqual(
        exporter.sanitizeFilename('Quarterly: Review / North?*'),
        'Quarterly Review North'
    );
}

function testBuildArchive() {
    const exporter = setupExport();
    const archive = exporter.buildDocxArchive({
        title: 'Quarterly planning thread',
        createdAt: '2026-03-24T09:00:00.000Z',
        updatedAt: '2026-03-24T10:00:00.000Z',
        messages: [
            {
                role: 'user',
                content: 'Summarize the quarter and highlight risks.',
                createdAt: '2026-03-24T09:00:00.000Z'
            },
            {
                role: 'assistant',
                content: 'Revenue improved, but retention dropped in enterprise accounts.',
                createdAt: '2026-03-24T09:01:00.000Z'
            }
        ]
    });

    assert.ok(archive instanceof Uint8Array);
    assert.strictEqual(String.fromCharCode(archive[0], archive[1]), 'PK');

    const bufferText = Buffer.from(archive).toString('utf8');
    assert.ok(bufferText.includes('[Content_Types].xml'));
    assert.ok(bufferText.includes('word/document.xml'));
    assert.ok(bufferText.includes('Quarterly planning thread'));
    assert.ok(bufferText.includes('Assistant'));
}

function run() {
    testSanitizeFilename();
    testBuildArchive();
    console.log('r7chat_thread_export.test.js: ok');
}

run();
