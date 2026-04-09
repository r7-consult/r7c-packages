const assert = require('assert');

function createStorage(seed) {
    const data = Object.assign({}, seed || {});
    return {
        getItem(key) {
            return Object.prototype.hasOwnProperty.call(data, key) ? data[key] : null;
        },
        setItem(key, value) {
            data[key] = String(value);
        },
        removeItem(key) {
            delete data[key];
        }
    };
}

function loadFresh(modulePath) {
    delete require.cache[require.resolve(modulePath)];
    return require(modulePath);
}

function setupRuntime(storage) {
    global.R7Chat = {
        platform: {
            storage: {
                getLocalStorage() {
                    return storage;
                }
            }
        },
        features: {}
    };
    return loadFresh('./r7chat_attachment_store');
}

function testFromFileNormalizesSupportedDocument() {
    const storage = createStorage();
    const attachments = setupRuntime(storage);
    const file = {
        name: 'Quarterly report.pdf',
        type: 'application/pdf',
        size: 16384,
        lastModified: Date.parse('2026-03-24T10:00:00.000Z')
    };

    const normalized = attachments.fromFile(file);

    assert.strictEqual(normalized.name, 'Quarterly report.pdf');
    assert.strictEqual(normalized.extension, 'pdf');
    assert.strictEqual(normalized.mimeType, 'application/pdf');
    assert.strictEqual(normalized.size, 16384);
    assert.strictEqual(normalized.source, 'local');
    assert.ok(normalized.id);
    assert.ok(normalized.fingerprint.includes('quarterly report.pdf'));
}

function testDedupeAttachmentsUsesFingerprint() {
    const storage = createStorage();
    const attachments = setupRuntime(storage);
    const base = attachments.normalizeAttachment({
        name: 'Budget.xlsx',
        size: 2048,
        modifiedAt: '2026-03-24T11:00:00.000Z'
    });
    const duplicate = attachments.normalizeAttachment({
        id: 'custom-id',
        name: 'Budget.xlsx',
        size: 2048,
        modifiedAt: '2026-03-24T11:00:00.000Z'
    });
    const second = attachments.normalizeAttachment({
        name: 'Deck.pptx',
        size: 4096,
        modifiedAt: '2026-03-24T12:00:00.000Z'
    });

    const deduped = attachments.dedupeAttachments([base], [duplicate, second]);

    assert.strictEqual(deduped.length, 2);
    assert.strictEqual(deduped[0].name, 'Budget.xlsx');
    assert.strictEqual(deduped[1].name, 'Deck.pptx');
}

function testRememberRecentKeepsLatestFilesFirstAndLimited() {
    const storage = createStorage();
    const attachments = setupRuntime(storage);
    const items = [];

    for (let index = 0; index < 10; index += 1) {
        items.push(attachments.normalizeAttachment({
            name: 'File-' + index + '.pdf',
            size: 1024 + index,
            modifiedAt: '2026-03-24T10:00:00.000Z',
            addedAt: new Date(Date.UTC(2026, 2, 24, 10, index, 0)).toISOString()
        }));
    }

    const recent = attachments.rememberRecent(items);

    assert.strictEqual(recent.length, attachments.STORAGE_LIMIT);
    assert.strictEqual(recent[0].name, 'File-9.pdf');
    assert.strictEqual(recent[recent.length - 1].name, 'File-2.pdf');
    assert.ok(storage.getItem(attachments.STORAGE_KEY));
}

function testImageAttachmentBuildsProviderContent() {
    const storage = createStorage();
    const attachments = setupRuntime(storage);
    const image = attachments.normalizeAttachment({
        attachmentType: 'image',
        name: 'Screenshot.png',
        mimeType: 'image/png',
        size: 24576,
        width: 1280,
        height: 720,
        source: 'clipboard',
        dataUrl: 'data:image/png;base64,AAAA'
    });

    const providerContent = attachments.buildProviderMessageContent('Analyze this UI', [image]);

    assert.strictEqual(image.attachmentType, 'image');
    assert.strictEqual(attachments.isImageAttachment(image), true);
    assert.strictEqual(attachments.getTypeLabel(image), 'IMG');
    assert.ok(Array.isArray(providerContent));
    assert.deepStrictEqual(providerContent[0], {
        type: 'text',
        text: 'Analyze this UI'
    });
    assert.deepStrictEqual(providerContent[1], {
        type: 'image_url',
        image_url: {
            url: 'data:image/png;base64,AAAA'
        }
    });
}

function run() {
    testFromFileNormalizesSupportedDocument();
    testDedupeAttachmentsUsesFingerprint();
    testRememberRecentKeepsLatestFilesFirstAndLimited();
    testImageAttachmentBuildsProviderContent();
    console.log('r7chat_attachment_store.test.js: ok');
}

run();
