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
        },
        dump() {
            return Object.assign({}, data);
        }
    };
}

function loadFresh(modulePath) {
    delete require.cache[require.resolve(modulePath)];
    return require(modulePath);
}

function setupRuntime(storage) {
    global.R7Chat = {
        runtime: {
            state: {
                chat: {
                    conversationHistory: []
                }
            },
            ui: {}
        },
        platform: {
            storage: {
                getLocalStorage() {
                    return storage;
                }
            }
        },
        features: {},
        services: {}
    };
    loadFresh('./r7chat_attachment_store');
    loadFresh('./r7chat_thread_title');
    return loadFresh('./r7chat_thread_store');
}

function testDefaultInitialization() {
    const storage = createStorage();
    const store = setupRuntime(storage);
    const state = store.initialize({ storage });

    assert.strictEqual(state.threads.length, 1);
    assert.ok(state.activeThreadId);
    assert.ok(storage.getItem(store.STORAGE_KEY));
}

function testAppendAndRenameFlow() {
    const storage = createStorage();
    const store = setupRuntime(storage);
    store.initialize({ storage });

    store.appendMessage('user', 'summarize revenue by region for q4 and suggest the main anomalies');
    const activeThread = store.getActiveThread();
    assert.strictEqual(activeThread.messages.length, 1);
    assert.strictEqual(activeThread.title, 'Summarize revenue by region for q4 and');

    store.replaceDraft(activeThread.id, 'Draft text');
    assert.strictEqual(store.getDraft(activeThread.id), 'Draft text');

    store.renameThread(activeThread.id, 'Revenue review');
    assert.strictEqual(store.getActiveThread().title, 'Revenue review');
    assert.strictEqual(global.R7Chat.runtime.state.chat.conversationHistory.length, 1);
}

function testDeleteFallsBackToFreshThread() {
    const storage = createStorage();
    const store = setupRuntime(storage);
    store.initialize({ storage });
    const firstThread = store.getActiveThread();

    store.deleteThread(firstThread.id);

    assert.strictEqual(store.getThreads().length, 1);
    assert.notStrictEqual(store.getActiveThread().id, firstThread.id);
    assert.strictEqual(store.getActiveThread().messages.length, 0);
}

function testConversationHistoryPreservesImagePayload() {
    const storage = createStorage();
    const store = setupRuntime(storage);
    const attachments = global.R7Chat.features.chatAttachments;
    store.initialize({ storage });

    store.appendMessage('user', 'Check this image', {
        meta: {
            attachments: [{
                attachmentType: 'image',
                name: 'Screenshot.png',
                mimeType: 'image/png',
                size: 5120,
                width: 800,
                height: 600,
                dataUrl: 'data:image/png;base64,BBBB'
            }]
        }
    });

    const history = store.syncConversationHistory();

    assert.strictEqual(history.length, 1);
    assert.ok(Array.isArray(history[0].content));
    assert.deepStrictEqual(history[0].content, attachments.buildProviderMessageContent('Check this image', [{
        attachmentType: 'image',
        name: 'Screenshot.png',
        mimeType: 'image/png',
        size: 5120,
        width: 800,
        height: 600,
        dataUrl: 'data:image/png;base64,BBBB'
    }]));
}

function run() {
    testDefaultInitialization();
    testAppendAndRenameFlow();
    testDeleteFallsBackToFreshThread();
    testConversationHistoryPreservesImagePayload();
    console.log('r7chat_thread_store.test.js: ok');
}

run();
