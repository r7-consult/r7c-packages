const assert = require('assert');

function loadFresh(modulePath) {
    delete require.cache[require.resolve(modulePath)];
    return require(modulePath);
}

async function testHostDiagnostics() {
    global.R7Chat = {};
    delete global.Asc;

    const hostBridge = loadFresh('./r7chat_host_bridge');

    assert.strictEqual(hostBridge.isPluginHostReady(), false);
    assert.strictEqual(hostBridge.getEditorTypeSafe(), '');
    assert.strictEqual(hostBridge.getHostDiagnostics(), 'Asc host is not available.');
    assert.strictEqual(await hostBridge.callEditorMethod('noop', []), null);
}

async function testCommandAndMethodWrappers() {
    global.R7Chat = {};
    global.Asc = {
        scope: {},
        plugin: {
            info: { editorType: 'cell' },
            callCommand(command, close, recalculate, callback) {
                assert.strictEqual(close, false);
                assert.strictEqual(recalculate, false);
                callback(command());
            },
            executeMethod(name, args, callback) {
                callback({ name, args });
            }
        }
    };

    const hostBridge = loadFresh('./r7chat_host_bridge');

    assert.strictEqual(hostBridge.isPluginHostReady(), true);
    assert.strictEqual(hostBridge.getEditorTypeSafe(), 'cell');
    assert.strictEqual(hostBridge.getHostDiagnostics(), '');

    const commandResult = await hostBridge.callEditorCommand(function () {
        return {
            sheetName: global.Asc.scope.sheetName
        };
    }, { sheetName: 'Sheet 1' }, { recalculate: false });

    assert.deepStrictEqual(commandResult, { sheetName: 'Sheet 1' });
    assert.strictEqual(global.Asc.scope.sheetName, 'Sheet 1');

    const methodResult = await hostBridge.callEditorMethod('GetVersion', ['v1']);
    assert.deepStrictEqual(methodResult, { name: 'GetVersion', args: ['v1'] });
}

async function testCommandTimeoutResolvesNull() {
    global.R7Chat = {};
    global.Asc = {
        scope: {},
        plugin: {
            info: { editorType: 'cell' },
            callCommand() {
                // Intentionally never calls the callback.
            },
            executeMethod(name, args, callback) {
                callback({ name, args });
            }
        }
    };

    const hostBridge = loadFresh('./r7chat_host_bridge');
    const startedAt = Date.now();
    const result = await hostBridge.callEditorCommand(function () {
        return 'never';
    }, null, { timeoutMs: 25 });
    const elapsedMs = Date.now() - startedAt;

    assert.strictEqual(result, null);
    assert.ok(elapsedMs >= 20);
}

async function testTranslateRespectsSavedLanguageOverride() {
    global.R7Chat = {
        features: {
            settings: {
                loadSettings() {
                    return { language: 'ru' };
                }
            }
        }
    };
    global.Asc = {
        scope: {},
        plugin: {
            info: { editorType: 'word', lang: 'en-US' },
            tr(text) {
                return 'host:' + text;
            },
            callCommand(command, close, recalculate, callback) {
                callback(command());
            },
            executeMethod(name, args, callback) {
                callback({ name, args });
            }
        }
    };

    const hostBridge = loadFresh('./r7chat_host_bridge');
    assert.strictEqual(hostBridge.getLanguagePreference(), 'ru');
    assert.strictEqual(hostBridge.getUiLanguageCode(), 'ru');
    assert.strictEqual(hostBridge.translate('Settings'), 'Настройки');
    assert.strictEqual(hostBridge.translate('Use editor language'), 'Использовать язык редактора');
}

async function run() {
    await testHostDiagnostics();
    await testCommandAndMethodWrappers();
    await testCommandTimeoutResolvesNull();
    await testTranslateRespectsSavedLanguageOverride();
    console.log('r7chat_host_bridge.test.js: ok');
}

run().catch((error) => {
    console.error(error);
    process.exit(1);
});
