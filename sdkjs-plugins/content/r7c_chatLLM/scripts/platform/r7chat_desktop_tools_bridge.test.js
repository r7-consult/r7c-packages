const assert = require('assert');

function loadFresh(modulePath) {
    delete require.cache[require.resolve(modulePath)];
    return require(modulePath);
}

function resetGlobal(editor) {
    global.R7Chat = {};
    if (editor) {
        global.AscDesktopEditor = editor;
    } else {
        delete global.AscDesktopEditor;
    }
}

function testUnavailableWhenDesktopEditorIsMissing() {
    resetGlobal(null);
    const bridge = loadFresh('./r7chat_desktop_tools_bridge');
    const status = bridge.getStatus();

    assert.strictEqual(bridge.isAvailable(), false);
    assert.strictEqual(bridge.canExecute(), false);
    assert.strictEqual(status.desktopEditorAvailable, false);
    assert.strictEqual(status.catalogAvailable, false);
    assert.strictEqual(status.executionAvailable, false);
    assert.strictEqual(status.toolCount, 0);
}

function testReadOnlyCatalogStatus() {
    resetGlobal({
        getToolFunctions() {
            return JSON.stringify([
                {
                    name: 'document_summarize',
                    description: 'Summarize the active document',
                    parameters: {
                        type: 'object',
                        properties: {
                            length: { type: 'string' }
                        },
                        required: ['length']
                    }
                }
            ]);
        }
    });
    const bridge = loadFresh('./r7chat_desktop_tools_bridge');
    const status = bridge.refreshCatalog();
    const catalog = bridge.getCatalog();

    assert.strictEqual(bridge.isAvailable(), true);
    assert.strictEqual(bridge.canExecute(), false);
    assert.strictEqual(status.desktopEditorAvailable, true);
    assert.strictEqual(status.catalogAvailable, true);
    assert.strictEqual(status.executionAvailable, false);
    assert.strictEqual(status.toolCount, 1);
    assert.strictEqual(catalog[0].name, 'document_summarize');
    assert.strictEqual(catalog[0].inputSchema.required[0], 'length');
}

function testBrokenCatalogStatus() {
    resetGlobal({
        getToolFunctions() {
            return '{broken json';
        }
    });
    const bridge = loadFresh('./r7chat_desktop_tools_bridge');
    const status = bridge.refreshCatalog();

    assert.strictEqual(status.catalogAvailable, false);
    assert.ok(/Failed to parse/i.test(status.catalogParseError));
    assert.strictEqual(status.toolCount, 0);
}

async function testCallToolValidatesAndExecutes() {
    let capturedName = '';
    let capturedArgs = '';
    resetGlobal({
        getToolFunctions() {
            return JSON.stringify([
                {
                    name: 'document_insert_summary',
                    description: 'Insert a summary block',
                    parameters: {
                        type: 'object',
                        properties: {
                            topic: { type: 'string' },
                            bulletCount: { type: 'number' }
                        },
                        required: ['topic']
                    }
                }
            ]);
        },
        callToolFunction(name, argsJson) {
            capturedName = name;
            capturedArgs = argsJson;
            return JSON.stringify({
                ok: true,
                inserted: true
            });
        }
    });
    const bridge = loadFresh('./r7chat_desktop_tools_bridge');
    bridge.refreshCatalog();

    const invalid = await bridge.callTool('document_insert_summary', { bulletCount: 3 });
    assert.strictEqual(invalid.ok, false);
    assert.ok(/Missing required input: topic/i.test(invalid.error));

    const valid = await bridge.callTool('document_insert_summary', {
        topic: 'Quarterly review',
        bulletCount: 3
    });

    assert.strictEqual(valid.ok, true);
    assert.strictEqual(capturedName, 'document_insert_summary');
    assert.deepStrictEqual(JSON.parse(capturedArgs), {
        topic: 'Quarterly review',
        bulletCount: 3
    });
    assert.deepStrictEqual(valid.data, {
        ok: true,
        inserted: true
    });
}

async function testCallToolFailsForUnknownTool() {
    resetGlobal({
        getToolFunctions() {
            return JSON.stringify([]);
        },
        callToolFunction() {
            throw new Error('should not execute');
        }
    });
    const bridge = loadFresh('./r7chat_desktop_tools_bridge');
    bridge.refreshCatalog();

    const result = await bridge.callTool('missing_tool', {});
    assert.strictEqual(result.ok, false);
    assert.ok(/runtime catalog/i.test(result.error));
}

async function run() {
    testUnavailableWhenDesktopEditorIsMissing();
    testReadOnlyCatalogStatus();
    testBrokenCatalogStatus();
    await testCallToolValidatesAndExecutes();
    await testCallToolFailsForUnknownTool();
    console.log('r7chat_desktop_tools_bridge tests passed');
}

run().catch((error) => {
    console.error(error);
    process.exitCode = 1;
});
