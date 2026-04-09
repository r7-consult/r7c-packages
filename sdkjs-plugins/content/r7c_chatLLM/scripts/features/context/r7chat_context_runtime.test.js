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

function createContextRuntime(hostBridge) {
    global.localStorage = createStorage();
    global.R7Chat = {
        runtime: {
            constants: {
                contextStateKey: 'r7chat.test.context',
                contextCacheTtlMs: 1800,
                contextToolLimits: {
                    maxCharsWord: 120,
                    maxCharsCell: 80,
                    maxRowsPerChunk: 2,
                    maxColsPerChunk: 2,
                    maxParagraphsPerChunk: 2,
                    maxInternalChunks: 2
                },
                contextLimits: {
                    maxDocumentChars: 12000,
                    maxSheetChars: 24,
                    maxPreviewChars: 1800,
                    maxRowsPerSheet: 2,
                    maxColsPerSheet: 2,
                    maxPreviewRows: 16,
                    maxPreviewCols: 8,
                    maxAttachedSheets: 6
                }
            },
            state: {
                context: {
                    autoModeWord: 'full_document',
                    autoModeCell: 'active_sheet',
                    attachedSheets: [],
                    smartLimit: true,
                    workbookSheetCache: []
                }
            }
        },
        platform: {
            hostBridge: hostBridge || null
        },
        context: {}
    };

    return loadFresh('./r7chat_context_runtime');
}

function testSanitizeAndHelpers() {
    const runtime = createContextRuntime(null);

    const sanitized = runtime.sanitizeContextState({
        autoModeWord: 'bad-value',
        autoModeCell: 'active_sheet',
        attachedSheets: ['Sheet 1', 'sheet 1', 'Sheet 2', '', null],
        smartLimit: false,
        workbookSheetCache: [{ name: 'Cached' }]
    });

    assert.deepStrictEqual(sanitized, {
        autoModeWord: 'full_document',
        autoModeCell: 'active_sheet',
        attachedSheets: ['Sheet 1', 'Sheet 2'],
        smartLimit: false,
        workbookSheetCache: [{ name: 'Cached' }]
        ,
        lastToolSummary: null
    });

    const table = runtime.formatTablePayload([
        ['A1', 'B1', 'C1'],
        ['A2', 'B2', 'C2'],
        ['A3', 'B3', 'C3']
    ], {
        maxRows: 2,
        maxCols: 2,
        maxChars: 11
    });

    assert.strictEqual(table.totalRows, 3);
    assert.strictEqual(table.totalCols, 3);
    assert.strictEqual(table.sampledRows, 2);
    assert.strictEqual(table.sampledCols, 2);
    assert.strictEqual(table.truncated, true);
    assert.strictEqual(table.text, 'A1\tB1\nA2\tB2');

    assert.deepStrictEqual(runtime.parseRangeDimensions('Sheet1!$B$2:$D$5'), {
        rows: 4,
        cols: 3
    });
}

async function testBuildEnvelopeAndRequestPreparation() {
    const runtime = createContextRuntime({
        getEditorTypeSafe() {
            return 'word';
        },
        async callEditorCommand(command, scopeData, options) {
            assert.strictEqual(options.timeoutMs, 2500);
            return {
                source: 'full_document',
                totalParagraphs: 2,
                text: 'Alpha\nBeta'
            };
        }
    });

    const envelope = await runtime.buildContextEnvelope('word', 'Summarize this');
    assert.ok(envelope.includes('PLUGIN_DATA_CONTEXT_START'));
    assert.ok(envelope.includes('"type": "document_context"'));
    assert.ok(envelope.includes('"payload": "Alpha\\nBeta"'));

    const prepared = await runtime.prepareRequestWithContext([
        { role: 'system', content: 'System message' },
        { role: 'user', content: 'Summarize this' }
    ], 'Summarize this');

    assert.strictEqual(prepared.length, 3);
    assert.strictEqual(prepared[1].role, 'user');
    assert.ok(prepared[1].content.includes('PLUGIN_DATA_CONTEXT_START'));
    assert.strictEqual(prepared[2].content, 'Summarize this');
}

async function testSpreadsheetContextParsesJsonStringPayloads() {
    const calls = [];
    const runtime = createContextRuntime({
        async callEditorCommand(command) {
            calls.push(command.toString());
            const source = command.toString();
            if (source.includes('safeSheetName') && source.includes('GetSelection')) {
                return JSON.stringify({
                    sheetName: 'Лист1',
                    address: 'A1:B2',
                    values: [['A', 'B'], ['1', '2']]
                });
            }
            if (source.includes('listSheets')) {
                return JSON.stringify({
                    sheets: [{ name: 'Лист1', address: 'A1:B2' }],
                    discovery_status: 'ok',
                    sources_tried: ['Api.Sheets']
                });
            }
            return null;
        }
    });

    const activeSheet = await runtime.collectActiveSheetContext({ force: true });
    const workbookSheets = await runtime.collectWorkbookSheetsList({ force: true });

    assert.deepStrictEqual(activeSheet, {
        sheetName: 'Лист1',
        address: 'A1:B2',
        values: [['A', 'B'], ['1', '2']]
    });
    assert.deepStrictEqual(workbookSheets, {
        sheets: [{ name: 'Лист1', address: 'A1:B2' }],
        discovery_status: 'ok',
        sources_tried: ['Api.Sheets']
    });
}

async function testDiscoverAndCollectContextForCellUsesChunking() {
    const rangeCalls = [];
    const runtime = createContextRuntime({
        getEditorTypeSafe() {
            return 'cell';
        },
        async callEditorCommand(command, scopeData) {
            const source = command.toString();
            if (source.includes('selectionAddress')) {
                return JSON.stringify({
                    sheetName: 'Sales',
                    address: 'A1:C5',
                    selectionAddress: 'A1:B2'
                });
            }
            if (source.includes('listSheets')) {
                return JSON.stringify({
                    sheets: [{ name: 'Sales', address: 'A1:C5' }],
                    discovery_status: 'ok',
                    sources_tried: ['Api.Sheets']
                });
            }
            if (scopeData && scopeData.sheetName === 'Sales' && source.includes('safeRangeAddress(usedRange)') && !source.includes('Asc.scope.range')) {
                return JSON.stringify({
                    sheetName: 'Sales',
                    address: 'A1:C5'
                });
            }
            if (source.includes('Asc.scope.range')) {
                rangeCalls.push(scopeData.range);
                if (scopeData.range === 'A1:B2') {
                    return JSON.stringify({
                        sheetName: scopeData.sheetName,
                        address: scopeData.range,
                        values: [['Month', 'Revenue'], ['Jan', '100']]
                    });
                }
                if (scopeData.range === 'A3:B4') {
                    return JSON.stringify({
                        sheetName: scopeData.sheetName,
                        address: scopeData.range,
                        values: [['Feb', '120'], ['Mar', '140']]
                    });
                }
                return JSON.stringify({
                    sheetName: scopeData.sheetName,
                    address: scopeData.range,
                    values: [['Apr', '160']]
                });
            }
            return null;
        }
    });

    const discover = await runtime.discoverAvailableContext({ editorType: 'cell', forceRefresh: true });
    const collected = await runtime.collectContext({
        editorType: 'cell',
        sheetName: 'Sales',
        maxRowsPerChunk: 2,
        maxColsPerChunk: 2,
        maxInternalChunks: 2,
        maxChars: 120
    });

    assert.strictEqual(discover.mode, 'discover');
    assert.strictEqual(discover.activeSheet, 'Sales');
    assert.deepStrictEqual(discover.sheets, [{
        name: 'Sales',
        address: 'A1:C5',
        rows: 5,
        cols: 3
    }]);

    assert.deepStrictEqual(rangeCalls, ['A1:B2', 'A3:B4']);
    assert.strictEqual(collected.mode, 'collect');
    assert.strictEqual(collected.sheetName, 'Sales');
    assert.strictEqual(collected.coverage.rows, 5);
    assert.strictEqual(collected.coverage.cols, 3);
    assert.strictEqual(collected.coverage.collectedRows, 4);
    assert.strictEqual(collected.coverage.collectedCols, 2);
    assert.strictEqual(collected.internalChunksUsed, 2);
    assert.strictEqual(collected.truncated, true);
    assert.deepStrictEqual(collected.preview.headers, ['Month', 'Revenue']);
    assert.deepStrictEqual(collected.preview.rows.slice(0, 3), [
        ['Month', 'Revenue'],
        ['Jan', '100'],
        ['Feb', '120']
    ]);
    assert.ok(collected.warnings.includes('columns_limited_to_first_2'));
    assert.ok(collected.warnings.includes('max_internal_chunks_reached'));
    assert.ok(collected.payload.includes('[A1:B2]'));
    assert.ok(collected.payload.includes('[A3:B4]'));
}

async function testCollectActiveSheetPrefersUsedRangeOverCurrentSelection() {
    const runtime = createContextRuntime({
        getEditorTypeSafe() {
            return 'cell';
        },
        async callEditorCommand(command) {
            global.Asc = { scope: {} };
            global.Api = {
                GetActiveSheet() {
                    return {
                        Name: 'Лист1',
                        GetUsedRange() {
                            return {
                                Address: '$A$1:$K$25',
                                GetValue2() {
                                    return [
                                        ['', '', '', '', '', '', '', '', '', '', ''],
                                        ['', '', '', '', '', '', '', '', '', '', ''],
                                        ['', '', '', '', '', '', '6', '', '', '', '']
                                    ];
                                }
                            };
                        },
                        GetSelection() {
                            return {
                                Address: '$A$1',
                                GetValue2() {
                                    return [['']];
                                }
                            };
                        }
                    };
                }
            };
            return command();
        }
    });

    const activeSheet = await runtime.collectActiveSheetContext({ force: true });

    assert.strictEqual(activeSheet.sheetName, 'Лист1');
    assert.strictEqual(activeSheet.address, '$A$1:$K$25');
    assert.deepStrictEqual(activeSheet.values, [
        ['', '', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '', '', '', '', ''],
        ['', '', '', '', '', '', '6', '', '', '', '']
    ]);
}

async function testCollectContextForWordUsesParagraphChunks() {
    const chunkCalls = [];
    const runtime = createContextRuntime({
        getEditorTypeSafe() {
            return 'word';
        },
        async callEditorCommand(command, scopeData) {
            const source = command.toString();
            if (source.includes("source: 'document_meta'")) {
                return JSON.stringify({
                    source: 'document_meta',
                    totalParagraphs: 5
                });
            }
            if (source.includes('Asc.scope.startIndex')) {
                chunkCalls.push(scopeData.startIndex);
                if (scopeData.startIndex === 0) {
                    return JSON.stringify({
                        source: 'document_chunk',
                        startIndex: 0,
                        endIndex: 2,
                        totalParagraphs: 5,
                        text: 'P1\nP2'
                    });
                }
                return JSON.stringify({
                    source: 'document_chunk',
                    startIndex: 2,
                    endIndex: 4,
                    totalParagraphs: 5,
                    text: 'P3\nP4'
                });
            }
            return null;
        }
    });

    const result = await runtime.collectContext({
        editorType: 'word',
        maxParagraphsPerChunk: 2,
        maxInternalChunks: 2,
        maxChars: 120
    });

    assert.deepStrictEqual(chunkCalls, [0, 2]);
    assert.strictEqual(result.mode, 'collect');
    assert.strictEqual(result.source, 'document');
    assert.strictEqual(result.coverage.totalParagraphs, 5);
    assert.strictEqual(result.coverage.collectedParagraphs, 4);
    assert.strictEqual(result.internalChunksUsed, 2);
    assert.strictEqual(result.truncated, true);
    assert.ok(result.warnings.includes('max_internal_chunks_reached'));
    assert.ok(result.payload.includes('[Paragraphs 1-2]'));
    assert.ok(result.payload.includes('[Paragraphs 3-4]'));
}

async function run() {
    testSanitizeAndHelpers();
    await testBuildEnvelopeAndRequestPreparation();
    await testSpreadsheetContextParsesJsonStringPayloads();
    await testDiscoverAndCollectContextForCellUsesChunking();
    await testCollectActiveSheetPrefersUsedRangeOverCurrentSelection();
    await testCollectContextForWordUsesParagraphChunks();
    console.log('r7chat_context_runtime.test.js: ok');
}

run().catch((error) => {
    console.error(error);
    process.exit(1);
});
