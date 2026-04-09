(function (rootFactory) {
    'use strict';

    if (typeof module === 'object' && module.exports) {
        module.exports = rootFactory(typeof globalThis !== 'undefined' ? globalThis : global, true);
        return;
    }

    rootFactory(typeof window !== 'undefined' ? window : this, false);
})(function (globalRoot, isNode) {
    'use strict';

    var root = globalRoot.R7Chat = globalRoot.R7Chat || {};
    root.features = root.features || {};
    root.services = root.services || {};
    root.context = root.context || {};
    root.runtime = root.runtime || {};
    root.runtime.state = root.runtime.state || {};
    root.runtime.state.context = root.runtime.state.context || {
        autoModeWord: 'full_document',
        autoModeCell: 'active_sheet',
        attachedSheets: [],
        smartLimit: true,
        workbookSheetCache: [],
        lastToolSummary: null
    };

    function constants() {
        return root.runtime.constants || {
            contextStateKey: 'openrouter_context_state_v1',
            contextCacheTtlMs: 1800,
            contextCommandTimeoutMs: 2500,
            contextLimits: {
                maxDocumentChars: 12000,
                maxSheetChars: 9000,
                maxPreviewChars: 1800,
                maxRowsPerSheet: 80,
                maxColsPerSheet: 20,
                maxPreviewRows: 16,
                maxPreviewCols: 8,
                maxAttachedSheets: 6
            },
            contextToolLimits: {
                maxCharsWord: 12000,
                maxCharsCell: 9000,
                maxRowsPerChunk: 80,
                maxColsPerChunk: 20,
                maxParagraphsPerChunk: 120,
                maxInternalChunks: 8
            }
        };
    }

    function contextState() {
        return root.runtime.state.context;
    }

    function getStorage() {
        if (typeof globalRoot.localStorage !== 'undefined' && globalRoot.localStorage) {
            return globalRoot.localStorage;
        }
        return null;
    }

    function getHostBridge() {
        return root.platform && root.platform.hostBridge ? root.platform.hostBridge : null;
    }

    function getCacheApi() {
        return root.context && root.context.cache ? root.context.cache : null;
    }

    function normalizeTextPayload(value) {
        if (value === null || value === undefined) return '';
        return String(value).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    }

    function applyTextSmartLimit(text, maxChars) {
        var sourceText = normalizeTextPayload(text);
        if (!contextState().smartLimit || sourceText.length <= maxChars) {
            return { text: sourceText, truncated: false };
        }
        return {
            text: sourceText.slice(0, maxChars),
            truncated: true
        };
    }

    function normalizeTableValues(values) {
        if (!Array.isArray(values)) {
            return [[normalizeTextPayload(values)]];
        }
        return values.map(function (row) {
            if (Array.isArray(row)) return row;
            return [row];
        });
    }

    function normalizePreviewRows(values, maxRows, maxCols) {
        var table = normalizeTableValues(values);
        var rows = [];
        var rowLimit = Math.max(1, Number(maxRows || 6) || 6);
        var colLimit = Math.max(1, Number(maxCols || 8) || 8);
        for (var r = 0; r < table.length && rows.length < rowLimit; r += 1) {
            var sourceRow = Array.isArray(table[r]) ? table[r] : [table[r]];
            var row = [];
            for (var c = 0; c < sourceRow.length && c < colLimit; c += 1) {
                row.push(normalizeTextPayload(sourceRow[c]));
            }
            rows.push(row);
        }
        return rows;
    }

    function formatTablePayload(values, options) {
        var table = normalizeTableValues(values);
        var settings = options && typeof options === 'object' ? options : {};
        var maxRows = Number(settings.maxRows || constants().contextLimits.maxRowsPerSheet);
        var maxCols = Number(settings.maxCols || constants().contextLimits.maxColsPerSheet);
        var maxChars = Number(settings.maxChars || constants().contextLimits.maxSheetChars);
        var totalRows = table.length;
        var totalCols = table.reduce(function (maxValue, row) {
            return Math.max(maxValue, Array.isArray(row) ? row.length : 0);
        }, 0);
        var sampledRows = contextState().smartLimit ? Math.min(totalRows, maxRows) : totalRows;
        var sampledCols = contextState().smartLimit ? Math.min(totalCols || 1, maxCols) : (totalCols || 1);
        var truncated = sampledRows < totalRows || sampledCols < totalCols;
        var rowsToRender = sampledRows;
        var renderedText = '';

        while (rowsToRender > 0) {
            var lines = [];
            for (var r = 0; r < rowsToRender; r += 1) {
                var row = Array.isArray(table[r]) ? table[r] : [table[r]];
                var cells = [];
                for (var c = 0; c < sampledCols; c += 1) {
                    var rawCell = row[c] === undefined ? '' : row[c];
                    var normalizedCell = normalizeTextPayload(rawCell).replace(/\n/g, ' ').trim();
                    cells.push(normalizedCell);
                }
                lines.push(cells.join('\t'));
            }
            renderedText = lines.join('\n');
            if (!contextState().smartLimit || renderedText.length <= maxChars) {
                break;
            }
            rowsToRender -= 1;
            truncated = true;
        }

        if (contextState().smartLimit && renderedText.length > maxChars) {
            renderedText = renderedText.slice(0, maxChars);
            truncated = true;
        }

        return {
            text: renderedText,
            totalRows: totalRows,
            totalCols: totalCols,
            sampledRows: rowsToRender,
            sampledCols: sampledCols,
            truncated: truncated
        };
    }

    function colLettersToNumber(value) {
        var result = 0;
        var text = String(value || '').toUpperCase();
        for (var i = 0; i < text.length; i += 1) {
            result = result * 26 + (text.charCodeAt(i) - 64);
        }
        return result;
    }

    function parseRangeDimensions(address) {
        if (!address || typeof address !== 'string') return { rows: 0, cols: 0 };
        var pureAddress = address.split('!').pop().replace(/\$/g, '');
        var match = pureAddress.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
        if (!match) return { rows: 0, cols: 0 };

        var startCol = colLettersToNumber(match[1]);
        var startRow = Number(match[2]);
        var endCol = colLettersToNumber(match[3] || match[1]);
        var endRow = Number(match[4] || match[2]);

        return {
            rows: Math.max(0, endRow - startRow + 1),
            cols: Math.max(0, endCol - startCol + 1)
        };
    }

    function numberToColLetters(value) {
        var num = Math.max(1, Number(value) || 1);
        var output = '';
        while (num > 0) {
            var mod = (num - 1) % 26;
            output = String.fromCharCode(65 + mod) + output;
            num = Math.floor((num - 1) / 26);
        }
        return output || 'A';
    }

    function parseRangeBounds(address) {
        if (!address || typeof address !== 'string') {
            return {
                address: '',
                startRow: 0,
                endRow: 0,
                startCol: 0,
                endCol: 0,
                rows: 0,
                cols: 0
            };
        }

        var pureAddress = address.split('!').pop().replace(/\$/g, '');
        var match = pureAddress.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i);
        if (!match) {
            return {
                address: pureAddress,
                startRow: 0,
                endRow: 0,
                startCol: 0,
                endCol: 0,
                rows: 0,
                cols: 0
            };
        }

        var startCol = colLettersToNumber(match[1]);
        var startRow = Number(match[2]);
        var endCol = colLettersToNumber(match[3] || match[1]);
        var endRow = Number(match[4] || match[2]);

        return {
            address: pureAddress,
            startRow: Math.min(startRow, endRow),
            endRow: Math.max(startRow, endRow),
            startCol: Math.min(startCol, endCol),
            endCol: Math.max(startCol, endCol),
            rows: Math.max(0, Math.max(startRow, endRow) - Math.min(startRow, endRow) + 1),
            cols: Math.max(0, Math.max(startCol, endCol) - Math.min(startCol, endCol) + 1)
        };
    }

    function buildRangeAddress(startRow, startCol, endRow, endCol) {
        if (!startRow || !startCol || !endRow || !endCol) return '';
        return numberToColLetters(startCol) + String(startRow) + ':' + numberToColLetters(endCol) + String(endRow);
    }

    function cloneSerializable(value) {
        if (value === null || value === undefined) return value;
        try {
            return JSON.parse(JSON.stringify(value));
        } catch (error) {
            return value;
        }
    }

    function clampPositiveInt(value, fallback, maxValue) {
        var parsed = Number(value);
        if (!Number.isFinite(parsed) || parsed <= 0) parsed = Number(fallback);
        if (!Number.isFinite(parsed) || parsed <= 0) parsed = 1;
        parsed = Math.floor(parsed);
        if (Number.isFinite(maxValue) && maxValue > 0) {
            parsed = Math.min(parsed, Math.floor(maxValue));
        }
        return Math.max(1, parsed);
    }

    function normalizeContextToolOptions(options, editorType) {
        var source = options && typeof options === 'object' ? options : {};
        var toolLimits = constants().contextToolLimits || {};
        var isWord = editorType === 'word';
        return {
            mode: source.mode === 'discover' ? 'discover' : 'collect',
            editorType: editorType === 'word' ? 'word' : 'cell',
            sheetName: normalizeTextPayload(source.sheetName || '').trim(),
            intent: normalizeTextPayload(source.intent || '').trim(),
            forceRefresh: source.forceRefresh === true,
            maxChars: clampPositiveInt(
                source.maxChars,
                isWord ? toolLimits.maxCharsWord : toolLimits.maxCharsCell,
                isWord ? 24000 : 24000
            ),
            maxRowsPerChunk: clampPositiveInt(source.maxRowsPerChunk, toolLimits.maxRowsPerChunk || 80, 400),
            maxColsPerChunk: clampPositiveInt(source.maxColsPerChunk, toolLimits.maxColsPerChunk || 20, 64),
            maxParagraphsPerChunk: clampPositiveInt(source.maxParagraphsPerChunk, toolLimits.maxParagraphsPerChunk || 120, 400),
            maxInternalChunks: clampPositiveInt(source.maxInternalChunks, toolLimits.maxInternalChunks || 8, 24)
        };
    }

    function appendBoundedChunk(currentText, chunkText, maxChars) {
        var safeCurrent = normalizeTextPayload(currentText || '');
        var safeChunk = normalizeTextPayload(chunkText || '');
        if (!safeChunk.length) {
            return {
                text: safeCurrent,
                appended: false,
                truncated: false
            };
        }
        var nextText = safeCurrent ? (safeCurrent + '\n\n' + safeChunk) : safeChunk;
        if (nextText.length <= maxChars) {
            return {
                text: nextText,
                appended: true,
                truncated: false
            };
        }
        if (safeCurrent.length >= maxChars) {
            return {
                text: safeCurrent.slice(0, maxChars),
                appended: false,
                truncated: true
            };
        }
        var available = maxChars - safeCurrent.length;
        if (safeCurrent.length && available > 2) {
            return {
                text: safeCurrent + '\n\n' + safeChunk.slice(0, Math.max(0, available - 2)),
                appended: true,
                truncated: true
            };
        }
        return {
            text: safeChunk.slice(0, maxChars),
            appended: true,
            truncated: true
        };
    }

    function setLastToolSummary(summary) {
        contextState().lastToolSummary = summary && typeof summary === 'object'
            ? cloneSerializable(summary)
            : null;
        return contextState().lastToolSummary;
    }

    function cloneMessages(messages) {
        if (!Array.isArray(messages)) return [];
        return messages.map(function (item) {
            var output = {
                role: item.role,
                content: item.content
            };
            if (Array.isArray(item.attachments)) {
                output.attachments = item.attachments.map(function (attachment) {
                    return attachment && typeof attachment === 'object' ? Object.assign({}, attachment) : attachment;
                }).filter(Boolean);
            }
            if (item.tool_call_id) output.tool_call_id = item.tool_call_id;
            if (item.name) output.name = item.name;
            if (Array.isArray(item.tool_calls)) output.tool_calls = item.tool_calls.slice();
            return output;
        });
    }

    function getDefaultContextState() {
        return {
            autoModeWord: 'full_document',
            autoModeCell: 'active_sheet',
            attachedSheets: [],
            smartLimit: true,
            workbookSheetCache: [],
            lastToolSummary: null
        };
    }

    function normalizeAttachedSheets(sheets) {
        if (!Array.isArray(sheets)) return [];
        var unique = [];
        var seen = {};
        sheets.forEach(function (item) {
            var name = String(item || '').trim();
            if (!name.length) return;
            var key = name.toLowerCase();
            if (seen[key]) return;
            seen[key] = true;
            unique.push(name);
        });
        return unique;
    }

    function sanitizeContextState(inputState) {
        var defaults = getDefaultContextState();
        var safeInput = inputState || {};
        return {
            autoModeWord: safeInput.autoModeWord === 'full_document' ? safeInput.autoModeWord : defaults.autoModeWord,
            autoModeCell: safeInput.autoModeCell === 'active_sheet' ? safeInput.autoModeCell : defaults.autoModeCell,
            attachedSheets: normalizeAttachedSheets(safeInput.attachedSheets),
            smartLimit: safeInput.smartLimit !== false,
            workbookSheetCache: Array.isArray(safeInput.workbookSheetCache) ? safeInput.workbookSheetCache.slice() : [],
            lastToolSummary: safeInput.lastToolSummary && typeof safeInput.lastToolSummary === 'object'
                ? JSON.parse(JSON.stringify(safeInput.lastToolSummary))
                : null
        };
    }

    function loadContextState() {
        var storage = getStorage();
        var nextState = getDefaultContextState();
        try {
            var raw = storage ? storage.getItem(constants().contextStateKey) : null;
            if (raw) {
                nextState = sanitizeContextState(JSON.parse(raw));
            }
        } catch (error) {
            console.error('Failed to parse context state, fallback to defaults', error);
        }
        root.runtime.state.context = nextState;
        return nextState;
    }

    function saveContextState() {
        var storage = getStorage();
        root.runtime.state.context = sanitizeContextState(root.runtime.state.context);
        if (storage) {
            storage.setItem(constants().contextStateKey, JSON.stringify({
                autoModeWord: contextState().autoModeWord,
                autoModeCell: contextState().autoModeCell,
                attachedSheets: contextState().attachedSheets,
                smartLimit: contextState().smartLimit
            }));
        }
        return contextState();
    }

    function callEditorCommand(command, scopeData, options) {
        var host = getHostBridge();
        if (!host || typeof host.callEditorCommand !== 'function') {
            return Promise.resolve(null);
        }
        var settings = Object.assign({
            timeoutMs: Number(constants().contextCommandTimeoutMs || 2500)
        }, options && typeof options === 'object' ? options : {});
        return host.callEditorCommand(command, scopeData, settings);
    }

    function callEditorMethod(name, args) {
        var host = getHostBridge();
        if (!host || typeof host.callEditorMethod !== 'function') {
            return Promise.resolve(null);
        }
        return host.callEditorMethod(name, args);
    }

    function parseJsonEditorPayload(value, fallbackValue, label) {
        if (value === null || value === undefined || value === '') {
            return fallbackValue;
        }
        if (typeof value === 'string') {
            try {
                return JSON.parse(value);
            } catch (error) {
                console.error('Failed to parse ' + label, error);
                return fallbackValue;
            }
        }
        return value;
    }

    async function collectWordDocumentContext() {
        var result = await callEditorCommand(function () {
            try {
                var doc = Api.GetDocument();
                if (!doc) {
                    return JSON.stringify({ source: 'full_document', totalParagraphs: 0, text: '' });
                }
                var paragraphs = [];
                if (typeof doc.GetAllParagraphs === 'function') {
                    paragraphs = doc.GetAllParagraphs() || [];
                }
                var chunks = [];
                for (var i = 0; i < paragraphs.length; i += 1) {
                    var text = paragraphs[i].GetText();
                    if (text !== null && text !== undefined && String(text).trim().length) {
                        chunks.push(String(text));
                    }
                }
                return JSON.stringify({
                    source: 'full_document',
                    totalParagraphs: paragraphs.length,
                    text: chunks.join('\n')
                });
            } catch (error) {
                return JSON.stringify({
                    source: 'full_document_error',
                    totalParagraphs: 0,
                    text: '',
                    error: String(error && error.message ? error.message : error)
                });
            }
        });
        return parseJsonEditorPayload(result, {
            source: 'full_document_error',
            totalParagraphs: 0,
            text: ''
        }, 'doc context');
    }

    async function collectWordDocumentMeta() {
        var payload = await callEditorCommand(function () {
            try {
                var doc = Api.GetDocument ? Api.GetDocument() : null;
                if (!doc) {
                    return JSON.stringify({
                        source: 'document_meta',
                        totalParagraphs: 0
                    });
                }
                var paragraphs = typeof doc.GetAllParagraphs === 'function' ? (doc.GetAllParagraphs() || []) : [];
                return JSON.stringify({
                    source: 'document_meta',
                    totalParagraphs: paragraphs.length
                });
            } catch (error) {
                return JSON.stringify({
                    source: 'document_meta_error',
                    totalParagraphs: 0,
                    error: String(error && error.message ? error.message : error)
                });
            }
        });

        return parseJsonEditorPayload(payload, {
            source: 'document_meta_error',
            totalParagraphs: 0
        }, 'word document meta');
    }

    async function collectWordDocumentChunk(startIndex, limit) {
        var payload = await callEditorCommand(function () {
            try {
                var doc = Api.GetDocument ? Api.GetDocument() : null;
                var start = Math.max(0, Number(Asc.scope.startIndex || 0));
                var size = Math.max(1, Number(Asc.scope.limit || 1));
                if (!doc) {
                    return JSON.stringify({
                        source: 'document_chunk',
                        startIndex: start,
                        endIndex: start,
                        totalParagraphs: 0,
                        text: ''
                    });
                }
                var paragraphs = typeof doc.GetAllParagraphs === 'function' ? (doc.GetAllParagraphs() || []) : [];
                var collected = [];
                var endIndex = Math.min(paragraphs.length, start + size);
                for (var i = start; i < endIndex; i += 1) {
                    var paragraph = paragraphs[i];
                    var text = paragraph && typeof paragraph.GetText === 'function' ? paragraph.GetText() : '';
                    if (text !== null && text !== undefined) {
                        text = String(text);
                        if (text.trim().length) collected.push(text);
                    }
                }
                return JSON.stringify({
                    source: 'document_chunk',
                    startIndex: start,
                    endIndex: endIndex,
                    totalParagraphs: paragraphs.length,
                    text: collected.join('\n')
                });
            } catch (error) {
                return JSON.stringify({
                    source: 'document_chunk_error',
                    startIndex: Number(Asc.scope.startIndex || 0),
                    endIndex: Number(Asc.scope.startIndex || 0),
                    totalParagraphs: 0,
                    text: '',
                    error: String(error && error.message ? error.message : error)
                });
            }
        }, {
            startIndex: Number(startIndex || 0) || 0,
            limit: Number(limit || 1) || 1
        });

        return parseJsonEditorPayload(payload, {
            source: 'document_chunk_error',
            startIndex: Number(startIndex || 0) || 0,
            endIndex: Number(startIndex || 0) || 0,
            totalParagraphs: 0,
            text: ''
        }, 'word document chunk');
    }

    async function collectActiveSheetMeta() {
        var payload = await callEditorCommand(function () {
            function safeSheetName(sheet) {
                try {
                    if (!sheet) return '';
                    if (sheet.Name) return String(sheet.Name);
                    if (typeof sheet.GetName === 'function') return String(sheet.GetName() || '');
                } catch (error) {}
                return '';
            }

            function safeRangeAddress(range) {
                if (!range) return '';
                try {
                    if (range.Address) return String(range.Address);
                } catch (error) {}
                try {
                    if (typeof range.GetAddress === 'function') {
                        return String(range.GetAddress(true, true, 'xlA1', false) || '');
                    }
                } catch (error2) {}
                return '';
            }

            var ws = null;
            try {
                ws = Api.GetActiveSheet ? Api.GetActiveSheet() : null;
            } catch (error3) {
                ws = null;
            }
            if (!ws) {
                return JSON.stringify({
                    sheetName: '',
                    address: '',
                    selectionAddress: ''
                });
            }

            var usedRange = null;
            var selection = null;
            try {
                if (typeof ws.GetUsedRange === 'function') usedRange = ws.GetUsedRange();
            } catch (error4) {}
            try {
                if (typeof ws.GetSelection === 'function') selection = ws.GetSelection();
            } catch (error5) {}

            return JSON.stringify({
                sheetName: safeSheetName(ws),
                address: safeRangeAddress(usedRange),
                selectionAddress: safeRangeAddress(selection || usedRange)
            });
        });

        return parseJsonEditorPayload(payload, {
            sheetName: '',
            address: '',
            selectionAddress: ''
        }, 'active sheet meta');
    }

    async function collectActiveSheetContext(options) {
        var settings = options && typeof options === 'object' ? options : {};
        var force = settings.force === true;
        var ttlMs = Math.max(0, Number(settings.ttlMs || constants().contextCacheTtlMs));
        var cacheApi = getCacheApi();
        if (cacheApi && typeof cacheApi.getActiveSheet === 'function') {
            var cached = cacheApi.getActiveSheet({ force: force, ttlMs: ttlMs });
            if (cached) {
                return JSON.parse(JSON.stringify(cached));
            }
        }

        var payload = await callEditorCommand(function () {
            function safeSheetName(sheet) {
                try {
                    if (!sheet) return '';
                    if (sheet.Name) return String(sheet.Name);
                    if (typeof sheet.GetName === 'function') return String(sheet.GetName() || '');
                } catch (error) {
                    return '';
                }
                return '';
            }

            function safeRangeAddress(range) {
                if (!range) return '';
                try {
                    if (range.Address) return String(range.Address);
                } catch (error) {}
                try {
                    if (typeof range.GetAddress === 'function') {
                        return String(range.GetAddress(true, true, 'xlA1', false) || '');
                    }
                } catch (error2) {}
                return '';
            }

            function normalizeValues(values) {
                if (Array.isArray(values)) return values;
                if (values === null || values === undefined) return [];
                return [[values]];
            }

            function safeRangeValues(range) {
                if (!range) return [];
                try {
                    if (typeof range.GetValue2 === 'function') {
                        return normalizeValues(range.GetValue2());
                    }
                } catch (error) {}
                try {
                    if (typeof range.GetValue === 'function') {
                        return normalizeValues(range.GetValue());
                    }
                } catch (error2) {}
                try {
                    if (typeof range.GetText === 'function') {
                        var text = String(range.GetText() || '');
                        if (text.length) return [[text]];
                    }
                } catch (error3) {}
                return [];
            }

            var ws = null;
            try {
                ws = Api.GetActiveSheet ? Api.GetActiveSheet() : null;
            } catch (error4) {
                ws = null;
            }
            if (!ws) {
                return JSON.stringify({ sheetName: 'Active sheet', address: '', values: [] });
            }

            var usedRange = null;
            var selection = null;
            try {
                if (typeof ws.GetUsedRange === 'function') usedRange = ws.GetUsedRange();
            } catch (error5) {}
            try {
                if (typeof ws.GetSelection === 'function') selection = ws.GetSelection();
            } catch (error6) {}
            var targetRange = usedRange || selection;

            return JSON.stringify({
                sheetName: safeSheetName(ws) || 'Active sheet',
                address: safeRangeAddress(targetRange),
                values: safeRangeValues(targetRange)
            });
        });

        payload = parseJsonEditorPayload(payload, { sheetName: 'Active sheet', address: '', values: [] }, 'active sheet context');

        if (cacheApi && typeof cacheApi.setActiveSheet === 'function' && payload) {
            cacheApi.setActiveSheet(JSON.parse(JSON.stringify(payload)));
        }
        return payload;
    }

    async function collectWorkbookSheetsList(options) {
        var settings = options && typeof options === 'object' ? options : {};
        var force = settings.force === true;
        var ttlMs = Math.max(0, Number(settings.ttlMs || constants().contextCacheTtlMs));
        var cacheApi = getCacheApi();
        if (cacheApi && typeof cacheApi.getWorkbookSheets === 'function') {
            var cached = cacheApi.getWorkbookSheets({ force: force, ttlMs: ttlMs });
            if (cached) {
                return JSON.parse(JSON.stringify(cached));
            }
        }

        var payload = await callEditorCommand(function () {
            function safeSheetName(sheet, fallback) {
                try {
                    if (!sheet) return fallback || '';
                    if (sheet.Name) return String(sheet.Name);
                    if (typeof sheet.GetName === 'function') return String(sheet.GetName() || fallback || '');
                } catch (error) {}
                return fallback || '';
            }

            function safeRangeAddress(range) {
                if (!range) return '';
                try {
                    if (range.Address) return String(range.Address);
                } catch (error) {}
                try {
                    if (typeof range.GetAddress === 'function') {
                        return String(range.GetAddress(true, true, 'xlA1', false) || '');
                    }
                } catch (error2) {}
                return '';
            }

            function listSheets() {
                if (typeof Api.GetSheets === 'function') {
                    try { return Api.GetSheets(); } catch (error) {}
                }
                if (Array.isArray(Api.Sheets)) return Api.Sheets;
                if (Api.Sheets && typeof Api.Sheets.length === 'number') {
                    var arr = [];
                    for (var i = 0; i < Api.Sheets.length; i += 1) {
                        if (Api.Sheets[i]) arr.push(Api.Sheets[i]);
                    }
                    return arr;
                }
                return [];
            }

            var sheets = listSheets();
            var normalized = [];
            for (var i = 0; i < sheets.length; i += 1) {
                var sheet = sheets[i];
                var usedRange = null;
                try {
                    if (sheet && typeof sheet.GetUsedRange === 'function') {
                        usedRange = sheet.GetUsedRange();
                    }
                } catch (error3) {}
                normalized.push({
                    name: safeSheetName(sheet, 'Sheet ' + (i + 1)),
                    address: safeRangeAddress(usedRange)
                });
            }
            return JSON.stringify({
                sheets: normalized,
                discovery_status: normalized.length ? 'ok' : 'failed',
                sources_tried: ['Api.Sheets']
            });
        });

        payload = parseJsonEditorPayload(payload, { sheets: [], discovery_status: 'failed', sources_tried: [] }, 'workbook sheets list');

        if (!payload) {
            payload = { sheets: [], discovery_status: 'failed', sources_tried: [] };
        }
        if (cacheApi && typeof cacheApi.setWorkbookSheets === 'function') {
            cacheApi.setWorkbookSheets(JSON.parse(JSON.stringify(payload)));
        }
        return payload;
    }

    async function collectSheetByName(name) {
        var payload = await callEditorCommand(function () {
            function normalizeValues(values) {
                if (Array.isArray(values)) return values;
                if (values === null || values === undefined) return [];
                return [[values]];
            }

            function safeRangeValues(range) {
                if (!range) return [];
                try {
                    if (typeof range.GetValue2 === 'function') return normalizeValues(range.GetValue2());
                } catch (error) {}
                try {
                    if (typeof range.GetValue === 'function') return normalizeValues(range.GetValue());
                } catch (error2) {}
                return [];
            }

            var ws = null;
            if (typeof Api.GetSheet === 'function') {
                try {
                    ws = Api.GetSheet(Asc.scope.sheetName);
                } catch (error3) {
                    ws = null;
                }
            }
            if (!ws) return '';

            var usedRange = null;
            try {
                if (typeof ws.GetUsedRange === 'function') usedRange = ws.GetUsedRange();
            } catch (error4) {}
            return JSON.stringify({
                sheetName: String(Asc.scope.sheetName || ''),
                address: usedRange && usedRange.Address ? String(usedRange.Address) : '',
                values: safeRangeValues(usedRange)
            });
        }, { sheetName: name });
        return parseJsonEditorPayload(payload, null, 'sheet by name context');
    }

    async function collectSheetRangeContext(sheetName, rangeText) {
        var payload = await callEditorCommand(function () {
            function normalizeValues(values) {
                if (Array.isArray(values)) return values;
                if (values === null || values === undefined) return [];
                return [[values]];
            }

            function safeRangeValues(range) {
                if (!range) return [];
                try {
                    if (typeof range.GetValue2 === 'function') return normalizeValues(range.GetValue2());
                } catch (error) {}
                try {
                    if (typeof range.GetValue === 'function') return normalizeValues(range.GetValue());
                } catch (error2) {}
                return [];
            }

            var ws = null;
            try {
                ws = Api.GetSheet ? Api.GetSheet(Asc.scope.sheetName) : null;
            } catch (error3) {
                ws = null;
            }
            if (!ws && typeof Api.GetActiveSheet === 'function') {
                try {
                    ws = Api.GetActiveSheet();
                } catch (error4) {
                    ws = null;
                }
            }
            if (!ws) return '';

            var range = null;
            try {
                if (typeof ws.GetRange === 'function' && String(Asc.scope.range || '').trim().length) {
                    range = ws.GetRange(Asc.scope.range);
                }
            } catch (error5) {
                range = null;
            }
            if (!range && typeof ws.GetUsedRange === 'function') {
                try { range = ws.GetUsedRange(); } catch (error6) {}
            }

            return JSON.stringify({
                sheetName: String(Asc.scope.sheetName || ''),
                address: range && range.Address ? String(range.Address) : String(Asc.scope.range || ''),
                values: safeRangeValues(range)
            });
        }, { sheetName: sheetName || '', range: rangeText || '' });
        return parseJsonEditorPayload(payload, null, 'sheet range context');
    }

    async function discoverAvailableContext(options) {
        var host = getHostBridge();
        var requestedEditor = options && options.editorType ? String(options.editorType) : '';
        var editorType = requestedEditor || (host && typeof host.getEditorTypeSafe === 'function' ? host.getEditorTypeSafe() : '');

        if (editorType === 'word') {
            var docMeta = await collectWordDocumentMeta();
            var docWarnings = [];
            if (docMeta && docMeta.error) docWarnings.push('document_meta_error');
            if (!docMeta || !docMeta.totalParagraphs) docWarnings.push('document_empty');
            var docResult = {
                editorType: 'word',
                mode: 'discover',
                source: 'document',
                discoveryStatus: docMeta && !docMeta.error ? 'ok' : 'failed',
                totalParagraphs: Number(docMeta && docMeta.totalParagraphs || 0) || 0,
                warnings: docWarnings
            };
            setLastToolSummary({
                editorType: 'word',
                mode: 'discover',
                source: 'document',
                discoveryStatus: docResult.discoveryStatus,
                totalParagraphs: docResult.totalParagraphs,
                warnings: docWarnings
            });
            return docResult;
        }

        if (editorType !== 'cell') {
            return {
                editorType: normalizeTextPayload(editorType || ''),
                mode: 'discover',
                source: 'unsupported',
                discoveryStatus: 'unsupported',
                warnings: ['editor_not_supported']
            };
        }

        var workbookMeta = await collectWorkbookSheetsList({
            force: options && options.forceRefresh === true
        });
        var activeMeta = await collectActiveSheetMeta();
        var sheets = Array.isArray(workbookMeta && workbookMeta.sheets) ? workbookMeta.sheets.map(function (sheet) {
            var bounds = parseRangeBounds(sheet && sheet.address || '');
            return {
                name: normalizeTextPayload(sheet && sheet.name || ''),
                address: normalizeTextPayload(sheet && sheet.address || ''),
                rows: bounds.rows,
                cols: bounds.cols
            };
        }) : [];
        contextState().workbookSheetCache = sheets.slice();
        var result = {
            editorType: 'cell',
            mode: 'discover',
            source: 'workbook',
            activeSheet: normalizeTextPayload(activeMeta && activeMeta.sheetName || ''),
            sheets: sheets,
            discoveryStatus: workbookMeta && workbookMeta.discovery_status ? String(workbookMeta.discovery_status) : (sheets.length ? 'ok' : 'failed'),
            warnings: []
        };
        if (!sheets.length) result.warnings.push('sheet_enumeration_failed');
        if (!result.activeSheet) result.warnings.push('active_sheet_unknown');
        setLastToolSummary({
            editorType: 'cell',
            mode: 'discover',
            source: 'workbook',
            activeSheet: result.activeSheet,
            sheetCount: sheets.length,
            discoveryStatus: result.discoveryStatus,
            warnings: result.warnings
        });
        return result;
    }

    async function collectCellContext(options) {
        var settings = normalizeContextToolOptions(options, 'cell');
        var warnings = [];
        var meta = settings.sheetName
            ? await callEditorCommand(function () {
                function safeSheetName(sheet) {
                    try {
                        if (!sheet) return '';
                        if (sheet.Name) return String(sheet.Name);
                        if (typeof sheet.GetName === 'function') return String(sheet.GetName() || '');
                    } catch (error) {}
                    return '';
                }

                function safeRangeAddress(range) {
                    if (!range) return '';
                    try {
                        if (range.Address) return String(range.Address);
                    } catch (error) {}
                    try {
                        if (typeof range.GetAddress === 'function') {
                            return String(range.GetAddress(true, true, 'xlA1', false) || '');
                        }
                    } catch (error2) {}
                    return '';
                }

                var ws = null;
                try {
                    ws = Api.GetSheet ? Api.GetSheet(Asc.scope.sheetName) : null;
                } catch (error3) {
                    ws = null;
                }
                if (!ws) {
                    return JSON.stringify({
                        sheetName: String(Asc.scope.sheetName || ''),
                        address: ''
                    });
                }
                var usedRange = null;
                try {
                    if (typeof ws.GetUsedRange === 'function') usedRange = ws.GetUsedRange();
                } catch (error4) {}
                return JSON.stringify({
                    sheetName: safeSheetName(ws) || String(Asc.scope.sheetName || ''),
                    address: safeRangeAddress(usedRange)
                });
            }, { sheetName: settings.sheetName })
            : await collectActiveSheetMeta();

        meta = parseJsonEditorPayload(meta, null, 'sheet meta');
        var bounds = parseRangeBounds(meta && meta.address || '');
        if (!meta || !meta.sheetName) {
            warnings.push('sheet_not_available');
        }
        if (!bounds.rows || !bounds.cols) {
            warnings.push('sheet_empty_or_unavailable');
        }

        var payloadText = '';
        var internalChunksUsed = 0;
        var payloadTruncated = false;
        var collectedRows = 0;
        var effectiveCols = bounds.cols ? Math.min(bounds.cols, settings.maxColsPerChunk) : 0;
        var previewRows = [];
        if (bounds.cols > effectiveCols) {
            warnings.push('columns_limited_to_first_' + effectiveCols);
        }

        if (meta && meta.sheetName && bounds.rows && effectiveCols) {
            var rowCursor = bounds.startRow;
            while (rowCursor <= bounds.endRow && internalChunksUsed < settings.maxInternalChunks && payloadText.length < settings.maxChars) {
                var chunkEndRow = Math.min(bounds.endRow, rowCursor + settings.maxRowsPerChunk - 1);
                var chunkAddress = buildRangeAddress(rowCursor, bounds.startCol, chunkEndRow, bounds.startCol + effectiveCols - 1);
                var chunkData = await collectSheetRangeContext(meta.sheetName, chunkAddress);
                var formattedChunk = formatTablePayload(chunkData && chunkData.values ? chunkData.values : [], {
                    maxRows: settings.maxRowsPerChunk,
                    maxCols: effectiveCols,
                    maxChars: Math.max(settings.maxChars, 4000)
                });
                if (previewRows.length < 8 && chunkData && Array.isArray(chunkData.values)) {
                    Array.prototype.push.apply(previewRows, normalizePreviewRows(chunkData.values, 8 - previewRows.length, effectiveCols));
                }
                var bounded = appendBoundedChunk(payloadText, '[' + chunkAddress + ']\n' + (formattedChunk.text || ''), settings.maxChars);
                payloadText = bounded.text;
                if (bounded.truncated) payloadTruncated = true;
                internalChunksUsed += 1;
                collectedRows += Math.max(0, chunkEndRow - rowCursor + 1);
                rowCursor = chunkEndRow + 1;
                if (bounded.truncated) break;
            }

            if (collectedRows < bounds.rows) {
                payloadTruncated = true;
                if (internalChunksUsed >= settings.maxInternalChunks) {
                    warnings.push('max_internal_chunks_reached');
                } else {
                    warnings.push('max_chars_reached');
                }
            }
        }

        if (!payloadText.length && !warnings.length) {
            warnings.push('sheet_empty');
        }

        var result = {
            editorType: 'cell',
            mode: 'collect',
            source: 'sheet',
            sheetName: normalizeTextPayload(meta && meta.sheetName || settings.sheetName || ''),
            coverage: {
                range: normalizeTextPayload(meta && meta.address || ''),
                rows: bounds.rows,
                cols: bounds.cols,
                collectedRows: Math.min(bounds.rows, collectedRows),
                collectedCols: effectiveCols
            },
            preview: {
                rows: previewRows,
                headers: previewRows.length ? previewRows[0] : []
            },
            payload: payloadText,
            truncated: payloadTruncated || bounds.cols > effectiveCols,
            internalChunksUsed: internalChunksUsed,
            warnings: warnings
        };
        setLastToolSummary({
            editorType: 'cell',
            mode: 'collect',
            source: 'sheet',
            sheetName: result.sheetName,
            truncated: result.truncated,
            internalChunksUsed: internalChunksUsed,
            coverage: cloneSerializable(result.coverage),
            warnings: warnings
        });
        return result;
    }

    async function collectWordToolContext(options) {
        var settings = normalizeContextToolOptions(options, 'word');
        var warnings = [];
        var meta = await collectWordDocumentMeta();
        var totalParagraphs = Number(meta && meta.totalParagraphs || 0) || 0;
        if (meta && meta.error) warnings.push('document_meta_error');
        if (!totalParagraphs) warnings.push('document_empty');

        var payloadText = '';
        var internalChunksUsed = 0;
        var payloadTruncated = false;
        var collectedParagraphs = 0;
        var paragraphCursor = 0;

        while (paragraphCursor < totalParagraphs && internalChunksUsed < settings.maxInternalChunks && payloadText.length < settings.maxChars) {
            var chunk = await collectWordDocumentChunk(paragraphCursor, settings.maxParagraphsPerChunk);
            var startLabel = paragraphCursor + 1;
            var endLabel = Math.max(startLabel, Number(chunk && chunk.endIndex || paragraphCursor));
            var bounded = appendBoundedChunk(
                payloadText,
                '[Paragraphs ' + startLabel + '-' + endLabel + ']\n' + normalizeTextPayload(chunk && chunk.text || ''),
                settings.maxChars
            );
            payloadText = bounded.text;
            if (bounded.truncated) payloadTruncated = true;
            internalChunksUsed += 1;
            if (chunk && Number(chunk.endIndex || 0) > paragraphCursor) {
                collectedParagraphs += Number(chunk.endIndex || 0) - paragraphCursor;
                paragraphCursor = Number(chunk.endIndex || 0);
            } else {
                break;
            }
        }

        if (collectedParagraphs < totalParagraphs) {
            payloadTruncated = true;
            if (internalChunksUsed >= settings.maxInternalChunks) {
                warnings.push('max_internal_chunks_reached');
            } else {
                warnings.push('max_chars_reached');
            }
        }

        var result = {
            editorType: 'word',
            mode: 'collect',
            source: 'document',
            coverage: {
                totalParagraphs: totalParagraphs,
                collectedParagraphs: Math.min(totalParagraphs, collectedParagraphs)
            },
            payload: payloadText,
            truncated: payloadTruncated,
            internalChunksUsed: internalChunksUsed,
            warnings: warnings
        };
        setLastToolSummary({
            editorType: 'word',
            mode: 'collect',
            source: 'document',
            truncated: result.truncated,
            internalChunksUsed: internalChunksUsed,
            coverage: cloneSerializable(result.coverage),
            warnings: warnings
        });
        return result;
    }

    async function collectContext(options) {
        var host = getHostBridge();
        var requestedEditor = options && options.editorType ? String(options.editorType) : '';
        var editorType = requestedEditor || (host && typeof host.getEditorTypeSafe === 'function' ? host.getEditorTypeSafe() : '');
        if (editorType === 'word') {
            return collectWordToolContext(options || {});
        }
        if (editorType === 'cell') {
            return collectCellContext(options || {});
        }
        return {
            editorType: normalizeTextPayload(editorType || ''),
            mode: 'collect',
            source: 'unsupported',
            payload: '',
            truncated: false,
            internalChunksUsed: 0,
            warnings: ['editor_not_supported']
        };
    }

    function isSheetAttached(name) {
        var key = String(name || '').toLowerCase();
        return contextState().attachedSheets.some(function (item) {
            return String(item).toLowerCase() === key;
        });
    }

    function toggleAttachedSheet(name) {
        var key = String(name || '').toLowerCase();
        var list = contextState().attachedSheets;
        var index = list.findIndex(function (item) {
            return String(item).toLowerCase() === key;
        });

        if (index >= 0) {
            list.splice(index, 1);
        } else {
            if (list.length >= constants().contextLimits.maxAttachedSheets) {
                if (!isNode && globalRoot.alert) {
                    globalRoot.alert('You can attach up to 6 sheets at once.');
                }
                return false;
            }
            list.push(name);
        }
        saveContextState();
        return true;
    }

    async function buildContextEnvelope(editorType, userText, baseMessages) {
        var contextItems = [];
        var limits = constants().contextLimits;

        // Add file attachments from the current conversation
        var safeMessages = Array.isArray(baseMessages) ? baseMessages : [];
        var recentAttachments = [];
        for (var i = Math.max(0, safeMessages.length - 4); i < safeMessages.length; i++) {
            var msg = safeMessages[i];
            if (msg && Array.isArray(msg.attachments)) {
                msg.attachments.forEach(function (att) {
                    if (att && att.attachmentType === 'file') {
                        recentAttachments.push(att);
                    }
                });
            }
        }
        
        // Deduplicate recent file attachments by ID
        var uniqueAttachments = [];
        var seenAttachments = {};
        recentAttachments.forEach(function (att) {
            if (att && att.id && !seenAttachments[att.id]) {
                seenAttachments[att.id] = true;
                uniqueAttachments.push(att);
            }
        });

        uniqueAttachments.forEach(function (att) {
            var summary = att.contextSummary;
            if (summary && summary.kind === 'xlsx_workbook') {
                var previewText = 'XLSX Workbook: ' + (att.name || 'unnamed.xlsx') + '\n';
                previewText += 'Sheets (' + summary.sheetCount + ' total): ' + summary.sheetNames.join(', ') + '\n\n';
                if (summary.sheets && summary.sheets.length) {
                    summary.sheets.forEach(function (sheet) {
                        previewText += 'Sheet: ' + sheet.name + ' | range: ' + sheet.range + ' | rows: ' + sheet.rows + ' | cols: ' + sheet.cols + '\n';
                        if (sheet.preview) {
                            previewText += 'Preview:\n' + sheet.preview + '\n\n';
                        }
                    });
                }
                
                var limitedPreview = applyTextSmartLimit(previewText, limits.maxDocumentChars || 12000);
                contextItems.push({
                    type: 'attached_workbook',
                    source: 'file_attachment',
                    sheetName: null,
                    range: null,
                    rows: summary.sheetCount || 0,
                    cols: 1,
                    truncated: limitedPreview.truncated || summary.status === 'metadata_only',
                    payload: limitedPreview.text
                });
            }
        });

        if (editorType === 'word' && contextState().autoModeWord === 'full_document') {
            var documentContext = await collectWordDocumentContext();
            if (documentContext) {
                var limitedDoc = applyTextSmartLimit(documentContext.text, limits.maxDocumentChars);
                contextItems.push({
                    type: 'document_context',
                    source: documentContext.error ? 'full_document_error' : 'full_document_auto',
                    sheetName: null,
                    range: null,
                    rows: documentContext.totalParagraphs || 0,
                    cols: 1,
                    truncated: limitedDoc.truncated,
                    payload: documentContext.error ? 'Error extracting document text: ' + documentContext.error : limitedDoc.text
                });
            }
        }

        if (editorType === 'cell' && contextState().autoModeCell === 'active_sheet') {
            var activeSheet = await collectActiveSheetContext();
            if (activeSheet) {
                var activePayload = formatTablePayload(activeSheet.values, {
                    maxRows: limits.maxRowsPerSheet,
                    maxCols: limits.maxColsPerSheet,
                    maxChars: limits.maxSheetChars
                });
                contextItems.push({
                    type: 'sheet_context',
                    source: 'active_sheet_auto',
                    sheetName: activeSheet.sheetName || '',
                    range: activeSheet.address || '',
                    rows: activePayload.totalRows,
                    cols: activePayload.totalCols,
                    truncated: activePayload.truncated,
                    payload: activePayload.text
                });
            }

            var workbookSheetsMeta = await collectWorkbookSheetsList();
            var workbookSheets = Array.isArray(workbookSheetsMeta && workbookSheetsMeta.sheets) ? workbookSheetsMeta.sheets : [];
            if (workbookSheets.length) {
                var listText = workbookSheets.map(function (sheet, index) {
                    var dimensions = parseRangeDimensions(sheet.address || '');
                    return (index + 1) + '. ' + sheet.name + ' | range: ' + (sheet.address || '-') + ' | rows: ' + dimensions.rows + ' | cols: ' + dimensions.cols;
                }).join('\n');
                var limitedListText = applyTextSmartLimit(listText, 2200);
                contextItems.push({
                    type: 'workbook_context',
                    source: 'workbook_sheet_list',
                    sheetName: null,
                    range: null,
                    rows: workbookSheets.length,
                    cols: 2,
                    truncated: limitedListText.truncated,
                    payload: limitedListText.text
                });
            } else {
                contextItems.push({
                    type: 'workbook_context',
                    source: 'workbook_sheet_list_unavailable',
                    sheetName: null,
                    range: null,
                    rows: 0,
                    cols: 1,
                    truncated: false,
                    payload: 'sheet_enumeration_failed status=' + (workbookSheetsMeta && workbookSheetsMeta.discovery_status ? workbookSheetsMeta.discovery_status : 'failed')
                });
            }

            var attachedSheetNames = normalizeAttachedSheets(contextState().attachedSheets).slice(0, limits.maxAttachedSheets);
            contextState().attachedSheets = attachedSheetNames;
            saveContextState();
            for (var i = 0; i < attachedSheetNames.length; i += 1) {
                var attachedName = attachedSheetNames[i];
                if (activeSheet && String(activeSheet.sheetName || '').toLowerCase() === attachedName.toLowerCase()) {
                    continue;
                }
                var sheetData = await collectSheetByName(attachedName);
                if (!sheetData) continue;
                var sheetPayload = formatTablePayload(sheetData.values, {
                    maxRows: limits.maxRowsPerSheet,
                    maxCols: limits.maxColsPerSheet,
                    maxChars: limits.maxSheetChars
                });
                contextItems.push({
                    type: 'sheet_context',
                    source: 'attached_sheet',
                    sheetName: sheetData.sheetName || attachedName,
                    range: sheetData.address || '',
                    rows: sheetPayload.totalRows,
                    cols: sheetPayload.totalCols,
                    truncated: sheetPayload.truncated,
                    payload: sheetPayload.text
                });
            }
        }

        if (!contextItems.length) return null;

        var envelope = {
            editorType: editorType,
            smartLimit: contextState().smartLimit === true,
            generatedAt: new Date().toISOString(),
            userMessage: normalizeTextPayload(userText || ''),
            contexts: contextItems
        };

        return [
            'PLUGIN_DATA_CONTEXT_START',
            'The following JSON contains document/sheet data extracted by the plugin.',
            'You MUST use this data when user asks about document, workbook, sheet, rows, columns, or cells.',
            'Do NOT claim you have no access to files/data if context is provided below.',
            JSON.stringify(envelope, null, 2),
            'When context item has truncated=true and details are missing, ask user to specify sheet name or range.',
            'PLUGIN_DATA_CONTEXT_END'
        ].join('\n');
    }

    async function prepareRequestWithContext(baseMessages, userText) {
        var request = cloneMessages(baseMessages);
        var host = getHostBridge();
        var editorType = host && typeof host.getEditorTypeSafe === 'function' ? host.getEditorTypeSafe() : '';
        var contextEnvelope = await buildContextEnvelope(editorType, userText, baseMessages);
        if (!contextEnvelope) return request;

        var contextMessage = {
            role: 'user',
            content: contextEnvelope
        };
        var lastItem = request[request.length - 1];
        if (lastItem && lastItem.role === 'user') {
            request.splice(request.length - 1, 0, contextMessage);
        } else {
            request.push(contextMessage);
        }
        return request;
    }

    var api = {
        getState: function () { return contextState(); },
        getDefaultContextState: getDefaultContextState,
        sanitizeContextState: sanitizeContextState,
        loadContextState: loadContextState,
        saveContextState: saveContextState,
        normalizeTextPayload: normalizeTextPayload,
        applyTextSmartLimit: applyTextSmartLimit,
        formatTablePayload: formatTablePayload,
        parseRangeDimensions: parseRangeDimensions,
        cloneMessages: cloneMessages,
        collectWordDocumentContext: collectWordDocumentContext,
        collectActiveSheetContext: collectActiveSheetContext,
        collectWorkbookSheetsList: collectWorkbookSheetsList,
        collectSheetByName: collectSheetByName,
        collectSheetRangeContext: collectSheetRangeContext,
        collectWordDocumentMeta: collectWordDocumentMeta,
        collectWordDocumentChunk: collectWordDocumentChunk,
        collectActiveSheetMeta: collectActiveSheetMeta,
        discoverAvailableContext: discoverAvailableContext,
        collectContext: collectContext,
        buildContextEnvelope: buildContextEnvelope,
        prepareRequestWithContext: prepareRequestWithContext,
        isSheetAttached: isSheetAttached,
        toggleAttachedSheet: toggleAttachedSheet,
        getWorkbookSheets: function (options) {
            return collectWorkbookSheetsList(options);
        },
        getActiveSheet: function (options) {
            return collectActiveSheetContext(options);
        },
        getLastToolSummary: function () {
            return cloneSerializable(contextState().lastToolSummary || null);
        }
    };

    root.features.contextRuntime = api;
    root.services.context = root.services.context || api;
    root.context.collectSheetByName = collectSheetByName;
    root.context.getWorkbookSheets = api.getWorkbookSheets;
    root.context.getActiveSheet = api.getActiveSheet;

    return api;
});
