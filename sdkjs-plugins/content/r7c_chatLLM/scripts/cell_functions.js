(function (window) {
    'use strict';

    if (!window || !window.Asc || !window.Asc.plugin) return;

    var core = window.R7CellAiCore || {
        normalizeScalar: function (v) { return String(v === null || v === undefined ? '' : v).trim(); },
        normalizeEnum: function (v) { return String(v === null || v === undefined ? '' : v).trim().toLowerCase(); },
        normalizeLabels: function (v) { return String(v === null || v === undefined ? '' : v).split(',').map(function (x) { return x.trim().toLowerCase(); }).filter(Boolean); },
        simpleHash: function (v) {
            var s = String(v || '');
            var h = 2166136261;
            for (var i = 0; i < s.length; i += 1) {
                h ^= s.charCodeAt(i);
                h = Math.imul(h, 16777619);
            }
            return (h >>> 0).toString(16);
        }
    };

    var LIBRARY_ID = 'R7AI_V1';
    var LIBRARY_MARKER = '/*R7AI_LIBRARY:' + LIBRARY_ID + '*/';
    var TRACE_PULL_INTERVAL_MS = 4000;
    var PROMPT_TEMPLATE_VERSION = 'v1';

    var FUNCTION_DEFS = [
        { name: 'R7_ASK', alias: 'R7.ASK', signature: 'R7_ASK(prompt; value)', example: '=R7_ASK("Сделай краткое резюме"; A2)', description: 'Универсальная AI-функция для одного значения.' },
        { name: 'R7_TRANSLATE', alias: 'R7.TRANSLATE', signature: 'R7_TRANSLATE(text; targetLang)', example: '=R7_TRANSLATE(A2; "RU")', description: 'Перевод текста в целевой язык.' },
        { name: 'R7_EXTRACT', alias: 'R7.EXTRACT', signature: 'R7_EXTRACT(kind; text)', example: '=R7_EXTRACT("email"; B2)', description: 'Извлечение сущностей: email/phone/company/name/date/amount.' },
        { name: 'R7_CLASSIFY', alias: 'R7.CLASSIFY', signature: 'R7_CLASSIFY(text; labels)', example: '=R7_CLASSIFY(C2; "lead,spam,client,partner")', description: 'Классификация по списку категорий.' },
        { name: 'R7_SUMMARIZE', alias: 'R7.SUMMARIZE', signature: 'R7_SUMMARIZE(text; mode)', example: '=R7_SUMMARIZE(D2; "short")', description: 'Суммаризация: short/medium/long.' }
    ];

    var state = {
        initialized: false,
        editorType: '',
        mode: 'disabled',
        version: '0.0.0',
        supportsAsync: false,
        supportsAddress: false,
        supportsSetCustomFunctions: false,
        hasCallCommand: false,
        registrationOk: false,
        registrationSource: 'none',
        registerInFlight: false,
        latestConfigHash: '',
        cacheEpoch: 1,
        trace: [],
        traceTimer: null
    };

    var ui = {
        button: null,
        panel: null,
        close: null,
        mode: null,
        status: null,
        formulaInput: null,
        functionList: null,
        insert: null,
        refreshSelected: null,
        refreshAll: null,
        clearSelected: null,
        clearAll: null,
        helperTrace: null,
        contextTrace: null
    };

    function tr(text) {
        if (window.Asc && window.Asc.plugin && typeof window.Asc.plugin.tr === 'function') {
            return window.Asc.plugin.tr(text);
        }
        return text;
    }

    function setStatus(message, isError) {
        if (!ui.status) return;
        ui.status.textContent = message || '';
        ui.status.classList.toggle('is-error', !!isError);
    }

    function withTimeout(run, timeoutMs, label) {
        return new Promise(function (resolve, reject) {
            var done = false;
            var timer = window.setTimeout(function () {
                if (done) return;
                done = true;
                reject(new Error((label || 'operation') + ' timeout'));
            }, timeoutMs || 2500);

            run(function (result) {
                if (done) return;
                done = true;
                window.clearTimeout(timer);
                resolve(result);
            }, function (error) {
                if (done) return;
                done = true;
                window.clearTimeout(timer);
                reject(error || new Error((label || 'operation') + ' failed'));
            });
        });
    }

    function executeMethod(name, args, timeoutMs) {
        return withTimeout(function (resolve, reject) {
            try {
                if (!window.Asc || !Asc.plugin || typeof Asc.plugin.executeMethod !== 'function') {
                    reject(new Error('executeMethod is unavailable'));
                    return;
                }
                Asc.plugin.executeMethod(name, args || [], function (result) {
                    resolve(result);
                });
            } catch (error) {
                reject(error);
            }
        }, timeoutMs || 3500, name);
    }

    function callCommand(command, scopeData, recalculate, timeoutMs) {
        return withTimeout(function (resolve, reject) {
            try {
                if (!window.Asc || !Asc.plugin || typeof Asc.plugin.callCommand !== 'function') {
                    reject(new Error('callCommand is unavailable'));
                    return;
                }
                if (scopeData && typeof scopeData === 'object') {
                    Object.keys(scopeData).forEach(function (key) {
                        Asc.scope[key] = scopeData[key];
                    });
                }
                Asc.plugin.callCommand(command, false, recalculate !== false, function (result) {
                    resolve(result);
                });
            } catch (error) {
                reject(error);
            }
        }, timeoutMs || 5000, 'callCommand');
    }

    function parseVersion(value) {
        var text = String(value || '0.0.0');
        var match = text.match(/(\d+)\.(\d+)(?:\.(\d+))?/);
        return {
            raw: text,
            major: match ? Number(match[1]) : 0,
            minor: match ? Number(match[2]) : 0,
            patch: match && match[3] ? Number(match[3]) : 0
        };
    }

    function versionAtLeast(parsed, major, minor, patch) {
        if (parsed.major > major) return true;
        if (parsed.major < major) return false;
        if (parsed.minor > minor) return true;
        if (parsed.minor < minor) return false;
        return parsed.patch >= patch;
    }

    function normalizeCustomFunctionsResult(rawResult) {
        var parsed = rawResult;
        if (typeof parsed === 'string') {
            try {
                var maybe = JSON.parse(parsed);
                if (Array.isArray(maybe)) parsed = maybe;
                else parsed = parsed.trim() ? [parsed] : [];
            } catch (e) {
                parsed = parsed.trim() ? [parsed] : [];
            }
        }
        if (!Array.isArray(parsed)) return { shape: 'array-string', entries: [] };
        if (!parsed.length) return { shape: 'array-string', entries: [] };
        if (parsed[0] && typeof parsed[0] === 'object') {
            return {
                shape: 'array-object',
                entries: parsed.map(function (item) {
                    return {
                        name: String(item && item.name ? item.name : ''),
                        value: String(item && item.value ? item.value : '')
                    };
                })
            };
        }
        return {
            shape: 'array-string',
            entries: parsed.map(function (item) { return String(item || ''); })
        };
    }

    function removeExistingLibrary(existingState) {
        if (existingState.shape === 'array-object') {
            return existingState.entries.filter(function (item) {
                return String(item.value || '').indexOf(LIBRARY_MARKER) === -1 && String(item.name || '') !== LIBRARY_ID;
            });
        }
        return existingState.entries.filter(function (item) {
            return String(item || '').indexOf(LIBRARY_MARKER) === -1;
        });
    }

    function buildConfig() {
        var settingsService = window.R7Chat && window.R7Chat.features && window.R7Chat.features.settings
            ? window.R7Chat.features.settings
            : null;
        var runtimeSettings = settingsService && typeof settingsService.loadSettings === 'function'
            ? settingsService.loadSettings()
            : {
                apiKey: (localStorage.getItem('apikey') || '').trim(),
                model: (localStorage.getItem('model') || 'openrouter/auto').trim()
            };
        var openrouterConfig = settingsService && typeof settingsService.getProviderConfig === 'function'
            ? settingsService.getProviderConfig(runtimeSettings, 'openrouter')
            : null;
        return {
            apiKey: openrouterConfig && openrouterConfig.apiKey ? openrouterConfig.apiKey : (runtimeSettings.apiKey || ''),
            model: openrouterConfig && openrouterConfig.model ? openrouterConfig.model : (runtimeSettings.model || 'openrouter/auto'),
            locale: (window.Asc && window.Asc.plugin && window.Asc.plugin.info && window.Asc.plugin.info.lang) ? window.Asc.plugin.info.lang : 'en-US',
            cacheEpoch: state.cacheEpoch,
            promptTemplateVersion: PROMPT_TEMPLATE_VERSION,
            libraryId: LIBRARY_ID
        };
    }

    function buildConfigHash(config) {
        return core.simpleHash([
            config.apiKey || '',
            config.model || '',
            config.locale || '',
            config.cacheEpoch || 1,
            config.promptTemplateVersion || ''
        ].join('|'));
    }

    function libraryRuntimeFactory(config) {
        var globalRef = (typeof globalThis !== 'undefined') ? globalThis : this;
        var SETTINGS = config || {};
        var STATUS = {
            INVALID_ARGS: '#R7.INVALID_ARGS',
            RATE_LIMIT: '#R7.RATE_LIMIT',
            TIMEOUT: '#R7.TIMEOUT',
            FAILED: '#R7.FAILED'
        };
        var EXTRACT_KINDS = {
            email: true,
            phone: true,
            company: true,
            name: true,
            date: true,
            amount: true,
            url: true,
            inn: true,
            iban: true
        };
        var SUMMARY_MODES = {
            short: true,
            medium: true,
            long: true
        };
        var CFG = {
            R7_ASK: { maxTokens: 180, maxChars: 4000 },
            R7_TRANSLATE: { maxTokens: 180, maxChars: 3000 },
            R7_EXTRACT: { maxTokens: 120, maxChars: 3000 },
            R7_CLASSIFY: { maxTokens: 90, maxChars: 2500 },
            R7_SUMMARIZE: { maxTokens: 220, maxChars: 6000 }
        };

        function normalizeText(value) {
            if (value === null || value === undefined) return '';
            if (Array.isArray(value)) {
                return value.map(function (row) {
                    if (Array.isArray(row)) return row.map(normalizeText).join('\t');
                    return normalizeText(row);
                }).join('\n').replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim();
            }
            return String(value).replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n').trim();
        }

        function normalizeEnum(value) {
            return normalizeText(value).toLowerCase();
        }

        function normalizeLabels(value) {
            var labels = normalizeText(value).split(',').map(function (x) { return normalizeText(x).toLowerCase(); }).filter(Boolean);
            var seen = {};
            var uniq = [];
            for (var i = 0; i < labels.length; i += 1) {
                if (seen[labels[i]]) continue;
                seen[labels[i]] = true;
                uniq.push(labels[i]);
            }
            return uniq;
        }

        function simpleHash(text) {
            var input = String(text || '');
            var hash = 2166136261;
            for (var i = 0; i < input.length; i += 1) {
                hash ^= input.charCodeAt(i);
                hash = Math.imul(hash, 16777619);
            }
            return (hash >>> 0).toString(16);
        }

        function redactSensitive(value) {
            var text = String(value || '');
            text = text.replace(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi, '[email]');
            text = text.replace(/\+?\d[\d\s\-()]{6,}\d/g, '[phone]');
            text = text.replace(/\b\d{12,19}\b/g, '[number]');
            return text;
        }

        function createRuntime() {
            var libraryId = String(SETTINGS.libraryId || 'R7AI_V1');
            return {
                libraryId: libraryId,
                model: SETTINGS.model || 'openrouter/auto',
                apiKey: SETTINGS.apiKey || '',
                locale: SETTINGS.locale || 'en-US',
                promptTemplateVersion: SETTINGS.promptTemplateVersion || 'v1',
                cacheEpoch: Number(SETTINGS.cacheEpoch || 1),
                cache: {},
                inflight: {},
                trace: [],
                bypassCells: {},
                queue: [],
                activeCount: 0
            };
        }

        var runtime = globalRef.__R7_CELL_AI_RUNTIME__;
        if (!runtime || runtime.cacheEpoch !== Number(SETTINGS.cacheEpoch || 1) || runtime.model !== SETTINGS.model || runtime.apiKey !== SETTINGS.apiKey) {
            runtime = createRuntime();
            globalRef.__R7_CELL_AI_RUNTIME__ = runtime;
        }

        runtime.clearAllCache = function () {
            runtime.cache = {};
            runtime.inflight = {};
            runtime.bypassCells = {};
        };
        runtime.clearTrace = function () {
            runtime.trace = [];
        };
        runtime.getTraceSnapshot = function () {
            return runtime.trace.slice(-120);
        };
        runtime.markBypassCells = function (addresses, ttlMs) {
            var ttl = Number(ttlMs || 30000);
            var now = Date.now();
            var list = Array.isArray(addresses) ? addresses : [];
            for (var i = 0; i < list.length; i += 1) {
                var key = String(list[i] || '').toUpperCase().replace(/\$/g, '');
                if (!key) continue;
                runtime.bypassCells[key] = now + ttl;
            }
        };

        function pushTrace(entry) {
            runtime.trace.push(entry);
            if (runtime.trace.length > 300) runtime.trace = runtime.trace.slice(runtime.trace.length - 300);
        }

        function shouldBypassCache(meta) {
            var cell = String(meta && meta.cell ? meta.cell : '').toUpperCase().replace(/\$/g, '');
            if (!cell) return false;
            var expires = runtime.bypassCells[cell];
            if (!expires) return false;
            if (expires < Date.now()) {
                delete runtime.bypassCells[cell];
                return false;
            }
            delete runtime.bypassCells[cell];
            return true;
        }

        function withTimeout(promise, timeoutMs) {
            return new Promise(function (resolve, reject) {
                var done = false;
                var timer = setTimeout(function () {
                    if (done) return;
                    done = true;
                    var timeoutError = new Error('timeout');
                    timeoutError.code = 'TIMEOUT';
                    reject(timeoutError);
                }, timeoutMs);

                promise.then(function (result) {
                    if (done) return;
                    done = true;
                    clearTimeout(timer);
                    resolve(result);
                }).catch(function (error) {
                    if (done) return;
                    done = true;
                    clearTimeout(timer);
                    reject(error);
                });
            });
        }

        function limitConcurrency(task) {
            return new Promise(function (resolve, reject) {
                function runTask() {
                    runtime.activeCount += 1;
                    Promise.resolve().then(task).then(resolve).catch(reject).finally(function () {
                        runtime.activeCount -= 1;
                        if (runtime.queue.length) {
                            var next = runtime.queue.shift();
                            next();
                        }
                    });
                }

                if (runtime.activeCount < 3) runTask();
                else runtime.queue.push(runTask);
            });
        }

        function buildPrompt(functionName, args) {
            if (functionName === 'R7_ASK') return 'Task: ' + args[0] + '\nInput: ' + args[1] + '\nReturn concise answer.';
            if (functionName === 'R7_TRANSLATE') return 'Translate text to ' + args[1] + '. Return translation only.\nText: ' + args[0];
            if (functionName === 'R7_EXTRACT') return 'Extract "' + args[0] + '" from text. Return only value or empty string.\nText: ' + args[1];
            if (functionName === 'R7_CLASSIFY') return 'Classify text into exactly one label from: ' + args[1].join(', ') + '. Return one label only.\nText: ' + args[0];
            if (functionName === 'R7_SUMMARIZE') return 'Summarize text in mode ' + args[1] + '. Keep concise and deterministic.\nText: ' + args[0];
            return 'Return concise answer.';
        }

        function normalizeArgs(functionName, rawArgs) {
            if (functionName === 'R7_ASK') {
                var prompt = normalizeText(rawArgs[0]);
                var value = normalizeText(rawArgs[1]);
                if (!prompt || !value) return { error: 'INVALID_ARGS' };
                return { args: [prompt, value] };
            }
            if (functionName === 'R7_TRANSLATE') {
                var text = normalizeText(rawArgs[0]);
                var target = normalizeText(rawArgs[1]).toUpperCase();
                if (!text || !target) return { error: 'INVALID_ARGS' };
                return { args: [text, target] };
            }
            if (functionName === 'R7_EXTRACT') {
                var kind = normalizeEnum(rawArgs[0]);
                var source = normalizeText(rawArgs[1]);
                if (!kind || !source || !EXTRACT_KINDS[kind]) return { error: 'INVALID_ARGS' };
                return { args: [kind, source] };
            }
            if (functionName === 'R7_CLASSIFY') {
                var clsText = normalizeText(rawArgs[0]);
                var labels = normalizeLabels(rawArgs[1]);
                if (!clsText || labels.length < 2) return { error: 'INVALID_ARGS' };
                return { args: [clsText, labels] };
            }
            if (functionName === 'R7_SUMMARIZE') {
                var sumText = normalizeText(rawArgs[0]);
                var mode = normalizeEnum(rawArgs[1]);
                if (!sumText || !mode || !SUMMARY_MODES[mode]) return { error: 'INVALID_ARGS' };
                return { args: [sumText, mode] };
            }
            return { error: 'INVALID_ARGS' };
        }

        function mapErrorToCode(error) {
            if (!error) return STATUS.FAILED;
            if (error.code === 'INVALID_ARGS') return STATUS.INVALID_ARGS;
            if (error.code === 'RATE_LIMIT') return STATUS.RATE_LIMIT;
            if (error.code === 'TIMEOUT') return STATUS.TIMEOUT;
            var status = Number(error.status || 0);
            if (status === 429) return STATUS.RATE_LIMIT;
            if (status === 408) return STATUS.TIMEOUT;
            return STATUS.FAILED;
        }

        function isTransient(error) {
            var status = Number(error && error.status ? error.status : 0);
            return status === 429 || status === 500 || status === 502 || status === 503 || status === 504;
        }

        function sanitizeResult(functionName, value, args) {
            var text = normalizeText(value || '');
            if (!text) return '';
            if (functionName === 'R7_CLASSIFY') {
                var first = normalizeEnum(text.split('\n')[0]);
                for (var i = 0; i < args[1].length; i += 1) {
                    if (first === args[1][i]) return args[1][i];
                }
                return args[1][0];
            }
            if (functionName === 'R7_EXTRACT') {
                return text.split('\n')[0].trim();
            }
            return text;
        }

        async function requestOpenRouter(prompt, maxTokens) {
            if (!runtime.apiKey) {
                var keyError = new Error('Missing API key');
                keyError.code = 'FAILED';
                throw keyError;
            }

            var response = await fetch('https://openrouter.ai/api/v1/chat/completions', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': 'Bearer ' + runtime.apiKey,
                    'X-Title': '{r7c}.ChatLLM (Cell Functions)'
                },
                body: JSON.stringify({
                    model: runtime.model || 'openrouter/auto',
                    messages: [
                        { role: 'system', content: 'You are deterministic spreadsheet AI operator. Return concise data output only.' },
                        { role: 'user', content: prompt }
                    ],
                    temperature: 0,
                    max_tokens: maxTokens,
                    stream: false
                })
            });

            if (!response.ok) {
                var bodyText = await response.text();
                var requestError = new Error(bodyText || 'Request failed');
                requestError.status = response.status;
                if (response.status === 429) requestError.code = 'RATE_LIMIT';
                throw requestError;
            }

            var payload = await response.json();
            var choice = payload && payload.choices && payload.choices[0] ? payload.choices[0] : null;
            var content = choice && choice.message ? choice.message.content : '';
            if (Array.isArray(content)) {
                content = content.map(function (part) {
                    if (typeof part === 'string') return part;
                    return part && part.text ? part.text : '';
                }).join('');
            }
            return {
                text: normalizeText(content || ''),
                usage: payload && payload.usage ? payload.usage : null
            };
        }

        async function executeFormula(functionName, rawArgs, meta) {
            var startedAt = Date.now();
            var normalized = normalizeArgs(functionName, rawArgs || []);
            var cell = String(meta && meta.cell ? meta.cell : '');

            if (normalized.error) {
                pushTrace({ time: startedAt, cell: cell, fn: functionName, status: 'invalid_arguments', cacheHit: false, inputHash: '', model: runtime.model, latencyMs: Date.now() - startedAt, errorCode: 'INVALID_ARGS' });
                return STATUS.INVALID_ARGS;
            }

            var args = normalized.args;
            var rawArgJson = JSON.stringify(args);
            var cfg = CFG[functionName] || { maxTokens: 160, maxChars: 3000 };
            if (rawArgJson.length > cfg.maxChars) {
                pushTrace({ time: startedAt, cell: cell, fn: functionName, status: 'invalid_arguments', cacheHit: false, inputHash: simpleHash(rawArgJson), model: runtime.model, latencyMs: Date.now() - startedAt, errorCode: 'INPUT_CAP' });
                return STATUS.INVALID_ARGS;
            }

            var cacheKey = JSON.stringify({
                fn: functionName,
                args: args,
                model: runtime.model,
                locale: runtime.locale,
                promptTemplateVersion: runtime.promptTemplateVersion,
                cacheEpoch: runtime.cacheEpoch
            });
            var bypassCache = shouldBypassCache(meta);
            if (!bypassCache && Object.prototype.hasOwnProperty.call(runtime.cache, cacheKey)) {
                var cached = runtime.cache[cacheKey];
                pushTrace({ time: startedAt, cell: cell, fn: functionName, status: 'cached', cacheHit: true, inputHash: simpleHash(rawArgJson), model: runtime.model, latencyMs: Date.now() - startedAt, tokenUsage: cached && cached.usage ? cached.usage : null });
                return cached.value;
            }
            if (!bypassCache && runtime.inflight[cacheKey]) {
                return runtime.inflight[cacheKey];
            }

            var prompt = redactSensitive(buildPrompt(functionName, args));
            var inputHash = simpleHash(rawArgJson);
            var task = limitConcurrency(async function () {
                var attempt = 0;
                while (true) {
                    try {
                        var result = await withTimeout(requestOpenRouter(prompt, cfg.maxTokens), 15000);
                        var value = sanitizeResult(functionName, result.text, args);
                        if (!bypassCache) runtime.cache[cacheKey] = { value: value, usage: result.usage || null, ts: Date.now() };
                        pushTrace({ time: startedAt, cell: cell, fn: functionName, status: 'ok', cacheHit: false, inputHash: inputHash, model: runtime.model, latencyMs: Date.now() - startedAt, tokenUsage: result.usage || null });
                        return value;
                    } catch (error) {
                        if (error && error.code === 'TIMEOUT') {
                            pushTrace({ time: startedAt, cell: cell, fn: functionName, status: 'timeout', cacheHit: false, inputHash: inputHash, model: runtime.model, latencyMs: Date.now() - startedAt, errorCode: 'TIMEOUT' });
                            return STATUS.TIMEOUT;
                        }
                        if (attempt < 1 && isTransient(error)) {
                            attempt += 1;
                            continue;
                        }
                        var code = mapErrorToCode(error);
                        pushTrace({ time: startedAt, cell: cell, fn: functionName, status: code === STATUS.RATE_LIMIT ? 'rate_limited' : 'failed', cacheHit: false, inputHash: inputHash, model: runtime.model, latencyMs: Date.now() - startedAt, errorCode: code });
                        return code;
                    }
                }
            });

            runtime.inflight[cacheKey] = task.finally(function () { delete runtime.inflight[cacheKey]; });
            return runtime.inflight[cacheKey];
        }

        function contextMeta(ctx) {
            var cellAddress = '';
            if (ctx && ctx.address) cellAddress = String(ctx.address);
            if (!cellAddress && ctx && ctx.args && Array.isArray(ctx.args) && ctx.args[0] && ctx.args[0].address) {
                cellAddress = String(ctx.args[0].address);
            }
            return { cell: cellAddress };
        }

        async function R7_ASK(prompt, value) { return executeFormula('R7_ASK', [prompt, value], contextMeta(this)); }
        async function R7_TRANSLATE(text, targetLang) { return executeFormula('R7_TRANSLATE', [text, targetLang], contextMeta(this)); }
        async function R7_EXTRACT(kind, text) { return executeFormula('R7_EXTRACT', [kind, text], contextMeta(this)); }
        async function R7_CLASSIFY(text, labels) { return executeFormula('R7_CLASSIFY', [text, labels], contextMeta(this)); }
        async function R7_SUMMARIZE(text, mode) { return executeFormula('R7_SUMMARIZE', [text, mode], contextMeta(this)); }

        try { Api.RemoveCustomFunction('R7_ASK'); } catch (e1) {}
        try { Api.RemoveCustomFunction('R7_TRANSLATE'); } catch (e2) {}
        try { Api.RemoveCustomFunction('R7_EXTRACT'); } catch (e3) {}
        try { Api.RemoveCustomFunction('R7_CLASSIFY'); } catch (e4) {}
        try { Api.RemoveCustomFunction('R7_SUMMARIZE'); } catch (e5) {}

        Api.AddCustomFunctionLibrary(String(SETTINGS.libraryId || 'R7AI_V1'), function () {
            Api.AddCustomFunction(R7_ASK);
            Api.AddCustomFunction(R7_TRANSLATE);
            Api.AddCustomFunction(R7_EXTRACT);
            Api.AddCustomFunction(R7_CLASSIFY);
            Api.AddCustomFunction(R7_SUMMARIZE);
        });
    }

    function buildLibraryCode(config) {
        return LIBRARY_MARKER + '\n(' + libraryRuntimeFactory.toString() + ')(' + JSON.stringify(config) + ');';
    }

    async function probeSetCustomFunctions() {
        try {
            var currentRaw = await executeMethod('GetCustomFunctions', [], 2200);
            var current = normalizeCustomFunctionsResult(currentRaw);
            await executeMethod('SetCustomFunctions', [JSON.stringify(current.entries)], 2200);
            return true;
        } catch (error) {
            return false;
        }
    }

    async function recalculateAllFormulas() {
        try {
            await executeMethod('RecalculateAllFormulas', [], 3500);
        } catch (error) {
            // no-op
        }
    }

    async function detectCapabilities() {
        state.editorType = (window.Asc && window.Asc.plugin && window.Asc.plugin.info && window.Asc.plugin.info.editorType) ? window.Asc.plugin.info.editorType : '';
        state.hasCallCommand = !!(window.Asc && window.Asc.plugin && typeof window.Asc.plugin.callCommand === 'function');
        try {
            state.version = String(await executeMethod('GetVersion', [], 2200) || '0.0.0');
        } catch (error) {
            state.version = '0.0.0';
        }
        var parsed = parseVersion(state.version);
        state.supportsAsync = versionAtLeast(parsed, 9, 0, 0);
        state.supportsAddress = versionAtLeast(parsed, 9, 0, 4);
        state.supportsSetCustomFunctions = await probeSetCustomFunctions();

        if (state.editorType !== 'cell') state.mode = 'disabled';
        else if (state.supportsAsync && state.supportsSetCustomFunctions) state.mode = 'native_async';
        else if (state.supportsAsync && state.hasCallCommand) state.mode = 'macro_register_fallback';
        else state.mode = 'bulk_only_fallback';
    }

    async function registerWithSetCustomFunctions(config) {
        var libraryCode = buildLibraryCode(config);
        await executeMethod('SetCustomFunctions', [JSON.stringify([libraryCode])], 5000);
        var verifyRaw = await executeMethod('GetCustomFunctions', [], 2600);
        var verify = normalizeCustomFunctionsResult(verifyRaw);
        var verifyText = JSON.stringify(verify.entries || []);
        if (verifyText.indexOf('R7_ASK') === -1) {
            throw new Error('native SetCustomFunctions applied, but R7_ASK is not visible in GetCustomFunctions');
        }
        await recalculateAllFormulas();
        state.registrationSource = 'SetCustomFunctions';
        state.registrationOk = true;
    }

    async function registerWithCallCommand(config) {
        var result = await callCommand(function () {
            try {
                var SETTINGS = Asc.scope.r7CellConfig || {};
                var STATUS = {
                    INVALID_ARGS: '#R7.INVALID_ARGS',
                    RATE_LIMIT: '#R7.RATE_LIMIT',
                    TIMEOUT: '#R7.TIMEOUT',
                    FAILED: '#R7.FAILED'
                };
                var EXTRACT_KINDS = { email: true, phone: true, company: true, name: true, date: true, amount: true, url: true, inn: true, iban: true };
                var SUMMARY_MODES = { short: true, medium: true, long: true };
                var CFG = {
                    R7_ASK: { maxTokens: 180, maxChars: 4000 },
                    R7_TRANSLATE: { maxTokens: 180, maxChars: 3000 },
                    R7_EXTRACT: { maxTokens: 120, maxChars: 3000 },
                    R7_CLASSIFY: { maxTokens: 90, maxChars: 2500 },
                    R7_SUMMARIZE: { maxTokens: 220, maxChars: 6000 }
                };

                function normalizeText(value) {
                    if (value === null || value === undefined) return '';
                    if (Array.isArray(value)) {
                        return value.map(function (row) {
                            if (Array.isArray(row)) return row.map(normalizeText).join('\t');
                            return normalizeText(row);
                        }).join('\n').replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim();
                    }
                    return String(value).replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n').trim();
                }
                function normalizeEnum(value) {
                    return normalizeText(value).toLowerCase();
                }
                function normalizeLabels(value) {
                    var labels = normalizeText(value).split(',').map(function (x) { return normalizeText(x).toLowerCase(); }).filter(Boolean);
                    var seen = {};
                    var uniq = [];
                    for (var i = 0; i < labels.length; i += 1) {
                        if (seen[labels[i]]) continue;
                        seen[labels[i]] = true;
                        uniq.push(labels[i]);
                    }
                    return uniq;
                }
                function simpleHash(input) {
                    var text = String(input || '');
                    var hash = 2166136261;
                    for (var i = 0; i < text.length; i += 1) {
                        hash ^= text.charCodeAt(i);
                        hash = Math.imul(hash, 16777619);
                    }
                    return (hash >>> 0).toString(16);
                }
                function redactSensitive(value) {
                    var text = String(value || '');
                    text = text.replace(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi, '[email]');
                    text = text.replace(/\+?\d[\d\s\-()]{6,}\d/g, '[phone]');
                    text = text.replace(/\b\d{12,19}\b/g, '[number]');
                    return text;
                }
                function createRuntime() {
                    return {
                        libraryId: String(SETTINGS.libraryId || 'R7AI_V1'),
                        model: String(SETTINGS.model || 'openrouter/auto'),
                        apiKey: String(SETTINGS.apiKey || ''),
                        locale: String(SETTINGS.locale || 'en-US'),
                        promptTemplateVersion: String(SETTINGS.promptTemplateVersion || 'v1'),
                        cacheEpoch: Number(SETTINGS.cacheEpoch || 1),
                        cache: {},
                        inflight: {},
                        trace: [],
                        bypassCells: {},
                        queue: [],
                        activeCount: 0
                    };
                }
                var globalRef = (typeof globalThis !== 'undefined') ? globalThis : this;
                var runtime = globalRef.__R7_CELL_AI_RUNTIME__;
                if (!runtime || runtime.cacheEpoch !== Number(SETTINGS.cacheEpoch || 1) || runtime.model !== SETTINGS.model || runtime.apiKey !== SETTINGS.apiKey) {
                    runtime = createRuntime();
                    globalRef.__R7_CELL_AI_RUNTIME__ = runtime;
                }
                runtime.clearAllCache = function () {
                    runtime.cache = {};
                    runtime.inflight = {};
                    runtime.bypassCells = {};
                };
                runtime.clearTrace = function () {
                    runtime.trace = [];
                };
                runtime.getTraceSnapshot = function () {
                    return runtime.trace.slice(-120);
                };
                runtime.markBypassCells = function (addresses, ttlMs) {
                    var ttl = Number(ttlMs || 30000);
                    var now = Date.now();
                    var list = Array.isArray(addresses) ? addresses : [];
                    for (var i = 0; i < list.length; i += 1) {
                        var key = String(list[i] || '').toUpperCase().replace(/\$/g, '');
                        if (!key) continue;
                        runtime.bypassCells[key] = now + ttl;
                    }
                };
                function pushTrace(entry) {
                    runtime.trace.push(entry);
                    if (runtime.trace.length > 300) runtime.trace = runtime.trace.slice(runtime.trace.length - 300);
                }
                function shouldBypassCache(meta) {
                    var cell = String(meta && meta.cell ? meta.cell : '').toUpperCase().replace(/\$/g, '');
                    if (!cell) return false;
                    var expires = runtime.bypassCells[cell];
                    if (!expires) return false;
                    if (expires < Date.now()) {
                        delete runtime.bypassCells[cell];
                        return false;
                    }
                    delete runtime.bypassCells[cell];
                    return true;
                }
                function withTimeout(promise, timeoutMs) {
                    return new Promise(function (resolve, reject) {
                        var done = false;
                        var timer = setTimeout(function () {
                            if (done) return;
                            done = true;
                            var timeoutError = new Error('timeout');
                            timeoutError.code = 'TIMEOUT';
                            reject(timeoutError);
                        }, timeoutMs);
                        promise.then(function (result) {
                            if (done) return;
                            done = true;
                            clearTimeout(timer);
                            resolve(result);
                        }).catch(function (error) {
                            if (done) return;
                            done = true;
                            clearTimeout(timer);
                            reject(error);
                        });
                    });
                }
                function limitConcurrency(task) {
                    return new Promise(function (resolve, reject) {
                        function runTask() {
                            runtime.activeCount += 1;
                            Promise.resolve().then(task).then(resolve).catch(reject).finally(function () {
                                runtime.activeCount -= 1;
                                if (runtime.queue.length) {
                                    var next = runtime.queue.shift();
                                    next();
                                }
                            });
                        }
                        if (runtime.activeCount < 3) runTask();
                        else runtime.queue.push(runTask);
                    });
                }
                function buildPrompt(functionName, args) {
                    if (functionName === 'R7_ASK') return 'Task: ' + args[0] + '\nInput: ' + args[1] + '\nReturn concise answer.';
                    if (functionName === 'R7_TRANSLATE') return 'Translate text to ' + args[1] + '. Return translation only.\nText: ' + args[0];
                    if (functionName === 'R7_EXTRACT') return 'Extract "' + args[0] + '" from text. Return only value or empty string.\nText: ' + args[1];
                    if (functionName === 'R7_CLASSIFY') return 'Classify text into exactly one label from: ' + args[1].join(', ') + '. Return one label only.\nText: ' + args[0];
                    if (functionName === 'R7_SUMMARIZE') return 'Summarize text in mode ' + args[1] + '. Keep concise and deterministic.\nText: ' + args[0];
                    return 'Return concise answer.';
                }
                function normalizeArgs(functionName, rawArgs) {
                    if (functionName === 'R7_ASK') {
                        var prompt = normalizeText(rawArgs[0]);
                        var value = normalizeText(rawArgs[1]);
                        if (!prompt || !value) return { error: 'INVALID_ARGS' };
                        return { args: [prompt, value] };
                    }
                    if (functionName === 'R7_TRANSLATE') {
                        var text = normalizeText(rawArgs[0]);
                        var target = normalizeText(rawArgs[1]).toUpperCase();
                        if (!text || !target) return { error: 'INVALID_ARGS' };
                        return { args: [text, target] };
                    }
                    if (functionName === 'R7_EXTRACT') {
                        var kind = normalizeEnum(rawArgs[0]);
                        var source = normalizeText(rawArgs[1]);
                        if (!kind || !source || !EXTRACT_KINDS[kind]) return { error: 'INVALID_ARGS' };
                        return { args: [kind, source] };
                    }
                    if (functionName === 'R7_CLASSIFY') {
                        var clsText = normalizeText(rawArgs[0]);
                        var labels = normalizeLabels(rawArgs[1]);
                        if (!clsText || labels.length < 2) return { error: 'INVALID_ARGS' };
                        return { args: [clsText, labels] };
                    }
                    if (functionName === 'R7_SUMMARIZE') {
                        var sumText = normalizeText(rawArgs[0]);
                        var mode = normalizeEnum(rawArgs[1]);
                        if (!sumText || !mode || !SUMMARY_MODES[mode]) return { error: 'INVALID_ARGS' };
                        return { args: [sumText, mode] };
                    }
                    return { error: 'INVALID_ARGS' };
                }
                function mapErrorToCode(error) {
                    if (!error) return STATUS.FAILED;
                    if (error.code === 'INVALID_ARGS') return STATUS.INVALID_ARGS;
                    if (error.code === 'RATE_LIMIT') return STATUS.RATE_LIMIT;
                    if (error.code === 'TIMEOUT') return STATUS.TIMEOUT;
                    var status = Number(error.status || 0);
                    if (status === 429) return STATUS.RATE_LIMIT;
                    if (status === 408) return STATUS.TIMEOUT;
                    return STATUS.FAILED;
                }
                function isTransient(error) {
                    var status = Number(error && error.status ? error.status : 0);
                    return status === 429 || status === 500 || status === 502 || status === 503 || status === 504;
                }
                function sanitizeResult(functionName, value, args) {
                    var text = normalizeText(value || '');
                    if (!text) return '';
                    if (functionName === 'R7_CLASSIFY') {
                        var first = normalizeEnum(text.split('\n')[0]);
                        for (var i = 0; i < args[1].length; i += 1) {
                            if (first === args[1][i]) return args[1][i];
                        }
                        return args[1][0];
                    }
                    if (functionName === 'R7_EXTRACT') return text.split('\n')[0].trim();
                    return text;
                }
                async function requestOpenRouter(prompt, maxTokens) {
                    if (!runtime.apiKey) {
                        var keyError = new Error('Missing API key');
                        keyError.code = 'FAILED';
                        throw keyError;
                    }
                    var response = await fetch('https://openrouter.ai/api/v1/chat/completions', {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json',
                            'Authorization': 'Bearer ' + runtime.apiKey,
                            'X-Title': '{r7c}.ChatLLM (Cell Functions)'
                        },
                        body: JSON.stringify({
                            model: runtime.model || 'openrouter/auto',
                            messages: [
                                { role: 'system', content: 'You are deterministic spreadsheet AI operator. Return concise data output only.' },
                                { role: 'user', content: prompt }
                            ],
                            temperature: 0,
                            max_tokens: maxTokens,
                            stream: false
                        })
                    });
                    if (!response.ok) {
                        var bodyText = await response.text();
                        var requestError = new Error(bodyText || 'Request failed');
                        requestError.status = response.status;
                        if (response.status === 429) requestError.code = 'RATE_LIMIT';
                        throw requestError;
                    }
                    var payload = await response.json();
                    var choice = payload && payload.choices && payload.choices[0] ? payload.choices[0] : null;
                    var content = choice && choice.message ? choice.message.content : '';
                    if (Array.isArray(content)) {
                        content = content.map(function (part) {
                            if (typeof part === 'string') return part;
                            return part && part.text ? part.text : '';
                        }).join('');
                    }
                    return { text: normalizeText(content || ''), usage: payload && payload.usage ? payload.usage : null };
                }
                async function executeFormula(functionName, rawArgs, meta) {
                    var startedAt = Date.now();
                    var normalized = normalizeArgs(functionName, rawArgs || []);
                    var cell = String(meta && meta.cell ? meta.cell : '');
                    if (normalized.error) {
                        pushTrace({ time: startedAt, cell: cell, fn: functionName, status: 'invalid_arguments', cacheHit: false, inputHash: '', model: runtime.model, latencyMs: Date.now() - startedAt, errorCode: 'INVALID_ARGS' });
                        return STATUS.INVALID_ARGS;
                    }
                    var args = normalized.args;
                    var rawArgJson = JSON.stringify(args);
                    var fnCfg = CFG[functionName] || { maxTokens: 160, maxChars: 3000 };
                    if (rawArgJson.length > fnCfg.maxChars) {
                        pushTrace({ time: startedAt, cell: cell, fn: functionName, status: 'invalid_arguments', cacheHit: false, inputHash: simpleHash(rawArgJson), model: runtime.model, latencyMs: Date.now() - startedAt, errorCode: 'INPUT_CAP' });
                        return STATUS.INVALID_ARGS;
                    }
                    var cacheKey = JSON.stringify({
                        fn: functionName,
                        args: args,
                        model: runtime.model,
                        locale: runtime.locale,
                        promptTemplateVersion: runtime.promptTemplateVersion,
                        cacheEpoch: runtime.cacheEpoch
                    });
                    var bypass = shouldBypassCache(meta);
                    if (!bypass && Object.prototype.hasOwnProperty.call(runtime.cache, cacheKey)) {
                        var cached = runtime.cache[cacheKey];
                        pushTrace({ time: startedAt, cell: cell, fn: functionName, status: 'cached', cacheHit: true, inputHash: simpleHash(rawArgJson), model: runtime.model, latencyMs: Date.now() - startedAt, tokenUsage: cached && cached.usage ? cached.usage : null });
                        return cached.value;
                    }
                    if (!bypass && runtime.inflight[cacheKey]) return runtime.inflight[cacheKey];
                    var prompt = redactSensitive(buildPrompt(functionName, args));
                    var inputHash = simpleHash(rawArgJson);
                    var task = limitConcurrency(async function () {
                        var attempt = 0;
                        while (true) {
                            try {
                                var result = await withTimeout(requestOpenRouter(prompt, fnCfg.maxTokens), 15000);
                                var value = sanitizeResult(functionName, result.text, args);
                                if (!bypass) runtime.cache[cacheKey] = { value: value, usage: result.usage || null, ts: Date.now() };
                                pushTrace({ time: startedAt, cell: cell, fn: functionName, status: 'ok', cacheHit: false, inputHash: inputHash, model: runtime.model, latencyMs: Date.now() - startedAt, tokenUsage: result.usage || null });
                                return value;
                            } catch (error) {
                                if (error && error.code === 'TIMEOUT') {
                                    pushTrace({ time: startedAt, cell: cell, fn: functionName, status: 'timeout', cacheHit: false, inputHash: inputHash, model: runtime.model, latencyMs: Date.now() - startedAt, errorCode: 'TIMEOUT' });
                                    return STATUS.TIMEOUT;
                                }
                                if (attempt < 1 && isTransient(error)) {
                                    attempt += 1;
                                    continue;
                                }
                                var code = mapErrorToCode(error);
                                pushTrace({ time: startedAt, cell: cell, fn: functionName, status: code === STATUS.RATE_LIMIT ? 'rate_limited' : 'failed', cacheHit: false, inputHash: inputHash, model: runtime.model, latencyMs: Date.now() - startedAt, errorCode: code });
                                return code;
                            }
                        }
                    });
                    runtime.inflight[cacheKey] = task.finally(function () { delete runtime.inflight[cacheKey]; });
                    return runtime.inflight[cacheKey];
                }
                function contextMeta(ctx) {
                    var cellAddress = '';
                    if (ctx && ctx.address) cellAddress = String(ctx.address);
                    if (!cellAddress && ctx && ctx.args && Array.isArray(ctx.args) && ctx.args[0] && ctx.args[0].address) {
                        cellAddress = String(ctx.args[0].address);
                    }
                    return { cell: cellAddress };
                }
                async function R7_ASK(prompt, value) { return executeFormula('R7_ASK', [prompt, value], contextMeta(this)); }
                async function R7_TRANSLATE(text, targetLang) { return executeFormula('R7_TRANSLATE', [text, targetLang], contextMeta(this)); }
                async function R7_EXTRACT(kind, text) { return executeFormula('R7_EXTRACT', [kind, text], contextMeta(this)); }
                async function R7_CLASSIFY(text, labels) { return executeFormula('R7_CLASSIFY', [text, labels], contextMeta(this)); }
                async function R7_SUMMARIZE(text, mode) { return executeFormula('R7_SUMMARIZE', [text, mode], contextMeta(this)); }
                
                try { Api.RemoveCustomFunction('R7_ASK'); } catch (e1) {}
                try { Api.RemoveCustomFunction('R7_TRANSLATE'); } catch (e2) {}
                try { Api.RemoveCustomFunction('R7_EXTRACT'); } catch (e3) {}
                try { Api.RemoveCustomFunction('R7_CLASSIFY'); } catch (e4) {}
                try { Api.RemoveCustomFunction('R7_SUMMARIZE'); } catch (e5) {}
                
                Api.AddCustomFunctionLibrary(String(SETTINGS.libraryId || 'R7AI_V1'), function () {
                    Api.AddCustomFunction(R7_ASK);
                    Api.AddCustomFunction(R7_TRANSLATE);
                    Api.AddCustomFunction(R7_EXTRACT);
                    Api.AddCustomFunction(R7_CLASSIFY);
                    Api.AddCustomFunction(R7_SUMMARIZE);
                });

                Api.RecalculateAllFormulas();
                return { ok: true };
            } catch (error) {
                return { ok: false, error: String(error && error.message ? error.message : error) };
            }
        }, { r7CellConfig: config }, true, 9000);

        if (!result || !result.ok) {
            throw new Error(result && result.error ? result.error : 'callCommand registration failed');
        }
        state.registrationSource = 'callCommand';
        state.registrationOk = true;
    }

    function updateModePill() {
        if (!ui.mode) return;
        ui.mode.textContent = state.mode + ' | v' + state.version;
        ui.mode.classList.toggle('is-error', state.mode === 'bulk_only_fallback' || state.mode === 'disabled');
    }

    function normalizeFormulaAlias(formula) {
        var value = String(formula || '').trim();
        if (!value) return '';
        if (value.charAt(0) === '=') value = value.slice(1).trim();
        value = value.replace(/^R7\./i, 'R7_');
        value = value.replace(/^R7\.(ASK|TRANSLATE|EXTRACT|CLASSIFY|SUMMARIZE)\b/i, function (_, fn) { return 'R7_' + fn.toUpperCase(); });
        if (!/^R7_(ASK|TRANSLATE|EXTRACT|CLASSIFY|SUMMARIZE)\b/i.test(value)) return '';
        return '=' + value;
    }

    async function registerCellFunctions(force) {
        if (state.registerInFlight) return;
        if (state.mode === 'disabled' || state.mode === 'bulk_only_fallback') return;

        var config = buildConfig();
        if (!config.apiKey) {
            state.registrationOk = false;
            setStatus(tr('Для AI-функций в ячейках укажите OpenRouter API Key в настройках.'), true);
            return;
        }

        var configHash = buildConfigHash(config);
        if (!force && state.registrationOk && state.latestConfigHash === configHash) return;

        state.registerInFlight = true;
        setStatus(tr('Регистрация AI-функций...'), false);

        try {
            if (state.mode === 'native_async') {
                try {
                    await registerWithSetCustomFunctions(config);
                } catch (nativeError) {
                    if (!state.hasCallCommand) throw nativeError;
                    state.mode = 'macro_register_fallback';
                    try {
                        await registerWithCallCommand(config);
                    } catch (fallbackError) {
                        throw new Error('native failed: ' + String(nativeError && nativeError.message ? nativeError.message : nativeError) + ' | fallback failed: ' + String(fallbackError && fallbackError.message ? fallbackError.message : fallbackError));
                    }
                }
            } else {
                await registerWithCallCommand(config);
            }

            state.latestConfigHash = configHash;
            updateModePill();
            setStatus(tr('AI-функции зарегистрированы: R7_ASK/R7_TRANSLATE/R7_EXTRACT/R7_CLASSIFY/R7_SUMMARIZE'), false);
        } catch (error) {
            state.registrationOk = false;
            setStatus(tr('Не удалось зарегистрировать AI-функции: ') + String(error && error.message ? error.message : error), true);
        } finally {
            state.registerInFlight = false;
        }
    }

    async function insertFormula(formula) {
        var normalized = normalizeFormulaAlias(formula);
        if (!normalized) {
            setStatus(tr('Формула должна начинаться с R7_... или R7....'), true);
            return;
        }
        try {
            await callCommand(function () {
                var ws = Api.GetActiveSheet();
                var selection = ws ? ws.GetSelection() : null;
                if (!selection) return { ok: false, error: 'No selection' };
                selection.SetValue(Asc.scope.r7Formula);
                Api.RecalculateAllFormulas();
                return { ok: true };
            }, { r7Formula: normalized }, true, 5000);
            setStatus(tr('Формула вставлена: ') + normalized, false);
        } catch (error) {
            setStatus(tr('Ошибка вставки формулы: ') + String(error && error.message ? error.message : error), true);
        }
    }

    async function refreshAll() {
        await recalculateAllFormulas();
        setStatus(tr('Refresh all выполнен.'), false);
        pullTrace();
    }

    async function refreshSelected(markBypass) {
        try {
            var result = await callCommand(function () {
                function colToName(col) {
                    var n = Number(col || 1);
                    var s = '';
                    while (n > 0) {
                        var m = (n - 1) % 26;
                        s = String.fromCharCode(65 + m) + s;
                        n = Math.floor((n - m) / 26);
                    }
                    return s;
                }
                function colToNumber(name) {
                    var result = 0;
                    var letters = String(name || '').toUpperCase();
                    for (var i = 0; i < letters.length; i += 1) {
                        var code = letters.charCodeAt(i);
                        if (code < 65 || code > 90) continue;
                        result = result * 26 + (code - 64);
                    }
                    return result;
                }
                function parseAddress(address) {
                    var text = String(address || '').replace(/\$/g, '');
                    var idx = text.lastIndexOf('!');
                    if (idx !== -1) text = text.slice(idx + 1);
                    var parts = text.split(':');
                    if (parts.length === 1) parts.push(parts[0]);
                    var left = parts[0].match(/^([A-Za-z]+)(\d+)$/);
                    var right = parts[1].match(/^([A-Za-z]+)(\d+)$/);
                    if (!left || !right) return null;
                    var sc = colToNumber(left[1]);
                    var ec = colToNumber(right[1]);
                    var sr = Number(left[2]);
                    var er = Number(right[2]);
                    return {
                        startCol: Math.min(sc, ec),
                        endCol: Math.max(sc, ec),
                        startRow: Math.min(sr, er),
                        endRow: Math.max(sr, er)
                    };
                }

                var ws = Api.GetActiveSheet();
                var selection = ws ? ws.GetSelection() : null;
                if (!selection) return { ok: false, error: 'No selection' };
                var parsed = parseAddress(selection.GetAddress(true, true, 'xlA1', false));
                if (!parsed) return { ok: false, error: 'Unsupported selection' };

                var refreshed = 0;
                var bypassCells = [];
                for (var row = parsed.startRow; row <= parsed.endRow; row += 1) {
                    for (var col = parsed.startCol; col <= parsed.endCol; col += 1) {
                        var address = colToName(col) + row;
                        var cell = ws.GetRange(address);
                        if (!cell || typeof cell.GetFormula !== 'function') continue;
                        var formula = cell.GetFormula();
                        if (!formula) continue;
                        var normalizedFormula = String(formula);
                        if (normalizedFormula.charAt(0) !== '=') normalizedFormula = '=' + normalizedFormula;
                        if (!/R7_(ASK|TRANSLATE|EXTRACT|CLASSIFY|SUMMARIZE)/i.test(normalizedFormula)) continue;
                        cell.SetValue(normalizedFormula);
                        refreshed += 1;
                        if (bypassCells.length < 500) bypassCells.push(address);
                    }
                }

                if (Asc.scope.markBypass) {
                    var globalRef = (typeof globalThis !== 'undefined') ? globalThis : this;
                    var runtime = globalRef.__R7_CELL_AI_RUNTIME__;
                    if (runtime && typeof runtime.markBypassCells === 'function') {
                        runtime.markBypassCells(bypassCells, 60000);
                    }
                }

                Api.RecalculateAllFormulas();
                return { ok: true, refreshed: refreshed };
            }, { markBypass: !!markBypass }, true, 8000);

            if (!result || !result.ok) {
                throw new Error(result && result.error ? result.error : 'Unknown refresh selected error');
            }
            setStatus(tr('Refresh selected: обновлено формул: ') + String(result.refreshed || 0), false);
            pullTrace();
        } catch (error) {
            setStatus(tr('Ошибка refresh selected: ') + String(error && error.message ? error.message : error), true);
        }
    }

    async function clearAllCache() {
        state.cacheEpoch += 1;
        try {
            await callCommand(function () {
                var globalRef = (typeof globalThis !== 'undefined') ? globalThis : this;
                var runtime = globalRef.__R7_CELL_AI_RUNTIME__;
                if (runtime && typeof runtime.clearAllCache === 'function') runtime.clearAllCache();
                if (runtime && typeof runtime.clearTrace === 'function') runtime.clearTrace();
                return { ok: true };
            }, null, false, 2500);
        } catch (error) {
            // ignore and continue by re-registering
        }
        await registerCellFunctions(true);
        await refreshAll();
        setStatus(tr('Кэш AI-функций очищен для всей книги.'), false);
    }

    async function clearSelectedCache() {
        await refreshSelected(true);
        setStatus(tr('Кэш выбранных AI-ячеек сброшен (bypass + refresh).'), false);
    }

    function renderFunctions() {
        if (!ui.functionList) return;
        var fragment = document.createDocumentFragment();
        FUNCTION_DEFS.forEach(function (def) {
            var item = document.createElement('div');
            item.className = 'cell-ai-function-item';

            var top = document.createElement('div');
            top.className = 'cell-ai-function-top';

            var title = document.createElement('div');
            title.className = 'cell-ai-function-name';
            title.textContent = def.signature;

            var insertButton = document.createElement('button');
            insertButton.type = 'button';
            insertButton.className = 'sheet-action-btn';
            insertButton.textContent = tr('Insert');
            insertButton.addEventListener('click', function () {
                insertFormula(def.example);
            });

            top.appendChild(title);
            top.appendChild(insertButton);

            var desc = document.createElement('div');
            desc.className = 'cell-ai-function-desc';
            desc.textContent = def.description;

            var example = document.createElement('div');
            example.className = 'cell-ai-function-example';
            example.textContent = def.example + '    (' + def.alias + ')';

            item.appendChild(top);
            item.appendChild(desc);
            item.appendChild(example);
            fragment.appendChild(item);
        });
        ui.functionList.innerHTML = '';
        ui.functionList.appendChild(fragment);
    }

    function renderTrace(target, records) {
        if (!target) return;
        if (!records || !records.length) {
            target.innerHTML = '<div class="context-empty">' + tr('Trace пока пустой.') + '</div>';
            return;
        }
        var fragment = document.createDocumentFragment();
        records.slice(-25).reverse().forEach(function (record) {
            var item = document.createElement('div');
            item.className = 'cell-ai-trace-item';

            var top = document.createElement('div');
            top.className = 'cell-ai-trace-top';
            top.textContent = (record.fn || 'unknown') + ' · ' + (record.status || 'n/a');

            var meta = document.createElement('div');
            meta.className = 'cell-ai-trace-meta';
            var cell = record.cell ? ('cell=' + record.cell + ' · ') : '';
            var cache = record.cacheHit ? 'cache=hit' : 'cache=miss';
            var latency = record.latencyMs !== undefined ? (' · ' + record.latencyMs + 'ms') : '';
            meta.textContent = cell + cache + latency;

            item.appendChild(top);
            item.appendChild(meta);
            fragment.appendChild(item);
        });
        target.innerHTML = '';
        target.appendChild(fragment);
    }

    async function pullTrace() {
        if (state.mode === 'disabled' || state.mode === 'bulk_only_fallback') return;
        try {
            var trace = await callCommand(function () {
                var globalRef = (typeof globalThis !== 'undefined') ? globalThis : this;
                var runtime = globalRef.__R7_CELL_AI_RUNTIME__;
                if (!runtime || typeof runtime.getTraceSnapshot !== 'function') return [];
                return runtime.getTraceSnapshot();
            }, null, false, 2500);

            state.trace = Array.isArray(trace) ? trace : [];
            renderTrace(ui.helperTrace, state.trace);
            renderTrace(ui.contextTrace, state.trace);
        } catch (error) {
            // no-op
        }
    }

    function openPanel() {
        if (!ui.panel || state.editorType !== 'cell') return;
        ui.panel.style.display = 'flex';
        if (ui.button) ui.button.setAttribute('aria-expanded', 'true');
        pullTrace();
    }

    function closePanel() {
        if (!ui.panel) return;
        ui.panel.style.display = 'none';
        if (ui.button) ui.button.setAttribute('aria-expanded', 'false');
    }

    function togglePanel() {
        if (!ui.panel) return;
        if (ui.panel.style.display === 'none' || !ui.panel.style.display) openPanel();
        else closePanel();
    }

    function applyTranslations() {
        var pairs = [
            ['cellAiPanelTitle', 'Cell AI Functions'],
            ['cellAiMode', 'mode'],
            ['cellAiHint', 'Используйте канонические формулы R7_* (helper принимает alias R7.*).'],
            ['cellAiInsertLabel', 'Formula'],
            ['cellAiInsert', 'Insert Formula'],
            ['cellAiRefreshSelected', 'Refresh selected AI formulas'],
            ['cellAiRefreshAll', 'Refresh all AI formulas'],
            ['cellAiClearSelected', 'Clear selected AI cache'],
            ['cellAiClearAll', 'Clear all AI cache'],
            ['cellAiTraceTitle', 'Cell AI Trace'],
            ['cellAiContextTraceTitle', 'Cell AI Trace']
        ];
        for (var i = 0; i < pairs.length; i += 1) {
            var element = document.getElementById(pairs[i][0]);
            if (!element) continue;
            element.textContent = tr(pairs[i][1]);
        }
        var button = document.getElementById('cellai-button');
        if (button) {
            button.title = tr('Cell AI Functions');
            button.setAttribute('aria-label', tr('Open cell AI helper'));
        }
    }

    function bindUi() {
        ui.button = document.getElementById('cellai-button');
        ui.panel = document.getElementById('cellAiPanel');
        ui.close = document.getElementById('cellAiClose');
        ui.mode = document.getElementById('cellAiMode');
        ui.status = document.getElementById('cellAiStatus');
        ui.formulaInput = document.getElementById('cellAiFormulaInput');
        ui.functionList = document.getElementById('cellAiFunctionList');
        ui.insert = document.getElementById('cellAiInsert');
        ui.refreshSelected = document.getElementById('cellAiRefreshSelected');
        ui.refreshAll = document.getElementById('cellAiRefreshAll');
        ui.clearSelected = document.getElementById('cellAiClearSelected');
        ui.clearAll = document.getElementById('cellAiClearAll');
        ui.helperTrace = document.getElementById('cellAiTraceList');
        ui.contextTrace = document.getElementById('cellAiContextTraceList');

        if (!ui.button || !ui.panel) return false;

        ui.button.addEventListener('click', togglePanel);
        if (ui.close) ui.close.addEventListener('click', closePanel);
        if (ui.insert) {
            ui.insert.addEventListener('click', function () {
                insertFormula(ui.formulaInput ? ui.formulaInput.value : '');
            });
        }
        if (ui.refreshSelected) ui.refreshSelected.addEventListener('click', function () { refreshSelected(false); });
        if (ui.refreshAll) ui.refreshAll.addEventListener('click', refreshAll);
        if (ui.clearSelected) ui.clearSelected.addEventListener('click', clearSelectedCache);
        if (ui.clearAll) ui.clearAll.addEventListener('click', clearAllCache);

        document.addEventListener('keydown', function (event) {
            if (event.key === 'Escape' && ui.panel && ui.panel.style.display !== 'none') {
                closePanel();
            }
        });

        return true;
    }

    function startTracePolling() {
        if (state.traceTimer) window.clearInterval(state.traceTimer);
        state.traceTimer = window.setInterval(function () {
            pullTrace();
        }, TRACE_PULL_INTERVAL_MS);
    }

    async function bootstrap() {
        if (state.initialized) return;
        if (!bindUi()) return;
        state.initialized = true;

        applyTranslations();
        renderFunctions();
        await detectCapabilities();
        updateModePill();

        if (state.editorType !== 'cell') {
            if (ui.button) ui.button.style.display = 'none';
            setStatus(tr('AI-функции доступны только в Spreadsheet Editor.'), true);
            return;
        }

        if (ui.formulaInput) ui.formulaInput.value = FUNCTION_DEFS[0].example;
        if (state.mode === 'bulk_only_fallback') {
            setStatus(tr('В текущей версии доступен fallback без формульного режима (bulk-only).'), true);
            return;
        }

        await registerCellFunctions(false);
        startTracePolling();
        pullTrace();
    }

    function patchPluginHooks() {
        var originalInit = window.Asc.plugin.init;
        window.Asc.plugin.init = function () {
            if (typeof originalInit === 'function') originalInit.apply(this, arguments);
            window.setTimeout(function () { bootstrap(); }, 240);
        };

        var originalTranslate = window.Asc.plugin.onTranslate;
        window.Asc.plugin.onTranslate = function () {
            if (typeof originalTranslate === 'function') originalTranslate.apply(this, arguments);
            applyTranslations();
            renderFunctions();
            updateModePill();
        };
    }

    function bindSettingsWatcher() {
        var saveButton = document.getElementById('inlineSaveSettings');
        if (!saveButton) return;
        saveButton.addEventListener('click', function () {
            window.setTimeout(function () {
                if (state.editorType === 'cell') registerCellFunctions(true);
            }, 450);
        });
    }

    patchPluginHooks();
    document.addEventListener('DOMContentLoaded', function () {
        bindSettingsWatcher();
        window.setTimeout(function () {
            bootstrap();
        }, 320);
    });
})(window);
