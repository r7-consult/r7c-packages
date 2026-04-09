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
    root.runtime = root.runtime || {};
    root.runtime.state = root.runtime.state || {};
    root.runtime.state.agent = root.runtime.state.agent || {
        mode: 'off',
        status: 'idle',
        currentStepIndex: -1,
        steps: [],
        trace: [],
        stopRequested: false,
        lastUserMessage: '',
        lastReadDocumentText: '',
        lastFailedStep: null,
        recoveryQueue: [],
        fastPathQueue: [],
        fastPathActive: false,
        currentRunContainer: null,
        abortController: null,
        runCounter: 0,
        pendingPlan: null,
        wordPlanMode: {
            enabled: false,
            awaitingApproval: false,
            approved: false,
            executionCountAfterApproval: 0,
            revisionCount: 0,
            approvedPlan: null
        }
    };

    function constants() {
        return root.runtime.constants || {
            defaultModel: 'openrouter/auto',
            agentLimits: {
                maxIterations: 50,
                stepTimeoutMs: 300000,
                maxMacroCodeChars: 20000,
                maxConsecutiveStepErrors: 4,
                maxPlannerContextChars: 120000,
                maxPlannerMessageChars: 6000,
                minPlannerTailMessages: 16,
                maxContextOverflowRecoveries: 2
            },
            plannerPolicyVersion: 'empty-probe-v3',
            apiGuidePath: 'docs/R7_MACRO_API_GUIDE.md',
            apiReferenceDefaultMethodLimit: 24,
            apiReferenceMaxMethodLimit: 120,
            agentLogPrefix: '[MacroRunner]'
        };
    }

    function agentState() {
        return root.runtime.state.agent;
    }

    function getContextService() {
        return root.services && root.services.context ? root.services.context : null;
    }

    function getChatService() {
        return root.services && root.services.chat ? root.services.chat : null;
    }

    function getChatThreadStore() {
        return root.services && root.services.chatThreads
            ? root.services.chatThreads
            : (root.features && root.features.chatThreadStore ? root.features.chatThreadStore : null);
    }

    function getSettingsService() {
        return root.features && root.features.settings ? root.features.settings : null;
    }

    function getHostBridge() {
        return root.platform && root.platform.hostBridge ? root.platform.hostBridge : null;
    }

    function getOpenRouterBridge() {
        return root.platform && root.platform.openrouter ? root.platform.openrouter : null;
    }

    function getProvidersBridge() {
        return root.platform && root.platform.providers ? root.platform.providers : null;
    }

    function getWebToolsBridge() {
        return root.platform && root.platform.webTools ? root.platform.webTools : null;
    }

    function getDesktopToolsBridge() {
        return root.platform && root.platform.desktopTools ? root.platform.desktopTools : null;
    }

    function getUi() {
        return root.ui || {};
    }

    function getCurrentEditorType() {
        var host = getHostBridge();
        return host && typeof host.getEditorTypeSafe === 'function' ? host.getEditorTypeSafe() : '';
    }

    function getErrorTools() {
        return root.shared && root.shared.errors ? root.shared.errors : null;
    }

    function normalizeTextPayload(value) {
        var context = getContextService();
        if (context && typeof context.normalizeTextPayload === 'function') {
            return context.normalizeTextPayload(value);
        }
        if (value === null || value === undefined) return '';
        return String(value).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    }

    function debugAgent(eventName, payload, level) {
        var logLevel = level || 'log';
        var logger = console && typeof console[logLevel] === 'function' ? console[logLevel] : console.log;
        var stamp = new Date().toISOString();
        if (payload === undefined) {
            logger(constants().agentLogPrefix + ' [' + stamp + '] ' + eventName);
            return;
        }
        logger(constants().agentLogPrefix + ' [' + stamp + '] ' + eventName, payload);
    }

    function cloneSerializable(value) {
        if (value === null || value === undefined) return value;
        try {
            return JSON.parse(JSON.stringify(value));
        } catch (error) {
            return value;
        }
    }

    function cloneMessages(messages) {
        var context = getContextService();
        if (context && typeof context.cloneMessages === 'function') {
            return context.cloneMessages(messages);
        }
        return Array.isArray(messages) ? messages.slice() : [];
    }

    function normalizePlannerConversationRole(role) {
        var normalized = normalizeTextPayload(role || '').toLowerCase();
        if (normalized === 'assistant' || normalized === 'user' || normalized === 'system') return normalized;
        return '';
    }

    function normalizePlannerConversationContent(content) {
        if (content === null || content === undefined) return '';
        if (typeof content === 'string') return normalizeTextPayload(content).trim();
        if (Array.isArray(content)) {
            return normalizeTextPayload(content.map(function (item) {
                if (typeof item === 'string') return item;
                if (!item || typeof item !== 'object') return '';
                if (typeof item.text === 'string') return item.text;
                if (typeof item.content === 'string') return item.content;
                if (typeof item.value === 'string') return item.value;
                return '';
            }).filter(Boolean).join('\n')).trim();
        }
        if (content && typeof content === 'object') {
            if (typeof content.text === 'string') return normalizeTextPayload(content.text).trim();
            if (typeof content.content === 'string') return normalizeTextPayload(content.content).trim();
            try {
                return normalizeTextPayload(JSON.stringify(content)).trim();
            } catch (error) {
                return '';
            }
        }
        return normalizeTextPayload(String(content)).trim();
    }

    function getConversationHistoryForPlanner() {
        var threadStore = getChatThreadStore();
        if (threadStore && typeof threadStore.syncConversationHistory === 'function') {
            var synced = threadStore.syncConversationHistory();
            if (Array.isArray(synced)) return synced;
        }
        var chat = getChatService();
        if (chat && typeof chat.getState === 'function') {
            var state = chat.getState();
            if (state && Array.isArray(state.conversationHistory)) {
                return state.conversationHistory;
            }
        }
        return [];
    }

    function buildPlannerConversationSeed(userMessage) {
        var currentUserMessage = normalizeTextPayload(userMessage || '').trim();
        var history = getConversationHistoryForPlanner();
        var tailLimit = 10;
        var tail = Array.isArray(history) ? history.slice(-tailLimit) : [];
        var seed = [];

        tail.forEach(function (item) {
            var role = normalizePlannerConversationRole(item && item.role);
            if (!role) return;
            var content = normalizePlannerConversationContent(item && item.content);
            if (Array.isArray(item && item.attachments)) {
                item.attachments.forEach(function (att) {
                    if (att && att.contextSummary && att.contextSummary.kind === 'xlsx_workbook') {
                        var summary = att.contextSummary;
                        var text = '\n[ATTACHED WORKBOOK: ' + (att.name || 'unnamed.xlsx') + ']\n';
                        text += 'Sheets (' + summary.sheetCount + ' total): ' + summary.sheetNames.join(', ') + '\n';
                        if (summary.sheets && summary.sheets.length) {
                            summary.sheets.forEach(function (sheet) {
                                text += 'Sheet: ' + sheet.name + ' | range: ' + sheet.range + ' | rows: ' + sheet.rows + ' | cols: ' + sheet.cols + '\n';
                                if (sheet.preview) {
                                    text += 'Preview:\n' + sheet.preview + '\n';
                                }
                            });
                        }
                        text += '[/ATTACHED WORKBOOK]\n';
                        content += text;
                    }
                });
            }
            if (!content.length) return;
            seed.push({
                role: role,
                content: sanitizePlannerMessageContent(content, 2400)
            });
        });

        if (!seed.length) {
            if (currentUserMessage.length) {
                seed.push({ role: 'user', content: currentUserMessage });
            }
            return seed;
        }

        if (currentUserMessage.length) {
            var last = seed[seed.length - 1];
            var lastText = normalizeTextPayload(last && last.content || '').trim();
            if (!(last && last.role === 'user' && lastText === currentUserMessage)) {
                seed.push({ role: 'user', content: currentUserMessage });
            }
        }

        return seed;
    }

    function createDefaultResearchPolicy() {
        return {
            required: false,
            query: '',
            completed: false,
            attempted: false,
            failed: false,
            unavailable: false,
            webSearchEnabled: false,
            activeProvider: '',
            braveApiConfigured: false,
            exaApiConfigured: false
        };
    }

    function createDefaultWordPlanState() {
        return {
            enabled: false,
            awaitingApproval: false,
            approved: false,
            executionCountAfterApproval: 0,
            revisionCount: 0,
            approvedPlan: null
        };
    }

    function createDefaultAgentState() {
        return {
            mode: 'off',
            status: 'idle',
            currentStepIndex: -1,
            steps: [],
            trace: [],
            stopRequested: false,
            lastUserMessage: '',
            lastReadDocumentText: '',
            lastFailedStep: null,
            recoveryQueue: [],
            fastPathQueue: [],
            fastPathActive: false,
            currentRunContainer: null,
            abortController: null,
            runCounter: 0,
            researchPolicy: createDefaultResearchPolicy(),
            lastHostToolDiscovery: null,
            pendingPlan: null,
            wordPlanMode: createDefaultWordPlanState()
        };
    }

    function loadAgentRuntimeSettings() {
        var chat = getChatService();
        if (chat && typeof chat.loadRuntimeSettings === 'function') {
            return chat.loadRuntimeSettings();
        }
        var settings = getSettingsService();
        if (settings && typeof settings.loadSettings === 'function') {
            return settings.loadSettings();
        }
        return {
            desktopTools: {
                automationMode: 'auto',
                disabledTools: []
            },
            webTools: {
                provider: '',
                providers: {
                    exa: { apiKey: '' },
                    brave: { apiKey: '' }
                }
            }
        };
    }

    function normalizeDesktopAutomationMode(value) {
        var mode = normalizeTextPayload(value || '').toLowerCase();
        if (mode === 'native_tools' || mode === 'macro_only') return mode;
        return 'auto';
    }

    function normalizeDesktopToolsSettings(value) {
        var source = value && typeof value === 'object' ? value : {};
        return {
            automationMode: normalizeDesktopAutomationMode(source.automationMode),
            disabledTools: Array.isArray(source.disabledTools) ? source.disabledTools.slice() : []
        };
    }

    function isExplicitMacroRequest(messageText) {
        var text = normalizeTextPayload(messageText || '');
        if (!text) return false;
        return /(\bmacro\b|\bmacros\b|макрос|макросы|напиши макрос|создай макрос|save macro|document macro|setmacros)/i.test(text);
    }

    function isLongFormResearchWritingRequest(messageText) {
        var text = normalizeTextPayload(messageText || '');
        if (!text.length) return false;
        var asksToWrite = /(напиши|написать|сделай|создай|подготовь|составь|оформи|write|draft|prepare|create|generate|compose)/i.test(text);
        var longFormType = /(диплом|дипломн|курсов|реферат|эссе|стать[ьяю]|доклад|отч[её]т|обзор|исследован|thesis|dissertation|article|essay|report|paper|review)/i.test(text);
        return asksToWrite && longFormType;
    }

    function isWordPlanModeRequest(messageText) {
        var text = normalizeTextPayload(messageText || '');
        if (!text.length) return false;
        if (isExplicitMacroRequest(text)) return false;
        var asksToWrite = /(напиши|написать|сделай|создай|подготовь|составь|оформи|write|draft|prepare|create|generate|compose)/i.test(text);
        if (!asksToWrite) return false;
        var longFormType = /(диплом|дипломн|курсов|реферат|эссе|стать[ьяю]|доклад|отч[её]т|обзор|исследован|thesis|dissertation|article|essay|report|paper|review|гайд|руководств|whitepaper|лендинг|презентац|proposal|brief)/i.test(text);
        var hasImageNeed = /(картин|изображен|иллюстрац|рисунк|обложк|visual|image|picture|illustration|hero image|cover)/i.test(text);
        var hasResearchNeed = /(интернет|веб|web|search|research|источник|источники|ссылк|пруф|факт|crawling|brave|exa)/i.test(text);
        var hasPlanNeed = /(план|структур|outline|section|абзац|раздел|глав)/i.test(text);
        return longFormType || hasImageNeed || hasResearchNeed || hasPlanNeed;
    }

    function buildResearchQuery(messageText) {
        var original = normalizeTextPayload(messageText || '').replace(/\s+/g, ' ').trim();
        if (!original.length) return '';

        var stripped = original
            .replace(/^(пожалуйста|please)\s+/i, '')
            .replace(/^(напиши|написать|сделай|создай|подготовь|составь|оформи|write|draft|prepare|create|generate|compose)\s+/i, '')
            .replace(/^(мне\s+)?(диплом(ную работу)?|курсов(ую|ая)?( работу)?|реферат|эссе|стать[ьяю]|доклад|отч[её]т|обзор|thesis|dissertation|article|essay|report|paper|review)\s*/i, '')
            .replace(/^(на тему|по теме|about|on)\s+/i, '')
            .trim();

        stripped = stripped
            .split(/[.!?\n;]/)[0]
            .replace(/\b(нужен|нужно|сделай|добавь|prepare|include)\b[\s\S]*$/i, '')
            .replace(/\b(с картинками|с изображениями|с иллюстрациями|3 картинки|три картинки|по разделам|с планом|с источниками|с пруфами|с фактами из интернета)\b[\s\S]*$/i, '')
            .replace(/^(про|о)\s+/i, '')
            .trim();

        var serviceLeadTopic = stripped.match(/^(?:с|with)\s+.+?(?:\s+про|\s+о|\s+about|\s+on)\s+(.+)$/i);
        if (serviceLeadTopic && serviceLeadTopic[1]) {
            stripped = normalizeTextPayload(serviceLeadTopic[1]).trim();
        }

        if (stripped.length < 8) {
            var topicMatch = original.match(/(?:^|\s)(?:про|о|about|on)\s+(.+)$/i);
            if (topicMatch && topicMatch[1]) {
                stripped = normalizeTextPayload(topicMatch[1])
                    .replace(/[.!?\n;]+$/g, '')
                    .trim();
            }
        }
        if (stripped.length < 8) stripped = original;
        return stripped.slice(0, 240);
    }

    function buildResearchPolicyForRequest(messageText, options) {
        var policy = createDefaultResearchPolicy();
        var opts = options && typeof options === 'object' ? options : {};
        var runtimeSettings = loadAgentRuntimeSettings();
        var webTools = getWebToolsBridge();
        var webToolState = runtimeSettings && runtimeSettings.webTools && typeof runtimeSettings.webTools === 'object' ? runtimeSettings.webTools : {};
        var providers = webToolState.providers && typeof webToolState.providers === 'object' ? webToolState.providers : {};

        policy.braveApiConfigured = !!String(providers.brave && providers.brave.apiKey || '').trim();
        policy.exaApiConfigured = !!String(providers.exa && providers.exa.apiKey || '').trim();

        if (webTools && typeof webTools.getCurrentProvider === 'function') {
            var currentProvider = webTools.getCurrentProvider() || {};
            policy.activeProvider = normalizeTextPayload(currentProvider.provider || '').toLowerCase();
        }
        if (webTools && typeof webTools.isEnabled === 'function') {
            policy.webSearchEnabled = !!webTools.isEnabled(runtimeSettings);
        }

        policy.required = policy.webSearchEnabled && (
            opts.forceRequired === true ||
            isLongFormResearchWritingRequest(messageText)
        );
        policy.query = policy.required ? buildResearchQuery(messageText) : '';
        return policy;
    }

    function buildWordPlanModeState(messageText) {
        return {
            enabled: getCurrentEditorType() === 'word' && isWordPlanModeRequest(messageText),
            awaitingApproval: false,
            approved: false,
            executionCountAfterApproval: 0,
            revisionCount: 0,
            approvedPlan: null
        };
    }

    function buildPlannerWebSearchStatusLine(policy) {
        var current = policy && typeof policy === 'object' ? policy : createDefaultResearchPolicy();
        return [
            'WEB_SEARCH_STATUS:',
            'enabled=' + (current.webSearchEnabled ? 'true' : 'false') + ';',
            'unavailable=' + (current.unavailable ? 'true' : 'false') + ';',
            'active_provider=' + (current.activeProvider || 'none') + ';',
            'brave_api_configured=' + (current.braveApiConfigured ? 'true' : 'false') + ';',
            'exa_api_configured=' + (current.exaApiConfigured ? 'true' : 'false')
        ].join(' ');
    }

    function isDesktopToolsEditorSupported(editorType) {
        return editorType === 'word' || editorType === 'cell';
    }

    function buildDesktopToolsStatusLine(state) {
        var current = state && typeof state === 'object' ? state : {};
        return [
            'available=' + (current.catalogAvailable ? 'true' : 'false') + ';',
            'execution=' + (current.executionAvailable ? 'true' : 'false') + ';',
            'tool_count=' + String(current.toolCount || 0) + ';',
            'mode=' + String(current.automationMode || 'auto') + ';',
            'explicit_macro_request=' + (current.explicitMacroRequest ? 'true' : 'false') + ';',
            'planner_enabled=' + (current.enabledForPlanner ? 'true' : 'false') + ';',
            'reason=' + String(current.reason || 'none')
        ].join(' ');
    }

    function buildDesktopToolsPlannerState(messageText) {
        var settings = loadAgentRuntimeSettings();
        var desktopSettings = normalizeDesktopToolsSettings(settings && settings.desktopTools);
        var bridge = getDesktopToolsBridge();
        var host = getHostBridge();
        var editorType = host && typeof host.getEditorTypeSafe === 'function' ? host.getEditorTypeSafe() : '';
        var status = bridge && typeof bridge.getStatus === 'function'
            ? bridge.getStatus()
            : {
                desktopEditorAvailable: false,
                catalogAvailable: false,
                executionAvailable: false,
                catalogParseError: '',
                toolCount: 0
            };
        var explicitMacroRequest = isExplicitMacroRequest(messageText);
        var supportedEditor = isDesktopToolsEditorSupported(editorType);
        var enabledForPlanner = false;
        var reason = 'disabled';

        if (!supportedEditor) {
            reason = 'editor_not_supported';
        } else if (desktopSettings.automationMode === 'macro_only') {
            reason = 'macro_only_mode';
        } else if (explicitMacroRequest) {
            reason = 'explicit_macro_request';
        } else if (!status.catalogAvailable) {
            reason = status.catalogParseError ? 'catalog_parse_error' : 'catalog_unavailable';
        } else if (!status.executionAvailable) {
            reason = 'execution_unavailable';
        } else if (!status.toolCount) {
            reason = 'empty_catalog';
        } else {
            enabledForPlanner = true;
            reason = desktopSettings.automationMode === 'native_tools' ? 'native_tools_mode' : 'auto_mode';
        }

        return {
            editorType: editorType,
            automationMode: desktopSettings.automationMode,
            disabledTools: desktopSettings.disabledTools.slice(),
            desktopEditorAvailable: status.desktopEditorAvailable === true,
            catalogAvailable: status.catalogAvailable === true,
            executionAvailable: status.executionAvailable === true,
            catalogParseError: normalizeTextPayload(status.catalogParseError || ''),
            toolCount: Number(status.toolCount || 0) || 0,
            catalogHash: normalizeTextPayload(status.catalogHash || ''),
            enabledForPlanner: enabledForPlanner,
            explicitMacroRequest: explicitMacroRequest,
            reason: reason,
            statusLine: buildDesktopToolsStatusLine({
                catalogAvailable: status.catalogAvailable === true,
                executionAvailable: status.executionAvailable === true,
                toolCount: status.toolCount || 0,
                automationMode: desktopSettings.automationMode,
                explicitMacroRequest: explicitMacroRequest,
                enabledForPlanner: enabledForPlanner,
                reason: reason
            }),
            catalogForPrompt: enabledForPlanner
                ? buildRelevantDesktopToolsPromptCatalog(messageText, 16)
                : ''
        };
    }

    function buildHostToolInputPreview(value) {
        try {
            var serialized = JSON.stringify(value && typeof value === 'object' ? value : {}, null, 2);
            if (serialized.length > 420) return serialized.slice(0, 420) + '\n...[truncated]';
            return serialized;
        } catch (error) {
            return '';
        }
    }

    function isImageGenerationRequest(messageText) {
        var text = normalizeTextPayload(messageText || '').trim();
        if (!text.length) return false;
        var hasAction = /(\u0441\u0433\u0435\u043d\u0435\u0440|\u0441\u043e\u0437\u0434\u0430\u0439|\u043d\u0430\u0440\u0438\u0441\u0443\u0439|\u0441\u0434\u0435\u043b\u0430\u0439|draw|paint|illustrat|render|sketch|generate|create|make)/i.test(text);
        var hasImageObject = /(\u043a\u0430\u0440\u0442\u0438\u043d|\u0438\u0437\u043e\u0431\u0440\u0430\u0436\u0435\u043d|\u0438\u043b\u043b\u044e\u0441\u0442\u0440\u0430\u0446|\u0440\u0438\u0441\u0443\u043d|\u0430\u0440\u0442|\u0444\u043e\u0442\u043e|\u0432\u0438\u0437\u0443\u0430\u043b|image|picture|illustration|drawing|art|photo|visual|cover|icon)/i.test(text);
        return hasAction && hasImageObject;
    }

    function extractImageRequestSubject(messageText) {
        var text = normalizeTextPayload(messageText || '').trim();
        if (!text.length) return '';
        var subject = text
            .replace(/^\/draw\s+/i, '')
            .replace(/^(\u043f\u043e\u0436\u0430\u043b\u0443\u0439\u0441\u0442\u0430|please)\s+/i, '')
            .replace(/(\u0441\u0433\u0435\u043d\u0435\u0440|\u0441\u043e\u0437\u0434\u0430\u0439|\u043d\u0430\u0440\u0438\u0441\u0443\u0439|\u0441\u0434\u0435\u043b\u0430\u0439|draw|paint|illustrat(?:e|ion)?|render|sketch|generate|create|make)/ig, ' ')
            .replace(/(\u043a\u0430\u0440\u0442\u0438\u043d(?:\u043a|\u0443|\u0430|\u0435)?|\u0438\u0437\u043e\u0431\u0440\u0430\u0436\u0435\u043d(?:\u0438|\u0438\u0435|\u0438\u044f)?|\u0438\u043b\u043b\u044e\u0441\u0442\u0440\u0430\u0446(?:\u0438|\u0438\u044e|\u0438\u044f)?|\u0440\u0438\u0441\u0443\u043d(?:\u043e\u043a|\u043a\u0430|\u043a\u0438)?|\u0430\u0440\u0442|\u0444\u043e\u0442\u043e|\u0432\u0438\u0437\u0443\u0430\u043b(?:\u0438\u0437\u0430\u0446\u0438\u044e)?|image|picture|illustration|drawing|art|photo|visual|cover|icon)/ig, ' ')
            .replace(/\s+/g, ' ')
            .replace(/^[\s,:;\-]+|[\s,:;\-]+$/g, '')
            .trim();
        return subject;
    }

    function buildImageGenerationPrompt(userMessage) {
        var subject = extractImageRequestSubject(userMessage);
        var documentText = normalizeTextPayload(agentState().lastReadDocumentText || '')
            .replace(/\s+/g, ' ')
            .trim();
        var docSnippet = documentText.length > 520 ? documentText.slice(0, 520) + '...' : documentText;

        if (subject.length >= 8) {
            if (docSnippet) {
                return 'Create a polished, professional illustration of ' + subject + '. Use this document context for relevance: ' + docSnippet + '. No text in the image.';
            }
            return 'Create a polished, professional illustration of ' + subject + '. No text in the image.';
        }
        if (docSnippet) {
            return 'Create a polished, professional editorial illustration for this document context: ' + docSnippet + '. No text in the image.';
        }
        return 'Create a polished, professional illustration that matches the user request. No text in the image.';
    }

    function extractImageCommand(answerText) {
        var text = normalizeTextPayload(answerText || '');
        var match = text.match(/\[GENERATE_IMAGE:\s*([^\]]+)\]/i);
        if (!match || !match[1]) return null;
        return {
            prompt: match[1].trim(),
            visibleText: text.replace(match[0], '').trim()
        };
    }

    function reset(options) {
        var settings = options && typeof options === 'object' ? options : {};
        var current = agentState();
        var mode = current.mode;
        var trace = Array.isArray(current.trace) ? current.trace.slice() : [];
        root.runtime.state.agent = createDefaultAgentState();
        root.runtime.state.agent.mode = mode;
        if (settings.preserveTrace) {
            root.runtime.state.agent.trace = trace;
        }
        if (getUi().scheduleRenderContextPanel) getUi().scheduleRenderContextPanel(true);
        if (getUi().scheduleRenderTrace) getUi().scheduleRenderTrace();
        return agentState();
    }

    function resolveAgentMode() {
        var editorType = getCurrentEditorType();
        agentState().mode = (editorType === 'cell' || editorType === 'word') ? 'macro_runner' : 'off';
        return agentState().mode;
    }

    function shouldUseMacroRunner() {
        return resolveAgentMode() === 'macro_runner';
    }

    function isBusy() {
        var status = agentState().status;
        return status === 'planning' || status === 'executing' || status === 'answering';
    }

    function canRetry() {
        return !isBusy() && !!agentState().lastFailedStep;
    }

    function canRunFromStart() {
        return !isBusy() && !!String(agentState().lastUserMessage || '').trim().length;
    }

    function setStatus(status) {
        agentState().status = status;
        if (getUi().scheduleRenderContextPanel) getUi().scheduleRenderContextPanel(true);
        if (getUi().scheduleRenderTrace) getUi().scheduleRenderTrace();
    }

    function addTraceRecord(record) {
        var item = Object.assign({
            ts: new Date().toISOString(),
            status: 'info'
        }, record || {});
        agentState().trace.push(item);
        debugAgent('trace_record', item);
        if (getUi().scheduleRenderTrace) getUi().scheduleRenderTrace();
    }

    function shouldShowModelReasoningTrace(runtimeSettings) {
        return !(runtimeSettings && runtimeSettings.trace && runtimeSettings.trace.showModelReasoning === false);
    }

    function addModelReasoningTraceFromPlanner(response, options) {
        var settings = options && typeof options === 'object' ? options : {};
        if (!shouldShowModelReasoningTrace(settings.runtimeSettings)) return;
        var reasoning = response && response.reasoning && typeof response.reasoning === 'object'
            ? response.reasoning
            : null;
        if (!reasoning || reasoning.available !== true) return;
        var summary = normalizeTextPayload(reasoning.summary || '').replace(/\s+/g, ' ').trim();
        if (!summary) return;
        if (summary.length > 320) summary = summary.slice(0, 317) + '...';
        var effort = normalizeTextPayload(reasoning.effort || '').toLowerCase();
        var tokens = reasoning.tokens !== undefined ? Number(reasoning.tokens) : undefined;
        addTraceRecord({
            step_id: 'planner_reasoning_' + Date.now(),
            step_type: 'model_reasoning',
            status: 'info',
            reason: 'Model thinking (summary): ' + summary,
            reasoning_summary: summary,
            reasoning_effort: effort,
            reasoning_tokens: Number.isFinite(tokens) ? tokens : undefined,
            provider: normalizeTextPayload(settings.provider || ''),
            duration_ms: Math.max(0, Number(settings.durationMs || 0))
        });
    }

    function sleep(ms) {
        return new Promise(function (resolve) {
            globalRoot.setTimeout(resolve, Number(ms || 0));
        });
    }

    function withTimeout(promise, timeoutMs, label) {
        return new Promise(function (resolve, reject) {
            var settled = false;
            var timer = globalRoot.setTimeout(function () {
                if (settled) return;
                settled = true;
                reject(new Error((label || 'Operation') + ' timed out in ' + timeoutMs + ' ms'));
            }, timeoutMs);

            promise.then(function (value) {
                if (settled) return;
                settled = true;
                globalRoot.clearTimeout(timer);
                resolve(value);
            }, function (error) {
                if (settled) return;
                settled = true;
                globalRoot.clearTimeout(timer);
                reject(error);
            });
        });
    }

    function isTransientPlannerError(error) {
        var status = Number(error && error.status ? error.status : 0);
        return status === 429 || status === 500 || status === 502 || status === 503 || status === 504;
    }

    function isPlannerOutputFormatError(error) {
        var text = normalizeTextPayload(error && error.message ? error.message : error).toLowerCase();
        return /empty response|not valid json|json must be an object|unsupported step type/.test(text);
    }

    function isContextOverflowError(error) {
        var status = Number(error && error.status ? error.status : 0);
        var text = normalizeTextPayload(error && error.message ? error.message : error).toLowerCase();
        var containsContextSignal = /maximum context length|context length|too many tokens|token limit|prompt is too long|requested about .* tokens|context-compression plugin/.test(text);
        if (status === 400 || status === 413) {
            return containsContextSignal || /context|token|payload too large|request too large|limit exceeded/.test(text);
        }
        return containsContextSignal;
    }

    function isTransientStepError(message) {
        var text = normalizeTextPayload(message || '').toLowerCase();
        return /timeout|temporar|rate limit|429|busy|throttl/.test(text);
    }

    function formatPlannerError(error) {
        var tools = getErrorTools();
        if (tools && typeof tools.buildProviderErrorMessage === 'function') {
            return tools.buildProviderErrorMessage(error);
        }
        return String(error && error.message ? error.message : error || 'Planner failed');
    }

    function sanitizeMacroCode(code) {
        var text = normalizeTextPayload(code || '').trim();
        if (!text.length) return '';
        if (text.indexOf('```') === 0) {
            text = text.replace(/^```[a-zA-Z0-9_-]*\n?/, '').replace(/```$/, '').trim();
        }
        return text;
    }

    function buildMacroTraceMeta(code, maxChars) {
        var text = sanitizeMacroCode(code);
        var limit = Number(maxChars || 420);
        return {
            length: text.length,
            preview: text.length > limit ? text.slice(0, limit) + '\n...[truncated]' : text
        };
    }

    function truncateInlineText(value, maxChars) {
        var text = normalizeTextPayload(value || '').replace(/\s+/g, ' ').trim();
        var limit = Math.max(24, Number(maxChars || 0) || 240);
        if (text.length <= limit) return text;
        return text.slice(0, limit - 3).trim() + '...';
    }

    function isDataImageUrl(value) {
        return /^data:image\/[a-z0-9.+-]+;base64,/i.test(normalizeTextPayload(value || '').trim());
    }

    function compactDataImageUrl(value) {
        var text = normalizeTextPayload(value || '').trim();
        if (!text.length) return '';
        return '[data-image-omitted chars=' + text.length + ']';
    }

    function sanitizeImageUrlForPlanner(value) {
        var text = normalizeTextPayload(value || '').trim();
        if (!text.length) return '';
        if (isDataImageUrl(text)) return compactDataImageUrl(text);
        return truncateInlineText(text, 220);
    }

    function compactArbitraryResultData(value, depth) {
        var currentDepth = Number(depth || 0);
        if (currentDepth > 4) return '[depth_limit]';
        if (value === null || value === undefined) return null;
        if (typeof value === 'string') {
            if (isDataImageUrl(value)) return compactDataImageUrl(value);
            return truncateInlineText(value, 420);
        }
        if (typeof value === 'number' || typeof value === 'boolean') return value;
        if (Array.isArray(value)) {
            return value.slice(0, 10).map(function (item) {
                return compactArbitraryResultData(item, currentDepth + 1);
            });
        }
        if (typeof value === 'object') {
            var output = {};
            Object.keys(value).slice(0, 20).forEach(function (key) {
                output[key] = compactArbitraryResultData(value[key], currentDepth + 1);
            });
            return output;
        }
        return truncateInlineText(String(value), 200);
    }

    function sanitizePlannerMessageContent(value, maxChars) {
        var text = normalizeTextPayload(value || '');
        var limit = Math.max(800, Number(maxChars || 0) || 6000);
        text = text.replace(/data:image\/[a-z0-9.+-]+;base64,[a-z0-9+/=\s]+/ig, function (match) {
            return compactDataImageUrl(match);
        });
        if (text.length <= limit) return text;
        return text.slice(0, Math.max(0, limit - 16)) + '\n...[truncated]';
    }

    function estimatePlannerMessagesChars(messages) {
        var total = 0;
        (Array.isArray(messages) ? messages : []).forEach(function (item) {
            var content = normalizeTextPayload(item && item.content || '');
            total += content.length + 32;
        });
        return total;
    }

    function summarizeDroppedPlannerMessages(messages) {
        var list = Array.isArray(messages) ? messages : [];
        if (!list.length) return '';
        var roleCounts = {};
        var firstUserExcerpt = '';
        var lastToolStep = '';
        list.forEach(function (item) {
            var role = normalizeTextPayload(item && item.role || 'user').toLowerCase() || 'user';
            roleCounts[role] = Number(roleCounts[role] || 0) + 1;
            var content = normalizeTextPayload(item && item.content || '');
            if (!firstUserExcerpt && role === 'user' && content.length) {
                firstUserExcerpt = truncateInlineText(content, 120);
            }
            if (role === 'user' && content.indexOf('"type": "tool_result"') !== -1) {
                var match = content.match(/"step_type"\s*:\s*"([^"]+)"/i);
                if (match && match[1]) lastToolStep = match[1];
            }
        });
        var roleSummary = Object.keys(roleCounts).map(function (role) {
            return role + '=' + roleCounts[role];
        }).join(', ');
        var output = 'RUNTIME_CONTEXT_SUMMARY: dropped_messages=' + list.length + '; roles=' + roleSummary;
        if (lastToolStep) output += '; last_tool=' + lastToolStep;
        if (firstUserExcerpt) output += '; first_user_excerpt="' + firstUserExcerpt + '"';
        return output;
    }

    function compactPlannerMessagesInPlace(plannerMessages, options) {
        var settings = options && typeof options === 'object' ? options : {};
        var limits = constants().agentLimits || {};
        var maxTotalChars = Math.max(32000, Number(limits.maxPlannerContextChars || 120000));
        var perMessageChars = Math.max(800, Number(limits.maxPlannerMessageChars || 6000));
        var keepTailMessages = Math.max(8, Number(limits.minPlannerTailMessages || 16));
        var changed = false;
        var removedCount = 0;
        if (!Array.isArray(plannerMessages) || !plannerMessages.length) {
            return {
                changed: false,
                removedMessages: 0,
                totalChars: 0
            };
        }

        for (var i = 0; i < plannerMessages.length; i += 1) {
            if (!plannerMessages[i] || typeof plannerMessages[i] !== 'object') continue;
            var nextContent = sanitizePlannerMessageContent(plannerMessages[i].content, perMessageChars);
            if (plannerMessages[i].content !== nextContent) {
                plannerMessages[i].content = nextContent;
                changed = true;
            }
        }

        var totalChars = estimatePlannerMessagesChars(plannerMessages);
        if (totalChars > maxTotalChars && plannerMessages.length > keepTailMessages) {
            var splitIndex = Math.max(1, plannerMessages.length - keepTailMessages);
            var droppedMessages = plannerMessages.slice(0, splitIndex);
            var preservedMessages = plannerMessages.slice(splitIndex);
            removedCount = droppedMessages.length;
            plannerMessages.length = 0;
            var summaryMessage = summarizeDroppedPlannerMessages(droppedMessages);
            if (summaryMessage) {
                plannerMessages.push({
                    role: 'system',
                    content: sanitizePlannerMessageContent(summaryMessage, perMessageChars)
                });
            }
            Array.prototype.push.apply(plannerMessages, preservedMessages);
            changed = true;
            totalChars = estimatePlannerMessagesChars(plannerMessages);
        }

        while (totalChars > maxTotalChars && plannerMessages.length > 2) {
            plannerMessages.splice(1, 1);
            removedCount += 1;
            changed = true;
            totalChars = estimatePlannerMessagesChars(plannerMessages);
        }

        return {
            changed: changed,
            removedMessages: removedCount,
            totalChars: totalChars,
            reason: normalizeTextPayload(settings.reason || '')
        };
    }

    function normalizePlanString(value, fallback, maxChars) {
        var text = truncateInlineText(value || '', maxChars || 240);
        if (text) return text;
        return normalizeTextPayload(fallback || '');
    }

    function normalizePlanPoints(list, maxItems, maxCharsPerItem) {
        var items = Array.isArray(list) ? list : [];
        return items.map(function (item) {
            return truncateInlineText(item, maxCharsPerItem || 180);
        }).filter(Boolean).slice(0, Math.max(1, Number(maxItems || 3) || 3));
    }

    function normalizeWordPlanSection(rawSection, index) {
        var section = rawSection && typeof rawSection === 'object' ? rawSection : {};
        var image = section.image && typeof section.image === 'object' ? section.image : {};
        var keyPoints = normalizePlanPoints(section.key_points || section.keyPoints || section.points || section.bullets, 4, 180);
        var title = normalizePlanString(section.title || section.heading, 'Section ' + (index + 1), 120);
        var purpose = normalizePlanString(section.purpose || section.goal || section.summary, '', 220);
        var prompt = normalizePlanString(image.prompt || section.image_prompt || section.imagePrompt, '', 320);
        var caption = normalizePlanString(image.caption || section.image_caption || section.imageCaption, '', 180);
        var placement = normalizePlanString(image.placement || section.image_placement || 'after_section', 'after_section', 48).toLowerCase();

        return {
            id: normalizePlanString(section.id, 'section_' + (index + 1), 48).replace(/\s+/g, '_').toLowerCase(),
            title: title,
            purpose: purpose,
            keyPoints: keyPoints,
            image: {
                needed: image.needed === true || !!prompt,
                placement: placement || 'after_section',
                purpose: normalizePlanString(image.purpose || section.image_purpose, '', 180),
                prompt: prompt,
                caption: caption
            }
        };
    }

    function normalizeWordPlanSources(rawSources) {
        var sources = Array.isArray(rawSources) ? rawSources : [];
        return sources.map(function (item) {
            var source = item && typeof item === 'object' ? item : {};
            return {
                title: normalizePlanString(source.title || source.label || source.name, '', 140),
                url: truncateInlineText(source.url || source.link || '', 280),
                note: normalizePlanString(source.note || source.summary || source.fact, '', 180)
            };
        }).filter(function (item) {
            return item.title || item.url || item.note;
        }).slice(0, 6);
    }

    function normalizeWordPlan(rawPlan, stepId) {
        var plan = rawPlan && typeof rawPlan === 'object' ? rawPlan : {};
        var sections = Array.isArray(plan.sections) ? plan.sections : [];
        var normalizedSections = sections.map(normalizeWordPlanSection).filter(function (section) {
            return section && section.title;
        }).slice(0, 8);

        if (!normalizedSections.length) {
            normalizedSections.push(normalizeWordPlanSection({}, 0));
        }

        return {
            planId: normalizePlanString(plan.planId || stepId || ('plan_' + Date.now()), 'plan_' + Date.now(), 80).replace(/\s+/g, '_'),
            title: normalizePlanString(plan.title, 'Document plan', 140),
            objective: normalizePlanString(plan.objective || plan.goal, '', 220),
            summary: normalizePlanString(plan.summary || plan.brief, '', 320),
            approvalHint: normalizePlanString(plan.approval_hint || plan.approvalHint, 'Approve to execute, or send edits in chat.', 180),
            sections: normalizedSections,
            sources: normalizeWordPlanSources(plan.sources),
            researchSummary: normalizePlanString(plan.research_summary || plan.researchSummary, '', 220)
        };
    }

    function getPlanApprovalCommandText() {
        return 'Начинай исполнять план';
    }

    function isPlanApprovalMessage(messageText) {
        var text = normalizeTextPayload(messageText || '').replace(/\s+/g, ' ').trim().toLowerCase();
        if (!text) return false;
        if (text === getPlanApprovalCommandText().toLowerCase()) return true;
        if (/(^|\s)(approve|approved|execute|implement|start|begin|run)(?:\s+the)?\s+plan($|\s)/i.test(text)) return true;
        if (/(^|\s)(approve|approved)\s+this\s+plan($|\s)/i.test(text)) return true;
        if (/(^|\s)(утверждаю|одобряю)\s+план($|\s)/i.test(text)) return true;
        if (/(^|\s)(начинай|исполняй|выполняй|приступай)(?:\s+[^\s]+){0,4}\s+план($|\s)/i.test(text)) return true;
        if (/^давай\s+начинай(?:\s+[^\s]+){0,4}\s+план$/i.test(text)) return true;
        return false;
    }

    function tryParseJsonObject(candidate) {
        try {
            return JSON.parse(candidate);
        } catch (error) {
            return null;
        }
    }

    function extractJsonObjectsFromText(text) {
        var source = normalizeTextPayload(text || '');
        var results = [];
        var depth = 0;
        var inString = false;
        var escapeNext = false;
        var startIndex = -1;

        for (var i = 0; i < source.length; i += 1) {
            var ch = source[i];
            if (escapeNext) {
                escapeNext = false;
                continue;
            }
            if (ch === '\\') {
                escapeNext = true;
                continue;
            }
            if (inString) {
                if (ch === '"') inString = false;
                continue;
            }
            if (ch === '"') {
                inString = true;
                continue;
            }
            if (ch === '{') {
                if (depth === 0) startIndex = i;
                depth += 1;
                continue;
            }
            if (ch === '}' && depth > 0) {
                depth -= 1;
                if (depth === 0 && startIndex >= 0) {
                    results.push(source.slice(startIndex, i + 1));
                    startIndex = -1;
                }
            }
        }
        return results;
    }

    function looksLikePlannerStepObject(obj) {
        var current = obj && typeof obj === 'object' && obj.step && typeof obj.step === 'object' ? obj.step : obj;
        return !!(current && typeof current === 'object' && typeof current.type === 'string' && current.type.trim().length);
    }

    function parsePlannerResponse(rawText) {
        var text = normalizeTextPayload(rawText || '').trim();
        if (!text.length) throw new Error('Planner returned empty response');

        var parsed = tryParseJsonObject(text);
        if (parsed) return parsed;

        var fence = text.match(/```(?:json)?\s*([\s\S]*?)\s*```/i);
        if (fence && fence[1]) {
            parsed = tryParseJsonObject(fence[1]);
            if (parsed) return parsed;
        }

        var candidates = extractJsonObjectsFromText(text);
        var parsedCandidates = [];
        candidates.forEach(function (candidate) {
            var value = tryParseJsonObject(candidate);
            if (value && typeof value === 'object') {
                parsedCandidates.push(value);
            }
        });

        for (var i = 0; i < parsedCandidates.length; i += 1) {
            if (looksLikePlannerStepObject(parsedCandidates[i])) {
                return parsedCandidates[i];
            }
        }
        if (parsedCandidates.length) return parsedCandidates[0];
        throw new Error('Planner response is not valid JSON');
    }

    function normalizePlannerStep(parsed, stepIndex) {
        var current = parsed && typeof parsed === 'object' && parsed.step && typeof parsed.step === 'object'
            ? parsed.step
            : parsed;
        var allowedTypes = {
            collect_context: true,
            read_document_snapshot: true,
            present_plan: true,
            generate_image_asset: true,
            get_current_time: true,
            list_sheets: true,
            read_active_sheet: true,
            read_document: true,
            read_sheet_range: true,
            list_host_tools: true,
            call_host_tool: true,
            web_search: true,
            web_crawling: true,
            analyze_reference_macros: true,
            read_attached_sheet: true,
            write_cell_text: true,
            write_cells_batch: true,
            write_data_to_sheet: true,
            copy_attached_sheet: true,
            run_macro_code: true,
            final_answer: true
        };

        if (!current || typeof current !== 'object') {
            throw new Error('Planner JSON must be an object');
        }
        var type = String(current.type || '').trim();
        if (!allowedTypes[type]) {
            throw new Error('Unsupported step type: ' + type);
        }

        var macroCode = undefined;
        if (type === 'run_macro_code') {
            macroCode = extractMacroCodeFromStep(current);
        }

        return {
            id: String(current.id || ('step_' + (stepIndex + 1))).trim() || ('step_' + (stepIndex + 1)),
            type: type,
            reason: normalizeTextPayload(current.reason || ''),
            args: current.args && typeof current.args === 'object' ? current.args : {},
            macro_code: macroCode
        };
    }

    function extractMacroCodeFromStep(stepObject) {
        var current = stepObject && typeof stepObject === 'object' ? stepObject : {};
        var args = current.args && typeof current.args === 'object' ? current.args : {};
        return normalizeTextPayload(
            current.macro_code ||
            current.macroCode ||
            current.code ||
            args.macro_code ||
            args.macroCode ||
            args.code ||
            args.script ||
            ''
        );
    }

    function extractHostToolName(args) {
        var safeArgs = args && typeof args === 'object' ? args : {};
        return normalizeTextPayload(safeArgs.tool || safeArgs.name || '');
    }

    function buildToolResultMessage(stepResult) {
        function compactResultData(stepType, data) {
            var safe = data && typeof data === 'object' ? data : null;
            if (!safe) return data || null;

            if (stepType === 'web_search' || stepType === 'web_crawling') {
                var results = Array.isArray(safe.results) ? safe.results.slice(0, 3).map(function (item) {
                    var source = item && typeof item === 'object' ? item : {};
                    return {
                        title: truncateInlineText(source.title || '', 120),
                        url: truncateInlineText(source.url || '', 220),
                        snippet: truncateInlineText(source.snippet || source.text || source.excerpt || '', 220)
                    };
                }) : [];
                var errors = Array.isArray(safe.errors) ? safe.errors.slice(0, 2).map(function (item) {
                    var source = item && typeof item === 'object' ? item : {};
                    return {
                        code: normalizeTextPayload(source.code || ''),
                        message: truncateInlineText(source.message || '', 260)
                    };
                }) : [];
                return {
                    provider: normalizeTextPayload(safe.provider || ''),
                    query: truncateInlineText(safe.query || '', 160),
                    results: results,
                    errors: errors,
                    result_count: Array.isArray(safe.results) ? safe.results.length : 0,
                    error_count: Array.isArray(safe.errors) ? safe.errors.length : 0
                };
            }

            if (stepType === 'collect_context') {
                return {
                    editorType: normalizeTextPayload(safe.editorType || ''),
                    mode: normalizeTextPayload(safe.mode || ''),
                    source: normalizeTextPayload(safe.source || ''),
                    sheetName: normalizeTextPayload(safe.sheetName || ''),
                    truncated: safe.truncated === true,
                    warnings: Array.isArray(safe.warnings) ? safe.warnings.slice(0, 4) : [],
                    coverage: cloneSerializable(safe.coverage || null),
                    preview: safe.preview ? {
                        headers: Array.isArray(safe.preview.headers) ? safe.preview.headers.slice(0, 8) : [],
                        rows: Array.isArray(safe.preview.rows) ? safe.preview.rows.slice(0, 4) : []
                    } : null,
                    payload_excerpt: truncateInlineText(safe.payload || '', 420)
                };
            }

            if (stepType === 'read_document_snapshot') {
                var chunks = Array.isArray(safe.textChunks)
                    ? safe.textChunks.slice(0, 3).map(function (chunk) {
                        var sourceChunk = chunk && typeof chunk === 'object' ? chunk : {};
                        return {
                            start: Number(sourceChunk.start || 0) || 0,
                            end: Number(sourceChunk.end || 0) || 0,
                            text_excerpt: truncateInlineText(sourceChunk.text || '', 260)
                        };
                    })
                    : [];
                var coverage = safe.coverage && typeof safe.coverage === 'object' ? safe.coverage : {};
                var objects = safe.objects && typeof safe.objects === 'object' ? safe.objects : {};
                return {
                    editorType: normalizeTextPayload(safe.editorType || ''),
                    mode: normalizeTextPayload(safe.mode || ''),
                    source: normalizeTextPayload(safe.source || ''),
                    coverage: {
                        totalParagraphs: Number(coverage.totalParagraphs || 0) || 0,
                        nonEmptyParagraphs: Number(coverage.nonEmptyParagraphs || 0) || 0,
                        collectedParagraphs: Number(coverage.collectedParagraphs || 0) || 0
                    },
                    objects: {
                        imageCount: Number(objects.imageCount || 0) || 0,
                        tableCount: Number(objects.tableCount || 0) || 0
                    },
                    truncated: safe.truncated === true,
                    warnings: Array.isArray(safe.warnings) ? safe.warnings.slice(0, 6) : [],
                    textChunks: chunks,
                    payload_excerpt: truncateInlineText(safe.payload || '', 420)
                };
            }

            if (stepType === 'read_active_sheet' || stepType === 'read_sheet_range') {
                return {
                    sheetName: normalizeTextPayload(safe.sheetName || ''),
                    range: normalizeTextPayload(safe.range || ''),
                    rows: Number(safe.rows || 0) || 0,
                    cols: Number(safe.cols || 0) || 0,
                    truncated: safe.truncated === true,
                    non_empty_count: Number(safe.nonEmptyCount || 0) || 0,
                    non_empty_cells: Array.isArray(safe.nonEmptyCells)
                        ? safe.nonEmptyCells.slice(0, 16).map(function (item) {
                            var source = item && typeof item === 'object' ? item : {};
                            return {
                                address: normalizeTextPayload(source.address || ''),
                                value: truncateInlineText(source.value || '', 120)
                            };
                        })
                        : [],
                    payload_excerpt: truncateInlineText(safe.payload || '', 420)
                };
            }

            if (stepType === 'read_attached_sheet') {
                return {
                    filename: normalizeTextPayload(safe.filename || ''),
                    sheetName: normalizeTextPayload(safe.sheetName || ''),
                    rows: Number(safe.rows || 0) || 0,
                    cols: Number(safe.cols || 0) || 0,
                    truncated: safe.truncated === true,
                    payload_excerpt: truncateInlineText(safe.payload || '', 420)
                };
            }

            if (stepType === 'write_cell_text') {
                return {
                    sheetName: normalizeTextPayload(safe.sheetName || ''),
                    cell: normalizeTextPayload(safe.cell || ''),
                    mode: normalizeTextPayload(safe.mode || ''),
                    written: safe.written === true,
                    value_excerpt: truncateInlineText(safe.value || '', 120)
                };
            }

            if (stepType === 'write_cells_batch') {
                return {
                    sheetName: normalizeTextPayload(safe.sheetName || ''),
                    written: Number(safe.written || 0) || 0,
                    skipped: Number(safe.skipped || 0) || 0,
                    item_count: Number(safe.itemCount || 0) || 0
                };
            }

            if (stepType === 'write_data_to_sheet' || stepType === 'copy_attached_sheet') {
                return {
                    sheetName: normalizeTextPayload(safe.sheetName || ''),
                    writtenRows: Number(safe.writtenRows || 0) || 0,
                    writtenCols: Number(safe.writtenCols || 0) || 0
                };
            }

            if (stepType === 'list_host_tools') {
                return {
                    query: truncateInlineText(safe.query || '', 180),
                    tool_count: safe.status && Number(safe.status.toolCount || 0) || 0,
                    relevant: Array.isArray(safe.relevant)
                        ? safe.relevant.slice(0, 8).map(function (item) {
                            var source = item && typeof item === 'object' ? item : {};
                            return {
                                name: normalizeTextPayload(source.name || ''),
                                description: truncateInlineText(source.description || '', 120),
                                score: Number(source.score || 0) || 0,
                                required: Array.isArray(source.required) ? source.required.slice(0, 6) : [],
                                suggestedInput: source.suggestedInput && typeof source.suggestedInput === 'object'
                                    ? compactArbitraryResultData(source.suggestedInput, 0)
                                    : {}
                            };
                        })
                        : []
                };
            }

            if (stepType === 'call_host_tool') {
                return {
                    tool: normalizeTextPayload(safe.tool || ''),
                    input_excerpt: truncateInlineText(JSON.stringify(compactArbitraryResultData(safe.input || {}, 0) || {}), 220),
                    output_excerpt: truncateInlineText(JSON.stringify(compactArbitraryResultData(safe.output || {}, 0) || {}), 260)
                };
            }

            if (stepType === 'list_sheets') {
                return {
                    sheets: Array.isArray(safe.sheets)
                        ? safe.sheets.slice(0, 12).map(function (item) {
                            var source = item && typeof item === 'object' ? item : {};
                            return {
                                name: normalizeTextPayload(source.name || ''),
                                address: normalizeTextPayload(source.address || ''),
                                rows: Number(source.rows || 0) || 0,
                                cols: Number(source.cols || 0) || 0
                            };
                        })
                        : [],
                    activeSheet: normalizeTextPayload(safe.activeSheet || ''),
                    discovery_status: normalizeTextPayload(safe.discovery_status || '')
                };
            }

            if (stepType === 'analyze_reference_macros') {
                return {
                    methods: Array.isArray(safe.methods)
                        ? safe.methods.slice(0, 12).map(function (item) {
                            var source = item && typeof item === 'object' ? item : {};
                            return {
                                name: normalizeTextPayload(source.name || ''),
                                category: normalizeTextPayload(source.category || ''),
                                description: truncateInlineText(source.description || '', 180)
                            };
                        })
                        : [],
                    category: normalizeTextPayload(safe.category || ''),
                    methodCount: Number(safe.methodCount || 0) || 0,
                    guide_excerpt: truncateInlineText(safe.guide || safe.guide_excerpt || '', 260)
                };
            }

            if (stepType === 'generate_image_asset') {
                var rawUrl = normalizeTextPayload(safe.url || '');
                return {
                    prompt_excerpt: truncateInlineText(safe.prompt || '', 220),
                    caption: truncateInlineText(safe.caption || '', 180),
                    sectionTitle: truncateInlineText(safe.sectionTitle || '', 140),
                    inserted: safe.inserted === true,
                    url: sanitizeImageUrlForPlanner(rawUrl),
                    url_kind: isDataImageUrl(rawUrl) ? 'data_url_omitted' : (rawUrl ? 'remote_url' : '')
                };
            }

            if (stepType === 'run_macro_code') {
                return {
                    keys: Object.keys(safe).slice(0, 12),
                    payload_excerpt: truncateInlineText(JSON.stringify(compactArbitraryResultData(safe, 0) || {}), 420)
                };
            }

            if (stepType === 'get_current_time') {
                return {
                    iso: normalizeTextPayload(safe.iso || ''),
                    local: normalizeTextPayload(safe.local || ''),
                    timezone: normalizeTextPayload(safe.timezone || ''),
                    unix_ms: Number(safe.unix_ms || 0) || 0
                };
            }

            return compactArbitraryResultData(safe, 0);
        }

        return JSON.stringify({
            type: 'tool_result',
            step_id: stepResult.step_id,
            ok: !!stepResult.ok,
            data: compactResultData(stepResult.step_type, stepResult.data || null),
            logs: Array.isArray(stepResult.logs) ? stepResult.logs : [],
            error: truncateInlineText(stepResult.error || '', 320) || null,
            duration_ms: stepResult.duration_ms
        }, null, 2);
    }

    function buildPlanDecisionMessage(decision, plan, note) {
        return JSON.stringify({
            type: 'plan_decision',
            decision: normalizeTextPayload(decision || ''),
            plan: cloneSerializable(plan) || null,
            note: normalizeTextPayload(note || '')
        }, null, 2);
    }

    function isWordPlanModeActive() {
        var mode = agentState().wordPlanMode;
        return !!(mode && mode.enabled);
    }

    function isAwaitingPlanApproval() {
        return agentState().status === 'awaiting_plan' && !!agentState().pendingPlan;
    }

    function shouldForcePlanPresentation(step) {
        if (!isWordPlanModeActive()) return false;
        if (agentState().wordPlanMode.awaitingApproval || agentState().wordPlanMode.approved) return false;
        if (!step || !step.type) return false;
        return !(
            step.type === 'collect_context' ||
            step.type === 'web_search' ||
            step.type === 'web_crawling' ||
            step.type === 'present_plan' ||
            step.type === 'analyze_reference_macros'
        );
    }

    function shouldDeferFinalAnswerUntilExecution(step) {
        if (!step || step.type !== 'final_answer') return false;
        if (!isWordPlanModeActive()) return false;
        var mode = agentState().wordPlanMode || createDefaultWordPlanState();
        return mode.approved === true && Number(mode.executionCountAfterApproval || 0) < 1;
    }

    function countWordPlanExecutionStep(stepType) {
        return stepType === 'run_macro_code' || stepType === 'generate_image_asset';
    }

    function shouldDegradeWebResearchFailure(step, stepResult, consecutiveStepErrors) {
        var policy = agentState().researchPolicy || createDefaultResearchPolicy();
        if (!step || step.type !== 'web_search') return false;
        if (policy.unavailable) return true;
        if (!policy.required) return false;
        return Number(consecutiveStepErrors || 0) >= 2;
    }

    function markWebResearchUnavailable(stepResult) {
        var policy = agentState().researchPolicy || createDefaultResearchPolicy();
        policy.attempted = true;
        policy.completed = false;
        policy.failed = true;
        policy.unavailable = true;
        agentState().researchPolicy = policy;
        return 'RUNTIME_POLICY: Web research is unavailable in this run after repeated failures. Do NOT emit more web_search or web_crawling. Continue with present_plan or final_answer, explicitly note that external facts could not be verified, and do not fabricate citations.';
    }

    function getPlannerSystemMessage(options) {
        var settings = options && typeof options === 'object' ? options : {};
        var mode = settings.mode || 'default';
        var editorType = getCurrentEditorType();
        var desktopTools = settings.desktopTools && typeof settings.desktopTools === 'object'
            ? settings.desktopTools
            : buildDesktopToolsPlannerState(agentState().lastUserMessage || '');
        var researchPolicy = settings.researchPolicy && typeof settings.researchPolicy === 'object'
            ? settings.researchPolicy
            : (agentState().researchPolicy || createDefaultResearchPolicy());
        var profiles = root.agent && root.agent.plannerProfiles ? root.agent.plannerProfiles : null;
        if (profiles && typeof profiles.buildSystemMessage === 'function') {
            return profiles.buildSystemMessage(mode, {
                apiGuidePath: constants().apiGuidePath,
                policyVersion: constants().plannerPolicyVersion,
                editorType: editorType,
                webSearchEnabled: researchPolicy.webSearchEnabled,
                activeWebSearchProvider: researchPolicy.activeProvider,
                braveApiConfigured: researchPolicy.braveApiConfigured,
                exaApiConfigured: researchPolicy.exaApiConfigured,
                researchRequired: researchPolicy.required,
                researchUnavailable: researchPolicy.unavailable,
                researchQuery: researchPolicy.query,
                webSearchStatusLine: buildPlannerWebSearchStatusLine(researchPolicy),
                desktopTools: desktopTools
            });
        }
        var schemaTypes = editorType === 'word'
            ? 'collect_context|read_document_snapshot|present_plan|generate_image_asset|get_current_time|list_host_tools|call_host_tool|web_search|web_crawling|run_macro_code|analyze_reference_macros|read_attached_sheet|write_cell_text|write_cells_batch|write_data_to_sheet|copy_attached_sheet|final_answer'
            : 'list_sheets|read_active_sheet|read_sheet_range|get_current_time|list_host_tools|call_host_tool|web_search|web_crawling|run_macro_code|analyze_reference_macros|read_attached_sheet|write_cell_text|write_cells_batch|write_data_to_sheet|copy_attached_sheet|final_answer';
        return [
            editorType === 'word'
                ? 'You are an R7 Office Document macro planner (ONLYOFFICE-compatible API alias).'
                : 'You are an R7 Office Spreadsheet macro planner (ONLYOFFICE-compatible API alias).',
            editorType === 'word'
                ? 'You are running inside the {r7c}.ChatLLM plugin in the R7 Office word processor.'
                : 'You are running inside the {r7c}.ChatLLM plugin in the R7 Office spreadsheet editor.',
            editorType === 'word'
                ? 'Help the user with document work in the active R7 Office environment and rely on the current document context available in this editor.'
                : 'Help the user with spreadsheet work in the active R7 Office environment and rely on the current workbook, sheet, rows, columns, ranges, and cells available in this editor.',
            'Return only one strict JSON object, no markdown.',
            'Schema: {"id":"step_n","type":"' + schemaTypes + '","reason":"...","args":{},"macro_code":"..."}',
            'TOOL_FIRST_POLICY: For spreadsheet actions, ALWAYS try predefined tools first and use run_macro_code only if no tool can satisfy the task or tool execution failed.',
            'HOST_TOOL_SEQUENCE: If a matching native desktop tool exists, use list_host_tools if needed to inspect the runtime catalog, then call_host_tool before analyze_reference_macros or run_macro_code.',
            'MACRO_QUALITY_RULES: Use concise deterministic macros. Validate objects before method calls (null checks). Prefer Api.GetSheet/AddSheet + worksheet.GetRange(range).SetValue(data[][]) for bulk writes. Do not call unsupported methods discovered by guesswork.',
            'MACRO_ANTI_PATTERNS: avoid oSheet.Clear() if unverified; avoid probing random window/plugin objects; avoid exploratory API discovery macros when predefined tools exist.',
            'IDEAL_MACRO_STYLE (fallback only): (function(){"use strict";try{if(typeof Api==="undefined") throw new Error("API unavailable"); var sheet=Api.GetActiveSheet(); if(!sheet) throw new Error("Sheet unavailable"); sheet.GetRange("A1").SetValue("..."); return {ok:true}; }catch(e){ return {ok:false,error:String(e&&e.message?e.message:e)}; }})();',
            'When user asks for current date/time/timezone, use get_current_time before final_answer.',
            'HOST_TOOL_STEPS: list_host_tools returns the runtime desktop tool catalog. call_host_tool executes one runtime tool with args: {"tool":"tool_name","input":{...}}.',
            'To read data from an attached file (e.g., XLSX workbook), use step type "read_attached_sheet" with args: {"filename": "attached.xlsx", "sheetName": "Sheet1"}. This returns the full data array from that sheet.',
            'To write text into a single cell, use step type "write_cell_text" with args: {"sheetName":"Sheet2","cell":"B3","text":"Hello","mode":"replace|append","createSheet":true}.',
            'To write many scattered cells at once, use step type "write_cells_batch" with args: {"sheetName":"Sheet2","items":[{"cell":"A1","value":"Header"},{"cell":"B2","value":42}],"createSheet":true}.',
            'CRITICAL: To copy ALL data from an attached file into a new or existing sheet, ALWAYS use step type "copy_attached_sheet" with args: {"filename": "attached.xlsx", "sourceSheetName": "Sheet1", "targetSheetName": "Sheet2", "clearFirst": true}. This BYPASSES token limits and copies up to 100k rows instantly!',
            'To write a small 2D array of newly generated data into a sheet, use step type "write_data_to_sheet" with args: {"sheetName": "Sheet2", "data": [["A", "B"], [1, 2]], "clearFirst": true}. Do NOT use this for attached files, use copy_attached_sheet instead.',
            editorType === 'word'
                ? 'Use read_document_snapshot for robust Word document reads and use collect_context as a fallback for workbook/sheet data.'
                : 'Use list_sheets, read_active_sheet, and read_sheet_range as the primary spreadsheet read path. These steps execute predefined verified ONLYOFFICE read macros/commands inside the plugin runtime and are safer than model-generated run_macro_code for inspection.',
            editorType === 'word'
                ? 'Use collect_context only as a fallback in Word when snapshot data is unavailable.'
                : 'For spreadsheet requests about sheet contents, rows, columns, values, headers, formulas, tables, or workbook structure, call one of these spreadsheet read steps before final_answer.',
            editorType === 'word'
                ? 'Use read_document_snapshot before drafting from document context.'
                : 'If spreadsheet targeting is ambiguous, call list_sheets first. If the user asks about the current sheet, call read_active_sheet. If the user names a sheet or gives a range, call read_sheet_range.',
            editorType === 'word'
                ? 'Do not skip document reading before document-grounded reasoning.'
                : 'Do not use run_macro_code just to inspect spreadsheet contents when list_sheets, read_active_sheet, or read_sheet_range can read them.',
            'For run_macro_code, put the JavaScript source in the TOP-LEVEL "macro_code" field, not only inside args.',
            'For final_answer use args.answer, concise plain text in user language.',
            'HOST_TOOLS_STATUS: ' + desktopTools.statusLine,
            'HOST_TOOLS_POLICY: Native desktop tools are a first-class planning path when runtime catalog and execution are available.',
            desktopTools.catalogForPrompt ? ('HOST_TOOLS_CATALOG:\n' + desktopTools.catalogForPrompt) : 'HOST_TOOLS_CATALOG: unavailable',
            buildPlannerWebSearchStatusLine(researchPolicy),
            'RESEARCH_GUIDE: Use web_search or web_crawling if current external information is needed for the request.',
            researchPolicy.required && researchPolicy.query
                ? ('RESEARCH_QUERY_HINT: ' + researchPolicy.query)
                : 'RESEARCH_QUERY_HINT: none',
            researchPolicy.required
                ? 'RESEARCH_FIRST_POLICY: For this specific request, emit web_search before substantive factual writing steps.'
                : 'RESEARCH_FIRST_POLICY: not_required',
            'The local spreadsheet API reference is available from scripts/api_reference.js as window.R7_API_REFERENCE_CATALOG/window.R7_API_REFERENCE_GUIDE.',
            'If a macro fails because a method is missing or unverified, use analyze_reference_macros before retrying run_macro_code.',
            'Use guide mirror: ' + constants().apiGuidePath + '.'
        ].join('\n');
    }

    function extractFinalAnswer(step) {
        if (!step || !step.args) return '';
        return normalizeTextPayload(step.args.answer || step.args.final_answer || '');
    }

    function hasExplicitSheetTarget(messageText) {
        var text = normalizeTextPayload(messageText || '');
        if (!text) return false;
        if (/(\u043d\u0430 \u044d\u0442\u043e\u043c \u043b\u0438\u0441\u0442\u0435|\u0432 \u044d\u0442\u043e\u043c \u043b\u0438\u0441\u0442\u0435|\u0442\u0435\u043a\u0443\u0449(\u0435\u043c|\u0438\u0439) \u043b\u0438\u0441\u0442|active sheet|current sheet|this sheet)/i.test(text)) return true;
        if (/(\u043b\u0438\u0441\u0442|sheet)\s*[:\"'`]?\s*[\w\- ]{1,60}/i.test(text)) return true;
        if (/(\u0434\u0438\u0430\u043f\u0430\u0437\u043e\u043d|range)\s*[:\"'`]?\s*[a-z]{1,3}\d{1,7}(:[a-z]{1,3}\d{1,7})?/i.test(text)) return true;
        if (/\b[a-z]{1,3}\d{1,7}(:[a-z]{1,3}\d{1,7})?\b/i.test(text)) return true;
        return false;
    }

    function isSpreadsheetReadIntent(messageText) {
        var text = normalizeTextPayload(messageText || '');
        if (!text.length) return false;
        var readSignal = /(\u0447\u0442\u043e \u043d\u0430 \u043b\u0438\u0441\u0442\u0435|\u0447\u0442\u043e \u0432 \u044f\u0447\u0435\u0439\u043a|\u043a\u0430\u043a\u0438\u0435 \u0434\u0430\u043d\u043d\u044b\u0435|\u043f\u043e\u043a\u0430\u0436\u0438 \u0434\u0430\u043d\u043d\u044b\u0435|\u043f\u043e\u043a\u0430\u0436\u0438 \u043b\u0438\u0441\u0442|\u0441\u0443\u043c\u043c\u0438\u0440|\u0441\u0432\u043e\u0434\u043a|\u043f\u0440\u043e\u0430\u043d\u0430\u043b\u0438\u0437|\u0430\u043d\u0430\u043b\u0438\u0437|\u043e\u043f\u0438\u0448\u0438 \u043b\u0438\u0441\u0442|\u0442\u0430\u0431\u043b\u0438\u0446|\u044f\u0447\u0435\u0439\u043a|\u0441\u0442\u0440\u043e\u043a|\u0441\u0442\u043e\u043b\u0431\u0446|\u0437\u043d\u0430\u0447\u0435\u043d\u0438|\u0444\u043e\u0440\u043c\u0443\u043b|\u0437\u0430\u0433\u043e\u043b\u043e\u0432|what is on the sheet|what data|show the sheet|show data|sheet data|table data|analy[sz]e|summari[sz]e|inspect|review|headers?|cells?|rows?|columns?|values?|formulas?)/i.test(text);
        if (!readSignal) return false;
        var writeSignal = /(\u0441\u043e\u0437\u0434\u0430\u0439|\u0441\u043e\u0437\u0434\u0430\u0442\u044c|\u0434\u043e\u0431\u0430\u0432\u044c|\u0432\u0441\u0442\u0430\u0432\u044c|\u0437\u0430\u043f\u043e\u043b\u043d\u0438|\u0438\u0437\u043c\u0435\u043d\u0438|\u043e\u0431\u043d\u043e\u0432\u0438|\u0443\u0434\u0430\u043b\u0438|\u043e\u0444\u043e\u0440\u043c\u0438|create|insert|fill|update|delete|format|populate|write into)/i.test(text);
        return !writeSignal;
    }

    function isSpreadsheetWriteIntent(messageText) {
        var text = normalizeTextPayload(messageText || '');
        if (!text.length) return false;
        var writeSignal = /(\u0441\u0434\u0435\u043b\u0430\u0439|\u0441\u043e\u0437\u0434\u0430\u0439|\u0441\u043e\u0437\u0434\u0430\u0442\u044c|\u0434\u043e\u0431\u0430\u0432\u044c|\u0432\u0441\u0442\u0430\u0432\u044c|\u0437\u0430\u043f\u043e\u043b\u043d\u0438|\u0438\u0437\u043c\u0435\u043d\u0438|\u043e\u0431\u043d\u043e\u0432\u0438|\u043f\u0435\u0440\u0435\u043d\u0435\u0441\u0438|\u0441\u043a\u043e\u043f\u0438\u0440\u0443\u0439|\u0437\u0430\u043f\u0438\u0448\u0438|\u0432\u043f\u0438\u0448\u0438|\u043e\u0444\u043e\u0440\u043c\u0438|\u043f\u043e\u0434\u0433\u043e\u0442\u043e\u0432\u044c|append|create|insert|fill|update|delete|format|populate|write|copy|move|paste|report)/i.test(text);
        var sheetSignal = /(\u043b\u0438\u0441\u0442|sheet|\u044f\u0447\u0435\u0439\u043a|cell|\u0442\u0430\u0431\u043b\u0438\u0446|workbook|range|\u0434\u0438\u0430\u043f\u0430\u0437\u043e\u043d)/i.test(text);
        return writeSignal && sheetSignal;
    }

    function isSpreadsheetWorkbookDiscoveryIntent(messageText) {
        var text = normalizeTextPayload(messageText || '');
        if (!text.length) return false;
        return /(\u043a\u0430\u043a\u0438\u0435 \u043b\u0438\u0441\u0442\u044b|\u0441\u043f\u0438\u0441\u043e\u043a \u043b\u0438\u0441\u0442\u043e\u0432|\u043f\u0435\u0440\u0435\u0447\u0438\u0441\u043b\u0438 \u043b\u0438\u0441\u0442\u044b|list sheets|which sheets|workbook structure|tabs)/i.test(text);
    }

    function extractSpreadsheetRangeTarget(messageText) {
        var match = normalizeTextPayload(messageText || '').match(/\b([A-Za-z]{1,3}\d{1,7}:[A-Za-z]{1,3}\d{1,7})\b/);
        return match && match[1] ? String(match[1]).toUpperCase() : '';
    }

    function extractSpreadsheetSheetName(messageText) {
        var text = normalizeTextPayload(messageText || '');
        var quoted = text.match(/["«“']([^"»”']{1,80})["»”']/);
        if (quoted && quoted[1]) return String(quoted[1]).trim();
        var named = text.match(/(?:\u043b\u0438\u0441\u0442|sheet)\s+([A-Za-z\u0410-\u042f\u0430-\u044f\u0401\u04510-9_\-]{1,40})/i);
        if (!named || !named[1]) return '';
        var candidate = String(named[1]).trim();
        if (/^(?:\u0430\u043a\u0442\u0438\u0432\u043d\u044b\u0439|\u0442\u0435\u043a\u0443\u0449\u0438\u0439|\u044d\u0442\u043e\u043c|\u0432|\u043d\u0430|\u043f\u043e|\u0438|current|active|this)$/i.test(candidate)) {
            return '';
        }
        return candidate;
    }

    function hasExecutedSpreadsheetReadStep(messageText) {
        var steps = Array.isArray(agentState().steps) ? agentState().steps : [];
        var expectsWorkbookList = isSpreadsheetWorkbookDiscoveryIntent(messageText || '');
        for (var i = steps.length - 1; i >= 0; i -= 1) {
            var step = steps[i] || {};
            if (expectsWorkbookList && step.type === 'list_sheets') return true;
            if (step.type === 'read_active_sheet' || step.type === 'read_sheet_range') return true;
        }
        return false;
    }

    function hasExecutedStructuredSpreadsheetToolStep() {
        var steps = Array.isArray(agentState().steps) ? agentState().steps : [];
        for (var i = steps.length - 1; i >= 0; i -= 1) {
            var type = normalizeTextPayload(steps[i] && steps[i].type || '');
            if (!type) continue;
            if (type === 'run_macro_code' || type === 'final_answer') continue;
            if (type === 'analyze_reference_macros') continue;
            return true;
        }
        return false;
    }

    function hasExecutedSpreadsheetWriteToolStep() {
        var steps = Array.isArray(agentState().steps) ? agentState().steps : [];
        for (var i = steps.length - 1; i >= 0; i -= 1) {
            var type = normalizeTextPayload(steps[i] && steps[i].type || '');
            if (type === 'write_cell_text' || type === 'write_cells_batch' || type === 'write_data_to_sheet' || type === 'copy_attached_sheet') {
                return true;
            }
        }
        return false;
    }

    function hasExecutedHostToolStep() {
        var steps = Array.isArray(agentState().steps) ? agentState().steps : [];
        for (var i = steps.length - 1; i >= 0; i -= 1) {
            var type = normalizeTextPayload(steps[i] && steps[i].type || '');
            if (type === 'call_host_tool') return true;
        }
        return false;
    }

    function getLatestAttachedWorkbookMeta() {
        var history = getConversationHistoryForPlanner();
        if (!Array.isArray(history) || !history.length) return null;
        for (var i = history.length - 1; i >= 0; i -= 1) {
            var message = history[i] || {};
            var attachments = Array.isArray(message.attachments) ? message.attachments : [];
            for (var j = attachments.length - 1; j >= 0; j -= 1) {
                var att = attachments[j] || {};
                var summary = att.contextSummary && typeof att.contextSummary === 'object' ? att.contextSummary : null;
                if (!summary || summary.kind !== 'xlsx_workbook') continue;
                var sheetNames = Array.isArray(summary.sheetNames) ? summary.sheetNames.filter(Boolean) : [];
                return {
                    filename: normalizeTextPayload(att.name || ''),
                    sheetNames: sheetNames,
                    defaultSourceSheet: normalizeTextPayload(sheetNames[0] || 'Sheet1') || 'Sheet1'
                };
            }
        }
        return null;
    }

    function isAttachedWorkbookCopyIntent(messageText) {
        var text = normalizeTextPayload(messageText || '');
        if (!text.length) return false;
        var attachmentSignal = /(прикрепл|вложен|attached|attachment|xlsx|excel)/i.test(text);
        var copySignal = /(скопир|перенес|встав|добав|import|copy|move|insert|paste)/i.test(text);
        var sheetSignal = /(лист|sheet|таблиц|workbook)/i.test(text);
        return attachmentSignal && copySignal && sheetSignal;
    }

    function extractTargetSheetNameForCopy(messageText) {
        var text = normalizeTextPayload(messageText || '');
        var explicit = text.match(/(?:в|to)\s+(?:лист|sheet)\s*([A-Za-zА-Яа-яЁё0-9_\-]{1,40})/i);
        if (explicit && explicit[1]) return String(explicit[1]).trim();
        var quoted = text.match(/(?:в|to)\s+["«“']([^"»”']{1,80})["»”']/i);
        if (quoted && quoted[1]) return String(quoted[1]).trim();
        return 'Sheet2';
    }

    function hasExecutedAttachedWorkbookCopyStep() {
        var steps = Array.isArray(agentState().steps) ? agentState().steps : [];
        for (var i = steps.length - 1; i >= 0; i -= 1) {
            var step = steps[i] || {};
            if (step.type === 'copy_attached_sheet') return true;
        }
        return false;
    }

    function isAttachedWorkbookCopyStrictMode() {
        if (getCurrentEditorType() !== 'cell') return false;
        if (!isAttachedWorkbookCopyIntent(agentState().lastUserMessage || '')) return false;
        var attached = getLatestAttachedWorkbookMeta();
        return !!(attached && attached.filename);
    }

    function shouldRestrictStepToAttachedCopyFlow(step) {
        if (!step || typeof step !== 'object') return false;
        if (!isAttachedWorkbookCopyStrictMode()) return false;
        if (hasExecutedAttachedWorkbookCopyStep()) return false;
        return !(step.type === 'copy_attached_sheet' || step.type === 'final_answer');
    }

    function shouldForceHostToolDiscoveryFirst(messageText) {
        var text = normalizeTextPayload(messageText || '');
        if (!text.length) return false;
        if (isExplicitMacroRequest(text)) return false;
        if (getCurrentEditorType() !== 'cell') return false;
        var desktopState = buildDesktopToolsPlannerState(text);
        if (!desktopState.enabledForPlanner) return false;
        if (isAttachedWorkbookCopyIntent(text)) return false;
        if (!isSpreadsheetWriteIntent(text)) return false;
        if (hasExecutedSpreadsheetWriteToolStep() || hasExecutedHostToolStep()) return false;
        return true;
    }

    function buildForcedSpreadsheetReadStep(messageText, stepIndex, intent) {
        var text = normalizeTextPayload(messageText || '');
        var range = extractSpreadsheetRangeTarget(text);
        var sheetName = extractSpreadsheetSheetName(text);
        var stepIntent = intent || 'read_before_answer';
        if (isSpreadsheetWorkbookDiscoveryIntent(text)) {
            return {
                id: 'step_' + (stepIndex + 1),
                type: 'list_sheets',
                reason: 'List workbook sheets with the predefined spreadsheet read macro before answering.',
                args: {
                    editorType: 'cell',
                    intent: stepIntent,
                    forceRefresh: true
                }
            };
        }
        if (range || (sheetName && !/(\u0442\u0435\u043a\u0443\u0449|current|active|this sheet)/i.test(text))) {
            return {
                id: 'step_' + (stepIndex + 1),
                type: 'read_sheet_range',
                reason: 'Read the requested spreadsheet target with the predefined spreadsheet read macro before answering.',
                args: {
                    editorType: 'cell',
                    sheetName: sheetName,
                    range: range,
                    intent: stepIntent,
                    forceRefresh: true
                }
            };
        }
        return {
            id: 'step_' + (stepIndex + 1),
            type: 'read_active_sheet',
            reason: 'Read the active sheet with the predefined spreadsheet read macro before answering.',
            args: {
                editorType: 'cell',
                intent: stepIntent,
                forceRefresh: true
            }
        };
    }

    function shouldForceSpreadsheetReadBeforeAnswer(step) {
        if (!step || step.type !== 'final_answer') return false;
        if (getCurrentEditorType() !== 'cell') return false;
        if (hasPendingPostMacroVerification()) return false;
        if (!isSpreadsheetReadIntent(agentState().lastUserMessage || '')) return false;
        return !hasExecutedSpreadsheetReadStep(agentState().lastUserMessage || '');
    }

    function shouldBlockGeneratedSpreadsheetInspectionMacro(step) {
        if (!step || step.type !== 'run_macro_code') return false;
        if (getCurrentEditorType() !== 'cell') return false;
        if (isAttachedWorkbookCopyIntent(agentState().lastUserMessage || '') && !hasExecutedAttachedWorkbookCopyStep()) {
            return true;
        }
        if (!isSpreadsheetReadIntent(agentState().lastUserMessage || '')) return false;
        return !hasExecutedSpreadsheetReadStep(agentState().lastUserMessage || '');
    }

    function shouldEnforceSpreadsheetToolsBeforeMacros(step) {
        if (!step || step.type !== 'run_macro_code') return false;
        if (getCurrentEditorType() !== 'cell') return false;
        if (isExplicitMacroRequest(agentState().lastUserMessage || '')) return false;
        if (!isSpreadsheetWriteIntent(agentState().lastUserMessage || '')) return false;
        return !(hasExecutedSpreadsheetWriteToolStep() || hasExecutedHostToolStep());
    }

    function shouldEnforceSpreadsheetToolsBeforeReference(step) {
        if (!step || step.type !== 'analyze_reference_macros') return false;
        if (getCurrentEditorType() !== 'cell') return false;
        if (isExplicitMacroRequest(agentState().lastUserMessage || '')) return false;
        if (!isSpreadsheetWriteIntent(agentState().lastUserMessage || '')) return false;
        return !(hasExecutedSpreadsheetWriteToolStep() || hasExecutedHostToolStep());
    }

    function shouldRequireCallHostToolAfterDiscovery(step) {
        if (!step || typeof step !== 'object') return false;
        if (getCurrentEditorType() !== 'cell') return false;
        if (isExplicitMacroRequest(agentState().lastUserMessage || '')) return false;
        if (!isSpreadsheetWriteIntent(agentState().lastUserMessage || '')) return false;
        if (hasExecutedHostToolStep()) return false;
        var discovery = agentState().lastHostToolDiscovery && typeof agentState().lastHostToolDiscovery === 'object'
            ? agentState().lastHostToolDiscovery
            : null;
        if (!discovery || !Array.isArray(discovery.relevant) || !discovery.relevant.length) return false;
        return step.type !== 'call_host_tool';
    }

    function buildHostToolCallRequiredPrompt() {
        var discovery = agentState().lastHostToolDiscovery && typeof agentState().lastHostToolDiscovery === 'object'
            ? agentState().lastHostToolDiscovery
            : null;
        var relevant = discovery && Array.isArray(discovery.relevant) ? discovery.relevant.slice(0, 3) : [];
        var lines = [
            'RUNTIME_POLICY: A relevant runtime host tool has already been discovered. Emit call_host_tool as the next step instead of analyze_reference_macros, run_macro_code, or local spreadsheet write tools.'
        ];
        relevant.forEach(function (item, index) {
            var source = item && typeof item === 'object' ? item : {};
            var suggested = source.suggestedInput && typeof source.suggestedInput === 'object'
                ? JSON.stringify(source.suggestedInput)
                : '{}';
            lines.push((index + 1) + '. tool=' + normalizeTextPayload(source.name || '') + '; required=' + (Array.isArray(source.required) ? source.required.join(',') : 'none') + '; suggested_input=' + suggested);
        });
        return lines.join('\n');
    }

    function getForcedInitialStep(userMessage, stepIndex) {
        var editorType = getCurrentEditorType();
        var researchPolicy = agentState().researchPolicy || createDefaultResearchPolicy();
        if (editorType === 'word' && isImageGenerationRequest(userMessage)) {
            if (stepIndex === 0) {
                return {
                    id: 'step_' + (stepIndex + 1),
                    type: 'read_document_snapshot',
                    reason: 'Capture a robust Word snapshot before building an image prompt.',
                    args: {
                        editorType: 'word',
                        intent: 'image_prompt',
                        maxChars: 12000
                    }
                };
            }
            if (stepIndex === 1) {
                return {
                    id: 'step_' + (stepIndex + 1),
                    type: 'final_answer',
                    reason: 'Generate an image that matches the current document context and the user request.',
                    args: {
                        answer: '[GENERATE_IMAGE: ' + buildImageGenerationPrompt(userMessage) + ']'
                    }
                };
            }
            return null;
        }
        if (stepIndex === 0 && isImageGenerationRequest(userMessage)) {
            return {
                id: 'step_' + (stepIndex + 1),
                type: 'final_answer',
                reason: 'Generate an image that matches the user request.',
                args: {
                    answer: '[GENERATE_IMAGE: ' + buildImageGenerationPrompt(userMessage) + ']'
                }
            };
        }
        if (editorType === 'word') {
            if (isExplicitMacroRequest(userMessage)) return null;
            if (stepIndex === 0) {
                return {
                    id: 'step_' + (stepIndex + 1),
                    type: 'read_document_snapshot',
                    reason: 'Collect a robust Word document snapshot before planning.',
                    args: {
                        editorType: 'word',
                        intent: 'initial_context',
                        maxChars: 12000
                    }
                };
            }
            if (stepIndex === 1 && researchPolicy.required && researchPolicy.query) {
                return {
                    id: 'step_' + (stepIndex + 1),
                    type: 'web_search',
                    reason: 'Research the topic on the web before drafting the document because web search is configured for the user.',
                    args: {
                        query: researchPolicy.query
                    }
                };
            }
            return null;
        }
        if (stepIndex !== 0) return null;
        if (editorType === 'slide') return null;

        if (editorType === 'cell' && isAttachedWorkbookCopyIntent(userMessage || '')) {
            var attached = getLatestAttachedWorkbookMeta();
            if (attached && attached.filename) {
                return {
                    id: 'step_' + (stepIndex + 1),
                    type: 'copy_attached_sheet',
                    reason: 'Copy all rows from the latest attached workbook sheet into target sheet using verified native tool.',
                    args: {
                        filename: attached.filename,
                        sourceSheetName: attached.defaultSourceSheet,
                        targetSheetName: extractTargetSheetNameForCopy(userMessage || ''),
                        clearFirst: true
                    }
                };
            }
        }

        if (editorType === 'cell' && shouldForceHostToolDiscoveryFirst(userMessage || '')) {
            return buildForcedInitialHostToolStep(userMessage || '', stepIndex);
        }

        if (isSpreadsheetReadIntent(userMessage)) {
            return buildForcedSpreadsheetReadStep(userMessage, stepIndex, 'initial_context');
        }

        if (isExplicitMacroRequest(userMessage)) return null;

        return null;
    }

    function isResearchBlockedForStep(step) {
        // Fix: Disable mandatory research blocking entirely as requested by the user.
        // It was causing too many false positives when web search returned data but wasn't marked correctly.
        return false;
    }

    function buildResearchBlockedResult() {
        var policy = agentState().researchPolicy || createDefaultResearchPolicy();
        return {
            ok: false,
            data: null,
            logs: ['research_required_before_write'],
            error: policy.attempted
                ? 'Mandatory web_search must complete successfully before writing the document.'
                : 'Mandatory web_search must run before writing the document.'
        };
    }

    function buildFastPathQueueForMessage(userMessage, startStepIndex) {
        var text = normalizeTextPayload(userMessage || '');
        if (!text.length) return [];
        var fastPath = root.agent && root.agent.fastPath ? root.agent.fastPath : null;
        if (!fastPath) {
            return [];
        }
        var queue = [];
        if (typeof fastPath.createQueueForMessage === 'function') {
            queue = fastPath.createQueueForMessage(text, startStepIndex || 0);
        } else {
            if (hasExplicitSheetTarget(text)) return [];
            if (typeof fastPath.isHeavyWorkbookRequest !== 'function' || typeof fastPath.createHeavyWorkbookQueue !== 'function') {
                return [];
            }
            if (!fastPath.isHeavyWorkbookRequest(text)) return [];
            queue = fastPath.createHeavyWorkbookQueue(text, startStepIndex || 0);
        }
        if (!Array.isArray(queue)) return [];
        return queue.map(function (item, index) {
            var step = item && typeof item === 'object' ? item : {};
            var args = step.args && typeof step.args === 'object' ? Object.assign({}, step.args) : {};
            args.fast_path = true;
            return {
                id: normalizeTextPayload(step.id || ('step_' + ((startStepIndex || 0) + index + 1))),
                type: normalizeTextPayload(step.type || ''),
                reason: normalizeTextPayload(step.reason || 'Fast-path step'),
                args: args,
                macro_code: step.macro_code ? sanitizeMacroCode(step.macro_code) : undefined
            };
        }).filter(function (step) { return !!step.type; });
    }

    async function plannerNextStep(plannerMessages, stepIndex) {
        var providers = getProvidersBridge();
        var openrouter = getOpenRouterBridge();
        var chat = getChatService();
        agentState().abortController = new AbortController();
        var runtimeSettings = chat && typeof chat.loadRuntimeSettings === 'function'
            ? chat.loadRuntimeSettings()
            : { model: constants().defaultModel };
        var activeProvider = normalizeTextPayload(
            runtimeSettings && (runtimeSettings.activeProvider || runtimeSettings.provider) || 'openrouter'
        ).toLowerCase() || 'openrouter';
        var settingsService = getSettingsService();
        var providerConfig = settingsService && typeof settingsService.getProviderConfig === 'function'
            ? settingsService.getProviderConfig(runtimeSettings, activeProvider)
            : null;
        var config = {
            provider: activeProvider,
            activeProvider: activeProvider,
            model: providerConfig && providerConfig.model ? providerConfig.model : (runtimeSettings.model || constants().defaultModel),
            reasoningEffort: providerConfig && providerConfig.reasoningEffort ? providerConfig.reasoningEffort : 'medium',
            temperature: 0.1
        };
        if (!providers || typeof providers.chatRequest !== 'function') {
            if (!openrouter || typeof openrouter.chatRequest !== 'function' || activeProvider !== 'openrouter') {
                throw new Error('Provider registry is unavailable for planner provider "' + activeProvider + '"');
            }
        }
        var maxAttempts = 3;

        try {
            for (var attempt = 1; attempt <= maxAttempts; attempt += 1) {
                var requestMessages = Array.isArray(plannerMessages) ? plannerMessages.slice() : [];
                if (attempt > 1) {
                    requestMessages.push({
                        role: 'user',
                        content: 'Previous planner output was invalid JSON. Return exactly ONE short valid JSON step object only. Keep args.answer concise.'
                    });
                }
                var plannerMode = attempt > 1 ? 'diagnostic' : (agentState().fastPathActive ? 'macro_batch' : 'default');
                try {
                    var desktopToolsState = buildDesktopToolsPlannerState(agentState().lastUserMessage || '');
                    var plannerCallStartedAt = Date.now();
                    var response = providers && typeof providers.chatRequest === 'function'
                        ? await providers.chatRequest(
                            requestMessages,
                            getPlannerSystemMessage({
                                mode: plannerMode,
                                researchPolicy: agentState().researchPolicy,
                                desktopTools: desktopToolsState
                            }),
                            false,
                            config,
                            agentState().abortController.signal
                        )
                        : await openrouter.chatRequest(
                            requestMessages,
                            getPlannerSystemMessage({
                                mode: plannerMode,
                                researchPolicy: agentState().researchPolicy,
                                desktopTools: desktopToolsState
                            }),
                            false,
                            config,
                            agentState().abortController.signal
                        );
                    addModelReasoningTraceFromPlanner(response, {
                        runtimeSettings: runtimeSettings,
                        provider: activeProvider,
                        durationMs: Date.now() - plannerCallStartedAt
                    });
                    var content = response && response.data && response.data[0] ? response.data[0].content : '';
                    return normalizePlannerStep(parsePlannerResponse(content), stepIndex);
                } catch (error) {
                    if (error && error.name === 'AbortError') throw error;
                    if ((!isTransientPlannerError(error) && !isPlannerOutputFormatError(error)) || attempt >= maxAttempts) throw error;
                    await sleep(700 * attempt);
                }
            }
            throw new Error('Planner request failed');
        } finally {
            agentState().abortController = null;
        }
    }

    async function executeMacroCode(macroCode) {
        var host = getHostBridge();
        var code = sanitizeMacroCode(macroCode);
        var expressionLikeCode = /^\s*\(\s*function\b[\s\S]*\)\s*\(\s*\)\s*;?\s*$/.test(code) || /^\s*\(\s*\([^)]*\)\s*=>[\s\S]*\)\s*\(\s*\)\s*;?\s*$/.test(code);
        var macroBody = expressionLikeCode ? ('return ' + code.replace(/;\s*$/, '') + ';') : code;
        if (!code.length) {
            return { ok: false, data: null, logs: [], error: 'macro_code is empty' };
        }
        if (code.length > constants().agentLimits.maxMacroCodeChars) {
            return { ok: false, data: null, logs: [], error: 'macro_code exceeds max length' };
        }
        if (!host || typeof host.callEditorCommand !== 'function') {
            return { ok: false, data: null, logs: [], error: 'Host bridge is unavailable' };
        }
        if (getCurrentEditorType() === 'word' && /\.SetHeading\s*\(/i.test(code)) {
            return {
                ok: false,
                data: null,
                logs: [],
                error: 'Macro validation failed: SetHeading is not a verified Word API method. Use Api.CreateParagraph() + AddText() + SetBold()/SetFontSize() for headings instead.'
            };
        }

        var commandString = [
            "var output = { ok: false, data: null, logs: [], error: null };",
            "function toSerializable(value, depth) {",
            "    var currentDepth = typeof depth === 'number' ? depth : 0;",
            "    if (currentDepth > 5) return '[depth_limit]';",
            "    if (value === null || value === undefined) return null;",
            "    if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return value;",
            "    if (Array.isArray(value)) return value.slice(0, 200).map(function(item) { return toSerializable(item, currentDepth + 1); });",
            "    if (typeof value === 'object') {",
            "        var normalized = {};",
            "        Object.keys(value).slice(0, 80).forEach(function(key) {",
            "            normalized[key] = toSerializable(value[key], currentDepth + 1);",
            "        });",
            "        return normalized;",
            "    }",
            "    return String(value);",
            "}",
            "try {",
            "    console.log('macroUserFunc: Starting evaluation for macro...');",
            "    var macroUserFunc = function() {",
            "        " + macroBody,
            "    };",
            "    var evalResult = macroUserFunc();",
            "    console.log('macroUserFunc: Finished successfully', evalResult);",
            "    output.ok = true;",
            "    output.data = toSerializable(evalResult, 0);",
            "    return JSON.stringify(output);",
            "} catch (error) {",
            "    console.error('macroUserFunc: Caught error during execution', error);",
            "    output.error = String(error && error.message ? error.message : error);",
            "    return JSON.stringify(output);",
            "}"
        ].join('\n');

        var compiledCommand;
        try {
            compiledCommand = new Function(commandString);
        } catch (compileError) {
            return { ok: false, data: null, logs: [], error: 'Macro compilation error: ' + compileError.message };
        }

        var result = await host.callEditorCommand(compiledCommand);
        if (typeof result === 'string') {
            try {
                result = JSON.parse(result);
            } catch (e) {
                // Return as data if it's somehow not JSON, though it should be.
                result = { ok: true, data: result, logs: [], error: null };
            }
        }
        return result || { ok: false, data: null, logs: [], error: 'Macro execution failed' };
    }

    async function executeListSheetsStep() {
        var context = getContextService();
        var meta = await context.getWorkbookSheets({ force: true });
        return {
            ok: true,
            data: meta,
            logs: [],
            error: null
        };
    }

    async function executeCollectContextStep(args) {
        var context = getContextService();
        if (!context || typeof context.collectContext !== 'function' || typeof context.discoverAvailableContext !== 'function') {
            return { ok: false, data: null, logs: [], error: 'Context service is unavailable' };
        }
        var safeArgs = args && typeof args === 'object' ? args : {};
        var requestedMode = normalizeTextPayload(safeArgs.mode || '').toLowerCase() === 'discover' ? 'discover' : 'collect';
        var editorType = safeArgs.editorType === 'word' ? 'word' : (safeArgs.editorType === 'cell' ? 'cell' : getCurrentEditorType());
        var result = requestedMode === 'discover'
            ? await context.discoverAvailableContext({
                mode: 'discover',
                editorType: editorType,
                forceRefresh: safeArgs.forceRefresh === true,
                intent: safeArgs.intent || ''
            })
            : await context.collectContext({
                mode: 'collect',
                editorType: editorType,
                sheetName: safeArgs.sheetName || '',
                intent: safeArgs.intent || '',
                maxChars: safeArgs.maxChars,
                maxRowsPerChunk: safeArgs.maxRowsPerChunk,
                maxColsPerChunk: safeArgs.maxColsPerChunk,
                maxParagraphsPerChunk: safeArgs.maxParagraphsPerChunk,
                maxInternalChunks: safeArgs.maxInternalChunks,
                forceRefresh: safeArgs.forceRefresh === true
            });

        if (requestedMode === 'collect' && editorType === 'word') {
            agentState().lastReadDocumentText = normalizeTextPayload(result && result.payload || '');
        }

        return {
            ok: true,
            data: result,
            logs: [requestedMode === 'discover' ? 'context_discovered' : 'context_collected'],
            error: null
        };
    }

    function spreadsheetColLetters(value) {
        var num = Math.max(1, Number(value) || 1);
        var output = '';
        while (num > 0) {
            var mod = (num - 1) % 26;
            output = String.fromCharCode(65 + mod) + output;
            num = Math.floor((num - 1) / 26);
        }
        return output || 'A';
    }

    function parseSpreadsheetStart(address) {
        var text = normalizeTextPayload(address || '').split('!').pop().replace(/\$/g, '');
        var match = text.match(/^([A-Z]+)(\d+)/i);
        if (!match) {
            return { row: 1, col: 1 };
        }
        var letters = String(match[1] || '').toUpperCase();
        var col = 0;
        for (var i = 0; i < letters.length; i += 1) {
            col = col * 26 + (letters.charCodeAt(i) - 64);
        }
        return {
            row: Math.max(1, Number(match[2]) || 1),
            col: Math.max(1, col || 1)
        };
    }

    function buildSpreadsheetNonEmptyPreview(values, address, maxItems) {
        var rows = Array.isArray(values) ? values : [];
        var start = parseSpreadsheetStart(address || '');
        var limit = Math.max(1, Number(maxItems || 24) || 24);
        var cells = [];
        var total = 0;

        for (var r = 0; r < rows.length; r += 1) {
            var row = Array.isArray(rows[r]) ? rows[r] : [rows[r]];
            for (var c = 0; c < row.length; c += 1) {
                var value = row[c];
                var normalized = normalizeTextPayload(value === undefined || value === null ? '' : value);
                if (!normalized.length) continue;
                total += 1;
                if (cells.length >= limit) continue;
                cells.push({
                    address: spreadsheetColLetters(start.col + c) + String(start.row + r),
                    value: normalized
                });
            }
        }

        return {
            total: total,
            cells: cells,
            truncated: total > cells.length
        };
    }

    async function executeReadActiveSheetStep() {
        var context = getContextService();
        var activeSheet = await context.collectActiveSheetContext({ force: true });
        if (!activeSheet) {
            return { ok: false, data: null, logs: [], error: 'Active sheet is not available' };
        }
        var payload = context.formatTablePayload(activeSheet.values, {
            maxRows: constants().contextLimits.maxRowsPerSheet,
            maxCols: constants().contextLimits.maxColsPerSheet,
            maxChars: constants().contextLimits.maxSheetChars
        });
        var nonEmptyPreview = buildSpreadsheetNonEmptyPreview(activeSheet.values, activeSheet.address, 24);
        return {
            ok: true,
            data: {
                sheetName: activeSheet.sheetName || 'Active sheet',
                range: activeSheet.address || '',
                rows: payload.totalRows,
                cols: payload.totalCols,
                truncated: payload.truncated,
                payload: payload.text,
                nonEmptyCount: nonEmptyPreview.total,
                nonEmptyCells: nonEmptyPreview.cells,
                nonEmptyCellsTruncated: nonEmptyPreview.truncated
            },
            logs: [],
            error: null
        };
    }

    async function executeReadAttachedSheetStep(args) {
        var safeArgs = args && typeof args === 'object' ? args : {};
        var filename = normalizeTextPayload(safeArgs.filename || '');
        var sheetName = normalizeTextPayload(safeArgs.sheetName || '');
        
        var threadStore = getChatThreadStore();
        if (!threadStore) return { ok: false, data: null, logs: [], error: 'Thread store unavailable' };

        var activeThread = threadStore.getActiveThread();
        if (!activeThread) return { ok: false, data: null, logs: [], error: 'Active thread unavailable' };

        var targetAttachment = null;
        for (var i = activeThread.messages.length - 1; i >= 0; i--) {
            var msg = activeThread.messages[i];
            if (Array.isArray(msg.attachments)) {
                for (var j = 0; j < msg.attachments.length; j++) {
                    var att = msg.attachments[j];
                    if (att.name === filename && att.attachmentType === 'file' && att.contextSummary && att.contextSummary.kind === 'xlsx_workbook') {
                        targetAttachment = att;
                        break;
                    }
                }
            }
            if (targetAttachment) break;
        }

        if (!targetAttachment) {
            return { ok: false, data: null, logs: [], error: 'Attached workbook not found in recent messages: ' + filename };
        }

        var attachmentService = root.features.chatAttachments;
        if (!attachmentService || typeof attachmentService.getSessionBlob !== 'function') {
            return { ok: false, data: null, logs: [], error: 'Attachment session cache unavailable' };
        }

        var blob = attachmentService.getSessionBlob(targetAttachment.id);
        if (!blob) {
            return { ok: false, data: null, logs: [], error: 'Attachment content is not available in memory. Please re-attach the file.' };
        }

        if (!root.shared || !root.shared.xlsxParser) {
            return { ok: false, data: null, logs: [], error: 'XLSX parser is unavailable' };
        }

        try {
            var summary = await root.shared.xlsxParser.parseXlsxFromBlob(blob, { full: true });
            var sheet = summary.sheets.find(function(s) { return s.name === sheetName; });
            if (!sheet) {
                return { ok: false, data: null, logs: [], error: 'Sheet not found in attachment: ' + sheetName };
            }
            
            var context = getContextService();
            var payload = context && typeof context.formatTablePayload === 'function' ? context.formatTablePayload(sheet.values, {
                maxRows: constants().contextLimits.maxRowsPerSheet,
                maxCols: constants().contextLimits.maxColsPerSheet,
                maxChars: constants().contextLimits.maxSheetChars
            }) : {
                text: sheet.preview,
                totalRows: sheet.rows,
                totalCols: sheet.cols,
                truncated: sheet.truncated
            };

            return {
                ok: true,
                data: {
                    filename: filename,
                    sheetName: sheet.name,
                    rows: payload.totalRows,
                    cols: payload.totalCols,
                    payload: payload.text,
                    truncated: payload.truncated
                },
                logs: [],
                error: null
            };
        } catch (error) {
            return { ok: false, data: null, logs: [], error: 'Failed to read attached sheet: ' + String(error) };
        }
    }

    async function executeWriteDataToSheetStep(args) {
        var safeArgs = args && typeof args === 'object' ? args : {};
        var sheetName = normalizeTextPayload(safeArgs.sheetName || 'Sheet1');
        var data = Array.isArray(safeArgs.data) ? safeArgs.data : [];
        var clearFirst = safeArgs.clearFirst === true;

        function normalizeMatrix(rows) {
            var source = Array.isArray(rows) ? rows : [];
            var maxCols = 0;
            for (var i = 0; i < source.length; i += 1) {
                var row = Array.isArray(source[i]) ? source[i] : [source[i]];
                if (row.length > maxCols) maxCols = row.length;
            }
            maxCols = Math.max(1, maxCols);
            return source.map(function (item) {
                var row = Array.isArray(item) ? item.slice() : [item];
                while (row.length < maxCols) row.push('');
                return row;
            });
        }

        if (!data.length) {
            return { ok: false, data: null, logs: [], error: 'Data array is empty' };
        }

        data = normalizeMatrix(data);

        var host = getHostBridge();
        if (!host || typeof host.callEditorCommand !== 'function') {
            return { ok: false, data: null, logs: [], error: 'Host bridge is unavailable' };
        }

        var macroFunction = function () {
            'use strict';
            try {
                function colName(index) {
                    var n = index;
                    var name = "";
                    while (n > 0) {
                        var rem = (n - 1) % 26;
                        name = String.fromCharCode(65 + rem) + name;
                        n = Math.floor((n - 1) / 26);
                    }
                    return name;
                }

                var sName = Asc.scope.sheetName;
                var sData = Asc.scope.data;
                var sClear = Asc.scope.clearFirst;

                var sheet = typeof Api.GetSheet === "function" ? Api.GetSheet(sName) : null;
                if (!sheet) {
                    if (typeof Api.AddSheet !== "function") throw new Error("Api.AddSheet is unavailable");
                    sheet = Api.AddSheet(sName);
                }
                if (!sheet) throw new Error("Failed to create or resolve sheet: " + sName);
                
                if (typeof sheet.SetActive === "function") sheet.SetActive();

                if (sClear && typeof sheet.GetUsedRange === "function") {
                    var used = sheet.GetUsedRange();
                    if (used) {
                        if (typeof used.ClearContents === "function") used.ClearContents();
                        else if (typeof used.Clear === "function") used.Clear();
                    }
                }

                var rowCount = sData.length;
                var colCount = 0;
                for (var i = 0; i < sData.length; i += 1) {
                    var row = Array.isArray(sData[i]) ? sData[i] : [sData[i]];
                    if (row.length > colCount) colCount = row.length;
                }
                colCount = Math.max(1, colCount);

                var endAddress = colName(colCount) + String(rowCount);
                var targetRange = sheet.GetRange("A1:" + endAddress);
                targetRange.SetValue(sData);

                return JSON.stringify({ ok: true, data: { writtenRows: rowCount, writtenCols: colCount, sheetName: sName } });
            } catch (e) {
                return JSON.stringify({ ok: false, error: String(e && e.message ? e.message : e) });
            }
        };

        try {
            var result = await host.callEditorCommand(macroFunction, {
                sheetName: sheetName,
                data: data,
                clearFirst: clearFirst
            }, { timeoutMs: 300000 });

            var parsedResult = parseAgentJsonPayload(result, null);
            if (parsedResult && parsedResult.ok) {
                return {
                    ok: true,
                    data: parsedResult.data,
                    logs: ['data_written_successfully'],
                    error: null
                };
            } else {
                return { ok: false, data: null, logs: [], error: parsedResult && parsedResult.error ? parsedResult.error : 'Execution failed' };
            }
        } catch (error) {
            return { ok: false, data: null, logs: [], error: String(error) };
        }
    }

    async function executeWriteCellTextStep(args) {
        var safeArgs = args && typeof args === 'object' ? args : {};
        var sheetName = normalizeTextPayload(safeArgs.sheetName || '');
        var cell = normalizeTextPayload(safeArgs.cell || safeArgs.address || '').toUpperCase();
        var mode = normalizeTextPayload(safeArgs.mode || 'replace').toLowerCase();
        var textValue = safeArgs.text;
        if (textValue === undefined || textValue === null) {
            textValue = safeArgs.value;
        }
        var text = normalizeTextPayload(textValue);
        var createSheet = safeArgs.createSheet === true;

        if (!/^[A-Z]{1,3}\d{1,7}$/.test(cell)) {
            return { ok: false, data: null, logs: [], error: 'Invalid cell address. Expected format like A1.' };
        }
        if (mode !== 'replace' && mode !== 'append') mode = 'replace';

        var host = getHostBridge();
        if (!host || typeof host.callEditorCommand !== 'function') {
            return { ok: false, data: null, logs: [], error: 'Host bridge is unavailable' };
        }

        var result = await host.callEditorCommand(function () {
            'use strict';
            try {
                var sName = Asc.scope.sheetName;
                var sCell = String(Asc.scope.cell || 'A1');
                var sText = String(Asc.scope.text || '');
                var sMode = String(Asc.scope.mode || 'replace').toLowerCase();
                var shouldCreateSheet = Asc.scope.createSheet === true;

                var sheet = null;
                if (sName && typeof Api.GetSheet === 'function') {
                    sheet = Api.GetSheet(sName);
                }
                if (!sheet && shouldCreateSheet && sName && typeof Api.AddSheet === 'function') {
                    sheet = Api.AddSheet(sName);
                }
                if (!sheet && typeof Api.GetActiveSheet === 'function') {
                    sheet = Api.GetActiveSheet();
                }
                if (!sheet) {
                    return JSON.stringify({ ok: false, error: 'Sheet is not available' });
                }

                if (typeof sheet.SetActive === 'function') sheet.SetActive();

                var range = sheet.GetRange(sCell);
                if (!range) {
                    return JSON.stringify({ ok: false, error: 'Cell range is not available: ' + sCell });
                }

                var valueToWrite = sText;
                if (sMode === 'append') {
                    var existing = '';
                    try {
                        if (typeof range.GetText === 'function') {
                            existing = String(range.GetText() || '');
                        } else if (typeof range.GetValue2 === 'function') {
                            existing = String(range.GetValue2() || '');
                        } else if (typeof range.GetValue === 'function') {
                            existing = String(range.GetValue() || '');
                        }
                    } catch (e1) {
                        existing = '';
                    }
                    valueToWrite = existing + sText;
                }

                range.SetValue(valueToWrite);
                return JSON.stringify({
                    ok: true,
                    data: {
                        sheetName: typeof sheet.GetName === 'function' ? String(sheet.GetName() || sName || '') : String(sName || ''),
                        cell: sCell,
                        mode: sMode,
                        value: valueToWrite,
                        written: true
                    }
                });
            } catch (error) {
                return JSON.stringify({ ok: false, error: String(error && error.message ? error.message : error) });
            }
        }, {
            sheetName: sheetName,
            cell: cell,
            text: text,
            mode: mode,
            createSheet: createSheet
        }, {
            timeoutMs: Number(constants().agentLimits.stepTimeoutMs || 300000)
        });

        var parsedWriteCellResult = parseAgentJsonPayload(result, null);
        if (parsedWriteCellResult && parsedWriteCellResult.ok) {
            return {
                ok: true,
                data: parsedWriteCellResult.data,
                logs: ['cell_text_written'],
                error: null
            };
        }

        return {
            ok: false,
            data: null,
            logs: [],
            error: parsedWriteCellResult && parsedWriteCellResult.error ? String(parsedWriteCellResult.error) : 'Failed to write cell text'
        };
    }

    async function executeWriteCellsBatchStep(args) {
        var safeArgs = args && typeof args === 'object' ? args : {};
        var sheetName = normalizeTextPayload(safeArgs.sheetName || '');
        var items = Array.isArray(safeArgs.items) ? safeArgs.items : [];
        var createSheet = safeArgs.createSheet === true;
        var clearFirst = safeArgs.clearFirst === true;

        if (!items.length) {
            return { ok: false, data: null, logs: [], error: 'items array is empty' };
        }

        var normalizedItems = items.slice(0, 20000).map(function (item) {
            var source = item && typeof item === 'object' ? item : {};
            var cell = normalizeTextPayload(source.cell || source.address || '').toUpperCase();
            var value = source.value;
            if (value === undefined) value = source.text;
            if (value === undefined || value === null) value = '';
            return {
                cell: cell,
                value: value
            };
        }).filter(function (item) {
            return /^[A-Z]{1,3}\d{1,7}$/.test(item.cell);
        });

        if (!normalizedItems.length) {
            return { ok: false, data: null, logs: [], error: 'No valid cell addresses in items.' };
        }

        var host = getHostBridge();
        if (!host || typeof host.callEditorCommand !== 'function') {
            return { ok: false, data: null, logs: [], error: 'Host bridge is unavailable' };
        }

        var result = await host.callEditorCommand(function () {
            'use strict';
            try {
                var sName = Asc.scope.sheetName;
                var sItems = Array.isArray(Asc.scope.items) ? Asc.scope.items : [];
                var shouldCreateSheet = Asc.scope.createSheet === true;
                var shouldClearFirst = Asc.scope.clearFirst === true;

                var sheet = null;
                if (sName && typeof Api.GetSheet === 'function') {
                    sheet = Api.GetSheet(sName);
                }
                if (!sheet && shouldCreateSheet && sName && typeof Api.AddSheet === 'function') {
                    sheet = Api.AddSheet(sName);
                }
                if (!sheet && typeof Api.GetActiveSheet === 'function') {
                    sheet = Api.GetActiveSheet();
                }
                if (!sheet) {
                    return JSON.stringify({ ok: false, error: 'Sheet is not available' });
                }

                if (typeof sheet.SetActive === 'function') sheet.SetActive();

                if (shouldClearFirst && typeof sheet.GetUsedRange === 'function') {
                    var used = sheet.GetUsedRange();
                    if (used) {
                        if (typeof used.ClearContents === 'function') {
                            used.ClearContents();
                        } else if (typeof used.Clear === 'function') {
                            used.Clear();
                        }
                    }
                }

                var written = 0;
                var skipped = 0;
                for (var i = 0; i < sItems.length; i += 1) {
                    var item = sItems[i] || {};
                    var cell = String(item.cell || '').toUpperCase();
                    if (!/^[A-Z]{1,3}\d{1,7}$/.test(cell)) {
                        skipped += 1;
                        continue;
                    }
                    var range = sheet.GetRange(cell);
                    if (!range) {
                        skipped += 1;
                        continue;
                    }
                    range.SetValue(item.value === null || item.value === undefined ? '' : item.value);
                    written += 1;
                }

                return JSON.stringify({
                    ok: true,
                    data: {
                        sheetName: typeof sheet.GetName === 'function' ? String(sheet.GetName() || sName || '') : String(sName || ''),
                        itemCount: sItems.length,
                        written: written,
                        skipped: skipped
                    }
                });
            } catch (error) {
                return JSON.stringify({ ok: false, error: String(error && error.message ? error.message : error) });
            }
        }, {
            sheetName: sheetName,
            items: normalizedItems,
            createSheet: createSheet,
            clearFirst: clearFirst
        }, {
            timeoutMs: Number(constants().agentLimits.stepTimeoutMs || 300000)
        });

        var parsedBatchResult = parseAgentJsonPayload(result, null);
        if (parsedBatchResult && parsedBatchResult.ok) {
            return {
                ok: true,
                data: parsedBatchResult.data,
                logs: ['batch_cells_written'],
                error: null
            };
        }

        return {
            ok: false,
            data: null,
            logs: [],
            error: parsedBatchResult && parsedBatchResult.error ? String(parsedBatchResult.error) : 'Failed to write cells batch'
        };
    }

    async function executeCopyAttachedSheetStep(args) {
        var safeArgs = args && typeof args === 'object' ? args : {};
        var filename = normalizeTextPayload(safeArgs.filename || '');
        var sourceSheetName = normalizeTextPayload(safeArgs.sourceSheetName || '');
        var targetSheetName = normalizeTextPayload(safeArgs.targetSheetName || sourceSheetName || 'Imported Data');
        var clearFirst = safeArgs.clearFirst === true;

        var threadStore = getChatThreadStore();
        if (!threadStore) return { ok: false, data: null, logs: [], error: 'Thread store unavailable' };

        var activeThread = threadStore.getActiveThread();
        if (!activeThread) return { ok: false, data: null, logs: [], error: 'Active thread unavailable' };

        var targetAttachment = null;
        for (var i = activeThread.messages.length - 1; i >= 0; i--) {
            var msg = activeThread.messages[i];
            if (Array.isArray(msg.attachments)) {
                for (var j = 0; j < msg.attachments.length; j++) {
                    var att = msg.attachments[j];
                    if (att.name === filename && att.attachmentType === 'file' && att.contextSummary && att.contextSummary.kind === 'xlsx_workbook') {
                        targetAttachment = att;
                        break;
                    }
                }
            }
            if (targetAttachment) break;
        }

        if (!targetAttachment) {
            return { ok: false, data: null, logs: [], error: 'Attached workbook not found in recent messages: ' + filename };
        }

        var attachmentService = root.features.chatAttachments;
        if (!attachmentService || typeof attachmentService.getSessionBlob !== 'function') {
            return { ok: false, data: null, logs: [], error: 'Attachment session cache unavailable' };
        }

        var blob = attachmentService.getSessionBlob(targetAttachment.id);
        if (!blob) {
            return { ok: false, data: null, logs: [], error: 'Attachment content is not available in memory. Please re-attach the file.' };
        }

        if (!root.shared || !root.shared.xlsxParser) {
            return { ok: false, data: null, logs: [], error: 'XLSX parser is unavailable' };
        }

        var sheetData;
        try {
            var summary = await root.shared.xlsxParser.parseXlsxFromBlob(blob, { full: true, maxRows: 100000 });
            var sheet = summary.sheets.find(function(s) { return s.name === sourceSheetName; });
            if (!sheet) {
                return { ok: false, data: null, logs: [], error: 'Sheet not found in attachment: ' + sourceSheetName };
            }
            sheetData = sheet.values || [];
        } catch (error) {
            return { ok: false, data: null, logs: [], error: 'Failed to read attached sheet: ' + String(error) };
        }

        if (!sheetData.length) {
            return { ok: false, data: null, logs: [], error: 'Attached sheet is empty' };
        }

        return executeWriteDataToSheetStep({
            sheetName: targetSheetName,
            data: sheetData,
            clearFirst: clearFirst
        });
    }

    async function executeReadSheetRangeStep(args) {
        var context = getContextService();
        var safeArgs = args && typeof args === 'object' ? args : {};
        var rangeData = await context.collectSheetRangeContext(safeArgs.sheetName || '', safeArgs.range || '');
        if (!rangeData) {
            return { ok: false, data: null, logs: [], error: 'Requested range is not available' };
        }
        var payload = context.formatTablePayload(rangeData.values, {
            maxRows: constants().contextLimits.maxRowsPerSheet,
            maxCols: constants().contextLimits.maxColsPerSheet,
            maxChars: constants().contextLimits.maxSheetChars
        });
        var nonEmptyPreview = buildSpreadsheetNonEmptyPreview(rangeData.values, rangeData.address || safeArgs.range || '', 24);
        return {
            ok: true,
            data: {
                sheetName: rangeData.sheetName || safeArgs.sheetName || 'Active sheet',
                range: rangeData.address || safeArgs.range || '',
                rows: payload.totalRows,
                cols: payload.totalCols,
                truncated: payload.truncated,
                payload: payload.text,
                nonEmptyCount: nonEmptyPreview.total,
                nonEmptyCells: nonEmptyPreview.cells,
                nonEmptyCellsTruncated: nonEmptyPreview.truncated
            },
            logs: [],
            error: null
        };
    }

    async function executeReadDocumentSnapshotStep(args) {
        var host = getHostBridge();
        if (!host || typeof host.callEditorCommand !== 'function') {
            return { ok: false, data: null, logs: [], error: 'Host bridge is unavailable' };
        }
        if (getCurrentEditorType() !== 'word') {
            return { ok: false, data: null, logs: [], error: 'read_document_snapshot is available only in Word editor' };
        }

        var safeArgs = args && typeof args === 'object' ? args : {};
        var contextLimits = constants().contextLimits || {};
        var maxChars = Number(safeArgs.maxChars || contextLimits.maxDocumentChars || 12000);
        var maxParagraphsPerChunk = Number(safeArgs.maxParagraphsPerChunk || 40);
        var maxChunks = Number(safeArgs.maxChunks || 8);
        maxChars = Math.max(800, Math.min(30000, Number.isFinite(maxChars) ? maxChars : 12000));
        maxParagraphsPerChunk = Math.max(1, Math.min(200, Number.isFinite(maxParagraphsPerChunk) ? Math.floor(maxParagraphsPerChunk) : 40));
        maxChunks = Math.max(1, Math.min(24, Number.isFinite(maxChunks) ? Math.floor(maxChunks) : 8));

        var rawResult = await host.callEditorCommand(function () {
            function readParagraphText(paragraph) {
                if (!paragraph || typeof paragraph.GetText !== 'function') return '';
                try {
                    var text = paragraph.GetText();
                    return text === null || text === undefined ? '' : String(text);
                } catch (error) {
                    return '';
                }
            }

            try {
                var doc = Api.GetDocument ? Api.GetDocument() : null;
                if (!doc) {
                    return JSON.stringify({
                        editorType: 'word',
                        mode: 'snapshot',
                        source: 'document_snapshot_error',
                        coverage: {
                            totalParagraphs: 0,
                            nonEmptyParagraphs: 0,
                            collectedParagraphs: 0
                        },
                        objects: {
                            imageCount: 0,
                            tableCount: 0
                        },
                        textChunks: [],
                        payload: '',
                        truncated: false,
                        warnings: ['document_not_available'],
                        error: 'Document is not available'
                    });
                }

                var paragraphs = typeof doc.GetAllParagraphs === 'function' ? (doc.GetAllParagraphs() || []) : [];
                var totalParagraphs = paragraphs.length;
                var nonEmptyParagraphs = 0;
                for (var p = 0; p < paragraphs.length; p += 1) {
                    if (readParagraphText(paragraphs[p]).trim().length) nonEmptyParagraphs += 1;
                }

                var payload = '';
                var textChunks = [];
                var truncated = false;
                var collectedParagraphs = 0;
                var cursor = 0;
                var chunkIndex = 0;

                while (cursor < totalParagraphs && chunkIndex < Number(Asc.scope.maxChunks || 8)) {
                    var start = cursor;
                    var endExclusive = Math.min(totalParagraphs, start + Number(Asc.scope.maxParagraphsPerChunk || 40));
                    var chunkLines = [];
                    for (var i = start; i < endExclusive; i += 1) {
                        chunkLines.push(readParagraphText(paragraphs[i]));
                    }
                    var chunkText = chunkLines.join('\n');
                    var block = '[Paragraphs ' + (start + 1) + '-' + endExclusive + ']\n' + chunkText;

                    if (payload.length + (payload.length ? 2 : 0) + block.length > Number(Asc.scope.maxChars || 12000)) {
                        var remaining = Math.max(0, Number(Asc.scope.maxChars || 12000) - payload.length - (payload.length ? 2 : 0));
                        if (remaining > 0) {
                            var limitedBlock = block.slice(0, remaining);
                            if (payload.length) payload += '\n\n';
                            payload += limitedBlock;
                            textChunks.push({
                                start: start + 1,
                                end: endExclusive,
                                text: limitedBlock
                            });
                        }
                        collectedParagraphs = endExclusive;
                        truncated = true;
                        break;
                    }

                    if (payload.length) payload += '\n\n';
                    payload += block;
                    textChunks.push({
                        start: start + 1,
                        end: endExclusive,
                        text: chunkText
                    });

                    collectedParagraphs = endExclusive;
                    cursor = endExclusive;
                    chunkIndex += 1;
                }

                if (collectedParagraphs < totalParagraphs) {
                    truncated = true;
                }

                var warnings = [];
                if (!totalParagraphs) warnings.push('document_empty');
                if (totalParagraphs > 0 && nonEmptyParagraphs === 0) warnings.push('all_paragraphs_empty_or_nontext');
                if (truncated && chunkIndex >= Number(Asc.scope.maxChunks || 8)) warnings.push('max_chunks_reached');
                if (truncated && payload.length >= Number(Asc.scope.maxChars || 12000)) warnings.push('max_chars_reached');

                var imageCount = 0;
                var tableCount = 0;
                try {
                    if (typeof doc.GetAllImages === 'function') {
                        var images = doc.GetAllImages() || [];
                        imageCount = Array.isArray(images) ? images.length : 0;
                    } else if (typeof doc.GetAllDrawings === 'function') {
                        var drawings = doc.GetAllDrawings() || [];
                        imageCount = Array.isArray(drawings) ? drawings.length : 0;
                    }
                } catch (imageError) {
                    warnings.push('image_count_unavailable');
                }
                try {
                    if (typeof doc.GetAllTables === 'function') {
                        var tables = doc.GetAllTables() || [];
                        tableCount = Array.isArray(tables) ? tables.length : 0;
                    }
                } catch (tableError) {
                    warnings.push('table_count_unavailable');
                }

                return JSON.stringify({
                    editorType: 'word',
                    mode: 'snapshot',
                    source: 'document_snapshot',
                    coverage: {
                        totalParagraphs: totalParagraphs,
                        nonEmptyParagraphs: nonEmptyParagraphs,
                        collectedParagraphs: collectedParagraphs
                    },
                    objects: {
                        imageCount: imageCount,
                        tableCount: tableCount
                    },
                    textChunks: textChunks,
                    payload: payload,
                    truncated: truncated,
                    warnings: warnings
                });
            } catch (error) {
                return JSON.stringify({
                    editorType: 'word',
                    mode: 'snapshot',
                    source: 'document_snapshot_error',
                    coverage: {
                        totalParagraphs: 0,
                        nonEmptyParagraphs: 0,
                        collectedParagraphs: 0
                    },
                    objects: {
                        imageCount: 0,
                        tableCount: 0
                    },
                    textChunks: [],
                    payload: '',
                    truncated: false,
                    warnings: ['snapshot_failed'],
                    error: String(error && error.message ? error.message : error)
                });
            }
        }, {
            maxChars: maxChars,
            maxParagraphsPerChunk: maxParagraphsPerChunk,
            maxChunks: maxChunks
        }, {
            timeoutMs: Number(constants().agentLimits.stepTimeoutMs || 300000)
        });

        var parsed = parseAgentJsonPayload(rawResult, null);
        if (!parsed || typeof parsed !== 'object') {
            return { ok: false, data: null, logs: [], error: 'Document snapshot returned an invalid payload' };
        }
        if (parsed.source === 'document_snapshot_error') {
            return {
                ok: false,
                data: parsed,
                logs: ['document_snapshot_failed'],
                error: normalizeTextPayload(parsed.error || 'Document snapshot failed')
            };
        }

        var coverage = parsed.coverage && typeof parsed.coverage === 'object' ? parsed.coverage : {};
        var objects = parsed.objects && typeof parsed.objects === 'object' ? parsed.objects : {};
        var normalized = {
            editorType: 'word',
            mode: 'snapshot',
            source: normalizeTextPayload(parsed.source || 'document_snapshot') || 'document_snapshot',
            coverage: {
                totalParagraphs: Number(coverage.totalParagraphs || 0) || 0,
                nonEmptyParagraphs: Number(coverage.nonEmptyParagraphs || 0) || 0,
                collectedParagraphs: Number(coverage.collectedParagraphs || 0) || 0
            },
            objects: {
                imageCount: Number(objects.imageCount || 0) || 0,
                tableCount: Number(objects.tableCount || 0) || 0
            },
            textChunks: Array.isArray(parsed.textChunks) ? parsed.textChunks.map(function (chunk) {
                var safeChunk = chunk && typeof chunk === 'object' ? chunk : {};
                return {
                    start: Number(safeChunk.start || 0) || 0,
                    end: Number(safeChunk.end || 0) || 0,
                    text: normalizeTextPayload(safeChunk.text || '')
                };
            }).slice(0, 16) : [],
            payload: normalizeTextPayload(parsed.payload || ''),
            truncated: parsed.truncated === true,
            warnings: Array.isArray(parsed.warnings) ? parsed.warnings.slice(0, 12) : []
        };

        agentState().lastReadDocumentText = normalizeTextPayload(normalized.payload || '');
        return {
            ok: true,
            data: normalized,
            logs: ['document_snapshot_collected'],
            error: null
        };
    }

    async function executeGetCurrentTimeStep() {
        var now = new Date();
        var local = '';
        try {
            local = now.toLocaleString();
        } catch (error) {
            local = now.toString();
        }
        var timezone = '';
        try {
            timezone = Intl && Intl.DateTimeFormat
                ? (Intl.DateTimeFormat().resolvedOptions().timeZone || '')
                : '';
        } catch (error2) {
            timezone = '';
        }
        return {
            ok: true,
            data: {
                iso: now.toISOString(),
                local: normalizeTextPayload(local),
                timezone: normalizeTextPayload(timezone),
                unix_ms: now.getTime()
            },
            logs: [],
            error: null
        };
    }

    async function executeReadDocumentStep() {
        var context = getContextService();
        var docData = await context.collectWordDocumentContext();
        if (!docData) {
            return { ok: false, data: null, logs: [], error: 'Document context is not available' };
        }
        agentState().lastReadDocumentText = normalizeTextPayload(docData.text || '');
        return {
            ok: true,
            data: {
                source: docData.source || 'full_document',
                totalParagraphs: docData.totalParagraphs || 0,
                text: docData.text || ''
            },
            logs: ['document_read_success'],
            error: null
        };
    }

    function getLastExecutableStepIndex(stepType) {
        var steps = Array.isArray(agentState().steps) ? agentState().steps : [];
        for (var i = steps.length - 1; i >= 0; i -= 1) {
            if (steps[i] && steps[i].type === stepType) return i;
        }
        return -1;
    }

    function hasPendingPostMacroVerification() {
        var lastMacroIndex = getLastExecutableStepIndex('run_macro_code');
        if (lastMacroIndex < 0) return false;
        
        var editorType = getCurrentEditorType();
        if (editorType === 'word') {
            var lastContextIndex = getLastExecutableStepIndex('collect_context');
            return lastContextIndex < lastMacroIndex;
        }
        
        if (editorType === 'cell') {
            var read1 = getLastExecutableStepIndex('read_active_sheet');
            var read2 = getLastExecutableStepIndex('read_sheet_range');
            var read3 = getLastExecutableStepIndex('list_sheets');
            var lastReadIndex = Math.max(read1, read2, read3);
            return lastReadIndex < lastMacroIndex;
        }
        
        return false;
    }

    function buildForcedPostMacroVerificationStep(stepIndex) {
        var editorType = getCurrentEditorType();
        if (editorType === 'word') {
            return {
                id: 'step_' + (stepIndex + 1),
                type: 'collect_context',
                reason: 'Verify the document state after macro execution before producing the final answer.',
                args: {
                    mode: 'collect',
                    editorType: 'word',
                    intent: 'verify_after_write',
                    forceRefresh: true
                }
            };
        }
        if (editorType === 'cell') {
            return buildForcedSpreadsheetReadStep(agentState().lastUserMessage || '', stepIndex, 'verify_after_write');
        }
        return null;
    }

    async function executeListHostToolsStep(args) {
        var bridge = getDesktopToolsBridge();
        if (!bridge || typeof bridge.getStatus !== 'function') {
            return { ok: false, data: null, logs: [], error: 'Desktop tools bridge is unavailable' };
        }

        var safeArgs = args && typeof args === 'object' ? args : {};
        if (safeArgs.forceRefresh === true && typeof bridge.refreshCatalog === 'function') {
            try { bridge.refreshCatalog(); } catch (error) {}
        }
        var status = bridge.getStatus();
        if (status.catalogAvailable !== true) {
            return {
                ok: false,
                data: {
                    status: status,
                    catalog: []
                },
                logs: ['host_tools_catalog_unavailable'],
                error: status.catalogParseError || 'Desktop tools catalog is unavailable'
            };
        }

        var query = normalizeTextPayload(safeArgs.query || agentState().lastUserMessage || '');
        var relevant = findRelevantDesktopTools(query, 12).map(function (entry) {
            return normalizeHostToolDiscoveryEntry(entry, query);
        });
        agentState().lastHostToolDiscovery = {
            query: query,
            relevant: cloneSerializable(relevant),
            status: cloneSerializable(status),
            at: new Date().toISOString()
        };
        return {
            ok: true,
            data: {
                status: status,
                query: query,
                catalog: getEnabledDesktopToolCatalog(),
                relevant: cloneSerializable(relevant)
            },
            logs: ['host_tools_catalog_loaded'],
            error: null
        };
    }

    async function executeCallHostToolStep(args) {
        var bridge = getDesktopToolsBridge();
        if (!bridge || typeof bridge.callTool !== 'function') {
            return { ok: false, data: null, logs: [], error: 'Desktop tools bridge is unavailable' };
        }

        var safeArgs = args && typeof args === 'object' ? args : {};
        var toolName = extractHostToolName(safeArgs);
        var input = safeArgs.input && typeof safeArgs.input === 'object'
            ? safeArgs.input
            : (safeArgs.args && typeof safeArgs.args === 'object' ? safeArgs.args : {});
        if (!toolName) {
            return { ok: false, data: null, logs: [], error: 'call_host_tool requires args.tool' };
        }

        var result = await bridge.callTool(toolName, input);
        return {
            ok: result && result.ok === true,
            data: {
                tool: toolName,
                input: cloneSerializable(input),
                output: result ? cloneSerializable(result.data) : null,
                descriptor: result && result.descriptor ? cloneSerializable(result.descriptor) : null
            },
            logs: [result && result.ok === true ? 'host_tool_called' : 'host_tool_failed'],
            error: result && result.error ? String(result.error) : null
        };
    }

    function normalizeReferenceMethodLimit(rawLimit) {
        var parsed = Number(rawLimit);
        if (!Number.isFinite(parsed) || parsed <= 0) {
            return constants().apiReferenceDefaultMethodLimit;
        }
        return Math.min(Math.floor(parsed), constants().apiReferenceMaxMethodLimit);
    }

    function makeLimitedMethodPack(items, limit) {
        var list = Array.isArray(items) ? items : [];
        var safeLimit = normalizeReferenceMethodLimit(limit);
        var limited = list.slice(0, safeLimit);
        return {
            methods: limited,
            total_count: list.length,
            returned_count: limited.length,
            truncated: list.length > limited.length
        };
    }

    function splitSearchTokens(value) {
        var source = normalizeTextPayload(value || '')
            .replace(/([a-z])([A-Z])/g, '$1 $2')
            .replace(/[^\p{L}\p{N}]+/gu, ' ')
            .toLowerCase()
            .trim();
        if (!source) return [];
        var seen = {};
        return source.split(/\s+/).filter(function (token) {
            if (!token || token.length < 2) return false;
            if (seen[token]) return false;
            seen[token] = true;
            return true;
        });
    }

    function getEnabledDesktopToolCatalog() {
        var bridge = getDesktopToolsBridge();
        var settings = loadAgentRuntimeSettings();
        var desktopSettings = normalizeDesktopToolsSettings(settings && settings.desktopTools);
        var disabled = {};
        desktopSettings.disabledTools.forEach(function (name) {
            var normalized = normalizeTextPayload(name || '');
            if (normalized) disabled[normalized] = true;
        });
        if (!bridge || typeof bridge.getCatalog !== 'function') return [];
        var catalog = bridge.getCatalog();
        if (!Array.isArray(catalog)) return [];
        return catalog.filter(function (tool) {
            var name = normalizeTextPayload(tool && tool.name || '');
            return !!name && !disabled[name];
        });
    }

    function buildDesktopToolSearchCorpus(tool) {
        var source = tool && typeof tool === 'object' ? tool : {};
        var schema = source.inputSchema && typeof source.inputSchema === 'object' ? source.inputSchema : {};
        var properties = schema.properties && typeof schema.properties === 'object' ? Object.keys(schema.properties) : [];
        return [
            normalizeTextPayload(source.name || '').replace(/[_-]+/g, ' '),
            normalizeTextPayload(source.description || ''),
            properties.join(' ')
        ].join(' ').toLowerCase();
    }

    function expandDesktopToolSearchTokens(tokens, rawNeedle) {
        var expanded = Array.isArray(tokens) ? tokens.slice() : [];
        var text = normalizeTextPayload(rawNeedle || '').toLowerCase();

        function add(token) {
            if (!token) return;
            if (expanded.indexOf(token) !== -1) return;
            expanded.push(token);
        }

        if (/отчет|отчёт|сводк|статист/i.test(text)) {
            add('report');
            add('summary');
            add('stats');
        }
        if (/лист|вкладк/i.test(text)) {
            add('sheet');
            add('worksheet');
        }
        if (/ячейк/i.test(text)) {
            add('cell');
        }
        if (/таблиц/i.test(text)) {
            add('table');
        }
        if (/встав|добав|запиш|заполн|перенес/i.test(text)) {
            add('insert');
            add('write');
            add('append');
            add('fill');
        }
        if (/скопир/i.test(text)) {
            add('copy');
        }
        if (/файл|вложен|прикрепл|xlsx|excel/i.test(text)) {
            add('file');
            add('attachment');
            add('xlsx');
        }

        return expanded;
    }

    function scoreDesktopToolDescriptor(tool, rawNeedle) {
        var needle = normalizeTextPayload(rawNeedle || '').trim();
        if (!needle) return 0;
        var tokens = expandDesktopToolSearchTokens(splitSearchTokens(needle), needle);
        if (!tokens.length) return 0;

        var source = tool && typeof tool === 'object' ? tool : {};
        var name = normalizeTextPayload(source.name || '').toLowerCase();
        var description = normalizeTextPayload(source.description || '').toLowerCase();
        var corpus = buildDesktopToolSearchCorpus(source);
        var score = 0;

        tokens.forEach(function (token) {
            if (name.indexOf(token) !== -1) score += 140;
            if (description.indexOf(token) !== -1) score += 60;
            if (corpus.indexOf(token) !== -1) score += 20;
        });

        if (/summary|report|stats|statistic|analy|summari|свод|отчет|отчёт|статист|сумм/i.test(needle)) {
            if (/summary|report|stats|analy|table|sheet|worksheet|workbook/i.test(corpus)) score += 80;
        }
        if (/insert|append|write|fill|copy|move|создай|добав|встав|запиш|скопир|перенес/i.test(needle)) {
            if (/insert|append|write|fill|copy|move|create|sheet|cell|range|worksheet/i.test(corpus)) score += 60;
        }

        return score;
    }

    function findRelevantDesktopTools(messageText, limit) {
        var safeLimit = Math.max(1, Number(limit || 12) || 12);
        var catalog = getEnabledDesktopToolCatalog();
        var needle = normalizeTextPayload(messageText || '');
        if (!catalog.length) return [];

        var scored = catalog.map(function (tool) {
            return {
                tool: cloneSerializable(tool),
                score: scoreDesktopToolDescriptor(tool, needle)
            };
        }).filter(function (entry) {
            return entry.score > 0;
        });

        scored.sort(function (left, right) {
            if (right.score !== left.score) return right.score - left.score;
            return String(left.tool && left.tool.name || '').localeCompare(String(right.tool && right.tool.name || ''));
        });

        return scored.slice(0, safeLimit);
    }

    function buildSuggestedHostToolInput(tool, messageText) {
        var source = tool && typeof tool === 'object' ? tool : {};
        var schema = source.inputSchema && typeof source.inputSchema === 'object' ? source.inputSchema : {};
        var properties = schema.properties && typeof schema.properties === 'object' ? schema.properties : {};
        var suggested = {};
        var text = normalizeTextPayload(messageText || agentState().lastUserMessage || '');
        var explicitSheetName = extractSpreadsheetSheetName(text);
        var attached = getLatestAttachedWorkbookMeta();

        Object.keys(properties).forEach(function (key) {
            var normalizedKey = normalizeTextPayload(key || '').toLowerCase();
            if (!normalizedKey) return;

            if (normalizedKey === 'sheetname') {
                suggested[key] = explicitSheetName || extractTargetSheetNameForCopy(text) || 'Sheet2';
                return;
            }
            if (normalizedKey === 'targetsheetname') {
                suggested[key] = extractTargetSheetNameForCopy(text) || explicitSheetName || 'Sheet2';
                return;
            }
            if (normalizedKey === 'sourcesheetname') {
                suggested[key] = attached && attached.defaultSourceSheet ? attached.defaultSourceSheet : (explicitSheetName || 'Sheet1');
                return;
            }
            if (normalizedKey === 'filename' || normalizedKey === 'filepath') {
                if (attached && attached.filename) suggested[key] = attached.filename;
                return;
            }
            if (normalizedKey === 'anchor') {
                if (/\b(bottom|below|footer|end)\b|внизу|внизу листа|в конце/i.test(text)) {
                    suggested[key] = 'bottom';
                }
                return;
            }
            if (normalizedKey === 'query' || normalizedKey === 'request' || normalizedKey === 'instruction' || normalizedKey === 'prompt') {
                suggested[key] = text;
                return;
            }
            if (normalizedKey === 'text' || normalizedKey === 'value') {
                suggested[key] = text;
                return;
            }
        });

        return suggested;
    }

    function normalizeHostToolDiscoveryEntry(entry, messageText) {
        var tool = entry && entry.tool ? entry.tool : {};
        var schema = tool.inputSchema && typeof tool.inputSchema === 'object' ? tool.inputSchema : {};
        var properties = schema.properties && typeof schema.properties === 'object' ? Object.keys(schema.properties) : [];
        return {
            name: normalizeTextPayload(tool.name || ''),
            description: normalizeTextPayload(tool.description || ''),
            score: Number(entry && entry.score || 0) || 0,
            required: Array.isArray(schema.required) ? schema.required.slice(0, 12) : [],
            properties: properties.slice(0, 12),
            inputSchema: cloneSerializable(schema),
            suggestedInput: buildSuggestedHostToolInput(tool, messageText)
        };
    }

    function canDirectlyCallHostTool(discoveryEntry) {
        var entry = discoveryEntry && typeof discoveryEntry === 'object' ? discoveryEntry : {};
        var required = Array.isArray(entry.required) ? entry.required : [];
        var suggested = entry.suggestedInput && typeof entry.suggestedInput === 'object' ? entry.suggestedInput : {};
        for (var i = 0; i < required.length; i += 1) {
            if (suggested[required[i]] === undefined || suggested[required[i]] === '') {
                return false;
            }
        }
        return true;
    }

    function buildForcedInitialHostToolStep(messageText, stepIndex) {
        var text = normalizeTextPayload(messageText || '');
        if (!shouldForceHostToolDiscoveryFirst(text)) return null;
        var relevant = findRelevantDesktopTools(text, 3).map(function (entry) {
            return normalizeHostToolDiscoveryEntry(entry, text);
        });
        if (!relevant.length) return null;

        var best = relevant[0];
        if (best && Number(best.score || 0) >= 180 && canDirectlyCallHostTool(best)) {
            return {
                id: 'step_' + (stepIndex + 1),
                type: 'call_host_tool',
                reason: 'Use the best matching runtime host tool directly for this spreadsheet write task before considering macros.',
                args: {
                    tool: best.name,
                    input: cloneSerializable(best.suggestedInput || {})
                }
            };
        }

        return {
            id: 'step_' + (stepIndex + 1),
            type: 'list_host_tools',
            reason: 'Inspect the runtime host tool catalog first because this spreadsheet write task may be handled by a native desktop tool.',
            args: {
                query: text,
                intent: 'host_tool_discovery',
                forceRefresh: true
            }
        };
    }

    function buildRelevantDesktopToolsPromptCatalog(messageText, limit) {
        var bridge = getDesktopToolsBridge();
        var relevant = findRelevantDesktopTools(messageText, limit || 16);
        if (relevant.length) {
            return relevant.map(function (entry) {
                var tool = entry.tool || {};
                var schema = tool.inputSchema && typeof tool.inputSchema === 'object' ? tool.inputSchema : {};
                var required = Array.isArray(schema.required) && schema.required.length ? schema.required.join(', ') : 'none';
                var properties = schema.properties && typeof schema.properties === 'object' ? Object.keys(schema.properties) : [];
                var props = properties.slice(0, 6).join(', ');
                var fragments = [tool.name, tool.description || 'No description', 'required=' + required];
                if (props) fragments.push('props=' + props);
                return '- ' + fragments.join(' | ');
            }).join('\n');
        }
        if (bridge && typeof bridge.getCatalogForPrompt === 'function') {
            return normalizeTextPayload(bridge.getCatalogForPrompt(limit || 12));
        }
        return '';
    }

    function buildReferenceEntry(item, categoryKey, categoryTitle) {
        var source = item && typeof item === 'object' ? item : {};
        return {
            object: source.object || '',
            method: source.method || '',
            args: source.args || '',
            description: source.description || '',
            category: categoryKey || '',
            category_title: categoryTitle || ''
        };
    }

    function scoreReferenceEntry(entry, rawNeedle, tokens) {
        var methodName = normalizeTextPayload(entry && entry.method || '').toLowerCase();
        var description = normalizeTextPayload(entry && entry.description || '').toLowerCase();
        var objectName = normalizeTextPayload(entry && entry.object || '').toLowerCase();
        var score = 0;
        var exactNeedle = normalizeTextPayload(rawNeedle || '').toLowerCase();

        if (exactNeedle) {
            if (methodName === exactNeedle) score += 1000;
            else if (methodName.indexOf(exactNeedle) !== -1) score += 700;
            if (description.indexOf(exactNeedle) !== -1) score += 320;
            if (objectName.indexOf(exactNeedle) !== -1) score += 180;
        }

        (tokens || []).forEach(function (token) {
            if (methodName.indexOf(token) !== -1) score += 120;
            if (description.indexOf(token) !== -1) score += 45;
            if (objectName.indexOf(token) !== -1) score += 35;
        });

        return score;
    }

    function findApiReferenceMatches(catalog, args) {
        var safeArgs = args && typeof args === 'object' ? args : {};
        var categories = catalog && catalog.categories && typeof catalog.categories === 'object' ? catalog.categories : {};
        var objects = catalog && catalog.objects && typeof catalog.objects === 'object' ? catalog.objects : {};
        var requestedCategory = normalizeTextPayload(safeArgs.category || '').trim();
        var rawNeedle = normalizeTextPayload(safeArgs.method || safeArgs.query || safeArgs.error || '').trim();
        var tokens = splitSearchTokens(rawNeedle);

        if (requestedCategory && categories[requestedCategory]) {
            var category = categories[requestedCategory];
            var methods = Array.isArray(category.methods) ? category.methods : [];
            return methods.map(function (item) {
                return buildReferenceEntry(item, requestedCategory, category.title || requestedCategory);
            });
        }

        if (!rawNeedle && !tokens.length) {
            return [];
        }

        var hits = [];
        Object.keys(objects).forEach(function (objectName) {
            var list = Array.isArray(objects[objectName]) ? objects[objectName] : [];
            list.forEach(function (item) {
                var entry = buildReferenceEntry(item, item && item.category || '', item && item.category_title || '');
                if (!entry.object) entry.object = objectName;
                var score = scoreReferenceEntry(entry, rawNeedle, tokens);
                if (score <= 0) return;
                entry.match_score = score;
                hits.push(entry);
            });
        });

        hits.sort(function (left, right) {
            if (right.match_score !== left.match_score) return right.match_score - left.match_score;
            return String(left.method || '').localeCompare(String(right.method || ''));
        });
        return hits;
    }

    function buildApiQuickReference(limit) {
        var maxItems = Math.max(20, Math.min(140, Number(limit || 100) || 100));
        var catalog = globalRoot.R7_API_REFERENCE_CATALOG && typeof globalRoot.R7_API_REFERENCE_CATALOG === 'object'
            ? globalRoot.R7_API_REFERENCE_CATALOG
            : null;
        if (!catalog || !catalog.objects || typeof catalog.objects !== 'object') return '';

        var preferredOrder = [
            'ApiInterface',
            'ApiWorksheet',
            'ApiRange',
            'ApiWorkbook',
            'ApiWorksheetFunction',
            'ApiName'
        ];
        var preferredMethods = {
            ApiInterface: {
                AddSheet: 220,
                GetSheet: 220,
                GetSheets: 210,
                GetActiveSheet: 210,
                RecalculateAllFormulas: 120
            },
            ApiWorksheet: {
                GetRange: 220,
                GetRangeByNumber: 200,
                GetUsedRange: 180,
                SetActive: 170,
                GetName: 120,
                GetRows: 110,
                GetCols: 110,
                GetCells: 110
            },
            ApiRange: {
                SetValue: 260,
                GetValue: 200,
                GetValue2: 200,
                GetText: 180,
                Clear: 150,
                ClearContents: 160,
                SetFormulaArray: 130,
                SetNumberFormat: 130,
                SetBold: 110
            }
        };

        var rows = [];
        Object.keys(catalog.objects).forEach(function (objectName) {
            var methods = Array.isArray(catalog.objects[objectName]) ? catalog.objects[objectName] : [];
            methods.forEach(function (item) {
                var method = normalizeTextPayload(item && item.method || '');
                if (!method) return;
                var args = normalizeTextPayload(item && item.args || '');
                var description = normalizeTextPayload(item && item.description || '');
                var objectPriority = preferredOrder.indexOf(objectName);
                var score = objectPriority === -1 ? 10 : (120 - (objectPriority * 12));
                if (preferredMethods[objectName] && preferredMethods[objectName][method]) {
                    score += preferredMethods[objectName][method];
                }
                if (/SetValue|GetRange|GetSheet|AddSheet|GetUsedRange|GetValue/i.test(method)) score += 60;
                rows.push({
                    object: objectName,
                    method: method,
                    args: args,
                    description: description,
                    score: score
                });
            });
        });

        rows.sort(function (left, right) {
            if (right.score !== left.score) return right.score - left.score;
            if (left.object !== right.object) return left.object.localeCompare(right.object);
            return left.method.localeCompare(right.method);
        });

        var seen = {};
        var selected = [];
        for (var i = 0; i < rows.length; i += 1) {
            var row = rows[i];
            var key = row.object + '.' + row.method;
            if (seen[key]) continue;
            seen[key] = true;
            selected.push(row);
            if (selected.length >= maxItems) break;
        }

        if (!selected.length) return '';
        return selected.map(function (row, index) {
            var signature = row.object + '.' + row.method + '(' + row.args + ')';
            var desc = row.description ? (' - ' + truncateInlineText(row.description, 90)) : '';
            return (index + 1) + '. ' + signature + desc;
        }).join('\n');
    }

    function executeAnalyzeReferenceMacrosStep(args) {
        var catalog = globalRoot.R7_API_REFERENCE_CATALOG && typeof globalRoot.R7_API_REFERENCE_CATALOG === 'object'
            ? globalRoot.R7_API_REFERENCE_CATALOG
            : null;
        var guide = globalRoot.R7_API_REFERENCE_GUIDE || 'No guide available.';
        if (!catalog) {
            return {
                ok: true,
                data: {
                    info: 'R7/ONLYOFFICE Spreadsheet API guide is available, but structured catalog is missing.',
                    guide: guide,
                    catalog_available: false
                },
                logs: ['api_reference_catalog_missing'],
                error: null
            };
        }
        var safeArgs = args && typeof args === 'object' ? args : {};
        var categories = catalog.categories && typeof catalog.categories === 'object' ? catalog.categories : {};
        var categoryList = Object.keys(categories).map(function (key) {
            var item = categories[key] || {};
            return {
                key: key,
                title: item.title || key,
                count: Number(item.count || (Array.isArray(item.methods) ? item.methods.length : 0) || 0)
            };
        });
        var hits = findApiReferenceMatches(catalog, safeArgs);
        var pack = makeLimitedMethodPack(hits, safeArgs && safeArgs.limit);
        return {
            ok: true,
            data: {
                guide: guide,
                requested_limit: normalizeReferenceMethodLimit(safeArgs && safeArgs.limit),
                requested_category: normalizeTextPayload(safeArgs.category || ''),
                available_categories: categoryList,
                method_search: {
                    query: normalizeTextPayload(safeArgs.method || safeArgs.query || safeArgs.error || ''),
                    total_count: pack.total_count,
                    returned_count: pack.returned_count,
                    truncated: pack.truncated,
                    methods: pack.methods
                }
            },
            logs: ['api_reference_loaded'],
            error: null
        };
    }

    function normalizeWebToolExecution(result, logLabel) {
        var payload = result && typeof result === 'object' ? result : null;
        var hasResults = !!(payload && Array.isArray(payload.results) && payload.results.length);
        var errors = payload && Array.isArray(payload.errors) ? payload.errors : [];
        var firstError = errors.length ? errors[0] : null;
        
        // Fix: If provider specifically set ok=true in the response, respect it.
        var explicitlyOk = payload && payload.ok === true;
        
        var isOk = !!(explicitlyOk || hasResults || errors.length === 0);
        
        debugAgent('normalization_details', {
            logLabel: logLabel,
            hasResults: hasResults,
            errorsCount: errors.length,
            explicitlyOk: explicitlyOk,
            finalOk: isOk
        });
        
        return {
            ok: isOk,
            data: payload,
            logs: [logLabel],
            error: (isOk || !firstError) ? null : String(firstError.message || firstError.code || 'Web tool failed')
        };
    }

    async function executeWebSearchStep(args) {
        var webTools = getWebToolsBridge();
        if (!webTools) {
            return { ok: false, data: null, logs: [], error: 'Web tools bridge is unavailable' };
        }
        
        var rawResult = await webTools.executeWebSearch(args || {});
        debugAgent('web_search_raw_result', rawResult);
        
        return normalizeWebToolExecution(rawResult, 'web_search_executed');
    }

    async function executeWebCrawlingStep(args) {
        var webTools = getWebToolsBridge();
        if (!webTools || typeof webTools.executeWebCrawling !== 'function') {
            return { ok: false, data: null, logs: [], error: 'Web tools client is unavailable' };
        }
        if (typeof webTools.isEnabled === 'function' && !webTools.isEnabled()) {
            return { ok: false, data: null, logs: [], error: 'Web tools provider is not configured' };
        }
        return normalizeWebToolExecution(await webTools.executeWebCrawling(args || {}), 'web_crawling_executed');
    }

    function parseAgentJsonPayload(value, fallbackValue) {
        if (value === null || value === undefined || value === '') {
            return fallbackValue;
        }
        if (typeof value === 'string') {
            try {
                return JSON.parse(value);
            } catch (error) {
                return fallbackValue;
            }
        }
        return value;
    }

    function extractGeneratedImageUrl(response) {
        var payload = response && typeof response === 'object' ? response : {};
        var item = payload && Array.isArray(payload.data) && payload.data[0] ? payload.data[0] : {};
        var directUrl = normalizeTextPayload(item.url || '');
        if (directUrl) return directUrl;
        var base64 = normalizeTextPayload(item.b64_json || item.base64 || '');
        if (!base64) return '';
        return base64.indexOf('data:image') === 0 ? base64 : ('data:image/png;base64,' + base64);
    }

    async function insertGeneratedImageIntoWord(url, args) {
        var host = getHostBridge();
        if (!host || typeof host.callEditorCommand !== 'function') {
            return { ok: false, data: null, logs: [], error: 'Host bridge is unavailable' };
        }
        var scopeData = {
            url: url,
            widthPx: Math.max(256, Math.min(1024, Number(args && args.widthPx || 768) || 768)),
            heightPx: Math.max(256, Math.min(1024, Number(args && args.heightPx || 512) || 512)),
            caption: normalizeTextPayload(args && args.caption || ''),
            sectionTitle: normalizeTextPayload(args && args.sectionTitle || '')
        };
        var result = await host.callEditorCommand(function () {
            try {
                var oDocument = Api.GetDocument ? Api.GetDocument() : null;
                if (!oDocument) {
                    return JSON.stringify({ inserted: false, error: 'Document is not available' });
                }

                var blocks = [];
                if (Asc.scope.sectionTitle) {
                    var heading = Api.CreateParagraph();
                    heading.AddText(String(Asc.scope.sectionTitle));
                    blocks.push(heading);
                }

                var imageParagraph = Api.CreateParagraph();
                var width = Number(Asc.scope.widthPx || 768) * (25.4 / 96.0) * 36000;
                var height = Number(Asc.scope.heightPx || 512) * (25.4 / 96.0) * 36000;
                var drawing = Api.CreateImage(String(Asc.scope.url), width, height);
                imageParagraph.AddDrawing(drawing);
                blocks.push(imageParagraph);

                if (Asc.scope.caption) {
                    var captionParagraph = Api.CreateParagraph();
                    captionParagraph.AddText(String(Asc.scope.caption));
                    blocks.push(captionParagraph);
                }

                oDocument.InsertContent(blocks);
                return JSON.stringify({
                    inserted: true,
                    blocks: blocks.length
                });
            } catch (error) {
                return JSON.stringify({
                    inserted: false,
                    error: String(error && error.message ? error.message : error)
                });
            }
        }, scopeData, {
            timeoutMs: Number(constants().agentLimits.stepTimeoutMs || 300000)
        });

        var parsed = parseAgentJsonPayload(result, { inserted: false, error: 'Image insertion returned empty result' });
        if (parsed && parsed.inserted === true) {
            return {
                ok: true,
                data: parsed,
                logs: ['image_inserted'],
                error: null
            };
        }
        return {
            ok: false,
            data: parsed,
            logs: ['image_insert_failed'],
            error: normalizeTextPayload(parsed && parsed.error || 'Image insertion failed')
        };
    }

    async function executeGenerateImageAssetStep(args) {
        var imageBridge = root.platform && root.platform.image ? root.platform.image : null;
        var prompt = normalizeTextPayload(args && args.prompt || '');
        if (!prompt.length) {
            return { ok: false, data: null, logs: [], error: 'Image prompt is empty' };
        }
        if (!imageBridge || typeof imageBridge.generate !== 'function') {
            return { ok: false, data: null, logs: [], error: 'Image generation bridge is unavailable' };
        }
        var response;
        try {
            response = await imageBridge.generate(prompt, args && typeof args === 'object' ? {
                size: args.size || ''
            } : null, null);
        } catch (error) {
            return {
                ok: false,
                data: null,
                logs: ['image_generation_failed'],
                error: String(error && error.message ? error.message : error)
            };
        }

        var url = extractGeneratedImageUrl(response);
        if (!url.length) {
            return { ok: false, data: response || null, logs: ['image_generation_empty'], error: 'Image generation returned no URL' };
        }

        var editorType = getCurrentEditorType();
        if (editorType === 'word') {
            var insertResult = await insertGeneratedImageIntoWord(url, args || {});
            if (!insertResult.ok) return insertResult;
        }

        return {
            ok: true,
            data: {
                prompt: prompt,
                url: url,
                caption: normalizeTextPayload(args && args.caption || ''),
                sectionTitle: normalizeTextPayload(args && args.sectionTitle || ''),
                inserted: editorType === 'word'
            },
            logs: ['image_generation_executed'],
            error: null
        };
    }

    async function executeAgentStep(step) {
        var startedAt = Date.now();
        var result = null;
        if (isResearchBlockedForStep(step)) {
            result = buildResearchBlockedResult();
            return {
                step_id: step.id,
                step_type: step.type,
                ok: false,
                data: null,
                logs: Array.isArray(result.logs) ? result.logs : [],
                error: result.error ? String(result.error) : null,
                duration_ms: Date.now() - startedAt
            };
        }
        try {
            result = await withTimeout((async function () {
                if (step.type === 'collect_context') return executeCollectContextStep(step.args || {});
                if (step.type === 'read_document_snapshot') return executeReadDocumentSnapshotStep(step.args || {});
                if (step.type === 'generate_image_asset') return executeGenerateImageAssetStep(step.args || {});
                if (step.type === 'get_current_time') return executeGetCurrentTimeStep();
                if (step.type === 'list_sheets') return executeListSheetsStep();
                if (step.type === 'read_active_sheet') return executeReadActiveSheetStep();
                if (step.type === 'read_document') return executeReadDocumentStep();
                if (step.type === 'read_sheet_range') return executeReadSheetRangeStep(step.args || {});
                if (step.type === 'read_attached_sheet') return executeReadAttachedSheetStep(step.args || {});
                if (step.type === 'write_cell_text') return executeWriteCellTextStep(step.args || {});
                if (step.type === 'write_cells_batch') return executeWriteCellsBatchStep(step.args || {});
                if (step.type === 'write_data_to_sheet') return executeWriteDataToSheetStep(step.args || {});
                if (step.type === 'copy_attached_sheet') return executeCopyAttachedSheetStep(step.args || {});
                if (step.type === 'list_host_tools') return executeListHostToolsStep(step.args || {});
                if (step.type === 'call_host_tool') return executeCallHostToolStep(step.args || {});
                if (step.type === 'web_search') return executeWebSearchStep(step.args || {});
                if (step.type === 'web_crawling') return executeWebCrawlingStep(step.args || {});
                if (step.type === 'analyze_reference_macros') return executeAnalyzeReferenceMacrosStep(step.args || {});
                if (step.type === 'run_macro_code') return executeMacroCode(step.macro_code);
                return { ok: false, data: null, logs: [], error: 'Unsupported executable step type: ' + step.type };
            })(), constants().agentLimits.stepTimeoutMs, 'Step "' + step.type + '"');
        } catch (error) {
            result = { ok: false, data: null, logs: [], error: String(error && error.message ? error.message : error) };
        }

        if (step.type === 'run_macro_code' && step.args && step.args.fast_path && (!result || result.ok !== true) && isTransientStepError(result && result.error)) {
            for (var retry = 1; retry <= 2; retry += 1) {
                await sleep(120 * retry);
                try {
                    result = await withTimeout(executeMacroCode(step.macro_code), constants().agentLimits.stepTimeoutMs, 'Fast-path retry "' + step.type + '"');
                } catch (error2) {
                    result = { ok: false, data: null, logs: [], error: String(error2 && error2.message ? error2.message : error2) };
                }
                if (result && result.ok === true) break;
                if (!isTransientStepError(result && result.error)) break;
            }
        }

        return {
            step_id: step.id,
            step_type: step.type,
            ok: result.ok === true,
            data: result.data === undefined ? null : result.data,
            logs: Array.isArray(result.logs) ? result.logs : [],
            error: result.error ? String(result.error) : null,
            duration_ms: Date.now() - startedAt
        };
    }

    function extractMissingMethodName(errorText) {
        var text = normalizeTextPayload(errorText || '');
        if (!text) return '';
        var match = text.match(/(?:[A-Za-z_$][\w$]*\.)*([A-Za-z_$][\w$]*)\s+is not a function/i);
        return match && match[1] ? String(match[1]) : '';
    }

    function buildMacroRecoveryPlannerHint(stepResult) {
        var errorText = normalizeTextPayload(stepResult && stepResult.error || '');
        var editorType = getCurrentEditorType();
        if (!errorText) return '';
        if (/macro_code is empty/i.test(errorText)) {
            return 'RECOVERY_HINT: The previous run_macro_code step was invalid because macro_code was empty. Return a corrected step with the JavaScript macro in the TOP-LEVEL field "macro_code", not inside args.';
        }
        if (editorType === 'word' && /setheading|verified word api method/i.test(errorText)) {
            return 'RECOVERY_HINT: The previous Word macro used an unverified heading method. In this runtime, create headings manually with verified methods only: var p = Api.CreateParagraph(); var r = p.AddText("Heading"); if (r && typeof r.SetBold === "function") r.SetBold(true); if (typeof p.SetFontSize === "function") p.SetFontSize(28); Api.GetDocument().InsertContent([p]); Never use SetHeading and never write error/debug text into the document.';
        }
        if (/is not a function/i.test(errorText)) {
            var missingMethod = extractMissingMethodName(errorText);
            if (editorType === 'word') {
                if (missingMethod) {
                    return 'RECOVERY_HINT: The previous Word macro used an unverified method "' + missingMethod + '". Use only verified Word methods from docs/R7_WORD_MACRO_API_GUIDE.md, such as Api.GetDocument(), Api.CreateParagraph(), paragraph.AddText(), run.SetBold(), paragraph.SetFontSize(), Api.CreateImage(), paragraph.AddDrawing(), and document.InsertContent(). Never use "' + missingMethod + '" unless it is explicitly verified.';
                }
                return 'RECOVERY_HINT: The previous Word macro used an unverified method. Use only verified Word methods from docs/R7_WORD_MACRO_API_GUIDE.md and create headings manually with bold text plus font size. Never write error/debug text into the document.';
            }
            if (missingMethod) {
                return 'RECOVERY_HINT: The previous macro used an unverified method "' + missingMethod + '". Before the next write attempt, inspect the local spreadsheet API reference loaded from scripts/api_reference.js. Prefer analyze_reference_macros with method "' + missingMethod + '" or a relevant category, then generate a corrected run_macro_code using only verified methods.';
            }
            return 'RECOVERY_HINT: The previous macro used an unverified method. Before the next write attempt, inspect the local spreadsheet API reference loaded from scripts/api_reference.js via analyze_reference_macros, then generate a corrected run_macro_code using only verified methods.';
        }
        if (editorType === 'word') {
            return 'RECOVERY_HINT: The previous Word run_macro_code step failed. Use only verified methods from docs/R7_WORD_MACRO_API_GUIDE.md, keep formatting simple, and do not write status/error text into the document body.';
        }
        return 'RECOVERY_HINT: The previous run_macro_code step failed. If you are unsure about the API, inspect the local spreadsheet API reference loaded from scripts/api_reference.js via analyze_reference_macros before emitting the next macro.';
    }

    function buildHostToolRecoveryPlannerHint(step, stepResult) {
        var host = getHostBridge();
        var editorType = host && typeof host.getEditorTypeSafe === 'function' ? host.getEditorTypeSafe() : '';
        var toolName = extractHostToolName(step && step.args);
        var errorText = normalizeTextPayload(stepResult && stepResult.error || '');
        var prefix = toolName
            ? ('RECOVERY_HINT: The previous call_host_tool step for "' + toolName + '" failed.')
            : 'RECOVERY_HINT: The previous call_host_tool step failed.';

        if (/unavailable|not available|catalog/i.test(errorText)) {
            if (editorType === 'word') {
                return prefix + ' The runtime does not expose a usable native tool path right now. Switch to run_macro_code for document automation.';
            }
            if (editorType === 'cell') {
                return prefix + ' The runtime does not expose a usable native tool path right now. Switch to run_macro_code for spreadsheet automation.';
            }
        }

        if (editorType === 'word') {
            return prefix + ' If a different catalog tool matches the request more clearly, use it. Otherwise switch to run_macro_code for precise insertion, formatting, or long-form writing.';
        }
        if (editorType === 'cell') {
            return prefix + ' If a different catalog tool matches the request more clearly, use it. Otherwise switch to run_macro_code for range-heavy, algorithmic, or verification-heavy spreadsheet work.';
        }
        return prefix + ' Either choose a different host tool from the runtime catalog or continue with run_macro_code if native tools are insufficient.';
    }

    function buildWebSearchRecoveryPlannerHint(step, stepResult) {
        var policy = agentState().researchPolicy || createDefaultResearchPolicy();
        var failedQuery = normalizeTextPayload(step && step.args && step.args.query || '');
        var simplifiedQuery = buildResearchQuery(agentState().lastUserMessage || '');
        var errorText = normalizeTextPayload(stepResult && stepResult.error || '');

        if (policy.required) {
            if (simplifiedQuery && simplifiedQuery !== failedQuery) {
                return 'RECOVERY_HINT: The previous web_search failed. Retry web_search with a shorter topic-only query such as "' + simplifiedQuery + '". Remove plan, image, and formatting instructions from the query.';
            }
            return 'RECOVERY_HINT: The previous web_search failed. If you retry, use a very short topic-only query. If search remains unavailable, do not fabricate facts and explain the web research failure in final_answer.';
        }

        if (/query/i.test(errorText)) {
            return 'RECOVERY_HINT: The previous web_search query was likely malformed or too verbose. Retry with a shorter topic-only query.';
        }
        return 'RECOVERY_HINT: The previous web_search failed. Retry with a simpler query or continue without web facts if they are not required.';
    }

    function buildForcedWebSearchRecoveryStep(step, nextStepIndex) {
        var policy = agentState().researchPolicy || createDefaultResearchPolicy();
        if (!policy.required) return null;
        var currentQuery = normalizeTextPayload(step && step.args && step.args.query || '');
        var simplifiedQuery = buildResearchQuery(agentState().lastUserMessage || '');
        if (!simplifiedQuery || simplifiedQuery === currentQuery) return null;
        return {
            id: 'step_' + (nextStepIndex + 1),
            type: 'web_search',
            reason: 'Retry web research with a shorter topic-only query after the previous search failed.',
            args: {
                query: simplifiedQuery
            }
        };
    }

    function buildForcedApiReferenceRecoveryStep(stepResult, nextStepIndex) {
        var errorText = normalizeTextPayload(stepResult && stepResult.error || '');
        if (!/is not a function/i.test(errorText)) return null;
        if (getCurrentEditorType() === 'word') return null;
        var missingMethod = extractMissingMethodName(errorText);
        return {
            id: 'step_' + (nextStepIndex + 1),
            type: 'analyze_reference_macros',
            reason: missingMethod
                ? ('Previous macro used unverified method "' + missingMethod + '". Search the local API catalog for verified alternatives before retrying.')
                : 'Previous macro used an unverified method. Search the local API catalog before retrying.',
            args: missingMethod
                ? { method: missingMethod, limit: 12 }
                : { category: 'sheet_workbook', limit: 12 }
        };
    }

    function clearPendingPlan() {
        agentState().pendingPlan = null;
        if (agentState().wordPlanMode) {
            agentState().wordPlanMode.awaitingApproval = false;
        }
    }

    function presentPlanToUser(step, plannerMessages) {
        var ui = getUi();
        var plan = normalizeWordPlan(step && step.args && step.args.plan, step && step.id);
        agentState().pendingPlan = {
            stepId: normalizeTextPayload(step && step.id || ''),
            plan: cloneSerializable(plan),
            plannerMessages: cloneMessages(plannerMessages),
            runContainer: agentState().currentRunContainer || null
        };
        agentState().wordPlanMode.awaitingApproval = true;
        setStatus('awaiting_plan');
        addTraceRecord({
            step_id: step.id,
            step_type: step.type,
            status: 'awaiting_approval',
            reason: 'Plan is ready and waiting for user approval.'
        });
        if (ui && typeof ui.displayWordPlanApproval === 'function') {
            ui.displayWordPlanApproval(plan, agentState().currentRunContainer);
        } else if (ui && typeof ui.displayMessage === 'function') {
            ui.displayMessage('Plan is ready. Approve it or send edits in chat.', 'assistant', true, agentState().currentRunContainer);
        }
    }

    function resumeFromPlanDecision(decision, note) {
        if (!agentState().pendingPlan) return Promise.resolve(null);
        var pending = agentState().pendingPlan;
        var plan = cloneSerializable(pending.plan);
        var plannerMessages = cloneMessages(pending.plannerMessages);
        var normalizedDecision = normalizeTextPayload(decision || '').toLowerCase();
        var normalizedNote = normalizeTextPayload(note || '');

        clearPendingPlan();
        agentState().stopRequested = false;
        if (normalizedDecision === 'approve') {
            agentState().wordPlanMode.approved = true;
            agentState().wordPlanMode.executionCountAfterApproval = 0;
            agentState().wordPlanMode.approvedPlan = cloneSerializable(plan);
            addTraceRecord({
                step_id: 'plan_approved',
                status: 'approved',
                reason: 'User approved the execution plan.'
            });
            plannerMessages.push({
                role: 'user',
                content: buildPlanDecisionMessage('approved', plan, 'Plan approved. Execute the document in order, generate the approved images, and finish only after the automation is done.')
            });
        } else if (normalizedDecision === 'revise') {
            agentState().wordPlanMode.approved = false;
            agentState().wordPlanMode.executionCountAfterApproval = 0;
            agentState().wordPlanMode.revisionCount = Number(agentState().wordPlanMode.revisionCount || 0) + 1;
            addTraceRecord({
                step_id: 'plan_revision',
                status: 'revised',
                reason: normalizedNote || 'User requested plan changes.'
            });
            plannerMessages.push({
                role: 'user',
                content: buildPlanDecisionMessage('revise', plan, normalizedNote || 'Revise the plan and present an updated version for approval.')
            });
        } else {
            agentState().wordPlanMode.approved = false;
            agentState().wordPlanMode.executionCountAfterApproval = 0;
            addTraceRecord({
                step_id: 'plan_cancelled',
                status: 'stopped',
                reason: normalizedNote || 'User cancelled the plan.'
            });
            setStatus('idle');
            return Promise.resolve({ cancelled: true });
        }

        setStatus('planning');
        return runAgentLoopCore(plannerMessages, agentState().steps.length);
    }

    async function runAgentLoopCore(plannerMessages, startIteration) {
        var chat = getChatService();
        var ui = getUi();
        var consecutiveStepErrors = 0;
        var plannerOverflowRecoveries = 0;

        for (var iteration = startIteration; iteration < constants().agentLimits.maxIterations; iteration += 1) {
            if (agentState().stopRequested) {
                addTraceRecord({ step_id: 'agent_stop', status: 'stopped', reason: 'Stop requested by user' });
                setStatus('idle');
                return;
            }

            setStatus('planning');
            var step = null;
            if (Array.isArray(agentState().recoveryQueue) && agentState().recoveryQueue.length) {
                step = agentState().recoveryQueue.shift();
            }
            if (!step && Array.isArray(agentState().fastPathQueue) && agentState().fastPathQueue.length) {
                step = agentState().fastPathQueue.shift();
                agentState().fastPathActive = true;
            }
            if (!step) {
                step = getForcedInitialStep(agentState().lastUserMessage, agentState().steps.length);
            }
            try {
                if (!step) {
                    var preflightCompaction = compactPlannerMessagesInPlace(plannerMessages, { reason: 'preflight' });
                    if (preflightCompaction.changed) {
                        addTraceRecord({
                            step_id: 'planner_context_compacted_preflight',
                            status: 'warning',
                            reason: 'Planner context compacted before planning request. removed=' + preflightCompaction.removedMessages + ', chars=' + preflightCompaction.totalChars
                        });
                    }
                    step = await plannerNextStep(plannerMessages, agentState().steps.length);
                }
            } catch (error) {
                if (agentState().stopRequested || (error && error.name === 'AbortError')) {
                    addTraceRecord({ step_id: 'agent_stop', status: 'stopped', reason: 'Stopped during planning' });
                    setStatus('idle');
                    return;
                }
                if (isContextOverflowError(error) && plannerOverflowRecoveries < Math.max(1, Number(constants().agentLimits.maxContextOverflowRecoveries || 2))) {
                    plannerOverflowRecoveries += 1;
                    var overflowCompaction = compactPlannerMessagesInPlace(plannerMessages, { reason: 'overflow_recovery_' + plannerOverflowRecoveries });
                    if (overflowCompaction.changed) {
                        addTraceRecord({
                            step_id: 'planner_context_overflow_recovery',
                            status: 'warning',
                            reason: 'Planner context overflow detected and compacted automatically. attempt=' + plannerOverflowRecoveries + ', removed=' + overflowCompaction.removedMessages + ', chars=' + overflowCompaction.totalChars
                        });
                        plannerMessages.push({
                            role: 'user',
                            content: 'RUNTIME_POLICY: Planner context was compacted after token overflow. Continue from current run state, do not restart, and keep next steps concise.'
                        });
                        continue;
                    }
                }
                addTraceRecord({
                    step_id: 'planner_error',
                    status: 'error',
                    error: formatPlannerError(error),
                    reason: 'Planner failed'
                });
                setStatus('error');
                if (ui.displayMessage) {
                    ui.displayMessage(formatPlannerError(error), 'assistant', true, agentState().currentRunContainer);
                }
                return;
            }

            if (shouldForcePlanPresentation(step)) {
                addTraceRecord({
                    step_id: 'word_plan_required',
                    status: 'recovery',
                    reason: 'Runtime requires a present_plan step before Word writing execution.'
                });
                plannerMessages.push({
                    role: 'user',
                    content: 'RUNTIME_POLICY: For this Word request you must first return a present_plan step with section titles, image slots, image prompts, and optional sources. Do not write the document yet.'
                });
                continue;
            }

            if (step.type === 'web_search' && (agentState().researchPolicy && agentState().researchPolicy.unavailable)) {
                addTraceRecord({
                    step_id: 'web_search_disabled',
                    status: 'recovery',
                    reason: 'Runtime disabled further web_search steps for this run after repeated failures.'
                });
                plannerMessages.push({
                    role: 'user',
                    content: markWebResearchUnavailable()
                });
                continue;
            }

            if (shouldDeferFinalAnswerUntilExecution(step)) {
                addTraceRecord({
                    step_id: 'word_plan_execute_required',
                    status: 'recovery',
                    reason: 'Plan was approved, but no execution step has run yet.'
                });
                plannerMessages.push({
                    role: 'user',
                    content: 'RUNTIME_POLICY: The plan is approved but no document automation has been executed yet. Return an executable step such as run_macro_code or generate_image_asset instead of final_answer.'
                });
                continue;
            }

            if (shouldBlockGeneratedSpreadsheetInspectionMacro(step)) {
                addTraceRecord({
                    step_id: step.id + '_spreadsheet_read_required',
                    step_type: step.type,
                    status: 'recovery',
                    reason: 'Runtime blocked a generated inspection macro and requires the predefined spreadsheet read step path first.'
                });
                plannerMessages.push({
                    role: 'user',
                    content: 'RUNTIME_POLICY: Do not generate exploratory macros for spreadsheet inspection/copy. Use predefined steps first. For attached workbook copy use copy_attached_sheet. For normal sheet reads use read_active_sheet, read_sheet_range, or list_sheets.'
                });
                continue;
            }

            if (shouldRestrictStepToAttachedCopyFlow(step)) {
                addTraceRecord({
                    step_id: step.id + '_attached_copy_required',
                    step_type: step.type,
                    status: 'recovery',
                    reason: 'Runtime enforces attached workbook copy flow via copy_attached_sheet for this request.'
                });
                plannerMessages.push({
                    role: 'user',
                    content: 'RUNTIME_POLICY: For this request do NOT use run_macro_code, analyze_reference_macros, list_sheets, read_sheet_range, or read_active_sheet. Emit copy_attached_sheet as the next step with filename/sourceSheetName/targetSheetName/clearFirst, then final_answer.'
                });
                continue;
            }

            if (shouldEnforceSpreadsheetToolsBeforeMacros(step)) {
                addTraceRecord({
                    step_id: step.id + '_tool_first_required',
                    step_type: step.type,
                    status: 'recovery',
                    reason: 'Runtime enforces tool-first policy before run_macro_code for spreadsheet write tasks.'
                });
                plannerMessages.push({
                    role: 'user',
                    content: 'RUNTIME_POLICY: Use tool-first spreadsheet execution. If a matching host tool exists, use list_host_tools then call_host_tool. Otherwise choose one of write_cell_text, write_cells_batch, write_data_to_sheet, or copy_attached_sheet based on task size. Use run_macro_code only if tool execution fails or no tool supports the requested action.'
                });
                continue;
            }

            if (shouldEnforceSpreadsheetToolsBeforeReference(step)) {
                addTraceRecord({
                    step_id: step.id + '_tool_first_reference_required',
                    step_type: step.type,
                    status: 'recovery',
                    reason: 'Runtime enforces tool-first policy before analyze_reference_macros for spreadsheet write tasks.'
                });
                plannerMessages.push({
                    role: 'user',
                    content: 'RUNTIME_POLICY: Skip analyze_reference_macros for now. First execute a host tool or predefined write tool: list_host_tools/call_host_tool, write_cell_text, write_cells_batch, write_data_to_sheet, or copy_attached_sheet. Use analyze_reference_macros only if the selected tool fails.'
                });
                continue;
            }

            if (shouldRequireCallHostToolAfterDiscovery(step)) {
                addTraceRecord({
                    step_id: step.id + '_host_tool_call_required',
                    step_type: step.type,
                    status: 'recovery',
                    reason: 'Runtime requires call_host_tool after successful host tool discovery for this spreadsheet write task.'
                });
                plannerMessages.push({
                    role: 'user',
                    content: buildHostToolCallRequiredPrompt()
                });
                continue;
            }

            agentState().currentStepIndex = agentState().steps.length;
            agentState().steps.push(step);
            addTraceRecord({
                step_id: step.id,
                step_type: step.type,
                status: 'planned',
                reason: step.reason || '',
                tool_name: step.type === 'call_host_tool' ? extractHostToolName(step.args) : '',
                tool_input_preview: step.type === 'call_host_tool' ? buildHostToolInputPreview(step.args && step.args.input) : ''
            });
            if (step.type === 'run_macro_code') {
                var macroMeta = buildMacroTraceMeta(step.macro_code, 420);
                addTraceRecord({
                    step_id: step.id + '_macro_generated',
                    step_type: step.type,
                    status: 'macro_generated',
                    reason: 'Planner generated macro and sent it to client executor',
                    macro_code_length: macroMeta.length,
                    macro_code_preview: macroMeta.preview
                });
            }

            plannerMessages.push({ role: 'assistant', content: JSON.stringify(step) });
            if (step.type === 'present_plan') {
                presentPlanToUser(step, plannerMessages);
                return;
            }
            if (step.type === 'final_answer') {
                var forcedVerificationStep = hasPendingPostMacroVerification()
                    ? buildForcedPostMacroVerificationStep(agentState().steps.length)
                    : null;
                if (forcedVerificationStep) {
                    agentState().steps.pop();
                    plannerMessages.pop();
                    if (Array.isArray(agentState().trace) && agentState().trace.length) {
                        var lastTrace = agentState().trace[agentState().trace.length - 1];
                        if (lastTrace && lastTrace.step_id === step.id && lastTrace.status === 'planned') {
                            agentState().trace.pop();
                        }
                    }
                    agentState().recoveryQueue = Array.isArray(agentState().recoveryQueue) ? agentState().recoveryQueue : [];
                    agentState().recoveryQueue.unshift(forcedVerificationStep);
                    addTraceRecord({
                        step_id: forcedVerificationStep.id,
                        step_type: forcedVerificationStep.type,
                        status: 'planned',
                        reason: 'Runtime forced verification after macro execution before allowing final_answer.'
                    });
                    continue;
                }
                var forcedSpreadsheetReadStep = shouldForceSpreadsheetReadBeforeAnswer(step)
                    ? buildForcedSpreadsheetReadStep(agentState().lastUserMessage || '', agentState().steps.length, 'read_before_answer')
                    : null;
                if (forcedSpreadsheetReadStep) {
                    agentState().steps.pop();
                    plannerMessages.pop();
                    if (Array.isArray(agentState().trace) && agentState().trace.length) {
                        var pendingTrace = agentState().trace[agentState().trace.length - 1];
                        if (pendingTrace && pendingTrace.step_id === step.id && pendingTrace.status === 'planned') {
                            agentState().trace.pop();
                        }
                    }
                    agentState().recoveryQueue = Array.isArray(agentState().recoveryQueue) ? agentState().recoveryQueue : [];
                    agentState().recoveryQueue.unshift(forcedSpreadsheetReadStep);
                    addTraceRecord({
                        step_id: forcedSpreadsheetReadStep.id,
                        step_type: forcedSpreadsheetReadStep.type,
                        status: 'planned',
                        reason: 'Runtime forced spreadsheet context collection before allowing final_answer.'
                    });
                    plannerMessages.push({
                        role: 'user',
                        content: 'RUNTIME_POLICY: Before answering spreadsheet read or analysis requests, you must use the predefined spreadsheet read step path such as read_active_sheet, read_sheet_range, or list_sheets. Do not answer from sheet metadata alone and do not switch to run_macro_code just to inspect cells.'
                    });
                    continue;
                }
                var finalAnswer = extractFinalAnswer(step) || 'Done.';
                var imageCommand = extractImageCommand(finalAnswer);
                setStatus('answering');
                addTraceRecord({
                    step_id: step.id,
                    step_type: step.type,
                    status: 'completed',
                    reason: 'Final answer produced'
                });
                if (ui.displayMessage) {
                    if (!imageCommand || imageCommand.visibleText) {
                        ui.displayMessage(imageCommand ? imageCommand.visibleText : finalAnswer, 'assistant', true, agentState().currentRunContainer);
                    }

                    if (imageCommand && imageCommand.prompt) {
                        var chat = getChatService();
                        if (chat && typeof chat.generateImageFromText === 'function') {
                            chat.generateImageFromText(imageCommand.prompt, agentState().currentRunContainer);
                        }
                    }
                }
                setStatus('idle');
                if (chat && typeof chat.stopChatRequest === 'function') {
                    chat.stopChatRequest();
                }
                return;
            }

            var plannerSnapshot = cloneMessages(plannerMessages);
            setStatus('executing');
            var stepResult = await executeAgentStep(step);
            if (step.type === 'web_search' && agentState().researchPolicy && agentState().researchPolicy.required) {
                agentState().researchPolicy.attempted = true;
                agentState().researchPolicy.completed = stepResult.ok === true;
                agentState().researchPolicy.failed = stepResult.ok !== true;
            }
            addTraceRecord({
                step_id: step.id,
                step_type: step.type,
                status: stepResult.ok ? 'ok' : 'error',
                duration_ms: stepResult.duration_ms,
                error: stepResult.error || '',
                tool_name: step.type === 'call_host_tool' ? extractHostToolName(step.args) : '',
                tool_input_preview: step.type === 'call_host_tool' ? buildHostToolInputPreview(step.args && step.args.input) : ''
            });
            plannerMessages.push({ role: 'user', content: buildToolResultMessage(stepResult) });
            if (!stepResult.ok && step.type === 'run_macro_code') {
                var macroRecoveryHint = buildMacroRecoveryPlannerHint(stepResult);
                if (macroRecoveryHint) {
                    plannerMessages.push({ role: 'user', content: macroRecoveryHint });
                }
            }
            if (!stepResult.ok && step.type === 'call_host_tool') {
                var hostToolRecoveryHint = buildHostToolRecoveryPlannerHint(step, stepResult);
                if (hostToolRecoveryHint) {
                    plannerMessages.push({ role: 'user', content: hostToolRecoveryHint });
                }
            }
            if (!stepResult.ok && step.type === 'web_search') {
                var webSearchRecoveryHint = buildWebSearchRecoveryPlannerHint(step, stepResult);
                if (webSearchRecoveryHint) {
                    plannerMessages.push({ role: 'user', content: webSearchRecoveryHint });
                }
            }
            var postToolResultCompaction = compactPlannerMessagesInPlace(plannerMessages, { reason: 'post_tool_result' });
            if (postToolResultCompaction.changed && postToolResultCompaction.removedMessages > 0) {
                addTraceRecord({
                    step_id: 'planner_context_compacted_post_tool_result',
                    status: 'warning',
                    reason: 'Planner context compacted after tool result. removed=' + postToolResultCompaction.removedMessages + ', chars=' + postToolResultCompaction.totalChars
                });
            }

            if (!stepResult.ok) {
                if (step && step.args && step.args.fast_path) {
                    agentState().fastPathQueue = [];
                    agentState().fastPathActive = false;
                }
                consecutiveStepErrors += 1;
                if (shouldDegradeWebResearchFailure(step, stepResult, consecutiveStepErrors)) {
                    var researchUnavailablePolicy = markWebResearchUnavailable(stepResult);
                    addTraceRecord({
                        step_id: step.id + '_research_unavailable',
                        step_type: step.type,
                        status: 'warning',
                        reason: 'Repeated web_search failures; continuing without further web research in this run.'
                    });
                    plannerMessages.push({ role: 'user', content: researchUnavailablePolicy });
                    consecutiveStepErrors = 0;
                    agentState().lastFailedStep = null;
                    continue;
                }
                agentState().lastFailedStep = {
                    step: step,
                    plannerMessages: plannerSnapshot,
                    iteration: iteration
                };
                var canRecover = iteration < constants().agentLimits.maxIterations - 1 &&
                    consecutiveStepErrors < constants().agentLimits.maxConsecutiveStepErrors;
                if (canRecover) {
                    var recoveryStep = step.type === 'run_macro_code'
                        ? buildForcedApiReferenceRecoveryStep(stepResult, agentState().steps.length)
                        : null;
                    if (!recoveryStep && step.type === 'web_search') {
                        recoveryStep = buildForcedWebSearchRecoveryStep(step, agentState().steps.length);
                    }
                    if (recoveryStep) {
                        agentState().recoveryQueue = Array.isArray(agentState().recoveryQueue) ? agentState().recoveryQueue : [];
                        agentState().recoveryQueue.push(recoveryStep);
                        var recoveryReason = step.type === 'web_search'
                            ? 'Forced web search retry with simplified query after search failure.'
                            : 'Forced API reference recovery after macro method failure.';
                        addTraceRecord({
                            step_id: recoveryStep.id,
                            step_type: recoveryStep.type,
                            status: 'planned',
                            reason: recoveryReason
                        });
                    }
                    addTraceRecord({
                        step_id: step.id + '_recovery',
                        step_type: step.type,
                        status: 'recovery',
                        reason: 'Step failed; planner will continue with corrective step.'
                    });
                    continue;
                }
                setStatus('error');
                if (ui.displayMessage) {
                    ui.displayMessage('Step failed repeatedly. Check trace and use Retry step.', 'assistant', true, agentState().currentRunContainer);
                }
                return;
            }

            consecutiveStepErrors = 0;
            plannerOverflowRecoveries = 0;
            agentState().lastFailedStep = null;
            if (agentState().wordPlanMode && agentState().wordPlanMode.approved && countWordPlanExecutionStep(step.type)) {
                agentState().wordPlanMode.executionCountAfterApproval = Number(agentState().wordPlanMode.executionCountAfterApproval || 0) + 1;
            }
            var postStepCompaction = compactPlannerMessagesInPlace(plannerMessages, { reason: 'post_step' });
            if (postStepCompaction.changed && postStepCompaction.removedMessages > 0) {
                addTraceRecord({
                    step_id: 'planner_context_compacted_post_step',
                    status: 'warning',
                    reason: 'Planner context compacted after step execution. removed=' + postStepCompaction.removedMessages + ', chars=' + postStepCompaction.totalChars
                });
            }
            if (!agentState().fastPathQueue.length) {
                agentState().fastPathActive = false;
            }
        }

        addTraceRecord({
            step_id: 'agent_limit',
            status: 'error',
            reason: 'Max iterations reached'
        });
        setStatus('error');
        if (getUi().displayMessage) {
            getUi().displayMessage('Agent loop reached max iterations. Use Retry step or Run from step 1.', 'assistant', true, agentState().currentRunContainer);
        }
    }

    async function run(userMessage, runContainer) {
        var host = getHostBridge();
        if (isBusy()) return;
        if (!host || host.getHostDiagnostics()) {
            setStatus('error');
            return;
        }
        resolveAgentMode();
        if (!shouldUseMacroRunner()) return;

        reset({ preserveTrace: false });
        resolveAgentMode();
        agentState().currentRunContainer = runContainer || null;
        agentState().lastUserMessage = normalizeTextPayload(userMessage || '');
        agentState().wordPlanMode = buildWordPlanModeState(agentState().lastUserMessage);
        agentState().researchPolicy = buildResearchPolicyForRequest(agentState().lastUserMessage, {
            forceRequired: agentState().wordPlanMode.enabled === true
        });
        agentState().fastPathQueue = buildFastPathQueueForMessage(agentState().lastUserMessage, agentState().steps.length);
        agentState().fastPathActive = Array.isArray(agentState().fastPathQueue) && agentState().fastPathQueue.length > 0;
        agentState().stopRequested = false;
        setStatus('planning');
        agentState().runCounter += 1;
        addTraceRecord({
            step_id: 'run_' + agentState().runCounter,
            status: 'started',
            reason: 'New prompt run started'
        });
        addTraceRecord({
            step_id: 'agent_start',
            status: 'started',
            reason: 'Macro runner started'
        });
        addTraceRecord({
            step_id: 'web_search_capability',
            status: 'info',
            reason: buildPlannerWebSearchStatusLine(agentState().researchPolicy)
        });
        addTraceRecord({
            step_id: 'word_plan_mode',
            status: agentState().wordPlanMode && agentState().wordPlanMode.enabled ? 'info' : 'warning',
            reason: agentState().wordPlanMode && agentState().wordPlanMode.enabled
                ? 'word_plan_mode=true; approval_required=true; execution will pause for user confirmation'
                : 'word_plan_mode=false'
        });
        var desktopToolsState = buildDesktopToolsPlannerState(agentState().lastUserMessage);
        addTraceRecord({
            step_id: 'desktop_tools_capability',
            status: desktopToolsState.enabledForPlanner ? 'info' : 'warning',
            reason: desktopToolsState.statusLine
        });
        if (desktopToolsState.catalogHash || desktopToolsState.toolCount) {
            addTraceRecord({
                step_id: 'desktop_tools_catalog',
                status: desktopToolsState.enabledForPlanner ? 'info' : 'warning',
                reason: 'tool_count=' + String(desktopToolsState.toolCount || 0) + '; catalog_hash=' + String(desktopToolsState.catalogHash || 'none')
            });
        }

        var plannerMessages = buildPlannerConversationSeed(agentState().lastUserMessage);
        var apiQuickRef = buildApiQuickReference(100);
        if (apiQuickRef) {
            plannerMessages.unshift({
                role: 'user',
                content: [
                    'API_QUICK_REFERENCE_START',
                    'Top verified spreadsheet API signatures from local catalog (read before planning macros):',
                    apiQuickRef,
                    'API_QUICK_REFERENCE_END'
                ].join('\n')
            });
            addTraceRecord({
                step_id: 'api_quick_reference_loaded',
                status: 'info',
                reason: 'Injected top API quick reference methods into planner context (count up to 100).'
            });
        }
        return runAgentLoopCore(plannerMessages, 0);
    }

    function stop() {
        agentState().stopRequested = true;
        if (agentState().abortController) {
            try { agentState().abortController.abort(); } catch (error) {}
        }
        addTraceRecord({
            step_id: 'agent_stop_requested',
            status: 'stopped',
            reason: 'Stop requested by user'
        });
        if (!isBusy()) {
            setStatus('idle');
        }
    }

    async function retryStep() {
        if (!canRetry()) return;
        var failed = agentState().lastFailedStep;
        agentState().stopRequested = false;
        setStatus('executing');

        var retryResult = await executeAgentStep(failed.step);
        addTraceRecord({
            step_id: failed.step.id,
            step_type: failed.step.type,
            status: retryResult.ok ? 'retry_ok' : 'retry_error',
            duration_ms: retryResult.duration_ms,
            error: retryResult.error || ''
        });

        var plannerMessages = cloneMessages(failed.plannerMessages);
        plannerMessages.push({ role: 'user', content: buildToolResultMessage(retryResult) });

        if (!retryResult.ok) {
            agentState().lastFailedStep = {
                step: failed.step,
                plannerMessages: failed.plannerMessages,
                iteration: failed.iteration
            };
            setStatus('error');
            return;
        }

        agentState().lastFailedStep = null;
        return runAgentLoopCore(plannerMessages, failed.iteration + 1);
    }

    async function runFromStepOne() {
        if (!canRunFromStart()) return;
        return run(agentState().lastUserMessage, agentState().currentRunContainer);
    }

    async function approvePendingPlan() {
        if (!isAwaitingPlanApproval()) return null;
        return resumeFromPlanDecision('approve', '');
    }

    async function revisePendingPlan(note) {
        if (!isAwaitingPlanApproval()) return null;
        return resumeFromPlanDecision('revise', note);
    }

    async function cancelPendingPlan() {
        if (!isAwaitingPlanApproval()) return null;
        return resumeFromPlanDecision('cancel', 'User cancelled the plan.');
    }

    var api = {
        getState: function () { return agentState(); },
        reset: reset,
        resolveAgentMode: resolveAgentMode,
        shouldUseMacroRunner: shouldUseMacroRunner,
        isBusy: isBusy,
        isAwaitingPlanApproval: isAwaitingPlanApproval,
        canRetry: canRetry,
        canRunFromStart: canRunFromStart,
        addTraceRecord: addTraceRecord,
        parsePlannerResponse: parsePlannerResponse,
        normalizePlannerStep: normalizePlannerStep,
        getPlannerSystemMessage: getPlannerSystemMessage,
        plannerNextStep: plannerNextStep,
        isContextOverflowError: isContextOverflowError,
        compactPlannerMessagesInPlace: compactPlannerMessagesInPlace,
        sanitizePlannerMessageContent: sanitizePlannerMessageContent,
        getForcedInitialStep: getForcedInitialStep,
        buildResearchPolicyForRequest: buildResearchPolicyForRequest,
        buildPlannerConversationSeed: buildPlannerConversationSeed,
        isWordPlanModeRequest: isWordPlanModeRequest,
        getPlanApprovalCommandText: getPlanApprovalCommandText,
        isPlanApprovalMessage: isPlanApprovalMessage,
        buildDesktopToolsPlannerState: buildDesktopToolsPlannerState,
        buildImageGenerationPrompt: buildImageGenerationPrompt,
        extractImageCommand: extractImageCommand,
        buildToolResultMessage: buildToolResultMessage,
        makeLimitedMethodPack: makeLimitedMethodPack,
        findApiReferenceMatches: findApiReferenceMatches,
        extractMissingMethodName: extractMissingMethodName,
        buildMacroRecoveryPlannerHint: buildMacroRecoveryPlannerHint,
        buildWebSearchRecoveryPlannerHint: buildWebSearchRecoveryPlannerHint,
        buildHostToolRecoveryPlannerHint: buildHostToolRecoveryPlannerHint,
        buildForcedApiReferenceRecoveryStep: buildForcedApiReferenceRecoveryStep,
        buildForcedWebSearchRecoveryStep: buildForcedWebSearchRecoveryStep,
        buildForcedPostMacroVerificationStep: buildForcedPostMacroVerificationStep,
        buildForcedSpreadsheetReadStep: buildForcedSpreadsheetReadStep,
        hasPendingPostMacroVerification: hasPendingPostMacroVerification,
        hasExecutedSpreadsheetReadStep: hasExecutedSpreadsheetReadStep,
        isSpreadsheetReadIntent: isSpreadsheetReadIntent,
        isSpreadsheetWorkbookDiscoveryIntent: isSpreadsheetWorkbookDiscoveryIntent,
        extractSpreadsheetRangeTarget: extractSpreadsheetRangeTarget,
        extractSpreadsheetSheetName: extractSpreadsheetSheetName,
        shouldForceSpreadsheetReadBeforeAnswer: shouldForceSpreadsheetReadBeforeAnswer,
        shouldBlockGeneratedSpreadsheetInspectionMacro: shouldBlockGeneratedSpreadsheetInspectionMacro,
        executeAgentStep: executeAgentStep,
        run: run,
        stop: stop,
        retryStep: retryStep,
        runFromStepOne: runFromStepOne,
        approvePendingPlan: approvePendingPlan,
        revisePendingPlan: revisePendingPlan,
        cancelPendingPlan: cancelPendingPlan
    };

    root.features.agentRuntime = api;
    root.services.agent = api;
    root.agent = root.agent || {};
    root.agent.run = api.run;
    root.agent.stop = api.stop;
    root.agent.retryStep = api.retryStep;
    root.agent.runFromStepOne = api.runFromStepOne;
    root.agent.approvePendingPlan = api.approvePendingPlan;
    root.agent.revisePendingPlan = api.revisePendingPlan;
    root.agent.cancelPendingPlan = api.cancelPendingPlan;
    root.agent.getPlanApprovalCommandText = api.getPlanApprovalCommandText;
    root.agent.isPlanApprovalMessage = api.isPlanApprovalMessage;

    return api;
});
