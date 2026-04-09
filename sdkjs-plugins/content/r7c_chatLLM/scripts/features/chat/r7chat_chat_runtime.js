(function (window) {
    'use strict';

    var root = window.R7Chat = window.R7Chat || {};
    root.features = root.features || {};
    root.services = root.services || {};
    root.runtime = root.runtime || {};
    root.runtime.state = root.runtime.state || {};
    root.runtime.state.chat = root.runtime.state.chat || {
        apiKey: '',
        hasKey: false,
        conversationHistory: [],
        lastRequest: null,
        abortController: null
    };

    function chatState() {
        return root.runtime.state.chat;
    }

    function dom() {
        return root.runtime && root.runtime.dom ? root.runtime.dom : {};
    }

    function constants() {
        return root.runtime.constants || { defaultModel: 'openrouter/auto' };
    }

    function tr(text) {
        var host = getHostBridge();
        if (host && typeof host.translate === 'function') {
            return host.translate(text);
        }
        return String(text || '');
    }

    function getSettingsService() {
        return root.features && root.features.settings ? root.features.settings : null;
    }

    function getSettingsPanel() {
        return root.features && root.features.settingsPanel ? root.features.settingsPanel : null;
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

    function getHostToolsClient() {
        return root.platform && root.platform.hostTools ? root.platform.hostTools : null;
    }

    function getImageBridge() {
        return root.platform && root.platform.image ? root.platform.image : null;
    }

    function getHostBridge() {
        return root.platform && root.platform.hostBridge ? root.platform.hostBridge : null;
    }

    function getToolLoop() {
        return root.features && root.features.chatToolLoop ? root.features.chatToolLoop : null;
    }

    function getThreadStore() {
        return root.services && root.services.chatThreads ? root.services.chatThreads : null;
    }

    function getAttachmentService() {
        return root.features && root.features.chatAttachments ? root.features.chatAttachments : null;
    }

    function getContextService() {
        return root.services && root.services.context ? root.services.context : null;
    }

    function getAgentService() {
        return root.services && root.services.agent ? root.services.agent : null;
    }

    function getUi() {
        return root.ui || {};
    }

    function isAbortLikeError(error) {
        var name = String(error && error.name || '');
        var message = String(error && error.message || error || '');
        return name === 'AbortError' || /abort|aborted|cancelled|canceled/i.test(message);
    }

    function isRuntimeClosing() {
        return !!(root.runtime && root.runtime.lifecycle && root.runtime.lifecycle.isClosing === true);
    }

    function getErrorTools() {
        return root.shared && root.shared.errors ? root.shared.errors : null;
    }

    function loadRuntimeSettings() {
        var service = getSettingsService();
        if (service && typeof service.loadSettings === 'function') {
            return service.loadSettings();
        }
        return {
            provider: 'openrouter',
            apiKey: '',
            model: constants().defaultModel,
            imageModel: 'google/gemini-2.5-flash-image-preview'
        };
    }

    function saveRuntimeSettings(nextSettings) {
        var service = getSettingsService();
        if (service && typeof service.saveSettings === 'function') {
            return service.saveSettings(null, nextSettings || {});
        }
        return Object.assign({}, loadRuntimeSettings(), nextSettings || {});
    }

    function getMissingAuthMessage() {
        var settings = loadRuntimeSettings();
        var service = getSettingsService();
        var providerLabel = service && typeof service.getProviderLabel === 'function'
            ? service.getProviderLabel(settings.activeProvider || settings.provider)
            : 'Provider';
        return tr('Set your {provider} credentials first').replace('{provider}', providerLabel);
    }

    function checkApiKey(silent) {
        var settings = loadRuntimeSettings();
        chatState().apiKey = settings.apiKey || '';
        chatState().hasKey = !!chatState().apiKey;
        if (!chatState().hasKey && !silent && getUi().displayMessage) {
            getUi().displayMessage(getMissingAuthMessage(), 'assistant', true);
        }
        return chatState().hasKey;
    }

    function decoratePrompt(prompt) {
        return [{
            role: 'user',
            content: prompt
        }];
    }

    function formatTable(data) {
        if (!Array.isArray(data)) return 'Input must be an array.';
        var table = '';
        for (var i = 0; i < data.length; i += 1) {
            var row = '';
            for (var j = 0; j < data[i].length; j += 1) {
                row += data[i][j] + '\t';
            }
            table += row.trim() + '\n';
        }
        return table.trim();
    }

    function getLanguageName() {
        var langMap = root.runtime.langMap || {};
        var lang = root.runtime.lang || 'en';
        return langMap[lang] || 'English';
    }

    function getSystemMessage() {
        var host = getHostBridge();
        var editorType = host && typeof host.getEditorTypeSafe === 'function' ? host.getEditorTypeSafe() : '';
        var language = getLanguageName();
        var skillDescription = '\n\n[SKILL: IMAGE_GENERATION]\n' +
            'You have a specialized skill to generate images. To use it, output the following command: [GENERATE_IMAGE: detailed prompt].\n' +
            'When a user asks for an image, a drawing, or a visualization, you MUST use the [GENERATE_IMAGE: ...] command. ' +
            'Do NOT use ASCII art or text drawings unless the user explicitly requests "ASCII art" or "text-based drawing". ' +
            'Describe what you are generating and then output the command on a new line.';

        var languagePolicy = 'Most users are from Russia/CIS. Default response language is Russian. If the user explicitly requests another language, switch to that language.';

        if (editorType === 'word') {
            return 'You are an assistant running inside the {r7c}.ChatLLM plugin in the R7 Office word processor. Help the user with document work in the active R7 Office environment. You can analyze and rewrite text. If native host tools are available, prefer them for simple high-level editor operations before inventing custom macros. ' + skillDescription + ' ' + languagePolicy + ' UI language: ' + language + '.';
        }
        if (editorType === 'cell') {
            return 'You are an assistant running inside the {r7c}.ChatLLM plugin in the R7 Office spreadsheet editor. Help the user with spreadsheet work in the active R7 Office environment. Use R7 terminology and methods first. If native host tools are available, prefer them for simple high-level spreadsheet actions before inventing custom macros. For spreadsheet operations, use the macro-runner workflow when a native tool is not sufficient. When the user asks about what is on the sheet, current values, rows, columns, headers, formulas, or table contents, use the predefined spreadsheet read path such as read_active_sheet, read_sheet_range, or list_sheets instead of inventing a read macro. ' + skillDescription + ' ' + languagePolicy + ' UI language: ' + language + '.';
        }
        if (editorType === 'slide') {
            return 'You are a good slide assistant, you have the ability to get the outline of a topic and more. ' + skillDescription + ' ' + languagePolicy + ' UI language: ' + language + '.';
        }
        return 'You are a good assistant. ' + skillDescription + ' ' + languagePolicy + ' UI language: ' + language + '.';
    }

    function startChat(request) {
        var refs = dom();
        chatState().lastRequest = request;
        chatState().busy = true;
        if (refs.regenButton) refs.regenButton.style.display = 'none';
        if (refs.stopButton) refs.stopButton.style.display = 'block';
    }

    function endChat() {
        var refs = dom();
        chatState().busy = false;
        if (refs.regenButton) refs.regenButton.style.display = 'block';
        if (refs.lastLoadElement) refs.lastLoadElement.style.display = 'none';
        if (refs.stopButton) refs.stopButton.style.display = 'none';
        chatState().abortController = null;
    }

    function buildProviderErrorMessage(error) {
        var tools = getErrorTools();
        if (tools && typeof tools.buildProviderErrorMessage === 'function') {
            return tools.buildProviderErrorMessage(error);
        }
        return String(error && error.message ? error.message : error || 'Request failed');
    }

    function insertInto(editorType, url, imgsize) {
        var host = getHostBridge();
        if (!host || typeof host.callEditorCommand !== 'function') return Promise.resolve(null);
        var scopeData = {
            url: url,
            imgsize: imgsize || { width: 256, height: 256 }
        };
        if (editorType === 'word') {
            return host.callEditorCommand(function () {
                var oDocument = Api.GetDocument();
                var oParagraph = Api.CreateParagraph();
                var width = Asc.scope.imgsize.width * (25.4 / 96.0) * 36000;
                var height = Asc.scope.imgsize.height * (25.4 / 96.0) * 36000;
                var oDrawing = Api.CreateImage(Asc.scope.url, width, height);
                oParagraph.AddDrawing(oDrawing);
                oDocument.InsertContent([oParagraph]);
            }, scopeData);
        }
        if (editorType === 'slide') {
            return host.callEditorCommand(function () {
                var oPresentation = Api.GetPresentation();
                var oSlide = oPresentation.GetCurrentSlide();
                var width = Asc.scope.imgsize.width * (25.4 / 96.0) * 36000;
                var height = Asc.scope.imgsize.height * (25.4 / 96.0) * 36000;
                var oDrawing = Api.CreateImage(Asc.scope.url, width, height);
                oSlide.AddObject(oDrawing);
            }, scopeData);
        }
        if (editorType === 'cell') {
            return host.callEditorCommand(function () {
                var oSheet = Api.GetActiveSheet();
                var width = Asc.scope.imgsize.width * (25.4 / 96.0) * 36000;
                var height = Asc.scope.imgsize.height * (25.4 / 96.0) * 36000;
                var oDrawing = Api.CreateImage(Asc.scope.url, width, height);
                oSheet.AddDrawing(oDrawing);
            }, scopeData);
        }
        return Promise.resolve(null);
    }

    function insertIntoDocument(text) {
        var host = getHostBridge();
        if (!host || typeof host.callEditorCommand !== 'function') return Promise.resolve(null);
        var editorType = typeof host.getEditorTypeSafe === 'function' ? host.getEditorTypeSafe() : '';
        var scopeData = { insert: text };
        if (editorType === 'word') {
            return host.callEditorCommand(function () {
                var oDocument = Api.GetDocument();
                var oParagraph = Api.CreateParagraph();
                oParagraph.AddText(Asc.scope.insert);
                oDocument.InsertContent([oParagraph]);
            }, scopeData);
        }
        if (editorType === 'cell') {
            return host.callEditorCommand(function () {
                var oWorksheet = Api.GetActiveSheet();
                var selection = oWorksheet.GetSelection();
                selection.SetValue(Asc.scope.insert);
            }, scopeData);
        }
        if (editorType === 'slide') {
            return host.callEditorCommand(function () {
                var oPresentation = Api.GetPresentation();
                var oSlide = oPresentation.GetCurrentSlide();
                var oShape = Api.CreateShape('rect', 4986000, 2419200, Api.CreateNoFill(), Api.CreateStroke(0, Api.CreateNoFill()));
                oShape.SetPosition(3834000, 3888000);
                var oContent = oShape.GetDocContent();
                oContent.RemoveAllElements();
                var oParagraph = Api.CreateParagraph();
                var oRun = Api.CreateRun();
                var oTextPr = oRun.GetTextPr();
                oTextPr.SetFontSize(50);
                var oFill = Api.CreateSolidFill(Api.CreateRGBColor(0, 0, 0));
                oTextPr.SetTextFill(oFill);
                oParagraph.SetJc('left');
                oRun.AddText(Asc.scope.insert);
                oParagraph.AddElement(oRun);
                oContent.Push(oParagraph);
                oSlide.AddObject(oShape);
            }, scopeData);
        }
        return Promise.resolve(null);
    }

    function copyToClipboard(text) {
        var payload = String(text || '');
        var successMessage = tr('Message copied to clipboard');
        var errorMessage = tr('Error copying message to clipboard');

        function fallbackCopy() {
            try {
                var buffer = document.createElement('textarea');
                buffer.value = payload;
                buffer.setAttribute('readonly', '');
                buffer.style.position = 'fixed';
                buffer.style.opacity = '0';
                document.body.appendChild(buffer);
                buffer.focus();
                buffer.select();
                var copied = document.execCommand && document.execCommand('copy');
                document.body.removeChild(buffer);
                return !!copied;
            } catch (error) {
                return false;
            }
        }

        function onCopySuccess() { alert(successMessage); }
        function onCopyError() { alert(errorMessage); }

        if (navigator.clipboard && window.isSecureContext) {
            navigator.clipboard.writeText(payload).then(onCopySuccess).catch(function () {
                if (fallbackCopy()) onCopySuccess();
                else onCopyError();
            });
            return;
        }
        if (fallbackCopy()) onCopySuccess();
        else onCopyError();
    }

    function pushAssistantMessage(message, runContainer, pin) {
        if (isRuntimeClosing()) return null;
        if (getUi().displayMessage) {
            return getUi().displayMessage(message, 'assistant', pin === true, runContainer);
        }
        return null;
    }

    function normalizeResponseText(value) {
        return (root.services.context && typeof root.services.context.normalizeTextPayload === 'function')
            ? root.services.context.normalizeTextPayload(value)
            : String(value === null || value === undefined ? '' : value);
    }

    function maybeGenerateInlineImage(responseText, runContainer) {
        var match = String(responseText || '').match(/\[GENERATE_IMAGE:\s*([^\]]+)\]/i);
        if (match && match[1]) {
            generateImageFromText(match[1].trim(), runContainer);
        }
    }

    function finalizeAssistantResponse(responseText, runContainer, warningText) {
        if (isRuntimeClosing()) return '';
        var response = normalizeResponseText(responseText || '');
        var warning = normalizeResponseText(warningText || '');
        var combined = warning ? (warning + (response ? '\n\n' + response : '')) : response;
        var safeResponseText = combined.length ? combined : tr('Empty response from provider');
        pushAssistantMessage(safeResponseText, runContainer, false);
        if (getUi().renderThreadHeader) getUi().renderThreadHeader();
        if (getUi().renderThreadList) getUi().renderThreadList();
        maybeGenerateInlineImage(safeResponseText, runContainer);
        return safeResponseText;
    }

    function buildChatConfig() {
        var runtimeSettings = loadRuntimeSettings();
        var providerId = String(runtimeSettings.activeProvider || runtimeSettings.provider || 'openrouter').trim().toLowerCase() || 'openrouter';
        var service = getSettingsService();
        var providerConfig = service && typeof service.getProviderConfig === 'function'
            ? service.getProviderConfig(runtimeSettings, providerId)
            : {};
        return {
            provider: providerId,
            activeProvider: providerId,
            model: providerConfig && providerConfig.model ? providerConfig.model : (runtimeSettings.model || constants().defaultModel),
            reasoningEffort: providerConfig && providerConfig.reasoningEffort ? providerConfig.reasoningEffort : 'medium',
            temperature: 0.5
        };
    }

    function shouldShowModelReasoningInTrace(runtimeSettings) {
        var settings = runtimeSettings && typeof runtimeSettings === 'object' ? runtimeSettings : loadRuntimeSettings();
        return !(settings && settings.trace && settings.trace.showModelReasoning === false);
    }

    function resolveReasoningFromResponse(config, providers, responseResult, rawPayload) {
        if (responseResult && responseResult.reasoning && responseResult.reasoning.available) {
            return responseResult.reasoning;
        }
        if (!rawPayload || !providers || typeof providers.normalizeResponse !== 'function') return null;
        try {
            var normalized = providers.normalizeResponse(config && config.provider ? config.provider : 'openrouter', rawPayload, config || {});
            return normalized && normalized.reasoning ? normalized.reasoning : null;
        } catch (error) {
            return null;
        }
    }

    function addModelReasoningTrace(config, providers, responseResult, rawPayload, durationMs, runtimeSettings) {
        if (!shouldShowModelReasoningInTrace(runtimeSettings)) return;
        var agent = getAgentService();
        if (!agent || typeof agent.addTraceRecord !== 'function') return;
        var reasoning = resolveReasoningFromResponse(config, providers, responseResult, rawPayload);
        var summary = reasoning && reasoning.summary ? String(reasoning.summary).replace(/\s+/g, ' ').trim() : '';
        if (!reasoning || reasoning.available !== true || !summary) return;
        if (summary.length > 320) summary = summary.slice(0, 317) + '...';
        var effort = reasoning && reasoning.effort ? String(reasoning.effort).trim().toLowerCase() : '';
        var tokens = reasoning && reasoning.tokens !== undefined ? Number(reasoning.tokens) : undefined;
        agent.addTraceRecord({
            step_id: 'chat_reasoning_' + Date.now(),
            step_type: 'model_reasoning',
            status: 'info',
            reason: tr('Model thinking (summary): ') + summary,
            reasoning_summary: summary,
            reasoning_effort: effort,
            reasoning_tokens: Number.isFinite(tokens) ? tokens : undefined,
            provider: config && config.provider ? config.provider : 'openrouter',
            duration_ms: Math.max(0, Number(durationMs || 0))
        });
    }

    function getActiveConversationHistory() {
        var threadStore = getThreadStore();
        if (threadStore && typeof threadStore.syncConversationHistory === 'function') {
            return threadStore.syncConversationHistory();
        }
        chatState().conversationHistory = Array.isArray(chatState().conversationHistory) ? chatState().conversationHistory : [];
        return chatState().conversationHistory;
    }

    function getChatCapabilities(config) {
        var providers = getProvidersBridge();
        if (providers && typeof providers.getChatCapabilities === 'function') {
            return providers.getChatCapabilities(Object.assign({}, buildChatConfig(), config || {}));
        }
        return {
            provider: 'openrouter',
            model: loadRuntimeSettings().model || constants().defaultModel,
            supportsVision: false,
            supportsStreaming: false,
            supportsTools: false
        };
    }

    function shouldUseWebToolLoop(options) {
        var settings = loadRuntimeSettings();
        var webTools = getWebToolsBridge();
        var hostTools = getHostToolsClient();
        var toolLoop = getToolLoop();
        var hasWebTools = !!(webTools && typeof webTools.isEnabled === 'function' && webTools.isEnabled(settings));
        var hasHostTools = !!(hostTools && typeof hostTools.isEnabled === 'function' && hostTools.isEnabled(settings));
        return !!(options && options.allowTools === true &&
            String(settings.activeProvider || settings.provider || '').toLowerCase() === 'openrouter' &&
            toolLoop && typeof toolLoop.run === 'function' &&
            (hasWebTools || hasHostTools));
    }

    function chatRequest(request, runContainer, options) {
        var providers = getProvidersBridge();
        if (!providers || typeof providers.chatRequest !== 'function') {
            throw new Error('Provider registry is unavailable');
        }
        startChat(request);
        var runtimeSettings = loadRuntimeSettings();
        var config = buildChatConfig();
        var requestStartedAt = Date.now();

        chatState().abortController = new AbortController();

        if (shouldUseWebToolLoop(options)) {
            var toolLoop = getToolLoop();
            return toolLoop.run({
                messages: request,
                systemMessage: getSystemMessage(),
                config: config,
                signal: chatState().abortController.signal,
                runtimeSettings: runtimeSettings
            }).then(function (result) {
                addModelReasoningTrace(config, providers, null, result && result.raw ? result.raw : null, Date.now() - requestStartedAt, runtimeSettings);
                var safeResponseText = finalizeAssistantResponse(result && result.response, runContainer, result && result.warning);
                endChat();
                return { response: safeResponseText };
            }).catch(function (error) {
                endChat();
                if (isAbortLikeError(error) || isRuntimeClosing()) {
                    return { response: '' };
                }
                pushAssistantMessage(buildProviderErrorMessage(error), runContainer, true);
                throw error;
            });
        }

        return providers.chatRequest(request, getSystemMessage(), true, config, chatState().abortController.signal)
            .then(function (resultOrReader) {
                var isStreamReader = resultOrReader && typeof resultOrReader.read === 'function';
                if (isStreamReader) {
                    var messageContentElement = getUi().createAIMessageFrame(false, runContainer);
                    return getUi().displaySSEMessage(resultOrReader, messageContentElement, null)
                        .then(function (result) {
                            var responseText = result.response;
                            var threadStore = getThreadStore();
                            if (threadStore && typeof threadStore.appendMessage === 'function') {
                                threadStore.appendMessage('assistant', responseText);
                            } else {
                                chatState().conversationHistory.push({ role: 'assistant', content: responseText });
                            }
                            maybeGenerateInlineImage(responseText, runContainer);

                            endChat();
                            if (getUi().renderThreadList) getUi().renderThreadList();
                            return result;
                        });
                }

                addModelReasoningTrace(config, providers, resultOrReader, null, Date.now() - requestStartedAt, runtimeSettings);
                var safeResponseText = finalizeAssistantResponse(
                    resultOrReader && resultOrReader.data && resultOrReader.data[0] ? resultOrReader.data[0].content : '',
                    runContainer
                );
                endChat();
                return { response: safeResponseText };
            })
            .catch(function (error) {
                endChat();
                if (isAbortLikeError(error) || isRuntimeClosing()) {
                    return { response: '' };
                }
                pushAssistantMessage(buildProviderErrorMessage(error), runContainer, true);
                throw error;
            });
    }

    function stopChatRequest() {
        var agent = getAgentService();
        if (agent && typeof agent.isBusy === 'function' && agent.isBusy() && typeof agent.stop === 'function') {
            agent.stop();
        }
        if (chatState().abortController) {
            chatState().abortController.abort();
        }
        endChat();
    }

    function clearHistory() {
        var refs = dom();
        var threadStore = getThreadStore();
        if (threadStore && typeof threadStore.clearThread === 'function') {
            var activeThread = threadStore.getActiveThread ? threadStore.getActiveThread() : null;
            threadStore.clearThread(activeThread ? activeThread.id : '');
        } else {
            chatState().conversationHistory = [];
        }
        if (refs.messageInput) {
            refs.messageInput.value = '';
        }
        var agent = getAgentService();
        if (agent && typeof agent.reset === 'function') {
            agent.reset();
        }
        if (getUi().renderChatThreads) {
            getUi().renderChatThreads();
        } else if (getUi().clearMessageHistory) {
            getUi().clearMessageHistory();
        }
        if (getUi().renderContextPanel) getUi().renderContextPanel();
        if (getUi().scheduleRenderTrace) getUi().scheduleRenderTrace();
    }

    function reSend() {
        var refs = dom();
        var agent = getAgentService();
        var context = getContextService();
        var history = getActiveConversationHistory();
        checkApiKey(true);

        if (refs.messageHistory && refs.messageHistory.lastChild) {
            refs.messageHistory.removeChild(refs.messageHistory.lastChild);
        }

        if (!chatState().hasKey) {
            pushAssistantMessage(getMissingAuthMessage(), null, true);
            return;
        }

        if (agent && typeof agent.shouldUseMacroRunner === 'function' && agent.shouldUseMacroRunner()) {
            if (typeof agent.runFromStepOne === 'function') {
                agent.runFromStepOne();
            }
            return;
        }

        if (context && typeof context.prepareRequestWithContext === 'function') {
            context.prepareRequestWithContext(history, '')
                .then(function (request) { return chatRequest(request, null, { allowTools: true }); })
                .catch(function () { return chatRequest(context.cloneMessages(history), null, { allowTools: true }); });
            return;
        }

        chatRequest(history, null, { allowTools: true });
    }

    function sendMessage(message, options) {
        var settings = options && typeof options === 'object' ? options : {};
        var text = String(message || '').trim();
        var attachments = Array.isArray(settings.attachments) ? settings.attachments.slice() : [];
        var attachmentService = getAttachmentService();
        var imageAttachments = attachmentService && typeof attachmentService.hasImageAttachments === 'function' && attachmentService.hasImageAttachments(attachments)
            ? attachments.filter(function (item) {
                return attachmentService.isImageAttachment && attachmentService.isImageAttachment(item);
            })
            : [];
        var context = getContextService();
        var agent = getAgentService();
        var ui = getUi();
        var threadStore = getThreadStore();
        var capabilities = imageAttachments.length ? getChatCapabilities() : null;

        chatState().lastRequest = null;
        if (!text.length && !attachments.length) return Promise.resolve(null);

        if (agent && typeof agent.isAwaitingPlanApproval === 'function' && agent.isAwaitingPlanApproval()) {
            if (attachments.length) {
                if (ui && typeof ui.showThreadToast === 'function') {
                    ui.showThreadToast(tr('Finish plan approval first. Use the next text message to edit the plan or approve it from the plan card.'), { kind: 'error', durationMs: 3200 });
                }
                return Promise.resolve({ sent: false, reason: 'plan_approval_pending' });
            }
            if (!text.length) {
                return Promise.resolve({ sent: false, reason: 'empty_plan_revision' });
            }
            var pendingRunContainer = agent.getState && agent.getState().currentRunContainer
                ? agent.getState().currentRunContainer
                : (ui.createRunContainer ? ui.createRunContainer() : null);
            var isApprovalMessage = typeof agent.isPlanApprovalMessage === 'function' && agent.isPlanApprovalMessage(text);
            ui.displayMessage(text, 'user', false, pendingRunContainer);
            if (ui.clearComposerAttachments) ui.clearComposerAttachments();
            if (ui.clearPlanRevisionPrompt) ui.clearPlanRevisionPrompt();
            if (threadStore && typeof threadStore.replaceDraft === 'function') {
                var pendingThread = threadStore.getActiveThread ? threadStore.getActiveThread() : null;
                if (pendingThread) {
                    threadStore.replaceDraft(pendingThread.id, '');
                }
            }
            if (isApprovalMessage && typeof agent.approvePendingPlan === 'function') {
                return Promise.resolve(agent.approvePendingPlan()).then(function (result) {
                    return { sent: true, response: result, reason: 'plan_approved' };
                });
            }
            return Promise.resolve(agent.revisePendingPlan(text)).then(function (result) {
                return { sent: true, response: result, reason: 'plan_revised' };
            });
        }

        if (imageAttachments.length && (!capabilities || capabilities.supportsVision !== true)) {
            if (ui && typeof ui.showThreadToast === 'function') {
                ui.showThreadToast(tr('The selected model does not support image input. Choose a vision-capable model and try again.'), { kind: 'error', durationMs: 3200 });
            }
            return Promise.resolve({ sent: false, reason: 'vision_not_supported' });
        }

        if (threadStore && typeof threadStore.replaceDraft === 'function') {
            var activeThread = threadStore.getActiveThread ? threadStore.getActiveThread() : null;
            if (activeThread) {
                threadStore.replaceDraft(activeThread.id, '');
            }
        }
        var runContainer = ui.createRunContainer ? ui.createRunContainer() : null;
        ui.displayMessage(text, 'user', false, runContainer, {
            meta: attachments.length ? { attachments: attachments } : null
        });
        if (ui.clearComposerAttachments) ui.clearComposerAttachments();
        if (ui.renderThreadHeader) ui.renderThreadHeader();
        if (ui.renderThreadList) ui.renderThreadList();
        
        // If message has no text but has attachments, provide a default intent for the agent/LLM
        var effectiveText = text;
        if (!effectiveText.trim() && attachments.length) {
            effectiveText = 'Проанализируй прикрепленные файлы';
        }

        if (!checkApiKey(true)) {
            ui.displayMessage(getMissingAuthMessage(), 'assistant', true, runContainer);
            return Promise.resolve({ sent: true, reason: 'missing_api_key' });
        }

        if (agent && typeof agent.shouldUseMacroRunner === 'function' && agent.shouldUseMacroRunner() && !imageAttachments.length) {
            return agent.run(effectiveText, runContainer).then(function (result) {
                return { sent: true, response: result };
            });
        }

        if (effectiveText.toLowerCase().indexOf('/draw ') === 0) {
            return generateImageFromText(effectiveText.substring(6).trim(), runContainer).then(function (result) {
                return { sent: true, response: result };
            });
        }

        if (context && typeof context.prepareRequestWithContext === 'function') {
            return context.prepareRequestWithContext(getActiveConversationHistory(), effectiveText)
                .then(function (request) { return chatRequest(request, runContainer, { allowTools: !imageAttachments.length }); })
                .catch(function () { return chatRequest(context.cloneMessages(getActiveConversationHistory()), runContainer, { allowTools: !imageAttachments.length }); })
                .then(function (result) {
                    return { sent: true, response: result };
                });
        }

        if (!effectiveText.trim() && attachments.length && imageAttachments.length) {
            return chatRequest(getActiveConversationHistory(), runContainer, { allowTools: false }).then(function (result) {
                return { sent: true, response: result };
            });
        }

        return chatRequest(getActiveConversationHistory(), runContainer, { allowTools: !imageAttachments.length }).then(function (result) {
            return { sent: true, response: result };
        });
    }

    function isBusy() {
        return !!chatState().busy;
    }

    function translateHelper(text, targetLanguage) {
        var prompt = translatePrompt(text, targetLanguage);
        var message = 'translate the selected text to ' + targetLanguage;
        if (!text && getHostBridge() && getHostBridge().getEditorTypeSafe() === 'cell') {
            pushAssistantMessage(tr('Sorry, please select text in a cell to proceed translate selected text.'), null, true);
            return;
        }
        pushAssistantMessage(tr(message), null, false);
        return chatRequest(decoratePrompt(prompt));
    }

    function generateImageFromText(text, runContainer) {
        var imageBridge = getImageBridge();
        var host = getHostBridge();
        if (!imageBridge || typeof imageBridge.generate !== 'function') {
            pushAssistantMessage(tr('Image generation bridge is unavailable'), runContainer, true);
            return Promise.resolve(null);
        }
        
        var messageFrame = pushAssistantMessage(tr('Generating image...'), runContainer, false);
        if (!messageFrame) return Promise.resolve(null);

        return imageBridge.generate(text).then(function (response) {
            try {
            console.log('[DEBUG] generateImageFromText raw response:', JSON.stringify(response));
            var url = (response && response.data && response.data[0] ? response.data[0].url : '') || '';
            console.log('[DEBUG] Image generation result url:', url.length > 100 ? url.slice(0, 100) + '...' : url);
            
            if (url && typeof url === 'string' && (url.indexOf('http') === 0 || url.indexOf('data:image') === 0)) {
                messageFrame.innerHTML = '';
                
                var topInfo = document.createElement('div');
                topInfo.style.marginBottom = '8px';
                topInfo.style.fontSize = '12px';
                topInfo.style.color = 'var(--accent-gold)';
                topInfo.textContent = tr('Image generated successfully');
                messageFrame.appendChild(topInfo);

                var img = document.createElement('img');
                img.src = url;
                img.style.maxWidth = '100%';
                img.style.display = 'block';
                img.style.borderRadius = 'var(--border-radius-md)';
                img.style.boxShadow = '0 4px 15px rgba(0,0,0,0.3)';
                img.style.cursor = 'zoom-in';
                img.style.marginBottom = '12px';
                
                img.onclick = function(e) { if (e && typeof e.preventDefault === 'function') { e.preventDefault(); e.stopPropagation(); } if (url && (url.indexOf('http') === 0 || url.indexOf('data:image') === 0)) { window.open(url, '_blank'); } };
                messageFrame.appendChild(img);

                var actionsRow = document.createElement('div');
                actionsRow.className = 'actions';
                actionsRow.style.display = 'flex';
                actionsRow.style.flexWrap = 'wrap';
                actionsRow.style.gap = '8px';
                
                // 1. Insert Button
                var insertBtn = document.createElement('button');
                insertBtn.type = 'button';
                insertBtn.className = 'sheet-action-btn';
                insertBtn.innerHTML = '&#x2398; ' + tr('Insert');
                insertBtn.onclick = function(e) { if (e && typeof e.preventDefault === 'function') { e.preventDefault(); e.stopPropagation(); } if (!url || (url.indexOf('http') !== 0 && url.indexOf('data:image') !== 0)) return;
                    insertInto(
                        host && typeof host.getEditorTypeSafe === 'function' ? host.getEditorTypeSafe() : '',
                        url,
                        { width: 420, height: 420 }
                    );
                };
                
                // 2. Copy Link Button
                var copyBtn = document.createElement('button');
                copyBtn.type = 'button';
                copyBtn.className = 'sheet-action-btn';
                copyBtn.textContent = tr('Copy URL');
                copyBtn.onclick = function(e) { if (e && typeof e.preventDefault === 'function') { e.preventDefault(); e.stopPropagation(); } if (!url) return;
                    copyToClipboard(url);
                };
                
                // 3. Download Button
                var saveBtn = document.createElement('button');
                saveBtn.type = 'button';
                saveBtn.className = 'sheet-action-btn';
                saveBtn.textContent = tr('Save');
                saveBtn.onclick = function(e) { if (e && typeof e.preventDefault === 'function') { e.preventDefault(); e.stopPropagation(); } if (!url || (url.indexOf('http') !== 0 && url.indexOf('data:image') !== 0)) return;
                    var link = document.createElement('a');
                    link.href = url;
                    link.download = 'generated_image_' + Date.now() + '.png';
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                };
                
                actionsRow.appendChild(insertBtn);
                actionsRow.appendChild(copyBtn);
                actionsRow.appendChild(saveBtn);
                messageFrame.appendChild(actionsRow);
                
                // Auto-insertion (as requested: "after selected text")
                insertBtn.onclick();
                
                if (getUi().scheduleRenderTrace) getUi().scheduleRenderTrace();
                endChat();
                return response;
            }
            throw new Error(tr('Model returned no image. Ensure an image generation model is selected in settings.'));
            } catch (innerErr) {
                console.error('[ERROR] generateImageFromText inner catch:', innerErr);
                endChat();
                if (messageFrame) {
                    messageFrame.innerHTML = '';
                    var errDiv = document.createElement('div');
                    errDiv.style.color = 'var(--accent-scar)';
                    errDiv.textContent = tr('Error: ') + (innerErr && innerErr.message ? innerErr.message : tr('Image generation failed'));
                    messageFrame.appendChild(errDiv);
                }
                throw innerErr;
            }
        }).catch(function (error) {
            console.error('[ERROR] generateImageFromText promise catch:', error);
            endChat();
            if (messageFrame) {
                messageFrame.innerHTML = '';
                var catchErrDiv = document.createElement('div');
                catchErrDiv.style.color = 'var(--accent-scar)';
                // ВАЖНО: используем textContent а не innerHTML, чтобы HTML из ответа API не рендерился
                var rawMsg = error && error.message ? error.message : '';
                // Если сообщение — это HTML-страница (404 от OpenRouter), показываем краткое описание
                if (rawMsg.indexOf('<!DOCTYPE') !== -1 || rawMsg.indexOf('<html') !== -1) {
                    var statusHint = error && error.status ? ' (HTTP ' + error.status + ')' : '';
                    catchErrDiv.textContent = tr('Error: ') + tr('Image generation API returned an error') + statusHint + '. ' + tr('Check that image model is set correctly in Settings.');
                } else {
                    catchErrDiv.textContent = tr('Error: ') + (rawMsg || tr('Image generation failed'));
                }
                messageFrame.appendChild(catchErrDiv);
            }
        });
    }

    root.features.chatRuntime = {
        getState: function () { return chatState(); },
        loadRuntimeSettings: loadRuntimeSettings,
        saveRuntimeSettings: saveRuntimeSettings,
        checkApiKey: checkApiKey,
        decoratePrompt: decoratePrompt,
        formatTable: formatTable,
        translateHelper: translateHelper,
        generateImageFromText: generateImageFromText,
        getSystemMessage: getSystemMessage,
        copyToClipboard: copyToClipboard,
        insertIntoDocument: insertIntoDocument,
        clearHistory: clearHistory,
        reSend: reSend,
        sendMessage: sendMessage,
        chatRequest: chatRequest,
        stopChatRequest: stopChatRequest,
        isBusy: isBusy,
        getChatCapabilities: getChatCapabilities
    };

    root.services.chat = root.features.chatRuntime;
})(window);
