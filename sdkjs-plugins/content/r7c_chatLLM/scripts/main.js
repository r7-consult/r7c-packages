// {r7c}.ChatLLM runtime coordinator
(function (window, undefined) {
    'use strict';

    function root() {
        return window.R7Chat = window.R7Chat || {};
    }

    function runtime() {
        return root().runtime || {};
    }

    function dom() {
        return runtime().dom || {};
    }

    function ui() {
        return root().ui || {};
    }

    function lifecycle() {
        root().runtime = root().runtime || {};
        root().runtime.lifecycle = root().runtime.lifecycle || {
            isClosing: false,
            closeScheduled: false
        };
        return root().runtime.lifecycle;
    }

    function chat() {
        return root().features && root().features.chatRuntime ? root().features.chatRuntime : null;
    }

    function context() {
        return root().features && root().features.contextRuntime ? root().features.contextRuntime : null;
    }

    function agent() {
        return root().features && root().features.agentRuntime ? root().features.agentRuntime : null;
    }

    function settingsService() {
        return root().features && root().features.settings ? root().features.settings : null;
    }

    function settingsPanel() {
        return root().features && root().features.settingsPanel ? root().features.settingsPanel : null;
    }

    function threadStore() {
        return root().services && root().services.chatThreads ? root().services.chatThreads : null;
    }

    function hostBridge() {
        return root().platform && root().platform.hostBridge ? root().platform.hostBridge : null;
    }

    function tr(text) {
        var host = hostBridge();
        if (host && typeof host.translate === 'function') {
            return host.translate(text);
        }
        if (window.Asc && window.Asc.plugin && typeof window.Asc.plugin.tr === 'function') {
            return window.Asc.plugin.tr(text);
        }
        return String(text || '');
    }

    function generateText(text) {
        return tr(text);
    }

    function currentLanguage() {
        var lang = runtime().lang || 'en';
        var langMap = runtime().langMap || {};
        return langMap[lang] || 'English';
    }

    function getHostLanguageCode() {
        var info = window.Asc && window.Asc.plugin && window.Asc.plugin.info ? window.Asc.plugin.info : {};
        var language = String(info.lang || runtime().lang || 'en').substring(0, 2).toLowerCase();
        return language || 'en';
    }

    function getLanguagePreference() {
        var settings = settingsService();
        if (!settings || typeof settings.loadSettings !== 'function') return 'auto';
        try {
            var loaded = settings.loadSettings();
            var preference = String(loaded && loaded.language || 'auto').trim().toLowerCase();
            return preference === 'en' || preference === 'ru' ? preference : 'auto';
        } catch (error) {
            return 'auto';
        }
    }

    function resolveUiLanguage(preference) {
        var normalized = String(preference || '').trim().toLowerCase();
        return normalized === 'en' || normalized === 'ru' ? normalized : getHostLanguageCode();
    }

    function applyLanguagePreference(preference, options) {
        var settings = options && typeof options === 'object' ? options : {};
        runtime().lang = resolveUiLanguage(preference);
        document.documentElement.setAttribute('lang', runtime().lang || 'en');
        if (!settings.skipRetranslate && window.Asc && window.Asc.plugin && typeof window.Asc.plugin.onTranslate === 'function') {
            window.Asc.plugin.onTranslate();
        }
        return runtime().lang;
    }

    function applyConfiguredLanguage(options) {
        return applyLanguagePreference(getLanguagePreference(), options);
    }

    function shutdownActiveWork() {
        var currentChat = chat();
        var currentAgent = agent();
        if (currentChat && typeof currentChat.stopChatRequest === 'function') {
            try {
                currentChat.stopChatRequest();
            } catch (error) {}
        }
        if (currentAgent && typeof currentAgent.stop === 'function') {
            try {
                currentAgent.stop();
            } catch (error2) {}
        }
    }

    function closePluginWindow() {
        if (!(window.Asc && window.Asc.plugin && typeof window.Asc.plugin.executeCommand === 'function')) return;
        window.Asc.plugin.executeCommand('close', '');
    }

    function requestPluginClose() {
        var state = lifecycle();
        if (state.closeScheduled) return;
        state.isClosing = true;
        state.closeScheduled = true;
        shutdownActiveWork();
        window.setTimeout(function () {
            closePluginWindow();
        }, 60);
    }

    function wireServiceContracts() {
        var settings = settingsService();
        var panel = settingsPanel();
        var currentUi = ui();
        var currentContext = context();
        var currentAgent = agent();
        var currentChat = chat();

        root().services = root().services || {};
        root().services.settings = {
            load: function () { return settings ? settings.loadSettings() : null; },
            save: function (nextSettings) { return settings ? settings.saveSettings(null, nextSettings || {}) : null; },
            validateApiKey: function (showRequiredError) {
                return panel ? panel.validateInlineToken(showRequiredError) : false;
            },
            open: function () { return panel && panel.openInlineSettingsPanel(); },
            close: function () { return panel && panel.closeInlineSettingsPanel(); },
            toggle: function () { return panel && panel.toggleInlineSettingsPanel(); }
        };
        root().services.chat = currentChat || {};
        root().services.context = currentContext || {};
        root().services.agent = currentAgent || {};
        root().services.i18n = {
            getPreference: getLanguagePreference,
            apply: function (language) { return applyLanguagePreference(language); },
            refresh: function () { return applyConfiguredLanguage(); }
        };
        root().services.uiShell = {
            openSettings: function () { return panel && panel.openInlineSettingsPanel(); },
            closeSettings: function () { return panel && panel.closeInlineSettingsPanel(); },
            toggleSettings: function () { return panel && panel.toggleInlineSettingsPanel(); },
            renderContextPanel: function () { return currentUi.renderContextPanel && currentUi.renderContextPanel(); },
            renderTrace: function () { return currentUi.renderTrace && currentUi.renderTrace(); }
        };
    }

    function updateGreeting() {
        if (dom().greeting) {
            dom().greeting.innerHTML = tr('Hi there! I am {r7c}.ChatLLM, how can I help you today?');
        }
    }

    function normalizeTheme(theme) {
        var t = String(theme || '').toLowerCase();
        return t.indexOf('light') !== -1 ? 'light' : 'dark';
    }

    function updateTheme(theme, isManual) {
        var normalized = normalizeTheme(theme);
        console.log('updateTheme: input=' + theme + ' normalized=' + normalized + ' isManual=' + !!isManual);
        
        document.documentElement.setAttribute('data-theme', normalized);
        if (ui().updateThemeIcon) {
            ui().updateThemeIcon(normalized);
        }
        
        if (isManual && root().platform && root().platform.storage && typeof root().platform.storage.writeJson === 'function') {
            root().platform.storage.writeJson('r7c_chat_theme', normalized);
            root().platform.storage.writeJson('r7c_chat_theme_manual', true);
        }
    }

    function extractClipboardImageFiles(clipboardData) {
        var files = [];
        if (!clipboardData) return files;

        var items = clipboardData.items || [];
        for (var i = 0; i < items.length; i += 1) {
            var item = items[i];
            if (!item || item.kind !== 'file' || !/^image\//i.test(String(item.type || ''))) continue;
            var file = typeof item.getAsFile === 'function' ? item.getAsFile() : null;
            if (file) files.push(file);
        }

        if (!files.length && clipboardData.files && clipboardData.files.length) {
            for (var j = 0; j < clipboardData.files.length; j += 1) {
                var clipboardFile = clipboardData.files[j];
                if (clipboardFile && /^image\//i.test(String(clipboardFile.type || ''))) {
                    files.push(clipboardFile);
                }
            }
        }

        return files;
    }

    function insertTextAtCursor(input, text) {
        if (!input) return;
        var value = String(input.value || '');
        var insertion = String(text || '');
        var start = typeof input.selectionStart === 'number' ? input.selectionStart : value.length;
        var end = typeof input.selectionEnd === 'number' ? input.selectionEnd : value.length;
        input.value = value.slice(0, start) + insertion + value.slice(end);
        var nextCursor = start + insertion.length;
        if (typeof input.setSelectionRange === 'function') {
            input.setSelectionRange(nextCursor, nextCursor);
        }
        input.dispatchEvent(new Event('input', { bubbles: true }));
    }

    function submitComposerMessage() {
        if (!chat()) return Promise.resolve({ sent: false, reason: 'chat_unavailable' });
        var input = dom().messageInput || null;
        var originalText = input ? String(input.value || '') : '';
        var attachments = ui().getActiveComposerAttachments ? ui().getActiveComposerAttachments() : [];
        var optimisticClear = !!(originalText.trim().length || (attachments && attachments.length));

        function restoreInput() {
            if (!input) return;
            input.value = originalText;
            input.dispatchEvent(new Event('input', { bubbles: true }));
        }

        if (optimisticClear && input) {
            input.value = '';
            input.dispatchEvent(new Event('input', { bubbles: true }));
        }

        return Promise.resolve(chat().sendMessage(originalText, {
            attachments: attachments
        })).then(function (result) {
            if (result && result.sent === false) {
                restoreInput();
            }
            return result;
        }).catch(function (error) {
            restoreInput();
            throw error;
        });
    }

    function bindThemeToggle() {
        if (!dom().themeToggle) return;
        dom().themeToggle.onclick = function () {
            var currentTheme = document.documentElement.getAttribute('data-theme') || 'dark';
            var nextTheme = currentTheme === 'dark' ? 'light' : 'dark';
            updateTheme(nextTheme, true);
        };
    }

    function getContextMenuItems(options) {
        var editorType = hostBridge() ? hostBridge().getEditorTypeSafe() : '';
        var hasKey = chat() && chat().getState ? !!chat().getState().hasKey : false;
        var menu = {
            guid: window.Asc.plugin.guid,
            items: [{
                id: 'r7c_chatllm',
                text: generateText('{r7c}.ChatLLM'),
                items: []
            }]
        };

        if (!hasKey) return menu;

        switch (options.type) {
            case 'Target':
                if (editorType === 'word') {
                    menu.items[0].items.push({
                        id: 'clear_history',
                        text: generateText('Clear chat history')
                    });
                }
                break;
            case 'Selection':
                if (editorType === 'word') {
                    menu.items[0].items.push(
                        { id: 'onAnalyse', text: generateText('Analyse text') },
                        { id: 'onComplete', text: generateText('Complete text') },
                        { id: 'onGenerateDraft', text: generateText('Generate draft') },
                        { id: 'onRewrite', text: generateText('Rewrite text') },
                        { id: 'onShorten', text: generateText('Shorten text') },
                        { id: 'onExpand', text: generateText('Expand text') },
                        { id: 'onTextToImage', text: generateText('Text to image') }
                    );
                } else if (editorType === 'cell') {
                    menu.items[0].items.push(
                        { id: 'onExplainC', text: generateText('Explain cells') },
                        { id: 'onSummariseC', text: generateText('Summarise cells') }
                    );
                } else if (editorType === 'slide') {
                    menu.items[0].items.push(
                        { id: 'onOutline', text: generateText('Get outline of a topic') },
                        { id: 'onTextToImage', text: generateText('Text to image') }
                    );
                }

                menu.items[0].items.push(
                    { id: 'onSummarize', text: generateText('Summarize text') },
                    { id: 'onExplain', text: generateText('Explain text') },
                    { id: 'onCorrect', text: generateText('Correct spelling and grammar') },
                    {
                        id: 'translate',
                        text: generateText('translate'),
                        items: [
                            { id: 'translate_to_en', text: generateText('translate to English') },
                            { id: 'translate_to_zh', text: generateText('translate to Chinese') },
                            { id: 'translate_to_fr', text: generateText('translate to French') },
                            { id: 'translate_to_de', text: generateText('translate to German') },
                            { id: 'translate_to_es', text: generateText('translate to Spanish') }
                        ]
                    },
                    { id: 'clear_history', text: generateText('clear chat history') }
                );
                break;
        }

        return menu;
    }

    function withSelectedText(callback) {
        window.Asc.plugin.executeMethod('GetSelectedText', null, function (text) {
            callback(text);
        });
    }

    function showPromptAction(label, prompt) {
        if (!chat() || !ui().displayMessage) return;
        ui().displayMessage(label, 'user', false);
        chat().chatRequest(chat().decoratePrompt(prompt));
    }

    async function getSelectedCellValues() {
        var host = hostBridge();
        if (!host) return [];
        return host.callEditorCommand(function () {
            var sheet = Api.GetActiveSheet();
            var selection = sheet.GetSelection();
            return selection.GetValue();
        }, null, { recalculate: false });
    }

    function bindContextMenuHandlers() {
        window.Asc.plugin.attachEvent('onContextMenuShow', function (options) {
            if (!options) return;
            if (options.type === 'Selection' || options.type === 'Target') {
                this.executeMethod('AddContextMenuItem', [getContextMenuItems(options)]);
            }
        });

        window.Asc.plugin.attachContextMenuClickEvent('clear_history', function () {
            if (chat()) chat().clearHistory();
        });

        window.Asc.plugin.attachContextMenuClickEvent('onExplainC', function () {
            getSelectedCellValues().then(function (data) {
                var prompt = explainCellPrompt(chat().formatTable(data), currentLanguage());
                ui().displayMessage(tr('Explain the selected cells'), 'user', false);
                chat().chatRequest(chat().decoratePrompt(prompt));
            });
        });

        window.Asc.plugin.attachContextMenuClickEvent('onSummariseC', function () {
            getSelectedCellValues().then(function (data) {
                var prompt = summariseCellPrompt(chat().formatTable(data), currentLanguage());
                ui().displayMessage(tr('Summarise the selected cells'), 'user', false);
                chat().chatRequest(chat().decoratePrompt(prompt));
            });
        });

        window.Asc.plugin.attachContextMenuClickEvent('onOutline', function () {
            withSelectedText(function (text) {
                showPromptAction(tr('Get outline of a topic'), slideOutlinePrompt(text, 5, currentLanguage()));
            });
        });

        window.Asc.plugin.attachContextMenuClickEvent('onCorrect', function () {
            withSelectedText(function (text) {
                if (!text && hostBridge() && hostBridge().getEditorTypeSafe() === 'cell') {
                    ui().displayMessage(tr('Sorry, please select text in a cell to proceed correct selected spelling and grammar.'), 'assistant', true);
                    return;
                }
                showPromptAction(tr('Correct spelling and grammar'), correctContentPrompt(text, currentLanguage()));
            });
        });

        window.Asc.plugin.attachContextMenuClickEvent('onGenerateDraft', function () {
            withSelectedText(function (text) {
                showPromptAction(tr('Generate draft'), generateDraftPrompt(text, currentLanguage()));
            });
        });

        window.Asc.plugin.attachContextMenuClickEvent('onSummarize', function () {
            withSelectedText(function (text) {
                if (!text && hostBridge() && hostBridge().getEditorTypeSafe() === 'cell') {
                    ui().displayMessage(tr('Sorry, please select text in a cell to proceed summarize the selected text.'), 'assistant', true);
                    return;
                }
                showPromptAction(tr('Summarize the selected text'), summarizeContentPrompt(text, currentLanguage()));
            });
        });

        window.Asc.plugin.attachContextMenuClickEvent('onExplain', function () {
            withSelectedText(function (text) {
                if (!text && hostBridge() && hostBridge().getEditorTypeSafe() === 'cell') {
                    ui().displayMessage(tr('Sorry, please select text in a cell to proceed explain the selected text.'), 'assistant', true);
                    return;
                }
                showPromptAction(tr('Explain the selected text'), explainContentPrompt(text, currentLanguage()));
            });
        });

        window.Asc.plugin.attachContextMenuClickEvent('onAnalyse', function () {
            withSelectedText(function (text) {
                showPromptAction(tr('Analyse the selected text'), analyzeTextPrompt(text, currentLanguage()));
            });
        });

        window.Asc.plugin.attachContextMenuClickEvent('onComplete', function () {
            withSelectedText(function (text) {
                showPromptAction(tr('Complete the selected text'), completeTextPrompt(text, currentLanguage()));
            });
        });

        window.Asc.plugin.attachContextMenuClickEvent('onRewrite', function () {
            withSelectedText(function (text) {
                showPromptAction(tr('Rewrite the selected text'), rewriteContentPrompt(text, currentLanguage()));
            });
        });

        window.Asc.plugin.attachContextMenuClickEvent('onShorten', function () {
            withSelectedText(function (text) {
                showPromptAction(tr('Shorten the selected text'), shortenContentPrompt(text, currentLanguage()));
            });
        });

        window.Asc.plugin.attachContextMenuClickEvent('onExpand', function () {
            withSelectedText(function (text) {
                showPromptAction(tr('Expand the selected text'), expandContentPrompt(text, currentLanguage()));
            });
        });

        window.Asc.plugin.attachContextMenuClickEvent('onTextToImage', function () {
            withSelectedText(function (text) {
                chat().generateImageFromText(textToImagePrompt(text));
            });
        });

        window.Asc.plugin.attachContextMenuClickEvent('translate_to_zh', function () { withSelectedText(function (text) { chat().translateHelper(text, 'Chinese'); }); });
        window.Asc.plugin.attachContextMenuClickEvent('translate_to_en', function () { withSelectedText(function (text) { chat().translateHelper(text, 'English'); }); });
        window.Asc.plugin.attachContextMenuClickEvent('translate_to_fr', function () { withSelectedText(function (text) { chat().translateHelper(text, 'French'); }); });
        window.Asc.plugin.attachContextMenuClickEvent('translate_to_de', function () { withSelectedText(function (text) { chat().translateHelper(text, 'German'); }); });
        window.Asc.plugin.attachContextMenuClickEvent('translate_to_es', function () { withSelectedText(function (text) { chat().translateHelper(text, 'Spanish'); }); });
    }

    function bindDomEvents() {
        if (dom().firstInsertButton) {
            dom().firstInsertButton.addEventListener('click', function () {
                chat().insertIntoDocument(dom().greeting ? dom().greeting.innerText : '');
            });
        }

        if (dom().threadSwitcherButton) {
            dom().threadSwitcherButton.addEventListener('click', function () {
                ui().toggleHistoryPanel();
            });
        }
        if (dom().threadHistoryCloseButton) {
            dom().threadHistoryCloseButton.addEventListener('click', function () {
                ui().closeHistoryPanel();
            });
        }
        if (dom().newChatButton) {
            dom().newChatButton.addEventListener('click', function () {
                if (chat() && typeof chat().isBusy === 'function' && chat().isBusy()) {
                    ui().showThreadToast(tr('Wait for the current response to finish first.'), { kind: 'error' });
                    return;
                }
                if (threadStore() && typeof threadStore().createThread === 'function') {
                    threadStore().createThread({ preferReuseEmpty: true });
                }
                if (ui().closeHistoryPanel) ui().closeHistoryPanel();
                if (ui().renderChatThreads) ui().renderChatThreads();
                if (dom().messageInput) dom().messageInput.focus();
            });
        }
        if (dom().threadMenuButton) {
            dom().threadMenuButton.addEventListener('click', function (event) {
                event.stopPropagation();
                if (threadStore() && typeof threadStore().getActiveThread === 'function') {
                    var activeThread = threadStore().getActiveThread();
                    if (activeThread) {
                        ui().toggleThreadActionMenu(dom().threadMenuButton, activeThread.id);
                    }
                }
            });
        }
        if (dom().threadSearchInput) {
            dom().threadSearchInput.addEventListener('input', function () {
                if (threadStore() && typeof threadStore().setUiState === 'function') {
                    threadStore().setUiState({ searchQuery: dom().threadSearchInput.value || '' });
                }
                if (ui().renderThreadList) ui().renderThreadList();
            });
        }
        if (dom().threadActionMenu) {
            dom().threadActionMenu.addEventListener('click', function (event) {
                var target = event.target && event.target.closest ? event.target.closest('[data-thread-action]') : null;
                var targetThreadId = ui().getMenuThreadId ? ui().getMenuThreadId() : '';
                if (!target || !targetThreadId) return;
                var action = target.getAttribute('data-thread-action');
                ui().closeThreadActionMenu();
                if (action === 'rename') {
                    ui().openRenameThreadDialog(targetThreadId);
                    return;
                }
                if (action === 'export') {
                    ui().exportThreadById(targetThreadId);
                    return;
                }
                if (action === 'delete') {
                    ui().openDeleteThreadDialog(targetThreadId);
                }
            });
        }
        if (dom().threadDialogClose) dom().threadDialogClose.addEventListener('click', function () { ui().closeThreadDialog(); });
        if (dom().threadDialogCancel) dom().threadDialogCancel.addEventListener('click', function () { ui().closeThreadDialog(); });
        if (dom().threadDialogConfirm) dom().threadDialogConfirm.addEventListener('click', function () { ui().confirmThreadDialog(); });
        if (dom().threadDialogInput) {
            dom().threadDialogInput.addEventListener('keydown', function (event) {
                if (event.key === 'Enter') {
                    event.preventDefault();
                    ui().confirmThreadDialog();
                }
            });
        }
        if (dom().attachmentTriggerButton) {
            dom().attachmentTriggerButton.addEventListener('click', function (event) {
                event.stopPropagation();
                ui().toggleAttachmentPopover();
            });
        }
        if (dom().attachmentAddFileAction) {
            dom().attachmentAddFileAction.addEventListener('click', function () {
                ui().promptLocalAttachmentSelection();
            });
        }
        if (dom().attachmentRecentToggleAction) {
            dom().attachmentRecentToggleAction.addEventListener('click', function () {
                ui().toggleAttachmentRecentPanel();
            });
        }
        if (dom().attachmentFileInput) {
            dom().attachmentFileInput.addEventListener('change', function () {
                ui().consumeSelectedFiles(dom().attachmentFileInput.files, { source: 'upload' });
                dom().attachmentFileInput.value = '';
            });
        }
        if (dom().attachmentRecentList) {
            dom().attachmentRecentList.addEventListener('click', function (event) {
                var target = event.target && event.target.closest ? event.target.closest('[data-attachment-fingerprint]') : null;
                if (!target) return;
                ui().attachRecentFile(target.getAttribute('data-attachment-fingerprint'));
            });
        }
        if (dom().composerAttachmentStrip) {
            dom().composerAttachmentStrip.addEventListener('click', function (event) {
                var target = event.target && event.target.closest ? event.target.closest('[data-attachment-id]') : null;
                if (!target) return;
                ui().removeComposerAttachment(target.getAttribute('data-attachment-id'));
            });
        }

        if (dom().sendButton) {
            dom().sendButton.addEventListener('click', function (event) {
                if (event) event.preventDefault();
                submitComposerMessage();
            });
        }
        if (dom().stopButton) dom().stopButton.addEventListener('click', function (event) { if (event) event.preventDefault(); chat().stopChatRequest(); });
        if (dom().regenButton) dom().regenButton.addEventListener('click', function (event) { if (event) event.preventDefault(); chat().reSend(); });
        if (dom().contextButton) dom().contextButton.addEventListener('click', function () { ui().toggleContextPanel(); });
        if (dom().settingsButton) dom().settingsButton.addEventListener('click', function () { settingsPanel().toggleInlineSettingsPanel(); });
        if (dom().sliderAiButton) dom().sliderAiButton.addEventListener('click', function (event) {
            if (event) event.preventDefault();
            try { window.open('https://data.slider-ai.ru/?utm_source=r7c_chatllm', '_blank', 'noopener'); } catch (_) {}
        });
        if (dom().inlineSaveSettings) dom().inlineSaveSettings.addEventListener('click', function () {
            settingsPanel().saveInlineSettings();
            chat().checkApiKey(true);
        });
        if (dom().inlineWebToolsProvider) dom().inlineWebToolsProvider.addEventListener('change', function () {
            settingsPanel().updateWebToolsHelp(dom().inlineWebToolsProvider.value);
        });
        if (dom().refreshSheetListButton) dom().refreshSheetListButton.addEventListener('click', function () {
            ui().renderWorkbookSheetList({ force: true });
        });
        if (dom().agentStopLoopButton) dom().agentStopLoopButton.addEventListener('click', function () { agent().stop(); });
        if (dom().agentRetryStepButton) dom().agentRetryStepButton.addEventListener('click', function () { agent().retryStep(); });
        if (dom().agentRunFromStartButton) dom().agentRunFromStartButton.addEventListener('click', function () { agent().runFromStepOne(); });
        if (dom().copyApiKeyButton) dom().copyApiKeyButton.addEventListener('click', function () {
            chat().copyToClipboard(dom().inlineTokenInput ? dom().inlineTokenInput.value.trim() : '');
        });

        if (dom().inlineTokenInput) {
            dom().inlineTokenInput.addEventListener('input', function () {
                if (dom().copyApiKeyButton) {
                    dom().copyApiKeyButton.disabled = !dom().inlineTokenInput.value.trim().length;
                }
                settingsPanel().validateInlineToken(false);
            });
            dom().inlineTokenInput.addEventListener('keydown', function (event) {
                if (event.key === 'Enter') {
                    event.preventDefault();
                    settingsPanel().saveInlineSettings();
                    chat().checkApiKey(true);
                }
            });
        }
        if (dom().inlineExaApiKeyInput) {
            dom().inlineExaApiKeyInput.addEventListener('keydown', function (event) {
                if (event.key === 'Enter') {
                    event.preventDefault();
                    settingsPanel().saveInlineSettings();
                    chat().checkApiKey(true);
                }
            });
        }
        if (dom().inlineBraveApiKeyInput) {
            dom().inlineBraveApiKeyInput.addEventListener('keydown', function (event) {
                if (event.key === 'Enter') {
                    event.preventDefault();
                    settingsPanel().saveInlineSettings();
                    chat().checkApiKey(true);
                }
            });
        }

        if (dom().settingsPanel) {
            dom().settingsPanel.addEventListener('click', function (event) {
                if (event.target === dom().settingsPanel) {
                    settingsPanel().closeInlineSettingsPanel();
                }
            });
        }

        if (dom().contextPanel) {
            dom().contextPanel.addEventListener('click', function (event) {
                if (event.target === dom().contextPanel) {
                    ui().closeContextPanel();
                }
            });
        }

        if (dom().threadHistoryPanel) {
            dom().threadHistoryPanel.addEventListener('click', function (event) {
                if (event.target === dom().threadHistoryPanel) {
                    ui().closeHistoryPanel();
                }
            });
        }
        if (dom().threadDialogBackdrop) {
            dom().threadDialogBackdrop.addEventListener('click', function (event) {
                if (event.target === dom().threadDialogBackdrop) {
                    ui().closeThreadDialog();
                }
            });
        }

        document.addEventListener('click', function (event) {
            if (dom().threadActionMenu && dom().threadActionMenu.style.display !== 'none') {
                var insideMenu = dom().threadActionMenu.contains(event.target);
                var isTrigger = event.target === dom().threadMenuButton || (dom().threadMenuButton && dom().threadMenuButton.contains(event.target));
                if (!insideMenu && !isTrigger) {
                    ui().closeThreadActionMenu();
                }
            }
            if (dom().attachmentPopover && dom().attachmentPopover.style.display !== 'none') {
                var insidePopover = dom().attachmentPopover.contains(event.target);
                var isAttachmentTrigger = event.target === dom().attachmentTriggerButton || (dom().attachmentTriggerButton && dom().attachmentTriggerButton.contains(event.target));
                if (!insidePopover && !isAttachmentTrigger) {
                    ui().closeAttachmentPopover();
                }
            }
        });

        document.addEventListener('keydown', function (event) {
            if (event.key === 'Escape' && dom().threadDialogBackdrop && dom().threadDialogBackdrop.style.display !== 'none') {
                ui().closeThreadDialog();
            } else if (event.key === 'Escape' && dom().threadActionMenu && dom().threadActionMenu.style.display !== 'none') {
                ui().closeThreadActionMenu();
            } else if (event.key === 'Escape' && dom().attachmentPopover && dom().attachmentPopover.style.display !== 'none') {
                ui().closeAttachmentPopover();
            } else if (event.key === 'Escape' && settingsPanel().isOpen()) {
                settingsPanel().closeInlineSettingsPanel();
            } else if (event.key === 'Escape' && ui().isContextPanelOpen()) {
                ui().closeContextPanel();
            } else if (event.key === 'Escape' && ui().isHistoryPanelOpen && ui().isHistoryPanelOpen()) {
                ui().closeHistoryPanel();
            }
        });

        if (dom().messageInput) {
            dom().messageInput.addEventListener('input', function () {
                if (!threadStore() || typeof threadStore().getActiveThread !== 'function') return;
                var activeThread = threadStore().getActiveThread();
                if (!activeThread) return;
                threadStore().replaceDraft(activeThread.id, dom().messageInput.value || '');
            });
            dom().messageInput.addEventListener('paste', function (event) {
                var clipboard = event.clipboardData || window.clipboardData;
                var imageFiles = extractClipboardImageFiles(clipboard);
                if (!imageFiles.length) return;

                event.preventDefault();
                var plainText = clipboard && typeof clipboard.getData === 'function'
                    ? String(clipboard.getData('text/plain') || '')
                    : '';
                if (plainText) {
                    insertTextAtCursor(dom().messageInput, plainText);
                }
                Promise.resolve(ui().consumeSelectedFiles(imageFiles, { source: 'clipboard' })).catch(function (error) {
                    if (ui().showThreadToast) {
                        ui().showThreadToast(error && error.message ? error.message : tr('Failed to paste image.'), { kind: 'error' });
                    }
                });
            });
            dom().messageInput.addEventListener('keydown', function (event) {
                if (event.key !== 'Enter') return;
                event.preventDefault();
                if (event.shiftKey) {
                    dom().messageInput.value += '\n';
                    if (threadStore() && typeof threadStore().getActiveThread === 'function') {
                        var activeThread = threadStore().getActiveThread();
                        if (activeThread) {
                            threadStore().replaceDraft(activeThread.id, dom().messageInput.value || '');
                        }
                    }
                    return;
                }
                submitComposerMessage();
            });
        }
    }

    function initializeRuntimeFromDom() {
        ui().bindDom();
        wireServiceContracts();
        
        var savedTheme = 'dark';
        if (root().platform && root().platform.storage && typeof root().platform.storage.readJson === 'function') {
            savedTheme = root().platform.storage.readJson('r7c_chat_theme', 'dark');
        }
        updateTheme(savedTheme, false);

        bindThemeToggle();
        if (ui().initializeThreadShell) {
            ui().initializeThreadShell();
        }
        bindDomEvents();
        settingsPanel().loadInlineSettings();
        chat().checkApiKey(true);
        ui().renderContextPanel();
    }

    window.Asc.plugin.init = function () {
        lifecycle().isClosing = false;
        lifecycle().closeScheduled = false;
        var info = window.Asc && window.Asc.plugin && window.Asc.plugin.info ? window.Asc.plugin.info : {};
        runtime().lang = String(info.lang || 'en').substring(0, 2).toLowerCase();
        applyConfiguredLanguage({ skipRetranslate: true });
        if (context()) context().loadContextState();
        if (agent()) {
            agent().reset();
            agent().resolveAgentMode();
        }
        if (chat()) {
            chat().checkApiKey(true);
        }
        updateGreeting();

        if (root().features && root().features.wordSlidesIntegration && root().features.wordSlidesIntegration.init) {
            root().features.wordSlidesIntegration.init();
        }
        if (window.Asc && window.Asc.plugin && typeof window.Asc.plugin.onTranslate === 'function') {
            window.Asc.plugin.onTranslate();
        }
    };

    window.Asc.plugin.onThemeChanged = function (theme) {
        var isManual = false;
        if (root().platform && root().platform.storage && typeof root().platform.storage.readJson === 'function') {
            isManual = root().platform.storage.readJson('r7c_chat_theme_manual', false);
        }
        
        if (isManual) {
            console.log('onThemeChanged: ignoring host theme because manual preference is set');
            return;
        }
        
        updateTheme(theme.type, false);
    };

    window.Asc.plugin.button = function () {
        requestPluginClose();
    };

    window.Asc.plugin.onTranslate = function () {
        updateGreeting();
        if (ui().renderThreadHeader) ui().renderThreadHeader();
        if (ui().renderThreadList) ui().renderThreadList();
        if (document.getElementById('thread-switcher-button')) {
            var historyLabel = tr('Chat history');
            document.getElementById('thread-switcher-button').setAttribute('aria-label', historyLabel);
            document.getElementById('thread-switcher-button').setAttribute('data-tooltip', historyLabel);
        }
        if (document.getElementById('threadHistoryKicker')) document.getElementById('threadHistoryKicker').innerText = tr('Chats');
        if (document.getElementById('threadHistoryCopy')) document.getElementById('threadHistoryCopy').innerText = tr('Recent conversations stay local on this device.');
        if (document.getElementById('new-chat-button')) {
            document.getElementById('new-chat-button').title = tr('New Chat');
            document.getElementById('new-chat-button').setAttribute('aria-label', tr('Create new chat'));
        }
        if (document.getElementById('context-button')) {
            document.getElementById('context-button').title = tr('Data Context');
            document.getElementById('context-button').setAttribute('aria-label', tr('Open data context'));
        }
        if (document.getElementById('settings-button')) {
            document.getElementById('settings-button').title = tr('Settings');
            document.getElementById('settings-button').setAttribute('aria-label', tr('Open settings'));
        }
        if (document.getElementById('slider-ai-button')) {
            document.getElementById('slider-ai-button').title = tr('Open Slider AI data portal');
            document.getElementById('slider-ai-button').setAttribute('aria-label', tr('Open Slider AI data portal'));
        }
        if (document.getElementById('theme-toggle')) {
            document.getElementById('theme-toggle').title = tr('Toggle Theme');
            document.getElementById('theme-toggle').setAttribute('aria-label', tr('Toggle Theme'));
        }
        if (document.getElementById('thread-menu-button')) {
            document.getElementById('thread-menu-button').title = tr('Thread Actions');
            document.getElementById('thread-menu-button').setAttribute('aria-label', tr('Open thread actions'));
        }
        if (document.getElementById('threadActionMenu')) {
            document.getElementById('threadActionMenu').setAttribute('aria-label', tr('Thread actions'));
        }
        var renameAction = document.querySelector('[data-thread-action="rename"]');
        if (renameAction) renameAction.innerText = tr('Rename');
        var exportAction = document.querySelector('[data-thread-action="export"]');
        if (exportAction) exportAction.innerText = tr('Export .docx');
        var deleteAction = document.querySelector('[data-thread-action="delete"]');
        if (deleteAction) deleteAction.innerText = tr('Delete chat');
        if (document.getElementById('attachmentTriggerButton')) {
            document.getElementById('attachmentTriggerButton').title = tr('Add attachment');
            document.getElementById('attachmentTriggerButton').setAttribute('aria-label', tr('Add attachment'));
        }
        if (document.getElementById('attachmentPopover')) {
            document.getElementById('attachmentPopover').setAttribute('aria-label', tr('Attachment menu'));
        }
        if (document.getElementById('attachmentAddFileAction')) document.getElementById('attachmentAddFileAction').querySelector('.attachment-popover-action-title').innerText = tr('Add file');
        if (document.getElementById('attachmentAddFileAction')) document.getElementById('attachmentAddFileAction').querySelector('.attachment-popover-action-meta').innerText = tr('PNG, JPG, WEBP, DOCX, XLSX, PPTX, PDF');
        if (document.getElementById('attachmentRecentToggleAction')) document.getElementById('attachmentRecentToggleAction').querySelector('.attachment-popover-action-title').innerText = tr('Recent Files');
        if (document.getElementById('attachmentRecentMeta')) document.getElementById('attachmentRecentMeta').innerText = tr('Quick access to recent documents');
        if (document.getElementById('attachmentRecentEmpty')) document.getElementById('attachmentRecentEmpty').innerText = tr('No recent files yet.');
        if (document.getElementById('send-button')) document.getElementById('send-button').title = tr('Send Message');
        if (document.getElementById('userInput')) document.getElementById('userInput').placeholder = tr('Type your message here...');
        if (document.getElementById('first-insert')) document.getElementById('first-insert').title = tr('insert document');
        if (document.getElementById('thread-history-close')) {
            document.getElementById('thread-history-close').title = tr('Close history');
            document.getElementById('thread-history-close').setAttribute('aria-label', tr('Close history'));
        }
        if (document.getElementById('thread-search-input')) document.getElementById('thread-search-input').placeholder = tr('Search chats by title');
        if (document.getElementById('threadDialogClose')) document.getElementById('threadDialogClose').setAttribute('aria-label', tr('Close dialog'));
        if (document.getElementById('threadDialogCancel')) document.getElementById('threadDialogCancel').innerText = tr('Cancel');
        if (document.getElementById('threadDialogConfirm')) document.getElementById('threadDialogConfirm').innerText = tr('Confirm');
        if (document.getElementById('inlineSettingsTitle')) document.getElementById('inlineSettingsTitle').innerText = tr('Operator');
        if (document.getElementById('inlineSettingsSummary')) document.getElementById('inlineSettingsSummary').innerText = tr('Choose a provider first, then configure only its credentials, model and capabilities.');
        if (document.getElementById('inlineLanguageSectionTitle')) document.getElementById('inlineLanguageSectionTitle').innerText = tr('Language');
        if (document.getElementById('inlineLanguageSectionHelp')) document.getElementById('inlineLanguageSectionHelp').innerText = tr('Choose the UI language and the default response language for the assistant.');
        if (document.getElementById('inlineLanguageLabel')) document.getElementById('inlineLanguageLabel').innerText = tr('Interface language');
        if (document.getElementById('inlineLanguageSelect')) {
            var languageSelect = document.getElementById('inlineLanguageSelect');
            Array.prototype.forEach.call(languageSelect.options, function (option) {
                if (option.value === 'auto') option.text = tr('Use editor language');
                if (option.value === 'en') option.text = tr('English');
                if (option.value === 'ru') option.text = tr('Russian');
            });
        }
        if (document.getElementById('inlineShowModelReasoningLabel')) document.getElementById('inlineShowModelReasoningLabel').innerText = tr('Show model thinking in trace');
        if (document.getElementById('inlineShowModelReasoningHelp')) document.getElementById('inlineShowModelReasoningHelp').innerText = tr('Shows a short reasoning summary from OpenRouter/OpenAI responses in the execution trace.');
        if (document.getElementById('inlineImageSectionTitle')) document.getElementById('inlineImageSectionTitle').innerText = tr('Image Generation');
        if (document.getElementById('inlineImageSectionHelp')) document.getElementById('inlineImageSectionHelp').innerText = tr('Image generation stays on the existing separate operator.');
        if (document.getElementById('inline_label_chooseModel')) document.getElementById('inline_label_chooseModel').innerText = tr('Choose operator');
        if (document.getElementById('inline_label_customModel')) document.getElementById('inline_label_customModel').innerText = tr('Enter model');
        if (document.getElementById('inlineAPILabel')) document.getElementById('inlineAPILabel').innerText = tr('OpenRouter API Key');
        if (document.getElementById('inlineTokenInput')) document.getElementById('inlineTokenInput').placeholder = tr('Set OpenRouter API key');
        if (document.getElementById('inline_label_chooseImageOperator')) document.getElementById('inline_label_chooseImageOperator').innerText = tr('Choose Image Operator');
        if (document.getElementById('inline_label_customImageModel')) document.getElementById('inline_label_customImageModel').innerText = tr('Image model');
        if (document.getElementById('inlineCustomImageModel')) document.getElementById('inlineCustomImageModel').placeholder = tr('e.g. google/gemini-2.5-flash-image');
        if (document.getElementById('inlineImageAPILabel')) document.getElementById('inlineImageAPILabel').innerText = tr('Image API Key (Optional)');
        if (document.getElementById('inlineImageTokenInput')) document.getElementById('inlineImageTokenInput').placeholder = tr('Leave empty to use the active provider key');
        if (document.getElementById('inlineDesktopToolsSectionTitle')) document.getElementById('inlineDesktopToolsSectionTitle').innerText = tr('Desktop Tools');
        if (document.getElementById('inlineDesktopToolsSectionHelp')) document.getElementById('inlineDesktopToolsSectionHelp').innerText = tr('Use native desktop host tools when the runtime exposes them. Macro fallback stays available.');
        if (document.getElementById('inlineDesktopAutomationModeLabel')) document.getElementById('inlineDesktopAutomationModeLabel').innerText = tr('Automation mode');
        if (document.getElementById('inlineDesktopAutomationMode')) {
            var desktopModeSelect = document.getElementById('inlineDesktopAutomationMode');
            Array.prototype.forEach.call(desktopModeSelect.options, function (option) {
                if (option.value === 'auto') option.text = tr('Auto: native first, macro fallback');
                if (option.value === 'native_tools') option.text = tr('Prefer native tools');
                if (option.value === 'macro_only') option.text = tr('Macro only');
            });
        }
        if (document.getElementById('inlineDesktopToolsStatus')) document.getElementById('inlineDesktopToolsStatus').innerText = tr('Desktop tools runtime status is checked when the panel opens.');
        if (document.getElementById('inlineWebToolsSectionTitle')) document.getElementById('inlineWebToolsSectionTitle').innerText = tr('Web Tools');
        if (document.getElementById('inlineWebToolsSectionHelp')) document.getElementById('inlineWebToolsSectionHelp').innerText = tr('Web tools are available only for OpenRouter in this release.');
        if (document.getElementById('inlineWebToolsProviderLabel')) document.getElementById('inlineWebToolsProviderLabel').innerText = tr('Web Tools Provider');
        if (document.getElementById('inlineExaApiLabel')) document.getElementById('inlineExaApiLabel').innerText = tr('Exa API Key');
        if (document.getElementById('inlineExaApiKeyInput')) document.getElementById('inlineExaApiKeyInput').placeholder = tr('Used when Exa web tools are enabled');
        if (document.getElementById('inlineBraveApiLabel')) document.getElementById('inlineBraveApiLabel').innerText = tr('Brave API Key');
        if (document.getElementById('inlineBraveApiKeyInput')) document.getElementById('inlineBraveApiKeyInput').placeholder = tr('Used when Brave web tools are enabled');
        if (document.getElementById('inlineWebToolsHelp')) document.getElementById('inlineWebToolsHelp').innerText = tr('Enable Exa or Brave to allow model web search and crawling tools.');
        if (document.getElementById('copyApiKeyButton')) document.getElementById('copyApiKeyButton').innerText = tr('Copy Main API Key');
        if (document.getElementById('inlineSaveSettings')) document.getElementById('inlineSaveSettings').innerText = tr('Save Settings');
        if (settingsPanel() && settingsService() && typeof settingsPanel().renderInlineSettings === 'function') settingsPanel().renderInlineSettings(settingsService().loadSettings());
        if (document.getElementById('contextPanelTitle')) document.getElementById('contextPanelTitle').innerText = tr('Data Context');
        if (document.getElementById('refreshSheetList')) document.getElementById('refreshSheetList').innerText = tr('Refresh sheets');
        if (document.getElementById('agentStopLoop')) document.getElementById('agentStopLoop').innerText = tr('Stop');
        if (document.getElementById('agentRetryStep')) document.getElementById('agentRetryStep').innerText = tr('Retry step');
        if (document.getElementById('agentRunFromStart')) document.getElementById('agentRunFromStart').innerText = tr('Run from step 1');
        if (document.getElementById('contextTraceTitle')) document.getElementById('contextTraceTitle').innerText = tr('Execution Trace');
        if (document.getElementById('stop-button')) document.getElementById('stop-button').innerHTML = '<span>&#x25A0;</span>' + tr('Stop');
        if (document.getElementById('regenerate-button')) document.getElementById('regenerate-button').innerHTML = '<span>&#x21BA;</span>' + tr('Regenerate');
        if (ui().renderContextPanel) ui().renderContextPanel();
    };

    bindContextMenuHandlers();

    document.addEventListener('DOMContentLoaded', function () {
        initializeRuntimeFromDom();
    });
})(window, undefined);
