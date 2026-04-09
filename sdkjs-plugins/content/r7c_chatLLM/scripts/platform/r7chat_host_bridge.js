(function (rootFactory) {
    'use strict';

    if (typeof module === 'object' && module.exports) {
        module.exports = rootFactory(typeof globalThis !== 'undefined' ? globalThis : global);
        return;
    }

    rootFactory(typeof window !== 'undefined' ? window : this);
})(function (globalRoot) {
    'use strict';

    var root = globalRoot.R7Chat = globalRoot.R7Chat || {};
    root.platform = root.platform || {};
    root.runtime = root.runtime || {};
    root.runtime.host = root.runtime.host || {
        editorCommandQueue: Promise.resolve(),
        editorCommandCounter: 0
    };
    var LOCAL_TRANSLATIONS = {
        ru: {
            'Add attachment': 'Добавить вложение',
            'Add file': 'Добавить файл',
            'Analyse the selected text': 'Проанализировать выделенный текст',
            'Agent mode: OFF': 'Режим агента: ВЫКЛ',
            'Agent mode: ON': 'Режим агента: ВКЛ',
            'Attach': 'Прикрепить',
            'Attached': 'Прикреплено',
            'Attachment': 'Вложение',
            'Attachment menu': 'Меню вложений',
            'Automatic data context is currently available for Word and Spreadsheet editors.': 'Автоматический контекст данных сейчас доступен только в редакторах Word и Spreadsheet.',
            'Base URL': 'Base URL',
            'Base64 client secret': 'Base64 client secret',
            'Brave API Key': 'Brave API Key',
            'Brave API key is required when Brave web tools are enabled': 'Нужен Brave API key, когда включены Brave web tools',
            'Brave crawling uses direct URL fetch fallback.': 'Для Brave crawling используется прямой fallback через загрузку URL.',
            'Build a formula for this table': 'Построить формулу для этой таблицы',
            'Cancel': 'Отмена',
            'Chat deleted.': 'Чат удалён.',
            'Chat history': 'История чатов',
            'Chat renamed.': 'Чат переименован.',
            'Chats': 'Чаты',
            'Check that image model is set correctly in Settings.': 'Проверьте, что image model правильно задана в настройках.',
            'Choose a provider first, then configure only its credentials, model and capabilities.': 'Сначала выберите провайдера, затем настройте только его credentials, model и capabilities.',
            'Choose the UI language and the default response language for the assistant.': 'Выберите язык интерфейса и язык ответов ассистента по умолчанию.',
            'Choose Image Operator': 'Выбрать image operator',
            'Choose operator': 'Выбрать оператор',
            'Close dialog': 'Закрыть диалог',
            'Close history': 'Закрыть историю',
            'Cloud folder ID': 'Cloud folder ID',
            'Complete the selected text': 'Дополнить выделенный текст',
            'Confirm': 'Подтвердить',
            'Copy Main API Key': 'Скопировать основной API key',
            'Copy URL': 'Скопировать URL',
            'Create an outline': 'Создать план',
            'Create new chat': 'Создать новый чат',
            'Correct spelling and grammar': 'Исправить орфографию и грамматику',
            'Data Context': 'Контекст данных',
            'Desktop Tools': 'Desktop tools',
            'Desktop tools are not exposed in this R7 runtime. Macro automation stays available.': 'В этом runtime R7 desktop tools не доступны. Макросная автоматизация остаётся доступной.',
            'Desktop tools bridge is not available in the plugin runtime.': 'Bridge для desktop tools недоступен в runtime плагина.',
            'Desktop tools catalog failed to parse. Check DevTools for runtime details.': 'Не удалось разобрать каталог desktop tools. Проверьте DevTools для деталей runtime.',
            'Desktop tools catalog is unavailable in this runtime. Macro automation stays available.': 'Каталог desktop tools недоступен в этом runtime. Макросная автоматизация остаётся доступной.',
            'Desktop tools catalog is readable ({count} tools), but callToolFunction is unavailable.': 'Каталог desktop tools доступен ({count} tools), но callToolFunction недоступен.',
            'Desktop tools ready: {count} tools available for execution.': 'Desktop tools готовы: для выполнения доступно {count} tools.',
            'Desktop tools runtime status is checked when the panel opens.': 'Статус desktop tools проверяется при открытии панели.',
            'Delete': 'Удалить',
            'Delete "{title}" and all of its messages? This action cannot be undone.': 'Удалить "{title}" и все сообщения этого чата? Это действие нельзя отменить.',
            'Delete chat': 'Удалить чат',
            'Detailed log': 'Подробный лог',
            'Done': 'Готово',
            'Draft a polished response': 'Подготовить аккуратный ответ',
            'Draft a reply in the same tone': 'Написать ответ в том же тоне',
            'Edit chat': 'Редактировать чат',
            'Empty message': 'Пустое сообщение',
            'Empty response from provider': 'Пустой ответ от провайдера',
            'Enable Exa or Brave to allow model web search and crawling tools.': 'Включите Exa или Brave, чтобы разрешить web search и crawling tools модели.',
            'English': 'English',
            'Enter model': 'Ввести модель',
            'Error:': 'Ошибка:',
            'Every chat request automatically includes the full document context with smart truncation.': 'Каждый запрос автоматически включает полный контекст документа с умным усечением.',
            'Execution failed': 'Выполнение завершилось с ошибкой',
            'Execution Trace': 'Трассировка выполнения',
            'Exa API Key': 'Exa API Key',
            'Exa API key is required when Exa web tools are enabled': 'Нужен Exa API key, когда включены Exa web tools',
            'Exa search and crawl use the Exa API.': 'Для поиска и crawl через Exa используется Exa API.',
            'Explain the selected cells': 'Объяснить выбранные ячейки',
            'Explain the selected text': 'Объяснить выделенный текст',
            'Explain this topic simply': 'Объяснить тему простыми словами',
            'Expand the selected text': 'Расширить выделенный текст',
            'Export .docx': 'Экспорт .docx',
            'Export failed.': 'Экспорт не удался.',
            'Failed to add image.': 'Не удалось добавить изображение.',
            'Failed to paste image.': 'Не удалось вставить изображение.',
            'Folder ID': 'Folder ID',
            'Generate draft': 'Сгенерировать черновик',
            'Generating image...': 'Генерация изображения...',
            'Get outline of a topic': 'Сделать план по теме',
            'GigaChat supports one image attachment per message in this release.': 'В этой версии GigaChat поддерживает только одно изображение на сообщение.',
            'Give this conversation a clearer name. The new title stays local on this device.': 'Дайте этому диалогу более понятное имя. Новое название сохранится только на этом устройстве.',
            'Hi there! I am {r7c}.ChatLLM, how can I help you today?': 'Привет! Я {r7c}.ChatLLM. Чем помочь?',
            'Hide': 'Скрыть',
            'Image': 'Изображение',
            'Image API Key (Optional)': 'Image API Key (необязательно)',
            'Image generated successfully': 'Изображение успешно создано',
            'Image generation bridge is unavailable': 'Мост генерации изображений недоступен',
            'Image generation failed': 'Генерация изображения не удалась',
            'Image generation API returned an error': 'API генерации изображений вернул ошибку',
            'Image generation stays on the existing separate operator.': 'Генерация изображений остаётся на отдельном операторе.',
            'Image Generation': 'Генерация изображений',
            'Image model': 'Image model',
            'Insert': 'Вставить',
            'insert document': 'вставить документ',
            'Interface language': 'Язык интерфейса',
            'Just now': 'Только что',
            'Language': 'Язык',
            'Leave empty to use the active provider key': 'Оставьте пустым, чтобы использовать ключ активного провайдера',
            'Load models': 'Загрузить модели',
            'Loading models...': 'Загрузка моделей...',
            'Loading preview...': 'Загрузка превью...',
            'Loading sheets...': 'Загрузка листов...',
            'Macro only': 'Только макросы',
            'Message copied to clipboard': 'Сообщение скопировано в буфер обмена',
            'Model': 'Модель',
            'Model discovery failed, custom model input stays available.': 'Не удалось получить список моделей, но поле для своей модели остаётся доступным.',
            'Model is required for the active provider': 'Для активного провайдера нужна модель',
            'Model returned no image. Ensure an image generation model is selected in settings.': 'Модель не вернула изображение. Убедитесь, что в настройках выбрана модель для генерации изображений.',
            'Models loaded': 'Модели загружены',
            'New Chat': 'Новый чат',
            'No auto context': 'Нет автоконтекста',
            'No chats yet': 'Чатов пока нет',
            'No models returned, you can still type a custom model.': 'Модели не вернулись, но вы всё равно можете ввести свою модель вручную.',
            'No recent files yet.': 'Пока нет недавних файлов.',
            'No steps yet.': 'Шагов пока нет.',
            'No supported files were added.': 'Не было добавлено ни одного поддерживаемого файла.',
            'Nothing found': 'Ничего не найдено',
            'Only PNG, JPEG, WebP, DOCX, XLSX, PPTX and PDF files are supported.': 'Поддерживаются только PNG, JPEG, WebP, DOCX, XLSX, PPTX и PDF.',
            'Open data context': 'Открыть контекст данных',
            'OpenRouter API Key': 'OpenRouter API Key',
            'Open settings': 'Открыть настройки',
            'Open thread actions': 'Открыть действия чата',
            'OpenRouter is the only provider that can use web tools in this release.': 'В этой версии только OpenRouter может использовать web tools.',
            'Operator': 'Оператор',
            'PNG, JPG, WEBP, DOCX, XLSX, PPTX, PDF': 'PNG, JPG, WEBP, DOCX, XLSX, PPTX, PDF',
            'Plugin host is not ready yet. Context collection and macro execution are disabled.': 'Plugin host ещё не готов. Сбор контекста и выполнение макросов отключены.',
            'Preview': 'Предпросмотр',
            'Quick access to recent documents': 'Быстрый доступ к недавним документам',
            'Quick access to recent files': 'Быстрый доступ к недавним файлам',
            'Recent conversations stay local on this device.': 'Недавние диалоги хранятся локально на этом устройстве.',
            'Recent Files': 'Недавние файлы',
            'Refresh sheets': 'Обновить листы',
            'Regenerate': 'Повторить',
            'Remove attachment': 'Удалить вложение',
            'Rename': 'Переименовать',
            'Rename chat': 'Переименовать чат',
            'Retry step': 'Повторить шаг',
            'rows': 'строки',
            'Russian': 'Русский',
            'Running...': 'Выполняется...',
            'Save': 'Сохранить',
            'Search chats by title': 'Искать чаты по названию',
            'Selected operator': 'Выбранный оператор',
            'Set API key': 'Укажите API key',
            'Set OpenRouter API key': 'Укажите OpenRouter API key',
            'Set your {provider} credentials first': 'Сначала укажите credentials для {provider}',
            'Send Message': 'Отправить сообщение',
            'Settings': 'Настройки',
            'Settings saved': 'Настройки сохранены',
            'Show model thinking in trace': 'Показывать ход мыслей модели в trace',
            'Shows a short reasoning summary from OpenRouter/OpenAI responses in the execution trace.': 'Показывает краткую сводку рассуждений модели из ответов OpenRouter/OpenAI в трассировке выполнения.',
            'Sheet is empty.': 'Лист пуст.',
            'Sheet navigation is available in spreadsheet mode.': 'Навигация по листам доступна в режиме spreadsheet.',
            'Sheets: active sheet': 'Листы: активный лист',
            'Shorten the selected text': 'Сократить выделенный текст',
            'Start a conversation and it will appear here automatically.': 'Начните диалог, и он автоматически появится здесь.',
            'Status': 'Статус',
            'Stop': 'Стоп',
            'Auto: native first, macro fallback': 'Auto: сначала native tools, потом macro fallback',
            'Automation mode': 'Режим автоматизации',
            'Sorry, please select text in a cell to proceed correct selected spelling and grammar.': 'Пожалуйста, сначала выберите текст в ячейке, чтобы исправить орфографию и грамматику.',
            'Sorry, please select text in a cell to proceed explain the selected text.': 'Пожалуйста, сначала выберите текст в ячейке, чтобы объяснить выделенный текст.',
            'Sorry, please select text in a cell to proceed summarize the selected text.': 'Пожалуйста, сначала выберите текст в ячейке, чтобы суммировать выделенный текст.',
            'Sorry, please select text in a cell to proceed translate selected text.': 'Пожалуйста, сначала выберите текст в ячейке, чтобы перевести выделенный текст.',
            'Summarize the active sheet': 'Суммировать активный лист',
            'Summarize the selected text': 'Суммировать выделенный текст',
            'Summarize this document': 'Суммировать этот документ',
            'Text': 'Текст',
            'The selected model does not support image input. Choose a vision-capable model and try again.': 'Выбранная модель не поддерживает ввод изображений. Выберите модель с поддержкой vision и попробуйте снова.',
            'This editor does not expose sheet/document context mode here.': 'Этот редактор не предоставляет здесь режим контекста листа/документа.',
            'This file is already attached.': 'Этот файл уже прикреплён.',
            'This ONLYOFFICE runtime has browser fetch only. Exa or Brave may fail if CORS blocks the request.': 'В этом ONLYOFFICE runtime доступен только browser fetch. Exa или Brave могут не сработать, если запрос заблокирует CORS.',
            'Thread Actions': 'Действия чата',
            'Thread actions': 'Действия чата',
            'Thread exported to .docx.': 'Чат экспортирован в .docx.',
            'Toggle Theme': 'Переключить тему',
            'Try a shorter title or clear the search field.': 'Попробуйте более короткое название или очистите поле поиска.',
            'Type your message here...': 'Введите сообщение...',
            'Unable to enumerate sheets.': 'Не удалось получить список листов.',
            'Unable to load this sheet.': 'Не удалось загрузить этот лист.',
            'Use editor language': 'Использовать язык редактора',
            'Used when Brave web tools are enabled': 'Используется, когда включены Brave web tools',
            'Used when Exa web tools are enabled': 'Используется, когда включены Exa web tools',
            'Vision': 'Vision',
            'Waiting for editor host...': 'Ожидание editor host...',
            'Wait for the current response to finish first.': 'Дождитесь завершения текущего ответа.',
            'Web tools': 'Web tools',
            'Web Tools': 'Web tools',
            'Web tools are available only for OpenRouter in this release.': 'В этой версии web tools доступны только для OpenRouter.',
            'Web tools are currently available only when OpenRouter is the active provider.': 'Сейчас web tools доступны только когда активным провайдером выбран OpenRouter.',
            'Web Tools Provider': 'Провайдер web tools',
            'Word: full document': 'Word: полный документ',
            'Prefer native tools': 'Предпочитать native tools',
            'Use native desktop host tools when the runtime exposes them. Macro fallback stays available.': 'Используйте native desktop host tools, когда runtime их предоставляет. Macro fallback остаётся доступным.',
            'YandexGPT is wired through the compatibility API and is text-only in this release.': 'YandexGPT подключён через compatibility API и в этой версии работает только с текстом.',
            '1 attachment': '1 вложение',
            'attachments': 'вложений',
            'cols': 'столбцы',
            'e.g. GigaChat-Pro': 'например, GigaChat-Pro',
            'e.g. google/gemini-2.5-flash-image': 'например, google/gemini-2.5-flash-image',
            'e.g. yandexgpt/latest': 'например, yandexgpt/latest',
            'images': 'изображения',
            'messages': 'сообщений',
            'recent files': 'недавних файлов'
        }
    };

    function hasAscHost() {
        return !!(globalRoot.Asc && globalRoot.Asc.plugin);
    }

    function isPluginHostReady() {
        return !!(globalRoot.Asc && globalRoot.Asc.plugin && globalRoot.Asc.plugin.info);
    }

    function getEditorTypeSafe() {
        if (!isPluginHostReady()) return '';
        return String(globalRoot.Asc.plugin.info.editorType || '');
    }

    function getHostDiagnostics() {
        if (!globalRoot.Asc) return 'Asc host is not available.';
        if (!globalRoot.Asc.plugin) return 'Asc.plugin is not available.';
        if (!globalRoot.Asc.plugin.info) return 'Asc.plugin.info is not available yet.';
        return '';
    }

    function getSettingsService() {
        return root.features && root.features.settings ? root.features.settings : null;
    }

    function getLanguagePreference() {
        var settingsService = getSettingsService();
        if (!settingsService || typeof settingsService.loadSettings !== 'function') return 'auto';
        try {
            var settings = settingsService.loadSettings();
            var preference = String(settings && settings.language || 'auto').trim().toLowerCase();
            return preference === 'en' || preference === 'ru' ? preference : 'auto';
        } catch (error) {
            return 'auto';
        }
    }

    function getHostLanguageCode() {
        var info = isPluginHostReady() ? globalRoot.Asc.plugin.info : null;
        var language = info && info.lang ? String(info.lang) : String(root.runtime && root.runtime.lang || 'en');
        return String(language || 'en').substring(0, 2).toLowerCase() || 'en';
    }

    function getUiLanguageCode() {
        var preference = getLanguagePreference();
        return preference === 'en' || preference === 'ru' ? preference : getHostLanguageCode();
    }

    function translate(text) {
        var key = String(text || '');
        var languagePreference = getLanguagePreference();
        var languageCode = getUiLanguageCode();
        var dictionary = LOCAL_TRANSLATIONS[languageCode] || null;

        if (dictionary && Object.prototype.hasOwnProperty.call(dictionary, key)) {
            return dictionary[key];
        }
        if (languagePreference === 'auto' && hasAscHost() && typeof globalRoot.Asc.plugin.tr === 'function') {
            return globalRoot.Asc.plugin.tr(key);
        }
        if (languageCode !== 'en' && hasAscHost() && typeof globalRoot.Asc.plugin.tr === 'function' && getHostLanguageCode() === languageCode) {
            return globalRoot.Asc.plugin.tr(key);
        }
        return key;
    }

    function callEditorCommand(command, scopeData, options) {
        var runtimeHost = root.runtime.host;
        var commandId = ++runtimeHost.editorCommandCounter;
        var run = function () {
            return new Promise(function (resolve) {
                var settled = false;
                var timerId = 0;
                try {
                    if (!hasAscHost() || typeof globalRoot.Asc.plugin.callCommand !== 'function') {
                        resolve(null);
                        return;
                    }

                    var settings = options && typeof options === 'object' ? options : {};
                    var recalculate = settings.recalculate === false ? false : true;
                    var timeoutMs = Number(settings.timeoutMs || 0);
                    function finalize(result) {
                        if (settled) return;
                        settled = true;
                        if (timerId) {
                            try { globalRoot.clearTimeout(timerId); } catch (clearError) {}
                        }
                        resolve(result);
                    }
                    if (scopeData && globalRoot.Asc.scope) {
                        try {
                            Object.keys(scopeData).forEach(function (key) {
                                globalRoot.Asc.scope[key] = scopeData[key];
                            });
                        } catch (scopeError) {
                            console.error('callEditorCommand: Failed to set Asc.scope data', scopeError);
                        }
                    }

                    if (timeoutMs > 0 && typeof globalRoot.setTimeout === 'function') {
                        timerId = globalRoot.setTimeout(function () {
                            console.warn('callEditorCommand: Timeout for command #' + commandId + ' after ' + timeoutMs + ' ms');
                            finalize(null);
                        }, timeoutMs);
                    }

                    globalRoot.Asc.plugin.callCommand(command, false, recalculate, function (result) {
                        if (result === undefined || result === null) {
                            console.error('callEditorCommand: Asc.plugin.callCommand returned ' + result + ' for command #' + commandId, command);
                        } else {
                            console.log('callEditorCommand: Success for command #' + commandId);
                        }
                        finalize(result);
                    });
                } catch (error) {
                    console.error('callEditorCommand: Caught synchronous error #' + commandId, error);
                    if (timerId) {
                        try { globalRoot.clearTimeout(timerId); } catch (clearError2) {}
                    }
                    resolve(null);
                }
            });
        };

        var queued = runtimeHost.editorCommandQueue.then(run, run);
        runtimeHost.editorCommandQueue = queued.then(function () { return null; }, function () { return null; });
        return queued;
    }

    function callEditorMethod(name, args) {
        return new Promise(function (resolve) {
            try {
                if (!hasAscHost() || typeof globalRoot.Asc.plugin.executeMethod !== 'function') {
                    resolve(null);
                    return;
                }
                globalRoot.Asc.plugin.executeMethod(name, args || [], function (result) {
                    resolve(result);
                });
            } catch (error) {
                console.error('callEditorMethod failed: ' + name, error);
                resolve(null);
            }
        });
    }

    var api = {
        isPluginHostReady: isPluginHostReady,
        getEditorTypeSafe: getEditorTypeSafe,
        getHostDiagnostics: getHostDiagnostics,
        getLanguagePreference: getLanguagePreference,
        getUiLanguageCode: getUiLanguageCode,
        translate: translate,
        callEditorCommand: callEditorCommand,
        callEditorMethod: callEditorMethod
    };

    root.platform.hostBridge = api;
    return api;
});
