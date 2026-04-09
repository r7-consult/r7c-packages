const assert = require('assert');
const settingsService = require('./r7chat_settings_service');

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

function testLegacyMigrationToV3() {
    const storage = createStorage({
        apikey: 'sk-or-v1-example-key-1234567890',
        model: 'openrouter/auto',
        image_model: 'openai/dall-e-3'
    });

    const settings = settingsService.loadSettings(storage);

    assert.strictEqual(settings.activeProvider, 'openrouter');
    assert.strictEqual(settings.providers.openrouter.apiKey, 'sk-or-v1-example-key-1234567890');
    assert.strictEqual(settings.providers.openrouter.model, 'openrouter/auto');
    assert.strictEqual(settings.providers.openrouter.reasoningEffort, 'medium');
    assert.strictEqual(settings.imageModel, 'google/gemini-2.5-flash-image');
    assert.strictEqual(settings.trace.showModelReasoning, true);
    assert.ok(storage.getItem(settingsService.STORAGE_KEY));
}

function testPreviousV2Migration() {
    const storage = createStorage({
        [settingsService.PREVIOUS_STORAGE_KEY]: JSON.stringify({
            provider: 'openrouter',
            apiKey: 'sk-or-v1-from-v2-1234567890',
            model: 'openrouter/auto',
            imageModel: 'google/gemini-2.5-flash-image-preview',
            webTools: {
                provider: 'exa',
                providers: {
                    exa: { apiKey: 'exa-secret' },
                    brave: { apiKey: '' }
                }
            }
        })
    });

    const settings = settingsService.loadSettings(storage);

    assert.strictEqual(settings.activeProvider, 'openrouter');
    assert.strictEqual(settings.providers.openrouter.apiKey, 'sk-or-v1-from-v2-1234567890');
    assert.strictEqual(settings.imageModel, 'google/gemini-2.5-flash-image-preview');
    assert.strictEqual(settings.webTools.provider, 'exa');
    assert.strictEqual(settings.webTools.providers.exa.apiKey, 'exa-secret');
}

function testProviderSwitchKeepsIndependentCredentials() {
    const storage = createStorage();
    settingsService.saveSettings(storage, {
        providers: {
            openrouter: {
                apiKey: 'sk-or-v1-openrouter-key-1234567890',
                model: 'openrouter/auto'
            },
            openai: {
                apiKey: 'sk-openai-key-1234567890',
                model: 'gpt-5-mini'
            }
        },
        activeProvider: 'openrouter'
    });

    const saved = settingsService.saveSettings(storage, {
        activeProvider: 'openai'
    });

    assert.strictEqual(saved.activeProvider, 'openai');
    assert.strictEqual(saved.apiKey, 'sk-openai-key-1234567890');
    assert.strictEqual(saved.model, 'gpt-5-mini');
    assert.strictEqual(saved.providers.openrouter.apiKey, 'sk-or-v1-openrouter-key-1234567890');
    assert.strictEqual(saved.providers.openrouter.model, 'openrouter/auto');
    assert.strictEqual(storage.getItem('apikey'), 'sk-openai-key-1234567890');
    assert.strictEqual(storage.getItem('model'), 'gpt-5-mini');
}

function testWebToolsRoundTrip() {
    const storage = createStorage();
    const saved = settingsService.saveSettings(storage, {
        webTools: {
            provider: 'brave',
            providers: {
                exa: { apiKey: 'exa-secret' },
                brave: { apiKey: 'brave-secret' }
            }
        }
    });

    const loaded = settingsService.loadSettings(storage);

    assert.deepStrictEqual(loaded.webTools, saved.webTools);
}

function testValidateProviderConfig() {
    assert.strictEqual(
        settingsService.validateProviderConfig('yandex', {
            apiKey: 'yandex-key-1234567890',
            folderId: ''
        }).isValid,
        false
    );
    assert.strictEqual(
        settingsService.validateProviderConfig('gigachat', {
            authKey: 'gigachat-auth-1234567890',
            scope: 'GIGACHAT_API_PERS'
        }).isValid,
        true
    );
}

function testSaveSettingsRemapsLegacyImageModel() {
    const storage = createStorage();
    const saved = settingsService.saveSettings(storage, {
        imageModel: 'openai/dall-e-3'
    });

    assert.strictEqual(saved.imageModel, 'google/gemini-2.5-flash-image');
    assert.strictEqual(settingsService.loadSettings(storage).imageModel, 'google/gemini-2.5-flash-image');
}

function testLanguageDefaultsToAuto() {
    const settings = settingsService.loadSettings(createStorage());
    assert.strictEqual(settings.language, 'auto');
    assert.strictEqual(settings.trace.showModelReasoning, true);
    assert.strictEqual(settings.providers.openai.reasoningEffort, 'medium');
}

function testLanguageRoundTripAndSanitization() {
    const storage = createStorage();
    const saved = settingsService.saveSettings(storage, {
        language: 'ru'
    });
    assert.strictEqual(saved.language, 'ru');
    assert.strictEqual(settingsService.loadSettings(storage).language, 'ru');

    const invalid = settingsService.saveSettings(storage, {
        language: 'es'
    });
    assert.strictEqual(invalid.language, 'auto');
}

function testTraceAndReasoningEffortRoundTrip() {
    const storage = createStorage();
    const saved = settingsService.saveSettings(storage, {
        trace: {
            showModelReasoning: false
        },
        providers: {
            openrouter: {
                reasoningEffort: 'off'
            },
            openai: {
                reasoningEffort: 'high'
            }
        }
    });

    assert.strictEqual(saved.trace.showModelReasoning, false);
    assert.strictEqual(saved.providers.openrouter.reasoningEffort, 'off');
    assert.strictEqual(saved.providers.openai.reasoningEffort, 'high');

    const loaded = settingsService.loadSettings(storage);
    assert.strictEqual(loaded.trace.showModelReasoning, false);
    assert.strictEqual(loaded.providers.openrouter.reasoningEffort, 'off');
    assert.strictEqual(loaded.providers.openai.reasoningEffort, 'high');
}

function run() {
    testLegacyMigrationToV3();
    testPreviousV2Migration();
    testProviderSwitchKeepsIndependentCredentials();
    testWebToolsRoundTrip();
    testValidateProviderConfig();
    testSaveSettingsRemapsLegacyImageModel();
    testLanguageDefaultsToAuto();
    testLanguageRoundTripAndSanitization();
    testTraceAndReasoningEffortRoundTrip();
    console.log('r7chat_settings_service.test.js: ok');
}

run();
