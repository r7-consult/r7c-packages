const assert = require('assert');

function loadFresh(modulePath) {
    delete require.cache[require.resolve(modulePath)];
    return require(modulePath);
}

function createJsonResponse(payload, status = 200) {
    return {
        ok: status >= 200 && status < 300,
        status,
        statusText: status === 200 ? 'OK' : 'ERR',
        headers: { 'content-type': 'application/json' },
        bodyText: JSON.stringify(payload),
        json() {
            return payload;
        },
        text() {
            return JSON.stringify(payload);
        }
    };
}

function createRuntime(overrides) {
    const requestLog = [];
    global.R7Chat = {
        platform: {
            http: {
                async request(options) {
                    requestLog.push(options);
                    if (overrides && typeof overrides.request === 'function') {
                        return overrides.request(options, requestLog);
                    }
                    return createJsonResponse({ ok: true });
                },
                async requestStream() {
                    return {
                        ok: true,
                        status: 200,
                        statusText: 'OK',
                        headers: { 'content-type': 'text/event-stream' },
                        reader: {
                            async read() {
                                return { done: true, value: '' };
                            }
                        }
                    };
                },
                requestJson: async function () {
                    const response = await this.request.apply(this, arguments);
                    return { response, data: response.json() };
                }
            }
        },
        features: {
            settings: {
                loadSettings() {
                    return {
                        activeProvider: 'openrouter',
                        provider: 'openrouter',
                        apiKey: '',
                        model: 'openrouter/auto',
                        imageModel: 'google/gemini-2.5-flash-image',
                        imageApiKey: '',
                        providers: {
                            openrouter: {
                                apiKey: 'sk-or-v1-default-openrouter-1234567890',
                                model: 'openrouter/auto',
                                baseUrl: 'https://openrouter.ai/api/v1'
                            },
                            openai: {
                                apiKey: 'sk-openai-key-1234567890',
                                model: 'gpt-5-mini',
                                baseUrl: 'https://api.openai.com/v1'
                            },
                            anthropic: {
                                apiKey: 'anthropic-key-1234567890',
                                model: 'claude-sonnet-4-5',
                                baseUrl: 'https://api.anthropic.com/v1'
                            },
                            gemini: {
                                apiKey: 'gemini-key-1234567890',
                                model: 'gemini-2.5-flash',
                                baseUrl: 'https://generativelanguage.googleapis.com/v1beta'
                            },
                            yandex: {
                                apiKey: 'yandex-key-1234567890',
                                folderId: 'folder-123',
                                model: 'yandexgpt/latest',
                                baseUrl: 'https://ai.api.cloud.yandex.net/v1'
                            },
                            gigachat: {
                                authKey: 'gigachat-auth-1234567890',
                                scope: 'GIGACHAT_API_PERS',
                                model: 'GigaChat-Pro',
                                apiBaseUrl: 'https://gigachat.devices.sberbank.ru/api/v1',
                                oauthBaseUrl: 'https://ngw.devices.sberbank.ru:9443/api/v2/oauth'
                            },
                            mistral: {
                                apiKey: 'mistral-key-1234567890',
                                model: 'pixtral-12b-latest',
                                baseUrl: 'https://api.mistral.ai/v1'
                            },
                            deepseek: {
                                apiKey: 'deepseek-key-1234567890',
                                model: 'deepseek-chat',
                                baseUrl: 'https://api.deepseek.com'
                            }
                        }
                    };
                }
            }
        }
    };
    return {
        registry: loadFresh('./openrouter_client'),
        requestLog
    };
}

function createImageAttachment() {
    return {
        attachmentType: 'image',
        name: 'screen.png',
        dataUrl: 'data:image/png;base64,aGVsbG8='
    };
}

function testRegistryCapabilities() {
    const { registry } = createRuntime();
    assert.deepStrictEqual(registry.getBuiltInProviderIds(), [
        'openrouter',
        'openai',
        'anthropic',
        'gemini',
        'yandex',
        'gigachat',
        'mistral',
        'deepseek'
    ]);
    assert.strictEqual(registry.getChatCapabilities({ provider: 'openai', model: 'gpt-5-mini' }).supportsVision, true);
    assert.strictEqual(registry.getChatCapabilities({ provider: 'deepseek', model: 'deepseek-chat' }).supportsVision, false);
    assert.strictEqual(registry.getChatCapabilities({ provider: 'openrouter', model: 'openrouter/auto' }).supportsTools, true);
}

function testOpenAiPayloadUsesImageUrl() {
    const { registry } = createRuntime();
    const request = registry.getProvider('openai').buildRequest({
        messages: [{
            role: 'user',
            content: 'Analyze this',
            attachments: [createImageAttachment()]
        }],
        config: {
            provider: 'openai',
            apiKey: 'sk-openai-key-1234567890',
            model: 'gpt-5-mini',
            baseUrl: 'https://api.openai.com/v1'
        }
    });

    assert.strictEqual(request.body.messages[0].content[1].type, 'image_url');
    assert.strictEqual(request.body.messages[0].content[1].image_url.url, 'data:image/png;base64,aGVsbG8=');
}

function testOpenAiRequestIncludesReasoningEffortByDefault() {
    const { registry } = createRuntime();
    const request = registry.getProvider('openai').buildRequest({
        messages: [{ role: 'user', content: 'Explain briefly' }],
        config: {
            provider: 'openai',
            apiKey: 'sk-openai-key-1234567890',
            model: 'gpt-5-mini',
            baseUrl: 'https://api.openai.com/v1'
        }
    });

    assert.deepStrictEqual(request.body.reasoning, { effort: 'medium' });
}

function testOpenAiRequestOmitsReasoningWhenEffortOff() {
    const { registry } = createRuntime();
    const request = registry.getProvider('openai').buildRequest({
        messages: [{ role: 'user', content: 'Explain briefly' }],
        config: {
            provider: 'openai',
            apiKey: 'sk-openai-key-1234567890',
            model: 'gpt-5-mini',
            baseUrl: 'https://api.openai.com/v1',
            reasoningEffort: 'off'
        }
    });

    assert.strictEqual(Object.prototype.hasOwnProperty.call(request.body, 'reasoning'), false);
}

function testOpenAiGpt5FamilyOmitsCustomTemperature() {
    const { registry } = createRuntime();
    const request = registry.getProvider('openai').buildRequest({
        messages: [{ role: 'user', content: 'Hello' }],
        config: {
            provider: 'openai',
            apiKey: 'sk-openai-key-1234567890',
            model: 'gpt-5.1-chat-latest',
            baseUrl: 'https://api.openai.com/v1',
            temperature: 0.1
        }
    });

    assert.strictEqual(Object.prototype.hasOwnProperty.call(request.body, 'temperature'), false);
}

function testOpenAiGpt52ProUsesResponsesApi() {
    const { registry } = createRuntime();
    const request = registry.getProvider('openai').buildRequest({
        messages: [{ role: 'user', content: 'Hello' }],
        config: {
            provider: 'openai',
            apiKey: 'sk-openai-key-1234567890',
            model: 'gpt-5.2-pro',
            baseUrl: 'https://api.openai.com/v1',
            temperature: 0.1
        }
    });

    assert.strictEqual(request.url, 'https://api.openai.com/v1/responses');
    assert.strictEqual(Array.isArray(request.body.input), true);
    assert.strictEqual(request.body.input[0].content, 'Hello');
    assert.strictEqual(request.capabilities.supportsStreaming, false);
}

function testOpenAiCodexModelIsRejectedForChatCompletions() {
    const { registry } = createRuntime();
    assert.throws(() => {
        registry.getProvider('openai').buildRequest({
            messages: [{ role: 'user', content: 'Hello' }],
            config: {
                provider: 'openai',
                apiKey: 'sk-openai-key-1234567890',
                model: 'gpt-5.1-codex',
                baseUrl: 'https://api.openai.com/v1'
            }
        });
    }, /Codex models are not supported/i);
}

async function testOpenAiGpt52ProNormalizesResponsesApiOutput() {
    const { registry } = createRuntime({
        request(options) {
            assert.strictEqual(options.url, 'https://api.openai.com/v1/responses');
            return createJsonResponse({
                output_text: 'Responses API ok'
            });
        }
    });

    const result = await registry.getProvider('openai').send({
        messages: [{ role: 'user', content: 'Hello' }],
        config: {
            provider: 'openai',
            apiKey: 'sk-openai-key-1234567890',
            model: 'gpt-5.2-pro',
            baseUrl: 'https://api.openai.com/v1'
        }
    });

    assert.strictEqual(result.data[0].content, 'Responses API ok');
}

async function testOpenAiReasoningMetadataIsNormalized() {
    const { registry } = createRuntime({
        request() {
            return createJsonResponse({
                choices: [{
                    message: {
                        role: 'assistant',
                        content: 'Answer ready',
                        reasoning: 'I compared two likely interpretations and selected the one matching the user request.'
                    }
                }],
                usage: {
                    completion_tokens_details: {
                        reasoning_tokens: 42
                    }
                }
            });
        }
    });

    const result = await registry.getProvider('openai').send({
        messages: [{ role: 'user', content: 'Explain briefly' }],
        config: {
            provider: 'openai',
            apiKey: 'sk-openai-key-1234567890',
            model: 'gpt-5-mini',
            baseUrl: 'https://api.openai.com/v1',
            reasoningEffort: 'high'
        }
    });

    assert.strictEqual(result.data[0].content, 'Answer ready');
    assert.strictEqual(result.reasoning.available, true);
    assert.ok(result.reasoning.summary.indexOf('I compared two likely interpretations') === 0);
    assert.strictEqual(result.reasoning.tokens, 42);
    assert.strictEqual(result.reasoning.effort, 'high');
    assert.strictEqual(result.reasoning.source, 'openai');
}

async function testMissingReasoningDoesNotBreakNormalization() {
    const { registry } = createRuntime({
        request() {
            return createJsonResponse({
                choices: [{
                    message: {
                        role: 'assistant',
                        content: 'No reasoning data'
                    }
                }]
            });
        }
    });

    const result = await registry.getProvider('openrouter').send({
        messages: [{ role: 'user', content: 'Ping' }],
        config: {
            provider: 'openrouter',
            apiKey: 'sk-or-v1-default-openrouter-1234567890',
            model: 'openrouter/auto',
            baseUrl: 'https://openrouter.ai/api/v1'
        }
    });

    assert.strictEqual(result.data[0].content, 'No reasoning data');
    assert.strictEqual(result.reasoning.available, false);
    assert.strictEqual(result.reasoning.summary, '');
    assert.strictEqual(result.reasoning.source, 'openrouter');
}

function testAnthropicPayloadUsesBase64ImageBlocks() {
    const { registry } = createRuntime();
    const request = registry.getProvider('anthropic').buildRequest({
        messages: [{
            role: 'user',
            content: 'Describe this',
            attachments: [createImageAttachment()]
        }],
        config: {
            provider: 'anthropic',
            apiKey: 'anthropic-key-1234567890',
            model: 'claude-sonnet-4-5',
            baseUrl: 'https://api.anthropic.com/v1'
        }
    });

    const imagePart = request.body.messages[0].content.find((item) => item.type === 'image');
    assert.ok(imagePart);
    assert.strictEqual(imagePart.source.type, 'base64');
    assert.strictEqual(imagePart.source.media_type, 'image/png');
}

function testGeminiPayloadUsesInlineData() {
    const { registry } = createRuntime();
    const request = registry.getProvider('gemini').buildRequest({
        messages: [{
            role: 'user',
            content: 'Classify this screenshot',
            attachments: [createImageAttachment()]
        }],
        config: {
            provider: 'gemini',
            apiKey: 'gemini-key-1234567890',
            model: 'gemini-2.5-flash'
        }
    });

    const inlineDataPart = request.body.contents[0].parts.find((item) => item.inline_data);
    assert.ok(inlineDataPart);
    assert.strictEqual(inlineDataPart.inline_data.mime_type, 'image/png');
}

function testYandexModelUsesFolderUri() {
    const { registry } = createRuntime();
    const request = registry.getProvider('yandex').buildRequest({
        messages: [{ role: 'user', content: 'Hello' }],
        config: {
            provider: 'yandex',
            apiKey: 'yandex-key-1234567890',
            folderId: 'folder-123',
            model: 'yandexgpt/latest'
        }
    });

    assert.strictEqual(request.body.model, 'gpt://folder-123/yandexgpt/latest');
}

async function testGigachatFlowUploadsAndCleansUpFiles() {
    const { registry, requestLog } = createRuntime({
        request(options) {
            if (/oauth$/i.test(options.url)) {
                return createJsonResponse({
                    access_token: 'gigachat-token',
                    expires_at: Date.now() + 600000
                });
            }
            if (/\/files$/i.test(options.url) && options.method === 'POST') {
                assert.ok(typeof FormData !== 'undefined' && options.body instanceof FormData);
                return createJsonResponse({ id: 'file-1' });
            }
            if (/chat\/completions$/i.test(options.url)) {
                const payload = JSON.parse(options.body);
                assert.deepStrictEqual(payload.messages[0].attachments, ['file-1']);
                return createJsonResponse({
                    choices: [{
                        message: {
                            role: 'assistant',
                            content: 'ready'
                        }
                    }]
                });
            }
            if (/\/files\/file-1$/i.test(options.url) && options.method === 'DELETE') {
                return createJsonResponse({});
            }
            throw new Error('Unexpected request: ' + options.method + ' ' + options.url);
        }
    });

    const result = await registry.getProvider('gigachat').send({
        messages: [{
            role: 'user',
            content: 'Describe image',
            attachments: [createImageAttachment()]
        }],
        config: {
            provider: 'gigachat',
            authKey: 'gigachat-auth-1234567890',
            scope: 'GIGACHAT_API_PERS',
            model: 'GigaChat-Pro'
        }
    });

    assert.strictEqual(result.data[0].content, 'ready');
    assert.ok(requestLog.some((item) => /\/files$/i.test(item.url) && item.method === 'POST'));
    assert.ok(requestLog.some((item) => /\/files\/file-1$/i.test(item.url) && item.method === 'DELETE'));
}

function testDeepSeekRejectsImageInput() {
    const { registry } = createRuntime();
    assert.throws(() => {
        registry.getProvider('deepseek').buildRequest({
            messages: [{
                role: 'user',
                content: 'Inspect this',
                attachments: [createImageAttachment()]
            }],
            config: {
                provider: 'deepseek',
                apiKey: 'deepseek-key-1234567890',
                model: 'deepseek-chat'
            }
        });
    }, /text-only input/i);
}

async function testOpenRouterImageGenerationUsesChatCompletions() {
    const { registry, requestLog } = createRuntime({
        request(options) {
            if (/chat\/completions$/i.test(options.url)) {
                const payload = JSON.parse(options.body);
                assert.strictEqual(payload.model, 'google/gemini-2.5-flash-image');
                assert.deepStrictEqual(payload.modalities, ['image', 'text']);
                assert.strictEqual(payload.messages[0].role, 'user');
                assert.strictEqual(payload.messages[0].content, 'Draw a lighthouse at sunset');
                assert.strictEqual(payload.image_config.aspect_ratio, '1:1');
                assert.strictEqual(payload.image_config.image_size, '1K');
                return createJsonResponse({
                    choices: [{
                        message: {
                            role: 'assistant',
                            content: 'done',
                            images: [{
                                image_url: {
                                    url: 'data:image/png;base64,abc123'
                                }
                            }]
                        }
                    }]
                });
            }
            throw new Error('Unexpected request: ' + options.method + ' ' + options.url);
        }
    });

    const result = await registry.getProvider('openrouter').imageRequest('Draw a lighthouse at sunset', {
        provider: 'openrouter',
        apiKey: 'sk-or-v1-default-openrouter-1234567890',
        baseUrl: 'https://openrouter.ai/api/v1',
        size: '1024x1024'
    });

    assert.strictEqual(result.data[0].url, 'data:image/png;base64,abc123');
    assert.ok(requestLog.some((item) => /chat\/completions$/i.test(item.url)));
    assert.ok(requestLog.every((item) => !/images\/generations$/i.test(item.url)));
}

async function testOpenRouterChatPrefersFetchTransport() {
    const { registry, requestLog } = createRuntime({
        request(options) {
            assert.strictEqual(options.preferFetch, true);
            assert.strictEqual(options.retryViaFetchOnJsonParseFailure, true);
            return createJsonResponse({
                choices: [{
                    message: {
                        role: 'assistant',
                        content: 'transport ok'
                    }
                }]
            });
        }
    });

    const result = await registry.getProvider('openrouter').send({
        messages: [{ role: 'user', content: 'Ping' }],
        config: {
            provider: 'openrouter',
            apiKey: 'sk-or-v1-default-openrouter-1234567890',
            model: 'openrouter/auto',
            baseUrl: 'https://openrouter.ai/api/v1'
        }
    });

    assert.strictEqual(result.data[0].content, 'transport ok');
    assert.strictEqual(requestLog.length, 1);
}

async function run() {
    testRegistryCapabilities();
    testOpenAiPayloadUsesImageUrl();
    testOpenAiRequestIncludesReasoningEffortByDefault();
    testOpenAiRequestOmitsReasoningWhenEffortOff();
    testOpenAiGpt5FamilyOmitsCustomTemperature();
    testOpenAiGpt52ProUsesResponsesApi();
    testOpenAiCodexModelIsRejectedForChatCompletions();
    await testOpenAiGpt52ProNormalizesResponsesApiOutput();
    await testOpenAiReasoningMetadataIsNormalized();
    await testMissingReasoningDoesNotBreakNormalization();
    testAnthropicPayloadUsesBase64ImageBlocks();
    testGeminiPayloadUsesInlineData();
    testYandexModelUsesFolderUri();
    await testGigachatFlowUploadsAndCleansUpFiles();
    testDeepSeekRejectsImageInput();
    await testOpenRouterImageGenerationUsesChatCompletions();
    await testOpenRouterChatPrefersFetchTransport();
    console.log('openrouter_client.test.js: ok');
}

run().catch((error) => {
    console.error(error);
    process.exit(1);
});
