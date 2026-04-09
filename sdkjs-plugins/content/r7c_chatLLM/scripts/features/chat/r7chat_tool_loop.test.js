const assert = require('assert');
const toolLoop = require('./r7chat_tool_loop');

function createLoop(overrides) {
    const options = overrides || {};
    return toolLoop.createToolLoop({
        openrouter: options.openrouter,
        webTools: options.webTools,
        hostTools: options.hostTools
    });
}

async function testResponseWithoutTools() {
    const loop = createLoop({
        openrouter: {
            async chatRequest(messages, systemMessage, stream) {
                assert.strictEqual(stream, false);
                return { data: [{ content: 'plain response' }] };
            },
            async chatRequestRaw() {
                throw new Error('raw path should not be used');
            }
        },
        webTools: {
            getToolDefinitions() { return []; },
            async executeToolCall() { throw new Error('unused'); }
        }
    });

    const result = await loop.run({
        messages: [{ role: 'user', content: 'hello' }],
        systemMessage: 'system',
        config: { model: 'test' }
    });

    assert.strictEqual(result.response, 'plain response');
    assert.strictEqual(result.usedTools, false);
}

async function testSingleToolCall() {
    let toolCalls = 0;
    const loop = createLoop({
        openrouter: {
            async chatRequest() {
                throw new Error('plain fallback should not be used');
            },
            async chatRequestRaw(messages) {
                if (messages.length === 1) {
                    return {
                        choices: [{
                            message: {
                                role: 'assistant',
                                content: '',
                                tool_calls: [{
                                    id: 'call_1',
                                    function: {
                                        name: 'web_search',
                                        arguments: JSON.stringify({ query: 'latest news' })
                                    }
                                }]
                            }
                        }]
                    };
                }
                return {
                    choices: [{
                        message: {
                            role: 'assistant',
                            content: 'tool-based answer'
                        }
                    }]
                };
            }
        },
        webTools: {
            getToolDefinitions() {
                return [{ type: 'function', function: { name: 'web_search' } }];
            },
            async executeToolCall(name, rawArgs) {
                toolCalls += 1;
                assert.strictEqual(name, 'web_search');
                assert.deepStrictEqual(JSON.parse(rawArgs), { query: 'latest news' });
                return { provider: 'exa', results: [{ url: 'https://example.com' }], errors: [], statuses: [], sources: [] };
            }
        }
    });

    const result = await loop.run({
        messages: [{ role: 'user', content: 'search the web' }],
        systemMessage: 'system',
        config: { model: 'test' }
    });

    assert.strictEqual(toolCalls, 1);
    assert.strictEqual(result.response, 'tool-based answer');
    assert.strictEqual(result.usedTools, true);
}

async function testMultipleToolCalls() {
    let count = 0;
    const loop = createLoop({
        openrouter: {
            async chatRequest() {
                throw new Error('plain fallback should not be used');
            },
            async chatRequestRaw(messages) {
                if (messages.length === 1) {
                    return {
                        choices: [{
                            message: {
                                role: 'assistant',
                                content: '',
                                tool_calls: [
                                    { id: 'call_1', function: { name: 'web_search', arguments: JSON.stringify({ query: 'one' }) } },
                                    { id: 'call_2', function: { name: 'web_crawling', arguments: JSON.stringify({ urls: ['https://example.com'] }) } }
                                ]
                            }
                        }]
                    };
                }
                return {
                    choices: [{
                        message: {
                            role: 'assistant',
                            content: 'combined answer'
                        }
                    }]
                };
            }
        },
        webTools: {
            getToolDefinitions() {
                return [
                    { type: 'function', function: { name: 'web_search' } },
                    { type: 'function', function: { name: 'web_crawling' } }
                ];
            },
            async executeToolCall() {
                count += 1;
                return { provider: 'brave', results: [], errors: [], statuses: [], sources: [] };
            }
        }
    });

    const result = await loop.run({
        messages: [{ role: 'user', content: 'search and crawl' }],
        systemMessage: 'system',
        config: { model: 'test' }
    });

    assert.strictEqual(count, 2);
    assert.strictEqual(result.response, 'combined answer');
}

async function testMaxIterationStop() {
    const loop = createLoop({
        openrouter: {
            async chatRequest() {
                throw new Error('plain fallback should not be used');
            },
            async chatRequestRaw() {
                return {
                    choices: [{
                        message: {
                            role: 'assistant',
                            content: '',
                            tool_calls: [{
                                id: 'call_1',
                                function: {
                                    name: 'web_search',
                                    arguments: JSON.stringify({ query: 'repeat' })
                                }
                            }]
                        }
                    }]
                };
            }
        },
        webTools: {
            getToolDefinitions() {
                return [{ type: 'function', function: { name: 'web_search' } }];
            },
            async executeToolCall() {
                return { provider: 'exa', results: [], errors: [], statuses: [], sources: [] };
            }
        }
    });

    const result = await loop.run({
        messages: [{ role: 'user', content: 'loop forever' }],
        systemMessage: 'system',
        config: { model: 'test' },
        maxIterations: 2
    });

    assert.strictEqual(result.stopped, true);
    assert.ok(/max iterations/i.test(result.response));
}

async function testFallbackWhenModelRejectsTools() {
    let fallbackCalled = false;
    const loop = createLoop({
        openrouter: {
            async chatRequest(messages, systemMessage, stream, config) {
                fallbackCalled = true;
                assert.strictEqual(stream, false);
                assert.strictEqual(config.tools, undefined);
                return { data: [{ content: 'fallback answer' }] };
            },
            async chatRequestRaw() {
                throw new Error('Model does not support tool calling');
            }
        },
        webTools: {
            getToolDefinitions() {
                return [{ type: 'function', function: { name: 'web_search' } }];
            },
            async executeToolCall() {
                throw new Error('unused');
            }
        }
    });

    const result = await loop.run({
        messages: [{ role: 'user', content: 'search' }],
        systemMessage: 'system',
        config: { model: 'test' }
    });

    assert.strictEqual(fallbackCalled, true);
    assert.strictEqual(result.response, 'fallback answer');
    assert.ok(/retried without web tools/i.test(result.warning));
}

async function testHostToolCall() {
    let hostCalls = 0;
    const loop = createLoop({
        openrouter: {
            async chatRequest() {
                throw new Error('plain fallback should not be used');
            },
            async chatRequestRaw(messages) {
                if (messages.length === 1) {
                    return {
                        choices: [{
                            message: {
                                role: 'assistant',
                                content: '',
                                tool_calls: [{
                                    id: 'call_1',
                                    function: {
                                        name: 'sheet_append_report',
                                        arguments: JSON.stringify({ sheetName: 'Sheet2', anchor: 'bottom' })
                                    }
                                }]
                            }
                        }]
                    };
                }
                return {
                    choices: [{
                        message: {
                            role: 'assistant',
                            content: 'host-tool answer'
                        }
                    }]
                };
            }
        },
        hostTools: {
            isEnabled() { return true; },
            getToolDefinitions() {
                return [{
                    type: 'function',
                    function: {
                        name: 'sheet_append_report',
                        description: 'Append report to sheet',
                        parameters: { type: 'object', properties: { sheetName: { type: 'string' } }, required: ['sheetName'] }
                    }
                }];
            },
            hasTool(name) {
                return name === 'sheet_append_report';
            },
            async executeToolCall(name, rawArgs) {
                hostCalls += 1;
                assert.strictEqual(name, 'sheet_append_report');
                assert.deepStrictEqual(JSON.parse(rawArgs), { sheetName: 'Sheet2', anchor: 'bottom' });
                return { provider: 'desktop-editor', tool: name, ok: true, output: { inserted: true } };
            }
        },
        webTools: {
            getToolDefinitions() { return []; },
            async executeToolCall() { throw new Error('unused'); }
        }
    });

    const result = await loop.run({
        messages: [{ role: 'user', content: 'make a report' }],
        systemMessage: 'system',
        config: { model: 'test' },
        runtimeSettings: {}
    });

    assert.strictEqual(hostCalls, 1);
    assert.strictEqual(result.response, 'host-tool answer');
    assert.strictEqual(result.usedTools, true);
}

async function run() {
    await testResponseWithoutTools();
    await testSingleToolCall();
    await testMultipleToolCalls();
    await testMaxIterationStop();
    await testFallbackWhenModelRejectsTools();
    await testHostToolCall();
    console.log('r7chat_tool_loop.test.js: ok');
}

run().catch((error) => {
    console.error(error);
    process.exit(1);
});
