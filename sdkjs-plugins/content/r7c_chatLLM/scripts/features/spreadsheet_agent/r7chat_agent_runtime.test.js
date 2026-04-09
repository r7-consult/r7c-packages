const assert = require('assert');

function loadFresh(modulePath) {
    delete require.cache[require.resolve(modulePath)];
    return require(modulePath);
}

function loadPlannerProfiles() {
    global.window = global;
    loadFresh('../../agent/r7chat_planner_profiles.js');
}

function createAgentRuntime() {
    global.R7Chat = {
        runtime: {
            constants: {
                defaultModel: 'openrouter/auto',
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
            },
            state: {}
        },
        platform: {},
        features: {},
        services: {},
        agent: {}
    };

    return loadFresh('./r7chat_agent_runtime');
}

function testPlannerParsing() {
    const runtime = createAgentRuntime();

    const parsed = runtime.parsePlannerResponse('```json\n{"step":{"type":"collect_context","reason":"Inspect workbook context","args":{"mode":"discover","sample":true}}}\n```');
    assert.strictEqual(parsed.step.type, 'collect_context');
    assert.deepStrictEqual(parsed.step.args, { mode: 'discover', sample: true });

    const normalized = runtime.normalizePlannerStep(parsed, 0);
    assert.deepStrictEqual(normalized, {
        id: 'step_1',
        type: 'collect_context',
        reason: 'Inspect workbook context',
        args: { mode: 'discover', sample: true },
        macro_code: undefined
    });
}

function testPlannerWebStepParsing() {
    const runtime = createAgentRuntime();
    const normalized = runtime.normalizePlannerStep({
        type: 'web_search',
        reason: 'Need current web information',
        args: { query: 'latest exa api changes' }
    }, 0);

    assert.deepStrictEqual(normalized, {
        id: 'step_1',
        type: 'web_search',
        reason: 'Need current web information',
        args: { query: 'latest exa api changes' },
        macro_code: undefined
    });
}

function testPlannerHostToolStepParsing() {
    const runtime = createAgentRuntime();
    const normalized = runtime.normalizePlannerStep({
        type: 'call_host_tool',
        reason: 'Use native word tool',
        args: {
            tool: 'document_insert_summary',
            input: {
                topic: 'Quarterly review'
            }
        }
    }, 0);

    assert.deepStrictEqual(normalized, {
        id: 'step_1',
        type: 'call_host_tool',
        reason: 'Use native word tool',
        args: {
            tool: 'document_insert_summary',
            input: {
                topic: 'Quarterly review'
            }
        },
        macro_code: undefined
    });
}

function testPlannerCollectContextStepParsing() {
    const runtime = createAgentRuntime();
    const normalized = runtime.normalizePlannerStep({
        type: 'collect_context',
        reason: 'Read selected sheet context',
        args: {
            mode: 'collect',
            sheetName: 'Sales'
        }
    }, 0);

    assert.deepStrictEqual(normalized, {
        id: 'step_1',
        type: 'collect_context',
        reason: 'Read selected sheet context',
        args: {
            mode: 'collect',
            sheetName: 'Sales'
        },
        macro_code: undefined
    });
}

function testPlannerReadDocumentSnapshotStepParsing() {
    const runtime = createAgentRuntime();
    const normalized = runtime.normalizePlannerStep({
        type: 'read_document_snapshot',
        reason: 'Capture a robust Word snapshot before reasoning.',
        args: {
            maxChars: 9000
        }
    }, 0);

    assert.deepStrictEqual(normalized, {
        id: 'step_1',
        type: 'read_document_snapshot',
        reason: 'Capture a robust Word snapshot before reasoning.',
        args: {
            maxChars: 9000
        },
        macro_code: undefined
    });
}

function testPlannerWordPlanStepParsing() {
    const runtime = createAgentRuntime();
    const normalized = runtime.normalizePlannerStep({
        type: 'present_plan',
        reason: 'Show plan before writing',
        args: {
            plan: {
                title: 'AI in education',
                sections: [{
                    title: 'Intro',
                    purpose: 'Set the context'
                }]
            }
        }
    }, 0);

    assert.deepStrictEqual(normalized, {
        id: 'step_1',
        type: 'present_plan',
        reason: 'Show plan before writing',
        args: {
            plan: {
                title: 'AI in education',
                sections: [{
                    title: 'Intro',
                    purpose: 'Set the context'
                }]
            }
        },
        macro_code: undefined
    });
}

function testPlannerImageStepParsing() {
    const runtime = createAgentRuntime();
    const normalized = runtime.normalizePlannerStep({
        type: 'generate_image_asset',
        reason: 'Generate the hero image',
        args: {
            prompt: 'Editorial illustration of AI in a classroom',
            caption: 'AI in education'
        }
    }, 0);

    assert.deepStrictEqual(normalized, {
        id: 'step_1',
        type: 'generate_image_asset',
        reason: 'Generate the hero image',
        args: {
            prompt: 'Editorial illustration of AI in a classroom',
            caption: 'AI in education'
        },
        macro_code: undefined
    });
}

function testPlannerRunMacroStepReadsMacroCodeFromArgs() {
    const runtime = createAgentRuntime();
    const normalized = runtime.normalizePlannerStep({
        type: 'run_macro_code',
        reason: 'Write a macro',
        args: {
            macro_code: 'return 42;'
        }
    }, 0);

    assert.deepStrictEqual(normalized, {
        id: 'step_1',
        type: 'run_macro_code',
        reason: 'Write a macro',
        args: { macro_code: 'return 42;' },
        macro_code: 'return 42;'
    });
}

function testPlannerErrorsAndMethodPack() {
    const runtime = createAgentRuntime();

    assert.throws(() => {
        runtime.parsePlannerResponse('not json');
    }, /Planner response is not valid JSON/);

    assert.throws(() => {
        runtime.normalizePlannerStep({ type: 'unsupported_step' }, 2);
    }, /Unsupported step type/);

    const pack = runtime.makeLimitedMethodPack(
        Array.from({ length: 30 }, (_, index) => ({ method: 'Method' + index })),
        5
    );

    assert.strictEqual(pack.total_count, 30);
    assert.strictEqual(pack.returned_count, 5);
    assert.strictEqual(pack.truncated, true);
    assert.strictEqual(pack.methods[0].method, 'Method0');
    assert.strictEqual(pack.methods[4].method, 'Method4');
}

function testApiReferenceFuzzyLookupAndRecoveryStep() {
    const runtime = createAgentRuntime();
    global.R7_API_REFERENCE_CATALOG = {
        categories: {
            sheet_workbook: {
                title: 'Sheet and Workbook',
                count: 1,
                methods: [{
                    object: 'ApiInterface',
                    method: 'AddSheet',
                    args: 'sName',
                    description: 'Creates a new worksheet.'
                }]
            }
        },
        objects: {
            ApiInterface: [{
                object: 'ApiInterface',
                method: 'AddSheet',
                args: 'sName',
                description: 'Creates a new worksheet.'
            }]
        }
    };

    const hits = runtime.findApiReferenceMatches(global.R7_API_REFERENCE_CATALOG, {
        method: 'CreateSheet'
    });
    assert.strictEqual(hits[0].method, 'AddSheet');

    const recoveryStep = runtime.buildForcedApiReferenceRecoveryStep({
        error: 'Api.CreateSheet is not a function'
    }, 2);
    assert.deepStrictEqual(recoveryStep, {
        id: 'step_3',
        type: 'analyze_reference_macros',
        reason: 'Previous macro used unverified method "CreateSheet". Search the local API catalog for verified alternatives before retrying.',
        args: { method: 'CreateSheet', limit: 12 }
    });
}

function testResearchPolicyUnderstandsConfiguredBrave() {
    const runtime = createAgentRuntime();
    global.R7Chat.platform.webTools = {
        isEnabled() {
            return true;
        },
        getCurrentProvider() {
            return {
                provider: 'brave',
                apiKey: 'brave-key'
            };
        }
    };
    global.R7Chat.services.chat = {
        loadRuntimeSettings() {
            return {
                model: 'openrouter/auto',
                webTools: {
                    provider: 'brave',
                    providers: {
                        brave: { apiKey: 'brave-key' },
                        exa: { apiKey: '' }
                    }
                }
            };
        }
    };

    const policy = runtime.buildResearchPolicyForRequest('Напиши диплом на тему квантовых вычислений');
    assert.strictEqual(policy.webSearchEnabled, true);
    assert.strictEqual(policy.activeProvider, 'brave');
    assert.strictEqual(policy.braveApiConfigured, true);
    assert.strictEqual(policy.required, true);
    assert.ok(/квантовых вычислений/i.test(policy.query));
}

function testWordPlanModeHeuristicAndForcedResearch() {
    const runtime = createAgentRuntime();
    global.R7Chat.platform.hostBridge = {
        getEditorTypeSafe() {
            return 'word';
        }
    };
    global.R7Chat.platform.webTools = {
        isEnabled() {
            return true;
        },
        getCurrentProvider() {
            return {
                provider: 'brave',
                apiKey: 'brave-key'
            };
        }
    };
    global.R7Chat.services.chat = {
        loadRuntimeSettings() {
            return {
                webTools: {
                    provider: 'brave',
                    providers: {
                        brave: { apiKey: 'brave-key' },
                        exa: { apiKey: '' }
                    }
                }
            };
        }
    };

    assert.strictEqual(runtime.isWordPlanModeRequest('Подготовь статью с картинками и фактами из интернета про ИИ в образовании'), true);
    const policy = runtime.buildResearchPolicyForRequest('Подготовь статью с картинками и фактами из интернета про ИИ в образовании', {
        forceRequired: true
    });
    assert.strictEqual(policy.required, true);
    assert.ok(/ИИ в образовании/i.test(policy.query));
    assert.ok(!/картинк|план|по разделам/i.test(policy.query));
}

function testWebSearchRecoveryHelpers() {
    const runtime = createAgentRuntime();
    global.R7Chat.platform.hostBridge = {
        getEditorTypeSafe() {
            return 'word';
        }
    };
    global.R7Chat.platform.webTools = {
        isEnabled() {
            return true;
        },
        getCurrentProvider() {
            return {
                provider: 'brave',
                apiKey: 'brave-key'
            };
        }
    };
    global.R7Chat.services.chat = {
        loadRuntimeSettings() {
            return {
                webTools: {
                    provider: 'brave',
                    providers: {
                        brave: { apiKey: 'brave-key' },
                        exa: { apiKey: '' }
                    }
                }
            };
        }
    };
    runtime.getState().lastUserMessage = 'Подготовь статью про влияние ИИ на образование. Нужен план, факты из интернета и 3 картинки по разделам';
    runtime.getState().researchPolicy = runtime.buildResearchPolicyForRequest(runtime.getState().lastUserMessage, {
        forceRequired: true
    });

    const hint = runtime.buildWebSearchRecoveryPlannerHint({
        args: {
            query: 'про влияние ИИ на образование. Нужен план, факты из интернета и 3 картинки по разделам'
        }
    }, {
        error: 'search failed'
    });
    assert.ok(/^RECOVERY_HINT:/i.test(hint));
    assert.ok(hint.length > 40);

    const recoveryStep = runtime.buildForcedWebSearchRecoveryStep({
        args: {
            query: 'про влияние ИИ на образование. Нужен план, факты из интернета и 3 картинки по разделам'
        }
    }, 3);
    assert.deepStrictEqual(recoveryStep, {
        id: 'step_4',
        type: 'web_search',
        reason: 'Retry web research with a shorter topic-only query after the previous search failed.',
        args: {
            query: 'влияние ИИ на образование'
        }
    });
}

function testForcedInitialStepAddsWebSearchForWordResearch() {
    const runtime = createAgentRuntime();
    global.R7Chat.platform.hostBridge = {
        getEditorTypeSafe() {
            return 'word';
        }
    };
    global.R7Chat.platform.webTools = {
        isEnabled() {
            return true;
        },
        getCurrentProvider() {
            return {
                provider: 'brave',
                apiKey: 'brave-key'
            };
        }
    };
    global.R7Chat.services.chat = {
        loadRuntimeSettings() {
            return {
                webTools: {
                    provider: 'brave',
                    providers: {
                        brave: { apiKey: 'brave-key' },
                        exa: { apiKey: '' }
                    }
                }
            };
        }
    };

    runtime.getState().researchPolicy = runtime.buildResearchPolicyForRequest('Напиши диплом на тему квантовых вычислений');

    const firstStep = runtime.getForcedInitialStep('Напиши диплом на тему квантовых вычислений', 0);
    const secondStep = runtime.getForcedInitialStep('Напиши диплом на тему квантовых вычислений', 1);

    assert.strictEqual(firstStep.type, 'read_document_snapshot');
    assert.strictEqual(firstStep.args.editorType, 'word');
    assert.strictEqual(secondStep.type, 'web_search');
    assert.ok(/квантовых вычислений/i.test(secondStep.args.query));
}

function testPlannerSystemMessageIncludesWebSearchStatus() {
    const runtime = createAgentRuntime();
    runtime.getState().researchPolicy = {
        required: true,
        query: 'квантовые вычисления',
        completed: false,
        attempted: false,
        failed: false,
        webSearchEnabled: true,
        activeProvider: 'brave',
        braveApiConfigured: true,
        exaApiConfigured: false
    };

    const message = runtime.getPlannerSystemMessage({ mode: 'default' });
    assert.ok(/WEB_SEARCH_STATUS:/i.test(message));
    assert.ok(/active_provider=brave/i.test(message));
    assert.ok(/brave_api_configured=true/i.test(message));
    assert.ok(/RESEARCH_FIRST_POLICY:/i.test(message));
}

function testPlannerSystemMessageMentionsPluginAndActiveOfficeEnvironment() {
    const runtime = createAgentRuntime();
    loadPlannerProfiles();

    global.R7Chat.platform.hostBridge = {
        getEditorTypeSafe() {
            return 'cell';
        }
    };

    const spreadsheetMessage = runtime.getPlannerSystemMessage({ mode: 'default' });
    assert.ok(/\{r7c\}\.ChatLLM plugin/i.test(spreadsheetMessage));
    assert.ok(/R7 Office spreadsheet editor/i.test(spreadsheetMessage));
    assert.ok(/read_active_sheet/i.test(spreadsheetMessage));
    assert.ok(/read_sheet_range/i.test(spreadsheetMessage));
    assert.ok(/predefined verified ONLYOFFICE read macros\/commands/i.test(spreadsheetMessage));
    assert.ok(/Do NOT use run_macro_code just to inspect spreadsheet contents/i.test(spreadsheetMessage));

    global.R7Chat.platform.hostBridge = {
        getEditorTypeSafe() {
            return 'word';
        }
    };

    const wordMessage = runtime.getPlannerSystemMessage({ mode: 'default' });
    assert.ok(/\{r7c\}\.ChatLLM plugin/i.test(wordMessage));
    assert.ok(/R7 Office word processor/i.test(wordMessage));
}

function testSpreadsheetReadIntentForcesSpreadsheetReadStepBeforeAnswer() {
    const runtime = createAgentRuntime();
    global.R7Chat.platform.hostBridge = {
        getEditorTypeSafe() {
            return 'cell';
        }
    };

    runtime.getState().lastUserMessage = 'Что на листе и какие данные в таблице?';
    assert.strictEqual(runtime.isSpreadsheetReadIntent(runtime.getState().lastUserMessage), true);
    assert.strictEqual(runtime.hasExecutedSpreadsheetReadStep(runtime.getState().lastUserMessage), false);
    assert.strictEqual(runtime.shouldForceSpreadsheetReadBeforeAnswer({ type: 'final_answer' }), true);

    const forcedStep = runtime.buildForcedSpreadsheetReadStep(runtime.getState().lastUserMessage, 0, 'read_before_answer');
    assert.strictEqual(forcedStep.type, 'read_active_sheet');
    assert.strictEqual(forcedStep.args.editorType, 'cell');

    runtime.getState().steps = [
        { id: 'step_1', type: 'read_active_sheet', args: { editorType: 'cell' } }
    ];
    assert.strictEqual(runtime.hasExecutedSpreadsheetReadStep(runtime.getState().lastUserMessage), true);
    assert.strictEqual(runtime.shouldForceSpreadsheetReadBeforeAnswer({ type: 'final_answer' }), false);
}

function testSpreadsheetMacroWordingStillForcesRecordedReadStep() {
    const runtime = createAgentRuntime();
    global.R7Chat.platform.hostBridge = {
        getEditorTypeSafe() {
            return 'cell';
        }
    };

    const firstStep = runtime.getForcedInitialStep('что на листе проверь через макросы', 0);
    assert.strictEqual(firstStep.type, 'read_active_sheet');

    runtime.getState().lastUserMessage = 'что на листе проверь через макросы';
    assert.strictEqual(runtime.shouldBlockGeneratedSpreadsheetInspectionMacro({ type: 'run_macro_code' }), true);

    runtime.getState().steps = [
        { id: 'step_1', type: 'read_active_sheet', args: { editorType: 'cell' } }
    ];
    assert.strictEqual(runtime.shouldBlockGeneratedSpreadsheetInspectionMacro({ type: 'run_macro_code' }), false);
}

function testDesktopToolsPlannerStateForWordAutoMode() {
    const runtime = createAgentRuntime();
    loadPlannerProfiles();

    global.R7Chat.platform.hostBridge = {
        getEditorTypeSafe() {
            return 'word';
        }
    };
    global.R7Chat.platform.desktopTools = {
        getStatus() {
            return {
                desktopEditorAvailable: true,
                catalogAvailable: true,
                executionAvailable: true,
                catalogParseError: '',
                toolCount: 2,
                catalogHash: 'dtb_demo'
            };
        },
        getCatalogForPrompt() {
            return '- document_insert_summary | Insert a summary block | required=topic';
        }
    };
    global.R7Chat.services.chat = {
        loadRuntimeSettings() {
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
    };

    const desktopToolsState = runtime.buildDesktopToolsPlannerState('Подготовь краткий summary документа');
    assert.strictEqual(desktopToolsState.enabledForPlanner, true);
    assert.strictEqual(desktopToolsState.reason, 'auto_mode');

    const message = runtime.getPlannerSystemMessage({
        mode: 'default',
        desktopTools: desktopToolsState
    });

    assert.ok(/read_document_snapshot/i.test(message));
    assert.ok(/HOST_TOOLS_POLICY:/i.test(message));
    assert.ok(/list_host_tools/i.test(message));
    assert.ok(/call_host_tool/i.test(message));
    assert.ok(/document_insert_summary/i.test(message));
    assert.ok(/Word document context/i.test(message) || /snapshot/i.test(message));
}

function testForcedInitialStepUsesHostToolDiscoveryForSpreadsheetWrite() {
    const runtime = createAgentRuntime();
    loadPlannerProfiles();

    global.R7Chat.platform.hostBridge = {
        getEditorTypeSafe() {
            return 'cell';
        }
    };
    global.R7Chat.platform.desktopTools = {
        getStatus() {
            return {
                desktopEditorAvailable: true,
                catalogAvailable: true,
                executionAvailable: true,
                catalogParseError: '',
                toolCount: 2,
                catalogHash: 'dtb_sheet'
            };
        },
        getCatalog() {
            return [
                {
                    name: 'sheet_append_report',
                    description: 'Append a report block at the bottom of the current sheet',
                    inputSchema: {
                        type: 'object',
                        properties: {
                            sheetName: { type: 'string' }
                        },
                        required: ['sheetName']
                    }
                },
                {
                    name: 'sheet_fill_cells',
                    description: 'Fill a matrix of cells on a worksheet',
                    inputSchema: {
                        type: 'object',
                        properties: {
                            sheetName: { type: 'string' }
                        },
                        required: ['sheetName']
                    }
                }
            ];
        },
        getCatalogForPrompt() {
            return '- sheet_append_report | Append a report block | required=sheetName';
        }
    };
    global.R7Chat.services.chat = {
        loadRuntimeSettings() {
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
    };

    const step = runtime.getForcedInitialStep('сделай отчет внизу листа по всем данным', 0);
    assert.ok(step);
    assert.ok(step.type === 'list_host_tools' || step.type === 'call_host_tool');
    assert.ok(/host tool/i.test(step.reason));
    if (step.type === 'list_host_tools') {
        assert.ok(/отчет/i.test(step.args.query));
    } else {
        assert.strictEqual(step.args.tool, 'sheet_append_report');
        assert.strictEqual(step.args.input.sheetName, 'Sheet2');
    }
}

function testHostToolDiscoveryStoresRelevantSuggestions() {
    const runtime = createAgentRuntime();

    global.R7Chat.platform.desktopTools = {
        getStatus() {
            return {
                desktopEditorAvailable: true,
                catalogAvailable: true,
                executionAvailable: true,
                catalogParseError: '',
                toolCount: 1,
                catalogHash: 'dtb_sheet'
            };
        },
        getCatalog() {
            return [
                {
                    name: 'sheet_append_report',
                    description: 'Append a report block at the bottom of the current sheet',
                    inputSchema: {
                        type: 'object',
                        properties: {
                            sheetName: { type: 'string' },
                            anchor: { type: 'string' }
                        },
                        required: ['sheetName']
                    }
                }
            ];
        }
    };

    return runtime.executeAgentStep({
        id: 'step_1',
        type: 'list_host_tools',
        args: {
            query: 'сделай отчет внизу листа по всем данным'
        }
    }).then(function (result) {
        assert.strictEqual(result.ok, true);
        assert.ok(Array.isArray(result.data.relevant));
        assert.strictEqual(result.data.relevant[0].name, 'sheet_append_report');
        assert.strictEqual(result.data.relevant[0].suggestedInput.anchor, 'bottom');
        assert.strictEqual(runtime.getState().lastHostToolDiscovery.relevant[0].name, 'sheet_append_report');
    });
}

function testDesktopToolsPlannerStateRespectsMacroOnlyAndExplicitMacroRequests() {
    const runtime = createAgentRuntime();
    loadPlannerProfiles();

    global.R7Chat.platform.hostBridge = {
        getEditorTypeSafe() {
            return 'word';
        }
    };
    global.R7Chat.platform.desktopTools = {
        getStatus() {
            return {
                desktopEditorAvailable: true,
                catalogAvailable: true,
                executionAvailable: true,
                catalogParseError: '',
                toolCount: 3,
                catalogHash: 'dtb_demo'
            };
        },
        getCatalogForPrompt() {
            return '- document_insert_summary | Insert a summary block | required=topic';
        }
    };
    global.R7Chat.services.chat = {
        loadRuntimeSettings() {
            return {
                desktopTools: {
                    automationMode: 'macro_only',
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
    };

    const macroOnlyState = runtime.buildDesktopToolsPlannerState('Подготовь краткий summary документа');
    assert.strictEqual(macroOnlyState.enabledForPlanner, false);
    assert.strictEqual(macroOnlyState.reason, 'macro_only_mode');

    const macroOnlyMessage = runtime.getPlannerSystemMessage({
        mode: 'default',
        desktopTools: macroOnlyState
    });
    assert.ok(/HOST_TOOLS_POLICY:/i.test(macroOnlyMessage));
    assert.ok(/run_macro_code/i.test(macroOnlyMessage));

    global.R7Chat.services.chat = {
        loadRuntimeSettings() {
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
    };

    const explicitMacroState = runtime.buildDesktopToolsPlannerState('Напиши макрос для форматирования документа');
    assert.strictEqual(explicitMacroState.enabledForPlanner, false);
    assert.strictEqual(explicitMacroState.reason, 'explicit_macro_request');
}

async function testPlannerUsesActiveProviderInsteadOfOpenRouterFallback() {
    const runtime = createAgentRuntime();
    let requestedConfig = null;

    global.R7Chat.services.chat = {
        loadRuntimeSettings() {
            return {
                activeProvider: 'openai',
                provider: 'openai',
                model: 'gpt-5',
                providers: {
                    openai: {
                        apiKey: 'sk-openai-test-1234567890',
                        model: 'gpt-5',
                        baseUrl: 'https://api.openai.com/v1',
                        reasoningEffort: 'high'
                    },
                    openrouter: {
                        apiKey: '',
                        model: 'openrouter/auto',
                        baseUrl: 'https://openrouter.ai/api/v1'
                    }
                }
            };
        }
    };
    global.R7Chat.features.settings = {
        getProviderConfig(settings, providerId) {
            assert.strictEqual(providerId, 'openai');
            return settings.providers.openai;
        }
    };
    global.R7Chat.platform.providers = {
        async chatRequest(messages, systemMessage, stream, config) {
            requestedConfig = {
                messages,
                systemMessage,
                stream,
                config
            };
            return {
                data: [{
                    content: JSON.stringify({
                        type: 'final_answer',
                        reason: 'done',
                        args: { answer: 'Done.' }
                    })
                }],
                reasoning: {
                    available: true,
                    summary: 'Planner compared available actions and selected final_answer.',
                    effort: 'high',
                    tokens: 18,
                    source: 'openai'
                }
            };
        }
    };
    global.R7Chat.platform.openrouter = {
        async chatRequest() {
            throw new Error('openrouter fallback should not be used');
        }
    };

    const step = await runtime.plannerNextStep([], 0);

    assert.strictEqual(step.type, 'final_answer');
    assert.ok(requestedConfig);
    assert.strictEqual(requestedConfig.stream, false);
    assert.strictEqual(requestedConfig.config.provider, 'openai');
    assert.strictEqual(requestedConfig.config.activeProvider, 'openai');
    assert.strictEqual(requestedConfig.config.model, 'gpt-5');
    assert.strictEqual(requestedConfig.config.reasoningEffort, 'high');
    assert.ok(runtime.getState().trace.some((entry) =>
        entry && entry.step_type === 'model_reasoning' && /Planner compared available actions/i.test(String(entry.reasoning_summary || ''))
    ));
}

async function testPlannerReasoningTraceCanBeDisabled() {
    const runtime = createAgentRuntime();

    global.R7Chat.services.chat = {
        loadRuntimeSettings() {
            return {
                activeProvider: 'openai',
                provider: 'openai',
                model: 'gpt-5',
                trace: {
                    showModelReasoning: false
                },
                providers: {
                    openai: {
                        apiKey: 'sk-openai-test-1234567890',
                        model: 'gpt-5',
                        baseUrl: 'https://api.openai.com/v1',
                        reasoningEffort: 'medium'
                    },
                    openrouter: {
                        apiKey: '',
                        model: 'openrouter/auto',
                        baseUrl: 'https://openrouter.ai/api/v1'
                    }
                }
            };
        }
    };
    global.R7Chat.features.settings = {
        getProviderConfig(settings, providerId) {
            assert.strictEqual(providerId, 'openai');
            return settings.providers.openai;
        }
    };
    global.R7Chat.platform.providers = {
        async chatRequest() {
            return {
                data: [{
                    content: JSON.stringify({
                        type: 'final_answer',
                        reason: 'done',
                        args: { answer: 'Done.' }
                    })
                }],
                reasoning: {
                    available: true,
                    summary: 'This summary should stay hidden from trace.',
                    effort: 'medium',
                    tokens: 10,
                    source: 'openai'
                }
            };
        }
    };
    global.R7Chat.platform.openrouter = {
        async chatRequest() {
            throw new Error('openrouter fallback should not be used');
        }
    };

    await runtime.plannerNextStep([], 0);

    assert.strictEqual(runtime.getState().trace.some((entry) => entry && entry.step_type === 'model_reasoning'), false);
}

function testForcedInitialStepGeneratesImageCommandForWordRequest() {
    const runtime = createAgentRuntime();
    global.R7Chat.platform.hostBridge = {
        getEditorTypeSafe() {
            return 'word';
        }
    };

    runtime.getState().lastReadDocumentText = 'OpenAI plans to double its workforce and prepare for an IPO in 2026.';

    const firstStep = runtime.getForcedInitialStep('сгенери картинку', 0);
    const secondStep = runtime.getForcedInitialStep('сгенери картинку', 1);

    assert.strictEqual(firstStep.type, 'read_document_snapshot');
    assert.strictEqual(firstStep.args.editorType, 'word');
    assert.strictEqual(secondStep.type, 'final_answer');
    assert.ok(/\[GENERATE_IMAGE:/i.test(secondStep.args.answer));
    assert.ok(/OpenAI/i.test(secondStep.args.answer));
}

function testForcedInitialStepUsesRecordedSpreadsheetReadStepForCell() {
    const runtime = createAgentRuntime();
    global.R7Chat.platform.hostBridge = {
        getEditorTypeSafe() {
            return 'cell';
        }
    };

    const firstStep = runtime.getForcedInitialStep('Суммируй активный лист в 5 пунктах', 0);
    assert.strictEqual(firstStep.type, 'read_active_sheet');
    assert.strictEqual(firstStep.args.editorType, 'cell');
}

function testPlanApprovalMessageHelpers() {
    const runtime = createAgentRuntime();

    assert.strictEqual(runtime.getPlanApprovalCommandText(), 'Начинай исполнять план');
    assert.strictEqual(runtime.isPlanApprovalMessage('Начинай исполнять план'), true);
    assert.strictEqual(runtime.isPlanApprovalMessage('Давай начинай исполнять план'), true);
    assert.strictEqual(runtime.isPlanApprovalMessage('approve plan'), true);
    assert.strictEqual(runtime.isPlanApprovalMessage('Измени второй раздел и замени картинку'), false);
}

function testExtractImageCommandSeparatesVisibleText() {
    const runtime = createAgentRuntime();
    const command = runtime.extractImageCommand('Generating image now.\n[GENERATE_IMAGE: futuristic newsroom illustration]');

    assert.strictEqual(command.prompt, 'futuristic newsroom illustration');
    assert.strictEqual(command.visibleText, 'Generating image now.');
}

async function testPlannerRetriesOnEmptyResponse() {
    const runtime = createAgentRuntime();
    let calls = 0;

    global.R7Chat.platform.providers = {
        async chatRequest() {
            calls += 1;
            return {
                data: [{
                    content: calls === 1 ? '' : JSON.stringify({
                        type: 'final_answer',
                        reason: 'done',
                        args: { answer: 'Done.' }
                    })
                }]
            };
        }
    };

    const step = await runtime.plannerNextStep([], 0);
    assert.strictEqual(step.type, 'final_answer');
    assert.strictEqual(calls, 2);
}

async function testRunMacroExecutionIsNotBlockedByResearchPolicy() {
    const runtime = createAgentRuntime();
    runtime.getState().researchPolicy = {
        required: true,
        query: 'квантовые вычисления',
        completed: false,
        attempted: true,
        failed: true,
        webSearchEnabled: true,
        activeProvider: 'brave',
        braveApiConfigured: true,
        exaApiConfigured: false
    };

    const result = await runtime.executeAgentStep({
        id: 'step_blocked',
        type: 'run_macro_code',
        reason: 'Try to write without research',
        args: {},
        macro_code: 'return 1;'
    });

    assert.strictEqual(result.ok, false);
    assert.strictEqual(result.step_type, 'run_macro_code');
    assert.ok(/host bridge is unavailable/i.test(result.error));
}

async function testWordMacroValidationBlocksUnsupportedHeadingMethod() {
    const runtime = createAgentRuntime();
    let called = false;

    global.R7Chat.platform.hostBridge = {
        getEditorTypeSafe() {
            return 'word';
        },
        async callEditorCommand() {
            called = true;
            return '{}';
        }
    };

    const result = await runtime.executeAgentStep({
        id: 'step_invalid_word_macro',
        type: 'run_macro_code',
        reason: 'Insert Word heading with invalid API',
        args: {},
        macro_code: 'var p = Api.CreateParagraph(); p.SetHeading("Heading 1"); return true;'
    });

    assert.strictEqual(result.ok, false);
    assert.strictEqual(called, false);
    assert.ok(/SetHeading is not a verified Word API method/i.test(result.error));

    const hint = runtime.buildMacroRecoveryPlannerHint({
        error: result.error
    });
    assert.ok(/Never use SetHeading/i.test(hint));
}

async function testRunMacroCodeSupportsIifeStyleMacros() {
    const runtime = createAgentRuntime();

    global.R7Chat.platform.hostBridge = {
        getEditorTypeSafe() {
            return 'cell';
        },
        async callEditorCommand(command) {
            return command();
        }
    };

    const result = await runtime.executeAgentStep({
        id: 'step_iife_macro',
        type: 'run_macro_code',
        reason: 'Run recorded IIFE macro',
        args: {},
        macro_code: '(function () { return { sheetName: "Лист1", value: 6, rows: [1, 2] }; })();'
    });

    assert.strictEqual(result.ok, true);
    assert.deepStrictEqual(result.data, {
        sheetName: 'Лист1',
        value: 6,
        rows: [1, 2]
    });
}

async function testReadActiveSheetCarriesSparseNonEmptyCells() {
    const runtime = createAgentRuntime();

    global.R7Chat.platform.hostBridge = {
        getEditorTypeSafe() {
            return 'cell';
        }
    };

    global.R7Chat.services.context = {
        async collectActiveSheetContext() {
            return {
                sheetName: 'Лист1',
                address: '$A$1:$K$25',
                values: [
                    ['чи', '', '', '', '', '', '', '', '', '', ''],
                    ['', '', '', '', '', '', '', '', '', '', ''],
                    ['', '', '', '', '', '', '6', '', '', '', ''],
                    ['', '', '', '', '6', '', '', '', '', '', ''],
                    ['', '', '', '', '', '', '', '', '', '', ''],
                    ['', '', '', '', '', '', '', '', '', '6', '']
                ]
            };
        },
        formatTablePayload(values) {
            return {
                text: 'A1=чи\nG3=6\nE4=6\nJ6=6',
                totalRows: values.length,
                totalCols: 11,
                sampledRows: values.length,
                sampledCols: 11,
                truncated: false
            };
        }
    };

    const result = await runtime.executeAgentStep({
        id: 'step_sparse_sheet',
        type: 'read_active_sheet',
        reason: 'Read active sheet',
        args: {}
    });

    assert.strictEqual(result.ok, true);
    assert.strictEqual(result.data.nonEmptyCount, 4);
    assert.deepStrictEqual(result.data.nonEmptyCells, [
        { address: 'A1', value: 'чи' },
        { address: 'G3', value: '6' },
        { address: 'E4', value: '6' },
        { address: 'J6', value: '6' }
    ]);
}

async function testWebToolExecutorSuccessAndError() {
    const runtime = createAgentRuntime();
    let searchCalls = 0;
    let crawlCalls = 0;

    global.R7Chat.platform.webTools = {
        isEnabled() {
            return true;
        },
        async executeWebSearch(args) {
            searchCalls += 1;
            assert.deepStrictEqual(args, { query: 'latest openrouter tools' });
            return {
                provider: 'exa',
                results: [{ title: 'Doc', url: 'https://example.com' }],
                errors: [],
                statuses: [{ status: 'ok' }],
                sources: [{ url: 'https://example.com', title: 'Doc' }]
            };
        },
        async executeWebCrawling(args) {
            crawlCalls += 1;
            assert.deepStrictEqual(args, { urls: ['http://127.0.0.1'] });
            return {
                provider: 'brave',
                results: [],
                errors: [{ code: 'private_host_blocked', message: 'URL validation failed' }],
                statuses: [],
                sources: []
            };
        }
    };

    const searchResult = await runtime.executeAgentStep({
        id: 'step_1',
        type: 'web_search',
        args: { query: 'latest openrouter tools' }
    });
    const crawlResult = await runtime.executeAgentStep({
        id: 'step_2',
        type: 'web_crawling',
        args: { urls: ['http://127.0.0.1'] }
    });

    assert.strictEqual(searchCalls, 1);
    assert.strictEqual(crawlCalls, 1);
    assert.strictEqual(searchResult.ok, true);
    assert.strictEqual(searchResult.step_type, 'web_search');
    assert.strictEqual(searchResult.data.provider, 'exa');
    assert.strictEqual(crawlResult.ok, false);
    assert.strictEqual(crawlResult.step_type, 'web_crawling');
    assert.strictEqual(crawlResult.error, 'URL validation failed');
}

async function testHostToolExecutorAndRecoveryHint() {
    const runtime = createAgentRuntime();
    let captured = null;

    global.R7Chat.platform.hostBridge = {
        getEditorTypeSafe() {
            return 'cell';
        }
    };
    global.R7Chat.platform.desktopTools = {
        async callTool(name, args) {
            captured = { name, args };
            return {
                ok: true,
                descriptor: { name: name },
                data: { done: true }
            };
        }
    };

    const result = await runtime.executeAgentStep({
        id: 'step_1',
        type: 'call_host_tool',
        args: {
            tool: 'worksheet_create_table',
            input: {
                title: 'Q1'
            }
        }
    });

    assert.strictEqual(result.ok, true);
    assert.strictEqual(result.step_type, 'call_host_tool');
    assert.deepStrictEqual(captured, {
        name: 'worksheet_create_table',
        args: { title: 'Q1' }
    });
    assert.strictEqual(result.data.tool, 'worksheet_create_table');

    const hint = runtime.buildHostToolRecoveryPlannerHint({
        args: { tool: 'worksheet_create_table' }
    }, {
        error: 'Tool is not present in the runtime catalog: worksheet_create_table'
    });
    assert.ok(/run_macro_code/i.test(hint));
}

async function testCollectContextExecutorAndVerificationHelper() {
    const runtime = createAgentRuntime();
    global.R7Chat.services.context = {
        async discoverAvailableContext(options) {
            assert.deepStrictEqual(options, {
                mode: 'discover',
                editorType: 'cell',
                forceRefresh: true,
                intent: 'initial_context'
            });
            return {
                editorType: 'cell',
                mode: 'discover',
                source: 'workbook',
                activeSheet: 'Sheet1',
                sheets: [{ name: 'Sheet1', address: 'A1:B2', rows: 2, cols: 2 }],
                discoveryStatus: 'ok',
                warnings: []
            };
        },
        async collectContext() {
            throw new Error('collectContext should not be called in this branch');
        }
    };
    global.R7Chat.platform.hostBridge = {
        getEditorTypeSafe() {
            return 'cell';
        }
    };

    const result = await runtime.executeAgentStep({
        id: 'step_1',
        type: 'collect_context',
        args: {
            mode: 'discover',
            editorType: 'cell',
            forceRefresh: true,
            intent: 'initial_context'
        }
    });

    assert.strictEqual(result.ok, true);
    assert.strictEqual(result.step_type, 'collect_context');
    assert.strictEqual(result.data.discoveryStatus, 'ok');

    runtime.getState().steps = [
        { id: 'step_1', type: 'run_macro_code' }
    ];
    assert.strictEqual(runtime.hasPendingPostMacroVerification(), true);
    const forcedStep = runtime.buildForcedPostMacroVerificationStep(1);
    assert.strictEqual(forcedStep.type, 'read_active_sheet');
}

async function testReadDocumentSnapshotExecutor() {
    const runtime = createAgentRuntime();
    let commandWasCalled = false;

    global.R7Chat.platform.hostBridge = {
        getEditorTypeSafe() {
            return 'word';
        },
        async callEditorCommand(command) {
            commandWasCalled = typeof command === 'function';
            return JSON.stringify({
                editorType: 'word',
                mode: 'snapshot',
                source: 'document_snapshot',
                coverage: {
                    totalParagraphs: 3,
                    nonEmptyParagraphs: 2,
                    collectedParagraphs: 3
                },
                objects: {
                    imageCount: 1,
                    tableCount: 0
                },
                textChunks: [{ start: 1, end: 3, text: 'Alpha\nBeta' }],
                payload: '[Paragraphs 1-3]\nAlpha\nBeta',
                truncated: false,
                warnings: []
            });
        }
    };

    const result = await runtime.executeAgentStep({
        id: 'step_snapshot',
        type: 'read_document_snapshot',
        args: {
            editorType: 'word'
        }
    });

    assert.strictEqual(commandWasCalled, true);
    assert.strictEqual(result.ok, true);
    assert.strictEqual(result.step_type, 'read_document_snapshot');
    assert.strictEqual(result.data.coverage.totalParagraphs, 3);
    assert.strictEqual(result.data.objects.imageCount, 1);
    assert.strictEqual(Array.isArray(result.data.textChunks), true);
    assert.strictEqual(result.data.truncated, false);
}

async function testReadDocumentSnapshotExecutor() {
    const runtime = createAgentRuntime();
    let commandWasCalled = false;

    global.R7Chat.platform.hostBridge = {
        getEditorTypeSafe() {
            return 'word';
        },
        async callEditorCommand(command) {
            commandWasCalled = typeof command === 'function';
            return JSON.stringify({
                editorType: 'word',
                mode: 'snapshot',
                source: 'document_snapshot',
                coverage: {
                    totalParagraphs: 3,
                    nonEmptyParagraphs: 2,
                    collectedParagraphs: 3
                },
                objects: {
                    imageCount: 1,
                    tableCount: 0
                },
                textChunks: [{ start: 1, end: 3, text: 'Alpha\nBeta' }],
                payload: '[Paragraphs 1-3]\nAlpha\nBeta',
                truncated: false,
                warnings: []
            });
        }
    };

    const result = await runtime.executeAgentStep({
        id: 'step_snapshot',
        type: 'read_document_snapshot',
        args: {
            editorType: 'word'
        }
    });

    assert.strictEqual(commandWasCalled, true);
    assert.strictEqual(result.ok, true);
    assert.strictEqual(result.step_type, 'read_document_snapshot');
    assert.strictEqual(result.data.coverage.totalParagraphs, 3);
    assert.strictEqual(result.data.objects.imageCount, 1);
    assert.strictEqual(Array.isArray(result.data.textChunks), true);
    assert.strictEqual(result.data.truncated, false);
}

async function testGenerateImageAssetExecutorForWord() {
    const runtime = createAgentRuntime();
    let insertedScope = null;

    global.R7Chat.platform.hostBridge = {
        getEditorTypeSafe() {
            return 'word';
        },
        async callEditorCommand(command, scopeData) {
            insertedScope = scopeData;
            return JSON.stringify({ inserted: true, blocks: 2 });
        }
    };
    global.R7Chat.platform.image = {
        async generate(prompt) {
            assert.strictEqual(prompt, 'Editorial illustration of AI in a classroom');
            return {
                data: [{
                    url: 'https://example.com/generated-image.png'
                }]
            };
        }
    };

    const result = await runtime.executeAgentStep({
        id: 'step_image',
        type: 'generate_image_asset',
        args: {
            prompt: 'Editorial illustration of AI in a classroom',
            caption: 'AI in education',
            sectionTitle: 'Section image'
        }
    });

    assert.strictEqual(result.ok, true);
    assert.strictEqual(result.step_type, 'generate_image_asset');
    assert.strictEqual(result.data.inserted, true);
    assert.strictEqual(insertedScope.caption, 'AI in education');
    assert.strictEqual(insertedScope.sectionTitle, 'Section image');
}

function testPlannerConversationSeedIncludesAssistantHistory() {
    const runtime = createAgentRuntime();
    global.R7Chat.services.chatThreads = {
        syncConversationHistory() {
            return [
                { role: 'user', content: 'Начни поиск web search про Ивана 3' },
                { role: 'assistant', content: 'Иван III (1440–1505) ...' },
                { role: 'user', content: 'ты взял эти факты откуда?' }
            ];
        }
    };

    const seed = runtime.buildPlannerConversationSeed('ты взял эти факты откуда?');
    assert.strictEqual(Array.isArray(seed), true);
    assert.strictEqual(seed.length, 3);
    assert.strictEqual(seed[0].role, 'user');
    assert.strictEqual(seed[1].role, 'assistant');
    assert.strictEqual(seed[2].role, 'user');
    assert.strictEqual(seed[2].content, 'ты взял эти факты откуда?');
}

function testPlannerConversationSeedFallbackToCurrentUserOnly() {
    const runtime = createAgentRuntime();
    global.R7Chat.services.chatThreads = {
        syncConversationHistory() {
            return [];
        }
    };
    const seed = runtime.buildPlannerConversationSeed('Проверь источник фактов');
    assert.deepStrictEqual(seed, [{
        role: 'user',
        content: 'Проверь источник фактов'
    }]);
}

function testBuildToolResultMessageCompactsImageDataUrl() {
    const runtime = createAgentRuntime();
    const dataUrl = 'data:image/png;base64,' + 'A'.repeat(24000);
    const payload = runtime.buildToolResultMessage({
        step_id: 'step_image',
        step_type: 'generate_image_asset',
        ok: true,
        data: {
            prompt: 'A large visual',
            url: dataUrl,
            caption: 'caption',
            inserted: true
        },
        logs: [],
        error: null,
        duration_ms: 321
    });

    assert.ok(payload.indexOf('data:image/png;base64,AAA') === -1);
    assert.ok(payload.indexOf('[data-image-omitted chars=') !== -1);
}

function testPlannerContextCompactionRemovesOversizedPayloads() {
    const runtime = createAgentRuntime();
    const messages = [{
        role: 'user',
        content: 'Write a long article about AI.'
    }];
    const hugeDataUrl = 'data:image/png;base64,' + 'B'.repeat(18000);

    for (let i = 0; i < 32; i += 1) {
        messages.push({
            role: 'assistant',
            content: JSON.stringify({ id: 'step_' + i, type: 'generate_image_asset' })
        });
        messages.push({
            role: 'user',
            content: JSON.stringify({
                type: 'tool_result',
                step_id: 'step_' + i,
                step_type: 'generate_image_asset',
                data: { url: hugeDataUrl, inserted: true }
            }, null, 2)
        });
    }

    const result = runtime.compactPlannerMessagesInPlace(messages, { reason: 'test_compaction' });
    const totalChars = messages.reduce((acc, item) => acc + String(item && item.content || '').length, 0);
    const joinedContent = messages.map((item) => String(item && item.content || '')).join('\n');

    assert.strictEqual(result.changed, true);
    assert.ok(result.removedMessages >= 0);
    assert.ok(totalChars <= 120000);
    assert.ok(joinedContent.indexOf('data:image/png;base64,BBB') === -1);
}

function testContextOverflowErrorDetection() {
    const runtime = createAgentRuntime();
    const detected = runtime.isContextOverflowError({
        status: 400,
        message: "This endpoint's maximum context length is 204800 tokens. However, you requested about 329989 tokens."
    });
    assert.strictEqual(detected, true);
}

async function run() {
    testPlannerParsing();
    testPlannerWebStepParsing();
    testPlannerConversationSeedIncludesAssistantHistory();
    testPlannerConversationSeedFallbackToCurrentUserOnly();
    testPlannerHostToolStepParsing();
    testPlannerCollectContextStepParsing();
    testPlannerReadDocumentSnapshotStepParsing();
    testPlannerWordPlanStepParsing();
    testPlannerImageStepParsing();
    testPlannerRunMacroStepReadsMacroCodeFromArgs();
    testPlannerErrorsAndMethodPack();
    testApiReferenceFuzzyLookupAndRecoveryStep();
    testResearchPolicyUnderstandsConfiguredBrave();
    testWordPlanModeHeuristicAndForcedResearch();
    testWebSearchRecoveryHelpers();
    testForcedInitialStepAddsWebSearchForWordResearch();
    testPlannerSystemMessageIncludesWebSearchStatus();
    testPlannerSystemMessageMentionsPluginAndActiveOfficeEnvironment();
    testSpreadsheetReadIntentForcesSpreadsheetReadStepBeforeAnswer();
    testSpreadsheetMacroWordingStillForcesRecordedReadStep();
    testDesktopToolsPlannerStateForWordAutoMode();
    testForcedInitialStepUsesHostToolDiscoveryForSpreadsheetWrite();
    await testHostToolDiscoveryStoresRelevantSuggestions();
    testDesktopToolsPlannerStateRespectsMacroOnlyAndExplicitMacroRequests();
    await testPlannerUsesActiveProviderInsteadOfOpenRouterFallback();
    await testPlannerReasoningTraceCanBeDisabled();
    testForcedInitialStepGeneratesImageCommandForWordRequest();
    testForcedInitialStepUsesRecordedSpreadsheetReadStepForCell();
    testPlanApprovalMessageHelpers();
    testExtractImageCommandSeparatesVisibleText();
    await testPlannerRetriesOnEmptyResponse();
    await testHostToolExecutorAndRecoveryHint();
    await testCollectContextExecutorAndVerificationHelper();
    await testReadDocumentSnapshotExecutor();
    await testGenerateImageAssetExecutorForWord();
    testBuildToolResultMessageCompactsImageDataUrl();
    testPlannerContextCompactionRemovesOversizedPayloads();
    testContextOverflowErrorDetection();
    await testWebToolExecutorSuccessAndError();
    await testRunMacroExecutionIsNotBlockedByResearchPolicy();
    await testWordMacroValidationBlocksUnsupportedHeadingMethod();
    await testRunMacroCodeSupportsIifeStyleMacros();
    await testReadActiveSheetCarriesSparseNonEmptyCells();
    console.log('r7chat_agent_runtime.test.js: ok');
}

run().catch((error) => {
    console.error(error);
    process.exit(1);
});
