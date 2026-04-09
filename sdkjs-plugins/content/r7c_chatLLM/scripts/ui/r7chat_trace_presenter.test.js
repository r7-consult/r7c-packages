const assert = require('assert');
const presenter = require('./r7chat_trace_presenter');

function testBuildsResearchStageFromPlannerTrace() {
    const viewModel = presenter.buildTraceViewModel([
        { step_type: 'run_start', status: 'info' },
        { step_type: 'web_search_capability', status: 'info' },
        { step_type: 'web_search', status: 'planned' },
        { step_type: 'web_search', status: 'ok', reason: 'search complete' }
    ], {
        status: 'executing',
        elapsedMs: 4800
    });

    assert.strictEqual(viewModel.stages.length, 1);
    assert.strictEqual(viewModel.stages[0].title, '\u041f\u043e\u0438\u0441\u043a \u0438\u043d\u0444\u043e\u0440\u043c\u0430\u0446\u0438\u0438');
    assert.strictEqual(viewModel.stages[0].summary, '\u0412\u044b\u043f\u043e\u043b\u043d\u0435\u043d \u0432\u0435\u0431-\u043f\u043e\u0438\u0441\u043a');
    assert.strictEqual(viewModel.summaryBarText, '\u0412\u044b\u043f\u043e\u043b\u043d\u044f\u0435\u0442\u0441\u044f: \u041f\u043e\u0438\u0441\u043a \u0438\u043d\u0444\u043e\u0440\u043c\u0430\u0446\u0438\u0438 \u2022 4.8\u0441');
}

function testBuildsAwaitingPlanStageAndSummaryBar() {
    const viewModel = presenter.buildTraceViewModel([
        { step_type: 'present_plan', status: 'awaiting_approval' }
    ], {
        status: 'awaiting_plan',
        elapsedMs: 6100
    });

    assert.strictEqual(viewModel.stages.length, 1);
    assert.strictEqual(viewModel.stages[0].title, '\u0421\u043e\u0433\u043b\u0430\u0441\u043e\u0432\u0430\u043d\u0438\u0435 \u043f\u043b\u0430\u043d\u0430');
    assert.strictEqual(viewModel.stages[0].summary, '\u041f\u043b\u0430\u043d \u043f\u043e\u0434\u0433\u043e\u0442\u043e\u0432\u043b\u0435\u043d \u0438 \u0436\u0434\u0451\u0442 \u043f\u043e\u0434\u0442\u0432\u0435\u0440\u0436\u0434\u0435\u043d\u0438\u044f');
    assert.strictEqual(viewModel.summaryBarText, '\u041e\u0436\u0438\u0434\u0430\u0435\u0442 \u043f\u043e\u0434\u0442\u0432\u0435\u0440\u0436\u0434\u0435\u043d\u0438\u044f \u043f\u043b\u0430\u043d\u0430 \u2022 1 \u044d\u0442\u0430\u043f \u2022 6.1\u0441');
}

function testBuildsErrorStageAndSummaryBar() {
    const viewModel = presenter.buildTraceViewModel([
        { step_type: 'planner_error', status: 'error', error: 'Provider timeout' }
    ], {
        status: 'error',
        elapsedMs: 7200
    });

    assert.strictEqual(viewModel.stages.length, 1);
    assert.strictEqual(viewModel.stages[0].title, '\u041e\u0448\u0438\u0431\u043a\u0430');
    assert.strictEqual(viewModel.stages[0].summary, '\u041d\u0435 \u0443\u0434\u0430\u043b\u043e\u0441\u044c \u043f\u0440\u043e\u0434\u043e\u043b\u0436\u0438\u0442\u044c \u0432\u044b\u043f\u043e\u043b\u043d\u0435\u043d\u0438\u0435');
    assert.strictEqual(viewModel.stages[0].details.indexOf('Provider timeout') >= 0, true);
    assert.strictEqual(viewModel.summaryBarText, '\u041e\u0448\u0438\u0431\u043a\u0430 \u043d\u0430 \u044d\u0442\u0430\u043f\u0435 "\u041e\u0448\u0438\u0431\u043a\u0430" \u2022 7.2\u0441');
}

function testBuildsErrorStageWithoutStepIdentifier() {
    const viewModel = presenter.buildTraceViewModel([
        { status: 'error', error: 'Missing planner step id' }
    ], {
        status: 'error',
        elapsedMs: 1200
    });

    assert.strictEqual(viewModel.stages.length, 1);
    assert.strictEqual(viewModel.stages[0].title, '\u041e\u0448\u0438\u0431\u043a\u0430');
    assert.strictEqual(viewModel.stages[0].details.indexOf('Missing planner step id') >= 0, true);
}

function testBuildsErrorStageForCapabilityFailure() {
    const viewModel = presenter.buildTraceViewModel([
        { step_type: 'web_search_capability', status: 'error', error: 'Capability probe failed' }
    ], {
        status: 'error',
        elapsedMs: 2300
    });

    assert.strictEqual(viewModel.stages.length, 1);
    assert.strictEqual(viewModel.stages[0].title, '\u041e\u0448\u0438\u0431\u043a\u0430');
    assert.strictEqual(viewModel.stages[0].details.indexOf('Capability probe failed') >= 0, true);
}

function testBuildsDoneSummaryStateAndCollapseSignal() {
    const viewModel = presenter.buildTraceViewModel([
        { step_id: 'run_1', status: 'started' },
        { step_id: 'final_answer_1', step_type: 'final_answer', status: 'completed', reason: 'Final answer produced' }
    ], {
        status: 'idle',
        elapsedMs: 9900
    });

    assert.strictEqual(viewModel.stages.length, 2);
    assert.strictEqual(viewModel.stages[0].title, '\u041f\u043e\u0434\u0433\u043e\u0442\u043e\u0432\u043a\u0430');
    assert.strictEqual(viewModel.stages[1].title, '\u041e\u0442\u0432\u0435\u0442');
    assert.strictEqual(viewModel.summaryBarText, '\u0413\u043e\u0442\u043e\u0432\u043e \u2022 2 \u044d\u0442\u0430\u043f\u0430 \u2022 9.9\u0441');
    assert.strictEqual(viewModel.summaryBarState, 'success');
    assert.strictEqual(viewModel.shouldCollapseOnSuccess, true);
}

function testMapsRunMacroCodeToApplyChangesStage() {
    const viewModel = presenter.buildTraceViewModel([
        { step_type: 'run_macro_code', status: 'ok', reason: 'Workbook updated successfully' }
    ], {
        status: 'executing',
        elapsedMs: 3500
    });

    assert.strictEqual(viewModel.stages.length, 1);
    assert.strictEqual(viewModel.stages[0].title, '\u041f\u0440\u0438\u043c\u0435\u043d\u0435\u043d\u0438\u0435 \u0438\u0437\u043c\u0435\u043d\u0435\u043d\u0438\u0439');
    assert.strictEqual(viewModel.summaryBarText, '\u0412\u044b\u043f\u043e\u043b\u043d\u044f\u0435\u0442\u0441\u044f: \u041f\u0440\u0438\u043c\u0435\u043d\u0435\u043d\u0438\u0435 \u0438\u0437\u043c\u0435\u043d\u0435\u043d\u0438\u0439 \u2022 3.5\u0441');
}

function testBuildsRunningSummaryBarDuringPlanning() {
    const viewModel = presenter.buildTraceViewModel([
        { step_type: 'collect_context', status: 'ok', reason: 'Loaded workbook context' }
    ], {
        status: 'planning',
        elapsedMs: 2700
    });

    assert.strictEqual(viewModel.stages.length, 1);
    assert.strictEqual(viewModel.summaryBarText, '\u0412\u044b\u043f\u043e\u043b\u043d\u044f\u0435\u0442\u0441\u044f: \u0410\u043d\u0430\u043b\u0438\u0437 \u043a\u043e\u043d\u0442\u0435\u043a\u0441\u0442\u0430 \u2022 2.7\u0441');
    assert.strictEqual(viewModel.summaryBarState, 'running');
}

function testBuildsRunningSummaryBarDuringAnswering() {
    const viewModel = presenter.buildTraceViewModel([
        { step_type: 'final_answer', status: 'started', reason: 'Drafting final response' }
    ], {
        status: 'answering',
        elapsedMs: 4100
    });

    assert.strictEqual(viewModel.stages.length, 1);
    assert.strictEqual(viewModel.summaryBarText, '\u0412\u044b\u043f\u043e\u043b\u043d\u044f\u0435\u0442\u0441\u044f: \u041e\u0442\u0432\u0435\u0442 \u2022 4.1\u0441');
    assert.strictEqual(viewModel.summaryBarState, 'running');
}

function testKeepsPlanFollowUpEventsVisibleInline() {
    const viewModel = presenter.buildTraceViewModel([
        { step_type: 'plan_approved', status: 'ok', reason: 'Plan approved by user' },
        { step_type: 'plan_revision', status: 'ok', reason: 'Plan revised after feedback' }
    ], {
        status: 'planning',
        elapsedMs: 5200
    });

    assert.strictEqual(viewModel.stages.length, 1);
    assert.strictEqual(viewModel.stages[0].title, '\u0421\u043e\u0433\u043b\u0430\u0441\u043e\u0432\u0430\u043d\u0438\u0435 \u043f\u043b\u0430\u043d\u0430');
    assert.strictEqual(viewModel.stages[0].traces.length, 2);
    assert.strictEqual(viewModel.summaryBarText, '\u0412\u044b\u043f\u043e\u043b\u043d\u044f\u0435\u0442\u0441\u044f: \u0421\u043e\u0433\u043b\u0430\u0441\u043e\u0432\u0430\u043d\u0438\u0435 \u043f\u043b\u0430\u043d\u0430 \u2022 5.2\u0441');
}

function testMapsSpreadsheetReadAndApiReferenceStages() {
    const viewModel = presenter.buildTraceViewModel([
        { step_type: 'read_active_sheet', status: 'ok', reason: 'Read active sheet data' },
        { step_type: 'analyze_reference_macros', status: 'ok', reason: 'Loaded API reference methods' }
    ], {
        status: 'planning',
        elapsedMs: 3300
    });

    assert.strictEqual(viewModel.stages.length, 2);
    assert.strictEqual(viewModel.stages[0].title, '\u0410\u043d\u0430\u043b\u0438\u0437 \u043a\u043e\u043d\u0442\u0435\u043a\u0441\u0442\u0430');
    assert.strictEqual(viewModel.stages[1].title, '\u041f\u0440\u043e\u0432\u0435\u0440\u043a\u0430 API \u0441\u043f\u0440\u0430\u0432\u043e\u0447\u043d\u0438\u043a\u0430');
    assert.strictEqual(viewModel.summaryBarText, '\u0412\u044b\u043f\u043e\u043b\u043d\u044f\u0435\u0442\u0441\u044f: \u041f\u0440\u043e\u0432\u0435\u0440\u043a\u0430 API \u0441\u043f\u0440\u0430\u0432\u043e\u0447\u043d\u0438\u043a\u0430 \u2022 3.3\u0441');
}

function testBuildsStoppedTerminalStateInsteadOfSuccessWhenRuntimeReturnsIdle() {
    const viewModel = presenter.buildTraceViewModel([
        { step_id: 'run_1', status: 'started' },
        { step_type: 'present_plan', status: 'awaiting_approval' },
        { step_type: 'plan_cancelled', status: 'cancelled', reason: 'User cancelled the plan' },
        { step_type: 'agent_stop_requested', status: 'stopped', reason: 'Stop requested by user' },
        { step_type: 'agent_stop', status: 'stopped', reason: 'Execution stopped by user' }
    ], {
        status: 'idle',
        elapsedMs: 6800
    });

    assert.strictEqual(viewModel.stages.length, 3);
    assert.strictEqual(viewModel.stages[2].title, '\u041e\u0441\u0442\u0430\u043d\u043e\u0432\u043b\u0435\u043d\u043e');
    assert.strictEqual(viewModel.stages[2].summary, '\u0412\u044b\u043f\u043e\u043b\u043d\u0435\u043d\u0438\u0435 \u043e\u0441\u0442\u0430\u043d\u043e\u0432\u043b\u0435\u043d\u043e \u0438\u043b\u0438 \u043e\u0442\u043c\u0435\u043d\u0435\u043d\u043e');
    assert.strictEqual(viewModel.summaryBarText, '\u0412\u044b\u043f\u043e\u043b\u043d\u0435\u043d\u0438\u0435 \u043e\u0441\u0442\u0430\u043d\u043e\u0432\u043b\u0435\u043d\u043e \u2022 6.8\u0441');
    assert.strictEqual(viewModel.summaryBarState, 'stopped');
    assert.strictEqual(viewModel.shouldCollapseOnSuccess, false);
}

function run() {
    testBuildsResearchStageFromPlannerTrace();
    testBuildsAwaitingPlanStageAndSummaryBar();
    testBuildsErrorStageAndSummaryBar();
    testBuildsErrorStageWithoutStepIdentifier();
    testBuildsErrorStageForCapabilityFailure();
    testBuildsDoneSummaryStateAndCollapseSignal();
    testMapsRunMacroCodeToApplyChangesStage();
    testBuildsRunningSummaryBarDuringPlanning();
    testBuildsRunningSummaryBarDuringAnswering();
    testKeepsPlanFollowUpEventsVisibleInline();
    testMapsSpreadsheetReadAndApiReferenceStages();
    testBuildsStoppedTerminalStateInsteadOfSuccessWhenRuntimeReturnsIdle();
    console.log('r7chat_trace_presenter.test.js: ok');
}

run();
