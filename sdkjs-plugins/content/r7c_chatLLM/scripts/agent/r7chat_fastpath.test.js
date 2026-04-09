const assert = require('assert');

function loadFresh(modulePath) {
    delete require.cache[require.resolve(modulePath)];
    return require(modulePath);
}

function loadFastPath() {
    global.window = global;
    global.R7Chat = { agent: {} };
    loadFresh('./r7chat_fastpath.js');
    return global.R7Chat.agent.fastPath;
}

function testSimpleWorksheetScaffoldQueue() {
    const fastPath = loadFastPath();
    const queue = fastPath.createQueueForMessage('Создай новый лист "Итоги" и заполни диапазон A1:C5 тестовой таблицей с заголовками.', 0);

    assert.strictEqual(queue.length, 3);
    assert.strictEqual(queue[0].type, 'run_macro_code');
    assert.strictEqual(queue[1].type, 'run_macro_code');
    assert.strictEqual(queue[2].type, 'final_answer');
    assert.strictEqual(queue[0].args.sheetName, 'Итоги');
    assert.strictEqual(queue[0].args.range, 'A1:C5');
    assert.ok(queue[0].macro_code.includes('Api.AddSheet'));
    assert.ok(queue[0].macro_code.includes('simple_sheet_scaffold_done'));
    assert.ok(queue[1].macro_code.includes('simple_sheet_scaffold_verified'));
    assert.ok(queue[2].args.answer.includes('Итоги'));
    assert.ok(queue[2].args.answer.includes('A1:C5'));
}

function testHeavyWorkbookQueueStillWorks() {
    const fastPath = loadFastPath();
    const queue = fastPath.createQueueForMessage('Создай excel с 2000 строк inventory для цветочного магазина', 0);

    assert.ok(queue.length >= 2);
    assert.strictEqual(queue[0].type, 'run_macro_code');
    assert.strictEqual(queue[1].type, 'read_sheet_range');
}

function run() {
    testSimpleWorksheetScaffoldQueue();
    testHeavyWorkbookQueueStillWorks();
    console.log('r7chat_fastpath.test.js: ok');
}

run();
