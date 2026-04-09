const assert = require('assert');
const threadTitles = require('./r7chat_thread_title');

function testBuildAutoTitle() {
    assert.strictEqual(
        threadTitles.buildAutoTitle('summarize revenue by region for q4 and suggest the main anomalies'),
        'Summarize revenue by region for q4 and'
    );

    assert.strictEqual(
        threadTitles.buildAutoTitle('/draw a clean onboarding illustration for a finance dashboard'),
        'A clean onboarding illustration for a'
    );

    assert.strictEqual(threadTitles.buildAutoTitle('   '), 'New chat');
}

function testResolveThreadTitle() {
    assert.strictEqual(
        threadTitles.resolveThreadTitle({
            autoTitle: 'summarize the uploaded report',
            customTitle: 'Weekly review'
        }),
        'Weekly review'
    );

    assert.strictEqual(
        threadTitles.resolveThreadTitle({
            autoTitle: 'summarize the uploaded report'
        }),
        'Summarize the uploaded report'
    );
}

function run() {
    testBuildAutoTitle();
    testResolveThreadTitle();
    console.log('r7chat_thread_title.test.js: ok');
}

run();
