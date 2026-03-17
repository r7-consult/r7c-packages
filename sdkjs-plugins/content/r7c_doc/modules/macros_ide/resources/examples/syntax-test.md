/**
 * R7 Office Syntax Test Macro
 * 
 * This macro demonstrates various syntax patterns and potential issues:
 * - Correct R7 Office API usage
 * - Common mistakes that trigger warnings
 * - JavaScript syntax validation
 * 
 * Use with: R7 Office-macro-cli -c syntax-test.js
 */

(function() {
    console.log("=== Syntax Test Macro ===");
    
    // ✓ Correct R7 Office API usage
    var oSheet = Api.GetActiveSheet();
    var oRange = oSheet.GetRange("A1:B2");
    oRange.SetFillColor(255, 255, 0);
    
    // ✓ Proper console usage
    console.log("This is correct usage");
    
    // ✓ Correct message display
    Api.ShowMessage("Test", "This is the correct way to show messages");
    
    // ⚠️ These lines will trigger warnings during syntax check:
    
    // Warning: Unknown R7 Office API call
    // Api.UnknownMethod();
    
    // Warning: Use Api.GetActiveDocument() instead of document
    // document.getElementById('test');
    
    // Warning: window object not available
    // window.location.href = 'http://example.com';
    
    // Warning: Use Api.ShowMessage() instead of alert()
    // alert('This will trigger a warning');
    
    // ✓ Proper parameter usage
    if (typeof PARAMS !== 'undefined') {
        console.log("Parameters available:", JSON.stringify(PARAMS));
    }
    
    // ✓ Proper document path usage
    if (typeof DOCUMENT_PATH !== 'undefined') {
        console.log("Document path:", DOCUMENT_PATH);
    }
    
    // ✓ Complex JavaScript that should parse correctly
    var data = {
        name: "Test",
        values: [1, 2, 3, 4, 5],
        process: function(item) {
            return item * 2;
        }
    };
    
    var processed = data.values.map(data.process);
    console.log("Processed values:", processed);
    
    // ✓ Conditional logic
    if (processed.length > 0) {
        console.log("Processing completed with " + processed.length + " items");
    }
    
    console.log("Syntax test completed!");
})();