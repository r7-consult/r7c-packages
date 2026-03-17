/**
 * R7 Office Parameterized Macro Example
 * 
 * This macro demonstrates parameter usage:
 * - Reading parameters passed from command line
 * - Using default values when parameters are not provided
 * - Conditional logic based on parameters
 * 
 * Usage: R7 Office-macro-cli -p name=Alice -p color=blue -p count=5 parameterized-macro.js
 */

(function() {
    // Read parameters with defaults
    var name = (typeof PARAMS !== 'undefined' && PARAMS.name) ? PARAMS.name : 'World';
    var color = (typeof PARAMS !== 'undefined' && PARAMS.color) ? PARAMS.color : 'red';
    var count = (typeof PARAMS !== 'undefined' && PARAMS.count) ? parseInt(PARAMS.count) : 3;
    
    console.log("=== Parameterized Macro Execution ===");
    console.log("Name: " + name);
    console.log("Color: " + color);
    console.log("Count: " + count);
    
    // Perform actions based on parameters
    if (Api.GetActiveSheet) {
        var oSheet = Api.GetActiveSheet();
        
        // Create colored ranges based on count
        for (var i = 0; i < count; i++) {
            var cellRange = "A" + (i + 1) + ":C" + (i + 1);
            var oRange = oSheet.GetRange(cellRange);
            
            // Set different colors based on parameter
            switch (color.toLowerCase()) {
                case 'red':
                    oRange.SetFillColor(255, 0, 0);
                    break;
                case 'blue':
                    oRange.SetFillColor(0, 0, 255);
                    break;
                case 'green':
                    oRange.SetFillColor(0, 255, 0);
                    break;
                case 'yellow':
                    oRange.SetFillColor(255, 255, 0);
                    break;
                default:
                    oRange.SetFillColor(128, 128, 128); // Gray
            }
            
            console.log("Filled range " + cellRange + " with " + color + " color");
        }
    }
    
    // Show personalized message
    var message = "Hello, " + name + "! Created " + count + " " + color + " ranges.";
    Api.ShowMessage("Parameterized Macro", message);
    
    console.log("Macro completed successfully!");
})();