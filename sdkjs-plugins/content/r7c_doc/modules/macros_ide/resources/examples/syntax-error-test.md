/**
 * Syntax Error Test
 * 
 * This file contains intentional syntax errors to test
 * the enhanced error reporting of the V8/JSC engines.
 */

(function() {
    console.log("This should not execute due to syntax errors below");
    
    // Syntax error 1: Missing closing parenthesis
    var result = Math.sqrt(16;
    
    // This line should not be reached
    console.log("Result:", result);
    
    // Syntax error 2: Invalid function syntax  
    function badFunction( {
        return "This is invalid";
    }
    
})();