/**
 * R7 Office Error Examples
 * 
 * This file contains intentional errors to test error handling:
 * - Syntax errors
 * - Runtime errors
 * - API misuse
 * 
 * Use with: R7 Office-macro-cli -c error-examples.js
 */

// Uncomment different sections to test different error types

// === SYNTAX ERRORS ===
// These will be caught during syntax checking:

// 1. Missing closing parenthesis
// console.log("This will cause a syntax error";

// 2. Missing closing brace
// (function() {
//     console.log("Missing closing brace");
// )();

// 3. Invalid variable name
// var 123invalid = "This is not a valid variable name";

// 4. Unclosed string
// var message = "This string is not closed;

// === RUNTIME ERRORS ===
// These will be caught during execution:

(function() {
    console.log("=== Runtime Error Examples ===");
    
    // 1. Undefined variable (comment out to test)
    // console.log(undefinedVariable);
    
    // 2. Calling method on undefined object
    // var obj = undefined;
    // obj.method();
    
    // 3. Invalid function call
    // nonExistentFunction();
    
    // 4. Type error
    // var num = 123;
    // num.push("This will fail");
    
    console.log("No runtime errors in this execution");
})();

// === API MISUSE EXAMPLES ===
// These will trigger warnings:

(function() {
    console.log("=== API Misuse Examples ===");
    
    // Using browser APIs that don't exist in R7 Office
    // alert("This should use Api.ShowMessage() instead");
    // document.getElementById("test");
    // window.location.href = "http://example.com";
    
    // Unknown R7 Office API calls
    // Api.NonExistentMethod();
    // Api.GetActiveSheet().UnknownMethod();
    
    console.log("API usage examples completed");
})();