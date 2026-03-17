/**
 * @file GetClassType_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiSchemeColor.GetClassType
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the class type of an ApiSchemeColor object.
 * It creates a scheme color, uses it in a solid fill for a shape, and then displays the class type of the scheme color object.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить тип класса объекта ApiSchemeColor.
 * Он создает цвет схемы, использует его в сплошной заливке для фигуры, а затем отображает тип класса объекта цвета схемы.
 *
 * @returns {void}
 *
 * @see https://r7-consult.com/
 */

(function() {
    'use strict';
    
    try {
        // Initialize R7 Office API
        const api = Api;
        if (!api) {
            throw new Error('R7 Office API not available');
        }
        
        // Original code enhanced with error handling:
        // This example gets a class type and inserts it into the document.
        
        // How to get a class type of ApiSchemeColor.
        
        // Get a class type of ApiSchemeColor and display it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let schemeColor = Api.CreateSchemeColor("dk1");
        let fill = Api.CreateSolidFill(schemeColor);
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        worksheet.AddShape("curvedUpArrow", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 1, 3 * 36000);
        let classType = schemeColor.GetClassType();
        worksheet.SetColumnWidth(0, 15);
        worksheet.SetColumnWidth(1, 10);
        worksheet.GetRange("A1").SetValue("Class Type = ");
        worksheet.GetRange("B1").SetValue(classType);
        
        // Success notification
        console.log('Macro executed successfully');
        
    } catch (error) {
        console.error('Macro execution failed:', error.message);
        // Optional: Show error to user
        if (typeof Api !== 'undefined' && Api.GetActiveSheet) {
            const sheet = Api.GetActiveSheet();
            if (sheet) {
                sheet.GetRange('A1').SetValue('Error: ' + error.message);
            }
        }
    }
})();
