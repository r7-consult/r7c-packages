/**
 * @file GetClassType_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiCustomProperties.GetClassType
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the class type of the ApiCustomProperties object.
 * It retrieves the class type and displays it in a shape on the worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить тип класса объекта ApiCustomProperties.
 * Он извлекает тип класса и отображает его в фигуре на листе.
 *
 * @returns {void}
 *
 * @see https://r7-consult.ru/
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
        // This example demonstrates how to get the class type of ApiCustomProperties.
        
        const worksheet = Api.GetActiveSheet();
        const customProps = Api.GetCustomProperties();
        const classType = customProps.GetClassType();
        
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(0, 100, 200));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        const shape = worksheet.AddShape("rect", 100 * 36000, 50 * 36000, fill, stroke, 0, 0, 5, 0);
        
        let paragraph = shape.GetDocContent().GetElement(0);
        paragraph.AddText("ApiCustomProperties class type: " + classType);
        
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
