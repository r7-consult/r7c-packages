/**
 * @file GetClassType_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiFill.GetClassType
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the class type of an ApiFill object.
 * It creates a linear gradient fill, applies it to a shape, and then displays the class type of the fill object.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить тип класса объекта ApiFill.
 * Он создает линейную градиентную заливку, применяет ее к фигуре, а затем отображает тип класса объекта заливки.
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
        // This example gets a class type and inserts it into the document.
        
        // How to get a class type of ApiFill.
        
        // Get a class type of ApiFill and display it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let gradientStop1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0);
        let gradientStop2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);
        let fill = Api.CreateLinearGradientFill([gradientStop1, gradientStop2], 5400000);
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        worksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 1, 3 * 36000);
        let classType = fill.GetClassType();
        worksheet.SetColumnWidth(0, 15);
        worksheet.SetColumnWidth(1, 10);
        worksheet.GetRange("A1").SetValue("Class Type = " + classType);
        
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
