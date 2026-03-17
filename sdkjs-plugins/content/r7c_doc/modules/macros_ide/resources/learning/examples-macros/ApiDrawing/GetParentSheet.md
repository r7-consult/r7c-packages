/**
 * @file GetParentSheet_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiDrawing.GetParentSheet
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get the parent sheet of a drawing object.
 * It creates a shape, gets its parent sheet, and then displays the name of the parent sheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить родительский лист объекта рисования.
 * Он создает фигуру, получает ее родительский лист, а затем отображает имя родительского листа.
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
        // This example show get drawing parent sheet.
        
        // How to know a parent sheet of a shape.
        
        // Get a shape's parent sheet and display it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let shape = worksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        let parentSheet = shape.GetParentSheet();
        let content = shape.GetDocContent();
        let paragraph = content.GetElement(0);
        paragraph.AddText("Parent sheet name is " + parentSheet.GetName());
        
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
