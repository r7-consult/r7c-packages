/**
 * @file SetLockValue_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiDrawing.SetLockValue
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the lock value for a specified lock type of a drawing.
 * It creates a shape, sets its size and position, sets a lock value for "noSelect", and then displays the lock value.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить значение блокировки для указанного типа блокировки рисунка.
 * Он создает фигуру, устанавливает ее размер и положение, устанавливает значение блокировки для "noSelect", а затем отображает значение блокировки.
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
        // This example sets the lock value to the specified lock type of the current drawing.
        
        // How to set a lock type of a drawing.
        
        // Create a drawing, set its lock value and display it in the worksheet.
        
        let worksheet = Api.GetActiveSheet();
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        let drawing = worksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        drawing.SetSize(120 * 36000, 70 * 36000);
        drawing.SetPosition(0, 2 * 36000, 1, 3 * 36000);
        drawing.SetLockValue("noSelect", true);
        let lockValue = drawing.GetLockValue("noSelect");
        worksheet.GetRange("A1").SetValue("This drawing cannot be selected: " + lockValue);
        
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
