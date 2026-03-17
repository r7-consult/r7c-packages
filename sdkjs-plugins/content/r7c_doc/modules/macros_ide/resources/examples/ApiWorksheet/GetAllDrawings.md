/**
 * @file GetAllDrawings_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiWorksheet.GetAllDrawings
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get all drawings from the sheet.
 * It creates a chart (which is a type of drawing), sets its properties, and then retrieves all drawings from the worksheet to modify the first one.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить все рисунки с листа.
 * Он создает диаграмму (которая является типом рисунка), устанавливает ее свойства, а затем извлекает все рисунки из листа для изменения первого.
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
        // This example shows how to get all drawings from the sheet.
        
        // How to get all drawings.
        
        // Get all drawings as an array.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("B1").SetValue(2014);
        worksheet.GetRange("C1").SetValue(2015);
        worksheet.GetRange("D1").SetValue(2016);
        worksheet.GetRange("A2").SetValue("Projected Revenue");
        worksheet.GetRange("A3").SetValue("Estimated Costs");
        worksheet.GetRange("B2").SetValue(200);
        worksheet.GetRange("B3").SetValue(250);
        worksheet.GetRange("C2").SetValue(240);
        worksheet.GetRange("C3").SetValue(260);
        worksheet.GetRange("D2").SetValue(280);
        worksheet.GetRange("D3").SetValue(280);
        let drawing = worksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);
        drawing.SetTitle("Financial Overview", 13);
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
        drawing.SetSeriesFill(fill, 0, false);
        fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        drawing.SetSeriesFill(fill, 1, false);
        let drawings = worksheet.GetAllDrawings();
        drawings[0].SetSize(150 * 36000, 100 * 36000);
        
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
