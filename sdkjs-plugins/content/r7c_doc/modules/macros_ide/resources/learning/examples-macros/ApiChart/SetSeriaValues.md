/**
 * @file SetSeriaValues_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiChart.SetSeriaValues
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to set the values for a series in a chart.
 * It creates a 3D bar chart, sets its title, and then sets the values for the second series from a different range.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как установить значения для ряда в диаграмме.
 * Он создает трехмерную гистограмму, устанавливает ее заголовок, а затем устанавливает значения для второго ряда из другого диапазона.
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
        // This example sets values from the specified range to the specified series.
        
        // How to add values to series from the indicated range using addresses.
        
        // Fill series with values obtained from the worksheet cells.
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("B1").SetValue(2014);
        worksheet.GetRange("C1").SetValue(2015);
        worksheet.GetRange("D1").SetValue(2016);
        worksheet.GetRange("A2").SetValue("Projected Revenue");
        worksheet.GetRange("A3").SetValue("Estimated Costs");
        worksheet.GetRange("B2").SetValue(200);
        worksheet.GetRange("B3").SetValue(250);
        worksheet.GetRange("B4").SetValue(260);
        worksheet.GetRange("C2").SetValue(240);
        worksheet.GetRange("C3").SetValue(260);
        worksheet.GetRange("C4").SetValue(270);
        worksheet.GetRange("D2").SetValue(280);
        worksheet.GetRange("D3").SetValue(280);
        worksheet.GetRange("D4").SetValue(300);
        let chart = worksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 7, 3 * 36000);
        chart.SetTitle("Financial Overview", 13);
        chart.SetSeriaValues("'Sheet1'!$B$4:$D$4", 1);
        chart.SetShowPointDataLabel(1, 0, false, false, true, false);
        chart.SetShowPointDataLabel(1, 1, false, false, true, false);
        chart.SetShowPointDataLabel(1, 2, false, false, true, false);
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
        chart.SetSeriesFill(fill, 0, false);
        fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        chart.SetSeriesFill(fill, 1, false);
        
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
