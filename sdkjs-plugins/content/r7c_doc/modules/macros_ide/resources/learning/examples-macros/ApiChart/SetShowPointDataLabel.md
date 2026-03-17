/**
 * @file SetShowPointDataLabel_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiChart.SetShowPointDataLabel
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to show or hide data labels for a specific data point in a chart.
 * It creates a 3D bar chart, sets its title, and then sets the data label for the first point of the second series to show only the value.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как показать или скрыть метки данных для определенной точки данных на диаграмме.
 * Он создает трехмерную гистограмму, устанавливает ее заголовок, а затем устанавливает метку данных для первой точки второго ряда для отображения только значения.
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
        // This example specifies the show options for the chart data labels.
        
        // How to hide chart data labels using options.
        
        // Show only values as chart lables with series index equal to 1 from the array of the data used to build the chart from.
        
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
        let chart = worksheet.AddChart("'Sheet1'!$A$1:$D$3", true, "bar3D", 2, 100 * 36000, 70 * 36000, 0, 2 * 36000, 5, 3 * 36000);
        chart.SetTitle("Financial Overview", 13);
        chart.SetShowPointDataLabel(1, 0, false, false, true, false);
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
