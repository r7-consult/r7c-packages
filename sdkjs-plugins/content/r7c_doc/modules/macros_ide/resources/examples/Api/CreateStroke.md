/**
 * @file CreateStroke_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.CreateStroke
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to create a stroke (outline) for a shape with a solid fill.
 * It defines two gradient stops and a linear gradient fill, then creates a solid fill for the stroke.
 * Finally, it adds a shape to the active worksheet with the defined fill and stroke.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как создать обводку (контур) для фигуры со сплошной заливкой.
 * Он определяет две точки градиента и линейную градиентную заливку, затем создает сплошную заливку для обводки.
 * Наконец, он добавляет фигуру на активный лист с определенной заливкой и обводкой.
 *
 * @param {number} width - The width of the stroke in EMU (English Metric Units). (Ширина обводки в EMU.)
 * @param {ApiFill} fill - The fill object for the stroke. (Объект заливки для обводки.)
 * @returns {ApiStroke} A new ApiStroke object. (Новый объект ApiStroke.)
 *
 * @see https://r7-consult.ru/
 */

(function() {
    'use strict';
    
    try {
        // Initialize R7 Office API
        // Инициализация API R7 Office
        const api = Api;
        if (!api) {
            throw new Error('R7 Office API not available');
        }
        
        // Get the active worksheet
        // Получение активного листа
        let worksheet = Api.GetActiveSheet();
        
        // Create gradient stops for the shape's fill
        // Создание точек градиента для заливки фигуры
        let gs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0);
        let gs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);
        
        // Create a linear gradient fill for the shape
        // Создание линейной градиентной заливки для фигуры
        let fill = Api.CreateLinearGradientFill([gs1, gs2], 5400000);
        
        // Create a solid fill for the stroke
        // Создание сплошной заливки для обводки
        let fill1 = Api.CreateSolidFill(Api.CreateRGBColor(51, 51, 51));
        
        // Create a stroke with a specified width and fill
        // Создание обводки с указанной шириной и заливкой
        let stroke = Api.CreateStroke(3 * 36000, fill1);
        
        // Add a shape to the worksheet with the defined fill and stroke
        // Добавление фигуры на лист с определенной заливкой и обводкой
        worksheet.AddShape("flowChartOnlineStorage", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 1, 3 * 36000);
        
        // Success notification
        // Уведомление об успешном выполнении
        console.log('Macro executed successfully');
        
    } catch (error) {
        // Error handling
        // Обработка ошибок
        console.error('Macro execution failed:', error.message);
        // Optional: Show error to user in cell A1 if API is available
        // Опционально: Показать ошибку пользователю в ячейке A1, если API доступен
        if (typeof Api !== 'undefined' && Api.GetActiveSheet) {
            const sheet = Api.GetActiveSheet();
            if (sheet) {
                sheet.GetRange('A1').SetValue('Error: ' + error.message);
            }
        }
    }
})();
