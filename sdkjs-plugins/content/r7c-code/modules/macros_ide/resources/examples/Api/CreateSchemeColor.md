/**
 * @file CreateSchemeColor_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.CreateSchemeColor
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to create a scheme color using `Api.CreateSchemeColor()`.
 * It creates a solid fill with a scheme color (dark 1), and then adds a shape
 * to the active worksheet with this fill and a no-fill stroke.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как создать цвет схемы с помощью `Api.CreateSchemeColor()`.
 * Он создает сплошную заливку с цветом схемы (темный 1), а затем добавляет фигуру
 * на активный лист с этой заливкой и обводкой без заливки.
 *
 * @param {string} schemeColorName - The name of the scheme color to create (e.g., "dk1", "lt1"). (Имя цвета схемы для создания.)
 * @returns {ApiColor} A new ApiColor object representing the scheme color. (Новый объект ApiColor, представляющий цвет схемы.)
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
        
        // Create a scheme color (dark 1)
        // Создание цвета схемы (темный 1)
        let schemeColor = Api.CreateSchemeColor("dk1");
        
        // Create a solid fill using the scheme color
        // Создание сплошной заливки с использованием цвета схемы
        let fill = Api.CreateSolidFill(schemeColor);
        
        // Create a no-fill stroke
        // Создание обводки без заливки
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        
        // Add a shape to the worksheet with the defined fill and stroke
        // Добавление фигуры на лист с определенной заливкой и обводкой
        worksheet.AddShape("curvedUpArrow", 60 * 36000, 35 * 36000, fill, stroke, 0, 2 * 36000, 1, 3 * 36000);
        
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
