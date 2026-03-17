/**
 * @file CreateGradientStop_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.CreateGradientStop
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to create a gradient stop, which is a point in a gradient
 * that specifies a color and its position. It defines two gradient stops with specific
 * RGB colors and positions, then uses them to create a linear gradient fill.
 * Finally, it adds a shape to the active worksheet with this fill and a no-fill stroke.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как создать точку градиента, которая является точкой в градиенте,
 * определяющей цвет и его положение. Он определяет две точки градиента с определенными
 * цветами RGB и позициями, затем использует их для создания линейной градиентной заливки.
 * Наконец, он добавляет фигуру на активный лист с этой заливкой и обводкой без заливки.
 *
 * @param {ApiColor} color - The color of the gradient stop. (Цвет точки градиента.)
 * @param {number} position - The position of the gradient stop (0-100000, where 0 is the start and 100000 is the end). (Позиция точки градиента (0-100000, где 0 - начало, а 100000 - конец).)
 * @returns {ApiGradientStop} A new ApiGradientStop object. (Новый объект ApiGradientStop.)
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
        
        // Create gradient stops with RGB colors and positions
        // Создание точек градиента с цветами RGB и позициями
        let gs1 = Api.CreateGradientStop(Api.CreateRGBColor(255, 213, 191), 0);
        let gs2 = Api.CreateGradientStop(Api.CreateRGBColor(255, 111, 61), 100000);
        
        // Create a linear gradient fill using the gradient stops
        // Создание линейной градиентной заливки с использованием точек градиента
        let fill = Api.CreateLinearGradientFill([gs1, gs2], 5400000);
        
        // Create a no-fill stroke
        // Создание обводки без заливки
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        
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
