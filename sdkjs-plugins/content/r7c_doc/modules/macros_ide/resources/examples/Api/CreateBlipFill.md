/**
 * @file CreateBlipFill_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.CreateBlipFill
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to create a blip fill (image fill) to apply to an object.
 * It creates a blip fill using a specified image URL and a tiling mode, then adds a shape
 * to the active worksheet with this blip fill and a no-fill stroke.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как создать растровую заливку (заливку изображением) для применения к объекту.
 * Он создает растровую заливку, используя указанный URL изображения и режим мозаики, затем добавляет фигуру
 * на активный лист с этой растровой заливкой и обводкой без заливки.
 *
 * @param {string} imageUrl - The URL of the image to use for the blip fill. (URL изображения для использования в растровой заливке.)
 * @param {string} tilingMode - The tiling mode for the image (e.g., "tile", "stretch"). (Режим мозаики для изображения.)
 * @returns {ApiFill} A new ApiFill object representing a blip fill. (Новый объект ApiFill, представляющий растровую заливку.)
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
        
        // Create a blip fill using an image URL and tiling mode
        // Создание растровой заливки с использованием URL изображения и режима мозаики
        let fill = Api.CreateBlipFill("https://api.R7 Office.com/content/img/docbuilder/examples/icon_DocumentEditors.png", "tile");
        
        // Create a no-fill stroke
        // Создание обводки без заливки
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        
        // Add a shape to the worksheet with the defined blip fill and stroke
        // Добавление фигуры на лист с определенной растровой заливкой и обводкой
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
