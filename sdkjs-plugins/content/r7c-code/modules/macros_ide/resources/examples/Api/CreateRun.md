/**
 * @file CreateRun_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.CreateRun
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to create and manipulate text runs within a paragraph.
 * It adds a shape to the active worksheet, then creates two text runs: one with default
 * formatting and another with a specified font family ("Comic Sans MS"). Both runs are
 * added to a paragraph within the shape's content.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как создавать и манипулировать текстовыми блоками (run) внутри абзаца.
 * Он добавляет фигуру на активный лист, затем создает два текстовых блока: один с форматированием по умолчанию
 * и другой с указанным семейством шрифтов ("Comic Sans MS"). Оба блока добавляются
 * в абзац внутри содержимого фигуры.
 *
 * @returns {ApiRun} A new empty ApiRun object. (Новый пустой объект ApiRun.)
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
        
        // Create fill and stroke for the shape
        // Создание заливки и обводки для фигуры
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(255, 111, 61));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        
        // Add a shape to the worksheet
        // Добавление фигуры на лист
        let shape = worksheet.AddShape("flowChartOnlineStorage", 120 * 36000, 70 * 36000, fill, stroke, 0, 2 * 36000, 0, 3 * 36000);
        
        // Get the document content of the shape and its first paragraph
        // Получение содержимого документа фигуры и ее первого абзаца
        let docContent = shape.GetContent();
        let paragraph = docContent.GetElement(0);
        
        // Create the first run and add text to it
        // Создание первого текстового блока и добавление в него текста
        let run1 = Api.CreateRun();
        run1.AddText("This is just a sample text. ");
        paragraph.AddElement(run1);
        
        // Create the second run, set its font family, and add text to it
        // Создание второго текстового блока, установка его семейства шрифтов и добавление в него текста
        let run2 = Api.CreateRun();
        run2.SetFontFamily("Comic Sans MS");
        run2.AddText("This is a text run with the font family set to 'Comic Sans MS'.");
        paragraph.AddElement(run2);
        
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
