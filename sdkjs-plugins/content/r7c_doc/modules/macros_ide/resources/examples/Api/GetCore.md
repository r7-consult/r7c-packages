/**
 * @file GetCore_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.GetCore
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to retrieve and manage core document metadata using `Api.GetCore()`.
 * It sets various metadata properties such as category, content status, creation date, author,
 * description, keywords, language, and more. Then, it retrieves these properties and displays
 * them within a shape added to the active worksheet.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получать и управлять основными метаданными документа с помощью `Api.GetCore()`.
 * Он устанавливает различные свойства метаданных, такие как категория, статус содержимого, дата создания, автор,
 * описание, ключевые слова, язык и многое другое. Затем он извлекает эти свойства и отображает
 * их внутри фигуры, добавленной на активный лист.
 *
 * @returns {ApiCore} An object representing the core properties of the document. (Объект, представляющий основные свойства документа.)
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
        
        const worksheet = Api.GetActiveSheet();
        
        // Get the core API object
        // Получение основного объекта API
        const core = Api.GetCore();
        
        // Set various core document properties
        // Установка различных основных свойств документа
        core.SetCategory("Examples");
        core.SetContentStatus("Final");
        core.SetCreated(new Date("2000-01-20"));
        core.SetCreator("John Smith");
        core.SetDescription("Sample spreadsheet demonstrating ApiCore methods.");
        core.SetIdentifier("#ID42");
        core.SetKeywords("Spreadsheet; ApiCore; Metadata");
        core.SetLanguage("en-US");
        core.SetLastModifiedBy("Mark Pottato");
        core.SetLastPrinted(new Date());
        core.SetModified(new Date("1990-03-10"));
        core.SetRevision("Rev. C");
        core.SetSubject("Spreadsheet Metadata Showcase");
        core.SetTitle("My Spreadsheet Title");
        core.SetVersion("v9.0");
        
        // Retrieve the core document properties
        // Извлечение основных свойств документа
        let category = core.GetCategory();
        let contentStatus = core.GetContentStatus();
        let createdDate = core.GetCreated().toDateString();
        let creator = core.GetCreator();
        let description = core.GetDescription();
        let identifier = core.GetIdentifier();
        let keywords = core.GetKeywords();
        let language = core.GetLanguage();
        let lastModifiedBy = core.GetLastModifiedBy();
        let lastPrintedDate = core.GetLastPrinted().toDateString();
        let lastModifiedDate = core.GetModified().toDateString();
        let revision = core.GetRevision();
        let subject = core.GetSubject();
        let title = core.GetTitle();
        let version = core.GetVersion();
        
        // Create a solid fill and a no-fill stroke for the shape
        // Создание сплошной заливки и обводки без заливки для фигуры
        let fill = Api.CreateSolidFill(Api.CreateRGBColor(100, 50, 200));
        let stroke = Api.CreateStroke(0, Api.CreateNoFill());
        
        // Add a rectangular shape to the worksheet
        // Добавление прямоугольной фигуры на лист
        const shape = worksheet.AddShape(
        	"rect",
        	100 * 36000, 100 * 36000,
        	fill, stroke,
        	0, 0, 3, 0
        );
        
        // Get the content of the shape and its first paragraph
        // Получение содержимого фигуры и ее первого абзаца
        let docContent = shape.GetContent();
        let paragraph = docContent.GetElement(0);
        
        // Add the core property values to the shape's text content
        // Добавление значений основных свойств в текстовое содержимое фигуры
        paragraph.AddText("Category: " + category + "\n");
        paragraph.AddText("Content Status: " + contentStatus + "\n");
        paragraph.AddText("Created: " + createdDate + "\n");
        paragraph.AddText("Creator: " + creator + "\n");
        paragraph.AddText("Description: " + description + "\n");
        paragraph.AddText("Identifier: " + identifier + "\n");
        paragraph.AddText("Keywords: " + keywords + "\n");
        paragraph.AddText("Language: " + language + "\n");
        paragraph.AddText("Last Modified By: " + lastModifiedBy + "\n");
        paragraph.AddText("Last Printed: " + lastPrintedDate + "\n");
        paragraph.AddText("Last Modified: " + lastModifiedDate + "\n");
        paragraph.AddText("Revision: " + revision + "\n");
        paragraph.AddText("Subject: " + subject + "\n");
        paragraph.AddText("Title: " + title + "\n");
        paragraph.AddText("Version: " + version);
        
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
