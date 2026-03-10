/**
 * @file GetMailMergeData_macroR7.js
 * @brief R7 Office JavaScript Macro - Api.GetMailMergeData
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to retrieve mail merge data from the spreadsheet.
 * It populates a sample dataset for mail merge (email, greeting, first name, last name),
 * then retrieves the mail merge data using an index, and displays it in cell A5.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить данные для слияния из электронной таблицы.
 * Он заполняет примерный набор данных для слияния (адрес электронной почты, приветствие, имя, фамилия),
 * затем извлекает данные для слияния по индексу и отображает их в ячейке A5.
 *
 * @param {number} index - The zero-based index of the mail merge data to retrieve. (Нулевой индекс данных для слияния для получения.)
 * @returns {string} The mail merge data as a string. (Данные для слияния в виде строки.)
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
        
        // Set column width for better readability
        // Установка ширины столбца для лучшей читаемости
        worksheet.SetColumnWidth(0, 20);
        
        // Populate sample mail merge data
        // Заполнение примерных данных для слияния
        worksheet.GetRange("A1").SetValue("Email address");
        worksheet.GetRange("B1").SetValue("Greeting");
        worksheet.GetRange("C1").SetValue("First name");
        worksheet.GetRange("D1").SetValue("Last name");
        worksheet.GetRange("A2").SetValue("user1@example.com");
        worksheet.GetRange("B2").SetValue("Dear");
        worksheet.GetRange("C2").SetValue("John");
        worksheet.GetRange("D2").SetValue("Smith");
        worksheet.GetRange("A3").SetValue("user2@example.com");
        worksheet.GetRange("B3").SetValue("Hello");
        worksheet.GetRange("C3").SetValue("Kate");
        worksheet.GetRange("D3").SetValue("Cage");
        
        // Retrieve mail merge data by index (e.g., the first record)
        // Получение данных для слияния по индексу (например, первая запись)
        let mailMergeData = Api.GetMailMergeData(0);
        
        // Display the retrieved mail merge data in cell A5
        // Отображение полученных данных для слияния в ячейке A5
        worksheet.GetRange("A5").SetValue("Mail merge data: " + mailMergeData);
        
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
