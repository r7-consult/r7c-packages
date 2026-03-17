/**
 * @file GetAllUsers_macroR7.js
 * @brief R7 Office JavaScript Macro - ApiProtectedRange.GetAllUsers
 * @author R7-Consult
 * @version 1.0.0
 * @date July 15, 2025
 *
 * @description
 * This macro demonstrates how to get all users of a protected range.
 * It adds a protected range to the worksheet, adds two users to it, and then displays the name of the first user.
 *
 * @description (Russian)
 * Этот макрос демонстрирует, как получить всех пользователей защищенного диапазона.
 * Он добавляет защищенный диапазон на лист, добавляет в него двух пользователей, а затем отображает имя первого пользователя.
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
        // This example gets all users of a protected range.
        
        // How to get an array of users of a protected range.
        
        // Get an active sheet, add protected range to it and diplay its first user. 
        
        let worksheet = Api.GetActiveSheet();
        worksheet.GetRange("A1").SetValue("1");
        worksheet.GetRange("B1").SetValue("2");
        worksheet.GetRange("C1").SetValue("3");
        worksheet.AddProtectedRange("Protected range", "$A$1:$C$1");
        let protectedRange = worksheet.GetProtectedRange("Protected range");
        protectedRange.AddUser("uid-1", "John Smith", "CanEdit");
        protectedRange.AddUser("uid-2", "Mark Potato", "CanView");
        let users = protectedRange.GetAllUsers();
        worksheet.GetRange("A3").SetValue(users[0].GetName());
        
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
