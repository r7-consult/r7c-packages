---
interactive:
  version: 1
  enabled: true
  lessonId: "macro-templates-28-replacetext"
  title: "Урок 28: Шаблон «Умная Замена» (ReplaceTextSmart)"
  draft: false
  targetEditor: "cell"
  starterCode: |
    (function() {
        'use strict';
        
        try {
            const api = Api;
            if (!api) throw new Error('API not available');
            
            let worksheet = Api.GetActiveSheet();
            
            // 1. Имитация старых данных
            worksheet.GetRange("A1").SetValue("Old Text 1");
            worksheet.GetRange("A2").SetValue("Old Text 2");
            worksheet.GetRange("C1").SetValue("Unrelated Text"); // Не выделяем
            
            // ВАША ЗАДАЧА ЗДЕСЬ (Управление выделением):
            // 2. Получите целевой диапазон для изменения (A1:A2)
            let targetRange = worksheet.GetRange("A1:A2");
            
            // 3. Выделите этот диапазон программно (Метод .Select())
            
            
            // 4. Используйте массив из двух новых значений для умной замены
            // Вызовите Api.ReplaceTextSmart(  ["New Value A1", "New Value A2"]  );
            
            
        } catch (error) {
            console.error('Macro execution failed:', error.message);
            if (typeof Api !== 'undefined' && Api.GetActiveSheet) {
                let sheet = Api.GetActiveSheet();
                if (sheet) sheet.GetRange('E1').SetValue('Error: ' + error.message);
            }
        }
    })();
  checks:
    - type: "range_values"
      range: "A1"
      expected: "New Value A1"
    - type: "range_values"
      range: "A2"
      expected: "New Value A2"
  beforeScript: |
    (function() {
    let sheet = Api.GetActiveSheet();
    sheet.GetRange("A1").SetValue("ООО Ромашка");
    sheet.GetRange("A2").SetValue("ИП Ромашка Плюс");
    sheet.GetRange("D1:H1").Merge(false);
    sheet.GetRange("D1:H1").SetValue("💡 Используйте ReplaceTextSmart, чтобы заменить 'Ромашка' на 'Василек'.");
    })();
---

# Урок 28. Анализ Эталона: Шаблон «Умная Замена» (ReplaceTextSmart)

Эталонный скрипт `ReplaceTextSmart_macroR7.js` открывает перед нами тайный мир так называемых **"Глобальных курсорных операций"**. 

В VBA почти все разработчики использовали `Selection.Replace(Что, На_что)`. Это считалось дурным тоном, так как макрос "прыгал" по экрану пользователя.
Но в Р7-Офис есть интеллектуальный метод глобальной замены `Api.ReplaceTextSmart(Массив_Строк)`. Он работает не с конкретно взятым `Range`, а с **Текущим выделением пользователя (Selection)**.

## 🎯 Архитектура Умной Замены

Механика работы этого шаблона кардинально отличается от `SetValue`:
1. Вы перехватываете управление интерфейсом пользователя (Программный Курсор).
2. Вы даете команду `myRange.Select()`. Таблица визуально выделит ячейки синим фоном!
3. Вы не перебираете циклом ячейки! Вы даете глобальную команду `Api.ReplaceTextSmart(["Текст1", "Текст2"])`.
4. Движок сам разберется, как "залить" массив строк в выделенную вами прямоугольную область.

Это так называемый "Непрерывный поток значений" (Stream Replacement).

---

## 📝 Постановка задачи

В заготовке кода вы уже видите архитектурную броню эталона Р7.
Мы сымитировали ситуацию: в `A1` и `A2` находится старый текст. 
Ваша задача (Ядро алгоритма):

1. Вызовите встроенный метод интерфейса `.Select()` на объекте `targetRange`. Это физически подвинет "мышь" пользователя на диапазон `A1:A2`.
2. Передайте глобальному API команду на заливку новых значений в это Выделение: `Api.ReplaceTextSmart(["New Value A1", "New Value A2"])`.
   *Обратите внимание: мы передаем массив `[ ]` внутрь функции. Это называется "Сброс блока данных" (Block drop).*

## 🧠 Глубинка API

Почему бы не использовать обычный цикл и `.SetValue()`? 
В 90% случаев мы так и делаем (Урок 23 "Batch Write"). 
Но метод `ReplaceTextSmart` — это UI-механизм (для плагинов). Если вы пишете визуальный помощник (например, плагин "Склонятор Фамилий"), который берет выделенный пользователем текст, склоняет его на внешнем сервере, и заменяет обратно в таблице:
Вы не знаете адреса `Range`! Пользователь сам выделил его мышкой. 
Вам остается только прочесть его, отправить на сервер, а ответ вернуть глобальной командой `Api.ReplaceTextSmart(Ответ_Массив)`. Это эталонный паттерн написания интерактивных Плагинов (Interactive Plugins) на движке Р7.

(Скопируйте структуру из Эталона в редактор и нажмите 'Проверить' для эмуляции замены).
