---
interactive:
  version: 1
  enabled: true
  lessonId: "macro-basics-03-read-write"
  title: "Урок 3: Практика. Чтение и запись значений"
  draft: false
  targetEditor: "cell"
  starterCode: |
    (function() {
        'use strict';
        try {
            let sheet = Api.GetActiveSheet();
            
            // ВАША ЗАДАЧА:
            // 1. Считайте три куска пароля из ячеек D5, F12 и Z20 (используйте .GetValue()).
            // 2. Склейте их в одну строку через тире (например: "СЛОВО1-СЛОВО2-СЛОВО3").
            // 3. Запишите итоговый пароль в ячейку A1.
            
            
            
        } catch (error) {
            console.error(error.message);
        }
    })();
  beforeScript: |
    (function() {
        let sheet = Api.GetActiveSheet();
        sheet.GetRange("A1:Z100").Clear();
        sheet.GetRange("D5").SetValue("R7");
        sheet.GetRange("F12").SetValue("MACRO");
        sheet.GetRange("Z20").SetValue("HACK");
        sheet.GetRange("A3:H3").Merge(false);
        sheet.GetRange("A3:H3").SetValue("Операция 'Шифр': Найдите 3 части пароля в D5, F12 и Z20, соедините их через тире и запишите в A1!");
        sheet.GetRange("A3:H3").SetFontColor(Api.CreateColorFromRGB(100, 100, 200));
        sheet.GetRange("A3:H3").SetBold(true);
    })();
  checks:
      - type: "cell_value"
        cell: "A1"
        expected: "R7-MACRO-HACK"
---
# Урок 3. Практика: Операция «Шифр» (Чтение и Запись)

Данные в реальном мире редко лежат красиво в одной таблице. Иногда нужные цифры или тексты разбросаны по разным углам листа.

## 🕵️ Ваша миссия
В системе скрыты три части секретного пароля. Координаты частей:
* Первая часть: ячейка **D5**
* Вторая часть: **F12**
* Третья часть: **Z20**

Напишите макрос, который прочитает все три значения из этих ячеек и склеит их в единую строку-пароль через дефис (строго в порядке D5 -> F12 -> Z20). Запишите в **A1**.

### Готовое решение

```javascript
(function() {
    'use strict';
    try {
        let sheet = Api.GetActiveSheet();
        
        let p1 = sheet.GetRange("D5").GetValue();
        let p2 = sheet.GetRange("F12").GetValue();
        let p3 = sheet.GetRange("Z20").GetValue();
        
        // Конкатенация строк через тире
        let finalPassword = p1 + "-" + p2 + "-" + p3;
        
        sheet.GetRange("A1").SetValue(finalPassword);
    } catch (error) {
        console.error(error.message);
    }
})();
```
