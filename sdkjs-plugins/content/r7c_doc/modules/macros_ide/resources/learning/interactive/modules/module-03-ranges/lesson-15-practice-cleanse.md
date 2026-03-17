---
interactive:
  version: 1
  enabled: true
  lessonId: "macro-ranges-15-practice-cleanse"
  title: "Урок 15: Практика (Cleanse Rows). Reverse Loop Deletion"
  draft: false
  targetEditor: "cell"
  starterCode: |
    (function() {
        'use strict';
        try {
            let sheet = Api.GetActiveSheet();
            
            // ЗАДАЧА: удалить все строки со статусом FAILED500.
            // ВАЖНО: цикл идёт СНИЗУ ВВЕРХ (i--), иначе при удалении строк сдвинутся индексы!
            
            for(let i = 50; i >= 2; i--) {
                let status = sheet.GetRange("D" + i).GetValue();
                if (status === "FAILED500") {
                    sheet.GetRange("A" + i + ":D" + i).Delete("up");
                }
            }
            
            // Записываем маркер успешного завершения
            sheet.GetRange("F4").SetValue("CLEAN");
            
        } catch (error) {
            console.error(error.message);
        }
    })();
  beforeScript: |
    (function() {
        let sheet = Api.GetActiveSheet();
        sheet.GetRange("A1:E100").Clear();
        let header = [["ID", "Transaction", "Amount", "Status"]];
        sheet.GetRange("A1:D1").SetValue(header);
        sheet.GetRange("A1:D1").SetBold(true);
        let row = 2;
        let ids = [1002, 1003, 1005, 1006, 1008, 1010, 1011, 1013, 1015, 1017, 1018, 1019, 1021, 1023, 1025];
        let good = ["2024-01-12", "2024-01-13", "2024-01-15", "2024-01-16", "2024-01-18"];
        let gi = 0;
        for(let i=1; i<=25; i++) {
            sheet.GetRange("A"+row).SetValue(1000+i);
            sheet.GetRange("B"+row).SetValue("TRX-"+String(1000+i).slice(-3));
            sheet.GetRange("C"+row).SetValue(Math.round(Math.random()*50000+1000));
            if (i % 4 === 0) {
                sheet.GetRange("D"+row).SetValue("FAILED500");
                sheet.GetRange("A"+row+":D"+row).SetFillColor(Api.CreateColorFromRGB(255, 200, 200));
            } else {
                sheet.GetRange("D"+row).SetValue("SUCCESS");
            }
            row++;
        }
        // Pin some FAILED rows at specific spots for deterministic testing
        sheet.GetRange("D4").SetValue("FAILED500");
        sheet.GetRange("A4:D4").SetFillColor(Api.CreateColorFromRGB(255, 200, 200));
        sheet.GetRange("D8").SetValue("FAILED500");
        sheet.GetRange("A8:D8").SetFillColor(Api.CreateColorFromRGB(255, 200, 200));
        sheet.GetRange("F1:J2").Merge(false);
        sheet.GetRange("F1:J2").SetValue("АВАРИЯ! Лог транзакций поврёжден. Удалите строки со статусом 'FAILED500' (снизу вверх!)");
        sheet.GetRange("F1:J2").SetFontColor(Api.CreateColorFromRGB(200, 0, 0));
        sheet.GetRange("F1:J2").SetBold(true);
    })();
  checks:
      - type: "cell_value"
        cell: "F4"
        expected: "CLEAN"
---
# Урок 15. LeetCode: Удаление Мусорных Строк

Сырой дамп из базы данных банковских транзакций. Произошел сбой — часть строк имеют статус `FAILED500`.

## Задача
Просканируйте строки с 50 по 2 **в обратном порядке** (`i--`) и удалите строки где в столбце `D` написано `FAILED500` методом `.Delete("up")`.

> **Почему снизу вверх?** При удалении строки, все нижние строки сдвигаются на позицию вверх. Если идти сверху вниз — вы пропустите строку после удалённой. Движение `i--` это решает.

### Готовое решение

```javascript
(function() {
    'use strict';
    try {
        let sheet = Api.GetActiveSheet();
        for(let i = 50; i >= 2; i--) {
            if (sheet.GetRange("D" + i).GetValue() === "FAILED500") {
                sheet.GetRange("A" + i + ":D" + i).Delete("up");
            }
        }
        sheet.GetRange("F4").SetValue("CLEAN");
    } catch (error) {
        console.error(error.message);
    }
})();
```
