---
interactive:
  version: 1
  enabled: true
  lessonId: "macro-templates-30-addsheet"
  title: "Урок 30: Анализ Эталона «Фабрика листов» (AddSheet)"
  draft: false
  targetEditor: "cell"
  starterCode: |
    (function() {
        'use strict';
        
        try {
            const api = Api;
            if (!api) throw new Error('API not available');
            
            // ВАША ЗАДАЧА ЗДЕСЬ (Фабрика Листов из Эталона):
            // 1. Создание нового листа через Api.AddSheet("ИмяЛиста")
            // Назовите лист "New sheet"
            let sheet = Api.AddSheet("New sheet");
            
            // 2. Метод AddSheet возвращает ССЫЛКУ на созданный лист (ApiWorksheet)
            // Докажите это: вызовите метод SetValue("Hello") на ячейке A1 этого sheet.
            
            
            // 3. Вывод лога (из эталона)
            console.log('Macro executed successfully');
            
        } catch (error) {
            console.error('Macro execution failed:', error.message);
            if (typeof Api !== 'undefined' && Api.GetActiveSheet) {
                let currentSheet = Api.GetActiveSheet();
                if (currentSheet) currentSheet.GetRange('A1').SetValue('Error: ' + error.message);
            }
        }
    })();
  checks:
    - type: "range_values"
      sheetName: "New sheet"
      range: "A1"
      expected: "Hello"
  beforeScript: |
    (function() {
    try {
        let sheet = Api.GetActiveSheet();
        sheet.GetRange("G1:J1").Merge(false);
        sheet.GetRange("G1:J1").SetValue("💡 Интерактивная песочница готова. Напишите макрос!");
        sheet.GetRange("G1:J1").SetFontColor(Api.CreateColorFromRGB(120, 120, 120));
        sheet.GetRange("G1:J1").SetItalic(true);
    } catch(e) {}
    })();
---

# Урок 30. Разбор до Костей: Шаблон «Добавление листа» (AddSheet)

В Уроке 21 мы с вами уже создавали "Секретную Базу" с помощью метода `Api.AddSheet`. Вы спросите: "Зачем мы возвращаемся к этому в самом сложном модуле?".
Потому что в Уроке 21 мы работали как любители, писали код "в чистом поле" без перехвата ошибок. А теперь мы читаем скрипт-эталон от разработчиков ядра `AddSheet_macroR7.js`. 

## 🎯 Архитектура Возвращаемых Значений (Returns)

Каждый раз, когда вы открываете любой эталонный скрипт от Р7, обращайте внимание на шапку его комментариев `.md` или `.js`.
Там написан так называемый контракт метода (JSDoc):
`@returns {ApiWorksheet} The newly created ApiWorksheet object.`

Это означает, что `Api.AddSheet` (как и большинство Фабрик в Р7-Офис) не просто где-то "в фоне" делает свою работу. **Он Возвращает в память полноценный Объект Листа**.
Вы сразу же можете его присвоить в переменную: `let newSheet = Api.AddSheet("New sheet");` 

**В чем гениальность эталонного макроса?**
Разработчик может моментально взять `newSheet` и продолжить цепочку вызовов: скрыть его `SetHidden()`, положить туда формулы `GetRange...`, закрасить лист, настроить столбцы `SetColumnWidth`. Раньше в VBA (в MS Excel) разработчикам часто приходилось писать `.Sheets.Add()` и потом отдельно искать добавленный лист через `.ActiveSheet`. Движок V8 возвращает Pointer (Указатель) мгновенно.

---

## 📝 Постановка задачи (Ядро Эталона)

Разберите эталонный код в редакторе (включающий `'use strict'` и `catch(error)`): 

1. Создайте лист с названием `"New sheet"`, как это написано в эталоне.
   (Здесь мы не делаем проверку на "уже существующий" лист, так как интерактивный контейнер сервера собирает чистые файлы каждый раз).
2. Заберите (перехватите) возвращенный объект в переменную (например, `sheet`).
3. Докажите, что у вас в руках оказался живой объект (С экземпляром класса `ApiWorksheet`): обратитесь к `sheet.GetRange("A1")` и заполните его словом `"Hello"`.

**Обратите внимание на Блок Обработки Ошибок `catch(error)`!**
Если разработчик вызовет макрос второй раз, имя `"New sheet"` УЖЕ будет существовать в файле пользователя. Движок бросит системное исключение `Error`. В вашем коде этот удар примет щит `catch(error)`. Вместо системного обрыва, код грациозно пойдет в блок `catch`, считает имя ошибки `error.message` ("Такое имя листа уже занято") и **печатает эту текстовую ошибку на активном листе пользователя**. Это уровень Professional Engineering! 

Жмите "Проверить", чтобы сдать Урок 30 и перейти к Финальному Экзамену Модуля 5. Архитектором Р7-Офис вы станете уже в следующем уроке!
