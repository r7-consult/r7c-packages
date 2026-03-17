# 🧪 Инструкция по тестированию Api.SqlNotebook.query() в R7 Office

## 📋 Шаг 1: Загрузите тестовые данные

1. Откройте **R7 Office**
2. Откройте **Macros IDE** (плагин)
3. Перейдите на вкладку **"SQLite Notebook"**
4. Нажмите кнопку **"Upload File"** или перетащите файл
5. Выберите файл **MOCK_DATA.csv** из папки:
   ```
   modules/macros_ide/resources/examples/MOCK_DATA.csv
   ```
6. Дождитесь загрузки - файл появится в списке источников как **"MOCK_DATA.csv"**

## 📋 Шаг 2: Проверьте данные в SQLite Notebook UI

Выполните тестовый запрос прямо в SQLite Notebook:

```sql
SELECT * FROM data LIMIT 5;
```

Если данные отобразились - значит файл загружен правильно.

## 📋 Шаг 3: Запустите минимальный тестовый макрос

1. Перейдите на вкладку **"Macros"** в Macros IDE
2. Создайте новый макрос (кнопка "New Macro")
3. Скопируйте код из файла **Test_MINIMAL.md**:

```javascript
(function() {
    var sheet = Api.GetActiveSheet();
    var result = Api.SqlNotebook.query("SELECT * FROM data", "MOCK_DATA.csv");

    if (!result.ok) {
        sheet.GetRange("A1").SetValue("ERROR: " + result.error.message);
        return;
    }

    var df = result.result;

    // Заголовки
    for (var i = 0; i < df.columns.length; i++) {
        sheet.GetRange(String.fromCharCode(65 + i) + "1").SetValue(df.columns[i]);
    }

    // Данные
    for (var r = 0; r < df.rows.length; r++) {
        for (var c = 0; c < df.rows[r].length; c++) {
            sheet.GetRange(String.fromCharCode(65 + c) + (r + 2)).SetValue(df.rows[r][c]);
        }
    }

    // Итого
    sheet.GetRange("A" + (df.rows.length + 3)).SetValue("Всего строк: " + df.rowCount);
})();
```

4. Нажмите **"Run"** (или F5)
5. Проверьте результат на листе Excel

## ✅ Ожидаемый результат:

В таблице Excel должно появиться:

```
A1: id         B1: name              C1: email                    D1: gender  ...
A2: 1          B2: Alice Johnson     C2: alice.j@example.com      D2: Female  ...
A3: 2          B3: Bob Smith         C3: bob.smith@example.com    D3: Male    ...
...
A51: (пусто)   B51: (пусто)          C51: (пусто)                 D51: (пусто)
A52: Всего строк: 50
```

## 🔍 Проверка работы:

### Тест 1: Все данные
```javascript
Api.SqlNotebook.query("SELECT * FROM data", "MOCK_DATA.csv");
```
**Ожидается:** 50 строк с 8 колонками

### Тест 2: Фильтрация
```javascript
Api.SqlNotebook.query("SELECT * FROM data WHERE value > 200", "MOCK_DATA.csv");
```
**Ожидается:** ~13 строк

### Тест 3: Группировка
```javascript
Api.SqlNotebook.query(
    "SELECT country, COUNT(*) as count FROM data GROUP BY country",
    "MOCK_DATA.csv"
);
```
**Ожидается:** 6 строк (USA, UK, Canada, Australia, Germany, France)

### Тест 4: Сортировка
```javascript
Api.SqlNotebook.query(
    "SELECT name, value FROM data ORDER BY value DESC LIMIT 10",
    "MOCK_DATA.csv"
);
```
**Ожидается:** 10 строк, отсортированных по убыванию value

## ❌ Возможные ошибки:

### Ошибка: "FILE_NOT_FOUND"
**Причина:** Файл не загружен в SQLite Notebook
**Решение:**
1. Откройте вкладку SQLite Notebook
2. Загрузите MOCK_DATA.csv
3. Проверьте что файл появился в списке источников

### Ошибка: "BACKEND_NOT_AVAILABLE"
**Причина:** SQLite Notebook backend не инициализирован
**Решение:**
1. Откройте вкладку SQLite Notebook хотя бы один раз
2. Дождитесь полной загрузки
3. Попробуйте запустить макрос снова

### Ошибка: "SQL_EXECUTION_ERROR"
**Причина:** Синтаксическая ошибка в SQL
**Решение:**
1. Проверьте синтаксис SQL запроса
2. Убедитесь что используется `data` как имя таблицы (не `MOCK_DATA`)
3. Проверьте имена колонок

### Ошибка: "EXTENSION_NOT_LOADED"
**Причина:** SqlNotebook extension не загружен
**Решение:**
1. Проверьте что файл sqlite-notebook-api-extension.js подключен в index.html
2. Перезагрузите страницу (F5)
3. Откройте консоль браузера (F12) и проверьте наличие ошибок

## 🐛 Отладка:

Откройте консоль браузера (F12) и выполните:

```javascript
// Проверка доступности extension
console.log("Extension:", window.SqlNotebookApiExtension);

// Проверка backend
console.log("Backend:", window.SQLiteWASMBackend);

// Проверка Api
console.log("Api.SqlNotebook:", Api.SqlNotebook);

// Проверка загруженных файлов
console.log("Sources:", window.SQLiteWASMBackend.listSources());

// Тест запроса
var result = Api.SqlNotebook.query("SELECT COUNT(*) FROM data", "MOCK_DATA.csv");
console.log("Result:", result);
```

## 📝 Примеры для копирования:

### Пример 1: Вывод топ-10 по значению
```javascript
(function() {
    var sheet = Api.GetActiveSheet();
    var result = Api.SqlNotebook.query(
        "SELECT name, value, country FROM data ORDER BY value DESC LIMIT 10",
        "MOCK_DATA.csv"
    );

    if (!result.ok) {
        sheet.GetRange("A1").SetValue("ERROR: " + result.error.message);
        return;
    }

    var df = result.result;

    // Заголовки
    for (var i = 0; i < df.columns.length; i++) {
        sheet.GetRange(String.fromCharCode(65 + i) + "1").SetValue(df.columns[i]);
        sheet.GetRange(String.fromCharCode(65 + i) + "1").SetBold(true);
    }

    // Данные
    for (var r = 0; r < df.rows.length; r++) {
        for (var c = 0; c < df.rows[r].length; c++) {
            sheet.GetRange(String.fromCharCode(65 + c) + (r + 2)).SetValue(df.rows[r][c]);
        }
    }
})();
```

### Пример 2: Статистика по странам
```javascript
(function() {
    var sheet = Api.GetActiveSheet();
    var result = Api.SqlNotebook.query(
        "SELECT country, COUNT(*) as users, AVG(value) as avg FROM data GROUP BY country",
        "MOCK_DATA.csv"
    );

    if (!result.ok) {
        sheet.GetRange("A1").SetValue("ERROR: " + result.error.message);
        return;
    }

    var df = result.result;

    // Заголовки
    sheet.GetRange("A1").SetValue("Страна");
    sheet.GetRange("B1").SetValue("Пользователей");
    sheet.GetRange("C1").SetValue("Среднее значение");
    sheet.GetRange("A1:C1").SetBold(true);

    // Данные
    for (var r = 0; r < df.rows.length; r++) {
        sheet.GetRange("A" + (r + 2)).SetValue(df.rows[r][0]);
        sheet.GetRange("B" + (r + 2)).SetValue(df.rows[r][1]);
        sheet.GetRange("C" + (r + 2)).SetValue(Math.round(df.rows[r][2]));
    }
})();
```

### Пример 3: Активные пользователи
```javascript
(function() {
    var sheet = Api.GetActiveSheet();
    var result = Api.SqlNotebook.query(
        "SELECT name, email, value FROM data WHERE status = 'active' AND value > 150",
        "MOCK_DATA.csv"
    );

    if (!result.ok) {
        sheet.GetRange("A1").SetValue("ERROR: " + result.error.message);
        return;
    }

    var df = result.result;

    sheet.GetRange("A1").SetValue("Активные пользователи (value > 150)");
    sheet.GetRange("A1").SetFontSize(14);
    sheet.GetRange("A1").SetBold(true);

    // Заголовки
    for (var i = 0; i < df.columns.length; i++) {
        sheet.GetRange(String.fromCharCode(65 + i) + "3").SetValue(df.columns[i]);
        sheet.GetRange(String.fromCharCode(65 + i) + "3").SetBold(true);
    }

    // Данные
    for (var r = 0; r < df.rows.length; r++) {
        for (var c = 0; c < df.rows[r].length; c++) {
            sheet.GetRange(String.fromCharCode(65 + c) + (r + 4)).SetValue(df.rows[r][c]);
        }
    }

    // Итого
    sheet.GetRange("A" + (df.rows.length + 5)).SetValue("Найдено: " + df.rowCount + " пользователей");
})();
```

---

**Важно:** В R7 Office не используйте `Api.ShowMessage()` - выводите результаты прямо в ячейки!
