# SQLite Notebook API - Тестовые данные и примеры

## 📊 MOCK_DATA.csv

Тестовый файл содержит 50 записей с следующими колонками:

| Колонка | Тип | Описание |
|---------|-----|----------|
| `id` | INTEGER | Уникальный идентификатор (1-50) |
| `name` | TEXT | Имя пользователя |
| `email` | TEXT | Email адрес |
| `gender` | TEXT | Пол (Male/Female) |
| `country` | TEXT | Страна (USA, UK, Canada, Australia, Germany, France) |
| `value` | INTEGER | Числовое значение (45-345) |
| `status` | TEXT | Статус (active/inactive) |
| `date` | TEXT | Дата (2024-01-15 до 2024-03-05) |

## 🚀 Быстрый старт

### 1. Загрузка данных в SQLite Notebook

1. Откройте **SQLite Notebook** вкладку в Macros IDE
2. Нажмите **"Upload File"** или перетащите файл
3. Выберите `MOCK_DATA.csv` из `resources/examples/`
4. Дождитесь загрузки (файл появится в списке источников)

### 2. Тестирование в UI

Попробуйте выполнить запросы прямо в SQLite Notebook:

```sql
-- Просмотр всех данных
SELECT * FROM data LIMIT 10;

-- Подсчет записей
SELECT COUNT(*) as total FROM data;

-- Фильтрация по значению
SELECT * FROM data WHERE value > 200;

-- Группировка по стране
SELECT country, COUNT(*) as count, AVG(value) as avg_value
FROM data
GROUP BY country
ORDER BY count DESC;

-- Активные пользователи
SELECT name, email, value
FROM data
WHERE status = 'active'
ORDER BY value DESC
LIMIT 5;
```

### 3. Использование в макросах

## Пример 1: Простой запрос

```javascript
(function() {
    'use strict';

    var result = Api.SqlNotebook.query(
        "SELECT * FROM data WHERE value > 200",
        "MOCK_DATA.csv"
    );

    if (!result.ok) {
        Api.ShowMessage("Ошибка", result.error.message);
        return;
    }

    var sheet = Api.GetActiveSheet();
    var df = result.result;

    // Вывод заголовков
    df.columns.forEach(function(col, i) {
        sheet.GetRange(String.fromCharCode(65 + i) + '1').SetValue(col);
    });

    // Вывод данных
    df.rows.forEach(function(row, rowIdx) {
        row.forEach(function(val, colIdx) {
            var cell = String.fromCharCode(65 + colIdx) + (rowIdx + 2);
            sheet.GetRange(cell).SetValue(val);
        });
    });

    Api.ShowMessage("Успех", "Загружено " + df.rowCount + " строк");
})();
```

## Пример 2: Агрегация по странам

```javascript
(function() {
    'use strict';

    var result = Api.SqlNotebook.query(
        "SELECT country, COUNT(*) as count, AVG(value) as avg_value " +
        "FROM data GROUP BY country ORDER BY count DESC",
        "MOCK_DATA.csv"
    );

    if (!result.ok) {
        Api.ShowMessage("Ошибка", result.error.message);
        return;
    }

    var sheet = Api.GetActiveSheet();
    var df = result.result;

    // Заголовки с форматированием
    sheet.GetRange("A1").SetValue("Страна");
    sheet.GetRange("B1").SetValue("Количество");
    sheet.GetRange("C1").SetValue("Среднее значение");
    sheet.GetRange("A1:C1").SetBold(true);

    // Данные
    df.rows.forEach(function(row, i) {
        sheet.GetRange("A" + (i + 2)).SetValue(row[0]);
        sheet.GetRange("B" + (i + 2)).SetValue(row[1]);
        sheet.GetRange("C" + (i + 2)).SetValue(Math.round(row[2]));
    });

    Api.ShowMessage("Статистика", "Обработано стран: " + df.rowCount);
})();
```

## Пример 3: Фильтрация активных пользователей

```javascript
(function() {
    'use strict';

    var result = Api.SqlNotebook.query(
        "SELECT name, email, value, country " +
        "FROM data WHERE status = 'active' AND value > 150 " +
        "ORDER BY value DESC",
        "MOCK_DATA.csv"
    );

    if (!result.ok) {
        Api.ShowMessage("Ошибка", result.error.message);
        return;
    }

    var sheet = Api.GetActiveSheet();
    var df = result.result;

    // Очистка области
    sheet.GetRange("A1:Z100").Clear();

    // Заголовок отчета
    sheet.GetRange("A1").SetValue("Топ активных пользователей (значение > 150)");
    sheet.GetRange("A1").SetBold(true);
    sheet.GetRange("A1").SetFontSize(14);

    // Заголовки колонок
    var headers = ["Имя", "Email", "Значение", "Страна"];
    headers.forEach(function(header, i) {
        var cell = String.fromCharCode(65 + i) + "3";
        sheet.GetRange(cell).SetValue(header);
        sheet.GetRange(cell).SetBold(true);
    });

    // Данные
    df.rows.forEach(function(row, rowIdx) {
        row.forEach(function(val, colIdx) {
            var cell = String.fromCharCode(65 + colIdx) + (rowIdx + 4);
            sheet.GetRange(cell).SetValue(val);
        });
    });

    // Итоговая строка
    var totalRow = df.rowCount + 5;
    sheet.GetRange("A" + totalRow).SetValue("Итого найдено:");
    sheet.GetRange("B" + totalRow).SetValue(df.rowCount + " пользователей");
    sheet.GetRange("A" + totalRow + ":B" + totalRow).SetBold(true);

    Api.ShowMessage("Отчет готов", "Найдено " + df.rowCount + " активных пользователей");
})();
```

## Пример 4: Обработка ошибок

```javascript
(function() {
    'use strict';

    // Функция для безопасного выполнения запроса
    function safeQuery(sql, filename) {
        var result = Api.SqlNotebook.query(sql, filename);

        if (!result.ok) {
            var errorMsg = "Ошибка выполнения запроса:\n\n";
            errorMsg += "Код: " + result.error.code + "\n";
            errorMsg += "Сообщение: " + result.error.message;

            console.error("[SQL Error]", result.error);
            Api.ShowMessage("Ошибка SQL", errorMsg);
            return null;
        }

        return result.result;
    }

    // Использование
    var df = safeQuery(
        "SELECT * FROM data WHERE value BETWEEN 100 AND 200",
        "MOCK_DATA.csv"
    );

    if (!df) {
        return; // Ошибка уже обработана
    }

    // Работа с результатом
    var sheet = Api.GetActiveSheet();
    console.log("Получено строк:", df.rowCount);
    console.log("Колонок:", df.columnCount);

    // Вывод в таблицу
    df.columns.forEach(function(col, i) {
        sheet.GetRange(String.fromCharCode(65 + i) + '1').SetValue(col);
    });

    df.rows.forEach(function(row, rowIdx) {
        row.forEach(function(val, colIdx) {
            var cell = String.fromCharCode(65 + colIdx) + (rowIdx + 2);
            sheet.GetRange(cell).SetValue(val);
        });
    });

    Api.ShowMessage("Успех", "Данные загружены успешно");
})();
```

## 📚 Полезные SQL запросы для тестирования

```sql
-- 1. Подсчет по полу
SELECT gender, COUNT(*) as count
FROM data
GROUP BY gender;

-- 2. Топ-5 пользователей по значению
SELECT name, value, country
FROM data
ORDER BY value DESC
LIMIT 5;

-- 3. Средние значения по странам
SELECT country,
       COUNT(*) as users,
       AVG(value) as avg_value,
       MIN(value) as min_value,
       MAX(value) as max_value
FROM data
GROUP BY country
ORDER BY avg_value DESC;

-- 4. Активные vs Неактивные
SELECT status, COUNT(*) as count, AVG(value) as avg_value
FROM data
GROUP BY status;

-- 5. Пользователи с высокими значениями по странам
SELECT country, name, value
FROM data
WHERE value > (SELECT AVG(value) FROM data)
ORDER BY country, value DESC;

-- 6. Данные за конкретный период
SELECT * FROM data
WHERE date BETWEEN '2024-02-01' AND '2024-02-28'
ORDER BY date;

-- 7. Email домены
SELECT SUBSTR(email, INSTR(email, '@') + 1) as domain,
       COUNT(*) as count
FROM data
GROUP BY domain
ORDER BY count DESC;
```

## 🔍 Коды ошибок

| Код | Описание | Решение |
|-----|----------|---------|
| `INVALID_SQL` | SQL запрос пустой или невалидный | Проверьте синтаксис SQL |
| `INVALID_FILENAME` | Имя файла пустое или невалидное | Укажите корректное имя файла |
| `BACKEND_NOT_AVAILABLE` | SQLite Notebook backend не загружен | Откройте SQLite Notebook вкладку |
| `FILE_NOT_FOUND` | Файл не найден среди загруженных | Загрузите файл через SQLite Notebook UI |
| `SQL_EXECUTION_ERROR` | Ошибка выполнения SQL запроса | Проверьте SQL синтаксис и имена колонок |
| `NO_FILES_LOADED` | Нет загруженных файлов | Загрузите хотя бы один файл |

## ⚙️ Опции query()

```javascript
Api.SqlNotebook.query(sql, filename, {
    autoAttach: true  // Автоматически активировать источник (по умолчанию: true)
});
```

## 💡 Советы и рекомендации

1. **Всегда проверяйте result.ok** перед использованием данных
2. **Используйте LIMIT** для ограничения результатов больших запросов
3. **Используйте console.log()** для отладки SQL запросов
4. **Проверяйте имя файла** - оно должно совпадать с именем в SQLite Notebook
5. **Обрабатывайте ошибки** с понятными сообщениями для пользователя

## 🎯 Следующие шаги

1. Попробуйте базовые примеры из этого руководства
2. Изучите примеры в `SqlNotebook_Query_Test.md`
3. Создайте свои SQL запросы на основе MOCK_DATA.csv
4. Интегрируйте запросы в существующие макросы

---

**Документация обновлена:** 02.02.2026
**Версия API:** 2.3.0
