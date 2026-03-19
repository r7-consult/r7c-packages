## Описание
- [ ] Обновления/Добавление плагинов
- [ ] Документация / инфраструктура
- [ ] Другое (указать ниже)

> Кратко опишите, что меняется и зачем. При работе с плагинами обязательно пройдите чеклист ниже.

---

## Чеклист плагина (обязательно)
- [ ] Папка плагина создана: `sdkjs-plugins/content/<plugin-name>/`
- [ ] `config.json` валиден (`name`, `guid`, `version`, `variations[0].url`)
- [ ] Точка входа из `config.json` существует (обычно `index.html`)
- [ ] Все пути из `config.json` существуют (icons/screenshots/translations/scripts)
- [ ] `store.categories` заполнены допустимыми значениями (`recommended`, `commercial`, `devTools`, `work`, `entertainment`, `communication`, `specAbilities`)
- [ ] `variations[0].description` описывает функциональность плагина
- [ ] `store.icons` корректно указывает иконки для `light` и `dark`
- [ ] `store.screenshots` содержит минимум один рабочий скриншот
- [ ] `store.background` заполнен для `light` и `dark`
- [ ] Плагин добавлен в `store/config.json` как `{ "name": "<plugin-name>", "discussion": "" }`
- [ ] В репозитории есть папка `deploy/`
- [ ] В `deploy/` лежит файл `<plugin-name>.plugin`

## Чеклист по умолчанию (рекомендуется)
- [ ] `README.md` в корне плагина
- [ ] `CHANGELOG.md` в корне плагина
- [ ] `LICENSE`/`LICENSE.md` и `3rd-Party.txt`
- [ ] Заполнено поле `offered` в `config.json`
- [ ] `store.categories` и `store.background` пересмотрены на соответствие UI

Если какой-то пункт не выполняется, добавьте объяснение в описание PR.
