# Hello World Plugin Template

Базовый шаблон плагина для `{r7c}.packages`.

## Что внутри
- Минимальный `config.json` с обязательными полями.
- `index.html` + `scripts/main.js` + `styles/main.css`.
- Поддержка light/dark через `onThemeChanged`.
- Структура `resources/`, `deploy/`, `README.md`, `CHANGELOG.md`, `LICENSE.md`, `3rd-Party.txt`.

## Как использовать как шаблон
1. Скопируйте папку `sdkjs-plugins/content/hello-world`.
2. Переименуйте папку в имя нового плагина.
3. Обновите `name`, `guid`, `description`, пути и категорию в `config.json`.
4. Добавьте запись в `store/config.json`:
   - `{ "name": "<plugin-name>", "discussion": "" }`
5. Подмените иконки и скриншоты в `resources/store/` и `resources/light|dark/`.
6. Соберите итоговый `.plugin` в `deploy/` и загрузите его вручную в {r7}/ONLYOFFICE через Plugins → Upload, чтобы проверить работоспособность.
