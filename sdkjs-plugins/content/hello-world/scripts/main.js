(function () {
    "use strict";

    var selectedTextEl = null;

    function setSelectedText(value) {
        if (!selectedTextEl) {
            return;
        }

        if (!value || !String(value).trim()) {
            selectedTextEl.textContent = "No text selected";
            return;
        }

        selectedTextEl.textContent = String(value);
    }

    window.Asc.plugin.init = function (text) {
        selectedTextEl = document.getElementById("selectedText");
        setSelectedText(text);
    };

    window.Asc.plugin.button = function () {
        this.executeCommand("close", "");
    };

    window.Asc.plugin.onThemeChanged = function (theme) {
        window.Asc.plugin.onThemeChangedBase(theme);

        var rules = "";
        if (theme["background-toolbar"]) {
            rules += "body { background-color: " + theme["background-toolbar"] + "; }\n";
        }
        if (theme["text-normal"]) {
            rules += "body, .subtitle, .content-block h2, #selectedText { color: " + theme["text-normal"] + "; }\n";
        }
        if (theme["background-normal"]) {
            rules += ".hello-card, .content-block { background-color: " + theme["background-normal"] + "; }\n";
        }
        if (theme["border-regular-control"]) {
            rules += ".hello-card, .content-block { border-color: " + theme["border-regular-control"] + "; }\n";
        }

        var styleEl = document.getElementById("pluginStyles");
        if (styleEl) {
            styleEl.innerHTML = rules;
        }
    };
})();
