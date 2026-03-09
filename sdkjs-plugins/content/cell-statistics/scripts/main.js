(function () {
    "use strict";

    var displayNoneClass = "hidden";

    function formatNumber(num) {
        if (num === null || num === undefined || isNaN(num)) return "—";
        if (Number.isInteger(num)) return num.toLocaleString();
        return num.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 4 });
    }

    function computeStats(values) {
        var nums = values.filter(function (v) { return typeof v === "number" && !isNaN(v); });
        if (nums.length === 0) return null;

        var sum = nums.reduce(function (a, b) { return a + b; }, 0);
        var avg = sum / nums.length;
        var min = Math.min.apply(null, nums);
        var max = Math.max.apply(null, nums);

        var sorted = nums.slice().sort(function (a, b) { return a - b; });
        var median;
        var mid = Math.floor(sorted.length / 2);
        if (sorted.length % 2 === 0) {
            median = (sorted[mid - 1] + sorted[mid]) / 2;
        } else {
            median = sorted[mid];
        }

        var variance = nums.reduce(function (acc, val) {
            return acc + Math.pow(val - avg, 2);
        }, 0) / nums.length;
        var stdDev = Math.sqrt(variance);

        return {
            sum: sum,
            avg: avg,
            min: min,
            max: max,
            median: median,
            stdDev: stdDev,
            count: values.length,
            numericCount: nums.length,
            emptyCount: values.filter(function (v) { return v === null || v === undefined || v === ""; }).length
        };
    }

    function updateUI(stats) {
        var emptyState = document.getElementById("emptyState");
        var statsContainer = document.getElementById("statsContainer");

        if (!stats) {
            emptyState.classList.remove(displayNoneClass);
            statsContainer.classList.add(displayNoneClass);
            return;
        }

        emptyState.classList.add(displayNoneClass);
        statsContainer.classList.remove(displayNoneClass);

        document.getElementById("statSum").textContent = formatNumber(stats.sum);
        document.getElementById("statAvg").textContent = formatNumber(stats.avg);
        document.getElementById("statMin").textContent = formatNumber(stats.min);
        document.getElementById("statMax").textContent = formatNumber(stats.max);
        document.getElementById("statMedian").textContent = formatNumber(stats.median);
        document.getElementById("statStd").textContent = formatNumber(stats.stdDev);
        document.getElementById("statCount").textContent = "Count: " + stats.count;
        document.getElementById("statEmpty").textContent = "Empty: " + stats.emptyCount;
        document.getElementById("statNumeric").textContent = "Numbers: " + stats.numericCount;
    }

    window.Asc.plugin.init = function (data) {
        // Called on selection change due to initOnSelectionChanged: true
        requestSelectionData();
    };

    window.Asc.plugin.onMethodReturn = function (returnValue) {
        // Handle method return from GetSelectionData or similar
        if (returnValue && returnValue._methodName === "GetSelectedRange") {
            processRange(returnValue);
        }
    };

    function requestSelectionData() {
        // Use the macro-based approach to read selected cell values
        window.Asc.plugin.callCommand(function () {
            var oSheet = Api.GetActiveSheet();
            var oRange = oSheet.GetSelection();
            var result = [];
            var rows = [];

            // Get the address and parse it
            var addr = oRange.GetAddress(true, true, "xlA1", false);
            if (!addr) return JSON.stringify({ values: [], address: "" });

            // Use ForEach to iterate
            oRange.ForEach(function (cell) {
                var val = cell.GetValue();
                result.push(val);
            });

            return JSON.stringify({ values: result, address: addr });
        }, false, false, function (result) {
            if (!result) {
                updateUI(null);
                return;
            }
            try {
                var data = JSON.parse(result);
                var values = data.values || [];

                // Convert string numbers to actual numbers
                var processed = values.map(function (v) {
                    if (v === "" || v === null || v === undefined) return "";
                    var num = Number(v);
                    if (!isNaN(num) && v !== "") return num;
                    return v;
                });

                var stats = computeStats(processed);
                updateUI(stats);
            } catch (e) {
                updateUI(null);
            }
        });
    }

    window.Asc.plugin.button = function (id) {
        this.executeCommand("close", "");
    };

    Asc.plugin.onThemeChanged = function (theme) {
        window.Asc.plugin.onThemeChangedBase(theme);
        var rules = "";
        if (theme["background-toolbar"]) {
            rules += "body { background-color: " + theme["background-toolbar"] + "; }\n";
        }
        if (theme["text-normal"]) {
            rules += "body, .stat-value, .stat-label, .stat-info { color: " + theme["text-normal"] + "; }\n";
        }
        if (theme["background-normal"]) {
            rules += ".stat-card { background-color: " + theme["background-normal"] + "; }\n";
        }
        if (theme["border-regular-control"]) {
            rules += ".stat-card { border-color: " + theme["border-regular-control"] + "; }\n";
        }
        if (theme["text-tertiary"]) {
            rules += ".stat-label, .stat-info, .empty-state { color: " + theme["text-tertiary"] + "; }\n";
        }
        if (theme["background-primary-dialog-button"]) {
            rules += ".stat-card.primary { border-left-color: " + theme["background-primary-dialog-button"] + "; }\n";
        }
        var styleEl = document.getElementById("pluginStyles");
        if (styleEl) styleEl.innerHTML = rules;
    };
})();
