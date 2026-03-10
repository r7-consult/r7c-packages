(function(window, undefined) {
	if (!window.Asc)
		window.Asc = {};
	if (!window.Asc.plugin)
		window.Asc.plugin = {};

	window.Asc.plugin.init = function() {
		// Visual plugin UI is loaded from variations[0].url (index.html).
		// Keep bootstrap minimal to satisfy plugin runtime lifecycle.
	};

	window.Asc.plugin.button = function(id, windowID) {
		if (windowID)
			window.Asc.plugin.executeMethod('CloseWindow', [windowID]);
		else
			this.executeCommand('close', '');
	};
})(window, undefined);
