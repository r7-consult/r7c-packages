/**
 *
 * (c) Copyright Ascensio System SIA 2020
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *	 http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */
(function (window, undefined) {
	const isLocal = ((window.AscDesktopEditor !== undefined) && (window.location.protocol.indexOf('file') !== -1));
	let interval = null;
	let errTimeout = null;
	let loader = null;
	let loaderTimeout = null;

	// create iframe
	const iframe = document.createElement('iframe');

	let warningWindow = null;
	let developerWindow = null;
	let removeGuid = null;
	let BFrameReady = false;
	let BPluginReady = false;
	let editorVersion = null;
	let marketplaceURl = null;
	let marketplaceRepo = '';
	const OOMarketplaceUrl = 'https://maikai-dev.github.io/r7-plugin-packages/store/index.html';

	function normalizeMarketplaceUrl(url) {
		let value = (url || '').trim();
		if (!value.length)
			return '';
		if (value === './store/index.html')
			return value;
		// External custom URL must point to the marketplace entrypoint.
		if (!/\/store\/index\.html([?#].*)?$/i.test(value))
			return '';
		return value;
	}

	function normalizeMarketplaceRepo(value) {
		let source = (value || '').trim();
		if (!source.length)
			return '';

		let match = source.match(/^([A-Za-z0-9_.-]+)\/([A-Za-z0-9_.-]+)(?:#([A-Za-z0-9_.\-/]+))?$/);
		if (!match) {
			match = source.match(/^https?:\/\/github\.com\/([^\/\s]+)\/([^\/\s#?]+)(?:\/tree\/([^#?]+))?\/?$/i);
		}
		if (!match)
			return '';

		let owner = match[1];
		let repo = match[2].replace(/\.git$/i, '');
		let branch = (match[3] || 'main').replace(/\/+$/, '');
		return owner + '/' + repo + '#' + branch;
	}

	function hasCustomSource() {
		return !!marketplaceRepo || marketplaceURl !== OOMarketplaceUrl;
	}
	try {
		// for incognito mode
		let savedRepo = normalizeMarketplaceRepo(localStorage.getItem('DeveloperMarketplaceRepo'));
		if (savedRepo.length) {
			marketplaceRepo = savedRepo;
			localStorage.setItem('DeveloperMarketplaceRepo', marketplaceRepo);
			localStorage.removeItem('DeveloperMarketplaceUrl');
		}
		let savedUrl = localStorage.getItem('DeveloperMarketplaceUrl');
		let normalizedUrl = normalizeMarketplaceUrl(savedUrl);
		marketplaceURl = marketplaceRepo ? OOMarketplaceUrl : (normalizedUrl || OOMarketplaceUrl);
		if (savedUrl && !normalizedUrl)
			localStorage.removeItem('DeveloperMarketplaceUrl');
	} catch (err) {
		marketplaceURl = OOMarketplaceUrl;
	}

	function applyThemeClass(type) {
		let mode = (type || '').indexOf('light') !== -1 ? 'light' : 'dark';
		document.body.classList.remove('theme-type-light', 'theme-type-dark');
		document.body.classList.add('theme-type-' + mode);
	}

	window.Asc.plugin.init = function () {
		applyThemeClass(window.Asc.plugin.theme && window.Asc.plugin.theme.type);
		window.Asc.plugin.executeMethod('ShowButton', ['developer', true, 'right']);
		// resize window
		window.Asc.plugin.resizeWindow(608, 600, 608, 600, 0, 0);
		if (!isLocal) {
			checkInternet(true);
			loaderTimeout = setTimeout(createLoader, 500);
		} else {
			initPlugin();
		}
	};

	function postMessage(message) {
		iframe.contentWindow.postMessage(JSON.stringify(message), '*');
	};

	function normalizeHostActionResult(action, guid, backup, result) {
		if (result && typeof result === 'object' && typeof result.type === 'string') {
			return result;
		}

		let isSuccess = (result === true) || (result && typeof result === 'object' && result.status === true);
		if (isSuccess) {
			switch (action) {
				case 'install':
					return { type: 'Installed', guid: guid };
				case 'update':
					return { type: 'Updated', guid: guid };
				case 'remove':
					return { type: 'Removed', guid: guid, backup: !!backup };
				default:
					return { type: 'Error', guid: guid, error: { message: 'Unknown successful operation result.' } };
			}
		}

		let errorMessage = '';
		if (typeof result === 'string' && result.length) {
			errorMessage = result;
		} else if (result && typeof result === 'object') {
			if (typeof result.error === 'string' && result.error.length)
				errorMessage = result.error;
			else if (typeof result.message === 'string' && result.message.length)
				errorMessage = result.message;
		} else if (result === false) {
			errorMessage = 'Operation returned false.';
		}

		if (!errorMessage.length)
			errorMessage = 'Problem with plugin ' + action + '.';

		return { type: 'Error', guid: guid || '', error: { message: errorMessage } };
	}

	function initPlugin() {
		document.body.appendChild(iframe);
		if (hasCustomSource())
			document.getElementById('notification').classList.remove('hidden');

		// send message that plugin is ready
		window.Asc.plugin.executeMethod('GetVersion', null, function (version) {
			editorVersion = version;
			BPluginReady = true;
			if (BFrameReady)
				postMessage({ type: 'PluginReady', version: editorVersion });
		});

		let divNoInt = document.getElementById('div_noIternet');
		let style = document.getElementsByTagName('head')[0].lastChild;
		let pageUrl = marketplaceURl;
		iframe.src = pageUrl + window.location.search;
		iframe.onload = function () {
			BFrameReady = true;
			if (BPluginReady) {
				if (!divNoInt.classList.contains('hidden')) {
					divNoInt.classList.add('hidden');
					clearInterval(interval);
					interval = null;
				}
				postMessage({ type: 'Theme', theme: window.Asc.plugin.theme, style: style.innerHTML });
				postMessage({ type: 'PluginReady', version: editorVersion });
			}
		};
	};

	window.Asc.plugin.button = function (id, windowID) {
		if (warningWindow && warningWindow.id == windowID) {
			switch (id) {
				case 0:
					removePlugin(false);
					break;
				case 1:
					removePlugin(true);
					break;
				default:
					postMessage({ type: 'Removed', guid: '' });
					break;
			}
			window.Asc.plugin.executeMethod('CloseWindow', [windowID]);
		} else if (developerWindow && developerWindow.id == windowID) {
			if (id == 0)
				developerWindow.command('onClickBtn');
			else
				window.Asc.plugin.executeMethod('CloseWindow', [windowID]);
		} else if (id == 'back') {
			window.Asc.plugin.executeMethod('ShowButton', ['back', false]);
			if (iframe && iframe.contentWindow)
				postMessage({ type: 'onClickBack' });
		} else if (id == 'developer') {
			createWindow('developer');
		} else {
			this.executeCommand('close', '');
		}
	};

	window.addEventListener('message', function (message) {
		// getting messages from marketplace
		let data = JSON.parse(message.data);
		if (!data || typeof data !== 'object' || !data.type) {
			postMessage({ type: 'Error', error: { message: 'Unknown command payload from marketplace frame.' } });
			return;
		}

		switch (data.type) {
			case 'getInstalled':
				// данное сообщение используется только при инициализации плагина и по умолчанию идёт парсинг и отрисовка плагинов из стора
				// добавлен флаг updateInstalled - в этом случае не загружаем плагины из стора повторно, работаем только с установленными

				window.Asc.plugin.executeMethod('GetInstalledPlugins', null, function (result) {
					postMessage({ type: 'InstalledPlugins', data: result, updateInstalled: data.updateInstalled });
				});
				break;
			case 'install':
				window.Asc.plugin.executeMethod('InstallPlugin', [data.config, data.guid], function (result) {
					postMessage(normalizeHostActionResult('install', data.guid, false, result));
				});
				break;
			case 'remove':
				removeGuid = data.guid;
				if (Number(editorVersion.split('.').join('') < 740))
					removePlugin(true);
				else if (!data.backup)
					removePlugin(data.backup);
				else
					createWindow('warning');
				break;
			case 'update':
				window.Asc.plugin.executeMethod('UpdatePlugin', [data.config, data.guid], function (result) {
					postMessage(normalizeHostActionResult('update', data.guid, false, result));
				});
				break;
			case 'showButton':
				window.Asc.plugin.executeMethod('ShowButton', ['back', true]);
				break;
		}

	}, false);

	window.Asc.plugin.onExternalMouseUp = function () {
		// mouse up event outside the plugin window
		if (iframe && iframe.contentWindow) {
			postMessage({ type: 'onExternalMouseUp' });
		}
	};

	window.Asc.plugin.onThemeChanged = function (theme) {
		// theme changed event
		if (theme.type.indexOf('light') !== -1) {
			theme['background-toolbar'] = '#f1f1f1';
		}
		applyThemeClass(theme.type);
		window.Asc.plugin.onThemeChangedBase(theme);
		let style = document.getElementsByTagName('head')[0].lastChild;
		if (iframe && iframe.contentWindow)
			postMessage({ type: 'Theme', theme: theme, style: style.innerHTML });
	};

	window.Asc.plugin.onTranslate = function () {
		let label = document.getElementById('lb_notification');
		if (label)
			label.innerHTML = window.Asc.plugin.tr(label.innerHTML);

		label = document.getElementById('lb_noInternet');
		if (label)
			label.innerHTML = window.Asc.plugin.tr(label.innerHTML);
	};

	function checkInternet(bSetTimeout) {
		try {
			let xhr = new XMLHttpRequest();
			let url = 'https://maikai-dev.github.io/r7-plugin-packages/store/translations/langs.json';
			xhr.open('GET', url, true);

			xhr.onload = function () {
				if (this.readyState == 4) {
					if (this.status >= 200 && this.status < 300) {
						endInternetChecking(true);
					}
				}
			};

			xhr.onerror = function (err) {
				endInternetChecking(false);
			};

			xhr.send(null);
		} catch (error) {
			endInternetChecking(false);
		}
		if (bSetTimeout) {
			errTimeout = setTimeout(function () {
				// if loading is too long show the error (because sometimes requests can not send error)
				endInternetChecking(false);
			}, 15000);
		}
	};

	function endInternetChecking(isOnline) {
		clearTimeout(errTimeout);
		errTimeout = null;
		destroyLoader();
		if (isOnline) {
			initPlugin();
		} else {
			document.getElementById('div_noIternet').classList.remove('hidden');
			if (!interval) {
				interval = setInterval(function () {
					checkInternet(false);
				}, 5000);
			}
		}
	};

	function createWindow(type) {
		let fileName = type + '.html';
		let description = window.Asc.plugin.tr('Warning');
		let size = [350, 100];
		let buttons = [
			{
				'text': window.Asc.plugin.tr('Yes'),
				'primary': true
			},
			{
				'text': window.Asc.plugin.tr('No'),
				'primary': false
			}
		];
		if (type == 'developer') {
			fileName = 'developer.html';
			description = window.Asc.plugin.tr('Developer Mode');
			size = [500, 150];
			buttons = [
				{
					'text': window.Asc.plugin.tr('Ok'),
					'primary': true
				},
				{
					'text': window.Asc.plugin.tr('Cancel'),
					'primary': false
				}
			];
		}
		let location = window.location;
		let start = location.pathname.lastIndexOf('/') + 1;
		let file = location.pathname.substring(start);

		let variation = {
			url: location.href.replace(file, fileName),
			description: description,
			isVisual: true,
			isModal: true,
			isViewer: true,
			EditorsSupport: ['word', 'cell', 'slide', 'pdf'],
			size: size,
			buttons: buttons
		};

		if (type == 'warning') {
			if (!warningWindow) {
				warningWindow = new window.Asc.PluginWindow();
			}
			warningWindow.show(variation);
		} else {
			if (!developerWindow) {
				developerWindow = new window.Asc.PluginWindow();
				developerWindow.attachEvent('onWindowMessage', function (message) {
					if (message.type == 'SetURL') {
						let source = message.url;
						let customRepo = normalizeMarketplaceRepo(source);
						let customUrl = normalizeMarketplaceUrl(source);
						if (customRepo.length) {
							marketplaceRepo = customRepo;
							marketplaceURl = OOMarketplaceUrl;
							localStorage.setItem('DeveloperMarketplaceRepo', marketplaceRepo);
							localStorage.removeItem('DeveloperMarketplaceUrl');
							document.getElementById('notification').classList.remove('hidden');
						} else if (customUrl.length) {
							marketplaceRepo = '';
							marketplaceURl = customUrl;
							localStorage.removeItem('DeveloperMarketplaceRepo');
							localStorage.setItem('DeveloperMarketplaceUrl', marketplaceURl);
							document.getElementById('notification').classList.remove('hidden');
						} else {
							marketplaceRepo = '';
							marketplaceURl = OOMarketplaceUrl;
							localStorage.removeItem('DeveloperMarketplaceRepo');
							localStorage.removeItem('DeveloperMarketplaceUrl');
							document.getElementById('notification').classList.add('hidden');
						}
						iframe.src = marketplaceURl + window.location.search;
					}
					window.Asc.plugin.executeMethod('CloseWindow', [developerWindow.id]);
				});
			}
			developerWindow.show(variation);
		}

	};

	function removePlugin(backup) {
		if (removeGuid) {
			let guid = removeGuid;
			window.Asc.plugin.executeMethod('RemovePlugin', [removeGuid, backup], function (result) {
				postMessage(normalizeHostActionResult('remove', guid, backup, result));
			});
		}

		removeGuid = null;
	};

	window.onresize = function () {
		// zoom for all elements in plugin window
		let zoom = 1;
		if (window.devicePixelRatio < 1)
			zoom = 1 / window.devicePixelRatio;

		document.getElementsByTagName('html')[0].style.zoom = zoom;
	};

	function createLoader() {
		$('#loader-container').removeClass('hidden');
		loader && (loader.remove ? loader.remove() : $('#loader-container')[0].removeChild(loader));
		loader = showLoader($('#loader-container')[0], window.Asc.plugin.tr('Checking internet connection...'));
	};

	function destroyLoader() {
		clearTimeout(loaderTimeout);
		loaderTimeout = null;
		$('#loader-container').addClass('hidden')
		loader && (loader.remove ? loader.remove() : $('#loader-container')[0].removeChild(loader));
		loader = undefined;
	};

})(window, undefined);
