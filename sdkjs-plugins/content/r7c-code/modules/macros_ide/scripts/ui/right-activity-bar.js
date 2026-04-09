try {
(function () {
  // Prevent duplicate bootstrap if this script is injected more than once
  try {
    if (typeof window !== 'undefined' && window.__RightActivityBarInitialized) {
      window.debug?.warn?.('RightActivityBar', 'Script already initialized, skipping duplicate bootstrap')
      return
    }
    if (typeof window !== 'undefined') {
      window.__RightActivityBarInitialized = true
    }
  } catch (_) {}
  /**
   * ADR References:
   * - ADR-018-right-activity-bar
   * - ADR-032-plugin-ai-chat-and-settings
   * - ADR-057-github-models-api-integration
   * - ADR-069-ai-job-manager-architecture
   * - ADR-070-ai-job-monitor-right-pane-tool
   */
  const SCRIPT_VERSION = '2024-10-25T12:05Z';
  try {
    console.info('[RightActivityBar] bootstrap start', { version: SCRIPT_VERSION });
  } catch (_) {}
  const STORAGE_KEY = 'right_pane_open';
  const WIDTH_KEY = 'right_pane_width';
  const PANE_WIDTH_VAR = '--right-pane-width';
  const DEFAULT_WIDTH = 450;
  const PANE_AI_CONNECTION = 'ai-connection';
  const PANE_AI_MANAGEMENT = 'ai-management';
  const PANE_CHAT = 'chat';
  const PANE_CODE_GRAPH = 'code-graph';
  const PANE_DEBUG = 'debug-insights';
  const PANE_AI_JOBS = 'ai-job-monitor';
  const AI_FREE_PROVIDER_ID = 'openai';
  const AI_FREE_LOCKED_MODEL_ID = 'gpt-4o-mini';
  const AI_FREE_RESTRICTIONS_ENABLED = false;
  const TestId = typeof window !== 'undefined' ? window.TestId : null;
  let currentPane = null; // e.g., 'ai-connection'
  let currentAgentsTab = 'connection';
  const CURRENT_KEY = 'right_pane_current';
  const SETTINGS_STORE = window.SettingsStore || null;
  // NOTE: AIPromptService is resolved lazily via window.AIPromptService to avoid script-order issues
  const PROVIDER_MODEL_CACHE = window.__aiPromptModels || (window.__aiPromptModels = new Map());
  const PROVIDER_TEXT_CACHE = window.__aiPromptTextCache || (window.__aiPromptTextCache = new Map());
  window.__aiPromptProvider = window.__aiPromptProvider || null;
  let agentTreeInstance = null;
  let initInProgress = false;
  let initCompleted = false;
  let codeGraphHost = null;
  let codeGraphViewMounted = false;
  let codeGraphContext = null;
  let aiJobMonitorHost = null;
  let aiJobMonitorMounted = false;
  let pendingDebugPluginsLoadedListener = null

  function normalizePaneKey(value) {
    if (!value) return null;
    if (value === 'ai' || value === PANE_AI_MANAGEMENT) return PANE_AI_CONNECTION;
    if (value === PANE_AI_CONNECTION || value === PANE_AI_MANAGEMENT || value === PANE_CHAT || value === PANE_CODE_GRAPH || value === PANE_DEBUG) {
      return value;
    }
    return null;
  }

  function showProFeatureNotice(options = {}) {
    const featureName = options.featureName || 'This feature';
    const subtitle = options.subtitle || 'Available in Pro version';
    if (typeof document === 'undefined') return;

    const existing = document.getElementById('ai-pro-feature-modal');
    if (existing) existing.remove();

    const overlay = document.createElement('div');
    overlay.id = 'ai-pro-feature-modal';
    overlay.style.cssText = [
      'position:fixed',
      'inset:0',
      'background:rgba(0,0,0,0.55)',
      'display:flex',
      'align-items:center',
      'justify-content:center',
      'z-index:10000'
    ].join(';');

    const card = document.createElement('div');
    card.style.cssText = [
      'width:min(520px,92vw)',
      'background:var(--vscode-editor-background,#1e1e1e)',
      'color:var(--vscode-foreground,#d4d4d4)',
      'border:1px solid var(--vscode-widget-border,#3c3c3c)',
      'border-radius:12px',
      'box-shadow:0 16px 48px rgba(0,0,0,0.45)',
      'padding:18px 18px 14px'
    ].join(';');

    card.innerHTML = `
      <div style="display:flex; align-items:flex-start; gap:12px;">
        <div style="font-size:20px; line-height:1;">🔒</div>
        <div style="flex:1; min-width:0;">
          <div style="font-weight:700; font-size:15px; margin-bottom:4px;">${featureName}</div>
          <div style="color:var(--vscode-descriptionForeground,#9da5b4); font-size:13px; line-height:1.45;">${subtitle}</div>
          <div style="margin-top:10px; font-size:12px; color:var(--vscode-descriptionForeground,#9da5b4);">In free mode AI Sputnik runs with OpenAI and fixed model preset for stable macro workflows.</div>
        </div>
      </div>
      <div style="display:flex; justify-content:flex-end; margin-top:14px;">
        <button type="button" class="settings-btn settings-btn-primary" id="ai-pro-feature-modal-ok">Got it</button>
      </div>
    `;

    const close = () => {
      try { overlay.remove(); } catch (_) {}
    };
    overlay.addEventListener('click', (event) => {
      if (event.target === overlay) close();
    });
    card.querySelector('#ai-pro-feature-modal-ok')?.addEventListener('click', close);
    overlay.appendChild(card);
    document.body.appendChild(overlay);
  }

  function buildModuleContext(overrides = {}) {
    const base = { pluginState: window.pluginState, ...overrides };
    try {
      if (window.ModuleContextFactory && typeof window.ModuleContextFactory.create === 'function') {
        return window.ModuleContextFactory.create(base);
      }
    } catch (_) {}
    return base;
  }

  function getRightPaneWidthBounds() {
    const viewport = window.innerWidth || document.documentElement.clientWidth || 1200
    const min = 220
    const max = Math.max(min, Math.min(700, viewport - 300))
    return { min, max }
  }

  function __isPrimarySelectActivationEvent(event) {
    if (!event) return true
    if (event.button !== undefined && event.button > 0) return false
    if (event.which !== undefined && event.which > 1) return false
    return true
  }

  function __closeSelectPicker(pickerEl) {
    if (!pickerEl) return
    try { pickerEl.__cleanup?.() } catch (_) {}
    try { pickerEl.remove() } catch (_) {}
  }

  // Desktop runtimes do not always deliver a native mousedown sequence reliably.
  // Bind a small activation shim so the custom picker still opens from click/pointer.
  function __bindStableSelectActivation(target, onActivate, options = {}) {
    if (!target || typeof target.addEventListener !== 'function' || typeof onActivate !== 'function') return
    const bindingKey = options.bindingKey || '__stableSelectActivationBound'
    if (target[bindingKey]) return
    target[bindingKey] = true

    let lastPressEventAt = 0
    let lastActivationAt = 0

    const handleActivate = (event, reason) => {
      const isKeyboard = reason === 'keydown'
      const now = Date.now()

      if (!isKeyboard && !__isPrimarySelectActivationEvent(event)) {
        return
      }

      if (reason === 'pointerdown' || reason === 'mousedown') {
        if ((now - lastPressEventAt) < 80) return
        lastPressEventAt = now
      } else if (reason === 'click' || reason === 'pointerup') {
        if ((now - lastPressEventAt) < 250) return
      }

      if ((now - lastActivationAt) < 80) return
      lastActivationAt = now

      if (event) {
        event.preventDefault()
        event.stopPropagation()
      }

      onActivate({
        byKeyboard: isKeyboard,
        event,
        reason
      })
    }

    target.addEventListener('pointerdown', (event) => handleActivate(event, 'pointerdown'), true)
    target.addEventListener('mousedown', (event) => handleActivate(event, 'mousedown'), true)
    target.addEventListener('pointerup', (event) => handleActivate(event, 'pointerup'), true)
    target.addEventListener('click', (event) => handleActivate(event, 'click'), true)
    target.addEventListener('keydown', (event) => {
      const key = event.key || event.which
      if (key !== 'Enter' && key !== ' ' && key !== 'Spacebar' && key !== 'ArrowDown' && key !== 'ArrowUp' && key !== 13 && key !== 32 && key !== 38 && key !== 40) {
        return
      }
      handleActivate(event, 'keydown')
    })
  }

  // ADR-023: Custom dropdown overlay helpers (for reliable selects inside panels)
  function __openSelectPicker(selectEl, hostEl, byKeyboard = false) {
    if (!selectEl || !hostEl) return
    const id = `picker-${selectEl.id || selectEl.name || selectEl.className || 'generic'}`
    const existing = document.getElementById(id)
    if (existing) {
      __closeSelectPicker(existing)
    }
    const hostRect = hostEl.getBoundingClientRect()
    const rect = selectEl.getBoundingClientRect()
    const picker = document.createElement('div')
    picker.id = id
    picker.className = 'settings-dropdown'
    try {
      const owner = selectEl?.id || selectEl?.name || 'generic'
      window.TestId?.setTestId?.(picker, `activity-right-picker-${owner}`)
    } catch (_) {}
    picker.style.position = 'absolute'
    picker.style.left = `${Math.max(0, rect.left - hostRect.left)}px`
    picker.style.top = `${Math.max(0, rect.bottom - hostRect.top + 4)}px`
    picker.style.minWidth = `${Math.max(rect.width, 200)}px`
    picker.style.width = `${Math.max(rect.width, 200)}px`
    picker.style.maxWidth = `${Math.max(rect.width, 200)}px`
    picker.style.maxHeight = '240px'
    picker.style.overflowY = 'auto'
    picker.style.overflowX = 'hidden'
    picker.style.background = 'var(--color-bg-primary, #1f1f1f)'
    picker.style.color = 'var(--color-text-primary, #d4d4d4)'
    picker.style.border = '1px solid var(--color-border-default, #3c3c3c)'
    picker.style.borderRadius = '6px'
    picker.style.boxShadow = '0 8px 24px rgba(0,0,0,0.4)'
    picker.style.zIndex = '3000'
    picker.style.padding = '4px'
    const list = document.createElement('div')
    list.setAttribute('role', 'listbox')
    try {
      const owner = selectEl?.id || selectEl?.name || 'generic'
      window.TestId?.setTestId?.(list, `activity-right-picker-list-${owner}`)
    } catch (_) {}
    list.style.display = 'flex'
    list.style.flexDirection = 'column'
    list.style.gap = '2px'
    picker.appendChild(list)

    const buildItem = (opt) => {
      const item = document.createElement('div')
      item.setAttribute('role', 'option')
      item.tabIndex = 0
      item.textContent = opt.textContent || opt.value
      item.dataset.value = opt.value
      try {
        const owner = selectEl?.id || selectEl?.name || 'generic'
        const value = window.TestId?.toKebab?.(opt.value) || 'option'
        window.TestId?.setTestId?.(item, `activity-right-picker-option-${owner}-${value}`)
      } catch (_) {}
      item.style.padding = '6px 10px'
      item.style.borderRadius = '4px'
      item.style.cursor = 'pointer'
      item.style.outline = 'none'
      item.style.display = 'block'
      item.style.width = '100%'
      item.style.minWidth = '0'
      item.style.boxSizing = 'border-box'
      item.style.whiteSpace = 'nowrap'
      item.style.overflow = 'hidden'
      item.style.textOverflow = 'ellipsis'
      item.title = opt.textContent || opt.value || ''
      const isActive = selectEl.value === opt.value
      item.style.background = isActive ? 'var(--vscode-list-focusBackground,#094771)' : 'transparent'
      item.style.color = isActive ? 'var(--vscode-list-focusForeground,#ffffff)' : 'var(--vscode-foreground,#d4d4d4)'
      item.addEventListener('mouseenter', () => { item.style.background = 'var(--vscode-list-hoverBackground, rgba(255,255,255,0.08))' })
      item.addEventListener('mouseleave', () => {
        const active = (selectEl.value === opt.value)
        item.style.background = active ? 'var(--vscode-list-focusBackground,#094771)' : 'transparent'
        item.style.color = active ? 'var(--vscode-list-focusForeground,#ffffff)' : 'var(--vscode-foreground,#d4d4d4)'
      })
      const commit = () => {
        try {
          selectEl.value = opt.value
          selectEl.dispatchEvent(new Event('change', { bubbles: true }))
        } catch (_) {}
        __closeSelectPicker(picker)
      }
      item.addEventListener('click', commit)
      item.addEventListener('keydown', (event) => {
        if (event.key === 'Enter' || event.key === ' ') {
          event.preventDefault()
          commit()
        }
      })
      return item
    }

    Array.from(selectEl.options || []).forEach((opt) => list.appendChild(buildItem(opt)))
    hostEl.appendChild(picker)

    const onOutside = (event) => {
      if (picker.contains(event.target) || event.target === selectEl) return
      __closeSelectPicker(picker)
    }
    const onKey = (event) => {
      if (event.key !== 'Escape') return
      event.preventDefault()
      __closeSelectPicker(picker)
      try { selectEl.focus?.() } catch (_) {}
    }

    picker.__cleanup = () => {
      window.removeEventListener('pointerdown', onOutside, true)
      window.removeEventListener('mousedown', onOutside, true)
      window.removeEventListener('click', onOutside, true)
      window.removeEventListener('keydown', onKey, true)
    }

    window.addEventListener('pointerdown', onOutside, true)
    window.addEventListener('mousedown', onOutside, true)
    window.addEventListener('click', onOutside, true)
    window.addEventListener('keydown', onKey, true)

    if (byKeyboard) {
      const active = list.querySelector('[role="option"]')
      try { active && active.focus && active.focus() } catch (_) {}
    }
  }

  function __attachCustomPicker(selectEl, hostEl) {
    if (!selectEl || !hostEl) return
    __bindStableSelectActivation(selectEl, ({ byKeyboard }) => {
      __openSelectPicker(selectEl, hostEl, byKeyboard)
    }, { bindingKey: '__customPickerBound' })
  }

  function performInit() {
    try { console.info('[RightActivityBar] init invoked', { version: SCRIPT_VERSION }); } catch (_) {}
    ensureContainers();
    logDebug('init:containers', {
      hasBar: !!document.getElementById('vscode-activity-bar-right'),
      hasPane: !!document.getElementById('vscode-right-pane')
    });
    // Ensure width variable exists
    try {
      let saved = null; try { saved = parseInt(localStorage.getItem(WIDTH_KEY), 10); } catch(_){ saved = null; }
      if (!Number.isFinite(saved) || saved < 120 || saved > 1200) saved = DEFAULT_WIDTH;
      const { min, max } = getRightPaneWidthBounds()
      if (saved < min) saved = min
      if (saved > max) saved = max
      document.documentElement.style.setProperty(PANE_WIDTH_VAR, saved + 'px');
      try { localStorage.setItem(WIDTH_KEY, String(saved)); } catch (_) {}
    } catch (_) {}
    renderBar();
    // Register global listeners once, with teardown handle
	    try {
	      if (!window.__RightActivityBarListeners) {
	        const onModulesChanged = (e) => { try { renderBar() } catch (_) {} }
	        window.addEventListener('modules:changed', onModulesChanged)
	        window.__RightActivityBarTeardown = function () {
	          try { window.removeEventListener('modules:changed', onModulesChanged) } catch (_) {}
	          try {
	            if (window.__aiConnectionsUpdatedListener) {
	              document.removeEventListener('aiConnections:updated', window.__aiConnectionsUpdatedListener)
	            }
	          } catch (_) {}
	          try { window.__aiConnectionsUpdatedListener = null } catch (_) {}
	          try { window.__aiConnectionsListenerBound = false } catch (_) {}
	          try {
	            if (window.__aiAgentSelectedListener) {
	              document.removeEventListener('ai:agent-selected', window.__aiAgentSelectedListener)
	            }
	          } catch (_) {}
	          try { window.__aiAgentSelectedListener = null } catch (_) {}
	          try { window.__aiAgentSelectedListenerBound = false } catch (_) {}
	          try {
	            if (window.__aiAgentStoreUnsubscribe && typeof window.__aiAgentStoreUnsubscribe === 'function') {
	              window.__aiAgentStoreUnsubscribe()
	            }
	          } catch (_) {}
	          try { window.__aiAgentStoreUnsubscribe = null } catch (_) {}
	          try {
	            if (pendingDebugPluginsLoadedListener) {
	              window.removeEventListener('debug-plugins-loaded', pendingDebugPluginsLoadedListener)
	              window.removeEventListener('debug-plugins-integration-ready', pendingDebugPluginsLoadedListener)
	            }
          } catch (_) {}
          try {
            if (window.__aiConnectionsSubscription && typeof window.__aiConnectionsSubscription === 'function') {
              window.__aiConnectionsSubscription()
            }
          } catch (_) {}
          try { window.__aiConnectionsSubscription = null } catch (_) {}
          pendingDebugPluginsLoadedListener = null
          window.__RightActivityBarListeners = false
        }
        window.__RightActivityBarListeners = true
      }
    } catch(_) {}
    // Restore pane state (default: open on first load)
    let storedPane = null;
    try {
      storedPane = localStorage.getItem(CURRENT_KEY) || null;
      currentPane = normalizePaneKey(storedPane);
    } catch(_) {}
    let storedOpen = null;
    try { storedOpen = localStorage.getItem(STORAGE_KEY); } catch(_) {}
    const registry = window.ModuleRegistry;
    const aiEnabled = !registry || registry.isEnabled?.('ai') !== false;
    // UI request: hide Code Graph button from right activity bar
    const codeGraphEnabled = false;
    logDebug('init:storageState', {
      storedPane,
      currentPane,
      storedOpen,
      registryEnabled: aiEnabled
    });
    if (storedPane && storedPane !== currentPane) {
      try {
        if (currentPane) localStorage.setItem(CURRENT_KEY, currentPane);
        else localStorage.removeItem(CURRENT_KEY);
      } catch (_) {}
    }
    if (!aiEnabled && (currentPane === PANE_AI_CONNECTION || currentPane === PANE_AI_MANAGEMENT)) {
      currentPane = codeGraphEnabled ? PANE_CODE_GRAPH : null;
      try {
        if (currentPane) localStorage.setItem(CURRENT_KEY, currentPane);
        else localStorage.removeItem(CURRENT_KEY);
      } catch(_) {}
    }
    if (!codeGraphEnabled && currentPane === PANE_CODE_GRAPH) {
      currentPane = aiEnabled ? PANE_AI_CONNECTION : null;
      try {
        if (currentPane) localStorage.setItem(CURRENT_KEY, currentPane);
        else localStorage.removeItem(CURRENT_KEY);
      } catch(_) {}
    }
    if (!currentPane) {
      currentPane = aiEnabled ? PANE_AI_CONNECTION : (codeGraphEnabled ? PANE_CODE_GRAPH : null);
    }
    const content = document.getElementById('right-pane-content');
    if (content) content.innerHTML = '';
    setOpen(false);
    try { localStorage.setItem(STORAGE_KEY, 'false'); } catch (_) {}
    initResizer();
  }

  const runInit = (reason) => {
    if (initCompleted || initInProgress) {
      try { console.info('[RightActivityBar] init skipped', { version: SCRIPT_VERSION, reason }); } catch (_) {}
      return;
    }
    initInProgress = true;
    try {
      console.info('[RightActivityBar] init trigger', { version: SCRIPT_VERSION, reason });
      performInit();
      initCompleted = true;
      console.info('[RightActivityBar] init complete', { version: SCRIPT_VERSION, reason });
    } catch (error) {
      try { console.error('[RightActivityBar] init error', { version: SCRIPT_VERSION, reason, message: error?.message, stack: error?.stack }); } catch (_) {}
    } finally {
      initInProgress = false;
    }
  };
  try {
    window.__RightActivityBar_runInit = runInit;
  } catch (_) {}

  try {
    console.info('[RightActivityBar] constants initialised', { version: SCRIPT_VERSION });
  } catch (_) {}
  try {
    setTimeout(() => {
      try {
        const state = typeof document !== 'undefined' ? document.readyState : 'no-document';
        console.info(`[RightActivityBar] async heartbeat → ${state}`, { version: SCRIPT_VERSION });
        if (state === 'loading') {
          console.info('[RightActivityBar] heartbeat deferred (document.loading)', { version: SCRIPT_VERSION });
        } else {
          window.__RightActivityBar_runInit?.('heartbeat');
        }
        setTimeout(() => {
          const state2 = typeof document !== 'undefined' ? document.readyState : 'no-document';
          console.info(`[RightActivityBar] async heartbeat (late) → ${state2}`, { version: SCRIPT_VERSION });
          if (!initCompleted && state2 !== 'loading') {
            window.__RightActivityBar_runInit?.('heartbeat-late');
          }
        }, 200);
      } catch (_) {}
    }, 0);
  } catch (error) {
    try { console.error('[RightActivityBar] heartbeat scheduling failed', error); } catch (_) {}
  }

  function logDebug(event, payload) {
    try {
      window.LogService?.logInfo('RightActivityBar', event, payload || undefined);
    } catch (_) {}
    try {
      window.debug?.info?.('RightActivityBar', event, payload || undefined);
    } catch (_) {}
  }
  try {
    console.info('[RightActivityBar] post-logDebug checkpoint', { version: SCRIPT_VERSION });
  } catch (_) {}

  function ensureContainers() {
    try {
      let bar = document.getElementById('vscode-activity-bar-right');
      if (!bar) {
        bar = document.createElement('div');
        bar.id = 'vscode-activity-bar-right';
        TestId?.setTestId?.(bar, 'activity-right-bar');
        document.body.appendChild(bar);
        console.info('[RightActivityBar] bar container created');
      } else {
        TestId?.setTestId?.(bar, 'activity-right-bar');
        console.info('[RightActivityBar] bar container reused');
      }
      let pane = document.getElementById('vscode-right-pane');
      if (!pane) {
        pane = document.createElement('div');
        pane.id = 'vscode-right-pane';
        pane.className = 'vscode-right-pane';
        TestId?.setTestId?.(pane, 'activity-right-pane');
        pane.innerHTML = `
        <div class="right-pane-title">
          <span class="title-text">AI Utilities</span>
          <div class="title-actions">
            <button class="btn-action" id="right-pane-close" title="Close" data-testid="activity-right-pane-close-btn">✕</button>
          </div>
        </div>
        <div class="right-pane-content" id="right-pane-content" data-testid="activity-right-pane-content"></div>
        <div class="right-pane-resizer" id="right-pane-resizer" title="Resize" data-testid="activity-right-pane-resizer"></div>
      `;
        document.body.appendChild(pane);
        pane.querySelector('#right-pane-close')?.addEventListener('click', () => setOpen(false));
        console.info('[RightActivityBar] pane container created');
      } else {
        TestId?.setTestId?.(pane, 'activity-right-pane');
        TestId?.setTestId?.(pane.querySelector('#right-pane-content'), 'activity-right-pane-content');
        TestId?.setTestId?.(pane.querySelector('#right-pane-resizer'), 'activity-right-pane-resizer');
        TestId?.setTestId?.(pane.querySelector('#right-pane-close'), 'activity-right-pane-close-btn');
        console.info('[RightActivityBar] pane container reused');
      }
    } catch (error) {
      try { console.error('[RightActivityBar] ensureContainers failed', error); } catch (_) {}
    }
  }
  try {
    console.info('[RightActivityBar] post-ensureContainers checkpoint', { version: SCRIPT_VERSION });
  } catch (_) {}

  function updateAgentsButtonHighlight() {
    const bar = document.getElementById('vscode-activity-bar-right');
    if (!bar) return;
    const btnConnection = bar.querySelector(`[data-button="${PANE_AI_CONNECTION}"]`);
    const btnManagement = bar.querySelector(`[data-button="${PANE_AI_MANAGEMENT}"]`);
    const open = isOpen();
    if (btnConnection) {
      const activeConnection = open && currentPane === PANE_AI_CONNECTION;
      btnConnection.classList.toggle('active', activeConnection);
      btnConnection.setAttribute('aria-selected', activeConnection ? 'true' : 'false');
    }
    if (btnManagement) {
      const activeManagement = open && currentPane === PANE_AI_MANAGEMENT;
      btnManagement.classList.toggle('active', activeManagement);
      btnManagement.setAttribute('aria-selected', activeManagement ? 'true' : 'false');
      if (activeManagement) {
        btnManagement.dataset.activeTab = currentAgentsTab || '';
      } else {
        delete btnManagement.dataset.activeTab;
      }
    }
  }

  function renderBar() {
    const bar = document.getElementById('vscode-activity-bar-right');
    if (!bar) return;
    const registry = window.ModuleRegistry;
    const aiEnabled = !registry || registry.isEnabled?.('ai') !== false;
    // UI request: hide Code Graph button from right activity bar
    const codeGraphEnabled = false;
    let paneChanged = false;
    if (!aiEnabled && (currentPane === PANE_AI_CONNECTION || currentPane === PANE_AI_MANAGEMENT)) {
      currentPane = codeGraphEnabled ? PANE_CODE_GRAPH : null;
      paneChanged = true;
      try {
        if (currentPane) localStorage.setItem(CURRENT_KEY, currentPane);
        else localStorage.removeItem(CURRENT_KEY);
      } catch (_) {}
    }
    if (!codeGraphEnabled && currentPane === PANE_CODE_GRAPH) {
      currentPane = aiEnabled ? PANE_AI_CONNECTION : null;
      paneChanged = true;
      try {
        if (currentPane) localStorage.setItem(CURRENT_KEY, currentPane);
        else localStorage.removeItem(CURRENT_KEY);
      } catch (_) {}
    }
    if (!currentPane && aiEnabled) {
      currentPane = PANE_AI_CONNECTION;
      paneChanged = true;
      try { localStorage.setItem(CURRENT_KEY, currentPane); } catch (_) {}
    }
    if (!currentPane && !aiEnabled && !codeGraphEnabled) {
      try { localStorage.removeItem(CURRENT_KEY); } catch (_) {}
    }
    bar.innerHTML = '';
    const frag = document.createDocumentFragment();
    // 1) Toggle pane button
    logDebug('renderBar:start', {
      open: isOpen(),
      pane: currentPane,
      tab: currentAgentsTab,
      aiEnabled,
      codeGraphEnabled
    });
    const btnToggle = document.createElement('button');
    btnToggle.className = 'activity-bar-icon';
    btnToggle.dataset.button = 'toggle';
    TestId?.setTestId?.(btnToggle, 'activity-right-toggle-pane-btn');
    btnToggle.title = 'Toggle AI Utilities';
    btnToggle.setAttribute('aria-label', 'Toggle AI Utilities');
    btnToggle.innerHTML = '<svg width="24" height="24" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><circle cx="12" cy="12" r="10" fill="currentColor"/><path d="M16.3 5.3a7.3 7.3 0 0 0-8.7.7 6.8 6.8 0 0 0-1.9 2.5l-.3.6-1.2-2.9C6.9 3.1 9.3 2 12 2c2.2 0 4.4.8 6.1 2.3l-1.8 1Z" fill="#ffffff"/><path d="M7.7 18.7a7.3 7.3 0 0 0 8.7-.7 6.8 6.8 0 0 0 1.9-2.5l.3-.6 1.2 2.9C17.1 20.9 14.7 22 12 22c-2.2 0-4.4-.8-6.1-2.3l1.8-1Z" fill="#ffffff"/></svg>';
    btnToggle.addEventListener('click', toggleRightPane);
    frag.appendChild(btnToggle);

    // 2) AI Connection button
    const btnConnection = document.createElement('button');
    btnConnection.className = 'activity-bar-icon';
    btnConnection.dataset.button = PANE_AI_CONNECTION;
    TestId?.setTestId?.(btnConnection, 'activity-right-pane-ai-connection-btn');
    btnConnection.title = 'AI Connection Settings';
    btnConnection.setAttribute('aria-label', 'AI Connection Settings');
    btnConnection.innerHTML = '<svg width="24" height="24" viewBox="0 0 24 24" fill="currentColor" xmlns="http://www.w3.org/2000/svg"><path d="M9 2h6v5h1a3 3 0 0 1 3 3v3a6 6 0 0 1-5 5.74V22h-4v-3.26A6 6 0 0 1 5 13V10a3 3 0 0 1 3-3h1V2zm2 0v3h2V2h-2z"/></svg>';
    btnConnection.addEventListener('click', openAIConnectionPane);
    btnConnection.disabled = !aiEnabled;
    btnConnection.setAttribute('aria-disabled', aiEnabled ? 'false' : 'true');
    frag.appendChild(btnConnection);

    // 4) Chat button
    const btnChat = document.createElement('button');
    btnChat.className = 'activity-bar-icon';
    btnChat.dataset.button = PANE_CHAT;
    TestId?.setTestId?.(btnChat, 'activity-right-pane-chat-btn');
    btnChat.title = 'AI Chat';
    btnChat.setAttribute('aria-label', 'AI Chat');
    btnChat.innerHTML = '<svg width="24" height="24" viewBox="0 0 24 24" fill="currentColor" xmlns="http://www.w3.org/2000/svg"><path d="M4 4h16v12H6l-2 2V4z"/></svg>';
    btnChat.addEventListener('click', () => {
      if (window.Trial && window.Trial.isExpired && window.Trial.isExpired()) {
        window.Trial.showExpiredOverlay();
        return;
      }
      openChatPane();
    });
    btnChat.disabled = !aiEnabled;
    btnChat.setAttribute('aria-disabled', aiEnabled ? 'false' : 'true');
    frag.appendChild(btnChat);

    let btnCodeGraph = null;
    let btnDebug = null;

    // Debug Insights button (always render; click will mount when available)
    btnDebug = document.createElement('button');
    btnDebug.className = 'activity-bar-icon';
    btnDebug.dataset.button = PANE_DEBUG;
    TestId?.setTestId?.(btnDebug, 'activity-right-pane-debug-insights-btn');
    btnDebug.title = 'Debug Insights';
    btnDebug.setAttribute('aria-label', 'Debug Insights');
    // Use graph icon to reflect embedded Code Graph report
    btnDebug.innerHTML = '<svg width="24" height="24" viewBox="0 0 24 24" fill="currentColor" xmlns="http://www.w3.org/2000/svg"><path d="M22 11V3h-7v3H9V3H2v8h7V8h2v10h4v3h7v-8h-7v3h-2V8h2v3h7zM7 9H4V5h3v4zm10 6h3v4h-3v-4zm0-10h3v4h-3V5z"/></svg>';
    btnDebug.addEventListener('click', openDebugPane);
    frag.appendChild(btnDebug);
    // Single DOM append to minimize layout thrash
    bar.appendChild(frag);

    const open = isOpen();
    if (open) btnToggle.classList.add('active');
    btnChat.classList.toggle('active', open && currentPane === PANE_CHAT);
    if (btnCodeGraph) btnCodeGraph.classList.toggle('active', open && currentPane === PANE_CODE_GRAPH);
    if (btnDebug) btnDebug.classList.toggle('active', open && currentPane === PANE_DEBUG);
    updateAgentsButtonHighlight();
    if (paneChanged && open) {
      if (currentPane) renderActivePane();
      else setOpen(false);
    }
    // Apply ModuleRegistry visibility (AI/CodeGraph/Debug can be disabled)
    try {
      const debugEnabled = !!window.debugPluginsIntegration || (!!window.ModuleRegistry && window.ModuleRegistry.isEnabled?.('debug-insights') !== false);
      const enabled = aiEnabled || codeGraphEnabled || debugEnabled;
      const barEl = document.getElementById('vscode-activity-bar-right');
      const paneEl = document.getElementById('vscode-right-pane');
      logDebug('renderBar:moduleState', {
        enabled,
        codeGraphEnabled,
        debugEnabled,
        hasRegistry: !!window.ModuleRegistry,
        storedModules: window.ModuleRegistry?.getEnabledMap?.()
      });
      if (barEl) {
        barEl.classList.toggle('module-disabled', !enabled);
        barEl.setAttribute('aria-disabled', enabled ? 'false' : 'true');
      }
      if (paneEl) paneEl.classList.toggle('is-disabled', !enabled);
      if (!enabled) {
        document.body.classList.remove('right-pane-open');
        try { localStorage.setItem('right_pane_open','false'); } catch(_) {}
      }
    } catch(_) {}
  }

  function isOpen() {
    try { return localStorage.getItem(STORAGE_KEY) === 'true'; } catch (_) { return false; }
  }

  function setOpen(open) {
    const nextOpen = !!open
    const wasOpen = document.body.classList.contains('right-pane-open')
    document.body.classList.toggle('right-pane-open', nextOpen)
    const pane = document.getElementById('vscode-right-pane');
    if (pane) {
      pane.setAttribute('aria-hidden', nextOpen ? 'false' : 'true');
      try { pane.setAttribute('data-testid-state', nextOpen ? 'open' : 'closed') } catch (_) {}
    }
    try { localStorage.setItem(STORAGE_KEY, String(nextOpen)); } catch (_) {}
    const bar = document.getElementById('vscode-activity-bar-right');
    const btnToggle = bar?.querySelector('[data-button="toggle"]');
    if (btnToggle) btnToggle.classList.toggle('active', nextOpen);
    const btnChat = bar?.querySelector(`[data-button="${PANE_CHAT}"]`);
    if (btnChat) btnChat.classList.toggle('active', nextOpen && currentPane === PANE_CHAT);
    const btnCodeGraph = bar?.querySelector(`[data-button="${PANE_CODE_GRAPH}"]`);
    if (btnCodeGraph) btnCodeGraph.classList.toggle('active', nextOpen && currentPane === PANE_CODE_GRAPH);
    const btnDebug = bar?.querySelector(`[data-button="${PANE_DEBUG}"]`);
    if (btnDebug) btnDebug.classList.toggle('active', nextOpen && currentPane === PANE_DEBUG);
    if (!nextOpen && agentTreeInstance && typeof agentTreeInstance.dispose === 'function') {
      try { agentTreeInstance.dispose(); } catch (_) {}
      agentTreeInstance = null;
    }

    // Align editor wrapper with right pane state (must override inline styles from tree toggle)
    try {
      const wrapper = document.getElementById('editorWrapper')
      if (wrapper) {
        if (nextOpen) {
          try { initResizer() } catch (_) {}
          const desired = 'calc(48px + var(--right-pane-width, 450px))'
          if (wrapper.style.right !== desired) wrapper.style.right = desired
        } else {
          const sideBar = document.getElementById('vscode-sidebar')
          let sidebarHidden = false
          try {
            if (!sideBar) sidebarHidden = true
            else sidebarHidden = window.getComputedStyle(sideBar).display === 'none'
          } catch (_) {}
          const desired = sidebarHidden ? '0' : ''
          if (wrapper.style.right !== desired) wrapper.style.right = desired
        }
      }
    } catch (_) {}

    // Trigger Monaco layout only when right pane visibility changes
    try {
      const editor = window.monacoEditorManager?.getEditor?.() ||
        (window.editor && typeof window.editor.getEditor === 'function'
          ? window.editor.getEditor()
          : window.editor);
      if (editor && typeof editor.layout === 'function' && nextOpen !== wasOpen) {
        setTimeout(() => {
          try { editor.layout(); } catch (_) {}
        }, 0);
      }
    } catch (_) {}

    if (!nextOpen) {
      disposeCodeGraph();
      disposeAIJobMonitor();
      // Aggressive Debug Insights cleanup on pane close
      try { cleanupDebugInsightsAggressive() } catch (_) {}
      // Revert Monaco theme when right pane closes
      applyDebugInsightsMonacoTheme(false);
      try { window.disableAllDebugPlugins?.() } catch (_) {}
      // Remove resizer listeners to avoid global mousemove/mouseup retention
      try { window.__RightPaneResizerTeardown?.() } catch (_) {}
      // Nudge Monaco to refresh layout (clears any lingering visuals)
      try { setTimeout(() => { const ed = window.monacoEditorManager?.getEditor?.() || window.editor?.getEditor?.() || window.monacoEditor; ed?.layout?.(); }, 0) } catch (_) {}
    }
    updateAgentsButtonHighlight();
  }

  function renderActivePane(options = {}) {
    if (!isOpen() || !currentPane) return;
    const container = document.getElementById('right-pane-content')
    const getRenderKey = () => {
      try { return container?.dataset?.paneRenderKey || null } catch (_) { return null }
    }
    const setRenderKey = (key) => {
      try {
        if (!container) return
        container.dataset.paneRenderKey = key || ''
      } catch (_) {}
    }
    if (currentPane === PANE_AI_CONNECTION) {
      // Leaving Debug Insights → cleanup
      try { cleanupDebugInsightsAggressive() } catch (_) {}
      applyDebugInsightsMonacoTheme(false);
      try { window.disableAllDebugPlugins?.() } catch (_) {}
      currentAgentsTab = 'connection';
      try { localStorage.setItem('right_pane_ai_tab', 'connection'); } catch (_) {}
      const expectedKey = 'ai-settings:connection:connection'
      if (getRenderKey() === expectedKey && container?.querySelector?.('.right-ai-tabs')) {
        return
      }
      renderAISettings('connection', { visibleTabs: ['connection'] });
      setRenderKey(expectedKey)
      return;
    }
    if (currentPane === PANE_AI_MANAGEMENT) {
      currentPane = PANE_AI_CONNECTION
      try { localStorage.setItem(CURRENT_KEY, currentPane); } catch (_) {}
      renderAISettings('connection', { visibleTabs: ['connection'] })
      setRenderKey('ai-settings:connection:connection')
      return;
    }
    if (currentPane === PANE_CHAT) {
      try { cleanupDebugInsightsAggressive() } catch (_) {}
      applyDebugInsightsMonacoTheme(false);
      try { window.disableAllDebugPlugins?.() } catch (_) {}
      safeRenderChat();
      setRenderKey(PANE_CHAT)
      return;
    }
    if (currentPane === PANE_AI_JOBS) {
      currentPane = PANE_AI_CONNECTION;
      try { localStorage.setItem(CURRENT_KEY, currentPane); } catch (_) {}
      renderActivePane();
      return;
    }
    if (currentPane === PANE_CODE_GRAPH) {
      try { cleanupDebugInsightsAggressive() } catch (_) {}
      applyDebugInsightsMonacoTheme(false);
      try { window.disableAllDebugPlugins?.() } catch (_) {}
      renderCodeGraphPane();
      setRenderKey(PANE_CODE_GRAPH)
    }
    if (currentPane === PANE_DEBUG) {
      try { performance.mark && performance.mark('debug:mount:start') } catch (_) {}
      applyDebugInsightsMonacoTheme(true);
      // Ensure DebugStore subscribed and integration initialized after aggressive cleanup
      try { window.DebugStore?.init?.() } catch (_) {}
      try {
        if (window.debugPluginsIntegration && typeof window.debugPluginsIntegration.isInitialized === 'function' && !window.debugPluginsIntegration.isInitialized()) {
          const p = window.debugPluginsIntegration.initialize?.();
          if (p && typeof p.then === 'function') {
            p.then(() => { try { window.applyDebugInsightsSavedSettings?.() } catch (_) {} renderDebugPane(); }).catch(() => renderDebugPane());
          } else {
            try { window.applyDebugInsightsSavedSettings?.() } catch (_) {}
            renderDebugPane();
          }
        } else {
          try { window.applyDebugInsightsSavedSettings?.() } catch (_) {}
          renderDebugPane();
        }
      } catch (_) { renderDebugPane(); }
      try {
        if (performance.mark && performance.measure) {
          performance.mark('debug:mount:end')
          performance.measure('debug:mount', 'debug:mount:start', 'debug:mount:end')
        }
      } catch (_) {}
      return;
    }
  }

  function toggleRightPane() {
    const willOpen = !isOpen();
    setOpen(willOpen);
    if (willOpen) {
      if (!currentPane) {
        try {
          const stored = localStorage.getItem(CURRENT_KEY) || null;
          currentPane = normalizePaneKey(stored) || PANE_AI_CONNECTION;
        } catch (_) {
          currentPane = PANE_AI_CONNECTION;
        }
        try { localStorage.setItem(CURRENT_KEY, currentPane); } catch (_) {}
      }
      renderActivePane();
    }
  }

  // Expose global for inline buttons
  window.toggleRightPane = toggleRightPane;

  function openDebugPane() {
    try { performance.mark && performance.mark('rightpane:open:start') } catch (_) {}
    currentPane = PANE_DEBUG;
    try { localStorage.setItem(CURRENT_KEY, currentPane); } catch(_) {}
    setOpen(true);
    applyDebugInsightsMonacoTheme(true);
    // Ensure store and integration are ready after aggressive cleanup
    try { window.DebugStore?.init?.() } catch (_) {}
    try {
      if (window.debugPluginsIntegration && typeof window.debugPluginsIntegration.isInitialized === 'function' && !window.debugPluginsIntegration.isInitialized()) {
        const p = window.debugPluginsIntegration.initialize?.();
        if (p && typeof p.then === 'function') {
          p.then(() => { try { window.applyDebugInsightsSavedSettings?.() } catch (_) {} renderDebugPane(); })
           .catch(() => renderDebugPane());
        } else {
          try { window.applyDebugInsightsSavedSettings?.() } catch (_) {}
          renderDebugPane();
        }
      } else {
        try { window.applyDebugInsightsSavedSettings?.() } catch (_) {}
        renderDebugPane();
      }
    } catch (_) { renderDebugPane(); }
    const btn = document.getElementById('vscode-activity-bar-right')?.querySelector(`[data-button="${PANE_DEBUG}"]`);
    if (btn) btn.classList.add('active');
    try {
      if (performance.mark && performance.measure) {
        performance.mark('rightpane:open:end')
        performance.measure('rightpane:open', 'rightpane:open:start', 'rightpane:open:end')
      }
    } catch (_) {}
  }

  function openChatPane() {
    try { performance.mark && performance.mark('rightpane:open:start') } catch (_) {}
    currentPane = PANE_CHAT;
    try { localStorage.setItem(CURRENT_KEY, currentPane); } catch(_) {}
    disposeCodeGraph();
    setOpen(true);
    safeRenderChat();
    const btnChat = document.getElementById('vscode-activity-bar-right')?.querySelector(`[data-button="${PANE_CHAT}"]`);
    if (btnChat) btnChat.classList.add('active');
    try {
      if (performance.mark && performance.measure) {
        performance.mark('rightpane:open:end')
        performance.measure('rightpane:open', 'rightpane:open:start', 'rightpane:open:end')
      }
    } catch (_) {}
  }

  function openAIConnectionPane() {
    try { performance.mark && performance.mark('rightpane:open:start') } catch (_) {}
    const container = document.getElementById('right-pane-content')
    if (isOpen() && currentPane === PANE_AI_CONNECTION && currentAgentsTab === 'connection' && container?.dataset?.paneRenderKey === 'ai-settings:connection:connection') {
      try {
        if (performance.mark && performance.measure) {
          performance.mark('rightpane:open:end')
          performance.measure('rightpane:open', 'rightpane:open:start', 'rightpane:open:end')
        }
      } catch (_) {}
      return
    }
    currentPane = PANE_AI_CONNECTION;
    currentAgentsTab = 'connection';
    try {
      localStorage.setItem(CURRENT_KEY, currentPane);
      localStorage.setItem('right_pane_ai_tab', 'connection');
    } catch (_) {}
    disposeCodeGraph();
    setOpen(true);
    renderActivePane({ defaultTab: 'connection' });
    try {
      if (performance.mark && performance.measure) {
        performance.mark('rightpane:open:end')
        performance.measure('rightpane:open', 'rightpane:open:start', 'rightpane:open:end')
      }
    } catch (_) {}
  }

  function openCodeGraphPane() {
    try { performance.mark && performance.mark('rightpane:open:start') } catch (_) {}
    currentPane = PANE_CODE_GRAPH;
    try { localStorage.setItem(CURRENT_KEY, currentPane); } catch(_) {}
    setOpen(true);
    // Leaving Debug Insights if previously mounted
    try { cleanupDebugInsightsAggressive() } catch (_) {}
    renderCodeGraphPane();
    const btnCodeGraph = document.getElementById('vscode-activity-bar-right')?.querySelector(`[data-button="${PANE_CODE_GRAPH}"]`);
    if (btnCodeGraph) btnCodeGraph.classList.add('active');
    try {
      if (performance.mark && performance.measure) {
        performance.mark('rightpane:open:end')
        performance.measure('rightpane:open', 'rightpane:open:start', 'rightpane:open:end')
      }
    } catch (_) {}
  }

  function cleanupDebugInsightsAggressive () {
    try { window.debugPluginsIntegration?.dispose?.() } catch (_) {}
    try { window.DebugStore?.dispose?.() } catch (_) {}
    try {
      const ids = ['debug-insights-styles','perf-di-css','cov-di-css','linter-di-css']
      ids.forEach((id) => { const el = document.getElementById(id); if (el) el.remove(); })
    } catch (_) {}
  }

  function disposeCodeGraph() {
    if (codeGraphViewMounted) {
      try {
        const view = window.ModuleViews?.['code-graph-pane'];
        if (view && typeof view.unmount === 'function') {
          view.unmount();
        }
      } catch (error) {
        console.warn('[RightActivityBar] disposeCodeGraph failed', error);
      }
    }
    if (codeGraphHost) {
      codeGraphHost.innerHTML = '';
      codeGraphHost = null;
    }
    codeGraphViewMounted = false;
    codeGraphContext = null;
  }

  function disposeAIJobMonitor () {
    if (aiJobMonitorMounted && window.AIJobMonitor && typeof window.AIJobMonitor.dispose === 'function') {
      try { window.AIJobMonitor.dispose(); } catch (_) {}
    }
    if (aiJobMonitorHost) {
      try { aiJobMonitorHost.innerHTML = ''; } catch (_) {}
      aiJobMonitorHost = null;
    }
    aiJobMonitorMounted = false;
  }

  async function renderCodeGraphPane() {
    const container = document.getElementById('right-pane-content');
    if (!container) return;

    disposeCodeGraph();

    const host = document.createElement('div');
    host.className = 'code-graph-right-host';
    host.style.display = 'flex';
    host.style.flexDirection = 'column';
    host.style.height = '100%';
    host.style.minHeight = '0';
    host.style.width = '100%';
    container.innerHTML = '';
    container.appendChild(host);
    codeGraphHost = host;

    try {
      const registry = window.ModuleRegistry;
      if (registry && typeof registry.ensureDescriptorLoaded === 'function') {
        await registry.ensureDescriptorLoaded('code-graph');
      }
      if (registry && typeof registry.loadModule === 'function') {
        await registry.loadModule('code-graph');
      }
    } catch (error) {
      console.error('[RightActivityBar] Failed to load Code Graph module', error);
      host.innerHTML = '<div style="padding:12px; line-height:1.4;">Unable to load Code Graph module.</div>';
      return;
    }

    const view = window.ModuleViews?.['code-graph-pane'];
    if (!view || typeof view.mount !== 'function') {
      host.innerHTML = '<div style="padding:12px; line-height:1.4;">Code Graph view is unavailable.</div>';
      return;
    }

    try {
      codeGraphContext = buildModuleContext();
      view.mount(host, codeGraphContext);
      codeGraphViewMounted = true;
    } catch (error) {
      console.error('[RightActivityBar] Code Graph mount failed', error);
      host.innerHTML = '<div style="padding:12px; line-height:1.4;">Failed to render Code Graph view.</div>';
    }
  }

  function setRightPaneTitle (text) {
    try {
      const pane = document.getElementById('vscode-right-pane')
      const title = pane && pane.querySelector('.right-pane-title .title-text')
      if (title) title.textContent = text || 'AI Utilities'
    } catch (_) {}
  }

  function renderDebugPane() {
    const container = document.getElementById('right-pane-content');
    if (!container) return;
    container.innerHTML = '';
    const host = document.createElement('div');
    host.className = 'debug-insights-right-host';
    host.style.display = 'block';
    host.style.height = '100%';
    host.style.overflow = 'auto';
    container.appendChild(host);

    // Respect explicit "do not load debug plugins" sessions (safe-mode)
    try {
      const urlParams = new URLSearchParams(window.location?.search || '')
      if (urlParams.get('debug-plugins') === 'false') {
        host.innerHTML = '<div style="padding:12px; line-height:1.4;">Debug Insights is disabled by URL parameter <code>?debug-plugins=false</code>.</div>'
        return
      }
    } catch (_) {}

    // Ensure the Debug Insights descriptor (and thus loader) is present
    let loaderReady = null
    try {
      const registry = window.ModuleRegistry;
      if (registry && typeof registry.ensureDescriptorLoaded === 'function') {
        loaderReady = registry.ensureDescriptorLoaded('debug-insights')
          .then(() => (typeof registry.loadModule === 'function' ? registry.loadModule('debug-insights') : true))
          .catch((e) => console.warn('[RightActivityBar] debug-insights descriptor load failed', e))
      }
    } catch (_) {}

    try {
      // Reuse the existing integration to mount insights tabs UI
      const attemptMount = () => {
        if (!window.debugPluginsIntegration || typeof window.debugPluginsIntegration.mountInsightsInto !== 'function') return false
        try {
          window.debugPluginsIntegration.mountInsightsInto(host)
          return true
        } catch (e) {
          console.error('[RightActivityBar] Debug Insights mount failed', e)
          return false
        }
      }

      if (attemptMount()) {
        try {
          if (pendingDebugPluginsLoadedListener) {
            window.removeEventListener('debug-plugins-loaded', pendingDebugPluginsLoadedListener)
            window.removeEventListener('debug-plugins-integration-ready', pendingDebugPluginsLoadedListener)
          }
        } catch (_) {}
        pendingDebugPluginsLoadedListener = null
      } else {
        // Late init path: wait for integration init, then retry mount
        host.innerHTML = '<div style="padding:12px; line-height:1.4;">Loading Debug Insights…</div>'

        // Trigger loader explicitly so we can surface failures in the UI (loader is idempotent)
        const startLoad = () => {
          try {
            const loader = window.DebugPluginsLoader
            if (loader && typeof loader.load === 'function') {
              const p = loader.load()
              if (p && typeof p.catch === 'function') {
                p.catch((e) => {
                  console.error('[RightActivityBar] Debug plugins load failed', e)
                  host.innerHTML = '<div style="padding:12px; line-height:1.4;">Unable to load Debug Insights.</div>'
                })
              }
            }
          } catch (e) {
            console.error('[RightActivityBar] Debug plugins load failed', e)
            host.innerHTML = '<div style="padding:12px; line-height:1.4;">Unable to load Debug Insights.</div>'
          }
        }

        if (loaderReady && typeof loaderReady.then === 'function') {
          loaderReady.then(startLoad).catch(startLoad)
        } else {
          startLoad()
        }

        const retry = () => {
          try { window.applyDebugInsightsSavedSettings?.() } catch (_) {}
          pendingDebugPluginsLoadedListener = null
          if (attemptMount()) return

          const loadError = window.DebugPluginsLoader?.getLoadError?.() || window.__DEBUG_PLUGINS_LOAD_ERROR__ || null
          if (loadError) {
            host.innerHTML = '<div style="padding:12px; line-height:1.4;">Unable to load Debug Insights.</div>'
          }
        };
        try {
          if (pendingDebugPluginsLoadedListener) {
            window.removeEventListener('debug-plugins-loaded', pendingDebugPluginsLoadedListener)
            window.removeEventListener('debug-plugins-integration-ready', pendingDebugPluginsLoadedListener)
          }
        } catch (_) {}
        pendingDebugPluginsLoadedListener = retry
        try { window.addEventListener('debug-plugins-integration-ready', retry, { once: true }); } catch (_) {}
      }
    } catch (error) {
      console.error('[RightActivityBar] Debug Insights mount failed', error);
      host.innerHTML = '<div style="padding:12px; line-height:1.4;">Failed to render Debug Insights.</div>';
    }
  }



  const CONNECTION_STATUS_OPTIONS = [
    { value: 'unknown', label: 'Unknown', className: 'status-unknown' },
    { value: 'ok', label: 'Ready', className: 'status-ok' },
    { value: 'warning', label: 'Warning', className: 'status-warning' },
    { value: 'error', label: 'Error', className: 'status-error' }
  ];

  const CONNECTION_SECRET_PREFIX = 'onlyoffice_macros_ai_connection_secret_';
  const AI_CONNECTION_SPLIT_KEY = 'right_pane_ai_connection_split';

  function mountConnectionManager({ elements, settingsStore }) {
    const { treeEl, inspectorEl, filterEl, searchEl, addButton, contentEl, splitterEl } = elements || {};
    if (!treeEl || !inspectorEl || !settingsStore) {
      return function noop() {};
    }

    const providerOptions = [{ id: AI_FREE_PROVIDER_ID, label: 'OpenAI', enabled: true }];
    const providerFactory = window.AIProviderFactory || null;
    const providerLabelCache = new Map();
    const providerModelsCache = new Map();
    const providerModelRequests = new Map();
    const connectionSearchCache = new Map();
    const disposers = [];
    let inspectorDisposers = [];
    let contextMenuEl = null;
    let suppressNextContextMenuDismiss = false;
    let lastTreeKey = '';

    function enforceFreeTierConnection(connection) {
      if (!connection || !AI_FREE_RESTRICTIONS_ENABLED) return connection;
      connection.providerId = AI_FREE_PROVIDER_ID;
      connection.models = [{ id: AI_FREE_LOCKED_MODEL_ID }];
      connection.primaryModelId = AI_FREE_LOCKED_MODEL_ID;
      return connection;
    }

    function getDefaultConnectionModel() {
      try {
        const fromConfig = window.AIConfiguration?.getDefaultModel?.(AI_FREE_PROVIDER_ID);
        if (typeof fromConfig === 'string' && fromConfig.trim()) {
          return fromConfig.trim();
        }
      } catch (_) {}
      return AI_FREE_LOCKED_MODEL_ID;
    }

    const state = {
      connections: [],
      filter: 'all',
      search: '',
      selectedId: null,
      editing: null,
      suppressStore: false,
      feedbackTone: 'info',
      feedbackMessage: '',
      modelFocusIndex: null,
      pendingModelFocusIndex: null,
      connectionsSignature: '',
      filterOptionsSignature: '',
      splitRatio: 0.2,
      previousSplitRatio: 0.2
    };

    function clampSplitRatio(value) {
      const numeric = Number(value);
      if (!Number.isFinite(numeric)) return 0.2;
      if (numeric <= 0.02) return 0;
      if (numeric >= 0.98) return 1;
      return Math.max(0.08, Math.min(0.92, numeric));
    }

    function applyConnectionSplit(nextRatio, options = {}) {
      if (!contentEl) return;
      const { persist = true } = options;
      const ratio = clampSplitRatio(nextRatio);
      state.splitRatio = ratio;
      if (ratio > 0 && ratio < 1) {
        state.previousSplitRatio = ratio;
      }
      contentEl.style.setProperty('--ai-connection-tree-width', `${Math.round(ratio * 1000) / 10}%`);
      contentEl.classList.toggle('collapsed-tree', ratio === 0);
      contentEl.classList.toggle('collapsed-inspector', ratio === 1);
      if (splitterEl) {
        splitterEl.setAttribute('aria-valuemin', '0');
        splitterEl.setAttribute('aria-valuemax', '100');
        splitterEl.setAttribute('aria-valuenow', String(Math.round(ratio * 100)));
      }
      if (persist) {
        try { localStorage.setItem(AI_CONNECTION_SPLIT_KEY, String(ratio)); } catch (_) {}
      }
    }

    function initConnectionSplitControls() {
      if (!contentEl || !splitterEl) return;
      let dragging = false;

      const handleDragMove = (event) => {
        if (!dragging) return;
        const rect = contentEl.getBoundingClientRect();
        const width = rect.width || 1;
        const offset = (event.clientX - rect.left) / width;
        applyConnectionSplit(offset, { persist: false });
      };

      const handleDragEnd = () => {
        if (!dragging) return;
        dragging = false;
        contentEl.classList.remove('is-resizing');
        applyConnectionSplit(state.splitRatio, { persist: true });
      };

      addListener(splitterEl, 'mousedown', (event) => {
        event.preventDefault();
        dragging = true;
        contentEl.classList.add('is-resizing');
      });
      addListener(document, 'mousemove', handleDragMove);
      addListener(document, 'mouseup', handleDragEnd);

      addListener(splitterEl, 'keydown', (event) => {
        if (event.key === 'ArrowLeft') {
          event.preventDefault();
          applyConnectionSplit(state.splitRatio - 0.05);
        } else if (event.key === 'ArrowRight') {
          event.preventDefault();
          applyConnectionSplit(state.splitRatio + 0.05);
        } else if (event.key === 'Home') {
          event.preventDefault();
          applyConnectionSplit(0);
        } else if (event.key === 'End') {
          event.preventDefault();
          applyConnectionSplit(1);
        }
      });

      addListener(splitterEl, 'dblclick', (event) => {
        event.preventDefault();
        if (state.splitRatio === 1) {
          applyConnectionSplit(state.previousSplitRatio || 0.2);
          return;
        }
        applyConnectionSplit(1);
      });

      let storedRatio = 0.2;
      try {
        const raw = localStorage.getItem(AI_CONNECTION_SPLIT_KEY);
        if (raw !== null) storedRatio = Number(raw);
      } catch (_) {}
      applyConnectionSplit(storedRatio, { persist: false });
    }

    function addListener(target, type, handler, options) {
      if (!target || typeof target.addEventListener !== 'function') return;
      target.addEventListener(type, handler, options || false);
      disposers.push(() => {
        try { target.removeEventListener(type, handler, options || false); } catch (_) {}
      });
    }

    function clearInspectorListeners() {
      inspectorDisposers.forEach((fn) => {
        try { fn(); } catch (_) {}
      });
      inspectorDisposers = [];
    }

    function addInspectorListener(target, type, handler, options) {
      if (!target || typeof target.addEventListener !== 'function') return;
      target.addEventListener(type, handler, options || false);
      inspectorDisposers.push(() => {
        try { target.removeEventListener(type, handler, options || false); } catch (_) {}
      });
    }

    function escapeHtml(value) {
      return String(value ?? '').replace(/[&<>]/g, (char) => ({
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;'
      }[char]));
    }

    function escapeAttr(value) {
      return escapeHtml(value).replace(/"/g, '&quot;');
    }

    function isPlainObject(value) {
      return value !== null && typeof value === 'object' && !Array.isArray(value);
    }

    function cloneConnection(connection) {
      if (!isPlainObject(connection)) return null;
      const base = { ...connection };
      base.models = Array.isArray(connection.models)
        ? connection.models.map((model) => isPlainObject(model) ? { ...model } : { id: String(model || '') })
        : [];
      base.tags = Array.isArray(connection.tags) ? [...connection.tags] : [];
      return base;
    }

    function classifyModel(providerId, modelId) {
      const id = typeof modelId === 'string' ? modelId.trim().toLowerCase() : '';
      if (!id) return { isText: false, isCoding: false };
      const provider = normalizeProviderId(providerId);
      if (provider === 'openai') {
        const nonTextTokens = ['embedding', 'image', 'vision', 'audio', 'sora', 'tts', 'speech', 'voice', 'transcribe', 'moderation', 'search'];
        if (nonTextTokens.some((token) => id.includes(token))) {
          return { isText: false, isCoding: false };
        }
        const codingTokens = ['code', 'codex', 'gpt-4o', 'gpt-4.1', 'gpt-4-', 'gpt-4', 'gpt-3.5', 'o1', 'o3'];
        const isCoding = codingTokens.some((token) => id.includes(token));
        return { isText: true, isCoding };
      }
      return { isText: true, isCoding: false };
    }

    function normalizeModelList(providerId, list) {
      if (!Array.isArray(list)) return [];
      return list.map((item) => {
        let id = null;
        let name = null;
        let maxTokens = undefined;
        let ownedBy = undefined;
        if (typeof item === 'string') {
          id = item;
          name = item;
        } else if (isPlainObject(item) && typeof item.id === 'string') {
          id = item.id;
          name = typeof item.name === 'string' && item.name.trim() ? item.name : item.id;
          maxTokens = item.maxTokens ?? item.max_tokens ?? undefined;
          ownedBy = item.ownedBy ?? item.owned_by ?? undefined;
          if (typeof item.isText === 'boolean' || typeof item.isCoding === 'boolean') {
            return {
              id,
              name,
              maxTokens,
              ownedBy,
              isText: item.isText !== false,
              isCoding: item.isCoding === true
            };
          }
        }
        if (!id) return null;
        const tags = classifyModel(providerId, id);
        return {
          id,
          name,
          maxTokens,
          ownedBy,
          isText: tags.isText,
          isCoding: tags.isCoding
        };
      }).filter(Boolean);
    }

    function persistConnections(nextConnections, { selectId } = {}) {
      const sanitized = Array.isArray(nextConnections)
        ? nextConnections.map((connection) => cloneConnection(connection)).filter(Boolean)
        : [];
      if (settingsStore && typeof settingsStore.update === 'function') {
        try {
          state.suppressStore = true;
          settingsStore.update({ aiConnections: sanitized });
          setTimeout(() => {
            state.suppressStore = false;
          }, 0);
        } catch (error) {
          state.suppressStore = false;
          console.error('[RightActivityBar] Failed to persist connections:', error);
          return false;
        }
      } else {
        console.warn('[RightActivityBar] SettingsStore unavailable; applying changes locally only.');
      }
      state.connections = sanitized;
      // Cache snapshot to survive incidental remounts/reloads
      try { window.__aiConnectionsCache = JSON.parse(JSON.stringify(state.connections)) } catch (_) {}
      syncConnectionCaches();
      if (typeof selectId === 'string' || selectId === null) {
        state.selectedId = selectId;
      } else if (state.selectedId && !sanitized.some((connection) => connection.id === state.selectedId)) {
        state.selectedId = sanitized[0]?.id || null;
      }
      state.editing = state.selectedId
        ? cloneConnection(sanitized.find((connection) => connection.id === state.selectedId))
        : null;
      populateFilterOptions();
      renderTree();
      renderInspector();
      try {
        document.dispatchEvent(new CustomEvent('aiConnections:updated', { detail: { connections: sanitized } }));
      } catch (_) {}
      return true;
    }

    function validateConnectionSecret(providerId, secret) {
      const normalized = normalizeProviderId(providerId);
      if (!secret || !secret.trim()) {
        if (normalized === 'ollama') {
          // Ollama typically does not require an API key
          return { ok: true, value: '' };
        }
        return { ok: false, message: 'Add an API key before testing the connection.' };
      }
      const value = secret.trim();
      if (normalized === 'openai' && !value.startsWith('sk-')) {
        return { ok: false, message: 'OpenAI keys should start with "sk-".' };
      }
      if (normalized === 'github') {
        const hasKnownPrefix = value.startsWith('ghp_') ||
          value.startsWith('gho_') ||
          value.startsWith('ghu_') ||
          value.startsWith('ghs_') ||
          value.startsWith('github_pat_');
        if (!hasKnownPrefix) {
          return { ok: false, message: 'GitHub tokens should start with gh* (e.g. “ghp_”, “github_pat_”).' };
        }
      }
      if (normalized === 'deepseek' ||
        normalized === 'litellm' ||
        normalized === 'minimax' ||
        normalized === 'qwen-code' ||
        normalized === 'zai' ||
        normalized === 'gemini') {
        // For these providers we only require a non-empty key; format
        // validation is delegated to the server response.
        if (!value) {
          return { ok: false, message: 'Add an API key before testing the connection.' };
        }
      }
      return { ok: true, value };
    }

    function ensureCopyName(baseName) {
      const existing = new Set(state.connections.map((connection) => (connection.name || '').toLowerCase()));
      const base = (baseName || 'Connection').trim() || 'Connection';
      let candidate = `${base} Copy`;
      let counter = 2;
      while (existing.has(candidate.toLowerCase())) {
        candidate = `${base} Copy ${counter}`;
        counter += 1;
      }
      return candidate;
    }

    function promptRename(connection) {
      if (!connection) return;
      const currentName = connection.name || '';
      const input = window.prompt('Rename connection', currentName);
      if (!input) return;
      const trimmed = input.trim();
      if (!trimmed || trimmed === currentName) return;
      const next = state.connections.map((entry) => {
        const clone = cloneConnection(entry);
        if (entry.id === connection.id) {
          clone.name = trimmed;
          clone.updatedAt = Date.now();
        }
        return clone;
      });
      if (!persistConnections(next, { selectId: connection.id })) return;
      if (state.editing && state.editing.id === connection.id) {
        setFeedback('Connection renamed.', 'success');
      } else {
        window.LogService?.logInfo?.('AI Connections', 'rename', { id: connection.id, name: trimmed });
      }
    }

    function duplicateConnection(connection) {
      if (!connection) return;
      const duplicate = cloneConnection(connection);
      const now = Date.now();
      duplicate.id = `conn-${Math.random().toString(36).slice(2, 9)}${now.toString(36)}`;
      duplicate.name = ensureCopyName(connection.name);
      duplicate.createdAt = now;
      duplicate.updatedAt = now;
      duplicate.lastTestedAt = null;
      duplicate.status = 'unknown';
      const next = state.connections.map((entry) => cloneConnection(entry));
      next.push(duplicate);
      if (!persistConnections(next, { selectId: duplicate.id })) return;
      const sourceSecret = loadSecret(connection.id);
      if (sourceSecret) {
        saveSecret(duplicate.id, duplicate.providerId, sourceSecret);
        scheduleProviderModelFetch(duplicate.providerId, {
          apiKey: sourceSecret,
          onComplete: () => {
            if (state.selectedId === duplicate.id) {
              renderTree();
              renderInspector();
            }
          }
        });
      }
      setTimeout(() => setFeedback('Connection copied.', 'success'), 0);
    }

    function testConnectionFromMenu(connection) {
      if (!connection) return;
      const secret = loadSecret(connection.id);
      const validation = validateConnectionSecret(connection.providerId, secret);
      if (!validation.ok) {
        if (state.editing && state.editing.id === connection.id) {
          setFeedback(validation.message, 'warning');
        } else {
          window.alert(validation.message);
        }
        return;
      }
      const now = Date.now();
      const next = state.connections.map((entry) => {
        const clone = cloneConnection(entry);
        if (entry.id === connection.id) {
          clone.status = 'ok';
          clone.lastTestedAt = now;
          clone.updatedAt = now;
        }
        return clone;
      });
      if (!persistConnections(next, { selectId: connection.id })) return;
      if (state.editing && state.editing.id === connection.id) {
        setFeedback('Connection passed basic validation.', 'success');
      } else {
        window.LogService?.logInfo?.('AI Connections', 'test', { id: connection.id, provider: connection.providerId });
      }
    }

    function deleteConnectionFromMenu(connection) {
      if (!connection) return;
      const confirmed = window.confirm(`Delete connection "${connection.name || connection.id}"?`);
      if (!confirmed) return;
      const index = state.connections.findIndex((entry) => entry.id === connection.id);
      const next = state.connections
        .filter((entry) => entry.id !== connection.id)
        .map((entry) => cloneConnection(entry));
      let nextSelection = state.selectedId;
      if (nextSelection === connection.id) {
        if (next.length) {
          const fallbackIndex = Math.max(0, index - 1);
          nextSelection = next[fallbackIndex]?.id || next[0]?.id || null;
        } else {
          nextSelection = null;
        }
      }
      removeSecret(connection.id);
      if (!persistConnections(next, { selectId: nextSelection })) return;
      window.LogService?.logInfo?.('AI Connections', 'delete', { id: connection.id });
    }

    function getProviderLabel(providerId) {
      if (providerLabelCache.has(providerId)) return providerLabelCache.get(providerId);
      if (!providerId) return 'Unknown';
      const catalogMatch = providerOptions.find((item) => item.id === providerId) || null;
      const label = catalogMatch?.label
        || providerId.replace(/[-_]+/g, ' ').replace(/\b\w/g, (c) => c.toUpperCase());
      providerLabelCache.set(providerId, label);
      return label;
    }

    /**
     * Debug-only helper to test the shared agent core inside AI Sputnik.
     * Uses AgentEngine.runSimpleAgentTask with a fixed goal.
     */
    function runAgentCoreSmokeTest () {
      if (!window.ENABLE_AGENT_CORE || !window.AgentEngine || typeof window.AgentEngine.runSimpleAgentTask !== 'function') {
        window.debug?.warn('RightActivityBar', 'Agent core smoke test not available', {
          ENABLE_AGENT_CORE: window.ENABLE_AGENT_CORE,
          hasAgentEngine: !!window.AgentEngine
        });
        return;
      }
      try {
        const goal = 'Explain in one sentence what this OnlyOffice macros IDE does for the user.'
        window.AgentEngine.runSimpleAgentTask({ goal }).then(function (result) {
          const msg = result && result.finalMessage && result.finalMessage.content
            ? result.finalMessage.content
            : '[no content]';
          window.debug?.info('RightActivityBar', 'Agent core smoke test result', { goal, msg });
        }).catch(function (error) {
          window.debug?.error('RightActivityBar', 'Agent core smoke test failed', error);
        });
      } catch (e) {
        window.debug?.error('RightActivityBar', 'Agent core smoke test threw', e);
      }
    }

    function normalizeProviderId(providerId) {
      if (providerFactory?.normalizeProviderId) {
        try { return providerFactory.normalizeProviderId(providerId); } catch (_) {}
      }
      return typeof providerId === 'string' && providerId.trim()
        ? providerId.trim().toLowerCase()
        : 'openai';
    }

    function getCachedProviderModels(providerId) {
      const normalized = normalizeProviderId(providerId);
      let entry = providerModelsCache.get(normalized);

      // Ensure GitHub Models always exposes at least the predefined
      // fallback catalog, even if a remote fetch has not yet run or
      // failed. This uses the provider's own static catalog so the
      // list stays in sync with ADR-057.
      if (normalized === 'github' &&
        (!entry || !Array.isArray(entry.models) || entry.models.length === 0)) {
        try {
          if (window.GitHubAIProvider) {
            const provider = new window.GitHubAIProvider();
            if (provider && typeof provider.getAvailableModels === 'function') {
              const models = provider.getAvailableModels() || [];
              const normalizedList = normalizeModelList('github', models);
              if (normalizedList.length) {
                entry = {
                  models: normalizedList,
                  source: 'static',
                  fetchedAt: Date.now()
                };
                providerModelsCache.set(normalized, entry);
              }
            }
          }
        } catch (_) {}
      }

      return Array.isArray(entry?.models) ? entry.models : [];
    }

    function setCachedProviderModels(providerId, models, source = 'static', metadata = {}) {
      const normalized = normalizeProviderId(providerId);
      const normalizedList = normalizeModelList(providerId, models);
      providerModelsCache.set(normalized, {
        models: normalizedList,
        source,
        fetchedAt: Date.now(),
        ...metadata
      });
      return providerModelsCache.get(normalized).models;
    }

    function loadProviderModels(providerId) {
      const normalized = normalizeProviderId(providerId);
      const cached = providerModelsCache.get(normalized);
      if (cached && Array.isArray(cached.models) && cached.models.length) {
        return cached.models;
      }
      let models = [];
      try {
        let instance = null;
        if (providerFactory?.createProvider) {
          try {
            instance = providerFactory.createProvider(normalized);
          } catch (error) {
            console.warn('[RightActivityBar] Failed to create provider for model catalog:', error);
          }
        }
        if (!instance) {
          if (normalized === 'openai' && window.OpenAIProvider) {
            instance = new window.OpenAIProvider();
          }
        }
        if (instance?.getAvailableModels) {
          const available = instance.getAvailableModels();
          if (Array.isArray(available)) {
            models = available;
          }
        }
      } catch (error) {
        console.warn('[RightActivityBar] Failed to load model list for provider:', providerId, error);
      }
      return setCachedProviderModels(providerId, models, 'static');
    }

    async function fetchOpenAIModels(apiKey) {
      if (!apiKey) return [];
      const results = [];
      const base = 'https://api.openai.com/v1/models';
      const headers = {
        Authorization: `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      };
      let url = base;
      const visited = new Set();
      while (url && !visited.has(url)) {
        visited.add(url);
        const response = await fetch(url, { method: 'GET', headers });
        if (!response.ok) {
          const text = await response.text().catch(() => '');
          throw new Error(`HTTP ${response.status} fetching models: ${text || response.statusText}`);
        }
        const payload = await response.json();
        const data = Array.isArray(payload?.data) ? payload.data : [];
        data.forEach((item) => {
          if (item && typeof item.id === 'string') {
            results.push({
              id: item.id,
              name: item.id,
              ownedBy: typeof item.owned_by === 'string' ? item.owned_by : undefined
            });
          }
        });
        if (payload?.has_more && payload?.last_id) {
          url = `${base}?after=${encodeURIComponent(payload.last_id)}`;
        } else if (payload?.next) {
          url = payload.next.startsWith('http')
            ? payload.next
            : `${base}?${payload.next.replace(/^\?/, '')}`;
        } else {
          break;
        }
      }
      return results;
    }

    function scheduleProviderModelFetch(providerId, options = {}) {
      const normalized = normalizeProviderId(providerId);
      if (normalized !== 'openai' &&
        normalized !== 'github' &&
        normalized !== 'gemini' &&
        normalized !== 'deepseek' &&
        normalized !== 'ollama' &&
        normalized !== 'litellm' &&
        normalized !== 'zai') return;
      const apiKey = typeof options.apiKey === 'string' && options.apiKey.trim() ? options.apiKey.trim() : '';
      if (normalized !== 'ollama' && normalized !== 'zai' && !apiKey) return;
      if (normalized !== 'zai' && typeof fetch !== 'function') return;
      const requestKey = normalized === 'ollama'
        ? `${normalized}:local`
        : `${normalized}:${apiKey.slice(0, 16)}`;
      const cacheEntry = providerModelsCache.get(normalized);
      const force = options.force === true;
      if (!force && cacheEntry) {
        if (cacheEntry.apiKeyFingerprint === requestKey && (cacheEntry.source === 'remote' || cacheEntry.source === 'error')) {
          return;
        }
        if (cacheEntry.source === 'remote' && Array.isArray(cacheEntry.models) && cacheEntry.models.length) {
          return;
        }
      }
      if (providerModelRequests.has(requestKey)) return;
      providerModelRequests.set(requestKey, true);

      let fetchFn = null;
      let fetchArgs = [apiKey];
      if (normalized === 'openai') {
        fetchFn = fetchOpenAIModels;
      } else if (normalized === 'github') {
        fetchFn = window.GitHubApi && typeof window.GitHubApi.fetchModelsCatalog === 'function'
          ? window.GitHubApi.fetchModelsCatalog
          : null;
      } else if (normalized === 'gemini') {
        fetchFn = window.GeminiApi && typeof window.GeminiApi.fetchModelsCatalog === 'function'
          ? window.GeminiApi.fetchModelsCatalog
          : null;
      } else if (normalized === 'deepseek') {
        fetchFn = window.DeepSeekApi && typeof window.DeepSeekApi.fetchModelsCatalog === 'function'
          ? window.DeepSeekApi.fetchModelsCatalog
          : null;
      } else if (normalized === 'ollama') {
        fetchFn = window.OllamaApi && typeof window.OllamaApi.fetchModelsCatalog === 'function'
          ? window.OllamaApi.fetchModelsCatalog
          : null;
        fetchArgs = [];
      } else if (normalized === 'litellm') {
        fetchFn = window.LiteLLMApi && typeof window.LiteLLMApi.fetchModelsCatalog === 'function'
          ? window.LiteLLMApi.fetchModelsCatalog
          : null;
        let baseUrl = options.baseUrl || options.baseURL || '';
        try {
          if ((!baseUrl || !baseUrl.trim()) &&
            window.AIConfiguration &&
            typeof window.AIConfiguration.get === 'function') {
            baseUrl = window.AIConfiguration.get('litellmBaseUrl');
          }
        } catch (_) {}
        fetchArgs = [apiKey, baseUrl];
      } else if (normalized === 'zai') {
        fetchFn = window.ZAiApi && typeof window.ZAiApi.getStaticModels === 'function'
          ? window.ZAiApi.getStaticModels
          : null;
        let apiLine = options.apiLine || options.zaiApiLine || '';
        try {
          if (!apiLine && window.AIConfiguration && typeof window.AIConfiguration.get === 'function') {
            apiLine = window.AIConfiguration.get('zaiApiLine');
          }
        } catch (_) {}
        fetchArgs = [apiLine];
      }

      if (!fetchFn) {
        providerModelRequests.delete(requestKey);
        return;
      }

      Promise.resolve(fetchFn.apply(null, fetchArgs))
        .then((models) => {
          const merged = Array.isArray(models) && models.length
            ? models
            : (Array.isArray(cacheEntry?.models) ? cacheEntry.models : []);
          const source = Array.isArray(models) && models.length
            ? 'remote'
            : (cacheEntry?.source || 'static');
          setCachedProviderModels(providerId, merged, source, {
            apiKeyFingerprint: requestKey,
            lastFetchError: undefined
          });
        })
        .catch((error) => {
          console.warn('[RightActivityBar] Failed to fetch models via REST:', error);
          const raw = error?.message || String(error);
          let friendly = raw;
          const lower = raw.toLowerCase ? raw.toLowerCase() : String(raw).toLowerCase();
          if (lower.includes('401') || lower.includes('403') || lower.includes('unauthorized') || lower.includes('invalid api key')) {
            const label = getProviderLabel(providerId);
            friendly = `Invalid or missing API key for ${label}.`;
          } else if (lower.includes('failed to fetch') ||
            lower.includes('networkerror') ||
            lower.includes('connection refused') ||
            lower.includes('ecconnrefused')) {
            if (normalized === 'ollama') {
              friendly = 'Ollama server not reachable. Ensure it is running on http://127.0.0.1:11434.';
            } else {
              friendly = 'Server not reachable (is it running and accessible?).';
            }
          }
          setCachedProviderModels(providerId, Array.isArray(cacheEntry?.models) ? cacheEntry.models : [], 'error', {
            apiKeyFingerprint: requestKey,
            lastFetchError: friendly
          });
        })
        .finally(() => {
          providerModelRequests.delete(requestKey);
          if (typeof options.onComplete === 'function') {
            try { options.onComplete(); } catch (_) {}
          }
        });
    }

    function loadSecret(connectionId) {
      if (!connectionId) return '';
      try {
        return localStorage.getItem(`${CONNECTION_SECRET_PREFIX}${connectionId}`) || '';
      } catch (error) {
        console.warn('[RightActivityBar] Failed to read connection secret:', error);
        return '';
      }
    }

    function saveSecret(connectionId, providerId, value) {
      if (!connectionId) return false;
      try {
        const key = `${CONNECTION_SECRET_PREFIX}${connectionId}`;
        if (value && value.trim()) {
          localStorage.setItem(key, value.trim());
          try { window.AIStorage?.setApiKey?.(providerId, value.trim()); } catch (_) {}
        } else {
          localStorage.removeItem(key);
        }
        return true;
      } catch (error) {
        console.error('[RightActivityBar] Failed to persist connection secret:', error);
        return false;
      }
    }

    function removeSecret(connectionId) {
      if (!connectionId) return;
      try {
        localStorage.removeItem(`${CONNECTION_SECRET_PREFIX}${connectionId}`);
      } catch (error) {
        console.warn('[RightActivityBar] Failed to remove connection secret:', error);
      }
    }

    function cacheConnectionsSnapshot(list) {
      try { window.__aiConnectionsCache = JSON.parse(JSON.stringify(Array.isArray(list) ? list : [])) } catch (_) {}
    }

    function rehydrateIfEmpty() {
      try {
        if (!Array.isArray(state.connections) || state.connections.length === 0) {
          const cached = Array.isArray(window.__aiConnectionsCache) ? window.__aiConnectionsCache : [];
          if (cached.length) {
            state.connections = cached.map((c) => cloneConnection(c));
            // Write back to SettingsStore to keep sources in sync
            try { settingsStore.update({ aiConnections: state.connections.map((c) => cloneConnection(c)) }) } catch (_) {}
          } else {
            // Fallback: try direct SettingsStore getter helpers
            const ensured = window.SettingsStore?.ensureAiConnections?.();
            if (Array.isArray(ensured) && ensured.length) {
              state.connections = ensured.map((c) => cloneConnection(c));
              cacheConnectionsSnapshot(state.connections);
            }
          }
        }
      } catch (_) {}
    }

    function loadConnections(snapshot) {
      let settings = snapshot;
      if (!settings || typeof settings !== 'object') {
        settings = settingsStore.getState?.() || {};
      }
      if ((!settings.aiConnections || !Array.isArray(settings.aiConnections))
        && typeof settingsStore.ensureAiConnections === 'function') {
        try {
          const ensured = settingsStore.ensureAiConnections();
          if (Array.isArray(ensured)) {
            settings = { ...settings, aiConnections: ensured };
          }
        } catch (error) {
          console.warn('[RightActivityBar] ensureAiConnections failed during load:', error);
        }
      }
      const list = Array.isArray(settings.aiConnections) ? settings.aiConnections : [];
      state.connections = list
        .map((entry) => cloneConnection(entry))
        .filter(Boolean)
        .map((entry) => enforceFreeTierConnection(entry));
      if (state.connections.length) cacheConnectionsSnapshot(state.connections);
      rehydrateIfEmpty();
      rebuildConnectionSearchCache();
      state.connectionsSignature = buildConnectionsSignature(state.connections);
    }

    function buildConnectionsSignature(list) {
      if (!Array.isArray(list) || !list.length) return '';
      return list.map((connection) => {
        const id = connection?.id || '';
        const updatedAt = connection?.updatedAt || '';
        const providerId = connection?.providerId || '';
        const status = connection?.status || '';
        const name = connection?.name || '';
        const primaryModelId = connection?.primaryModelId || '';
        const commentLen = typeof connection?.comment === 'string' ? connection.comment.length : 0;
        return [id, updatedAt, providerId, status, primaryModelId, name, commentLen].join(':');
      }).join('|');
    }

    function rebuildConnectionSearchCache() {
      const seen = new Set();
      state.connections.forEach((connection) => {
        if (!connection || !connection.id) return;
        seen.add(connection.id);
        const haystack = [
          connection.name,
          connection.providerId,
          connection.comment,
          ...(Array.isArray(connection.tags) ? connection.tags : []),
          ...(Array.isArray(connection.models) ? connection.models.map((model) => model.id) : [])
        ].join(' ').toLowerCase();
        connectionSearchCache.set(connection.id, haystack);
      });
      Array.from(connectionSearchCache.keys()).forEach((key) => {
        if (!seen.has(key)) connectionSearchCache.delete(key);
      });
    }

    function syncConnectionCaches() {
      rebuildConnectionSearchCache();
      state.connectionsSignature = buildConnectionsSignature(state.connections);
      lastTreeKey = '';
    }

    function populateFilterOptions() {
      if (!filterEl) return;
      const providerIds = new Set();
      providerOptions.forEach((provider) => providerIds.add(provider.id));
      state.connections.forEach((connection) => providerIds.add(connection.providerId));
      const sorted = Array.from(providerIds).sort((a, b) => getProviderLabel(a).localeCompare(getProviderLabel(b)));
      const signature = sorted.join('|');
      const current = state.filter || 'all';
      if (signature !== state.filterOptionsSignature) {
        const options = ['<option value="all">All Providers</option>']
          .concat(sorted.map((id) => `<option value="${escapeAttr(id)}">${escapeHtml(getProviderLabel(id))}</option>`));
        filterEl.innerHTML = options.join('');
        state.filterOptionsSignature = signature;
      }
      if (current !== 'all' && !providerIds.has(current)) {
        state.filter = 'all';
        filterEl.value = 'all';
      } else {
        filterEl.value = current;
      }
    }

    function renderStatusToken(connection) {
      const status = connection.status || 'unknown';
      const option = CONNECTION_STATUS_OPTIONS.find((item) => item.value === status) || CONNECTION_STATUS_OPTIONS[0];
      return `<span class="ai-connection-status-pill ${option.className}">${escapeHtml(option.label)}</span>`;
    }

    function formatPrimaryModel(connection) {
      if (!Array.isArray(connection.models) || !connection.models.length) return 'No model';
      const primary = connection.primaryModelId
        ? connection.models.find((model) => model.id === connection.primaryModelId)
        : connection.models[0];
      return primary?.id || connection.models[0]?.id || 'No model';
    }

    function renderTree() {
      const searchTerm = state.search.trim().toLowerCase();
      const providerFilter = state.filter;
      const treeKey = `${state.connectionsSignature}::${providerFilter}::${searchTerm}::${state.selectedId || ''}`;
      if (treeKey === lastTreeKey) return;
      const grouped = new Map();
      state.connections.forEach((connection) => {
        if (providerFilter !== 'all' && connection.providerId !== providerFilter) {
          return;
        }
        if (searchTerm) {
          const haystack = connectionSearchCache.get(connection.id) || '';
          if (!haystack.includes(searchTerm)) {
            return;
          }
        }
        const bucket = grouped.get(connection.providerId) || [];
        bucket.push(connection);
        grouped.set(connection.providerId, bucket);
      });

      if (!grouped.size) {
        treeEl.innerHTML = '<div class="connection-empty">No connections found. Create one to get started.</div>';
        lastTreeKey = treeKey;
        return;
      }

      const providerIds = Array.from(grouped.keys()).sort((a, b) => getProviderLabel(a).localeCompare(getProviderLabel(b)));
      let html = '';
      providerIds.forEach((providerId) => {
        const label = getProviderLabel(providerId);
        const connections = grouped.get(providerId) || [];
        connections.sort((a, b) => a.name.localeCompare(b.name));
        html += `<div class="connection-provider" data-provider="${escapeAttr(providerId)}">`;
        html += `<div class="connection-provider-header">${escapeHtml(label)}<span class="badge">${connections.length}</span></div>`;
        html += '<div class="connection-provider-body">';
        connections.forEach((connection) => {
          const isActive = connection.id === state.selectedId;
          const commentSnippet = (connection.comment || '').split('\n').map((line) => line.trim()).find(Boolean) || '';
          html += `<div class="connection-item${isActive ? ' active' : ''}" data-connection="${escapeAttr(connection.id)}" role="button" tabindex="0">`;
          html += '<div class="connection-item-heading">';
          html += `<span class="connection-item-title">${escapeHtml(connection.name)}</span>`;
          html += `<span class="connection-item-menu-trigger" role="button" tabindex="0" aria-label="Connection options" aria-haspopup="menu" data-connection="${escapeAttr(connection.id)}" title="More actions">&#8942;</span>`;
          html += '</div>';
          html += `<div class="connection-item-meta"><span>${escapeHtml(formatPrimaryModel(connection))}</span>${renderStatusToken(connection)}</div>`;
          if (commentSnippet) {
            html += `<div class="connection-item-comment">${escapeHtml(commentSnippet)}</div>`;
          }
          html += '</div>';
        });
        html += '</div></div>';
      });
      treeEl.innerHTML = html;
      lastTreeKey = treeKey;
    }

    function ensureContextMenu() {
      if (contextMenuEl) return contextMenuEl;
      contextMenuEl = document.createElement('div');
      contextMenuEl.className = 'ai-connection-context-menu hidden';
      contextMenuEl.setAttribute('role', 'menu');
      contextMenuEl.addEventListener('mousedown', (event) => event.stopPropagation());
      contextMenuEl.addEventListener('contextmenu', (event) => event.preventDefault());
      document.body.appendChild(contextMenuEl);
      disposers.push(() => {
        try { contextMenuEl.remove(); } catch (_) {}
        contextMenuEl = null;
      });
      const handleDocumentClick = (event) => {
        if (suppressNextContextMenuDismiss) {
          suppressNextContextMenuDismiss = false;
          return;
        }
        if (contextMenuEl && contextMenuEl.contains(event.target)) {
          return;
        }
        hideContextMenu();
      };
      addListener(document, 'click', handleDocumentClick);
      addListener(window, 'blur', hideContextMenu);
      addListener(window, 'resize', hideContextMenu);
      addListener(document, 'scroll', hideContextMenu, true);
      addListener(document, 'keydown', (event) => {
        if (event.key === 'Escape') hideContextMenu();
      });
      return contextMenuEl;
    }

    function hideContextMenu() {
      if (!contextMenuEl) return;
      contextMenuEl.classList.add('hidden');
      contextMenuEl.innerHTML = '';
      suppressNextContextMenuDismiss = false;
    }

    function openContextMenu(event, connection) {
      if (!connection) return;
      if (event) {
        event.preventDefault();
        event.stopPropagation();
      }
      const menu = ensureContextMenu();
      hideContextMenu();
      suppressNextContextMenuDismiss = true;
      if (state.selectedId !== connection.id) {
        state.selectedId = connection.id;
        state.editing = cloneConnection(connection);
        renderTree();
        renderInspector();
      }
      const releaseHandler = () => {
        suppressNextContextMenuDismiss = false;
        window.removeEventListener('mouseup', releaseHandler, true);
      };
      window.addEventListener('mouseup', releaseHandler, true);
      const items = [
        {
          id: 'rename',
          label: 'Rename…',
          handler: () => promptRename(connection)
        },
        {
          id: 'copy',
          label: 'Copy',
          handler: () => duplicateConnection(connection)
        },
        {
          id: 'test',
          label: 'Test Connection',
          handler: () => testConnectionFromMenu(connection)
        },
        {
          id: 'delete',
          label: 'Delete',
          className: 'danger',
          handler: () => deleteConnectionFromMenu(connection)
        }
      ];
      menu.innerHTML = items.map((item) => `<button type="button" class="menu-item${item.className ? ` ${item.className}` : ''}" data-action="${item.id}">${escapeHtml(item.label)}</button>`).join('');
      items.forEach((item) => {
        const button = menu.querySelector(`[data-action="${item.id}"]`);
        if (!button) return;
        button.addEventListener('click', (event) => {
          event.preventDefault();
          event.stopPropagation();
          hideContextMenu();
          try { item.handler(); } catch (error) { console.error('[RightActivityBar] Context menu action failed:', error); }
        });
      });
      const { clientX, clientY } = event;
      menu.style.left = `${clientX}px`;
      menu.style.top = `${clientY}px`;
      menu.classList.remove('hidden');
      requestAnimationFrame(() => {
        const rect = menu.getBoundingClientRect();
        const margin = 8;
        const viewportWidth = window.innerWidth || document.documentElement?.clientWidth || rect.right;
        const viewportHeight = window.innerHeight || document.documentElement?.clientHeight || rect.bottom;
        const maxLeft = Math.max(margin, viewportWidth - rect.width - margin);
        const maxTop = Math.max(margin, viewportHeight - rect.height - margin);
        let left = Math.min(Math.max(rect.left, margin), maxLeft);
        let top = Math.min(Math.max(rect.top, margin), maxTop);
        menu.style.left = `${left}px`;
        menu.style.top = `${top}px`;
      });
    }

    function openContextMenuFromTrigger(trigger, connection) {
      if (!trigger || !connection) return;
      const rect = trigger.getBoundingClientRect();
      const syntheticEvent = {
        preventDefault: () => {},
        stopPropagation: () => {},
        clientX: rect.left + rect.width / 2,
        clientY: rect.bottom + 4,
        target: trigger
      };
      openContextMenu(syntheticEvent, connection);
    }

    function setFeedback(message, tone = 'info') {
      state.feedbackMessage = message || '';
      state.feedbackTone = tone;
      const feedbackEl = inspectorEl.querySelector('#ai-connection-feedback');
      if (!feedbackEl) return;
      feedbackEl.textContent = state.feedbackMessage;
      const color = tone === 'error'
        ? 'var(--vscode-errorForeground,#f48771)'
        : tone === 'success'
          ? 'var(--vscode-testing-iconPassed,#89d185)'
          : tone === 'warning'
            ? 'var(--vscode-testing-iconQueued,#c8c874)'
            : 'var(--vscode-descriptionForeground,#8a8a8a)';
      feedbackEl.style.color = color;
    }

    function formatTimestamp(timestamp) {
      if (!timestamp) return '—';
      try {
        return new Date(timestamp).toLocaleString();
      } catch (_) {
        return String(timestamp);
      }
    }

    function renderModelRows(connection) {
      const fallbackModel = getDefaultConnectionModel();
      const currentModel = typeof connection.primaryModelId === 'string' && connection.primaryModelId.trim().length
        ? connection.primaryModelId.trim()
        : fallbackModel;
      connection.primaryModelId = currentModel;
      connection.models = [{ id: currentModel }];
      return `
        <label class="ai-connection-model-label" for="ai-connection-model-input">Model ID</label>
        <div class="ai-connection-model-row ai-connection-model-row-input">
          <input id="ai-connection-model-input" class="vscode-input" type="text" value="${escapeAttr(currentModel)}" placeholder="${escapeAttr(fallbackModel)}" spellcheck="false" autocomplete="off" autocapitalize="off" />
          <button class="settings-btn settings-btn-secondary" id="ai-connection-model-reset" type="button" title="Reset to default model">Reset</button>
        </div>
      `;
    }

    function renderModelSuggestions(connection, catalog, isLoading = false, cacheEntry = null) {
      const suggestionsEl = inspectorEl.querySelector('#ai-connection-model-suggestions');
      if (!suggestionsEl) return;

      const list = Array.isArray(catalog) ? catalog : [];
      const loadingHtml = isLoading
        ? '<div class="ai-connection-model-suggestions-empty">Loading models…</div>'
        : '';
      const errorHtml = cacheEntry && cacheEntry.lastFetchError
        ? `<div class="ai-connection-model-suggestions-empty" style="color: var(--vscode-errorForeground,#f48771);">${escapeHtml(cacheEntry.lastFetchError)}</div>`
        : '';

      if (!list.length) {
        const fallbackModel = getDefaultConnectionModel();
        suggestionsEl.innerHTML = `${loadingHtml}${errorHtml}<div class="ai-connection-model-suggestions-empty">Type any supported OpenAI model id (e.g. gpt-4o, gpt-4.1, o4-mini). Leave the field empty to revert to ${escapeHtml(fallbackModel)}.</div>`;
        return;
      }

      const rows = list
        .slice(0, 30)
        .map((model) => {
          const id = typeof model?.id === 'string' ? model.id.trim() : '';
          if (!id) return '';
          return `<button type="button" class="settings-btn settings-btn-secondary ai-model-suggestion" data-model-id="${escapeAttr(id)}">${escapeHtml(id)}</button>`;
        })
        .filter(Boolean)
        .join('');

      const truncatedHint = list.length > 30
        ? `<div class="ai-connection-model-suggestions-empty">Showing first 30 of ${list.length} models.</div>`
        : '';

      suggestionsEl.innerHTML = `${loadingHtml}${errorHtml}<div class="ai-connection-model-suggestions-list">${rows}</div>${truncatedHint}`;
    }

    function renderInspector() {
      clearInspectorListeners();
      if (!state.editing) {
        inspectorEl.innerHTML = `
          <div class="ai-connection-empty">
            <p>Select a connection in the list or use “New Connection” to create one.</p>
          </div>
        `;
        setFeedback('');
        return;
      }
      const connection = state.editing;
      const secret = loadSecret(connection.id);
      const meta = isPlainObject(connection.metadata) ? connection.metadata : {};
      const reqConfig = isPlainObject(meta.requestConfig) ? meta.requestConfig : {};
      const initialTemp = typeof reqConfig.temperature === 'number' && Number.isFinite(reqConfig.temperature)
        ? reqConfig.temperature
        : 0.7;
      const initialMaxTokens = typeof reqConfig.maxTokens === 'number' && Number.isFinite(reqConfig.maxTokens) && reqConfig.maxTokens > 0
        ? reqConfig.maxTokens
        : 4096;
      enforceFreeTierConnection(connection);
      const providerOptionsHtml = `<option value="${AI_FREE_PROVIDER_ID}" selected>OpenAI</option>`;
      inspectorEl.innerHTML = `
        <section>
          <div class="ai-connection-field">
            <label for="ai-connection-name-input">Connection Name</label>
            <input id="ai-connection-name-input" type="text" class="vscode-input" value="${escapeAttr(connection.name || '')}" placeholder="e.g. Main Production" />
          </div>
          <div class="ai-connection-field">
            <label for="ai-connection-provider-select">Provider</label>
            <select id="ai-connection-provider-select" class="vscode-select" title="Provider switching is available in Pro version">
              ${providerOptionsHtml}
            </select>
            <div class="ai-connection-field-note">Provider switching is available in Pro version.</div>
          </div>
          <div class="ai-connection-field">
            <label>Status</label>
            <div class="ai-connection-status-static">${connection.status || 'unknown'}</div>
          </div>
          <div class="ai-connection-field">
            <label for="ai-connection-api-key">API Key</label>
            <div class="ai-connection-secret">
              <input id="ai-connection-api-key" type="password" class="vscode-input" value="${escapeAttr(secret)}" placeholder="${connection.providerId === 'github' ? 'ghp_… or github_pat_…' : 'sk-...'}" autocomplete="new-password" />
              <button class="settings-btn settings-btn-secondary" id="ai-connection-toggle-secret" title="Show/Hide secret">👁️</button>
              <button class="settings-btn settings-btn-secondary" id="ai-connection-save-secret" title="Save secret">💾</button>
            </div>
          </div>
          <div class="ai-connection-field">
            <label>
              <input id="ai-connection-advanced-logging" type="checkbox"${connection.metadata && connection.metadata.advancedLogging ? ' checked' : ''}>
              <span>Advanced logging for this connection (debug API requests)</span>
            </label>
          </div>
        </section>
        <section>
          <div class="section-title">Models</div>
          <div class="ai-connection-models" id="ai-connection-models-list">
            ${renderModelRows(connection)}
          </div>
          <div class="ai-connection-model-suggestions" id="ai-connection-model-suggestions"></div>
        </section>
        <section>
          <div class="section-title">Request parameters</div>
          <div class="ai-connection-field">
            <label for="ai-connection-temp">Temperature</label>
            <input id="ai-connection-temp" type="number" min="0" max="2" step="0.1" class="vscode-input" value="${escapeAttr(String(initialTemp))}" />
          </div>
          <div class="ai-connection-field">
            <label for="ai-connection-max-tokens">Max tokens</label>
            <input id="ai-connection-max-tokens" type="number" min="1" class="vscode-input" value="${escapeAttr(String(initialMaxTokens))}" />
          </div>
        </section>
        <section>
          <div class="ai-connection-field">
            <label for="ai-connection-comment-input">Comments</label>
            <textarea id="ai-connection-comment-input" class="vscode-textarea" rows="5" placeholder="Purpose, usage notes, limits…">${escapeHtml(connection.comment || '')}</textarea>
          </div>
        </section>
        <section>
          <div class="ai-connection-meta">
            <div><strong>Provider:</strong> ${escapeHtml(getProviderLabel(connection.providerId))}</div>
            <div><strong>Primary model:</strong> ${escapeHtml(formatPrimaryModel(connection))}</div>
            <div><strong>Created:</strong> ${formatTimestamp(connection.createdAt)}</div>
            <div><strong>Updated:</strong> ${formatTimestamp(connection.updatedAt)}</div>
            <div><strong>Last tested:</strong> ${formatTimestamp(connection.lastTestedAt)}</div>
          </div>
          <div class="ai-connection-inspector-footer">
            <div id="ai-connection-feedback" class="ai-connection-feedback">${escapeHtml(state.feedbackMessage)}</div>
            <div class="connection-actions">
              <button class="settings-btn settings-btn-secondary" id="ai-connection-test">Test</button>
              <button class="settings-btn settings-btn-primary" id="ai-connection-save">Save</button>
              <button class="settings-btn settings-btn-secondary" id="ai-connection-delete">Delete</button>
            </div>
          </div>
        </section>
      `;
      setFeedback(state.feedbackMessage, state.feedbackTone);
      let applyModelValue = null;
      const bindModelSuggestionButtons = () => {
        if (typeof applyModelValue !== 'function') return;
        Array.from(inspectorEl.querySelectorAll('.ai-model-suggestion')).forEach((button) => {
          addInspectorListener(button, 'click', (event) => {
            event.preventDefault();
            const nextModel = button.dataset.modelId || '';
            if (!nextModel) return;
            if (modelInput) {
              modelInput.value = nextModel;
            }
            applyModelValue(nextModel);
          });
        });
      };
      const updateModelSuggestions = (isLoading = false) => {
        const normalizedProviderId = normalizeProviderId(connection.providerId);
        let cachedList = getCachedProviderModels(connection.providerId);
        if (!Array.isArray(cachedList) || !cachedList.length) {
          cachedList = loadProviderModels(connection.providerId);
        }
        const cacheEntry = providerModelsCache.get(normalizedProviderId) || null;
        renderModelSuggestions(connection, cachedList, isLoading, cacheEntry);
        bindModelSuggestionButtons();
      };
      updateModelSuggestions(false);

      const nameInput = inspectorEl.querySelector('#ai-connection-name-input');
      addInspectorListener(nameInput, 'input', () => {
        connection.name = nameInput.value;
      });

      const providerSelect = inspectorEl.querySelector('#ai-connection-provider-select');
      __bindStableSelectActivation(providerSelect, () => {
        showProFeatureNotice({ featureName: 'Provider switching', subtitle: 'Switching provider is available in Pro version.' });
      }, { bindingKey: '__providerSelectNoticeBound' });
      addInspectorListener(providerSelect, 'change', () => {
        connection.providerId = AI_FREE_PROVIDER_ID;
        providerSelect.value = AI_FREE_PROVIDER_ID;
      });

      Array.from(inspectorEl.querySelectorAll('.connection-model-select')).forEach((selectNode) => {
        __bindStableSelectActivation(selectNode, () => {
          showProFeatureNotice({ featureName: 'Model selection', subtitle: 'Model customization is available in Pro version.' });
        }, { bindingKey: '__modelSelectNoticeBound' });
      });

      // Status now reflects the most recent test result and is not manually editable.

      const commentInput = inspectorEl.querySelector('#ai-connection-comment-input');
      addInspectorListener(commentInput, 'input', () => {
        connection.comment = commentInput.value;
      });

      const modelInput = inspectorEl.querySelector('#ai-connection-model-input');
      const resetModelBtn = inspectorEl.querySelector('#ai-connection-model-reset');
      applyModelValue = (rawValue) => {
        const trimmed = (rawValue || '').trim();
        const effective = trimmed || AI_FREE_LOCKED_MODEL_ID;
        connection.primaryModelId = effective;
        connection.models = [{ id: effective }];
        if (!trimmed && modelInput && modelInput.value !== effective) {
          modelInput.value = effective;
        }
        renderTree();
      };
      bindModelSuggestionButtons();
      addInspectorListener(modelInput, 'input', () => {
        applyModelValue(modelInput.value);
      });
      addInspectorListener(modelInput, 'blur', () => {
        applyModelValue(modelInput.value);
      });
      addInspectorListener(resetModelBtn, 'click', (event) => {
        event.preventDefault();
        if (modelInput) {
          modelInput.value = getDefaultConnectionModel();
        }
        applyModelValue(getDefaultConnectionModel());
      });

      const toggleSecretBtn = inspectorEl.querySelector('#ai-connection-toggle-secret');
      const secretInput = inspectorEl.querySelector('#ai-connection-api-key');
      addInspectorListener(toggleSecretBtn, 'click', (event) => {
        event.preventDefault();
        if (!secretInput) return;
        secretInput.type = secretInput.type === 'password' ? 'text' : 'password';
      });

      const saveSecretBtn = inspectorEl.querySelector('#ai-connection-save-secret');
      addInspectorListener(saveSecretBtn, 'click', (event) => {
        event.preventDefault();
        const value = secretInput?.value || '';
        const ok = saveSecret(connection.id, connection.providerId, value);
        if (ok) {
          setFeedback('Secret saved locally for this connection.', 'success');
          providerModelsCache.delete(normalizeProviderId(connection.providerId));
          updateModelSuggestions(true);
          scheduleProviderModelFetch(connection.providerId, {
            apiKey: value,
            onComplete: () => {
              if (state.editing && state.editing.id === connection.id) {
                renderTree();
                updateModelSuggestions(false);
              }
            }
          });
        } else {
          setFeedback('Failed to store secret.', 'error');
        }
      });

      const currentSecret = typeof secret === 'string' ? secret.trim() : '';
      if (currentSecret) {
        const normalizedProviderId = normalizeProviderId(connection.providerId);
        const cacheEntry = providerModelsCache.get(normalizedProviderId) || null;
        const hasRemoteModels = cacheEntry?.source === 'remote' && Array.isArray(cacheEntry?.models) && cacheEntry.models.length > 0;
        if (!hasRemoteModels) {
          updateModelSuggestions(true);
          scheduleProviderModelFetch(connection.providerId, {
            apiKey: currentSecret,
            onComplete: () => {
              if (state.editing && state.editing.id === connection.id) {
                updateModelSuggestions(false);
              }
            }
          });
        }
      }

      const advancedLoggingEl = inspectorEl.querySelector('#ai-connection-advanced-logging');
      if (advancedLoggingEl) {
        addInspectorListener(advancedLoggingEl, 'change', () => {
          if (!isPlainObject(connection.metadata)) {
            connection.metadata = {};
          }
          connection.metadata.advancedLogging = advancedLoggingEl.checked === true;
        });
      }

      const tempInput = inspectorEl.querySelector('#ai-connection-temp');
      if (tempInput) {
        addInspectorListener(tempInput, 'change', () => {
          if (!isPlainObject(connection.metadata)) connection.metadata = {};
          if (!isPlainObject(connection.metadata.requestConfig)) connection.metadata.requestConfig = {};
          const value = parseFloat(tempInput.value);
          if (Number.isFinite(value)) {
            connection.metadata.requestConfig.temperature = value;
          } else {
            delete connection.metadata.requestConfig.temperature;
          }
        });
      }

      const maxTokensInput = inspectorEl.querySelector('#ai-connection-max-tokens');
      if (maxTokensInput) {
        addInspectorListener(maxTokensInput, 'change', () => {
          if (!isPlainObject(connection.metadata)) connection.metadata = {};
          if (!isPlainObject(connection.metadata.requestConfig)) connection.metadata.requestConfig = {};
          const value = parseInt(maxTokensInput.value, 10);
          if (Number.isFinite(value) && value > 0) {
            connection.metadata.requestConfig.maxTokens = value;
          } else {
            delete connection.metadata.requestConfig.maxTokens;
          }
        });
      }

      const testBtn = inspectorEl.querySelector('#ai-connection-test');
      addInspectorListener(testBtn, 'click', async (event) => {
        event.preventDefault();
        const key = secretInput?.value?.trim();
        if (!key) {
          setFeedback('Add an API key before testing the connection.', 'warning');
          return;
        }
        if (connection.providerId === 'openai' && !key.startsWith('sk-')) {
          setFeedback('OpenAI keys should start with “sk-”.', 'error');
          return;
        }
        if (connection.providerId === 'github') {
          const hasKnownPrefix = key.startsWith('ghp_') ||
            key.startsWith('gho_') ||
            key.startsWith('ghu_') ||
            key.startsWith('ghs_') ||
            key.startsWith('github_pat_');
          if (!hasKnownPrefix) {
            setFeedback('GitHub tokens should start with gh* (e.g. “ghp_”, “github_pat_”).', 'error');
            return;
          }
        }
        if (!providerFactory || typeof providerFactory.createProvider !== 'function') {
          connection.lastTestedAt = Date.now();
          connection.status = 'ok';
          setFeedback('Connection passed basic validation (provider unavailable for live test).', 'warning');
          renderTree();
          return;
        }
        setFeedback('Running live test request…', 'info');
        testBtn.disabled = true;
        try {
          const normalized = normalizeProviderId(connection.providerId);
          const provider = providerFactory.createProvider(normalized, {
            advancedLogging: isPlainObject(connection.metadata) && connection.metadata.advancedLogging === true,
            connectionId: connection.id
          });
          const messages = [
            { role: 'user', content: 'What is your name and version?' }
          ];
          const requestOptions = {};
          if (connection.primaryModelId) {
            requestOptions.model = connection.primaryModelId;
          }
          const metaCfg = isPlainObject(connection.metadata) && isPlainObject(connection.metadata.requestConfig)
            ? connection.metadata.requestConfig
            : null;
          if (metaCfg) {
            if (typeof metaCfg.temperature === 'number' && Number.isFinite(metaCfg.temperature)) {
              requestOptions.temperature = metaCfg.temperature;
            }
            if (typeof metaCfg.maxTokens === 'number' && Number.isFinite(metaCfg.maxTokens) && metaCfg.maxTokens > 0) {
              requestOptions.max_tokens = metaCfg.maxTokens;
            }
          }
          const answer = await provider.sendMessage(messages, requestOptions);
          const snippet = typeof answer === 'string' ? answer.slice(0, 160) : '';
          connection.lastTestedAt = Date.now();
          connection.status = 'ok';
          setFeedback(snippet ? `Test OK: ${snippet}` : 'Test OK (empty response).', 'success');
          try {
            window.LogService?.logInfo?.('AI Connections', 'test-success', {
              id: connection.id,
              provider: connection.providerId
            });
          } catch (_) {}
        } catch (error) {
          connection.lastTestedAt = Date.now();
          connection.status = 'error';
          setFeedback(`Test failed: ${error?.message || error}`, 'error');
          try {
            window.LogService?.logError?.('AI Connections', 'test-failed', {
              id: connection.id,
              provider: connection.providerId,
              error: error?.message || String(error)
            });
          } catch (_) {}
        } finally {
          testBtn.disabled = false;
          try {
            document.dispatchEvent(new CustomEvent('aiConnections:updated', { detail: { connections: state.connections } }));
          } catch (_) {}
          try {
            if (typeof refreshConnectionsFromSettings === 'function') {
              refreshConnectionsFromSettings({ aiConnections: state.connections });
            }
          } catch (_) {}
          renderTree();
          renderInspector();
        }
      });

      const saveBtn = inspectorEl.querySelector('#ai-connection-save');
      addInspectorListener(saveBtn, 'click', (event) => {
        event.preventDefault();
        saveCurrentConnection();
      });

      const deleteBtn = inspectorEl.querySelector('#ai-connection-delete');
      addInspectorListener(deleteBtn, 'click', (event) => {
        event.preventDefault();
        deleteCurrentConnection();
      });

      if (typeof state.pendingModelFocusIndex === 'number' && state.pendingModelFocusIndex >= 0) {
        const targetRow = modelsList?.querySelector(`.ai-connection-model-row[data-index="${state.pendingModelFocusIndex}"]`);
        const targetInput = targetRow?.querySelector('.connection-model-input');
        if (targetInput) {
          const length = targetInput.value?.length || 0;
          try {
            targetInput.focus();
            targetInput.setSelectionRange(length, length);
          } catch (_) {}
          state.modelFocusIndex = state.pendingModelFocusIndex;
        }
        state.pendingModelFocusIndex = null;
      }
    }

    function sanitiseModels(connection) {
      if (!Array.isArray(connection.models)) {
        return [];
      }
      const seen = new Set();
      const models = [];
      connection.models.forEach((model) => {
        const id = typeof model?.id === 'string' ? model.id.trim() : '';
        if (!id || seen.has(id)) return;
        seen.add(id);
        models.push({ id, note: typeof model?.note === 'string' ? model.note : '' });
      });
      return models;
    }

    function saveCurrentConnection() {
      if (!state.editing) return;
      const draft = state.editing;
      const name = (draft.name || '').trim();
      if (!name) {
        setFeedback('Connection name is required.', 'error');
        return;
      }
      draft.name = name;
      draft.providerId = AI_FREE_PROVIDER_ID;
      const fallbackModel = getDefaultConnectionModel();
      const models = sanitiseModels(draft);
      draft.models = models.length ? models : [{ id: (draft.primaryModelId || fallbackModel) }];
      draft.primaryModelId = (typeof draft.primaryModelId === 'string' && draft.primaryModelId.trim())
        ? draft.primaryModelId.trim()
        : (draft.models[0]?.id || fallbackModel);
      draft.status = draft.status || 'unknown';
      draft.comment = draft.comment || '';
      draft.tags = Array.isArray(draft.tags)
        ? draft.tags.filter((tag) => typeof tag === 'string' && tag.trim()).map((tag) => tag.trim())
        : [];
      if (draft.isNew) {
        delete draft.isNew;
      }

      const now = Date.now();
      const payload = {
        id: draft.id || `conn-${Math.random().toString(36).slice(2, 9)}${Date.now().toString(36)}`,
        providerId: draft.providerId,
        name: draft.name,
        models: draft.models.map((model) => ({ ...model })),
        primaryModelId: draft.primaryModelId || '',
        comment: draft.comment,
        tags: draft.tags,
        status: draft.status,
        createdAt: draft.createdAt || now,
        updatedAt: now,
        lastTestedAt: draft.lastTestedAt || null,
        metadata: isPlainObject(draft.metadata) ? { ...draft.metadata } : {}
      };

      const next = state.connections.slice();
      const existingIndex = next.findIndex((connection) => connection.id === payload.id);
      if (existingIndex >= 0) {
        payload.createdAt = next[existingIndex].createdAt || payload.createdAt;
        next[existingIndex] = cloneConnection(payload);
      } else {
        next.push(cloneConnection(payload));
      }

      try {
        state.suppressStore = true;
        settingsStore.update({ aiConnections: next.map((connection) => {
          const copy = cloneConnection(connection);
          copy.metadata = isPlainObject(copy.metadata) ? { ...copy.metadata } : {};
          return copy;
        }) });
        // Update in-memory cache immediately to survive intermediate remounts
        try { window.__aiConnectionsCache = JSON.parse(JSON.stringify(next)) } catch (_) {}
        setTimeout(() => { state.suppressStore = false; }, 0);
      } catch (error) {
        state.suppressStore = false;
        console.error('[RightActivityBar] Failed to persist AI connections:', error);
        setFeedback(`Save failed: ${error?.message || error}`, 'error');
        return;
      }

      state.connections = next;
      syncConnectionCaches();
      state.selectedId = payload.id;
      state.editing = cloneConnection(payload);
      setFeedback('Connection saved.', 'success');
      populateFilterOptions();
      renderTree();
      renderInspector();
      try {
        document.dispatchEvent(new CustomEvent('aiConnections:updated', { detail: { connections: state.connections } }));
      } catch (_) {}
      try {
        if (typeof refreshConnectionsFromSettings === 'function') {
          refreshConnectionsFromSettings({ aiConnections: state.connections });
        }
      } catch (_) {}
    }

    function deleteCurrentConnection() {
      if (!state.editing) return;
      const target = state.editing;
      const persisted = state.connections.some((connection) => connection.id === target.id && !connection.isNew);
      const proceed = persisted ? window.confirm(`Delete connection "${target.name}"?`) : true;
      if (!proceed) return;

      let next = state.connections.filter((connection) => connection.id !== target.id);
      if (persisted) {
        try {
          state.suppressStore = true;
          settingsStore.update({ aiConnections: next.map((connection) => cloneConnection(connection)) });
          setTimeout(() => { state.suppressStore = false; }, 0);
        } catch (error) {
          state.suppressStore = false;
          console.error('[RightActivityBar] Failed to delete connection:', error);
          setFeedback(`Delete failed: ${error?.message || error}`, 'error');
          return;
        }
      }
      removeSecret(target.id);
      state.connections = next;
      syncConnectionCaches();
      state.selectedId = next[0]?.id || null;
      state.editing = state.selectedId ? cloneConnection(next[0]) : null;
      setFeedback('Connection removed.', 'info');
      populateFilterOptions();
      renderTree();
      renderInspector();
    }

    function startCreateConnection(defaultProvider) {
      const providerId = AI_FREE_PROVIDER_ID;
      const defaultModel = getDefaultConnectionModel();
      const now = Date.now();
      const draft = {
        id: `conn-${Math.random().toString(36).slice(2, 9)}${now.toString(36)}`,
        providerId,
        name: `New ${getProviderLabel(providerId)}`,
        models: [{ id: defaultModel }],
        primaryModelId: defaultModel,
        comment: '',
        tags: [],
        status: 'unknown',
        createdAt: now,
        updatedAt: now,
        lastTestedAt: null,
        metadata: {},
        isNew: true
      };
      state.connections = state.connections.concat([cloneConnection(draft)]);
      syncConnectionCaches();
      state.editing = draft;
      state.selectedId = draft.id;
      populateFilterOptions();
      renderTree();
      renderInspector();
    }

    addListener(treeEl, 'click', (event) => {
      const trigger = event.target.closest('.connection-item-menu-trigger');
      if (trigger) {
        event.preventDefault();
        event.stopPropagation();
        const id = trigger.dataset.connection;
        if (!id) return;
        const connection = state.connections.find((entry) => entry.id === id);
        if (!connection) return;
        openContextMenuFromTrigger(trigger, connection);
        return;
      }
      const target = event.target.closest('.connection-item');
      if (!target) return;
      const id = target.dataset.connection;
      if (!id) return;
      const connection = state.connections.find((entry) => entry.id === id);
      if (!connection) return;
      state.selectedId = id;
      state.editing = cloneConnection(connection);
      renderTree();
      renderInspector();
    });

    addListener(treeEl, 'keydown', (event) => {
      if (event.key !== 'Enter' && event.key !== ' ') return;
      const trigger = event.target.closest('.connection-item-menu-trigger');
      if (trigger) {
        event.preventDefault();
        event.stopPropagation();
        const id = trigger.dataset.connection;
        if (!id) return;
        const connection = state.connections.find((entry) => entry.id === id);
        if (!connection) return;
        openContextMenuFromTrigger(trigger, connection);
        return;
      }
      const target = event.target.closest('.connection-item');
      if (!target) return;
      const id = target.dataset.connection;
      if (!id) return;
      const connection = state.connections.find((entry) => entry.id === id);
      if (!connection) return;
      event.preventDefault();
      event.stopPropagation();
      state.selectedId = id;
      state.editing = cloneConnection(connection);
      renderTree();
      renderInspector();
    });

    addListener(treeEl, 'contextmenu', (event) => {
      const item = event.target.closest('.connection-item');
      if (!item) return;
      event.preventDefault();
      event.stopPropagation();
      const id = item.dataset.connection;
      if (!id) return;
      const connection = state.connections.find((entry) => entry.id === id);
      if (!connection) return;
      openContextMenu(event, connection);
    });

    if (filterEl) {
      addListener(filterEl, 'change', () => {
        state.filter = filterEl.value || 'all';
        renderTree();
      });
      try {
        const host = filterEl.closest?.('.ai-connection-pane') || document.getElementById('right-pane-content')
        if (host) __attachCustomPicker(filterEl, host)
      } catch (_) {}
    }

    if (searchEl) {
      addListener(searchEl, 'input', () => {
        state.search = searchEl.value || '';
        renderTree();
      });
    }

    if (addButton) {
      addListener(addButton, 'click', (event) => {
        event.preventDefault();
        startCreateConnection();
      });
    }

    let unsubscribe = null;
    if (typeof settingsStore.subscribe === 'function') {
      unsubscribe = settingsStore.subscribe((snapshot) => {
        if (state.suppressStore) return;
        loadConnections(snapshot);
        populateFilterOptions();
        if (!state.selectedId || !state.connections.some((connection) => connection.id === state.selectedId)) {
          state.selectedId = state.connections[0]?.id || null;
        }
        state.editing = state.selectedId
          ? cloneConnection(state.connections.find((connection) => connection.id === state.selectedId))
          : null;
        renderTree();
        renderInspector();
      });
      if (unsubscribe) {
        disposers.push(() => {
          try { unsubscribe(); } catch (_) {}
        });
      }
    }

    loadConnections();
    initConnectionSplitControls();
    populateFilterOptions();
    if (state.connections.length && !state.selectedId) {
      state.selectedId = state.connections[0].id;
    }
    state.editing = state.selectedId
      ? cloneConnection(state.connections.find((connection) => connection.id === state.selectedId))
      : null;
    renderTree();
    renderInspector();

    return function cleanup() {
      clearInspectorListeners();
      disposers.forEach((dispose) => {
        try { dispose(); } catch (_) {}
      });
    };
  }

  function renderAISettings(defaultTab, options = {}) {
    const container = document.getElementById('right-pane-content');
    if (!container) return;
    disposeCodeGraph();

    if (agentTreeInstance && typeof agentTreeInstance.dispose === 'function') {
      try { agentTreeInstance.dispose(); } catch (_) {}
      agentTreeInstance = null;
    }

    let preservedPromptDraft = '';
    if (window.__rightPromptEditor && typeof window.__rightPromptEditor.getValue === 'function') {
      try { preservedPromptDraft = window.__rightPromptEditor.getValue(); } catch (_) {}
      try { window.__rightPromptEditor.dispose(); } catch (_) {}
      window.__rightPromptEditor = null;
    }
    if (window.__rightPromptEditorDisposer && typeof window.__rightPromptEditorDisposer.dispose === 'function') {
      try { window.__rightPromptEditorDisposer.dispose(); } catch (_) {}
    }
    window.__rightPromptEditorDisposer = null;
    if (window.__rightPromptTextArea) {
      try { if (!preservedPromptDraft) preservedPromptDraft = window.__rightPromptTextArea.value || ''; } catch (_) {}
      if (window.__rightPromptTextAreaListener) {
        try { window.__rightPromptTextArea.removeEventListener('input', window.__rightPromptTextAreaListener); } catch (_) {}
      }
      window.__rightPromptTextArea = null;
    }
    window.__rightPromptTextAreaListener = null;

    container.innerHTML = `
      <div class="right-ai-tabs" style="display:flex; gap:6px; padding:6px 10px; border-bottom:1px solid var(--vscode-sideBar-border,#2e2e2e);">
        <button class="settings-btn settings-btn-secondary" data-tab="connection" id="ai-tab-connection-btn">Connections</button>
        <button class="settings-btn settings-btn-secondary" data-tab="agents" id="ai-tab-agents-btn">Agents</button>
      </div>
      <div id="ai-tab-connection" class="ai-tab" style="padding:0; height:100%; display:flex; flex-direction:column;">
        <div class="ai-connection-pane">
          <div class="ai-connection-toolbar">
            <select id="ai-connection-provider-filter" class="vscode-select" title="Filter by provider">
              <option value="all">All Providers</option>
            </select>
            <div class="ai-connection-search">
              <input type="search" id="ai-connection-search" class="vscode-input" placeholder="Search connections…" aria-label="Search connections">
            </div>
            <button class="settings-btn settings-btn-primary" id="ai-connection-add">+ New Connection</button>
          </div>
          <div class="ai-connection-content">
            <div class="ai-connection-tree" id="ai-connection-tree" role="tree"></div>
            <div class="ai-connection-splitter" id="ai-connection-splitter" role="separator" tabindex="0" aria-orientation="vertical" aria-label="Resize connection panels">
              <div class="ai-connection-splitter-grip" aria-hidden="true"></div>
            </div>
            <div class="ai-connection-inspector" id="ai-connection-inspector">
              <div class="ai-connection-empty">
                <p>Select a connection on the left or create a new one to get started.</p>
              </div>
            </div>
          </div>
        </div>
      </div>

      <div id="ai-tab-agents" class="ai-tab ai-agents-view" style="display:none; height:100%;">
        <div class="ai-agents-split">
          <div class="ai-agents-tree-pane">
            <div id="ai-agent-tree-container" class="ai-agents-panel"></div>
          </div>
          <div class="ai-agent-detail">
            <div id="ai-agent-empty-state" class="ai-agent-empty-state">
              <p>Select an agent from the list to configure settings and prompts.</p>
            </div>
            <div id="ai-agent-detail-body" class="ai-agent-detail-body">
              <div class="ai-agent-detail-tabs">
                <button class="settings-btn settings-btn-primary" data-agent-pane-btn="config">Configuration</button>
                <button class="settings-btn settings-btn-secondary" data-agent-pane-btn="prompt">Prompt</button>
              </div>
              <div id="ai-agent-config-panel" class="ai-agent-pane" data-agent-pane="config">
                <div class="right-pane-section ai-agent-config-section">
                  <div class="section-row ai-agent-config-row">
                    <div class="settings-label" style="min-width: 180px;">
                      <span>Agent mode</span>
                      <div class="ai-connection-status-static">Assistant (single mode)</div>
                    </div>
                    <label class="settings-label">
                      <span>Stored Connection</span>
                      <select id="ai-agent-connection" class="vscode-select"></select>
                    </label>
                  </div>
                  <div class="section-row ai-agent-description-row">
                    <label class="settings-label">
                      <span>Description</span>
                      <textarea id="ai-agent-description" class="vscode-textarea" rows="3" placeholder="Short summary for the agent"></textarea>
                    </label>
                  </div>
                  <div class="section-row ai-agent-actions-row">
                    <button class="settings-btn settings-btn-primary" id="ai-agent-config-save">Save Configuration</button>
                  </div>
                </div>
              </div>
              <div id="ai-agent-prompt-panel" class="ai-agent-pane" data-agent-pane="prompt" style="display:none;">
                <div class="right-pane-section ai-agent-prompt-section">
                  <div class="section-header">System Prompt (XML interpreted)</div>
                <div class="ai-agent-connection-hint" id="ai-agent-connection-hint">Using provider <strong>OpenAI</strong></div>
                  <div id="ai-right-prompt-editor" class="ai-agent-prompt-editor"></div>
                  <div class="section-row ai-agent-prompt-footer">
                    <span id="ai-right-prompt-counter">0 / 10,000 characters</span>
                    <div class="btn-row">
                      <button class="settings-btn settings-btn-secondary" id="ai-right-load-default" disabled title="Available in Pro version">📋 Load Default</button>
                      <button class="settings-btn settings-btn-secondary" id="ai-right-clear" disabled title="Available in Pro version">🗑️ Clear</button>
                      <button class="settings-btn settings-btn-primary" id="ai-right-save-prompt" disabled title="Available in Pro version">💾 Save</button>
                      <button class="settings-btn settings-btn-secondary" data-pro-lock="prompt">🔒 Prompt editing — Pro</button>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    `;
    const allowedTabs = ['connection', 'agents'];
    const requestedVisible = Array.isArray(options?.visibleTabs) ? options.visibleTabs : null;
    const normalizedVisible = requestedVisible && requestedVisible.length
      ? requestedVisible.filter(tab => allowedTabs.includes(tab))
      : allowedTabs;
    const visibleTabs = normalizedVisible.length ? normalizedVisible : ['connection'];
    const visibleSet = new Set(visibleTabs);
    let initialTab = defaultTab && visibleSet.has(defaultTab) ? defaultTab : null;
    if (!initialTab) {
      try {
        const storedTab = localStorage.getItem('right_pane_ai_tab');
        if (storedTab && visibleSet.has(storedTab)) initialTab = storedTab;
      } catch (_) {}
    }
    if (!initialTab) initialTab = visibleTabs[0] || 'connection';
    currentAgentsTab = initialTab;
    const multiTab = visibleSet.size > 1;
    const tabStrip = container.querySelector('.right-ai-tabs');
    if (tabStrip) {
      tabStrip.style.display = multiTab ? 'flex' : 'none';
      tabStrip.setAttribute('aria-hidden', multiTab ? 'false' : 'true');
    }
    // Tabs
    const btnConn = document.getElementById('ai-tab-connection-btn');
    const btnPrompt = document.getElementById('ai-tab-prompt-btn');
    const tabConn = document.getElementById('ai-tab-connection');
    const tabPrompt = document.getElementById('ai-tab-prompt');
    const btnAgents = document.getElementById('ai-tab-agents-btn');
    const tabAgents = document.getElementById('ai-tab-agents');
    if (!visibleSet.has('connection') && btnConn) {
      btnConn.style.display = 'none';
      btnConn.setAttribute('aria-hidden', 'true');
    } else {
      btnConn?.setAttribute('aria-hidden', 'false');
    }
    if (btnPrompt) {
      btnPrompt.style.display = 'none';
      btnPrompt.setAttribute('aria-hidden', 'true');
    }
    if (!visibleSet.has('agents') && btnAgents) {
      btnAgents.style.display = 'none';
      btnAgents.setAttribute('aria-hidden', 'true');
    } else {
      btnAgents?.setAttribute('aria-hidden', 'false');
    }
    if (!visibleSet.has('connection') && tabConn) tabConn.style.display = 'none';
    if (tabPrompt) tabPrompt.style.display = 'none';
    if (!visibleSet.has('agents') && tabAgents) tabAgents.style.display = 'none';
    const agentTreeContainer = document.getElementById('ai-agent-tree-container');
    const agentTypeGroup = Array.from(container.querySelectorAll('input[name="ai-agent-type"]'));
    const agentConnectionSelect = document.getElementById('ai-agent-connection');
    try {
      const host = document.body || document.getElementById('right-pane-content') || container;
      if (agentConnectionSelect && host) __attachCustomPicker(agentConnectionSelect, host)
    } catch (_) {}
    const agentDescriptionInput = document.getElementById('ai-agent-description');
    const agentConfigSaveBtn = document.getElementById('ai-agent-config-save');
    const agentPaneButtons = Array.from(container.querySelectorAll('[data-agent-pane-btn]'));
    const agentPanes = Array.from(container.querySelectorAll('.ai-agent-pane'));
    const agentConnectionHint = document.getElementById('ai-agent-connection-hint');
    const agentDetailBody = document.getElementById('ai-agent-detail-body');
    const agentEmptyState = document.getElementById('ai-agent-empty-state');
    const counter = document.getElementById('ai-right-prompt-counter');
    const loadDefaultBtn = document.getElementById('ai-right-load-default');
    const clearBtn = document.getElementById('ai-right-clear');
    const saveBtn = document.getElementById('ai-right-save-prompt');
    const btnConvertVba = document.getElementById('ai-agent-action-convert-vba');
    const btnRefactorMacro = document.getElementById('ai-agent-action-refactor');
    const btnExplainMacro = document.getElementById('ai-agent-action-explain');
    const btnReviewMacro = document.getElementById('ai-agent-action-review');
    const jobStatusEl = document.getElementById('ai-agent-job-status');
    const jobOutputEl = document.getElementById('ai-agent-job-output');
    const jobHistorySelect = document.getElementById('ai-agent-job-history-select');
    const jobDetailsCopyBtn = document.getElementById('ai-agent-job-details-copy');
    try {
      const host = document.getElementById('right-pane-content') || container;
      if (jobHistorySelect && host) __attachCustomPicker(jobHistorySelect, host)
    } catch (_) {}
    let promptEditor = window.__rightPromptEditor || null;
    let currentProviderId = null;
    let currentAgentId = null;
    let currentConnectionId = null;
    let suppressAgentFormEvents = false;

    Array.from(container.querySelectorAll('[data-pro-lock]')).forEach((btn) => {
      btn.addEventListener('click', (event) => {
        event.preventDefault();
        const type = btn.getAttribute('data-pro-lock') || 'feature';
        const labels = {
          provider: 'Provider switching',
          model: 'Model selection',
          prompt: 'System prompt editing'
        };
        showProFeatureNotice({
          featureName: labels[type] || 'This feature',
          subtitle: 'Available in Pro version'
        });
      });
    });

    function cacheKey(providerId, agentId = currentAgentId) {
      const provider = (providerId || currentProviderId || 'openai').toLowerCase();
      const agent = agentId || currentAgentId || '__global__';
      return `${agent}::${provider}`;
    }

    if (window.AIPromptService?.migrateLegacyPrompt) {
      try { window.AIPromptService.migrateLegacyPrompt(); } catch (_) {}
    }

    const providerCatalog = getProviderCatalog();
    currentProviderId = resolveInitialProvider(providerCatalog);
    window.__aiPromptProvider = currentProviderId;
    try { localStorage.setItem('ai_prompt_provider', currentProviderId); } catch (_) {}

    if (preservedPromptDraft && typeof preservedPromptDraft === 'string') {
      PROVIDER_TEXT_CACHE.set(cacheKey(currentProviderId), preservedPromptDraft);
    }

    const escapeHtml = (value) => String(value ?? '').replace(/[&<>"']/g, (char) => ({
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    }[char]));

    let connectionCatalog = [];
    const cachedProviders = new Map();

    function getConnectionCatalog() {
      try {
        const list = window.SettingsStore?.getAiConnections?.();
        connectionCatalog = Array.isArray(list) ? list : [];
      } catch (error) {
        console.warn('[RightActivityBar] Failed to read stored connections:', error);
        connectionCatalog = [];
      }
      return connectionCatalog;
    }

    function findConnection(connectionId) {
      if (!connectionId) return null;
      if (!connectionCatalog.length) getConnectionCatalog();
      return connectionCatalog.find((entry) => entry && entry.id === connectionId) || null;
    }

    function populateAgentConnections(activeConnectionId, sourceList) {
      if (!agentConnectionSelect) return;
      let list = Array.isArray(sourceList) ? sourceList.slice() : null;
      if ((!list || !list.length) && window.SettingsStore?.ensureAiConnections) {
        try {
          list = window.SettingsStore.ensureAiConnections();
        } catch (error) {
          console.warn('[RightActivityBar] ensureAiConnections failed:', error);
        }
      }
      if (!Array.isArray(list)) {
        list = getConnectionCatalog();
      }
      connectionCatalog = Array.isArray(list) ? list.map((entry) => ({ ...entry })) : [];
      const placeholder = '<option value="">Select connection…</option>';
      const options = connectionCatalog.map((connection) => {
        const name = escapeHtml(connection.name || connection.id);
        return `<option value="${escapeHtml(connection.id)}">${name}</option>`;
      });
      agentConnectionSelect.innerHTML = [placeholder, ...options].join('');
      if (activeConnectionId && connectionCatalog.some((connection) => connection.id === activeConnectionId)) {
        agentConnectionSelect.value = activeConnectionId;
      } else {
        agentConnectionSelect.value = '';
      }
      connectionCatalog.forEach((connection) => {
        cachedProviders.set(connection.id, normalizeProvider(connection.providerId || currentProviderId));
      });
    }

    function resolvePromptProviderLabel(providerId) {
      const normalized = normalizeProvider(providerId);
      if (Array.isArray(providerCatalog)) {
        const match = providerCatalog.find((entry) => normalizeProvider(entry.id) === normalized);
        if (match && match.label) return match.label;
      }
      return typeof providerId === 'string' && providerId.length
        ? providerId.replace(/[-_]+/g, ' ').replace(/\b\w/g, (c) => c.toUpperCase())
        : 'Unknown provider';
    }

    function updateConnectionHint(connection, providerId) {
      if (!agentConnectionHint) return;
      const providerLabel = escapeHtml(resolvePromptProviderLabel(providerId) || '');
      if (connection) {
        const name = escapeHtml(connection.name || connection.id);
        agentConnectionHint.innerHTML = `Using connection <strong>${name}</strong> (${providerLabel || 'Unknown provider'})`;
      } else {
        agentConnectionHint.innerHTML = providerLabel
          ? `Using provider <strong>${providerLabel}</strong>`
          : 'No connection selected.';
      }
    }

    let connectionsSubscription = null;
    if (window.__aiConnectionsSubscription && typeof window.__aiConnectionsSubscription === 'function') {
      try { window.__aiConnectionsSubscription(); } catch (_) {}
      window.__aiConnectionsSubscription = null;
    }
    function refreshConnectionsFromSettings (settings) {
      if (!agentConnectionSelect) return;
      const list = Array.isArray(settings?.aiConnections) ? settings.aiConnections : null;
      if (!Array.isArray(list)) return;
      const hasChanged = list.length !== connectionCatalog.length
        || list.some((connection, index) => {
          const existing = connectionCatalog[index];
          return !existing || existing.id !== connection.id || existing.updatedAt !== connection.updatedAt;
        });
      if (!hasChanged) return;
      suppressAgentFormEvents = true;
      populateAgentConnections(currentConnectionId, list);
      suppressAgentFormEvents = false;
      const activeConnection = currentConnectionId ? findConnection(currentConnectionId) : null;
      if (currentConnectionId && !activeConnection) {
        applyConnectionSelection('', { persist: true, fallbackProviderId: currentProviderId });
        persistAgentConfiguration();
      } else {
        updateConnectionHint(activeConnection, currentProviderId);
      }
    }

    if (window.SettingsStore?.subscribe) {
      try {
        connectionsSubscription = window.SettingsStore.subscribe((settings) => {
          if (!settings || typeof settings !== 'object') return;
          refreshConnectionsFromSettings(settings);
        });
        window.__aiConnectionsSubscription = connectionsSubscription;
      } catch (error) {
        console.warn('[RightActivityBar] Failed to subscribe to SettingsStore for connections:', error);
      }
    }

    function applyConnectionSelection(connectionId, options = {}) {
      const { persist = false, fallbackProviderId = null } = options;
      const connection = connectionId ? findConnection(connectionId) : null;
      if (agentConnectionSelect) {
        if (connection && agentConnectionSelect.value !== connection.id) {
          agentConnectionSelect.value = connection.id;
        } else if (connectionId && !connection) {
          agentConnectionSelect.value = '';
        }
      }
      const providerCandidate = connection?.providerId
        || cachedProviders.get(connectionId || '')
        || fallbackProviderId
        || currentProviderId
        || 'openai';
      const normalizedProvider = normalizeProvider(providerCandidate);
      const providerChanged = normalizedProvider && normalizedProvider !== currentProviderId;
      if (providerChanged) {
        persistCurrentProviderState();
      }
      currentConnectionId = connection ? connection.id : null;
      currentProviderId = normalizedProvider || currentProviderId || 'openai';
      window.__aiPromptProvider = currentProviderId;
      try { localStorage.setItem('ai_prompt_provider', currentProviderId); } catch (_) {}
      updateConnectionHint(connection, currentProviderId);
      loadProviderPrompt(currentProviderId);
      updateCount();
      if (persist && currentAgentId) {
        try {
          window.AIAgentStore?.updateAgent?.(currentAgentId, {
            defaultProviderId: currentProviderId,
            metadata: {
              connectionId: currentConnectionId || undefined
            }
          });
          if (currentConnectionId) {
            cachedProviders.set(currentConnectionId, currentProviderId);
          }
        } catch (error) {
          console.error('[RightActivityBar] Failed to persist connection selection:', error);
        }
      }
    }

    function debounce(fn, wait = 200) {
      let timer = null;
      return function debounced(...args) {
        if (timer) clearTimeout(timer);
        timer = setTimeout(() => fn.apply(this, args), wait);
      };
    }

    const persistAgentConfiguration = debounce(() => {
      if (!currentAgentId) return;
      const activeType = agentTypeGroup.find((input) => input.checked)?.value;
      const type = activeType || 'assistant';
      const description = agentDescriptionInput?.value ? agentDescriptionInput.value.trim() : '';
      const connectionId = agentConnectionSelect?.value || '';
      const connection = connectionId ? findConnection(connectionId) : null;
      const metadataPatch = {};
      metadataPatch.connectionId = connectionId || undefined;
      const patch = {
        type,
        description,
        metadata: metadataPatch,
        defaultProviderId: currentProviderId
      };
      if (connection?.providerId) {
        patch.defaultProviderId = connection.providerId;
      }
      const snapshot = window.AIAgentStore?.getSelectedAgent?.();
      if (snapshot?.agent) {
        const existing = snapshot.agent;
        const existingDescription = existing.description || '';
        const existingType = existing.type || 'assistant';
        const existingConnectionId = existing.metadata?.connectionId || null;
        const existingProvider = normalizeProvider(existing.defaultProviderId || 'openai');
        const nextProvider = normalizeProvider(patch.defaultProviderId || existingProvider);
        if (existingType === type
          && existingDescription === description
          && existingConnectionId === metadataPatch.connectionId
          && existingProvider === nextProvider) {
          return;
        }
      }
	      try {
	        window.AIAgentStore?.updateAgent?.(currentAgentId, patch);
	        // Also persist provider-specific model from the selected connection so
	        // chat uses the same model as AI Sputnik connection settings.
	        if (connection && connection.providerId && connection.primaryModelId) {
	          try {
	            window.AIAgentStore?.updateProviderState?.(currentAgentId, connection.providerId, {
	              model: connection.primaryModelId
	            });
	          } catch (e) {
	            console.error('[RightActivityBar] Failed to persist agent provider model state:', e);
	          }
	        }
	      } catch (error) {
	        console.error('[RightActivityBar] Failed to persist agent metadata:', error);
	      }
    }, 300);

    const availableAgentPanes = ['config', 'prompt'];
    function setAgentPane(pane) {
      const targetPane = availableAgentPanes.includes(pane) ? pane : 'config';
      agentPanes.forEach((node) => {
        const isActive = node.dataset.agentPane === targetPane;
        node.style.display = isActive ? '' : 'none';
        node.setAttribute('aria-hidden', isActive ? 'false' : 'true');
      });
      agentPaneButtons.forEach((button) => {
        const isActive = button.dataset.agentPaneBtn === targetPane;
        button.classList.toggle('settings-btn-primary', isActive);
        button.classList.toggle('settings-btn-secondary', !isActive);
        button.classList.toggle('is-active', isActive);
        button.setAttribute('aria-pressed', isActive ? 'true' : 'false');
      });
      try { localStorage.setItem('ai_agent_detail_pane', targetPane); } catch (_) {}
      if (targetPane === 'prompt') {
        setTimeout(() => {
          try { promptEditor?.layout?.(); } catch (_) {}
        }, 0);
      }
    }

    let initialAgentPane = 'config';
    try {
      const storedPane = localStorage.getItem('ai_agent_detail_pane');
      if (availableAgentPanes.includes(storedPane)) {
        initialAgentPane = storedPane;
      }
    } catch (_) {}
    setAgentPane(initialAgentPane);
    agentPaneButtons.forEach((button) => {
      button.addEventListener('click', () => {
        setAgentPane(button.dataset.agentPaneBtn || 'config');
      });
    });

    function fetchPromptText(providerId, force) {
      const id = providerId || currentProviderId;
      const key = cacheKey(id);
      if (!force && PROVIDER_TEXT_CACHE.has(key)) {
        const cached = PROVIDER_TEXT_CACHE.get(key);
        return typeof cached === 'string' ? cached : '';
      }
      if (currentAgentId) {
        try {
          const snapshot = window.AIAgentStore?.getSelectedAgent?.();
          const agent = snapshot?.agent;
          const providerState = agent?.providers?.[id];
          if (providerState && typeof providerState.prompt === 'string') {
            PROVIDER_TEXT_CACHE.set(key, providerState.prompt);
            return providerState.prompt;
          }
        } catch (_) {}
      }
      let value = '';
      try { value = window.AIPromptService?.loadPrompt?.(id) || ''; } catch (_) {}
      PROVIDER_TEXT_CACHE.set(key, value || '');
      return value || '';
    }

    function ensureModelForProvider(providerId, text) {
      if (typeof monaco === 'undefined' || !window.monaco || !monaco.editor) return null;
      const id = providerId || currentProviderId;
      const key = cacheKey(id);
      const existing = PROVIDER_MODEL_CACHE.get(key);
      if (existing && typeof existing.isDisposed === 'function' && existing.isDisposed()) {
        PROVIDER_MODEL_CACHE.delete(key);
      } else if (existing && existing.getValue) {
        if (typeof text === 'string' && existing.getValue() !== text) {
          existing.setValue(text);
        }
        return existing;
      }
      const initial = typeof text === 'string' ? text : fetchPromptText(id, false);
      const model = monaco.editor.createModel(initial, 'xml');
      PROVIDER_MODEL_CACHE.set(key, model);
      return model;
    }

    function loadProviderPrompt(providerId) {
      const id = providerId || currentProviderId;
      if (!id) return;
      if (promptEditor && typeof monaco !== 'undefined' && monaco.editor) {
        const model = ensureModelForProvider(id);
        if (model && promptEditor.getModel() !== model) {
          promptEditor.setModel(model);
        }
      } else if (window.__rightPromptTextArea) {
        window.__rightPromptTextArea.value = fetchPromptText(id, false);
      }
      updateCount();
    }

    promptEditor = ensurePromptEditor(currentProviderId);
    attachEditorListeners();
    loadProviderPrompt(currentProviderId);
    updateCount();

    function setAgentMetadataFields(agent) {
      const hasAgent = !!agent;
      if (agentDetailBody) {
        agentDetailBody.style.display = hasAgent ? 'flex' : 'none';
      }
      if (agentEmptyState) {
        agentEmptyState.style.display = hasAgent ? 'none' : 'flex';
      }
      [loadDefaultBtn, clearBtn, saveBtn].forEach((btn) => {
        if (btn) btn.disabled = !hasAgent;
      });
      if (AI_FREE_RESTRICTIONS_ENABLED) {
        [loadDefaultBtn, clearBtn, saveBtn].forEach((btn) => {
          if (btn) btn.disabled = true;
        });
      }
      if (promptEditor && typeof promptEditor.updateOptions === 'function') {
        promptEditor.updateOptions({ readOnly: AI_FREE_RESTRICTIONS_ENABLED ? true : !hasAgent });
      } else if (window.__rightPromptTextArea) {
        window.__rightPromptTextArea.disabled = AI_FREE_RESTRICTIONS_ENABLED ? true : !hasAgent;
      }

      suppressAgentFormEvents = true;
      const connectionId = hasAgent ? (agent?.metadata?.connectionId || '') : '';
      populateAgentConnections(connectionId);
      agentTypeGroup.forEach((input) => {
        input.disabled = !hasAgent;
        if (hasAgent) {
          const isMatch = (agent?.type || 'assistant') === input.value;
          input.checked = isMatch;
          input.setAttribute('aria-checked', isMatch ? 'true' : 'false');
        } else {
          const isDefault = input.value === 'assistant';
          input.checked = isDefault;
          input.setAttribute('aria-checked', isDefault ? 'true' : 'false');
        }
      });
      if (agentConnectionSelect) {
        agentConnectionSelect.disabled = !hasAgent;
        if (connectionId && connectionCatalog.some((connection) => connection.id === connectionId)) {
          agentConnectionSelect.value = connectionId;
        } else {
          agentConnectionSelect.value = '';
        }
      }
      if (agentDescriptionInput) {
        agentDescriptionInput.disabled = !hasAgent;
        agentDescriptionInput.value = hasAgent ? (agent?.description || '') : '';
      }
      suppressAgentFormEvents = false;

      if (!hasAgent) {
        currentConnectionId = null;
        updateConnectionHint(null, currentProviderId);
        setEditorValue(fetchPromptText(currentProviderId, false));
        return;
      }

      const fallbackProviderId = AI_FREE_PROVIDER_ID;
      const normalizedDefaultProvider = AI_FREE_PROVIDER_ID;
      applyConnectionSelection(connectionId, { persist: false, fallbackProviderId });
      const providerMismatch = hasAgent && normalizedDefaultProvider !== normalizeProvider(currentProviderId);
      if ((connectionId && !currentConnectionId) || providerMismatch) {
        persistAgentConfiguration();
      }
      setTimeout(() => {
        try { promptEditor?.layout?.(); } catch (_) {}
      }, 0);
    }

    function getSelectedAgent() {
      if (!window.AIAgentStore || typeof window.AIAgentStore.getSelectedAgent !== 'function') return null;
      try { return window.AIAgentStore.getSelectedAgent(); } catch (_) { return null; }
    }

    setAgentMetadataFields(null);

    function applyAgentSnapshot(snapshot) {
      const previousAgentId = currentAgentId;
      if (!snapshot) {
        currentAgentId = null;
        setAgentMetadataFields(null);
        return;
      }
      if (previousAgentId && previousAgentId !== snapshot.agent?.id) {
        persistCurrentProviderState();
      }
      currentAgentId = snapshot.agent?.id || null;
      const agent = snapshot.agent || null;
      setAgentMetadataFields(agent);
      loadProviderPrompt(currentProviderId);
      updateCount();
    }

    agentConnectionSelect?.addEventListener('change', (event) => {
      handleConnectionChange(event.target.value);
    });
    agentTypeGroup.forEach((input) => {
      input.addEventListener('change', () => {
        if (suppressAgentFormEvents) return;
        agentTypeGroup.forEach((radio) => {
          if (radio.checked) {
            radio.setAttribute('aria-checked', 'true');
          } else {
            radio.setAttribute('aria-checked', 'false');
          }
        });
        persistAgentConfiguration();
      });
    });
    agentDescriptionInput?.addEventListener('input', () => {
      if (suppressAgentFormEvents) return;
      persistAgentConfiguration();
    });

    agentConfigSaveBtn?.addEventListener('click', () => {
      persistAgentConfiguration();
      if (window.LogService && typeof window.LogService.logInfo === 'function') {
        try {
          window.LogService.logInfo('ai-agents', 'agent-config-saved', { agentId: currentAgentId });
        } catch (_) {}
      }
    });

	    if (!window.__aiConnectionsListenerBound) {
	      const handleConnectionsUpdated = (event) => {
	        const list = event?.detail?.connections;
	        suppressAgentFormEvents = true;
	        populateAgentConnections(currentConnectionId, Array.isArray(list) ? list : undefined);
	        suppressAgentFormEvents = false;
	        updateConnectionHint(currentConnectionId ? findConnection(currentConnectionId) : null, currentProviderId);
	      }
	      window.__aiConnectionsUpdatedListener = handleConnectionsUpdated
	      document.addEventListener('aiConnections:updated', handleConnectionsUpdated)
	      window.__aiConnectionsListenerBound = true;
	    }

    function handleAgentSelected(event) {
      const detail = event?.detail;
      const agent = detail?.agent || detail?.agent?.agent;
      applyAgentSnapshot(detail || agent ? { agent } : null);
    }

    const agentStore = window.AIAgentStore;
    let agentStoreUnsubscribe = window.__aiAgentStoreUnsubscribe || null;
    if (agentStoreUnsubscribe) {
      try { agentStoreUnsubscribe(); } catch (_) {}
      agentStoreUnsubscribe = null;
    }
    if (agentStore && typeof agentStore.on === 'function') {
      try {
        agentStoreUnsubscribe = agentStore.on(agentStore.events?.AGENT_SELECTED || 'ai-agent:selected', (evt) => {
          const detail = evt?.detail || evt;
          applyAgentSnapshot(detail);
        });
        const snapshot = agentStore.getSelectedAgent?.();
        if (snapshot) {
          applyAgentSnapshot(snapshot);
        }
      } catch (error) {
        console.error('[RightActivityBar] Failed to bind agent store listener:', error);
      }
    }
    window.__aiAgentStoreUnsubscribe = agentStoreUnsubscribe;

	    if (!window.__aiAgentSelectedListenerBound) {
	      window.__aiAgentSelectedListener = handleAgentSelected
	      document.addEventListener('ai:agent-selected', handleAgentSelected);
	      window.__aiAgentSelectedListenerBound = true;
	    }

    function getCurrentMacroIndex () {
      try {
        const state = window.pluginState?.getState?.()
        if (!state) return -1
        if (state.currentMode === 1 && typeof state.userMacros?.current === 'number' && state.userMacros.current >= 0) {
          return state.userMacros.current
        }
        return -1
      } catch (_) {
        return -1
      }
    }

    function getCurrentMacroName () {
      try {
        const idx = getCurrentMacroIndex()
        if (idx < 0) return null
        const macros = window.pluginState.getMacros?.() || []
        return macros[idx] && macros[idx].name ? macros[idx].name : null
      } catch (_) {
        return null
      }
    }

    let macroJobPollTimer = null;
    let macroJobId = null;
    const macroJobHistory = [];

    function updateMacroJobView (type, job) {
      if (!jobStatusEl || !jobOutputEl || !job) return
      const status = job.status || 'unknown'
      const phase = job.phase || null
      const pipeline = job.result && job.result.pipeline ? job.result.pipeline : type
      const phaseLabel = phase ? `${status}:${phase}` : status
      jobStatusEl.textContent = `Status: ${phaseLabel} (${pipeline})`
      let output = ''
      const result = job.result || {}
      if (type === 'refactor-macro' && result.refactorSuggestion) {
        output = result.refactorSuggestion
      } else if (type === 'explain-macro' && result.explanation) {
        output = result.explanation
      } else if (type === 'review-macro-quality') {
        const parts = []
        if (result.apiUsage) {
          const calls = Array.isArray(result.apiUsage.calls) ? result.apiUsage.calls.length : (result.apiUsage.callCount || 0)
          parts.push(`API usage calls: ${calls}`)
        }
        if (result.performance) {
          const score = typeof result.performance.score === 'number' ? result.performance.score : null
          if (score !== null) parts.push(`Performance score: ${score}`)
        }
        if (result.linter) {
          const issues = Array.isArray(result.linter.issues) ? result.linter.issues.length : (result.linter.issueCount || 0)
          parts.push(`Lint issues: ${issues}`)
        }
        output = parts.length ? parts.join('\n') : JSON.stringify(result, null, 2)
      } else if (result && Object.keys(result).length) {
        output = JSON.stringify(result, null, 2)
      } else {
        output = '(no result payload)'
      }
      jobOutputEl.textContent = output
      try {
        if (jobHistorySelect && job.id) {
          const existingIndex = macroJobHistory.findIndex((entry) => entry && entry.id === job.id)
          const snapshot = {
            id: job.id,
            type: job.type,
            pipeline,
            status: job.status || 'unknown',
            phase: job.phase || null,
            updatedAt: job.updatedAt || Date.now()
          }
          if (existingIndex >= 0) {
            macroJobHistory[existingIndex] = snapshot
          } else {
            macroJobHistory.unshift(snapshot)
          }
          while (macroJobHistory.length > 8) macroJobHistory.pop()
          const selectedId = jobHistorySelect.value
          jobHistorySelect.innerHTML = ''
          if (!macroJobHistory.length) {
            const opt = document.createElement('option')
            opt.value = ''
            opt.textContent = 'No jobs yet'
            jobHistorySelect.appendChild(opt)
          } else {
            macroJobHistory.forEach((entry) => {
              const opt = document.createElement('option')
              opt.value = entry.id
              const ts = new Date(entry.updatedAt || Date.now()).toLocaleTimeString()
              opt.textContent = `${entry.type || 'job'} – ${entry.status}${entry.phase ? ':' + entry.phase : ''} @ ${ts}`
              jobHistorySelect.appendChild(opt)
            })
          }
          const hasMatching = macroJobHistory.some((entry) => entry.id === selectedId)
          jobHistorySelect.value = hasMatching ? selectedId : (job.id || '')
        }
      } catch (_) {}
    }

    function pollMacroJob (type) {
      if (!macroJobId || !window.AgentJobs || typeof window.AgentJobs.getJob !== 'function') return
      try {
        const job = window.AgentJobs.getJob(macroJobId)
        if (job) {
          updateMacroJobView(type, job)
          if (job.status === 'completed' || job.status === 'failed') {
            if (job.status === 'failed') {
              const msg = job.error || `Agent job "${type}" failed.`
              try {
                window.alert(`AI job failed: ${msg}`)
              } catch (_) {}
            }
            macroJobPollTimer = null
            return
          }
        }
      } catch (_) {}
      macroJobPollTimer = setTimeout(() => pollMacroJob(type), 500)
    }

    function runMacroJob (type) {
      if (!window.AgentJobs || typeof window.AgentJobs.startJob !== 'function') {
        window.debug?.warn('RightActivityBar', 'AgentJobs not available for macro action', { type })
        try {
          window.alert('AI agent jobs are not available. Please reload the plugin or enable developer mode.')
        } catch (_) {}
        return
      }
      const macroName = getCurrentMacroName()
      if (!macroName) {
        window.alert('Select a macro in the tree before running this action.')
        return
      }
      try {
        const jobId = window.AgentJobs.startJob(type, { macroName })
        window.debug?.info('RightActivityBar', 'Agent job started', { type, jobId, macroName })
        macroJobId = jobId
        if (jobStatusEl) {
          jobStatusEl.textContent = `Status: running (${type})`
        }
        if (jobOutputEl) {
          jobOutputEl.textContent = ''
        }
        if (macroJobPollTimer) {
          clearTimeout(macroJobPollTimer)
          macroJobPollTimer = null
        }
        pollMacroJob(type)
      } catch (error) {
        window.debug?.error('RightActivityBar', 'Failed to start agent job', { type, error: error?.message || String(error) })
        try {
          window.alert(`Failed to start AI job: ${error?.message || String(error)}`)
        } catch (_) {}
      }
    }

    function runAgenticVbaConversionForCurrentMacro () {
      try {
        const state = window.pluginState?.getState?.()
        if (!state) {
          window.alert('Plugin state is not available. Please reload the plugin.')
          return
        }
        try {
          window.debug?.info?.('RightActivityBar', 'runAgenticVbaConversionForCurrentMacro invoked', {
            currentMode: state.currentMode,
            importedCurrent: state.importedMacros?.current,
            userMacroCurrent: state.userMacros?.current
          })
          console.log('[RightActivityBar] runAgenticVbaConversionForCurrentMacro', {
            currentMode: state.currentMode,
            importedCurrent: state.importedMacros?.current,
            userMacroCurrent: state.userMacros?.current
          })
        } catch (_) {}

        // External VBA modules (currentMode === 2) use the existing translateVBAToJS pipeline
        if (state.currentMode === 2) {
          const currentExternalIndex = typeof state.importedMacros?.current === 'number'
            ? state.importedMacros.current
            : -1
          if (currentExternalIndex < 0) {
            window.alert('Select a VBA module in the External VBA tree before converting.')
            return
          }
          if (typeof window.translateVBAToJS === 'function') {
            window.translateVBAToJS(currentExternalIndex)
            return
          }
          window.alert('VBA translation function is not available for external modules.')
          return
        }

        // Document macros (currentMode === 1) use MacroManager.convertVBAMacro (agentic path when enabled)
        if (!window.macroManager) {
          window.alert('VBA conversion is not available in this build.')
          return
        }
        const idx = getCurrentMacroIndex()
        if (idx < 0) {
          window.alert('Select a macro in the Macros tree before converting.')
          return
        }
        const fn = window.macroManager.convertVBAMacro
        if (typeof fn !== 'function') {
          window.alert('VBA conversion method is not available.')
          return
        }
        fn.call(window.macroManager, idx)
      } catch (error) {
        window.debug?.error('RightActivityBar', 'Failed to start agentic VBA conversion', { error: error?.message || String(error) })
        window.alert('Failed to start VBA conversion. Check the console for details.')
      }
    }

    if (btnConvertVba) {
      btnConvertVba.addEventListener('click', () => {
        try {
          window.debug?.info?.('RightActivityBar', 'Convert VBA to JS macro button clicked', { tab: currentAgentsTab });
        } catch (_) {}
        runAgenticVbaConversionForCurrentMacro();
      });
    }
    btnRefactorMacro?.addEventListener('click', () => runMacroJob('refactor-macro'));
    btnExplainMacro?.addEventListener('click', () => runMacroJob('explain-macro'));
    btnReviewMacro?.addEventListener('click', () => runMacroJob('review-macro-quality'));
    const btnSearchMacros = document.getElementById('ai-agent-action-search');
    if (btnSearchMacros) {
      btnSearchMacros.addEventListener('click', () => {
        try {
          const query = window.prompt('Search macros (text query for agentic search):', '');
          if (!query || !query.trim()) {
            return;
          }
          if (!window.AgentJobs || typeof window.AgentJobs.startJob !== 'function') {
            window.alert('Agent job manager is not available in this build.');
            return;
          }
          window.debug?.info?.('RightActivityBar', 'macro-search agent job requested', { queryLength: query.length });
          const jobId = window.AgentJobs.startJob('macro-search', {
            query: query.trim(),
            options: { limit: 50 }
          });
          macroJobId = jobId;
          pollMacroJob('macro-search');
        } catch (error) {
          window.debug?.error('RightActivityBar', 'Failed to start macro-search agent job', { error: error?.message || String(error) });
        }
      });
    }

    // Fallback: if the expected action buttons are present visually but do not
    // have the canonical IDs (or the IDs were changed by a different build),
    // use event delegation on the Agentic Macros section and infer actions
    // from the button labels. This keeps the dev UX working even when markup
    // diverges slightly.
    const hasDirectAgentButtons = !!(btnConvertVba || btnRefactorMacro || btnExplainMacro || btnReviewMacro || btnSearchMacros);
    if (!hasDirectAgentButtons && container) {
      const actionsSection = container.querySelector('.ai-agent-actions-section');
      if (actionsSection) {
        actionsSection.addEventListener('click', (event) => {
          const target = event.target;
          if (!target || target.tagName !== 'BUTTON') return;
          const label = (target.textContent || '').toLowerCase();
          if (label.includes('convert') && label.includes('vba')) {
            runAgenticVbaConversionForCurrentMacro();
            return;
          }
          if (label.includes('refactor')) {
            runMacroJob('refactor-macro');
            return;
          }
          if (label.includes('explain')) {
            runMacroJob('explain-macro');
            return;
          }
          if (label.includes('review') || label.includes('quality')) {
            runMacroJob('review-macro-quality');
            return;
          }
          if (label.includes('search')) {
            try {
              const query = window.prompt('Search macros (text query for agentic search):', '');
              if (!query || !query.trim()) {
                return;
              }
              if (!window.AgentJobs || typeof window.AgentJobs.startJob !== 'function') {
                window.alert('Agent job manager is not available in this build.');
                return;
              }
              window.debug?.info?.('RightActivityBar', 'macro-search agent job requested (delegated)', { queryLength: query.length });
              const jobId = window.AgentJobs.startJob('macro-search', {
                query: query.trim(),
                options: { limit: 50 }
              });
              macroJobId = jobId;
              pollMacroJob('macro-search');
            } catch (error) {
              window.debug?.error('RightActivityBar', 'Failed to start macro-search agent job (delegated)', { error: error?.message || String(error) });
            }
          }
        });
      }
    }

    jobHistorySelect?.addEventListener('change', () => {
      const id = jobHistorySelect.value
      if (!id) return
      try {
        const job = window.AgentJobs && typeof window.AgentJobs.getJob === 'function'
          ? window.AgentJobs.getJob(id)
          : null
        if (!job) return
        // Reuse current type to render; when selected from history, type may not match current buttons but pipeline will.
        const inferredType = job.type || 'job'
        macroJobId = job.id
        updateMacroJobView(inferredType, job)
      } catch (_) {}
    })

    jobDetailsCopyBtn?.addEventListener('click', () => {
      if (!macroJobId || !window.AgentJobs || typeof window.AgentJobs.getJob !== 'function') {
        window.alert('No agent job selected to copy.')
        return
      }
      try {
        const job = window.AgentJobs.getJob(macroJobId)
        if (!job) {
          window.alert('Selected job is no longer available.')
          return
        }
        const payload = {
          id: job.id,
          type: job.type,
          status: job.status,
          phase: job.phase,
          pipeline: job.result && job.result.pipeline,
          error: job.error || null,
          result: job.result || null,
          logs: job.logs || []
        }
        const text = JSON.stringify(payload, null, 2)
        if (navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(text).then(
            () => { window.alert('Agent job details copied to clipboard.') },
            () => { window.alert('Failed to copy job details. You can copy from the console instead.') }
          )
        } else {
          // Fallback: show in prompt so user can copy manually
          // eslint-disable-next-line no-alert
          window.alert(text)
        }
      } catch (error) {
        window.debug?.error('RightActivityBar', 'Failed to copy job details', { error: error?.message || String(error) })
        window.alert('Failed to copy job details. Check the console for details.')
      }
    })

    loadDefaultBtn?.addEventListener('click', () => {
      let value = '';
      try { value = window.AIPromptService?.getDefaultPrompt?.(currentProviderId) || ''; } catch (_) {}
      if (!value) {
        try { value = window.AIInstructionManager?.getRawTemplate?.('vba-conversion') || ''; } catch (_) {}
      }
      setEditorValue(value);
      updateCount();
      persistCurrentProviderState();
    });

    clearBtn?.addEventListener('click', () => {
      setEditorValue('');
      updateCount();
      try {
        if (window.AIPromptService?.resetPrompt) {
          window.AIPromptService.resetPrompt(currentProviderId);
        } else {
          legacyReset();
        }
      } catch (error) {
        legacyReset();
        window.LogService?.logWarning?.('AI Prompt', 'Prompt reset failed', { providerId: currentProviderId, message: error?.message });
      }
      persistCurrentProviderState();
    });

    saveBtn?.addEventListener('click', () => {
      const text = (getEditorValue() || '').trim();
      try {
        if (window.AIPromptService?.savePrompt) {
          window.AIPromptService.savePrompt(currentProviderId, text);
        } else {
          legacySave(text);
        }
        window.LogService?.logInfo?.('AI Prompt', 'Prompt saved', { providerId: currentProviderId, length: text.length });
      } catch (error) {
        legacySave(text);
        window.LogService?.logWarning?.('AI Prompt', 'Prompt save failed', { providerId: currentProviderId, message: error?.message });
      }
      updateCount();
      persistCurrentProviderState();
    });

    function legacySave(text) {
      let ok = false;
      try { window.settingsManager?.setTranslationPrompt?.(text); ok = true; } catch (_) {}
      if (!ok && SETTINGS_STORE?.update) {
        try { SETTINGS_STORE.update({ translationPrompt: text || undefined }); ok = true; } catch (_) {}
      }
      if (!ok) {
        try {
          const raw = localStorage.getItem('macros_plugin_settings');
          const obj = raw ? JSON.parse(raw) : {};
          if (text) obj.translationPrompt = text;
          else delete obj.translationPrompt;
          localStorage.setItem('macros_plugin_settings', JSON.stringify(obj));
        } catch (_) {}
      }
    }

    function legacyReset() {
      let cleared = false;
      try { window.settingsManager?.resetTranslationPrompt?.(); cleared = true; } catch (_) {}
      if (!cleared && SETTINGS_STORE?.update) {
        try { SETTINGS_STORE.update({ translationPrompt: undefined }); cleared = true; } catch (_) {}
      }
      if (!cleared) {
        try {
          const raw = localStorage.getItem('macros_plugin_settings');
          const obj = raw ? JSON.parse(raw) : {};
          delete obj.translationPrompt;
          localStorage.setItem('macros_plugin_settings', JSON.stringify(obj));
        } catch (_) {}
      }
    }

    function getProviderCatalog() {
      if (AI_FREE_RESTRICTIONS_ENABLED) {
        return [{ id: AI_FREE_PROVIDER_ID, label: 'OpenAI', enabled: true }];
      }
      if (window.AIPromptService?.getProviderCatalog) {
        try {
          const catalog = window.AIPromptService.getProviderCatalog();
          if (Array.isArray(catalog) && catalog.length) {
            return catalog;
          }
        } catch (_) {}
      }
      return [
        { id: 'openai', label: 'OpenAI', enabled: true },
        { id: 'github', label: 'GitHub Models', enabled: true },
        { id: 'gemini', label: 'Google Gemini', enabled: true }
      ];
    }

    function resolveInitialProvider(catalog) {
      const available = Array.isArray(catalog) && catalog.length ? catalog.map((item) => item.id) : ['openai'];
      const candidates = [];
      if (window.__aiPromptProvider) candidates.push(window.__aiPromptProvider);
      try {
        const stored = localStorage.getItem('ai_prompt_provider');
        if (stored) candidates.push(stored);
      } catch (_) {}
      if (window.AIPromptService?.getDefaultProviderId) {
        try { candidates.push(window.AIPromptService.getDefaultProviderId()); } catch (_) {}
      }
      candidates.push('openai');
      for (let i = 0; i < candidates.length; i += 1) {
        const normalized = normalizeProvider(candidates[i]);
        if (available.includes(normalized)) return normalized;
      }
      return available[0];
    }

    function handleConnectionChange(connectionId) {
      if (suppressAgentFormEvents) return;
      getConnectionCatalog();
      applyConnectionSelection(connectionId || '', { persist: true });
      persistAgentConfiguration();
      try {
        document.dispatchEvent(new CustomEvent('rightPane:providerChanged', { detail: { providerId: currentProviderId, connectionId: currentConnectionId } }));
      } catch (_) {}
      if (window.LogService && typeof window.LogService.logInfo === 'function') {
        try {
          window.LogService.logInfo('ai-agents', 'connection-switched', {
            connectionId: currentConnectionId,
            providerId: currentProviderId,
            agentId: currentAgentId
          });
        } catch (_) {}
      }
    }

    function normalizeProvider(value) {
      if (window.AIPromptService?.normalizeProviderId) {
        try { return window.AIPromptService.normalizeProviderId(value); } catch (_) {}
      }
      return (typeof value === 'string' && value.trim()) ? value.trim().toLowerCase() : 'openai';
    }

    function ensurePromptEditor(providerId) {
      const node = document.getElementById('ai-right-prompt-editor');
      if (!node) return null;
      if (promptEditor && promptEditor.getDomNode()) {
        const domNode = promptEditor.getDomNode();
        if (domNode && domNode.parentNode !== node) {
          try { node.appendChild(domNode); } catch (_) {}
        }
        ensureXmlLanguage();
        setTimeout(() => {
          try { promptEditor.layout(); } catch (_) {}
        }, 0);
        return promptEditor;
      }
      if (typeof monaco === 'undefined' || !window.monaco || !monaco.editor) {
        const ta = document.createElement('textarea');
        ta.style.cssText = 'width:100%;height:100%;background:var(--vscode-input-background,#1e1e1e);color:var(--vscode-input-foreground,#d4d4d4);border:0;';
        node.appendChild(ta);
        window.__rightPromptTextArea = ta;
        window.__rightPromptEditor = null;
        promptEditor = null;
        return null;
      }
      promptEditor = monaco.editor.create(node, {
        value: '',
        language: 'plaintext',
        minimap: { enabled: false },
        wordWrap: 'on',
        automaticLayout: true,
        theme: (document.body.classList.contains('theme-type-dark') ? 'vs-dark' : 'vs')
      });
      window.__rightPromptEditor = promptEditor;
      ensureXmlLanguage();
      return promptEditor;
    }

    function ensureXmlLanguage() {
      if (!promptEditor) return;
      try {
        tryLoadXML().then(() => {
          try { monaco.editor.setModelLanguage(promptEditor.getModel(), 'xml'); } catch (_) {}
        });
      } catch (_) {}
    }

    function tryLoadXML() {
      return new Promise((resolve) => {
        try {
          if (typeof monaco !== 'undefined' && window.monaco && monaco.languages) {
            const exists = monaco.languages.getLanguages?.().some((l) => l.id === 'xml');
            if (exists) return resolve(true);
            try {
              monaco.languages.register({ id: 'xml' });
              monaco.languages.setMonarchTokensProvider('xml', {
                tokenizer: {
                  root: [
                    [/\<\!--[\s\S]*?--\>/, 'comment'],
                    [/\<\!\[CDATA\[[\s\S]*?\]\]\>/, 'string'],
                    [/\<\/?[A-Za-z_][\w\-\.]*\b/, 'tag'],
                    [/\s+[A-Za-z_:][-A-Za-z0-9_:.]*\s*=\s*"[^"]*"/, 'attribute'],
                    [/\s+[A-Za-z_:][-A-Za-z0-9_:.]*\s*=\s*'[^']*'/, 'attribute'],
                    [/\s+[A-Za-z_:][-A-Za-z0-9_:.]*/, 'attribute.name'],
                    [/\/?\>/, 'tag']
                  ]
                }
              });
              monaco.languages.setLanguageConfiguration('xml', {
                brackets: [['<', '>']],
                autoClosingPairs: [{ open: '<', close: '>' }, { open: '"', close: '"' }, { open: "'", close: "'" }]
              });
              return resolve(true);
            } catch (_) {}
          }
        } catch (_) {}
        resolve(false);
      });
    }

    function attachEditorListeners() {
      if (window.__rightPromptEditorDisposer && typeof window.__rightPromptEditorDisposer.dispose === 'function') {
        try { window.__rightPromptEditorDisposer.dispose(); } catch (_) {}
      }
      if (promptEditor && typeof promptEditor.onDidChangeModelContent === 'function') {
        window.__rightPromptEditorDisposer = promptEditor.onDidChangeModelContent(() => updateCount());
      } else {
        window.__rightPromptEditorDisposer = null;
      }
      if (window.__rightPromptTextArea && window.__rightPromptTextAreaListener) {
        window.__rightPromptTextArea.removeEventListener('input', window.__rightPromptTextAreaListener);
        window.__rightPromptTextAreaListener = null;
      }
      if (!promptEditor && window.__rightPromptTextArea) {
        const handler = () => updateCount();
        window.__rightPromptTextArea.addEventListener('input', handler);
        window.__rightPromptTextAreaListener = handler;
      }
    }

    function setEditorValue(text) {
      const value = typeof text === 'string' ? text : '';
      if (promptEditor && typeof monaco !== 'undefined' && monaco.editor) {
        const model = ensureModelForProvider(currentProviderId, value);
        if (model && promptEditor.getModel() !== model) {
          promptEditor.setModel(model);
        }
        if (model && model.getValue && model.getValue() !== value) {
          model.setValue(value);
        }
      } else if (window.__rightPromptTextArea) {
        window.__rightPromptTextArea.value = value;
      }
      PROVIDER_TEXT_CACHE.set(cacheKey(currentProviderId), value);
      updateCount();
    }

    function getEditorValue() {
      if (promptEditor && promptEditor.getValue) return promptEditor.getValue();
      if (window.__rightPromptTextArea) return window.__rightPromptTextArea.value;
      return '';
    }

    function updateCount() {
      if (!counter) return;
      const value = getEditorValue() || '';
      counter.textContent = `${value.length} / 10,000 characters`;
      if (currentProviderId) {
      PROVIDER_TEXT_CACHE.set(cacheKey(currentProviderId), value);
      }
    }

    function persistCurrentProviderState() {
      if (!currentProviderId) return;
      const value = getEditorValue() || '';
      PROVIDER_TEXT_CACHE.set(cacheKey(currentProviderId), value);
      if (currentAgentId) {
        try {
          window.AIAgentStore?.updateProviderState?.(currentAgentId, currentProviderId, {
            prompt: value
          });
        } catch (error) {
          console.error('[RightActivityBar] Failed to persist provider state:', error);
        }
      }
      if (promptEditor && promptEditor.getModel && promptEditor.getModel()) {
        const key = cacheKey(currentProviderId);
        const model = promptEditor.getModel();
        if (model && typeof model.isDisposed === 'function' && model.isDisposed()) {
          PROVIDER_MODEL_CACHE.delete(key);
        } else if (model) {
          PROVIDER_MODEL_CACHE.set(key, model);
        }
      }
    }

    async function ensureAgentTreeMounted() {
      if (!agentTreeContainer) return;
      if (agentTreeInstance) return;
      const api = window.AgentTreeView;
      if (!api || typeof api.createAgentTreeView !== 'function') {
        agentTreeContainer.innerHTML = '<div class="tree-empty-message">AI Agents view unavailable.</div>';
        return;
      }
      agentTreeContainer.innerHTML = '<div class="tree-empty-message">Loading agents…</div>';
      try {
        if (typeof api.waitForStore === 'function') {
          await api.waitForStore();
        }
        agentTreeInstance = api.createAgentTreeView(agentTreeContainer);
      } catch (error) {
        console.error('[RightActivityBar] Failed to initialize agent tree:', error);
        agentTreeContainer.innerHTML = '<div class="tree-empty-message">Failed to load agents.</div>';
      }
    }

    function activate(tab) {
      if (!visibleSet.has(tab)) {
        const fallback = visibleTabs[0] || 'connection';
        if (!visibleSet.has(fallback) || tab === fallback) return;
        return activate(fallback);
      }
      if (tab === 'connection') {
        persistCurrentProviderState();
        if (tabConn) tabConn.style.display = '';
        if (tabPrompt) tabPrompt.style.display = 'none';
        if (tabAgents) tabAgents.style.display = 'none';
        if (agentTreeInstance && typeof agentTreeInstance.dispose === 'function') {
          try { agentTreeInstance.dispose(); } catch (_) {}
          agentTreeInstance = null;
        }
      } else if (tab === 'agents') {
        persistCurrentProviderState();
        if (tabConn) tabConn.style.display = 'none';
        if (tabPrompt) tabPrompt.style.display = 'none';
        if (tabAgents) tabAgents.style.display = '';
        const mountPromise = ensureAgentTreeMounted();
        if (mountPromise && typeof mountPromise.catch === 'function') {
          mountPromise.catch(() => {});
        }
        promptEditor = ensurePromptEditor(currentProviderId);
        attachEditorListeners();
        loadProviderPrompt(currentProviderId);
        const snapshot = agentStore?.getSelectedAgent?.();
        if (snapshot) {
          applyAgentSnapshot(snapshot);
        } else {
          setAgentMetadataFields(null);
        }
        setTimeout(() => { try { promptEditor?.layout?.(); } catch (_) {} }, 0);
      } else {
        return;
      }
      currentAgentsTab = tab;
      btnConn?.classList.toggle('settings-btn-primary', tab === 'connection');
      btnAgents?.classList.toggle('settings-btn-primary', tab === 'agents');
      updateAgentsButtonHighlight();
      try { localStorage.setItem('right_pane_ai_tab', tab); } catch (_) {}
    }

    activate(initialTab);
    if (visibleSet.has('connection')) {
      btnConn?.addEventListener('click', () => activate('connection'));
    }
    if (visibleSet.has('agents')) {
      btnAgents?.addEventListener('click', () => activate('agents'));
    }

    // Bind + populate (Connection)
    try {
      if (typeof window.__aiConnectionManagerCleanup === 'function') {
        try { window.__aiConnectionManagerCleanup(); } catch (_) {}
        window.__aiConnectionManagerCleanup = null;
      }
      const cleanup = mountConnectionManager({
        elements: {
          treeEl: document.getElementById('ai-connection-tree'),
          inspectorEl: document.getElementById('ai-connection-inspector'),
          filterEl: document.getElementById('ai-connection-provider-filter'),
          searchEl: document.getElementById('ai-connection-search'),
          addButton: document.getElementById('ai-connection-add'),
          contentEl: document.querySelector('.ai-connection-content'),
          splitterEl: document.getElementById('ai-connection-splitter')
        },
        settingsStore: SETTINGS_STORE
      });
      window.__aiConnectionManagerCleanup = cleanup;
    } catch (e) {
      console.warn('[AI Pane] Failed to initialize connection manager:', e);
    }
  }

async function renderChat() {
    const container = document.getElementById('right-pane-content');
    if (!container) return;
    disposeCodeGraph();

    const providerFactory = window.AIProviderFactory || null;
    const CHAT_IDB_DB_NAME = 'r7c_ai_chat';
    const CHAT_IDB_STORE_NAME = 'chat_state';
    const CHAT_IDB_STATE_KEY = 'main';
    const CHAT_IDB_VERSION = 1;

    const isObject = (value) => value !== null && typeof value === 'object' && !Array.isArray(value);

    function createEmptySelectionState() {
      return {
        connectionId: null,
        sessionByConnection: {}
      };
    }

    function createSessionId() {
      return `chat-${Math.random().toString(36).slice(2, 8)}${Date.now().toString(36)}`;
    }

    function sanitizeMessages(messages) {
      if (!Array.isArray(messages)) return [];
      return messages.map((entry) => {
        if (!isObject(entry)) return null;
        const role = typeof entry.role === 'string' ? entry.role : 'assistant';
        if (!['user', 'assistant', 'system'].includes(role)) return null;
        const content = typeof entry.content === 'string' ? entry.content : String(entry.content || '');
        return { role, content };
      }).filter(Boolean);
    }

    function sanitizeConversationStore(raw) {
      if (!isObject(raw)) return {};
      const out = {};
      Object.entries(raw).forEach(([key, bucket]) => {
        if (!isObject(bucket)) return;
        const sessions = Array.isArray(bucket.sessions) ? bucket.sessions : [];
        const normalizedSessions = sessions.map((session) => {
          if (!isObject(session)) return null;
          const id = typeof session.id === 'string' && session.id.trim() ? session.id.trim() : createSessionId();
          const title = typeof session.title === 'string' && session.title.trim() ? session.title : 'New chat';
          const createdAt = Number.isFinite(session.createdAt) ? session.createdAt : Date.now();
          const updatedAt = Number.isFinite(session.updatedAt) ? session.updatedAt : createdAt;
          const messages = sanitizeMessages(session.messages);
          return { id, title, createdAt, updatedAt, messages };
        }).filter(Boolean);
        out[String(key)] = { sessions: normalizedSessions };
      });
      return out;
    }

    function sanitizeSelectionState(raw) {
      const safe = createEmptySelectionState();
      if (!isObject(raw)) return safe;
      if (typeof raw.connectionId === 'string' && raw.connectionId.trim()) {
        safe.connectionId = raw.connectionId;
      }
      if (isObject(raw.sessionByConnection)) {
        Object.entries(raw.sessionByConnection).forEach(([key, value]) => {
          if (typeof value === 'string' && value.trim()) {
            safe.sessionByConnection[String(key)] = value;
          }
        });
      }
      return safe;
    }

    function supportsIndexedDb() {
      return typeof window !== 'undefined' && typeof window.indexedDB !== 'undefined';
    }

    function openChatDb() {
      return new Promise((resolve, reject) => {
        if (!supportsIndexedDb()) {
          reject(new Error('IndexedDB is not supported'));
          return;
        }
        let request;
        try {
          request = window.indexedDB.open(CHAT_IDB_DB_NAME, CHAT_IDB_VERSION);
        } catch (error) {
          reject(error);
          return;
        }
        request.onerror = () => reject(request.error || new Error('Failed to open IndexedDB'));
        request.onupgradeneeded = () => {
          const db = request.result;
          if (!db.objectStoreNames.contains(CHAT_IDB_STORE_NAME)) {
            db.createObjectStore(CHAT_IDB_STORE_NAME);
          }
        };
        request.onsuccess = () => resolve(request.result);
      });
    }

    async function loadChatStateFromIndexedDb() {
      if (!supportsIndexedDb()) return null;
      let db = null;
      try {
        db = await openChatDb();
        const payload = await new Promise((resolve, reject) => {
          const tx = db.transaction(CHAT_IDB_STORE_NAME, 'readonly');
          const store = tx.objectStore(CHAT_IDB_STORE_NAME);
          const request = store.get(CHAT_IDB_STATE_KEY);
          request.onerror = () => reject(request.error || new Error('Failed to load chat state'));
          request.onsuccess = () => resolve(request.result || null);
        });
        if (!isObject(payload)) return null;
        return {
          conversationStore: sanitizeConversationStore(payload.conversationStore),
          chatSelectionState: sanitizeSelectionState(payload.chatSelectionState)
        };
      } catch (_) {
        return null;
      } finally {
        try { db?.close?.(); } catch (_) {}
      }
    }

    async function persistChatStateToIndexedDb(conversationStoreRef, chatSelectionStateRef) {
      if (!supportsIndexedDb()) return;
      let db = null;
      try {
        db = await openChatDb();
        const payload = {
          version: CHAT_IDB_VERSION,
          updatedAt: Date.now(),
          conversationStore: sanitizeConversationStore(conversationStoreRef),
          chatSelectionState: sanitizeSelectionState(chatSelectionStateRef)
        };
        await new Promise((resolve, reject) => {
          const tx = db.transaction(CHAT_IDB_STORE_NAME, 'readwrite');
          tx.oncomplete = () => resolve();
          tx.onerror = () => reject(tx.error || new Error('Failed to persist chat state'));
          tx.objectStore(CHAT_IDB_STORE_NAME).put(payload, CHAT_IDB_STATE_KEY);
        });
      } catch (_) {
      } finally {
        try { db?.close?.(); } catch (_) {}
      }
    }

    const inMemoryConversationStore = sanitizeConversationStore(window.__aiChatConversationsByConnection || {});
    const inMemorySelectionState = sanitizeSelectionState(window.__aiChatSelectionState || createEmptySelectionState());
    const persistedChatState = await loadChatStateFromIndexedDb();

    const hasPersistedData = !!(persistedChatState
      && isObject(persistedChatState.conversationStore)
      && Object.keys(persistedChatState.conversationStore).length);

    const conversationStore = hasPersistedData
      ? persistedChatState.conversationStore
      : inMemoryConversationStore;
    const chatSelectionState = hasPersistedData
      ? persistedChatState.chatSelectionState
      : inMemorySelectionState;

    window.__aiChatConversationsByConnection = conversationStore;
    window.__aiChatSelectionState = chatSelectionState;

    let persistTimer = null;
    const scheduleChatStatePersist = (delay = 180) => {
      try {
        if (persistTimer) clearTimeout(persistTimer);
      } catch (_) {}
      persistTimer = setTimeout(() => {
        persistTimer = null;
        persistChatStateToIndexedDb(conversationStore, chatSelectionState);
      }, Math.max(0, delay));
    };

    function listConnections() {
      try {
        const ensured = window.SettingsStore?.ensureAiConnections?.();
        if (Array.isArray(ensured)) return ensured;
      } catch (_) {}
      try {
        const list = window.SettingsStore?.getAiConnections?.();
        if (Array.isArray(list)) return list;
      } catch (_) {}
      return [];
    }

    function findConnectionById(connections, connectionId) {
      if (!connectionId) return null;
      return (connections || []).find((entry) => entry && entry.id === connectionId) || null;
    }

    let connectionCatalog = listConnections();
    let selectedConnectionId = findConnectionById(connectionCatalog, chatSelectionState.connectionId)
      ? chatSelectionState.connectionId
      : (connectionCatalog[0]?.id || null);
    chatSelectionState.connectionId = selectedConnectionId;
    let currentChatContext = null;

    function getConversationBucket(connectionId) {
      const key = connectionId || '__no_connection__';
      if (!conversationStore[key]) {
        conversationStore[key] = {
          sessions: []
        };
      }
      return conversationStore[key];
    }

    function createSession(connectionId, title = 'New chat') {
      const bucket = getConversationBucket(connectionId);
      const id = createSessionId();
      const now = Date.now();
      const session = {
        id,
        title,
        createdAt: now,
        updatedAt: now,
        messages: []
      };
      bucket.sessions.unshift(session);
      scheduleChatStatePersist();
      return session;
    }

    function ensureConversation(connectionId, sessionId) {
      const bucket = getConversationBucket(connectionId);
      let session = bucket.sessions.find((item) => item.id === sessionId) || null;
      if (!session) {
        if (bucket.sessions.length) {
          session = bucket.sessions[0];
        } else {
          session = createSession(connectionId);
        }
      }
      return session;
    }

    function getSelectedSessionId(connectionId) {
      const key = connectionId || '__no_connection__';
      return chatSelectionState.sessionByConnection[key] || null;
    }

    function setSelectedSessionId(connectionId, sessionId) {
      const key = connectionId || '__no_connection__';
      chatSelectionState.sessionByConnection[key] = sessionId;
      scheduleChatStatePersist();
    }

    function loadConnectionSecret(connectionId) {
      if (!connectionId) return '';
      try {
        return localStorage.getItem(`${CONNECTION_SECRET_PREFIX}${connectionId}`) || '';
      } catch (_) {
        return '';
      }
    }

    let selectedSessionId = getSelectedSessionId(selectedConnectionId);
    let chatState = ensureConversation(selectedConnectionId, selectedSessionId);
    selectedSessionId = chatState.id;
    setSelectedSessionId(selectedConnectionId, selectedSessionId);

    container.innerHTML = `
      <div class="chat-pane">
        <div class="chat-header" style="display:flex;justify-content:space-between;align-items:center;gap:8px;">
          <div style="display:flex;flex-direction:column;">
            <span class="title-text" style="font-weight:600;">Chat</span>
            <span class="subtitle" id="chat-provider-label"></span>
          </div>
          <div class="chat-controls" style="display:flex;gap:6px;align-items:center;min-width:0;">
            <select id="chat-connection-select" class="vscode-select" style="width:180px;max-width:180px;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;"></select>
            <select id="chat-session-select" class="vscode-select" style="width:180px;max-width:180px;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;"></select>
            <button id="chat-new" class="settings-btn settings-btn-secondary" title="New chat">＋ New</button>
            <button id="chat-export" class="settings-btn settings-btn-secondary" title="Export conversation">🗒 Export</button>
          </div>
        </div>
        <div id="chat-messages" class="chat-messages"></div>
        <div class="chat-composer">
          <div class="chat-input-shell">
            <textarea id="chat-input" class="vscode-textarea chat-input" placeholder="Type a message..."></textarea>
            <div class="chat-actions-inside">
              <button id="chat-send" class="settings-btn settings-btn-primary chat-action-btn" title="Send" aria-label="Send">
                <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor" aria-hidden="true"><path d="M3.4 20.6 22 12 3.4 3.4l.1 6.2 11.4 2.4-11.4 2.4-.1 6.2Z"/></svg>
              </button>
              <button id="chat-stop" class="settings-btn settings-btn-secondary chat-action-btn hidden" title="Stop" aria-label="Stop" disabled>
                <svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor" aria-hidden="true"><rect x="6" y="6" width="12" height="12" rx="2"/></svg>
              </button>
            </div>
          </div>
        </div>
      </div>`;

    const messagesEl = document.getElementById('chat-messages');
    const inputEl = document.getElementById('chat-input');
    const sendBtn = document.getElementById('chat-send');
    const stopBtn = document.getElementById('chat-stop');
    const exportBtn = document.getElementById('chat-export');
    const newChatBtn = document.getElementById('chat-new');
    const providerLabelEl = document.getElementById('chat-provider-label');
    const connectionSelect = document.getElementById('chat-connection-select');
    const sessionSelect = document.getElementById('chat-session-select');
    try { if (connectionSelect) { const host = container.querySelector('.chat-header') || container; __attachCustomPicker(connectionSelect, host) } } catch (_) {}
    try { if (sessionSelect) { const host = container.querySelector('.chat-header') || container; __attachCustomPicker(sessionSelect, host) } } catch (_) {}

    function escapeHtml(str) {
      return String(str).replace(/[&<>]/g, (s) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[s]));
    }

    function truncateSelectLabel(value, maxChars = 22) {
      const text = typeof value === 'string' ? value : String(value || '');
      if (text.length <= maxChars) return text;
      return `${text.slice(0, Math.max(1, maxChars - 1))}…`;
    }

    function fallbackChatContext() {
      const promptService = window.AIPromptService;
      const defaultId = AI_FREE_PROVIDER_ID;
      let prompt = '';
      try {
        prompt = window.AIInstructionManager?.getRawTemplate?.('vba-conversion') || '';
      } catch (_) {}
      if (!prompt) {
        try { prompt = promptService?.getDefaultPrompt?.(defaultId) || ''; } catch (_) {}
      }
      const label = providerFactory?.getProviderLabel
        ? providerFactory.getProviderLabel(defaultId)
        : defaultId.replace(/[-_]+/g, ' ').replace(/\b\w/g, (c) => c.toUpperCase());
      return {
        providerId: defaultId,
        providerLabel: label,
        model: window.AIConfiguration?.getDefaultModel?.(defaultId) || '',
        prompt,
        connectionId: null,
        connectionName: 'No connection selected'
      };
    }

    function resolveChatContext(connection) {
      if (!connection) {
        return fallbackChatContext();
      }
      const providerId = connection.providerId || AI_FREE_PROVIDER_ID;
      const providerLabel = providerFactory?.getProviderLabel
        ? providerFactory.getProviderLabel(providerId)
        : providerId;
      let prompt = '';
      try {
        prompt = window.AIPromptService?.loadPrompt?.(providerId) || '';
      } catch (_) {}
      if (!prompt) {
        try {
          prompt = window.AIInstructionManager?.getRawTemplate?.('vba-conversion') || '';
        } catch (_) {}
      }

      const base = {
        providerId,
        providerLabel,
        model: connection.primaryModelId || window.AIConfiguration?.getDefaultModel?.(providerId) || '',
        prompt,
        connectionId: connection.id,
        connectionName: connection.name || connection.id,
        advancedLogging: false
      };

      const meta = (connection.metadata && typeof connection.metadata === 'object' && !Array.isArray(connection.metadata))
        ? connection.metadata
        : {};
      base.advancedLogging = meta.advancedLogging === true;
      const cfg = meta.requestConfig && typeof meta.requestConfig === 'object' && !Array.isArray(meta.requestConfig)
        ? meta.requestConfig
        : null;
      if (cfg) {
        base.requestOptions = {};
        if (typeof cfg.temperature === 'number' && Number.isFinite(cfg.temperature)) {
          base.requestOptions.temperature = cfg.temperature;
        }
        if (typeof cfg.maxTokens === 'number' && Number.isFinite(cfg.maxTokens) && cfg.maxTokens > 0) {
          base.requestOptions.max_tokens = cfg.maxTokens;
        }
      }
      return base;
    }

    function renderMessages() {
      messagesEl.innerHTML = (chatState.messages || []).map((m) => {
        const roleName = m.role === 'user' ? 'User' : (m.role === 'assistant' ? 'Assistant' : 'System');
        const cls = m.role === 'user' ? 'chat-msg user' : (m.role === 'assistant' ? 'chat-msg assistant' : 'chat-msg');
        return `<div class="${cls}"><div class="role">${roleName}</div><div>${escapeHtml(m.content || '')}</div></div>`;
      }).join('');
      messagesEl.scrollTop = messagesEl.scrollHeight;
    }

    function refreshConnectionSelect() {
      if (!connectionSelect) return;
      connectionCatalog = listConnections();
      const options = connectionCatalog.map((connection) => {
        const fullLabel = connection.name || connection.id;
        return `<option value="${escapeHtml(connection.id)}" title="${escapeHtml(fullLabel)}">${escapeHtml(truncateSelectLabel(fullLabel))}</option>`;
      });
      connectionSelect.innerHTML = options.join('');
      if (!findConnectionById(connectionCatalog, selectedConnectionId)) {
        selectedConnectionId = connectionCatalog[0]?.id || null;
      }
      chatSelectionState.connectionId = selectedConnectionId;
      connectionSelect.value = selectedConnectionId || '';
    }

    function refreshSessionSelect() {
      if (!sessionSelect) return;
      const bucket = getConversationBucket(selectedConnectionId);
      const options = (bucket.sessions || []).map((session) => {
        const title = session.title || 'New chat';
        return `<option value="${escapeHtml(session.id)}" title="${escapeHtml(title)}">${escapeHtml(truncateSelectLabel(title))}</option>`;
      });
      sessionSelect.innerHTML = options.join('');
      if (!(bucket.sessions || []).some((session) => session.id === selectedSessionId)) {
        selectedSessionId = bucket.sessions[0]?.id || null;
      }
      if (selectedSessionId) {
        sessionSelect.value = selectedSessionId;
      }
    }

    function applyChatSelection(connectionId, sessionId) {
      selectedConnectionId = connectionId || null;
      chatSelectionState.connectionId = selectedConnectionId;
      chatState = ensureConversation(selectedConnectionId, sessionId || getSelectedSessionId(selectedConnectionId));
      selectedSessionId = chatState.id;
      setSelectedSessionId(selectedConnectionId, selectedSessionId);
      scheduleChatStatePersist();
      currentChatContext = resolveChatContext(findConnectionById(connectionCatalog, selectedConnectionId));
      refreshConnectionSelect();
      refreshSessionSelect();
      if (providerLabelEl) {
        if (currentChatContext.connectionId) {
          const modelText = currentChatContext.model ? ` • ${currentChatContext.model}` : '';
          providerLabelEl.textContent = `${currentChatContext.providerLabel}${modelText}`;
        } else {
          providerLabelEl.textContent = 'Select connection in AI Connection Settings';
        }
      }
      renderMessages();
    }

    connectionSelect?.addEventListener('change', (event) => {
      const nextConnectionId = event.target.value || null;
      applyChatSelection(nextConnectionId, getSelectedSessionId(nextConnectionId));
    });

    sessionSelect?.addEventListener('change', (event) => {
      const nextSessionId = event.target.value || null;
      applyChatSelection(selectedConnectionId, nextSessionId);
    });

    newChatBtn?.addEventListener('click', () => {
      const next = createSession(selectedConnectionId);
      applyChatSelection(selectedConnectionId, next.id);
    });

    applyChatSelection(selectedConnectionId, selectedSessionId);

    function buildConversationTranscript() {
      const context = currentChatContext || resolveChatContext(findConnectionById(connectionCatalog, selectedConnectionId));
      const chatName = chatState?.title || 'Chat';
      const timestamp = new Date().toISOString();
      const header = [
        '# AI Chat Transcript',
        `Chat: ${chatName}`,
        `Connection: ${context.connectionName || 'n/a'}`,
        `Provider: ${context.providerLabel}`,
        `Model: ${context.model || 'default'}`,
        `Timestamp: ${timestamp}`,
        ''
      ];
      const body = (chatState.messages || []).map((entry) => `## ${entry.role.toUpperCase()}\n${entry.content || ''}\n`);
      return header.concat(body).join('\n');
    }

    exportBtn?.addEventListener('click', () => {
      if (!chatState.messages || !chatState.messages.length) {
        window.LogService?.logWarning?.('AI Chat', 'Export skipped (no messages)', { connectionId: selectedConnectionId, sessionId: selectedSessionId });
        return;
      }
      const transcript = buildConversationTranscript();
      const blob = new Blob([transcript], { type: 'text/markdown;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      const chatName = (chatState?.title || 'chat').replace(/\s+/g, '-').toLowerCase();
      link.href = url;
      link.download = `ai-chat-${chatName}-${Date.now()}.md`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
      window.LogService?.logInfo?.('AI Chat', 'export', { connectionId: selectedConnectionId, sessionId: selectedSessionId, messages: chatState.messages.length });
    });

    function setChatBusy(isBusy) {
      sendBtn.disabled = !!isBusy;
      stopBtn.disabled = !isBusy;
      sendBtn.classList.toggle('hidden', !!isBusy);
      stopBtn.classList.toggle('hidden', !isBusy);
    }

    async function send() {
      if (window.Trial && !window.Trial.enforce('aiChatSend')) { return; }
      const text = (inputEl.value || '').trim();
      if (!text) return;
      inputEl.value = '';
      chatState.messages.push({ role: 'user', content: text });
      chatState.messages.push({ role: 'assistant', content: '' });
      renderMessages();
      scheduleChatStatePersist();

      const context = currentChatContext || resolveChatContext(findConnectionById(connectionCatalog, selectedConnectionId));
      const connectionIdForLog = selectedConnectionId || null;
      const systemPrompt = context.prompt || '';

      const providerApiKey = (context.connectionId ? loadConnectionSecret(context.connectionId) : '')
        || window.AIStorage?.getApiKey?.(context.providerId || AI_FREE_PROVIDER_ID);
      if (!providerApiKey) {
        const last = chatState.messages[chatState.messages.length - 1];
        last.content = '[Error] API key is not configured. Open AI Connection Settings and save your key first.';
        renderMessages();
        setChatBusy(false);
        return;
      }

      if (!providerFactory || typeof providerFactory.createProvider !== 'function') {
        const last = chatState.messages[chatState.messages.length - 1];
        last.content = `${last.content || ''}\n[Error] AI provider factory unavailable.`;
        renderMessages();
        window.LogService?.logError?.('AI Chat', 'provider-factory-missing', { connectionId: connectionIdForLog });
        return;
      }

      const conversation = [];
      if (systemPrompt) conversation.push({ role: 'system', content: systemPrompt });
      const pending = chatState.messages.slice(0, chatState.messages.length - 1);
      pending.forEach((entry) => conversation.push({ role: entry.role, content: entry.content }));

      window.LogService?.logInfo?.('AI Chat', 'send', {
        connectionId: connectionIdForLog,
        sessionId: selectedSessionId,
        providerId: context.providerId,
        messageLength: text.length
      });

      setChatBusy(true);
      let provider = null;
      try {
        provider = providerFactory.createProvider(context.providerId, {
          advancedLogging: context.advancedLogging === true,
          connectionId: context.connectionId || null
        });
      } catch (error) {
        console.error('[RightActivityBar] Failed to create AI provider:', error);
      }

      let aborted = false;
      stopBtn.onclick = () => { aborted = true; };
      const onDelta = (delta) => {
        if (aborted) return;
        const last = chatState.messages[chatState.messages.length - 1];
        last.content += delta || '';
        renderMessages();
        scheduleChatStatePersist(300);
      };

      try {
        let finalText = '';
        if (provider && typeof provider.sendMessageStream === 'function') {
          const streamOptions = context.model ? { model: context.model } : {};
          if (context.requestOptions && typeof context.requestOptions === 'object') {
            Object.assign(streamOptions, context.requestOptions);
          }
          finalText = await provider.sendMessageStream(conversation, streamOptions, onDelta);
        } else if (provider && typeof provider.sendMessage === 'function') {
          const requestOptions = context.model ? { model: context.model } : {};
          if (context.requestOptions && typeof context.requestOptions === 'object') {
            Object.assign(requestOptions, context.requestOptions);
          }
          finalText = await provider.sendMessage(conversation, requestOptions);
        } else {
          throw new Error('AI provider unavailable');
        }
        if (!aborted) {
          const last = chatState.messages[chatState.messages.length - 1];
          last.content = finalText || last.content;
          renderMessages();
          scheduleChatStatePersist();
          window.LogService?.logInfo?.('AI Chat', 'response', {
            connectionId: connectionIdForLog,
            sessionId: selectedSessionId,
            providerId: context.providerId,
            length: (finalText || '').length
          });
        }
      } catch (error) {
        const last = chatState.messages[chatState.messages.length - 1];
        const rawError = error?.message || String(error);
        let friendlyError = rawError;
        const lower = typeof rawError === 'string' ? rawError.toLowerCase() : String(rawError).toLowerCase();
        if ((lower.includes('permission denied') || lower.includes('does not have access to model')) && context?.connectionId) {
          const requestedModel = context.model || 'selected model';
          friendlyError = `${rawError}\nTip: model "${requestedModel}" is not available for your project. Open AI Connection Settings, pick one of available models from suggestions, save connection, then retry.`;
        }
        last.content = `${last.content || ''}\n[Error] ${friendlyError}`;
        renderMessages();
        scheduleChatStatePersist();
        window.LogService?.logError?.('AI Chat', 'error', {
          connectionId: connectionIdForLog,
          sessionId: selectedSessionId,
          providerId: context.providerId,
          message: rawError
        });
      } finally {
        chatState.updatedAt = Date.now();
        if ((chatState.title || 'New chat') === 'New chat') {
          const firstUser = (chatState.messages || []).find((entry) => entry.role === 'user' && typeof entry.content === 'string' && entry.content.trim());
          if (firstUser) {
            chatState.title = firstUser.content.trim().slice(0, 48);
            refreshSessionSelect();
          }
        }
        scheduleChatStatePersist();
        setChatBusy(false);
        stopBtn.onclick = null;
      }
    }

    sendBtn.addEventListener('click', send);
    inputEl.addEventListener('keydown', (event) => {
      if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') {
        event.preventDefault();
        send();
      }
    });

    renderMessages();
  }

  function safeRenderChat() {
    try {
      const result = renderChat();
      if (result && typeof result.catch === 'function') {
        result.catch((error) => {
          console.error('[RightActivityBar] renderChat failed', error);
        });
      }
      return result;
    } catch (error) {
      console.error('[RightActivityBar] renderChat failed', error);
      return null;
    }
  }

  function initResizer() {
    const resizer = document.getElementById('right-pane-resizer');
    if (!resizer) return;
    try { window.__RightPaneResizerTeardown?.() } catch (_) {}
    let dragging = false;
    function onMouseMove(e) {
      if (!dragging) return;
      // compute new width: right pane anchored to right with 48px bar
      const viewport = window.innerWidth || document.documentElement.clientWidth;
      const barWidth = 48; // right activity bar
      let width = Math.round((viewport - barWidth) - e.clientX);
      const { min, max } = getRightPaneWidthBounds();
      if (width < min) width = min;
      if (width > max) width = max;
      document.documentElement.style.setProperty(PANE_WIDTH_VAR, width + 'px');
      try { localStorage.setItem(WIDTH_KEY, String(width)); } catch(_) {}
      try { if (window.__rightPromptEditor) window.__rightPromptEditor.layout(); } catch(_) {}
    }
    function onMouseUp() { if (dragging) { dragging = false; document.body.classList.remove('resizing-right-pane'); } }
    const onMouseDown = (e) => {
      e.preventDefault(); dragging = true; document.body.classList.add('resizing-right-pane');
    }
    const onDoubleClick = () => {
      const { min, max } = getRightPaneWidthBounds()
      const width = Math.max(min, Math.min(max, DEFAULT_WIDTH))
      document.documentElement.style.setProperty(PANE_WIDTH_VAR, width + 'px');
      try { localStorage.setItem(WIDTH_KEY, String(width)); } catch(_) {}
    }

    resizer.addEventListener('mousedown', onMouseDown);
    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
    resizer.addEventListener('dblclick', onDoubleClick);

    window.__RightPaneResizerTeardown = function () {
      try { resizer.removeEventListener('mousedown', onMouseDown) } catch (_) {}
      try { document.removeEventListener('mousemove', onMouseMove) } catch (_) {}
      try { document.removeEventListener('mouseup', onMouseUp) } catch (_) {}
      try { resizer.removeEventListener('dblclick', onDoubleClick) } catch (_) {}
      window.__RightPaneResizerListeners = false
    }
    window.__RightPaneResizerListeners = true
  }

  try {
    console.info('[RightActivityBar] pre-ready marker', { version: SCRIPT_VERSION });
  } catch (_) {}

  try {
    console.info('[RightActivityBar] pre-readyState checkpoint', { version: SCRIPT_VERSION });
  } catch (_) {}

	  try {
	    console.info('[RightActivityBar] readyState', document.readyState);
	  } catch (error) {
	    try { console.error('[RightActivityBar] readyState check failed', error); } catch (_) {}
	  }
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
      try {
        console.info('[RightActivityBar] DOMContentLoaded fired', { version: SCRIPT_VERSION });
        window.__RightActivityBar_runInit?.('dom-content-loaded');
      } catch (error) {
        try { console.error('[RightActivityBar] init failed (DOMContentLoaded)', error); } catch (_) {}
      }
    }, { once: true });
  } else {
    try {
      console.info('[RightActivityBar] init immediate');
      window.__RightActivityBar_runInit?.('ready-state');
    } catch (error) {
      try { console.error('[RightActivityBar] init failed (immediate)', error); } catch (_) {}
    }
  }
  try {
    console.info('[RightActivityBar] bootstrap end', { version: SCRIPT_VERSION });
  } catch (_) {}
  try {
    setTimeout(() => {
      window.__RightActivityBar_runInit?.('safety-timeout');
    }, 500);
  } catch (_) {}
})();
} catch (error) {
  try {
    console.error('[RightActivityBar] bootstrap failed', error);
  } catch (_) {}
}
  // Monaco editor theme toggle for Debug Insights
  function applyDebugInsightsMonacoTheme(enable) {
    try {
      if (typeof monaco === 'undefined' || !monaco.editor) return;
      // Determine current default theme to restore later
      const isDark = document.body.classList.contains('theme-type-dark');
      const defaultTheme = isDark ? 'vs-dark' : 'vs';
      if (!enable) {
        // Restore original theme if we previously switched
        const prev = window.__prevMonacoTheme || defaultTheme;
        monaco.editor.setTheme(prev);
        return;
      }
      // Compute colors from CSS variables (fallbacks included)
      const cs = getComputedStyle(document.documentElement);
      const bg = cs.getPropertyValue('--vscode-editor-background').trim() || (isDark ? '#1e1e1e' : '#ffffff');
      const fg = cs.getPropertyValue('--vscode-editor-foreground').trim() || (isDark ? '#d4d4d4' : '#333333');
      const accent = cs.getPropertyValue('--accent-color').trim() || '#4b9cff';
      const sel = cs.getPropertyValue('--selection-bg').trim() || (isDark ? '#264f78' : '#bcdcff');
      const lineHl = cs.getPropertyValue('--line-highlight-bg').trim() || (isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)');
      const gutter = cs.getPropertyValue('--border-color').trim() || (isDark ? '#3c3c3c' : '#e0e0e0');
      // Remember previous theme once when switching on
      if (!window.__prevMonacoTheme) window.__prevMonacoTheme = defaultTheme;
      monaco.editor.defineTheme('debug-insights', {
        base: isDark ? 'vs-dark' : 'vs',
        inherit: true,
        rules: [
          { token: 'identifier', foreground: fg.replace('#','') },
          { token: 'type.identifier', foreground: accent.replace('#',''), fontStyle: 'bold' },
          { token: 'keyword', foreground: accent.replace('#','') }
        ],
        colors: {
          'editor.background': bg,
          'editor.foreground': fg,
          'editorLineNumber.foreground': gutter,
          'editorCursor.foreground': accent,
          'editor.selectionBackground': sel,
          'editor.lineHighlightBackground': lineHl,
          'editorIndentGuide.activeBackground': gutter,
          'editorBracketMatch.border': accent
        }
      });
      monaco.editor.setTheme('debug-insights');
    } catch (_) {}
  }
