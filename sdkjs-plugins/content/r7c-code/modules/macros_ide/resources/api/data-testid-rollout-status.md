# Data-testid Rollout Status

## Plan checklist verification

- [x] Naming convention approved (kebab-case)
- [x] Helper added (`modules/macros_ide/scripts/ui/testid.js`)
- [x] Static DOM covered (`index.html` + welcome popup html)
- [x] Dynamic DOM covered (sidebar/tree/activity bars/right pane/tabs/console/settings)
- [x] Additional dynamic zones covered (context menu/conversion dialog/AI Job Monitor/agent tree/welcome popup)
- [x] Over-tagging policy respected (no per-cell/per-char tagging)
- [x] Attribute-only changes (no business logic branch rewrites)
- [x] Selector registry documented (`data-testid-registry.md`)
- [ ] Selenium page objects switched to testid-only locators
- [ ] E2E smoke run completed in this session
- [x] Coverage/status report prepared (this file)

## Notes on remaining unchecked items
1. Selenium page objects are not present in the repository tree (no dedicated Selenium/e2e package found), so migration cannot be applied directly in-code here.
2. E2E smoke was not run because there is no explicit Selenium test runner config in this repository snapshot.

## Recommended next execution step (outside this plugin repo)
1. Update Selenium Page Objects to use only `By.cssSelector('[data-testid="..."]')`.
2. Apply explicit waits:
   - `presenceOfElementLocated`
   - `visibilityOfElementLocated`
   - `invisibilityOfElementLocated`
3. Re-run smoke flows for:
   - explorer/tree rerender
   - settings open/close
   - right pane tools
   - context menu
   - conversion dialog
   - AI Job Monitor
   - agent tree interactions
