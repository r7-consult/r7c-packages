# Changelog

## 1.0.3 - 2026-03-13

### Navigation and UX
- Removed redundant landing layers (`Разделы` / `Курсы`) from the primary start flow and open the default course immediately via `defaultCourseId`.
- Reworked left navigation so users see the full content tree of `Разработка макросов с нуля` сразу при входе.
- Added an additional bottom navigation section sourced from the template library course (`r7-macros`) and rendered as `Библиотека шаблонных макросов для обучения.`.
- Hid the back control (`Назад к курсам`) in module view to avoid dead space and unnecessary hierarchy hops.

### Content structure
- Updated learning manifest:
  - kept two active learning tracks: interactive macro course + template library.
  - renamed `r7-macros` title/description to `Библиотека шаблонных макросов для обучения.`.
  - removed extra high-level courses that were marked as unnecessary in UI flow (`javascript`, `dax`, `sql`, `mdx`) from the visible manifest structure.
- Added interactive freeform sandbox module:
  - `module-sandbox` (`Редактор макросов`) in `interactive-manifest.json`.
  - lesson content file `module-07-sandbox/sandbox-editor.md`.

### Template library source
- Switched template library source from:
  - `modules/macros_ide/resources/learning/examples-manifest.json`
  - `modules/macros_ide/resources/learning/examples-macros`
- to:
  - `modules/macros_ide/resources/examples-manifest.json`
  - `modules/macros_ide/resources/examples`
- Excluded legacy lesson bucket `r7-macros-lessons` from the aggregated "library at bottom" navigation block, so the library reflects the actual `resources/examples` content.

### Header and floating action button
- Simplified top header by removing non-functional textual branding block.
- Tightened header layout spacing and aligned action controls to the right.
- Added floating Slider Data button to `index.html` and connected `slider-flyout.js`.
- Updated floating button visuals (circular style, elevated hover, shadow).
- Updated training branding target to Slider Data URL and logo in runtime constants.

### Packaging and tooling
- Plugin metadata version bumped to `1.0.3` in `config.json`.
- Added `archiver` to dev dependencies for packaging tooling consistency.
- Updated package artifact naming target to `SmartDocumentation-v1.0.3.plugin`.
