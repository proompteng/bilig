# React Spectrum UI Philosophy For `bilig`

This document distills the parts of React Spectrum's architecture, theming model, and API discipline that should govern `bilig` UI work.

The source of truth for this philosophy is the local React Spectrum repo at `/Users/gregkonush/github.com/react-spectrum`, especially:

- `rfcs/2019-v3-architecture.md`
- `rfcs/2019-v3-theming.md`
- `specs/api/Guidelines.md`

`bilig` is not React Spectrum, and it should not copy Spectrum visuals mechanically. The goal is to copy the product and system discipline that makes React Spectrum polished, calm, consistent, and maintainable.

## Core Philosophy

### 1. Separate state, behavior, and themed rendering

React Spectrum's central idea is that a mature component system should split into:

- state management
- behavior and accessibility
- themed rendering

For `bilig`, this means:

- render-heavy components should not own complex state machines
- keyboard and pointer policy should live in controllers, hooks, or dedicated modules
- styling and layout components should stay mostly declarative

If a component mixes interaction logic, business state, geometry math, and styling, it is already drifting toward tech debt.

### 2. Theme through tokens, not one-off styling

React Spectrum's theming model is built on reusable variables, not scattered hardcoded values.

For `bilig`, shared tokens should define at minimum:

- app background
- surfaces
- muted surfaces
- borders
- text hierarchy
- accent and focus states
- radii
- shadows
- control heights
- density spacing

Rules:

- prefer one token system reused everywhere over local fixes
- do not introduce raw hex values into components unless promoting a new token
- do not let one control invent its own radius, height, or focus treatment
- nested or alternate surfaces should still come from the same token vocabulary

### 3. Quiet by default

React Spectrum components are intentionally calm. Most controls do not shout until the user hovers, focuses, selects, or presses them.

For `bilig`, enterprise polish means:

- default controls are quiet and legible
- selection and emphasis are explicit and rare
- strong accent fills are reserved for active state, primary actions, or current selection
- backgrounds, shadows, and rounding stay restrained

If a surface needs blur, glow, oversized rounding, or heavy decoration to feel "designed", the hierarchy is probably wrong.

### 4. Density is a system, not a component-level choice

React Spectrum treats sizing and spacing as deliberate scales.

For `bilig`, each surface should pick one density rhythm and stay on it:

- toolbar controls should share one height
- formula bar controls should align with that same height unless there is a product reason not to
- tabs, chips, popovers, and menus should reuse the same radius family
- vertical stacking between toolbar, formula bar, grid, and footer should be tight and intentional

Do not let local tweaks create hidden 30px, 32px, and 36px systems on the same screen.

### 5. Accessibility is component behavior, not cleanup

React Spectrum treats accessibility as part of behavior hooks rather than a late audit.

For `bilig`, every interactive control must have:

- explicit labeling
- predictable focus visibility
- keyboard reachability
- correct popup semantics where relevant
- stable focus continuity when overlays open and close

Menus, tabs, color pickers, popovers, and toolbar buttons must behave correctly before they are considered polished.

### 6. Public APIs must be boring and stable

React Spectrum's API guidance is disciplined and predictable.

For `bilig`, component APIs should follow these rules:

- state booleans use `is...`
- capability booleans use `allows...`
- callbacks use `on...`
- prop-specific changes use `on...Change`
- uncontrolled props use `default...`
- prefer enums or string unions over booleans when the API may grow
- prefer `start` and `end` over `left` and `right` when the meaning is alignment rather than literal side

If a prop name is ambiguous, it will rot.

## Styling Rules For `bilig`

### Required rules

- Use shared CSS variables in `apps/web/src/index.css` or theme modules before styling components.
- Reuse the same control height, radius, focus ring, and shadow tokens across the workbook shell.
- Popovers should read like tools, not floating marketing cards.
- Use subtle separators and grouping instead of large framed panels.
- Keep hover and selected states visible but restrained.
- Use one accent language across the workbook shell.

### Forbidden patterns

- ad hoc `rounded-xl`, `rounded-2xl`, blur, or glow on enterprise controls
- component-local hardcoded colors when an existing token can express the state
- one-off control sizes that break rhythm
- styling through private child selectors when an explicit prop or token should exist
- mixing business logic, input policy, and visual structure in a single giant component

## Architecture Rules For UI Work

When building or refactoring workbook UI:

1. Extract stateful behavior first.
2. Introduce or normalize tokens second.
3. Tighten spacing and hierarchy third.
4. Add only the minimal visual emphasis needed.

Preferred shape:

- state/controller module
- behavior hook or controller
- themed primitive
- composed product surface

Avoid:

- giant render component with state, geometry, keyboard handling, pointer handling, and styling mixed together

## `bilig`-Specific Application

### Workbook shell

The workbook shell should feel calm, dense, and tool-like:

- thin borders
- restrained radii
- consistent 32px-class controls unless a documented exception exists
- shallow shadows only on true overlays
- muted surfaces that frame the grid without competing with it

### Toolbar

The toolbar should follow a single grammar:

- quiet default buttons
- consistent trigger and button heights
- separators instead of boxed groups
- icon buttons only when the meaning is obvious; otherwise use icon plus text
- active state via shared accent tokens, not bespoke background colors

### Popovers and menus

Popovers should:

- align tightly to their trigger
- use the same panel radius and border system
- keep keyboard focus predictable
- avoid oversized padding or decorative styling

### Grid

The grid should preserve visual calm even when performance work gets more aggressive:

- GPU rendering should reproduce the same product tokens, not invent a new visual language
- selection chrome should remain precise and restrained
- body fills, borders, and grid lines must stay consistent with the shell token system

## Review Checklist

Before shipping a UI change, check:

- Is state separated from rendering?
- Are colors, radii, and shadows coming from shared tokens?
- Do controls align on one height and radius system?
- Are hover, focus, selected, and pressed states consistent?
- Does the popup or menu behave like a tool instead of a card?
- Are `aria-*` attributes and focus behavior explicit?
- Are prop names stable and unsurprising?
- Did the change reduce ad hoc styling or add more of it?

## What We Intentionally Copy From React Spectrum

- architectural separation
- token-first theming
- density discipline
- quiet default styling
- accessibility-first behavior
- API naming discipline

## What We Do Not Copy Blindly

- Adobe or Spectrum branding
- literal visual styling
- component names or package structure
- generic enterprise UI clichés that do not fit the workbook product

The requirement is not "look like Spectrum." The requirement is "be as deliberate and systematized as a mature design system."
