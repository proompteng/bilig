# Luckysheet Lessons For `bilig`

This document captures what `bilig` should learn from the local Luckysheet repo at `/Users/gregkonush/github.com/luckysheet`.

Reviewed source:

- `/Users/gregkonush/github.com/luckysheet/README.md`
- `/Users/gregkonush/github.com/luckysheet/src/core.js`
- `/Users/gregkonush/github.com/luckysheet/src/store/index.js`
- `/Users/gregkonush/github.com/luckysheet/src/global/refresh.js`
- `/Users/gregkonush/github.com/luckysheet/src/global/draw.js`
- `/Users/gregkonush/github.com/luckysheet/src/global/createsheet.js`

## Executive Summary

Luckysheet is not a codebase to copy structurally.

It is useful as:

- a reference for feature expectations in a browser spreadsheet
- a reminder of the interaction breadth users expect
- a warning about monolithic global-state architecture

The most valuable Luckysheet lessons for `bilig` are negative lessons:

- do not centralize everything in one mutable global store
- do not let refresh logic own data mutation, undo, server sync, and paint in one path
- do not let renderer and app logic collapse into giant imperative modules

There are still a few positive ideas worth keeping:

- direct canvas drawing on visible ranges
- practical feature completeness for spreadsheet interactions
- explicit visible-row and visible-column caches

## What Luckysheet Gets Right

### 1. It takes spreadsheet interactions seriously

Even the old codebase makes one thing clear: a real spreadsheet is more than just cell editing.

The feature and source layout reflect:

- fill handle
- multi-selection
- freeze
- formulas
- sorting and filtering
- validation
- comments
- charts
- collaboration hooks

That is useful because it keeps the product bar honest.

For `bilig`, the lesson is:

- spreadsheet quality is defined by interaction completeness, not only formula support

### 2. It uses visible row and column caches

`src/store/index.js` keeps:

- `visibledatarow`
- `visibledatacolumn`
- width and height accumulators

That is the right category of optimization even if the architecture around it is messy.

For `bilig`, the transferable lesson is:

- viewport math should be cached explicitly
- row and column visibility data should be cheap to query
- visible-range structures matter for scroll performance

### 3. It renders directly against canvas

`src/global/draw.js` is a reminder that spreadsheet rendering usually wants:

- direct control over painting
- precise text and border geometry
- visible-range clipping
- DPR handling

That is conceptually aligned with where `bilig` needs to go if it wants more control and more performance than a generic grid wrapper can give.

## What Luckysheet Does Poorly

### 1. Giant mutable global state

`src/store/index.js` is a massive singleton state bag containing:

- workbook data
- viewport data
- selection state
- copy/paste state
- chart state
- undo/redo state
- plugin loading state
- caches
- collaborative editing state

That is exactly the kind of architecture `bilig` should avoid.

Why it is bad:

- hidden coupling
- hard-to-reason updates
- leak risk
- impossible ownership boundaries
- very high regression risk as the product grows

### 2. Refresh logic is overloaded

`src/global/refresh.js` is doing too much at once:

- data mutation
- server sync
- undo/redo management
- formula execution
- chart refresh
- canvas refresh
- selection sync

That is not clean spreadsheet architecture. It is a warning sign.

For `bilig`, the lesson is:

- edit application, engine invalidation, transport sync, and repaint scheduling should be separate concerns

### 3. Rendering, DOM, state, and behavior are not cleanly separated

The codebase mixes:

- direct DOM logic
- jQuery-era event wiring
- global state mutation
- canvas drawing
- workbook operations

This is precisely the kind of entanglement that makes long-term maintenance painful.

### 4. Initialization is very broad and imperative

`src/core.js` and `src/global/createsheet.js` show a product boot flow that wires many subsystems through global state and delayed imperative restoration.

That can get a demo working, but it is not a strong model for a modern maintainable workbook system.

## What `bilig` Should Copy

### A. Visible-range caches and direct render awareness

Keep strengthening:

- visible row/column caches
- viewport math
- cheap render-range queries
- direct render ownership

Those are legitimate spreadsheet needs.

### B. Interaction completeness as a product bar

Luckysheet is a useful reminder that users notice:

- fill handle quality
- selection behavior
- freeze behavior
- paste semantics
- filter/sort responsiveness

Those details matter at least as much as adding exotic features.

## What `bilig` Should Explicitly Reject

### 1. Singleton global stores

Do not collapse workbook runtime, UI state, render caches, transport state, and feature state into one global bag.

### 2. One giant refresh path

Do not let one function manage:

- engine edits
- server sync
- undo bookkeeping
- formula execution
- paint scheduling

Those need explicit ownership boundaries.

### 3. jQuery-style imperative surface architecture

`bilig` should not move toward:

- implicit DOM lookups
- ambient globals
- mutation-heavy controller webs

## Recommended Actions For `bilig`

### Near term

- keep visible-range and viewport caches tight and measurable
- keep spreadsheet interaction polish a top priority

### Medium term

- continue splitting runtime, render, and UI state ownership
- keep edit application, sync, and paint scheduling separate

### Long term

- own the grid render plane directly without inheriting Luckysheet-style global-state debt

## Bottom Line

Luckysheet is useful mostly as a cautionary reference.

Copy:

- visible-range awareness
- direct render thinking
- interaction completeness pressure

Reject:

- global singleton architecture
- giant refresh/update funnels
- mixed rendering, state, and transport concerns
