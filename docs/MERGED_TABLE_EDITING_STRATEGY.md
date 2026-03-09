# Merged Table Editing Strategy

## Problem

The current structured HWPX engine can edit table cells safely, but row insertion/deletion still rejects any table that contains merged cells anywhere in the table.

Today that limitation comes from two choices in the engine:

- `tableSupportsRowMutations()` rejects a table when any cell has `rowSpan > 1` or `colSpan > 1`.
- `reindexSimpleTable()` rewrites every `cellAddr` and force-resets every `cellSpan` to `1x1`, which is only valid for fully non-merged tables.

Relevant code:

- [`structuredHwpx.ts`](../src/core/structuredHwpx.ts)
- [`templateFill.ts`](../src/core/templateFill.ts)

This is too coarse for real-world Korean forms. Many practical forms have:

- merged title/header rows
- non-merged data-entry rows below the header
- repeat-safe body regions inside an otherwise merged table

The real `.hwp` tax form fixture used during validation falls into that category: its body table contains merged header rows, but the data rows below them are logically repeatable.

## External Findings

### 1. HWPX is structured enough for safe merged-table editing

Hancom documents HWPX as an XML/OWPML-based open package format rather than an opaque binary editing surface. That means merge-safe table editing is a data-structure problem, not a format impossibility.

Source:

- [Hancom HWPX format overview](https://tech.hancom.com/hwpxformat/)

### 2. Existing HWPX tooling already uses a logical grid model

`python-hwpx` documents and implements:

- `iter_grid()`
- `get_cell_map()`
- `set_cell_text(..., logical=True, split_merged=True)`
- `split_merged_cell()`
- `merge_cells()`

That confirms the right model is:

- map logical `(row, col)` coordinates to physical cells
- track anchor cells and covered cells separately
- split or preserve merges intentionally instead of flattening the table

Sources:

- [python-hwpx usage docs](https://airmang.github.io/python-hwpx/usage/)
- [python-hwpx API docs](https://airmang.github.io/python-hwpx/api_reference/)
- [python-hwpx repository](https://github.com/airmang/python-hwpx)

Note:

- `python-hwpx` is useful as an algorithm reference, but its repository currently contains a non-commercial license file that should not be copied into a production commercial path without explicit license review.

### 3. OpenHWP already models merge semantics explicitly

OpenHWP's HWPX/HWP conversion layers track:

- `row_span`
- `column_span`
- `is_merged`
- merge origin semantics

This supports the same conclusion: the right long-term core is an explicit table IR, not raw row cloning.

Source:

- [OpenHWP repository](https://github.com/openhwp/openhwp)

## Decision

We should replace the current "whole table must be non-merged" rule with a region-based merged-table strategy.

The default behavior must remain deterministic and fail-closed.

## Design

### 1. Introduce a logical table grid

Add a table-level logical map in the structured engine:

- `TableGridEntry`
- `logicalRow`
- `logicalColumn`
- `anchorRow`
- `anchorColumn`
- `rowSpan`
- `colSpan`
- `isAnchor`
- `physicalCellRef`

This grid is derived from:

- `cellAddr`
- `cellSpan`
- physical `<tr>/<tc>` ordering

This must exist independently of the current flat `table.cells` list.

### 2. Distinguish physical rows from logical rows

Current code assumes:

- one physical row mutation = one logical row mutation

That is false once merges exist.

The engine must instead reason about:

- physical rows in XML
- logical grid rows
- cells that originate above and cover the current row

### 3. Replace whole-table rejection with boundary analysis

For any requested repeat/delete/insert region, compute whether merges cross the region boundary.

Definitions:

- `repeatStart`
- `repeatEnd`
- `insertAt = repeatEnd + 1`

Classify every merged anchor cell as one of:

- `outside`: fully above or below the region
- `internal`: starts and ends inside the region
- `crosses_top_boundary`
- `crosses_bottom_boundary`
- `covers_insert_boundary`

Rules:

- internal merges are safe to clone with the region
- outside merges are safe to keep unchanged
- boundary-crossing merges are not safe under strict mode

### 4. Support three explicit mutation modes

#### Mode A: `strict`

Default.

Behavior:

- fail if any merge crosses the repeat region boundary
- allow merged tables when the selected region itself is repeat-safe

This is the first production target.

#### Mode B: `split_boundary_merges`

Optional explicit mode.

Behavior:

- split merges that cross the boundary before region duplication
- then perform the repeat on normalized rows

This matches the `python-hwpx` style `split_merged` concept.

#### Mode C: `extend_boundary_merges`

Do not implement in v1.5.

Behavior:

- if a merge covers the insertion boundary, extend its `rowSpan`

This is valid in some layouts but too risky as a default because it changes visual semantics.

## Mutation Primitives

### 1. `buildTableGrid(tableRef)`

Returns a complete logical map for the table.

Needed for:

- safe cell targeting
- region safety analysis
- future logical-row APIs

### 2. `analyzeRepeatRegion(tableRef, startRow, endRow)`

Returns:

- `safe: boolean`
- `hasMergedCells: boolean`
- `boundaryCrossings: ...`
- `internalMergeGroups: ...`
- `recommendedMode: 'strict' | 'split_boundary_merges'`

This should be used by both:

- `fill-template`
- future host-editor UI

### 3. `shiftRowAddressesPreserveSpans(tableRef, insertAt, delta)`

Updates `cellAddr.rowAddr` for physical cells below the insertion point without rewriting all spans.

This replaces the current `reindexSimpleTable()` behavior for general tables.

### 4. `clonePhysicalRegionPreserveTopology(tableRef, startRow, endRow)`

Clones the physical rows for a repeat-safe region and preserves:

- `cellSpan`
- `cellSz`
- `cellMargin`
- paragraph structure
- border/fill refs

Internal merged cells inside the repeated region remain merged.

### 5. `splitMergedCellAtLogicalPosition(tableRef, row, col)`

Required for `split_boundary_merges`.

Behavior:

- find the anchor cell covering `(row, col)`
- if already `1x1`, no-op
- partition width/height across the covered rectangle
- emit new physical cells with `rowSpan=1`, `colSpan=1`
- preserve style/margin/paragraph defaults

### 6. `repeatTableRegion(tableRef, startRow, endRow, count, mode)`

High-level operation:

1. build grid
2. analyze region
3. if `strict` and boundary crossings exist, fail
4. if `split_boundary_merges`, split boundary overlaps first
5. clone region
6. shift downstream row addresses
7. refresh block refs
8. validate package

## API Changes

Extend `TableRepeatInstruction` to support region semantics explicitly.

Proposed shape:

```ts
type TableRepeatBoundaryPolicy = 'reject' | 'split_boundary_merges';

interface TableRepeatInstructionV2 {
  tableBlockId: string;
  templateRowIndex: number;
  templateEndRowIndex?: number;
  rows: Array<string[] | Record<string, string>>;
  logical?: boolean;
  boundaryPolicy?: TableRepeatBoundaryPolicy;
}
```

Rules:

- `templateEndRowIndex` omitted means single-row repeat
- `logical=true` means analyze logical grid rows
- `boundaryPolicy` defaults to `reject`

## Reference Analysis Changes

`analyze-reference` should expose repeatability diagnostics, not just raw table summaries.

Add:

- `repeatableRegions`
- `tableBlockId`
- `startRowIndex`
- `endRowIndex`
- `logicalColumnCount`
- `hasInternalMerges`
- `boundaryCrossingMerges`
- `supportedBoundaryPolicies`
- `reason`

That lets the host app suggest:

- "this row band is safe to repeat"
- "this row band needs merge splitting"
- "this region is not automatable safely"

## Implementation Phases

### Phase 1: Safe merged-table region support

Goal:

- support repeating rows inside a merged table when merges do not cross the selected region boundary

Work:

- add logical grid builder
- add region analysis
- replace whole-table rejection with region-based rejection
- replace `reindexSimpleTable()` with address-preserving downstream shift for region insertions

Expected impact:

- the current real tax-form fixture should allow repeating the body rows in table `0:3`
- header merges remain preserved

### Phase 2: Boundary split support

Status:

- Implemented for single-row repeats via `boundaryPolicy: 'split_boundary_merges'`
- Real fixture verified: vertically merged row-band boundaries can now be repeated without 500s
- Band-covering labels are restored by extending the original anchor cell `rowSpan` instead of leaving fragmented split cells

Goal:

- support repeating rows even when a merged cell overlaps the repeat boundary

Work:

- implement `splitMergedCellAtLogicalPosition`
- add `boundaryPolicy: 'split_boundary_merges'`
- add regression tests for horizontal, vertical, and rectangular merges

### Phase 3: Logical editing surface

Goal:

- make merged tables editable from host UI without exposing raw XML details

Work:

- use logical row/column targeting in mutation API
- expose merge-aware cell map in session payloads
- surface repeat-safe region suggestions in the host editor

Current status:

- region-level repeat is now implemented for contiguous template row bands
- single-row and multi-row repeats both support `boundaryPolicy: 'split_boundary_merges'`
- internal merges inside the repeated region are preserved because the engine clones the physical row band as a unit

## Testing Matrix

We need fixtures for:

- fully non-merged table
- merged header + simple body rows
- vertical merge crossing insert boundary
- horizontal merge inside repeated region
- rectangular merge inside repeated region
- boundary split case
- real imported `.hwp -> .hwpx` forms

Acceptance criteria:

- no 500s for merged-table repeat attempts
- fail-closed with explicit diagnostics when a region is unsafe
- safe regions repeat without changing unrelated merges
- downstream row addresses remain valid after multiple repeats
- markdown export reflects repeated rows correctly after save

## Immediate Next Step

Implement Phase 1 first.

That gives the highest-value production win:

- merged tables are no longer globally blocked
- real government/business forms with merged headers become usable
- risky merge-splitting logic stays out of the default path until explicitly added
