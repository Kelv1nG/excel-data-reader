# Multilevel table header discovery — discussion notes

Status: draft idea from the 2026-10-26 discussion. This is not an approved design or a change to `SPEC.md`.

Planning follow-up (2026-10-29): see the [reviewed draft design](../superpowers/specs/2026-10-29-multilevel-table-headers-design.md) and [implementation plan](../superpowers/plans/2026-10-29-multilevel-table-headers.md) for the consolidated decisions and proposed API. Earlier open questions below are historical where the design now resolves them.

## Motivation

Some ordinary tables have several header rows. A header may be merged down across rows, while a group header may be merged across several columns. For example:

```text
                 header row 1             header row 2          header row 3
column B         col1 (merged down through all header rows)
columns C–H      main_header (merged across C–H)
columns C–E                               subheader1 (C–E)     cola, colb, colc
columns F–H                               subheader2 (F–H)     cold, cole, coldf
```

The existing `TableQuery` matches headers on one row. `MatrixQuery` understands grouped headers, but its discovery contract is for sectioned matrices. `read_range()` can retrieve known cells without discovering their header hierarchy. A table query for this layout would fill a real API gap.

## Direction discussed

- Support a caller-declared header depth with more than two rows, rather than hard-coding a two-row layout. Discovery would look for a consecutive header band of that depth.
- Treat a column's ordered labels as its full header path. A vertically merged `col1` is one column identified by `col1`; a leaf under the example group has a path such as `("main_header", "subheader1", "cola")`.
- Let a caller select a parent group and receive all its descendant leaf columns. Selecting both `subheader1` and `subheader2` in the example returns six columns. Selecting a deeper path can narrow the result to one branch or one leaf.
- Produce flat record keys formed from the full header path, such as `main_header__subheader1__cola`, rather than nested data records. The caller wants control over how much of the hierarchy they specify when searching; a top-level header might be the only identifier they know.
- Use `__` between header levels in flat output names. Include every authored label inside the declared header band, even a top label merged across the entire table.
- A single-column label followed by empty header cells has the same logical path whether it is vertically merged through those cells or not. For example, both a plain `Store` in C2 with empty C3 and a `Store` merged across C2:C3 yield `Sales Report__Store`.
- When one authored group spans several physical columns and its child columns have no labels, give those leaves synthetic names by position. A two-column `Store` group yields `Sales Report__Store__col1` and `Sales Report__Store__col2` from left to right. Numbering restarts at `col1` inside each merged group and each separate table match. Positions count all physical columns in that group: if a three-column group has `New` in its first leaf and no labels in its next two leaves, the latter are `col2` and `col3`. This keeps the names tied to the visible layout. This does not yet define behavior for a wholly unlabeled column outside a named group.
- Preserve typed data values and the physical source coordinates of data and header labels. Repeated or overlapping selections should not duplicate a physical column.
- Report ambiguous labels or paths instead of arbitrarily choosing one group. For example, a bare `subheader1` may need a parent path if that label occurs under multiple parents.
- If the same multilevel header band appears more than once, return one match per occurrence, each with its own header coordinates and data bounds. Do not combine their rows into one result; callers can extract or choose the matches explicitly.
- A complete repeated header in the same physical columns starts a new table even without a blank separator. The previous table ends immediately before that header block. The user accepted this default before requesting an implementation plan.

The nested mapping suggested during discussion is a possible query syntax, for example `{"main_header": {"subheader1": ["cola", "colb", "colc"]}}`. Full paths may be the simpler internal representation. Neither public query shape nor exact path matching semantics has been chosen.

## Still to decide

- Whether a parent label selects all descendants, identifies a table without selecting columns, or can do either through separate query fields. Group expansion was received positively, but table identification also matters.
- How to represent a vertically merged leaf such as `col1` in a nested query, and how explicit the merge and header-depth rules should be.
- How callers include an adjacent, possibly unnamed data column (`col0`) to the left of a matched header column.
- How to choose the first of several matching tables without relying on implicit selection; location constraints or explicit physical ordering may be appropriate.
- Whether several blank-separated data blocks may share one header band. Repeated headers have the boundary rule above; shared headers remain separate future scope.
- How to form collision-free flat keys when authored labels contain separators, repeat, or normalize to the same text.
- Whether this belongs in a separate query type or extends `TableQuery`, whose current one-row semantics should remain stable unless the public contract is deliberately revised.

Before implementation, turn these notes into a reviewed design, update `SPEC.md` for the chosen public contract, and add contract-level tests.
