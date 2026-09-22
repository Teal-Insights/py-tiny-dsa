# Series bindings

Author `inputs.bindings.yaml`, `outputs.bindings.yaml`, `internals.bindings.yaml`,
and (when needed) `constants.bindings.yaml` here.

Bootstrap extract (`--extract-graph` / `--stop-after-stage extract`) is graph-first: empty `series: []` placeholder shards are fine so you can review `artifacts/dependency-graph/` before bindings exist. excel-grapher 5.1.4+ also loads and merges those placeholders (including divergent `concept_scheme` blocks). Author real series before export so the public API and leaf coverage are complete.

Use schema version `1.19.0` and the vendored skill in [.agents/skills/author-bindings](../.agents/skills/author-bindings/SKILL.md). Prefer authoring from an extracted graph rather than guessing sheet geometry up front.

## Constant bindings (reader-only leaves)

Use `constant: {}` for graph **leaves** that formulas read as fixed parameters,
scenario knobs, or lookup seeds, but that should **not** appear as required
keyword-only `compute_*` arguments or published outputs. Without a constant
binding, those leaves stay as unnamed cell reads in generated formula bodies even
when the domain model already names them.

Put constant series in `constants.bindings.yaml` (or any mergeable
`*.bindings.yaml` shard). Same YAML shape as public bindings; declare
`constant: {}` instead of `input` / `output` / `internal`:

```yaml
schema_version: 1.19.0
series:
  - id: shock_year_anchor
    sheet: Engine
    data_range: Engine!C5
    layout: scalar
    constant: {}
    structure:
      measure:
        concept: OBS_VALUE
        dtype: float
        bind: {kind: data_cell, read: float}
      dimensions: []
    key: []
```

| Rule | Detail |
|---|---|
| Leaf-only | `data_range` must intersect graph **leaves** (inverse of `internal`, which requires formula nodes). Non-leaf overlap → `non_leaf_constant_overlap`; no leaf overlap → `no_leaf_constant_targets`. |
| Exclusive | Mutually exclusive with `input`, `output`, and `internal` on the same series. |
| Codegen | Names the leaf for inverted-tree export: values land in `data.py` and as defaulted `{Output}Inputs.from_defaults(...)` fields. Optional `constant.reader.name` is gone; inverted-tree does not emit public readers or extra `compute_*` for constants. |
| Mutability | Values still live in `data.py` and as Inputs defaults; there is no public write surface. |
| Validate | `validate_series_bindings(...)`, then `derive_constant_series(...)`. Include `constant` when running `scripts.binding_resolution_audit`. |

**Leaf classification vs constant bindings.** `CONSTRAINTS` with a single-value
`Literal[...]` classifies a leaf as `constant` for codegen `CONSTANTS` vs
`DEFAULT_INPUTS`. Fail closed: every `constant` leaf must appear in
`constants.bindings.yaml`, and every mutable `input` leaf in
`inputs.bindings.yaml`. The configure tests assert an empty unbound list even
when a workbook has no constant leaves (do not skip-if-empty). A `constant: {}`
series also emits a semantic `read_*` so formulas do not keep bare `xl_cell`.

**Structural blanks are not constants.** Padding inside `INDEX`/`MATCH` arrays,
far-right `NPV`/`SUM` overflow, unused ladder copies, and separator rows belong
in `workbook_config.BLANK_RANGES`, not `CONSTRAINTS` + `constants.bindings.yaml`.
`Literal[None]` classifies a still-present leaf; it does not omit the node.

**Do not confuse** the `constant` **direction** with `bind.kind: constant` (a
fixed dimension / attribute scalar in `structure`). The direction binds a
spreadsheet leaf; the bind kind fills a coordinate without reading a cell.

Synthetic example: [tests/fixtures/synthetic/constants.bindings.yaml](../tests/fixtures/synthetic/constants.bindings.yaml)
(`Inputs!B1` bias leaf). Schema reference: excel-grapher
`user_guide/05-series-bindings.qmd` (constant direction, schema 1.11.0+).

## Output compute helpers

When an internals helper covers a published output series' leaves, declare
`output.compute.helper` so generated `compute_*` calls the helper from record
dims instead of `xl_cell(address)` (excel-grapher schema 1.13.0+):

```yaml
output:
  compute:
    name: compute_scenario_primary_expenditure_pct_gdp
    helper:
      name: scenario_primary_expenditure_pct_gdp_hot
      dims: [TIME_PERIOD]
```

`dims` defaults to the series `key` when omitted. Leaves without helper coverage
still use `xl_cell`. Declare helper blocks on every output series id that matches a same-named function in the generated
`internals.py`.

After authoring (or when export fails in codegen), run
`uv run python -m scripts.binding_resolution_audit`
([issue #101](https://github.com/Teal-Insights/extraction-pipeline-template/issues/101))
to catch bind-resolution errors that `validate_series_bindings` /
`derive_*_series` can miss — for example sparse year headers without
`fill: true`, or output series that resolve only a subset of their `data_range`.
Codegen requires resolution `ok=True`; validation alone does not. This is a
correctness audit of authored bindings; burndown
(`uv run python -m scripts.internal_binding_burndown`) is the coverage worklist for
cells that still lack a binding.

## Sparse labels, measure columns, and graph targets

Summary tables often place a year (or other group label) only on the first column
of a repeating measure triplet, with blank cells under the remaining measures:

```text
header row:  2050   (blank) (blank)  2075   (blank) (blank)  ...
measure row: Base   Alt     Gap      Base   Alt     Gap      ...
data:        ...    ...     ...      ...    ...     ...      ...
```

That layout trips three authoring mistakes that pass `validate_series_bindings`
but fail at output/input codegen:

1. **Sparse `column_header` / `row_label` cells need `fill: true`.**
   Without `fill`, blank label cells raise `bind_resolution_failed` for every
   column/row that has no source label, even when the contiguous `data_range`
   looked fine. Set `fill: true` when labels appear only on the first
   column/row of a group and should propagate across the blank span.

2. **Do not bind one contiguous rectangle across mixed measures** when the
   public series key is only something like `(SCENARIO, TIME_PERIOD)`.
   Columns from different measures then collide on the same key, or only a
   subset of columns resolve. Prefer either:
   - **one shard per measure column** (or per milestone column); or
   - a **richer key** that includes the measure dimension (common for
     internals that triangulate the whole triplet table).

   When you shard, choose `output.compute.name` / input series `id`
   deliberately:

   | Choice | When | Effect |
   |---|---|---|
   | **Share** the same name / id across shards | Shards are complementary slices of one logical public series (e.g. Gap columns for 2050 / 2075 that should become one `compute_gap_milestones`) | Export merges shards into one public function |
   | **Uniquify** per shard | Each shard is a distinct scenario / engine path (e.g. Paris vs Moderate expenditure rows on separate sheets) | Each path keeps its own `output.compute.name` / input series id |

   Complementary input shards share a series `id` (schema 1.14.0+); `input: {}`
   marks them as editable leaves. Extra `setter` / `reader` keys are rejected.
   Callers pass those series as `{Output}Inputs.from_defaults(...)` fields into
   `compute_*`.

   Sharing a name across distinct engine paths is the failure mode: export
   merges the colliding definitions, so most scenario paths become
   **unreachable** in the exported library even though every shard still
   looks valid in YAML. Share only when a merge is intentional.

3. **`workbook_config.TARGETS` (and overlapping internal ranges) must cover
   every bound data cell.** Sharding a “gap-only” column that sits past the
   previous target end requires widening the target (and any internal
   `data_range` that must stay on-graph for those cells). Bound cells outside
   the extracted graph never resolve cleanly.

Pedagogical catalog fragment (not used by the synthetic smoke workbook):
[.agents/skills/author-bindings/assets/measure-shards.example.yaml](../.agents/skills/author-bindings/assets/measure-shards.example.yaml).
The example shows the intentional **shared-name** merge for milestone Gap
columns; see its header comments for the **unique-name** alternative used
for per-scenario engine shards.

## Dimension `id` vs `concept`

| Field | Role |
|---|---|
| `concept` | SDMX-style meaning category (e.g. `TIME_PERIOD`) |
| `id` | Dimension identity used in records, cell keys, and generated helper parameter names |

Give every dimension an explicit `id`. When `id` is omitted, the effective id falls back to `concept`. If two dimensions share a concept, they must have distinct ids:

```yaml
dimensions:
  - id: PROJECTION_PERIOD
    concept: TIME_PERIOD
    role: key
    scope: cell
    bind:
      kind: column_header
      header_row: 5
      read: int
  - id: REFERENCE_PERIOD
    concept: TIME_PERIOD
    role: key
    scope: cell
    bind:
      kind: value_map
      values:
        0: C:G
      read: int
key: [PROJECTION_PERIOD, REFERENCE_PERIOD]
```

Effective ids drive parameter names (`projection_period`, `reference_period`). Concepts remain semantic metadata for documentation and concept-scheme dtype inheritance.
