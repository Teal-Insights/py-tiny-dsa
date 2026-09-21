# Tiny DSA graph-oracle parity validation

This folder ships the exported-library differential test, the workbook fixture,
and reference parity reports produced against excel-grapher's FormulaEvaluator.

## Reference results

`results/reference/` contains the last committed parity report from the
extraction pipeline. These reports document that the exported `tiny_dsa`
package matched the extraction graph for the configured scenario sweep.

## Re-run from the extraction repository

The test compares keyword-only `compute_*` results to FormulaEvaluator. It does
not drive Microsoft Excel.

```pwsh
uv run python -m tests.differential.differential_test_exported_library
```

Local reruns from this exported project write to `results/local/` when the
extraction graph cache and `excel-grapher` are available:

```pwsh
uv run python -m tests.differential.differential_test_exported_library --layout exported
```
