# Excel parity validation

This folder ships the exported-library differential test, the illustrative
workbook fixture, and reference parity reports produced in a maintainer
Windows environment with Microsoft Excel installed.

## Reference results

`results/reference/` contains the last committed parity report from the
extraction pipeline. These reports document that the exported `tiny_dsa`
package matched Excel across 118 scenarios and 15 output cells at
`atol = 1e-6`.

## Re-run locally (Windows + Excel only)

The test drives Excel through `xlwings` and cannot run in Linux CI.

```pwsh
uv run --project . --group validation python tests/differential_test_exported_library.py --layout exported
```

Local reruns write to `results/local/` by default. To refresh the shipped
reference reports after a passing run:

```pwsh
uv run --project . --group validation python tests/differential_test_exported_library.py --layout exported --report-dir tests/results/reference
```
