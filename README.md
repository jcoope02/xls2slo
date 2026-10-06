# xls2slo

Create Nobl9 SLO YAML from an Excel spreadsheet. Each row defines one objective;
rows sharing the same **project + SLO name** become one SLO.

The included sample contains **fictional data only**: four objective rows create
two SLOs, each containing two objectives. Its project, service and data-source
references are placeholders and must be replaced before deployment.

## Quick start

Requires Python 3.9 or later.

```bash
git clone https://github.com/jcoope02/xls2slo.git
cd xls2slo
python -m venv .venv
```

Activate the virtual environment on macOS/Linux:

```bash
source .venv/bin/activate
```

Or on Windows PowerShell:

```powershell
.venv\Scripts\Activate.ps1
```

Install dependencies and generate the sample YAML:

```bash
python -m pip install -r requirements.txt
python slo_sheet_to_yaml.py examples/sample_slos.xlsx --output-dir generated
```

This produces:

- `generated/demo-commerce__demo-store-latency.yaml`
- `generated/demo-commerce__demo-store-reliability.yaml`

The script creates files locally. It does not call Splunk or Nobl9, populate the
sheet through MCP, or apply SLOs.

## Spreadsheet structure

**Each row represents one objective.** The `slo_name` column determines which
SLO it belongs to, within the specified `project`. Rows with the same project
and SLO name are grouped into one YAML SLO, with each row added to
`spec.objectives`.

For example:

| project | slo_name | objective_name |
| --- | --- | --- |
| demo-commerce | demo-store-latency | catalog-items-get-p95 |
| demo-commerce | demo-store-latency | cart-post-p95 |
| demo-commerce | demo-store-reliability | catalog-items-get-success |

These three rows create **two SLOs**: latency with two objectives, and
reliability with one. This shortened example illustrates grouping; the sample
workbook includes four rows, giving each SLO two objectives. The same SLO name
in a different project belongs to a separate SLO.

Edit the **SLO Template** worksheet. The **Instructions** and **Field map**
worksheets explain the columns. Add rows to define additional objectives.

| Columns | Purpose |
| --- | --- |
| `endpoint`, `http_method` | Endpoint in column A and its HTTP method |
| `project`, `slo_name` | Grouping key for one SLO |
| `slo_display_name`, `service`, `description`, `label_*` | SLO identity and optional descriptive fields |
| `data_source_name`, `data_source_project`, `data_source_kind` | Existing Nobl9 data-source reference |
| `budgeting_method`, `window_unit`, `window_count`, `window_is_rolling` | Shared SLO configuration |
| `objective_name`, `objective_display_name`, `target`, `objective_value` | Objective settings |
| `metric_type`, `operator`, `incremental` | Select raw or count metrics and their settings |
| `splunk_environment`, `splunk_service`, `splunk_operation` | Splunk query context |

Shared SLO fields must agree within a group. Objective names and values must
be unique within each SLO. Keep `target` as a numeric fraction such as `0.999`;
Excel displays this as `99.9%`. Use `|` to separate multiple values in a label
cell. Additional `label_<key>` columns are supported.

For raw metrics, `objective_value` is the threshold in source metric units.
No unit conversion is performed. The sample's latency values are illustrative;
verify your Splunk metric units and intended thresholds. For count metrics,
`objective_value` is the legacy objective identifier, not a reliability threshold.

## Automatic query generation

For `rawMetric` rows, the script generates SignalFlow from the endpoint, method,
environment, service and operation columns. `splunk_operation` must equal
`http_method + space + endpoint`.

The current query template uses:

- The `spans` histogram and p95.
- `SERVER` and `CONSUMER` spans with `sf_error=false`.
- Exclusions for dimensionalized and service-mesh spans.

The sample contains no program columns. All latency and reliability queries are
generated from the context inputs. Advanced users may add optional `raw_program`,
`good_program`, and `total_program` columns to compare existing queries against
the generated ones. These references are never copied into output. A nonblank
reference must match by default; for intentional differences, update/clear
the optional reference or explicitly allow differences:

```bash
python slo_sheet_to_yaml.py examples/sample_slos.xlsx \
  --output-dir generated --overwrite --allow-reference-differences
```

Use `--latency-percentile 99` to generate p99 for all raw metric rows; also update
objective names, thresholds and references as appropriate. The generator does
not infer the percentile from objective names.

For `countMetrics`, the script automatically generates both programs using
`data('spans.count', ...).sum()`:

- **Good:** count matching spans with `sf_error=false`, published as `good`.
- **Total:** count all matching spans without the error filter, published as `total`.

Both use the context inputs, the same `SERVER`/`CONSUMER` span-kind filter, and
the same dimensionalized/service-mesh exclusions. `incremental` must be TRUE
or FALSE. The sample uses FALSE. The latency percentile option does not affect
reliability queries.

## Compare with existing YAML

You can compare generated files against local desired-state exports:

```bash
python slo_sheet_to_yaml.py examples/sample_slos.xlsx \
  --output-dir generated --overwrite \
  --compare-original baseline-latency.yaml baseline-reliability.yaml \
  --report comparison.txt
```

Supply your own baseline files; they are not included. Comparison is based on
parsed configuration, rather than YAML formatting. It excludes `status`,
`spec.createdAt`, `spec.createdBy`, and exported rolling-window `period` dates.
An empty optional description is equivalent to an omitted description. All
remaining configuration and objective ordering are compared.

Exit codes: **0** for success, **1** for a baseline comparison mismatch, and
**2** for input or validation errors. Existing output files require `--overwrite`.

## Scope and validation

This version supports Splunk Observability, Occurrences budgeting, one rolling
window, and objective-level `rawMetric` or `countMetrics`. Calendar windows,
Timeslices, composites and other sources need additional template/parser support.

Validation checks headers, required fields, names, numeric and boolean values,
group consistency, objective uniqueness, metric branches, context consistency and
reference-program consistency when an optional reference is supplied. Formula cells are rejected; use literal values.
Validation is offline and does not confirm that referenced projects, services
or data sources exist in Nobl9, or that Splunk queries execute successfully.

Generated YAML is intended for review before deployment. Applying a definition
with the same project and SLO name addresses that existing SLO; choose distinct
names when creating additional SLOs.

Schema reference: [Nobl9 YAML guide](https://docs.nobl9.com/yaml-guide/).
