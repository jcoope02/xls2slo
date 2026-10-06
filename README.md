# xls2slo

Create Nobl9 SLO YAML from an Excel spreadsheet. Each row defines one objective; rows sharing the same **project + SLO name** become one SLO.

The included sample contains **fictional data only**: four objective rows create two SLOs, each containing two objectives. Its project, service and data-source references are placeholders and must be replaced before deployment.

## Quick start

Requires Python 3.9 or later.

```bash
git clone https://github.com/jcoope02/xls2slo.git
cd xls2slo
python3 -m venv .venv
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
python3 -m pip install -r requirements.txt
python3 slo_sheet_to_yaml.py examples/sample_slos.xlsx --output-dir generated
```

This produces:

- `generated/demo-commerce__demo-store-latency.yaml`
- `generated/demo-commerce__demo-store-reliability.yaml`

The script creates files locally. It does not call Splunk or Nobl9, populate the sheet through MCP, or apply SLOs.

## Spreadsheet structure

**Each row represents one objective.** The `slo_name` column determines which SLO it belongs to, within the specified `project`. Rows with the same project and SLO name are grouped into one YAML SLO, with each row added to `spec.objectives`.

For example:

| project | slo_name | objective_name |
| --- | --- | --- |
| demo-commerce | demo-store-latency | catalog-items-get-p95 |
| demo-commerce | demo-store-latency | cart-post-p95 |
| demo-commerce | demo-store-reliability | catalog-items-get-success |

These three rows create **two SLOs**: latency with two objectives, and reliability with one. This shortened example illustrates grouping; the sample workbook includes four rows, giving each SLO two objectives. The same SLO name in a different project belongs to a separate SLO.

Edit the **SLO Template** worksheet. The **Instructions** and **Field map** worksheets explain the columns. Add rows to define additional objectives.

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

Shared SLO fields must agree within a group. Objective names and values must be unique within each SLO. Keep `target` as a numeric fraction such as `0.999`; Excel displays this as `99.9%`. Use `|` to separate multiple values in a label cell. Additional `label_<key>` columns are supported.

For raw metrics, `objective_value` is the threshold in source metric units. No unit conversion is performed. The sample’s latency values are illustrative; verify your Splunk metric units and intended thresholds. For count metrics, `objective_value` is the legacy objective identifier, not a reliability threshold.

## Automatic query generation

For `rawMetric` rows, the script generates SignalFlow from the endpoint, method, environment, service and operation columns. `splunk_operation` must equal `http_method + space
