# AWS Bedrock Cost Report

A Python portfolio demo that turns AWS Cost Explorer service costs into an Excel workbook and a machine-readable JSON snapshot. The supplied example uses fictional accounts and costs. It is not evidence of a production deployment or real customer spending.

The report selects service names containing `Bedrock`, `Claude` or `Anthropic`, then groups monthly `UnblendedCost` by service for each selected linked account. Keywords are configurable. Service labels are whatever Cost Explorer returns, not a guaranteed model taxonomy.

## Scope

One run uses one named AWS profile and its authenticated billing context. `LINKED_ACCOUNT` filters select accounts visible within that context. Providing another independent payer account ID does not grant access to its bills. This tool does not assume roles across payer accounts.

This is a selected-service cost report, not token usage, model-level attribution, a Bedrock inventory or invoice reconciliation. It does not split `BILLING_ENTITY` or infer direct versus Marketplace charges from service names. A matching service name can include costs for models beyond Claude. A nonmatching name can be missed. Review the discovered service labels in the output.

## Try it without AWS

Python 3.10 or later:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements-test.txt
python -B -m unittest discover -s tests -v
python usage_report.py --demo --chart --output reports/synthetic-demo
```

The demo reads only `fixtures/synthetic_ce.json`. It never creates an AWS client. Its two complete calendar months contain six service-month records, including zero cost, a negative adjustment and an estimated cost. The observed total is 36.00 USD. `Synthetic Claude Service` is a fictional label, not an assertion about an actual AWS service.

## Read your billing data

Install `requirements.txt`, then configure an AWS CLI named profile outside this repository. Prefer short-lived SSO credentials. Do not put keys in code or input files. Running the following command makes billed AWS Cost Explorer API requests:

```bash
python usage_report.py \
  --accounts linked_accounts.csv \
  --profile billing-readonly \
  --months 3 \
  --keywords "Bedrock,Claude,Anthropic" \
  --chart \
  --output reports/selected-service-costs
```

The CSV must contain one column named `linked_account_id`, with one 12-digit identifier per row. The example CSV contains fictional identifiers. XLSX input is also supported: the active sheet must have that same single header, and identifiers must be stored as text. Numeric, duplicate and malformed IDs are rejected rather than padded or silently dropped.

`--months` accepts 1 to 13 and defaults to 3. The default end is the first day of the current month, exclusive, so the report excludes the partial current month. `--end-month 2026-02` includes February as the last month. Future or current-month endpoints are rejected. Date arithmetic follows calendar boundaries, including leap years. Available history still depends on your Cost Explorer configuration and access, so the application limit is not an availability guarantee.

The principal needs authorized billing access and the read actions `ce:GetDimensionValues` and `ce:GetCostAndUsage`. Access rules differ between management and member accounts. The tool uses `us-east-1` by default for the Cost Explorer endpoint, with an optional `--region`. It changes no AWS resources, but reads can incur charges and reports contain sensitive billing information.

## Output and incomplete data

Each run creates a new `.xlsx` and `.json` pair. Existing files are never overwritten.

| Worksheet | Content |
| --- | --- |
| Summary | Observed cost by linked account and unit, estimated flag and coverage status |
| Monthly Trend | Monthly observed cost by linked account and unit |
| Service Breakdown | Observed cost by linked account, service and unit |
| Coverage | A status for every requested account-month, including failures and missing records |

The optional chart is a line chart inside Excel. It is included only when coverage is complete and costs have one unit. It is not a PNG export. Report cells are a static snapshot; rerun the program to refresh them. JSON retains decimal amounts as strings, while Excel displays numeric costs to two decimal places and has normal spreadsheet precision limits.

Both API operations paginate. Repeated tokens, duplicate groups and malformed responses fail closed. If a later page fails, the account's partial pages are discarded. A missing month, empty grouped response or no matching services is unavailable data, not an invented zero. An explicit zero returned in a cost group remains zero. Negative adjustments remain negative, and different cost units are never added together. `Estimated` is retained, including for completed months that AWS has not finalized.

Exit status is `0` for complete retrieval, `2` for an exported but incomplete report, or `1` for invalid inputs/setup/output failure. Incomplete reports show observed partial amounts and Coverage warnings. Check those warnings before relying on totals. A complete retrieval only means all requested account-months returned valid cost groups; it does not prove keyword coverage, billing permissions, invoice agreement or cost finality.

## Validation and limits

The offline regression suite uses mocked Cost Explorer responses, blocks socket connections and writes generated files only in temporary directories. It covers calendar boundaries, identifiers, pagination, aggregation, missing data, late-page failures, currency separation, literal spreadsheet text, chart creation and CLI exit codes. The included GitHub Actions workflow is newly added validation-only CI, with no AWS credentials or deployment steps.

No live AWS billing account was queried to validate this public demo. Dependency ranges are intentionally small, but not a complete production dependency lock or security guarantee. Large billing estates, resumable jobs, concurrency, Cost and Usage Reports, tax treatment and cross-payer role orchestration are outside this package.

API references: [GetCostAndUsage](https://docs.aws.amazon.com/aws-cost-management/latest/APIReference/API_GetCostAndUsage.html), [GetDimensionValues](https://docs.aws.amazon.com/aws-cost-management/latest/APIReference/API_GetDimensionValues.html), [Cost Explorer pricing](https://aws.amazon.com/aws-cost-management/aws-cost-explorer/pricing/).

No license grant is included. The project owner must choose and approve a license before one is added.
