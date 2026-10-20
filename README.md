# Workspace Lakehouse and Warehouse Audit

This repository contains a reproducible notebook and Python module for reviewing the supplied `Workspace_LH_and_DW_Audit.xlsx` workbook.

The action sheet is detected by its `Keep` header. Rows with a blank Keep value are ignored. Rows marked `NO` are joined to the workbook's `lineage view check` sheet. Dependency evidence is written to column G of the copied workbook before any deletion decision is made.

Deletion eligibility is deliberately narrow:

- the action row is marked `NO`;
- the lineage row has no direct or indirect dependency evidence, including shortcuts, reports, semantic models, dataflows, notebooks, pipelines, or non-zero impact counts;
- the row is not already marked `DELETED`; and
- the item and workspace resolve exactly in the live Fabric catalog.

The notebook defaults to a dry run. Live deletion requires both `FABRIC_ALLOW_DELETIONS=1` and `FABRIC_DELETE_CONFIRMATION=DELETE_APPROVED_NO_DEPENDENCY_ITEMS`. It uses Azure CLI/managed identity/default Azure credentials without storing tokens. The source workbook is never changed; the copied workbook under `data/` receives the column-G annotations.

The supplied workbook names `ws_prod_legacy`, but the current Fabric catalog must be queried before deletion. If that workspace cannot be resolved exactly, the notebook records `BLOCKED_NO_LIVE_WORKSPACE` and does not guess a replacement workspace.

## Run

```bash
python -m venv .venv
. .venv/bin/activate
pip install -r requirements.txt
jupyter notebook notebooks/workspace_lh_dw_cleanup.ipynb
```

For a command-line dry run:

```bash
PYTHONPATH=src python -m audit_cleanup --input data/Workspace_LH_and_DW_Audit.xlsx --output data/Workspace_LH_and_DW_Audit_annotated.xlsx
```

For a live run, first inspect the dry-run output and verify the exact Fabric workspace/item identities. Only then set the two deletion environment variables and rerun with `--execute-deletions`.
