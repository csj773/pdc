# GitHub Actions Automation

This directory can run the same logbook update in GitHub Actions.

## Required repository secrets

- `GOOGLE_SERVICE_ACCOUNT_JSON_B64` preferred, or `GOOGLE_SERVICE_ACCOUNT_JSON`.
- `GMAIL_USERNAME`: Gmail account used to send the completed files.
- `GMAIL_APP_PASSWORD`: Gmail app password for SMTP sending.

Share both Google Sheets with the service account email from the JSON credentials:

- authoritative `log filled`: `1tRvMJQeoqpGvekJ3xzs_Z80e9QnXoGsIEdIuchY7Wqw`
- upstream `PILOTLOG`: `1mKjEd__zIoMJaa6CLmDE-wALGhtlG-USLTAiQBZnioc`

To create the base64 secret locally:

```sh
base64 -i service-account.json | pbcopy
```

## Workflow

`.github/workflows/update-pilot-logbook.yml` runs every 10 days at `15:00 UTC`, which is midnight in Korea during standard time, and can also be started manually with `workflow_dispatch`.

The workflow:

1. exports both Google Sheets files to XLSX,
2. ensures the authoritative tab is named `flt_log`,
3. writes `work/PILOTLOG_export.xlsx`,
4. runs `work/sync_authoritative_from_pilotlog.py`,
5. updates the authoritative Google Sheet from `work/log_filled_authoritative_synced.xlsx`,
6. runs `work/build_final_deliverables.py`,
7. emails exactly the three final deliverables,
8. uploads the same files as a GitHub Actions artifact.
