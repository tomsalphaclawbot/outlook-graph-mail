---
name: outlook-graph-mail
description: Manage Outlook/Hotmail mailbox via Microsoft Graph OAuth2 (modern auth) for read, send, archive, delete, move, spam/junk checks, and folder operations. Use when working with tomsalphaclaw@outlook.com mail workflows, Cloudflare routing probes, inbox triage, or any Outlook automation that must avoid basic IMAP/POP/SMTP auth.
---

# Outlook Graph Mail

Use bundled scripts for full mailbox management with modern auth.

## Scripts

- `scripts/outlook-graph-auth.sh`
  - Run device-code auth and store token at `state/outlook_graph_token.json`.
- `scripts/outlook-graph-read.sh [top] [subject_contains]`
  - Read recent inbox messages (returns message IDs).
- `scripts/outlook-graph-read-folder.sh [folder] [top] [subject_contains]`
  - Read a specific folder (`inbox`, `junkemail`, `archive`, `deleteditems`, etc.).
- `scripts/outlook-graph-list-folders.sh`
  - List mailbox folders with counts.
- `scripts/outlook-graph-send.sh <to> <subject> <body>`
  - Send email via Graph `sendMail`.
- `scripts/outlook-graph-archive.sh <message_id>`
  - Move message to Archive.
- `scripts/outlook-graph-delete.sh <message_id>`
  - Delete message.
- `scripts/outlook-graph-move.sh <message_id> <destination>`
  - Move message to any folder.
- `scripts/outlook-graph-keepalive.sh`
  - Refresh/validate token and keep session warm for unattended automation.

## Operator quickstart

Run from workspace root:

```bash
# 1) Authenticate (device code)
skills/outlook-graph-mail/scripts/outlook-graph-auth.sh

# 2) Read latest inbox mail (with message IDs)
skills/outlook-graph-mail/scripts/outlook-graph-read.sh 20

# 3) Check spam/junk
skills/outlook-graph-mail/scripts/outlook-graph-read-folder.sh junkemail 20

# 4) Send email
skills/outlook-graph-mail/scripts/outlook-graph-send.sh "someone@example.com" "Subject" "Body text"

# 5) Archive/Delete/Move by message ID
skills/outlook-graph-mail/scripts/outlook-graph-archive.sh "<message_id>"
skills/outlook-graph-mail/scripts/outlook-graph-delete.sh "<message_id>"
skills/outlook-graph-mail/scripts/outlook-graph-move.sh "<message_id>" archive

# 6) Keep token fresh for unattended runs
skills/outlook-graph-mail/scripts/outlook-graph-keepalive.sh
```

## Default config

- Client ID default: `4a8603ea-1f64-4963-85bc-27c5758fd6f3`
- Tenant default: `consumers`
- Token file default: `state/outlook_graph_token.json`
- Scopes include: `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`, `offline_access`, `openid`, `profile`

Override with env vars:
- `OUTLOOK_GRAPH_CLIENT_ID`
- `OUTLOOK_GRAPH_TENANT`
- `OUTLOOK_GRAPH_TOKEN_FILE`
- `OUTLOOK_GRAPH_SCOPES`

## Operational notes

- Prefer Graph scripts over basic IMAP/POP/SMTP auth for Outlook.
- Scripts auto-refresh access tokens (refresh_token + expiry-aware logic) and retry once on 401.
- For spam checks, use `outlook-graph-read-folder.sh junkemail ...`.
- For deterministic triage: read -> act by `message_id` -> verify folder counts.

### Optional cron keepalive
Run every 30 minutes to keep tokens warm for unattended automations:

```bash
*/30 * * * * cd ~/.openclaw/workspace && skills/outlook-graph-mail/scripts/outlook-graph-keepalive.sh >> state/outlook-graph-keepalive.log 2>&1
```
