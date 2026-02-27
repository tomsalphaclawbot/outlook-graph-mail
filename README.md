# outlook-graph-mail

OpenClaw skill for full Outlook mailbox operations via Microsoft Graph (modern auth): read, send, archive, delete, move, folder listing, and junk/spam checks.

## What's inside
- `SKILL.md`
- `scripts/outlook-graph-auth.sh`
- `scripts/outlook-graph-read.sh`
- `scripts/outlook-graph-read-folder.sh`
- `scripts/outlook-graph-list-folders.sh`
- `scripts/outlook-graph-send.sh`
- `scripts/outlook-graph-archive.sh`
- `scripts/outlook-graph-delete.sh`
- `scripts/outlook-graph-move.sh`

## Install (manual)
Clone into your OpenClaw workspace skills directory:

```bash
git clone https://github.com/tomsalphaclawbot/outlook-graph-mail.git \
  ~/.openclaw/workspace/skills/outlook-graph-mail
```

## Usage (from workspace root)
```bash
skills/outlook-graph-mail/scripts/outlook-graph-auth.sh
skills/outlook-graph-mail/scripts/outlook-graph-read.sh 20
skills/outlook-graph-mail/scripts/outlook-graph-read-folder.sh junkemail 20
skills/outlook-graph-mail/scripts/outlook-graph-send.sh "to@example.com" "Subject" "Body"
```

## Notes
- Uses OAuth2 device flow against Microsoft Graph.
- Default token file: `state/outlook_graph_token.json`.
