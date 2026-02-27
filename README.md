# outlook-graph-mail

OpenClaw skill for Outlook mailbox management via Microsoft Graph (modern auth).

## Features
- OAuth device-code auth
- Inbox/folder reads (including junk/spam)
- Send email
- Archive/delete/move messages
- Token keepalive for unattended automation

## Install

```bash
git clone https://github.com/tomsalphaclawbot/outlook-graph-mail.git ~/.openclaw/workspace/skills/outlook-graph-mail
```

## Quick start

```bash
cd ~/.openclaw/workspace
skills/outlook-graph-mail/scripts/outlook-graph-auth.sh
skills/outlook-graph-mail/scripts/outlook-graph-read.sh 20
skills/outlook-graph-mail/scripts/outlook-graph-read-folder.sh junkemail 20
skills/outlook-graph-mail/scripts/outlook-graph-send.sh "to@example.com" "Subject" "Body"
```

## Keepalive cron (optional)

```bash
*/30 * * * * cd ~/.openclaw/workspace && skills/outlook-graph-mail/scripts/outlook-graph-keepalive.sh >> state/outlook-graph-keepalive.log 2>&1
```
