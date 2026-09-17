# Security Policy

## Reporting a vulnerability

Report security issues privately to the repository owner. Do not open a public issue containing credentials, tenant identifiers, device data, function keys, tokens, or SharePoint list contents.

## Credential handling

- Never commit `local.settings.json`, `.env`, client secrets, certificates, access tokens, function keys, or storage connection strings.
- Prefer certificate authentication or Azure Key Vault references.
- Rotate any credential disclosed in logs, screenshots, chat, tickets, or documentation.
- Use Microsoft Graph `Sites.Selected` and grant the Function app access only to the reporting site.

## Supported deployment

Review Microsoft Graph permissions, SharePoint site grants, OS-build thresholds, and risk rules before each production deployment.
