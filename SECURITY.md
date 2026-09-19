# Security Policy

## Supported version

Security fixes should target the latest portable release:

| Version | Supported |
|---|---|
| 3.1.x | ✅ |
| 3.0.x and older | ❌ Legacy / archived |

## Reporting a vulnerability

Please avoid publishing sensitive corporate information, recipient addresses, internal file paths or real customer documents in a public issue.

For a security-sensitive report, use GitHub's private vulnerability reporting feature if it is enabled for the repository.

## Enterprise controls

Outlook Part Sender is not designed to bypass endpoint security, Outlook security prompts, DLP, antivirus, application allow-lists, group policy, message-size controls or tenant policies.

If your organization blocks the launcher, local compilation, `%LOCALAPPDATA%` executables or Outlook COM automation, request approval from your IT/security team.

## Credentials

The application uses the user's already authenticated Classic Outlook session. It does not ask for, store or transmit an Outlook password.

## Saved `.opsparts.json` layouts

Saved layouts may contain recipient addresses, subject/body text and local file paths. Treat those files as potentially sensitive business data.

## Sending

Automatic Send is opt-in and requires an additional confirmation. Draft creation is the default mode and remains the recommended workflow.
