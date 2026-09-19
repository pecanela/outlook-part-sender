# Troubleshooting

## The `.cmd` opens and closes immediately

Open Command Prompt, drag `ABRIR Outlook Part Sender.cmd` into the window and press Enter. Read the error shown.

Also inspect:

```text
%LOCALAPPDATA%\OutlookPartSenderPortable\compile.log
```

## `.NET Framework 4` not found

The launcher checks:

```text
%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe
```

If neither exists, the Windows/.NET configuration or an enterprise policy may be preventing local compilation.

## Outlook is not detected

Use **Classic Outlook for Windows**, configured and signed in. New Outlook is not supported.

## The application is blocked by the company PC

Do not try to bypass the restriction. Ask IT/security to approve the source/launcher or provide an allowed deployment method.

## Wrong signature or signature formatting

Generate Drafts first and inspect the messages. Signature behavior depends on Outlook profile/account policies.

## Recipient cannot be resolved

Use an address that Outlook can resolve normally. The application intentionally stops when `ResolveAll()` fails.

## Message is too large

Reduce attachments in that Part. The app's size is an estimate; the recipient's server may enforce a smaller limit.

## Remove local cache

Run:

```text
LIMPAR CACHE LOCAL.cmd
```

Then the next launcher run performs a clean local compilation.
