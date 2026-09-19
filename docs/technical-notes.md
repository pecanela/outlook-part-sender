# Technical Notes

## Target

- Windows
- Classic Outlook
- .NET Framework 4.x compiler (`csc.exe`)
- WinForms
- Outlook COM late binding

## Entry point

The portable release uses `OutlookPartSender.PortableProgram` as the executable entry point. The launcher explicitly passes:

```text
/main:OutlookPartSender.PortableProgram
```

## Launcher behavior

`ABRIR Outlook Part Sender.cmd`:

1. defines the current package version;
2. finds the included `.cs` source;
3. creates `%LOCALAPPDATA%\OutlookPartSenderPortable`;
4. locates 64-bit or 32-bit .NET Framework `csc.exe`;
5. checks cached `version.txt`;
6. recompiles only when the cached version differs;
7. starts the resulting WinForms executable.

## No Outlook registration

v3.1.0 intentionally does not:

- register a COM add-in;
- write Outlook add-in registry keys;
- inject or overwrite VBA;
- deploy VSTO manifests.

## Batch safety

The application creates a snapshot before generation. During the batch, controls are disabled.

Recipient validation happens twice:

- before the batch;
- on every generated message.

The source contains one intended automatic Send path. Draft creation remains the default.

## Message-size estimate

The estimate is intentionally conservative and accounts for Base64/MIME overhead. It is still only an estimate; the authoritative limit is determined by the mail systems involved.

## COM lifecycle

Important Outlook objects are explicitly released where practical to reduce orphaned Outlook processes and long-lived RCWs.

## Signature

When enabled, the application displays the item so Outlook can load the configured signature, then uses `Inspector.WordEditor` to insert the user message before it.
