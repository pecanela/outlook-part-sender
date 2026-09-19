# Contributing

Thanks for helping improve Outlook Part Sender.

## Recommended development target

Changes should be based on the current portable architecture in:

```text
src/OutlookPartSenderPortable.cs
launcher/
```

The older PowerShell, VSTO and COM add-in branches are historical architectures.

## Pull-request checklist

- Keep Drafts as the default.
- Do not introduce automatic sending without explicit confirmation.
- Do not bypass Outlook or enterprise security controls.
- Preserve dynamic part numbering.
- Preserve multiple attachments per Part.
- Validate recipients before a batch.
- Avoid storing credentials.
- Document changes in `CHANGELOG.md`.
- Update relevant docs and release notes.
- Test on Classic Outlook for Windows.

## Style

The current implementation targets the Windows .NET Framework compiler and deliberately avoids dependencies that would require NuGet, Visual Studio or a separate runtime on the end user's PC.

When proposing a new dependency, explain why the extra deployment complexity is justified.
