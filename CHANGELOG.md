# Changelog

All notable project milestones are documented here.

## [3.1.0] - 2026-09-11 — Portable Corporate — **Latest**

- Replaced the installed COM add-in approach with a fully portable workflow.
- No Outlook add-in registration, no VBA changes, no VSTO and no administrator requirement.
- Added a lightweight launcher that compiles the included C# source with the Windows .NET Framework compiler.
- Reuses a versioned local executable cache in `%LOCALAPPDATA%\OutlookPartSenderPortable`.
- Launcher uses no PowerShell.
- Preserved dynamic parts, multiple attachments per part, account selection, Outlook signature, review, JSON layouts, size estimates, recipient resolution and Draft/Send protections.
- Designed specifically to reduce friction on managed corporate PCs.
- Confirmed working successfully in a real corporate Windows + Classic Outlook environment.

## [3.0.0] - 2026-09-11 — One Click COM Add-in

- Experimental COM add-in architecture.
- Added an Outlook Ribbon tab and `Enviar em Partes` button.
- Compiled locally using .NET Framework, registered per-user in HKCU.
- No Visual Studio/VSTO/admin required.
- Added diagnostics and uninstall workflow.
- Ultimately replaced because registry/add-in installation is less suitable for restricted corporate PCs.

## [2.0.2] - 2026-09-11 — VSTO hardened

- Fixed ClickOnce certificate thumbprint flow before Rebuild/Publish.
- Added safer exception handling for layout save/load and file selection.
- Preserved selected account on account refresh.
- Hardened Outlook shutdown handling.
- Increased WordEditor/signature initialization wait.
- Improved Visual Studio/VSTO tooling selection when multiple VS installations exist.
- Final archived VSTO source branch.

## [2.0.1] - 2026-09-11 — VSTO safety pass

- Added immutable batch snapshot.
- Locked all interactive UI while processing.
- Added fail-safe return to Drafts after automatic sending.
- Forced loaded layouts back to Drafts.
- Improved COM cleanup.
- Discarded incomplete email on failure.
- Restricted Ribbon to Outlook Explorer.
- Normalized `GetDefaultFolder` to `Outlook.MAPIFolder`.
- Improved build/publish validation.

## [2.0.0] - 2026-09-11 — First VSTO migration

- First application-level VSTO Outlook add-in.
- Added Outlook Ribbon integration.
- Ported dynamic parts and multi-attachment mapping into a native Outlook-hosted architecture.
- Added build/publish scripts for ClickOnce.
- Superseded by 2.0.1/2.0.2 before becoming the recommended architecture.

## [1.4.0] - 2026-09-11 — Final PowerShell architecture

- Single-instance protection.
- Better COM cleanup on account reload/testing.
- Incomplete email auto-discard on failure.
- More conservative MIME size estimate.
- Added `REVISAR PARTE -> ANEXOS`.
- Added `Limpar tudo`.
- Improved DPI/smaller notebook support.
- Saved account selection in layouts.
- Added backward compatibility for older layout files.
- Added unsaved-change warning.
- Hardened PowerShell 5.1 / STA behavior.

## [1.3.0] - 2026-09-11 — Dynamic parts

- Changed the data model from `1 file = 1 email` to `Part = email`.
- Removed the fixed-part assumption.
- Added multiple attachments per Part.
- Added `Arquivos -> partes`.
- Added `+ Nova parte`.
- Added attachment panel per selected Part.
- Added part reordering with automatic renumbering.
- Added save/load `.opsparts.json`.
- Updated size estimation to sum every attachment in the Part.

## [1.2.0] - 2026-09-11 — Robustness and Outlook controls

- Improved Windows PowerShell 5.1 compatibility.
- Added selectable Outlook account and Outlook test.
- Added `SendUsingAccount`.
- Added To / CC / BCC.
- Added pre-validation and per-message `ResolveAll`.
- Added size and Base64/MIME estimates.
- Added configurable size warning.
- Added duplicate-file protection and natural sorting.
- Added manual reorder controls.
- Switched signature insertion to WordEditor.
- Added batch ID and safer success wording.

## [1.0.1] - 2026-09-11 — Compatibility hotfix

- Removed WinForms `PlaceholderText` usage from recipient text boxes for Windows PowerShell 5.1/.NET Framework compatibility.
- Packaging and workflow otherwise remained the same as the initial prototype.

## [1.0.0] - 2026-09-11 — Prototype

- First functional Outlook Part Sender.
- Same To/CC, subject base and message across a sequence.
- One attachment per email.
- Automatic `Parte 01/NN` subject numbering.
- Outlook account selection.
- Drafts as safe default and optional automatic Send.
- Installer/shortcut workflow based on PowerShell and Outlook COM.

> All 1.x, 2.x and 3.0.0 releases are historical. **v3.1.0 is the recommended public release.**
