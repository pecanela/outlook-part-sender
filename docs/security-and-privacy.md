# Security and Privacy

## What the application accesses

The program accesses the local Classic Outlook profile through the Outlook Object Model to:

- enumerate configured accounts;
- resolve recipients;
- create email items;
- save drafts;
- optionally send messages;
- access the Outlook Word editor for signature-aware body insertion.

## Credentials

The application does not request an Outlook password and does not contain an authentication flow. It relies on the existing signed-in Outlook session.

## Persistent data

By default, recipients/body/attachment mappings exist only in memory.

If the user chooses **Save layout**, a `.opsparts.json` file is created by the user at a chosen location. That JSON may contain:

- account label/key;
- recipient addresses;
- subject;
- message body;
- local attachment paths;
- layout options.

It does not contain an Outlook password.

## Local executable cache

The portable launcher creates:

```text
%LOCALAPPDATA%\OutlookPartSenderPortable
```

Typical contents include:

- compiled EXE;
- copied C# source;
- compile log;
- version marker.

`LIMPAR CACHE LOCAL.cmd` removes this cache.

## Enterprise restrictions

The project deliberately does not implement workarounds for security policies. If an organization blocks `.cmd`, `csc.exe`, LocalAppData executables, Outlook automation, sending or file access, those controls should be respected.

## Recommended operational mode

For business/customer deliveries, use **Drafts** and review the generated messages before sending.
