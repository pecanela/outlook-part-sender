# Publishing to GitHub

This repository package is already arranged for GitHub.

## Suggested repository name

```text
outlook-part-sender
```

## Suggested description

```text
Portable Windows utility for splitting large Outlook attachments into numbered multi-part emails — no install, no admin, no Outlook add-in.
```

## Suggested topics

```text
outlook
outlook-classic
windows
csharp
dotnet-framework
email
attachments
automation
portable
corporate-tools
winforms
```

## First push

```bash
git init
git add .
git commit -m "Release Outlook Part Sender v3.1.0"
git branch -M main
git remote add origin <YOUR_REPOSITORY_GIT_URL>
git push -u origin main
```

## GitHub settings

Recommended:

- Add the repository description above.
- Add the suggested topics.
- Enable Issues.
- Enable Discussions only if you want user support/conversation.
- Enable private vulnerability reporting under Security if available.
- Keep `main` as the default branch.

## Create the latest release

Create a GitHub Release with:

```text
Tag: v3.1.0
Title: Outlook Part Sender v3.1.0 — Portable Corporate
```

Use the prepared text:

```text
release-notes/v3.1.0.md
```

Attach:

```text
Outlook_Part_Sender_v3.1.0_PORTABLE_CORPORATE.zip
```

Mark it as **Latest release**, not pre-release.

## Historical releases

Release-note files for every historical version are in `release-notes/`.

Recommended presentation:

- v3.1.0 — normal/latest release.
- v3.0.0 — historical/legacy.
- v2.x — mark as pre-release or clearly label VSTO experiment.
- v1.x — historical/legacy.

You can upload the corresponding original ZIP assets from the separate `GitHub Release Assets` package.

## Do not commit release ZIPs into `main`

Use GitHub Release assets for ZIP packages. Keep the repository focused on source and documentation.
