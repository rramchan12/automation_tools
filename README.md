# automation_tools

Automation utilities for exporting selected Microsoft OneNote sections to Markdown and keeping those exports up to date.

## Contents

| Path | Purpose |
|---|---|
| `onenote-md\onenote-md.cmd` | Batch launcher that runs the PowerShell exporter with execution-policy bypass. |
| `onenote-md\onenote-md.ps1` | OneNote desktop COM automation script for listing sections, exporting pages/sections/subtrees, and syncing changed pages. |
| `examples\sync-ravios-sections.cmd` | Direct batch automation for syncing the `Ravi OS` and `Personal Coaching` sections. |
| `scout-automations\sync-onenote-markdown-exports.json` | Microsoft Scout automation definition for the daily 6 PM sync. |

## Requirements

- Windows
- OneNote desktop installed and signed in
- Target notebooks opened and synced in OneNote desktop
- PowerShell 5.1+

The exporter uses local OneNote COM automation, so it must run on a Windows machine where OneNote desktop can access the notebooks.

## Quick start

List available OneNote sections:

```cmd
onenote-md\onenote-md.cmd list
```

Export a full section:

```cmd
onenote-md\onenote-md.cmd export-section -Section "Ravi OS" -Destination "C:\Users\rramchandran\OneDrive - Microsoft\work\RaviOS\OneNoteExports"
```

Sync changed pages only:

```cmd
onenote-md\onenote-md.cmd sync -Section "Ravi OS" -Destination "C:\Users\rramchandran\OneDrive - Microsoft\work\RaviOS\OneNoteExports"
```

Run the included batch sync for both configured sections:

```cmd
examples\sync-ravios-sections.cmd
```

## Commands

| Command | Description |
|---|---|
| `list` | Lists available OneNote sections as a Markdown table. |
| `export-page` | Exports one page by section and page name. |
| `export-subtree` | Exports a page and its child pages. |
| `export-section` | Exports all pages in a section. |
| `sync` | Exports changed pages only, using the manifest to skip unchanged pages. |

Common parameters:

| Parameter | Description |
|---|---|
| `-Notebook` | Optional notebook name filter. Useful when multiple sections share a name. |
| `-Section` | OneNote section name. Required for export and sync commands. |
| `-Page` | Page name for `export-page` and `export-subtree`. |
| `-Destination`, `-OutputPath`, `-Out` | Output root folder. |
| `-LogFile`, `-Log`, `-LogPath` | Optional validation log path. |
| `-Versioned` | Also writes a timestamped copy under `versions`. |
| `-IncludeXml` | Saves raw OneNote page XML beside Markdown for debugging. |

## Output layout

```text
<Destination>\<Section>\latest\index.md
<Destination>\<Section>\latest\*.md
<Destination>\<Section>\.onenote-md-manifest.json
<Destination>\<Section>\logs\onenote-md-<yyyyMMdd-HHmmss>.log
<Destination>\<Section>\versions\<yyyyMMdd-HHmmss>\   # when -Versioned is used
```

Each run writes a validation log with status, counts, output paths, and per-page export status.

## Scout automation

The Scout automation in `scout-automations\sync-onenote-markdown-exports.json` captures the configured daily sync:

- Schedule: every day at 6 PM
- Sections: `Ravi OS`, `Personal Coaching`
- Destination: `C:\Users\rramchandran\OneDrive - Microsoft\work\RaviOS\OneNoteExports`
- Mode: sync changed pages only

## Privacy

This repository contains automation scripts and configuration only. It should not include generated OneNote Markdown exports, validation logs, manifests, or raw page XML.
