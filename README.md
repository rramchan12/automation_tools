# Automations

Automations is the Git repository to automate the hell out of your life.

Below that are various tools to make it happen. **OneNote Exporter** is one tool; **Scout Automations** is another. The goal is to keep small, reusable automation building blocks in one place, with enough documentation to install, register, schedule, and safely run them again.

## Contents

| Path | Purpose |
|---|---|
| `onenote-md\onenote-md.cmd` | Batch launcher that runs the PowerShell exporter with execution-policy bypass. |
| `onenote-md\onenote-md.ps1` | OneNote desktop COM automation script for listing sections, exporting pages/sections/subtrees, and syncing changed pages. |
| `examples\sync-onenote-sections.cmd` | Direct batch automation template for syncing one or more OneNote sections. |
| `scout-automations\sync-onenote-markdown-exports.json` | Microsoft Scout automation template for scheduled OneNote Markdown syncs. |

## Tools

### OneNote Exporter

OneNote Exporter converts OneNote desktop content into readable Markdown. It supports listing sections, exporting a single page, exporting a page subtree, exporting a whole section, and syncing changed pages only.

### Scout Automations

Microsoft Scout Automations are scheduled or condition-based jobs that run Scout prompts for you. In this repo, the Scout automation template runs the OneNote Exporter on a schedule, inspects the validation log, and reports status without exposing note body content.

## Requirements

- Windows
- OneNote desktop installed and signed in
- Target notebooks opened and synced in OneNote desktop
- PowerShell 5.1+

The exporter uses local OneNote COM automation, so it must run on a Windows machine where OneNote desktop can access the notebooks.

## Register the exporter with Scout

Create a Scout custom skill named `/onenote-md` that points Scout at the local exporter CLI. Use the path where this repository is cloned:

```text
Use this skill when exporting, syncing, converting, or inspecting OneNote content as Markdown.

Local CLI:
<repo>\onenote-md\onenote-md.cmd

Default behavior:
- Prefer read-only OneNote to Markdown exports.
- Use the .cmd launcher rather than invoking the .ps1 directly.
- For sync jobs, run:
  & "<repo>\onenote-md\onenote-md.cmd" sync -Section "<SectionName>" -Destination "<Destination>"
- After every export or sync, inspect the validation log and report Status, Errors count, output paths, and log path.
- Do not paste private note contents into chat unless the user asks for a specific snippet.
```

After registering the skill, Scout can run prompts such as:

```text
Use the onenote-md skill to sync changed pages only for "<SectionName>" into "<Destination>".
```

## Quick start

List available OneNote sections:

```cmd
onenote-md\onenote-md.cmd list
```

Export a full section:

```cmd
onenote-md\onenote-md.cmd export-section -Section "<SectionName>" -Destination "C:\Path\To\OneNoteExports"
```

Sync changed pages only:

```cmd
onenote-md\onenote-md.cmd sync -Section "<SectionName>" -Destination "C:\Path\To\OneNoteExports"
```

Run the included batch sync template:

```cmd
examples\sync-onenote-sections.cmd
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

The Scout automation template in `scout-automations\sync-onenote-markdown-exports.json` captures a scheduled changed-page sync:

- Schedule: every day at 6 PM
- Sections: replace `<SectionName1>` and `<SectionName2>` with the sections you want to sync
- Destination: replace `<Destination>` with your Markdown export root
- Mode: sync changed pages only

Use Scout's automation UI or automation tools to create a scheduled automation from this template, then adjust the schedule, sections, destination, and notification policy for your environment.

## Privacy

This repository contains automation scripts and configuration only. It should not include generated OneNote Markdown exports, validation logs, manifests, or raw page XML.
