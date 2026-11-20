# PowerPoint COM Templates

[简体中文](README.md) · English

A Codex skill for creating or redesigning editable .pptx decks with **COM in desktop Microsoft PowerPoint on Windows**. The author chooses one template and supplies the content and visuals. Codex plans the sections and individual slides before editing.

Requires a local Windows Codex task, a working installation of desktop Microsoft PowerPoint, and PowerShell. No MCP server, Python dependency, or Office add-in is required. Browser-only PowerPoint, a Linux host, or a chat-only assistant cannot run local COM.

## Copy and paste this into your AI assistant

```text
In my local Windows Codex environment, download this repository:
https://github.com/yc0203/powerpoint-com-template
Read the README and powerpoint-com-template/SKILL.md, then install the entire
powerpoint-com-template subfolder in my Codex user skills directory. Keep its
scripts, references, agents, and assets together. Use the skills path already
configured on this machine. If the skill already exists, tell me; do not overwrite
it or change global settings. If GitHub access requires sign-in, use my existing
authorized access without asking for a token. Report the absolute path to SKILL.md
and show me how to invoke $powerpoint-com-template.
```

After installing, paste this into a new Codex chat:

```text
Use $powerpoint-com-template to create an editable deck in my local desktop PowerPoint.
First ask me to choose one template and provide a content document with text,
illustrations, captions, tables, and source data. Organize the content into sections
and plan the placement, sizing, alignment, and borders of figures and tables.
Use only the selected template, avoid identical body layouts on consecutive slides,
and apply a consistent red to key terms and takeaway sentences.
Build the deck using COM and let me see the pages change. Compare each rendered
slide with its template reference, fix issues, and reopen the final deck to verify it.
```

## Installation and templates

1. Clone this repository or choose **Code → Download ZIP** on GitHub.
2. Copy the entire [powerpoint-com-template/](powerpoint-com-template/) folder into your Codex user skills directory. A standard user scope is ~/.agents/skills/; use the destination recognized by your installation. The installed file should be at `.../skills/powerpoint-com-template/SKILL.md`.
3. Start a new Codex chat and invoke `$powerpoint-com-template`. If it is not listed, give Codex the absolute path to the local SKILL.md and confirm that the task runs on your local Windows machine.

The [original starter PPTX](powerpoint-com-template/assets/starter-research/reference.pptx), ID `starter-research`, includes a cover, an agenda, an evidence slide, a native flow diagram, a native table, and a plan. You can also supply a local PPTX you are authorized to use. Third-party templates are not distributed here.

The skill plans pages from the source material, works on a copy, and reuses native objects. It exports slide previews from PowerPoint, compares them with the chosen template, and reopens the result after corrections. Creating new native equations depends on the Office version and COM; the workflow reports this limitation when applicable. Input files remain unchanged.

[Skill instructions](powerpoint-com-template/SKILL.md) · [Template catalog](powerpoint-com-template/references/templates.md)
