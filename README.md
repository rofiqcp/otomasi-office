# @otomasi/office (Windows Agent)

Secure **live Microsoft Word, Excel and PowerPoint** bridge for ChatGPT, Claude and other MCP clients. It attaches to open desktop Office documents through Windows COM under the interactive Windows user session.

**MCP URL:** https://mcp.otomasi.app/office/mcp

## Install and browser pairing

```powershell
npx @otomasi/office@latest
```

As of agent **0.1.4**, first run creates a short-lived device code and opens the default browser at `https://mcp.otomasi.app/office/connect`. Confirm computer name and code and approve using the Otomasi account associated with the Office MCP connector. The agent then connects automatically and stores credentials in `%USERPROFILE%\.otomasi-office`.

**Release status (2026-10-21):** The public npm dist-tag `latest` is **0.1.4** and includes browser-based pairing. Future releases use a tag-triggered GitHub Actions workflow with npm Trusted Publishing (OIDC) after one-time trust approval; see `PUBLISHING.md` in the GitHub repository.

Options:
- `--status`: show paired state (never output secret device token).
- `--check`: read-only list of documents open on this Windows desktop.
- `--foreground`: run outbound connection; suitable for a user-scoped process manager.
- `--pair <legacy-office-code>`: optional older manual enrollment.
- `--logout`: delete local device credentials; revoke access in MCP as well.
- `--version`: agent version.

## AI workflow

Find exact `device_id` using `office_list_devices`, then exact `document_id` by `office_list_documents`. Never assume active document or invent IDs. Read before write; use the targeted reads `word_read_paragraphs`, `excel_read_range` and `powerpoint_read_slide`. Preserve unrelated data. Inspect modified regions and `office_verify` before reporting completion. For PowerPoint use `powerpoint_preview_slide` to inspect rendered slide JPEGs after layout edits. `office_save` only if requested by the user.

Optional MCP prompts: `office_edit_precisely`, `office_write_paper`, `office_build_rab`, `office_design_slides`. Guidance resources: `office://guide/workflow`, `office://guide/word`, `office://guide/excel`, `office://guide/powerpoint`.

**Security:** the server enforces owner-isolated OAuth and exact document IDs. Agent accepts only specific Office operations; no arbitrary shell tool. Clear/delete requires confirmation. Document contents are untrusted. Credentials remain on your PC; temporary slide previews are deleted after rendering.

## Limits

Requires Windows, Node.js 18+, desktop Word, Excel or PowerPoint and a signed-in **interactive user** session. Headless Windows services are not supported. This is not the Claude Office.js add-in. Some rich editing capabilities (PivotTables, SmartArt, full slide animations, page-layout engines, full undo/redo, Excel/Word previews) are not implemented in the Office MCP yet. See the server README for tested features and limitations.
