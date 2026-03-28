# @zyx1121/apple-powerpoint-mcp

MCP server for Microsoft PowerPoint — create, edit, and export presentations via Claude Code.

## Install

```bash
claude mcp add apple-powerpoint -- npx @zyx1121/apple-powerpoint-mcp
```

## Prerequisites

- macOS with Microsoft PowerPoint installed
- Node.js >= 18
- First run will prompt for Automation permission (System Settings > Privacy & Security > Automation)

## Tools

| Tool | Description |
|------|-------------|
| `powerpoint_create` | Create a new blank presentation |
| `powerpoint_open` | Open an existing .pptx file |
| `powerpoint_get_info` | Get active presentation info (name, path, slide count) |
| `powerpoint_list_layouts` | List available layouts from the slide master |
| `powerpoint_add_slide` | Add a slide with a specific layout |
| `powerpoint_set_text` | Set text of a shape on a slide |
| `powerpoint_format_text` | Set font size, bold, italic, color, font name |
| `powerpoint_set_bullets` | Enable/disable bullet points |
| `powerpoint_delete_slide` | Delete a slide by number |
| `powerpoint_list_slides` | List all slides with titles |
| `powerpoint_save` | Save the presentation |
| `powerpoint_export_pdf` | Export as PDF |
| `powerpoint_preview` | Export PDF to /tmp for visual inspection |

## Usage

```
"Create a new presentation"        → powerpoint_create
"Open template.pptx"               → powerpoint_open { path: "/path/to/template.pptx" }
"Show available layouts"           → powerpoint_list_layouts
"Add a title slide"                → powerpoint_add_slide { layout_index: 1, title: "Hello" }
"Set body text"                    → powerpoint_set_text { slide_number: 1, shape_index: 2, text: "Content" }
"Make title bold 36pt"             → powerpoint_format_text { slide_number: 1, shape_index: 1, font_size: 36, bold: true }
"Preview the result"               → powerpoint_preview
"Save to desktop"                  → powerpoint_save { path: "/Users/me/Desktop/deck.pptx" }
```

## Template Support

Works with any .pptx template — open the template, use `powerpoint_list_layouts` to discover available layouts, then `powerpoint_add_slide` with the desired `layout_index`.

## Limitations

- macOS only (uses AppleScript/JXA via `osascript`)
- PowerPoint.app must be running
- File exports go through PowerPoint's sandbox (`~/Library/Containers/com.microsoft.Powerpoint/Data/`) — the `preview` and `export_pdf` tools handle this automatically

## License

MIT
