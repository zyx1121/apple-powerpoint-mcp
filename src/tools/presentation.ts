import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { execFile } from "node:child_process";
import { copyFileSync, existsSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { z } from "zod";
import { runAppleScript, runJxa, escapeForAppleScript } from "../applescript.js";
import { PowerPointError } from "../applescript.js";
import { success, withErrorHandling } from "../helpers.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const PPTX_HELPER = join(__dirname, "pptx-helper.py");

async function runPythonHelper(args: string[]): Promise<string> {
  return new Promise((resolve, reject) => {
    execFile("uv", ["run", "--with", "python-pptx", "python3", PPTX_HELPER, ...args], { timeout: 30_000 }, (err, stdout, stderr) => {
      if (err) return reject(new PowerPointError(stderr || err.message));
      resolve(stdout.trimEnd());
    });
  });
}

// PowerPoint sandbox writes to ~/Library/Containers/com.microsoft.Powerpoint/Data/<path>
// We need to copy from sandbox to the real target path after export.
const PPT_SANDBOX = join(homedir(), "Library/Containers/com.microsoft.Powerpoint/Data");

function copySandboxFile(requestedPath: string, realTarget: string): string {
  // The file lands inside the sandbox at the same relative path
  const sandboxPath = join(PPT_SANDBOX, requestedPath);
  if (existsSync(sandboxPath)) {
    copyFileSync(sandboxPath, realTarget);
    return realTarget;
  }
  // Sometimes the file ends up at the real path (non-sandboxed PowerPoint)
  if (existsSync(requestedPath)) {
    if (requestedPath !== realTarget) copyFileSync(requestedPath, realTarget);
    return realTarget;
  }
  return sandboxPath; // return sandbox path as fallback
}

// Replace newlines with AppleScript return character
function prepareText(text: string): string {
  return escapeForAppleScript(text)
    .replace(/\n/g, '" & return & "')       // actual newline chars (from JSON)
    .replace(/\\\\n/g, '" & return & "');   // literal \n (escaped by escapeForAppleScript)
}

export function registerPresentationTools(server: McpServer) {
  // ── 建立新簡報 ─────────────────────────────────────────────
  server.registerTool(
    "powerpoint_create",
    {
      description: "Create a new blank PowerPoint presentation and optionally save it to a path.",
      inputSchema: z.object({
        path: z.string().optional().describe("Full file path to save the .pptx (e.g. /Users/me/Desktop/deck.pptx). If omitted, creates an unsaved presentation."),
      }),
    },
    withErrorHandling(async ({ path }) => {
      const saveScript = path
        ? `save newPres in "${escapeForAppleScript(path)}"`
        : "";
      const script = `
tell application "Microsoft PowerPoint"
  activate
  set newPres to make new presentation
  ${saveScript}
  return name of newPres
end tell`;
      const name = await runAppleScript(script);
      return success({ created: true, name: name.trim(), path: path ?? null });
    }),
  );

  // ── 開啟簡報 ──────────────────────────────────────────────
  server.registerTool(
    "powerpoint_open",
    {
      description: "Open an existing PowerPoint file (.pptx).",
      inputSchema: z.object({
        path: z.string().describe("Full path to the .pptx file"),
      }),
    },
    withErrorHandling(async ({ path }) => {
      const script = `
tell application "Microsoft PowerPoint"
  activate
  open "${escapeForAppleScript(path)}"
  delay 1
  tell active presentation
    return name & "\\t" & (count of slides)
  end tell
end tell`;
      const raw = await runAppleScript(script);
      const [name, count] = raw.split("\t");
      return success({ opened: true, name, slide_count: parseInt(count, 10) });
    }),
  );

  // ── 取得目前開啟的簡報資訊 ────────────────────────────────
  server.registerTool(
    "powerpoint_get_info",
    {
      description: "Get info about the currently active PowerPoint presentation (name, path, slide count).",
      inputSchema: z.object({}),
    },
    withErrorHandling(async () => {
      const script = `
tell application "Microsoft PowerPoint"
  tell active presentation
    set slideCount to count of slides
    set presName to name
    try
      set presPath to full name
    on error
      set presPath to ""
    end try
    return presName & "\\t" & presPath & "\\t" & slideCount
  end tell
end tell`;
      const raw = await runAppleScript(script);
      const [name, path, count] = raw.split("\t");
      return success({ name, path: path || null, slide_count: parseInt(count, 10) });
    }),
  );

  // ── 列出可用 Layouts ──────────────────────────────────────
  server.registerTool(
    "powerpoint_list_layouts",
    {
      description:
        "List all available layouts from the active presentation's slide master. " +
        "MUST be called before adding slides to know which layout_index to use. " +
        "Returns each layout's index (1-based), shape count, and shape details.",
      inputSchema: z.object({}),
    },
    withErrorHandling(async () => {
      const RS = "\u001e";
      const FS = "\u001f";
      // Create a temp slide for each layout to get accurate shape info,
      // because layout objects don't expose placeholder shapes directly.
      const script = `
tell application "Microsoft PowerPoint"
  tell active presentation
    set m to slide master
    set n to 0
    repeat
      set n to n + 1
      try
        set cl to custom layout n of m
      on error
        set n to n - 1
        exit repeat
      end try
    end repeat

    set output to ""
    repeat with i from 1 to n
      set cl to custom layout i of m
      set tempSlide to make new slide at end
      set custom layout of tempSlide to cl
      set sc to count of shapes of tempSlide
      set shapeInfo to ""
      repeat with j from 1 to sc
        try
          set t to content of text range of text frame of shape j of tempSlide
          if t is missing value then set t to "(placeholder)"
          if t is "" then set t to "(placeholder)"
          set shapeInfo to shapeInfo & "s" & j & "=" & t
        on error
          set shapeInfo to shapeInfo & "s" & j & "=(non-text)"
        end try
        if j < sc then set shapeInfo to shapeInfo & "${FS}"
      end repeat
      set output to output & i & "\\t" & sc & "\\t" & shapeInfo & "${RS}"
      delete tempSlide
    end repeat
    return output
  end tell
end tell`;
      const raw = await runAppleScript(script);
      if (!raw.trim()) {
        return success({ layout_count: 0, layouts: [], hint: "No custom layouts found. This may be a blank presentation without a template." });
      }
      const layouts = raw
        .split(RS)
        .filter(Boolean)
        .map((line) => {
          const [idx, shapeCount, shapesRaw] = line.split("\t");
          const shapes = (shapesRaw ?? "").split(FS).filter(Boolean).map((s) => {
            const [key, ...val] = s.split("=");
            return { shape: key, content: val.join("=") || "(empty)" };
          });
          return { index: parseInt(idx, 10), shape_count: parseInt(shapeCount, 10), shapes };
        });
      return success({ layout_count: layouts.length, layouts });
    }),
  );

  // ── 新增投影片 ────────────────────────────────────────────
  server.registerTool(
    "powerpoint_add_slide",
    {
      description:
        "Add a slide to the active presentation. Use powerpoint_list_layouts first to see available layouts. " +
        "Use \\n for line breaks in text fields.",
      inputSchema: z.object({
        layout_index: z.coerce.number().describe(
          "Layout index (1-based) from powerpoint_list_layouts. Required."
        ),
        title: z.string().optional().describe("Title text (shape 1 in most layouts)"),
        body: z.string().optional().describe("Body text (shape 2 in most layouts). Use \\n for line breaks."),
        position: z.coerce.number().optional().describe("Insert at this position (1-based). Defaults to end."),
      }),
    },
    withErrorHandling(async ({ layout_index, title, body, position }) => {
      const safeTitle = title ? prepareText(title) : null;
      const safeBody = body ? prepareText(body) : null;

      const titleScript = safeTitle
        ? `
      try
        set content of text range of text frame of shape 1 to "${safeTitle}"
      end try`
        : "";

      const bodyScript = safeBody
        ? `
      try
        set content of text range of text frame of shape 2 to "${safeBody}"
      end try`
        : "";

      const positionScript = position
        ? `move slide newSlide to before slide ${position} of active presentation`
        : "";

      const script = `
tell application "Microsoft PowerPoint"
  tell active presentation
    set newSlide to make new slide at end
    try
      set custom layout of newSlide to custom layout ${layout_index} of slide master
    end try
    tell newSlide${titleScript}${bodyScript}
    end tell
    ${positionScript}
    return slide number of newSlide
  end tell
end tell`;
      const slideNum = await runAppleScript(script);
      return success({ added: true, slide_number: parseInt(slideNum, 10), layout_index });
    }),
  );

  // ── 設定投影片文字 ────────────────────────────────────────
  server.registerTool(
    "powerpoint_set_text",
    {
      description: "Set the text of a specific shape on a slide. Use \\n for line breaks.",
      inputSchema: z.object({
        slide_number: z.coerce.number().describe("Slide number (1-based)"),
        shape_index: z.coerce.number().describe("Shape index on the slide (from powerpoint_list_layouts)"),
        text: z.string().describe("Text to set. Use \\n for line breaks."),
      }),
    },
    withErrorHandling(async ({ slide_number, shape_index, text }) => {
      const safeText = prepareText(text);
      const script = `
tell application "Microsoft PowerPoint"
  tell active presentation
    tell slide ${slide_number}
      set content of text range of text frame of shape ${shape_index} to "${safeText}"
    end tell
  end tell
end tell`;
      await runAppleScript(script);
      return success({ updated: true, slide_number, shape_index });
    }),
  );

  // ── 設定文字格式 ──────────────────────────────────────────
  server.registerTool(
    "powerpoint_format_text",
    {
      description: "Set font properties (size, bold, italic, color, font name) for a shape's text on a slide.",
      inputSchema: z.object({
        slide_number: z.coerce.number().describe("Slide number (1-based)"),
        shape_index: z.coerce.number().describe("Shape index on the slide"),
        font_size: z.coerce.number().optional().describe("Font size in points"),
        bold: z.boolean().optional().describe("Set bold"),
        italic: z.boolean().optional().describe("Set italic"),
        font_name: z.string().optional().describe("Font name (e.g. 'Arial', 'Helvetica Neue')"),
        color: z.string().optional().describe("Font color as hex (e.g. '#FFFFFF' for white)"),
      }),
    },
    withErrorHandling(async ({ slide_number, shape_index, font_size, bold, italic, font_name, color }) => {
      const props: string[] = [];
      if (font_size !== undefined) props.push(`set font size of f to ${font_size}`);
      if (bold !== undefined) props.push(`set bold of f to ${bold}`);
      if (italic !== undefined) props.push(`set italic of f to ${italic}`);
      if (font_name !== undefined) props.push(`set name of f to "${escapeForAppleScript(font_name)}"`);
      if (color !== undefined) {
        const hex = color.replace("#", "");
        const r = parseInt(hex.substring(0, 2), 16);
        const g = parseInt(hex.substring(2, 4), 16);
        const b = parseInt(hex.substring(4, 6), 16);
        props.push(`set color of f to {${r}, ${g}, ${b}}`);
      }

      const script = `
tell application "Microsoft PowerPoint"
  tell active presentation
    tell slide ${slide_number}
      set f to font of text range of text frame of shape ${shape_index}
      ${props.join("\n      ")}
    end tell
  end tell
end tell`;
      await runAppleScript(script);
      return success({ formatted: true, slide_number, shape_index });
    }),
  );

  // ── 設定 Bullets ──────────────────────────────────────────
  server.registerTool(
    "powerpoint_set_bullets",
    {
      description: "Enable or disable bullet points for a shape's text on a slide.",
      inputSchema: z.object({
        slide_number: z.coerce.number().describe("Slide number (1-based)"),
        shape_index: z.coerce.number().describe("Shape index on the slide"),
        enabled: z.boolean().default(true).describe("Enable (true) or disable (false) bullets"),
      }),
    },
    withErrorHandling(async ({ slide_number, shape_index, enabled }) => {
      const jxa = `
var ppt = Application("Microsoft PowerPoint");
var pres = ppt.activePresentation;
var slide = pres.slides[${slide_number - 1}];
var shape = slide.shapes[${shape_index - 1}];
shape.textFrame.textRange.paragraphFormat.bullet.visible = ${enabled};
"done";`;
      await runJxa(jxa);
      return success({ updated: true, slide_number, shape_index, bullets: enabled });
    }),
  );

  // ── 設定多層級文字（python-pptx）─────────────────────────
  server.registerTool(
    "powerpoint_set_text_levels",
    {
      description:
        "Set multi-level text with indent levels on a shape (uses python-pptx). " +
        "The template's bullet styles are preserved — only the indent level is set. " +
        "IMPORTANT: PowerPoint must save the file first, and will need to reopen after.",
      inputSchema: z.object({
        slide_number: z.coerce.number().describe("Slide number (1-based)"),
        shape_index: z.coerce.number().describe("Shape index on the slide"),
        paragraphs: z.array(z.object({
          text: z.string().describe("Paragraph text"),
          level: z.coerce.number().default(0).describe("Indent level (0=top, 1=sub, 2=sub-sub, ...)"),
        })).describe("Array of paragraphs with text and indent level"),
      }),
    },
    withErrorHandling(async ({ slide_number, shape_index, paragraphs }) => {
      // 1. Save from PowerPoint first
      const saveScript = `
tell application "Microsoft PowerPoint"
  save active presentation
  return full name of active presentation
end tell`;
      const filePath = (await runAppleScript(saveScript)).trim();

      // 2. Run python helper to set levels
      const parasJson = JSON.stringify(paragraphs);
      await runPythonHelper(["set-levels", filePath, String(slide_number), String(shape_index), parasJson]);

      // 3. Reopen in PowerPoint
      const reopenScript = `
tell application "Microsoft PowerPoint"
  close active presentation saving no
  delay 0.5
  open "${escapeForAppleScript(filePath)}"
  delay 1
  return "reopened"
end tell`;
      await runAppleScript(reopenScript);

      return success({ updated: true, slide_number, shape_index, paragraph_count: paragraphs.length });
    }),
  );

  // ── 刪除投影片 ────────────────────────────────────────────
  server.registerTool(
    "powerpoint_delete_slide",
    {
      description: "Delete a slide from the active presentation by slide number.",
      inputSchema: z.object({
        slide_number: z.coerce.number().describe("Slide number to delete (1-based)"),
      }),
    },
    withErrorHandling(async ({ slide_number }) => {
      const script = `
tell application "Microsoft PowerPoint"
  tell active presentation
    delete slide ${slide_number}
    return count of slides
  end tell
end tell`;
      const remaining = await runAppleScript(script);
      return success({ deleted: true, slide_number, remaining_slides: parseInt(remaining, 10) });
    }),
  );

  // ── 儲存簡報 ─────────────────────────────────────────────
  server.registerTool(
    "powerpoint_save",
    {
      description: "Save the active PowerPoint presentation. Optionally save to a new path.",
      inputSchema: z.object({
        path: z.string().optional().describe("Save to this path. If omitted, saves in place."),
      }),
    },
    withErrorHandling(async ({ path }) => {
      const script = path
        ? `
tell application "Microsoft PowerPoint"
  save active presentation in "${escapeForAppleScript(path)}"
  return full name of active presentation
end tell`
        : `
tell application "Microsoft PowerPoint"
  save active presentation
  return full name of active presentation
end tell`;
      const savedPath = await runAppleScript(script);
      return success({ saved: true, path: savedPath.trim() });
    }),
  );

  // ── 匯出 PDF ─────────────────────────────────────────────
  server.registerTool(
    "powerpoint_export_pdf",
    {
      description: "Export the active PowerPoint presentation as a PDF file. Handles PowerPoint sandbox path automatically.",
      inputSchema: z.object({
        path: z.string().describe("Full path for the exported PDF (e.g. /Users/me/Desktop/deck.pdf)"),
      }),
    },
    withErrorHandling(async ({ path }) => {
      // Use /tmp inside sandbox, then copy out
      const tmpName = `/tmp/ppt-export-${Date.now()}.pdf`;
      const script = `
tell application "Microsoft PowerPoint"
  save active presentation in "${escapeForAppleScript(tmpName)}" as save as PDF
end tell`;
      await runAppleScript(script);
      const finalPath = copySandboxFile(tmpName, path);
      return success({ exported: true, path: finalPath });
    }),
  );

  // ── 列出所有投影片 ────────────────────────────────────────
  server.registerTool(
    "powerpoint_list_slides",
    {
      description: "List all slides in the active presentation with their titles and layout info.",
      inputSchema: z.object({}),
    },
    withErrorHandling(async () => {
      const RS = "\u001e";
      const script = `
tell application "Microsoft PowerPoint"
  tell active presentation
    set output to ""
    repeat with i from 1 to count of slides
      tell slide i
        set shapeCount to count of shapes
        try
          set t to content of text range of text frame of shape 1
        on error
          set t to "(no title)"
        end try
        set output to output & i & "\\t" & t & "\\t" & shapeCount & "${RS}"
      end tell
    end repeat
    return output
  end tell
end tell`;
      const raw = await runAppleScript(script);
      const slides = raw
        .split(RS)
        .filter(Boolean)
        .map((line) => {
          const [num, title, shapes] = line.split("\t");
          return { slide_number: parseInt(num, 10), title: (title ?? "").trim(), shape_count: parseInt(shapes, 10) };
        });
      return success({ slide_count: slides.length, slides });
    }),
  );

  // ── 預覽投影片（匯出 PDF 到 /tmp）─────────────────────────
  server.registerTool(
    "powerpoint_preview",
    {
      description:
        "Export the active presentation as PDF to /tmp/ppt-preview.pdf for visual inspection. " +
        "Use this after making changes to verify the result visually. " +
        "The returned path can be read with the Read tool to view the slides.",
      inputSchema: z.object({
        slides: z.string().optional().describe("Slide range to hint (e.g. '1-5'). PDF always exports all slides, but this helps the caller know which pages to read."),
      }),
    },
    withErrorHandling(async ({ slides }) => {
      const tmpName = `/tmp/ppt-preview-${Date.now()}.pdf`;
      const script = `
tell application "Microsoft PowerPoint"
  save active presentation in "${escapeForAppleScript(tmpName)}" as save as PDF
  return count of slides of active presentation
end tell`;
      const slideCount = await runAppleScript(script);
      const realPath = `/tmp/ppt-preview.pdf`;
      copySandboxFile(tmpName, realPath);
      return success({
        path: realPath,
        slide_count: parseInt(slideCount, 10),
        hint: slides ? `Focus on pages ${slides}` : "Read the PDF with the Read tool to see the slides visually.",
      });
    }),
  );
}
