---
name: deskseeker
description: "DeskSeeker is a Windows desktop screenshot grounding skill that captures the current desktop, locates one target from a concrete natural-language description, and returns one logical desktop coordinate without performing any click. Triggers include desktop UI grounding, screenshot-to-coordinates, taskbar icon coordinates, input box coordinates, and similar screen-positioning tasks."
---

# DeskSeeker

Use the runner in `scripts/` when the task is:

- capture the current Windows desktop
- locate one target from a text description
- return one logical desktop coordinate only
- never perform the click
- work with icons and general desktop controls such as input boxes, dropdowns, tabs, and list rows

Runner:

```bash
node scripts/run.mjs --description "Click the browser icon on the taskbar. Please return the coordinate most likely to succeed when clicked."
```

Claude Sonnet 3.5:

```bash
node scripts/run.mjs --description "Click the browser icon on the taskbar. Please return the coordinate most likely to succeed when clicked." --model claude-sonnet-3.5
```

Enable final review:

```bash
node scripts/run.mjs --description "Click the browser icon on the taskbar. Please return the coordinate most likely to succeed when clicked." --review
```

Enable verbose logs:

```bash
node scripts/run.mjs --description "Click the browser icon on the taskbar. Please return the coordinate most likely to succeed when clicked." --verbose
```

Safe mode:

```bash
node scripts/run.mjs --description "Click the browser icon on the taskbar. Please return the coordinate most likely to succeed when clicked." --dry-run
```

Arguments:

- `--description <text>` or `--task <text>`: required target description; describe only the target
- `--model <name>`: optional model selector. Default is `gemini-3-flash-preview`. Built-in aliases: `gpt-5.4`, `claude-sonnet-3.5`, `claude-sonnet-4.6`, `gemini-3-flash`, `gemini-3-flash-preview`, `gemini-3.1-flash-lite`. Direct OpenRouter model ids are also accepted.
- `--reasoning-effort <level>`: optional reasoning effort. Default is `medium`. Supported values: `none`, `minimal`, `low`, `medium`, `high`
- `--screenshot <filepath>`: optional existing full-screen PNG to reuse instead of capturing a new desktop screenshot
- `--review`: enable the final review vote stage; disabled by default
- `--verbose`: print verbose progress logs to `stderr`; disabled by default
- `--dump-raw`: write raw visible model responses and trace to a sidecar JSON file while keeping the main output as coordinate-only JSON
- `--dry-run`: capture screenshot only, no network call
- `--out <filepath>`: optional JSON result path
- `--help`: show help

Authentication:

- Set the OpenRouter API key in the environment variable `OPENROUTER_API_KEY`.
- The public version of this skill does not read encrypted key files or local secret stores.

Behavior:

1. Capture the virtual desktop to a PNG under `saves/deskseeker/`, or reuse an existing screenshot when `--screenshot` is provided.
   The runner prefers a Python screenshot backend (`mss`, with `PIL.ImageGrab` fallback) and rescales to the logical Windows desktop size when the backend returns physical pixels. This avoids missing the Windows taskbar on systems where GDI `CopyFromScreen` drops shell-composited surfaces.
2. Split the full screenshot into an adaptive `n x n` coarse grid, draw high-contrast labels inside each cell, and send the labeled full-screen grid image to OpenRouter.
3. Run 6 parallel model calls for the coarse-grid stage and start the vote as soon as 4 successful replies arrive.
4. Crop the selected coarse cell plus its neighboring coarse cells, subdivide only the selected coarse cell into an adaptive `m x m` fine grid, and ask the model to vote on the best stage-2 fine cell from that single grid image.
5. If the stage-2 vote says the target is actually in a neighboring coarse cell, shift by one cell and redo stage 2 instead of restarting the whole run immediately.
6. Crop the selected stage-2 fine cell plus its surrounding fine cells from the original stage-2 crop, upscale that local image before drawing a final grid on it, and send two images to OpenRouter for stage 3: the stage-2 no-grid crop and the stage-3 local grid image.
7. Compute the candidate coordinate from the voted coarse cell, stage-2 fine cell, stage-3 cell, and stage-3 in-cell position.
8. If `--review` is enabled, draw the candidate point back onto the full screenshot and run another 6-way parallel review vote, again starting the vote as soon as 4 successful replies arrive; otherwise skip this stage by default.
9. Resize model-facing images to roughly 1024px on the longest side before upload by default, but allow the stage-3 local refinement image to stay larger after the pre-grid upscale so tiny targets remain readable.
10. Retry transient OpenRouter failures automatically instead of failing on the first transport or provider hiccup, and retry desktop screenshot capture a few times before giving up.
11. Return only one coordinate object with an explicit logical-desktop marker and note, for example `{ "coordinate_space": "logical_desktop", "coordinate_note": "This is a logical desktop coordinate. Use this coordinate directly.", "x": ..., "y": ... }`. This is a logical desktop coordinate; use it directly for the click. Never perform the click.

Runtime diagnostics:

- The runner is silent by default and keeps `stdout` reserved for the JSON result.
- Pass `--verbose` to print progress logs to `stderr`.
- Logs include stage start and end, elapsed time, OpenRouter request phases, crop and rotate operations, optional review passes, retry reasons, and fatal failure context.
- On success, the JSON result on `stdout` is only the final logical desktop coordinate.
- When `--dump-raw` is enabled, extra debug data is written to a sidecar `*.trace.json` file instead of expanding the main result shape.

The runner tells the selected model to identify the exact target named in the description, reject lookalikes, return the coordinate most likely to succeed when actually clicked, and return `none` instead of guessing when the target remains ambiguous.

Return shape:

```json
{
  "coordinate_space": "logical_desktop",
  "coordinate_note": "This is a logical desktop coordinate. Use this coordinate directly.",
  "x": 148,
  "y": 1041
}
```

These values are logical desktop coordinates. Please operate directly at this logical desktop coordinate.

Operational rules:

- When calling this skill, the description must be specific enough to identify the exact target and rule out nearby lookalikes. Do not use a vague label alone when similar neighbors may exist.
- Keep the description concrete. Mention visible labels, icon position, nearby anchors, and distinguishing features.
- Keep the description short and target-specific. Prefer one or two strong anchors over a long paragraph.
- Provide a sufficiently detailed and exact description so the model can precisely lock onto the right target.
- Include this fixed sentence in the description: `Please return the coordinate most likely to succeed when clicked.` The runner also auto-appends it internally if omitted.
- Aside from that fixed sentence, the external description should only identify the target itself. Do not put bbox instructions, click-surface rules, center-point requirements, output-format requirements, or neighbor-avoidance wording into `--description`; the runner handles those internally.
- Good example: `Windows taskbar Feishu icon`
- Better example when lookalikes exist: `Windows taskbar centered Feishu (Lark) icon, blue-and-white bird or paper-plane logo, not the adjacent Codex icon and not other nearby blue icons`
- Avoid: `Return the center of the full clickable bbox and avoid neighboring icons`
- Recommended description template: `[system or area] + [exact target name] + [stable visual feature] + [neighbor anchor or relative order] + [explicit exclusion of lookalikes]`
- Example template instantiation: `Windows taskbar centered Feishu (Lark) icon, blue-and-white bird/paper-plane logo, to the left of Codex, not the adjacent Codex icon`
- Use stable structural anchors when they are naturally part of the target description, such as `taskbar`, `desktop icon`, `window title`, `sidebar`, or `toolbar`.
- Prefer the default model and default reasoning effort first. Raise reasoning effort only when the target is unusually tiny, cluttered, rotated, or visually ambiguous.
- Prefer stable and easy-to-click points, such as the visual center of an icon, the center of a taskbar button, the body center of a button, or the editable body center of an input field.
- Avoid fragile edge points, thin borders, tiny badges, decorative pixels, text carets, or placeholder-text fragments when a safer center point exists.
- Return a clickable `bbox` that fully contains the requested target body or full clickable surface and excludes unrelated neighboring controls when possible; the runner will use the center of that `bbox` as the final coordinate.
- Do not return a tiny bbox that covers only the most visually obvious fragment of the target. If the target is an icon, include the full practical hit area or taskbar button body. If the target is an input box, search field, tab, row, menu item, or dropdown, include the full interactive body a user would naturally click.
- Keep the `bbox` tight around the full intended hit area, not around one symbol stroke, one colored corner, one letter, or one inner decoration; the runner performs one combined precision-and-context review pass before returning coordinates.
- Return `none` when the target is ambiguous, hidden, blocked, or not confidently identified.

Contact: `amart@novaserene.com`

