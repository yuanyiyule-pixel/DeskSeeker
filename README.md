# DeskSeeker

Windows desktop screenshot grounding that locates one UI target from a natural-language description and returns one logical desktop coordinate without performing any click.

The project is designed for desktop-agent workflows where a model needs to identify a precise target such as a taskbar icon, input box, toolbar button, tab, or list row, then return a stable screen coordinate for downstream automation.

## Core Idea

- Capture the desktop screenshot
- Run coarse-grid grounding on the full screen
- Run fine-grid grounding on the selected neighborhood
- Run a third local refinement stage on an upscaled crop
- Use parallel model calls and majority voting for robustness
- Return one logical desktop coordinate only

## Repository Layout

- `README.md`: project overview and quick start
- `LICENSE`: MIT license
- `SKILL.md`: compact skill instructions
- `scripts/run.mjs`: main runner

## Public Copy Notes

This repository copy was prepared for public release.

- Encrypted key-file fallback logic was removed
- Authentication uses only `OPENROUTER_API_KEY`
- No private network paths or local secret-store references remain in the copied runner

## Requirements

- Windows
- Node.js 18+
- PowerShell
- OpenRouter API key in `OPENROUTER_API_KEY`
- Python screenshot dependencies when using the preferred screenshot backend:
  - `mss`
  - `Pillow`

## Setup

Install the Python screenshot dependencies:

```powershell
python -m pip install mss Pillow
```

Set the API key:

```powershell
$env:OPENROUTER_API_KEY = "your_key_here"
```

## Quick Start

Run one grounding request:

```powershell
node scripts/run.mjs --description "Windows taskbar browser icon. 请注意一定要返回最有可能操作成功的坐标的位置。"
```

Run with final review:

```powershell
node scripts/run.mjs --description "Windows taskbar browser icon. 请注意一定要返回最有可能操作成功的坐标的位置。" --review
```

Run with verbose logs:

```powershell
node scripts/run.mjs --description "Windows taskbar browser icon. 请注意一定要返回最有可能操作成功的坐标的位置。" --verbose
```

## Output

Successful output:

```json
{
  "coordinate_space": "logical_desktop",
  "coordinate_note": "这是逻辑桌面坐标，请按此坐标直接操作。",
  "x": 148,
  "y": 1041
}
```

This output is a logical desktop coordinate, not a physical capture-pixel coordinate.

## Operational Notes

- The runner is silent by default. Use `--verbose` for progress logs.
- Use `--dump-raw` to write a sidecar trace JSON file for debugging.
- The runner never performs the click itself.
- The default voting setup runs 6 parallel model calls per round and starts voting after 4 successful replies.

## Limitations

- Windows-only workflow
- Requires an external VLM through OpenRouter
- Small, low-contrast, or highly ambiguous targets can still fail
- Similar neighboring icons remain a hard case even with multi-stage refinement

## License

MIT. Commercial use, modification, redistribution, and private use are allowed.

## Pre-Publish Check

- Review screenshots, examples, and saved artifacts before making the repository public
- Decide which screenshots or benchmark samples can be shared publicly
- Replace placeholder repository screenshots or figures before announcement
