---
name: doki-p
description: "Use for Doki / 代行者 Japanese otome game paid static creative production. Handles source-file discovery from the pre-registration asset library, role info and voice-line spreadsheets, character KV/SR/SSR references, image2/api2img prompts, strict copy audit, character consistency checks, logo/no-logo decisions, delivery renaming, and final asset packaging."
---

# Doki-P

## When to use

Use this skill for Doki / 代行者 Japanese otome game paid static creatives, especially 1:1 image templates made from the pre-registration asset library, character references, and Japanese copy tables.

Default source library when no path is supplied:

```text
Z:\创新事业部\AIGC项目视频组\代行者素材\预注册可用素材
```

## Core rules

1. Latest user instruction always wins.
2. Do not silently change production method.
   - If user says image2 only, do not use local compositing.
   - If user later allows local logo placement, local work is allowed only for that operation.
3. Do exactly the requested scope.
   - If user says only change logo, do not regenerate the whole image.
   - If user says only redo image 1, do not touch the batch.
4. No pretending: failed drafts must be called failed drafts.
5. Copy must come only from the approved Japanese copy table or current user instruction.
6. Preserve character identity: outfit, hair, face, accessories, temperament, core design.
7. Do not add logo unless the active instruction says to add logo. If user says no logo, no logo.
8. Keep the requested production method. If the accepted direction is image2/api2img direct generation, do not switch to local typography/compositing unless the user explicitly approves it.

## Standard workflow

1. Resolve source paths:
   - use the user-provided folder when present;
   - otherwise use the default pre-registration asset library above;
   - identify the target output folder before generating or packaging files.
2. Build a source inventory before writing prompts:
   - list top-level folders and key files;
   - count files by extension and by folder;
   - inspect image reference folders before choosing references;
   - inspect spreadsheets for sheet names, headers, and a few sample rows.
3. Read the core tables when available:
   - `角色信息表.xlsx`: role name, Japanese name, CV, crest, MBTI, birthday, height, motto, self-introduction, personality/core tags;
   - `2026年6月3日全角色台词.xlsx`: each character is usually one sheet; extract only the approved Japanese line from the `日文台词` column and use Chinese only for understanding intent.
4. Match character names across inconsistent naming:
   - match Chinese aliases, Japanese names, romanized filenames, and shortened names;
   - confirm ambiguous names with nearby evidence instead of guessing;
   - if a requested character cannot be matched to a reference image, say so before generating.
5. Select visual references from the library:
   - `角色kv`: primary identity reference, usually one PNG per character named in Chinese;
   - `角色立绘卡面/SR/单角色`: clean single-character SR references;
   - `角色立绘卡面/SSR`: high-impact card art, including `进化前`, `进化后`, `剧情差分`, and `主页适配` variants;
   - `前五章剧情场景`: background references such as 森罗大厅, 代行部办公室, 代行部走廊, 森罗宿舍走廊, 任务大厅, 研究中心, 训练室;
   - `Q版主线图` and `Q版宣传漫画`: use only when the creative direction is chibi/comic;
   - `徽章`: use only when badges/crests are explicitly requested;
   - `UI/logo.png` or top-level `LOGO.png`: use only when logo is required.
6. Extract exact Japanese copy per image:
   - CV line;
   - headline;
   - subcopy;
   - role name;
   - dialogue / interaction;
   - CTA.
7. Make a per-image prompt for image2/api2img.
8. Generate one image first after any new constraint; inspect before batching.
9. Audit every output at full size:
   - extra text / pseudo text;
   - CV correctness and size;
   - character outfit consistency;
   - role tags accidentally added;
   - logo/no-logo compliance;
   - layout issues such as empty corners or logo blocking content.
10. Only package approved files.

## Source reading commands

Prefer fast shell inventory first. On Windows / PowerShell:

```powershell
$root = 'Z:\创新事业部\AIGC项目视频组\代行者素材\预注册可用素材'
Get-ChildItem -LiteralPath $root -Force | Select-Object Mode,Length,LastWriteTime,Name
Get-ChildItem -LiteralPath $root -Recurse -File | Group-Object Extension | Sort-Object Count -Descending | Select-Object Count,Name
Get-ChildItem -LiteralPath $root -Directory | ForEach-Object { [PSCustomObject]@{ Name=$_.Name; Files=(Get-ChildItem -LiteralPath $_.FullName -Recurse -File | Measure-Object).Count } }
```

For Excel, use a structured reader. If Python path encoding fails on Chinese network paths, use Excel COM in PowerShell to print sheet names, dimensions, headers, and sample rows. Do not infer spreadsheet columns from filenames alone.

## Known asset library structure

The default library currently contains:

- top-level files: `2026年6月3日全角色台词.xlsx`, `角色信息表.xlsx`, `LOGO.png`;
- `角色kv`: 17 primary character KV PNGs, named by Chinese character names;
- `角色立绘卡面`: SR/SSR card art and `SR/单角色` cutouts;
- `前五章剧情场景`: official scene backgrounds for office, lobby, dormitory, training, research, shopping, park, and fantasy settings;
- `Q版主线图`, `Q版宣传漫画`: chibi and comic references;
- `UI`: contains `logo.png`;
- `徽章`: crest/badge assets;
- `预注册声优签名福利`: signed-benefit assets;
- `wav配音文件`: archives plus `已解压_按角色命名` audio organized by character.

Use this structure as a map, not as a substitute for reading the actual folder at runtime.

## Audio and CV handling

- For static image creatives, use `角色信息表.xlsx` as the source of CV names and character metadata.
- For audio-backed concepts, inspect `wav配音文件/已解压_按角色命名` and the rename reports before selecting lines.
- Never claim a static asset has matching voice audio unless the wav file was located and the line was matched.
- Do not use `.wav`, `.ogg`, `.zip`, or `.rar` assets for static image generation unless the user specifically asks for audio/video packaging.


## Generation execution rules

- For image2/api2img direct-generation workflows, generate the final artwork directly with the model. Do not add local text layers, local typography, or local compositing unless the user explicitly allows that production method.
- Write long prompts to a UTF-8 no-BOM temporary `.txt` file, then read the file content into the api2img call. Avoid passing long multilingual prompts with quotes directly in one command line because PowerShell/wrapper argument parsing can split Japanese/Chinese text into invalid arguments.
- Generate to an ASCII-only temporary output path and filename first, especially on Windows. After a successful image is created and inspected, copy or rename it to the required Chinese delivery filename. This avoids api2img wrapper issues with Chinese paths, mojibake, and false `output already exists` errors.
## Prompt pattern

```text
Use case: illustration-story
Asset type: Japanese otome game paid UA static creative, square 1:1.
Style: premium Japanese otome game collage ad, polaroids, blank sticky notes, tape, paper clips, torn paper, low-to-mid saturation, no cyberpunk.
Character reference: Use uploaded references strictly. Keep outfit, hair, face, accessories, temperament. Pose/expression variation only if allowed.
Source references: List the exact local reference files used before generation. Prefer `角色kv` + `SR/单角色` for identity, add SSR/card/background references only when they match the concept.

Use ONLY these exact Japanese texts:
<CV>
<title>
<subtitle>
<role-name>
<dialogue>
<CTA>

No other readable text. No pseudo text. Blank notes only. No crest. No character tags. No logo unless currently required.
CV must be large and eye-catching. Role info only <role-name>.
```

## Logo handling

Follow only the current instruction.

- If no logo: do not add logo.
- If image2 must add logo: input `LOGO.png` as reference, but test one image first because image2 may redraw the logo.
- If local logo placement is explicitly allowed:
  - use original `LOGO.png` only;
  - resize proportionally only;
  - no recolor, stroke, glow, shadow, rotation, filter, or sticker effect;
  - manually choose a placement that does not block character/copy;
  - preserve images the user already approved.

## Delivery renaming rule

When delivering approved Doki assets, rename files with this exact pattern:

```text
日期-Doki-CV-素材类型-素材名-GG-设计师-比例.扩展名
```

Example:

```text
260603-Doki-CV-P-森罗招聘广告青天目隐-GG-ZHM-1_1.png
```

Field meanings:

- `日期`: 6-digit date, e.g. `260527`, `260529`, `260603`.
- `Doki`: fixed project name.
- `CV`: fixed field.
- `素材类型`: `V` for video, `P` for picture/image.
- `素材名`: use `图片剧情名称+角色名称`; add a number only to avoid duplicates.
- `GG`: fixed field.
- `设计师`: designer code, e.g. `ZHM`, `LY`, `DLT`.
- `比例`: e.g. `1_1`, `9_16`, `16_9`.
- `扩展名`: preserve actual file format, e.g. `.mp4`, `.png`, `.jpg`.

Default assumptions:

- date uses current `yyMMdd`;
- static images use `P`;
- designer defaults to `ZHM` unless told otherwise;
- square images use `1_1`;
- preserve PNG/JPG/MP4 extension.

Do not create contact sheets / preview boards in the final delivery folder unless the user explicitly asks.

## Failure modes to avoid

- Producing prompts from user copy alone when the Doki asset library is available and relevant.
- Guessing character identity without reading `角色信息表.xlsx` and checking image references.
- Using Chinese translation text in the final creative when Japanese copy is required.
- Mixing character aliases incorrectly across KV, SR/SSR filenames, spreadsheet sheet names, and audio folders.
- Local compositing or local typography when the accepted workflow is image2/api2img direct generation.
- Passing long multilingual prompts directly through the command line instead of a UTF-8 prompt file.
- Writing api2img output directly to Chinese paths/filenames when an ASCII temp path plus final rename would be safer.
- Local compositing when user demanded image2 only.
- Whole-image regeneration when user asked to modify only logo.
- Adding logo after user said no logo.
- Omitting logo after user says logo is required.
- Placing logo blindly over important content.
- Claiming an image is final when copy/character/logo audit failed.
- Batch-running 8 images immediately after a new constraint.

## Response style

Be concise. When correcting mistakes, acknowledge briefly and act. Return final paths and exactly what changed.
