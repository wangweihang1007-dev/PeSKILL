---
name: investment-meeting-minutes
description: Convert an investment interview transcript draft plus clean company BP/background materials in DOCX, TXT, PDF, or PPTX into a corrected transcript, professional Q&A, and formal Chinese investment meeting minutes with relevant BP images inserted into matching sections. Use when the user asks Codex to process an investment meeting folder, repair transcript terminology, organize interview Q&A, summarize company positioning/products/market/technology/financials/financing/development, or reproduce the 冯源资本 meeting-minutes workflow. Do not use for audio transcription, handwritten-note recognition, external company research, project intake spreadsheets, or investment recommendations.
---

# Investment Meeting Minutes

Process one project folder end to end. Preserve source files and write each run to a new timestamped output directory.

## Required references

Read these before processing:

- `references/original-prompts.md`: authoritative wording and business rules from the user's guide.
- `references/workflow-and-schemas.md`: file discovery, stage interfaces, and JSON schemas.
- `references/quality-rules.md`: evidence, conflict, completeness, and delivery gates.

## Workflow

1. Resolve the bundled workspace Python runtime. Do not use unverified global Python packages.
2. Create an external temporary work directory. Run:

   ```powershell
   python scripts/prepare_project.py --project "<absolute-project-folder>" --work-dir "<absolute-temp-work-folder>"
   ```

3. Read `manifest.json`. Stop and ask only when its status is `blocked`. Ignore audio, standalone images, prior output folders, and filenames containing `会议笔记`, `手写`, or `批注`. Keep images extracted from BP PDF/PPTX through `background_images.json`.
4. Read `background.txt` once and create a compact source ledger at `<work>/source_facts.json`: canonical company/product/technology names, dated facts, numeric facts, meeting metadata, and source-attributed conflicts. Reuse this ledger instead of rereading the full background at every stage. Read `transcript.txt` and `correction_chunks.json`.
5. Correct every chunk using the transcript-correction prompt and the compact terminology portion of the source ledger. Keep every substantive turn, timestamp, speaker, number, company name, and sequence position. Write a single UTF-8 `corrected_transcript.txt` in the temp work directory.
6. Run the correction gate:

   ```powershell
   python scripts/quality_check.py corrected --source "<work>/transcript.txt" --corrected "<work>/corrected_transcript.txt"
   ```

   Fix all errors before continuing. Treat warnings as review items.
7. Convert the corrected transcript into Q&A using the Q&A prompt. Process complete topic/question boundaries rather than arbitrary character slices. Merge globally, remove duplicate questions, and write `qa.json` using the required schema.
8. Run the Q&A gate. Fix all errors:

   ```powershell
   python scripts/quality_check.py qa --qa "<work>/qa.json"
   ```

9. Generate `minutes.json` from `source_facts.json` plus final Q&A; consult `background.txt` only to verify a disputed or missing claim. Use no web research. Resolve relative years from the system date. Preserve conflicting source claims as attributed parallel statements; never silently choose one.
10. Run the minutes gate, then build all three DOCX files. The build script inserts relevant BP images into matching sections of the final meeting minutes:

   ```powershell
   python scripts/quality_check.py minutes --minutes "<work>/minutes.json"
   python scripts/build_documents.py --work-dir "<work>" --project "<project-folder>" --template "assets/meeting-minutes-template.docx"
   ```

11. Render all three DOCX files with the installed documents skill. Inspect every page at 100%. Fix clipping, overlap, broken headings, excessive blank space, or inconsistent Q/A spacing, rebuild, and rerender.
12. Return only the timestamped output folder and its three DOCX files. Do not expose temp files unless the user asks.

## Token discipline

- Extract each source once; do not repeatedly reopen DOCX/PPTX/PDF files.
- Keep correction chunks under the deterministic limit and process only one chunk at a time.
- Carry terminology and evidence forward through `source_facts.json`; do not paste the full BP into correction or Q&A prompts.
- Use the page context in `background_images.json`; do not spend model tokens reviewing every extracted image.
- Build Q&A only from the corrected transcript. Use background facts only during final summary and conflict verification.
- Do not spend tokens describing intermediate work unless blocked or the user asks.

## Non-negotiable behavior

- Do not modify, rename, or overwrite source files.
- Do not summarize during transcript correction.
- Do not introduce facts from the BP into the corrected transcript when the speaker did not say them; use the BP only to repair obvious terminology, names, and malformed expressions.
- Do not omit material detail from Q&A merely to make it concise.
- Do not use handwritten/annotated notes, audio, standalone images, or external websites. Only use images embedded in accepted BP PDF/PPTX files.
- Insert BP images only into the final meeting minutes, after relevant section text. Limit to one image per section and six images total; include the BP filename and page in the caption.
- Mark unavailable meeting metadata as `待确认`.
- When a source conflict changes commercial stage, financing, customer status, financial performance, or core technical claims, retain both attributed versions and add `待核实`.

## Output contract

Create, without overwriting prior runs:

```text
<project>/AI会议纪要输出/<YYYYMMDD-HHMMSS>/
├─ 01_修正转录.docx
├─ 02_QA整理.docx
└─ 03_<项目名称>会议纪要.docx
```

If the output directory already exists, add `-02`, `-03`, and so on.
