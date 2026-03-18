# Brightspace Quiz Review Extractor

Turn Brightspace quiz export XML into reviewable Excel, JSON, and Markdown—with clear traceability back to the source.

## Why This Exists

Brightspace quiz exports are distributed across multiple XML files and are difficult to review in their raw form. Questions, feedback, pools, and banked content are often split across quiz files and `questiondb.xml`, with relationships that are not always explicit or stable.

This tool restructures that data into reviewer-friendly outputs while preserving traceability back to the original export. Ambiguities are surfaced as diagnostics instead of being silently resolved.

It is built for review, audit, and inspection—not for rebuilding importable quizzes.

This work also emerged alongside efforts to build a more deliberate and reliable import workflow. Existing Brightspace import paths provide limited control over how quizzes and question banks are reconstructed and placed.

A review-first layer makes it possible to inspect, verify, and reason about quiz structure before attempting import or migration, rather than treating the export as a black box.

This extractor is designed to support that broader workflow, where review and traceability come first, and reconstruction follows from verified structure.

## What It Produces

This project generates three review-first artifacts:

- `quiz_review.xlsx` — primary reviewer workbook
- `quiz_review.json` — canonical machine-readable output
- `quiz_review_summary.md` — concise per-quiz summary

The workbook is the primary artifact. JSON is the stable machine-readable output. Markdown is a quick inspection layer.

## What It Does

Use this extractor when you want to:

- review quiz wording, answer keys, pools, and feedback outside Brightspace
- trace each review row back to its quiz XML or question bank source
- preserve ambiguous Brightspace structures as diagnostics instead of guessing
- keep a stable JSON artifact for downstream comparison or automation

This tool does **not** try to rebuild importable Brightspace quizzes.

## Quick Start

Create a virtual environment, install the runtime dependency, and run the extractor against an unpacked Brightspace export folder or ZIP:

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python brightspace_quiz_review_extractor_v2.py /path/to/export --out ./quiz_review_out
```

On macOS, you can also run the included `.command` launcher. The CLI works cross-platform.

## Runtime and Test Prerequisites

The script depends on `openpyxl`.

Recommended local setup:

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

Run tests with:

```bash
.venv/bin/pip install -r requirements-dev.txt
.venv/bin/python -m pytest -q
```

## macOS Convenience Launcher

`Run Brightspace Quiz Review Extractor.command` is a double-click launcher for less technical users.

On first run it:

- creates a local `.venv`
- installs the runtime dependency from `requirements.txt`
- prompts for the Brightspace export
- prompts for the destination folder
- opens the generated output folder

It still requires Python 3 to be installed on the Mac.

## Supported Inputs

The extractor accepts either:

- a Brightspace export ZIP
- an unpacked Brightspace export folder

The unpacked folder should contain:

- `imsmanifest.xml`
- one or more `quiz_d2l_*.xml` files

If `questiondb.xml` is present, the extractor will use it for traceability and bank resolution.

## What You Need

To run the extractor, you need:

- Python 3.10+ recommended
- the files in this repository
- a Brightspace export containing `imsmanifest.xml` and `quiz_d2l_*.xml`
- `questiondb.xml` if the quiz uses banks or pools and you want the best traceability

You do **not** need:

- `pytest`, unless you want to run tests
- Excel to generate outputs
- any database, web server, or extra application stack

## Outputs

### `quiz_review.xlsx`

Primary reviewer workbook.

Workbook behavior:

- frozen header row on every sheet
- filters enabled on every sheet
- wrapped text for long fields
- minimum row heights for matching, ordering, long-answer, and image-bearing question rows
- stable sheet ordering
- reviewer-oriented column ordering
- reviewer-friendly columns grouped left and raw traceability columns grouped right
- question-level image trace fields and optional local hyperlinks

Sheet order:

1. `quiz_overview`
2. `sections_pools`
3. `questions`
4. `matching_pairs_expanded`
5. `pool_members`
6. `diagnostics`
7. `source_map`

### `quiz_review.json`

Canonical machine-readable artifact containing the same core data as the workbook, including image reference and resolution fields on question rows and the additive `matching_pairs_expanded` collection.

### `quiz_review_summary.md`

Per-quiz summary including:

- quiz title and ID
- storage and delivery type
- points
- section and pool counts
- diagnostic counts and top issue types

## Core Design Principle

The extractor does not guess.

If a relationship, especially bank resolution, is not confident:

- the row stays `inline` or `unresolved`
- a diagnostic is emitted
- the source clue is preserved

This is deliberate. Review fidelity is prioritized over forced completeness.

## Source Model

### Quiz-level `storage_type`

- `inline`: reviewable content came from quiz XML only
- `banked`: reviewable content came from `questiondb.xml` only
- `hybrid`: quiz XML content had a confident question-bank relationship
- `unresolved`: bank-backed intent was detected but could not be resolved confidently

`has_questiondb` only indicates that `questiondb.xml` exists in the package. It does **not** mean a quiz is automatically `hybrid`.

### Row-level `source_location`

Question and pool rows use:

- `inline`
- `hybrid`
- `questiondb`
- `unresolved`

The extractor also preserves:

- `source_quiz_file`
- `source_bank_file`
- `source_hint`

`source_hint` is the human-readable trace string that explains where the row came from, for example:

- `assessment/section[SECT_01]/item`
- `objectbank/section[SECT_47992]/item`
- `assessment/section[RAND_48803] resolved to questiondb section SECT_47992 via quiz-to-bank section relation`

## Bank Resolution Heuristics

When `questiondb.xml` is present, the extractor tries to match banked content in this order:

1. stable keys such as label, ident, local ID, display ID, or global ID
2. title or stem evidence plus question type
3. quiz-title to bank-section relationship heuristics

If a match is not confident:

- the row stays `inline` or `unresolved`
- a diagnostic is emitted
- the source clue is preserved

This is intentional. The extractor prefers explicit uncertainty over silent assumptions.

## Question Parsing Coverage

### Stronger direct support

- Multiple Choice
- True/False
- Multi-select / multiple response
- Matching
- Short Answer
- Fill in the Blanks
- Long Answer

### Best-effort support

- Ordering
- Numeric / arithmetic

For partially resolved structures, the extractor preserves extra detail in:

- `question_payload_json`
- `diagnostics`

## Feedback Parsing

The extractor keeps feedback channels separate whenever Brightspace exports enough structure:

- `general_feedback`
- `correct_feedback`
- `incorrect_feedback`
- `answer_specific_feedback`

If Brightspace feedback cannot be mapped cleanly, the extractor preserves what it can and emits a diagnostic instead of silently collapsing fields.

## Reviewer-Facing Workbook Fields

The `questions` sheet keeps the raw traceability fields, but also adds reviewer-facing display fields so matching and ordering questions are readable without opening JSON payloads.

Reviewer-facing additions on question rows:

- `question_title_review`: preserves the original title unless it is effectively an internal ID such as `question_id`, `QUES_*`, `ITEM_*`, or `OBJ_*`; only then does it fall back to a short stem-derived title
- `matching_review_display`: multiline `Prompt -> Correct match` display that preserves Brightspace prompt order
- `ordering_review_display`: multiline numbered sequence display; when ordering labels are missing, placeholder labels such as `Item 1` are shown instead of leaving the reviewer with only raw payload data
- `has_image`: `yes` / `no` reviewer filter derived from resolved image count

Traceability fields remain intact on the same row, including:

- `question_title`
- `matching_pairs`
- `accepted_answers`
- `question_payload_json`
- `source_location`
- `source_hint`
- `source_quiz_file`
- `source_bank_file`

The `questions` sheet is laid out to support review flow:

- identity and reviewer filters first
- reviewer-facing display fields next
- raw and technical payload fields farther right with separate header styling

## Matching Review Expansion

The workbook includes `matching_pairs_expanded`, a reviewer-facing expansion sheet with one row per prompt or match pair.

Each row includes:

- quiz and section context
- `question_order`, `question_id`, `question_title_review`, and raw `question_title`
- `prompt_order` and `match_order`
- `prompt`
- `correct_match`
- question-level image trace fields and primary image link when present
- source traceability fields

This sheet is built from the same parsed matching data used for the main `questions` sheet. It does not replace `matching_pairs`; it complements it.

## Image Handling

The extractor scans question-level Brightspace XML for image references in:

- stem and presentation material
- answer choice material
- `itemfeedback`
- `answer_key_material`

It supports the Brightspace patterns observed in exported quizzes and question banks:

- escaped HTML `&lt;img src="..."&gt;` inside `mattext`
- explicit `matimage` nodes

Question rows include:

- `image_refs`: raw refs as extracted after HTML-unescape
- `image_paths_resolved`: resolved package-relative file paths
- `image_count`: count of unique resolved images for the question
- `image_link_primary`: first local link target used in the workbook

Resolution order is deterministic:

1. exact `csfiles/home_dir/...`
2. exact export-root path
3. unique fallback filename match anywhere in the export package

If multiple files share the same fallback filename, the extractor does not guess. It preserves the raw ref and emits diagnostics.

Optional reviewer-portable asset copies:

```bash
.venv/bin/python brightspace_quiz_review_extractor_v2.py \
  /path/to/unpacked_export \
  --out ./quiz_review_out \
  --copy-images-to-assets
```

When `--copy-images-to-assets` is enabled, resolved images are copied into `assets/` beside the workbook and `image_link_primary` points to the copied file. Without that flag, the workbook links to the original file inside the export package.

Image-related diagnostics include:

- `malformed_image_ref`
- `unresolved_image_ref`
- `missing_image_file`
- `duplicate_image_filename`
- `image_copy_failed`

## Traceability and Diagnostics

Diagnostics are meant for reviewers, not just developers. They use quiz title context and plain language for issues such as:

- ambiguous bank matches
- unresolved bank-backed pools
- duplicate question IDs
- best-effort ordering or numeric parsing
- missing assets
- unresolved or missing image files
- partially collapsed feedback

`source_map` provides the audit trail from workbook or JSON rows back to the originating XML object and hint.

## Usage

### Run on an unpacked Brightspace export

```bash
.venv/bin/python brightspace_quiz_review_extractor_v2.py \
  /path/to/unpacked_export \
  --out ./quiz_review_out
```

### Run on an unpacked export and copy resolved images beside the workbook

```bash
.venv/bin/python brightspace_quiz_review_extractor_v2.py \
  /path/to/unpacked_export \
  --out ./quiz_review_out \
  --copy-images-to-assets
```

### Run on a Brightspace ZIP export

```bash
.venv/bin/python brightspace_quiz_review_extractor_v2.py \
  ./D2LExport.zip \
  --out ./quiz_review_out
```

### Create only the workbook template

```bash
.venv/bin/python brightspace_quiz_review_extractor_v2.py \
  --template-only \
  --out ./quiz_review_template_v2
```

## Source Repo and Releases

This repository is intended to track source, tests, and the workbook template.

It intentionally does **not** track:

- local virtual environments such as `.venv`
- unpacked sample course exports
- generated review outputs
- packaged `dist/` bundles



## Repository Scope

This repo tracks:

- source code
- tests
- workbook template

It does not track:

- `.venv`
- unpacked exports
- generated outputs
- `dist/` bundles

Use GitHub Releases for distributable builds.

## Limitations

- images are linked, not embedded, in the workbook
- the extractor only surfaces question-level image references; decorative course-content files outside quiz or question-bank content remain out of scope
- external or unsupported image URLs are preserved as raw refs and diagnosed instead of being rewritten
- fallback filename resolution only succeeds when there is exactly one package match
- ordering reconstruction remains bounded by the semantics Brightspace actually exports; when only partial structure is available, the workbook uses a clear best-effort display and preserves the raw payload

## Regenerating Outputs Locally

If you keep a local sample export outside git, regenerate outputs with:

```bash
.venv/bin/python brightspace_quiz_review_extractor_v2.py \
  /path/to/local_sample_export \
  --out ./quiz_review_out \
  --copy-images-to-assets
```

Refresh the template workbook with:

```bash
.venv/bin/python brightspace_quiz_review_extractor_v2.py \
  --template-only \
  --out ./quiz_review_template_v2
```

## Scope Guardrails

This extractor intentionally does **not**:

- generate re-import packages
- rebuild Brightspace quizzes
- interpret gradebook or rubric integrations
- add a service, UI, or migration framework

It is a local, review-first parser.

## Known Heuristics

- bank-only pool resolution remains heuristic when Brightspace does not expose stable section linkage
- ordering and numeric or arithmetic questions are preserved as best-effort structures when exact semantics are not explicit in export conditions
- `question_count_resolved` reflects reviewable question rows, which may include pool candidates when that is the only way to surface question content for review
