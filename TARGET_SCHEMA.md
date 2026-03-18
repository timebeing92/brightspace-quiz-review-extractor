# TARGET_SCHEMA.md

## Purpose

This project is a **Brightspace quiz review extractor**.

It is designed to parse Brightspace course export packages and produce **review-oriented artifacts** for human and machine inspection.

This is **not** an LMS rebuild tool and **not** a Brightspace re-import tool.

The extractor should optimize for:

- human review
- auditability
- source traceability
- legibility
- stable machine-readable output
- practical use in VS Code on local files

---

## Primary Artifacts

The extractor should generate three outputs for each run.

### 1. Workbook (`quiz_review.xlsx`)
This is the **primary human-review artifact**.

It should be easy for reviewers to:

- filter
- sort
- scan question wording
- verify scoring
- inspect answer keys
- review feedback
- annotate issues

### 2. JSON (`quiz_review.json`)
This is the **canonical machine-readable artifact**.

It should contain the same core information as the workbook, in a stable normalized structure suitable for:

- future automation
- transformation
- comparison/diffing
- downstream tooling

### 3. Markdown Summary (`quiz_review_summary.md`)
This is a concise narrative overview.

It should summarize:

- quiz structure
- storage type
- delivery type
- points
- question counts
- pool counts
- major diagnostics

---

## Review-First Design Principles

The extractor should behave like a **review/audit pipeline**, not just a file copier.

It should:

1. resolve Brightspace quiz structures into human-usable form
2. preserve source traceability to XML/question bank origins
3. expose ambiguities explicitly in diagnostics
4. prefer completeness and clarity over minimal package extraction
5. preserve raw source context when full normalization is uncertain

When the source is ambiguous, the extractor should prefer:

- explicit diagnostics
- source references
- best-effort structured output

instead of silent guesses.

---

## Brightspace Structural Model

The extractor should assume three common quiz storage patterns.

### Type A — Inline / self-contained
The quiz contains questions directly in `quiz_d2l_*.xml`.

### Type B — Banked
The quiz structure is in `quiz_d2l_*.xml` and question content is in `questiondb.xml`.

### Type C — Hybrid
The quiz includes both inline question content and bank relationships.

The extractor should classify each quiz with:

- `storage_type`: `inline`, `banked`, `hybrid`, `unresolved`
- `delivery_type`: `fixed`, `pooled`, `mixed`, `unresolved`

---

## Required Output Fields

### Quiz-level required fields
These should appear in all relevant outputs.

- `quiz_id`
- `quiz_title`
- `source_quiz_file`
- `storage_type`
- `delivery_type`
- `has_questiondb`
- `time_limit_minutes`
- `attempts_allowed`
- `shuffle_questions`
- `shuffle_answers`
- `declared_total_points`
- `resolved_total_points`
- `question_count_resolved`
- `section_count`
- `pool_count`
- `instructions_text`

`quiz_title` is required everywhere a reviewer would otherwise need to infer identity from `quiz_id`.

### Section/pool required fields

- `quiz_id`
- `quiz_title`
- `section_id`
- `section_title`
- `section_order`
- `section_type`
- `draw_count`
- `pool_size`
- `section_points_total`
- `question_count`
- `source_quiz_file`
- `source_bank_file`

### Question required fields

- `quiz_id`
- `quiz_title`
- `section_id`
- `section_title`
- `question_order`
- `question_id`
- `question_title`
- `question_title_review`
- `question_type`
- `source_location`
- `source_quiz_file`
- `source_bank_file`
- `points`
- `has_image`
- `stem_text`
- `matching_review_display`
- `ordering_review_display`
- `correct_answer_key`
- `correct_answer_text`
- `all_correct_keys`
- `general_feedback`
- `correct_feedback`
- `incorrect_feedback`
- `answer_specific_feedback`
- `grading_notes`
- `asset_refs`
- `image_refs`
- `image_paths_resolved`
- `image_count`
- `image_link_primary`
- `review_status`
- `reviewer_notes`

### Diagnostics required fields

- `severity`
- `quiz_id`
- `quiz_title`
- `section_id`
- `question_id`
- `issue_type`
- `message`
- `source_file`
- `suggested_action`

### Source mapping required fields

- `object_type`
- `object_id`
- `object_title`
- `quiz_id`
- `quiz_title`
- `source_file`
- `source_hint`
- `resolved_to_sheet`
- `resolved_to_key`

---

## Workbook Design Requirements

The workbook must be optimized for actual reviewer use.

Required workbook behavior:

- freeze top row on each sheet
- enable filters on each header row
- wrap long text columns
- use a consistent sheet order
- preserve stable column ordering
- use reviewer-friendly column names
- keep one row per question on the main review sheet
- preserve reviewer-traceable image references without embedding binaries into the workbook
- keep reviewer-facing columns left and raw traceability columns right on the `questions` sheet
- include a reviewer-facing matching expansion sheet derived from the same parsed question data

Recommended sheet order:

1. `quiz_overview`
2. `sections_pools`
3. `questions`
4. `matching_pairs_expanded`
5. `pool_members`
6. `diagnostics`
7. `source_map`

The `questions` sheet is the primary review surface.

The `matching_pairs_expanded` sheet should preserve prompt order and provide one row per prompt/match pair for human review while leaving the raw `matching_pairs` field intact on the main question row.

---

## Question-Type Support Expectations

The extractor should strongly support these question types:

- multiple choice
- true/false
- multi-select
- matching
- ordering
- short answer
- fill-in-the-blanks
- numeric/arithmetic

If a question type cannot be fully normalized, the extractor should:

- preserve as much structured data as possible
- include raw payload fragments if useful
- log a diagnostic explaining the uncertainty

---

## Feedback Parsing Requirements

The extractor should keep these feedback channels separate whenever possible:

- `general_feedback`
- `correct_feedback`
- `incorrect_feedback`
- `answer_specific_feedback`

Do not collapse them into a single field unless unavoidable.

If collapsing is required, emit a diagnostic.

---

## Image Handling Requirements

Question-level image handling should preserve traceability rather than attempt aggressive transformation.

The extractor should:

- detect image references from Brightspace `mattext` HTML and `matimage` nodes
- scan question stems, answer choices, and question-level feedback/answer-key blocks
- preserve raw image refs separately from resolved local file paths
- prefer deterministic local resolution over fuzzy guessing
- emit diagnostics instead of guessing when filenames collide or paths are malformed

`image_link_primary` may point either to the original export-package file or to a copied output-side asset when the caller opts into asset copying.

---

## Bank and Pool Resolution Requirements

If `questiondb.xml` exists and a quiz uses pools or bank references, the extractor should attempt to resolve:

- pool draw counts
- candidate pool sizes
- bank question content
- source relationship (`inline`, `questiondb`, `hybrid`, `unresolved`)

If exact resolution is uncertain, preserve:

- question identifiers
- source file references
- raw structural hints
- diagnostics

---

## Diagnostics Philosophy

Diagnostics should be understandable by a human reviewer.

Required diagnostic categories include:

- unresolved bank reference
- duplicate question ID
- unsupported question structure
- points mismatch
- missing asset
- ambiguous answer parsing
- collapsed feedback parsing
- unresolved pool semantics
- bank question not linked to reviewed quiz

Diagnostics should explain what happened and what a reviewer should do next.

---

## Non-Goals

This project should **not** prioritize:

- Brightspace re-import packaging
- full LMS reconstruction
- gradebook/rubric reconstruction beyond quiz-review needs
- web app interfaces
- unnecessary framework complexity

This is a **review-first local tool**.

---

## Implementation Guidance

The code should remain:

- readable
- locally runnable in VS Code
- lightly documented
- practical rather than over-abstracted

Prefer:

- clear parsing functions
- explicit heuristics
- comments near Brightspace-specific logic
- predictable output schemas

When using heuristics, label them clearly in code and diagnostics.
