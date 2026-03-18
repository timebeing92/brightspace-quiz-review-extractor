# CHANGELOG

## Brightspace Quiz Review Extractor v2 Refresh

### Review-first schema and traceability

- Added `quiz_title` across all reviewer-facing row types and outputs.
- Standardized `source_hint` usage and removed the older mixed `source_xpath_or_hint` naming.
- Expanded `source_map` rows to carry `quiz_id` and `quiz_title`.
- Preserved `source_quiz_file` and `source_bank_file` consistently across question and pool outputs.

### Source classification and bank resolution

- Fixed `storage_type` so it is based on resolved quiz/question sources instead of `questiondb.xml` presence alone.
- Standardized row-level `source_location` values to `inline`, `hybrid`, `questiondb`, and `unresolved`.
- Added explicit diagnostics for ambiguous bank matches and unresolved bank-backed pool sections.
- Preserved pool draw counts and candidate pool sizes in workbook and JSON outputs.

### Parsing improvements

- Improved matching parsing to keep prompt/option structure and correct pairs.
- Added reviewer-facing `matching_review_display` output that preserves source prompt order while leaving raw `matching_pairs` intact.
- Added `matching_pairs_expanded` workbook/JSON output with one row per prompt/match pair, including `prompt_order` and `match_order`.
- Improved short answer and fill-in-the-blank parsing for multiple accepted values, case flags, and interleaved prompt text.
- Improved ordering parsing so ordered `response_grp` exports can be reconstructed into reviewer-facing numbered sequences when Brightspace preserves ranking conditions.
- Added `ordering_review_display` so ordering questions are readable in the workbook, including placeholder labels such as `Item 1` when sequence positions exist but labels are missing.
- Preserved numeric/arithmetic and remaining partial ordering structures as best-effort payloads with clear diagnostics when full semantics are not exposed.
- Improved feedback handling so general, correct, incorrect, and answer-specific feedback remain separate when possible.
- Added deterministic question-level image detection from Brightspace `mattext` HTML and `matimage` nodes across stems, choices, feedback, and answer keys.
- Added reviewer-facing question fields for raw image refs, resolved package paths, resolved image counts, and a primary local image link.
- Added optional `--copy-images-to-assets` output mode so workbook links can point to copied local assets instead of the original export bundle.
- Added explicit diagnostics for malformed image refs, unresolved image refs, missing image files, duplicate fallback filenames, and image-copy failures.

### Workbook usability

- Enabled filters on every worksheet.
- Kept frozen header rows and wrapped text across the workbook.
- Added `question_title_review` so internal ID-like titles fall back to short stem-derived reviewer titles without overwriting raw `question_title`.
- Added `has_image` reviewer filtering and moved `image_link_primary` into the reviewer-facing column group near image status columns.
- Reordered `questions` so reviewer-facing columns stay left and raw/traceability fields move farther right with separate header styling.
- Added reviewer-facing matching and ordering display columns so reviewers do not need to inspect JSON payloads for common audit tasks.
- Added the `matching_pairs_expanded` worksheet directly after `questions` for one-row-per-pair matching review.
- Applied stable sheet ordering, stronger visual separation between reviewer and raw columns, more useful widths for long review fields, and minimum row heights for matching, ordering, long-answer, and image-bearing questions.
- Improved diagnostics readability with severity-aware styling.

### Testing and artifacts

- Added focused `pytest` coverage for parser behavior, bank-match confidence handling, and integration output checks.
- Refreshed the sample output bundle in `quiz_review_example_v2`.
- Refreshed the template workbook in `quiz_review_template_v2`.

### Remaining heuristics

- Bank-only pool resolution is still heuristic when Brightspace does not expose stable section linkage.
- Ordering and numeric/arithmetic questions remain best-effort when exports only preserve partial scoring semantics.
- When the source is ambiguous, the extractor now prefers diagnostics and preserved clues over silent assumptions.
