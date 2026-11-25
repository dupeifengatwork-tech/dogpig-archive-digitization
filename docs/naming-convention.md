# Naming convention (DP2)

Stage 6b applies Asia Art Archive's established naming convention for the second Dogpig deposit (DP2). Namer v1.0 turns the convention into explicit, testable rules.

```
<prefix><group>_<seq><ext>
DP2_99_04_200311a_01.jpg
DP2_99_04_20022004_07.jpg
DP2_99_04_00_03.pdf
```

| Part | Rule |
|---|---|
| prefix | One per series folder, `DP2_99_` + the series number (`04 …` → `DP2_99_04_`). Editable. |
| group | `YYYYMM` + a lowercase letter for each event in that month (`a`, `b`, `c` …), or a code assigned by a person (`00` for general documents, `20022004` for a multi-year folder). |
| seq | Two digits or more, counted within a group. |
| ext | Lowercase; `.jpeg` → `.jpg`, `.tiff` → `.tif`. |

Rules that came out of real folders:

- Existing DP2 names are kept and their letters and numbers are never reused; new files take the lowest free letter or number.
- Nested folders inherit the date of the nearest dated parent.
- Sibling folders under the same parent can share one group when a person says they are one event.
- Only archival formats get a number: `.jpg .tif .pdf .mp4 .m4a`. Anything else is marked `NOT_ARCHIVAL_FORMAT` and waits for conversion.
- A file with no date in its folder path is marked `DATE_NOT_DETERMINABLE`.
- An email's `<name>_attachments` folder follows the email PDF's new name; its contents are not numbered.

## Review codes

| Code | Meaning | What happens |
|---|---|---|
| `RENAME` | A new name is planned | Renamed after Gate B |
| `KEEP` | Already correctly named | Nothing |
| `REVIEW` | A rule cannot decide (no date, non-archival format, name collision, file directly under the root) | Left as is, with the reason |
| `SKIP` | System file | Ignored, takes no number |

## Safety rules

| ID | Rule |
|---|---|
| S-1 | Preview is read-only. |
| S-2 | Never overwrite: an occupied target name becomes `REVIEW`. |
| S-3 | Never guess: undecidable files become `REVIEW` and take no letter. |
| S-4 | Execution requires typing the confirmation phrase. |
| S-5 | Rescan before execution; any difference from the approved preview aborts with nothing changed. |
| S-6 | Check source size and target absence before each file; on mismatch, stop the whole run. |
| S-7 | Write the log line for each file immediately (append, flush, fsync). |
| S-8 | Case-only renames go through a temporary name in two steps, reverted on failure. |
| S-9 | Logs live in the tool's own folder, never in the archive. |
| S-10 | Rollback uses only the log, and only where the original name is still free. |
| S-11 | The tool has no code path that deletes a file. |

## Results

Two series have been renamed. The first (678 files) was renamed with the earlier planning script that Namer replaced, and the result was checked against a re-exported directory listing. The second was the first run of Namer v1.0: 1,654 previewed files, 1,448 renamed, 9 sent to review (7 with no determinable date, 2 needing conversion first), 197 skipped as system or attachment files. All 1,448 renames completed with no errors and passed verification.

Namer has 33 unit and integration tests. Tests that used a real directory listing are being replaced with synthetic data before they are published.
