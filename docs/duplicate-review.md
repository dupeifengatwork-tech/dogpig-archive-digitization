# Duplicate review: one worked example

Tools propose; a person decides. This page shows how that worked on one review pass carried out on 30 September 2026; the flagged files came from two series, live music and theatre. Folder and file names are replaced with placeholders.

## What the review sheet records

Every candidate gets one row with six fields:

| Field | Content |
|---|---|
| Category | One of the categories below |
| Folder | Where the file sits |
| File | The file under review |
| Counterpart | The file it duplicates or resembles |
| Evidence and recommendation | Why the pair was flagged and what to do (EXIF time, pixel match, text overlap, file type) |
| Status | Not yet handled, kept, or deleted (with date) |

## Result of the pass

48 rows, in six categories.

| Category | Rows | Evidence used | Decision |
|---|---:|---|---|
| Identical photo | 1 | Same EXIF capture time to the second, identical pixels; the date matches only one of the two event folders | The copy in the wrong event folder can be deleted; left for a person to action |
| Word lock file (`~$*.doc`) | 3 | 162-byte temporary files left by Word | Can be deleted; left for a person to action |
| Near-duplicate documents | 3 | Text overlap of 83% and 86% for two pairs; one proposal fully contained in a longer one | Kept both in the two overlapping pairs (one is an excerpt or a newsletter version, used differently); for the third pair, removed two converted `.docx` copies and kept the original `.doc` and `.rtf` files |
| Similar but distinct photos | 39 | Burst shots or adjacent frames: different pose, expression, head angle or lighting | Kept all 39 |
| `.docx` with no original | 1 | The only surviving copy of a 2012 document | Kept |
| Other converted `.docx` copies | 1 row (9 files) | Short documents under 30 Chinese characters, too short for text comparison; each has a same-name `.doc` or `.rtf` original | Removed the `.docx` copies; originals kept |

## Rules this pass applied

- **An original outranks a converted copy.** When a `.docx` exists only because it was converted from a `.doc` or `.rtf`, the original stays.
- **A burst is not a duplicate.** Frames a second apart that differ in pose or expression are kept: they record different moments.
- **High overlap is not enough to delete.** An 86% overlap between a full proposal and its local excerpt means two documents with two uses, so both stay.
- **Location is evidence.** For an identical photo filed under two events, the capture date decides which folder it belongs to.
- **Nothing is deleted by a tool.** Two items marked deletable were still "not yet handled" at the end of the pass, waiting for a person.

The rules match the keep order the image review tool uses (original first, then resolution, then file size, otherwise review). See [workflow.md](workflow.md), Stage 3.
