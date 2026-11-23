# Safety and QA

## Principles

1. **No tool deletes or moves an original.** Review stages write spreadsheets. The only stage that changes source files is the in-place rename (Stage 6b), and only after Gate B. Deletion is always done by a person.
2. **Preview, then confirm.** Every change is listed in full before it runs.
3. **Never guess.** A file a rule cannot decide is marked for review with a reason code.
4. **Leave a record.** Rename logs, old-to-new name maps and conversion checksums make every run auditable and reversible.

## Human gates

| Gate | After | A person must |
|---|---|---|
| A | Stages 2–4 | Open every "needs review" item, decide keep or delete, tick "reviewed" |
| B | Stage 6b preview | Read the full rename preview and type the confirmation phrase |
| C | Stage 7 | Review every flagged metadata field |

## QA checks on the CMS import sheet

A first QA pass on one series' import sheet (about 1,300 records, 89 columns) found recurring typos such as "Perfomance" and "Enseumble" in hundreds of cells, and category months that disagreed with event dates. Stage 8 turns those one-off checks into a repeatable script:

- spelling against a list of known typos
- controlled terms against the dropdown sheet
- start date before end date; end date blank for single-day events
- category year and month vs. filename year and month
- video duration field vs. the duration written in the description
- name and role columns filled in pairs

Every finding records the cell, the original value, the reason and a suggested fix. The script never edits the sheet.

## Publishing this repository

Nothing in this repository can be used to reconstruct the archive's file structure. Before each push:

- no network or local paths (`Aaa_nas`, drive letters, `/Users/`, `C:\Users\`)
- no API keys
- no real directory listings, rename logs, name maps or spreadsheets
- no institution-internal IDs; examples use synthetic data
- performers' names are removed from case studies
