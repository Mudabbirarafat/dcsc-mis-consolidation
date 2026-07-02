# Extension Note

To extend this from the sample pack to the remaining airline formats, I would keep the consolidation engine stable and move airline differences into configuration plus small adapters.

## Recommended design

1. **Core scanner**
   - Recursively discovers Excel files.
   - Derives stable metadata from folder, filename, workbook, and sheet names.
   - Sends each workbook to the best matching parser.

2. **Config-driven schema mapping**
   - Maintain a YAML/JSON config per airline/direction with:
     - accepted detail sheet names or regex patterns,
     - header aliases,
     - mandatory fields,
     - amount/rate sheet names,
     - Service Challan total rules,
     - known quirks such as in-workbook fortnights or combined AWB+suffix.

3. **Adapter layer**
   - Use the generic parser for normal annexure-style files.
   - Add lightweight adapters only where the structure differs materially, such as transfer/COSYS layouts, master invoice workbooks, or import lifecycle files.
   - Each adapter should return the same canonical row schema so downstream summary/reconciliation code never changes.

4. **Validation and reconciliation**
   - Validate every output row for required metadata, numeric weights, AWB format, and duplicate uniqueness key.
   - Reconcile at source-file level and optionally at flight/date/service level.
   - Store mismatch reason codes: missing total, rounding tolerance, duplicate source, unmapped service, missing rate, or adapter required.

5. **Onboarding a new airline**
   - Add two or three sample files to a fixture folder.
   - Run a sheet/header profiler to identify detail tabs, total tabs, and rate tables.
   - Add config aliases first.
   - If config is insufficient, add a small adapter function and a regression test.
   - Confirm output row count, month summary totals, duplicate keys, and Service Challan reconciliation before release.

This approach keeps new airline effort predictable: most formats should need config only, while unusual layouts get isolated adapters without making the common parser brittle.
