# P&L A08 Project — Read This First

## Current decision

**NO-GO for production.** The architecture is substantially implemented, but
the real NASCA workbook has not completed the required Windows Excel COM proof.

## Repaired in this review

- The missing working-copy fingerprint that guaranteed execution failure.
- The incorrect September expectation (`A09`) was changed to the correct `T09`.
- The diagnostic now handles the leading `★` workbook character safely.
- The idempotency integration test now reruns the produced A08 workbook and
  requires explicit execution approval.
- A dependency-free quick protection/container probe now routes NASCA files to
  attach mode and normal XLSB files to isolated Excel automatically.
- A portable Windows build now bundles Python and the runtime packages, so the
  production PC does not need Python, pip, internet, admin rights, or PATH edits.
- A self-check verifies Windows, 64-bit runtime, Excel registration, COM
  packages, configuration support, and a writable work folder before Excel use.
- A locked offline wheelhouse, SHA-256 manifest, private repair environment,
  and deterministic one-folder build scripts were added.
- The first general Excel-agent layer now exposes a machine-readable catalogue
  of 21 capabilities without pretending planned tools are already callable.
- Verified byte backups and workbook inventory/full JSON snapshots are now
  available for recognized Excel files, including NASCA attach routing.

## Still required

1. ~~Clean full unit-test run with zero failures.~~ **Completed: 59 passed.**
2. Read-only diagnostic and dry-run on the authorized Windows/NASCA PC.
3. Complete Total PL row-lineage audit.
4. Disposable-copy COM execution, calculation, reconciliation, save, close,
   reopen, and corruption check.
5. Idempotency and forced-failure tests.
6. Finance-owner comparison and written production approval.
7. Build the portable executable from the supplied verified offline wheelhouse
   on Windows, then run its self-check. PyInstaller cannot produce a Windows
   executable from this Linux review environment.

The authoritative updated plan is:

`PL_A08_AUTOMATION_DETAILED_EXECUTION_PLAN_V2_STATUS_UPDATED.md`
