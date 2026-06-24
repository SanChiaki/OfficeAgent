# Limit Template Import Cancellation to Download

Template import shows a progress dialog with a cancel action, but the first implementation only guarantees cancellation while downloading the business-system Excel export. Once OfficeAgent starts applying the exported `Business Data` sheet through Excel COM, cancellation is not guaranteed because interrupting workbook copy operations safely is significantly harder than aborting the network I/O.

## Consequences

The progress UI should make the current phase clear. Cancel should abort the export download and leave the current Work Sheet unchanged, but after the import enters Excel COM copy/application work the UI should no longer imply that the operation can be safely stopped.
