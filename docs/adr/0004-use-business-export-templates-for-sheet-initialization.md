# Use Business Export Templates for Sheet Initialization

The template list shown during "Initialize current sheet" template import comes from the business system for the active Project. It is independent from OfficeAgent's existing Sync Configuration Template feature, even though the user-facing UI may continue to use the word "template" for both concepts.

## Consequences

The implementation should use separate domain types, connector methods, dialogs, and tests for Business Export Templates. It must not reuse the local template catalog that backs "Apply Setting", "Save Setting", and "Save as Setting".
