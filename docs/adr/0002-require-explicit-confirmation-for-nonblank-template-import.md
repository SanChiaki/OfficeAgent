# Require Explicit Confirmation for Nonblank Template Import

When a Work Sheet is not blank, OfficeAgent may still allow importing a Business Export Template into the current worksheet, but this overwrite path must not be selected by default and must clearly warn that existing worksheet content can be replaced. This protects local work while still allowing users to intentionally reset a sheet from a business template.

## Consequences

The initialization dialog should default to configuration-only initialization for nonblank worksheets. If the user chooses template import on a nonblank worksheet, the primary action and warning copy must make the destructive nature visible before execution.
