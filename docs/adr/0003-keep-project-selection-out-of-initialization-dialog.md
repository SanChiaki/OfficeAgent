# Keep Project Selection out of the Initialization Dialog

OfficeAgent keeps Project selection in the Ribbon project dropdown instead of duplicating it in the initialization dialog. If the user clicks "Initialize current sheet" without a selected Project, the existing project-required prompt remains the entry point, and the initialization dialog focuses only on configuration-only initialization versus template import for the already selected Project.

## Consequences

The initialization dialog does not need to load or search Projects. It should list only templates for the active Project and rely on the existing Ribbon project state as the source of truth.
