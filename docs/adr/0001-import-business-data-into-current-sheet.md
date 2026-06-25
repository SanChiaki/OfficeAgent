# Import Business Data into the Current Sheet

When initializing a blank Work Sheet from a Business Export Template, OfficeAgent imports the `Business Data` sheet content into the current worksheet and preserves the current worksheet name. This keeps the Ribbon action "Initialize current sheet" aligned with the user's target sheet and avoids leaving behind an unused `Sheet1` next to a newly copied `Business Data` sheet.

## Considered Options

- Copy `Business Data` as a new worksheet and use the exported sheet name.
- Import the exported sheet content into the current worksheet and keep the current worksheet name.

## Consequences

The import implementation must preserve the useful content and formatting from the exported `Business Data` sheet without changing the identity of the current Work Sheet. Metadata such as sheet bindings and field mappings should be written for the current worksheet name.
