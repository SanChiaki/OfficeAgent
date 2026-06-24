# OfficeAgent Context

This context defines the domain language for OfficeAgent workbook setup, project synchronization, and template-driven sheet initialization.

## Language

**Project**:
A business-system work scope that owns the rows, fields, and synchronization target for a worksheet.
_Avoid_: Workbook, file, account

**Business Export Template**:
A business-system template selected with a Project to export an Excel workbook whose data sheet is the authority for initial worksheet content.
_Avoid_: Local template, configuration template

**Business Export Template ID**:
A stable business-system identifier used to request an exported workbook for a Business Export Template.
_Avoid_: Template name, display label

**Business Data Sheet**:
The worksheet named `Business Data` inside a business-system export workbook that OfficeAgent imports as initial Work Sheet content.
_Avoid_: First sheet, active sheet, guessed data sheet

**Sync Configuration Template**:
A local OfficeAgent template that stores reusable synchronization configuration such as layout and field mappings.
_Avoid_: Business export template, exported Excel

**Work Sheet**:
The worksheet in the user's active workbook that is bound to a Project and used for editing, downloading, and uploading business data.
_Avoid_: Export workbook, template file

**Blank Work Sheet**:
A Work Sheet with no user-entered or imported cell content. Formatting alone does not make a Work Sheet nonblank.
_Avoid_: Empty UsedRange, unformatted sheet

**Settings Sheet**:
The workbook sheet that stores OfficeAgent metadata such as project bindings, field mappings, and template bindings.
_Avoid_: Work sheet, data sheet

**Field Mapping**:
The metadata that connects visible worksheet headers to business-system field identities used for download and upload.
_Avoid_: Column name, display header
