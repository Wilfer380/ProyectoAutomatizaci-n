# document-generation-pipeline Specification

## Purpose

Rebuild the Excel-to-Word generation flow while preserving the current UI, buttons, filters, window behavior, and manual-review sequence.

Non-goals: UI redesign, filter/visual-flow changes, or automatic printing.

## Requirements

### Requirement: Preserve current user flow and contract

The system MUST keep the existing user-facing flow unchanged and MUST keep the worker/controller contract compatible unless a breaking change is unavoidable.

#### Scenario: Existing flow remains available

- GIVEN the user opens the app and selects the same inputs as before
- WHEN generation starts
- THEN the same start, review, and cancel/continue flow remains available

### Requirement: Validate inputs before generating blocks

The system MUST validate the Excel path, Word template path, selected filter, required columns, record availability, image availability, and template placeholders before generating output.

#### Scenario: Missing or invalid input is rejected early

- GIVEN the Excel file, Word template, or selected filter is invalid
- WHEN the user starts generation
- THEN the process stops before any block is generated
- AND a user-friendly error is shown

#### Scenario: Template or data is incomplete

- GIVEN required columns, images, or expected placeholders are missing
- WHEN validation runs
- THEN the affected run is rejected or the affected block is skipped with a logged warning, depending on what can still be produced safely

### Requirement: Read Excel without a visible window

The system MUST read Excel in the background without showing a visible Excel window to the user.

#### Scenario: Excel is loaded headlessly

- GIVEN a valid Excel source file
- WHEN filters or records are loaded
- THEN Excel remains hidden from the user interface

### Requirement: Split selected records into blocks of 27

The system MUST split the selected records into blocks of at most 27 records and process each block independently.

#### Scenario: More than 27 records create multiple blocks

- GIVEN 28 selected records
- WHEN generation starts
- THEN the system creates 2 blocks
- AND no block exceeds 27 records

### Requirement: Generate one Word document per block

The system MUST create one Word document from the template for each block and replace `<img1>...<img27>` in order with the block images.

#### Scenario: Partial block clears unused placeholders

- GIVEN a block with fewer than 27 records
- WHEN the document is generated
- THEN only the used placeholders are populated in order
- AND the remaining placeholders are cleared

### Requirement: Support manual review and resilient continuation

The system MUST leave each generated document available for manual review before proceeding to the next block. It SHOULD continue with later blocks after recoverable missing-image or block-level failures when a usable document can still be produced.

#### Scenario: User reviews a generated block

- GIVEN a block document has been generated
- WHEN manual review is requested
- THEN the document remains openable for inspection and saving
- AND the process continues only after the review step completes

#### Scenario: A block fails but later blocks can still run

- GIVEN one block hits a recoverable generation error
- WHEN the run continues
- THEN the system logs the failure internally
- AND later blocks continue when possible

### Requirement: Log technical detail internally and show friendly errors

The system MUST log detailed technical failures internally and MUST present concise user-friendly messages to the UI.

#### Scenario: An internal exception occurs

- GIVEN a COM, file, or placeholder operation fails
- WHEN the failure is handled
- THEN the technical details are written to logs
- AND the UI receives a non-technical error message
