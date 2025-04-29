# Word Document Analysis Tool Add-in

This Word Add-in allows you to analyze a document by assigning unique IDs to paragraphs and tables, and displaying their content in a list.

## Features

- **Analyze Document**: Assigns unique IDs to paragraphs and tables
  - Each paragraph (not inside a table) gets a unique ID with format `para-<timestamp>-<random>`
  - Each table is treated as one block with unique ID format `table-<timestamp>-<random>`
  - Shows list of all unique IDs with their text content

- **Delete Context**: Removes all content controls and resets the state
  - Cleans up all content controls while preserving document content
  - Allows re-analysis with new IDs

## Requirements

- Word (Office) 2016 or later, or Microsoft 365
- For development: Node.js

## Running the Add-in

### Development

1. Clone the repository
2. Run `npm install` to install dependencies
3. Run `npm start` to start the development server and sideload the add-in in Word

### Usage

1. Open a Word document
2. Click the "Show Task Pane" button in the ribbon to open the add-in
3. Click "Analyse Document" to add unique IDs to paragraphs and tables
4. View the list of IDs and content in the task pane
5. Click "Delete Context" to remove all IDs and reset

## Implementation Details

- Uses Office.js API to interact with Word documents
- Content controls are used to mark and identify paragraphs and tables
- React is used for the UI components
- Each element gets exactly one content control (no nesting or duplication)
- Paragraphs inside tables are not assigned IDs

## Troubleshooting

- If content controls are not being properly removed, try:
  1. Click "Delete Context"
  2. Save and reopen the document
  3. Try the analysis again

## Copyright

Copyright (c) 2024 Microsoft Corporation. All rights reserved.

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**