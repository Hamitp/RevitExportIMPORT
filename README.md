# Hamit Schedule Sync

A Revit add-in for seamless two-way synchronization between Revit schedules and Excel files.

## Features
- **Export to Excel**: Exports the active Revit schedule to a formatted Excel file.
- **Import from Excel**: Reads modifications from the Excel file and updates the corresponding Revit elements.
- **Custom UI**: Integrates into the Revit ribbon via a custom "Hamit" panel with dedicated Export and Import buttons.

## Setup
1. Compile the project in Visual Studio.
2. Copy the `.addin` manifest file and the compiled `.dll` to your Revit Addins folder (usually `%PROGRAMDATA%\Autodesk\Revit\Addins\<Version>\`).
3. Launch Revit. You will find the tools under the "Hamit" tab in the ribbon.

## Usage
- Open a schedule view in Revit.
- Click **Export Schedule** to generate an Excel file.
- Make necessary updates in Excel.
- Click **Import Schedule**, select the modified Excel file, and apply the changes back to your Revit model.
