# Manage-DSF.ps1
A script for automating product upkeep with EFI's Digital StoreFront.

## Goal
Bulk add, modify, delete non-printed products non-interactively using data from a spreadsheet.

## Current State
Alpha quality, at best.  Using Selenium library, I'm working through the DSF web forms in order to automate these tedious tasks.

## Compatibility
Digital StoreFront 9.8 series.

## Notes on input data
To clear the E-mail Notification field, put `-` as the only data in the cell.