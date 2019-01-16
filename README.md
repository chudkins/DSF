# Manage-DSF.ps1
A script for automating product upkeep with EFI's Digital StoreFront.

## Goal
Bulk add, modify, delete non-printed products non-interactively using data from a spreadsheet.

## Current State
Alpha quality, but it's becoming usable.  Script can now add and update products, including thumbnail images.  (All images will be set to the same file.)

Categories and security groups aren't handled yet.

## Compatibility
Digital StoreFront 9.8 series.

## Notes on input data
To clear the E-mail Notification field, put `-` as the only data in the cell.

## Usage
### Test 3 hashes
#### Test 4 hashes

## News
### 2019-01-14
To celebrate the first successful production run (updating threshold on 50 products in 24 minutes), I'm bumping the version up to 0.6-alpha!
