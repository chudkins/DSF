# Manage-DSF
A set of scripts for automating product upkeep with EFI's Digital StoreFront.

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
### `Manage-DsfProduct.ps1 -UserName fred -Password 'blah%293' -ProductFile 'C:\Somewhere\somedata.csv'`

## Parameters
### -UserName
The Digital StoreFront account name you would use when logging into the site.
### -Password
Password for your DSF account.
### -ProductFile
Full path to a data file containing the products to be handled.  (#17 provide a sample data file.)
### -SkipImageUpload
Causes the script to ignore any image paths provided in the input file; it will not touch the icon section of any product.
### -Debug
Causes the script to emit lots of detailed and possibly useful information about what it's doing.

## News
### 2019-01-28
Very broken due to my first ever attempt at writing a PowerShell module.  Eventually there will be more than one script, so shared functions need to live in a module.
### 2019-01-16
Image upload works now!  `Manage-DSF` can upload an image that will replace all thumbnails for a product.
### 2019-01-14
To celebrate the first successful production run (updating threshold on 50 products in 24 minutes), I'm bumping the version up to 0.6-alpha!
