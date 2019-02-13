# Manage-DSF
A set of scripts for automating product upkeep with EFI's Digital StoreFront.

## This project contains:
### Manage-DSF.psm1
PowerShell module containing various functions for scripts to use.
### Manage-DsfProduct.ps1
Using input from a CSV file, you can...
* Add or change details of non-printed products.  (Coming soon, ability to delete products.)
* Publish product to one or more categories.

## Planned:
### Manage-DsfCategory
Add, change, or remove storefront categories, using input from a CSV file.
### Manage-DsfGroup
Add, change, or remove user groups.
### Manage-DsfUser
Add, change, or remove user accounts.
### Manage-DsfKitProduct
Add, change, or remove kits that consist of other products.

## Goal
Bulk add, modify, delete non-printed products non-interactively using data from a spreadsheet.

## Current State
Alpha quality, but it's becoming usable.  Script can now add and update products, including thumbnail images.  (All images will be set to the same file.)

Security groups aren't handled yet.

## Compatibility
Digital StoreFront 9.8 series.

## Notes on input data
You may rearrange the spreadsheet columns however you like; data is sorted by column name, not position.

**Advanced Quantity** is a text field that supports regex-like notation, such as "5..20[5]|100|500" and will be entered verbatim from the input data.  This example would produce a drop-down list containing 5, 10, 15, 20, 100, 500.

To clear text fields, put `-` as the only data in the cell.  Currently, this works for:  Brief Description, Long Description, Notify Emails, Production Notes, and Keywords.

To publish a product into multiple categories, put them in the **Category** field and separate them with a semicolon (`;`) just like DSF's e-mail fields.

## Usage
First, search the file for `# Main site URL to start from` and replace the `$SiteURL` with your own site.  Open the Excel file, fill in your product info and export it to CSV.  Then, run something like the following.

```powershell
Manage-DsfProduct.ps1 -UserName fred -Password 'blah%293' -ProductFile 'C:\Somewhere\somedata.csv'
```

## Parameters
### -UserName
The Digital StoreFront account name you would use when logging into the site.
### -Password
Password for your DSF account.
### -ProductFile
Full path to a data file containing the products to be handled.  (Issue #17 provide a sample data file.)
### -SkipImageUpload
Causes the script to ignore any image paths provided in the input file; it will not touch the icon section of any product.
### -Debug
Causes the script to emit lots of detailed and possibly useful information about what it's doing.

## News
### 2019-02-05
Added ability to publish products into categories.
### 2019-01-29
Stabilized after module split.  `Manage-DsfProduct` can now set Advanced Quantity field in product settings.
### 2019-01-28
Very broken due to my first ever attempt at writing a PowerShell module.  Eventually there will be more than one script, so shared functions need to live in a module.
### 2019-01-16
Image upload works now!  `Manage-DSF` can upload an image that will replace all thumbnails for a product.
### 2019-01-14
To celebrate the first successful production run (updating threshold on 50 products in 24 minutes), I'm bumping the version up to 0.6-alpha!
