#requires -Module Selenium

<#
	.Synopsis
	Given a CSV file, add or update products in EFI's Digital Storefront.
	
	.Parameter ProductFile
	The file containing product details.
	
	.Parameter UserName
	Account name to use when signing into the web site.
	
	.Parameter Password
	Password associated with the account.
	
	.Parameter SkipImageUpload
	Don't upload product thumbnail images.  Default is to upload them if a valid file path is provided.
	
	.Parameter Debug
	Emit lots of information in the hope of aiding troubleshooting.
#>

<#	Selenium class documentation for .NET:
		https://seleniumhq.github.io/selenium/docs/api/dotnet/index.html
#>

<#
	List of web sites set up specifically for automation testing!
		https://www.ultimateqa.com/best-test-automation-websites-to-practice-using-selenium-webdriver/
#>

<#	To install Selenium, I had to set up NuGet.org as a package source.
	After hours, I found the correct instructions on that here:  https://github.com/OneGet/oneget

	Register-PackageSource -name "Nuget.org" -providername NuGet -Location https://www.nuget.org/api/v2
	Set-PackageSource -Name "Nuget.org" -Trusted
	$pkg = find-package -Source Nuget.org -Name "selenium*"
	$pkg | where Name -eq "Selenium.WebDriver" | install-package
	$pkg | where Name -eq "Selenium.Support" | install-package
	$pkg | where Name -eq "Selenium.WebDriver.IEDriver" | install-package
	#$pkg | where Name -eq "Selenium.WebDriver.IEDriver64" | install-package
	
	To update Selenium, we need to grab currently installed packages and install them:
	
	find-package -Source Nuget.org -Name "Selenium.WebDriver" | install-package -InstallUpdate
	find-package -Source Nuget.org -Name "Selenium.Support" | install-package -InstallUpdate
	find-package -Source Nuget.org -Name "Selenium.WebDriver.IEDriver" | install-package -InstallUpdate
	find-package -Source Nuget.org -Name "Selenium.WebDriver.IEDriver64" | install-package -InstallUpdate
	find-package -Source Nuget.org -Name "Selenium.WebDriver.GeckoDriver.Win64" | install-package -InstallUpdate
#>

#[cmdletbinding()]

Param (
	[ValidateScript({
		if ( ( $_ -eq $null ) -or ( test-path $_ ) ) {
			$true
		} else {
			throw "ProductFile - Supplied path not found: $_!"
		}
	})]
	[string] $ProductFile,

	[ValidateNotNullOrEmpty()]
	[string] $UserName = "DefaultUser",

	[ValidateNotNullOrEmpty()]
	[string] $Password,
	
	[switch] $SkipImageUpload,
	
	[switch] $Debug
)

Begin {
	<#	Things in here will be done only once per run of this script.
		This is sort of like a class constructor, but geared more toward functions that process
		pipeline input.  Those functions would have some setup in the BEGIN block, then everything
		in the PROCESS block runs once per object received, and finally the END stuff runs for any
		final actions, sort of like a destructor.
		
		I've chosen to use Begin/Process/End in case this script evolves to handle multiple inputs,
		which could very well be useful, but even if that never happens it makes sense to me.
		Setup, such as global variables and function declarations, goes in BEGIN.  Main loop goes
		in PROCESS, followed by cleanup and summaries in END.
	#>

	# Put setup stuff BELOW functions!

[string]$ScriptLocation = Split-Path $MyInvocation.MyCommand.Path	# Script is in this folder
$ManageDSFModule = join-path $ScriptLocation "Manage-DSF.psm1"
Import-Module $ManageDSFModule

Function Handle-Exception {
	# Custom error handling for this script
	param (
		$Exc
	)

	if ( $Exc.Exception.WasThrownFromThrowStatement ) {
		# Throw statement means we did it on purpose.  Examine error code and
		#  print appropriate message.

		write-log -fore mag "Caught custom exception:" 
		write-log -fore mag $( $Exc | fl * -force | out-string )
		
		$exMsg = $Exc.Exception.Message

		switch -wildcard ( $exMsg ) {
			"BFCache not set!" { 
				Write-Log -fore mag $exMsg 
				Write-Log -fore yel "ERROR: Registry entry is required for proper operation!"
				Write-Log -fore yel "Please run the following in an Administrator window:"
				Write-Log -fore cyan '$IEFeatureControl = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Internet Explorer\Main\FeatureControl"'
				Write-Log -fore cyan '$BFCache = "FEATURE_BFCACHE'
				Write-Log -fore cyan 'New-Item ( join-path $IEFeatureControl $BFCache )'
				Write-Log -fore cyan 'New-ItemProperty ( join-path $IEFeatureControl $BFCache ) -Name "iexplore.exe" -Value 0 -PropertyType Dword'
				Write-Log " "
				Write-Log -fore yel $( get-psdrive -psprovider filesystem | format-table -property name, root, description -autosize | out-string )
			}
			"Timed out while waiting*" {
				Write-Log -fore mag $exMsg 
				Write-Log -fore red "Time limit exceeded while waiting for web form element."
				Write-Log -fore red "Execution cannot continue."
			}

			default { 
				write-log -fore red "Unhandled exception:" 
				write-log ( $Exc.Exception | fl | out-string )
			}
		}
	
	} else {
		write-log -fore mag "Caught standard exception:"
		write-log -fore mag $_.Exception.ErrorRecord.Exception
		write-log -fore gray $( $_ | fl * -force | out-string )
	}

}

function Find-Product {
	<#
		.Synopsis
		Given an object containing product details, navigate to the product details page.
		
		.Description
		This function takes a custom object populated with all details pertaining to a DSF product.
		It will then search the system for that product, and return a link to the product details page
		if it's found.  If product is not found, it will log an error and return nothing.
		
		.Parameter Product
		Custom object containing the details of a product, typically populated via spreadsheet import.
		
		.Parameter BrowserObject
		Selenium driver object representing the browser we're automating.
	#>

	param (
		[Parameter( Mandatory, ValueFromPipeLine )]
		[OpenQA.Selenium.Remote.RemoteWebDriver] $BrowserObject,
		
		[Parameter( Mandatory )]
		[PSCustomObject] $Product
	)
	
	$Fn = (Get-Variable MyInvocation -Scope 0).Value.MyCommand.Name
	Write-DebugLog "${Fn}: Search for product '$( $Product.'Product Id' )'"
	
	<#
		To find a product, we need to...
			Go to Products page.
				Select Products from picklist.
					ctl00_ctl00_TabNavigatorSFAdministration_QuickMenuSearch
			Search for it.
				Search box
					ctl00_ctl00_C_M_TextBoxSearch
				Enter product SKU
				Click Search button
					ctl00_ctl00_C_M_ButtonSearch
			Look in the resulting table for <a> where ID matches "*_HyperLinkManageProduct" and link text 
			matches product Name.
				If search term isn't found, you get a result page with a table that has headers only.
				So, check for an empty table and log failure.
			Click that link and it should take us to the product details page.
			Verify the details page loaded.  If <div class="ctr-bc-links"> exists, we're good.
	#>
	
	[bool] $FoundResult = $false
	
	# Navigate to Products list page.
	#$OpsList = $BrowserObject | Get-Control -Type List -ID "ctl00_ctl00_TabNavigatorSFAdministration_QuickMenuSearch"
	#$OpsList | Select-FromList "Products"
	$ManageProductsURL = $SiteURL + "Admin/ManageProducts.aspx"
	Enter-SeUrl $ManageProductsURL -Driver $BrowserObject
	
	# Search for requested product.
	$SearchBox = $BrowserObject | Get-Control -Type Text -ID "ctl00_ctl00_C_M_TextBoxSearch"
	$SearchButton = $BrowserObject | Get-Control -Type Button -ID "ctl00_ctl00_C_M_ButtonSearch"
	Set-TextField $SearchBox $Product.'Product Id'.Trim()
	Click-Link $SearchButton
	
	# Wait for results.
	$ResultsTable = WaitFor-ElementExists -WebDriver $BrowserObject -ID "ctl00_ctl00_C_M_GridProducts"
	if ( $ResultsTable ) {
		# We got something back, however it may not have any products listed.
		Write-DebugLog "${Fn}: Got a result table back."
		# Verify table actually contains results by counting the rows that are in "bg-AdS-001000" class.
		# The table header has a different class name.
		$ResultHitRows = $ResultsTable.FindElementsByClassName("bg-AdS-001000")
		$ResultCount = ( $ResultHitRows | Measure-Object ).Count
		if ( $ResultCount -ge 1 ) {
			# Table has some result rows, meaning we got some hits back.
			Write-DebugLog "${Fn}: Got $ResultCount results back."
			# Check through the rows and find the one where ID exactly matches our Product.
			foreach ( $row in $ResultHitRows ) {
				# Enumerate links for troubleshooting.
				foreach ( $link in $row.FindElementsByTagName("a") ) {
					Write-DebugLog "Links found:"
					Write-DebugLog "`tID $( $link.GetProperty('id') )"
					Write-DebugLog "`tHref $( $link.GetProperty('href') )"
					Write-DebugLog "`tText $( $link.Text )`n"
				}
			#>
				# FindElementByLinkText will throw an exception if nothing is found.
				# Catch this and continue.
				try {
					if ( $row.FindElementByLinkText($Product.'Product Id'.Trim()) ) {
						$ProductFoundRow = $row
						$FoundResult = $true
						break
					}
				}
				catch [OpenQA.Selenium.NoSuchElementException] {
					Write-DebugLog "${Fn}: Element not found in current result row."
				}
				catch {
					Handle-Exception $_
				}
			}
			# Extract the product management link.
			$ProductLink = $ProductFoundRow.FindElementByTagName("a") | Where-Object { $_.GetProperty("id") -like "*_HyperLinkManageProduct" }
		} else {
			Write-DebugLog "${Fn}: Table doesn't seem to contain any hits."
		}
	} else {
		# We got nothing back, which probably means WaitFor-ElementExists timed out.
		Write-DebugLog "${Fn}: Something went wrong trying to retrieve search results."
	}
	
	$ProductLink
}

function Get-PriceRow {
	<#
		.Synopsis
		Find a pricing row on the product details page; output a WebElement containing it.
		
		.Parameter WebDriver
		WebDriver object to search within.
		
		.Parameter PriceSheetName
		Name, such as "Contoso Base Price Sheet" that identifies the table containing the price row
		you want.
		
		.Parameter RangeStart
		Integer representing the beginning of the range we want.  Default is 1.
		
		.Example
		$PriceElement = $Browser | Get-PriceRow "My Price Sheet" 101
		
		Returns the row in "My Price Sheet" where the range starts at 101.
	#>

	param (
		[Parameter( Position=1, Mandatory, ValueFromPipeLine )]
		[OpenQA.Selenium.Remote.RemoteWebDriver] $WebDriver,

		[Parameter( Position=2, Mandatory )]
		[ValidateNotNullOrEmpty()]
		[string] $PriceSheetName,
		
		[Parameter( Position=3 )]
		[int] $RangeStart = 1
	)
	
	<#	Pricing Structure
	
		Even if a product has no price yet, it will have at least one Price Sheet with one row in it.
	
		Price tables are contained in <div id="ctl00_ctl00_C_M_ctl00_W_ctl01_PricingPanel">.
		
		Now things get tricky because these are dynamically generated and will be different for each
		DSF setup.  Hard-coding cell names or whatever will cause the script to break if anything changes.
		Therefore, find the right parts relative to the page structure.
		
		Within that, each price sheet seems to be in <div class="ctr-contentcontainer" style="margin-top:0px;">.
		There will be one for each price sheet shown.
		
		Default "price sheet" can be addressed by name, if you get all "ctr-contentcontainer" and then find the
		one containing a "span" with text equal to whatever your price sheet is named.
		I think we could do this by getting each one into a WebElement, then searching within that for 
		the matching span.
		
		Price Sheet name must be a variable.  Options for how it's set include:
			Global script variable.
			Script parameter.
			Product property in input data.  (This might be best if we want to allow modifying different sheets.)
		
	#>
	
	<#
		Finding the right row:
			Get all tables with class="border-Ads-000001".  This is used for other tables besides
			price sheets.
			
			Within each table, find <td class="bg-Ads-010000">.  This is used only for price sheets.
			
			Within price sheet, find <span class="bold">$PriceSheetName</span> to get the right one.
			
			Inside that, row can be found using input fields.  Range Begin field will have an ID matching
			"*_rngbegin_*", and its value will be set to an integer we can match.
				In the row, Regular Price has ID matching "*_PriceCatalog_regularprice_*".
				Setup Price ID matches "*_PriceCatalog_setupprice_*".
	#>
	
	try {
		$Fn = (Get-Variable MyInvocation -Scope 0).Value.MyCommand.Name
		Write-DebugLog "${Fn}: Get price row in '$PriceSheetName' with range starting at $RangeStart."

<#
		# Find all objects with class="border-Ads-000001".
		$colTables = $WebDriver.FindElementsByClassName("border-Ads-000001")
		#$colTables = $WebDriver.FindElementsByTagName("table") | Where-Object { $_.GetProperty("class") -eq "border-Ads-000001" }
		
		Write-DebugLog "${Fn}: Got $( ($colTables | Measure-Object).Count ) document tables, $( $colTables.GetProperty('id') )"
		
		# Now search through those for a price sheet.
		# We can't just use FindElementsByTagName because that will give us a <td> when we really 
		#	want the entire table containing this data.
		[OpenQA.Selenium.Remote.RemoteWebElement[]]$colPriceSheets = $null
		foreach ( $table in $colTables ) {
#			if ( ( $table.FindElementsByTagName("td") | Where-Object { $_.GetProperty("class") -eq "bg-Ads-010000" } ) -notlike $null ) {
			if ( $table.FindElementsByClassName("bg-Ads-010000") -notlike $null ) {
				$colPriceSheets += $table
			}
		}
#>
		# Find the div that holds the price sheets.
		$PriceSheetGrid = $WebDriver.FindElementByID("ctl00_ctl00_C_M_ctl00_W_ctl01_GridViewPricesheetsUpdatePanel")
		Write-DebugLog "${Fn}: Got $( ($PriceSheetGrid | Measure-Object).Count ) price grids."
		
		# Within that, there will be some number of <div class="ctr-contentcontainer" ...>, one for each price sheet.
		$colPriceSheets = $PriceSheetGrid.FindElementsByClassName("ctr-contentcontainer")
			
		Write-DebugLog "${Fn}: Got $( ($colPriceSheets | Measure-Object).Count ) price sheets:"
		foreach ( $sht in $colPriceSheets ) {
			Write-DebugLog "$( $sht.ToString() )`n"
		}
		
		# One of these should have $PriceSheetName in a span.
		foreach ( $sheet in $colPriceSheets ) {
			# Check each element in collection to see if it contains a span matching $PriceSheetName.
			if ( $sheet.FindElementByTagName("span") | Where-Object { $_.Text -eq $PriceSheetName } ) {
				$PriceSheetSubGrid = $sheet
			}
		}
		
		# Within this, get the contained table that holds the actual rows.
		$PriceSheet = $PriceSheetSubGrid.FindElementByClassName("bg-AdS-000110") | Where-Object { $_.GetProperty("id") -like "*_GridViewPricesheets_ctl*_PriceItemFrame_*" }
		
		Write-DebugLog "${Fn} Got final price sheet, ID = $( $PriceSheet.GetAttribute('id') )"
		
		# Now we've got the right sheet; find the row based on the start of the range.
		# Again, FindElementByTagName is going to get the actual element, an input field in this case.
		# Search the rows in $PriceSheet to find the row containing the range start value we need.
		
		# Get rows in this table.  They don't necessarily have the class ID attached.
		$rows = $PriceSheet.FindElementsByTagName("tr")
		
		# This should get the actual <tr> containing the input fields.
		:rowscan foreach ( $row in $rows ) {
			# Iterate through the <td> in the row.
			foreach ( $td in $row.FindElementsByTagName("td") ) {
				# The <td> we want has an <input> with ID matching "_rngbegin_" and value set to range start target.
				if ( $td.FindElementsByTagName("input") | Where-Object { ( $_.GetProperty("id") -like "*_rngbegin_*" ) -and ( $_.GetProperty("value") -eq $RangeStart ) } ) {
					# Found the row!
					$PriceRow = $row
					# Don't bother checking any others.
					break rowscan
				}
			}
		}
		
		if ( $PriceRow ) { 
			Write-DebugLog "${Fn} Got price row with range starting at $( $RangeStart ): $( $PriceRow.ToString() )"
			#Write-DebugLog ( $PriceRow | out-string )
			# For debugging, iterate through each <td> and output IDs of any input objects found in them.
			Write-DebugLog "Price row contains these Input fields:"
			foreach ( $td in $PriceRow.FindElementsByTagName("td") ) {
				Write-DebugLog ( ( $td.FindElementsByTagName("input") ).GetProperty("id") | out-string )
			}
		} else {
			# $PriceRow still uninitialized by now, so we didn't find anything.
			Write-DebugLog "${Fn} Failed to find price row with range starting at $( $RangeStart )"
		}
		
		$PriceRow
	}
	
	catch {
		throw $_
	}
	
	finally {}
	
}

function Manage-Product {
	<#	
		.Synopsis
		Add or modify a product in DSF.
		
		.Description
		Add or change a product in DSF.  Once added, pass to Update-Product which will handle all the details.
		
		.Parameter Product
		Object containing the name and other properties of the target product.
		
		.Parameter BrowserObject
		Object containing the web browser object we're using.
		
		.Parameter Mode
		Add (create new product), Change (modify existing product), Skip (ignore this row).
		May include "Delete" in future.
	#>

	param (
		[Parameter( Mandatory )]
		[PSCustomObject] $Product,
		
		[Parameter( Mandatory )]
		[OpenQA.Selenium.Remote.RemoteWebDriver] $BrowserObject,
		
		[Parameter( Mandatory )]
		[ValidateSet("Add","Change","Skip")]
		$Mode
	)
	
	$Fn = (Get-Variable MyInvocation -Scope 0).Value.MyCommand.Name

	# Need code that will handle Add mode or Change mode.
	# For each property, if Change, check if a value was supplied; if so, change it, otherwise leave it.
	# This is probably best handled by a function so we don't need an "if" on every line.
	# We want to avoid having multiple functions that each have to be rewritten for a DSF change.
	
	switch ( $Mode ) {
		"Add"	{
			write-log "${Fn}: Add product, $($Product.'Product Id')"
			
			# Check if product already exists and refuse to add if so.
			$ExistingProduct = Find-Product -BrowserObject $BrowserObject -Product $Product
			if ( $ExistingProduct ) {
				Write-Log -fore red "Warning: '$( $Product.'Product Id' )' already exists; not added!"
				continue
			} else {
				# Press Create Product button, but check for system message if it already exists.
				Click-Link ( $BrowserObject | Get-Control -Type Button -ID "ctl00_ctl00_C_M_ButtonCreateProduct" )
				
				# If this is a dupe, system will return the same page again with a text message.
				# Both of these will be true:
				#	Still on same page.
				#	"ItemNameMustBeUnique" span is visible.
				if ( ( $BrowserObject.Url -like "*/Admin/CreateNewCatalogItem.aspx" ) -and ( $BrowserObject.FindElementByID("ctl00_ctl00_C_M__ItemNameMustBeUnique").Displayed -eq $true ) ) {
					Write-Log -fore red "Error: '$( $Product.'Product Id' )' not added because '$( $Product.'Product Name' )' is not unique.  Please change the Name and try again."
					continue
				}
				
				# Handle first page, which only asks for name and type.
				$ProductNameField = $BrowserObject | Get-Control -Type Text -ID "ctl00_ctl00_C_M_txtName"
				Set-TextField $ProductNameField $Product.'Product Name'
				# We only deal with non-printed products, so set Type to that.
				$Picklist = $BrowserObject | Get-Control -Type List -ID "ctl00_ctl00_C_M_drpProductTypes"
				$Picklist | Select-FromList -Item "Non Printed Products"
				
				# Click Next to get to product creation page
				Click-Link ( $BrowserObject | Get-Control -Type Button -ID "ctl00_ctl00_C_M_btnNext" )
				
				<# Issue 12
					Function must check to make sure there wasn't a problem.  For example...
						Attempting to provide a product name that already exists results in this same page
						loading again, with a message next to the Product Name field reading "This name already exists"
				#>
				
				# The rest of the work is the same whether adding or updating, so let another function do it.
				$BrowserObject | Update-Product -Product $Product
			}
		}
		
		"Change"	{
			write-log "${Fn}: Change details for product, $($Product.'Product Id')"

			# Locate product; Find-Product will ensure a unique match.
			$EditProduct = Find-Product -BrowserObject $BrowserObject -Product $Product
			if ( $EditProduct ) {
				Click-Link $EditProduct
				# Now we're on the product details page.
				$BrowserObject | Update-Product -Product $Product
			} else {
				Write-Log -fore yellow "${Fn}: Unable to find ID '$( $Product.'Product Id' )'; skipping."
			}
		}
		
		"Skip"	{
			# Ignore this entry; provided in case user wants to leave the row in the file for some reason.
			Write-Log "${Fn}: Skip input row, $($Product.'Product Id')"
		}
		
		default {
			# This should never be reached due to parameter validation!
		}
	}
	

	<#	Here are the input fields we need to deal with.
		
		First Page:
			Product Name, input id="ctl00_ctl00_C_M_txtName" maxlength="50"
			Type, select id="ctl00_ctl00_C_M_drpProductTypes"
				<option value="0">Static Document</option>
				<option value="3">Non Printed Products</option>
				<option value="4">Kit</option>
				<option value="6">Ad Hoc</option>
				<option value="7">DSFdesign Studio</option>
				<option value="11">Product Matrix</option>
				<option value="12">Digital Download</option>
				<option value="20">EFI SmartCanvas Powered by DirectSmile</option>
			Next, input id="ctl00_ctl00_C_M_btnNext"		
	#>
}

function Set-PriceRow {
	<#
		.Synopsis
		Given a WebElement containing a pricing row, and price data, sets the prices.
		
		.Description
		Given a WebElement containing a pricing row, and price data, sets the prices.  This function can do
		both Regular and Setup at once, or individually.  There is no default for either, because if not
		specified we don't want to overwrite a value that's already present.
		
		.Parameter RegularPrice
		Number representing the normal price for the item, such as 2.55.
		
		.Parameter SetupPrice
		Number representing the setup fee for the item.
	#>
	
	param (
		[Parameter( Mandatory, ValueFromPipeLine )]
		[OpenQA.Selenium.Remote.RemoteWebElement] $PriceRow,
		
		[float] $RegularPrice,
		
		[float] $SetupPrice
	)
	
	$Fn = (Get-Variable MyInvocation -Scope 0).Value.MyCommand.Name

	Write-DebugLog "${Fn}: Got price row, $( $PriceRow.GetHashCode() )"
<#	Write-DebugLog "'PriceRow' object type '$($PriceRow.GetType().FullName)'"
	# For debugging, iterate through each <td> and output IDs of any input objects found in them.
	Write-DebugLog "${Fn}: Price row contains these Input fields:"
	foreach ( $td in $PriceRow.FindElementsByTagName("td") ) {
		Write-DebugLog ( ( $td.FindElementsByTagName("input") ).GetProperty("id") | out-string )
	}
	Write-DebugLog ( $PriceRow | out-string )
#>
	
	# Check if something was supplied, and act on it if so.
	
	if ( $PSBoundParameters.ContainsKey('RegularPrice') ) {
		Write-DebugLog "${Fn}: Set Regular Price to `$${RegularPrice}"
		
		# Find the input box for Regular Price.  It will have ID like "*_PriceCatalog_regularprice_*"
		$RegPriceTxt = $PriceRow.FindElementsByTagName("input") | Where-Object { $_.GetProperty("id") -like "*_PriceCatalog_regularprice_*" }
		Write-DebugLog "'RegPriceTxt' object type '$($RegPriceTxt.GetType().FullName)'"
		Set-TextField $RegPriceTxt $RegularPrice
	}
	
	if ( $PSBoundParameters.ContainsKey('SetupPrice') ) {
		Write-DebugLog "${Fn}: Set Setup Price to `$${SetupPrice}"
		
		# Find the input box for Setup Price.  It will have ID like "*_PriceCatalog_setupprice_*"
		$SetPriceTxt = $PriceRow.FindElementsByTagName("input") | Where-Object { $_.GetProperty("id") -like "*_PriceCatalog_setupprice_*" }
		Write-DebugLog "'SetPriceTxt' object type '$($SetPriceTxt.GetType().FullName)'"
		Set-TextField $SetPriceTxt $SetupPrice
	}
}

function Update-Product {
	<#
		.Synopsis
		Given an object containing product details, fill in the product pages as needed.
		
		.Description
		This function takes a custom object populated with all details pertaining to a DSF product.
		It will then step through all the pages necessary to update the details of that product,
		provided that you're already at the first of those pages when it's called.
		
		Non-empty properties will be populated as needed.
		
		Currently, this function only handles "Non-printed products," not printed, kits, etc.
		Other types may be added later.
		
		.Parameter Product
		Custom object containing the details of a product, typically populated via spreadsheet import.
		
		.Parameter BrowserObject
		Selenium driver object representing the browser we're automating.
	#>

	param (
		[Parameter( Mandatory )]
		[PSCustomObject] $Product,
		
		[Parameter( Mandatory, ValueFromPipeLine )]
		[OpenQA.Selenium.Remote.RemoteWebDriver] $BrowserObject
	)

	<#	Lots of conditionals here, but this way we have one function that handles the form filling.
		If a DSF update or process change alters what we have to deal with, we only have to 
		change one section instead of multiple parts of the code.
	
		Whether we are making a new product or changing an existing one, we're going to use the
		Product Information page.  For each piece of data, if it's empty we don't change the web
		form; if it's not empty, we set the appropriate value in the form.
		
		In this way, the same spreadsheet can be used to update products by just running the
		script against it after changing some of the values.
	#>
	
	<#	Product page:
			Navigation box at upper left:
				Information, <span class="rtsTxt">Information</span>
				Details, <span class="rtsTxt">Details</span>
				Settings, <span class="rtsTxt">Settings</span>
				Pricing, <span class="rtsTxt">Settings</span>
				Security, <span class="rtsTxt">Security</span>
	#>

	<#		Information section
			x	Product Name, input id="ctl00_ctl00_C_M_ctl00_W_ctl01__Name"
			x	Display As, input id="ctl00_ctl00_C_M_ctl00_W_ctl01__StorefrontName" maxlength="2000"
			x	Product Id, input id="ctl00_ctl00_C_M_ctl00_W_ctl01__SKU" maxlength="50"
			x	Brief Description, <iframe title="Rich text editor with ID ctl00_ctl00_C_M_ctl00_W_ctl01__Description" id="ctl00_ctl00_C_M_ctl00_W_ctl01__Description_contentIframe" src="javascript:'<html></html>';" frameborder="0" style="margin: 0px; padding: 0px; width: 100%; height: 166px;"></iframe>
					<body style="font-family: Arial;" contenteditable="true"><p>insert text here...</p></body>
			/	Product Icon - Smart Storefront (upload may be required)
					Edit, input id="ctl00_ctl00_C_M_ctl00_W_ctl01__BigIconByItself_EditProductImage"
						Upload Custom Icon, input id="ctl00_ctl00_C_M_ctl00_W_ctl01__BigIconByItself_ProductIcon_rdbUploadIcon" 
							<input name="ctl00$ctl00$C$M$ctl00$W$ctl01$_BigIconByItself$ProductIcon$_uploadedFile$ctl01" type="file">
								Test image: "C:\Users\Carl\Pictures\cute_moogle_by_negocio_plz-d5dw8o1.png"
							Use same image for all icons, input id="ctl00_ctl00_C_M_ctl00_W_ctl01__BigIconByItself_ProductIcon_ChkUseSameImageIcon" type="checkbox"
							Upload, input id="ctl00_ctl00_C_M_ctl00_W_ctl01__BigIconByItself_ProductIcon_Upload"
					If custom icon hasn't been provided, the Smart Storefront image will be:
						<img id="ctl00_ctl00_C_M_ctl00_W_ctl01__BigIconByItself_ProductIconImage" style="height: 50%;" src="/DSF/Images/44d84a1f-8850-4627-bc55-9402d10ae145/Blanks/27.gif">
	#>

	# Product Name, max length 50
	#	This is the name customers see most of the time.
	if ( $null -notlike $Product.'Product Name' ) {
		$Field = Find-SeElement -Driver $BrowserObject -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__Name"
		Set-TextField $Field $Product.'Product Name'.Trim()
	}

	# Display As, max length unknown
	#	Supposedly, product name as customer sees it in the storefront catalog.
	#	In reality, rarely seen except when editing product.
	if ( $null -notlike $Product.'Display Name' ) {
		$Field = Find-SeElement -Driver $BrowserObject -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__StorefrontName"
		Set-TextField $Field $Product.'Display Name'.Trim()
	}
	
	# Product ID (SKU), 50 chars max
	if ( $null -notlike $Product.'Product ID' ) {
		$Field = Find-SeElement -Driver $BrowserObject -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__SKU"
		Set-TextField $Field $Product.'Product Id'.Trim()
	}
	
	<#	Dealing with rich text editors!
		The product Brief Description field is a rich text editor in an iFrame.
		Unlike "textbox" elements, you can't just set the value and move on.
		Each one is an iFrame, which you can get like any other element, however
		after that you have to drill down a bit.
		
		$iFrame.ContentWindow.Document.Body.innerHTML or .innerText
		
		This is fine when you're manipulating the objects directly, but now with Selenium
		we're using browser-independent methods so we can't do that.
	#>
	
	# Brief Description, rich text field
	if ( $null -notlike $Product.'Brief Description' ) {
		$iFrame = Find-SeElement -Driver $BrowserObject -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__Description_contentIframe"
		if ( $Product.'Brief Description'.Trim() -in $DashValues ) {
			# Input is just "-" so clear it by setting value to a space.
			Set-RichTextField -FieldObject $iFrame -XPath "/html/body" -Text ' '
		} else {
			Set-RichTextField -FieldObject $iFrame -XPath "/html/body" -Text $Product.'Brief Description'.Trim()
		}
	}
	
	# Now, if there's a thumbnail image to upload, do that.
	#	Issue 1
	#	As long as the web form accepts a valid file path WITHOUT the user populating it via a dialog,
	#	this should be possible.  Upload-Thumbnail will validate the path before attempting to upload, logging 
	#	an error if it's bad.
	
	# If we have a path to image file, and SkipImageUpload isn't set, try to upload.
	if ( $null -notlike $Product.'Product Icon' ) {
		# Log a message if SkipImageUpload is set.
		if ( $SkipImageUpload ) {
			write-log -fore yellow "Warning: File path provided but SkipImageUpload is set; ignoring for '$($Product.'Product ID')'."
		} else {
			# Upload the image file 
			Upload-Thumbnail -BrowserObject $BrowserObject -ImageURI $Product.'Product Icon'
		}
	}
	
	# Switch to Details section.
	#	<a class="rtsLink rtsAfter" id="TabDetails" href="#">...
	$NavTab = $BrowserObject | Wait-Link -TagName "a" -Property "id" -Pattern "TabDetails"
	$NavTab | Click-Link
	
	<#		Details section
			x	Long Description, div id="ctl00_ctl00_C_M_ctl00_W_ctl01__LongDescription_contentDiv"
					Limit 4000 chars
	#>

	# Long Description
	if ( $null -notlike $Product.'Long Description' ) {
		# This editor isn't in an iFrame.
		$LongDescField = Find-SeElement -Driver $BrowserObject -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__LongDescription_contentDiv"
		if ( $Product.'Long Description'.Trim() -in $DashValues ) {
			Set-RichTextField -RichEditFrame $LongDescField -Text ' '
		} else {
			Set-RichTextField -RichEditFrame $LongDescField -Text $Product.'Long Description'.Trim()
		}
	}
	
	# Switch to Settings section.
	$NavTab = $BrowserObject | Wait-Link -TagName "a" -Property "id" -Pattern "TabSettings"
	$NavTab | Click-Link
	
	# Display priority
	<#
		x	Display Priority, select id="ctl00_ctl00_C_M_ctl00_W_ctl01__Rank_DropDownListRank"
		*** If blank, leave this at default.
				<option value="2147483647">Highest</option>
				<option value="2147483646">Higher</option>
				<option value="100000000">High</option>
				<option value="1">Standard - High</option>
				<option selected="selected" value="0">Standard</option>
				<option value="-1">Standard - Low</option>
				<option value="-100000000">Low</option>
				<option value="-2147483647">Lower</option>
				<option value="-2147483648">Lowest</option>
	#>
	if ( $null -notlike $Product.'Display Priority' ) {	
		# If a value is specified, try to set the selection to a matching value.
		# If match fails, print a warning and set it to Standard.
		$Picklist = $BrowserObject | Get-Control -Type List -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__Rank_DropDownListRank"
		$Set = $Picklist | Select-FromList $Product.'Display Priority'
		# Check result of request; log a message if it defaults to Standard.
		if ( $Set -ne $true ) {
			write-log -fore yellow "Warning: No Display Priority option matched the imported data; setting to Standard for '$($Product.'Product ID')'."
			$null = $Picklist | Select-FromList "Standard"
		}
	}
	
	<#
		Valid Dates:
		*** If not specified, product becomes active immediately and forever.
		x	Active, input id="ctl00_ctl00_C_M_ctl00_W_ctl01__ProductActivationCtrl__YesNo_1" type="radio" checked="checked" value="True"
		x	Start Date, input id="ctl00_ctl00_C_M_ctl00_W_ctl01__ProductActivationCtrl__Begin_dateInput_text" 
				defaults to current date; may be in the future
	#>
	
	<#	Note about radio buttons and checkboxes:
		In this web form, most of the radio buttons have an OnClick property, which causes
		something to happen when a button is selected, such as activating a text field.
		Similarly, checkboxes often control sections of the form.
		In order to ensure the form works properly, we'll use Click() on these controls,
		instead of just setting their values.
		This also reduces the amount of code required, as clicking one radio button
		usually causes the others in that set to become unselected so we don't have to do it.
	#>
	
	# Valid Dates:  Here, we do things a little differently.
	#	If Start Date is empty, product should become active immediately, so we need to 
	#	make sure Active is set.
	if ( $null -notlike $Product.'Start Date' ) {
		<#	To make sure we give the form valid data, we will cast the input from CSV as a 
			PowerShell DateTime object, which can do smart conversion of things like "2017-07-21", 
			"4 Jul 2017" etc.  It will even handle things like "7/22", though it must assume the 
			current year is meant.
			Then, we call ToShortDateString(), which will always output "7/22/2017", because
			that is the format DSF is expecting.
		#>
		$StartDate = [DateTime]$Product.'Start Date'
		$StartDateField = $BrowserObject | Get-Control -Type Text -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__ProductActivationCtrl__Begin_dateInput_text"
		Set-TextField $StartDateField $StartDate.ToShortDateString()
	} else {
		# Start Date is empty, so set product to Active.
		$RadioButton = $BrowserObject | Get-Control -Type RadioButton -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__ProductActivationCtrl__YesNo_1"
		$RadioButton | Set-RadioButton
	}
	
<#
		x	End Date
				To specify a date, click the radio button and then set the text field.
				Radio button: 
					id="ctl00_ctl00_C_M_ctl00_W_ctl01__ProductActivationCtrl_rdbEndDate"
				Text field:
					id="ctl00_ctl00_C_M_ctl00_W_ctl01__ProductActivationCtrl__End_dateInput_text" 
					Date must be in the future.
		
			Never
				Radio button:
					id="ctl00_ctl00_C_M_ctl00_W_ctl01__ProductActivationCtrl_rdbNever"
#>

	# Using similar logic, if End Date is empty, product will be active forever.
	if ( $null -notlike $Product.'End Date' ) {
		# Click the button to select End Date.
		$RadioButton = $BrowserObject | Get-Control -Type RadioButton -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__ProductActivationCtrl_rdbEndDate"
		$RadioButton | Set-RadioButton
		# Now set the date.
		$StopDate = [DateTime]$Product.'End Date'
		$StopDateField = $BrowserObject | Get-Control -Type Text -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__ProductActivationCtrl__End_dateInput_text"
		Set-TextField $StopDateField $StopDate.ToShortDateString()
	} else {
		# End Date is empty, so set to Never.
		$RadioButton = $BrowserObject | Get-Control -Type RadioButton -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__ProductActivationCtrl_rdbNever"
		$RadioButton | Set-RadioButton
	}
	
	<#
		x	Turn Around Time, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_TurnAroundTimeCtrl_rdbNone" type="radio" checked="checked" value="rdbNone"
				either "None" or number of days, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_TurnAroundTimeCtrl__Value"
				number of days field is disabled if "None" is selected
				To supply a value, first click:
					id="ctl00_ctl00_C_M_ctl00_W_ctl01_TurnAroundTimeCtrl_rdbValue"
	#>

	# Turnaround time is the same deal -- combo radio button and text field.
	if ( $null -notlike $Product.'Turnaround Time' ) {
		# Set Value (the second radio button)
		$RadioButton = $BrowserObject | Get-Control -Type RadioButton -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_TurnAroundTimeCtrl_rdbValue"
		$RadioButton | Set-RadioButton
		# Now fill in the number of days.
		$TurnaroundField = $BrowserObject | Get-Control -Type Text -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_TurnAroundTimeCtrl__Value"
		Set-TextField $TurnaroundField $Product.'Turnaround Time'
	} else {
		# None specified; set radio button to None.
		$RadioButton = $BrowserObject | Get-Control -Type RadioButton -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_TurnAroundTimeCtrl_rdbNone"
		$RadioButton | Set-RadioButton
	}
	
	<#
			x	Shipping Charges:
				Exempt Shipping Charges, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_chkShippingExempt"
		x	Tax:
				Exempt Taxes, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_chkTaxExempt"
		x	Mobile Supported
				Yes, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_IsMobileSupportedList_0"
				No, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_IsMobileSupportedList_1"
	#>
	
	# Exempt from Shipping Charge?
	if ( $null -notlike $Product.'Exempt Shipping' ) {
		$ExemptShipChk = $BrowserObject | Get-Control -Type Checkbox -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_chkShippingExempt"
		if ( $Product.'Exempt Shipping' -in $YesValues ) {
			Set-CheckBox $ExemptShipChk
		} else {
			Set-CheckBox $ExemptShipChk -Off
		}
	}
	
	# Exempt from Sales Tax?
	if ( $null -notlike $Product.'Exempt Tax' ) {
		$ExemptTaxChk = $BrowserObject | Get-Control -Type Checkbox -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_chkTaxExempt"
		if ( $Product.'Exempt Tax' -in $YesValues ) {
			Set-CheckBox $ExemptTaxChk
		} else {
			Set-CheckBox $ExemptTaxChk -Off
		}
	}
	
	# Show on the mobile version of the site?
	if ( $null -notlike $Product.'Mobile' ) {
		$MobileSupportYes = $BrowserObject | Get-Control -Type RadioButton -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_IsMobileSupportedList_0"
		$MobileSupportNo = $BrowserObject | Get-Control -Type RadioButton -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_IsMobileSupportedList_1"
		if ( $Product.'Mobile' -in $YesValues ) {
			# Set if Yes or unspecified.
			$MobileSupportYes | Set-RadioButton
		} else {
			# Set No if we specifically don't want mobile support.
			$MobileSupportNo | Set-RadioButton
		}
	}

	<#
		Manage inventory:
			Enabled, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_chkManageInventory"
			if checked, more fields become available...
				
				Replenish inventory - either Add or Reset:
					Add XXX to existing
						Radio button to activate, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_RbInvAddToExistingInv" type="radio" value="RbInvAddToExistingInv"
						Number field, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_RbInvAddToExistingInvTextBox"
					Reset to XXX,
						Radio button to activate, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_RbInvReset" type="radio" value="RbInvReset">
						Number field, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_RbResetInvTextBox"
	#>
	
	# If any values are given for inventory management, turn this on and fill them in.
	# If none of these values are specified, leave the setting alone.
	$ManageInventory = $null
	
	# Get the checkbox control; ID is the same whether checked or not.
	$MgInvenChk = $BrowserObject | Get-Control -Type Checkbox -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_chkManageInventory"
	
	switch ( $Product.'Manage Inventory' ) {
		# If explicitly set to "Yes," turn the checkbox on.
		{ $_ -in $YesValues }	{ Write-DebugLog "ManageInv = True" ; $ManageInventory = $true ; continue }
		
		# If explicitly set to "No," turn the checkbox off.
		{ $_ -in $NoValues }	{ Write-DebugLog "ManageInv = False" ; $ManageInventory = $false ; continue }
		
		# If we got here, Manage Inventory wasn't specified, but if any inventory management values are given,
		#	enable it anyway.  (Sanity check!)
		{ $Product.Threshold -or
			$Product.'Allow Back Order' -or
			$Product.'Show Inventory with Back Order' -or
			$Product.'Add to Inventory' -or
			$Product.'Reset Inventory'
		}						{ 
									# Enable and log a warning.
									$ManageInventory = $true
									write-log -fore yellow "Warning: Manage Inventory not specified, but management items were; enabling for '$($Product.'Product ID')'."
								}
	}
	
	# So if none of the conditions are met, $ManageInventory will still be NULL.  As it's neither True nor False,
	#	following check will drop through and do nothing.
	
	# Now, if ManageInventory is TRUE, handle the values that depend on it.
	if ( $ManageInventory -eq $true ) {
		# Check the box for Manage Inventory = Enabled, otherwise this part of the form will be disabled.
		Set-CheckBox $MgInvenChk
		
		# Threshold, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_txtThQty"
		if ( $null -notlike $Product.Threshold ) {
			$InvThreshold = $BrowserObject | Get-Control -Type Text -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_txtThQty"
			Set-TextField $InvThreshold [int]$Product.Threshold
		}
		
		# Allow Back Order, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_chkBackOrderAllowed"
		if ( $null -notlike $Product.'Allow Back Order' ) {
			$AllowBkOrdChk = $BrowserObject | Get-Control -Type Checkbox -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_chkBackOrderAllowed"
			if ( $Product.'Allow Back Order' -in $YesValues ) {
				Set-CheckBox $AllowBkOrdChk
			} else {
				Set-CheckBox $AllowBkOrdChk -Off
			}
		}
		
		# Show inventory when back order is allowed, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_chkShowInventoryWhenBackOrderIsAllowed"
		if ( $null -notlike $Product.'Show Inventory with Back Order' ) {
			$ShowInvBOChk = $BrowserObject | Get-Control -Type Checkbox -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_chkShowInventoryWhenBackOrderIsAllowed"
			if ( $Product.'Show Inventory with Back Order' -in $YesValues ) {
				Set-CheckBox $ShowInvBOChk
			} else {
				Set-CheckBox $ShowInvBOChk -Off
			}
		}
		
		# Notification Email Id, id="ctl00_ctl00_C_M_ctl00_W_ctl01_txtEmailId"
		$NotifyEmailField = $BrowserObject | Get-Control -Type TextArea -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_txtEmailId"
		if ( $null -notlike $Product.'Notify Emails' ) {
			if ( $Product.'Notify Emails' -in $DashValues ) {
				# Clear field if cell contains only "-"
				Set-TextField $NotifyEmailField ""
			} else {
				Set-TextField $NotifyEmailField $Product.'Notify Emails'.Trim()
			}
		}
		
		<#	Replenish inventory - Note this is either one or the other!
				o Add XXX to existing
				o Reset to XXX
		#>
		
		if ( $null -notlike $Product.'Add to Inventory' ) {
			$RadioButton = $BrowserObject | Get-Control -Type RadioButton -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_RbInvAddToExistingInv"
			$RadioButton | Set-RadioButton
			$AddToInvText = $BrowserObject | Get-Control -Type Text -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_RbInvAddToExistingInvTextBox" 
			Set-TextField $AddToInvText $Product.'Add to Inventory'
		} elseif ( $null -notlike $Product.'Reset Inventory' ) {
			$RadioButton = $BrowserObject | Get-Control -Type RadioButton -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_RbInvReset"
			$RadioButton | Set-RadioButton
			$ResetInvText = $BrowserObject | Get-Control -Type Text -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_RbResetInvTextBox" 
			Set-TextField $ResetInvText $Product.'Reset Inventory'
		}
	} elseif ( $ManageInventory -eq $false ) {
		# Turn the checkbox off.
		Set-CheckBox $MgInvenChk -Off
	}
	
	<#
		Settings section
			Order Quantities:
				Set of 4 radio buttons, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_OrderQuantitiesCtrl__AnyQuantities" 
					Any qty, id="ctl00_ctl00_C_M_ctl00_W_ctl01_OrderQuantitiesCtrl__AnyQuantities"
					Fixed qtys, value="_FixedQuantities"
						Enables text box, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_OrderQuantitiesCtrl__FixedQuantitiesValues_ctl02__Value"
					By multiples, value="_Multiples"
						Enables checkbox & 3 text fields
							Allow buyer to edit quantity, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_OrderQuantitiesCtrl__AllowBuyerToEditQty"
							Minimum, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_OrderQuantitiesCtrl__Minimum" 
							Maximum, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_OrderQuantitiesCtrl__Maximum"
							Multiple, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_OrderQuantitiesCtrl__Multiple"
					Advanced (textbox), input id="ctl00_ctl00_C_M_ctl00_W_ctl01_OrderQuantitiesCtrl__Expression"
						Done (button), input id="ctl00_ctl00_C_M_ctl00_W_ctl01_OrderQuantitiesCtrl_btnDone"
				Multiple Recipients:
					Set of 2 radio buttons, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_OrderQuantitiesCtrl_OrderQtysMultRecipient_EachRecipient"
						Enforce max qty permitted in cart, value="OrderQtysMultRecipient_EachRecipient"
						Total qty must add up to a valid qty, value="OrderQtysMultRecipient_TotalQty"
	#>
	
	<#
		Possibilities for Order Quantities section are one, and only one, of these options:
			Any
			Fixed
			By Multiples
			Advanced
	#>
	
	# Are any of the Order Quantity options specified?
	if (	( $Product.'Any Qty' ) -or
			( $Product.'Fixed Qty' ) -or
			( $Product.'Min Qty' ) -or
			( $Product.'Max Qty' ) -or
			( $Product.'Mult Qty' ) -or
			( $Product.'Advanced Qty' ) ) {
		# Something was specified, so now let's act on it.
		Write-DebugLog "Handling Order Quantity options..."
		
		switch ( $true ) {
			# Any quantity should be set if explicitly set or if none of the other options have values,
			#	however if nothing is specified in the input file we should do nothing.
			( $Product.'Any Qty' -in $YesValues ) {
				Write-DebugLog "${Fn}: Any qty"
				$RadioButton = $BrowserObject | Get-Control -Type RadioButton -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_OrderQuantitiesCtrl__AnyQuantities"
				$RadioButton | Set-RadioButton
				continue
			}

			# Fixed Quantities
			( $Product.'Fixed Qty' ) {
				Write-DebugLog "${Fn}: Fixed qty"
				<#  Fixed Quantity actually creates a set of valid values, which you edit using a GUI.
					It's like the pricing sheet, except each row contains only one value.
					We don't handle this yet, so log a warning and move on.
					You will need to update the allowed quantities by hand.
				#>
				# Set the radio button.
				$RadioButton = $BrowserObject | Get-Control -Type RadioButton -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_OrderQuantitiesCtrl__FixedQuantities"
				$RadioButton | Set-RadioButton
				# Wait for field and buttons to appear.
				$UpdateButton = $BrowserObject | Get-Control -Type Button -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_OrderQuantitiesCtrl__FixedQuantitiesValues_ctl02_LinkButton1"
				# Set the input field to whatever our value is.
				$InvFixedQtyText = $BrowserObject | Get-Control -Type Text -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_OrderQuantitiesCtrl__FixedQuantitiesValues_ctl02__Value"
				Set-TextField $InvFixedQtyText $Product.'Fixed Qty'
				# Now click the Update button, which will post the value to the server.
				$UpdateButton | Click-Link
				# Issue 9:  Add logic to handle this option.
				write-log -fore red "Warning [Issue 9]: Fixed Quantities selected for $($Product.'Product ID'); values must be entered by hand!"
				continue
			}
			
			# By Multiples
			( ( $Product.'Min Qty' ) -or
			( $Product.'Max Qty' ) -or
			( $Product.'Mult Qty' ) ) {
				Write-DebugLog "${Fn}: Min/Max/Mult qty"
				# Click the radio button.
				$RadioButton = $BrowserObject | Get-Control -Type RadioButton -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_OrderQuantitiesCtrl__Multiples"
				$RadioButton | Set-RadioButton
				# Set the minimum quantity from input.
				$MinQtyText = $BrowserObject | Get-Control -Type Text -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_OrderQuantitiesCtrl__Minimum"
				Set-TextField $MinQtyText [int]$Product.'Min Qty'
				# If max wasn't given, use the default.
				$MaxQtyText = $BrowserObject | Get-Control -Type Text -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_OrderQuantitiesCtrl__Maximum"
				if ( $Product.'Max Qty' ) {
					Set-TextField $MaxQtyText [int]$Product.'Max Qty'
					# Check the box for Enforce Max Quantity Permitted in Cart.
					$Checkbox = $BrowserObject | Get-Control -Type Checkbox -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_OrderQuantitiesCtrl_chkEnforceMaxQtyInCart"
					Set-CheckBox $Checkbox
				} else {
					Set-TextField $MaxQtyText $DefaultMaxQty
				}
				# If no multiple was given, use the default.
				$MultQtyText = $BrowserObject | Get-Control -Type Text -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_OrderQuantitiesCtrl__Multiple"
				if ( $Product.'Mult Qty' ) {
					Set-TextField $MultQtyText [int]$Product.'Mult Qty'
				} else {
					Set-TextField $MultQtyText $DefaultQtyMult
				}
				continue
			}
			
			# Advanced is a text field, so just enter whatever was given.
			( $Product.'Advanced Qty' ) {
				Write-DebugLog "${Fn}: Advanced qty"
				$AdvQtyText = $BrowserObject | Get-Control -Type Text -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_OrderQuantitiesCtrl__Expression"
				Set-TextField $AdvQtyText $Product.'Advanced Qty'
				# Click Done button.
				$AdvDoneBtn = Get-Control -Type Button -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_OrderQuantitiesCtrl_btnDone"
				$AdvDoneBtn | Click-Link
				continue
			}
		} # end switch
		
	} else {
		Write-DebugLog "No Order Quantity options given."
	}
	
	<#
		Production Notes, <textarea name="ctl00$ctl00$C$M$ctl00$W$ctl01$_ProductionNotes" id="ctl00_ctl00_C_M_ctl00_W_ctl01__ProductionNotes" style="width: 90%;" rows="10" cols="20"></textarea>
		Keywords, <textarea name="ctl00$ctl00$C$M$ctl00$W$ctl01$_Keywords" id="ctl00_ctl00_C_M_ctl00_W_ctl01__Keywords" style="width: 90%;" rows="10" cols="20"></textarea>
	#>
	
	if ( $null -notlike $Product.'Production Notes' ) {
		$Field = $BrowserObject | Get-Control -Type TextArea -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__ProductionNotes"
		if ( $Product.'Production Notes'.Trim() -in $DashValues ) {
			Set-TextField $Field ' '
		} else {
			Set-TextField $Field $Product.'Production Notes'.Trim()
		}
	}
	
	if ( $null -notlike $Product.Keywords ) {
		$Field = $BrowserObject | Get-Control -Type TextArea -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__Keywords"
		if ( $Product.Keywords.Trim() -in $DashValues ) {
			Set-TextField $Field ' '
		} else {
			Set-TextField $Field $Product.Keywords.Trim()
		}
	}
	
	<#
				Weight:
					Weight, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_WeightCtrl__Weight"
					Picklist, select id="ctl00_ctl00_C_M_ctl00_W_ctl01_WeightCtrl__Unit"
						<option selected="selected" value="0">oz</option>
						<option value="1">lb</option>
				Shipping Subcontainer:
					Ship item separately, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_ShipmentDimensionCtrl_chkPackSeparately"
					Width, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_ShipmentDimensionCtrl__BoxX__Length"
					Width units, select id="ctl00_ctl00_C_M_ctl00_W_ctl01_ShipmentDimensionCtrl__BoxX__Unit"
						<option selected="selected" value="0">Inches</option>
						<option value="1">Feet</option>
						<option value="5">Yard</option>
					Length, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_ShipmentDimensionCtrl__BoxY__Length"
					Length units, select id="ctl00_ctl00_C_M_ctl00_W_ctl01_ShipmentDimensionCtrl__BoxY__Unit"
						<option selected="selected" value="0">Inches</option>
						<option value="1">Feet</option>
						<option value="5">Yard</option>
					Height, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_ShipmentDimensionCtrl__BoxZ__Length"
					Height units, select id="ctl00_ctl00_C_M_ctl00_W_ctl01_ShipmentDimensionCtrl__BoxZ__Unit"
						<option selected="selected" value="0">Inches</option>
						<option value="1">Feet</option>
						<option value="5">Yard</option>
					Max qty per subcontainer, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_ShipmentDimensionCtrl_txtLotSize"
	#>

	# Product weight
	$Field = $BrowserObject | Get-Control -Type Text -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_WeightCtrl__Weight"
	if ( [string]::IsNullOrWhiteSpace( $Product.Weight ) ) {
		# OK, somehow the weight IS null, so if Weight is currently empty set it to zero.
		# If Weight is already filled in, leave it alone.
		if ( $Field.GetAttribute("value") -eq $null ) {
			# A zero weight will allow the product to be created, though it won't work for shipping quotes.
			Set-TextField $Field "0"
		}
	} else {
		Set-TextField $Field $Product.Weight
	}
	# Weight units - Get the list object, then select the right value.
	if ( $Product.'Weight Unit' ) {
		$WeightList = $BrowserObject | Get-Control -Type List -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_WeightCtrl__Unit"
		$WeightList | Select-FromList -Item ( $Product.'Weight Unit' | FixUp-Unit )
	}
	
	# Ship item separately?
	if ( $Product.'Ship Separately' ) {
		$Checkbox = $BrowserObject | Get-Control -Type Checkbox -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_ShipmentDimensionCtrl_chkPackSeparately"
		if ( $Product.'Ship Separately' -in $YesValues ) {
			Set-CheckBox $Checkbox
		} else {
			Set-CheckBox $Checkbox -Off
		}
	}
	
	# Width
	if ( $Product.Width ) {
		$NumField = $BrowserObject | Get-Control -Type Text -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_ShipmentDimensionCtrl__BoxX__Length"
		Set-TextField $NumField $Product.Width
	if ( $Product.'Width Unit' ) {
			$UnitList = $BrowserObject | Get-Control -Type List -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_ShipmentDimensionCtrl__BoxX__Unit"
			$UnitList | Select-FromList -Item ( $Product.'Width Unit' | FixUp-Unit )
		}
	}
	
	# Length
	if ( $Product.Length ) {
		$NumField = $BrowserObject | Get-Control -Type Text -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_ShipmentDimensionCtrl__BoxY__Length"
		Set-TextField $NumField $Product.Length
	if ( $Product.'Length Unit' ) {
			$UnitList = $BrowserObject | Get-Control -Type List -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_ShipmentDimensionCtrl__BoxY__Unit"
			$UnitList | Select-FromList -Item ( $Product.'Length Unit' | FixUp-Unit )
		}
	}
	
	# Height
	if ( $Product.Height ) {
		$NumField = $BrowserObject | Get-Control -Type Text -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_ShipmentDimensionCtrl__BoxZ__Length"
		Set-TextField $NumField $Product.Height
	if ( $Product.'Height Unit' ) {
			$UnitList = $BrowserObject | Get-Control -Type List -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_ShipmentDimensionCtrl__BoxZ__Unit"
			$UnitList | Select-FromList -Item ( $Product.'Height Unit' | FixUp-Unit )
		}
	}
	
	# Max Qty per Subcontainer
	if ( $Product.'Max Qty Per Subcontainer' ) {
		$NumField = $BrowserObject | Get-Control -Type Text -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_ShipmentDimensionCtrl_txtLotSize"
		Set-TextField $NumField $Product.'Max Qty Per Subcontainer'
	}
	
	<#		Pricing section:
				Note:  Tiered pricing will require creation of table rows, which in the GUI is done
					by clicking buttons and then modifying the Range Unit fields.
					For now, we don't handle tiered pricing.
				ADS Base Price Sheet
					Regular Price, input id="tbl_0_PriceCatalog_regularprice_1"
					Setup Price, input id="tbl_0_PriceCatalog_setupprice_1"
	#>
	
	<#
		Functions needed for pricing rows:
			New-PriceRow, to make a new one.
			Get-PriceRow, to find a row based on some reliable criteria.
			Set-PriceRow, to modify an existing one, such as the default row you get with a new product.
			Remove-PriceRow, to delete one.
				Needs to know how to identify the row.
	#>
	
	<#		Security section:
				There are some things here, but we probably don't need to change them, except for
				Owner.  That should be set to a Group, for easier management particularly in
				environments where more than one person administrates, without sharing credentials.
			
		Finish (button), input id="ctl00_ctl00_C_M_ctl00_W_FinishNavigationTemplateContainerID_FinishButton"
	#>
	
	# Switch to Pricing section.
	$NavTab = $BrowserObject | Wait-Link -TagName "a" -Property "id" -Pattern "TabPricing"
	$NavTab | Click-Link
	
	# Issue 10:  Add price handling.
	# Don't cast these as [float] because then an empty value becomes 0.
	if ( ( $Product.'Regular Price' ) -or ( $Product.'Setup Price' ) ) {
		$BasePriceRow = $BrowserObject | Get-PriceRow -PriceSheetName "ADS Base Price Sheet" -RangeStart 1
		# Must check numeric values against null because 0 counts as False.
		# Also, Set-PriceRow accepts [float] so we don't want to send a null to it because that would become 0.
		if ( $Product.'Regular Price' ) {
			$BasePriceRow | Set-PriceRow -RegularPrice $Product.'Regular Price'
		}
		if ( $Product.'Setup Price' ) {
			$BasePriceRow | Set-PriceRow -SetupPrice $Product.'Setup Price'
		}
	}
	
	# Switch to Security section.
	$NavTab = $BrowserObject | Wait-Link -TagName "a" -Property "id" -Pattern "TabSecurity"
	$NavTab | Click-Link
	
	# Issue 11:  Add code to add, verify and change security groups.
	
	# For now, just click the Finish button to save the updated product.
	# Also, hit the Done button on the following page, to get through Category assignment.
	
	$FinishButton = $BrowserObject | Get-Control -Type Button -ID "ctl00_ctl00_C_M_ctl00_W_FinishNavigationTemplateContainerID_FinishButton"
	$FinishButton | Click-Link
	
	# There are many situations that lead to DSF simply refusing to add or update a product.
	# In most of these cases, there is absolutely no indication on the user-facing site as to why.
	# Also, Stage 2 is still on 'ManageProduct.aspx' so checking URL doesn't help.
	# Therefore, after pressing 'Finish' we need to see if the 'Done' button appears.
	#	If it didn't, then assume the operation failed.
	$CategoryDoneBtn = $BrowserObject | Get-Control -Type Button -ID "ctl00_ctl00_C_M_ctl00_W_ctl02__Done"
	# Gotcha:  There exists [System.Management.Automation.Internal.AutomationNull], which will actually
	#	produce $null as the result when using "-like $null", instead of returning $true as we'd expect!
	# For that reason, test using -eq instead, which will work as expected.
	# See:  https://stackoverflow.com/questions/30016949/why-and-how-are-these-two-null-values-different
	# Also, -like is for string comparisons and will therefore cast things as strings if they aren't already.
	# We shouldn't use it for non-string data.
	$CategoryDoneBtn | Dump-ElementInfo -WebInfo
	if ( $CategoryDoneBtn -eq $null ) {
		Write-Log -fore red "${Fn}: Error!  DSF refused to save product info for '$( $Product.'Product Id' )'."
		return
	}
	
	$CategoryDoneBtn | Click-Link
}

function Upload-Thumbnail {

	<# Issue 26:  Make this function not specific to products; needs to accept a control instead of being
		hard-coded to the product form.
	#>
	
	param (
		[Parameter( Mandatory )]
		[OpenQA.Selenium.Remote.RemoteWebDriver] $BrowserObject,

		[ValidateNotNullOrEmpty()]
		[string] $ImageURI
	)
	
	$Fn = (Get-Variable MyInvocation -Scope 0).Value.MyCommand.Name

	# Verify file actually exists.
	if ( test-path $ImageURI ) {
		# Seems legit; proceed with upload process.
		# Start by clicking "Edit" button.
		$EditButton = $BrowserObject | Get-Control -Type Button -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__BigIconByItself_EditProductImage"
		if ( $EditButton ) {
			$EditButton | Click-Link
			# Once clicked, image graphic is replaced with a set of radio buttons.
			# Select "Upload Custom Icon" to proceed.
			$UploadIconButton = $BrowserObject | Get-Control -Type Button -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__BigIconByItself_ProductIcon_rdbUploadIcon"
			$UploadIconButton | Click-Link
			# Now we have a checkbox and a text field to manipulate.
			# Check the box to use this image for all of this product's thumbnails.
			$SameImageForAllChk = $BrowserObject | Get-Control -Type CheckBox -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__BigIconByItself_ProductIcon_ChkUseSameImageIcon"
			Set-CheckBox $SameImageForAllChk
			# Set the text field because we can't mess with a file dialog.
			$ThumbnailField = $BrowserObject | Get-Control -Type File -Name 'ctl00$ctl00$C$M$ctl00$W$ctl01$_BigIconByItself$ProductIcon$_uploadedFile$ctl01'
			Set-TextField $ThumbnailField $ImageURI
			# Click the "Upload" button, which will cause the page to reload.
			$UploadButton = $BrowserObject | Get-Control -Type Button -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__BigIconByItself_ProductIcon_Upload"
			$UploadButton | Click-Link
		} else {
			throw "Error: Couldn't find Edit button for image upload!"
		}
	}
}

	# Save start time for calculating elapsed time later.
	$StartTime = Get-Date

	[string]$ScriptName = Split-Path $MyInvocation.MyCommand.Path -Leaf	# Name of script file
	[string]$ScriptLocation = Split-Path $MyInvocation.MyCommand.Path	# Script is in this folder
	[string]$ScriptDrive = split-path $MyInvocation.MyCommand.Path -qualifier	# Drive (ex. "C:")

	# Suppress Write-Debug's "Confirm" prompt, which occurs because -Debug causes DebugPreference
	#	to change from SilentlyContinue to Inquire.
	If ($PSBoundParameters['Debug']) {
		$DebugPreference = 'Continue'
		$Global:DebugLogging = $true
	}
	
	$LoggingPreference = "Continue"		# Log all output
	# Log file name for Write-Log function
	$LoggingFilePreference = join-path $ScriptLocation "DSF_Task_Log.txt"

<#	######
	######	Global Defaults
	######
#>

	# Main site URL to start from
	$SiteURL = "https://store.adocument.net/DSF/"

	# Snips of text we need to match, to find various controls or input fields
	#	Use '' instead of "" to avoid accidental substitutions.
	#$AdminLinkSnip = 'myadmin-link'
	$AdminLinkText = "Administration"
	$ProductsLinkSnip = "ctl00_ctl00_C_M_LinkColumn3_RepeaterCategories_ctl00_RepeaterItems_ctl02_HyperLinkItem"

	# Some sets of values for flexibility
	$YesValues = "yes","y","true","x"
	$NoValues = "no","n","false"
	# Hyphen, en dash, em dash
	$DashValues = "-",[char]0x2013,[char]0x2014
	# Default Max Quantity if not specified
	$DefaultMaxQty = 10000
	# Default Multiple if not specified
	$DefaultQtyMult = 1
	
	# Selenium script snippets, to run using ExecuteScript:
	
	# Wait for element to become visible if hidden
	# From:  https://stackoverflow.com/questions/44724185/element-myelement-is-not-clickable-at-point-x-y-other-element-would-receiv
	$scrWaitUntilVisible = @"
WebDriverWait wait3 = new WebDriverWait(driver, 10);
wait3.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("ele_to_inv")));
"@
	
<#	######
	######	End Global Defaults
	######
#>
	# Create Firefox instance
	$Browser = Start-SeFirefox
	#$Browser = Start-SeChrome

	$StorefrontURL = Invoke-Login -WebDriver $Browser -SiteURL $SiteURL -UserName $UserName -Password $Password
	# Just a string, the URL of the page that loads after logging in.
	Write-DebugLog "URL loaded: $StorefrontURL"
	
	# Wait a few seconds and check if LoadingSpinner is visible.
	# If "display" attribute is "none" then it's hidden and shouldn't obscure the link.
	$LoadingSpinner = $Browser.FindElementByID("loadingSpinner")
	# Wait for spinner to be hidden.
	$WaitCount = 1
	while ( $LoadingSpinner.GetAttribute("style") -notlike "*display: none*" ) {
		Write-DebugLog "Waiting for Loading Spinner:  $WaitCount"
		Write-DebugLog "Spinner attribute 'style' = $($LoadingSpinner.GetAttribute("style"))"
		$WaitCount++
		Start-Sleep -Seconds 1
	}

	# Verify that we're logged in.  There won't be an Administration link if we aren't.
	#$AdminControl = $Browser.FindElementByCssSelector(".myadmin-link")
	$AdminLink = $Browser.FindElementsByTagName("span") | where { $_.GetAttribute("ng-localize") -eq "StoreFront.Administration" }

	# Admin link exists; now we have to wait until it's not obscured by "Loading" gizmo.
	# By now, the element should no longer be obscured.
	$AdminClickable = WaitFor-ElementToBeClickable -WebElement $AdminLink -TimeInSeconds 30
	if ( $AdminClickable ) {
		write-log -fore green "Admin link found; successfully logged in!"
		$AdminClickable | Click-Link
	} else {
		Dump-ElementInfo $AdminClickable -WebInfo
		throw "Error: Unable to log in; clickable link to Administration page not found."
	}
}

Process {
	<#	This section will be run once for each object passed to the script.
	#>

	try {
		# Find Products link and click it.  Probably unnecessary because Manage-Product calls Find-Product,
		#	which always navigates to Products list to start its job.
		#$ProductsLink = Find-SeElement -Driver $Browser -ID $ProductsLinkSnip
		#Click-Link $ProductsLink
		
		# Grab details from CSV
		
		# *** NOTE ***
		# We should probably bring this in from Excel instead.
		
		# Import-CSV will bring in empty fields as zero-length strings, which are -like but not -eq $null.
		# However, remember to trim leading/trailing whitespace from all text values!
		$ProductList = import-csv $ProductFile
		
		# Some counters for information when done.
		$ProcessCount = 0
		$SkipCount = 0
		
		# Issue 8:  Check $Product.ProcessedStatus (better name?) and if already set, skip this item.
		#	When product has been processed, update this property in the data file.
	
		foreach ( $prItem in $ProductList ) {
			# We use Product ID as the key here, so if it's empty skip this one.
			# Should help in the case of input files with unnoticed blank lines, too.
			if ( [string]::IsNullOrWhiteSpace( $prItem.'Product ID' ) ) {
				if ( ( [string]::IsNullOrWhiteSpace( $prItem.'Product Name' ) ) -and ( [string]::IsNullOrWhiteSpace( $prItem.'Display Name' ) ) ) {
					# If both Name fields are also blank, the whole line probably is, so skip it.
					Write-Log "Skipping probable empty row."
				} else {
					# ID is empty but Name fields aren't; log warning for user.
					Write-Log -fore yellow "Warning: Skipping product with empty ID; please check input file."
				}
				$SkipCount++
			} else {
				Manage-Product -BrowserObject $Browser -Mode $prItem.Operation -Product $prItem
				$ProcessCount++
			}
		}
		
	}
	
	catch {
		Handle-Exception $_
	}
	
}

End {
	# Output count of processed & skipped items.
	write-log -fore green "Done!  Counts for this run:"
	write-log -fore green "Items Processed = $ProcessCount"
	write-log -fore green "Skipped (incl blanks) = $SkipCount"
	
	# Output how much time elapsed during the run.
	$StopTime = Get-Date
	$ElapsedTime = $StopTime - $StartTime
	write-log "Elapsed time:  $( $ElapsedTime.ToString().SubString(0,8) )"
	
	read-host "Press Enter to close browser and quit"
	# Shut down driver and close browser
	Stop-SeDriver $Browser
	# if ('browser still running') get-process $BrowserPID | stop-process
	
	Remove-Module -FullyQualifiedName $ManageDSFModule
}

<#
	Steps to create new product:
		Go to Products page.
		"Create Product"
			Product Name: ("Ace - Sample Book")
			Type: Non Printed Products
			Next
		Information section
			Display As: ("Sample Book")
			Product Id: ("ACE Sample Book") limit 50 chars
			Brief Description: free-form text with formatting
			Etc.
		Settings section
			Tax Exempt
			Weight oz/lb
			Width inch/feet/yard
			Length inch/feet/yard
			Height inch/feet/yard
			Max quantity per subcontainer
		Pricing
			ADS Base Price Sheet
				Regular Price
				Setup price
		"Finish"
		Publish It
			Categories > Ace > Ace Items
			Publish
		Done

#>