<#	This script is horrible alpha quality and will not work!
	It's very much a work in progress as I learn how to automate things,
	in order to maybe, one day, not have to add dozens of products by hand.
	
	Maybe.
#>

#requires -Module Selenium

<# Can we automate product creation in DSF?
	Possibly -- see this page for ideas:
		http://www.westerndevs.com/simple-powershell-automation-browser-based-tasks/
	Use case similar to what I want to do:
		http://cmdrkeene.com/automating-internet-explorer-with-powershell
	
	See also this, which uses Invoke-Webrequest for a different approach:
		https://www.gngrninja.com/script-ninja/2016/7/8/powershell-getting-started-utilizing-the-web
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

[cmdletbinding()]

Param (
	#[ValidateNotNullOrEmpty()]
	[ValidateScript({
		if ( ( $_ -like $null ) -or ( test-path $_ ) ) {
			$true
		} else {
			throw "ProductFile - Supplied path not found: $_!"
		}
	})]
#>
	$ProductFile = "C:\Users\Carl\Documents\ADS Work\Automation\Test_Product_List.csv",
	[ValidateNotNullOrEmpty()]
	# Account name of user to log in.
	[string] $UserName = "DefaultUser",
	[ValidateNotNullOrEmpty()]
	# Password for user account.
	[string] $Password
)

Begin {

	# Put any random stuff BELOW functions!

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

function Click-Link {
	<#
		.Synopsis
		Given a link, wait until browser is not busy, then click it.
		
		.Parameter Link
		The link to click.
	#>
	
	param (
		[Parameter( Mandatory, ValueFromPipeLine )]
		[ValidateNotNullOrEmpty()]
		[OpenQA.Selenium.Remote.RemoteWebElement] $Link
	)
	
	Begin {
	}
	
	Process {
		if ( $Link -notlike $null ) {
			# 
			#$Link.Click()
			Invoke-SeClick $Link
			# Now wait for browser to process the click and load the next page
			$Link.WrappedDriver | Invoke-Wait
		} else {
			write-log -fore yellow "Link is empty?"
			write-log "Link: $($Link.href)"
		}
	}
	
	End {
	}
}

function Click-Wait {
	<#
		.Synopsis
		Given an object, call its Click() method and then sleep.  (Default sleep time = 1 second.)
		
		.Parameter ClickMe
		The target object.
	#>
	
	param (
		[Parameter( Mandatory, ValueFromPipeLine, Position=1 )]
		$ClickMe,
		[Parameter( ValueFromPipeLine, Position=2 )]
		[int] $SleepTime = 1
	)
	
	Begin {
	}
	
	Process {
		if ( $ClickMe -notlike $null ) {
			$ClickMe.Click()
			# Now wait for a bit
			Start-Sleep -Seconds $SleepTime
		} else {
			write-log -fore yellow "ClickMe is empty?"
		}
	}
	
	End {
	}
}

function FixUp-Unit {
	<#
		.Synopsis
		Given a string, make sure it's a valid unit abbreviation, then return the unit.
		
		.Description
		User data from CVS will not be consistent, so this function receives an input such as "pounds"
		and returns a standardized unit like "lb".  This allows making comparisons later with less stress.
		
		.Parameter Input
		The unit name to be standardized.
	#>
	
	# 
	param (
		[Parameter( Mandatory, ValueFromPipeLine )]
		$Input
	)
	
	$InchValues = "in", "inch", "inches"
	$FootValues = "ft", "foot", "feet"
	$YardValues = "yd", "yard", "yards", "yds"
	$PoundValues = "lb", "pound", "pounds", "lbs"
	$OunceValues = "oz", "ounce", "ounces", "ozs"
	
	switch ( $Input ) {
		# DSF is not consistent with abbreviations or plurals of units!
		{ $_ -in $InchValues }	{ $Output = "inches"}
		{ $_ -in $FootValues }	{ $Output = "feet"}
		{ $_ -in $YardValues }	{ $Output = "yard"}
		{ $_ -in $PoundValues }	{ $Output = "lb"}
		{ $_ -in $OunceValues }	{ $Output = "oz"}
		default					{ $Output = "undefined"}
	}
	
	$Output
}

function Get-Link {
	<#
		.Synopsis
		Given a document, find a link based on a property, such as "href".
		
		.Parameter Links
		Collection of links to search.
		
		.Parameter Href
		String containing the pattern to match in "href" property.
	#>
	
	param (
		$Links,
		[string] $Href
	)
	
	Begin {
		# Initialize $Collection to an empty array
		$Collection = @()
	}
	
	Process {
		foreach ( $l in $Links ) {
			if ( $l.href -like $href ) {
				#write-host $l.href
				$Collection += $l
			}
		}
	}
	
	End {
		$Collection
	}
}

function Invoke-Wait {

	<#
		.Synopsis
		Wait for browser object to return "complete" and then return.
		
		.Parameter BrowserObject
		A Selenium browser (driver) object
	#>
	
	[cmdletbinding()]
	
	Param(
		[Parameter( Mandatory, ValueFromPipeLine )]
		[OpenQA.Selenium.Remote.RemoteWebDriver] $BrowserObject
	)
	
	<#	Wait until the main document returns "complete"
	#>
	$DocState = $BrowserObject.ExecuteScript("return document.readyState")
	while ( $DocState -notlike "complete" ) {
		# Not ready yet, so wait 1 second and check again.
		write-debug "Waiting for page load to complete..."
		Start-Sleep -Seconds 1
		$DocState = $BrowserObject.ExecuteScript("return document.readyState")
	}
}

function Manage-Product {
	<#	
		.Synopsis
		Add or modify a product in DSF.
		
		.Description
		Add or change a product in DSF.  Once added, pass to Update-Product which will handle all the details.
		
		.Parameter Product
		Object containing the name and other properties of the target product.
		
		.Parameter Document
		Object containing the web page (Document) we're using.
		
		.Parameter Mode
		Add (create new product) or Change (modify existing product).
		May include "Delete" in future.
	#>

	param (
		[Parameter( Mandatory )]
		$Product,
		
		[Parameter( Mandatory )]
		$Document,
		
		[Parameter( Mandatory )]
		[ValidateSet("Add","Change")]
		$Mode
	)
	
	# Need code that will handle Add mode or Change mode.
	# For each property, if Change, check if a value was supplied; if so, change it, otherwise leave it.
	# This is probably best handled by a function so we don't need an "if" on every line.
	# We want to avoid having multiple functions that each have to be rewritten for a DSF change.
	
	switch ( $Mode ) {
		"Add"	{
			write-log "Add product: $($Product.'Product Name')"
			# Press Create Product button, go about new product stuff.
			Click-Link ( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*ButtonCreateProduct" )
			
			# Handle first page, which only asks for name and type.
			( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*txtName" ).Value = $Product.'Product Name'
			# We only deal with non-printed products, so set Type to 3
			$Picklist = $Document | Wait-Link -TagName "select" -Property "id" -Pattern "*drpProductTypes"
			( $Picklist | where innerHTML -eq "Non Printed Products" ).Selected = $true
			
			# Click Next to get to product creation page
			Click-Link ( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*btnNext" )
			
			# The rest of the work is the same whether adding or updating, so let another function do it.
			$Document | Update-Product -Product $Product
		}
		
		"Change"	{
			# Go to list of All Products and click the link to it; make changes.
		}
		
		"Other"	{
			# Some other mode to be added later.
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

function Select-FromList {
	# Given an element that is a pick-list, select the named item from it.
	
	param (
		[Parameter( Mandatory, ValueFromPipeLine )]
		$List,
		
		[Parameter( Mandatory )]
		[string] $Item
	)
	
	$Target = $List | where innerText -eq $Item 
	if ( $Target -notlike $null ) { 
		$Target.Selected = $true 
	} else {
		write-log -fore red "Error: Option `'$Item`' not found in list!"
	}
}

function Update-Product {

	param (
		[Parameter( Mandatory )]
		$Product,
		
		[Parameter( Mandatory, ValueFromPipeLine )]
		$Document
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

	# Product Name will already be filled in, based on the previous page.
	
	# Display As, max length unknown
	#	Supposedly, product name as customer sees it in the storefront catalog.
	#	In reality, rarely seen except when editing product.
	if ( $Product.'Display Name' -notlike $null ) {
		( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*StorefrontName" ).Value = $Product.'Display Name'
	}
	
	# Product ID (SKU), 50 chars max
	if ( $Product.'Product ID' -notlike $null ) {
		( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*SKU" ).Value = $Product.'Product Id'
	}
	
	<#	Dealing with rich text editors!
		The product Brief description field is a rich text editor in an iFrame.
		Unlike "textbox" elements, you can't just set the value and move on.
		Each one is an iFrame, which you can get like any other element, however
		after that you have to drill down a bit.
		
		$iFrame.ContentWindow.Document.Body.innerHTML or .innerText
		
		We will use innerHTML, as the Word form allows formatting.
	#>
	
	# Brief Description, rich text field
	if ( $Product.'Brief Description' -notlike $null ) {
		$iFrame = $Document | Wait-Link -TagName "iframe" -Property "id" -Pattern "*Description_contentIframe"
		$iFrame.ContentWindow.Document.Body.innerHTML = $Product.'Brief Description'
	}
	
	# Now, if there's a thumbnail image to upload, do that.
	#	Issue 1
	#	As long as the web form accepts a valid file path WITHOUT the user populating it via a dialog,
	#	this should be possible.  We'll need to validate the path before attempting to upload, logging 
	#	an error if it's bad.
	if ( $Product.'Product Icon' -notlike $null ) {
		#Upload-Thumbnail -Document $Document -URL $Product.'Product Icon'
	}
	
	# Switch to Details section.
	$NavTab = $Document | Wait-Link -TagName "a" -Property "id" -Pattern "*TabDetails"
	$NavTab.Click()
	
	<#		Details section
			x	Long Description, div id="ctl00_ctl00_C_M_ctl00_W_ctl01__LongDescription_contentDiv"
					Limit 4000 chars
	#>

	# Long Description
	if ( $Product.'Long Description' -notlike $null ) {
		$RichTextEdit = $Document | Wait-Link -TagName "div" -Property "id" -Pattern "*LongDescription_contentDiv"
		$RichTextEdit.innerHTML = $Product.'Long Description'
	}
	
	# Switch to Settings section.
	$NavTab = $Document | Wait-Link -TagName "a" -Property "id" -Pattern "*TabSettings"
	$NavTab.Click()
	
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
	if ( $Product.'Display Priority' -notlike $null ) {	
		# If a value is specified, try to set the selection to a matching value.
		# If match fails, print a warning and set it to Standard.
		$Picklist = $Document | Wait-Link -TagName "select" -Property "id" -Pattern "*DropDownListRank"
		if ( $Picklist.innerHTML -eq $Product.'Display Priority' ) {
			( $Picklist | where innerHTML -eq $Product.'Display Priority' ).Selected = $true 
		} else {
			write-log -fore yellow "Warning: No Display Priority option found to match `'$($Product.'Display Priority')`'; setting to Standard."
			( $Picklist | where innerHTML -eq "Standard" ).Selected = $true 
		}
	}
	
	<#
		Valid Dates:
		*** If not specified, product becomes active immediately and forever.
		x	Active, input id="ctl00_ctl00_C_M_ctl00_W_ctl01__ProductActivationCtrl__YesNo_1" type="radio" checked="checked" value="True"
		x	Start Date, input id="ctl00_ctl00_C_M_ctl00_W_ctl01__ProductActivationCtrl__Begin_dateInput_text" 
				defaults to current date; may be in the future
		x	End Date, input id="ctl00_ctl00_C_M_ctl00_W_ctl01__ProductActivationCtrl__End_dateInput_text" 
				Radio button: input id="ctl00_ctl00_C_M_ctl00_W_ctl01__ProductActivationCtrl_rdbNever" type="radio" checked="checked" value="rdbNever"
				defaults to "Never", may be in the future
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
	if ( $Product.'Start Date' -notlike $null ) {
		<#	To make sure we give the form valid data, we will cast the input from CSV as a 
			PowerShell DateTime object, which can do smart conversion of things like "2017-07-21", 
			"4 Jul 2017" etc.  It will even handle things like "7/22", though it must assume the 
			current year is meant.
			Then, we call ToShortDateString(), which will always output "7/22/2017", because
			that is the format DSF is expecting.
		#>
		$StartDate = [DateTime]$Product.'Start Date'
		( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*Begin_dateInput_text" ).Value = $StartDate.ToShortDateString()
	} else {
		# Start Date is empty, so set product to Active.
		$RadioButton = ( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*ProductActivationCtrl__YesNo_1" )
		$RadioButton.Click()
	}
	
	# Using similar logic, if End Date is empty, product will be active forever.
	if ( $Product.'End Date' -notlike $null ) {
		# Click the button to select End Date.
		$RadioButton = ( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*ProductActivationCtrl_rdbEndDate" )
		$RadioButton.Click()
		# Now set the date.
		$StopDate = [DateTime]$Product.'End Date'
		( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*End_dateInput_text" ).Value = $StopDate.ToShortDateString()
	} else {
		# End Date is empty, so set to Never.
		$RadioButton = ( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*ProductActivationCtrl_rdbNever" )
		$RadioButton.Click()
	}
	
	<#
		x	Turn Around Time, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_TurnAroundTimeCtrl_rdbNone" type="radio" checked="checked" value="rdbNone"
				either "None" or number of days, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_TurnAroundTimeCtrl__Value"
				number of days field is disabled if "None" is selected
		x	Shipping Charges:
				Exempt Shipping Charges, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_chkShippingExempt"
		x	Tax:
				Exempt Taxes, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_chkTaxExempt"
		x	Mobile Supported
				Yes, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_IsMobileSupportedList_0"
	#>

	# Turnaround time is the same deal -- combo radio button and text field.
	if ( $Product.'Turnaround Time' -notlike $null ) {
		# Set Value (the second radio button)
		$RadioButton = ( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*TurnAroundTimeCtrl_rdbValue" )
		$RadioButton.Click()
		# Now fill in the number of days.
		( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*TurnAroundTimeCtrl__Value" ).Value = $Product.'Turnaround Time'
	} else {
		# None specified; set radio button to None.
		$RadioButton = ( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*TurnAroundTimeCtrl_rdbNone" )
		$RadioButton.Click()
	}
	
	# Exempt from Shipping Charge?
	if ( $Product.'Exempt Shipping' -in $YesValues ) {
		( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*chkShippingExempt" ).Checked = $true
	} else {
		( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*chkShippingExempt" ).Checked = $false
	}
	
	# Exempt from Sales Tax?
	if ( $Product.'Exempt Tax' -in $YesValues ) {
		( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*chkTaxExempt" ).Checked = $true
	} else {
		( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*chkTaxExempt" ).Checked = $false
	}
	
	# Show on the mobile version of the site?
	if ( ( $Product.'Mobile' -in $YesValues ) -or ( $Product.'Mobile' -like $null ) ) {
		# Set if Yes or unspecified.
		( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*IsMobileSupportedList_0" ).Checked = $true
	} else {
		( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*IsMobileSupportedList_0" ).Checked = $false
	}

	<#
		Manage inventory:
			Enabled, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_chkManageInventory"
			if checked, more fields become available...
				Threshold, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_txtThQty"
				Allow back Order, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_chkBackOrderAllowed"
				Show inventory when back order is allowed, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_chkShowInventoryWhenBackOrderIsAllowed"
				Notification Email Id, <textarea name="ctl00$ctl00$C$M$ctl00$W$ctl01$txtEmailId" id="ctl00_ctl00_C_M_ctl00_W_ctl01_txtEmailId" rows="3" cols="45"></textarea>
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
	$Checkbox = $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*chkManageInventory"
	
	switch ( $Product.'Manage Inventory' ) {
		# If explicitly set to "Yes," turn the checkbox on.
		{ $_ -in $YesValues }	{ write-debug "ManageInv = True" ; $ManageInventory = $true }
		# If explicitly set to "No," turn the checkbox off.
		{ $_ -in $NoValues }	{ write-debug "ManageInv = False" ; $ManageInventory = $false }
		# Check the box if any inventory management values are given,
		#	even if Manage isn't specified.  (Sanity check!)
		{ $Product.Threshold -or
			$Product.'Allow Back Order' -or
			$Product.'Show Inventory with Back Order' -or
			$Product.'Add to Inventory' -or
			$Product.'Reset Inventory'
		}						{ $ManageInventory = $true }
	}
	
	# So if none of the conditions are met, $ManageInventory will still be NULL.
	
	# Now, if ManageInventory is TRUE, handle the values that depend on it.
	# In order to do this, we may need to take steps to activate the section of the form
	#	that is deactivated if this checkbox is not checked.
	# For now, see what happens if we submit the form anyway.
	if ( $ManageInventory -eq $true ) {
		$Checkbox = ( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*chkManageInventory" )
		#$Checkbox.SetActive()
		# Set "checked" state to False, because Click will toggle it.
		$Checkbox.Checked = $false
		Click-Wait $Checkbox
		
		# Threshold
		if ( $Product.Threshold -notlike $null ) {
			( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*txtThQty" ).Value = $Product.Threshold
		}
		
		# Allow back Order
		$Checkbox = $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*chkBackOrderAllowed"
		$Checkbox.Checked = ( $Product.'Allow Back Order' -in $YesValues )
		
		# Show inventory when back ordered
		$Checkbox = $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*chkShowInventoryWhenBackOrderIsAllowed"
		$Checkbox.Checked = ( $Product.'Show Inventory with Back Order' -in $YesValues )
		
		# Notification Email Id
		# If you want to CLEAR this field, input a blank space instead of leaving it empty.
		if ( $Product.'Notify Emails' -notlike $null ) {
			( $Document | Wait-Link -TagName "textarea" -Property "id" -Pattern "*txtEmailId" ).Value = $Product.'Notify Emails'
		}
		
		# Replenish inventory - Note this is either one or the other!
		#	o Add XXX to existing
		#	o Reset to XXX
		
		# For some reason, calling SetActive or Click on these radio buttons causes the web form
		#	to freeze -- at least from the GUI.  So, try proceeding without doing that.
		
		if ( $Product.'Add to Inventory' -notlike $null ) {
			$RadioButton = ( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*RbInvAddToExistingInv" )
			$RadioButton.isDisabled = $false
			#$RadioButton.SetActive()
			$RadioButton | Click-Wait
			( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*RbInvAddToExistingInvTextBox" ).Value = $Product.'Add to Inventory'
		} elseif ( $Product.'Reset Inventory' -notlike $null ) {
			$RadioButton = ( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*RbInvReset" )
			$RadioButton.isDisabled = $false
			#$RadioButton.SetActive()
			$RadioButton | Click-Wait
			( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*RbResetInvTextBox" ).Value = $Product.'Reset Inventory'
		}
	} elseif ( $ManageInventory -eq $false ) {
		# Turn the checkbox off.
		$Checkbox = ( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*chkManageInventory" )
		#$Checkbox.SetActive()
		$Checkbox.Checked = $true
		$Checkbox.Click()
	}
	
	<#
		Settings section
			Order Quantities:
				Set of 4 radio buttons, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_OrderQuantitiesCtrl__AnyQuantities" 
					Any qty, value="_AnyQuantities"
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
	
	# Here, "Any Quantity" is the default.  If neither Fixed, Min, Max, Mult, or Advanced has a value,
	#	just do nothing and leave it.  If one of those has a value, we need to handle it.
	if ( $Product.'Fixed Qty' -or
			$Product.'Min Qty' -or
			$Product.'Max Qty' -or
			$Product.'Mult Qty' -or
			$Product.'Advanced Qty' ) {
		# One of these is not empty, so act accordingly.
		write-log -fore red "TODO: Fixed/Mult/Advanced section is incomplete."
		
		if ( $Product.'Fixed Qty' -notlike $null ) {
			<#  Fixed Quantity actually creates a set of valid values, which you edit using a GUI.
				It's like the pricing sheet, except each row contains only one value.
				We don't handle this yet, so log a warning and move on.
				Admin will need to update the allowed quantities by hand.
			#>
			$RadioButton = ( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*FixedQuantities" )
			#$RadioButton.SetActive()
			$RadioButton | Click-Wait
			( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*FixedQuantitiesValues_ctl02__Value" ).Value = $Product.'Fixed Qty'
			<# Now click the Update button, which will post the value to the server.
			<input name="ctl00$ctl00$C$M$ctl00$W$ctl01$OrderQuantitiesCtrl$_FixedQuantitiesValues$ctl02$LinkButton1" class="button-mouseout" id="ctl00_ctl00_C_M_ctl00_W_ctl01_OrderQuantitiesCtrl__FixedQuantitiesValues_ctl02_LinkButton1" onmouseover="this.className='button-mouseover'" onmouseout="this.className='button-mouseout'" onclick='javascript:WebForm_DoPostBackWithOptions(new WebForm_PostBackOptions("ctl00$ctl00$C$M$ctl00$W$ctl01$OrderQuantitiesCtrl$_FixedQuantitiesValues$ctl02$LinkButton1", "", true, "_FixedQuantitiesValues", "", false, false))' type="submit" value="Update">
			#>
			# TODO: Handle quantities...
			write-log -fore red "Warning: Fixed Quantities selected for $($Product.'Product Name'); values must be entered by hand!"
			# This option doesn't coexist with others, so break out after doing this.
			break
		}
		
		if ( ( $Product.'Min Qty' -or $Product.'Max Qty' -or $Product.'Mult Qty' ) -notlike $null ) {
			write-log "$($Product.'Product Name') has Min/Max/Mult Qty."
			# Min quantity can be used by itself or in conjunction with Max.
			# For this to be available, "By Multiples" button must be clicked.
			$RadioButton = ( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*_Multiples" )
			$RadioButton | Click-Wait
			( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*OrderQuantitiesCtrl__Minimum" ).Value = $Product.'Min Qty'
			( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*OrderQuantitiesCtrl__Maximum" ).Value = $Product.'Max Qty'
			( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*OrderQuantitiesCtrl__Multiple" ).Value = $Product.'Mult Qty'
		}
		
		# If a Max quantity was specified, check the box to enforce this in shopping cart.
		if ( $Product.'Max Qty' -notlike $null ) {
			$Checkbox = ( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*chkEnforceMaxQtyInCart" )
			# Set to False and then click, to ensure control is recognized.
			$Checkbox.Checked = $false
			$Checkbox | Click-Wait
		}		
	}	
	
	<#
		Production Notes, <textarea name="ctl00$ctl00$C$M$ctl00$W$ctl01$_ProductionNotes" id="ctl00_ctl00_C_M_ctl00_W_ctl01__ProductionNotes" style="width: 90%;" rows="10" cols="20"></textarea>
		Keywords, <textarea name="ctl00$ctl00$C$M$ctl00$W$ctl01$_Keywords" id="ctl00_ctl00_C_M_ctl00_W_ctl01__Keywords" style="width: 90%;" rows="10" cols="20"></textarea>
	#>
	
	if ( $Product.'Production Notes' -notlike $null ) {
		( $Document | Wait-Link -TagName "textarea" -Property "id" -Pattern "*ProductionNotes" ).Value = $Product.'Production Notes'
	}
	
	if ( $Product.Keywords -notlike $null ) {
		( $Document | Wait-Link -TagName "textarea" -Property "id" -Pattern "*Keywords" ).Value = $Product.Keywords
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
	if ( $Product.Weight -notlike $null ) {
		( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*WeightCtrl__Weight" ).Value = $Product.Weight
	} else {
		# OK, somehow the weight IS null, so ensure it's set to zero.
		# A zero weight will allow the product to be created, though it won't work for shipping quotes.
		( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*WeightCtrl__Weight" ).Value = "0"
	}
	# Weight units - Get the list object, then select the right value.
	$WeightList = $Document | Wait-Link -TagName "select" -Property "id" -Pattern "*WeightCtrl__Unit"
	$WeightList | Select-FromList -Item ( $Product.'Weight Unit' | FixUp-Unit )
	
	# Ship item separately?
	$Checkbox = $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*chkPackSeparately"
	$Checkbox.Checked = ( $Product.'Ship Separately' -in $YesValues )
	
	# Width
	$Width = $Product.Width | ForEach-Object { if ( $_ -notlike $null ) { $_ } else { 0 } }
	( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*BoxX__Length" ).Value = $Product.Width
	$UnitList = $Document | Wait-Link -TagName "select" -Property "id" -Pattern "*ShipmentDimensionCtrl__BoxX__Unit"
	$UnitList | Select-FromList -Item ( $Product.'Width Unit' | FixUp-Unit )
	
	# Length
	$Length = $Product.Length | ForEach-Object { if ( $_ -notlike $null ) { $_ } else { 0 } }
	( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*BoxY__Length" ).Value = $Product.Length
	$UnitList = $Document | Wait-Link -TagName "select" -Property "id" -Pattern "*BoxY__Unit"
	$UnitList | Select-FromList -Item ( $Product.'Length Unit' | FixUp-Unit )
	
	# Height
	$Height = $Product.Height | ForEach-Object { if ( $_ -notlike $null ) { $_ } else { 0 } }
	( $Document | Wait-Link -TagName "input" -Property "id" -Pattern "*BoxZ__Length" ).Value = $Product.Height
	$UnitList = $Document | Wait-Link -TagName "select" -Property "id" -Pattern "*BoxZ__Unit"
	$UnitList | Select-FromList -Item ( $Product.'Height Unit' | FixUp-Unit )
			
	<#		Pricing section:
				Note:  Tiered pricing will require creation of table rows, which in the GUI is done
					by clicking buttons and then modifying the Range Unit fields.
					For now, we don't handle tiered pricing.
				ADS Base Price Sheet
					Regular Price, input id="tbl_0_PriceCatalog_regularprice_1"
					Setup Price, input id="tbl_0_PriceCatalog_setupprice_1"
			Security section:
				There are some things here, but we probably don't need to change them.
			
		Finish (button), input id="ctl00_ctl00_C_M_ctl00_W_FinishNavigationTemplateContainerID_FinishButton"
	#>
	
}

function Wait-Link {
	<#	Loop until the specified object becomes available.
		So, where you might do this...
			$UserField = $CurrentDoc.IHTMLDocument3_getElementsByTagName('input') | Where-Object {$_.name -like $UserFieldSnip }
		...instead you do this:
			$UserField = $CurrentDoc | Wait-Link -TagName "input" -Property "name" -Pattern $UserFieldSnip
			
		TODO: Return something unique if wait times out, so the calling function has a chance to handle
			the lack of link, such as forcing a page reload or navigating to some known location.
	#>
	
	<#
		.Synopsis
		Find an element in an IE Document by waiting until it becomes available.
		
		.Parameter Document
		IE Document to search.  May be passed via pipeline.
		
		.Parameter TagName
		Tag name, such as "input" or "span"; will be passed to IHTMLDocument3_getElementsByTagName.
		
		.Parameter Property
		The property of the found element to search.
		
		.Parameter Pattern
		Wildcard pattern to match when searching, such as "*my_UserName".
		
		.Parameter Timeout
		Number of seconds after which to give up waiting.  Default is 30.
	#>
	
	Param(
		[Parameter( Mandatory, ValueFromPipeLine )]
		$Document,
		
		[Parameter( Mandatory )]
		[ValidateNotNullOrEmpty()]
		$TagName,
		
		[Parameter( Mandatory )]
		[ValidateNotNullOrEmpty()]
		$Property,
		
		[Parameter( Mandatory )]
		[ValidateNotNullOrEmpty()]
		$Pattern,
		
		$Timeout = 30
		
	)

	Write-Debug "Wait for $TagName element with $Property matching `'$Pattern`'"
	
	# Create a Stopwatch object to keep track of time
	$Stopwatch = New-Object System.Diagnostics.Stopwatch
	$Stopwatch.Start()
	
	$TimedOut = $false
	
	do {
		# Check if too much time has elapsed; break out if so.
		if ( $Stopwatch.Elapsed.Seconds -ge $Timeout ) {
			$TimedOut = $true
			break
		}
		
		# Test the result of the requested search.
		# If the document hasn't loaded yet, or otherwise isn't populated, this should return nothing.
		# Therefore, only when document is complete will we get our result.
		$result = $Document.IHTMLDocument3_getElementsByTagName( $TagName ) | Where-Object { $_.$Property -like $Pattern }
	}
	until ( $result -notlike $null )
	
	if ( $TimedOut ) {
		write-log -fore yellow "Timeout reached. Add better error handling to Wait-Link!"
		throw "Timed out while waiting for $TagName element with $Property matching `'$Pattern`'"
	}

	# We made it this far, so presumably we got what we need.  Return it.
	$result

}

function Wait-LinkSe {
	<#	Loop until the specified object becomes available.
		So, where you might do this...
			$UserField = $CurrentDoc.IHTMLDocument3_getElementsByTagName('input') | Where-Object {$_.name -like $UserFieldSnip }
		...instead you do this:
			$UserField = $CurrentDoc | Wait-Link -TagName "input" -Property "name" -Pattern $UserFieldSnip
			
		TODO: Return something unique if wait times out, so the calling function has a chance to handle
			the lack of link, such as forcing a page reload or navigating to some known location.
	#>
	
	<#
		.Synopsis
		Find an element in a Selenium Browser Object by waiting until it becomes available.
		
		.Parameter SeObject
		Browser Object to search.  May be passed via pipeline.
		
		.Parameter TagName
		Tag name, such as "input" or "span"; will be passed to Find-SeElement.
		
		.Parameter Property
		The property of the found element to search.
		
		.Parameter Pattern
		Wildcard pattern to match when searching, such as "*my_UserName".
		
		.Parameter Timeout
		Number of seconds after which to give up waiting.  Default is 30.
	#>
	
	Param(
		[Parameter( Mandatory, ValueFromPipeLine )]
		$SeObject,
		
		[Parameter( Mandatory )]
		[ValidateNotNullOrEmpty()]
		$TagName,
		
		[Parameter( Mandatory )]
		[ValidateNotNullOrEmpty()]
		$Property,
		
		[Parameter( Mandatory )]
		[ValidateNotNullOrEmpty()]
		$Pattern,
		
		$Timeout = 30
		
	)

	Write-Debug "Wait for $TagName element with $Property matching `'$Pattern`'"
	
	# Create a Stopwatch object to keep track of time
	$Stopwatch = New-Object System.Diagnostics.Stopwatch
	$Stopwatch.Start()
	
	$TimedOut = $false
	
	do {
		# Check if too much time has elapsed; break out if so.
		if ( $Stopwatch.Elapsed.Seconds -ge $Timeout ) {
			$TimedOut = $true
			break
		}
		
		# Test the result of the requested search.
		# If the document hasn't loaded yet, or otherwise isn't populated, this should return nothing.
		# Therefore, only when document is complete will we get our result.
		$result = $SeObject.FindElementsByTagName( $TagName ) | Where-Object { $_.GetProperty($Property) -like $Pattern }
	}
	until ( $result -notlike $null )
	
	if ( $TimedOut ) {
		write-log -fore yellow "Timeout reached. Add better error handling to Wait-Link!"
		throw "Timed out while waiting for $TagName element with $Property matching `'$Pattern`'"
	}

	# We made it this far, so presumably we got what we need.  Return it.
	$result

}

function Write-DebugLog {
	param ( [string] $Text = " " )
	
	if ( $Debug ) { write-log -fore darkyellow $Text }
}

Function Write-Log {
# Write-Log based on code from http://jdhitsolutions.com/blog/2011/03/powershell-automatic-logging/
<#
	.Synopsis
	Write a message to a log file. 
	
	.Description
	Write-Log can be used to write text messages to a log file. It can be used like Write-Verbose,
	and looks for two variables that you can define in your scripts and functions. If the function
	finds $LoggingPreference with a value of "Continue", the message text will be written to the file.
	The default file is PowerShellLog.txt in your %TEMP% directory. You can specify a different file
	path by parameter or set the $LoggingFilePreference variable. See the help examples.

	.Parameter Message
	The message string to write to the log file. It will be prepended with a date time stamp.
	
	.Parameter Path
	The filename and path for the log file. The default is $env:temp\PowerShellLog.txt, 
	unless the $loggingFilePreference variable is found. If so, then this value will be
	used.

	.Notes
	NAME: Write-Log
	AUTHOR: Jeffery Hicks
	VERSION: 1.0
	LASTEDIT: 03/02/2011

	Learn more with a copy of Windows PowerShell 2.0: TFM (SAPIEN Press 2010)

	.Link
	http://jdhitsolutions.com/blog/2011/03/powershell-automatic-logging/

	.Link
	Write-Verbose
	
	.Inputs
	None

	.Outputs
	None
#>

	#[cmdletbinding()]

	Param(
		[ConsoleColor]$ForegroundColor = "white",
		[Parameter(Mandatory = $true, Position = 0)]
		[AllowEmptyString()]
		[string]$Message = "",
		[string]$Path
	)

	#Pass on the message to Write-Verbose if -Verbose was detected
	Write-Host -fore $ForegroundColor -object $Message

	#only write to the log file if the $LoggingPreference variable is set to Continue
	if ($LoggingPreference -eq "Continue")
	{

		#if a $loggingFilePreference variable is found in the scope
		#hierarchy then use that value for the file, otherwise use the default
		#$path
		if ($loggingFilePreference)
		{
			$LogFile=$loggingFilePreference
		}
		else
		{
			$LogFile=$Path
		}
		
		# Don't bother to log an empty message
		if ( $Message ) {
			Write-Output "$(Get-Date) $Message" | Out-File -FilePath $LogFile -Append
		}
	}

} #end function


	[string]$ScriptName = Split-Path $MyInvocation.MyCommand.Path -Leaf	# Name of script file
	[string]$ScriptLocation = Split-Path $MyInvocation.MyCommand.Path	# Script is in this folder
	[string]$ScriptDrive = split-path $MyInvocation.MyCommand.Path -qualifier	# Drive (ex. "C:")

	# Suppress Write-Debug's "Confirm" prompt, which occurs because -Debug causes DebugPreference
	#	to change from SilentlyContinue to Inquire.
	If ($PSBoundParameters['Debug']) {
		$DebugPreference = 'Continue'
		$Debug = $true
	}
	
	$LoggingPreference = "Continue"		# Log all output
	# Log file name for Write-Log function
	$LoggingFilePreference = join-path $ScriptLocation "DSF_Task_Log.txt"	

<#	# Setup for Selenium control of IE
	# Web Driver - location of DLL
	$SeWebDriverPath = join-path ( split-path (get-package -name Selenium.WebDriver).Source ) "\lib\net45"
	Add-Type -Path ( join-path $SeWebDriverPath "WebDriver.dll" )
	# IE Driver - location of EXE
	$SeIEDriverPath = join-path ( split-path (get-package -name Selenium.WebDriver.IEDriver).Source ) "\driver"
	$env:PATH += ";$($SeIEDriverPath)"
	# IE 64 Driver - location of EXE
	#$SeIEDriverPath = join-path ( split-path (get-package -name Selenium.WebDriver.IEDriver64).Source ) "\driver"
	#$env:PATH += ";$($SeIEDriverPath)"
	# Support - location of DLL
	$SeSupportPath = join-path ( split-path (get-package -name Selenium.Support).Source ) "\lib\net45"
	Add-Type -Path ( join-path $SeSupportPath "WebDriver.Support.dll" )
	
	# #### Required setup for IE11 with Selenium ####
	# #### According to:  https://github.com/SeleniumHQ/selenium/wiki/InternetExplorerDriver
	# "...set a registry entry ... so that the driver can maintain a connection to the instance 
	#	of Internet Explorer it creates"
	$IEFeatureControl = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Internet Explorer\Main\FeatureControl"
	$BFCache = "FEATURE_BFCACHE"
	if ( ( get-itemproperty ( join-path $IEFeatureControl $BFCache ) -EA SilentlyContinue | select -ExpandProperty "iexplore.exe" ) -ne 0 ) {
		throw "BFCache not set!"
	}
#>

	# Main site URL to start from
	$SiteURL = "https://store.adocument.net/DSF/"

	# To encrypt a password, put the password on the Clipboard and then run this:
	# Get-Clipboard | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Out-Clipboard

	# Make a Credential object using account name & password supplied to this script.
	$CryptPass = $Password | ConvertTo-SecureString -AsPlainText -Force
	$Credential = New-Object System.Management.Automation.PSCredential ($UserName,$CryptPass)

	# Snips of text we need to match, to find various controls or input fields
	#	Use '' instead of "" to avoid accidental substitutions.
	$ProductsLinkSnip = 'ctl00_ctl00_C_M_LinkColumn3_RepeaterCategories_ctl00_RepeaterItems_ctl02_HyperLinkItem'
	$LoginButtonSnip = 'ctl00_ctl00_C_W__loginWP__myLogin_Login'
	$UserFieldSnip = 'ctl00_ctl00_C_W__loginWP__myLogin__userNameTB'
	$PassFieldSnip = 'ctl00_ctl00_C_W__loginWP__myLogin__passwordTB'

	# What counts as a yes/no value?
	$YesValues = "yes","y","true","x"
	$NoValues = "no","n","false"
	
<#	# Create IE instance
	$IE = New-Object -ComObject 'InternetExplorer.Application'
	# Check if there are extra tabs open, and close them if so.
#>

<#	# Create IE instance
	$IE = New-Object OpenQA.Selenium.IE.InternetExplorerDriver
	# Check if there are extra tabs open, and close them if so.
#>
	
	# Create Firefox instance
	$Browser = Start-SeFirefox

	# Show the window -- not necessary for this to work, but useful to see what's going on.
#	$IE.Visible = $true
	# May want to set this, to prevent any IE popups like "Do you want to..."
	#$IE.Silent = $true
}

Process {
	try {
		write-log -fore cyan "Loading main URL: $SiteURL"
		Enter-SeUrl $SiteURL -Driver $Browser
		#$IE.Navigate( $SiteURL )
		#$IE | Invoke-IEWait
		#$CurrentDoc = $IE.Document	# for IE
		#$CurrentDoc = $IE	# for Selenium (don't know how, or if, we should get just the Document part)

		<#
		# Go to main Administration page
		if ( $IE.LocationName -ne "Home" ) {
			# Get links containing the admin snippet and click the first one, as there may be multiple.
			$AdminLink = @( Get-Link $IE.Document.Links -Href $AdminHref )[0]
			Click-Link $AdminLink
		}
		#>

		##### Log in
		# Get input fields
		$UserField = Find-SeElement -Driver $Browser -ID $UserFieldSnip
		$PassField = Find-SeElement -Driver $Browser -ID $PassFieldSnip
		
		# Fill in values from stored credential
		Send-SeKeys -Element $UserField -Keys $Credential.UserName
		Send-SeKeys -Element $PassField -Keys $Credential.GetNetworkCredential().Password
		# Find the Login button and click it
		$LoginButton = Find-SeElement -Driver $Browser -ID $LoginButtonSnip
		# Note, for some forms it may be better to match like this:
		#	$CurrentDoc.IHTMLDocument3_getElementsByTagName('input') | Where-Object {$_.type -eq "Submit" }
		#$LoginButton.Click()
		write-log -fore cyan "Logging in..."
		Click-Link $LoginButton

		# Verify that we're logged in.  There won't be an Administration link if we aren't.
		$AdminLinkSnip = 'myadmin-link'
		$AdminLink = Find-SeElement -Driver $Browser -ClassName $AdminLinkSnip
		if ( $AdminLink -notlike $null ) {
			write-log -fore green "Admin link found; successfully logged in!"
		} # else Wait-Link should have timed out and thrown an exception.
		
		##### Logged in
		
		# Now we'll be on the Storefront page, but we need to get to the Administration page.
		# We could do it by loading a URL, but it's more flexible to find the link and click it.
		Click-Link $AdminLink

		# Find Products link and click it.
		$ProductsLink = $Browser | Wait-LinkSe -TagName "a" -Property "id" -Pattern $ProductsLinkSnip
		Click-Link $ProductsLink
		exit
		
		# Now we're on the Product Management page.
		
		# Grab details from CSV
		# *** NOTE ***
		# We should probably bring this in from Excel instead.
		# Remember to trim leading/trailing whitespace from all values!
		$Products = import-csv $ProductFile
		$Counter = 1
		
		foreach ( $product in $Products ) {
			Manage-Product -Document $CurrentDoc -Mode Add -Product $product
			$Counter++
		}
		
	}
	
	catch {
		Handle-Exception $_
	}
	
	<#
		$evt = Register-ObjectEvent -InputObject $CurrentDoc -EventName DocumentComplete
	#>
	
	
	<#exit

	# Go to Products page
	$ProductsLink = @( Get-Link $IE.Document.Links -Href $ProductsHref )[0]
	Click-Link $ProductsLink
	#>
}

End {
	read-host "Press Enter to close IE and quit"
	# Close IE application
	$IE.Close()
	$IE.Dispose()
	$IE.Quit()
	
	#Release COM Object (not needed when using Selenium)
	#[void][Runtime.Interopservices.Marshal]::ReleaseComObject($IE)
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