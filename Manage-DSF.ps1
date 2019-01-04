<#	This script is horrible alpha quality and will not work!
	It's very much a work in progress as I learn how to automate things,
	in order to maybe, one day, not have to add dozens of products by hand.
	
	Maybe.
#>

#requires -Module Selenium

<#	Selenium class documentation for .NET:
		https://seleniumhq.github.io/selenium/docs/api/dotnet/index.html
#>

<# Can we automate product creation in DSF?
	Possibly -- see this page for ideas:
		http://www.westerndevs.com/simple-powershell-automation-browser-based-tasks/
	Use case similar to what I want to do:
		http://cmdrkeene.com/automating-internet-explorer-with-powershell
	
	See also this, which uses Invoke-Webrequest for a different approach:
		https://www.gngrninja.com/script-ninja/2016/7/8/powershell-getting-started-utilizing-the-web
	
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

[cmdletbinding()]

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
#>

Param (
	[ValidateScript({
		if ( ( $_ -like $null ) -or ( test-path $_ ) ) {
			$true
		} else {
			throw "ProductFile - Supplied path not found: $_!"
		}
	})]
	$ProductFile = "C:\Users\Carl\Documents\ADS Work\Automation\Test_Product_List.csv",

	[ValidateNotNullOrEmpty()]
	[string] $UserName = "DefaultUser",

	[ValidateNotNullOrEmpty()]
	[string] $Password,
	
	[switch] $SkipImageUpload = $false
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
			# Make sure link is clickable before trying it.
			$ClickableLink = WaitFor-ElementToBeClickable -Element $Link
			Invoke-SeClick $ClickableLink
			# Now wait for browser to process the click and load the next page
			$ClickableLink.WrappedDriver | Invoke-Wait
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
			Invoke-SeClick $ClickMe
			# Now wait for a bit
			Start-Sleep -Seconds $SleepTime
		} else {
			write-log -fore yellow "ClickMe is empty?"
		}
	}
	
	End {
	}
}

function Dump-ElementInfo {
	<#
		.Description
		Print out information about the supplied web element, for debugging purposes.
		Call this function instead of sticking "blah | gm | ft-auto | out-string" into the code.
		
		If you only want one section, specify it with the appropriate switch.
		If you want everything, specify -All.
		
		.Parameter WebElement
		Selenium web element to examine.
		
		.Parameter WebInfo
		Dump the element's properties, the ones you can see in the shell.
		
		.Parameter MemberInfo
		Dump the output of Get-Member for this object.
		
		.Parameter All
		Dump all types of info.
	#>
	
	param (
		[Parameter( Mandatory, ValueFromPipeLine )]
		[OpenQA.Selenium.Remote.RemoteWebElement] $WebElement,
		
		[switch] $WebInfo,
		
		[switch] $MemberInfo,
		
		[switch] $All
	)
	
	begin {}
	
	process {
		# Check if object is empty; if so, log a message but don't bother trying to dump info.
		if ( $WebElement -like $null ) {
			write-log -fore yellow "Warning: Attempted to dump info from a null element."
		} else {
			# WebElement info section.
			if ( $WebInfo -or $All ) {
				$Output = $WebElement | out-string
				write-log -fore gray $Output
			}
			
			# Member info section.
			if ( $MemberInfo -or $All ) {
				$Output = $WebElement | get-member | format-table -auto | out-string
				write-log -fore gray $Output
			}
		}
	}
	
	end {}
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

function Invoke-Login {
	<#
		.Synopsis
		Log into web site.  Return the URL of the page that loads after login.
		
		.Parameter SiteURL
		Link to login page.
		
		.Parameter UserName
		Account name to use when logging in.
		
		.Parameter Password
		Password to use when logging in.
		
		.Example
		$HomePage = Invoke-Login "https://www.test.com" "Ford42" "/towel/0"
	#>
	
	param (
		[Parameter(Position=0)]
		[string] $SiteURL,
		
		[Parameter(Position=1)]
		[string] $UserName,
		
		[Parameter(Position=2)]
		[string] $Password
	)
	
	try {
		# What do we return?
		$ReturnLink = $null
		
		# Web control names
		$LoginButtonSnip = 'ctl00_ctl00_C_W__loginWP__myLogin_Login'
		$UserFieldSnip = 'ctl00_ctl00_C_W__loginWP__myLogin__userNameTB'
		$PassFieldSnip = 'ctl00_ctl00_C_W__loginWP__myLogin__passwordTB'

		# Navigate to page and attempt to sign in.
		write-log -fore cyan "Loading site: $SiteURL"
		Enter-SeUrl $SiteURL -Driver $Browser
		# Issue 7: $Browser is a global variable but not intentionally so.  Fix this for better scoping.
		
		##### Log in
		# Get input fields
		$UserField = Find-SeElement -Driver $Browser -ID $UserFieldSnip
		$PassField = Find-SeElement -Driver $Browser -ID $PassFieldSnip
		
		# Fill in values from stored credential
		Set-TextField $UserField $UserName
		Set-TextField $PassField $Password
		# Find the Login button and click it
		$LoginButton = Find-SeElement -Driver $Browser -ID $LoginButtonSnip

		write-log -fore cyan "Logging in..."
		#Click-Link $LoginButton
		# Sleep after clicking, because we don't yet know how to reliably detect when storefront page is complete.
		#	Issue 2.
		#Click-Wait $LoginButton 30
		$LoginButton.Click()
		
		$ReturnLink = $Browser.Url

		return $ReturnLink
	}

	catch {}
	
	finally {
		# Do something, maybe emit optional debugging info about page details, load time, ???.
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
	$DocState = $BrowserObject.ExecuteScript($scrGetReadyState)
	while ( $DocState -notlike "complete" ) {
		# Not ready yet, so wait 1 second and check again.
		write-debug "Waiting for page load to complete..."
		Start-Sleep -Seconds 1
		$DocState = $BrowserObject.ExecuteScript($scrGetReadyState)
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
		
		.Parameter BrowserObject
		Object containing the web browser object we're using.
		
		.Parameter Mode
		Add (create new product) or Change (modify existing product).
		May include "Delete" in future.
	#>

	param (
		[Parameter( Mandatory )]
		[PSCustomObject] $Product,
		
		[Parameter( Mandatory )]
		[OpenQA.Selenium.Remote.RemoteWebDriver] $BrowserObject,
		
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
			Invoke-SeClick ( $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "*ButtonCreateProduct" )
			
			# Handle first page, which only asks for name and type.
			$ProductNameField = $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "*txtName"
			Set-TextField $ProductNameField $product.'Product Name'
			# We only deal with non-printed products, so set Type to 3
			$Picklist = Find-SeElement -Driver $BrowserObject -ID "ctl00_ctl00_C_M_drpProductTypes"
			$Picklist | Select-FromList -Item "Non Printed Products"
			
			# Click Next to get to product creation page
			Click-Link ( $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "*btnNext" )
			
			# The rest of the work is the same whether adding or updating, so let another function do it.
			$BrowserObject | Update-Product -Product $Product
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

function Run-JavaScript {
	<#
		.Synopsis
		Execute JavaScript using the specified web driver object.
		
		.Parameter WebDriver
		The driver/browser object that will execute the script.
		
		.Parameter Arguments
		Object or collection containing arguments to the script.
	#>
	
	# Nothing here yet; placeholder function in case we do need a JavaScript runner.
}

function Select-FromList {
	<#
		.Synopsis
		Given a target text string, find it in a list and select it.
		
		.Parameter ListObject
		Web list object, such as you'd get from Find-SeElement.
		
		.Parameter Item
		Item to select from the list of choices.
	#>

	param (
		[Parameter( Mandatory, ValueFromPipeLine )]
		$ListObject,
		
		[Parameter( Mandatory )]
		[string] $Item
	)
	
	Begin {}
	
	Process {
		try {
			# Status indicator
			$Success = $false
			
			# Create a Selenium Select object to find what we want.
			$Selector = New-Object -TypeName OpenQA.Selenium.Support.UI.SelectElement( $ListObject )
			# Have it select out target out of the list.
			# This will return NoSuchElementException if the option isn't found.
			$Selector.SelectByText( $Item )
			
			# Now verify the item is actually selected.
			# This is different from not being found; the item exists but wasn't selected for some reason.
			if ( $Selector.SelectedOption.Text -ne $Item ) {
				throw "Couldn't select `'$TargetItem`' as requested!"
			}
		}
		
		catch [OpenQA.Selenium.NoSuchElementException] {
			# SelectElement will throw this if the requested item wasn't found.
			# This could be a critical failure, so we want to log it.
			write-log -fore red "Error: Attempt to select '$Item' from list failed because it wasn't found."
			write-log -fore red "Actual list contains:"
			write-log -fore red "$( $ListObject | Out-String )"
		}
		
		catch {
			write-log -fore red "Error: Attempt to select '$Item' from list failed because the selection didn't stick."
		}
		
		finally {
			# Do we need to do anything for cleanup here, such as destroy the Select object?
		}
	}
	
	End {}
}

function Set-CheckBox {
	<#
		.Description
		Given a checkbox object and a desired state, check the checkbox' current state and 
		change it if necessary.  Default is to set it "Checked;" supply -Off parameter if you
		want to un-check it.
		
		.Synopsis
		Given a checkbox object and desired state, set the object accordingly.
		
		.Parameter CheckBoxObject
		Selenium web element representing the checkbox form object.
		
		.Parameter Off
		Set this if you want to uncheck the box.
	#>
	
	param (
		[Parameter( Mandatory, ValueFromPipeLine )]
		[OpenQA.Selenium.Remote.RemoteWebElement] $CheckBoxObject,
		
		[switch] $Off = $false
	)
	
	begin {}
	
	process {
		# Grab the ID attribute and save it for later.
		$ChkBoxID = $CheckBoxObject.GetAttribute("Id")
		$ChkBoxDriver = $CheckBoxObject.WrappedDriver
		
		# Check if we're turning it on or off.
		if ( $Off ) {
			# User requested uncheck.
			if ( $CheckBoxObject.Selected -eq $True ) {
				# Currently checked, so click to uncheck.
				$CheckBoxObject.Click()
				
				# Verify it's now unchecked.
				$RealCheckBox = WaitFor-ElementToBeClickable $ChkBoxDriver -ID $ChkBoxID 5
				if ( $RealCheckBox.Selected -ne $False ) {
					throw "Error:  Unable to uncheck $($RealCheckBox.GetAttribute("ID"))!"
				}
			}
		} else {
			# User requested check.
			if ( $CheckBoxObject.Selected -eq $False ) {
				# Currently not checked, so click to check.
				$CheckBoxObject.Click()
				
				# Sometimes, changing checkbox state causes the page to be refreshed.
				# If this happens, the element reference will become stale.
				
				# Verify it's now checked.
				$RealCheckBox = WaitFor-ElementToBeClickable $ChkBoxDriver -ID $ChkBoxID 5
				if ( $RealCheckBox.Selected -ne $True ) {
					throw "Error:  Unable to check $($RealCheckBox.GetAttribute("ID"))!"
				}
			}
		}
	}
	
	end {}
}

function Set-RichTextField {
	<#
		.Synopsis
		Find a rich text field, then set its value to the supplied string.
		
		.Description
		This function will place text into a rich text editor, when supplied with a way of
		finding it.  In the case of an editor in an iFrame, you'll need to supply the iFrame 
		itself, plus either ID or XPath of the edit field.
		
		If the editor isn't in an iFrame, just supply the ID as a named parameter.
		
		.Parameter FieldObject
		The iFrame containing the rich text editor.
		
		.Parameter ID
		ID of the edit field within the iFrame.
		
		.Parameter XPath
		XPath to the edit field within the iFrame.
		
		.Parameter Text
		String to put into the field.  If not supplied, defaults to empty string.
		Note that rich text can include formatting.  Ideally, find some way to pass
		formatted text through to the editor without losing formatting.
	#>
	
	<#	Parameter planning
	
		***	Note, apparently for parameter sets to work, each one must have one parameter
			that is not used in any other set.  So, in our case the "ID" and "XPath" sets
			should work because they have unique parameters, but "NonIFrame" needs a 
			unique parameter.  Let's solve that by calling it RichEditFrame instead.
	
		Expected uses:
			Editor is in an iFrame, and can be identified by ID after switching to it.
			We'd require:
				FieldObject (web element, specifically the iFrame in question)
				ID (ID tag to search for)
				Text (string to put into the edit field)
				
			Editor is in an iFrame, and can be identified by XPath after switching to it.
			We'd require:
				FieldObject (web element, specifically the iFrame in question)
				XPath (XPath to follow)
				Text (string to put into the edit field)
				
			Editor is *NOT* in an iFrame, so we don't need to switch to it; caller will
			supply the FieldObject similar to calling Set-TextField.  We need:
				FieldObject (the editor itself)
				Text
	#>
	
	param(
		[Parameter( Mandatory, ParameterSetName="ID" )]
		[Parameter( Mandatory, ParameterSetName="XPath" )]
		[OpenQA.Selenium.Remote.RemoteWebElement] $FieldObject,
		
		[Parameter( Mandatory, ParameterSetName="NonIFrame" )]
		[OpenQA.Selenium.Remote.RemoteWebElement] $RichEditFrame,
		
		[Parameter( Mandatory, ParameterSetName="ID" )]
		[string] $ID,
		
		[Parameter( Mandatory, ParameterSetName="XPath" )]
		[string] $XPath,
		
		[Parameter( Mandatory, ParameterSetName="ID" )]
		[Parameter( Mandatory, ParameterSetName="XPath" )]
		[Parameter( Mandatory, ParameterSetName="NonIFrame" )]
		[string] $Text
	)
	
	if ( $PSCmdlet.ParameterSetName -eq "NonIFrame" ) {
		# We have a rich text editor that isn't in an iFrame, so the caller
		#	has supplied only the FieldObject and Text parameters.
		$Editor = $RichEditFrame
		
		# Make sure we actually got something we can use.
		if ( $Editor -notlike $null ) {
			# Clear anything that's already there.
			$Editor.Clear()
			# Send text to editor field.
			$Editor.SendKeys( $Text )
		} else {
			throw "Editor object seems to be empty."
		}
	} else {
		# The editor will be inside an iFrame.  To navigate to the actual edit field,
		#	we first need to switch focus to the iFrame.
		$BrowserObject = $FieldObject.WrappedDriver
		$NewFrame = $BrowserObject.SwitchTo().Frame($FieldObject)
		# $NewFrame is, on Firefox at least, of type OpenQA.Selenium.Firefox.FirefoxDriver.
		#$NewFrame | get-member | format-table -auto | out-string | write-host
		#$NewFrame | Out-string | write-host
		
		# Next, search for the element in the new context.
		switch ( $PSCmdlet.ParameterSetName ){
			"ID"	{ $EditorIFrame = Find-SeElement -Element $NewFrame -ID $ID }
			"XPath"	{ $EditorIFrame = Find-SeElement -Element $NewFrame -XPath $XPath }
		}
		
		# Make sure we actually found something.
		if ( $EditorIFrame -notlike $null ) {
			# By snooping in Dev Mode, we find the editor HTML looks for a Click event,
			#	on which it sets focus to the text area.
			# So, let's invoke its Click event.
			$EditorIFrame.Click()
			# Clear the field first, in case it already has text.
			$EditorIFrame.Clear()
			
			# Set it to the string we were given.
			$EditorIFrame.SendKeys( $Text )
		} else {
			# Throw an error specific to which set was used.
			switch ( $PSCmdlet.ParameterSetName ){
				"ID"	{ throw "Couldn't find an iFrame matching ID `'$ID`'." }
				"XPath"	{ throw "Couldn't find an iFrame matching XPath `'$XPath`'." }
			}
		}
		
		# Switch browser back to the parent frame so it will be able to find stuff
		#	outside the iFrame.
		$null = $BrowserObject.SwitchTo().parentFrame()
	}
}

function Set-TextField {
	<#
		.Synopsis
		Given a fillable text field, set its value to the supplied string.
		
		.Parameter FieldObject
		Text field to fill.
		
		.Parameter Text
		String to put into the field.  If not supplied, defaults to empty string.
	#>
	
	param(
		[Parameter( Mandatory, Position=1 )]
		[OpenQA.Selenium.Remote.RemoteWebElement] $FieldObject,
		
		[Parameter( Position=2 )]
		[string] $Text = ""
	)
	
	# Clear the field first.
	$FieldObject.Clear()
	
	# Set it to the string we were given.
	$FieldObject.SendKeys( $Text )
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

	# Product Name will already be filled in, based on the previous page.
	
	# Display As, max length unknown
	#	Supposedly, product name as customer sees it in the storefront catalog.
	#	In reality, rarely seen except when editing product.
	if ( $Product.'Display Name' -notlike $null ) {
		$Field = Find-SeElement -Driver $BrowserObject -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__StorefrontName"
		Set-TextField $Field $Product.'Display Name'
		#( $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "*StorefrontName" ).Value = $Product.'Display Name'
	}
	
	# Product ID (SKU), 50 chars max
	if ( $Product.'Product ID' -notlike $null ) {
		#( $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "*SKU" ).Value = $Product.'Product Id'
		$Field = Find-SeElement -Driver $BrowserObject -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__SKU"
		Set-TextField $Field $Product.'Product Id'
	}
	
	<#	Dealing with rich text editors!
		The product Brief description field is a rich text editor in an iFrame.
		Unlike "textbox" elements, you can't just set the value and move on.
		Each one is an iFrame, which you can get like any other element, however
		after that you have to drill down a bit.
		
		$iFrame.ContentWindow.Document.Body.innerHTML or .innerText
		
		This is fine when you're manipulating the objects directly, but now with Selenium
		we're using browser-independent methods so we can't do that.
	#>
	
	# Brief Description, rich text field
	if ( $Product.'Brief Description' -notlike $null ) {
		#$iFrame = $BrowserObject | Wait-Link -TagName "iframe" -Property "id" -Pattern "*Description_contentIframe"
		#$iFrame.ContentWindow.Document.Body.innerHTML = $Product.'Brief Description'
		$iFrame = Find-SeElement -Driver $BrowserObject -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__Description_contentIframe"
		Set-RichTextField -FieldObject $iFrame -XPath "/html/body" -Text $Product.'Brief Description'
	}
	
	# Now, if there's a thumbnail image to upload, do that.
	#	Issue 1
	#	As long as the web form accepts a valid file path WITHOUT the user populating it via a dialog,
	#	this should be possible.  Upload-Thumbnail will validate the path before attempting to upload, logging 
	#	an error if it's bad.
	
	# If we have a path to image file, and SkipImageUpload isn't set, try to upload.
	if ( $Product.'Product Icon' -notlike $null ) {
		# Log a message if SkipImageUpload is set.
		if ( $SkipImageUpload ) {
			write-log -fore yellow "Warning: File path provided but SkipImageUpload is set; ignoring."
		} else {
			# Upload the image file 
			Upload-Thumbnail -BrowserObject $BrowserObject -ImageURI $Product.'Product Icon'
		}
	}
	
	# Switch to Details section.
	#	<a class="rtsLink rtsAfter" id="TabDetails" href="#">...
	#$NavTab = Find-SeElement -Driver $BrowserObject -ID "TabDetails"
	$NavTab = $BrowserObject | Wait-Link -TagName "a" -Property "id" -Pattern "TabDetails"
	$NavTab.Click()
	
	<#		Details section
			x	Long Description, div id="ctl00_ctl00_C_M_ctl00_W_ctl01__LongDescription_contentDiv"
					Limit 4000 chars
	#>

	# Long Description
	if ( $Product.'Long Description' -notlike $null ) {
		$LongDescField = Find-SeElement -Driver $BrowserObject -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__LongDescription_contentDiv"
		# This editor isn't in an iFrame.
		Set-RichTextField -RichEditFrame $LongDescField -Text $Product.'Long Description'
	}
	
	# Switch to Settings section.
	$NavTab = $BrowserObject | Wait-Link -TagName "a" -Property "id" -Pattern "TabSettings"
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
		$Picklist = $BrowserObject | Wait-Link -TagName "select" -Property "id" -Pattern "ctl00_ctl00_C_M_ctl00_W_ctl01__Rank_DropDownListRank"
		$Set = $Picklist | Select-FromList $Product.'Display Priority'
		# Check result of request; log a message if it defaults to Standard.
		if ( $Set -ne $true ) {
			write-log -fore yellow "Warning: No Display Priority option matched the imported data; setting to Standard."
			$null = $Picklist | Select-FromList "Standard"
		}
<#		if ( $Picklist.innerHTML -eq $Product.'Display Priority' ) {
			( $Picklist | where innerHTML -eq $Product.'Display Priority' ).Selected = $true 
		} else {
			write-log -fore yellow "Warning: No Display Priority option found to match `'$($Product.'Display Priority')`'; setting to Standard."
			( $Picklist | where innerHTML -eq "Standard" ).Selected = $true 
		}
#>
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
	if ( $Product.'Start Date' -notlike $null ) {
		<#	To make sure we give the form valid data, we will cast the input from CSV as a 
			PowerShell DateTime object, which can do smart conversion of things like "2017-07-21", 
			"4 Jul 2017" etc.  It will even handle things like "7/22", though it must assume the 
			current year is meant.
			Then, we call ToShortDateString(), which will always output "7/22/2017", because
			that is the format DSF is expecting.
		#>
		$StartDate = [DateTime]$Product.'Start Date'
		$StartDateField = $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "ctl00_ctl00_C_M_ctl00_W_ctl01__ProductActivationCtrl__Begin_dateInput_text"
		Set-TextField $StartDateField $StartDate.ToShortDateString()
	} else {
		# Start Date is empty, so set product to Active.
		$RadioButton = $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "ctl00_ctl00_C_M_ctl00_W_ctl01__ProductActivationCtrl__YesNo_1"
		$RadioButton.Click()
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
	if ( $Product.'End Date' -notlike $null ) {
		# Click the button to select End Date.
		$RadioButton = $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "ctl00_ctl00_C_M_ctl00_W_ctl01__ProductActivationCtrl_rdbEndDate"
		$RadioButton.Click()
		# Now set the date.
		$StopDate = [DateTime]$Product.'End Date'
		$StopDateField = $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "ctl00_ctl00_C_M_ctl00_W_ctl01__ProductActivationCtrl__End_dateInput_text"
		Set-TextField $StopDateField $StopDate.ToShortDateString()
	} else {
		# End Date is empty, so set to Never.
		$RadioButton = $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "ctl00_ctl00_C_M_ctl00_W_ctl01__ProductActivationCtrl_rdbNever"
		$RadioButton.Click()
	}
	
	<#
		x	Turn Around Time, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_TurnAroundTimeCtrl_rdbNone" type="radio" checked="checked" value="rdbNone"
				either "None" or number of days, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_TurnAroundTimeCtrl__Value"
				number of days field is disabled if "None" is selected
				To supply a value, first click:
					id="ctl00_ctl00_C_M_ctl00_W_ctl01_TurnAroundTimeCtrl_rdbValue"
	#>

	# Turnaround time is the same deal -- combo radio button and text field.
	if ( $Product.'Turnaround Time' -notlike $null ) {
		# Set Value (the second radio button)
		$RadioButton = $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "ctl00_ctl00_C_M_ctl00_W_ctl01_TurnAroundTimeCtrl_rdbValue"
		$RadioButton.Click()
		# Now fill in the number of days.
		$TurnaroundField = $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "ctl00_ctl00_C_M_ctl00_W_ctl01_TurnAroundTimeCtrl__Value"
		Set-TextField $TurnaroundField $Product.'Turnaround Time'
	} else {
		# None specified; set radio button to None.
		$RadioButton = $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "ctl00_ctl00_C_M_ctl00_W_ctl01_TurnAroundTimeCtrl_rdbNone"
		$RadioButton.Click()
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
	$ExemptShipChk = $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "ctl00_ctl00_C_M_ctl00_W_ctl01_chkShippingExempt"
	if ( $Product.'Exempt Shipping' -in $YesValues ) {
		Set-CheckBox $ExemptShipChk
	} else {
		Set-CheckBox $ExemptShipChk -Off
	}
	
	# Exempt from Sales Tax?
		$ExemptTaxChk = $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "ctl00_ctl00_C_M_ctl00_W_ctl01_chkTaxExempt"
	if ( $Product.'Exempt Tax' -in $YesValues ) {
		Set-CheckBox $ExemptTaxChk
	} else {
		Set-CheckBox $ExemptTaxChk -Off
	}
	
	# Show on the mobile version of the site?
	$MobileSupportYes = $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "ctl00_ctl00_C_M_ctl00_W_ctl01_IsMobileSupportedList_0"
	$MobileSupportNo = $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "ctl00_ctl00_C_M_ctl00_W_ctl01_IsMobileSupportedList_1"
	if ( ( $Product.'Mobile' -in $YesValues ) -or ( $Product.'Mobile' -like $null ) ) {
		# Set if Yes or unspecified.
		$MobileSupportYes.Click()
	} else {
		# Set No if we specifically don't want mobile support.
		$MobileSupportNo.Click()
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
	$MgInvenChk = $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "ctl00_ctl00_C_M_ctl00_W_ctl01_chkManageInventory"
	
	switch ( $Product.'Manage Inventory' ) {
		# If explicitly set to "Yes," turn the checkbox on.
		{ $_ -in $YesValues }	{ write-debug "ManageInv = True" ; $ManageInventory = $true }
		
		# If explicitly set to "No," turn the checkbox off.
		{ $_ -in $NoValues }	{ write-debug "ManageInv = False" ; $ManageInventory = $false }
		
		# Enable management if any inventory management values are given,
		#	even if Manage isn't specified.  (Sanity check!)
		{ $Product.Threshold -or
			$Product.'Allow Back Order' -or
			$Product.'Show Inventory with Back Order' -or
			$Product.'Add to Inventory' -or
			$Product.'Reset Inventory'
		}						{ 
									# Enable and log a warning.
									$ManageInventory = $true
									write-log -fore yellow "Warning: Manage Inventory not specified, but management items were; enabling."
								}
	}
	
	# So if none of the conditions are met, $ManageInventory will still be NULL.
	
	# Now, if ManageInventory is TRUE, handle the values that depend on it.
	# In order to do this, we may need to take steps to activate the section of the form
	#	that is deactivated if this checkbox is not checked.
	# For now, see what happens if we submit the form anyway.
	if ( $ManageInventory -eq $true ) {
		# Check the box for Manage Inventory = Enabled.
		Set-CheckBox $MgInvenChk
		
		# Threshold, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_txtThQty"
		if ( $Product.Threshold -notlike $null ) {
			$InvThreshold = $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "ctl00_ctl00_C_M_ctl00_W_ctl01_txtThQty"
			Set-TextField $InvThreshold $Product.Threshold
		}
		
		# Allow Back Order, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_chkBackOrderAllowed"
		$AllowBkOrdChk = $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "ctl00_ctl00_C_M_ctl00_W_ctl01_chkBackOrderAllowed"
		if ( $Product.'Allow Back Order' -in $YesValues ) {
			Set-CheckBox $AllowBkOrdChk
		} else {
			Set-CheckBox $AllowBkOrdChk -Off
		}
		
		# Show inventory when back order is allowed, input id="ctl00_ctl00_C_M_ctl00_W_ctl01_chkShowInventoryWhenBackOrderIsAllowed"
		$ShowInvBOChk = $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "ctl00_ctl00_C_M_ctl00_W_ctl01_chkShowInventoryWhenBackOrderIsAllowed"
		if ( $Product.'Show Inventory with Back Order' -in $YesValues ) {
			Set-CheckBox $ShowInvBOChk
		} else {
			Set-CheckBox $ShowInvBOChk -Off
		}
		
		# Notification Email Id, id="ctl00_ctl00_C_M_ctl00_W_ctl01_txtEmailId"
		$NotifyEmailField = $BrowserObject | Wait-Link -TagName "textarea" -Property "id" -Pattern "ctl00_ctl00_C_M_ctl00_W_ctl01_txtEmailId"
		if ( $Product.'Notify Emails' -notlike $null ) {
			Set-TextField $NotifyEmailField $Product.'Notify Emails'
		}
		
		<#	Replenish inventory - Note this is either one or the other!
				o Add XXX to existing
				o Reset to XXX
			
			For some reason, calling SetActive or Click on these radio buttons causes the web form
			to freeze -- at least from the GUI.  So, try proceeding without doing that.
		#>
		#[nwch]
		if ( $Product.'Add to Inventory' -notlike $null ) {
			$RadioButton = ( $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "*RbInvAddToExistingInv" )
			$RadioButton.isDisabled = $false
			#$RadioButton.SetActive()
			$RadioButton | Click-Wait
			( $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "*RbInvAddToExistingInvTextBox" ).Value = $Product.'Add to Inventory'
		} elseif ( $Product.'Reset Inventory' -notlike $null ) {
			$RadioButton = ( $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "*RbInvReset" )
			$RadioButton.isDisabled = $false
			#$RadioButton.SetActive()
			$RadioButton | Click-Wait
			( $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "*RbResetInvTextBox" ).Value = $Product.'Reset Inventory'
		}
	} elseif ( $ManageInventory -eq $false ) {
		# Turn the checkbox off.
		Set-CheckBox $MgInvenChk -Off
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
			$RadioButton = ( $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "*FixedQuantities" )
			#$RadioButton.SetActive()
			$RadioButton | Click-Wait
			( $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "*FixedQuantitiesValues_ctl02__Value" ).Value = $Product.'Fixed Qty'
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
			$RadioButton = ( $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "*_Multiples" )
			$RadioButton | Click-Wait
			( $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "*OrderQuantitiesCtrl__Minimum" ).Value = $Product.'Min Qty'
			( $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "*OrderQuantitiesCtrl__Maximum" ).Value = $Product.'Max Qty'
			( $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "*OrderQuantitiesCtrl__Multiple" ).Value = $Product.'Mult Qty'
		}
		
		# If a Max quantity was specified, check the box to enforce this in shopping cart.
		if ( $Product.'Max Qty' -notlike $null ) {
			$Checkbox = ( $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "*chkEnforceMaxQtyInCart" )
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
		$Field = Find-SeElement -Driver $BrowserObject -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__ProductionNotes"
		Set-TextField $Field $Product.'Production Notes'
	}
	
	if ( $Product.Keywords -notlike $null ) {
		$Field = Find-SeElement -Driver $BrowserObject -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__Keywords"
		Set-TextField $Field $Product.Keywords
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
	$Field = Find-SeElement -Driver $BrowserObject -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_WeightCtrl__Weight"
	if ( $Product.Weight -notlike $null ) {
		Set-TextField $Field $Product.Weight
	} else {
		# OK, somehow the weight IS null, so ensure it's set to zero.
		# A zero weight will allow the product to be created, though it won't work for shipping quotes.
		Set-TextField $Field "0"
	}
	# Weight units - Get the list object, then select the right value.
	$WeightList = Find-SeElement -Driver $BrowserObject -ID "ctl00_ctl00_C_M_ctl00_W_ctl01_WeightCtrl__Unit"
	$WeightList | Select-FromList -Item ( $Product.'Weight Unit' | FixUp-Unit )
	
	# Ship item separately?
	$Checkbox = $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "*chkPackSeparately"
	$Checkbox.Checked = ( $Product.'Ship Separately' -in $YesValues )
	
	# Width
	$Width = $Product.Width | ForEach-Object { if ( $_ -notlike $null ) { $_ } else { 0 } }
	( $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "*BoxX__Length" ).Value = $Product.Width
	$UnitList = $BrowserObject | Wait-Link -TagName "select" -Property "id" -Pattern "*ShipmentDimensionCtrl__BoxX__Unit"
	$UnitList | Select-FromList -Item ( $Product.'Width Unit' | FixUp-Unit )
	
	# Length
	$Length = $Product.Length | ForEach-Object { if ( $_ -notlike $null ) { $_ } else { 0 } }
	( $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "*BoxY__Length" ).Value = $Product.Length
	$UnitList = $BrowserObject | Wait-Link -TagName "select" -Property "id" -Pattern "*BoxY__Unit"
	$UnitList | Select-FromList -Item ( $Product.'Length Unit' | FixUp-Unit )
	
	# Height
	$Height = $Product.Height | ForEach-Object { if ( $_ -notlike $null ) { $_ } else { 0 } }
	( $BrowserObject | Wait-Link -TagName "input" -Property "id" -Pattern "*BoxZ__Length" ).Value = $Product.Height
	$UnitList = $BrowserObject | Wait-Link -TagName "select" -Property "id" -Pattern "*BoxZ__Unit"
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

function Upload-Thumbnail {
	param (
		[Parameter( Mandatory )]
		[OpenQA.Selenium.Remote.RemoteWebDriver] $BrowserObject,

		[ValidateNotNullOrEmpty()]
		[string] $ImageURI
	)
	
	# Verify file actually exists.
	if ( test-path $ImageURI ) {
		# Seems legit; proceed with upload process.
		# Start by clicking "Edit" button.
		$EditButton = Find-SeElement -Driver $BrowserObject -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__BigIconByItself_EditProductImage"
		if ( $EditButton -notlike $null ) {
			$EditButton.Click()
			# Once clicked, image graphic is replaced with a set of radio buttons.
			# Select "Upload Custom Icon" to proceed.
			$UploadIconButton = Find-SeElement -Driver $BrowserObject -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__BigIconByItself_ProductIcon_rdbUploadIcon"
			$UploadIconButton.Click()
			# Now we have a checkbox and a text field to manipulate.
			# Check the box to use this image for all of this product's thumbnails.
			$SameImageForAllChk = Find-SeElement -Driver $BrowserObject -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__BigIconByItself_ProductIcon_ChkUseSameImageIcon"
			Set-CheckBox $SameImageForAllChk
			# Try to set the text field so we don't have to mess with a file dialog.
			$ThumbnailField = Find-SeElement -Driver $BrowserObject -Name "ctl00$ctl00$C$M$ctl00$W$ctl01$_BigIconByItself$ProductIcon$_uploadedFile$ctl01"
			# However, the element may not accept keyboard input.  Try to change its display style
			#	to work around this problem.
			<#
				field = driver.find_element_by_id("selectedFile")
				driver.execute_script("arguments[0].style.display = 'block';", field)
			#>
			$ScriptString = 'document.getElementsByName("ctl00$ctl00$C$M$ctl00$W$ctl01$_BigIconByItself$ProductIcon$_uploadedFile$ctl01")[0].style.display = ''block'';'
			$BrowserObject.ExecuteScript( $ScriptString )
			# If that worked, we should be able to set the text field now.
			# Find it again because it has now changed?
			$ThumbnailField = Find-SeElement -Driver $BrowserObject -Name "ctl00$ctl00$C$M$ctl00$W$ctl01$_BigIconByItself$ProductIcon$_uploadedFile$ctl01"
			Set-TextField $ThumbnailField $ImageURI
			# Click the "Upload" button, which will cause the page to reload.
			$UploadButton = Find-SeElement -Driver $BrowserObject -ID "ctl00_ctl00_C_M_ctl00_W_ctl01__BigIconByItself_ProductIcon_Upload"
			$UploadButton.Click()
		} else {
			throw "Error: Couldn't find Edit button for image upload!"
		}
	}
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

	Write-DebugLog "Wait for $TagName element with $Property matching `'$Pattern`'"
	
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

function WaitFor-ElementExists {
	<#
		.Synopsis
		Given some way of finding a web element, wait until it exists somewhere in the page.
		
		.Description
		Try to find a web element using the specified information.  When it exists, return the element.
		
		Note that just because it exists, it's not necessarily usable yet.  See WaitFor-ElementToBeClickable.
		
		.Parameter WebDriver
		Web driver (browser) to use in finding the element.
		
		.Parameter TimeInSeconds
		Number of seconds to wait before giving up.  Default is 10.
		
		.Parameter ID
		ID tag to search by.
		
		.Parameter XPath
		XPath to search by.
		
		.Parameter LinkText
		Link text string to search by.
	#>
	
	param (
		[Parameter( Mandatory, Position=1, ParameterSetName="ID" )]
		[Parameter( Mandatory, Position=1, ParameterSetName="XPath" )]
		[Parameter( Mandatory, Position=1, ParameterSetName="LinkText" )]
		[OpenQA.Selenium.Remote.RemoteWebDriver] $WebDriver,

		[Parameter( ParameterSetName="ID" )]
		[Parameter( ParameterSetName="XPath" )]
		[Parameter( ParameterSetName="LinkText" )]
		[int] $TimeInSeconds = 10,

		[Parameter( ParameterSetName="ID" )]
		[string] $ID,
		
		[Parameter( ParameterSetName="XPath" )]
		[string] $XPath,
		
		[Parameter( ParameterSetName="LinkText" )]
		[string] $LinkText
	)
	
	# This object's job is to wait for something.
	$Waiter = New-Object OpenQA.Selenium.Support.UI.WebDriverWait($WebDriver, $TimeInSeconds)
	
	try {
		# Check which info we're given, then wait until it exists or times out.
		# If the waiter times out, it will throw an exception, which we'll handle.
		switch ( $PSCmdlet.ParameterSetName ) {
			"ID"		{
				$Locator = "ID: $ID"
				$Gotcha = $Waiter.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists( [OpenQA.Selenium.by]::Id($ID)))
			}
			"XPath"		{
				$Locator = "XPath: $XPath"
				$Gotcha = $Waiter.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists( [OpenQA.Selenium.by]::XPath($XPath)))
			}
			"LinkText"	{
				$Locator = "LinkText: $LinkText"
				$Gotcha = $Waiter.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists( [OpenQA.Selenium.by]::LinkText($LinkText)))
			}
		}
		
		# We got something, so return the element to caller.
		return $Gotcha
	}
	
	catch [OpenQA.Selenium.WebDriverTimeoutException] {
		# Timed out waiting for element.  What should we do here?
		# Nothing.  Caller must check if something was returned.
		write-log -fore yellow "Timed out waiting for $Locator"
	}
}

function WaitFor-ElementToBeClickable {
	<#
		.Synopsis
		Given some way of finding a web element, wait until it exists and can be clicked.
		
		.Description
		Try to find a web element using the specified information.  When it is clickable, return the element.
		
		.Parameter WebDriver
		Web driver (browser) to use in finding the element.
		
		.Parameter WebElement
		Web element to wait for.
		
		.Parameter ID
		ID tag to search by.
		
		.Parameter XPath
		XPath to search by.
		
		.Parameter LinkText
		Link text string to search by.
		
		.Parameter TimeInSeconds
		Number of seconds to wait before giving up.  Default is 10.
	#>
	
	param (
		[Parameter( Mandatory, Position=1, ParameterSetName="ID" )]
		[Parameter( Mandatory, Position=1, ParameterSetName="XPath" )]
		[Parameter( Mandatory, Position=1, ParameterSetName="LinkText" )]
		[OpenQA.Selenium.Remote.RemoteWebDriver] $WebDriver,

		[Parameter( Position=1, ParameterSetName="Element" )]
		[OpenQA.Selenium.Remote.RemoteWebElement] $WebElement,
		
		[Parameter( Position=2, ParameterSetName="ID" )]
		[string] $ID,
		
		[Parameter( Position=2, ParameterSetName="XPath" )]
		[string] $XPath,
		
		[Parameter( Position=2, ParameterSetName="LinkText" )]
		[string] $LinkText,
		
		[Parameter( Position=3, ParameterSetName="ID" )]
		[Parameter( Position=3, ParameterSetName="XPath" )]
		[Parameter( Position=3, ParameterSetName="LinkText" )]
		[Parameter( Position=2, ParameterSetName="Element" )]
		[int] $TimeInSeconds = 10
	)
	
	if ( $WebElement -notlike $null ) {
		$WebDriver = $WebElement.WrappedDriver
	}
	
	# This object's job is to wait for something.
	$Waiter = New-Object OpenQA.Selenium.Support.UI.WebDriverWait($WebDriver, $TimeInSeconds)
	
	try {
		# Check which info we're given, then wait until it exists or times out.
		# If the waiter times out, it will throw an exception, which we'll handle.
		switch ( $PSCmdlet.ParameterSetName ) {
			"ID"		{
				$Locator = "ID: $ID"
				$Gotcha = $Waiter.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementToBeClickable( [OpenQA.Selenium.by]::Id($ID)))
			}
			"XPath"		{
				$Locator = "XPath: $XPath"
				$Gotcha = $Waiter.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementToBeClickable( [OpenQA.Selenium.by]::XPath($XPath)))
			}
			"LinkText"	{
				$Locator = "LinkText: $LinkText"
				$Gotcha = $Waiter.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementToBeClickable( [OpenQA.Selenium.by]::LinkText($LinkText)))
			}
			"Element"	{
				$Locator = "Element: $($WebElement.TagName)"
				$Gotcha = $Waiter.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementToBeClickable( $WebElement ))
			}
		}
		
		# We got something, so return the element to caller.
		return $Gotcha
	}
	
	catch [OpenQA.Selenium.WebDriverTimeoutException] {
		# Timed out waiting for element.  What should we do here?
		# Nothing.  Caller must check if something was returned.
		write-log -fore yellow "Timed out waiting for clickable $Locator"
	}
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

	# Snips of text we need to match, to find various controls or input fields
	#	Use '' instead of "" to avoid accidental substitutions.
	#$AdminLinkSnip = 'myadmin-link'
	$AdminLinkText = "Administration"
	$ProductsLinkSnip = 'ctl00_ctl00_C_M_LinkColumn3_RepeaterCategories_ctl00_RepeaterItems_ctl02_HyperLinkItem'

	# What counts as a yes/no value?
	$YesValues = "yes","y","true","x"
	$NoValues = "no","n","false"
	
	# Selenium script snippets, to run using ExecuteScript:
	$scrGetReadyState = "return document.readyState"
	
	# Wait for element to become visible if hidden
	# From:  https://stackoverflow.com/questions/44724185/element-myelement-is-not-clickable-at-point-x-y-other-element-would-receiv
	$scrWaitUntilVisible = @"
WebDriverWait wait3 = new WebDriverWait(driver, 10);
wait3.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("ele_to_inv")));
"@
	
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

	$StorefrontURL = Invoke-Login $SiteURL -UserName $UserName -Password $Password
	# Just a string, the URL of the page that loads after logging in.
	
	# Wait a few seconds and check if LoadingSpinner is visible.
	# If "display" attribute is "none" then it's hidden and shouldn't obscure the link.
	$LoadingSpinner = $Browser.FindElementByID("loadingSpinner")
	Dump-ElementInfo $LoadingSpinner -WebInfo
	write-host ( $LoadingSpinner.GetAttribute("outerHTML") | out-string )
	# Wait for spinner to be hidden.
	$WaitCount = 1
	while ( $LoadingSpinner.GetAttribute("display") -notlike "none" ) {
		write-host "Waiting for Loading Spinner:  $WaitCount"
		write-host "Spinner attribute 'display' = $($LoadingSpinner.GetAttribute("display"))"
		$WaitCount++
		Start-Sleep -Seconds 1
	}

	# Verify that we're logged in.  There won't be an Administration link if we aren't.
	#$AdminControl = $Browser.FindElementByCssSelector(".myadmin-link")
	$AdminLink = $Browser.FindElementsByTagName("span") | where { $_.GetAttribute("ng-localize") -eq "StoreFront.Administration" }
	Dump-ElementInfo $AdminLink -WebInfo
	# Admin link exists; now we have to wait until it's not obscured by "Loading" gizmo.
	# By now, the element should no longer be obscured.
	#$AdminClickable = WaitFor-ElementToBeClickable $Browser -LinkText "Administration" -TimeInSeconds 30
	$AdminClickable = WaitFor-ElementToBeClickable -WebElement $AdminLink -TimeInSeconds 30
	if ( $AdminClickable -notlike $null ) {
		write-log -fore green "Admin link found; successfully logged in!"
		$AdminClickable.Click()
	} else {
		Dump-ElementInfo $AdminClickable -WebInfo
		throw "Error: Unable to log in; clickable link to Administration page not found."
	}
}

Process {
	<#	This section will be run once for each object passed to the script.
	#>

	try {
		# BEGIN should have gotten us to the Administration page.
		#Click-Link ( $Browser | Wait-Link -TagName "a" -Property "text" -Pattern "Administration" )

		# Find Products link and click it.
		#$ProductsLink = $Browser | Wait-LinkSe -TagName "a" -Property "id" -Pattern $ProductsLinkSnip
		$ProductsLink = Find-SeElement -Driver $Browser -ID $ProductsLinkSnip
		Click-Link $ProductsLink
		
		# Now we're on the Product Management page.
		
		# Grab details from CSV
		# *** NOTE ***
		# We should probably bring this in from Excel instead.
		# Remember to trim leading/trailing whitespace from all values!
		$Products = import-csv $ProductFile
		$Counter = 1
		
		foreach ( $product in $Products ) {
#			Manage-Product -Document $CurrentDoc -Mode Add -Product $product
			Manage-Product -BrowserObject $Browser -Mode Add -Product $product
			$Counter++
		}
		
	}
	
	catch {
		Handle-Exception $_
	}
	
}

End {
	read-host "Press Enter to close browser and quit"
	# Shut down driver and close browser
	Stop-SeDriver $Browser
	# if ('browser still running') get-process $BrowserPID | stop-process
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