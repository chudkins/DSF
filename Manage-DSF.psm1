#requires -Module Selenium

#[cmdletbinding()]

<#
	Are any of these functions used ONLY within other functions of this module?
	
	If so, we need to use an Export statement so that they aren't exposed.
#>

function Click-Link {
	<#
		.Synopsis
		Given something that's clickable, wait until browser is not busy, then click it.
		
		.Description
		Given an object with a clickable link, wait for browser to not be busy, then click the object.
		A WebElement may contain a clickable link despite not having an HREF property.
		
		.Parameter Link
		The link to click.
	#>
	
	param (
		[Parameter( Mandatory, ValueFromPipeLine )]
		[ValidateNotNullOrEmpty()]
		[OpenQA.Selenium.Remote.RemoteWebElement] $Link
	)
	
	Begin {
		$Fn = (Get-Variable MyInvocation -Scope 0).Value.MyCommand.Name
	}
	
	Process {
		try {
			# Make sure link is clickable before trying it.
			Write-DebugLog "${Fn}: Try to click link:"
			Dump-ElementInfo $Link -WebInfo
			$ClickableLink = WaitFor-ElementToBeClickable -WebElement $Link
			Invoke-SeClick $ClickableLink
			# Now wait for browser to process the click and load the next page
			$ClickableLink.WrappedDriver | Invoke-Wait
		} 
		catch {
			write-log -fore yellow "${Fn}: Problem clicking link, '$($Link.href)'"
			Handle-Exception $_
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
		[ValidateNotNullOrEmpty()]
		$ClickMe,
		
		[Parameter( ValueFromPipeLine, Position=2 )]
		[int] $SleepTime = 1
	)
	
	Begin {
		$Fn = (Get-Variable MyInvocation -Scope 0).Value.MyCommand.Name
	}
	
	Process {
		try {
			Invoke-SeClick $ClickMe
			# Now wait for a bit
			Start-Sleep -Seconds $SleepTime
		}
		catch {
			write-log -fore yellow "${Fn}: Some problem clicking object."
			Handle-Exception $_
		}
	}
	
	End {
	}
}

function Dump-ElementInfo {
	<#
		.Description
		Print information into the Debug log about the supplied web element, for debugging purposes.
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
	
	# Check if object is empty; if so, log a message but don't bother trying to dump info.
	if ( $WebElement ) {
		# WebElement info section.
		if ( $WebInfo -or $All ) {
			$Output = $WebElement | out-string
			Write-DebugLog -fore gray $Output
		}
		
		# Member info section.
		if ( $MemberInfo -or $All ) {
			$Output = $WebElement | get-member | format-table -auto | out-string
			Write-DebugLog -fore gray $Output
		}
	} else {
		Write-DebugLog -fore yellow "Warning: Attempted to dump info from a null element."
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
	
	param (
		[Parameter( Mandatory, ValueFromPipeLine )]
		$Input
	)
	
	Begin {
		$InchValues = "in", "inch", "inches"
		$FootValues = "ft", "foot", "feet"
		$YardValues = "yd", "yard", "yards", "yds"
		$PoundValues = "lb", "pound", "pounds", "lbs"
		$OunceValues = "oz", "ounce", "ounces", "ozs"
	}
	
	Process {
		switch ( $Input ) {
			# DSF is not consistent with abbreviations or plurals of units!
			{ $_ -in $InchValues }	{ $Output = "Inches"}
			{ $_ -in $FootValues }	{ $Output = "Feet"}
			{ $_ -in $YardValues }	{ $Output = "Yard"}
			{ $_ -in $PoundValues }	{ $Output = "lb"}
			{ $_ -in $OunceValues }	{ $Output = "oz"}
			default					{ $Output = "undefined"}
		}
		
		$Output
	}
	
	End {}
}

function Get-Control {
	<#
		.Synopsis
		Find a control, such as a checkbox, in the provided web driver or element.
		
		.Description
		Given a web driver or web element, and search criteria such as ID, find the control element.
		Return the control as a web element.
		
		.Parameter WebDriver
		Web driver to search.  May be passed via pipeline.
		
		.Parameter WebElement
		Web element to search.  May be passed via pipeline.
		
		.Parameter Type
		Type of control, such as "checkbox" or "radiobutton".
		
		.Parameter ID
		The ID of the found element to search.
		
		.Parameter Timeout
		Number of seconds after which to give up waiting.  Default is 10.
		
		.Example
		Get a checkbox whose ID is "GiftWrap".
		
		$GiftWrapChk = $Browser | Get-Control -Type Checkbox -ID "GiftWrap"
		
		.Example
		Get a picklist whose ID is "SelectProductA".
		
		$Picklist = $Browser | Get-Control -Type List -ID "SelectProductA"
		
		.Example
		Get an input text box within an existing element.
		
		$PriceField = $BigTable | Get-Control -Type Text -ID "UnitPrice"
	#>
	
	<#	
		TODO: Throw unique exception if wait times out, so the calling function has a chance to handle
			the lack of link, such as forcing a page reload or navigating to some known location.
			
		Issue 27:  Function needs to handle the case where neither WebDriver nor WebElement is supplied,
			because the coder messed up, instead of going ahead and returning $null.
	#>
	
	Param(
		[Parameter( Position=1, ValueFromPipeLine, ParameterSetName="ID" )]
		[Parameter( Position=1, ValueFromPipeLine, ParameterSetName="Name" )]
		[OpenQA.Selenium.Remote.RemoteWebDriver] $WebDriver,

		[Parameter( Position=1, ValueFromPipeLine, ParameterSetName="ID" )]
		[Parameter( Position=1, ValueFromPipeLine, ParameterSetName="Name" )]
		[OpenQA.Selenium.Remote.RemoteWebElement] $WebElement,

		[Parameter( Mandatory, Position=2, ParameterSetName="ID" )]
		[Parameter( Mandatory, Position=2, ParameterSetName="Name" )]
		[ValidateNotNullOrEmpty()]
		[string] $Type,
		
		[Parameter( Mandatory, Position=3, ParameterSetName="ID" )]
		[ValidateNotNullOrEmpty()]
		[string] $ID,
		
		[Parameter( Mandatory, Position=3, ParameterSetName="Name" )]
		[ValidateNotNullOrEmpty()]
		[string] $Name,
		
		[int] $Timeout = 10
	)

	if ( $WebElement ) {
		$WebDriver = $WebElement.WrappedDriver
	}
	
	$Fn = (Get-Variable MyInvocation -Scope 0).Value.MyCommand.Name
	Write-DebugLog -fore gray "${Fn}: Find element with ID matching `'$ID`'"
	<#
		Typical controls would include:
			Radio button, input type="radio" id="whatever"
			Checkbox, input type="checkbox" id="whatever"
			List, select  id="whatever"
	#>
	
	<#
		Typical use of Wait-Link:
			$Picklist = $BrowserObject | Wait-Link -TagName "select" -Property "id" -Pattern "ctl00_ctl00_C_M_ctl00_W_ctl01__Rank_DropDownListRank"
	#>

	$TypeTag = switch ( $Type ) {
		"Button"			{ "input" ; continue }
		"Checkbox"			{ "input" ; continue }
		"File"				{ "input" ; continue }
		"List"				{ "select" ; continue }
		"RadioButton"		{ "input" ; continue }
		"RichText"			{ "input" ; continue }
		"Text"				{ "input" ; continue }
		"TextArea"			{ "textarea" ; continue }
		default	{
			throw "${Fn}: Unexpected control type '$Type'"
		}
	}
	
	if ( $Name ) {
		$SearchBy = $Name
	} else {
		$SearchBy = $ID
	}
	
	# Debug:  Output the TypeTag
	Write-DebugLog -fore gray "${Fn}: TypeTag '$TypeTag' trying to match Type '$Type'"
		
	try {
		# Create a Stopwatch object to keep track of time
		$Stopwatch = New-Object System.Diagnostics.Stopwatch
		$Stopwatch.Start()
		
		$TimedOut = $false
		
		do {
			# Check if too much time has elapsed; break out if so.
			if ( $Stopwatch.Elapsed.Seconds -ge $Timeout ) {
				$TimedOut = $true
				Write-DebugLog -fore gray "${Fn}: Timed out after $Timeout seconds waiting for control element."
				break
			}
			
			# May also want to handle NoSuchElementException or similar...
			switch ( $PSCmdlet.ParameterSetName ) {
				"ID"	{
					# For most DSF web controls, use the tag to find them; double-check control type for sanity.
					$result = $WebDriver.FindElementsByID( $SearchBy ) | Where-Object { $_.TagName -eq $TypeTag }
				}
				"Name"	{
					$result = $WebDriver.FindElementsByName( $SearchBy ) | Where-Object { $_.TagName -eq $TypeTag }
				}
			}
		}
		until ( $result )
		
		if ( $TimedOut ) {
			throw "Timed out while waiting for '$TypeTag' element matching '$SearchBy'"
		}
	}
	catch {
		# Handle any exceptions thrown within function, or send them to main exception handler.
		switch -wildcard ( $_.Exception.Message ) {
			"Timed out while waiting for*"	{
				write-log -fore yellow "${Fn}: Timeout reached. Either element wasn't found or browser took more than $Timeout seconds to return it."
			}
			default	{
				Handle-Exception $_
			}
		}
	}
	
	# We made it this far, so presumably we got what we need.  Return it.
	$result
}

Function Handle-Exception {
	# Custom error handling for this script
	param (
		$Exc
	)

	if ( $Exc.Exception.WasThrownFromThrowStatement ) {
		# Throw statement means we did it on purpose.  Examine error code and
		#  print appropriate message.

		write-log -fore mag "Caught custom exception:" 
		write-log -fore mag $( $Exc | Format-List * -force | out-string )
		
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
			}
			"Timed out while waiting*" {
				Write-Log -fore mag $exMsg 
				Write-Log -fore red "Time limit exceeded while waiting for web form element."
				Write-Log -fore red "Execution cannot continue."
			}

			default { 
				write-log -fore red "Unhandled exception:" 
				write-log ( $Exc.Exception | Format-List | Out-String )
			}
		}
	
	} else {
		write-log -fore mag "Caught standard exception:"
		write-log -fore mag $_.Exception.ErrorRecord.Exception
		write-log -fore gray $( $_ | fl * -force | out-string )
	}

}

function Invoke-Login {
	<#
		.Synopsis
		Log into web site.  Return the URL of the page that loads after login.
		
		.Parameter WebDriver
		Browser object to control.
		
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
		[Parameter( Mandatory )]
		[OpenQA.Selenium.Remote.RemoteWebDriver] $WebDriver,
		
		[string] $SiteURL,
		
		[string] $UserName,
		
		[string] $Password
	)
	
	$Fn = (Get-Variable MyInvocation -Scope 0).Value.MyCommand.Name
	
	try {
		# What do we return?
		$ReturnLink = $null
		
		# Web control names
		$LoginButtonSnip = 'ctl00_ctl00_C_W__loginWP__myLogin_Login'
		$UserFieldSnip = 'ctl00_ctl00_C_W__loginWP__myLogin__userNameTB'
		$PassFieldSnip = 'ctl00_ctl00_C_W__loginWP__myLogin__passwordTB'

		# Navigate to page and attempt to sign in.
		write-log -fore cyan "${Fn}: Loading site: $SiteURL"
		Enter-SeUrl $SiteURL -Driver $WebDriver
		
		##### Log in
		# Get input fields
		$UserField = Find-SeElement -Driver $WebDriver -ID $UserFieldSnip
		$PassField = Find-SeElement -Driver $WebDriver -ID $PassFieldSnip
		
		# Fill in values from stored credential
		Set-TextField $UserField $UserName
		Set-TextField $PassField $Password
		# Find the Login button and click it
		$LoginButton = Find-SeElement -Driver $WebDriver -ID $LoginButtonSnip

		write-log -fore cyan "${Fn}: Logging in..."
		#Click-Link $LoginButton
		# Sleep after clicking, because we don't yet know how to reliably detect when storefront page is complete.
		#	Issue 2.
		#Click-Wait $LoginButton 30
		$LoginButton | Click-Link
		
		$ReturnLink = $WebDriver.Url

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
	
	$Fn = (Get-Variable MyInvocation -Scope 0).Value.MyCommand.Name

	$scrGetReadyState = "return document.readyState"

	<#	Wait until the main document returns "complete"
	#>
	$DocState = $BrowserObject.ExecuteScript($scrGetReadyState)
	while ( $DocState -notlike "complete" ) {
		# Not ready yet, so wait 1 second and check again.
		Write-DebugLog "${Fn}: Waiting for page load to complete..."
		Start-Sleep -Seconds 1
		$DocState = $BrowserObject.ExecuteScript($scrGetReadyState)
	}
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
			
			Write-DebugLog "Select-FromList: Try to select '$Item' from list containing `n'$($ListObject.Text)'"
			
			# Create a Selenium Select object to find what we want.
			$Selector = New-Object -TypeName OpenQA.Selenium.Support.UI.SelectElement( $ListObject )
			# Have it select out target out of the list.
			# This will return NoSuchElementException if the option isn't found.
			$Selector.SelectByText( $Item )
			
			# Now verify the item is actually selected.
			# This is different from not being found; the item exists but wasn't selected for some reason.
			if ( $Selector.SelectedOption.Text -ne $Item ) {
				throw "Select-FromList: Couldn't select '$TargetItem' as requested!"
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
		Write-DebugLog "Set-CheckBox: Set checkbox with ID matching '$ChkBoxID' to $( $Off -eq $false )"
		
		# Check if we're turning it on or off.
		if ( $Off ) {
			# User requested uncheck.
			if ( $CheckBoxObject.Selected -eq $True ) {
				# Currently checked, so click to uncheck.
				$CheckBoxObject.Click()
				
				# Verify it's now unchecked.
				$RealCheckBox = WaitFor-ElementToBeClickable $ChkBoxDriver -ID $ChkBoxID 5
				if ( $RealCheckBox.Selected -ne $False ) {
					throw "Error:  Unable to uncheck $($RealCheckBox.GetAttribute('ID'))!"
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
					throw "Error:  Unable to check $($RealCheckBox.GetAttribute('ID'))!"
				}
			}
		}
	}
	
	end {}
}

function Set-RadioButton {
	<#
		.Description
		Given a radio button object, set the button to be clicked.  Considering that radio buttons
		cannot be turned off, and must be cleared by clicking a different button in the set,
		and that normal user action via the GUI cannot clear a button without clicking another one,
		this function has no provision for clearing a button.
		
		.Synopsis
		Given a radio button object, click the button.
		
		.Parameter RadioButton
		Selenium web element representing the radio button object.
	#>
	
	param (
		[Parameter( Mandatory, ValueFromPipeLine )]
		[OpenQA.Selenium.Remote.RemoteWebElement] $RadioButton
	)
	
	begin {}
	
	process {
		# Grab the ID attribute and save it for later.
		$RButnID = $RadioButton.GetAttribute("Id")
		$RButnDriver = $RadioButton.WrappedDriver
		Write-DebugLog "Set-RadioButton: Click radio button with ID matching '$RButnID'"
		
		# Check if button is already clicked.
		if ( $RadioButton.Selected -eq $False ) {
			# Currently not checked, so click to check.
			$RadioButton.Click()
			
			# Sometimes, changing button state causes the page to be refreshed.
			# If this happens, the element reference will become stale.
			
			# Verify it's now checked.
			$RealRButton = WaitFor-ElementToBeClickable $RButnDriver -ID $RButnID 5
			if ( $RealRButton.Selected -ne $True ) {
				throw "Error:  Unable to select $($RealRButton.GetAttribute('ID'))!"
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
		if ( $Editor ) {
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
		if ( $EditorIFrame ) {
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
		[OpenQA.Selenium.Remote.RemoteWebDriver] $SeObject,
		
		[Parameter( Mandatory )]
		[ValidateNotNullOrEmpty()]
		[string] $TagName,
		
		[Parameter( Mandatory )]
		[ValidateNotNullOrEmpty()]
		[string] $Property,
		
		[Parameter( Mandatory )]
		[ValidateNotNullOrEmpty()]
		[string] $Pattern,
		
		[int] $Timeout = 30
		
	)

	$Fn = (Get-Variable MyInvocation -Scope 0).Value.MyCommand.Name
	Write-DebugLog -fore gray "${Fn}: Wait for '$TagName' element with '$Property' matching '$Pattern'"
	
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
		$TryCount = 1
		$MaxTryCount = 3
		do {
			try {
				$result = $SeObject.FindElementsByTagName( $TagName ) | Where-Object { $_.GetProperty($Property) -like $Pattern }
			}
			catch [OpenQA.Selenium.StaleElementReferenceException] {
				Write-DebugLog "${Fn}: Stale element reference, try $TryCount of $MaxTryCount."
				Start-Sleep -Seconds 1
			}
			$TryCount++
		} until ( $TryCount -ge $MaxTryCount )
	}
	until ( $result )
	
	if ( $TimedOut ) {
		write-log -fore yellow "${Fn}: Timeout reached. Add better error handling to Wait-Link!"
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
		
		.Parameter Class
		Class name to search for.  Only use Class if you know it's unique!
		
		.Parameter ID
		ID tag to search by.
		
		.Parameter XPath
		XPath to search by.
		
		.Parameter LinkText
		Link text string to search by.
	#>
	
	param (
		[Parameter( Mandatory, Position=1, ParameterSetName="Class" )]
		[Parameter( Mandatory, Position=1, ParameterSetName="ID" )]
		[Parameter( Mandatory, Position=1, ParameterSetName="XPath" )]
		[Parameter( Mandatory, Position=1, ParameterSetName="LinkText" )]
		[OpenQA.Selenium.Remote.RemoteWebDriver] $WebDriver,

		[Parameter( Position=1, ParameterSetName="Element" )]
		[OpenQA.Selenium.Remote.RemoteWebElement] $WebElement,
		
		[Parameter( Position=2, ParameterSetName="Class" )]
		[string] $Class,
		
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
	
	$Fn = (Get-Variable MyInvocation -Scope 0).Value.MyCommand.Name

	if ( $WebElement -notlike $null ) {
		$WebDriver = $WebElement.WrappedDriver
	}
	
	# This object's job is to wait for something.
	$Waiter = New-Object OpenQA.Selenium.Support.UI.WebDriverWait($WebDriver, $TimeInSeconds)
	
	try {
		# Check which info we're given, then wait until it exists or times out.
		# If the waiter times out, it will throw an exception, which we'll handle.
		switch ( $PSCmdlet.ParameterSetName ) {
			"Class"		{
				$Locator = "Class: $Class"
				$Gotcha = $Waiter.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists( [OpenQA.Selenium.by]::ClassName($Class)))
			}
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
			"Element"	{
				$Locator = "Element: $($WebElement.TagName)"
				$Gotcha = $Waiter.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists( $WebElement ))
			}
		}
		
		# We got something, so return the element to caller.
		return $Gotcha
	}
	
	catch [OpenQA.Selenium.WebDriverTimeoutException] {
		# Timed out waiting for element.  What should we do here?
		# Nothing.  Caller must check if something was returned.
		write-log -fore yellow "${Fn}: Timed out waiting for '$Locator'"
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
	
	$Fn = (Get-Variable MyInvocation -Scope 0).Value.MyCommand.Name

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
		write-log -fore yellow "${Fn}: Timed out waiting for clickable '$Locator'"
	}
}

function Write-DebugLog {
	param ( [string] $Text = " " )
	
	if ( $DebugLogging ) { write-log -fore darkyellow $Text }
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
