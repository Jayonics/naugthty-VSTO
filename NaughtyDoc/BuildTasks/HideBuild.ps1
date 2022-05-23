#!/usr/bin/env pwsh
param($BuildDir, $DocumentName, $Verbose)

[String[]]$CommonOfficeExtensions = @(".docm", ".docx", ".doc")
[String[]]$HiddenAttributes = @("Hidden", "System")

Function HideFiles() {
	$LibraryFiles = (Get-ChildItem -Force -Recurse -Path:$BuildDir -Exclude:$($CommonOfficeExtensions.ForEach({$_ = "*$_"})) -Verbose)

	If ($Verbose -ne $null) {
		$LibraryFiles | % { 
			$_.Attributes = @HiddenAttributes
			Select-Object -Properties: Name, Attributes | Format-Table
		}
	} Else {
		$_.Attributes = @HiddenAttributes
	}
}