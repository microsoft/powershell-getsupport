[CmdletBinding()]
param (
	[string]
	$Path
)

$ErrorActionPreference = 'Stop'
trap {
	Stop-FileWriter
	Write-Warning "Script failed: $_"
	throw $_
}

#region Functions

function Show-SaveFileDialog {
	[CmdletBinding()]
	param (
		[string]
		$InitialDirectory = '.',

		[string]
		$Filter = '*.*',
		
		$Filename
	)
	
	Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
	$saveFileDialog = [Windows.Forms.SaveFileDialog]::new()
	$saveFileDialog.FileName = $Filename
	$saveFileDialog.InitialDirectory = Resolve-Path -Path $InitialDirectory
	$saveFileDialog.Title = "Save File to Disk"
	$saveFileDialog.Filter = $Filter
	$saveFileDialog.ShowHelp = $True
	
	$result = $saveFileDialog.ShowDialog()
	if ($result -eq "OK") {
		$saveFileDialog.FileName
	}
}

function Resolve-ExportPath {
	[CmdletBinding()]
	param (
		[AllowEmptyString()]
		[string]
		$Path,

		[string]
		$DefaultFileName = 'default.txt',

		[string]
		$Filter = 'Text File (*.txt)|*.txt'
	)

	$parentPath = ''
	$fileName = $DefaultFileName
	if ($Path) {
		if ($Path -like "*.zip") {
			$fileName = Split-Path -Path $Path -Leaf
			$Path = Split-Path $Path
		}
		try { $resolved = Resolve-Path -Path $Path }
		catch { Write-Warning "Failed to resolve $Path : $_" }

		if ($resolved) {
			$parentPath = $resolved | Select-Object -First 1
		}
	}

	if (-not $parentPath) {
		$newPath = Show-SaveFileDialog -Filename $fileName -Filter $Filter
		if (-not $newPath) {
			throw "Export Path not resolvable!"
		}

		$fileName = Split-Path -Path $newPath -Leaf
		$parentPath = Split-Path -Path $newPath
	}

	Join-Path -Path $parentPath -ChildPath $fileName
}

function Start-FileWriter {
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]
		$Path
	)

	Stop-FileWriter
	$command = { Set-Content -Path $Path }.GetSteppablePipeline()
	$command.Begin($true)
	$script:_FileWriter = $command
}
function Stop-FileWriter {
	[CmdletBinding()]
	param (
		
	)

	if (-not $script:_FileWriter) {
		return
	}

	$script:_FileWriter.End()
	$script:_FileWriter = $null
}
function Write-File {
	[CmdletBinding()]
	param (
		[Parameter(ValueFromPipeline = $true)]
		[AllowEmptyString()]
		[string[]]
		$Content
	)

	begin {
		if (-not $script:_FileWriter) {
			throw "No file opened to write to! call Start-FileWriter to begin writing to file first!"
		}
	}
	process {
		foreach ($entry in $Content) {
			$script:_FileWriter.Process($entry)
		}
	}
}

function Write-DebugInformation {
	[CmdletBinding()]
	param (
		
	)

	Write-Header
	Write-History
	Write-Module
	Write-Assembly
	Write-PSSnapIn
	Write-ErrorText
	Write-ConsoleBuffer
}
function Write-Header {
	[CmdletBinding()]
	param (
		
	)

	$header = @'
################################################################################
Date: {0:yyyy-MM-dd} | PS Version: {1} | x64: {2}

Last Command:
{3}

Last Error:
{4}

'@
	$date = Get-Date
	$psversion = $PSVersionTable.PSVersion.ToString()
	$isx64 = [intptr]::Size -eq 8
	$lastCommand = (Get-History)[-1].CommandLine
	$lastError = $error[0] | Out-String
	$header -f $date, $psversion, $isx64, $lastCommand, $lastError | Write-File
}

function Write-History {
	param (

	)

	$header = @'
################################################################################
History:

'@
	$header | Write-File
	foreach ($entry in Get-History) {
		'  {0:D4}: {1}' -f $entry.Id, $entry.CommandLine | Write-File
	}

	Write-File -Content ''
}

function Write-Module {
	[CmdletBinding()]
	param (
		
	)

	Write-File @'
################################################################################
Modules:

'@
	foreach ($module in Get-Module) {
		'  {0}: {1} ({2})' -f $module.Name, $module.Version, $module.ModuleBase | Write-File
	}

	Write-File ''
}

function Write-Assembly {
	[CmdletBinding()]
	param (
		
	)

	Write-File -Content @'
################################################################################
Assemblies:

'@
	foreach ($assembly in [System.AppDomain]::CurrentDomain.GetAssemblies()) {
		'  {0}: {1}' -f $assembly.FullName, $assembly.Location | Write-File
	}

	Write-File ''
}

function Write-PSSnapIn {
	[CmdletBinding()]
	param (
		
	)

	if ($PSVersionTable.PSVersion.Major -gt 5) {
		return
	}

	Write-File -Content @'
################################################################################
PSSnapins:

'@
	foreach ($snapIn in Get-PSSnapIn) {
		'  {0} ({1} | {2})' -f $snapIn.Name, $snapIn.Version, $snapIn.AssemblyName | Write-File
	}

	Write-File -Content ''
}

function Write-ErrorText {
	[CmdletBinding()]
	param (
		
	)

	$header = @'
################################################################################
Errors:

'@
	If ($PSVersionTable.PSVersion.Major -ge 7) {
		Get-Error | Write-File
	}
	else {
		$error | Format-List -Force | Out-String | Write-File
	}

	Write-File -Content ''
}

function Write-ConsoleBuffer {
	[CmdletBinding()]
	param (
		
	)

	$header = @'
################################################################################
Console:

'@
	Write-File -Content $header
	Get-ConsoleText | Write-File
}


function Get-ShellBuffer {
	[CmdletBinding()]
	param ()
	
	if ($Host.Name -eq 'Windows PowerShell ISE Host') {
		return $psIse.CurrentPowerShellTab.ConsolePane.Text
	}
	
	try {
		# Define limits
		$rec = New-Object System.Management.Automation.Host.Rectangle
		$rec.Left = 0
		$rec.Right = $host.ui.rawui.BufferSize.Width - 1
		$rec.Top = 0
		$rec.Bottom = $host.ui.rawui.BufferSize.Height - 1
		
		# Load buffer
		$buffer = $host.ui.rawui.GetBufferContents($rec)
		
		# Convert Buffer to list of strings
		$int = 0
		$lines = @()
		while ($int -le $rec.Bottom) {
			$n = 0
			$line = ""
			while ($n -le $rec.Right) {
				$line += $buffer[$int, $n].Character
				$n++
			}
			$line = $line.TrimEnd()
			$lines += $line
			$int++
		}
		
		# Measure empty lines at the beginning
		$int = 0
		$temp = $lines[$int]
		while ($temp -eq "") { $int++; $temp = $lines[$int] }
		
		# Measure empty lines at the end
		$z = $rec.Bottom
		$temp = $lines[$z]
		while ($temp -eq "") { $z--; $temp = $lines[$z] }
		
		# Skip the line launching this very function
		$z--
		
		# Measure empty lines at the end (continued)
		$temp = $lines[$z]
		while ($temp -eq "") { $z--; $temp = $lines[$z] }
		
		# Cut results to the limit and return them
		return $lines[$int .. $z]
	}
	catch { }
}
#endregion Functions

$exportPath = Resolve-ExportPath -Path $Path -DefaultFileName "PowerShell_Support_$(Get-Date -Format yyyy-MM-dd).txt"
Start-FileWriter -Path $exportPath
Write-DebugInformation
Stop-FileWriter