[CmdletBinding()]
param (
	[string]
	$OutPath
)

#region Helper functions
function New-TempFolder {
	[CmdletBinding()]
	param (
		
	)

	$item = New-Item -Path $env:TEMP -Name "PSSupport-$(Get-Random)" -ItemType Directory -Force
	$item.FullName
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

function Export-DebugInformation {
	[CmdletBinding()]
	param (
		[string]
		$Path
	)

	Get-ShellBuffer | Set-Content -Path "$Path\console-buffer.txt"

	$osData = if ($IsLinux -or $IsMacOs) {
		[PSCustomObject]@{
			OSVersion       = [System.Environment]::OSVersion
			ProcessorCount  = [System.Environment]::ProcessorCount
			Is64Bit         = [System.Environment]::Is64BitOperatingSystem
			LogicalDrives   = [System.Environment]::GetLogicalDrives()
			SystemDirectory = [System.Environment]::SystemDirectory
		}
	}
	else {
		Get-CimInstance -ClassName Win32_OperatingSystem
	}
	$osData | Export-Clixml -Path "$Path\os.clixml"

	$cpuData = if ($IsLinux -and (Test-Path -Path /proc/cpuinfo)) {
		Get-Content -Raw -Path /proc/cpuinfo
	}
	else {
		Get-CimInstance -ClassName Win32_Processor
	}
	$cpuData | Export-Clixml -Path "$Path\cpu.clixml"

	$memoryData = if ($IsLinux -and (Test-Path -Path /proc/meminfo)) {
		Get-Content -Raw -Path /proc/meminfo
	}
	else {
		Get-CimInstance -ClassName Win32_PhysicalMemory
	}
	$memoryData | Export-Clixml -Path "$Path\memory.clixml"

	$PSVersionTable | Out-String | Set-Content -Path "$Path\psversion.txt"

	Get-History | Export-Clixml -Path "$Path\history.clixml"
	
	Get-Module | Export-Clixml -Path "$Path\modules.clixml"
	
	if (Get-Command -Name Get-PSSnapIn -ErrorAction SilentlyContinue) {
		Get-PSSnapin | Export-Clixml -Path "$Path\pssnapins.clixml"
	}

	[appdomain]::CurrentDomain.GetAssemblies() |
		Select-Object CodeBase, FullName, Location, ImageRuntimeVersion, GlobalAssemblyCache, IsDynamic |
			Export-Clixml -Path "$Path\assemblies.clixml"
	
	
	$error | Export-Clixml "$Path\errors.clixml"

	if ($PSVersionTable.PSVersion.Major -ge 7) {
		Get-Error | Set-Content -Path "$Path\errors.txt"
	}
}

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
		$DefaultFileName = 'default.zip'
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
		$newPath = Show-SaveFileDialog -Filename $fileName -Filter '*.zip'
		if (-not $newPath) {
			throw "Export Path not resolvable!"
		}

		$fileName = Split-Path -Path $newPath -Leaf
		$parentPath = Split-Path -Path $newPath
	}

	Join-Path -Path $parentPath -ChildPath $fileName
}
#endregion Helper functions

$tempPath = New-TempFolder
Export-DebugInformation -Path $tempPath
$resolvedOutPath = Resolve-ExportPath -Path $OutPath -DefaultFileName "powershell_support_$(Get-Date -Format 'yyyy-MM-dd').zip"
Compress-Archive -Path "$tempPath\*" -DestinationPath $resolvedOutPath
Remove-Item -Path $tempPath -Recurse -Force