function Reset-Printers{

param ($TempDirectory, [switch]$Force,[switch]$light,[switch]$full,[switch]$quiet=$false)

#Updated 2018/01/09
$ScriptVersion = "8"


#Windows 7 and Windows 2008 R2 registry
(new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/DrDrrae/Powershell/master/Printers/registry/Windows7%26Server2008R2.ps1') | Invoke-Expression

#Windows 2012 R2 registry
(new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/DrDrrae/Powershell/master/Printers/registry/AMD_64_6_3_9600.ps1') | Invoke-Expression

#Windows 10 1607
#. ".\registry\x86_10_0_14393.ps1"
#. ".\registry\AMD_64_10_0_14393.ps1"

#Windows 10 1703
#. ".\registry\x86_10_0_15063.ps1"
(new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/DrDrrae/Powershell/master/Printers/registry/AMD_64_10_0_15063.ps1') | Invoke-Expression

#Windows 10 1709
#. ".\registry\x86_10_0_16299.ps1"
(new-object Net.WebClient).DownloadString('https://raw.githubusercontent.com/DrDrrae/Powershell/master/Printers/registry/AMD_64_10_0_16299.ps1') | Invoke-Expression


function Writeto-ExecutionLog ([Parameter(Mandatory=$true, Position=0, ValueFromPipeline=$true)][String[]][AllowEmptyString()]$InputObject)
#([switch]$quiet, [string]$inputObject)
{
	if (-not $quiet)
	{
		write-host $InputObject
		$InputObject | Out-File $ExecutionLog -Append
	}
	else
	{
		$InputObject | Out-File $ExecutionLog -Append
	}
}

function previousexecution()
{
	if (Test-Path $ExecutionDirectory)
	{
		return $true
	}
	else
	{
		return $false
	}
}

#deploy registration files for correct OS version and machine architecture.
function deployregfiles()
{
	$arrRegFiles = @()
	$os = Get-WmiObject win32_operatingsystem
	if ($os.ProductType -gt 1)
	{
		$regvarsubstring = ($env:PROCESSOR_ARCHITECTURE + "_" + ($os.Version).replace(".","_") + "_server")
	}
	else
	{
		$regvarsubstring = ($env:PROCESSOR_ARCHITECTURE + "_" + ($os.Version).replace(".","_") + "_client")
	}

	if (-not (Get-Variable -erroraction SilentlyContinue -Name ("PRINT_" + $regvarsubstring)).value)
	{
		"[error]:  support for the detected OS or architecture is not available in this script - $regvarsubstring" | Writeto-ExecutionLog
		return $null
	}
	else
	{
		$outputRegFile = "$tempdirectory\defaultprintkey.reg"
		(Get-Variable -Name ("PRINT_" + $regvarsubstring)).value | Out-File $outputRegFile
		if (-not (Test-Path $outputRegFile))
		{
			"[error]:  not able to successfully deploy reg file $outputregfile" | Writeto-ExecutionLog
			return $null
		}
		else
		{
			$arrRegFiles += $outputRegFile
		}
	}

	if (-not (Get-Variable -erroraction SilentlyContinue -Name ("SERVICES_" + $regvarsubstring)).value)
	{
		"[error]:  support for the detected OS or architecture is not available in this script - $regvarsubstring" | Writeto-ExecutionLog
		return $null
	}
	else
	{
		$outputRegFile = "$tempdirectory\defaultserviceskey.reg"
		(Get-Variable -Name ("SERVICES_" + $regvarsubstring)).value | Out-File $outputRegFile
		if (-not (Test-Path $outputRegFile))
		{
			"[error]:  not able to successfully deploy reg file $outputregfile" | Writeto-ExecutionLog
			return $null
		}
		else
		{
			$arrRegFiles += $outputRegFile
		}
	}

	if (-not (Get-Variable -erroraction SilentlyContinue -Name ("MONITORS_" + $regvarsubstring)).value)
	{
		"[error]:  support for the detected OS or architecture is not available in this script - $regvarsubstring" | Writeto-ExecutionLog
		return $null
	}
	else
	{
		$outputRegFile = "$tempdirectory\defaultmonitorskey.reg"
		(Get-Variable -Name ("MONITORS_" + $regvarsubstring)).value | Out-File $outputRegFile
		if (-not (Test-Path $outputRegFile))
		{
			"[error]:  not able to successfully deploy reg file $outputregfile" | Writeto-ExecutionLog
			return $null
		}
		else
		{
			$arrRegFiles += $outputRegFile
		}
	}

	if (-not (Get-Variable -erroraction SilentlyContinue -Name ("WINPRINT_" + $regvarsubstring)).value)
	{
		"[error]:  support for the detected OS or architecture is not available in this script - $regvarsubstring" | Writeto-ExecutionLog
		return $null
	}
	else
	{
		$outputRegFile = "$tempdirectory\defaultwinprintkey.reg"
		(Get-Variable -Name ("WINPRINT_" + $regvarsubstring)).value | Out-File $outputRegFile
		if (-not (Test-Path $outputRegFile))
		{
			"[error]:  not able to successfully deploy reg file $outputregfile" | Writeto-ExecutionLog
			return $null
		}
		else
		{
			$arrRegFiles += $outputRegFile
		}
	}
	return $arrRegFiles
}

Function PsCreate($Process, $Arguments = "", $WorkingDirectory = $null)
{
	
	"[info]: PsCreate($Process, $Arguments) called." 
	
	$Error.Clear()
	$processStartInfo  = new-object System.Diagnostics.ProcessStartInfo
	$processStartInfo.fileName = $Process
	if ($Arguments.Length -ne 0) { $processStartInfo.Arguments = $Arguments }
	if ($WorkingDirectory -eq $null) {$processStartInfo.WorkingDirectory = (Get-Location).Path}
	$processStartInfo.UseShellExecute = $false
	$processStartInfo.RedirectStandardOutput = $true
	$processStartInfo.REdirectStandardError = $true
	
	$process = [System.Diagnostics.Process]::Start($processStartInfo)

	if ($Error.Count -gt 0)
	{
		$errorMessage = $Error[0].Exception.Message
		$errorCode = $Error[0].Exception.ErrorRecord.FullyQualifiedErrorId
		$PositionMessage = $Error[0].InvocationInfo.PositionMessage
		"[error]: " + $errorCode + " on: " + $line + ": $errorMessage" | Writeto-ExecutionLog
		$Error.Clear()
	}

	Return $process
}

function Get-ServiceDependencyList($ServiceName)
{	
	$arrServiceList = @()
	$arrServiceList += $ServiceName
	foreach ($ServiceName in $arrServiceList)
	{
		$Dependents = (Get-Service $ServiceName).DependentServices
		foreach ($ServiceName in $Dependents)
		{	if (-not ($arrServiceList -contains $ServiceName.name))
			{
				$arrServiceList += $ServiceName.name
				Get-ServiceDependencyList($ServiceName.name)
			}
		}
	}
	return $arrServiceList 
}

#################################################################################################################################################################################################################
#"light" mode.  Corrects common printing issues.
#1.  Stop spooler service and dependents.
#2.  Do not backup spooler service configuration:  hklm\system\currentcontrolset\services  - warn user -printbrm/PMC
#3.  Do not Backup print configuration:  hklm\system\currentcontrolset\control\print 
#4.  Do not backup files in %windir%\system32\spool\printers\ directory - warn user - back up manually
#5.  Import default print monitors (USB, Local Port, TCP/IP, WSD, and Microsoft Shared Fax registry configuration)
#6.  Import spooler service default configuration
#7.  Import print processor configuration for correct machine architecture:  hklm\system\currentcontrolset\control\print\environments\(machine architecture)\winprint
#8.  Detect if spoolsv.exe, spoolss.dll, localspl.dll, or win32spl.dll are not present in %windir%\system32, write error to log.
#9.  Copy NTPrint.inf from %windir%\driverstore if not present in %windir%\inf.  If (also) not present in driverstore, write error to log.
#10. Test the path for hklm\system\currentcontrolset\control\print\printers\.DefaultSpoolDirectory. If it is not a valid path, write error to log.  correct this condition by setting it to the default path.
#11. Delete HKLM:/Software/Microsoft/Windows NT/CurrentVersion/Print/Providers
#12. Delete all files from the spooler directory 
#13. Start the spooler service and dependents
#################################################################################################################################################################################################################

function lightmode()
{
	$arrServiceList = get-ServiceDependencyList ("spooler") | select -Unique
	$arrStartedServices = @()
	foreach ($service in $arrServiceList)
	{
		$CurrentService = Get-Service -Name $service
		
		if ($CurrentService.status -eq "Running")
		{	
			$CurrentServiceName = $CurrentService.Name
			$arrStartedServices += $CurrentServiceName
			$SvcStopTimeout = "300"
			stop-service $CurrentServiceName -Force
			"[Info]:  Waiting up to 5 minutes to stop service $CurrentServiceName" | Writeto-ExecutionLog
			for ($i = 1; $i -le $SvcStopTimeout; $i++)
			{	
				if ($CurrentService.Status -ne "stopped")
				{
					Start-Sleep 1
					$CurrentService = Get-Service -Name $CurrentServiceName
				}
				else
				{
					Start-Sleep 5
					break
				}
			}
			if ($CurrentService.Status -ne "stopped")
			{
				Write-Host "[error]:  Could not stop service $CurrentServiceName" | Writeto-ExecutionLog
			}
		}
	}
	Stop-Process -name "printisolationhost" -force -erroraction silentlycontinue
	Start-Sleep 5

	"[info]:  Spooler service and dependents stopped" | Writeto-ExecutionLog
	
	$spooldir = (get-itemproperty "HKLM:/System/CurrentControlSet/Control/Print/Printers").DefaultSpoolDirectory
	
	if (-not (test-path $spooldir) -or ($spooldir -match "\\\\"))
	{
		"[info]:  Invalid spooler path $spooldir detected.  Resetting to default spooler path $env:windir\system32\spool\PRINTERS" | Writeto-ExecutionLog
		Set-ItemProperty -Name "DefaultSpoolDirectory" -Path "HKLM:/System/CurrentControlSet/Control/Print/Printers" -Value "$env:windir\system32\spool\PRINTERS"
	}
	
	if (Test-Path "HKLM:/Software/Microsoft/Windows NT/CurrentVersion/Print/Providers")
	{
		Remove-Item -Path "HKLM:/Software/Microsoft/Windows NT/CurrentVersion/Print/Providers" -Recurse -Force | Out-Null
	}

	#"[info]:  Backup directory: $ExecutionDirectory" | Writeto-ExecutionLog
	#"[info]:  Backing up hklm\system\currentcontrolset\services\spooler to spooler.reg" | Writeto-ExecutionLog
	#$backup_Spooler_filename = (Join-Path $ExecutionDirectory "spooler.reg")
	#PsCreate -process "reg" -arguments "export hklm\system\currentcontrolset\services\spooler $backup_Spooler_filename" | Out-Null

	#"[info]:  Backing up hklm\system\currentcontrolset\control\print to print.reg" | Writeto-ExecutionLog
	#$backup_printkey_filename = (Join-Path $ExecutionDirectory "print.reg")
	#PsCreate -process "reg" -arguments "export hklm\system\currentcontrolset\control\print $backup_printkey_filename" | Out-Null
	#$arrSplFiles = Get-ChildItem "$SpoolDir\*.*" 
	
	#if ($arrSplFiles -ne $null)
	#{
	#	foreach ($file in $arrSplFiles)
	#	{
	#		"[info]:  copying $file to $ExecutionDirectory" | Writeto-ExecutionLog
	#		Copy-Item -Path $file -Destination $ExecutionDirectory
	#	}
	#}
	$RegFilesDeployed = deployregfiles
	if ($RegFilesDeployed)
	{
		"[info]:  Importing default registry information" | Writeto-ExecutionLog
		foreach ($_ in $RegFilesDeployed)
		{
			if (-not ($_ -match "defaultprintkey.reg"))
			{
				PsCreate -process "reg" -arguments "import $_" | Out-Null
			}
		}
	}
	else
	{
		"[error]:  unable to import default registry information" | Writeto-ExecutionLog
	}
	
	Wait-Process -Name "reg" -ErrorAction SilentlyContinue
	
	"[info]:  removing temporary files"| Writeto-ExecutionLog
	foreach ($file in $RegFilesDeployed)
	{
		Remove-Item $file -ErrorAction SilentlyContinue
	}
	
	if (-not (Test-Path "$env:windir\system32\spoolsv.exe"))
	{
		"[error]:  $env:windir\system32\spoolsv.exe not detected.  Please run SFC /SCANNOW to correct this issue." | Writeto-ExecutionLog
	}
	if (-not (Test-Path "$env:windir\system32\spoolss.dll"))
	{
		"[error]:  $env:windir\system32\spoolss.dll not detected.  Please run SFC /SCANNOW to correct this issue." | Writeto-ExecutionLog
	}
	if (-not (Test-Path "$env:windir\system32\localspl.dll"))
	{
		"[error]:  $env:windir\system32\localspl.dll not detected.  Please run SFC /SCANNOW to correct this issue." | Writeto-ExecutionLog
	}
	if (-not (Test-Path "$env:windir\system32\win32spl.dll"))
	{
		"[error]:  $env:windir\system32\win32spl.dll not detected.  Please run SFC /SCANNOW to correct this issue." | Writeto-ExecutionLog
	}
	if (-not (Test-Path "$env:windir\inf\ntprint.inf"))
	{
		"[error]:  $env:windir\inf\ntprint.inf not detected.  Attempting to copy from $env:windir\System32\DriverStore\FileRepository" | Writeto-ExecutionLog
		$ntprintinfpath = join-path "$env:windir\System32\DriverStore\FileRepository" (Get-ChildItem "$env:windir\System32\DriverStore\FileRepository" | where {$_.name -match "ntprint.inf"}) 
		if (-not (Test-Path ($ntprintinfpath + "\ntprint.inf")))
		{
			"[error]:  ntprint.inf not detected in $env:windir\System32\DriverStore\FileRepository." | Writeto-ExecutionLog
		}
		else
		{
			Copy-Item -Path (join-path $ntprintinfpath "ntprint.inf") -Destination "$env:windir\inf" -OutVariable $cmdoutput
			if (-not (Test-Path ($ntprintinfpath + "\ntprint.inf")))
			{
				"[error]:  failed to copy ntprint.inf from FileRepository." | Writeto-ExecutionLog
			}
			else 
			{
				"[info]:  ntprint.inf successfully copied to $env:windir\inf." | Writeto-ExecutionLog
			}
		}
	}	
	if ($arrSplFiles -ne $null)
	{
		"[info]:  deleting pending print jobs" | Writeto-ExecutionLog
		foreach ($file in $arrSplFiles)
		{
			#"[info]:  deleting $file" | Writeto-ExecutionLog
			Remove-Item $file
		}
	}
	
	foreach ($service in $arrStartedServices)
	{
		$CurrentService = Get-Service -Name $service
		$CurrentServiceName = $CurrentService.Name
		if ($CurrentService.status -eq "stopped")
		{
			$SvcStartTimeout = "300"
			start-service $CurrentServiceName 
			Write-Host "[Info]:  Waiting up to 5 minutes to start service $CurrentServiceName" | Writeto-ExecutionLog
			for ($i = 1; $i -le $SvcStopTimeout; $i++)
			{	
				if ($CurrentService.Status -ne "running")
				{
					Start-Sleep 1
					$CurrentService = Get-Service -Name $CurrentServiceName
				}
				else
				{
					"[info]:  $CurrentServiceName is running" | Writeto-ExecutionLog
					break
				}
			}
			if ($CurrentService.Status -ne "running")
			{
				Write-Host "[error]:  Could not start service $CurrentServiceName" | Writeto-ExecutionLog
			}
		}
		elseif ($CurrentService.Status -eq "running")
		{
			"[info]:  $CurrentServiceName is running" | Writeto-ExecutionLog
		}
		else
		{
			"[warning]:  $CurrentServiceName is not running" | Writeto-ExecutionLog
			$CurrentService.Status | Writeto-ExecutionLog
		}
	}
	
	"[info]:  Finished light mode" | Writeto-ExecutionLog
}


$TempDirectory = "$env:windir\temp"
$ExecutionDirectory = "$Env:windir\printreset\"
$ExecutionLogName = ("printreset_executionlog_" + (Get-Date -Format o).replace(":",".") + ".txt")
$ExecutionLog = (Join-Path $ExecutionDirectory $ExecutionLogName)

#################################################################
#"full" mode.  Resets spooler service and print key to defaults.
#################################################################
#1.   Stop spooler service and dependents.
#2.   Do not Backup spooler service configuration:  hklm\system\currentcontrolset\services  
#3.   Do not Backup print configuration:  hklm\system\currentcontrolset\control\print 
#4.   Do not Backup Lanmanserver/Shares key
#5.   Do not Backup HKCU\Printers\Connections
#6.   Import default print key for correct machine architecture
#7.   Import spooler service default configuration
#8.   Detect if spoolsv.exe, spoolss.dll, localspl.dll, or win32spl.dll are not present in %windir%\system32, write error to log.
#9.   Copy NTPrint.inf from %windir%\driverstore if not present in %windir%\inf.  If (also) not present in driverstore, write error to log.
#10. Do not backup files in %windir%\system32\spool\printers\ directory
#11. Delete all files from %windir%\system32\spool\printers\ directory
#12. Delete all printer shares 
#13. Delete HKCU\Printers\Connections.  This key will be recreated the next time the user runs Control Panel\Hardware and Sound\Devices and Printers
#14. Delete HKLM:/Software/Microsoft/Windows NT/CurrentVersion/Print/Providers
#15. Delete non-default files from “%windir%\System32\spool\drivers\<arch>", “%windir%\System32\spool\drivers\<arch>\3”, and "%windir%\System32\spool\prtprocs\<arch>". Ignore subdirectories. 
#16. Start the spooler service and dependents.  It is expected that some third-party dependent services may fail because of missing drivers/configuration information.  These should be reinstalled.

function fullmode()
{
	$arrServiceList = get-ServiceDependencyList ("spooler") | select -Unique
	$arrStartedServices = @()
	foreach ($service in $arrServiceList)
	{
		$CurrentService = Get-Service -Name $service
		
		if ($CurrentService.status -eq "Running")
		{	
			$CurrentServiceName = $CurrentService.Name
			$arrStartedServices += $CurrentServiceName
			$SvcStopTimeout = "300"
			stop-service $CurrentServiceName -Force
			"[Info]:  Waiting up to 5 minutes to stop service $CurrentServiceName" | Writeto-ExecutionLog
			for ($i = 1; $i -le $SvcStopTimeout; $i++)
			{	
				if ($CurrentService.Status -ne "stopped")
				{
					Start-Sleep 1
					$CurrentService = Get-Service -Name $CurrentServiceName
				}
				else
				{
					Start-Sleep 5
					break
				}
			}
			if ($CurrentService.Status -ne "stopped")
			{
				Write-Host "[error]:  Could not stop service $CurrentServiceName" | Writeto-ExecutionLog
			}
		}
	}
	Stop-Process -name "printisolationhost" -force -erroraction silentlycontinue
	Start-Sleep 5

	"[info]:  Spooler service and dependents stopped" | Writeto-ExecutionLog
	
	#$spooldir = (get-itemproperty "HKLM:/System/CurrentControlSet/Control/Print/Printers").DefaultSpoolDirectory
	
	#if (-not (test-path $spooldir) -or ($spooldir -match "\\\\"))
	#{
	#	"Invalid spooler path $spooldir detected.  Resetting to default spooler path $env:windir\system32\spool\PRINTERS" | Writeto-ExecutionLog
	#	Set-ItemProperty -Name "DefaultSpoolDirectory" -Path "HKLM:/System/CurrentControlSet/Control/Print/Printers" -Value "$env:windir\system32\spool\PRINTERS"
	#	$spooldir = (get-itemproperty "HKLM:/System/CurrentControlSet/Control/Print/Printers").DefaultSpoolDirectory
	#}
	
	#"[info]:  Backup directory: $ExecutionDirectory" | Writeto-ExecutionLog
	#"[info]:  Backing up hklm\system\currentcontrolset\services\spooler to spooler.reg" | Writeto-ExecutionLog
	#$backup_Spooler_filename = (Join-Path $ExecutionDirectory "spooler.reg")
	#PsCreate -process "reg" -arguments "export hklm\system\currentcontrolset\services\spooler $backup_Spooler_filename" | Out-Null

	#"[info]:  Backing up hklm\system\currentcontrolset\control\print to print.reg" | Writeto-ExecutionLog
	#$backup_printkey_filename = (Join-Path $ExecutionDirectory "print.reg")
	#PsCreate -process "reg" -arguments "export hklm\system\currentcontrolset\control\print $backup_printkey_filename" | Out-Null
	#$arrSplFiles = Get-ChildItem "$SpoolDir\*.*" 
	
	#if ($arrSplFiles -ne $null)
	#{
	#	foreach ($file in $arrSplFiles)
	#	{
	#		"[info]:  copying $file to $ExecutionDirectory" | Writeto-ExecutionLog
	#		Copy-Item -Path $file -Destination $ExecutionDirectory
	#	}
	#}
	$RegFilesDeployed = deployregfiles
	if ($RegFilesDeployed)
	{
		"[info]:  Importing default registry information" | Writeto-ExecutionLog
		foreach ($_ in $RegFilesDeployed)
		{
			if ($_ -match "defaultprintkey.reg")
			{
				if (Test-Path -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Print")
				{
					Remove-Item -path "HKLM:\SYSTEM\CurrentControlSet\Control\Print" -force -Recurse -ErrorAction SilentlyContinue
					PsCreate -process "reg" -arguments "import $_" | Out-Null
				}	
			}
			if ($_ -match "defaultserviceskey.reg")
			{
				if (Test-Path -Path "HKLM:\SYSTEM\CurrentControlSet\Services\Spooler")
				{
					Remove-Item -path "HKLM:\SYSTEM\CurrentControlSet\Services\Spooler" -force -Recurse -ErrorAction SilentlyContinue
					PsCreate -process "reg" -arguments "import $_" | Out-Null
				}	
			}
		}
	}
	else
	{
		"[error]:  unable to import default registry information" | Writeto-ExecutionLog
	}
	
	Wait-Process -Name "reg" -ErrorAction SilentlyContinue
	
	if (Test-Path "HKLM:/Software/Microsoft/Windows NT/CurrentVersion/Print/Providers")
	{
		Remove-Item -Path "HKLM:/Software/Microsoft/Windows NT/CurrentVersion/Print/Providers" -Recurse -Force | Out-Null
	}
	
	"[info]:  removing temporary files" | Writeto-ExecutionLog
	foreach ($file in $RegFilesDeployed)
	{
		if (Test-Path $file)
		{
			Remove-Item $file -ErrorAction SilentlyContinue
		}
	}
	
	if (-not (Test-Path "$env:windir\system32\spoolsv.exe"))
	{
		"[error]:  $env:windir\system32\spoolsv.exe not detected.  Please run SFC /SCANNOW to correct this issue." | Writeto-ExecutionLog
	}
	if (-not (Test-Path "$env:windir\system32\spoolss.dll"))
	{
		"[error]:  $env:windir\system32\spoolss.dll not detected.  Please run SFC /SCANNOW to correct this issue." | Writeto-ExecutionLog
	}
	if (-not (Test-Path "$env:windir\system32\localspl.dll"))
	{
		"[error]:  $env:windir\system32\localspl.dll not detected.  Please run SFC /SCANNOW to correct this issue." | Writeto-ExecutionLog
	}
	if (-not (Test-Path "$env:windir\system32\win32spl.dll"))
	{
		"[error]:  $env:windir\system32\win32spl.dll not detected.  Please run SFC /SCANNOW to correct this issue." | Writeto-ExecutionLog
	}
	if (-not (Test-Path "$env:windir\inf\ntprint.inf"))
	{
		"[error]:  $env:windir\inf\ntprint.inf not detected.  Attempting to copy from $env:windir\System32\DriverStore\FileRepository" | Writeto-ExecutionLog
		$ntprintinfpath = join-path "$env:windir\System32\DriverStore\FileRepository" (Get-ChildItem "$env:windir\System32\DriverStore\FileRepository" | where {$_.name -match "ntprint.inf"}) 
		if (-not (Test-Path ($ntprintinfpath + "\ntprint.inf")))
		{
			"[error]:  ntprint.inf not detected in $env:windir\System32\DriverStore\FileRepository." | Writeto-ExecutionLog
		}
		else
		{
			Copy-Item -Path (join-path $ntprintinfpath "ntprint.inf") -Destination "$env:windir\inf" -OutVariable $cmdoutput
			if (-not (Test-Path ($ntprintinfpath + "\ntprint.inf")))
			{
				"[error]:  failed to copy ntprint.inf from FileRepository." | Writeto-ExecutionLog
			}
			else 
			{
				"[info]:  ntprint.inf successfully copied to $env:windir\inf." | Writeto-ExecutionLog
			}
		}
	}	
	if ($arrSplFiles -ne $null)
	{
		"[info]:  removing files from spooler folder" | Writeto-ExecutionLog
		foreach ($file in $arrSplFiles)
		{
			#"[info]:  deleting $file" | Writeto-ExecutionLog
			Remove-Item $file | Out-Null
		}
	}
	
	#12. Delete all printer shares 
	"[info]:  deleting printer shares." | Writeto-ExecutionLog
	$arrShares = Get-WmiObject Win32_Share | ? {$_.Type -eq "1"}
	if ($arrShares.Length -gt 0)
	{
		foreach ($Share in $arrShares)
		{
			$Share.Delete()
		}
	}
	#13. Delete HKCU\Printers\Connections.  This key will be recreated the next time the user runs Control Panel\Hardware and Sound\Devices and Printers
	
	if (Test-Path -Path "HKCU:\printers\connections")
	{
		"[info]:  deleting hkcu\printers\connections key" | Writeto-ExecutionLog
		Remove-Item -path "HKCU:\printers\connections" -force
	}
	
	#14. Delete non-default files from “%windir%\System32\spool\drivers\<arch>" and “%windir%\System32\spool\drivers\<arch>\3”. Ignore subdirectories. 
	
	$arch = $env:PROCESSOR_ARCHITECTURE

	if ($arch -eq "AMD64")
	{
		$DriversArchFolder = (Join-Path $env:windir "\System32\spool\drivers\$arch").replace("AMD", "x")
		$PrtProcsArchFolder = (Join-Path $env:windir "\System32\spool\prtprocs\$arch").replace("AMD", "x")
	}
	elseif ($arch -eq "x86")
	{
		$DriversArchFolder = (Join-Path $env:windir "\System32\spool\drivers\$arch").replace("x86", "W32X86")
		$PrtProcsArchFolder = (Join-Path $env:windir "\System32\spool\prtprocs\$arch").replace("x86", "W32X86")
	}
	
	$Driversv3Folder = (Join-Path $DriversArchFolder "3")

	$arrprtprocsxclusionList = 
	"en-us",
	"winprint.dll",
	"jnwppr.dll"
	
	$arrArchExclusionList = 
	"3",
	"PCC"

	$arrv3ExclusionList = 
	"mxdwdrv.dll",
	"unidrvui.dll",
	"mxdwdui.gpd",
	"mxdwdui.ini",
	"mxdwdui.dll",
	"UNIDRV.DLL",
	"UNIRES.DLL",
	"STDNAMES.GPD",
	"STDDTYPE.GDL",
	"STDSCHEM.GDL",
	"STDSCHMX.GDL",
	"XPSSVCS.DLL",
	"tsprint.dll",
	"tsprint-datafile.dat",
	"tsprint-PipelineConfig.xml",
	"unidrv.hlp",
	"FXSAPI.DLL",
	"FXSDRV.DLL",
	"FXSRES.DLL",
	"FXSTIFF.DLL",
	"FXSUI.DLL",
	"FXSWZRD.DLL",
	"JNWDRV.dll",
	"jnwdui.dll",
	"unidrvui.dll",
	"en-us",
	"mui"

	"[info]:  deleting non-default files and directories from spool" | Writeto-ExecutionLog
	$archFiles = Get-ChildItem $DriversArchFolder
	foreach ($fsobject in $archFiles)
	{
		if (-not ($arrArchExclusionList -contains $fsobject.Name))
			{
				trap [Exception] 
				{
				    $errorMessage = $_.Exception.Message
				    "[error]: "+ $errorMessage | Writeto-ExecutionLog
				    $Error.Clear()
				    Continue
				}
				remove-item (join-path $DriversArchFolder $fsobject) -Force -Recurse -ErrorAction Stop
			}
	}

	$v3Files = Get-ChildItem $Driversv3Folder
	foreach ($fsobject in $v3Files)
	{
		if (-not ($arrv3ExclusionList -contains $fsobject.Name))
			{
				trap [Exception] 
				{
				    $errorMessage = $_.Exception.Message
				    "[error]: "+ $errorMessage | Writeto-ExecutionLog
				    $Error.Clear()
				    Continue
				}
				remove-item (join-path $Driversv3Folder $fsobject) -Force -Recurse -ErrorAction Stop
			}
	}
	
	$PrtProcsFiles = Get-ChildItem $PrtProcsArchFolder
	foreach ($fsobject in $PrtProcsFiles)
	{
		if (-not ($arrprtprocsxclusionList -contains $fsobject.Name))
			{
				trap [Exception] 
				{
				    $errorMessage = $_.Exception.Message
				    "[error]: "+ $errorMessage | Writeto-ExecutionLog
				    $Error.Clear()
				    Continue
				}
				remove-item (join-path $PrtProcsArchFolder $fsobject) -Force -Recurse -ErrorAction Stop
			}
	}
	
	#15. Start the spooler service and dependents.  It is expected that some third-party dependent services may fail because of missing drivers/configuration information.  These should be reinstalled.
	
	foreach ($service in $arrStartedServices)
	{
		$CurrentService = Get-Service -Name $service
		
		if ($CurrentService.status -eq "stopped")
		{
			$CurrentServiceName = $CurrentService.Name
			$SvcStartTimeout = "300"
			start-service $CurrentServiceName 
			Write-Host "[Info]:  Waiting up to 5 minutes to start service $CurrentServiceName" | Writeto-ExecutionLog
			for ($i = 1; $i -le $SvcStopTimeout; $i++)
			{	
				if ($CurrentService.Status -ne "running")
				{
					Start-Sleep 1
					$CurrentService = Get-Service -Name $CurrentServiceName
				}
				else
				{
					break
				}
			}
			if ($CurrentService.Status -ne "running")
			{
				Write-Host "[error]:  Could not start service $CurrentServiceName" | Writeto-ExecutionLog
			}
		}
	}
	
	"[info]:  Finished full mode" | Writeto-ExecutionLog
}

function printusage()
{
	"-light: Light mode - Corrects common printing issues."
	"1.  Stop spooler service and dependents.  Stop printisolationhost.exe"
    "2.  Do not backup spooler service configuration:  hklm\system\currentcontrolset\services  - warn user -printbrm/PMC"
	"3.  Do not Backup print configuration:  hklm\system\currentcontrolset\control\print "
	"4.  Do not backup files in %windir%\system32\spool\printers\ directory - warn user - back up manually"
	"5.  Import default print monitors (USB, Local Port, TCP/IP, WSD, and Microsoft Shared Fax registry configuration)"
	"6.  Import spooler service default configuration"
	"7.  Import print processor configuration for correct machine architecture:  hklm\system\currentcontrolset\control\print\environments\(machine architecture)\winprint"
	"8.  Detect if spoolsv.exe, spoolss.dll, localspl.dll, or win32spl.dll are not present in %windir%\system32, write error to log."
	"9.  Copy NTPrint.inf from %windir%\driverstore if not present in %windir%\inf.  If (also) not present in driverstore, write error to log."
	"10. Test the path for hklm\system\currentcontrolset\control\print\printers\.DefaultSpoolDirectory. If it is not a valid path, write error to log.  correct this condition by setting it to the default path."
	"11. Delete all files from the spooler directory "
	"12. Start the spooler service and dependents"
	
	"-full: Full mode.  Resets spooler service and print key to defaults."
	"1.   Stop spooler service and dependents.  Stop printisolationhost.exe"
	"2.   Do not Backup spooler service configuration:  hklm\system\currentcontrolset\services  "
	"3.   Do not Backup print configuration:  hklm\system\currentcontrolset\control\print "
	"4.   Do not Backup Lanmanserver/Shares key"
	"5.   Do not Backup HKCU\Printers\Connections"
	"6.   Import default print key for correct machine architecture"
	"7.   Import spooler service default configuration"
	"8.   Detect if spoolsv.exe, spoolss.dll, localspl.dll, or win32spl.dll are not present in %windir%\system32, write error to log."
	"9.   Copy NTPrint.inf from %windir%\driverstore if not present in %windir%\inf.  If (also) not present in driverstore, write error to log."
	"10. Do not backup files in %windir%\system32\spool\printers\ directory"
	"11. Delete all files from %windir%\system32\spool\printers\ directory"
	"12. Delete all printer shares "
	"13. Delete HKCU\Printers\Connections.  This key will be recreated the next time the user runs Control Panel\Hardware and Sound\Devices and Printers"
	"14. Delete non-default files from %windir%\System32\spool\drivers\<arch>, %windir%\System32\spool\drivers\<arch>\3, and %windir%\System32\spool\prtprocs\<arch>. Ignore subdirectories. "
	"15. Start the spooler service and dependents.  It is expected that some third-party dependent services may fail because of missing drivers/configuration information.  These should be reinstalled."
	
	"-force:  Do not prompt for confirmation"
	
	"-quiet:  Do not display console output"
}
###########################
#Process script parameters
###########################

if (-not $Force.IsPresent)
{
	if (test-path $ExecutionDirectory)
	{
#		$PreviousExecutionUserResponse = read-host "Warning:  Data from a previous execution of this tool was detected at $ExecutionDirectory.  `n If you continue, the data in this location will be overwritten.  This is an irreversible operation.  `n Do you wish to continue? [Y] Yes [N] No"
#		if ($PreviousExecutionUserResponse -ne "y")
#		{
#			"exiting..."
#			break
#		}
#		else
#		{
			Remove-Item -Path $ExecutionDirectory -Force -Recurse
			New-Item -Path $ExecutionDirectory -ItemType directory  | Out-null
		}
#	}
	else
	{
		New-Item -Path $ExecutionDirectory -ItemType directory  | Out-null
	}
}
else
{
	if (test-path $ExecutionDirectory)
	{
		Remove-Item -Path $ExecutionDirectory -Force -Recurse
	}
	New-Item -Path $ExecutionDirectory -ItemType directory  | Out-null
}


#block execution on:
#1.  Servers without RDS role.
#2.  Servers with RDS role AND print server role.

$os = Get-WmiObject -Class win32_operatingsystem
if ($os.ProductType -gt 1)
{
	$isServer = $true
}
else
{
	$isServer = $false
}

$ValidateOSConfig = $false
if (((([Environment]::OSVersion.Version).major -eq 6) -and (([Environment]::OSVersion.Version).minor -ge 1)) -or (([Environment]::OSVersion.Version).major -gt 6))
{
	if ($isServer)
	{
		if ((Get-WmiObject -Class Win32_TerminalServiceSetting -Namespace root\cimv2\TerminalServices).TerminalServerMode -eq 1) 
		{
			$isRDSSessionHost = $true
		}
		else
		{
			$isRDSSessionHost = $false
		}
		if ((Get-WmiObject -Class Win32_ServerFeature | Where-Object {$_.ID -eq 7}) -ne $null)
		{
			$isPrintServer = $true 
		}
		else
		{
			$isPrintServer = $false
		}
		if ($isRDSSessionHost)
		{
			if ($isPrintServer)
			{
				#This is a server with RDS session host, but print server is installed.  Block execution.
			}
			else
			{
				#This is a server with RDS session host and it does not have print server role installed.  Allow execution.
				$ValidateOSConfig = $true
			}
		}
		else
		{
			#This is a server without RDS session host.  Block execution.	
		}
	}
	else
	{
		#This is a client.  Allow execution.
		$ValidateOSConfig = $true
	}
}
else
{
	#script is not supported on Windows XP/Vista.  Use Previous fixit - http://support.microsoft.com/kb/324757/lt
	$shell = new-object -comobject wscript.shell
 	$response = $shell.popup(“Script is not supported on this version of Windows.  Please use the tool located at this URL - http://support.microsoft.com/kb/324757/lt “,0,”Unsupported OS Version”,1)
}

$ValidateParams = $false
if ($light.IsPresent -and $full.IsPresent)
{
	#please choose either light or full mode, but not both.
}
if ((-not ($light.IsPresent)) -and (-not ($full.IsPresent)))
{
	#please choose either light or full mode.
}
if ((($light.IsPresent) -and (-not ($full.IsPresent))) -or ((-not ($light.IsPresent)) -and ($full.IsPresent)))
{
	$ValidateParams = $true
}

if (($ValidateOSConfig -eq $true) -and ($ValidateParams -eq $true))
{
	if ($light.IsPresent)
	{
		if (-not $Force.IsPresent)
		{	
			Write-Host -foregroundcolor yellow "Warning:  Running the Print Reset Tool in LIGHT mode performs non-reversible operations."
			Write-Host -foregroundcolor yellow "If you have custom printer configurations, it is strongly recommended that you run printbrm -b to backup the printing environment before you run this tool"
			Write-Host -foregroundcolor yellow "Please refer to this knowledge base article for information about how to backup your printing environment:"
			Write-Host -foregroundcolor yellow "http://support.microsoft.com/kb/938923"
			$LightModeExecutionResponse = Read-Host " `n Do you wish to continue? [Y] Yes [N] No"
			if ($LightModeExecutionResponse -ne "y")
			{
				"exiting..."
				break
			}
			else
			{
				lightmode
			}
		}
		else
		{
			lightmode
		}
	}
	elseif ($full.IsPresent)
	{
		if (-not $Force.IsPresent)
		{	
			Write-Host -foregroundcolor yellow "Warning:  Running the Print Reset Tool in FULL mode performs non-reversible operations."
			Write-Host -foregroundcolor yellow "If you have custom printer configurations, it is strongly recommended that you run printbrm -b to backup the printing environment before you run this tool"
			Write-Host -foregroundcolor yellow "Please refer to this knowledge base article for information about how to backup your printing environment:"
			Write-Host -foregroundcolor yellow "http://support.microsoft.com/kb/938923"
			Write-Host -foregroundcolor red "Running programs with pending print jobs may prevent files from being removed."
			$FullModeExecutionResponse = Read-Host " `n Do you wish to continue? [Y] Yes [N] No"
			if ($FullModeExecutionResponse -ne "y")
			{
				"exiting..."
				break
			}
			else
			{
				fullmode
			}		
		}
		else
		{
			fullmode
		}
	}
	else
	{
		printusage
	}
}
else
{
	if($ValidateOSConfig -eq $false)
	{
		"[error:]  An unsupported role services configuration, or an unsupported Windows version was detected. This tool is targeted at Windows 7, or RDS Session Host servers running Windows Server 2008 R2 without the Print Server role installed. "
	}
	else
	{
		printusage
	}
}

}