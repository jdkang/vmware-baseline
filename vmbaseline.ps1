<#
	.SYNOPSIS
	Applies VM baseline settings from KEY=VALUE files for vmhosts, vms, vswitches, and vportgroups (using PowerCLI)

	.DESCRIPTION
	This script is designed to help apply VM Baselines to vmhosts, vms, vswitches, and vportgroups using "desired state" txt files. These files are essentially KEY=VALUE files which get processed into hashtables and then applied against various objects.

	These "desired tate" files can be chained and merged in an array, ergo you could theortically have multiple levels of settings.

	* A HTML report gets generated with the objects which were scanned as well as changes are made.
	* If chanegs are made, "rollback" files are generated. These txt files can be used to try to rollback VMs 
	* If changes are made, download the current .VMX file before changes.
	* Some basic audits are run and added to the HTML report

	-----------------------------------
	| About Relative Paths
	-----------------------------------
	The script performs a pushd on the script location, so relative paths are relative to the script root directory not PWD.

	e.g. logDir, dsFile_*, etc.

	-----------------------------------
	| IMPORTANT! About VM Settings
	-----------------------------------
	Settings applied to VMs are not instant and require the VM to re-read the VMX settings. This is most commonly done by powering the machine off (not reboot) and on again. There are also other methods to "stun" the VM such as creating a snapshot and removing it or using storage vMotion. 

	There is an EXPERIMENTAL parameter -stunVms parameter which will initiate a vMotion to host the VM is currently on. It's advised you test this feature out in manual mode to see how your VMs behave/tolerate this.

	-----------------------------------
	Modes
	-----------------------------------
	A mode must be specified in the parameters.

	There are two main modes:

	-auto = Utilizes the default settings in vmbaseline.ps1 parameters + code from auto-settings.ps1 (to populate vm objects). This is most useful for automated running. 

	This will automatically connect to the specified vCenters.

	-manual = The main differance is this mode supports explicitly passing the PowerCLI VI objects. 

	This will NOT connect to the vCenters and ASSUMES you already connected to VIServers

	.EXAMPLE
	vmbaseline.ps1 -auto -logging -Verbose

	This will use the logic in auto-settings.ps1 to populate objects, use settings with default values in the parameter list of vmbaselinse.ps1, and enable logging.

	.EXAMPLE
	vmbaseline.ps1 -auto -logging -Verbose -testInitialize
	
	The script runs through the initialization sequence. This is useful for testing desired state, sanity check on what objects would be acted upon, etc.

	.EXAMPLE
	Connect-VIServer vc01.casterlyrock.westeros.local
	$vmHosts = Get-VMHost LealEsxi*
	$vms = Get-VM Vassal*
	$vswitches = Get-VirtualSwitch
	$virtualPortGroups = Get-VirtualPortGroup 
	vmbaseline.ps1 -manual -logging -verbose -vmHosts $vmHosts -vms $vms -vswitches $vswitches -vportgroups $vportgroups

	This will use the VI objects passed via parameters (in lieu of using the auto-settings.ps1) and enable logging,

	.EXAMPLE
	vmbaseline.ps1 -auto -logging -Verbose -stunVms
	
	This will use the logic in auto-settings.ps1 to populate objects, use settings with default values in the parameter list of vmbaselinse.ps1, and enable logging.

	the -stunVMs switch will use the experimental "VM Stun" which is a same-host vMotion. Please consult -detailed help for more information.

	.EXAMPLE
	vmbaseline.ps1 -auto -logging -Verbose -stunVms -useTasks
	
	This will use the logic in auto-settings.ps1 to populate objects, use settings with default values in the parameter list of vmbaselinse.ps1, and enable logging.

	the -stunVMs switch will use the experimental "VM Stun" which is a same-host vMotion. Please consult -detailed help for more information.

	the -useTasks switch will wait on tasks to finish. Not using the the -useTasks switch will cause the script to block between each operation rather than running the tasks in parallel (like when you use vCenter fat client)

	.EXAMPLE
	$vms = Get-VM
	$backedUpVMs =
	foreach ($vm in $vms) {
		$veeamNote = $vm | Get-Annotation -Name 'VEEAM BACKUP OK'
		if ($veeamNote.Value) {
			$vm
		}
	}
	vmbaseline.ps1 -manual -logging -verbose -vm $backedUpVms [...]

	An example of using Veeam's ability to annotate VMs with backup job information to choose VMs. Though, one should probably check the results of the notes/backups beforehand ...
	
	.EXAMPLE
	$vmsOnLun = Get-Datastore -Name VMFSLUN01 | Get-VM
	vmbaseline.ps1 -manual -logging -verbose -vm $vmsOnLun

	Example of selecting VMs on a specific LUN.
#>

#Requires -Version 3
[CmdletBinding(DefaultParameterSetName="nomode")]
param(
<# Modes #>
	
	# Does not connect to VIServers and populate objects via passed PowerCLI VI objects
	[Parameter(ParameterSetName="passedObjs",Mandatory=$true)]
	[switch]
	$manual,
	# Connects to the VIServer(s) specified -and- uses the logic supplied in auto-settings.ps1 once connected to populate VI objects.
	[Parameter(ParameterSetName="hardcodedObjs",Mandatory=$true)]
	[switch]
	$auto,
	
<# Manual Mode #>
	
	# (manual mode) VMHost objects via PowerCLI
	[Parameter(ParameterSetName="passedObjs")]
	[VMware.VimAutomation.ViCore.Impl.V1.Inventory.VMHostImpl[]]
	$vmHosts,
	# (manual mode) VM objects via PowerCLI
	[Parameter(ParameterSetName="passedObjs")]
	[VMware.VimAutomation.ViCore.Impl.V1.Inventory.VirtualMachineImpl[]]
	$vms,
	# (manual mode) VirtualSwitch objects via PowerCLI
	[Parameter(ParameterSetName="passedObjs")]
	[VMware.VimAutomation.ViCore.Impl.V1.Host.Networking.VirtualSwitchImpl[]]
	$vSwitches,
	# (manual mode) VirtualPortGroup objects via PowerCLI
	[Parameter(ParameterSetName="passedObjs")]
	[VMware.VimAutomation.ViCore.Impl.V1.Host.Networking.VirtualPortGroupImpl[]]
	$vPortgroups,
		
<# Logging #>
	
	# Enable logging utilizing the PowershellLoggingModule ( https://github.com/dlwyatt/PowerShellLoggingModule ) which utilizes .NET reflection to capture and "tee" PS output streams.
	[Parameter(ParameterSetName="passedObjs")]
	[Parameter(ParameterSetName="hardcodedObjs")]
	[switch]
	$logging,
	# Log directory. DEFAULT is SCRIPT_DIRECTORY\logs
	[Parameter(ParameterSetName="passedObjs")]
	[Parameter(ParameterSetName="hardcodedObjs")]
	[string]
	$logDir = "",
	
	# The script will only go through the initilizastion phase which can be useful for seeing if it reads desired state and which objects it will act upon.
	[Parameter(ParameterSetName="passedObjs")]
	[Parameter(ParameterSetName="hardcodedObjs")]
	[switch]
	$testInitialize,

<#
!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
!! Ensure default values set correctly when using -auto mode
!! ALSO, make sure you EDIT auto-settings.ps1
!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
#>

<# Desired State Files #>

	# Relative paths are relative to the script root not PWD. VM Host Desired State path. Path(s) to "desired state" txt file(s) in KEY=VALUE format. If multiple files are specified, settings get merged in order they are processed, so conflicting values will let the "newest" value win. Ensure this has a default value when using auto mode.
	[Parameter(ParameterSetName="passedObjs")]
	[Parameter(ParameterSetName="hardcodedObjs")]	
	[string[]]
	$dsFile_vmhosts=@('.\harden_settings\harden_LV2_vmHost_AdvSettings.txt'),
	# Relative paths are relative to the script root not PWD. VM Desired State path. Path(s) to "desired state" txt file(s) in KEY=VALUE format. If multiple files are specified, settings get merged in order they are processed, so conflicting values will let the "newest" value win. Ensure this has a default value when using auto mode.
	[Parameter(ParameterSetName="passedObjs")]
	[Parameter(ParameterSetName="hardcodedObjs")]
	[string[]]
	$dsFile_vm=@('.\harden_settings\harden_LV2_vm_extraconfig.txt'),
	# Relative paths are relative to the script root not PWD. vSwitch Desired State path. Path(s) to "desired state" txt file(s) in KEY=VALUE format. If multiple files are specified, settings get merged in order they are processed, so conflicting values will let the "newest" value win. Values should correspond to parameters in the Set-SecurityPolicy cmdlet as these get converted into SPLAT-able hashtables and passed to this cmdlet. Ensure this has a default value when using auto mode.
	[Parameter(ParameterSetName="passedObjs")]
	[Parameter(ParameterSetName="hardcodedObjs")]
	[string[]]
	$dsFile_vswitch=@('.\harden_settings\harden_LV2_vSwitchSec.txt'),
	# Relative paths are relative to the script root not PWD. vPortgroup Desired State path. Path(s) to "desired state" txt file(s) in KEY=VALUE format. If multiple files are specified, settings get merged in order they are processed, so conflicting values will let the "newest" value win. Values should correspond to parameters in the Set-SecurityPolicy cmdlet as these get converted into SPLAT-able hashtables and passed to this cmdlet. Ensure this has a default value when using auto mode.
	[Parameter(ParameterSetName="passedObjs")]
	[Parameter(ParameterSetName="hardcodedObjs")]
	[string[]]
	$dsFile_vportgroup=@('.\harden_settings\harden_LV2_vPortGroupSec.txt'),

<# Msc. Options #>
	
	# vCenter servers to connect to. Ensure this has a default value when using auto mode.
	[Parameter(ParameterSetName="hardcodedObjs")]
	[string[]]
	$vCenterServers = @('vc01.contoso.local'),

	# This is the directory where rollbacks and reports are saved. Each time the script is run, it will get its own timestamp subdirectory. DEFAULT value (if none is passed) is SCRIPT_DIRECTORY\out
	[Parameter(ParameterSetName="passedObjs")]
	[Parameter(ParameterSetName="hardcodedObjs")]
	[string]
	$saveDirectory = "",
	# EXPERIMENTAL option to stun VMs after applying settings by initiating a same-host vMotion.
	[Parameter(ParameterSetName="passedObjs")]
	[Parameter(ParameterSetName="hardcodedObjs")]
	[switch]
	$stunVms,	

	# Whether to use tasks for applying VM settings and vmStun (vMotion). Tasks are parallel but the script will still block to wait for all the tasks to finish.
	[Parameter(ParameterSetName="passedObjs")]
	[Parameter(ParameterSetName="hardcodedObjs")]
	[switch]
	$useTasks
)

# Default desired state file locations (relative to script path)
# if (!$dsFile_vmhosts) {
# 	$dsFile_vmhosts = @("$currentScriptPath\harden_settings\harden_LV2_vmHost_AdvSettings.txt")
# }
# if (!$dsFile_vm) {
# 	$dsFile_vm = @("$currentScriptPath\harden_settings\harden_LV2_vm_extraconfig.txt")
# }
# if (!$dsFile_vswitch) {
# 	$dsFile_vswitch = @("$currentScriptPath\harden_settings\harden_LV2_vSwitchSec.txt")
# }
# if (!$dsFile_vportgroup) {
# 	$dsFile_vportgroup = @("$currentScriptPath\harden_settings\harden_LV2_vPortGroupSec.txt")
# }

#=========================================================================================
# Init
#=========================================================================================
# Tasks SPLAT
$taskSplat = @{}
if ($useTasks) {
	write-verbose "-useTasks : Using tasks"
	$taskSplat.Add('useTasks',$true)
}

# Verbose SPLAT
$verboseSplat = @{}
if($PSBoundParameters['Verbose']) {
	$verboseSplat.Add('verbose',$true)
}

# Proper exit codes
function Exit-WithCode
{
param([int]$exitcode)
	if (!$exitcode) {
		if ($script:CUSTOM_EXIT_CODE) {
			$exitcode = $script:CUSTOM_EXIT_CODE
		} else {
			$exitcode = 0
		}
	}

	# powershell.exe command
	$powershellProcessCommand = (gwmi win32_process | ? { $_.processid -eq $PID }) | select commandline
	# callstack string
	[string]$callersStr = ""
	[string[]]$callers = Get-PsCallStack | Select -Expand Command
	($callers.length - 2)..1 | % {
		$callersStr += "[$($callers[$_])]"
	}
	if ($powershellProcessCommand -notlike '*-nointeractive*') {
		write-host "$callersStr Exit $exitcode" -foregroundcolor black -backgroundcolor green
		break
	} else {
		$host.SetShouldExit($exitcode)
	}
}
trap {
	$script:CUSTOM_EXIT_CODE = 1

	$_ | select *
	write-verbose "Disabling log file before exit."
	if ($logging) {
		$LogFile | Disable-LogFile
	}
	Pop-Location
	Exit-WithCode $script:CUSTOM_EXIT_CODE
}

$initialTsStr = Get-Date -f 'yyyyMMdd_HHmmss'
$scriptStartTime = Get-Date
$currentScriptPath = Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path
$autoSettingsPs1 = "$currentScriptPath\auto-settings.ps1" # auto-settings.ps1
Push-Location $currentScriptPath

# Logging
$loggingModulePath = "$currentScriptPath\lib\PowerShellLogging\PowerShellLogging.psd1"
if (!(Test-Path $loggingModulePath)) {
	throw "Cannot find logging module $loggingModulePath"
}
ipmo $loggingModulePath

if ($logging) {
	write-verbose "-logging"
	write-verbose "LOGGING initializing ..."
	# Log dir
	if (!$logDir) {
		$logDir = "$currentScriptPath\logs"
		write-verbose "Setting log directory to $logDir"
	}
	if (!(Test-Path $logDir)) {
		write-verbose "Creating directory $logDir"
		md $logDir -force | out-null
		if (!$?) {
			throw "Unable to create $logDir"
		}
	}
	$logFilePath = "$logDir\$($initialTsStr).log"
	$LogFile = Enable-LogFile -Path $logFilePath
	$VerbosePreference = 'Continue' 
	$DebugPreference = 'Continue'
}

# Kinder parametersetname message
if ($PsCmdlet.ParameterSetName -eq "nomode") {
	throw "No mode parameter specified. Please see help for more details about modes."
}

# auto-settings.ps1
if ($PsCmdlet.ParameterSetName -eq 'hardcodedObjs') {
	if (!(Test-Path $autoSettingsPs1)) {
		throw "Cannot find auto-setings.ps1 at $autoSettingsPs1"
	}
}

# Load \Libs
Remove-Module vmware.extraconfig -ErrorAction 0
write-verbose "Loading \libs"
$libs = @(
	"$currentScriptPath\lib\vmware.extraconfig\VMWare.Extraconfig.psd1",
	"$currentScriptPath\lib\HTML-Reportify.ps1"
)
foreach ($lib in $libs) {
	write-verbose "Checking $lib"
	if (!(Test-Path $lib)) {
		throw "Cannot find dep $lib"
	} else {
		$ext = (gi $lib).Extension
		switch ($ext) {
			".ps1" {
				write-verbose "Dot sourcing $lib"
				. "$lib"
				if (!$?) {
					throw "Error loading $lib"
				}
				break
			}
			".psd1" {
				write-verbose "Importing module $lib"
				ipmo $lib
				if (!$?) {
					throw "Err loading $lib"
				}
				break
			}
			default { throw "Cannot determine extension of $lib" }
		}
	}
}

# Validate/Load PSSnapin
#Validate-PSSnapin -name "vmware.vimautomation.core" -Verbose -Important -VersionMajor 5 -VersionMinor 5 -VersionBuild 0 -VersionRevision 0

# -auto
# Load PSSnappin and Connect to vCenter 
if ($PsCmdlet.ParameterSetName -eq 'hardcodedObjs') {
	write-verbose "Loading snappin vmware.VimAutomation.core"
	Add-PSSnapin vmware.VimAutomation.core -ErrorAction SilentlyContinue
	
	# Clear connected VI Servers
	if ($global:DefaultVIServers) {
		write-verbose "Disconnecting from existing servers: $($global:DefaultVIServers)"
		Disconnect-VIServer -Server $global:DefaultVIServers -Force -confirm:$false
	}

	# Connect to vCenters
	write-verbose "Connecting to vCenters $vCenterServers"
	foreach ($vCenter in $vCenterServers) {
		write-verbose "Connecting to $vCenter"
		Connect-VIServer $vCenter | out-null
		if (!$?) {
			throw "Issue connecting to vCenter $vCenter"
		}
	}
}

# Try Test-Connection
write-verbose "Testing VIServer connection (fetching datacenters)"
Get-Datacenter | out-null
if (!$?) {
	write-warning "Could not gather datacenter information. Please check connection."
	write-host "Clearing your VI connection." -f magenta
	# Clear connected VI Servers
	if ($global:DefaultVIServers) {
		write-verbose "Disconnecting from existing servers: $($global:DefaultVIServers)"
		Disconnect-VIServer -Server $global:DefaultVIServers -Force -confirm:$false
	}
	$script:CUSTOM_EXIT_CODE = 2
	throw "Could not get datacenter information. Please check connection"
}

# run auto-settings.ps1
if ($PsCmdlet.ParameterSetName -eq 'hardcodedObjs') {
	write-verbose "Importing hard-coded logic from "
	. $autoSettingsPs1
	
	if (!$?) {
		throw "Issue importing auto-settings.ps1"
	}
}

# map to hard coded values
switch ($PsCmdlet.ParameterSetName) {
	"hardcodedObjs" {
		write-verbose "-auto mode (using hard-coded logic)"

		$vmHosts = $hardcoded_vmHosts
		$vms = $hardcoded_vms
		$vSwitches = $hardcoded_vSwitches
		$vPortGroups = $hardcoded_vPortgroups
	}
	"passedObjs" {
		write-verbose "-manual mode (using passed objects)"
	}
}

# sanity check
if (!$vmHosts -and !$vms -and !$vSwitches -and !$vPortGroups) {
	throw "No objects to process."
}

# Load desired state from files
write-verbose "#========================================================================================="
write-verbose "Loading desired state."
write-verbose "#========================================================================================="
# Load Desired State
$ds_vmhosts = Load-HashtableFromFile $dsFile_vmhosts @verboseSplat
$ds_vm = Load-HashtableFromFile $dsFile_vm @verboseSplat
$ds_vswitch = Load-HashtableFromFile $dsFile_vswitch @verboseSplat
$ds_vportgroup = Load-HashtableFromFile $dsFile_vportgroup @verboseSplat

write-verbose "#========================================================================================="
write-verbose "Finishing initialization"
write-verbose "#========================================================================================="
# Setup save directory structure
if (!$saveDirectory) {
	$saveDirectory = "$currentScriptPath\out"
	write-verbose "Setting report dir to $saveDirectory"
	if (!$?) {
		throw "Unable to create $saveDirectory"
	}
}
if (!(Test-Path $saveDirectory)) {
	write-verbose "Creating directory $saveDirectory"
	md $saveDirectory -force | out-null
}
$savePath = "$saveDirectory\$($initialTsStr)"
$vmHostSaveDir = "$savepath\rollbacks\vmhost"
$vmSaveDir = "$savepath\rollbacks\vm"
$vSwitchSaveDir = "$savepath\rollbacks\vswitch"
$vPortGroupSaveDir = "$savepath\rollbacks\vportgroup"

write-verbose "Creating rollback directories"
md $savePath -force | Out-Null
md $vmHostSaveDir -force | Out-Null
md $vmSaveDir -force | Out-Null
md $vSwitchSaveDir -force | Out-Null
md $vPortGroupSaveDir -force | Out-Null

write-verbose "------------------------------------"
write-verbose "Objects:"
write-verbose "------------------------------------"
write-verbose ('VMHosts = ' + ($vmHosts | ft | out-string))
write-verbose ('VMs = ' + ($vms | ft | out-string))
write-verbose ('vSwitches = ' + ($vSwitches | ft | out-string))
write-verbose ('vPortgroups = ' + ($vPortGroups | ft | out-string))
write-verbose "------------------------------------"
write-verbose "Desired State"
write-verbose "------------------------------------"
write-verbose ($ds_vmhosts | ft | out-string)
write-verbose ($ds_vm | ft | out-string)
write-verbose ($ds_vswitch | ft | out-string)
write-verbose ($ds_vportgroup | ft | out-string)

if ($testInitialize) {
	write-verbose "-testInitialize : Initlization Done. Exiting"
	if ($logging) {
		write-verbose "And now my watch has ended."
		$LogFile | Disable-LogFile 
	}
	Exit-WithCode 0
}


#=========================================================================================
# MAIN
#=========================================================================================

$script:vmTaskList = @()
$script:vmTaskListStun = @()
$script:scannedObjs = @()
$script:appliedVms = @()
$script:stunnedVms = @()

$saveManifestSuffix = '_viManifest.xml'

#-------------------------------------------------------
# Apply Desired State
#-------------------------------------------------------

write-verbose "#========================================================================================="
write-verbose "APPLYING Desired State"
write-verbose "#========================================================================================="

# VmHost
write-verbose "------------------------------"
write-verbose "VM Hosts"
write-verbose "------------------------------"

if ($vmHosts) {
	Try {
		Foreach ($vh in $vmHosts) {
			write-verbose "---- VMHost: $vh -----"
			$vmHostApply = $null

			# Apply and create rollbacks
			$vmHostApply = Update-VMHostAdvancedSettings -vmhosts $vh -ds $ds_vmhosts -saveDir $vmHostSaveDir @verboseSplat
			if ($vmHostApply) {
				$script:scannedObjs += $vmHostApply
			}
		}
	}
	Catch {
		write-error "Error applying VM host state."
		throw $_
	}
}

# VMs
write-verbose "------------------------------"
write-verbose "VMs"
write-verbose "------------------------------"

if ($vms) {
	Try {
		foreach ($vm in $vms) {
			write-verbose "---- VM: $vm ----"
			$vmApply = $null

			# Apply and create rollbacks (+ copy .vmx file)
			$vmApply = Update-VMAdvancedConfiguration -vms $vm -ds $ds_vm -saveDir $vmSaveDir @taskSplat @verboseSplat
			if ($vmApply) {
				$script:scannedObjs += $vmApply
				if ($vmApply.applied) {
					$script:appliedVms += $vm

					# Queue task
					if ($vmApply.viTask) {
						$script:vmTaskList += $vmApply.viTask
					}
				}
			}			
		}
		# Wait for task queue
		if ($script:vmTaskList) {
			write-verbose "Waiting on tasks."
			write-verbose ($script:vmTaskList | ft | out-string)
			wait-task $script:vmTaskList
		}
	}
	Catch {
		write-error "Error applying VM state."
		throw $_
	}
}

# vSwitches
write-verbose "------------------------------"
write-verbose "vSwitches"
write-verbose "------------------------------"

if ($vSwitches) {
	Try {
		foreach ($vsw in $vSwitches) {
			write-verbose "---- vSwitch: $vsw ----"
			$vSwitchApply = $null

			# Apply and create rollbacks
			$vSwitchApply = Update-vSwitchSettings -vSwitches $vsw -ds $ds_vswitch -saveDir $vSwitchSaveDir @verboseSplat
			if ($vSwitchApply) {
				$script:scannedObjs += $vSwitchApply
			}
		}
    }
	Catch {
		write-error "Error applying vSwitch VM state."
		throw $_
	}
}

# vPortGroups
write-verbose "------------------------------"
write-verbose "vPortGroups"
write-verbose "------------------------------"

if ($vPortgroups) {
	Try {
		foreach ($vpg in $vPortgroups) {
			write-verbose "---- vPortGroup: $vpg ----"
			$vPortGroupApply = $null

			# Apply and create rollbacks
			$vPortGroupApply = Update-vPortgroupSettings -vPortgroups $vpg -ds $ds_vportgroup -saveDir $vPortGroupSaveDir @verboseSplat
			if ($vPortGroupApply) {
				$script:scannedObjs += $vPortGroupApply
            }
        }
	}
	Catch {
		write-error "Error applying vPortgroup state."
		throw $_
	}
}

# stun VMs
write-verbose "------------------------------"
write-verbose "VM Stuns"
write-verbose "------------------------------"
if ($stunVms -and $script:appliedVms) {
	write-verbose "Initiating VM stuns on VMs with applied (new) settings."
	foreach ($vm in $script:appliedVms) {
		write-verbose "Stunning VM $($vm.name) ..."
		$vmStun = $vm | Start-VMStun -Confirm:$false @taskSplat @verboseSplat
		if ($?) {
			$script:stunnedVms += $vm
		}
		if ($vmStun.viTask) {
			$script:vmTaskListStun += $vmStun.viTask
		}
	}
}
if ($script:vmTaskListStun) {
	write-verbose "Waiting on tasks."
	write-verbose ($script:vmTaskListStun | ft | out-string)
	wait-task $script:vmTaskListStun
}

# Save manifest
$manifestSavePath = "$savePath\Manifest_$($initialTsStr).xml"
$scannedObjs | Export-Clixml -Path $manifestSavePath -Depth 5 -Force

#=========================================================================================
# REPORTING and Return Object
#=========================================================================================
$script:scriptRetObjHt = @{}
$script:reportRetObjHt = @{}
$reportPath = "$savePath\VMware_Baseline_Report_$($initialTsStr).html"
$deltaReportName = "VMware_Baseline_Report_$($initialTsStr)_delta.html"
$deltaReportPath = "$savePath\$deltaReportName"

write-verbose "#========================================================================================="
write-verbose "GENERATING Report"
write-verbose "#========================================================================================="
$htmlBody += "<center><h1>VMWare Baseline Report</h1></center>`n"

#-------------------------------------------------------
# Env Info
#-------------------------------------------------------
write-verbose "Writng top information"
$CurrentIPs = @()
get-wmiobject win32_networkadapterconfiguration | ? { $_.IPAddress -ne $null } | Sort-Object IPAddress -Unique | % { 
   $CurrentIPs+=$_.IPAddress 
} 

$scriptParametersHt = (@{} + $psBoundParameters).GetEnumerator()
$scriptParamHtml = ""
$scriptParametersHt | % {
	$scriptParamHtml += "<br /> -$($_.key) $($_.value)"
}

# General Info
$htmlBody += @"
<B>Script Mode</B>: $($PsCmdlet.ParameterSetName) <br />
<B>Generated</B>: $((Get-Date).ToUniversalTime()) UTC <br />
<B>Endpoint:</B>: $($ENV:Computername) <br />
<B>Endpoint IP(S):</B>: $CurrentIPs <br />
<B>User:</B>: $($ENV:USERNAME)@$($ENV:USERDNSDOMAIN) <br />
<B>Script Parameters</B>: <code>$scriptParamHtml</code><br />
<B>Designation:</B> TK-421<br />
<B>Thermal Exhaust Port:</B> Open<br /> 
"@

$logStatus = ""
if ($logging) {
	$script:scriptRetObjHt.Add('logFilePath',$logFilePath)
	$logStatus = $logFilePath
} else {
	$logStatus = "NO LOGGING"
}
$htmlBody += "<h3>Save Paths</h3>"
$htmlBody += "<B>LOG FILE:</B>: <code>$($logStatus)</code> <br />"
$htmlBody += "<B>OUTPUT PATH:</B>: <code>$($savePath)</code> <br />"

$script:scriptRetObjHt.Add('savePath',$savePath)

# vCenters
$htmlBody += "<h3>Connected vCenters</h3>"
$htmlBody += ($Global:DefaultVIServers | % { "$($_.Name):$($_.Port)" } | ConvertTo-HtmlList)
$script:scriptRetObjHt.Add('viServers',$Global:DefaultVIServers)

#-------------------------------------------------------
# Scanned & Remedied Objects
#-------------------------------------------------------
$scannedObjCount = ($script:scannedObjs | measure).count
$htmlBody += "<h2>Scanned & Remedied Objects ($($scannedObjCount))</h2>`n"
$colorFilters = @()
$colorFilters += new-object psobject -property @{
	markType = "cell"
	property = 'Applied'
	filter = '$_.Applied -eq $False'
	class = 'good'
}
$colorFilters += new-object psobject -property @{
	markType = "cell"
	property = 'Applied'
	filter = '$_.Applied -eq $True'
	class = 'bad'
}

write-verbose "--- Processed Objects ----"
$script:scannedObjs | ft | out-string | write-verbose

$htmlBody +=
	$script:scannedObjs |
	ConvertTo-HTML -Fragment -Property Name,Type,Applied,vmUuid,delta |
	Format-PrettyTable -colorFilters $colorFilters -sourceObj $script:scannedObjs

$script:scriptRetObjHt.Add('scannedObjs',$script:scannedObjs)

# Stunned Vms

$stunVmMessage = "<blockquote>VMX settings are not automagically applied. VMX settings are read during a COMPLETE power off/on, snapshots, and during vMotion/Storage vMotion. VM Stunning is basically a vMotion to the host a VM is already on, which initiates a 'stun' which causes the VMX-read. As the name implies, however it does cause a slight 'stun' on the VM, so test this in your enviornment.</blockquote>"
write-verbose "---- Stunned VMs ----"
if ($stunVms -and $script:stunnedVms) {
	$htmlBody += "<h2>Stunned VMs</h2>"
	$htmlBody += $stunVmMessage
	$htmlBody +=
		$script:stunnedVms |
		ConvertTo-HTML -Fragment -Property Name,Folder,VMHost,PowerState |
		Format-PrettyTable -sourceObj $script:stunnedVms
} elseif ($script:appliedVms) {
	$htmlBody += "<h2>Applied (NOT STUNNED) VMs</h2>"
	$htmlBody += $stunVmMessage
	$htmlBody +=
		$script:appliedVms |
		ConvertTo-HTML -Fragment -Property Name,Folder,VMHost,PowerState |
		Format-PrettyTable -sourceObj $script:appliedVms
}
$script:scriptRetObjHt.Add('appliedVms',$script:appliedVms)
if ($script:stunnedVms) {
	$script:scriptRetObjHt.Add('stunnedVms',$script:stunnedVms)
}

#-------------------------------------------------------
# Audits
#-------------------------------------------------------
$htmlBody += "<h2>Audits</h2>`n"
	
# Host Domain Membership
write-verbose "------------------------------"
write-verbose "Host Domain Membership"
write-verbose "------------------------------"
$htmlBody += "<h3>Host Domain Memmbership</h3>`n"
if ($vmHosts) {
	$vmhostDomain = $vmHosts | Get-VMHostAuthentication | Select VmHost,Domain,DomainMembershipStatus
	$colorFilters = @()
	$colorFilters += new-object psobject -property @{
		markType = "cell"
		property = 'Domain'
		filter = '[string]::IsNullOrEmpty($_.Domain)'
		class = 'bad'
	}
	$htmlBody += 
		$vmhostDomain |
		ConvertTo-HTML -Fragment -Property VmHost, Domain, DomainMembershipStatus |
		Format-PrettyTable -colorFilters $colorFilters -sourceObj $vmhostDomain

	$script:reportRetObjHt.Add('vmHostDomain',$vmHostDomain)
} else {
	$htmlBody += '<p>No Hosts Specified</p>'
}


# Unrestricted Enabled Services
write-verbose "------------------------------"
write-verbose "Unrestricted Enabled Services"
write-verbose "------------------------------"
$htmlBody += '<h3>Host Firewall: AllIP, Enabled Services</h3>'
$htmlBody += 'These services are enabled allow from ALL IPs <br /><br />'
if ($vmHosts) {
	$unrestrictedHostServices = $null
	$unrestrictedHostServices = $vmHosts | Get-VMHostFirewallException | ? {$_.Enabled -and ($_.ExtensionData.AllowedHosts.AllIP)} | Sort -Property ServiceRunning -descending | Select VMHost,Name,IncomingPorts,OutgoingPorts,Protocols,ServiceRunning


	if ($unrestrictedHostServices) {
		$colorFilters = @()
		$colorFilters += new-object psobject -property @{
			markType = "cell"
			property = 'ServiceRunning'
			filter = '$_.ServiceRunning -eq $true'
			class = 'bad'
		}
		$htmlBody += 
			$unrestrictedHostServices |
			ConvertTo-HTML -Fragment -Property VMHost,Name,IncomingPorts,OutgoingPorts,Protocols,ServiceRunning |
			Format-PrettyTable -colorFilters $colorFilters -sourceObj $unrestrictedHostServices

		$script:reportRetObjHt.Add('unrestrictedHostServices',$unrestrictedHostServices)
	} else {
		$htmlBody += "n/a <br />"
	}
} else {
	$htmlBody += '<p>No Hosts Specified</p>'
}

# NTP
write-verbose "------------------------------"
write-verbose "NTP"
write-verbose "------------------------------"
$htmlBody += "<h3>Host NTP Settings</h3>`n"
if ($vmHosts) {
	$hostNtp = $null
	$hostNtp = $vmHosts | Select Name, @{N="NTPSetting";E={ ($_ | Get-VMHostNtpServer) }}
	$colorFilters = @()
	$colorFilters += new-object psobject -property @{
		markType = "cell"
		property = 'NTPSetting'
		filter = '[string]::IsNullOrEmpty($_.NTPSetting)'
		class = 'bad'
	}
	$htmlBody += 
		$hostNtp |
		ConvertTo-HTML -Fragment -Property Name,NTPSetting |
		Format-PrettyTable -colorFilters $colorFilters -sourceObj $hostNtp

	$script:reportRetObjHt.Add('hostNtp',$hostNtp)
} else {
	$htmlBody += '<p>No Hosts Specified</p>'
}

# SNMP
write-verbose "------------------------------"
write-verbose "SNMP"
write-verbose "------------------------------"
$htmlBody += "<h3>Host SNMP</h3>`n"
if ($vmHosts) {
	$hostSnmp = @()
	foreach ($vmHost in $vmHosts) {
		$esxcli = Get-EsxCli -vmhost $vmHost.Name
		$hostSnmp += $esxcli.system.snmp.get() | Select @{N='vmhost';E={$vmHost.Name}},*
	}
	$colorFilters = @()
	$colorFilters += new-object psobject -property @{
		markType = "cell"
		property = 'communities'
		filter = '[string]::IsNullOrEmpty($_.communities)'
		class = 'bad'
	}
	$colorFilters += new-object psobject -property @{
		markType = "cell"
		property = 'enable'
		filter = '$_.enable -eq "false"' # esxcli is returning a string not a bool
		class = 'bad'
	}
	$htmlBody += 
		$hostSnmp |
		ConvertTo-HTML -Fragment -Property vmhost,* |
		Format-PrettyTable -colorFilters $colorFilters -sourceObj $hostSnmp

	$script:reportRetObjHt.Add('hostSnmp',$hostSnmp)
} else {
	$htmlBody += '<p>No Hosts Specified</p>'
}

# SSL certs
write-verbose "------------------------------"
write-verbose "SSL Cert Check"
write-verbose "------------------------------"
$htmlBody += "<h3>Host SSL Certificates</h3>`n"
if ($vmHosts) {
	$hostCertStatus = $vmHosts | % { 
		Test-WebServerSSL -URL $_.Name | Select OriginalURi, CertificateIsValid, Issuer, @{N="Expires";E={$_.Certificate.NotAfter} }, @{N="DaysTillExpire";E={(New-TimeSpan -Start (Get-Date) -End ($_.Certificate.NotAfter)).Days} }
	}
	$colorFilters = @()
	$colorFilters += new-object psobject -property @{
		markType = "cell"
		property = 'CertificateIsValid'
		filter = '$_.CertificateIsValid -eq $false'
		class = 'bad'
	}
	$htmlBody += 
		$hostCertStatus |
		ConvertTo-HTML -Fragment -Property OriginalURi, CertificateIsValid, Issuer |
		Format-PrettyTable -colorFilters $colorFilters -sourceObj $hostCertStatus

	$script:reportRetObjHt.Add('hostCertStatus',$hostCertStatus)
} else {
	$htmlBody += '<p>No Hosts Specified</p>'
}

# Host Running Services
write-verbose "------------------------------"
write-verbose "Host Running Services"
write-verbose "------------------------------"
if ($vmHosts) {
	$htmlBody += "<h3>Host Shell/SSH Services</h3>"
	$hostServices = $vmHosts | Get-VMHostService | ? { $_.key -like 'TSM*' } | Sort -Property Running -descending | Select VMHost, Key, Label, Policy, Running, Required
	$colorFilters = @()
	$colorFilters += new-object psobject -property @{
		markType = "cell"
		property = 'Running'
		filter = '$_.Running -eq $true'
		class = 'bad'
	}
	$htmlBody += 
		$hostServices |
		ConvertTo-HTML -Fragment -Property VMHost, Key, Label, Policy, Running, Required |
		Format-PrettyTable -colorFilters $colorFilters -sourceObj $hostServices

	$script:reportRetObjHt.Add('hostServices',$hostServices)
} else {
	$htmlBody += '<p>No Hosts Specified</p>'
}
	
# vCenter NFC SSL
write-verbose "------------------------------"
write-verbose "vCenter NFC SSL"
write-verbose "------------------------------"
$htmlBody += "<h3>vCenter NFC SSL</h3>`n"
$vCenters = $Global:DefaultVIServers | Select -expand Name
$vcNfcSsl = Test-VcenterNfcSsl -vcenters $vcenters
$colorFilters = @()
$colorFilters += new-object psobject -property @{
	markType = "cell"
	property = 'nfcssl'
	filter = '$_.nfcssl -eq $false'
	class = 'bad'
}
$htmlBody += 
	$vcNfcSsl |
	ConvertTo-HTML -Fragment -property vCenter,nfcssl |
	Format-PrettyTable -colorFilters $colorFilters -sourceObj $vcNfcSsl

$script:reportRetObjHt.Add('vcNfcSsl',$vcNfcSsl)

#-------------------------------------------------------
# Add boilerplate
#-------------------------------------------------------
write-verbose "------------------------------"
write-verbose "Report boilerplate / write-out"
write-verbose "------------------------------"
$elapsedTime = New-Timespan $scriptStartTime (Get-Date)
$htmlBody += "<br /><hr /><i>Script ran for $($elapsedTime.TotalSeconds) seconds with $(Get-Random -min 1 -max 9999) units of macho maddness <i>"

write-verbose "Script ran for $($elapsedTime.TotalSeconds) seconds"

write-verbose "Assembling HTML report"
$voltron = (New-HtmlReportHead) + $htmlBody + (New-HtmlReportBottom -extraFun)
	
#-------------------------------------------------------	
# Save report
#-------------------------------------------------------
write-verbose "SAVING Report $reportPath"
$voltron | sc -path $reportPath -force

$script:scriptRetObjHt.Add('reportPath',$reportPath)

#-------------------------------------------------------	
# Cleanup
#-------------------------------------------------------
if ( (gci "$savePath\rollbacks" -recurse | ? { !$_.PsIsContainer } | measure).Count -eq 0 ) {
	write-verbose "rollback directory $savePath\rollbacks is empty. Removing."
	rm "$savePath\rollbacks" -recurse -force -confirm:$false
}

if ($logging) {
	write-verbose "And now my watch has ended."
	$LogFile | Disable-LogFile 
}

#-------------------------------------------------------	
# Return
#-------------------------------------------------------
$reportRetObj = new-object psobject -property $script:reportRetObjHt
$script:scriptRetObjHt.Add('audit',$reportRetObj)
new-object psobject -property $script:scriptRetObjHt
Pop-Location
Exit-WithCode 0