#------------------------------------------
# filesystem io
#------------------------------------------
function Remove-InvalidFileNameChars {
[CmdletBinding()] 
param(
    [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
    [String]
    $Name
)
	write-verbose "Removing invalid filename chars from $Name"
	$invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
	$re = "[{0}]" -f [RegEx]::Escape($invalidChars)
	$ret = $Name -replace $re
	write-verbose "Sanatized version: $ret"
	$ret
}

#------------------------------------------
# html
#------------------------------------------
function ConvertTo-HtmlList {
param(
	[Parameter(Mandatory=$true,ValueFromPipeline=$true)]
	[ValidateNotNullOrEmpty()]
	$object,
	[Parameter(Mandatory=$false)]
	[switch]
	$orderedList
)
	BEGIN {
		if (!$orderedList) {
			"<ul>"
		} else {
			"<ol>"
		}
	}
	PROCESS {
		foreach ($obj in $object) {
			if ($obj) {
				"<li>$($obj)</li>"
			}
		}
	}
	END {
		if (!$orderedList) {
			"</ul>"
		} else {
			"</ol>"
		}
	}
}

#------------------------------------------
# Hashtable
#------------------------------------------

function Get-HashtableRollback {
param(
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[hashtable]
	$currHt,
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[hashtable]
	$desiredHt
)

	write-verbose "Creating rollback HT"
	$rollbackHt = @{}
		foreach ($item in $desiredHt.GetEnumerator()) {
			$itemKey = $null
			$itemValue = $null
			$itemKey = $item.Key
			$itemValue = $item.Value
			
			write-verbose ($currHt | fl | out-string)
			
			write-verbose "DESIRED VALUE: $itemKey = $itemValue"
			
			if ($currHt.ContainsKey($itemKey)) {
				$currItemValue = $currHt[$itemKey]
				
				write-verbose "* CURRENT VALUE: $itemKey = $currItemValue"
				if ($itemValue -ne $currItemValue) {
					write-verbose "!! CHANGE: $itemKey = $itemValue"
					$rollbackHt.Add($itemKey,$currItemValue)
				} else {
					write-verbose "** No change"
				}
			} else {
				write-verbose "KEY DOES NOT EXIST."
				write-verbose "CURRENT VALUE: $itemKey = [not set]"
				$rollbackValue = " "
				if (($itemValue -ne 'false') -and ($itemValue -ne 'true')) {
					$rollbackValue = "0"
				} else {
					$boolItemValue = [System.Convert]::ToBoolean($itemValue)
					$rollbackValue = "$(!$boolItemValue)"
				}
				
				write-verbose "Setting rollback value for $itemKey to $rollbackValue"
				$rollbackHt.Add($itemKey,$rollbackValue)
			}
		} # foreach $item
	$rollbackHt
}

function Write-HashtableToFile {
param(
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[hashtable]
	$ht,
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[string]
	$path
)
	write-verbose "Writing HT to $path"
	[string[]]$fileContents = @()
	$ht.GetEnumerator() | % {
		$fileContents += "$($_.Key)=$($_.value)"
	}
	$fileContents | sc -path $path -force
	$fileContents
}

function Convert-HashtableValuesToTypes {
param(
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[hashtable]$hashtable
)
	[hashtable]$ret = @{}
	
	write-verbose "Converting hashtable types from strings to types"
	Foreach($item in $hashtable.GetEnumerator()) {
		$newVal = $null
		write-verbose "KEY: $($item.key) | VALUE: $($item.value)"
		# parse strings
		if ($item.Value.GetType().Name -eq 'string') {
			write-verbose "Value type = STRING"
			# bool
			if (($item.Value -eq 'true') -or ($item.Value -eq 'false')) {
				write-verbose "Probably Boolean. Attempting to convert"
				$val = [System.Convert]::ToBoolean($item.value)
				if ($?) {
					$newVal = $val
					write-verbose "$newVal OK $($val.GetType().Name)" 
				} else {
					write-error "Error converting."
				}
			} else {
				# numbers
				$valNum = 0
				$isNum = [System.Int32]::TryParse($item.Value, [ref]$valNum)
				if ($isNum) {
					$newVal = $valNum
				} else {
					# default
					$newVal = $item.value
				}
			}
		} else {
			write-verbose "Type NOT a string. Using existing value."
			$newVal = $item.value
		}
		
		if ($newVal -ne $null) {
			$ret.Add($item.Key,$newVal)
		} else {
			throw "Error conveting hashtable value $($item.value) from key $($item.key)"
		}
	}
	$ret
}

function Convert-ObjPropertiesToHashtable {
param(
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	$obj
)

	[hashtable]$ret = @{}
	$objProperties = $obj | gm -membertype *property | select -expand name
	Foreach ($objProperty in $objProperties) {
		$ret.Add($objProperty,$obj.$objProperty)
		if (!$?) { return $null }
	}
	$ret
}

function Load-HashtableFromFile {
[CmdletBinding()]
param(
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[string[]]
	$files
)
	[hashtable]$desiredStateConfig = @{}

	foreach($file in $files) {
		write-verbose "Generating desired state config from files: $file"
		if (!(Test-Path $file)) {
			throw "$file could not be found."
		} else {
			$dsFullPath = (gi $file).FullName
			$encoding = New-Object System.Text.ASCIIEncoding
			write-verbose "Loading $dsFullPath"
			$tmpHt = ConvertFrom-StringData ([system.io.file]::ReadAllText($dsFullPath,$encoding))
			if (!$?) {
				throw "Error encoding hash table from file $dsFullPath"
			} else {
				$desiredStateConfig = Merge-Hashtables -refHt $desiredStateConfig -newHt $tmpHt -overrideValues
				if (!$?) {
					throw "Issue merging hashtables."
				}
			}
		}
	}
	write-verbose "HT Loaded:"
	write-verbose ($desiredStateConfig | ft | out-string)
	$desiredStateConfig
}


function Merge-Hashtables {
param(
	[Parameter(Mandatory=$true)]
	[hashtable]
	$refHt,
	[Parameter(Mandatory=$true)]
	[hashtable]
	$newHt,
	[switch]
	$overrideValues
)
	[hashtable]$ret = $refHt.Clone()
	write-verbose "Merging hashtables"
	write-verbose "Checking new HT against ref HT"
	foreach ($item in $newHt.GetEnumerator()) {
		write-verbose "[new ht] $($item.key) = $($item.value)"
		if ($ret.Contains($item.Key)) {
			write-verbose "Ref hash table already contains key."
			if ($overrideValues) {
				write-verbose "Overriding to new value $($item.Value)"
				$ret.Remove($item.Key)
				$ret.Add($item.Key,$item.Value)
			} else {
				write-verbose "Ignoring existing key."
			}
		} else {
			write-verbose "Key does not exist in ref HT. Adding key."
			$ret.Add($item.Key,$item.Value)
		}
	}
	$ret
}

function Compare-HashtableChanges {
param(
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[hashtable]
	$currHt,
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[hashtable]
	$desiredHt
)
	write-verbose "Checking HT to desired state"
	[string[]]$htChanges = $()

		foreach ($item in $desiredHt.GetEnumerator()) {
			$itemKey = $null
			$itemValue = $null
			$itemKey = $item.Key
			$itemValue = $item.Value
			write-verbose "[ds] $itemKey = $itemValue"
			if ($currHt.ContainsKey($itemKey)) {
				
				$currItemValue = $currHt[$itemKey]
				write-verbose "[ht] $itemKey = $currItemValue"
				if ($itemValue -ne $currItemValue) {
					write-verbose "CHANGE - DS value != HT value"
					$htChanges += "$itemKey = $itemValue [CHANGE FROM $($currItemValue)]"
				} else {
					write-verbose "OK - DS value == HT value"
				}
			} else {
				write-verbose "NEW - Key/Value"
				$htChanges += "$itemKey = $itemValue [NEW KEY]"
			}
		} # foreach $item

	$htChanges
}

#------------------------------------------
# VMWare
#------------------------------------------
#-------------------------------
# Stun unstun cycle
# http://purple-screen.com/?p=307
#-------------------------------
function Start-VMStun {
[CmdletBinding(SupportsShouldProcess,ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$false,HelpMessage="VM Object")]
    [ValidateNotNullOrEmpty()]
    [VMware.VimAutomation.ViCore.Impl.V1.Inventory.VirtualMachineImpl[]]
    $vm,
    [Parameter(Mandatory=$false)]
    [switch]$useTasks
)
    PROCESS {
        Foreach ($vmObj in $vm) {
            if ($PsCmdlet.ShouldProcess($vm, "Stun-Cycle VM")) {
                write-verbose "Performing stun/unstun cycle via same-host migration of $($vm.name) on $($vm.runtime.host)"
                $vmObjView = $vmObj | Get-View
                if ($useTasks) {
                	write-verbose "async stun (tasks)"
                	$task = $null
                	$viTask = $null
                	$task = $vmObjView.MigrateVM_Task($null, $vmObj.Runtime.Host, 'highPriority', $null)
                	if ($task) {
                		$viTask = Get-Task -Id "Task-$($task.Value)"
                	}
                } else {
                	write-verbose "blocking stun"
                	$vmObjView.MigrateVM($null, $vmObj.Runtime.Host, 'highPriority', $null) | write-verbose
                }
                if (!$?) {
                    write-error "Error with migrate_VM() task"
                }
            }
            new-object psobject -property @{
            	vitask = $viTask
            }
        }
    }
}

#------------------------------------------
# VMWare GET
#------------------------------------------

function Get-VMAdvancedConfiguration {
param(
	[Parameter(Mandatory=$true,ValueFromPipeline=$true)]
	[ValidateNotNullOrEmpty()]
	$vm,
	[String]
	$key
)
  PROCESS {
	if ($key) {
		$VM | Foreach {
			$_.ExtensionData.Config.ExtraConfig | Select * -ExcludeProperty DynamicType, DynamicProperty | Where { $_.Key -eq $key }
		}
	} Else {
		$VM | Foreach {
				$_.ExtensionData.Config.ExtraConfig | Select * -ExcludeProperty DynamicType, DynamicProperty
			}
	}
  }
}

#------------------------------------------
# VMWare Rollbacks & Backup
#------------------------------------------
function Backup-VmVmxFile {
param(
	[Parameter(Mandatory=$true,ValueFromPipeline=$true)]
	[ValidateNotNullOrEmpty()]
	[VMware.VimAutomation.ViCore.Impl.V1.Inventory.VirtualMachineImpl[]]
	$vms,
	[Parameter(Mandatory=$false)]
	[ValidateNotNullOrEmpty()]
	[string]
	$saveFolder
)
	BEGIN {
		if (!(Test-Path $saveFolder)) {
			write-error "Cannot find folder $saveFolder"
			return $null
		}
	}
	PROCESS {
		foreach ($vm in $vms) {
			$vmView = $vm | Get-View
			$vmxFile = $vmView.Config.Files.VmPathName
			$dsname = $vmxFile.split(" ")[0].TrimStart("[").TrimEnd("]")
			$vmxPath = $vmxFile.split(']')[1].TrimStart(' ')
			write-verbose "BACKUP VM : VMX File = $vmxFile, $dsName\$vmxPath"
			
			# mount drive
			write-verbose "Mounting datastore: $dsname"
			Remove-PSDrive -Name tmpVmMnt -ErrorAction silentlycontinue
			$VmPsDrive = new-psdrive -name tmpVmMnt -location (get-datastore $dsname) -psprovider vimdatastore -root ‘/’ | out-null
			
			$vmxfileName = $vmxPath.split('/')[1]
			$destVmx = ("$saveFolder\") + $vmxfileName
			# copy .vmx
			write-verbose "Copying VMX to $destVmx"
			Copy-DatastoreItem -Item tmpVmMnt:\$vmxPath -Destination $destVmx -Force
			if ($? -and (Test-Path $destVmx)) {
				new-object psobject -property @{
					'vmxFilePath' = $destVmx
					'vmxFileName' = $vmxfileName
				}
			} else {
				write-error "Error copying VMX file $vmxFileName"
			}
			$VmPsDrive | Remove-PSDrive -ErrorAction silentlycontinue
		}
	}
}


function New-VMHostRollback {
param(
	[Parameter(Mandatory=$true,ValueFromPipeline=$true)]
	[ValidateNotNullOrEmpty()]
	[VMware.VimAutomation.ViCore.Impl.V1.Inventory.VMHostImpl[]]
	$vmHosts,
	[ValidateNotNullOrEmpty()]
	[hashtable]
	$ds,
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]	
	[string]
	$saveFolder,
	[Parameter(Mandatory=$false)]
	[hashtable]
	$hostCurrConfigHt=@{}
)
	PROCESS {
		Foreach ($vmhost in $vmhosts) {
			$vmUuid = $null
			$vmUuid = $vmHost | Get-vmHostUniqueId
			
			if (!$hostCurrConfigHt) {
				write-verbose "Populating current settings"
				$hostCurrConfig = $vmhost | Get-AdvancedSetting
				$hostCurrConfig | % {
					$hostCurrConfigHt.Add($_.Name,$_.Value)
				}
			}
			$rollbackHt = Get-HashtableRollback -currHt $hostCurrConfigHt -desiredHt $ds
			
			$fileName = $null
			$filePath = $null
			$fileName = "$($vmUuid)_$($vmhost.name)" | Remove-InvalidFileNameChars
			$filePath = "$saveFolder\$($fileName).txt"
			write-verbose "Writing rollback to $filePath"
			$keyValueStr = Write-HashtableToFile -ht $rollbackHt -path $filePath
			new-object psobject -property @{
				'rollbackFilePath' = $filePath
				'rollbackFileName' = $fileName
				'rollbackValues' = $keyValueStr
			}
		}
	}
}

function New-VMRollback {
param(
	[Parameter(Mandatory=$true,ValueFromPipeline=$true)]
	[ValidateNotNullOrEmpty()]
	[VMware.VimAutomation.ViCore.Impl.V1.Inventory.VirtualMachineImpl[]]
	$vms,
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[hashtable]
	$ds,
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[string]
	$saveFolder,
	[Parameter(Mandatory=$false)]
	[hashtable]
	$vmCurrentConfigHt=@{}
)
	PROCESS {
		Foreach ($vm in $vms) {
			$vmUuid = $null
			$vmUuid = $vm | Get-VMUniqueId
			if (!$vmCurrentConfigHt) {
				write-verbose "Populating current settings"
				$vmCurrentConfig = $vm | Get-VMAdvancedConfiguration
				[hashtable]$vmCurrentConfigHt = @{}
				$vmCurrentConfig | % {
					$vmCurrentConfigHt.Add($_.Key,$_.Value)
				}
			}
			$rollbackHt = Get-HashtableRollback -currHt $vmCurrentConfigHt -desiredHt $ds
			
			$fileName = $null
			$filePath = $null
			$fileName = "$($vm.name)" | Remove-InvalidFIleNameChars
			$filePath = "$saveFolder\$($fileName).txt"
			write-verbose "Writing rollback to $filePath"
			$keyValueStr = Write-HashtableToFile -ht $rollbackHt -path $filePath			
			new-object psobject -property @{
				'rollbackFilePath' = $filePath
				'rollbackFileName' = $fileName
				'rollbackValues' = $keyValueStr
			}
		}
	}
}

function New-VswitchRollback {
param(
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]	
	[VMware.VimAutomation.ViCore.Impl.V1.Host.Networking.VirtualSwitchImpl[]]
	$vSwitches,
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[hashtable]
	$ds,
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[string]
	$saveFolder,
	[Parameter(Mandatory=$false)]
	[hashtable]
	$vSwitchSecPolHt=@{}
)
	PROCESS {
		foreach ($vSwitch in $vSwitches) {
			if ($vSwitchSecPolHt) {
				write-verbose "Populating current settings"
				$vSwitchSecPol = $vSwitch | Get-SecurityPolicy
				$vSwitchSecPolHt = Convert-ObjPropertiesToHashtable $vSwitchSecPol
			}
			$rollbackHt = Get-HashtableRollback -currHt $vSwitchSecPolHt -desiredHt $ds

			$vmUuid = $null
			$vmUuid = $vSwitch | Get-vSwitchUniqueId
			$fileName = "$($vmUuid)_$($vSwitch.name)" | Remove-InvalidFileNameChars
			$filePath = "$saveFolder\$($fileName).txt"
			write-verbose "Writing rollback to $filePath"
			$keyValueStr = Write-HashtableToFile -ht $rollbackHt -path $filePath
			new-object psobject -property @{
				"rollbackFilePath" = $filePath
				'rollbackFileName' = $fileName
				'rollbackValue' = $keyValueStr
			}
		}
	}
}

function New-VportgroupRollback {
param(
	[Parameter(Mandatory=$true,ValueFromPipeline=$true)]
	[ValidateNotNullOrEmpty()]
	[VMware.VimAutomation.ViCore.Impl.V1.Host.Networking.VirtualPortGroupImpl]
	$vPortgroups,
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[hashtable]
	$ds,
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[string]
	$saveFolder,
	[Parameter(Mandatory=$false)]
	[hashtable]
	$vPortgroupSecPolHt=@{}
)
	PROCESS {
		foreach ($vPortgroup in $vPortGroups) {
			if (!$vPortgroupSecPolHt) {
				write-verbose "Populating current settings"
				$vPortgroupSecPol = $vPortgroup | Get-SecurityPolicy
				$vPortgroupSecPolHt = Convert-ObjPropertiesToHashtable $vPortgroupSecPol
			}
			$rollbackHt = Get-HashtableRollback -currHt $vPortgroupSecPolHt -desiredHt $ds

			$vmUuid = $null
			$fileName = $null
			$filePath = $null
			$vmUuid = $vPortGroup | Get-vPortgroupUniqueId
			$fileName = "$($vmUuid)_$($vPortgroup.name)" | Remove-InvalidFileNameChars
			$filePath = "$saveFolder\$($fileName).txt"
			write-verbose "Writing rollback to $filePath"
			$keyValueStr = Write-HashtableToFile -ht $rollbackHt -path $filePath
			new-object psobject -property @{
				'rollbackFilePath' = $filePath
				'rollbackFileName' = $fileName
				'rollbackValue' = $keyValueStr
			}
		}
	}
}

#------------------------------------------
# VMWare Apply Desired State
#------------------------------------------

function Update-VMHostAdvancedSettings {
[CmdletBinding()]
param(
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[VMware.VimAutomation.ViCore.Impl.V1.Inventory.VMHostImpl[]]
	$vmHosts,
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[hashtable]
	$ds,
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]	
	[string]
	$saveDir
)
	PROCESS {
		Foreach ($vmhost in $vmHosts) {
			$tsNow = Get-Date
			$applied = $null
			$vmUuid = $vmHost | Get-vmHostUniqueId

			write-verbose "Populating current settings"
			$hostCurrConfig = $vmHost | Get-AdvancedSetting
			$hostCurrConfigHt = @{}
			$hostCurrConfig | % {
				$hostCurrConfigHt.Add($_.Name,$_.Value)
			}
			
			write-verbose "Checking delta from DS"
			$changeDelta = $null
			$changeDeltaHtml = $null
			$changeDelta = Compare-HashtableChanges -currHt $hostCurrConfigHt -desiredHt $ds

			write-verbose "Applying Adv Settings VMHost: $($vmHost.Name) UUID $vmUuid"
			if ($changeDelta) {
				$changeDeltaHtml = ($changeDelta | ConvertTo-HtmlList)
				write-verbose "Changes needed"
				
				$tsNow = Get-Date
				write-verbose "Creating a rollback"
				$rollback = $null
				$rollback = New-VMHostRollback -vmHosts $vmHost -ds $ds -saveFolder $saveDir -hostCurrConfigHt $hostCurrConfigHt
				if (!$? -or !$rollback) { 
					throw "Error creating VM host rollback for $($vmHost.name)"
				}

				write-verbose "Applying changes"
				foreach ($Option in $ds.GetEnumerator()) {
					write-verbose "Applying $($Option.Key) = $($Option.Value)"
					$vmHost | Get-AdvancedSetting -name $Option.Key | Set-AdvancedSetting -Value $Option.Value -Confirm:$false | write-verbose
				}
				$applied = $true
			} else {
				write-verbose "No Changes Needed"
				$applied = $false
			}	
			 new-object psobject -property @{
				type = "vmhost"
				name = $vmHost.Name
				vmUuid = $vmUuid
				uName = "$($vmUuid)_$($vmHost.name)"
				applied = $applied
				deltaArray = $changeDelta
				delta = $changeDeltaHtml
				rollback = $rollback
				scanTs = $tsNow
				scanTz = [TimeZoneInfo]::Local
			}
		}
	}
}

function Update-VMAdvancedConfiguration {
[CmdletBinding()]
param(
	[Parameter(Mandatory=$true,ValueFromPipeline=$true)]
	[ValidateNotNullOrEmpty()]
	[VMware.VimAutomation.ViCore.Impl.V1.Inventory.VirtualMachineImpl[]]
	$vms,
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[hashtable]
	$ds,
	[switch]
	$useTasks,
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[string]
	$saveDir
)
	PROCESS {
		foreach ($vm in $vms) {
			$tsNow = Get-Date
			$applied = $null
			$vmUuid = $null
			$vmUuid = $vm | Get-VMUniqueId
			write-verbose "Setting Adv Settings VM: $($vm.Name) UUID: $vmUuid"
			
			write-verbose "Populating current settings"
			$vmCurrentConfig = $vm | Get-VMAdvancedConfiguration
			[hashtable]$vmCurrentConfigHt = @{}
			$vmCurrentConfig | % {
				$vmCurrentConfigHt.Add($_.Key,$_.Value)
			}
			
			write-verbose "Checking delta from DS"
			$changeDelta = $null
			$changeDeltaHtml = ""
			$changeDelta = Compare-HashtableChanges -currHt $vmCurrentConfigHt -desiredHt $ds
			
			if ($changeDelta) {
				$changeDeltaHtml = ($changeDelta | ConvertTo-HtmlList)
				write-verbose "Chaanges needed"

				$folderName = "$($vmUuid)_$($vm.name)" | Remove-InvalidFileNameChars
				$saveFolder = "$saveDir\$folderName\"
				md $saveFolder -force | out-null

				write-verbose "Creating rollback"
				$rollback = $null
				$rollback = New-VMRollback -vms $vm -ds $ds -saveFolder $saveFolder -vmCurrentConfigHt $vmCurrentConfigHt
				if (!$? -or !$rollback) {
					throw "Error creating VM Rollback for $($vm.Name)"
				}

				write-verbose "Copying VMX File"
				$vmxFileCopy = $null
				$vmxFileCopy = Backup-VmVmxFile -vm $vm -saveFolder $saveFolder
				if (!$? -or !$vmxFileCopy) {
					throw "Error copying VMX file for $($vm.Name)"
				}

				# Construct the vm spec
				write-verbose "Applying changes"
				write-verbose "Constructing VMSpec"
				$vmConfigSpec = New-Object VMware.Vim.VirtualMachineConfigSpec
				Foreach ($Option in $ds.GetEnumerator()) {
					$OptionValue = New-Object VMware.Vim.optionvalue
					$OptionValue.Key = $Option.Key
					$OptionValue.Value = $Option.Value
					$vmConfigSpec.extraconfig += $OptionValue
				}

				$vmView = $vm | Get-View

				if ($useTasks) {
					write-verbose "Applying Changes (async)"
					$task = $null
					$viTask = $null
					$task = $vmView.ReconfigVM_Task($vmConfigSpec)
					if ($task) {
						$viTask = Get-Task -Id "Task-$($task.Value)"
					}
				} else {
					write-verbose "Applying Changes (blocking) ..."
					$vmView.ReconfigVM($vmConfigSpec) | write-verbose
				}
				$applied = $true
			} else {
				write-verbose "No Changes Needed"
				$applied = $false
			}

			new-object psobject -property @{
				type = "vm"
				name = $vm.Name
				vmUuid = $vmUuid
				uName = "$($vmUuid)_$($vm.Name)"
				applied = $applied
				deltaArray = $changeDelta
				delta = $changeDeltaHtml
				rollback = $rollback
				rollbackVmx = $vmxFileCopy
				vitask = $viTask
				scanTs = $tsNow
				scanTz = [TimeZoneInfo]::Local
			}
		}
	}
}


function Update-vSwitchSettings {
[CmdletBinding()]
param(
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]	
	[VMware.VimAutomation.ViCore.Impl.V1.Host.Networking.VirtualSwitchImpl[]]
	$vSwitches,
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[hashtable]
	$ds,
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[string]
	$saveDir
)
	PROCESS {
		foreach ($vSwitch in $vSwitches) {
			$tsNow = Get-Date
			$applied = $null
			$vmUuid = $null
			$vmUuid = $vSwitch | Get-vSwitchUniqueId
			write-verbose "vSwitch: $($vsiwtch.name) UUID: $vmUuid"
			
			write-verbose "Populating current settings"
			$vSwitchSecPol = $vSwitch | Get-SecurityPolicy
			$vSwitchSecPolHt = Convert-ObjPropertiesToHashtable $vSwitchSecPol

			write-verbose "Checking delta from DS"
			$changeDelta = $null
			$changeDeltaHtml = ""
			$changeDelta = Compare-HashtableChanges -currHt $vSwitchSecPolHt -desiredHt $ds

			if ($changeDelta) {
				$changeDeltaHtml = ($changeDelta | ConvertTo-HtmlList)
				write-verbose "Changes needed"

				write-verbose "Creating rollback"
				$rollback = $null
				$rollback = New-VswitchRollback -vSwitches $vSwitch -saveFolder $saveDir -ds $ds -vSwitchSecPolHt $vSwitchSecPolHt
				if (!$? -or !$rollback) { 
					throw "Error creating vSwitch rollback for $($vSwitch.name)"
				}

				write-verbose "Converting HT to SPLAT-able HT"
				$dsSplat = Convert-HashtableValuesToTypes $ds
				
				write-verbose "Setting security policies"
				$vSwitchSecPol | Set-SecurityPolicy @dsSplat | write-verbose
				$applied = $true
			} else {
				write-verbose "No Changes Needed"
				$applied = $false
			}

			new-object psobject -property @{
				type = "vswitch"
				name = $vswitch.Name
				vmUuid = $vmUuid
				uName = "$($vmUuid)_$($vSwitch.Name)"
				applied = $applied
				deltaArray = $changeDelta
				delta = $changeDeltaHtml
				rollback = $rollback
				scanTs = $tsNow
				scanTz = [TimeZoneInfo]::Local
			}
		}
	}
}

function Update-vPortgroupSettings {
[CmdletBinding()]
param(
	[Parameter(Mandatory=$true,ValueFromPipeline=$true)]
	[ValidateNotNullOrEmpty()]
	[VMware.VimAutomation.ViCore.Impl.V1.Host.Networking.VirtualPortGroupImpl]
	$vPortgroups,
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[hashtable]
	$ds,
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[string]
	$saveDir
)
	PROCESS {
		foreach ($vPortgroup in $vPortGroups) {
			$tsNow = Get-Date
			$applied = $null
			$vmUuid = $null
			$vmUuid = $vPortgroup | Get-vPortgroupUniqueId
			write-verbose "vPortGroup: $($vportgroup.name) UUID: $vmUuid"
			
			write-verbose "Populating current settings"
			$vPortgroupSecPol = $vPortgroup | Get-SecurityPolicy
			$vPortgroupSecPolHt = Convert-ObjPropertiesToHashtable $vPortgroupSecPol
			
			write-verbose "Checking delta from DS"	
			$changeDelta = $null
			$changeDeltaHtml = $null
			$changeDelta = Compare-HashtableChanges -currHt $vPortgroupSecPolHt -desiredHt $ds

			if ($changeDelta) {
				$changeDeltaHtml = ($changeDelta | ConvertTo-HtmlList)
				write-verbose "Changes needed"

				write-verbose "Creating rollback"
				$rollback = $null
				$rollback = New-VportgroupRollback -vPortGroups $vPortgroup -saveFolder $saveDir -ds $ds -vPortgroupSecPolHt $vPortgroupSecPolHt
				if (!$? -or !$rollback) { 
					throw "Error creating vPortGroup rollback for $($vPortgroup.name)"
				}

				write-verbose "Converting HT to SPLAT-able HT"
				$dsSplat = Convert-HashtableValuesToTypes $ds
				
				write-verbose "Setting security policies"
				$vPortgroupSecPol | Set-SecurityPolicy @dsSplat | write-verbose
				$applied = $true
			} else {
				$applied = $false
			}
			new-object psobject -property @{
				type = "vportgroup"
				name = $vportgroup.Name
				vmUuid = $vmUuid
				uName = "$($vmUuid)_$($vportgroup.Name)"
				applied = $applied
				deltaArray = $changeDelta
				delta = $changeDeltaHtml
				rollback = $rollback
				scanTs = $tsNow
				scanTz = [TimeZoneInfo]::Local
			}
		}
	}
}

#------------------------------------------
# VMWare UniqueIDs
#------------------------------------------

function Get-VMUniqueId {
param(
	[Parameter(Mandatory=$true,ValueFromPipeline=$true)]  
	[VMware.VimAutomation.ViCore.Impl.V1.Inventory.VirtualMachineImpl]
	$vm
)
	PROCESS {
		# Uniquely identifying VMs
		# vCenter UUID + MoRef
		# http://blogs.vmware.com/vsphere/tag/moref
		$vm | % {
			$vmView = $_ | Get-View
			$vmVcenterUuid = $vmView.config.uuid
			$vmMoRef = ($vmView | select -expand moref).Value
			$vmUniqueId = "$vmVcenterUuid-$vmMoRef"

			$vmUniqueId
		}
	}
}

function Get-vmHostUniqueId {
param(
	[Parameter(Mandatory=$true,ValueFromPipeline=$true)]
	[VMware.VimAutomation.ViCore.Impl.V1.Inventory.VMHostImpl]
	$vmHost
)
	PROCESS {
		$vmHost | % {
			$hostView = $_ | Get-View
			$hostClusterUuid = $hostView.Hardware.SystemInfo.uuid
			$hostMoRef = ($hostView | select -expand moref).Value
			$hostUniqueId = "$hostClusterUuid-$hostMoRef"
			
			$hostUniqueId
		}
	}
}

function Get-vSwitchUniqueId {
param(
	[Parameter(Mandatory=$true,ValueFromPipeline=$true)]
	[VMware.VimAutomation.ViCore.Impl.V1.Host.Networking.VirtualSwitchImpl]
	$vSwitch
)
	PROCESS {
		$vSwitch | % {
			$vmHostUniqueId = $_.VMHost | Get-vmHostUniqueId
			$vSwitchUniqueId = "$vmHostUniqueId-$($_.name)"
			
			$vSwitchUniqueId
		}
	}
}

function Get-vPortgroupUniqueId {
param(
	[Parameter(Mandatory=$true,ValueFromPipeline=$true)]
	[VMware.VimAutomation.ViCore.Impl.V1.Host.Networking.VirtualPortGroupImpl]
	$vPortgroup
)
	PROCESS {
		$vPortgroup | % {
			$vSwitch = Get-VirtualSwitch -Name $vPortgroup.VirtualSwitchName
			$vSwitchUniqueId = ($vSwitch | Get-vSwitchUniqueId)
			$vPortgroupUniqueId = "$vSwitchUniqueId-$($vPortgroup.name)"
			$vPortgroupUniqueId
		}
	}
}

#------------------------------------------
# VMWare Reporting
#------------------------------------------

function Test-WebServerSSL {
# Function original location: http://en-us.sysadmins.lv/Lists/Posts/Post.aspx?List=332991f0-bfed-4143-9eea-f521167d287c&ID=60
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 0)]
    [string]
    $URL,
    [Parameter(Position = 1)]
    [ValidateRange(1,65535)]
    [int]
    $Port = 443,
    [Parameter(Position = 2)]
    [Net.WebProxy]
    $Proxy,
    [Parameter(Position = 3)]
    [int]
    $Timeout = 15000,
    [switch]
    $UseUserContext
)
Add-Type @"
using System;
using System.Net;
using System.Security.Cryptography.X509Certificates;
namespace PKI {
    namespace Web {
        public class WebSSL {
            public Uri OriginalURi;
            public Uri ReturnedURi;
            public X509Certificate2 Certificate;
            //public X500DistinguishedName Issuer;
            //public X500DistinguishedName Subject;
            public string Issuer;
            public string Subject;
            public string[] SubjectAlternativeNames;
            public bool CertificateIsValid;
            //public X509ChainStatus[] ErrorInformation;
            public string[] ErrorInformation;
            public HttpWebResponse Response;
        }
    }
}
"@
    $ConnectString = "https://$url`:$port"
    $WebRequest = [Net.WebRequest]::Create($ConnectString)
    $WebRequest.Proxy = $Proxy
    $WebRequest.Credentials = $null
    $WebRequest.Timeout = $Timeout
    $WebRequest.AllowAutoRedirect = $true
    [Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
    try {$Response = $WebRequest.GetResponse()}
    catch {}
    if ($WebRequest.ServicePoint.Certificate -ne $null) {
        $Cert = [Security.Cryptography.X509Certificates.X509Certificate2]$WebRequest.ServicePoint.Certificate.Handle
        try {$SAN = ($Cert.Extensions | Where-Object {$_.Oid.Value -eq "2.5.29.17"}).Format(0) -split ", "}
        catch {$SAN = $null}
        $chain = New-Object Security.Cryptography.X509Certificates.X509Chain -ArgumentList (!$UseUserContext)
        [void]$chain.ChainPolicy.ApplicationPolicy.Add("1.3.6.1.5.5.7.3.1")
        $Status = $chain.Build($Cert)
        New-Object PKI.Web.WebSSL -Property @{
            OriginalUri = $ConnectString;
            ReturnedUri = $Response.ResponseUri;
            Certificate = $WebRequest.ServicePoint.Certificate;
            Issuer = $WebRequest.ServicePoint.Certificate.Issuer;
            Subject = $WebRequest.ServicePoint.Certificate.Subject;
            SubjectAlternativeNames = $SAN;
            CertificateIsValid = $Status;
            Response = $Response;
            ErrorInformation = $chain.ChainStatus | ForEach-Object {$_.Status}
        }
        $chain.Reset()
        [Net.ServicePointManager]::ServerCertificateValidationCallback = $null
    } else {
        Write-Error $Error[0]
    }
}


function Test-VcenterNfcSsl {
param(
	[Parameter(Mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[string[]]
	$vCenters
)	

	foreach ($vCenter in $vCenters) {
		# nfc ssl
		# Check Network File Copy NFC uses SSL
		$vpxdPath = "\\$vCenter\C$\ProgramData\VMware\VMware VirtualCenter\vpxd.cfg"

		$nfcSsl = $null
		if (Test-Path $vpxdPath) {
			
			[XML]$file = Get-Content "\\$vCenter\C$\ProgramData\VMware\VMware VirtualCenter\vpxd.cfg"
			
			if ($file.config.nfc.Usessl) { 
				$nfcSsl = $true
			} Else { 
				$nfcSsl = $false
			}
		}
		new-object psobject -Property @{
			'vCenter' = $vCenter
			'nfcSsl' = $nfcSsl
		}
	}
}
