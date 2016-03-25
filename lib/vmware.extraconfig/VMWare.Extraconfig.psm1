$currentScriptPath = Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path

. $currentScriptPath\SnapinCheck.ps1
# This will prevent the module from being loaded if the snappins aren't registered.
Validate-PSSnapin -name "vmware.vimautomation.core" -Verbose -Important -VersionMajor 5 -VersionMinor 5 -VersionBuild 0 -VersionRevision 0
write-host "Remember this EXTENDS PowerCLI so load it." -f black -b yellow

# Dot source .ps1
. $currentScriptPath\VMWare.Extraconfig.ps1

###########################################################
# Exported
###########################################################

Export-ModuleMember -function @(
	'Start-VMStun',
	'Get-VMAdvancedConfiguration',
	'Backup-VmVmxFile',
	'Update-VMHostAdvancedSettings',
	'Update-VMAdvancedConfiguration',
	'Update-vSwitchSettings',
	'Update-vPortgroupSettings',
	'Get-VMUniqueId',
	'Get-vmHostUniqueId',
	'Get-vSwitchUniqueId',
	'Get-vPortgroupUniqueId',
	'Test-WebServerSSL',
	'Test-VcenterNfcSsl',
	'Load-HashtableFromFile'
)