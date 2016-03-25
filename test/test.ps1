param(
	[switch]$all,
	[switch]$test,
	[switch]$newTestVms,
	[switch]$wipeTestVms,
	[string[]]$vCenterServers=@('vc01.contoso.local')
)

Remove-MOdule vmware.extraconfig -erroraction 0

if ($wipeTestVms -or $all) {
	write-host "Wiping test VMs" -f green
	$vmsWipe = Get-VM "test_upgrayedd*"
	if ($vmsWipe) {
		$vmsWipe | select -expand name | write-host -f yellow
		$vmsWipe | Stop-VM -Kill -Confirm:$false -ErrorAction SilentlyContinue
		start-sleep -seconds 4
		$vmsWipe | Remove-VM -DeletePermanently -Confirm:$false
	}
	if (!$?) {
		throw "Error wiping VMs"
	}
}

if ($newTestVms -or $all) {
	$newVmSuffix = (get-date -f 'yyyyMMMddHHmmss')
	$randWord = (Invoke-WebRequest -Uri 'http://randomword.setgetgo.com/get.php').Content
	if ($randWord) {
		$newVmSuffix = $randWord
	}
	$baseVm = 'upgrayedd_base_ubuntu'
	$upgradedUpgrayedd = "test_upgrayedd_$($newVmSuffix)"
	write-host "Cloning $baseVm --> $upgradedUpgrayedd" -f green
	$vmhost = get-vmhost
	$newVm = New-VM -Name $upgradedUpgrayedd -VM $baseVm -VMHost $vmhost[0]
	$newVm | Start-VM
	if (!$?) {
		throw "Error creating test Vms."
	}
}

if ($stageBadState -or $all) {
	write-host "Setting some bad vaues." -f green

	# vmHost
	get-vmhost | Get-AdvancedSetting 'UserVars.ESXiShellInteractiveTimeOut' | Set-AdvancedSetting -Value 0 -Confirm:$false

	# vSwitch
	Get-VirtualSwitch -Name vSwitch1 | Get-SecurityPolicy | Set-SecurityPolicy -AllowPromiscuous $true #-ForgedTransmits $true -MacChanges $true

	# vPortGroup
	Get-VirtualPortGroup -Name vportgroup_voltron | Get-SecurityPolicy | Set-SecurityPolicy -AllowPromiscuous $true #-ForgedTransmits $true -MacChanges $true
}

if ($test -or $all) {
	write-host "Starting Test." -f green
	$vms = get-vm test*
	& ..\vmbaseline.ps1 -manual -logging -Verbose -stunVms -vms $vms -vmhosts (get-vmhost) -vswitches (get-virtualswitch) -vportgroups (get-virtualportgroup)
	# & ..\vmbaseline.ps1 -auto -logging -Verbose -stunVms -vCenterServers $vCenterServers
}