# Populate hardcoded objects.
# Keep the variable names the same.
$hardcoded_vmHosts = Get-VMHost
$hardcoded_vms = Get-VM
$hardcoded_vSwitches = Get-Virtualswitch
$hardcoded_vPortGroups = Get-VirtualPortGroup

# EXAMPLES backing up based on annotation (e.g. Veeam) b/s vCenter VM
# function Get-VeeamBackedUp {
#     foreach ($vm in (Get-VM)) {
#         $veeamNote = ($vm | Get-Annotation -Name 'VEEAM BACKUP OK').Value
#         if ($veeamNote) {
#             $vm | Add-Member -MemberType NoteProperty -Name 'Veeam' -Value $veeamNote
#             $vm
#         }
#     }
# }
# $hardcoded_vms = Get-VeeamBackedUp | ? { $_.name -notlike '*vcenter*' }