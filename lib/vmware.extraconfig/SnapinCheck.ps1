# http://stackoverflow.com/questions/14614305/powershell-modules-and-snapins
function Validate-PSSnapin {
[CmdletBinding(DefaultParameterSetName = 'Default')]
param(
    [Parameter(ParameterSetName="Default",Mandatory=$true)]
    [Parameter(ParameterSetName="ReqVersion")]
    [ValidateNotNullOrEmpty()]
    [string]
    $name,
    [Parameter(ParameterSetName="Default",Mandatory=$false)]
    [Parameter(ParameterSetName="ReqVersion")]
    [switch]
    $important,
    [Parameter(ParameterSetName="ReqVersion",Mandatory=$true)]
    [int]$VersionMajor,
    [Parameter(ParameterSetName="ReqVersion",Mandatory=$true)]
    [int]$versionMinor,
    [Parameter(ParameterSetName="ReqVersion",Mandatory=$false)]
    [int]$VersionBuild,
    [Parameter(ParameterSetName="ReqVersion",Mandatory=$false)]
    [int]$VersionRevision
)
	if (Get-PSSnapin $name -Registered -ErrorAction SilentlyContinue) {
        if ($PsCmdlet.ParameterSetName -eq "ReqVersion") {
            $targetVersion = ""
            $targetversion += "$($versionMajor).$($versionMinor)"
            if ($versionBuild -ne $null) {
                $targetVersion += " Build $versionBuild"
            }
            if ($VersionRevision -ne $null) {
                $targetVersion += " Revision $VersionRevision"
            }
            $snappinVersion = Get-PSSnapin $name -Registered | Select -expand Version
            $versionOk = $false
            if ($snappinVersion.major -ge $versionMajor) {
                if ($snappinVersion.minor -ge $versionMinor) {
                    if ($versionBuild -ne $null) {
                        if ($snappinVersion.Build -ge $versionBuild) {
                            if ($VersionRevision -ne $null) {
                                if ($snappinVersion.Revision -ge $VersionRevision) {
                                    $versionOk = $true
                                }
                            } else {
                                $versionOk = $true
                            }
                        }
                    } else {
                        $versionOk = $true
                    }
                }
            }
            if (!$versionOk) {
                write-warning "Your $name version is $snappinVersion whereas the module was authored against version $targetVersion"
            } else {
                write-verbose "$name version $snappinVersion OK (Target = $targetVersion)"
            }
        }
        <#
		if (!(Get-PSSnapin $name -ErrorAction SilentlyContinue)) {
            write-verbose "Loading $name"
			Add-PSSnapin $name
		}
		else {
			Write-Verbose "Skipping $name, already loaded."
		}
        #>
	} else {
		write-warning "Cannot find module $name"
        if ($important) {
            write-warning "This module DEPENDS on the functionality of this snappin. Please unload/reload this module after installing the module."
            throw "Module not found."
        }
	}
}