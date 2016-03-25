function New-HtmlReportHead {
param(
	[string]$moreStyle=""
)
$docType = '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"> <html xmlns="http://www.w3.org/1999/xhtml">'
$html_top = @"
$docType
<head>
<Title>$report_title</Title>
<style>
	* {
		font-family: Calibri, Verdana;
	}
	h1 { font-size: 350%; color: #4682b4 }
	h2 { font-size: 250%; color: #4876ff; }
	h3 { font-size: 150%; color: #4f94cd; }
	TABLE {width: 1280px; border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
	TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #4876ff; color:#ffffff; }
	TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black; vertical-align: top; }
	li { margin-left:10px; padding: 0px; }
	ul { margin-left:10px; padding: 0px; }
	ol { margin-left:10px; padding: 0px; }
	code { 
		font-family: 'Courier New';
		background-color:#DDDDDD;
		background-clip: padding-box;
		margin: 2px 2px;
		padding: 2px 2px; 
	}
	blockquote {
	  background: #f9f9f9;
	  border-left: 10px solid #ccc;
	  margin: 1.5em 10px;
	  padding: 0.5em 10px;
	  width: 1024px;
	}
	blockquote p {
	  display: inline;
	}
	.even { background-color:#eeeeee; }
	.odd  { background-color:#dddddd; }
	.good { background-color:#BADA55; }
	.bad { background-color:#990000; color:#e0b2b2; }
	.mark-row { background-color:#999900; }
	$moreStyle
</style>
</head><body>
"@

$html_top
}

function New-HtmlReportBottom {
param (
	[switch]$extraFun,
	[string]$pre
)
	$url = ""
	if ($extraFun) {
		switch ([int](get-date).dayofweek) {
			1 {
				$url = 'http://youtu.be/hkDD03yeLnU'
				break
			}
			2 {
				$url = 'http://youtu.be/nkLtXfsPqVQ'
				break
			}
			3 {
				$url = 'http://youtu.be/O2rGTXHvPCQ'
				break
			}
			4 {
				$url = 'http://youtu.be/h0ZgED70FMg'
				break
			}
			5 {
				$url = 'http://youtu.be/kfVsfOSbJY0'
				break
			}
			6 {
				$url = 'http://youtu.be/BIUQw1w5OqM'
				break
			}
			7 {
				$url = 'http://youtu.be/P9dpTTpjymE'
				break
			}
			default {
				$url = 'http://youtu.be/8To-6VIJZRE'
			}
		}
		$extraFunStr = "<br />It's <a target='_new' href='$($url)'>$((get-date).dayofweek)</a>"
	}
	
$html_bottom = @"
$pre
$extraFunStr
<br /></body></html>
"@
$html_bottom
}

# http://community.spiceworks.com/scripts/show/1745-set-alternatingrows-function-modify-your-html-table-to-have-alternating-row-colors
# http://community.spiceworks.com/scripts/show/2450-change-cell-color-in-html-table-with-powershell-set-cellcolor


Function Format-PrettyTable {
    [CmdletBinding()]
   	Param(
        [Parameter(Mandatory,ValueFromPipeline)]
        [Object[]]$InputObject,
       
	   [Parameter(Mandatory)]
	   [Object[]]$sourceObj,
	   
		[psobject[]]$colorFilters,
		[string]$CSSEvenClass='even',
		[string]$CSSOddClass='odd'
   	)
	BEGIN {
		write-verbose "--- PrettyTable Running ----"
		
		# Even/Odd Coloring
		write-verbose "Initializing even/odd coloring"
		$ClassName = $CSSEvenClass
				
		$colorFiltersActive = @()
		
		# Table Header Lookup
		$Filter = $null
		If ($colorFilters) {
			write-verbose "Processing Color Filters"	
			Foreach ($cf in $colorFilters) {
				write-verbose "Processing filter: $($cf | fl | out-string)"
			
				# basic filter validation
				$validFilter = $false
				if ($cf.markType -eq "cell") {
					write-verbose "Processing filter for ""$($cf.property)"""
					If ($cf.property) { 
						$validFilter = $true
					} else {
						write-error "colorFilter markType is ""cell"", the ""property"" field must be set."
					}
				} elseif ($cf.markType -eq "row") {
					if ($cf.property) {
						write-warning "NOTE that when markType is ""row"" the property field is IGNORED."
					}
					$validFilter = $true
				} else {
					write-error "No valid markType specified for colorFilter"
				}
				
				# add filter
				if ($validFilter) {
					write-verbose "Adding filter."
					
					write-verbose "Processing FILTER: $Filter"
					Try {
						[scriptblock]$Filter = [scriptblock]::Create($cf.filter)
					}
					Catch {
						$_
						throw "Cannot process filter $($cf.filter)" 
					}

					write-verbose "Filter OK. Adding to active filters."
					$colorFiltersActive += new-object psobject -property @{
						markType = $cf.markType
						filter = $Filter
						property = $cf.property
						class = $cf.class
						index = -1
					}
				} else {
					write-warning "Invalid filter not added."
				}
			}
        } else {
			write-verbose "Color filtering off"
		}

		$propertyFilterLookup = @{}
		$rowFilterLookup = @()
		$headerLookup = @{}
		$trNum = 0
	}
	PROCESS {
		
		Foreach ($Line in $InputObject) {
			write-verbose "INPUT: $Line"
			write-verbose "--------------"

			$newLine = "$($Line)`n"
			$trClass = ""
			$tdMiddle = ""
		
			# Even/Odd Coloring
			If ($Line.Contains("<tr><td>")) {
				write-verbose "Setting TR Class (Even/Odd) to $ClassName"
				$trClass = $ClassName
				If ($ClassName -eq $CSSEvenClass) {
					$ClassName = $CSSOddClass
				} else {
					$ClassName = $CSSEvenClass
				}
			}			
						
			# Table Headers
			If ($colorFiltersActive -and ($Line.IndexOf("<tr><th") -ge 0))
			{
				Write-Verbose "PARSING Table Headers & Constructing Filter Lookup Tables"
				$Search = $Line | Select-String -Pattern '<th ?[a-z\-:;"=]*>(.*?)<\/th>' -AllMatches

				# Property Table Filter Lookup
				write-verbose "CONSTRUCTING property filter lookup table"
				$i = 0
				Foreach ($Match in $Search.Matches) {
					$headerName = $Match.Groups[1].Value
					write-verbose "---- Processing Header ""$headerName"" Index $i"
					$matchingPropertyNames = $colorFiltersActive | ? { ($_.markType -eq 'cell') -and ($_.property -eq $headerName) }
					if ($matchingPropertyNames) {
						write-verbose "------ Adding color filters with property matching ""$headerName"""
						$propertyFilterLookup.Add($i,$matchingPropertyNames)
					} else {
						write-verbose "------ No matching properties for this header."
					}
					write-verbose "ADDING $i $headerName to header lookup table"
					$headerLookup.Add($i,$headerName)
					$i++
				}
				# Row lookups
				write-verbose "CONSTRUCTING row lookup table"
				$rowFilterLookup = $colorFiltersActive | ? { $_.markType -eq 'row' }
				
				# results
				write-verbose "CONSTRUCTED tables"
				write-verbose "-- Property Lookup Table"
				write-verbose ($propertyFilterLookup | fl | out-string)
				write-verbose "-- Row Lookup Table"
				write-verbose ($rowFilterLookup | fl | out-string)
				write-verbose "-- Header lookup table"
				write-verbose ($headerLookup | fl | out-string)
			}
						
			# Color filters
			if ($colorFiltersActive -and ($Line -like "*<tr><td*")) {
				$newLine = ""
				$dummyObj = $null
				$dummyObjProps = @{}
												
				write-verbose "SELECTING Row Data - `$obj[$($trNum)]"
				$dummyObj = $sourceObj[$trNum]
				write-verbose ($dummyObj | ft | out-string)
				
				# Row Coloring
				if ($dummyObj) {
					write-verbose "Running Row Filters"
					$dummyObj | % {
						Foreach ($f in $rowFilterLookup) {
							write-verbose "EVAL $($f.filter)"
							if (Invoke-Command $f.filter) {
								write-verbose "-- Criteria met. Setting TR class to $($f.class)"
								$trClass = $f.class
							} else {
								write-verbose "-- Criteria not met."
							}
						}
					}
				}
				
				$newline += "<tr class=""$trClass"">"
				
				# Cell Coloring
				if ($dummyObj) {
					write-verbose "Running Cell Filters"
					for ($i=0;$i -lt $headerLookup.Count;$i++) {
						$tdClass = ""
						$applicableFilters = $null
						if ($propertyFilterLookup[$i]) {
							$dummyObj | % {
								$applicableFilters = $propertyFilterLookup[$i]
								Foreach ($f in $applicableFilters) {
									write-verbose "EVAL: $($f.filter)"
									write-verbose "handles $($_.handles)"
									if (Invoke-Command $f.filter) {
										write-verbose "-- Criteira Met. Setting TD class to $($f.class)"
										$tdClass = " class=""$($f.class)"""
									} else {
										write-verbose "-- Criteria not met."
									}
								}
							}
						}
						#$v = $tdMatches.Matches[$i].Groups[1].Value
						$cellValue = $dummyObj.($headerLookup[$i])
						$newLine += "`n`t<td$($tdClass)>$($cellValue)</td>"
					}
				}
				
				# close row
				$newLine += "`n</tr>`n"
				
				$trNum++
			} # if colorfiltering
			write-verbose "OUTPUT: $newLine"
			$newLine
			write-verbose "--------------"
		}
	}
}

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