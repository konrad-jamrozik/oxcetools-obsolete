<#
.Synopsis A library module for oxcetools.ps1.
#>
Set-StrictMode -Off

<# 
.SYNOPSIS
Sets the keys on $item to values ("defaults") if not present.
The keys to set, the values to set, and if to set, is determined based on
$defaults.
#>
function Set-Defaults([hashtable] $item, [hashtable] $defaults) {
    foreach ($defaultName in $defaults.Keys) {
        Set-Default $item $defaultName `
            $defaults.$defaultName.value `
            $defaults.$defaultName.dependentKey
    }
}

<#
.SYNOPSIS
Sets value of $key on $item to $defaultValue, only if the value is not already present AND $dependentKey 
is present on the item.
#>
function Set-Default([hashtable] $item, [string] $key, [int] $defaultValue, [string] $dependentKey = "") {
    if (("" -eq $dependentKey) -or ($item.ContainsKey($dependentKey))) {
        if (!$item.ContainsKey($key)) {
            $item.$key = $defaultValue
        }
    }
}

<#
.SYNOPSIS
Reads the "key: value" pair present on current $line into $item.
It will read only keys present in $keysIncluded.
It will skip the line if it is not of form "key: value"
It will recognize lists of form "[a, b, c], like categories, and read them into array of strings instead of a string.
#>
function Read-KeyValue([string] $line, [string[]] $keysIncluded, [hashtable] $item) {

    $keyValuePair = $line.Split(": ")

    # Return if the current line is not a key value pair ("key: value")
    if ($keyValuePair.Count -lt 2) {
        return;
    }

    $key = $keyValuePair[0]

    if (!$keysIncluded.Contains($key)) {
        return
    }

    $valueString = $line.Substring(($key+": ").Length)
    if ($valueString.StartsWith("[")) {
        $value = $valueString.Replace("[","").Replace("]","").Split(", ")
    } else {
        $value = $valueString
    }
    $item.$key = $value
}

function Read-LineData([string] $line) {

    ($key, $valueString, $type) = if ($line.Contains(": ")) {
        $keyValuePair = $line -split ": "
        $key = $keyValuePair[0]
        ($key, $line.Substring(($key + ": ").Length), "keyValue")
    } elseif ($line.StartsWith("- ")) {
        ("-", $line.Substring("- ".Length), "listItem") 
    } elseif ($line.EndsWith(":") -and !($line -contains "- ")) {
        $key = $line -replace ":"
        ($key, "", "listHeader")
    } else {
        throw "Unexpected `$line format: $line"
    }

    $value = if ($valueString.StartsWith("[")) {
        $valueString.Replace("[", "").Replace("]", "") -split ", "
    } else {
        $valueString
    }

    return @($key, $value, $type)
}

function Read-Item([string[]] $lines, [hashtable] $item, [string[]] $categoriesIncluded, [string[]] $categoriesExcluded) {

    if ($lines.Count -lt 1) { throw "Expected for the item being read to have at least one line." }
    if (!$lines[0].TrimStart().StartsWith("- type: ")) { throw "Expected for the first line to start with '- type: '. Instead got: '$($lines[0])'." }

    $lines = $lines 
        | ForEach-Object { $_.Trim() }
        | Where-Object { !($_.StartsWith("#")) } # Remove comments

    # If categories are missing
    if ($lines.Count -lt 2) { return $null }
    if (!$lines[1].TrimStart().StartsWith("categories: ")) { return $null }

    $name = $lines[0].Substring("- type: ".Length)
    $categories = (Read-LineData $lines[1])[1]

    $nocategoriesIncludedPresent = $null -eq (Compare-Object $categories $categoriesIncluded -IncludeEqual -ExcludeDifferent -PassThru)
    if ($nocategoriesIncludedPresent) { 
        Write-Debug "No included categories ($categoriesIncluded) present in $name. Categories present: $categories"
        return $null 
    }
    $excludedCategoriesPresent = $null -ne (Compare-Object $categories $categoriesExcluded -IncludeEqual -ExcludeDifferent -PassThru)
    if ($excludedCategoriesPresent) { 
        Write-Debug "Excluded categories ($categoriesExcluded) present in $name. Categories present: $categories"
        return $null 
    }

    $lineData = $lines[2..$lines.Length] | ForEach-Object { ,@(Read-LineData $_) }

    $ht = @{}

    for ($i = 0; $i -lt $lineData.Count; $i++) {
        ($key, $value, $type) = $lineData[$i]

        if ($type -eq "keyValue") {
            $ht += @{ $key = $value }
        }
    }

    return $ht
}

# TODO handle special ammo cases. See e.g. STR_HUMAN_SONIC_HEAVY_CANNON
function Read-ItemsDebug()
{
    $StartTime = Get-Date
    [string[]] $categoriesIncluded = @("STR_CONCEALABLE")
    
    [string[]] $categoriesExcluded = @("")

    $itemsRulFilePath = "$HOME\OneDrive\Documents\OpenXcom\mods\XComFiles\Ruleset\items_XCOMFILES.rul"
    $itemsRulLines = Get-Content $itemsRulFilePath
    $lineCount = $itemsRulLines.Count

    $i = 0;
    while (-not $itemsRulLines[$i].Contains(" - type: ")) {
        $i++;
    }
    $lastItemFirstLine = $i;
    
    [System.Collections.Generic.LinkedList`1[hashtable]] $linkedList = New-Object System.Collections.Generic.LinkedList[HashTable]

    for ($i = $lastItemFirstLine; $i -lt $lineCount; $i++) {
        
        if ($itemsRulLines[$i].TrimStart().StartsWith("- type: ")) {
        
            $item = Read-Item -lines $itemsRulLines[$lastItemFirstLine..($i-1)] -categoriesIncluded $categoriesIncluded -categoriesExcluded $categoriesExcluded
        
            if ($null -ne $item) {
                $linkedList.Add($item)
            }

            $lastItemFirstLine = $i
            
        }

        if ($i % 5000 -eq 0) {
            Write-Host "Reading line $i/$lineCount"
        }

    }

    Write-Host "Elapsed: $(Elapsed $StartTime)"
    Write-Host $linkedList.Count;
}

function Elapsed([DateTime] $StartTime) {
    $ElapsedTime = $(Get-Date) - $StartTime
    $FormattedElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$ElapsedTime.Ticks)
    return $FormattedElapsedTime
}

Export-ModuleMember -Function * -Cmdlet * -Alias * -Variable *