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
function Set-Default([hashtable] $item, [string] $key, [object] $defaultValue, [string] $dependentKey = "") {
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

    ($key, $valueString, $type) = if ($line.StartsWith(" ")) {

        if ($line.Contains("- ")) {
            ("-", ($line -replace "- ").Trim(), "listItem")
        } elseif ($line.Contains(": ")) {
            $keyValuePair = $line.Trim() -split ": "
            $key = $keyValuePair[0]
            ($key, $line.Trim().Substring(($key + ": ").Length), "listItem")
        } else { throw "Invalid case!"}
    } elseif ($line.Contains(": ")) {
        $keyValuePair = $line -split ": "
        $key = $keyValuePair[0]
        ($key, $line.Substring(($key + ": ").Length), "keyValue")
    } elseif ($line.EndsWith(":") -and !($line.Contains("- "))) {
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

    # KJA TODO pass in as param
    # note that list headers cannot be included here
    if (@("bigSprite", "floorSprite", "handSprite", "bulletSprite", "fireSound", "hitSound", "hitAnimation",
    "armor", "attraction", "invHeight", "invWidth", "listOrder", "recoveryPoints") -contains $key) {
        return $null
    }

    return @($key, $value, $type)
}

function Read-Item([string[]] $lines, [hashtable] $item, [string[]] $categoriesIncluded, [string[]] $categoriesExcluded) {

    if ($lines.Count -lt 1) { throw "Expected for the item being read to have at least one line." }
    if (!$lines[0].Substring(4).StartsWith("type: ")) { throw "Expected for the first line to start with '  - type: '. Instead got: '$($lines[0])'." }

    $lines = $lines
        | Where-Object { !($_.StartsWith("#")) } # Remove comments
        | Where-Object { $_.Length -ge 4 }
        | ForEach-Object {
            $_.Substring(4).TrimEnd() }

    # If categories are missing
    if ($lines.Count -lt 2) { return $null }
    if (!$lines[1].StartsWith("categories: ")) { return $null }

    $name = $lines[0].Substring("type: ".Length)
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

    $lineData = $lines[2..$lines.Length]
    | ForEach-Object { ,@(Read-LineData $_) }
    | Where-Object { $null -ne $_[0] }

    $ht = @{ "name" = $name; "categories" = $categories }
    $listMode = $false
    $listKey = ""
    $list = @{}

    for ($i = 0; $i -lt $lineData.Count; $i++) {
        ($key, $value, $type) = $lineData[$i]

        if ($type -eq "listHeader") {
            if ($listMode -eq $true) {
                if ($list.ContainsKey("-")) {
                    $ht += @{ $listKey = $list["-"] }
                } else {
                    $ht += @{ $listKey = $list }
                }
            }
            $listMode = $true
            $listKey = $key
            $list = @{}
        }

        if ($type -eq "listItem") {
            if ($listMode -eq $false) { throw "Invalid state!" }

            if ($key -eq "-") {
                if ($list.ContainsKey("-")) {
                    $list["-"] += $value
                } else {
                    $list["-"] = @($value)
                }
            } else {
                $list += @{ $key = $value }
            }
        }

        if ($type -eq "keyValue") {
            if ($listMode -eq $true) {
                if ($list.ContainsKey("-")) {
                    $ht += @{ $listKey = $list["-"] }
                } else {
                    $ht += @{ $listKey = $list }
                }
                $listMode = $false
            } else {
                # There was no list, so nothing to do with that regard.
            }
            $ht += @{ $key = $value }
        }
    }

    return $ht
}

# KJA Be aware there are not yet handled special ammo and double-nesting cases. See e.g. STR_HUMAN_SONIC_HEAVY_CANNON
function Read-ItemsDebug()
{
    [string[]] $categoriesIncluded = @("STR_CONCEALABLE")

    [string[]] $categoriesExcluded = @("STR_MELEE", "STR_GRENADES", "STR_MEDICAL")

    $itemsRulFilePath = "$HOME\OneDrive\Documents\OpenXcom\mods\XComFiles\Ruleset\items_XCOMFILES.rul"
    $itemsRulLines = Get-Content $itemsRulFilePath
    $lineCount = $itemsRulLines.Count

    $i = 0;
    while (-not $itemsRulLines[$i].Contains(" - type: ")) {
        $i++;
    }
    $lastItemFirstLine = $i;

    $items = @{}
    $itemsWithAmmoList = @{}
    $clips = @{}
    $itemsOther = @{}

    for ($i = $lastItemFirstLine; $i -lt $lineCount; $i++) {

        if ($itemsRulLines[$i].TrimStart().StartsWith("- type: ")) {

            $item = Read-Item -lines $itemsRulLines[$lastItemFirstLine..($i-1)] -categoriesIncluded $categoriesIncluded -categoriesExcluded $categoriesExcluded
            $lastItemFirstLine = $i
            if ($null -eq $item) {
                continue;
            }
            $items += @{ $item.name = $item }
            if ($item.ContainsKey("compatibleAmmo")) {
                $itemsWithAmmoList += @{ $item.name = $item }
                if ($item.categories -contains "STR_CLIPS") {
                    throw "Invalid item!"
                }
            } elseif ($item.categories -contains "STR_CLIPS") {
                $clips += @{ $item.name = $item }
            } else {
                if ($item.ContainsKey("battleType") -and $item.battleType -ne 0) {
                    $itemsOther += @{ $item.name = $item }
                    Write-Host "other item: $($item.name)"
                } else {
                    # Exclude Geoscape-only items
                    # https://www.ufopaedia.org/index.php/Ruleset_Reference_Nightly_(OpenXcom)#Naming.2C_Categorization_and_Storage
                }
            }
        }

        if ($i % 5000 -eq 0) {
            Write-Host "Reading line $i/$lineCount"
        }
    }

    $itemsWithLoadedClip = @{}

    foreach ($itemWithAmmoList in $itemsWithAmmoList.values) {
        Write-Host "Loading ammo into $($itemWithAmmoList.name)"
        foreach ($clipName in $itemWithAmmoList.compatibleAmmo) {
            [Hashtable] $itemClone = $itemWithAmmoList.Clone()
            [Hashtable] $clipClone = $clips[$clipName].Clone()

            $name = $itemClone.name + "_" + ($clipClone.name -replace $itemClone.name)
            $itemClone.Remove("name")
            $clipClone.Remove("name")
            $size = [math]::Round([float]$itemClone.size + $clipClone.size, 3)
            $itemClone.Remove("size")
            $clipClone.Remove("size")
            $weight = [int]$itemClone.weight + $clipClone.weight
            $itemClone.Remove("weight")
            $clipClone.Remove("weight")
            $costBuy = [int]$itemClone.costBuy + $clipClone.costBuy
            $itemClone.Remove("costBuy")
            $clipClone.Remove("costBuy")
            $costSell = [int]$itemClone.costSell + $clipClone.costsell
            $itemClone.Remove("costSell")
            $clipClone.Remove("costSell")
            $clipClone.Remove("costThrow")
            $itemClone.Remove("requires")
            $clipClone.Remove("requires")
            $itemClone.Remove("requiresBuy")
            $clipClone.Remove("requiresBuy")
            $itemClone.Remove("compatibleAmmo")
            $clipClone.Remove("categories")
            $clipClone.Remove("battleType")

            $itemClone += $clipClone

            $itemClone.name = $name
            $itemClone.size = $size
            $itemClone.weight = $weight
            $itemClone.costBuy = $costBuy
            $itemClone.costSell = $costSell

            $itemsWithLoadedClip[$itemClone.name] = $itemClone

        }
    }

    return ($items, $itemsWithAmmoList, $clips, $itemsWithLoadedClip, $itemsOther)
}

function Read-ItemsMain {

    $StartTime = Get-Date

    $defaults = @{
        "size"       = @{ value = 0.0 };
        "weight"     = @{ value = 3   };
        "autoShots"  = @{ value = 1    ; dependentKey = "accuracyAuto"  }
        "kneelBonus" = @{ value = 115 };
        "minRange"   = @{ value = 0   };
        "autoRange"  = @{ value = 7    ; dependentKey = "accuracyAuto"  };
        "snapRange"  = @{ value = 15   ; dependentKey = "accuracySnap"  };
        "aimRange"   = @{ value = 200  ; dependentKey = "accuracyAimed" };
        "dropoff"    = @{ value = 2   }
        "twoHanded"  = @{ value = "false" };
        "shotgunBehavior" = @{ value = 0   ; dependentKey = "shotgunPellets" }
        "shotgunSpread"   = @{ value = 100 ; dependentKey = "shotgunPellets" }
        "shotgunChoke"    = @{ value = 100 ; dependentKey = "shotgunPellets" }        
    }


    ($items, $itemsWithAmmoList, $clips, $itemsWithLoadedClip, $itemsOther) = Read-ItemsDebug
    Write-Host "Items count: $($items.Count)"
    Write-Host "Items with ammo list count: $($itemsWithAmmoList.Count)"
    Write-Host "Clips count: $($clips.Count)"
    Write-Host "Items with loaded clip count: $($itemsWithLoadedClip.Count)"
    Write-Host "Other items count: $($itemsOther.Count)"

    $itemsToOutput = ($itemsWithLoadedClip + $itemsOther)

    Write-Host "Setting defaults"
    $itemsToOutput.Keys | ForEach-Object { Set-Defaults $itemsToOutput.$_ $defaults }

    Write-ItemsDataToCsv $itemsToOutput

    Write-Host "Elapsed: $(Elapsed $StartTime)"
}

function Write-ItemsDataToCsv($items) {

    $keysIncluded = @(
    # Common keys
    "name", "categories", "weight", "size", "costSell", "costBuy", "battleType",

    # Weapon keys
    "twoHanded", "autoShots", "sprayWaypoints", "shotgunChoke",
    "accuracyCloseQuarters", "accuracyAuto", "accuracySnap", "accuracyAimed",
    "kneelBonus",
    "tuLoad", "tuAuto", "tuSnap", "tuAimed",
    "minRange", "autoRange", "snapRange", "aimRange", "dropoff",

    # Ammo keys
    "power", "damageType", "shotgunBehavior", "shotgunSpread", "shotgunPellets"
    )

    $outputCsv = "~/items_2.csv"
    Write-Host "Saving to $outputCsv"
    # Note: Export-CSV doesn't support hashtables, per:
    # https://github.com/PowerShell/PowerShell/issues/10999
    # But it supports [PSCustomObject]s, so we can convert hashtables to them first, per:
    # https://docs.microsoft.com/en-us/powershell/scripting/learn/deep-dives/everything-about-hashtable?view=powershell-7.1#creating-objects
    # https://powershellexplained.com/2016-10-28-powershell-everything-you-wanted-to-know-about-pscustomobject/
    $keysIncluded | ForEach-Object { $csvHeaderRow = [ordered]@{} } { $csvHeaderRow.$_ = "" }
    # Need to output this empty row to force ordering of properties upon calls to Export-Csv below, per:
    # https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/export-csv?view=powershell-7.1#notes
    [PSCustomObject]$csvHeaderRow | Export-CSV -Path $outputCsv -UseQuotes Never -Delimiter ","
    $items.Keys | ForEach-Object {
        $item = $items.$_
        $csvItem = $item.Clone()
        $csvItem.categories = $item.categories -join ";"

        [PSCustomObject]$csvItem | Export-Csv -Path $outputCsv -Append -Force -UseQuotes Never -Delimiter ","
    }
}

function Elapsed([DateTime] $StartTime) {
    $ElapsedTime = $(Get-Date) - $StartTime
    $FormattedElapsedTime = "{0:HH:mm:ss}" -f ([datetime]$ElapsedTime.Ticks)
    return $FormattedElapsedTime
}

Export-ModuleMember -Function * -Cmdlet * -Alias * -Variable *

# KJA TODO: damageAlter / armorEffectiveness
# KJA TODO: max range