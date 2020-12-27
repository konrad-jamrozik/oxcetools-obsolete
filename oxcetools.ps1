<#
.Synopsis A script that reads X-COM files weapon data and puts it into .csv file, ready to be post-processed by Excel.

The data is read from items_XCOMFILES.rul file of your X-COM files installation. 
You have to pass the path as a $dirXcomFiles parameter.

The script reads only $keysIncluded parameter keys.

This scipt reads in main loop the items .rul file line by line, reading the data for 
each item into $currItem and adding it to $items map when the item data is read in full.
To speed up execution, lines of filtered out items are skipped. Given item is filtered out if:
- it has no "categories" key
- or it doesn't have at least one category in $categoriesIncluded parameter
- or it has at least one category in $categoriesExcluded parameter

Currently known limitations:
- no support for nested keys, like "accuracyMultiplier" or "damageBonus".
- no support for attaching data from clips, one or more. Like "power".
#>
param(
    [Parameter(Mandatory=$true)]
    [ValidateScript({Test-Path $_ -PathType Container})] 
    $dirXcomFiles, # On Windows a path like "C:\Users\username\OneDrive\Documents\OpenXcom\mods\XComFiles"
    
    [string[]] 
    $keysIncluded = @("name", "categories", "weight", "accuracyAuto", "accuracySnap", "accuracyAimed", "kneelBonus", 
    "tuAuto", "tuSnap", "tuAimed", "minRange", "autoRange", "snapRange", "aimRange", "dropoff"),
    
    [string[]] 
    $categoriesIncluded = @("STR_FIREARMS", "STR_LAUNCHERS", "STR_HEAVY_WEAPONS", "STR_INCENDIARIES", "STR_PISTOLS", 
    "STR_RIFLES", "STR_CANNONS"),
    
    [string[]] 
    $categoriesExcluded = @("STR_GRENADES", "STR_CLIPS"),
    
    [string] 
    $outputCsv = "~/items.csv"
)
# Set to off to enable referencing $null hashmap keys.
# https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/set-strictmode?view=powershell-7.1
Set-StrictMode -Off 

$currentDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location) }
Import-Module "$currentDir/oxcetools-lib.psm1" -Force

# Inputs
$itemsRulFilePath = "$dirXcomFiles\Ruleset\items_XCOMFILES.rul"

<# For explanation of this data structure, 
please see the comment on the function Set-Defaults in kj-project-oxcetools-lib.psm1. 
#>
$defaults = @{
    "weight"     = @{ value = 3   };
    "kneelBonus" = @{ value = 115 };
    "minRange"   = @{ value = 0   };
    "autoRange"  = @{ value = 7    ; dependentKey = "accuracyAuto"  };
    "snapRange"  = @{ value = 15   ; dependentKey = "accuracySnap"  };
    "aimRange"   = @{ value = 200  ; dependentKey = "accuracyAimed" };
    "dropoff"    = @{ value = 2   }
}

$itemsRulLines = Get-Content $itemsRulFilePath
$lineCount = $itemsRulLines.Count

# This map will contain data of the items read from items_XCOMFILES.rul when the foreach loop below finishes.
$items = @{}
# Denotes that when we encounter next item (a line with "- type: "), then do not add currently constructed ite,
# to $items map. At first there is no item being constructed, so it is set to true.
$skipUntilNextItem = $true
# Once we read "- type: " line, we will expect for the next line to be "categories: ", and will set this variable
# to true. This is optimization. If we expect categories, and we won't find any, we will skip all lines of given item,
# by setting $skipUntilNextItem to true.
$expectCategoriesOnNextLine = $false
# line index, used only in debug messages. You can enable debug with:
# $DebugPreference = 'Continue'
# and disable with
# $DebugPreference = 'SilentlyContinue'
$lineIndex = 0

# The main script loop. See the script doc for details.
foreach ($line in $itemsRulLines) {

    $lineIndex++;

    # When in debug mode, allow reviewing the script run interactively, after each 5000 lines read.
    if ($lineIndex % 5000 -eq 0) {
        Write-Host "Reading line $lineIndex/$lineCount"
        if ($DebugPreference -eq "Continue") {
            Write-Host "DEBUG Please press any key to continue"
            Read-Host
        }
    }

    # Discard whiespace at the beginning and end of the line
    $line = $line.Trim();

    # Skip comments
    if ($line.StartsWith("#")) {
        continue;
    }

    # Encounter new item data
    if ($line.StartsWith("- type: ")) {
        # Add previously read item data to the output map,
        # unless we skipped current item, or were expecting to see 
        # categories line (but instead suddenly got next item entry)
        if (!$skipUntilNextItem -and !$expectCategoriesOnNextLine) {
             
            $items.($currItem.name) = $currItem
            Write-Debug "Added new item: $($currItem.name). Categories: $($currItem.categories)"
        }
        # Reset currently read item
        $currItem = @{} 
        # Set the currently read item name to the value of "- type: "
        $currItem.name = $line.Substring("- type: ".Length)
        $expectCategoriesOnNextLine = $true
        $skipUntilNextItem = $false
        Write-Verbose "Read new item: $lineIndex $line"
        continue;
    }

    # If we already deduced the currently processed item has to be skipped,
    # then skip current line.
    if ($skipUntilNextItem) {
        Write-Verbose "Skipping $lineIndex $line"
        continue;
    }

    # If we expected to see "categories" line but didn't find one,
    # then we skip the item.
    if ($expectCategoriesOnNextLine -and !$line.StartsWith("categories"))
    {
        Write-Debug "No categories at $lineIndex $($currItem.name)"
        $skipUntilNextItem = $true
        continue;
    # If we expected and found "categories" line, then filter the item based on our categories included & excluded filters.
    } elseif ($expectCategoriesOnNextLine) { 
        $expectCategoriesOnNextLine = $false
        Read-KeyValue $line $keysIncluded $currItem
        
        $nocategoriesIncludedPresent = $null -eq (Compare-Object $currItem.categories $categoriesIncluded -IncludeEqual -ExcludeDifferent -PassThru)
        $excludedCategoriesPresent = $null -ne (Compare-Object $currItem.categories $categoriesExcluded -IncludeEqual -ExcludeDifferent -PassThru)

        # If the file has to be filtered out due to not matching our categories filters.
        if ($nocategoriesIncludedPresent -or $excludedCategoriesPresent) {
            if ($nocategoriesIncludedPresent) {
                Write-Debug "No included categories ($categoriesIncluded) present in $lineIndex $($currItem.name)"
            } elseif ($excludedCategoriesPresent) {
                Write-Debug "Excluded categories ($categoriesIncluded) present in $lineIndex $($currItem.name)"
            }
            $skipUntilNextItem = $true
            continue;
        }
    } else { # If we didn't expect "categories" line, then just read the key
        Read-KeyValue $line $keysIncluded $currItem
    }
}

Write-Host "Setting defaults"
$items.Keys | ForEach-Object { Set-Defaults $items.$_ $defaults }

Write-Host "Saving to $outputCsv"
# Note: Export-CSV doesn't support hashtables, per:
# https://github.com/PowerShell/PowerShell/issues/10999
# But it supports [PSCustomObject]s, so we can convert hashtables to them first, per:
# https://docs.microsoft.com/en-us/powershell/scripting/learn/deep-dives/everything-about-hashtable?view=powershell-7.1#creating-objects
# https://powershellexplained.com/2016-10-28-powershell-everything-you-wanted-to-know-about-pscustomobject/
$keysIncluded | ForEach-Object { $csvHeaderRow = [ordered]@{} } { $csvHeaderRow.$_ = "" } 
[PSCustomObject]$csvHeaderRow | Export-CSV -Path $outputCsv -UseQuotes Never -Delimiter ","
$items.Keys | ForEach-Object { 
    $item = $items.$_
    $csvItem = $item.Clone()
    $csvItem.categories = $item.categories -join ";"
    [PSCustomObject]$csvItem | Export-Csv -Path $outputCsv -Append -Force -UseQuotes Never -Delimiter ","
}

# Remove the extra 2nd line left over after forcing the .csv header
(Get-Content $outputCsv) 
| Where-Object { !$_.StartsWith(",") } 
| Out-File $outputCsv

# Returning $items just so if you want to inspect it interactively in terminal.
return $items
