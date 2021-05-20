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

Export-ModuleMember -Function * -Cmdlet * -Alias * -Variable *