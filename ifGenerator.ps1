
<# 

Utility to create horrible, horrible =IF() statement chains for Sharepoint List Derived Columns, or Excel. 
Splits output into groups of 4+remainder to allow as deep as an IF statement chain as needed. 

Usage: Read the IMPORTANT heading below

Example output:

=IF((A1=11),"item11Name",
IF((A1=10),"item10Name",
IF((A1=9),"item9Name",
IF((A1=8),"item8Name",
("")))))
&
IF((A1=7),"item7Name",
IF((A1=6),"item6Name",
IF((A1=5),"item5Name",
IF((A1=4),"item4Name",
("")))))
&
IF((A1=3),"item3Name",
IF((A1=2),"item2Name",
IF((A1=1),"item1Name",
(""))))

#>

$itemsHash = @{}

# Enter Key/Value pairs needed in the =IF() statement. You can dynamically add/remove as many =IF() 
# conditions as you need with this hash array. 
$itemsHash[1] = "item1Name" 
$itemsHash[2] = "item2Name"
$itemsHash[3] = "item3Name"
$itemsHash[4] = "item4Name"
$itemsHash[5] = "item5Name"
$itemsHash[6] = "item6Name"
$itemsHash[7] = "item7Name"
$itemsHash[8] = "item8Name"
$itemsHash[9] = "item9Name"
$itemsHash[10] = "item10Name"
$itemsHash[11] = "item11Name"

$nl = [Environment]::NewLine

# IMPORTANT
# Leave only one of the following $ifType variables (lines) uncommented. Excel and Sharepoint derived 
# columns use different formats. The specific cell or column name needs to be changed inside the 
# $ifType string. 

# Use for Excel
# Replace `A1` in the string below with the Excel cell in question
$ifType = "IF((A1=" # Uncomment this line for Excel. Comment out this line if creating a Sharepoint =IF() 
 
# Use for Sharepoint
# Replace `item#` in the string below with the Sharepoint derived column name in question
#$ifType = "IF(([item#]=" # Uncomment this line for Sharepoint. Comment out this line if creating an Excel =IF()

Function ifLine ($hashKeyName, $hashValueName) {
    $outString = ""
    $outString = 'IF(([item#]=' + $hashKeyName + '),"' + $hashValueName + '",'
    $outString = $ifType + $hashKeyName + '),"' + $hashValueName + '",'
    $outString += "$nl"
    return $outString
}

$totalCount = $itemsHash.Count
[double]$remainder = $totalCount / 4 # Divide total items by 4
$remainder = [math]::floor($remainder) # Round that total down; floor()
$subCount = $totalCount - ($remainder * 4) # Amount remaining after 'groups of 4' iterations

<#
Write-Host
Write-Host "`$itemsHash.Count:" $itemsHash.Count
Write-Host "`$remainder:" $remainder
Write-Host "`$subCount:" $subCount
Write-Host
#>

$finalString = "="
$outerCount = 1

$itemsHash.GetEnumerator() | % {
    $key = $($_.key) -as [string]
    $value = $($_.value) -as [string]
    $finalString += ifLine $key $value
    
    if ((($outerCount % 4) -eq 0) -and 
    ($outerCount -ne 0) -and 
    ($outerCount -ne $itemsHash.Count)) {
        #Write-Host "On an every 4th entry iteration" #Debugging
        $finalString += '("")))))'
        $finalString += "$nl"
        $finalString += '&'
        $finalString += "$nl"
    } 
    if (($outerCount) -eq $itemsHash.Count) {
        #Write-Host "Last Entry" #Debugging 
        $finalString += '("")'
        for ($j = 0; $j -lt $subCount; $j++) {
            $finalString += ')'
        }        
    }
    $outerCount += 1
}
Write-Host $finalString


