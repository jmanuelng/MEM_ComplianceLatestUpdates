<#
.SYNOPSIS
    Check for the update status of the latest cumulative update for Windows.

.DESCRIPTION
    Script checks for the latest patch Tuesday and adds an offset of days that is used to provide the user with the install window. 
    Based on the installation status of the latest cumulative update, the Windows device will return if it’s Up-to-date or Not up-to-date. 
    That information is returned in a compressed single line JSON-format.

    Author: Peter van der Woude tw:@pvanderwoude
    Source: https://www.petervanderwoude.nl/post/using-update-status-as-part-of-the-compliance-of-windows-devices/


.NOTES
    
    Author: Peter van der Woude tw:@pvanderwoude
    Source: https://www.petervanderwoude.nl/post/using-update-status-as-part-of-the-compliance-of-windows-devices/

    PowerShell script logics for determining the second Tuesday of the month are based on this example provided by Travis Roberts on his GitHub:
    https://github.com/tsrob50/Get-PatchTuesday/blob/master/Get-PatchTuesday.ps1

    Me, added support for multi-language information message (in json).

    Deploy script as Custm Compliance Policy.

#>
[datetime]$dtToday = [datetime]::NOW
$strCurrentMonth = $dtToday.Month.ToString()
$strCurrentYear = $dtToday.Year.ToString()
[datetime]$dtMonth = $strCurrentMonth + '/1/' + $strCurrentYear

while ($dtMonth.DayofWeek -ne 'Tuesday') { 
      $dtMonth = $dtMonth.AddDays(1) 
}

$strPatchTuesday = $dtMonth.AddDays(7)
$intOffSet = 7

if ([datetime]::NOW -lt $strPatchTuesday -or [datetime]::NOW -ge $strPatchTuesday.AddDays($intOffSet)) {
    $objUpdateSession = New-Object -ComObject Microsoft.Update.Session
    $objUpdateSearcher = $objUpdateSession.CreateupdateSearcher()
    $arrAvailableUpdates = @($objUpdateSearcher.Search("IsAssigned=1 and IsHidden=0 and IsInstalled=0").Updates)
    $strAvailableCumulativeUpdates = $arrAvailableUpdates | Where-Object {$_.title -like "*cumulative*"}

    if ($strAvailableCumulativeUpdates -eq $null) {
        $strUpdateStatus = @{"Update status" = "Up-to-date"}
    } 
    else {
        $strUpdateStatus = @{"Update status" = "Not up-to-date"}
    }
} 
else {
    $strUpdateStatus = @{"Update status" = "Up-to-date"}
}

return $strUpdateStatus | ConvertTo-Json -Compress