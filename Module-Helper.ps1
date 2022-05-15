function Display-HelperProcessingTime
{
    [cmdletbinding()]
	param (
        [Parameter(Mandatory=$true)][DateTime]$StartDate,
        [Parameter(Mandatory=$false)][DateTime]$EndDate,
        [Parameter(Mandatory=$false)][String]$AdditionalText,
        [Parameter(Mandatory=$false)][String]$BatchInformation,
        [Parameter(Mandatory=$false)][string]$Indentation
    );
    if ($EndDate -eq $null) { $EndDate = (Get-Date); }
    $TotalTime = ($EndDate - $StartDate);

    if ($AdditionalText -eq $null) { $AdditionalText = ""; }
    if ($BatchInformation -eq $null) { $BatchInformation = ""; }
    

    $Hours = $TotalTime.Hours.ToString(); if ($TotalTime.Hours -lt 10) { $Hours = ("0" + $Hours);  }
    $Minutes = $TotalTime.Minutes.ToString(); if ($TotalTime.Minutes -lt 10) { $Minutes = ("0" + $Minutes); }
    $Seconds = $TotalTime.Seconds.ToString(); if ($TotalTime.Seconds -lt 10) { $Seconds = ("0" + $Seconds); }
    if ($Indentation -eq $null) { $Indentation = ""; }

    $DisplayStr = (
        $Indentation +
        $BatchInformation +
        $AdditionalText +         
        $TotalTime.Days.ToString() +
        "." + $Hours + 
        ":" + $Minutes +
        ":" + $Seconds + 
        "." + $TotalTime.Milliseconds.ToString()
    );

    Write-Host ($DisplayStr);
}