<#
.SYNOPSIS
    Returns the days, months, and years of all days between 2 days
.DESCRIPTION
    Calculates a timespan between 2 dates, then returns each day between those 2 dates as datetime objects
.EXAMPLE
    Get-DaysInRange -StartDate 1/1/2020 -EndDate 30/1/2020
    Returns all days in January 2020 as an array
.INPUTS
    Datetime strings
.OUTPUTS
    Array of Datetime objects
.NOTES

#>

function Get-DaysInRange {
    [CmdletBinding()]
    param (
        # The Start Date for the array to calculate
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]
        $StartDate,

        # The end date for the array to calculate
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]
        $EndDate
    )

    begin {
        $Timespan = New-Timespan -Start $StartDate -End $EndDate
        $DateArray = @()
    }

    process {
        $DateArray += Get-Date $StartDate
        $WorkingTime = Get-Date $StartDate
        foreach ($_ in 1..$Timespan.Days) {
            $WorkingTime = $WorkingTime.AddDays(1)
            $DateArray += $WorkingTime
        }
    }

    end {
        return $DateArray
    }
}
