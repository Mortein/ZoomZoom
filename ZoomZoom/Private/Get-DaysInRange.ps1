<#
.SYNOPSIS
    Returns the days, months, and years of all days between 2 days
.DESCRIPTION
    Long description
.EXAMPLE
    Get-DaysInRange -StartDate 1/1/2020 -EndDate 30/1/2020
    Returns all days in January 2020 as an array
.INPUTS
    Datetime objects
.OUTPUTS
    Array of Datetime objects
.NOTES

#>

function Get-DaysInRange {
    [CmdletBinding()]
    param (
        # The Start Date for the array to calculate
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [datetime]
        $StartDate,

        # The end date for the array to calculate
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [datetime]
        $EndDate

    )

    begin {
        $Timespan = New-Timespan -Start $Start -End $End
        $DateArray = @()
    }

    process {
        $DateArray += $Start
        $WorkingTime = $Start
        foreach ($_ in 1..$Timespan.Days) {
            $WorkingTime = $WorkingTime.AddDays(1)
            $DateArray += $WorkingTime
        }
    }

    end {
        return $DateArray
    }
}
