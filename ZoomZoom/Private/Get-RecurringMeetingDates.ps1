#RecurrenceType values
<#
CdoRecurTypeDaily/0
CdoRecurTypeWeekly/1 DayOfWeekMask Interval
CdoRecurTypeMonthly/2 DayOfMonth Interval
CdoRecurTypeMonthlyNth/3 DayOfWeekMask Instance Interval
CdoRecurTypeYearly/5 DayOfMonth Interval MonthOfYear
CdoRecurTypeYearlyNth/6
#>

function Get-RecurringMeetingDates {
    param (
        # An Outlook appointment item
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [object]
        $Appointment
    )

    $DateArray = Get-DaysinRange -StartDate ($Appointment.Start).ToString() -EndDate (Get-Date).ToString()
    $Results = @()

    Switch (($Appointment.GetRecurrencePattern()).DayOfWeekMask) {
        #CdoMonday/2 - The appointment recurs on Mondays.
        2 {
            $DayOfWeekMask = $DateArray | Where-Object -Property DayOfWeek -eq 'Monday'
        }
        #CdoTuesday/4 - The appointment recurs on Tuesdays.
        4 {
            $DayOfWeekMask = $DateArray | Where-Object -Property DayOfWeek -eq 'Tuesday'
        }
        #CdoWednesday/8 - The appointment recurs on Wednesdays.
        8 {
            $DayOfWeekMask = $DateArray | Where-Object -Property DayOfWeek -eq 'Wednesday'
        }
        #CdoThursday/16 - The appointment recurs on Thursdays.
        16 {
            $DayOfWeekMask = $DateArray | Where-Object -Property DayOfWeek -eq 'Thursday'
        }
        #CdoFriday/32 - The appointment recurs on Fridays.
        32 {
            $DayOfWeekMask = $DateArray | Where-Object -Property DayOfWeek -eq 'Friday'
        }
        #Meeting that occur on every week day have the value of 62 (2+4+8+16+32)
        #62 {
        #    $DayOfWeekMask += $DateArray | Where-Object -Property DayOfWeek -NotMatch ^[S*]
        #    return $Results
        #}
    }

    Switch (($Appointment.GetRecurrencePattern()).RecurrenceType) {
        #CdoRecurTypeDaily/0 - Appointments that recur daily
        0 {
            $Results += $DateArray | Where-Object -Property DayOfWeek -NotMatch ^[S*]
            return $Results
        }
        #CdoRecurTypeWeekly/1 - Appointment recurs weekly (DayOfWeekMask,Interval)
        1 {
            $AppointmentDates = @()
            for ($i = 0; $i -lt ($DayOfWeekMask).Count; $i += (($Appointment.GetRecurrencePattern()).Interval)) {
                $AppointmentDates += $DayOfWeekMask[$i]
            }
            $Results += [PSCustomObject]@{
                Appointment      = $Appointment
                AppointmentDates = $AppointmentDates
            }
        }
        #CdoRecurTypeMonthly/2 - Appointment recurs monthly (DayOfMonth Interval)
        2 {

        }
        #CdoRecurTypeMonthlyNth/3 - Appointment recurs every Nth month
        3 {

        }
        #CdoRecurTypeYearly/5 - Appointment recurs every year
        5 {

        }
        #CdoRecurTypeYearlyNth/6 - Appointment recurs every Nth year
        6 {

        }
    }
    return $results
}
