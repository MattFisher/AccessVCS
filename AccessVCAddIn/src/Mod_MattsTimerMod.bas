' Simple Timer debugging routines
' Useful to get a rough idea how long different chunks of code take.
' Author:           Matt Fisher
' Date Created:     09-03-08
' Date Modified:    27-10-09

Option Compare Database
Option Explicit

Const SECONDS_PER_DAY As Double = 86400
Dim startTime As Double
Dim lastCheckedTime As Double
Dim newCheckedTime As Double
Dim timeDiff As Double

Public Sub StartTimer()
' Starts timer and prints info to debug. eg:
' "Timer Started at 10-Mar-2008 12:51:00.67"
' NOTE: On dev machine, Timer() increments 0.015s or 0.016s
' Author:           Matt Fisher
' Date Created:     09-03-08
' Date Modified:    10-03-08

startTime = Timer() + (CDbl(Date) * SECONDS_PER_DAY)
lastCheckedTime = startTime
Debug.Print "Timer Started at " & _
            Format(CDate(Int(startTime) / SECONDS_PER_DAY), _
                   "dd-mmm-yyyy Hh:Nn:Ss") & _
            "." & Right(Format((startTime), "#0.00"), 2)
End Sub

Public Function CheckTimer() As Double
' Prints timing info to debug output. eg:
' "Time since last check: 00h 00m 00.56s"
' "Total time elapsed:    00h 00m 04.16s"
' Appends " # days" if over 25 hrs has passed
' Returns total number of days since timer started as double
' NOTE: On dev machine, Timer() increments 0.015s or 0.016s
' CheckTimer itself takes approx 5 milliseconds (0.005s)
' Author:           Matt Fisher
' Date Created:     09-03-08
' Date Modified:    10-03-08

newCheckedTime = Timer() + (CDbl(Date) * SECONDS_PER_DAY)
timeDiff = newCheckedTime - lastCheckedTime
CheckTimer = newCheckedTime - startTime
Debug.Print "Time since last check: " & _
    GetTimeString(timeDiff) & _
    "  Total: " & _
    GetTimeString(CheckTimer)
lastCheckedTime = newCheckedTime
newCheckedTime = -1
End Function

Public Function GetTimeString(timePeriod As Double) As String
GetTimeString = _
    Format(CDate(Int(timePeriod) / SECONDS_PER_DAY), _
           "Hh\h Nn\m Ss") & _
    "." & Right(Format(timePeriod, "#0.00"), 2) & "s" & _
    IIf(timePeriod >= SECONDS_PER_DAY, _
        (" " & (timePeriod \ SECONDS_PER_DAY) & " days"), "")
End Function