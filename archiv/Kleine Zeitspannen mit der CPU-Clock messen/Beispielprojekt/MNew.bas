Attribute VB_Name = "MNew"
Option Explicit

Public Function CPUClock(ByVal Rate As Double) As CPUClock
    Set CPUClock = New CPUClock: CPUClock.New_ Rate
End Function


