VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MappingLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public CountPerfect As Integer
Public CountGeneric As Integer
Public CountNoMatch As Integer

Sub CompareLines(mapLine As Object, dataLine As Object, checkType As String)

Dim key As Variant, mapValue As String, dataValue As String, localMatchType As MatchType

For Each key In mapLine.keys()
    mapValue = mapLine(key)
    dataValue = dataLine(key)
    
    Select Case True
        Case CStr(key) = checkType
            'key column, skip
        Case mapValue = dataValue
            CountPerfect = CountPerfect + 1
        Case InStr(1, mapValue, Chr(42)) > 0
            CountGeneric = CountGeneric + 1
        Case Else
            CountNoMatch = CountNoMatch + 1
            GoTo ExitSub
    End Select
Next key

ExitSub:
End Sub
