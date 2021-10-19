VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_cMapLines As Collection
Private p_itemName As String
Private p_BestMatchLine As Object

Sub SetMapping(itemName As String)

Dim mapRange As Range
Set mapRange = WS_MAPPING.Range("mapping_" & itemName)
Set mapRange = mapRange.Resize(, mapRange.Columns.Count - 1)

Set p_cMapLines = RangeToCollection(mapRange)
p_itemName = itemName

End Sub

Property Get GetMappingValue(dDataLine As Object) As String

Dim mapLine As MappingLine
Dim dMappingLine As Object
Dim maxPerfectMatch As Integer
Dim maxGenericMatch As Integer

For Each dMappingLine In p_cMapLines
    Set mapLine = New MappingLine
    mapLine.CompareLines dMappingLine, dDataLine, p_itemName
    
    Select Case True
        Case mapLine.CountNoMatch > 0
            GoTo next_dMappingLine
        Case mapLine.CountPerfect < maxPerfectMatch
            GoTo next_dMappingLine
        Case mapLine.CountPerfect > maxPerfectMatch
            Set p_BestMatchLine = dMappingLine
            maxPerfectMatch = mapLine.CountPerfect
        Case mapLine.CountGeneric > maxGenericMatch
            Set p_BestMatchLine = dMappingLine
            maxGenericMatch = mapLine.CountGeneric
    End Select
    
next_dMappingLine:
Next dMappingLine

If p_BestMatchLine Is Nothing Then
    GetMappingValue = "NO MATCH"
Else
    GetMappingValue = p_BestMatchLine(p_itemName)
End If

Set p_BestMatchLine = Nothing

End Property
