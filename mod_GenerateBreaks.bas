Attribute VB_Name = "mod_GenerateBreaks"
Option Explicit

Public Const DEBUG_MODE As Boolean = False


Sub Button_GenerateBreaks()

Dim cData As Collection

If Not DEBUG_MODE Then On Error GoTo ErrorHandler
SetMacroMode True
   
    Set cData = RangeToCollection(WS_DATA.Range("A1:J49"))
    
    addGenericColumnsToData cData
    generateBreaks cData, "RISKCLASS"
    generateBreaks cData, "RISKWEIGHT"

    WS_CONTROL.Activate
    MsgBox "Break Generated!", vbInformation
    
SetMacroMode False
Exit Sub

ErrorHandler:
    SetMacroMode False
    MsgBox "Unexpected error occured!", vbExclamation

End Sub

Private Sub addGenericColumnsToData(cData As Collection)

Dim dDataLine As Object

For Each dDataLine In cData
    dDataLine.Add "Mapping", ""
    dDataLine.Add "Owner", ""
    dDataLine.Add "Investigation Status", ""
    dDataLine.Add "Prior Comment Date", ""
    dDataLine.Add "Prior Comment", ""
    dDataLine.Add "Comment Date", ""
    dDataLine.Add "Comment", ""
Next dDataLine

End Sub

Private Sub generateBreaks(cData As Collection, itemName As String)

Dim dDataLine As Object
Dim mappingValue As String
Dim itemMapping As mapping

Set itemMapping = New mapping
itemMapping.SetMapping itemName

RemoveWorksheet "breaks_" & itemName

For Each dDataLine In cData
    If dDataLine(itemName) = itemName Then
        GoTo next_dDataLine
    End If
    mappingValue = itemMapping.GetMappingValue(dDataLine)
    
    If mappingValue <> dDataLine(itemName) Then
        generateBreak itemName, dDataLine, mappingValue
    End If
next_dDataLine:
Next dDataLine

End Sub

Private Sub generateBreak(itemName As String, dDataLine As Object, mappingValue As String)

Dim wsBreak As Worksheet, lastRow As Long, includeHeader As Boolean

Set wsBreak = SetWorksheet("breaks_" & itemName)
dDataLine("Mapping") = mappingValue
lastRow = GetLastRow(wsBreak.Cells(1, 1))

If lastRow = 1 Then
    includeHeader = True
Else
    lastRow = lastRow + 1
End If

PrintDictionaryToRange dDataLine, wsBreak.Cells(lastRow, 1), includeHeader

End Sub
