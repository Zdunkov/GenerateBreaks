Attribute VB_Name = "mod_GenericFunctions"
Option Explicit

Sub SetMacroMode(mode As Boolean)

If mode Then
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False
Else
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.StatusBar = False
End If

End Sub


Function GetLastRow(rng As Range, Optional columnsCount As Long) As Long

Dim iRow As Long, iCol As Long, curLastRow As Long, finalLastRow As Long, curWs As Worksheet, curCellValue As String

Set curWs = rng.Worksheet
If columnsCount < 1 Then columnsCount = rng.Columns.Count

For iCol = 1 To columnsCount
    iRow = curWs.Cells(1048576, rng.Column).Offset(0, iCol - 1).End(xlUp).Row
    curCellValue = GetSafeCellValue(curWs.Cells(iRow, rng.Column).Offset(0, iCol - 1), True)
    
    Do Until curCellValue <> ""
        iRow = iRow - 1
        If iRow < 1 Then GoTo next_iCol
        curCellValue = GetSafeCellValue(curWs.Cells(iRow, rng.Column).Offset(0, iCol - 1), True)
    Loop
    
    curLastRow = iRow
    finalLastRow = WorksheetFunction.Max(curLastRow, finalLastRow)
next_iCol:
Next iCol

If finalLastRow = 0 Then finalLastRow = 1
GetLastRow = finalLastRow

End Function
Function GetSafeCellValue(ByVal rng As Range, Optional OnErrorReturn_ERROR_ As Boolean) As String

Dim tempStr As String

Set rng = rng.Resize(1, 1)

tempStr = "tempCellValue_xxxxxx"
On Error Resume Next
    tempStr = rng.Value
On Error GoTo 0

If tempStr = "tempCellValue_xxxxxx" Then
    If OnErrorReturn_ERROR_ Then
        GetSafeCellValue = "_ERROR_"
    Else
        GetSafeCellValue = ""
    End If
Else
    GetSafeCellValue = tempStr
End If

End Function


Function RangeToCollection(rng As Range) As Collection

Dim rowsCount As Long, colsCount As Long, iRow As Long, iCol As Long, tempCollection As Collection, tempDictionary As Object, header As String, cellValue As String

Set tempCollection = New Collection
rowsCount = GetLastRow(rng) - rng.Row + 1
colsCount = rng.Columns.Count

For iRow = 1 To rowsCount
    Set tempDictionary = Nothing
    Set tempDictionary = CreateObject("Scripting.Dictionary")
    
    For iCol = 1 To colsCount
        header = rng.Cells(1, iCol).Value
        cellValue = GetSafeCellValue(rng.Cells(iRow, iCol))
        tempDictionary.Add header, cellValue
    Next iCol
    
    tempCollection.Add tempDictionary, CStr(iRow)
    
Next iRow

Set RangeToCollection = tempCollection

End Function

Function SetWorksheet(itemName As String) As Worksheet

Dim tempSheet As Worksheet

On Error Resume Next
    Set tempSheet = ThisWorkbook.Sheets(itemName)
On Error GoTo 0

If tempSheet Is Nothing Then
    Set tempSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    tempSheet.Name = itemName
End If

Set SetWorksheet = tempSheet

End Function

Sub RemoveWorksheet(wsName As String)

Dim currentAlerts As Boolean
currentAlerts = Application.DisplayAlerts

On Error Resume Next
Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(wsName).Delete
Application.DisplayAlerts = currentAlerts
On Error GoTo 0

End Sub

Sub PrintDictionaryToRange(dict As Object, startCell As Range, includeHeader As Boolean)

Dim key As Variant, rowOffset As Integer, colOffset As Long

If includeHeader Then
    For Each key In dict.keys()
        startCell.Offset(rowOffset, colOffset).Value = CStr(key)
        colOffset = colOffset + 1
    Next key
    rowOffset = rowOffset + 1
End If

colOffset = 0
For Each key In dict.keys()
    startCell.Offset(rowOffset, colOffset).Value = dict(key)
    colOffset = colOffset + 1
Next key

End Sub

