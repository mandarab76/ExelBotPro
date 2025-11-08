Attribute VB_Name = "ExcelBotProUtils"
' ExcelBot Pro - Utility Functions Module
' This module contains helper functions for data analysis and manipulation

Option Explicit

' Calculate basic statistics for a range
Function CalculateStats(dataRange As Range) As String
    Dim result As String
    Dim total As Double
    Dim count As Long
    Dim avg As Double
    Dim maxVal As Double
    Dim minVal As Double
    
    On Error GoTo ErrorHandler
    
    total = Application.WorksheetFunction.Sum(dataRange)
    count = Application.WorksheetFunction.Count(dataRange)
    avg = Application.WorksheetFunction.Average(dataRange)
    maxVal = Application.WorksheetFunction.Max(dataRange)
    minVal = Application.WorksheetFunction.Min(dataRange)
    
    result = "Statistics for selected range:" & vbCrLf & vbCrLf
    result = result & "Count: " & count & vbCrLf
    result = result & "Sum: " & Format(total, "#,##0.00") & vbCrLf
    result = result & "Average: " & Format(avg, "#,##0.00") & vbCrLf
    result = result & "Maximum: " & Format(maxVal, "#,##0.00") & vbCrLf
    result = result & "Minimum: " & Format(minVal, "#,##0.00")
    
    CalculateStats = result
    Exit Function
    
ErrorHandler:
    CalculateStats = "Error calculating statistics: " & Err.Description
End Function

' Show statistics for selected range
Sub ShowRangeStatistics()
    Dim rng As Range
    Dim stats As String
    
    ' Check if range is selected
    If Selection.Count < 1 Then
        MsgBox "Please select a range with numeric data.", vbExclamation
        Exit Sub
    End If
    
    Set rng = Selection
    stats = CalculateStats(rng)
    
    MsgBox stats, vbInformation, "Range Statistics"
End Sub

' Remove duplicate rows
Sub RemoveDuplicates()
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim lastRow As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "Not enough data to remove duplicates.", vbExclamation
        Exit Sub
    End If
    
    Set dataRange = ws.Range("A1").CurrentRegion
    
    ' Remove duplicates
    dataRange.RemoveDuplicates Columns:=1, Header:=xlYes
    
    MsgBox "Duplicate rows removed!", vbInformation
End Sub

' Find and replace text
Sub FindAndReplaceText()
    Dim findText As String
    Dim replaceText As String
    Dim rng As Range
    
    ' Get search and replace text from user
    findText = InputBox("Enter text to find:", "Find Text")
    If findText = "" Then Exit Sub
    
    replaceText = InputBox("Enter replacement text:", "Replace Text")
    
    ' Perform replacement
    Set rng = ActiveSheet.UsedRange
    rng.Replace What:=findText, Replacement:=replaceText, LookAt:=xlPart, _
                SearchOrder:=xlByRows, MatchCase:=False
    
    MsgBox "Find and replace completed!", vbInformation
End Sub

' Convert text to proper case
Sub ConvertToProperCase()
    Dim cell As Range
    
    If Selection.Count < 1 Then
        MsgBox "Please select cells to convert.", vbExclamation
        Exit Sub
    End If
    
    For Each cell In Selection
        If Not IsEmpty(cell.Value) Then
            cell.Value = Application.WorksheetFunction.Proper(cell.Value)
        End If
    Next cell
    
    MsgBox "Text converted to proper case!", vbInformation
End Sub

' Trim whitespace from cells
Sub TrimWhitespace()
    Dim cell As Range
    
    If Selection.Count < 1 Then
        MsgBox "Please select cells to trim.", vbExclamation
        Exit Sub
    End If
    
    For Each cell In Selection
        If Not IsEmpty(cell.Value) Then
            cell.Value = Application.WorksheetFunction.Trim(cell.Value)
        End If
    Next cell
    
    MsgBox "Whitespace trimmed from selected cells!", vbInformation
End Sub

' Create pivot table
Sub CreatePivotTable()
    Dim ws As Worksheet
    Dim pivotWs As Worksheet
    Dim dataRange As Range
    Dim pivotCache As PivotCache
    Dim pivotTable As pivotTable
    
    On Error GoTo ErrorHandler
    
    Set ws = ActiveSheet
    Set dataRange = ws.Range("A1").CurrentRegion
    
    ' Validate data
    If dataRange.Rows.Count < 2 Then
        MsgBox "Not enough data to create a pivot table.", vbExclamation
        Exit Sub
    End If
    
    ' Create new worksheet for pivot table
    Set pivotWs = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    pivotWs.Name = "Pivot_" & Format(Now, "hhmmss")
    
    ' Create pivot cache and pivot table
    Set pivotCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    
    Set pivotTable = pivotCache.CreatePivotTable( _
        TableDestination:=pivotWs.Range("A3"), _
        TableName:="ExcelBotPivot")
    
    MsgBox "Pivot table created in '" & pivotWs.Name & "' sheet!" & vbCrLf & _
           "Configure the pivot table by dragging fields to the appropriate areas.", vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating pivot table: " & Err.Description, vbCritical
End Sub

' Generate random data for testing
Sub GenerateRandomData()
    Dim ws As Worksheet
    Dim i As Long
    Dim rows As Long
    Dim startRow As Long
    
    Set ws = ActiveSheet
    
    rows = Application.InputBox("How many rows of random data to generate?", "Generate Data", 100, Type:=1)
    If rows < 1 Then Exit Sub
    
    startRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If startRow = 2 Then startRow = 1
    
    Application.ScreenUpdating = False
    
    ' Generate headers
    ws.Cells(startRow, 1).Value = "ID"
    ws.Cells(startRow, 2).Value = "Name"
    ws.Cells(startRow, 3).Value = "Value"
    ws.Cells(startRow, 4).Value = "Date"
    
    ' Generate random data
    For i = 1 To rows
        ws.Cells(startRow + i, 1).Value = i
        ws.Cells(startRow + i, 2).Value = "Item " & i
        ws.Cells(startRow + i, 3).Value = Int(Rnd() * 1000)
        ws.Cells(startRow + i, 4).Value = Date - Int(Rnd() * 365)
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox rows & " rows of random data generated!", vbInformation
End Sub

' Protect/Unprotect worksheet
Sub ToggleSheetProtection()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    If ws.ProtectContents Then
        ws.Unprotect
        MsgBox "Worksheet unprotected!", vbInformation
    Else
        ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
        MsgBox "Worksheet protected!", vbInformation
    End If
End Sub

' Add timestamp to selected cell
Sub InsertTimestamp()
    If Selection.Count <> 1 Then
        MsgBox "Please select a single cell.", vbExclamation
        Exit Sub
    End If
    
    Selection.Value = Now()
    Selection.NumberFormat = "mm/dd/yyyy hh:mm:ss"
    
    MsgBox "Timestamp inserted!", vbInformation
End Sub

' Color alternate rows
Sub ColorAlternateRows()
    Dim rng As Range
    Dim i As Long
    Dim color1 As Long
    Dim color2 As Long
    
    If Selection.Count < 1 Then
        MsgBox "Please select a range.", vbExclamation
        Exit Sub
    End If
    
    Set rng = Selection
    color1 = RGB(255, 255, 255) ' White
    color2 = RGB(230, 240, 255) ' Light blue
    
    For i = 1 To rng.Rows.Count
        If i Mod 2 = 0 Then
            rng.Rows(i).Interior.Color = color2
        Else
            rng.Rows(i).Interior.Color = color1
        End If
    Next i
    
    MsgBox "Alternate row coloring applied!", vbInformation
End Sub

' Freeze panes helper
Sub FreezePanes()
    Dim freezeRow As Long
    Dim freezeCol As Long
    
    freezeRow = Application.InputBox("Freeze rows above row number (0 for none):", "Freeze Rows", 1, Type:=1)
    freezeCol = Application.InputBox("Freeze columns to the left of column number (0 for none):", "Freeze Columns", 1, Type:=1)
    
    If freezeRow < 0 Or freezeCol < 0 Then Exit Sub
    
    ActiveWindow.FreezePanes = False
    
    If freezeRow > 0 Or freezeCol > 0 Then
        Cells(freezeRow + 1, freezeCol + 1).Select
        ActiveWindow.FreezePanes = True
    End If
    
    MsgBox "Freeze panes applied!", vbInformation
End Sub

' Create data validation list
Sub AddValidationList()
    Dim listItems As String
    Dim rng As Range
    
    If Selection.Count < 1 Then
        MsgBox "Please select cells to add validation.", vbExclamation
        Exit Sub
    End If
    
    listItems = InputBox("Enter list items separated by commas:", "Validation List", "Yes,No,Maybe")
    If listItems = "" Then Exit Sub
    
    Set rng = Selection
    
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=listItems
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    MsgBox "Data validation list added!", vbInformation
End Sub

' Merge and center cells
Sub MergeAndCenterSelection()
    If Selection.Count < 2 Then
        MsgBox "Please select multiple cells to merge.", vbExclamation
        Exit Sub
    End If
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = True
    End With
    
    MsgBox "Cells merged and centered!", vbInformation
End Sub
