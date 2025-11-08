Attribute VB_Name = "ExcelBotProCore"
' ExcelBot Pro - Core VBA Module
' This module contains core automation functions

Option Explicit

' Main automation function
Sub AutomateTask(taskType As String)
    On Error GoTo ErrorHandler
    
    Select Case taskType
        Case "HighlightCells"
            Call HighlightCells
        Case "SortData"
            Call SortData
        Case "FilterData"
            Call FilterData
        Case "CreateChart"
            Call CreateChart
        Case "FormatRange"
            Call FormatRange
        Case Else
            MsgBox "Unknown task type: " & taskType, vbExclamation
    End Select
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in AutomateTask: " & Err.Description, vbCritical
End Sub

' Highlight cells based on value thresholds
Sub HighlightCells()
    Dim rng As Range
    Dim cell As Range
    Dim threshold1 As Double
    Dim threshold2 As Double
    
    ' Set thresholds
    threshold1 = 50
    threshold2 = 100
    
    ' Get the selected range or used range
    If Selection.Count > 1 Then
        Set rng = Selection
    Else
        Set rng = ActiveSheet.UsedRange
    End If
    
    ' Loop through each cell
    For Each cell In rng
        If IsNumeric(cell.Value) Then
            If cell.Value > threshold2 Then
                cell.Interior.Color = RGB(255, 200, 200) ' Light red
                cell.Font.Bold = True
            ElseIf cell.Value > threshold1 Then
                cell.Interior.Color = RGB(255, 255, 200) ' Light yellow
            End If
        End If
    Next cell
    
    MsgBox "Highlighting completed! Cells above " & threshold2 & " are red, above " & threshold1 & " are yellow.", vbInformation
End Sub

' Sort data in the active range
Sub SortData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim sortRange As Range
    Dim sortColumn As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Ask user for sort column
    sortColumn = Application.InputBox("Enter column number to sort by (1 for A, 2 for B, etc.):", "Sort Column", 1, Type:=1)
    
    If sortColumn < 1 Then Exit Sub
    
    ' Define the range to sort
    Set sortRange = ws.Range("A1").CurrentRegion
    
    ' Sort by specified column
    sortRange.Sort Key1:=ws.Cells(1, sortColumn), _
                    Order1:=xlAscending, _
                    Header:=xlYes
    
    MsgBox "Data sorted successfully by column " & sortColumn & "!", vbInformation
End Sub

' Apply AutoFilter to data
Sub FilterData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim filterColumn As Long
    Dim filterValue As String
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Turn on AutoFilter if not already on
    If Not ws.AutoFilterMode Then
        ws.Range("A1").CurrentRegion.AutoFilter
    End If
    
    MsgBox "AutoFilter applied! Use the dropdown arrows in the header row to filter data.", vbInformation
End Sub

' Create a chart from selected data
Sub CreateChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim dataRange As Range
    Dim chartType As Long
    
    Set ws = ActiveSheet
    
    ' Check if data is selected
    If Selection.Count < 2 Then
        MsgBox "Please select data range for the chart.", vbExclamation
        Exit Sub
    End If
    
    Set dataRange = Selection
    
    ' Ask user for chart type
    chartType = Application.InputBox("Select chart type:" & vbCrLf & _
                                     "1 = Column Chart" & vbCrLf & _
                                     "2 = Line Chart" & vbCrLf & _
                                     "3 = Pie Chart" & vbCrLf & _
                                     "4 = Bar Chart", "Chart Type", 1, Type:=1)
    
    ' Create a new chart
    Set chartObj = ws.ChartObjects.Add(Left:=300, Top:=50, Width:=400, Height:=300)
    
    With chartObj.Chart
        .SetSourceData Source:=dataRange
        
        Select Case chartType
            Case 1
                .ChartType = xlColumnClustered
            Case 2
                .ChartType = xlLine
            Case 3
                .ChartType = xlPie
            Case 4
                .ChartType = xlBarClustered
            Case Else
                .ChartType = xlColumnClustered
        End Select
        
        .HasTitle = True
        .ChartTitle.Text = "Data Chart"
        .HasLegend = True
    End With
    
    MsgBox "Chart created successfully!", vbInformation
End Sub

' Format selected range with professional styling
Sub FormatRange()
    Dim rng As Range
    
    ' Check if range is selected
    If Selection.Count < 1 Then
        MsgBox "Please select a range to format.", vbExclamation
        Exit Sub
    End If
    
    Set rng = Selection
    
    With rng
        ' Font formatting
        .Font.Name = "Calibri"
        .Font.Size = 11
        
        ' Add borders
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        
        ' Alignment
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        
        ' Format first row as header
        If rng.Rows.Count > 1 Then
            With rng.Rows(1)
                .Font.Bold = True
                .Font.Size = 12
                .Interior.Color = RGB(68, 114, 196)
                .Font.Color = RGB(255, 255, 255)
            End With
        End If
    End With
    
    MsgBox "Formatting applied successfully!", vbInformation
End Sub

' Generate summary report
Sub GenerateSummaryReport()
    Dim ws As Worksheet
    Dim reportWs As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Create new worksheet for report
    On Error Resume Next
    Set reportWs = Worksheets("Summary Report")
    If reportWs Is Nothing Then
        Set reportWs = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        reportWs.Name = "Summary Report"
    Else
        reportWs.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Add report headers
    With reportWs
        .Range("A1").Value = "Summary Report"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        
        .Range("A3").Value = "Source Sheet:"
        .Range("B3").Value = ws.Name
        
        .Range("A4").Value = "Total Rows:"
        .Range("B4").Value = lastRow
        
        .Range("A5").Value = "Generated:"
        .Range("B5").Value = Now()
    End With
    
    MsgBox "Summary report generated in '" & reportWs.Name & "' sheet!", vbInformation
End Sub

' Clear all formatting from selected range
Sub ClearFormatting()
    If Selection.Count < 1 Then
        MsgBox "Please select a range.", vbExclamation
        Exit Sub
    End If
    
    Selection.ClearFormats
    MsgBox "Formatting cleared!", vbInformation
End Sub

' Export data to CSV
Sub ExportToCSV()
    Dim ws As Worksheet
    Dim filePath As String
    Dim fileNum As Integer
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long, j As Long
    Dim lineText As String
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Get file path from user
    filePath = Application.GetSaveAsFilename(FileFilter:="CSV Files (*.csv), *.csv", Title:="Export to CSV")
    
    If filePath = "False" Then Exit Sub
    
    ' Open file for writing
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    
    ' Write data
    For i = 1 To lastRow
        lineText = ""
        For j = 1 To lastCol
            lineText = lineText & ws.Cells(i, j).Value
            If j < lastCol Then lineText = lineText & ","
        Next j
        Print #fileNum, lineText
    Next i
    
    Close #fileNum
    
    MsgBox "Data exported to: " & filePath, vbInformation
End Sub
