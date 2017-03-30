Attribute VB_Name = "Module3"
Sub Main()
'
' Main Macro
' Runs all of the different parts to format the spreadsheet
'
    Dim Filename As String
    Filename = Right(ThisWorkbook.FullName, Len(ThisWorkbook.FullName) - InStrRev(ThisWorkbook.FullName, "\"))

    Application.Run "'" + Filename + "'!SortRows"
    Application.Run "'" + Filename + "'!DeleteBlankRows"
    Application.Run "'" + Filename + "'!AddBlankRows"
    Application.Run "'" + Filename + "'!FormatCells"
    Application.Run "'" + Filename + "'!AddBorders"
End Sub

Sub SortRows()
'
' SortDeleteandAddBlankRow Macro
' Sorts the Rows according to the date of the party,
' then the start time, then the end time
'

    Dim Sheet1 As Worksheet
    Dim SF As SortFields
    
    Set Sheet1 = ActiveWorkbook.Worksheets("Sheet1")
    Set SF = Sheet1.Sort.SortFields
    
    SF.Clear
    SF.Add Key:=Columns("D") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    SF.Add Key:=Columns("B") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
        "8:00am,9:00am,10:00am,11:00am,11:30am,12:00pm,1:00pm,1:30pm,2:00pm,3:00pm,4:00pm,4:30pm,5:00pm,6:00pm,7:00pm,8:00pm", DataOption:=xlSortNormal
    SF.Add Key:=Columns("C") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, CustomOrder:= _
        "8:00am,9:00am,10:00am,11:00am,11:30am,12:00pm,1:00pm,1:30pm,2:00pm,3:00pm,4:00pm,4:30pm,5:00pm,6:00pm,7:00pm,8:00pm", DataOption:=xlSortNormal
    With Sheet1.Sort
        .SetRange Range("A1:E75")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
End Sub

Sub DeleteBlankRows()
'
' DeleteBlankRows Macro
' Deletes all the blank rows in the spreadsheet
'

  Dim r As Range, i As Long
  Set r = ActiveSheet.Range("A1:Z50")
  For i = r.rows.Count To 1 Step (-1)
    If WorksheetFunction.CountA(r.rows(i)) = 0 Then r.rows(i).Delete
  Next
End Sub

Sub AddBlankRows()
'
' AddBlankRows Macro
' Adds blank rows in between different dates,
' so that all parties on the same date are grouped
'
    Dim LastRow As Long, i As Long
    
    LastRow = Cells(rows.Count, "D").End(xlUp).Row
    Application.ScreenUpdating = False
    For i = LastRow To 3 Step -1
        If Cells(i, "D") <> Cells(i - 1, "D") _
                And Not IsEmpty(Cells(i, "D")) _
                And Not IsEmpty(Cells(i - 1, "D")) Then
            rows(i).Insert
        End If
    Next i
    Application.ScreenUpdating = True
End Sub

Sub FormatCells()
'
' FormatCells Macro
' Formats the cells to your liking
'

    Columns("A").HorizontalAlignment = xlLeft
    Columns("B:D").HorizontalAlignment = xlCenter
    Columns("E").HorizontalAlignment = xlLeft
    rows("1").HorizontalAlignment = xlCenter
    ' Change dates to correct date format
    Columns("D").NumberFormat = "[$-x-sysdate]dddd, mmmm dd, yyyy"
    ' Change font of regular text to desired
    With rows("2:" & rows.Count).Font
        .Name = "Baskerville Old Face"
        .Size = 11
    End With
    With rows("2:" & rows.Count)
        .VerticalAlignment = xlCenter
    End With
End Sub

Sub AddBorders()
'
' AddBorders Macro
' Adds borders around all of the used cells
'
    With Range("A1:E" & Range("A" & rows.Count).End(xlUp).Row)
        .BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
         On Error Resume Next 'used in case there are no inside borders
         With .Borders(xlInsideHorizontal)
             .LineStyle = xlContinuous
             .Weight = xlThin
             .ColorIndex = xlAutomatic
         End With
         With .Borders(xlInsideVertical)
             .LineStyle = xlContinuous
             .Weight = xlThin
             .ColorIndex = xlAutomatic
         End With
    End With
End Sub
