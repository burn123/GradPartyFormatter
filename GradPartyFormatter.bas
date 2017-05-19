Attribute VB_Name = "GradPartyFormatter"
Public WB As Workbook
Public Sheet As Worksheet
' The number of rows when there are no blank rows between dates
Public RowCount As Integer
' The number of rows when there are blank rows between dates
Public SpacedRowCount As Integer



Sub FormatGradParties()
'
' FormatGradParties Macro
' Runs all of the different parts to format the spreadsheet
'
    Dim Filename As String
    Filename = Right(ThisWorkbook.FullName, Len(ThisWorkbook.FullName) - InStrRev(ThisWorkbook.FullName, "\"))

    Application.Run "'" + Filename + "'!SortRows"
    Application.Run "'" + Filename + "'!DeleteBlankRows"
    Application.Run "'" + Filename + "'!StyleCells"
    Application.Run "'" + Filename + "'!AddBlankRows"
    Application.Run "'" + Filename + "'!AddBorders"
End Sub

Sub SortRows()
'
' SortDeleteandAddBlankRow Macro
' Sorts the Rows according to the date of the party,
' then the start time, then the end time
'
    Set WB = ActiveWorkbook
    Set Sheet = WB.Worksheets("Sheet1")
    
    Dim SF As SortFields
    Set SF = Sheet.Sort.SortFields
    
    SF.Clear
    ' The date column
    SF.Add Key:=Columns("D") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ' The start time column
    SF.Add Key:=Columns("B") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
        "8:00am,9:00am,10:00am,11:00am,11:30am,12:00pm,1:00pm,1:30pm,2:00pm,3:00pm,4:00pm,4:30pm,5:00pm,6:00pm,7:00pm,8:00pm", DataOption:=xlSortNormal
    ' The end time column
    SF.Add Key:=Columns("C") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, CustomOrder:= _
        "8:00am,9:00am,10:00am,11:00am,11:30am,12:00pm,1:00pm,1:30pm,2:00pm,3:00pm,4:00pm,4:30pm,5:00pm,6:00pm,7:00pm,8:00pm", DataOption:=xlSortNormal
    ' The name column
    SF.Add Key:=Columns("A") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
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
  Set r = Sheet.Range("1:50")
  For i = r.Rows.Count To 1 Step (-1)
    ' Deletes the row if it is empty
    If WorksheetFunction.CountA(r.Rows(i)) = 0 Then r.Rows(i).Delete
  Next
  RowCount = Cells(Sheet.Rows.Count, "A").End(xlUp).row
End Sub

Sub AddBlankRows()
'
' AddBlankRows Macro
' Adds blank rows in between different dates,
' so that all parties on the same date are grouped
'
    Dim i As Long
    
    LastRow = Cells(Rows.Count, "D").End(xlUp).row
    ' Iterate backwards, so row numbers stay absolute
    For i = RowCount To 3 Step -1
        ' Checks if the date is equal between 2 columns
        If Cells(i, "D") <> Cells(i - 1, "D") Then
            Rows(i).Insert
        End If
    Next i
    ' Set how many rows there are
    SpacedRowCount = Cells(Sheet.Rows.Count, "A").End(xlUp).row
End Sub

Sub StyleCells()
'
' StyleCells Macro
' Formats the cells to your liking
'
    Columns("A").HorizontalAlignment = xlLeft
    Columns("B:D").HorizontalAlignment = xlCenter
    Columns("E").HorizontalAlignment = xlLeft
    Rows("1").HorizontalAlignment = xlCenter
    ' Change dates to correct date format
    Columns("D").NumberFormat = "[$-x-sysdate]dddd, mmmm dd, yyyy"
    ' Change font of regular text to desired
    With Rows("2:" & RowCount).Font
        .Name = "Perpetua"
        .Size = 13
    End With
    
    ' Bolden the city name in location, if applicable
    Dim commaPos As Integer, row As Integer
    Dim cell As Range
    
    For row = 2 To RowCount
        Set cell = Cells(row, "E")
        commaPos = InStr(cell.Value, ",")
        If commaPos <> 0 Then
            cell.Characters(Start:=commaPos).Font.FontStyle = "Italic"
        End If
    
    
        Next row
    
    With Rows("2:" & RowCount)
        .VerticalAlignment = xlCenter
    End With
End Sub

Sub AddBorders()
'
' AddBorders Macro
' Adds borders around all of the used cells
'
    With Range("A1:E" & SpacedRowCount).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 1
    End With
End Sub
