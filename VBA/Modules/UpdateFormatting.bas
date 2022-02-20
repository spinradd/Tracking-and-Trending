Attribute VB_Name = "UpdateFormatting"
Sub UpdateCountermeasureFormatting()
Dim counter_tbl As ListObject
Dim counter_sheet As Worksheet


On Error Resume Next

Set counter_sheet = ThisWorkbook.Worksheets("Countermeasures")
Set counter_tbl = ThisWorkbook.Worksheets("Countermeasures").ListObjects("Tbl_Counter")

    
 counter_tbl.Range.ClearFormats

 counter_tbl.TableStyle = "TableStyleMedium9"

For Each cell In counter_tbl.HeaderRowRange
    cell.Interior.Color = RGB(0, 176, 240)
    cell.Font.Color = RGB(0, 0, 0)
Next

For Each cell In counter_tbl.HeaderRowRange
    Select Case cell.value
            
            Case "Issue ID"
                IssueID_val = cell.Column
            Case "Questions"
                questions_colval = cell.Column
            Case "Issue Date"
                issue_date_colval = cell.Column
            Case "KPI"
                KPI_ColVal = cell.Column
            Case "Issue"
                issue_val = cell.Column
            Case "Status"
                status_colval = cell.Column
            Case "Owner"
                owner_colval = cell.Column
            Case "Early and Overdue Differential"
                Differential_colval = cell.Column
    End Select
Next cell

col = 0
For Each cell In counter_tbl.HeaderRowRange
    col = col + 1
    If cell.Column > IssueID_val And cell.Column < issue_date_colval Then
    
        counter_tbl.ListColumns(col).DataBodyRange.Width = 21.5
    
        For Each cell2 In counter_tbl.ListColumns(col).DataBodyRange
            If cell2 = Empty Then
                cell2.Interior.Color = RGB(255, 204, 102)
                cell2.Font.Color = RGB(0, 0, 0)
            End If
        Next
    End If
    
    If cell.Column > KPI_ColVal And cell.Column < issue_val Then
        
        counter_tbl.ListColumns(col).DataBodyRange.Width = 21.5
    
        For Each cell2 In counter_tbl.ListColumns(col).DataBodyRange
            If cell2 = Empty Then
                cell2.Interior.Color = RGB(255, 204, 102)
                cell2.Font.Color = RGB(0, 0, 0)
            End If
        Next
    End If
    
    If cell.Column >= issue_val And cell.Column < owner_colval Then
            
       counter_tbl.ListColumns(col).DataBodyRange.Width = 55

           For Each cell2 In counter_tbl.ListColumns(col).DataBodyRange
            If cell2 = Empty Then
                cell2.Interior.Color = RGB(255, 204, 102)
                cell2.Font.Color = RGB(0, 0, 0)
            End If
           Next
    End If
Next

For Each cell In counter_tbl.ListColumns("Category").DataBodyRange
    If cell = Empty Then
        cell.Interior.Color = RGB(225, 0, 0)
        cell.Font.Color = RGB(0, 0, 0)
    End If
Next

For Each cell In counter_tbl.ListColumns("KPI").DataBodyRange
    If cell = Empty Then
        cell.Interior.Color = RGB(255, 0, 0)
        cell.Font.Color = RGB(0, 0, 0)
    End If
Next

For Each cell In counter_tbl.ListColumns("Issue Date").DataBodyRange
    If cell = Empty Then
        cell.Interior.Color = RGB(235, 0, 0)
        cell.Font.Color = RGB(0, 0, 0)
    End If
Next

 counter_tbl.ListColumns("Issue Date").DataBodyRange.NumberFormat = "dd-mmm-yy"
 counter_tbl.ListColumns("Date Due").DataBodyRange.NumberFormat = "dd-mmm-yy"
 counter_tbl.ListColumns("Date Closed").DataBodyRange.NumberFormat = "dd-mmm-yy"
 
 For Each cell In counter_tbl.ListColumns("Date Closed").DataBodyRange
    row = row + 1
    If cell = Empty Or cell = "" Then
        counter_tbl.ListColumns("Status").DataBodyRange(row, 1).value = "Open"
        counter_tbl.ListColumns("Status").DataBodyRange(row, 1).Interior.Color = RGB(235, 0, 0)
        counter_tbl.ListColumns("Status").DataBodyRange(row, 1).Font.Color = RGB(0, 0, 0)
    Else
        counter_tbl.ListColumns("Status").DataBodyRange(row, 1).value = "Closed"
        counter_tbl.ListColumns("Status").DataBodyRange(row, 1).Interior.Color = RGB(0, 176, 80)
        counter_tbl.ListColumns("Status").DataBodyRange(row, 1).Font.Color = RGB(0, 0, 0)
    End If
Next

    With counter_tbl.Range
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
           .Font.Name = "Calibri"
           .WrapText = True
       For Each iCells In counter_tbl.Range
            iCells.BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThin
        Next iCells
    End With
    
rowcount = 1

Dim CatCell As Variant
Dim KPICell As Variant
Dim IssueDateCell As Variant
Dim IssueCell As Variant
Dim CauseCell As Variant
Dim CounterCell As Variant
Dim OwnerCell As Variant
Dim DateDueCell As Variant '
Dim DateClosedCell As Variant
Dim StatusCell As Variant

Dim RowColl As New Collection

rowcount = 1
For row = 1 To counter_tbl.DataBodyRange.Rows.Count
    'Debug.Print Row.value
    EmptyCheck = 0
    Set RowColl = Nothing

    CatCell = counter_tbl.ListColumns("Category").DataBodyRange(row, 1).value
    KPICell = counter_tbl.ListColumns("KPI").DataBodyRange(row, 1).value
    IssueDateCell = counter_tbl.ListColumns("Issue Date").DataBodyRange(row, 1).value
    IssueCell = counter_tbl.ListColumns("Issue").DataBodyRange(row, 1).value
    CauseCell = counter_tbl.ListColumns("Cause").DataBodyRange(row, 1).value
    CounterCell = counter_tbl.ListColumns("Countermeasure").DataBodyRange(row, 1).value
    OwnerCell = counter_tbl.ListColumns("Owner").DataBodyRange(row, 1).value
    DateDueCell = counter_tbl.ListColumns("Date Due").DataBodyRange(row, 1).value
    DateClosedCell = counter_tbl.ListColumns("Date Closed").DataBodyRange(row, 1).value
    StatusCell = counter_tbl.ListColumns("Status").DataBodyRange(row, 1).value
    
    
    RowColl.Add (CatCell)
    RowColl.Add (KPICell)
    RowColl.Add (IssueDateCell)
    RowColl.Add (IssueCell)
    RowColl.Add (CauseCell)
    RowColl.Add (CounterCell)
    RowColl.Add (OwnerCell)
    RowColl.Add (DateDueCell)
    RowColl.Add (DateClosedCell)
    RowColl.Add (StatusCell)
    
    For Each cell In RowColl
        'Debug.Print Cell
        If cell = Empty Then
        EmptyCheck = EmptyCheck + 1
        End If
    Next
    
    If EmptyCheck = 10 Then
            counter_tbl.DataBodyRange.EntireRow(row).Delete
    End If
Next

End Sub

