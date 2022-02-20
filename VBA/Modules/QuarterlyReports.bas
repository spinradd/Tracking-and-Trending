Attribute VB_Name = "QuarterlyReports"
Sub CreateQuarterly(start_date As Date, end_date As Date, filter_val As String)
    Dim ListObj As ListObject
    Dim MyName As String
    Dim MyRangeString As String
    Dim MyListExists As Boolean
    Dim TagTable As ListObject

'On Error GoTo MyErrorTrap

        Dim SheetName As String
        SheetName = "Quarterly Report"
        
        Application.DisplayAlerts = False
    'If sheet with report already exists, delete for new one
    For Each sheet In ThisWorkbook.Worksheets
    If sheet.Name = SheetName Then
        Sheets(SheetName).Delete
    End If
    Next sheet
    
        'creates new sheet for pivot table and chart
    Sheets.Add After:=ThisWorkbook.Worksheets("Control Center")
    ActiveSheet.Name = SheetName
    Set mreport_sheet = ThisWorkbook.Worksheets(SheetName)
        'turns on alerts after initial deletion/creation
    Application.DisplayAlerts = True
    
    Sheets(SheetName).Activate
    ActiveWindow.DisplayGridlines = False
    
    ' if excel table already exists, delete and replace with new one
    MyName = "QuarterlyTable"
    MyRangeString = "A2:I2"

    MyListExists = False
    For Each ListObj In Sheets(SheetName).ListObjects
       
        If ListObj.Name = MyName Then MyListExists = True
        
    Next ListObj
    
    If Not (MyListExists) Then
        Set TagTable = ActiveSheet.ListObjects.Add(xlSrcRange, Range(MyRangeString), , xlYes)
        TagTable.Name = MyName
    End If

'setting countermeasures table as source for monthly report data
Dim counter_tbl As ListObject
   Set counter_sheet = Worksheets("Countermeasures")
   Set counter_tbl = counter_sheet.ListObjects("Tbl_Counter")
   Set counter_col = counter_tbl.ListColumns

'formatting new monthly report table column headers
Set TagTable = Sheets(SheetName).ListObjects(MyName)
With TagTable
    With .HeaderRowRange
            .ColumnWidth = 11
           .VerticalAlignment = xlCenter
           .HorizontalAlignment = xlCenter
           .VerticalAlignment = xlCenter
           .Orientation = xlHorizontal
           .Font.Size = 12
           .Font.Bold = True
           .RowHeight = 20
    End With
End With


'naming new monthly report table column header
TagTable.HeaderRowRange(1, 1) = "Issue ID"
TagTable.HeaderRowRange(1, 2) = "Issue Date"
TagTable.HeaderRowRange(1, 3) = "Category"
TagTable.HeaderRowRange(1, 4) = "KPI"
TagTable.HeaderRowRange(1, 5) = "Issue"
TagTable.HeaderRowRange(1, 6) = "Cause"
TagTable.HeaderRowRange(1, 7) = "Countermeasure"
TagTable.HeaderRowRange(1, 8) = "Owner"
TagTable.HeaderRowRange(1, 10) = "Status"
TagTable.HeaderRowRange(1, 9) = "Date Closed"


Dim catval As Variant
catval = Sheets("Lookup Tag").Pivot_DD_Box.value
   
   'collect input for report
   TimeFrame = Application.InputBox( _
                            prompt:="Type in month and year for report", _
                            Title:="Report Info", _
                            Default:="Ex: September 2020", _
                            Type:=2)
   Dim SortBy As String
   SortBy = filter_val
    
    
    'collect input for report
   MyInput = start_date
   SortBy = InputBox("Sort by: 'Issue ID', 'Issue Date', 'Category', or 'KPI'  (Case Sensitive :( )")
    
    ' Allow user to end macro with Cancel in InputBox.
    If MyInput = "" Then Exit Sub
    ' Get the date value of the beginning of inputted month.
    StartDay = start_date
    ' Check if valid date but not the first of the month
    ' -- if so, reset StartDay to first day of month.
    If Day(StartDay) <> 1 Then
        StartDay = DateValue(Month(StartDay) & "/1/" & _
            Year(StartDay))
     End If
    
    'set month and year label to range
    Set MonthYearRange = mreport_sheet.Range("A1:I1")
       
    'formatting month and year label
    With MonthYearRange
          .HorizontalAlignment = xlCenterAcrossSelection
           .VerticalAlignment = xlCenter
           .Font.Size = 18
           .Font.Bold = True
           .RowHeight = 35
           .NumberFormat = "mmmm yyyy"
    End With
    
    
'Put inputted month and year fully spelling out into "a1".
   MonthYearRange.Cells(1, 1).value = Application.Text(TimeFrame, "mmmm yyyy")
   ' Set variable and get which day of the week the month starts.
   DayofWeek = Weekday(StartDay)
   ' Set variables to identify the year and month as separate
   ' variables.
   CurYear = Year(StartDay)
   CurMonth = Month(StartDay)
   ' Set variable and calculate the first day of the next month.
   FinalDay = DateSerial(CurYear, CurMonth + 1, 1)
   ' Place a "1" in cell position of the first day of the chosen
   ' month based on DayofWeek.
   
'Dimming arrays to contain information from countermeasures table. All entries are included, so essentially the
'4th spot in any array is that columns cell for the 4th row of the data body range of the countermeasures table
Dim Issue_ID() As Variant
Dim Questions() As Variant
Dim IssueTag1() As Variant
Dim IssueTag2() As Variant
Dim CauseCategory() As Variant
Dim CauseDetail() As Variant
Dim Issue_Date() As Variant
Dim category() As Variant
Dim KPI() As Variant
Dim Issue() As Variant
Dim Cause() As Variant
Dim Countermeasure() As Variant
Dim Owner() As Variant
Dim Status() As Variant

Dim ArrBase() As Variant
Dim col_val As String

 
            
            For Each Column In counter_tbl.HeaderRowRange
                col_val = Column.value
                'entry count is to determine if first cell and first row or not. If first, must dim array appropriately
                entry_count = 0
                'Will run through each cell in the countermeasures table
                For Each cell In counter_tbl.ListColumns(col_val).DataBodyRange
                    CRow = cell.row
                    i = 0
                            ' If issue date is within requested month, then that data will be added to the array according to the list column the cell is under
                            'crow - 1 because crow = 1 is actually row = 2 in data body range
                           If counter_tbl.ListColumns("Issue Date").DataBodyRange.Cells(CRow - 1, 1).value >= StartDay _
                           And counter_tbl.ListColumns("Issue Date").DataBodyRange.Cells(CRow - 1, 1).value < FinalDay Then
                                       'if first entry, redim to hold one spot "(0)"
                                       If entry_count = 0 Then
                                           ReDim Preserve ArrBase(0)
                                           ArrBase(0) = cell
                                        'For all subsequent entries extend array by 1 and enter contents in cell
                                       Else
                                           ReDim Preserve ArrBase(UBound(ArrBase) + 1)
                                           ArrBase(UBound(ArrBase)) = cell
                                       End If
                                       entry_count = entry_count + 1
                            
                            'If not during month but IS an Open issue with date before month, include in array and in report
                            ElseIf counter_tbl.ListColumns("Status").DataBodyRange.Cells(CRow - 1, 1).value = "Open" _
                            And counter_tbl.ListColumns("Issue Date").DataBodyRange.Cells(CRow - 1, 1).value < FinalDay Then
                                        If entry_count = 0 Then
                                            ReDim Preserve ArrBase(0)
                                            ArrBase(0) = cell
                                        Else
                                            ReDim Preserve ArrBase(UBound(ArrBase) + 1)
                                            ArrBase(UBound(ArrBase)) = cell
                                        End If
                                        entry_count = entry_count + 1
                            End If
                Next cell
                
                'Assign array to correct collection, named after the list columns in countermeasures report
                Select Case counter_tbl.ListColumns(col_val)
                    Case "Issue ID"
                        Issue_ID = ArrBase
                    Case "Questions"
                        Questions = ArrBase
                    Case "Issue Tier 1 Tag"
                        IssueTag1 = ArrBase
                    Case "Issue Tier 2 Tag"
                        IssueTag2 = ArrBase
                    Case "Category Cause"
                        CauseCategory = ArrBase
                    Case "Category Detail"
                        CauseDetail = ArrBase
                    Case "Issue Date"
                        Issue_Date = ArrBase
                    Case "Category"
                        category = ArrBase
                    Case "KPI"
                        KPI = ArrBase
                    Case "Issue"
                        Issue = ArrBase
                    Case "Cause"
                        Cause = ArrBase
                    Case "Countermeasure"
                        Countermeasure = ArrBase
                    Case "Owner"
                        Owner = ArrBase
                    Case "Status"
                        Status = ArrBase
                    Case "Date Closed"
                        DateClosed = ArrBase
                        
                End Select
            '''''CLEAR ARRAY
            Erase ArrBase
            Next
    'Adding arrays to collection
    Dim MArrays As New Collection
    MArrays.Add item:=Issue_ID, key:="Issue ID"
    MArrays.Add item:=Issue_Date, key:="Issue Date"
    MArrays.Add item:=category, key:="Category"
    MArrays.Add item:=KPI, key:="KPI"
    MArrays.Add item:=Issue, key:="Issue"
    MArrays.Add item:=Cause, key:="Cause"
    MArrays.Add item:=Countermeasure, key:="Countermeasure"
    MArrays.Add item:=Owner, key:="Owner"
    MArrays.Add item:=Status, key:="Status"
    MArrays.Add item:=DateClosed, key:="Date Closed"
    
    
    Dim UsedArray() As Variant
    
    'Add rows equal to number of entries or spots in array (can use any any array because they will be the same length)
    For Each Entry In MArrays("Issue ID")
        TagTable.ListRows.Add
    Next
    
    'For each column in monthly report table, populate column with entries from associate array from collection
    For Each Column In TagTable.HeaderRowRange
        col_val = Column.value
        x = 0
        UsedArray = MArrays(col_val)
        
        For Each cell In TagTable.ListColumns(col_val).DataBodyRange
             cell.value = UsedArray(x)
            x = x + 1
        Next
    Next
    
    'Formatting monthly report table
    With TagTable.HeaderRowRange
        .WrapText = True
    End With
    With TagTable.DataBodyRange
           .VerticalAlignment = xlCenter
           .HorizontalAlignment = xlCenter
           .Font.Size = 8
           .RowHeight = 40
           .Font.Bold = False
           .ColumnWidth = 5
           .WrapText = True
    End With
      With TagTable.ListColumns("Issue ID").DataBodyRange
           .VerticalAlignment = xlCenter
           .Font.Size = 8
           .Font.Bold = False
           .ColumnWidth = 12.88
    End With
      With TagTable.ListColumns("Category").DataBodyRange
           .VerticalAlignment = xlCenter
           .Font.Size = 8
           .Font.Bold = False
           .ColumnWidth = 8.88
    End With
     With TagTable.ListColumns("KPI").DataBodyRange
           .VerticalAlignment = xlCenter
           .Font.Size = 8
           .Font.Bold = False
           .ColumnWidth = 9.88
    End With
     With TagTable.ListColumns("Issue Date").DataBodyRange
           .VerticalAlignment = xlCenter
           .Font.Size = 8
           .Font.Bold = False
           .ColumnWidth = 8.88
    End With
    With TagTable.ListColumns("Issue").DataBodyRange
           .VerticalAlignment = xlCenter
           .Font.Size = 8
           .Font.Bold = False
           .ColumnWidth = 55
    End With
    With TagTable.ListColumns("Cause").DataBodyRange
           .VerticalAlignment = xlCenter
           .Font.Size = 8
           .Font.Bold = False
           .ColumnWidth = 55
    End With
    With TagTable.ListColumns("Countermeasure").DataBodyRange
           .VerticalAlignment = xlCenter
           .Font.Size = 8
           .Font.Bold = False
           .ColumnWidth = 55
    End With
    With TagTable.ListColumns("Owner").DataBodyRange
           .VerticalAlignment = xlCenter
           .Font.Size = 8
           .Font.Bold = False
           .ColumnWidth = 6.5
    End With
       With TagTable.ListColumns("Status").DataBodyRange
           .VerticalAlignment = xlCenter
           .Font.Size = 8
           .Font.Bold = False
           .ColumnWidth = 6.88
    End With
      With TagTable.ListColumns("Date Closed").DataBodyRange
           .VerticalAlignment = xlCenter
           .Font.Size = 8
           .Font.Bold = False
           .ColumnWidth = 10.13
    End With
    
                
'Implementing sort logic depending on user entry at beginning of script
With TagTable.Sort
                    .SortFields.Add key:=Range("MonthlyTable[Status]"), SortOn:=xlSortOnValues, Order:=xlDescending
                    .Header = xlYes
    
    Select Case SortBy
        Case "Issue ID"
                .SortFields.Add key:=Range("MonthlyTable[Issue ID]"), SortOn:=xlSortOnValues, Order:=xlAscending
                    .Header = xlYes
        Case "Issue Date"
               .SortFields.Add key:=Range("MonthlyTable[Issue Date]"), SortOn:=xlSortOnValues, Order:=xlAscending
                    .Header = xlYes
        Case "Category"
                .SortFields.Add key:=Range("MonthlyTable[Category]"), SortOn:=xlSortOnValues, Order:=xlAscending
                    .Header = xlYes
        Case "KPI"
                .SortFields.Add key:=Range("MonthlyTable[KPI]"), SortOn:=xlSortOnValues, Order:=xlAscending
                    .Header = xlYes
    End Select
    
        .Apply
    
End With
   
    'find last row of monthly report table
    lastrow = TagTable.DataBodyRange.End(xlDown).row
    Debug.Print lastrow
    
    'find bottom left cell of monthly report table
    Dim left As Range
    Dim right As Range
    Set left = Cells(lastrow, 1).Offset(2, 0)
    Set right = Cells(lastrow, 1).Offset(2, 1)
    
    'Create a new excel table two cells below monthly table, table is for category summary
    Dim both As Variant
    both = left.Address & ":" & right.Address
    
    CatName = "MonthlyCAT"
    
    MyListExists = False
    For Each ListObj In Sheets(SheetName).ListObjects
       
        If ListObj.Name = CatName Then MyListExists = True
        
    Next ListObj
    
    If Not (MyListExists) Then
        Set Cattbl = ActiveSheet.ListObjects.Add(xlSrcRange, Range(both), , xlYes)
        Cattbl.Name = CatName
    End If
    
    Dim CatArr() As Variant
    entry_count = 0
                    'create array to hold categories pulled from monthly report table
                    For Each cell In TagTable.ListColumns("Category").DataBodyRange
                            If entry_count = 0 Then
                                ReDim Preserve CatArr(0)
                                CatArr(0) = cell
                                'For all subsequent entries
                            Else
                                ReDim Preserve CatArr(UBound(CatArr) + 1)
                                CatArr(UBound(CatArr)) = cell
                            End If
                            'If ArrBase(UBound(ArrBase)) = Empty Then
                                'ReDim Preserve ArrBase(UBound(ArrBase) - 1)
                            'End If
                            entry_count = entry_count + 1
                    Next cell
                    
'format category table
Cattbl.HeaderRowRange(1, 1) = "Category"
Cattbl.HeaderRowRange(1, 2) = "Open"
Cattbl.HeaderRowRange(1, 3) = "Closed"
Cattbl.HeaderRowRange(1, 4) = "Total"
   
Dim cdict As Scripting.Dictionary
Set cdict = New Scripting.Dictionary

'rids duplicates and counts new dictionary
Set cdict = DuplicateCountToScript(CatArr)

For Each item In cdict.Keys
    Cattbl.ListRows.Add
Next

Worksheets(SheetName).ListObjects("MonthlyCAT").DataBodyRange.Cells(1, 1).Resize(cdict.Count, 1).Value2 = Application.Transpose(cdict.Keys)
Worksheets(SheetName).ListObjects("MonthlyCAT").DataBodyRange.Cells(1, 4).Resize(cdict.Count, 1).Value2 = Application.Transpose(cdict.Items)

    'counting each category instance of open and closed, for stats
    For Each CatCell In Worksheets(SheetName).ListObjects("MonthlyCAT").ListColumns("Category").DataBodyRange
       Debug.Print CatCell.value
       Debug.Print CatCell.Address
        For Each cell In TagTable.ListColumns("Category").DataBodyRange
             Debug.Print cell.value
             Debug.Print cell.Address
             Debug.Print cell.row
            If CatCell.value = cell.value Then
                Debug.Print TagTable.ListColumns("Status").DataBodyRange.Cells(cell.row, 1).value
                If TagTable.ListColumns("Status").DataBodyRange.Cells(cell.row, 1) = "Open" Then
                    o_kpi_count = o_kpi_count + 1
                Else
                    c_kpi_count = c_kpi_count + 1
                End If
            End If
        Next cell
    
    CatCell.Offset(0, 1) = o_kpi_count
    CatCell.Offset(0, 2) = c_kpi_count
    
    o_kpi_count = 0
    c_kpi_count = 0
    
    Next CatCell
    
    'inserting KPI summary table using same process as above
    lastrow = Cattbl.DataBodyRange.End(xlDown).row
    Debug.Print lastrow
    
    Set left = Cells(lastrow, 1).Offset(2, 0)
    Set right = Cells(lastrow, 1).Offset(2, 1)
    
    both = left.Address & ":" & right.Address
    
    KPIName = "MonthlyKPI"
    
    MyListExists = False
    For Each ListObj In Sheets(SheetName).ListObjects
       
        If ListObj.Name = KPIName Then MyListExists = True
        
    Next ListObj
    
    If Not (MyListExists) Then
        Set KPItbl = ActiveSheet.ListObjects.Add(xlSrcRange, Range(both), , xlYes)
        KPItbl.Name = KPIName
    End If
    
    Dim KPIarr() As Variant
    entry_count = 0
    
                    For Each cell In TagTable.ListColumns("KPI").DataBodyRange
                            If entry_count = 0 Then
                                ReDim Preserve KPIarr(0)
                                KPIarr(0) = cell
                                'For all subsequent entries
                            Else
                                ReDim Preserve KPIarr(UBound(KPIarr) + 1)
                                KPIarr(UBound(KPIarr)) = cell
                            End If
                            'If ArrBase(UBound(ArrBase)) = Empty Then
                                'ReDim Preserve ArrBase(UBound(ArrBase) - 1)
                            'End If
                            entry_count = entry_count + 1
                    Next cell
                    
                    'MsgBox Join(KPIarr, vbCrLf)
                    
KPItbl.HeaderRowRange(1, 1) = "KPI"
KPItbl.HeaderRowRange(1, 2) = "Open"
KPItbl.HeaderRowRange(1, 3) = "Closed"
KPItbl.HeaderRowRange(1, 4) = "Total"
   
Dim kdict As Scripting.Dictionary
Set kdict = New Scripting.Dictionary

'rids duplicates and counts new dictionary
Set kdict = DuplicateCountToScript(KPIarr)

For Each item In kdict.Keys
    KPItbl.ListRows.Add
Next

Worksheets(SheetName).ListObjects("MonthlyKPI").DataBodyRange.Cells(1, 1).Resize(kdict.Count, 1).Value2 = Application.Transpose(kdict.Keys)
Worksheets(SheetName).ListObjects("MonthlyKPI").DataBodyRange.Cells(1, 4).Resize(kdict.Count, 1).Value2 = Application.Transpose(kdict.Items)


    For Each KPICell In Worksheets(SheetName).ListObjects("MonthlyKPI").ListColumns("KPI").DataBodyRange
       Debug.Print KPICell.value
       Debug.Print KPICell.Address
        For Each cell In TagTable.ListColumns("KPI").DataBodyRange
             Debug.Print cell.value
             Debug.Print cell.Address
             Debug.Print cell.row
            If KPICell.value = cell.value Then
                Debug.Print TagTable.ListColumns("Status").DataBodyRange.Cells(cell.row, 1).value
                If TagTable.ListColumns("Status").DataBodyRange.Cells(cell.row, 1) = "Open" Then
                    o_kpi_count = o_kpi_count + 1
                Else
                    c_kpi_count = c_kpi_count + 1
                End If
            End If
        Next cell
    
    KPICell.Offset(0, 1) = o_kpi_count
    KPICell.Offset(0, 2) = c_kpi_count
    
    o_kpi_count = 0
    c_kpi_count = 0
    
    Next KPICell
    
    Exit Sub


MyErrorTrap:
       MsgBox "You may not have entered your Year correctly." _
           & Chr(13) & "Please input full month name with 4 digits for the Year." _
           & Chr(13) & Chr(13) & "OR the month and year you entered has no entries, please select another month and year" _
           & Chr(13) & Chr(13) & "OR you have mis-typed the 'sort by' option, please try again"
   Exit Sub
       Resume

End Sub

