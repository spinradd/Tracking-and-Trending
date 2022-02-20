Attribute VB_Name = "Basic_Reports"
Sub CreateBasicReport(start_date As Date, end_date As Date, filter_val As String)
    'creates monthly report sheet, and table
    
    
    Dim ListObj As ListObject
    Dim MyName As String
    Dim MyRangeString As String
    Dim MyListExists As Boolean
    Dim TagTable As ListObject

'On Error GoTo MyErrorTrap

        Dim SheetName As String
        SheetName = "Monthly Report"
        
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
    MyName = "MonthlyTable"
    MyRangeString = "A2:I2"

    MyListExists = False                    'assume table doesn't exist
    For Each ListObj In Sheets(SheetName).ListObjects
       
        If ListObj.Name = MyName Then MyListExists = True
        
    Next ListObj
    
    If Not (MyListExists) Then          'is table doesn;t exist, create
        Set TagTable = ActiveSheet.ListObjects.Add(xlSrcRange, Range(MyRangeString), , xlYes)
        TagTable.Name = MyName
        
        Else                            'if it does assign variable
        Set TagTable = Sheets(SheetName).ListObject(MyName)
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
           .Font.Name = "Arial"
           '.Font.Color = RGB(0, 153, 220)
           '.Font.Color = RGB(0, 38, 76)
           .Font.Color = RGB(10, 50, 84)
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


   'collect input for report filtering
   MyInput = start_date
   Dim SortBy As String
   SortBy = filter_val
    
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
           .Font.Name = "Arial"
           '.Font.Color = RGB(0, 153, 220)
           '.Font.Color = RGB(0, 38, 76)
           .Font.Color = RGB(10, 50, 84)
           .Font.Bold = True
           .RowHeight = 35
           .NumberFormat = "mmmm yyyy"
    End With
    
    
'Put inputted month and year fully spelling out into "cell a1".
   MonthYearRange.Cells(1, 1).value = CStr(MonthName(Month(start_date)) & " " & Year(start_date) & " - " & MonthName(Month(end_date)) & " " & Year(end_date))
   ' Set variable and get which day of the week the month starts.
   DayofWeek = Weekday(StartDay)
   ' Set variables to identify the year and month as separate
   ' variables.
   CurYear = Year(StartDay)
   CurMonth = Month(StartDay)
   ' Set variable and calculate the first day of the next month.
   FinalDay = DateValue(end_date)
   ' Place a "1" in cell position of the first day of the chosen
   ' month based on DayofWeek.
   
'Dimming arrays to contain information from countermeasures table. All entries are included
'so essentially the 4th spot in any array is that array/columns 4th cell from the 4th data body row of the countermeasures table
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
                row_counter = 0
                For Each cell In counter_tbl.ListColumns(col_val).DataBodyRange     'for each cell in countermeasures column
                    row_counter = row_counter + 1
                    i = 0
                            ' If issue date of row/entry is within requested month, then that data will be added to the array
                            ' The array will then be added to the the collection of arrays to be pulled later into the monthly report table
                           If counter_tbl.ListColumns("Issue Date").DataBodyRange.Cells(row_counter, 1).value >= StartDay _
                           And counter_tbl.ListColumns("Issue Date").DataBodyRange.Cells(row_counter, 1).value < FinalDay Then
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
                                       
                  'unlock loop if you want all open entries included!! Currently, code gathers entries where a new issue was opened during month.
                  'Code does not gather entries that opened in previous months (and remain open). Code below adds that feature:
                  '
                  '          'If not during month but IS an Open issue with date before month, include in array and in report
                  '          ElseIf counter_tbl.ListColumns("Status").DataBodyRange.Cells(row_counter, 1).value = "Open" _
                  '          And counter_tbl.ListColumns("Issue Date").DataBodyRange.Cells(row_counter, 1).value < FinalDay Then
                  '                      If entry_count = 0 Then
                  '                          ReDim Preserve ArrBase(0)
                  '                          ArrBase(0) = Cell
                  '                      Else
                  '                          ReDim Preserve ArrBase(UBound(ArrBase) + 1)
                  '                          ArrBase(UBound(ArrBase)) = Cell
                  '                      End If
                  '                      entry_count = entry_count + 1
                            End If
                Next cell
                
                ArrayResults = DoesArrayExist(ArrBase)
                If ArrayResults(0) = False Then
                    GoTo ArrayDoesntExist  'if array does not exist then go to next column
                End If
                
                'Assign array to correct variable, later, add variable to appropriate collection
                Select Case counter_tbl.ListColumns(col_val)
                    Case "Issue ID"
                        Issue_ID = ArrBase
                    Case "Questions"
                        Questions = ArrBase
                    Case "Issue Tier 1 Tag"
                        IssueTag1 = ArrBase
                    Case "Issue Tier 2 Tag"
                        IssueTag2 = ArrBase
                    Case "Cause Category"
                        CauseCategory = ArrBase
                    Case "Cause Detail"
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
            
ArrayDoesntExist:
            
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
            .Font.Name = "Arial"
           '.Font.Color = RGB(0, 153, 220)
           '.Font.Color = RGB(0, 38, 76)
           .Font.Color = RGB(10, 50, 84)
           .ColumnWidth = 5
           .WrapText = True
    End With
      With TagTable.ListColumns("Issue ID").DataBodyRange
           .VerticalAlignment = xlCenter
           .Font.Size = 8
           .Font.Bold = False
           .ColumnWidth = 12.88
            .Font.Name = "Arial"
           '.Font.Color = RGB(0, 153, 220)
           '.Font.Color = RGB(0, 38, 76)
           .Font.Color = RGB(10, 50, 84)
    End With
      With TagTable.ListColumns("Category").DataBodyRange
           .VerticalAlignment = xlCenter
           .Font.Size = 8
           .Font.Bold = False
           .ColumnWidth = 8.88
            .Font.Name = "Arial"
           '.Font.Color = RGB(0, 153, 220)
           '.Font.Color = RGB(0, 38, 76)
           .Font.Color = RGB(10, 50, 84)
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
           .NumberFormat = "dd-mmm-yyyy"
            .Font.Name = "Arial"
           '.Font.Color = RGB(0, 153, 220)
           '.Font.Color = RGB(0, 38, 76)
           .Font.Color = RGB(10, 50, 84)
    End With
    With TagTable.ListColumns("Issue").DataBodyRange
           .VerticalAlignment = xlCenter
           .Font.Size = 8
           .Font.Bold = False
           .ColumnWidth = 55
            .Font.Name = "Arial"
           '.Font.Color = RGB(0, 153, 220)
           '.Font.Color = RGB(0, 38, 76)
           .Font.Color = RGB(10, 50, 84)
    End With
    With TagTable.ListColumns("Cause").DataBodyRange
           .VerticalAlignment = xlCenter
           .Font.Size = 8
           .Font.Bold = False
           .ColumnWidth = 55
            .Font.Name = "Arial"
           '.Font.Color = RGB(0, 153, 220)
           '.Font.Color = RGB(0, 38, 76)
           .Font.Color = RGB(10, 50, 84)
    End With
    With TagTable.ListColumns("Countermeasure").DataBodyRange
           .VerticalAlignment = xlCenter
           .Font.Size = 8
           .Font.Bold = False
           .ColumnWidth = 55
            .Font.Name = "Arial"
           '.Font.Color = RGB(0, 153, 220)
           '.Font.Color = RGB(0, 38, 76)
           .Font.Color = RGB(10, 50, 84)
    End With
    With TagTable.ListColumns("Owner").DataBodyRange
           .VerticalAlignment = xlCenter
           .Font.Size = 8
           .Font.Bold = False
           .ColumnWidth = 6.5
            .Font.Name = "Arial"
           '.Font.Color = RGB(0, 153, 220)
           '.Font.Color = RGB(0, 38, 76)
           .Font.Color = RGB(10, 50, 84)
    End With
       With TagTable.ListColumns("Status").DataBodyRange
           .VerticalAlignment = xlCenter
           .Font.Size = 8
           .Font.Bold = False
           .ColumnWidth = 6.88
            .Font.Name = "Arial"
           '.Font.Color = RGB(0, 153, 220)
           '.Font.Color = RGB(0, 38, 76)
           .Font.Color = RGB(10, 50, 84)
    End With
      With TagTable.ListColumns("Date Closed").DataBodyRange
           .VerticalAlignment = xlCenter
           .Font.Size = 8
           .Font.Bold = False
           .ColumnWidth = 10.13
            .Font.Name = "Arial"
           '.Font.Color = RGB(0, 153, 220)
           '.Font.Color = RGB(0, 38, 76)
           .Font.Color = RGB(10, 50, 84)
    End With
    
    TagTable.TableStyle = "TableStyleMedium23"
    
                
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
    lastrow = TagTable.Range.Rows.Count
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
    
    MyListExists = False                                    'check if table exists, if it doesnt create it
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
                    
                ArrayResults = DoesArrayExist(CatArr)
                If ArrayResults(0) = False Then
                            MsgBox "There must be a 'Category' column to continue" 'if array does not exist then go to next column
                            Exit Sub
                End If
                    
                
                    
'format category table
Cattbl.HeaderRowRange(1, 1) = "Category"
Cattbl.HeaderRowRange(1, 2) = "Open"
Cattbl.HeaderRowRange(1, 3) = "Closed"
Cattbl.HeaderRowRange(1, 4) = "Total"

Cattbl.TableStyle = "TableStyleMedium23"
   
Dim cdict As Scripting.Dictionary
Set cdict = New Scripting.Dictionary

'rids duplicates and counts new dictionary
Set cdict = DuplicateCountToScript(CatArr)          'turn array into dictionary
                                                    'this gets rid of and adds up the duplicates

For Each item In cdict.Keys
    Cattbl.ListRows.Add
Next
                                    'paste into table
Worksheets(SheetName).ListObjects("MonthlyCAT").DataBodyRange.Cells(1, 1).Resize(cdict.Count, 1).Value2 = Application.Transpose(cdict.Keys)
Worksheets(SheetName).ListObjects("MonthlyCAT").DataBodyRange.Cells(1, 4).Resize(cdict.Count, 1).Value2 = Application.Transpose(cdict.Items)

    'counting each category instance of open and closed, for stats
    For Each CatCell In Worksheets(SheetName).ListObjects("MonthlyCAT").ListColumns("Category").DataBodyRange       'for each cell in category col of category table
        For Each cell In TagTable.ListColumns("Category").DataBodyRange                                             'for each cell in main monthly report table
            If CatCell.value = cell.value Then                                                                      'if categories match, keep track if entry is open or closed
                If TagTable.ListColumns("Status").DataBodyRange.Cells(cell.row, 1) = "Open" Then
                    o_kpi_count = o_kpi_count + 1
                Else
                    c_kpi_count = c_kpi_count + 1
                End If
            End If
            
        Next cell
                                        'display results in table
    CatCell.Offset(0, 1) = o_kpi_count
    CatCell.Offset(0, 2) = c_kpi_count
    
    o_kpi_count = 0
    c_kpi_count = 0
    
    Next CatCell
    
     With Cattbl.HeaderRowRange         'futher formatting for cat table
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
          .Font.Name = "Arial"
           '.Font.Color = RGB(0, 153, 220)
           '.Font.Color = RGB(0, 38, 76)
           .Font.Color = RGB(10, 50, 84)
    End With
    With Cattbl.DataBodyRange
           .VerticalAlignment = xlCenter
           .HorizontalAlignment = xlCenter
           .WrapText = True
             .Font.Name = "Arial"
           '.Font.Color = RGB(0, 153, 220)
           '.Font.Color = RGB(0, 38, 76)
           .Font.Color = RGB(10, 50, 84)
    End With

    
    
    
    'inserting KPI summary table using same process as above
    lastrow = Cattbl.DataBodyRange.End(xlDown).row
    
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
KPItbl.HeaderRowRange(1, 1) = "Category"
KPItbl.HeaderRowRange(1, 2) = "KPI"
KPItbl.HeaderRowRange(1, 3) = "Opened Issues"
   
Dim kdict As Scripting.Dictionary
Set kdict = New Scripting.Dictionary

'rids duplicates and counts new dictionary
Set kdict = DuplicateCountToScript(KPIarr)

For Each item In kdict.Keys
    KPItbl.ListRows.Add
Next

Worksheets(SheetName).ListObjects("MonthlyKPI").DataBodyRange.Cells(1, 2).Resize(kdict.Count, 1).Value2 = Application.Transpose(kdict.Keys)
'Worksheets(SheetName).ListObjects("MonthlyKPI").DataBodyRange.Cells(1, 5).Resize(kdict.Count, 1).Value2 = Application.Transpose(kdict.Items)


    For Each KPICell In Worksheets(SheetName).ListObjects("MonthlyKPI").ListColumns("KPI").DataBodyRange
        
        Debug.Print KPICell.value
        
        row_count = 0
        For Each cell In TagTable.ListColumns("KPI").DataBodyRange
             row_count = row_count + 1
             Count = 0
             N = 0
             Do While Count <> 1
                
                If KPI(N) = KPICell.value Then
                
                KPICell.Cells.Offset(0, -1) = category(N)
                Count = Count + 1
                Else
                
                N = N + 1
                
                End If
             Loop
             
            If KPICell.value = cell.value Then
                'Debug.Print TagTable.ListColumns("Status").DataBodyRange.Cells(Cell.row, 1).value
                If TagTable.ListColumns("Issue Date").DataBodyRange.Cells(row_count, 1) >= start_date And _
                    TagTable.ListColumns("Issue Date").DataBodyRange.Cells(row_count, 1) < end_date Then
                    
                     'Debug.Print TagTable.ListColumns("Issue Date").DataBodyRange.Cells(row_count, 1).Address
                    'Debug.Print TagTable.ListColumns("Issue Date").DataBodyRange.Cells(row_count, 1).Address
                    
                    o_kpi_count = o_kpi_count + 1
                Else
                    
                    'Debug.Print TagTable.ListColumns("Issue Date").DataBodyRange.Cells(row_count, 1).Address
                    'Debug.Print TagTable.ListColumns("Issue Date").DataBodyRange.Cells(row_count, 1).Address
                    
                    'c_kpi_count = c_kpi_count + 1
                End If
            End If
        Next cell
    
    KPICell.Offset(0, 1) = o_kpi_count
    'kpicell.Offset(0, 2) = c_kpi_count
    
    o_kpi_count = 0
    'c_kpi_count = 0
    
    Next KPICell
    
      With KPItbl.HeaderRowRange
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
    End With
    With KPItbl.DataBodyRange
           .VerticalAlignment = xlCenter
           .HorizontalAlignment = xlCenter
           .WrapText = True
             .Font.Name = "Arial"
           '.Font.Color = RGB(0, 153, 220)
           '.Font.Color = RGB(0, 38, 76)
           .Font.Color = RGB(10, 50, 84)
    End With
    
    row_count = 0
    x = 1
    Do While x <= KPItbl.ListColumns("Opened Issues").DataBodyRange.Rows.Count
        cell = KPItbl.ListColumns("Opened Issues").DataBodyRange(x, 1)
        row_counter = row_counter + 1
        If KPItbl.ListColumns("Opened Issues").DataBodyRange(x, 1).value = 0 Then
            KPItbl.DataBodyRange.EntireRow(x).Delete
            x = x - 1
        End If
        x = 1 + x
     Loop
    
    KPItbl.TableStyle = "TableStyleMedium23"
    
    With KPItbl.Sort
    
        .SortFields.Add key:=Range("MonthlyKPI[Category]"), Order:=xlAscending
        .SortFields.Add key:=Range("MonthlyKPI[KPI]"), Order:=xlAscending
        .Header = xlYes
        .Apply
    
End With
    Exit Sub


MyErrorTrap:
       MsgBox "You may not have entered your Year correctly." _
           & Chr(13) & "Please input full month name with 4 digits for the Year." _
           & Chr(13) & Chr(13) & "OR the month and year you entered has no entries, please select another month and year" _
           & Chr(13) & Chr(13) & "OR you have mis-typed the 'sort by' option, please try again"
   Exit Sub
       Resume

End Sub

Sub CreatePPTReport(num_per_slides As Integer, FontSize As Double, start_date As Date)
'takes out Category table, takes out trend chart,changes running table to house KPI not tags,
'applies template to slides

     Dim eApp As Excel.Application
    Dim wb As Excel.Workbook
    'Set eApp = New Excel.Application
    Dim PPApp As PowerPoint.Application
    Dim PPPresentation As PowerPoint.Presentation
    Dim PPLayout As CustomLayout
    Dim PPSlide  As PowerPoint.Slide
    Dim PPShape As PowerPoint.Shape
    Dim PPCharts As Excel.ChartObject
    Dim MyTitleSlide As Variant
    Dim MyTitleBox As Object
    Dim MyTitle As String
    Dim MyTitleBox2 As Object
    Dim MyTitle2 As String
    Dim MyReportSlide As Variant
    Dim MyReportTitle As Object
    Dim Report_tbl As ListObject
    Dim SummarySlide As Variant
    Dim SummaryTitle As Object
    Dim SummaryTblC1 As Variant
    Dim SummaryTblC2 As Table
    Dim SummaryTblK1 As Variant
    Dim SummaryTblK2 As Table
    Dim Summary_tbl2 As ListObject
    Dim AddReport As Object
 
    'Set variable to designate Powerpoint Application
    Set PPApp = New PowerPoint.Application
    
    Excel.Application.DisplayAlerts = False
    PPApp.DisplayAlerts = False
    
    'Add Presentation
     Set PPPresentation = PPApp.Presentations.Add
    
    AppActivate Application.Caption
    
    'create title and subtitle for presentation '''''''''''''''''''''''''''''
    Set MyTitleSlide = PPApp.ActivePresentation.Slides.Add(1, ppLayoutTitle)
    
    'link new powerpoint presentation to template
    PPApp.ActivePresentation.ApplyTemplate _
    Environ("USERPROFILE") & "\EXAMPLE\EXAMPLE\Powerpoint Design Template .potx"
    
    
    MyTitle2 = DateValue(Worksheets("Monthly Report").Range("A1").value) 'title is the month/year on excel monthly report sheet
    
    'Get month and year for title slide
    Dim mn As String
        mn = MonthName(Month(MyTitle2))
    Dim yr As Variant
        yr = Year(MyTitle2)
    MyTitle2 = mn & " " & yr
    
     'Input and format title
     With MyTitleSlide
        Set MyTitleBox = .Shapes.Title
            With MyTitleBox
                With .TextFrame.TextRange
                    .Text = "Monthly MDI Report"
                    
                    With .Font
                        .Bold = msoFalse
                        '.Name = "Times New Roman"
                        .Size = 50
                        '.Color = RGB(0, 0, 0)
                    End With
                End With
            End With
        'Input and format subtitle
        Set MyTitleBox2 = .Shapes(2)
            With MyTitleBox2
                With .TextFrame.TextRange
                    .Text = MyTitle2
                    With .Font
                        .Bold = msoFalse
                        '.Name = "Times New Roman"
                        .Size = 25
                        '.Color = RGB(0, 0, 0)
                    End With
                End With
            End With
        '.FollowMasterBackground = msoFalse
        '.Background.Fill.Solid
        '.Background.Fill.ForeColor.RGB = RGB(222, 220, 216)
    End With

    'Create Category and KPI Chart Slide
    Set SummarySlide = PPApp.ActivePresentation.Slides.Add(PPApp.ActivePresentation.Slides.Count + 1, ppLayoutTitleOnly)

    With SummarySlide
        'Copy the chart in excel
        Worksheets("Monthly Report").ListObjects("MonthlyKPI").Range.Copy
        'wait action to allow VBA to copy required entries [if removed may cause error]
        Application.Wait (Now + TimeValue("0:00:01"))
    
        'paste chart in powerpoint, assign variable to table
        Set SummaryTblK1 = .Shapes.Paste
        Set SummaryTblK2 = .Shapes(2).Table
        Set SummaryTitle = .Shapes.Title
    End With
            
    'Change title format of summary slide'''''''''''''''''''''''''''''''''''''''''''''''
    With SummarySlide
                With SummaryTitle
                    With .TextFrame.TextRange
                      'Logic for naming convention of new slide
                        .Text = "Monthly Summary"
                         With .Font
                                .Bold = msoFalse
                                '.Name = "Times New Roman"
                                .Size = 45
                                '.Color = RGB(0, 0, 0)
                        End With
                End With
            End With
    
            pptcolumns = SummaryTblK2.Columns.Count
            pptrows = SummaryTblK2.Rows.Count
            
        ' alter format of table within slide''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        With SummaryTblK2
            
            'scale up table size, must be this way ->
            PPApp.ActivePresentation.Slides(2).Shapes(2).Table.ScaleProportionally (2)
            
            With PPApp.ActivePresentation.PageSetup
                'Center Horizontally
                    PPApp.ActiveWindow.View.GotoSlide (PPApp.ActivePresentation.Slides.Count)
                    PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count).Shapes(2).Width = .SlideWidth
                    PPApp.ActivePresentation.Slides(2).Shapes(2).left = 0
                'Center Vertically
                    PPApp.ActivePresentation.Slides(2).Shapes(2).Top = PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count).Shapes(1).Top + PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count).Shapes(1).Height
                    PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count).Shapes(2).Height = (.SlideHeight - PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count).Shapes(1).Height - PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count).Shapes(1).Top)
            End With
            
            
            'for each column and row (each cell) change width to acommodate all columns
            For col = 1 To pptcolumns
                SummaryTblK2.Columns(col).Width = PPApp.ActivePresentation.PageSetup.SlideWidth / pptcolumns
            Next    ' column
           
            
            
        End With
    
    'Formatting slide and contents''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'center table within slide
    
    End With

            Dim Tag_Val As String
            Tag_Val = "KPI"
    
    
                                        'Create charts Slide
                                        Set SummarySlide2 = PPApp.ActivePresentation.Slides.Add(PPApp.ActivePresentation.Slides.Count + 1, ppLayoutTitleOnly)
                                        'Copy trend chart from excel
                                        
                                                    For Each sheet In ThisWorkbook.Worksheets       'delete previously made trend and running total charts
                                                    If sheet.Name = "Trend Table " & Tag_Val Then
                                                        Sheets("Trend Table " & Tag_Val).Delete
                                                    ElseIf sheet.Name = "Running " & Tag_Val Then
                                                        Sheets("Running " & Tag_Val).Delete
                                                    End If
                                                    Next sheet
                                    
                                        With SummarySlide2
                                            'paste chart in powerpoint, assign variable to table
                                            
                                            Create_Pivot_Running_Total_for_ppt start_date, Tag_Val   'macro creates running total KPI chart
                                        
                                        Worksheets("Running " & Tag_Val).Shapes("RunningPivotTable " & Tag_Val).Copy
                                        
                                        'wait action to allow VBA to copy required entries [if removed may cause error]
                                        Application.Wait (Now + TimeValue("0:00:01"))
                                        
                                            'paste chart in powerpoint, assign variable to table
                                            Set running_chart1 = .Shapes.Paste
                                            
                                            Set SummaryTitle2 = .Shapes.Title   'assign variable to title of chart
                                            
                                            Excel.Application.DisplayAlerts = False
                                            PPApp.DisplayAlerts = False
                                            
                            
                                            
                                        End With
                                                
                                        'Change title format of chart slide'''''''''''''''''''''''''''''''''''''''''''''''
                                        With SummarySlide2
                                                    With SummaryTitle2
                                                        With .TextFrame.TextRange
                                                          'Logic for naming convention of new slide
                                                            .Text = "Monthly Summary"
                                                             With .Font
                                                                    .Bold = msoFalse
                                                                    '.Name = "Times New Roman"
                                                                    .Size = 45
                                                                    '.Color = RGB(0, 0, 0)
                                                            End With
                                                    End With
                                                End With
                                            
                                            ' alter format of chart within slide'''''''''''''''''''''''''''
                                            With running_chart1
                                                With PPApp.ActivePresentation.PageSetup
                                                    'Center Horizontally
                                                        'must view slide to set table to center
                                                        PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count).Shapes(2).Width = .SlideWidth
                                                        PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count).Shapes(2).left = (.SlideWidth \ 2) - (running_chart1.Width \ 2)
                                                    'Center Vertically
                                                        PPApp.ActiveWindow.View.GotoSlide (PPApp.ActivePresentation.Slides.Count)
                                                        'PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count).Shapes(3).Select
                                                        PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count).Shapes(2).Top = PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count).Shapes(1).Top + PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count).Shapes(1).Height
                                                        PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count).Shapes(2).Height = (.SlideHeight - PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count).Shapes(1).Height - PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count).Shapes(1).Top)
                                                End With
                                                End With
                                        
                                        End With


    
    
    'set var for monthly report in excel to excel table for counting, copying, and pasting'''''''''''''''''''''''''''''''''''''''''''
    Set Report_tbl = Worksheets("Monthly Report").ListObjects(1)
    
    'count how many "open" and "closed" entries there are to account for how many slides will be in ppt
    open_count = 0
    closed_count = 0
    For Each cell In Report_tbl.ListColumns("Status").DataBodyRange
        If cell.value = "Open" Then
            open_count = open_count + 1
        Else
            closed_count = closed_count + 1
        End If
    Next
        
    'based on numbers per slide, calculating total "Open" and "Closed" Entry slides
    Open_slides = Application.WorksheetFunction.RoundUp(open_count / num_per_slides, 0)
    closed_slides = Application.WorksheetFunction.RoundUp(closed_count / num_per_slides, 0)
    
    total_slides = Open_slides + closed_slides  '=total slide sin ppt
     
''''''create slides for tables, copy, paste, and delete irrelevant or already used entries from chart into differnt worksheets''''''''''

    ' calculate total rows and columns in monthly report table
    ro = Report_tbl.DataBodyRange.Rows.Count
    cols = Report_tbl.DataBodyRange.Columns.Count
    
    'entry slide start count
    Slide = 1
    ' creation of slides, containing open and closed entries according to the desired numbers per slide
    For Slide = 1 To total_slides
        
        
        ' create a new sheet to house copyable and pastable content for every new slide
        'new sheet will copy and paste entire monthly report, and then eliminate rows that won't be able to fit on current slide
        'code then does the same thing for the next slide - eliminates rows already used or won't fit
        'shets are deleted after info is pasted into ppt
        
        Dim TempReportSheet As String
        TempReportSheet = "Sub Monthly Report for Slide  " & PPApp.ActivePresentation.Slides.Count  'each new slide gets a temporary sheet
            
        Application.DisplayAlerts = False
                                                    ' delete sheet if already exists
        For Each sheet In ThisWorkbook.Worksheets
            If sheet.Name = TempReportSheet Then
                Sheets(TempReportSheet).Delete
            End If
        Next sheet
        
        'creates new sheet for pivot table and chart
        Sheets.Add After:=ThisWorkbook.Worksheets("Control Center")
        ActiveSheet.Name = TempReportSheet
        Set mreport_sheet = ThisWorkbook.Worksheets(TempReportSheet)
        
        'turns on alerts after initial deletion/creation
        Sheets(TempReportSheet).Activate
        ActiveWindow.DisplayGridlines = False
        
        'identifies first cell of first row and last cell of last row of main monthly report table
        Dim firstcell As Range
        Dim endcell As Range
        Set firstcell = Sheets("Monthly Report").Range("A2")
        Set endcell = Sheets("Monthly Report").Range("A2").Offset(ro, cols - 1)
        
        'set monthly report as variable
        Dim report_rng As Range
        'set range that new table will talke up
        Set report_rng = ThisWorkbook.Sheets("Monthly Report").Range(firstcell.Address & ":" & endcell.Address)
        
        'copy and paste report into new sheet
        report_rng.Copy
        Application.Wait (Now + TimeValue("0:00:01"))
        Sheets(TempReportSheet).Range("A1").PasteSpecial xlValues
        Sheets(TempReportSheet).Range("A1").PasteSpecial xlFormats
        Sheets(TempReportSheet).Range("A1").PasteSpecial xlPasteColumnWidths
        
        ActiveWindow.Zoom = 25
        
        'create new excel table from copied table (makes it easier to paste within powerpoint->)''''''''''''''''''''
        'following vlock of code takes regular cells and turns them into an excel table
        MyName = "TempTable" & Slide
        MyRangeString = "A1:I1"             'to accomodate each column in main monthly table
        
        MyListExists = False                'if the table xists already, delete
        For Each ListObj In Sheets(TempReportSheet).ListObjects
            If ListObj.Name = MyName Then MyListExists = True
        Next ListObj
        
        If Not (MyListExists) Then          'if table doesn't exist, create a new one (will always be the case), the sheets and tables get deleted before the next slide is created
            Set NewTempTable = Sheets(TempReportSheet).ListObjects.Add
            NewTempTable.Name = MyName
        End If
    
        
        'Delete "open" entries if closed slide; delete "closed" entries if open slide
        Dim rowcount As Double
        rowcount = NewTempTable.Range.Rows.Count
        i = 1               'row being analyzed
        Dim C As Variant
        Do While i <= rowcount
            '"Open" entries are listed first, so "Open" slides are also displayed first
            'So if the current slide number "Slide" is less or equal to the total amount of "Open" slides sets to be displayed
            'AND if the row of the i row is a "closed" entry, then the code will delete the "closed" entry row
            '
            '
            'since rowcount gets smaller with every deletion the i = i - 2 is necessary to backtrack over rows that have
            '"moved up" post deletion, because otherwise excel will skip the "next" row, so not all rows will be analyzed
            
            
            If Slide <= Open_slides Then        'if current slide is less than or equal to number of open slides, then:
                SlideType = "Open"              'Slide type = open
                i = i + 1                       'go to next row
                If NewTempTable.ListColumns("Status").DataBodyRange.Cells(i, 1).value = "Closed" Then   'if row is closed
                    NewTempTable.ListColumns("Status").DataBodyRange.Cells(i, 1).EntireRow.Delete       'delete row
                    i = i - 2                                                                           'cycle back row
                End If
            
            'Same thing here, If current slide is greater than the number of open slides then that must mean it is a "closed" slide
            ' because it is a closed slide, the code will delete any rows with "Open" entries
            ElseIf Slide > Open_slides Then
                SlideType = "Closed"
                i = i + 1
                If NewTempTable.ListColumns("Status").DataBodyRange.Cells(i, 1).value = "Open" Then
                    NewTempTable.ListColumns("Status").DataBodyRange.Cells(i, 1).EntireRow.Delete
                    i = i - 2
                End If
            
            End If
        Loop
        
        
        'This logic adds keeps track of the current slide (5th "Open" slide, 4th "Closed" slide, etc)
        ' E.g. First "Open" Slide, Second "Closed" Slide, etc
        'This helps keep track of how many cells need to be offset to capture the correct num per slides in the table
        'as seen below
        
        If SlideType = "Open" Then
        SlideOpenNum = SlideOpenNum + 1
        ElseIf SlideType = "Closed" Then
        SlideClosedNum = SlideClosedNum + 1
        End If
        
        
        'Logic will delete items existing physically above desired entries (in temp table report) and output to slide''''''''''''''''''''''''''''''''''''''''''''''''
        '   EX: If there are 6 entries in main table, and entries 1-4 are already on a slide then
        '   entries 5 and 6 will need to be isolated to be put on their own slide
        '   code will use the current slide number, the num_per_slide (number of entries per slide as deisgnated by the useform)
        '   and the slide type to delete appropriate rows that exist above the desired rows
        '
        '
        '''''Logic to keep track of how many cells to offset for deletion
        '''''T as in "TOffsetCount" = Top entries to delete -> "Top" [1-4, if staying true to example above]
        '''''B = Bottom entries to delete -> "Bottom [7-end if staying true to example above ]
        
        ''''''''If statement asks "are there any remaining "Open" slides?
        If SlideOpenNum > 1 And SlideType = "Open" Then
        
        'Offset count (how to isolate rows that need to be deleted) is one less than the slide type count because at this point,
        'when this logic is triggered the count will be starting at 2 (previous if statement asks if >1 not >=1), the "-1" is to bring that number back to 1 and increment from there
        'because the first time this logic is used will be the FIRST time when cells to be deleted will exist ABOVE the entires we wish to display;
        'we have to bring the offset coefficient back to 1
        If SlideType = "Open" Then
        TOffsetCount = SlideOpenNum - 1
        ElseIf SlideType = "Closed" Then
        TOffsetCount = SlideClosedNum - 1
        End If
        
        'this is the row RIGHT BEFORE the group of rows we want to display. If the rows we want to display are 4, 5, and 6,
        ' this row will yield the first cell in row 3
        Dim topdelete_row As Range
        Set topdelete_row = Sheets(TempReportSheet).Range("A1").Offset(num_per_slides * TOffsetCount, 0)
        
        'The first row in the group of cells we want to delete will always be the second row, or the row of cell (2, 1)
        top_row = Cells(2, 1).Address

        'the delete range is the rows in between the top row (row 2) and the last row before the rows we want to display
        Set topdelete_rng = Range(top_row & ":" & topdelete_row.Address)
        End If
        
        
        'same logic and reasoning as above.. ">1" not ">=1"
        If SlideClosedNum > 1 And SlideType = "Closed" Then
        
        'offset counts to get to the desired row
        If SlideType = "Open" Then
        TOffsetCount = SlideOpenNum - 1
        ElseIf SlideType = "Closed" Then
        TOffsetCount = SlideClosedNum - 1
        End If
        
        'this is the row RIGHT BEFORE the group of rows we want to display. If the rows we want to display are 4, 5, and 6, this row will yield the first cell in row 3
        Set topdelete_row = Sheets(TempReportSheet).Range("A1").Offset(num_per_slides * TOffsetCount, 0)
        
        'The first row in the group of cells we want to delete will always be the second row, or the row of cell (2, 1)
        top_row = Cells(2, 1).Address

        'the delete range is the rows in between the top row (row 2) and the last row before the rows we want to display
        Set topdelete_rng = Range(top_row & ":" & topdelete_row.Address)
        End If
        
        Dim OffsetCount As Double
        
        If SlideType = "Open" Then
        BOffsetCount = SlideOpenNum
        ElseIf SlideType = "Closed" Then
        BOffsetCount = SlideClosedNum
        End If
        
        'procedure to delete the rows of the table located after the rows we wish to display''''''''''''''''''''''''''''''''
        
        'this is the row RIGHT BEFORE the group of rows we want to display. If the rows we want to display are 4, 5, and 6, this row will yield the first cell in row 7
        Dim bottomdelete_row As Range
        Set bottomdelete_row = Sheets(TempReportSheet).Range("A1").Offset(num_per_slides * BOffsetCount + 1, 0)
        
        'The last row in the group of cells
        bottom_row = Sheets(TempReportSheet).Cells(Rows.Count, 1).End(xlUp).row + 1
        
        'the delete range is the rows in between the row after the rows we wish to display and the last row of the tablle
        Dim bottomdelete_rng As Range
        Set bottomdelete_rng = Range(bottomdelete_row.Address & ":" & "A" & bottom_row)
        
        'If first entry slide in deck, only bottom range needs to be deleted (no unwanted entries will exist above the ones we want)
        If SlideOpenNum <= 1 And SlideType = "Open" Then
            bottomdelete_rng.EntireRow.Delete
            ElseIf SlideClosedNum <= 1 And SlideType = "Closed" Then
            bottomdelete_rng.EntireRow.Delete
            ElseIf SlideOpenNum > 1 And SlideType = "Open" Then
            bottomdelete_rng.EntireRow.Delete
            topdelete_rng.EntireRow.Delete
            ElseIf SlideClosedNum > 1 And SlideType = "Closed" Then
            bottomdelete_rng.EntireRow.Delete
            topdelete_rng.EntireRow.Delete
            
        End If
        
        
        'Create a new slide for the entries'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set PPSlide = PPApp.ActivePresentation.Slides.Add(PPApp.ActivePresentation.Slides.Count + 1, ppLayoutTitleOnly)
    
        'Copy the chart and paste in Powerpoint
        NewTempTable.Range.Copy
        'wait action to allow VBA to copy required entries [if removed may cause error]
        Application.Wait (Now + TimeValue("0:00:01"))
         
        'Paste title on slide, paste table (entries) on slide
        With PPSlide
            Set ReportTable = .Shapes.Paste
            Set ReportTable = .Shapes(2).Table
        End With
        
        'Delete temporary sheet holding temporary table
        Sheets(TempReportSheet).Delete
        
        'Change title format of reporting slide'''''''''''''''''''''''''''''
    With PPSlide
            Set ReportSlideTitle = .Shapes.Title
                With ReportSlideTitle
                    With .TextFrame.TextRange
                      'Logic for naming convention of new slide
                        If SlideType = "Open" Then
                        .Text = "Opened Issues"
                        ElseIf SlideType = "Closed" Then
                        .Text = "Closed Issues"
                        End If
                         With .Font
                                .Bold = msoFalse
                                '.Name = "Times New Roman"
                                .Size = 45
                                '.Color = RGB(0, 0, 0)
                        End With
                End With
            End With
    End With
    
    pptcolumns = ReportTable.Columns.Count
    pptrows = ReportTable.Rows.Count
    
    With ReportTable
            'for each column and row (each cell)
            For col = 1 To pptcolumns
                For roh = 1 To pptrows
                    If col <= pptcolumns And roh <= pptrows Then
                        With .cell(roh, col).Shape
                            'fix column width
                            If .HasTextFrame Then
                                If .TextFrame.HasText Then
                                        .TextFrame.TextRange.Font.Size = FontSize
                                End If
                            End If
                        End With
                    With .cell(pptrows, pptcolumns).Shape.TextFrame
                    If minW = 0 Then minW = .TextRange.BoundWidth + .MarginLeft + .MarginRight + 1
                    If minW < .TextRange.BoundWidth + .MarginLeft + .MarginRight + 1 Then _
                    minW = .TextRange.BoundWidth + .MarginLeft + .MarginRight + 1
                    End With
                    Else
                    End If
                Next
                .Columns(pptcolumns).Width = minW
            Next    ' column
            ' row
    End With
            
        ' alter format of table within slide'''''''''''''''''''''''''''
    
    'Formatting slide and contents''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'center table within slide
    Set Obj = PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count).Shapes(2)
            
            With PPApp.ActivePresentation.PageSetup
    
    'Center Horizontally
        'must view slide to set table to center
        PPApp.ActiveWindow.View.GotoSlide (PPApp.ActivePresentation.Slides.Count)
        'Center Horizontally
            'must view slide to set table to center
            Obj.Width = .SlideWidth
            Obj.left = (.SlideWidth \ 2) - (Obj.Width \ 2)
        'Center Vertically
            PPApp.ActiveWindow.View.GotoSlide (PPApp.ActivePresentation.Slides.Count)
            Obj.Select
            Obj.Top = PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count).Shapes(1).Top + PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count).Shapes(1).Height
            Obj.Height = (.SlideHeight - PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count).Shapes(1).Height - PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count).Shapes(1).Top)
            End With
        'set background color
        With PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count)
        '.FollowMasterBackground = msoFalse
        '.Background.Fill.Solid
        '.Background.Fill.ForeColor.RGB = RGB(222, 220, 216)
        End With
        
        pptcolumns = ReportTable.Columns.Count
        pptrows = ReportTable.Rows.Count
        
        With ReportTable
            For col = 1 To pptcolumns
                For roh = 1 To pptrows
                If col <= pptcolumns And roh <= pptrows Then
                        With .cell(roh, col).Shape
                            'fix column width
                            If .HasTextFrame Then
                                If .TextFrame.HasText Then
                                        If .TextFrame.TextRange.Text = "Category" Then
                                            ReportTable.Columns(col).Width = ReportTable.Columns(col).Width + (0.185 * 72)
                                            ReportTable.Columns(col + 1).Width = ReportTable.Columns(col + 1).Width - (0.185 * 72)
                                        ElseIf .TextFrame.TextRange.Text = "KPI" Then
                                            ReportTable.Columns(col).Width = ReportTable.Columns(col).Width + ((0.5 + 0.185) * 72)
                                            ReportTable.Columns(col + 1).Width = ReportTable.Columns(col + 1).Width - (0.5 * 72)
                                            ReportTable.Columns(col + 2).Width = ReportTable.Columns(col + 2).Width - (0.185 * 72)
                                        ElseIf .TextFrame.TextRange.Text = "Owner" Then
                                            ReportTable.Columns(col).Width = ReportTable.Columns(col).Width + (0.9 * 72)
                                            ReportTable.Columns(col - 1).Width = ReportTable.Columns(col - 1).Width - (0.9 * 72)
                                        End If
                                End If
                            End If
                        End With '13.33
                Else
                End If
                Next
            Next
            
            PPApp.ActiveWindow.View.GotoSlide (PPApp.ActivePresentation.Slides.Count)
            Obj.Select
            Obj.Top = PPApp.ActivePresentation.PageSetup.SlideHeight - PPApp.ActivePresentation.Slides(PPApp.ActivePresentation.Slides.Count).Shapes(2).Height
        End With
    Next
    
    MsgBox "Please make sure the powerpoint is formatted correctly prior to presenting, not all inputs will yield a perfect format."
   

     Excel.Application.DisplayAlerts = False
    PPApp.DisplayAlerts = False
        
End Sub





Sub Create_Pivot_Trend_for_ppt(start_date As Date, Tag_Val As String)

Dim trend_sheet As Worksheet 'variable for worksheet for general trends
Dim counter_sheet As Worksheet 'variable for worksheet with trend data
Dim trend_cache As PivotCache 'variable for trend data as a pivot cache
Dim trend_table As PivotTable 'variable for trend pivot table
Dim counter_range As Range 'variable for the range of the data of the trned pivot table being the countermeasures trend data
Dim lastrow As Long 'last row of countermeasures trend data
Dim lastcol As Long 'last col of countermeasures trend data
Dim counter_tbl As ListObject 'excel table of countermeasure data found within name manager
Dim cat_val As String 'variable for category value when choosing from drop down
Dim filterval As Integer 'variable for filter value when choosing from drop down
Dim pivot_item As PivotItem 'variable for assessing if pivotitems = blank
Dim issue_field As PivotField ' variable for assessing pivot field from drop down box
Dim sheet_names As New Collection

'start_date = "September 2020"

'On Error GoTo ErrorHandler 'if error, displays error message

    'turns off "are you sure you want to delete (previous worksheet and chart??"
    'turns off excel blink
Application.DisplayAlerts = False
Application.ScreenUpdating = False

'deletes previous pivot table and chart, if it exists
For Each sheet In ThisWorkbook.Worksheets
    If sheet.Name = "Trend Table " & Tag_Val Then
        Sheets("Trend Table " & Tag_Val).Delete
    End If
Next sheet
    
    'creates new sheet for pivot table and chart
Sheets.Add After:=ThisWorkbook.Worksheets("Control Center")
ActiveSheet.Name = "Trend Table " & Tag_Val

sheet_names.Add item:="Trend Table " & Tag_Val, key:=Tag_Val

    'turns on alerts after initial deletion/creation
Application.DisplayAlerts = True

    'sets variables and names to worksheets
    'sets variable and name to countermeasures table
Set trend_sheet = Worksheets("Trend Table " & Tag_Val)
Set counter_sheet = Worksheets("Countermeasures")
Set counter_tbl = counter_sheet.ListObjects("Tbl_Counter")


    'Sets pivot cache to countermeasures table
Set trend_cache = ThisWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:="Tbl_Counter")

    'creates the blank pivot table in the top left cell of spreadsheet
Set trend_table = trend_cache.CreatePivotTable _
(TableDestination:=Worksheets("Trend Table " & Tag_Val).Cells(1, 1), _
TableName:="TrendPivotTable " & Tag_Val)

    'sets variables to value within drop down boxes on "Create Pivot Table" spreadsheet
'cat_val = Sheets("Control Center").Pivot_DD_Box.value
'issue_val = Sheets("Control Center").Issue_DD_Box.value
filterval = 1



    'Inputs pivot table fields into pivot table,
    'Logic to determine which category fields should be shown or not
   

With trend_table.PivotFields("Issue Year")
    .Orientation = xlRowField
    .Position = 1
    End With
    
     For Each pivot_item In trend_table.PivotFields("Issue Year").PivotItems
            If pivot_item.Name <> Year(start_date) Then
                pivot_item.Visible = False
                Else
                pivot_item.Visible = True
            End If
        Next

With trend_table.PivotFields("Issue Month")
    .Orientation = xlRowField
    .Position = 2
    End With
    
    For Each pivot_item In trend_table.PivotFields("Issue Month").PivotItems
            If pivot_item.Name <> Month(start_date) Then
                pivot_item.Visible = False
                Else
                pivot_item.Visible = True
            End If
    Next
    
With trend_table.PivotFields("Issue Date")
    .Orientation = xlDataField
    .Position = 1
    .Function = xlCount
    .Name = "Count of Issues"
    End With

Set issue_field = trend_table.PivotFields(Tag_Val)
    
With issue_field
  .Orientation = xlRowField
   .Position = 3
   .PivotItems("(blank)").Visible = False
   End With
   
   
    For Each pivot_item In issue_field.PivotItems
            If pivot_item.Name = "(blank)" Then
                pivot_item.Visible = False
                Else
                pivot_item.Visible = True
            End If
    Next
    
    
    'Adds value filter, found on "Create Pivot Table" spreadsheet, to look for issues with frequency greater than or equal to drop down box value
With ThisWorkbook.Worksheets("Trend Table " & Tag_Val).PivotTables("TrendPivotTable " & Tag_Val).PivotFields(Tag_Val)
            .PivotFilters.Add2 Type:=xlValueIsGreaterThanOrEqualTo, DataField:=ThisWorkbook.Worksheets("Trend Table " & Tag_Val).PivotTables("TrendPivotTable " & Tag_Val).PivotFields("Count of Issues"), _
            Value1:=filterval
    End With
   
''''''''PIVOT CHART CREATION'''''''''''''''''''''''''''''
Dim Trend_chart As ChartObject 'shape housing the chart
Dim trend_range As Range 'range of the pivot table, and therefore data source of pivot chart
Dim PT_First_Cell As Range 'first cell of the pivot table data ~TableRange1~
Dim TopRow As Long 'first row of pivot table ~TableRange1~
Dim TopCol As Long 'first column of pivot table ~TableRange1~
Dim lastrow2 As Long 'last row of pivot table
Dim lastcol2 As Long 'last column of pivot table
Dim Chart_Add As Range 'cell where chart will be anchored


    'find range of pivot table to use for pivot chart
lastrow2 = trend_table.TableRange1.Rows.Count 'last row of pivot table
lastcol2 = trend_table.TableRange1.Columns.Count 'last column of pivot table
Set trend_range = Worksheets("Trend Table " & Tag_Val).PivotTables("TrendPivotTable " & Tag_Val).TableRange1.Cells(1, 1).Resize(lastrow2, lastcol2) 'table 1 range (data range) of pivot table
TopRow = Worksheets("Trend Table " & Tag_Val).PivotTables("TrendPivotTable " & Tag_Val).TableRange1.row 'top row of pivot table (data range)
TopCol = Worksheets("Trend Table " & Tag_Val).PivotTables("TrendPivotTable " & Tag_Val).TableRange1.Column 'top column of pivot table (table range)

Set PT_First_Cell = Cells(TopRow, TopCol) 'first cell of pivot table range (data range)
Set Chart_Add = PT_First_Cell.Offset(rowOffset:=trend_table.TableRange1.Rows.Count, ColumnOffset:=0) 'takes first cell and adds # of rows in pivot table for first cell of chart placement

With Worksheets("Trend Table " & Tag_Val).Shapes.AddChart2(297, xlColumnStacked)
    
    .Name = "TrendPivotChart " & Tag_Val
    .Chart.SetSourceData Source:=trend_range
    .Chart.HasTitle = True
    .Chart.ChartTitle.Text = "Monthly Trends"
    .Chart.Axes(xlValue, xlPrimary).HasTitle = True
    .Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Frequency"
    .Chart.Axes(xlValue, xlPrimary).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 12
    '.Chart.Axes(xlCategory, xlPrimary).HasTitle = True
    '.Chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time"
    '.Chart.Axes(xlCategory, xlPrimary).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 12
    .left = Chart_Add.left
    .Top = Chart_Add.Top
    .Height = 250
    .Width = 1000
    .Chart.Legend.Position = xlLegendPositionCorner
    .Chart.Legend.Top = 7
    .Chart.Legend.Height = 200
End With


''''''''''''On Error:'''''''''''''''''''''''''''''
'Exit Sub

'ErrorHandler:

'MsgBox ("Please select value in both drop down boxes")

'Resume Next

End Sub
Sub Create_Pivot_Running_Total_for_ppt(start_date As Date, Tag_Val As String)

Dim running_sheet As Worksheet 'variable for worksheet for general trends
Dim counter_sheet As Worksheet 'variable for worksheet with trend data
Dim running_cache As PivotCache 'variable for trend data as a pivot cache
Dim running_table As PivotTable 'variable for trend pivot table
Dim counter_range As Range 'variable for the range of the data of the trned pivot table being the countermeasures trend data
Dim lastrow As Long 'last row of countermeasures trend data
Dim lastcol As Long 'last col of countermeasures trend data
Dim counter_tbl As ListObject 'excel table of countermeasure data found within name manager
Dim cat_val As String 'variable for category value when choosing from drop down
Dim filterval As Integer 'variable for filter value when choosing from drop down
Dim pivot_item As PivotItem 'variable for assessing if pivotitems = blank
Dim issue_field As PivotField ' variable for assessing pivot field from drop down box

'On Error GoTo ErrorHandler 'if error, displays error message
'start_date = "September 2020"
'Tag_Val = "Issue Tier 1 Tag"

    'turns off "are you sure you want to delete (previous worksheet and chart??"
    'turns off excel blink
Application.DisplayAlerts = False
Application.ScreenUpdating = False


    'deletes previous pivot table and chart, if it exists
For Each sheet In ThisWorkbook.Worksheets
    If sheet.Name = "Running " & Tag_Val Then
        Sheets("Running " & Tag_Val).Delete
    End If
Next sheet
    
    'creates new sheet for pivot table and chart
Sheets.Add After:=ThisWorkbook.Worksheets("Control Center")
ActiveSheet.Name = "Running " & Tag_Val

    'turns on alerts after initial deletion/creation
Application.DisplayAlerts = True

    'sets variables and names to worksheets
    'sets variable and name to countermeasures table
Set running_sheet = Worksheets("Running " & Tag_Val)
Set counter_sheet = Worksheets("Countermeasures")
Set counter_tbl = counter_sheet.ListObjects("Tbl_Counter")


    'Sets pivot cache to countermeasures table
Set running_cache = ThisWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:="Tbl_Counter")

    'creates the blank pivot table in the top left cell of spreadsheet
Set running_table = running_cache.CreatePivotTable _
(TableDestination:=Sheets("Running " & Tag_Val).Cells(1, 1), _
TableName:="RunningPivotTable " & Tag_Val)

    'sets variables to value within drop down boxes on "Create Pivot Table" spreadsheet
'cat_val = Sheets("Control Center").Pivot_DD_Box.value
'issue_val = Sheets("Control Center").Issue_DD_Box.value
filterval = 1

    'Inputs pivot table fields into pivot table,
    'Logic to determine which category fields should be shown or not

'running_table.PivotFields.ClearAllFilters

With running_table.PivotFields("Month Name")
    .Orientation = xlRowField
    .Position = 1
    End With
             For Each pivot_item In running_table.PivotFields("Month Name").PivotItems
            If pivot_item.Name <> MonthName(Month(start_date)) Then
                pivot_item.Visible = False
                Else
                pivot_item.Visible = True
            End If
    Next
    
With running_table.PivotFields("Day of Month")
    .Orientation = xlRowField
    .Position = 2
    End With
    
    'For Each pivot_item In running_table.PivotFields("Day of Month").PivotItems
            'If pivot_item.Name = "(blank)" Then
              '  pivot_item.Visible = False
              '  Else
              '  pivot_item.Visible = True
           ' End If
    'Next
    
     Set issue_field = running_table.PivotFields(Tag_Val)
   
With issue_field
    .Orientation = xlColumnField
   .Position = 1
   .PivotItems("(blank)").Visible = False
   Subtotals = Array(False, True, True, False, False, False, False, False, False, False, _
        False, False)
   End With
   
    For Each pivot_item In issue_field.PivotItems
            If pivot_item.Name = "(blank)" Then
                pivot_item.Visible = False
                Else
                pivot_item.Visible = True
            End If
    Next
     
With running_table.PivotFields("Issue Date")
    .Orientation = xlDataField
    .Position = 1
    .Function = xlCount
    .Name = "Count of Issues"
    .Calculation = xlRunningTotal
    .BaseField = "Day of Month"
    End With
    
 'For Each pivot_item In running_table.PivotFields("Issue Date").PivotItems
            'If pivot_item.Name = "(blank)" Then
                'pivot_item.Visible = False
                'Else
                'pivot_item.Visible = True
            'End If
    'Next
    
    
    'Adds value filter, found on "Create Pivot Table" spreadsheet, to look for issues with frequency greater than or equal to drop down box value
With ThisWorkbook.Worksheets("Running " & Tag_Val).PivotTables("RunningPivotTable " & Tag_Val).PivotFields(Tag_Val)
            .PivotFilters.Add2 Type:=xlValueIsGreaterThanOrEqualTo, DataField:=ThisWorkbook.Worksheets("Running " & Tag_Val).PivotTables("RunningPivotTable " & Tag_Val).PivotFields("Count of Issues"), _
            Value1:=filterval
    End With
   
''''''''PIVOT CHART CREATION'''''''''''''''''''''''''''''
Dim running_chart As ChartObject 'shape housing the chart
Dim running_range As Range 'range of the pivot table, and therefore data source of pivot chart
Dim PT_First_Cell As Range 'first cell of the pivot table data ~TableRange1~
Dim TopRow As Long 'first row of pivot table ~TableRange1~
Dim TopCol As Long 'first column of pivot table ~TableRange1~
Dim lastrow2 As Long 'last row of pivot table
Dim lastcol2 As Long 'last column of pivot table
Dim Chart_Add As Range 'cell where chart will be anchored


    'find range of pivot table to use for pivot chart
lastrow2 = running_table.TableRange1.Rows.Count 'last row of pivot table
lastcol2 = running_table.TableRange1.Columns.Count 'last column of pivot table
Set running_range = Worksheets("Running " & Tag_Val).PivotTables("RunningPivotTable " & Tag_Val).TableRange1.Cells(1, 1).Resize(lastrow2, lastcol2) 'table 1 range (data range) of pivot table
TopRow = Worksheets("Running " & Tag_Val).PivotTables("RunningPivotTable " & Tag_Val).TableRange1.row 'top row of pivot table (data range)
TopCol = Worksheets("Running " & Tag_Val).PivotTables("RunningPivotTable " & Tag_Val).TableRange1.Column 'top column of pivot table (table range)

Set PT_First_Cell = Cells(TopRow, TopCol) 'first cell of pivot table range (data range)
Set Chart_Add = PT_First_Cell.Offset(rowOffset:=running_table.TableRange1.Rows.Count, ColumnOffset:=0) 'takes first cell and adds # of rows in pivot table for first cell of chart placement



running_range.Select

With Worksheets("Running " & Tag_Val).Shapes.AddChart2(227, xlLine)
    .Name = "RunningPivotTable " & Tag_Val
    .Chart.SetSourceData Source:=running_range
    .Chart.HasTitle = True
    .Chart.ChartTitle.Text = "Running Total for " & Tag_Val
    .Chart.Axes(xlValue, xlPrimary).HasTitle = True
    .Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Frequency"
    .Chart.Axes(xlValue, xlPrimary).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 12
    .Chart.Axes(xlCategory, xlPrimary).HasTitle = True
    .Chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Day of the Month"
    .Chart.Axes(xlCategory, xlPrimary).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 12
    .left = Chart_Add.left
    .Top = Chart_Add.Top
    .Height = 250
    .Width = 1000
    .Chart.Legend.Position = xlLegendPositionCorner
    .Chart.Legend.Top = 7
    .Chart.Legend.Height = 200
    .Chart.Legend.Width = 70
End With


Worksheets("Monthly Report").Activate

''''''''''''On Error:'''''''''''''''''''''''''''''
Exit Sub

'ErrorHandler:

'MsgBox ("Please select value in both drop down boxes")

'Resume Next

End Sub



