Attribute VB_Name = "CreateCharts"
Sub CreateKPIChart(cat_val As String, year_val As Long)

Dim SheetName As String
        SheetName = "KPI Chart"
        
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
    
    
    'setting countermeasures table as source for monthly report data
Dim counter_tbl As ListObject
   Set counter_sheet = Worksheets("Countermeasures")
   Set counter_tbl = counter_sheet.ListObjects("Tbl_Counter")
   
   Set KPI_rng = counter_tbl.ListColumns("KPI").DataBodyRange
Dim KPI_Cal_Rng As Range
 
    Dim ArrBase() As Variant
        row_count = 0
            For Each Cell3 In counter_tbl.ListColumns("Category").DataBodyRange
            row_count = 1 + row_count
                If Cell3 = cat_val Then
                                        
                    KPICell = counter_tbl.ListColumns("KPI").DataBodyRange(row_count, 1).value
                    
                    'if first entry, redim to hold one spot "(0)"
                    If entry_count = 0 Then
                        ReDim Preserve ArrBase(0)
                        ArrBase(0) = KPICell
                     'For all subsequent entries extend array by 1 and enter contents in cell
                    Else
                        ReDim Preserve ArrBase(UBound(ArrBase) + 1)
                        ArrBase(UBound(ArrBase)) = KPICell
                    End If
                    entry_count = entry_count + 1
                End If
             Next Cell3
             
            If (Not Not ArrBase) = 0 Then       'if array never intitialized then it doesn't exist,
                     IsArrayEmpty = True         'if list doesn't exist, then it's empty (by default)
                 Else
                     IsArrayEmpty = True         'if list does exist, assume it is empty
                     For Each item_in_array In ArrBase       'test if array is empty. If there is one non-blank cell, change bool value
                         If item_in_array <> Empty Then
                             IsArrayEmpty = False
                         End If
                     Next
             End If
                     
             If IsArrayEmpty = False Then            'if array isn't empty, remove blanks, remove dups
                     ArrBase = BlankRemover(ArrBase)     'if not empty, remove blanks
                     ArrBase = ArrayRemoveDups(ArrBase)  'if not empty, remove duplicates
                 ElseIf IsArrayEmpty = True Then     'if array is empty,  resize to one entry and create one entry "No List Available"
                     ReDim Preserve ArrBase(0)
                     ArrBase(0) = "No List Available"
            End If
                
                'MsgBox Join(ArrBase, vbCrLf)
                
    Dim ArrDict As Scripting.Dictionary
    Set ArrDict = New Scripting.Dictionary
    
    'rids duplicates and counts new dictionary
    Set ArrDict = DuplicateCountToScript(ArrBase)
    
    'Call Functions_M.SortDictionary(ArrDict, False, True, vbBinaryCompare)
        'place KPI text at top of sheet starting from col D and moving right
    OffsetCount = 0
    For Each item In ArrDict.Keys
        Sheets(SheetName).Range("D1").Offset(0, OffsetCount).value = item
        OffsetCount = OffsetCount + 1
    Next
    
    Range("A1").value = year_val
    Range("B1").value = cat_val
    Range("C1").value = "KPIs: "
    
    With Range("A1:C1")
        .HorizontalAlignment = xlCenter
            Range("C1").HorizontalAlignment = xlLeft
        .WrapText = True
    End With
    
    
    Dim ColorColl As New Collection

    'Format unique KPIs in the first row (or the KPI range)
    For i = 0 To ArrDict.Count - 1
        'red = 200 - (25 * (i))
        'If red <= 200 Then
            Red = GenerateRandomInt(255, 50)
            green = GenerateRandomInt(255, 50)
            Blue = GenerateRandomInt(255, 50)
            'Else
            'green = 0
        'End If
        With Range("D1").Offset(0, i)
           .VerticalAlignment = xlCenter
           .Font.Size = 9
           .Font.Bold = False
           .RowHeight = 25
           .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThick
            .Borders.Color = RGB(255, 0, 0)
            .WrapText = True
            .Interior.Color = RGB(Red, green, Blue)
           End With
        
        ColorColl.Add Range("D1").Offset(0, i).Interior.Color, Range("D1").Offset(0, i).value
    Next
    
    
    ActiveWindow.FreezePanes = False
    'Freeze top row so KPIs are visible
    Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    
    
            'create tally table!!
            MyName = "TallyTable"     'name for monthly table
            
            ExtraColumns = 3 'Number of extra columns after KPIs ("Empty, Total, Running Total")
            
           Set Top_Left = Range("C3")    'set top left of table
           Set Bottom_Right = Top_Left.Offset(13, OffsetCount + ExtraColumns)     'set bottom right = 12 down, # of different KPIs across
        
            MyRangeString = Top_Left.Address & ":" & Bottom_Right.Address   'make table address string
        
            MyListExists = False                                'assume table does not exist
            For Each ListObj In Sheets(SheetName).ListObjects
                If ListObj.Name = MyName Then MyListExists = True   'if table exists, true
            Next ListObj
            
            If Not (MyListExists) Then      'if table does not exist, create
                Set Tally_Tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Range(MyRangeString), , xlYes)
                Tally_Tbl.Name = MyName
                Else                        'if exists, designate variable
                Set Tally_Tbl = ActiveSheet.ListObjects(MyName)
            End If
            
            
            Tally_Tbl.HeaderRowRange(1, 1).value = "Month"   'put table name in table
            
            key = 0
            For col = 2 To Tally_Tbl.Range.Columns.Count - ExtraColumns    'label columns in monthly table
                Tally_Tbl.HeaderRowRange(1, col).value = ArrDict.Keys(key)
                key = key + 1
            Next col
            
            Tally_Tbl.HeaderRowRange(1, Tally_Tbl.Range.Columns.Count - 2).value = "[Empty]"
            Tally_Tbl.HeaderRowRange(1, Tally_Tbl.Range.Columns.Count - 1).value = "Total" 'create total column
            Tally_Tbl.HeaderRowRange(1, Tally_Tbl.Range.Columns.Count).value = "Running Total" 'create running total column
            
            For row = 1 To Tally_Tbl.ListColumns(1).DataBodyRange.Rows.Count - 1      'label frequency columns
                Count = Count + 1
                Tally_Tbl.ListColumns(1).DataBodyRange(Count, 1) = MonthName(Count)
            Next row
            
            Tally_Tbl.HeaderRowRange(Tally_Tbl.ListColumns(1).Range.Rows.Count, 1).value = "Total"  'label total KPI row
            
            With Tally_Tbl.Range
                .BorderAround xlContinuous, xlMedium, xlColorIndexAutomatic
            End With
            Count = 0
            
            col_count = 0
            For Each col In Tally_Tbl.HeaderRowRange
                col_count = col_count + 1
                For Each KPI In ArrDict.Keys
                    If col = KPI Then
                        Tally_Tbl.ListColumns(col_count).Range(1).Interior.Color = ColorColl(KPI)
                    End If
                Next
                
                If col = "[Empty]" Then
                    Tally_Tbl.ListColumns("[Empty]").Range(1).Interior.Color = RGB(234, 246, 148)
                End If
                
            Next
    
    Dim Mnth As Integer
    Dim DateVal As Date
    
    For Mnth = 1 To 12
            
            
            'Create monthly table!!
            ' if excel table already exists, delete and replace with new one
            MyName = MonthName(Mnth) & "_Table"     'name for monthly table
            
                DateVal = DateValue(MonthName(Mnth) & " 1 " & Str(year_val))  'date value for first day of month
            
               Set Top_Left = Range("C20").Offset((Mnth - 1) * 11, 0)    'set top leftof table, offsetting by 11 for months > 1
               Set Bottom_Right = Top_Left.Offset(10, MonthDays(DateVal))
            
            MyRangeString = Top_Left.Address & ":" & Bottom_Right.Address   'make table address string
        
            MyListExists = False                                'assume table does not exist
            For Each ListObj In Sheets(SheetName).ListObjects
                If ListObj.Name = MyName Then MyListExists = True   'if table exists, true
            Next ListObj
            
            If Not (MyListExists) Then      'if table does not exist, create
                Set NewTable = ActiveSheet.ListObjects.Add(xlSrcRange, Range(MyRangeString), , xlYes)
                NewTable.Name = MyName
                Else                        'if exists, designate variable
                Set NewTable = ActiveSheet.ListObjects(MyName)
            End If
            
            NewTable.HeaderRowRange(1, 1).value = MonthName(Mnth)   'put month name in table
            
            For col = 2 To NewTable.Range.Columns.Count + 1     'label columns in monthly table
                NewTable.HeaderRowRange(1, col).value = col - 1
            Next col
            
            For Each cell In NewTable.ListColumns(1).DataBodyRange      'label frequency columns
                Count = Count + 1
                cell.value = Count
            Next cell
            Count = 0
                
                KDay = 0        'initialize NewTable column count at 0
                For Each Column In NewTable.ListColumns                     'for each column in month table
                    
                    If IsNumeric(Column) = False Then
                                'do nothing    'if not numeric, e.g. "January", skip
                                    Else
                                CRow = 0   'row counter for countermeasures table
                                KRow = 0    'row counterfor NewTable (monthly KPI)
                                KMonth = Mnth   'KPI month = month of current table
                                KYear = year_val    'KPI year = year val of macro
                                KDay = KDay + 1      'KPI day = column header
                                
                                
                                    For Each cell In counter_tbl.ListColumns("Issue Date").DataBodyRange    'for each cell in countermeasures table "Issue Date" column
                                        
                                        CRow = CRow + 1     'Row count for countermeasures table
                                        
                                        CMonth = Month(cell.value)  'month of issue date
                                        CDay = Day(cell.value)      'day of issue date
                                        CYear = Year(cell.value)    'year of issue date
                                        
                                            'if issue date Month = NewTable Month, and issue day = column and isue year = year of chart
                                            ' and category or issue date = category value then
                                            'add KPI row + 1
                                            'add KPI value from countermeasures table to NewTable
                                        
                                        If CMonth = KMonth And CDay = KDay And CYear = KYear And counter_tbl.ListColumns("Category").DataBodyRange(CRow, 1).value = cat_val Then
                                                KRow = KRow + 1     'add NewTable row count
                                                
                                                If counter_tbl.ListColumns("KPI").DataBodyRange(CRow, 1).value <> Empty Then
                                                        
                                                        With NewTable.ListColumns(KDay + 1).DataBodyRange(KRow, 1)
                                                                .value = counter_tbl.ListColumns("KPI").DataBodyRange(CRow, 1).value     'add KPI value from countermeasures table to NewTable(KDay + 1 because in each month table there is one extra column for month name)
                                                                .Interior.Color = ColorColl(counter_tbl.ListColumns("KPI").DataBodyRange(CRow, 1).value)
                                                        End With
                                                        
                                                        Old_Tally = Tally_Tbl.ListColumns(counter_tbl.ListColumns("KPI").DataBodyRange(CRow, 1).value).DataBodyRange(Mnth, 1).value     'old tally = old total in tally table data body range
                                                        New_Tally = Old_Tally + 1                                                                                                       'add one
                                                        Tally_Tbl.ListColumns(counter_tbl.ListColumns("KPI").DataBodyRange(CRow, 1).value).DataBodyRange(Mnth, 1).value = New_Tally     'replace old tally with new tally in tally table
                                                        
                                                    Else
                                                        With NewTable.ListColumns(KDay + 1).DataBodyRange(KRow, 1)
                                                            .value = "[Empty]"     'add KPI value from countermeasures table to NewTable(KDay + 1 because in each month table there is one extra column for month name)
                                                            .Interior.Color = RGB(234, 246, 148)
                                                        End With
                                                        
                                                        Old_Tally = Tally_Tbl.ListColumns("[Empty]").DataBodyRange(Mnth, 1).value     'old tally = old total in tally table data body range
                                                        New_Tally = Old_Tally + 1                                                                                                       'add one
                                                        Tally_Tbl.ListColumns("[Empty]").DataBodyRange(Mnth, 1).value = New_Tally     'replace old tally with new tally in tally table
                                                        
                                                End If
                                        End If
                                    Next        'go to next cell in countermeasures table
                    End If
                    Next Column                 'go to next column (day) in month (newtable)
                    
                    NewTable.Range.WrapText = True
                
    Next Mnth                                   'go to next month in Calendar
    
        'add totals in TallyTable for each row ("Totals" column)
    Sum = 0
    For row = 1 To Tally_Tbl.DataBodyRange.Rows.Count - 1
        Sum = 0
        
        For col = 2 To Tally_Tbl.Range.Columns.Count - 2
            Sum = Sum + Tally_Tbl.DataBodyRange(row, col).value
        Next col
        
        Tally_Tbl.ListColumns("Total").DataBodyRange(row, 1).value = Sum
        
    Next row

        'add totals in TallyTable for each column ("Totals" row)
     Sum = 0
    For col = 2 To Tally_Tbl.Range.Columns.Count - 1
        Sum = 0
        
        For row = 1 To Tally_Tbl.DataBodyRange.Rows.Count - 1
            Sum = Sum + Tally_Tbl.DataBodyRange(row, col).value
        Next row
        
        Tally_Tbl.ListColumns(col).DataBodyRange(Tally_Tbl.DataBodyRange.Rows.Count, 1).value = Sum
        
    Next col
    
        'add running totals in TallyTable for last column ("Running Totals" col)
     RunningTotal = 0
    For row = 1 To Tally_Tbl.DataBodyRange.Rows.Count - 1
        
        If row = 1 Then
            RunningTotal = Tally_Tbl.ListColumns("Total").DataBodyRange(row, 1).value
            Else
            RunningTotal = RunningTotal + Tally_Tbl.ListColumns("Total").DataBodyRange(row, 1).value
        End If
        
        Tally_Tbl.ListColumns("Running Total").DataBodyRange(row, 1).value = RunningTotal
        
    Next row
    
        'format tally table
    Tally_Tbl.TableStyle = "TableStyleLight1"
    Tally_Tbl.DataBodyRange.HorizontalAlignment = xlCenter
    Tally_Tbl.ListColumns(1).DataBodyRange.HorizontalAlignment = xlLeft
    Tally_Tbl.Range.WrapText = True
    
    
        'create bar chart
    
    Set Tally_Top_Right = Tally_Tbl.Range(1, Tally_Tbl.Range.Columns.Count)
    
    Set BarTopLeft = Tally_Top_Right.Offset(0, 2)
    Set BarBottomRight = BarTopLeft.Offset(13, 8)
    
    Set BarChartRange = Range(BarTopLeft.Address & ":" & BarBottomRight.Address)
    
    Set SourceTopLeft = Tally_Tbl.ListColumns(1).Range.Cells(1)
    Set SourceBottomRight = Tally_Tbl.ListColumns(Tally_Tbl.Range.Columns.Count - 2).DataBodyRange(Tally_Tbl.DataBodyRange.Rows.Count - 1, 1)
    
    Set SourceDataRange = Range(SourceTopLeft.Address & ":" & SourceBottomRight.Address)

    Set BarChart = Worksheets(SheetName).Shapes.AddChart2(XlChartType:=xlColumnStacked, _
                                            left:=BarTopLeft.left, Top:=BarTopLeft.Top, _
                                            Width:=BarChartRange.Width, Height:=BarChartRange.Height).Chart
        'format bar chart
        With BarChart
            .SetSourceData Source:=SourceDataRange
            .SeriesCollection(1).XValues = Range(Tally_Tbl.ListColumns(1).DataBodyRange(1, 1).Address & ":" & Tally_Tbl.ListColumns(1).DataBodyRange(Tally_Tbl.DataBodyRange.Rows.Count - 1, 1).Address) 'set x axis
            'Debug.Print Range(Tally_Tbl.ListColumns(Tally_Tbl.ListColumns.Count).DataBodyRange(1, 1).Address & ":" & Tally_Tbl.ListColumns(Tally_Tbl.ListColumns.Count).DataBodyRange(Tally_Tbl.DataBodyRange.Rows.Count - 1, Tally_Tbl.ListColumns.Count - 2).Address).Address
            .SeriesCollection.NewSeries.values = Range(Tally_Tbl.ListColumns(Tally_Tbl.ListColumns.Count).DataBodyRange(1, 1).Address & ":" & Tally_Tbl.ListColumns(Tally_Tbl.ListColumns.Count).DataBodyRange(Tally_Tbl.DataBodyRange.Rows.Count - 1, 1).Address)
            .SeriesCollection(Tally_Tbl.ListColumns.Count - 2).XValues = Range(Tally_Tbl.ListColumns(1).DataBodyRange(1, 1).Address & ":" & Tally_Tbl.ListColumns(1).DataBodyRange(Tally_Tbl.DataBodyRange.Rows.Count - 1, 1).Address) 'set x axis
            '.SeriesCollection(2).HasDataLabels = True
            .SeriesCollection(Tally_Tbl.ListColumns.Count - 2).Name = "Running Total"
            .FullSeriesCollection(1).ChartType = xlColumnClustered
            .FullSeriesCollection(1).AxisGroup = 1
            .FullSeriesCollection(Tally_Tbl.ListColumns.Count - 2).ChartType = xlLine
            .FullSeriesCollection(Tally_Tbl.ListColumns.Count - 2).AxisGroup = 2
            .HasTitle = True
            .ChartTitle.Text = "KPIs: " & CStr(Worksheets(SheetName).Range("B1").value) & " " & year_val
            .SetElement (msoElementLegendBottom)
            .SetElement msoElementPrimaryValueAxisShow
            .SetElement msoElementPrimaryValueAxisTitleHorizontal
            .Axes(xlValue).AxisTitle.Caption = "Totals"
            
            Count = 0
            For Each x In ColorColl
                Count = Count + 1
                .SeriesCollection(Count).Format.Line.ForeColor.RGB = RGB(0, 0, 0)   'outline bar with black line
                .SeriesCollection(Count).Format.Fill.ForeColor.RGB = x          'match bar with KPI color
            Next
                
            If .SeriesCollection.Count > ColorColl.Count Then
                .SeriesCollection(Count + 1).Format.Line.ForeColor.RGB = RGB(0, 0, 0)
               .SeriesCollection(Count + 1).Format.Fill.ForeColor.RGB = RGB(234, 246, 148) 'format "Empty"
            End If
        End With

With Worksheets(SheetName).Cells
    .EntireColumn.AutoFit
    .EntireRow.AutoFit
End With

End Sub

Sub KPIChartsForm()
    KPI_Chart_Choice.Show vbModeless
End Sub

Sub CreateKPIQuarter(start_date As Date, end_date As Date, cat_val As String)

Dim start_month As String
Dim start_year As Long
Dim end_month As String
Dim end_year As Long
Dim Num_Of_Months As Integer
Dim Num_Of_Years As Integer
Dim DatesColl As New Collection
Dim datetoadd As Date
Dim key As Long

startdateval = start_date
enddateval = end_date
'cat_val = "Safety"
Num_Of_Months = DateDiff("m", startdateval, enddateval)
Num_Of_Years = DateDiff("yyyy", startdateval, enddateval)

If Num_Of_Months = 0 Then
    Num_Of_Months = 1
End If
If Num_Of_Years = 0 Then
    Num_Of_Years = 1
End If

For N = 0 To Num_Of_Months - 1
    If N = 0 Then
        datetoadd = startdateval
        Else
        datetoadd = DateAdd("m", 1, datetoadd)
    End If
    
    key = N + 1
    
    DatesColl.Add datetoadd, CStr(N + 1)
Next N



Dim SheetName As String
        SheetName = "KPI Chart"
        
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
    
    
    'setting countermeasures table as source for monthly report data
Dim counter_tbl As ListObject
   Set counter_sheet = Worksheets("Countermeasures")
   Set counter_tbl = counter_sheet.ListObjects("Tbl_Counter")
   
   Set KPI_rng = counter_tbl.ListColumns("KPI").DataBodyRange
Dim KPI_Cal_Rng As Range
 
    Dim ArrBase() As Variant
        row_count = 0
            For Each Cell3 In counter_tbl.ListColumns("Category").DataBodyRange
            row_count = 1 + row_count
                If Cell3 = cat_val Then
                                        
                    KPICell = counter_tbl.ListColumns("KPI").DataBodyRange(row_count, 1).value
                    
                    'if first entry, redim to hold one spot "(0)"
                    If entry_count = 0 Then
                        ReDim Preserve ArrBase(0)
                        ArrBase(0) = KPICell
                     'For all subsequent entries extend array by 1 and enter contents in cell
                    Else
                        ReDim Preserve ArrBase(UBound(ArrBase) + 1)
                        ArrBase(UBound(ArrBase)) = KPICell
                    End If
                    entry_count = entry_count + 1
                End If
             Next Cell3
             
            If (Not Not ArrBase) = 0 Then       'if array never intitialized then it doesn't exist,
                     IsArrayEmpty = True         'if list doesn't exist, then it's empty (by default)
                 Else
                     IsArrayEmpty = True         'if list does exist, assume it is empty
                     For Each item_in_array In ArrBase       'test if array is empty. If there is one non-blank cell, change bool value
                         If item_in_array <> Empty Then
                             IsArrayEmpty = False
                         End If
                     Next
             End If
                     
             If IsArrayEmpty = False Then            'if array isn't empty, remove blanks, remove dups
                     ArrBase = BlankRemover(ArrBase)     'if not empty, remove blanks
                     ArrBase = ArrayRemoveDups(ArrBase)  'if not empty, remove duplicates
                 ElseIf IsArrayEmpty = True Then     'if array is empty,  resize to one entry and create one entry "No List Available"
                     ReDim Preserve ArrBase(0)
                     ArrBase(0) = "No List Available"
            End If
                
                'MsgBox Join(ArrBase, vbCrLf)
                
    Dim ArrDict As Scripting.Dictionary
    Set ArrDict = New Scripting.Dictionary
    
    'rids duplicates and counts new dictionary
    Set ArrDict = DuplicateCountToScript(ArrBase)
    
    'Call Functions_M.SortDictionary(ArrDict, False, True, vbBinaryCompare)
        'place KPI text at top of sheet starting from col D and moving right
    OffsetCount = 0
    For Each item In ArrDict.Keys
        Sheets(SheetName).Range("D1").Offset(0, OffsetCount).value = item
        OffsetCount = OffsetCount + 1
    Next
    
    Range("A1").value = year_val
    Range("B1").value = cat_val
    Range("C1").value = "KPIs: "
    
    With Range("A1:C1")
        .HorizontalAlignment = xlCenter
            Range("C1").HorizontalAlignment = xlLeft
        .WrapText = True
    End With
    
    
    Dim ColorColl As New Collection

    'Format unique KPIs in the first row (or the KPI range)
    For i = 0 To ArrDict.Count - 1
        'red = 200 - (25 * (i))
        'If red <= 200 Then
            Red = GenerateRandomInt(255, 50)
            green = GenerateRandomInt(255, 50)
            Blue = GenerateRandomInt(255, 50)
            'Else
            'green = 0
        'End If
        With Range("D1").Offset(0, i)
           .VerticalAlignment = xlCenter
           .Font.Size = 9
           .Font.Bold = False
           .RowHeight = 25
           .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThick
            .Borders.Color = RGB(255, 0, 0)
            .WrapText = True
            .Interior.Color = RGB(Red, green, Blue)
           End With
        
        ColorColl.Add Range("D1").Offset(0, i).Interior.Color, Range("D1").Offset(0, i).value   'adding color with key equal to KPI text
    Next
    
    
    ActiveWindow.FreezePanes = False
    'Freeze top row so KPIs are visible
    Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    
    
            'create tally table!!
            MyName = "TallyTable"     'name for monthly table
            
            ExtraColumns = 4 'Number of extra columns after KPIs ("Empty, Total, Running Total, and Year")
            
           Set Top_Left = Range("C3")    'set top left of table
           Set Bottom_Right = Top_Left.Offset(Num_Of_Months + 1, OffsetCount + ExtraColumns)    'set bottom right = 12 down, # of different KPIs across
        
            MyRangeString = Top_Left.Address & ":" & Bottom_Right.Address   'make table address string
        
            MyListExists = False                                'assume table does not exist
            For Each ListObj In Sheets(SheetName).ListObjects
                If ListObj.Name = MyName Then MyListExists = True   'if table exists, true
            Next ListObj
            
            If Not (MyListExists) Then      'if table does not exist, create
                Set Tally_Tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Range(MyRangeString), , xlYes)
                Tally_Tbl.Name = MyName
                Else                        'if exists, designate variable
                Set Tally_Tbl = ActiveSheet.ListObjects(MyName)
            End If
            
            
            Tally_Tbl.HeaderRowRange(1, 1).value = "Year"   'put table name in table
            Tally_Tbl.HeaderRowRange(1, 2).value = "Month"
            
            key = 0
            For col = 3 To Tally_Tbl.Range.Columns.Count - 3    'label columns in monthly table
                Tally_Tbl.HeaderRowRange(1, col).value = ArrDict.Keys(key)
                key = key + 1
            Next col
            
            Tally_Tbl.HeaderRowRange(1, Tally_Tbl.Range.Columns.Count - 2).value = "[Empty]"
            Tally_Tbl.HeaderRowRange(1, Tally_Tbl.Range.Columns.Count - 1).value = "Total" 'create total column
            Tally_Tbl.HeaderRowRange(1, Tally_Tbl.Range.Columns.Count).value = "Running Total" 'create running total column
            
            For row = 1 To Tally_Tbl.ListColumns(1).DataBodyRange.Rows.Count - 1      'label frequency columns
                Count = Count + 1
                Tally_Tbl.ListColumns(1).DataBodyRange(Count, 1) = Year(DatesColl(Count))
                Tally_Tbl.ListColumns(2).DataBodyRange(Count, 1) = MonthName(Month(DatesColl(Count)))
            Next row
            
            Tally_Tbl.HeaderRowRange(Tally_Tbl.ListColumns(1).Range.Rows.Count, 1).value = "Total"  'label total KPI row
            
            With Tally_Tbl.Range
                .BorderAround xlContinuous, xlMedium, xlColorIndexAutomatic
            End With
            Count = 0
            
            col_count = 0
            For Each col In Tally_Tbl.HeaderRowRange
                col_count = col_count + 1
                For Each KPI In ArrDict.Keys
                    If col = KPI Then
                        Tally_Tbl.ListColumns(col_count).Range(1).Interior.Color = ColorColl(KPI)
                    End If
                Next
                
                If col = "[Empty]" Then
                    Tally_Tbl.ListColumns("[Empty]").Range(1).Interior.Color = RGB(234, 246, 148)
                End If
                
            Next
    
    Dim Mnth As Integer
    Dim DateVal As Date
    
    key = 0
    For Each monthyear In DatesColl
       key = key + 1
            'Create monthly table!!
            ' if excel table already exists, delete and replace with new one
               
            MyName = MonthName(Month(monthyear)) & "_" & Year(monthyear) & "_Table"      'name for monthly table
            
                DateVal = DateValue(MonthName(Month(monthyear)) & "/1/" & Year(monthyear))  'date value for first day of month
              
               Set Top_Left = Tally_Tbl.ListColumns(1).DataBodyRange(Tally_Tbl.ListColumns(1).DataBodyRange.Rows.Count, 1).Offset((11 * (key - 1)) + 3, 0)  'set top leftof table, offsetting by bottom of tally tale by 2 rows
               Set Bottom_Right = Top_Left.Offset(10, MonthDays(DateVal))
            
            MyRangeString = Top_Left.Address & ":" & Bottom_Right.Address   'make table address string
        
            MyListExists = False                                'assume table does not exist
            For Each ListObj In Sheets(SheetName).ListObjects
                If ListObj.Name = MyName Then MyListExists = True   'if table exists, true
            Next ListObj
            
            If Not (MyListExists) Then      'if table does not exist, create
                Set NewTable = ActiveSheet.ListObjects.Add(xlSrcRange, Range(MyRangeString), , xlYes)
                NewTable.Name = MyName
                Else                        'if exists, designate variable
                Set NewTable = ActiveSheet.ListObjects(MyName)
            End If
            
            NewTable.HeaderRowRange(1, 1).value = MonthName(Month(monthyear))   'put month name in table
            NewTable.HeaderRowRange(1, 1).Offset(0, -1) = CStr(Year(DateVal))
            
            For col = 2 To NewTable.Range.Columns.Count + 1     'label columns in monthly table
                NewTable.HeaderRowRange(1, col).value = col - 1
            Next col
            
            For Each cell In NewTable.ListColumns(1).DataBodyRange      'label frequency columns
                Count = Count + 1
                cell.value = Count
            Next cell
            Count = 0
                
                KDay = 0        'initialize NewTable column count at 0
                For Each Column In NewTable.ListColumns                     'for each column in month table
                    
                    If IsNumeric(Column) = False Then
                                'do nothing    'if not numeric, e.g. "January", skip
                                    Else
                                CRow = 0   'row counter for countermeasures table
                                KRow = 0    'row counterfor NewTable (monthly KPI)
                                KMonth = Month(monthyear)   'KPI month = month of current table
                                KYear = Year(monthyear)    'KPI year = year val of macro
                                KDay = KDay + 1      'KPI day = column header
                                
                                
                                    For Each cell In counter_tbl.ListColumns("Issue Date").DataBodyRange    'for each cell in countermeasures table "Issue Date" column
                                        
                                        CRow = CRow + 1     'Row count for countermeasures table
                                        
                                        CMonth = Month(cell.value)  'month of issue date
                                        CDay = Day(cell.value)      'day of issue date
                                        CYear = Year(cell.value)    'year of issue date
                                        
                                            'if issue date Month = NewTable Month, and issue day = column and isue year = year of chart
                                            ' and category or issue date = category value then
                                            'add KPI row + 1
                                            'add KPI value from countermeasures table to NewTable
                                        
                                        If CMonth = KMonth And CDay = KDay And CYear = KYear And counter_tbl.ListColumns("Category").DataBodyRange(CRow, 1).value = cat_val Then
                                                KRow = KRow + 1     'add NewTable row count
                                                
                                                If counter_tbl.ListColumns("KPI").DataBodyRange(CRow, 1).value <> Empty Then
                                                        
                                                        With NewTable.ListColumns(KDay + 1).DataBodyRange(KRow, 1)
                                                                .value = counter_tbl.ListColumns("KPI").DataBodyRange(CRow, 1).value     'add KPI value from countermeasures table to NewTable(KDay + 1 because in each month table there is one extra column for month name)
                                                                .Interior.Color = ColorColl(counter_tbl.ListColumns("KPI").DataBodyRange(CRow, 1).value)
                                                        End With
                                                        
                                                        Old_Tally = Tally_Tbl.ListColumns(counter_tbl.ListColumns("KPI").DataBodyRange(CRow, 1).value).DataBodyRange(key, 1).value     'old tally = old total in tally table data body range
                                                        New_Tally = Old_Tally + 1                                                                                                       'add one
                                                        Tally_Tbl.ListColumns(counter_tbl.ListColumns("KPI").DataBodyRange(CRow, 1).value).DataBodyRange(key, 1).value = New_Tally     'replace old tally with new tally in tally table
                                                        
                                                    Else
                                                        With NewTable.ListColumns(KDay + 1).DataBodyRange(KRow, 1)
                                                            .value = "[Empty]"     'add KPI value from countermeasures table to NewTable(KDay + 1 because in each month table there is one extra column for month name)
                                                            .Interior.Color = RGB(234, 246, 148)
                                                        End With
                                                        
                                                        Old_Tally = Tally_Tbl.ListColumns("[Empty]").DataBodyRange(key, 1).value     'old tally = old total in tally table data body range
                                                        New_Tally = Old_Tally + 1                                                                                                       'add one
                                                        Tally_Tbl.ListColumns("[Empty]").DataBodyRange(key, 1).value = New_Tally     'replace old tally with new tally in tally table
                                                        
                                                End If
                                        End If
                                    Next        'go to next cell in countermeasures table
                    End If
                    Next Column                 'go to next column (day) in month (newtable)
                    
                    NewTable.Range.WrapText = True
                
    Next monthyear                                   'go to next month in Calendar
    
        'add totals in TallyTable for each row ("Totals" column)
    Sum = 0
    For row = 1 To Tally_Tbl.DataBodyRange.Rows.Count - 1
        Sum = 0
        
        For col = 3 To Tally_Tbl.Range.Columns.Count - 2
            Sum = Sum + Tally_Tbl.DataBodyRange(row, col).value
        Next col
        
        Tally_Tbl.ListColumns("Total").DataBodyRange(row, 1).value = Sum
        
    Next row

        'add totals in TallyTable for each column ("Totals" row)
     Sum = 0
    For col = 3 To Tally_Tbl.Range.Columns.Count - 1
        Sum = 0
        
        For row = 1 To Tally_Tbl.DataBodyRange.Rows.Count - 1
            Sum = Sum + Tally_Tbl.DataBodyRange(row, col).value
        Next row
        
        Tally_Tbl.ListColumns(col).DataBodyRange(Tally_Tbl.DataBodyRange.Rows.Count, 1).value = Sum
        
    Next col
    
        'add running totals in TallyTable for last column ("Running Totals" col)
     RunningTotal = 0
    For row = 1 To Tally_Tbl.DataBodyRange.Rows.Count - 1
        
        If row = 1 Then
            RunningTotal = Tally_Tbl.ListColumns("Total").DataBodyRange(row, 1).value
            Else
            RunningTotal = RunningTotal + Tally_Tbl.ListColumns("Total").DataBodyRange(row, 1).value
        End If
        
        Tally_Tbl.ListColumns("Running Total").DataBodyRange(row, 1).value = RunningTotal
        
    Next row
    
        'format tally table
    Tally_Tbl.TableStyle = "TableStyleLight1"
    Tally_Tbl.DataBodyRange.HorizontalAlignment = xlCenter
    Tally_Tbl.ListColumns(1).DataBodyRange.HorizontalAlignment = xlLeft
    Tally_Tbl.Range.WrapText = True
    
    
        'create bar chart
    
    Set Tally_Top_Right = Tally_Tbl.Range(1, Tally_Tbl.Range.Columns.Count)
    
    Set BarTopLeft = Tally_Top_Right.Offset(0, 2)
    Set BarBottomRight = BarTopLeft.Offset(13, 8)
    
    Set BarChartRange = Range(BarTopLeft.Address & ":" & BarBottomRight.Address)
    
    Set SourceTopLeft = Tally_Tbl.ListColumns(1).Range.Cells(1)
    Set SourceBottomRight = Tally_Tbl.ListColumns(Tally_Tbl.Range.Columns.Count - 2).DataBodyRange(Tally_Tbl.DataBodyRange.Rows.Count - 1, 1)
    
    Set SourceDataRange = Range(SourceTopLeft.Address & ":" & SourceBottomRight.Address)

    Set BarChart = Worksheets(SheetName).Shapes.AddChart2(XlChartType:=201, _
                                            left:=BarTopLeft.left, Top:=BarTopLeft.Top, _
                                            Width:=BarChartRange.Width, Height:=BarChartRange.Height).Chart
        'format bar chart
        With BarChart
            .SetSourceData Source:=SourceDataRange
            .SeriesCollection(1).XValues = Range(Tally_Tbl.ListColumns(1).DataBodyRange(1, 1).Address & ":" & Tally_Tbl.ListColumns(1).DataBodyRange(Tally_Tbl.DataBodyRange.Rows.Count - 1, 1).Address) 'set x axis
            'Debug.Print Range(Tally_Tbl.ListColumns(Tally_Tbl.ListColumns.Count).DataBodyRange(1, 1).Address & ":" & Tally_Tbl.ListColumns(Tally_Tbl.ListColumns.Count).DataBodyRange(Tally_Tbl.DataBodyRange.Rows.Count - 1, Tally_Tbl.ListColumns.Count - 2).Address).Address
            .SeriesCollection.NewSeries.values = Range(Tally_Tbl.ListColumns(Tally_Tbl.ListColumns.Count).Range(1, 1).Address & ":" & Tally_Tbl.ListColumns(Tally_Tbl.ListColumns.Count).DataBodyRange(Tally_Tbl.DataBodyRange.Rows.Count - 1, 1).Address)
            .SeriesCollection(Tally_Tbl.ListColumns.Count - 2).XValues = Range(Tally_Tbl.ListColumns(Tally_Tbl.ListColumns.Count).DataBodyRange(1, 1).Address & ":" & Tally_Tbl.ListColumns(Tally_Tbl.ListColumns.Count).DataBodyRange(Tally_Tbl.DataBodyRange.Rows.Count - 1, 1).Address) 'set x axis
            '.SeriesCollection(2).HasDataLabels = True
            .SeriesCollection(Tally_Tbl.ListColumns.Count - 2).Name = "Running Total"
            .FullSeriesCollection(1).ChartType = xlColumnClustered
            .FullSeriesCollection(1).AxisGroup = 1
            .FullSeriesCollection(Tally_Tbl.ListColumns.Count - 2).ChartType = xlLine
            .FullSeriesCollection(Tally_Tbl.ListColumns.Count - 2).AxisGroup = 2
            .HasTitle = True
            If Month(enddateval) = 1 Then
            .ChartTitle.Text = "KPIs: " & CStr(MonthName(Month(startdateval))) & " " & CStr(Year(startdateval)) & " - " & CStr(MonthName(1)) & " " & CStr(Year(enddateval) - 1)
            Else
            .ChartTitle.Text = "KPIs: " & CStr(MonthName(Month(startdateval))) & " " & CStr(Year(startdateval)) & " - " & CStr(MonthName(Month(enddateval))) & " " & CStr(Year(enddateval))
            End If
            .SetElement (msoElementLegendBottom)
            .SetElement msoElementPrimaryValueAxisShow
            .SetElement msoElementPrimaryValueAxisTitleHorizontal
            .Axes(xlValue).AxisTitle.Caption = "Totals"
            
            Count = 0
            For Each x In ColorColl
                Count = Count + 1
                .SeriesCollection(Count).Format.Line.ForeColor.RGB = RGB(0, 0, 0)   'outline bar with black line
                .SeriesCollection(Count).Format.Fill.ForeColor.RGB = x          'match bar with KPI color
            Next
                
            If .SeriesCollection.Count > ColorColl.Count Then
                .SeriesCollection(Count + 1).Format.Line.ForeColor.RGB = RGB(0, 0, 0)
               .SeriesCollection(Count + 1).Format.Fill.ForeColor.RGB = RGB(234, 246, 148) 'format "Empty"
            End If
        End With
        
With Worksheets(SheetName).Cells
    .EntireColumn.AutoFit
    .EntireRow.AutoFit
End With

End Sub


Sub CreateKPIMonth(start_date As Date, end_date As Date, cat_val As String)

Dim start_month As String
Dim start_year As Long
Dim end_month As String
Dim end_year As Long
Dim Num_Of_Months As Integer
Dim Num_Of_Years As Integer
Dim DatesColl As New Collection
Dim datetoadd As Date
Dim key As Long

startdateval = start_date
enddateval = end_date
'cat_val = "Safety"
Num_Of_Months = DateDiff("m", startdateval, enddateval)
Num_Of_Years = DateDiff("yyyy", startdateval, enddateval)

If Num_Of_Months = 0 Then
    Num_Of_Months = 1
End If
If Num_Of_Years = 0 Then
    Num_Of_Years = 1
End If

For N = 0 To Num_Of_Months - 1
    If N = 0 Then
        datetoadd = startdateval
        Else
        datetoadd = DateAdd("m", 1, datetoadd)
    End If
    
    key = N + 1
    
    DatesColl.Add datetoadd, CStr(N + 1)
Next N



Dim SheetName As String
        SheetName = "KPI Chart"
        
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
    
    
    'setting countermeasures table as source for monthly report data
Dim counter_tbl As ListObject
   Set counter_sheet = Worksheets("Countermeasures")
   Set counter_tbl = counter_sheet.ListObjects("Tbl_Counter")
   
   Set KPI_rng = counter_tbl.ListColumns("KPI").DataBodyRange
Dim KPI_Cal_Rng As Range
 
    Dim ArrBase() As Variant
        row_count = 0
            For Each Cell3 In counter_tbl.ListColumns("Category").DataBodyRange
            row_count = 1 + row_count
                If Cell3 = cat_val Then
                                        
                    KPICell = counter_tbl.ListColumns("KPI").DataBodyRange(row_count, 1).value
                    
                    'if first entry, redim to hold one spot "(0)"
                    If entry_count = 0 Then
                        ReDim Preserve ArrBase(0)
                        ArrBase(0) = KPICell
                     'For all subsequent entries extend array by 1 and enter contents in cell
                    Else
                        ReDim Preserve ArrBase(UBound(ArrBase) + 1)
                        ArrBase(UBound(ArrBase)) = KPICell
                    End If
                    entry_count = entry_count + 1
                End If
             Next Cell3
             
            If (Not Not ArrBase) = 0 Then       'if array never intitialized then it doesn't exist,
                     IsArrayEmpty = True         'if list doesn't exist, then it's empty (by default)
                 Else
                     IsArrayEmpty = True         'if list does exist, assume it is empty
                     For Each item_in_array In ArrBase       'test if array is empty. If there is one non-blank cell, change bool value
                         If item_in_array <> Empty Then
                             IsArrayEmpty = False
                         End If
                     Next
             End If
                     
             If IsArrayEmpty = False Then            'if array isn't empty, remove blanks, remove dups
                     ArrBase = BlankRemover(ArrBase)     'if not empty, remove blanks
                     ArrBase = ArrayRemoveDups(ArrBase)  'if not empty, remove duplicates
                 ElseIf IsArrayEmpty = True Then     'if array is empty,  resize to one entry and create one entry "No List Available"
                     ReDim Preserve ArrBase(0)
                     ArrBase(0) = "No List Available"
            End If
                
                'MsgBox Join(ArrBase, vbCrLf)
                
    Dim ArrDict As Scripting.Dictionary
    Set ArrDict = New Scripting.Dictionary
    
    'rids duplicates and counts new dictionary
    Set ArrDict = DuplicateCountToScript(ArrBase)
    
    'Call Functions_M.SortDictionary(ArrDict, False, True, vbBinaryCompare)
        'place KPI text at top of sheet starting from col D and moving right
    OffsetCount = 0
    For Each item In ArrDict.Keys
        Sheets(SheetName).Range("D1").Offset(0, OffsetCount).value = item
        OffsetCount = OffsetCount + 1
    Next
    
    Range("A1").value = year_val
    Range("B1").value = cat_val
    Range("C1").value = "KPIs: "
    
    With Range("A1:C1")
        .HorizontalAlignment = xlCenter
            Range("C1").HorizontalAlignment = xlLeft
        .WrapText = True
    End With
    
    
    Dim ColorColl As New Collection

    'Format unique KPIs in the first row (or the KPI range)
    For i = 0 To ArrDict.Count - 1
        'red = 200 - (25 * (i))
        'If red <= 200 Then
            Red = GenerateRandomInt(255, 50)
            green = GenerateRandomInt(255, 50)
            Blue = GenerateRandomInt(255, 50)
            'Else
            'green = 0
        'End If
        With Range("D1").Offset(0, i)
           .VerticalAlignment = xlCenter
           .Font.Size = 9
           .Font.Bold = False
           .RowHeight = 25
           .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThick
            .Borders.Color = RGB(255, 0, 0)
            .WrapText = True
            .Interior.Color = RGB(Red, green, Blue)
           End With
        
        ColorColl.Add Range("D1").Offset(0, i).Interior.Color, Range("D1").Offset(0, i).value   'adding color with key equal to KPI text
    Next
    
    
    ActiveWindow.FreezePanes = False
    'Freeze top row so KPIs are visible
    Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    
    
            'create tally table!!
            MyName = "TallyTable"     'name for monthly table
            
            ExtraColumns = 4 'Number of extra columns after KPIs ("Empty, Total, Running Total, and Year")
            
           Set Top_Left = Range("C3")    'set top left of table
           Set Bottom_Right = Top_Left.Offset(Num_Of_Months + 1, OffsetCount + ExtraColumns)    'set bottom right = 12 down, # of different KPIs across
        
            MyRangeString = Top_Left.Address & ":" & Bottom_Right.Address   'make table address string
        
            MyListExists = False                                'assume table does not exist
            For Each ListObj In Sheets(SheetName).ListObjects
                If ListObj.Name = MyName Then MyListExists = True   'if table exists, true
            Next ListObj
            
            If Not (MyListExists) Then      'if table does not exist, create
                Set Tally_Tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Range(MyRangeString), , xlYes)
                Tally_Tbl.Name = MyName
                Else                        'if exists, designate variable
                Set Tally_Tbl = ActiveSheet.ListObjects(MyName)
            End If
            
            
            Tally_Tbl.HeaderRowRange(1, 1).value = "Year"   'put table name in table
            Tally_Tbl.HeaderRowRange(1, 2).value = "Month"
            
            key = 0
            For col = 3 To Tally_Tbl.Range.Columns.Count - 3    'label columns in monthly table
                Tally_Tbl.HeaderRowRange(1, col).value = ArrDict.Keys(key)
                key = key + 1
            Next col
            
            Tally_Tbl.HeaderRowRange(1, Tally_Tbl.Range.Columns.Count - 2).value = "[Empty]"
            Tally_Tbl.HeaderRowRange(1, Tally_Tbl.Range.Columns.Count - 1).value = "Total" 'create total column
            Tally_Tbl.HeaderRowRange(1, Tally_Tbl.Range.Columns.Count).value = "Running Total" 'create running total column
            
            For row = 1 To Tally_Tbl.ListColumns(1).DataBodyRange.Rows.Count - 1      'label frequency columns
                Count = Count + 1
                Tally_Tbl.ListColumns(1).DataBodyRange(Count, 1) = Year(DatesColl(Count))
                Tally_Tbl.ListColumns(2).DataBodyRange(Count, 1) = MonthName(Month(DatesColl(Count)))
            Next row
            
            Tally_Tbl.HeaderRowRange(Tally_Tbl.ListColumns(1).Range.Rows.Count, 1).value = "Total"  'label total KPI row
            
            With Tally_Tbl.Range
                .BorderAround xlContinuous, xlMedium, xlColorIndexAutomatic
            End With
            Count = 0
            
            col_count = 0
            For Each col In Tally_Tbl.HeaderRowRange
                col_count = col_count + 1
                For Each KPI In ArrDict.Keys
                    If col = KPI Then
                        Tally_Tbl.ListColumns(col_count).Range(1).Interior.Color = ColorColl(KPI)
                    End If
                Next
                
                If col = "[Empty]" Then
                    Tally_Tbl.ListColumns("[Empty]").Range(1).Interior.Color = RGB(234, 246, 148)
                End If
                
            Next
    
    Dim Mnth As Integer
    Dim DateVal As Date
    
    key = 0
    For Each monthyear In DatesColl
       key = key + 1
            'Create monthly table!!
            ' if excel table already exists, delete and replace with new one
               
            MyName = MonthName(Month(monthyear)) & "_" & Year(monthyear) & "_Table"      'name for monthly table
            
                DateVal = DateValue(MonthName(Month(monthyear)) & "/1/" & Year(monthyear))  'date value for first day of month
              
               Set Top_Left = Tally_Tbl.ListColumns(1).DataBodyRange(Tally_Tbl.ListColumns(1).DataBodyRange.Rows.Count, 1).Offset((11 * (key - 1)) + 3, 0)  'set top leftof table, offsetting by bottom of tally tale by 2 rows
               Set Bottom_Right = Top_Left.Offset(10, MonthDays(DateVal))
            
            MyRangeString = Top_Left.Address & ":" & Bottom_Right.Address   'make table address string
        
            MyListExists = False                                'assume table does not exist
            For Each ListObj In Sheets(SheetName).ListObjects
                If ListObj.Name = MyName Then MyListExists = True   'if table exists, true
            Next ListObj
            
            If Not (MyListExists) Then      'if table does not exist, create
                Set NewTable = ActiveSheet.ListObjects.Add(xlSrcRange, Range(MyRangeString), , xlYes)
                NewTable.Name = MyName
                Else                        'if exists, designate variable
                Set NewTable = ActiveSheet.ListObjects(MyName)
            End If
            
            NewTable.HeaderRowRange(1, 1).value = MonthName(Month(monthyear))   'put month name in table
            NewTable.HeaderRowRange(1, 1).Offset(0, -1) = CStr(Year(DateVal))
            
            For col = 2 To NewTable.Range.Columns.Count + 1     'label columns in monthly table
                NewTable.HeaderRowRange(1, col).value = col - 1
            Next col
            
            For Each cell In NewTable.ListColumns(1).DataBodyRange      'label frequency columns
                Count = Count + 1
                cell.value = Count
            Next cell
            Count = 0
                
                KDay = 0        'initialize NewTable column count at 0
                For Each Column In NewTable.ListColumns                     'for each column in month table
                    
                    If IsNumeric(Column) = False Then
                                'do nothing    'if not numeric, e.g. "January", skip
                                    Else
                                CRow = 0   'row counter for countermeasures table
                                KRow = 0    'row counterfor NewTable (monthly KPI)
                                KMonth = Month(monthyear)   'KPI month = month of current table
                                KYear = Year(monthyear)    'KPI year = year val of macro
                                KDay = KDay + 1      'KPI day = column header
                                
                                
                                    For Each cell In counter_tbl.ListColumns("Issue Date").DataBodyRange    'for each cell in countermeasures table "Issue Date" column
                                        
                                        CRow = CRow + 1     'Row count for countermeasures table
                                        
                                        CMonth = Month(cell.value)  'month of issue date
                                        CDay = Day(cell.value)      'day of issue date
                                        CYear = Year(cell.value)    'year of issue date
                                        
                                            'if issue date Month = NewTable Month, and issue day = column and isue year = year of chart
                                            ' and category or issue date = category value then
                                            'add KPI row + 1
                                            'add KPI value from countermeasures table to NewTable
                                        
                                        If CMonth = KMonth And CDay = KDay And CYear = KYear And counter_tbl.ListColumns("Category").DataBodyRange(CRow, 1).value = cat_val Then
                                                KRow = KRow + 1     'add NewTable row count
                                                
                                                If counter_tbl.ListColumns("KPI").DataBodyRange(CRow, 1).value <> Empty Then
                                                        
                                                        With NewTable.ListColumns(KDay + 1).DataBodyRange(KRow, 1)
                                                                .value = counter_tbl.ListColumns("KPI").DataBodyRange(CRow, 1).value     'add KPI value from countermeasures table to NewTable(KDay + 1 because in each month table there is one extra column for month name)
                                                                .Interior.Color = ColorColl(counter_tbl.ListColumns("KPI").DataBodyRange(CRow, 1).value)
                                                        End With
                                                        
                                                        Old_Tally = Tally_Tbl.ListColumns(counter_tbl.ListColumns("KPI").DataBodyRange(CRow, 1).value).DataBodyRange(key, 1).value     'old tally = old total in tally table data body range
                                                        New_Tally = Old_Tally + 1                                                                                                       'add one
                                                        Tally_Tbl.ListColumns(counter_tbl.ListColumns("KPI").DataBodyRange(CRow, 1).value).DataBodyRange(key, 1).value = New_Tally     'replace old tally with new tally in tally table
                                                        
                                                    Else
                                                        With NewTable.ListColumns(KDay + 1).DataBodyRange(KRow, 1)
                                                            .value = "[Empty]"     'add KPI value from countermeasures table to NewTable(KDay + 1 because in each month table there is one extra column for month name)
                                                            .Interior.Color = RGB(234, 246, 148)
                                                        End With
                                                        
                                                        Old_Tally = Tally_Tbl.ListColumns("[Empty]").DataBodyRange(key, 1).value     'old tally = old total in tally table data body range
                                                        New_Tally = Old_Tally + 1                                                                                                       'add one
                                                        Tally_Tbl.ListColumns("[Empty]").DataBodyRange(key, 1).value = New_Tally     'replace old tally with new tally in tally table
                                                        
                                                End If
                                        End If
                                    Next        'go to next cell in countermeasures table
                    End If
                    Next Column                 'go to next column (day) in month (newtable)
                    
                    NewTable.Range.WrapText = True
                
    Next monthyear                                   'go to next month in Calendar
    
        'add totals in TallyTable for each row ("Totals" column)
    Sum = 0
    For row = 1 To Tally_Tbl.DataBodyRange.Rows.Count - 1
        Sum = 0
        
        For col = 3 To Tally_Tbl.Range.Columns.Count - 2
            Sum = Sum + Tally_Tbl.DataBodyRange(row, col).value
        Next col
        
        Tally_Tbl.ListColumns("Total").DataBodyRange(row, 1).value = Sum
        
    Next row

        'add totals in TallyTable for each column ("Totals" row)
     Sum = 0
    For col = 3 To Tally_Tbl.Range.Columns.Count - 1
        Sum = 0
        
        For row = 1 To Tally_Tbl.DataBodyRange.Rows.Count - 1
            Sum = Sum + Tally_Tbl.DataBodyRange(row, col).value
        Next row
        
        Tally_Tbl.ListColumns(col).DataBodyRange(Tally_Tbl.DataBodyRange.Rows.Count, 1).value = Sum
        
    Next col
    
        'add running totals in TallyTable for last column ("Running Totals" col)
     RunningTotal = 0
    For row = 1 To Tally_Tbl.DataBodyRange.Rows.Count - 1
        
        If row = 1 Then
            RunningTotal = Tally_Tbl.ListColumns("Total").DataBodyRange(row, 1).value
            Else
            RunningTotal = RunningTotal + Tally_Tbl.ListColumns("Total").DataBodyRange(row, 1).value
        End If
        
        Tally_Tbl.ListColumns("Running Total").DataBodyRange(row, 1).value = RunningTotal
        
    Next row
    
        'format tally table
    Tally_Tbl.TableStyle = "TableStyleLight1"
    Tally_Tbl.DataBodyRange.HorizontalAlignment = xlCenter
    Tally_Tbl.ListColumns(1).DataBodyRange.HorizontalAlignment = xlLeft
    Tally_Tbl.Range.WrapText = True
    
    
        'create bar chart
    
    Set Tally_Top_Right = Tally_Tbl.Range(1, Tally_Tbl.Range.Columns.Count)
    
    Set BarTopLeft = Tally_Top_Right.Offset(0, 2)
    Set BarBottomRight = BarTopLeft.Offset(13, 8)
    
    Set BarChartRange = Range(BarTopLeft.Address & ":" & BarBottomRight.Address)
    
    Set SourceTopLeft = Tally_Tbl.ListColumns(1).Range.Cells(1)
    Set SourceBottomRight = Tally_Tbl.ListColumns(Tally_Tbl.Range.Columns.Count - 2).DataBodyRange(Tally_Tbl.DataBodyRange.Rows.Count - 1, 1)
    
    Set SourceDataRange = Range(SourceTopLeft.Address & ":" & SourceBottomRight.Address)

    Set BarChart = Worksheets(SheetName).Shapes.AddChart2(XlChartType:=xlColumnStacked, _
                                            left:=BarTopLeft.left, Top:=BarTopLeft.Top, _
                                            Width:=BarChartRange.Width, Height:=BarChartRange.Height).Chart
        'format bar chart
        With BarChart
            .SetSourceData Source:=SourceDataRange
            '.SeriesCollection(1).XValues = Range(Tally_Tbl.ListColumns(1).DataBodyRange(2, 1).Address & ":" & Tally_Tbl.ListColumns(1).DataBodyRange(Tally_Tbl.DataBodyRange.Rows.Count - 1, 1).Address) 'set x axis
            'Debug.Print Range(Tally_Tbl.ListColumns(Tally_Tbl.ListColumns.Count).DataBodyRange(1, 1).Address & ":" & Tally_Tbl.ListColumns(Tally_Tbl.ListColumns.Count).DataBodyRange(Tally_Tbl.DataBodyRange.Rows.Count - 1, Tally_Tbl.ListColumns.Count - 2).Address).Address
            .FullSeriesCollection(1).ChartType = xlColumnClustered
            .FullSeriesCollection(1).AxisGroup = 1
            .HasTitle = True
            .ChartTitle.Text = "KPIs: " & CStr(Worksheets(SheetName).Range("B1").value) & " " & CStr(MonthName(Month(startdateval))) & " " & CStr(Year(startdateval))
            .SetElement (msoElementLegendBottom)
            .SetElement msoElementPrimaryValueAxisShow
            .SetElement msoElementPrimaryValueAxisTitleHorizontal
            .SetElement msoElementPrimaryCategoryAxisShow   'x axis exists
            .SetElement msoElementPrimaryCategoryAxisTitleHorizontal    'show x axis
            .Axes(xlCategory, xlPrimary).HasTitle = True    'title exists
            .Axes(xlCategory, xlPrimary).AxisTitle.Caption = CStr(Tally_Tbl.ListColumns(2).DataBodyRange(1, 1).value)  'title
            .Axes(xlValue).AxisTitle.Caption = "Totals"
            
            Count = 0
            For Each x In ColorColl
                Count = Count + 1
                .SeriesCollection(Count).Format.Line.ForeColor.RGB = RGB(0, 0, 0)   'outline bar with black line
                .SeriesCollection(Count).Format.Fill.ForeColor.RGB = x          'match bar with KPI color
            Next
                
            If .SeriesCollection.Count > ColorColl.Count Then
                .SeriesCollection(Count + 1).Format.Line.ForeColor.RGB = RGB(0, 0, 0)
               .SeriesCollection(Count + 1).Format.Fill.ForeColor.RGB = RGB(234, 246, 148) 'format "Empty"
            End If
        End With

With Worksheets(SheetName).Cells
    .EntireColumn.AutoFit
    .EntireRow.AutoFit
End With

End Sub

Sub CreateKPIChartCustMonth(start_date As Date, end_date As Date, cat_val As String)

Dim start_month As String
Dim start_year As Long
Dim end_month As String
Dim end_year As Long
Dim Num_Of_Months As Integer
Dim Num_Of_Years As Integer
Dim DatesColl As New Collection
Dim datetoadd As Date
Dim key As Long

startdateval = start_date
enddateval = end_date
'cat_val = "Safety"
Num_Of_Months = DateDiff("m", startdateval, enddateval)
Num_Of_Years = DateDiff("yyyy", startdateval, enddateval)

If Num_Of_Months = 0 Then
    Num_Of_Months = 1
End If
If Num_Of_Years = 0 Then
    Num_Of_Years = 1
End If

For N = 0 To Num_Of_Months - 1
    If N = 0 Then
        datetoadd = startdateval
        Else
        datetoadd = DateAdd("m", 1, datetoadd)
    End If
    
    key = N + 1
    
    DatesColl.Add datetoadd, CStr(N + 1)
Next N



Dim SheetName As String
        SheetName = "KPI Chart"
        
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
    
    
    'setting countermeasures table as source for monthly report data
Dim counter_tbl As ListObject
   Set counter_sheet = Worksheets("Countermeasures")
   Set counter_tbl = counter_sheet.ListObjects("Tbl_Counter")
   
   Set KPI_rng = counter_tbl.ListColumns("KPI").DataBodyRange
Dim KPI_Cal_Rng As Range
 
    Dim ArrBase() As Variant
        row_count = 0
            For Each Cell3 In counter_tbl.ListColumns("Category").DataBodyRange
            row_count = 1 + row_count
                If Cell3 = cat_val Then
                                        
                    KPICell = counter_tbl.ListColumns("KPI").DataBodyRange(row_count, 1).value
                    
                    'if first entry, redim to hold one spot "(0)"
                    If entry_count = 0 Then
                        ReDim Preserve ArrBase(0)
                        ArrBase(0) = KPICell
                     'For all subsequent entries extend array by 1 and enter contents in cell
                    Else
                        ReDim Preserve ArrBase(UBound(ArrBase) + 1)
                        ArrBase(UBound(ArrBase)) = KPICell
                    End If
                    entry_count = entry_count + 1
                End If
             Next Cell3
             
            If (Not Not ArrBase) = 0 Then       'if array never intitialized then it doesn't exist,
                     IsArrayEmpty = True         'if list doesn't exist, then it's empty (by default)
                 Else
                     IsArrayEmpty = True         'if list does exist, assume it is empty
                     For Each item_in_array In ArrBase       'test if array is empty. If there is one non-blank cell, change bool value
                         If item_in_array <> Empty Then
                             IsArrayEmpty = False
                         End If
                     Next
             End If
                     
             If IsArrayEmpty = False Then            'if array isn't empty, remove blanks, remove dups
                     ArrBase = BlankRemover(ArrBase)     'if not empty, remove blanks
                     ArrBase = ArrayRemoveDups(ArrBase)  'if not empty, remove duplicates
                 ElseIf IsArrayEmpty = True Then     'if array is empty,  resize to one entry and create one entry "No List Available"
                     ReDim Preserve ArrBase(0)
                     ArrBase(0) = "No List Available"
            End If
                
                'MsgBox Join(ArrBase, vbCrLf)
                
    Dim ArrDict As Scripting.Dictionary
    Set ArrDict = New Scripting.Dictionary
    
    'rids duplicates and counts new dictionary
    Set ArrDict = DuplicateCountToScript(ArrBase)
    
    'Call Functions_M.SortDictionary(ArrDict, False, True, vbBinaryCompare)
        'place KPI text at top of sheet starting from col D and moving right
    OffsetCount = 0
    For Each item In ArrDict.Keys
        Sheets(SheetName).Range("D1").Offset(0, OffsetCount).value = item
        OffsetCount = OffsetCount + 1
    Next
    
    Range("A1").value = year_val
    Range("B1").value = cat_val
    Range("C1").value = "KPIs: "
    
    With Range("A1:C1")
        .HorizontalAlignment = xlCenter
            Range("C1").HorizontalAlignment = xlLeft
        .WrapText = True
    End With
    
    
    Dim ColorColl As New Collection

    'Format unique KPIs in the first row (or the KPI range)
    For i = 0 To ArrDict.Count - 1
        'red = 200 - (25 * (i))
        'If red <= 200 Then
            Red = GenerateRandomInt(255, 50)
            green = GenerateRandomInt(255, 50)
            Blue = GenerateRandomInt(255, 50)
            'Else
            'green = 0
        'End If
        With Range("D1").Offset(0, i)
           .VerticalAlignment = xlCenter
           .Font.Size = 9
           .Font.Bold = False
           .RowHeight = 25
           .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThick
            .Borders.Color = RGB(255, 0, 0)
            .WrapText = True
            .Interior.Color = RGB(Red, green, Blue)
           End With
        
        ColorColl.Add Range("D1").Offset(0, i).Interior.Color, Range("D1").Offset(0, i).value   'adding color with key equal to KPI text
    Next
    
    
    ActiveWindow.FreezePanes = False
    'Freeze top row so KPIs are visible
    Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    
    
            'create tally table!!
            MyName = "TallyTable"     'name for monthly table
            
            ExtraColumns = 4 'Number of extra columns after KPIs ("Empty, Total, Running Total, and Year")
            
           Set Top_Left = Range("C3")    'set top left of table
           Set Bottom_Right = Top_Left.Offset(Num_Of_Months + 1, OffsetCount + ExtraColumns)    'set bottom right = 12 down, # of different KPIs across
        
            MyRangeString = Top_Left.Address & ":" & Bottom_Right.Address   'make table address string
        
            MyListExists = False                                'assume table does not exist
            For Each ListObj In Sheets(SheetName).ListObjects
                If ListObj.Name = MyName Then MyListExists = True   'if table exists, true
            Next ListObj
            
            If Not (MyListExists) Then      'if table does not exist, create
                Set Tally_Tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Range(MyRangeString), , xlYes)
                Tally_Tbl.Name = MyName
                Else                        'if exists, designate variable
                Set Tally_Tbl = ActiveSheet.ListObjects(MyName)
            End If
            
            
            Tally_Tbl.HeaderRowRange(1, 1).value = "Year"   'put table name in table
            Tally_Tbl.HeaderRowRange(1, 2).value = "Month"
            
            key = 0
            For col = 3 To Tally_Tbl.Range.Columns.Count - 3    'label columns in monthly table
                Tally_Tbl.HeaderRowRange(1, col).value = ArrDict.Keys(key)
                key = key + 1
            Next col
            
            Tally_Tbl.HeaderRowRange(1, Tally_Tbl.Range.Columns.Count - 2).value = "[Empty]"
            Tally_Tbl.HeaderRowRange(1, Tally_Tbl.Range.Columns.Count - 1).value = "Total" 'create total column
            Tally_Tbl.HeaderRowRange(1, Tally_Tbl.Range.Columns.Count).value = "Running Total" 'create running total column
            
            For row = 1 To Tally_Tbl.ListColumns(1).DataBodyRange.Rows.Count - 1      'label frequency columns
                Count = Count + 1
                Tally_Tbl.ListColumns(1).DataBodyRange(Count, 1) = Year(DatesColl(Count))
                Tally_Tbl.ListColumns(2).DataBodyRange(Count, 1) = MonthName(Month(DatesColl(Count)))
            Next row
            
            Tally_Tbl.HeaderRowRange(Tally_Tbl.ListColumns(1).Range.Rows.Count, 1).value = "Total"  'label total KPI row
            
            With Tally_Tbl.Range
                .BorderAround xlContinuous, xlMedium, xlColorIndexAutomatic
            End With
            Count = 0
            
            col_count = 0
            For Each col In Tally_Tbl.HeaderRowRange
                col_count = col_count + 1
                For Each KPI In ArrDict.Keys
                    If col = KPI Then
                        Tally_Tbl.ListColumns(col_count).Range(1).Interior.Color = ColorColl(KPI)
                    End If
                Next
                
                If col = "[Empty]" Then
                    Tally_Tbl.ListColumns("[Empty]").Range(1).Interior.Color = RGB(234, 246, 148)
                End If
                
            Next
    
    Dim Mnth As Integer
    Dim DateVal As Date
    
    key = 0
    For Each monthyear In DatesColl
       key = key + 1
            'Create monthly table!!
            ' if excel table already exists, delete and replace with new one
               
            MyName = MonthName(Month(monthyear)) & "_" & Year(monthyear) & "_Table"      'name for monthly table
            
                DateVal = DateValue(MonthName(Month(monthyear)) & "/1/" & Year(monthyear))  'date value for first day of month
              
               Set Top_Left = Tally_Tbl.ListColumns(1).DataBodyRange(Tally_Tbl.ListColumns(1).DataBodyRange.Rows.Count, 1).Offset((11 * (key - 1)) + 3, 0)  'set top leftof table, offsetting by bottom of tally tale by 2 rows
               Set Bottom_Right = Top_Left.Offset(10, MonthDays(DateVal))
            
            MyRangeString = Top_Left.Address & ":" & Bottom_Right.Address   'make table address string
        
            MyListExists = False                                'assume table does not exist
            For Each ListObj In Sheets(SheetName).ListObjects
                If ListObj.Name = MyName Then MyListExists = True   'if table exists, true
            Next ListObj
            
            If Not (MyListExists) Then      'if table does not exist, create
                Set NewTable = ActiveSheet.ListObjects.Add(xlSrcRange, Range(MyRangeString), , xlYes)
                NewTable.Name = MyName
                Else                        'if exists, designate variable
                Set NewTable = ActiveSheet.ListObjects(MyName)
            End If
            
            NewTable.HeaderRowRange(1, 1).value = MonthName(Month(monthyear))   'put month name in table
            NewTable.HeaderRowRange(1, 1).Offset(0, -1) = CStr(Year(DateVal))
            
            For col = 2 To NewTable.Range.Columns.Count + 1     'label columns in monthly table
                NewTable.HeaderRowRange(1, col).value = col - 1
            Next col
            
            For Each cell In NewTable.ListColumns(1).DataBodyRange      'label frequency columns
                Count = Count + 1
                cell.value = Count
            Next cell
            Count = 0
                
                KDay = 0        'initialize NewTable column count at 0
                For Each Column In NewTable.ListColumns                     'for each column in month table
                    
                    If IsNumeric(Column) = False Then
                                'do nothing    'if not numeric, e.g. "January", skip
                                    Else
                                CRow = 0   'row counter for countermeasures table
                                KRow = 0    'row counterfor NewTable (monthly KPI)
                                KMonth = Month(monthyear)   'KPI month = month of current table
                                KYear = Year(monthyear)    'KPI year = year val of macro
                                KDay = KDay + 1      'KPI day = column header
                                
                                
                                    For Each cell In counter_tbl.ListColumns("Issue Date").DataBodyRange    'for each cell in countermeasures table "Issue Date" column
                                        
                                        CRow = CRow + 1     'Row count for countermeasures table
                                        
                                        CMonth = Month(cell.value)  'month of issue date
                                        CDay = Day(cell.value)      'day of issue date
                                        CYear = Year(cell.value)    'year of issue date
                                        
                                            'if issue date Month = NewTable Month, and issue day = column and isue year = year of chart
                                            ' and category or issue date = category value then
                                            'add KPI row + 1
                                            'add KPI value from countermeasures table to NewTable
                                        
                                        If CMonth = KMonth And CDay = KDay And CYear = KYear And counter_tbl.ListColumns("Category").DataBodyRange(CRow, 1).value = cat_val Then
                                                KRow = KRow + 1     'add NewTable row count
                                                
                                                If counter_tbl.ListColumns("KPI").DataBodyRange(CRow, 1).value <> Empty Then
                                                        
                                                        With NewTable.ListColumns(KDay + 1).DataBodyRange(KRow, 1)
                                                                .value = counter_tbl.ListColumns("KPI").DataBodyRange(CRow, 1).value     'add KPI value from countermeasures table to NewTable(KDay + 1 because in each month table there is one extra column for month name)
                                                                .Interior.Color = ColorColl(counter_tbl.ListColumns("KPI").DataBodyRange(CRow, 1).value)
                                                        End With
                                                        
                                                        Old_Tally = Tally_Tbl.ListColumns(counter_tbl.ListColumns("KPI").DataBodyRange(CRow, 1).value).DataBodyRange(key, 1).value     'old tally = old total in tally table data body range
                                                        New_Tally = Old_Tally + 1                                                                                                       'add one
                                                        Tally_Tbl.ListColumns(counter_tbl.ListColumns("KPI").DataBodyRange(CRow, 1).value).DataBodyRange(key, 1).value = New_Tally     'replace old tally with new tally in tally table
                                                        
                                                    Else
                                                        With NewTable.ListColumns(KDay + 1).DataBodyRange(KRow, 1)
                                                            .value = "[Empty]"     'add KPI value from countermeasures table to NewTable(KDay + 1 because in each month table there is one extra column for month name)
                                                            .Interior.Color = RGB(234, 246, 148)
                                                        End With
                                                        
                                                        Old_Tally = Tally_Tbl.ListColumns("[Empty]").DataBodyRange(key, 1).value     'old tally = old total in tally table data body range
                                                        New_Tally = Old_Tally + 1                                                                                                       'add one
                                                        Tally_Tbl.ListColumns("[Empty]").DataBodyRange(key, 1).value = New_Tally     'replace old tally with new tally in tally table
                                                        
                                                End If
                                        End If
                                    Next        'go to next cell in countermeasures table
                    End If
                    Next Column                 'go to next column (day) in month (newtable)
                    
                    NewTable.Range.WrapText = True
                
    Next monthyear                                   'go to next month in Calendar
    
        'add totals in TallyTable for each row ("Totals" column)
    Sum = 0
    For row = 1 To Tally_Tbl.DataBodyRange.Rows.Count - 1
        Sum = 0
        
        For col = 3 To Tally_Tbl.Range.Columns.Count - 2
            Sum = Sum + Tally_Tbl.DataBodyRange(row, col).value
        Next col
        
        Tally_Tbl.ListColumns("Total").DataBodyRange(row, 1).value = Sum
        
    Next row

        'add totals in TallyTable for each column ("Totals" row)
     Sum = 0
    For col = 3 To Tally_Tbl.Range.Columns.Count - 1
        Sum = 0
        
        For row = 1 To Tally_Tbl.DataBodyRange.Rows.Count - 1
            Sum = Sum + Tally_Tbl.DataBodyRange(row, col).value
        Next row
        
        Tally_Tbl.ListColumns(col).DataBodyRange(Tally_Tbl.DataBodyRange.Rows.Count, 1).value = Sum
        
    Next col
    
        'add running totals in TallyTable for last column ("Running Totals" col)
     RunningTotal = 0
    For row = 1 To Tally_Tbl.DataBodyRange.Rows.Count - 1
        
        If row = 1 Then
            RunningTotal = Tally_Tbl.ListColumns("Total").DataBodyRange(row, 1).value
            Else
            RunningTotal = RunningTotal + Tally_Tbl.ListColumns("Total").DataBodyRange(row, 1).value
        End If
        
        Tally_Tbl.ListColumns("Running Total").DataBodyRange(row, 1).value = RunningTotal
        
    Next row
    
        'format tally table
    Tally_Tbl.TableStyle = "TableStyleLight1"
    Tally_Tbl.DataBodyRange.HorizontalAlignment = xlCenter
    Tally_Tbl.ListColumns(1).DataBodyRange.HorizontalAlignment = xlLeft
    Tally_Tbl.Range.WrapText = True
    
    
        'create bar chart
    
    Set Tally_Top_Right = Tally_Tbl.Range(1, Tally_Tbl.Range.Columns.Count)
    
    Set BarTopLeft = Tally_Top_Right.Offset(0, 2)
    Set BarBottomRight = BarTopLeft.Offset(13, 8)
    
    Set BarChartRange = Range(BarTopLeft.Address & ":" & BarBottomRight.Address)
    
    Set SourceTopLeft = Tally_Tbl.ListColumns(1).Range.Cells(1)
    Set SourceBottomRight = Tally_Tbl.ListColumns(Tally_Tbl.Range.Columns.Count - 2).DataBodyRange(Tally_Tbl.DataBodyRange.Rows.Count - 1, 1)
    
    Set SourceDataRange = Range(SourceTopLeft.Address & ":" & SourceBottomRight.Address)

    Set BarChart = Worksheets(SheetName).Shapes.AddChart2(XlChartType:=201, _
                                            left:=BarTopLeft.left, Top:=BarTopLeft.Top, _
                                            Width:=BarChartRange.Width, Height:=BarChartRange.Height).Chart
        'format bar chart
        With BarChart
            .SetSourceData Source:=SourceDataRange
            .SeriesCollection(1).XValues = Range(Tally_Tbl.ListColumns(2).DataBodyRange(1, 1).Address & ":" & Tally_Tbl.ListColumns(1).DataBodyRange(Tally_Tbl.DataBodyRange.Rows.Count - 1, 1).Address) 'set x axis
            .SeriesCollection.NewSeries.values = Range(Tally_Tbl.ListColumns(Tally_Tbl.ListColumns.Count).DataBodyRange(1, 1).Address & ":" & Tally_Tbl.ListColumns(Tally_Tbl.ListColumns.Count).DataBodyRange(Tally_Tbl.DataBodyRange.Rows.Count - 1, 1).Address)
            .SeriesCollection(Tally_Tbl.ListColumns.Count - 3).XValues = Range(Tally_Tbl.ListColumns(2).DataBodyRange(1, 1).Address & ":" & Tally_Tbl.ListColumns(1).DataBodyRange(Tally_Tbl.DataBodyRange.Rows.Count - 1, 1).Address) 'set x axis
            
            '.SeriesCollection(2).HasDataLabels = True
            .SeriesCollection(Tally_Tbl.ListColumns.Count - 3).Name = "Running Total"
            .FullSeriesCollection(1).ChartType = xlColumnClustered
            .FullSeriesCollection(1).AxisGroup = 1
            .FullSeriesCollection(Tally_Tbl.ListColumns.Count - 3).ChartType = xlLine
            .FullSeriesCollection(Tally_Tbl.ListColumns.Count - 3).AxisGroup = 2
            .HasTitle = True
            If Month(enddateval) = 1 Then
            .ChartTitle.Text = "KPIs: " & CStr(Worksheets(SheetName).Range("B1").value) & " " & CStr(Worksheets(SheetName).Range("B1").value) & " " & CStr(MonthName(Month(startdateval))) & " " & CStr(Year(startdateval)) & " - " & CStr(MonthName(Month(12))) & " " & CStr(Year(enddateval) - 1)
           Else
            .ChartTitle.Text = "KPIs: " & CStr(Worksheets(SheetName).Range("B1").value) & " " & CStr(Worksheets(SheetName).Range("B1").value) & " " & CStr(MonthName(Month(startdateval))) & " " & CStr(Year(startdateval)) & " - " & CStr(MonthName(Month(enddateval))) & " " & CStr(Year(enddateval))
            End If
            .SetElement (msoElementLegendBottom)
            .SetElement msoElementPrimaryValueAxisShow
            .SetElement msoElementPrimaryValueAxisTitleHorizontal
            .Axes(xlValue).AxisTitle.Caption = "Totals"
            
            Count = 0
            For Each x In ColorColl
                Count = Count + 1
                .SeriesCollection(Count).Format.Line.ForeColor.RGB = RGB(0, 0, 0)   'outline bar with black line
                .SeriesCollection(Count).Format.Fill.ForeColor.RGB = x          'match bar with KPI color
            Next
                
            If .SeriesCollection.Count > ColorColl.Count Then
                .SeriesCollection(Count + 1).Format.Line.ForeColor.RGB = RGB(0, 0, 0)
               .SeriesCollection(Count + 1).Format.Fill.ForeColor.RGB = RGB(234, 246, 148) 'format "Empty"
            End If
        End With
        

With Worksheets(SheetName).Cells
    .EntireColumn.AutoFit
    .EntireRow.AutoFit
End With

End Sub


