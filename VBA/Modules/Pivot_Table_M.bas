Attribute VB_Name = "Pivot_Table_M"
Sub Create_Pivot_Trend(cat_val As String, issue_val As String, filterval As Long)

Dim trend_sheet As Worksheet 'variable for worksheet for general trends
Dim counter_sheet As Worksheet 'variable for worksheet with trend data
Dim trend_cache As PivotCache 'variable for trend data as a pivot cache
Dim trend_table As PivotTable 'variable for trend pivot table
Dim counter_range As Range 'variable for the range of the data of the trned pivot table being the countermeasures trend data
Dim lastrow As Long 'last row of countermeasures trend data
Dim lastcol As Long 'last col of countermeasures trend data
Dim counter_tbl As ListObject 'excel table of countermeasure data found within name manager
Dim pivot_item As PivotItem 'variable for assessing if pivotitems = blank
Dim issue_field As PivotField ' variable for assessing pivot field from drop down box

'On Error GoTo ErrorHandler 'if error, displays error message


    'turns off "are you sure you want to delete (previous worksheet and chart??"
    'turns off excel blink
Application.DisplayAlerts = False
Application.ScreenUpdating = False


    'deletes previous pivot table and chart, if it exists
For Each sheet In ThisWorkbook.Worksheets
    If sheet.Name = "Trend Table" Then
        Sheets("Trend Table").Delete
    End If
Next sheet
    
    'creates new sheet for pivot table and chart
Sheets.Add After:=ThisWorkbook.Worksheets("Control Center")
ActiveSheet.Name = "Trend Table"

    'turns on alerts after initial deletion/creation
Application.DisplayAlerts = True

    'sets variables and names to worksheets
    'sets variable and name to countermeasures table
Set trend_sheet = Worksheets("Trend Table")
Set counter_sheet = Worksheets("Countermeasures")
Set counter_tbl = counter_sheet.ListObjects("Tbl_Counter")


    'Sets pivot cache to countermeasures table
Set trend_cache = ThisWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:="Tbl_Counter")

    'creates the blank pivot table in the top left cell of spreadsheet
Set trend_table = trend_cache.CreatePivotTable _
(TableDestination:=trend_sheet.Cells(1, 1), _
TableName:="TrendPivotTable")


    'Inputs pivot table fields into pivot table,
    'Logic to determine which category fields should be shown or not

    
With trend_table.PivotFields("Category")
    .Orientation = xlPageField
    .Position = 1
End With
    
   
    ActiveSheet.PivotTables("TrendPivotTable").RefreshTable
    For Each pivot_item In trend_table.PivotFields("Category").PivotItems
            If pivot_item.Name = "(blank)" Then
                pivot_item.Visible = False
                ElseIf pivot_item.Name = cat_val Then
                pivot_item.Visible = True
                Else
                pivot_item.Visible = False
            End If
        Next
   

With trend_table.PivotFields("Issue Year")
    .Orientation = xlRowField
    .Position = 1
    End With
    
     For Each pivot_item In trend_table.PivotFields("Issue Year").PivotItems
            If pivot_item.Name = "(blank)" Then
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
            If pivot_item.Name = "(blank)" Then
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

Set issue_field = trend_table.PivotFields(issue_val)
    
With issue_field
  .Orientation = xlColumnField
   .Position = 1
   For Each pivot_item In issue_field.PivotItems
        If pivot_item.Name = "(blank)" Then
            pivot_item.Visible = False
        End If
    Next
   End With
   
   
    For Each pivot_item In issue_field.PivotItems
            If pivot_item.Name = "(blank)" Then
                pivot_item.Visible = False
                Else
                pivot_item.Visible = True
            End If
    Next
    
    
    'Adds value filter, found on "Create Pivot Table" spreadsheet, to look for issues with frequency greater than or equal to drop down box value
With ThisWorkbook.Worksheets("Trend Table").PivotTables("TrendPivotTable").PivotFields("Issue Tier 2 Tag")
            .PivotFilters.Add2 Type:=xlValueIsGreaterThanOrEqualTo, DataField:=ThisWorkbook.Worksheets("Trend Table").PivotTables("TrendPivotTable").PivotFields("Count of Issues"), _
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
Set trend_range = Worksheets("Trend Table").PivotTables("TrendPivotTable").TableRange1.Cells(1, 1).Resize(lastrow2, lastcol2) 'table 1 range (data range) of pivot table
TopRow = Worksheets("Trend Table").PivotTables("TrendPivotTable").TableRange1.row 'top row of pivot table (data range)
TopCol = Worksheets("Trend Table").PivotTables("TrendPivotTable").TableRange1.Column 'top column of pivot table (table range)

Set PT_First_Cell = Cells(TopRow, TopCol) 'first cell of pivot table range (data range)
Set Chart_Add = PT_First_Cell.Offset(rowOffset:=trend_table.TableRange1.Rows.Count, ColumnOffset:=0) 'takes first cell and adds # of rows in pivot table for first cell of chart placement

Debug.Print trend_range.Address

trend_range.Select

With Worksheets("Trend Table").Shapes.AddChart2(297, xlColumnStacked) '297
    .Name = "Trend Chart"
    .Chart.SetSourceData Source:=trend_range
    .Chart.HasTitle = True
    .Chart.ChartTitle.Text = ActiveSheet.Name
    .Chart.Axes(xlValue, xlPrimary).HasTitle = True
    .Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Frequency"
    .Chart.Axes(xlValue, xlPrimary).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 12
    .Chart.Axes(xlCategory, xlPrimary).HasTitle = True
    .Chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time"
    .Chart.Axes(xlCategory, xlPrimary).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 12
    .left = Chart_Add.left
    .Top = Chart_Add.Top
    .Height = 250
    .Width = 1000
    .Chart.Legend.Position = xlLegendPositionCorner
    .Chart.Legend.Top = 7
    .Chart.Legend.Height = 200
End With


''''''''''''On Error:'''''''''''''''''''''''''''''
Exit Sub

ErrorHandler:

MsgBox ("Please ensure:" & vbCrLf & " - each pivot table drop down box on the " & Chr(34) & "Control Center" & Chr(34) & " sheet contains a selection" & vbCrLf & _
         "- The selected category and trend exist within the Countermeasures table" & vbCrLf & _
         "- the Countermeasure table has a " & Chr(34) & "Category," & Chr(34) & " " & Chr(34) & "Issue Date," & Chr(34) & " " & Chr(34) & "Issue Month," & Chr(34) & "and " & Chr(34) & " " & Chr(34) & "Issue Year," & Chr(34) & " column" & vbCrLf & _
         "- The Countermeasure table's name is " & Chr(34) & "Tbl_Counter" & Chr(34) & vbCrLf & _
         "- The Countermeasure table has relavant columns non-empty")
 
End Sub

Sub Create_Pivot_Running_Total(cat_val As String, issue_val As String, filterval As Long)

Dim running_sheet As Worksheet 'variable for worksheet for general trends
Dim counter_sheet As Worksheet 'variable for worksheet with trend data
Dim running_cache As PivotCache 'variable for trend data as a pivot cache
Dim running_table As PivotTable 'variable for trend pivot table
Dim counter_range As Range 'variable for the range of the data of the trned pivot table being the countermeasures trend data
Dim lastrow As Long 'last row of countermeasures trend data
Dim lastcol As Long 'last col of countermeasures trend data
Dim counter_tbl As ListObject 'excel table of countermeasure data found within name manager
Dim pivot_item As PivotItem 'variable for assessing if pivotitems = blank
Dim issue_field As PivotField ' variable for assessing pivot field from drop down box

'On Error GoTo ErrorHandler 'if error, displays error message


    'turns off "are you sure you want to delete (previous worksheet and chart??"
    'turns off excel blink
Application.DisplayAlerts = False
Application.ScreenUpdating = False


    'deletes previous pivot table and chart, if it exists
For Each sheet In ThisWorkbook.Worksheets
    If sheet.Name = "Running Total Table" Then
        Sheets("Running Total Table").Delete
    End If
Next sheet
    
    'creates new sheet for pivot table and chart
Sheets.Add After:=ThisWorkbook.Worksheets("Control Center")
ActiveSheet.Name = "Running Total Table"

    'turns on alerts after initial deletion/creation
Application.DisplayAlerts = True

    'sets variables and names to worksheets
    'sets variable and name to countermeasures table
Set running_sheet = Worksheets("Running Total Table")
Set counter_sheet = Worksheets("Countermeasures")
Set counter_tbl = counter_sheet.ListObjects("Tbl_Counter")


    'Sets pivot cache to countermeasures table
Set running_cache = ThisWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:="Tbl_Counter")

    'creates the blank pivot table in the top left cell of spreadsheet
Set running_table = running_cache.CreatePivotTable _
(TableDestination:=running_sheet.Cells(1, 1), _
TableName:="RunningPivotTable")

    'sets variables to value within drop down boxes on "Create Pivot Table" spreadsheet

    'Inputs pivot table fields into pivot table,
    'Logic to determine which category fields should be shown or not
With running_table.PivotFields("Category")
    .Orientation = xlPageField
    .Position = 1
    End With
    
     ActiveSheet.PivotTables("RunningPivotTable").RefreshTable
    For Each pivot_item In running_table.PivotFields("Category").PivotItems
            If pivot_item.Name = "(blank)" Then
                pivot_item.Visible = False
                ElseIf pivot_item.Name = cat_val Then
                pivot_item.Visible = True
                Else
                pivot_item.Visible = False
            End If
        Next
    

With running_table.PivotFields("Yr-Month")
    .Orientation = xlRowField
    .Position = 1
    End With
    
    For Each pivot_item In running_table.PivotFields("Yr-Month").PivotItems
            If pivot_item.Name = "(blank)" Then
                pivot_item.Visible = False
                Else
                pivot_item.Visible = True
            End If
    Next
    
With running_table.PivotFields("Issue Year")
    .Orientation = xlRowField
    .Position = 1
    End With
    
     For Each pivot_item In running_table.PivotFields("Issue Year").PivotItems
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
    .BaseField = "Yr-Month"
    End With
    
 For Each pivot_item In running_table.PivotFields("Issue Date").PivotItems
            If pivot_item.Name = "(blank)" Then
                pivot_item.Visible = False
                Else
                pivot_item.Visible = True
            End If
    Next
    
Set issue_field = running_table.PivotFields(issue_val)
   
With issue_field
    .Orientation = xlColumnField
   .Position = 1
    For Each pivot_item In issue_field.PivotItems
        If pivot_item.Name = "(blank)" Then
            pivot_item.Visible = False
        End If
    Next
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
    
    
    'Adds value filter, found on "Create Pivot Table" spreadsheet, to look for issues with frequency greater than or equal to drop down box value
With ThisWorkbook.Worksheets("Running Total Table").PivotTables("RunningPivotTable").PivotFields("Issue Tier 2 Tag")
            .PivotFilters.Add2 Type:=xlValueIsGreaterThanOrEqualTo, DataField:=ThisWorkbook.Worksheets("Running Total Table").PivotTables("RunningPivotTable").PivotFields("Count of Issues"), _
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
Set running_range = Worksheets("Running Total Table").PivotTables("RunningPivotTable").TableRange1.Cells(1, 1).Resize(lastrow2, lastcol2) 'table 1 range (data range) of pivot table
TopRow = Worksheets("Running Total Table").PivotTables("RunningPivotTable").TableRange1.row 'top row of pivot table (data range)
TopCol = Worksheets("Running Total Table").PivotTables("RunningPivotTable").TableRange1.Column 'top column of pivot table (table range)

Set PT_First_Cell = Cells(TopRow, TopCol) 'first cell of pivot table range (data range)
Set Chart_Add = PT_First_Cell.Offset(rowOffset:=running_table.TableRange1.Rows.Count, ColumnOffset:=0) 'takes first cell and adds # of rows in pivot table for first cell of chart placement

Debug.Print running_range.Address

running_range.Select

With Worksheets("Running Total Table").Shapes.AddChart2(227, xlLine)
    .Name = "Running Total Chart"
    .Chart.SetSourceData Source:=running_range
    .Chart.HasTitle = True
    .Chart.ChartTitle.Text = ActiveSheet.Name
    .Chart.Axes(xlValue, xlPrimary).HasTitle = True
    .Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Frequency"
    .Chart.Axes(xlValue, xlPrimary).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 12
    .Chart.Axes(xlCategory, xlPrimary).HasTitle = True
    .Chart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Time"
    .Chart.Axes(xlCategory, xlPrimary).AxisTitle.Format.TextFrame2.TextRange.Font.Size = 12
    .left = Chart_Add.left
    .Top = Chart_Add.Top
    .Height = 250
    .Width = 1000
    .Chart.Legend.Position = xlLegendPositionCorner
    .Chart.Legend.Top = 7
    .Chart.Legend.Height = 200
End With


''''''''''''On Error:'''''''''''''''''''''''''''''
Exit Sub

'ErrorHandler:

'MsgBox ("Please ensure:" & vbCrLf & " - each pivot table drop down box on the " & Chr(34) & "Control Center" & Chr(34) & " sheet contains a selection" & vbCrLf & _
 '        "- The selected category and trend exist within the Countermeasures table" & vbCrLf & _
  '       "- the Countermeasure table has a " & Chr(34) & "Category," & Chr(34) & " " & Chr(34) & "Issue Date," & Chr(34) & " and " & Chr(34) & "Yr-Month," & Chr(34) & "column" & vbCrLf & _
   '      "- The Countermeasure table's name is " & Chr(34) & "Tbl_Counter" & Chr(34) & vbCrLf & _
    '     "- The Countermeasure table has relavant columns non-empty")
 

End Sub

Sub ShowPivotTableFrm()
PivotTable_Frm.Show vbModeless
End Sub


Sub FixLegends()
  Dim C As Chart
  Dim i As Long, Max As Long
  Dim Red As Integer, Blue As Integer
  'Get the chart from the sheet
  Set C = ActiveSheet.ChartObjects(1).Chart
  With C.Legend
    Max = .LegendEntries.Count
    For i = 1 To .LegendEntries.Count
      Red = (255 / (Max - 1)) * (Max - i)
      Blue = (255 / (Max - 1)) * (i - 1)
      .LegendEntries(i).LegendKey.Interior.Color = RGB(Red, 0, Blue)
    Next
  End With
End Sub


