Attribute VB_Name = "Calendar_M"
Sub CalendarMakerv(cat_val As String, MyInput As Long)

   Dim counter_tbl As ListObject
   Set counter_sheet = Worksheets("Countermeasures")
   Set counter_tbl = counter_sheet.ListObjects("Tbl_Counter")
    
      ' Set up error trapping.
    'On Error GoTo MyErrorTrap
    
       ' Prevent screen flashing while drawing calendar.
       Application.DisplayAlerts = False
       Application.ScreenUpdating = False
       ' Set up error trapping.
       
       ' make sheet name according to category value selected
        Dim SheetName As String
        SheetName = cat_val & " Calendar"
        
          ' Use InputBox to get desired year and set variable
       ' MyInput.
       ' Allow user to end macro with Cancel in InputBox.
       If MyInput = 0 Then Exit Sub
       
      ' Delete old calendar sheet
    For Each sheet In ThisWorkbook.Worksheets
    If sheet.Name = SheetName Then
        Sheets(SheetName).Delete
    End If
    Next sheet
    
        'creates new sheet for pivot table and chart
    Sheets.Add After:=ThisWorkbook.Worksheets("Control Center")
    ActiveSheet.Name = SheetName
    Set cal_sheet = ThisWorkbook.Worksheets(SheetName)
        'turns on alerts after initial deletion/creation
    Application.DisplayAlerts = True

       
       ' Get the date value of the beginning of inputted month.
      Dim mn As String
      Dim yr As Long

For k = 1 To 12
       'Get starting day for each month for each year, where MyInput is year as str
       StartDay = DateValue(k & " 1," & MyInput)
        yr = Year(StartDay)
        mn = MonthName(k)
       
       ' Prepare cell for Month and Year as fully spelled out.
        If k = 1 Then
            Set Mn_Name_Rng = cal_sheet.Range("A2:G9")
            Set first_cell = Mn_Name_Rng.Cells(1, 1)
        Else
            Set Mn_Name_Rng = cal_sheet.Range("A2:G9")
            'set range to new value based on number of rows for previous month
            Set Mn_Name_Rng = Mn_Name_Rng.Offset(Offset_row_val, 0)
        End If
       
       ' Center the Month and Year label across a2:g2 with appropriate
       ' size, height and bolding.
       With Mn_Name_Rng
           .HorizontalAlignment = xlCenterAcrossSelection
           .VerticalAlignment = xlCenter
           .Font.Size = 18
           .Font.Bold = True
           .RowHeight = 35
           .NumberFormat = "mmmm yyyy"
       End With
       ' Prepare a4:g4 for day of week labels with centering, size,
       ' height and bolding.
       If k = 1 Then
            Set Day_Name_Rng = cal_sheet.Range("A3:G3")
        Else
            Set Day_Name_Rng = cal_sheet.Range("A3:G3")
            'set range to new value based on number of rows for previous month
            Set Day_Name_Rng = Day_Name_Rng.Offset(Offset_row_val, 0)
       End If
       With Day_Name_Rng
           .ColumnWidth = 11
           .VerticalAlignment = xlCenter
           .HorizontalAlignment = xlCenter
           .VerticalAlignment = xlCenter
           .Orientation = xlHorizontal
           .Font.Size = 12
           .Font.Bold = True
           .RowHeight = 20
           .Interior.Color = RGB(0, 51, 103)
           .Font.Color = RGB(255, 255, 255)
           .NumberFormat = "mmmm yyyy"
       End With
       ' Put days of week in a3:g3.
       Day_Name_Rng.Cells(1, 1) = "Sunday"
       Day_Name_Rng.Cells(1, 2) = "Monday"
       Day_Name_Rng.Cells(1, 3) = "Tuesday"
       Day_Name_Rng.Cells(1, 4) = "Wednesday"
       Day_Name_Rng.Cells(1, 5) = "Thursday"
       Day_Name_Rng.Cells(1, 6) = "Friday"
       Day_Name_Rng.Cells(1, 7) = "Saturday"
       ' Prepare a4:g9 for dates with left/top alignment, size, height
       ' and bolding.
       If k = 1 Then
            Set Day_Num_Rng = cal_sheet.Range("A4:G9")
       Else
            Set Day_Num_Rng = cal_sheet.Range("A4:G9")
            Set Day_Num_Rng = Day_Num_Rng.Offset(Offset_row_val, 0)
       End If
       With Day_Num_Rng
           .HorizontalAlignment = xlRight
           .VerticalAlignment = xlTop
           .Font.Size = 18
           .Font.Bold = True
           .RowHeight = 21
           .NumberFormat = "####"
       End With
       
       ' Put inputted month and year fully spelling out into "a1".
         Mn_Name_Rng.Cells(1, 1).value = Application.Text(mn & MyInput, "mmmm yyyy")
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
       If k = 1 Then
            Set First_Day = cal_sheet.Range("A4:G4")
        Else
            Set First_Day = cal_sheet.Range("A4:G4")
            'set range to new value based on number of rows for previous month
           Set First_Day = First_Day.Offset(Offset_row_val, 0)
        End If
       
       Select Case DayofWeek
           Case 1
                First_Day.Cells(1, 1).value = 1
           Case 2
                First_Day.Cells(1, 2).value = 1
           Case 3
                First_Day.Cells(1, 3).value = 1
           Case 4
                First_Day.Cells(1, 4).value = 1
           Case 5
                First_Day.Cells(1, 5).value = 1
           Case 6
                First_Day.Cells(1, 6).value = 1
           Case 7
                First_Day.Cells(1, 7).value = 1
       End Select
       ' Loop through range a3:g8 incrementing each cell after the "1"
       ' cell.
       
       For Each cell In Day_Num_Rng
       If k = 0 Then
        Else
        End If
           RowCell = cell.row
           ColCell = cell.Column
           ' Do if "1" is in first column.
           If cell.Column = 1 And cell.row = (4 + Offset_row_val) Then
           ' Do if current cell is not in 1st column.
           ElseIf cell.Column <> 1 Then
               If cell.Offset(0, -1).value >= 1 Then
                   cell.value = cell.Offset(0, -1).value + 1
                   ' Stop when the last day of the month has been
                   ' entered.
                   If cell.value > (FinalDay - StartDay) Then
                       cell.value = ""
                       ' Exit loop when calendar has correct number of
                       ' days shown.
                       Exit For
                   End If
               End If
           ' Do only if current cell is not in Row 3 and is in Column 1.
           ElseIf cell.row > (4 + Offset_row_val) And cell.Column = 1 Then
               cell.value = cell.Offset(-1, 6).value + 1
               ' Stop when the last day of the month has been entered.
               If cell.value > (FinalDay - StartDay) Then
                   cell.value = ""
                   ' Exit loop when calendar has correct number of days
                   ' shown.
                   Exit For
               End If
           End If
       Next
       

        'Create Entry cells, format them centered, wrap text, and border
       ' around days.
       
       For x = 0 To 5
           Day_Num_Rng.Cells(2, 1).Offset(x * 2, 0).EntireRow.Insert
           With Day_Num_Rng.Rows(2).Offset(x * 2, 0)
               .RowHeight = 35
               .HorizontalAlignment = xlCenter
               .VerticalAlignment = xlTop
               .WrapText = True
               .Font.Size = 10
               .Font.Bold = False
               ' Unlock these cells to be able to enter text later after
               ' sheet is protected.
               .Locked = False
           End With
           ' Put border around the block of dates.
           With Day_Num_Rng.Cells(1, 1).Offset(x * 2, 0).Resize(2, _
           7).Borders(xlLeft)
               .Weight = xlThick
               .ColorIndex = xlAutomatic
           End With
           With Day_Num_Rng.Cells(1, 1).Offset(x * 2, 0).Resize(2, _
           7).Borders(xlRight)
               .Weight = xlThick
               .ColorIndex = xlAutomatic
           End With
           Day_Num_Rng.Cells(1, 1).Offset(x * 2, 0).Resize(2, 7).BorderAround _
              Weight:=xlThick, ColorIndex:=xlAutomatic
       Next x
       
        'set variables to the number of rows in the calendar entries (day num and space)
        ' set variables for the number of rows past cal into new cal
        ro = Day_Num_Rng.Rows.Count
        Ro1 = Day_Num_Rng.Rows.Count + 1
        
        'Delete extra empty rows (with borders) after calendar
        'If cell is in range and is blank, check to see if its a empty cell that correlates with a number,
        'if not delete
       If Day_Num_Rng.Cells(Day_Num_Rng.Rows.Count, 1) <> "" Then
            Set Day_Num_Rng = Day_Num_Rng.Cells(1, 1).Resize(Day_Num_Rng.Rows.Count + 1, 8)
                ' If calendar range last left cell is empty, delete until row above that has number for a date
            ElseIf Day_Num_Rng.Cells(Day_Num_Rng.Rows.Count, 1) = "" Then
                While Day_Num_Rng.Cells(Day_Num_Rng.Rows.Count, 1) = ""
                    Day_Num_Rng.Cells(Day_Num_Rng.Rows.Count, 1).EntireRow.Delete
                Wend
                        ' Add extra empty row past last number to calendar range, for fitting, following code will delete all empty
                        ' entries, leaving one row with blanks for the last row with number dates.
                        ' set range to include that extra empty row, so when other empty rows are deleted this one will remain
                    Set Day_Num_Rng = Day_Num_Rng.Cells(1, 1).Resize(Day_Num_Rng.Rows.Count + 1, 8)
        End If
       
       ' for each empty cell, plant in corresponding date value for
       ' comparison against countermeasures chart
        For Each box In Day_Num_Rng
        If box = "" Then
            'checks if box is below a day of the week, is part of the month name range, or displays a calendar number
            'which already has a date
                If right(box.Offset(-1, 0), 3) = "day" Or box.Offset(-1, 0) = "" Or IsDate(box.Offset(-1, 0)) = "True" Then
                Else
            box.value = DateValue(Month(StartDay) & "/" & box.Offset(-1, 0).value & "/" & yr)
            box.Font.Color = RGB(255, 255, 255)
            box.NumberFormat = "mm/dd/yyyy"
        End If
        End If
        Next
        
       ' Turn off gridlines.
        ActiveWindow.DisplayGridlines = False
          
        'Tally offset value as calendar gets longer, offset KPI row at top
        If k = 1 Then
         Offset_row_val = Day_Num_Rng.Rows.Count + 2
        Else
        Offset_row_val = Offset_row_val + Day_Num_Rng.Rows.Count + 2
        End If
        
'loop to next month
Next k

    'insert KPI row that displays category of cal and KPI of cat
    Range("B1") = cat_val
    Range("C1") = "KPIs:"
        With Range("B1:C1")
           .VerticalAlignment = xlCenter
           .Font.Size = 10
           .Font.Bold = False
           .RowHeight = 25
           .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThick
            .Borders.Color = RGB(0, 255, 0)
        End With
        
    'set variables to sheets and other list objects for further functions
    Set counter_sheet = Worksheets("Countermeasures")
    Set counter_tbl = counter_sheet.ListObjects("Tbl_Counter")
    Set KPI_rng = counter_tbl.ListColumns("KPI").DataBodyRange
    Dim KPI_Cal_Rng As Range

    'copy and paste KPI column in Countermeasures to top row
    i = 1
    For Each KPI In KPI_rng
        If KPI.Offset(0, -1).value = cat_val Then
            Range("C1").Offset(0, i) = KPI
            i = i + 1
        Else
        End If
    Next
    
    'Set variable to new row range that contains KPIs within Cal sheet,
    'going to be used to find unique values
    
    'Set KPI range to collection, call macro to get unique values range
    Set KPI_Cal_Rng = Range("D1", Range("D1").Offset(0, i))
    Dim uniques As Collection
    Set uniques = GetUniqueValues(KPI_Cal_Rng.value)
    
    'clear row where KPIs were housed
    Range("D1:AA1").Clear

    'Place unique KPIs in the first row (or the KPI range)
    Dim it
    i = 0
    For Each it In uniques
        Range("D1").Offset(0, i) = it
        With Range("D1").Offset(0, i)
           .VerticalAlignment = xlCenter
           .Font.Size = 9
           .Font.Bold = False
           .RowHeight = 25
           .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThick
            .Borders.Color = RGB(255, 0, 0)
            .WrapText = True
            
           End With
        
        i = i + 1
    Next
    
    
    'Freeze top row so KPIs are visible
    Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    
    'Set last cell of calendar range
    'Set calendar range
    Set last_cell = Day_Num_Rng.Cells(Day_Num_Rng.Rows.Count, Day_Num_Rng.Columns.Count - 1)
      
      'Set calendar range as first month name
      'cell to dec empty cell
      Dim cal_range As Range
      Set cal_range = Range(first_cell, last_cell)
     
        'Use check counter to identify and redden
        'cells where KPI were missed
       
       'Dim date_rng As Variant
   Set date_rng = counter_tbl.ListColumns("Issue Date").DataBodyRange
        'set countermeasures category range as var
   Dim cat_rng As Variant
   Set cat_rng = counter_tbl.ListColumns("Category").DataBodyRange
       
       ' for each cell in calendar, checks if "empty" cell (white cell accompanying
        ' date numeral) and if cell is blank and matches with a date in countermeasures
        ' with the same category, then the cell is marked red
     For Each caldat In cal_range
        If IsDate(caldat) = "True" Then
        If right(caldat.Offset(1, 0), 3) <> "day" Then
            For Each datum In date_rng
                i = 1
                If datum = caldat.value And datum.Offset(0, 1).value = cat_val Then
                    With caldat
                        .Interior.Color = RGB(255, 0, 0)
                        .Font.Color = RGB(255, 0, 0)
                        End With
                i = i + 1
                Else
                End If
            Next
        Else: End If
    Else
    End If
    Next
   
   Check_Counter_for_Cal cal_range
    
    AddButtonandCode2
       
       
       ' Resize window to show all of calendar (may have to be adjusted
       ' for video configuration).
       ActiveWindow.WindowState = xlMaximized
       ActiveWindow.ScrollRow = 1
       Exit Sub
       ' Allow screen to redraw with calendar showing.
       Application.ScreenUpdating = True
       ' Prevent going to error trap unless error found by exiting Sub
       ' here.
      ' Error causes msgbox to indicate the problem, provides new input box,
   ' and resumes at the line that caused the error.
'MyErrorTrap:
       'MsgBox "You may not have entered your Year correctly." _
           '& Chr(13) & "Please input 4 digits for the Year"
       'MyInput = InputBox("Type in year for Calendar")
       'If MyInput = "" Then Exit Sub
       'Resume
      
   End Sub
   
   
   
   
   Sub Check_Counter_for_Cal(cal_range As Range)
   
   
   'set list objects
   Dim counter_tbl As ListObject
   Set counter_sheet = Worksheets("Countermeasures")
   Set counter_tbl = counter_sheet.ListObjects("Tbl_Counter")
   'set countermeasures issue date range as var
   Dim date_rng As Variant
   Set date_rng = counter_tbl.ListColumns("Issue Date").DataBodyRange
   'set countermeasures category range as var
   Dim cat_rng As Variant
   Set cat_rng = counter_tbl.ListColumns("Category").DataBodyRange
   'set drop down value as var
   cat_val = ActiveSheet.Range("B1").value
  
   
   
   ' for each cell in calendar, checks if "empty" cell (white cell accompanying
   ' date numeral) and if cell is blank and matches with a date in countermeasures
   ' with the same category, then the cell is marked red
    For Each caldat In cal_range
    If caldat < Date And IsDate(caldat) = "True" And right(caldat.Offset(1, 0), 3) <> "day" Then
             i = 0
            For Each datum In date_rng
                If datum = caldat And datum.Offset(0, 1).value = cat_val Then
                        i = 1
                    Else
                End If
            Next
                If i < 1 Then
                With caldat
                        .Interior.Color = RGB(33, 101, 31)
                        .Font.Color = RGB(33, 101, 31)
                        End With
                Else
                With caldat
                        .Interior.Color = RGB(255, 0, 0)
                        .Font.Color = RGB(255, 0, 0)
                        End With
                End If
    End If
    Next
   

  End Sub
   Sub AddButtonAndCode()
     ' Declare variables
    Dim i As Long, Hght As Long
    Dim Name As String, NName As String
     ' Set the button properties
    i = 0
    Hght = 305.25
     ' Set the name for the button
    NName = "Update_Cal" & i
     ' Test if there is a button already and if so, increment its name
    For Each OLEObject In ActiveSheet.OLEObjects
        If left(OLEObject.Name, 9) = "Update_Cal" Then
            Name = right(OLEObject.Name, Len(OLEObject.Name) - 9)
            If Name >= i Then
                i = Name + 1
            End If
            NName = "Update_Cal" & i
            Hght = Hght + 27
        End If
    Next
     ' Add button
    Dim myCmdObj As OLEObject, N%
    Set myCmdObj = ActiveSheet.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
    Link:=False, DisplayAsIcon:=False, left:=0, Top:=0, _
    Width:=60, Height:=26.25)
     ' Define buttons name
    myCmdObj.Name = NName
     ' Define buttons caption
    myCmdObj.Object.Caption = "Update"
     ' Inserts code for the button
    With ThisWorkbook.VBProject.VBComponents(ActiveSheet.CodeName).CodeModule
        N = .CountOfLines
        .insertlines N + 1, "Private Sub " & NName & "_Click()"
        .insertlines N + 2, vbNewLine
        .insertlines N + 3, vbTab & "Dim cal_range as Range"
        .insertlines N + 4, vbTab & "Set cal_range = ActiveSheet.UsedRange"
        .insertlines N + 5, vbTab & "Call Calendar_M.Check_Counter_for_Cal(cal_range)"
        .insertlines N + 6, vbNewLine
        .insertlines N + 7, "End Sub"
    End With
End Sub

Sub AddButtonandCode2()

    Dim i As Long, Hght As Long
    Dim Name As String, NName As String
     ' Set the button properties
    i = 0
    Hght = 305.25
     ' Set the name for the button
    NName = "Update"
    
    Dim myCmdObj As Button
    Set myCmdObj = ActiveSheet.Buttons.Add(0, 0, _
   60, 26.25)
    
    With myCmdObj
    .OnAction = "Update_Cal_Button"
    .Caption = NName
    End With
    
End Sub

 Sub Update_Cal_Button()

    Dim cal_range As Range
    Set cal_range = ActiveSheet.UsedRange
    Call Calendar_M.Check_Counter_for_Cal(cal_range)

End Sub

Sub ShowCalForm()
Calendar_Frm.Show vbModeless
End Sub

