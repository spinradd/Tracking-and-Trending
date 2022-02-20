Attribute VB_Name = "Identifier_Updater_M"
Sub OpenTandDForm()
TandDChoiceForm.Show vbModeless
End Sub

Sub CreateTagDescriptorSheet(start_date As Date, end_date As Date)
'creates Tag and descriptor sheet based off of specified time interval for data

Dim wb As Workbook
Dim SheetNm As String
Dim buttonname As String
Dim buttoncaption As String
Dim buttonheight As Double
Dim buttonwidth As Double
Dim buttontop As Double
Dim buttonleft As Double
Dim subtitle As String
Dim subprocedure As String
Dim Obj As OLEObject
Dim tbl As ListObject


SheetNm = "Tag and Descriptor Tables"

Application.DisplayAlerts = False
Application.ScreenUpdating = False


    For Each sheet In ThisWorkbook.Worksheets   'if sheet already exists, delete it
    If sheet.Name = SheetNm Then
        Sheets(SheetNm).Delete
    End If
    Next sheet

                                                'creates new sheet
    Sheets.Add After:=ThisWorkbook.Worksheets("Control Center")
    ActiveSheet.Name = SheetNm
    Set Tnd_sheet = ThisWorkbook.Worksheets(SheetNm)
        'turns on alerts after initial deletion/creation

    'create each individual table within sheet (creates color Key too)
Call MainTagTableUpdate(start_date, end_date)
Call TagsbyCatTblUpdate(start_date, end_date)
Call Tbl_CategoriesUpdate(start_date, end_date)
Call Tbl_KPIUpdate(start_date, end_date)
Call Tbl_IdentifierUpdate(start_date, end_date)


Sheets(SheetNm).Columns("A:B").ColumnWidth = 16    'increase column size for a larger button
Sheets(SheetNm).Rows(1).RowHeight = 39.5           'increase row size for button placement
Sheets(SheetNm).Rows(5).RowHeight = 39.5            'increase row size for tag finder button

        'set variables for reset button specifications
buttonname = "ResetButton"
buttoncaption = "Reset Sheet"
buttonwidth = Sheets(SheetNm).Range("A1:B1").Width
buttonheight = Sheets(SheetNm).Range("A1:B1").Height
buttonleft = Sheets(SheetNm).Columns("A").left
buttontop = 0

subtitle = buttonname & "_click()"
subprocedure = "OpenTandDForm"
            
            'function creates a basic button with assigned variables (and macros)
AddButtonandCode2 SheetNm, buttonname, buttoncaption, buttonleft, buttontop, buttonwidth, buttonheight, subprocedure, subtitle

                'creates tag tester feature
Call Identifier_Updater_M.CreateTagTester
Call Identifier_Updater_M.CreateTagTestwithCodeButton
    
 Application.DisplayAlerts = True
 Application.ScreenUpdating = False
End Sub
Sub AddButtonandCode2(SheetName As String, buttonname As String, buttoncaption As String, buttonleft As Double, buttontop As Double, buttonwidth As Double, buttonheight As Double, subprocedure As String, subtitle As String)
    'sub to create a simple button with the name of the subprocedure the button should do upon activation
    
    Dim myCmdObj As Button
    Set myCmdObj = ActiveSheet.Buttons.Add(buttonleft, buttontop, buttonwidth, buttonheight)
    
    With myCmdObj
    .OnAction = subprocedure
    .Caption = buttonname
    End With
    
End Sub

Sub MainTagTableUpdate(start_date As Date, end_date As Date)
'creates first main table in sheet depicting all the tag columns and all values within them

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
Dim Tags() As Variant
Dim counter_tbl As ListObject

Dim ArrBase() As Variant
Dim col_val As String

    
        'set variables for countermeasures sheet, table, and columns
   Set counter_sheet = Worksheets("Countermeasures")
   Set counter_tbl = counter_sheet.ListObjects("Tbl_Counter")
   Set counter_col = counter_tbl.ListColumns
            
            'identify issue ID column and Issue Date column. All columns in between are Tag columns
For Each cell In counter_tbl.HeaderRowRange
    If cell.value = "Issue ID" Then
    IssueIDVal = cell.Column
    End If
    If cell.value = "Issue Date" Then
    Category_ColVal = cell.Column
    End If
Next cell


entry_count = 0
For Each cell2 In counter_tbl.HeaderRowRange                                    'if column exists between IssueID and Issue Date, add column title to array
    If cell2.Column > IssueIDVal And cell2.Column < Category_ColVal Then
    
    If entry_count = 0 Then
        ReDim Preserve Tags(0 To 0)
        Tags(0) = cell2
     'For all subsequent entries extend array by 1 and enter contents in cell
    Else
        ReDim Preserve Tags(UBound(Tags) + 1)
        Tags(UBound(Tags)) = cell2
    End If
    entry_count = entry_count + 1
    End If
Next cell2

'MsgBox Join(Tags, vbCrLf)
Debug.Print LBound(Tags)

If (Not Not Tags) = 0 Then       'if array never intitialized then it doesn't exist,
        IsTagsEmpty = True         'if list doesn't exist, then it's empty (by default)
    Else
        IsTagsEmpty = True         'if list does exist, assume it is empty
        For Each item_in_array In Tags       'test if array is empty. If there is one non-blank cell, change bool value
            If item_in_array <> Empty Then
                IsTagsEmpty = False
            End If
        Next
End If
        
If IsTagsEmpty = False Then            'if array isn't empty, remove blanks, remove dups
        Tags = BlankRemover(Tags)     'if not empty, remove blanks
        Tags = ArrayRemoveDups(Tags)
    ElseIf IsTagsEmpty = True Then
        Exit Sub                        'if array is empty,  no tags exist, exit sub
End If

                                            ' if excel table already exists, delete and replace with new one
   SheetName = "Tag and Descriptor Tables"
   MyName = "MainTagTable"
   
   Set Top_Left = Worksheets(SheetName).Range("D7")     'initialize top left cell for table start
   Set Bottom_Right = Top_Left.Offset(1, ((UBound(Tags) + 1) * 2) - 1)  'specify last cell of table (= 2 * number of tag columns)
   
   MyRangeString = Top_Left.Address & ":" & Bottom_Right.Address    'place addresses into a string
   
   
  MyListExists = False                                  'assume table does not exist
    For Each ListObj In Sheets(SheetName).ListObjects   'check if table exists (redundant) new sheet is deleted every time
        If ListObj.Name = MyName Then
            ListObj.Range.ClearFormats
            ListObj.Range.ClearContents
            MyListExists = False
        End If
    Next ListObj
    
    If Not (MyListExists) Then                          'if table doesn't exist, make a new one
        Set MainTagTbl = ActiveSheet.ListObjects.Add(xlSrcRange, Range(MyRangeString), , xlYes)
        MainTagTbl.Name = MyName
    End If

For x = 1 To MainTagTbl.Range.Columns.Count     'loop assigns the column titles for table
    
    If x = 1 Then                               'if first column, assign first tag
        MainTagTbl.HeaderRowRange(1, x) = Tags(0)
    ElseIf x Mod 2 = 0 Then                     'if every other column, make column header the previous column + "- count"
                                                'will contain the frequency of the corresponding tag form the left column
        MainTagTbl.HeaderRowRange(1, x) = MainTagTbl.HeaderRowRange(1, x - 1).value & " Count"
    Else                                        'if not first and not every other column, assign column header to appropriate tag
        Tag_Multiplier = Application.WorksheetFunction.RoundUp(x / 2, 0)
        MainTagTbl.HeaderRowRange(1, x) = Tags(Tag_Multiplier - 1)
    
    
    End If
Next

Set tagtitlerange = Range(MainTagTbl.HeaderRowRange(1, 1).Offset(-1, 0), MainTagTbl.HeaderRowRange(1, MainTagTbl.Range.Columns.Count).Offset(-1, 0))
        If tagtitlerange.MergeCells Then
            tagtitlerange.Cells.UnMerge
        End If

    StartDay = start_date
    ' Check if valid date but not the first of the month
    ' -- if so, reset StartDay to first day of month.
    If Day(StartDay) <> 1 Then
        StartDay = DateValue(Month(StartDay) & "/1/" & _
            Year(StartDay))
    End If
     
    FinalDay = DateValue(end_date)

Column = Column + 1
For Each tblcell In MainTagTbl.HeaderRowRange    'For each column           '
        tblcell.Offset(-1, 0) = "All Tags"      'cell above it = "All Tags" (to merge later)
    If InStr(1, tblcell.value, "Count") Then    'if cell is a "count" column, ignore, loop to next column
    Column = Column + 1
    Else                                        'if cell is not a count column, continue:
    entry_count = 0
    counter_row_count = 0
                For Each Cell3 In counter_tbl.ListColumns(tblcell.value).DataBodyRange      'for each cell in corresponding column in countermeasure table, if row (entry) has issue date
                                                                                            'between start day and final day, then add the column's tag to an array
                    counter_row_count = counter_row_count + 1
                         If counter_tbl.ListColumns("Issue Date").DataBodyRange.Cells(counter_row_count, 1).value >= StartDay _
                            And counter_tbl.ListColumns("Issue Date").DataBodyRange.Cells(counter_row_count, 1).value < FinalDay Then
                                       'if first entry, redim to hold one spot "(0)"
                                       If entry_count = 0 Then
                                           ReDim Preserve ArrBase(0)
                                           ArrBase(0) = Cell3
                                        'For all subsequent entries extend array by 1 and enter contents in cell
                                       Else
                                           ReDim Preserve ArrBase(UBound(ArrBase) + 1)
                                           ArrBase(UBound(ArrBase)) = Cell3
                                       End If
                                       entry_count = entry_count + 1
                        End If
                Next Cell3
                
                'MsgBox Join(ArrBase, vbCrLf)
                
            If (Not Not ArrBase) = 0 Then       'if array never intitialized then it doesn't exist,
                    IsTagsEmpty = True         'if list doesn't exist, then it's empty (by default)
                Else
                    IsTagsEmpty = True         'if list does exist, assume it is empty
                    For Each item_in_array In ArrBase       'test if array is empty. If there is one non-blank cell, change bool value
                        If item_in_array <> Empty Then
                            IsTagsEmpty = False
                        End If
                    Next
            End If
                    
            If IsTagsEmpty = False Then            'if array isn't empty, remove blanks, remove dups
                    'ArrBase = BlankRemover(ArrBase)     'if not empty, remove blanks
                    'ArrBase = ArrayRemoveDups(ArrBase)  'if not empty, remove duplicates
                ElseIf IsTagsEmpty = True Then GoTo FollowingColumn   'if array is empty,  go to next column
            End If
                
    Dim ArrDict As Scripting.Dictionary
    Set ArrDict = New Scripting.Dictionary
    
    'rids duplicates and counts new dictionary
    Set ArrDict = DuplicateCountToScript(ArrBase)
    
    Call Functions_M.SortDictionary(ArrDict, False, True, vbBinaryCompare)  'sorts dictionary by frequency of tag
    
    MainTagTbl.DataBodyRange.Cells(1, Column).Resize(ArrDict.Count, 1).Value2 = Application.Transpose(ArrDict.Keys) 'paste tags in the left column
    MainTagTbl.DataBodyRange.Cells(1, Column + 1).Resize(ArrDict.Count, 1).Value2 = Application.Transpose(ArrDict.Items)    'paste frequency in corresponding right column
    
    If Column = 1 Then              'if first column, then resize column to the size of the array
    resize_num = ArrDict.Count + 1
    Else                            'if not first column see if array is bigger than current row count
        If ArrDict.Count > resize_num Then
            resize_num = ArrDict.Count + 1
        End If
    End If
    
    With MainTagTbl.Range           'resize range to current or greater row count
    MainTagTbl.Resize .Resize(resize_num)
    End With
    
FollowingColumn:
Column = Column + 1
End If
    
    Erase ArrBase       'erase arrbase for new column
Next tblcell
'Formatting
    With MainTagTbl.Range
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
            
       For Each iCells In MainTagTbl.Range
            iCells.BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThin
        Next iCells
        
        MainTagTbl.TableStyle = "TableStyleMedium9"
        
    End With
    
    For x = 1 To MainTagTbl.Range.Columns.Count                 'formatting colored cells
        If x Mod 2 = 0 Then                                     'if count column, then
            For Each num In MainTagTbl.ListColumns(x).DataBodyRange     'for each cell in data body range of cell change
                                                                        'color based on count
                
                If num.value >= 0 And num.value <= 3 And num <> Empty Then
                num.Interior.Color = RGB(84, 130, 53)
                num.Font.Color = RGB(255, 255, 255)
                ElseIf num.value >= 4 And num.value <= 6 Then
                num.Interior.Color = RGB(255, 204, 102)
                num.Font.Color = RGB(0, 0, 0)
                ElseIf num.value >= 7 Then
                num.Interior.Color = RGB(255, 124, 128)
                num.Font.Color = RGB(0, 0, 0)
                End If
                
                If num = Empty Then                                     'if empty cell, input "blank"
                    If num = Empty And num.Offset(0, -1) <> Empty Then
                    num.value = "(blank)"
                    Else
                    num.Interior.Color = RGB(255, 255, 255)
                    End If
                End If
                
            Next num
        
        Else                                                             'if a tag column (not count)
            
            For Each num In MainTagTbl.ListColumns(x).DataBodyRange
                
                For Each box In MainTagTbl.DataBodyRange                    'check each box in table to see if it matches (duplicate)
                                                                            'if duplicate among different columns, then change formatting
                    If num.value = box.value And box.Address <> num.Address And box <> Empty Then
                    num.Interior.Color = RGB(0, 32, 96)
                    num.Font.Color = RGB(255, 255, 255)
                    End If
                 Next box
                    
                                                                            'change formatting for blank cells
                If x = 1 And num = Empty And num.Offset(0, 1) = Empty Then
                    num.Interior.Color = RGB(255, 255, 255)
                End If
                If x = 1 And num = Empty And num.Offset(0, 1) <> Empty Then
                    num.value = "(blank)"
                    num.Interior.Color = RGB(255, 255, 255)
                End If
                If x > 1 And num = Empty And num.Offset(0, 1) = Empty Then
                    num.Interior.Color = RGB(255, 255, 255)
                End If
                If x > 1 And num = Empty And num.Offset(0, 1) <> Empty Then
                    num.value = "(blank)"
                    num.Interior.Color = RGB(255, 255, 255)
                    num.Font.Color = RGB(0, 0, 0)
                End If
                    
                    
            Next num
        End If
    Next
    
    Application.DisplayAlerts = False
    If tagtitlerange.MergeCells Then
                tagtitlerange.VerticalAlignment = xlCenter
                tagtitlerange.HorizontalAlignment = xlCenter
                Else
                tagtitlerange.Merge Across:=True
                tagtitlerange.VerticalAlignment = xlCenter
                tagtitlerange.HorizontalAlignment = xlCenter
            End If
            With tagtitlerange.Cells
                 .BorderAround LineStyle:=xlContinuous, _
                    Weight:=xlThin
            End With
            
    MainTagTbl.Range.EntireColumn.AutoFit

Set colorrng = Range(ActiveSheet.Columns(MainTagTbl.ListColumns(1).Range.Column).Cells(1, 1).Address, ActiveSheet.Columns(MainTagTbl.ListColumns(1).Range.Column).Cells(4, 1))
If colorrng.MergeCells Then
            colorrng.Cells.UnMerge
        End If
        
With colorrng.Cells(1, 1).Offset(0, 1)
    .value = "Format"
    .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
End With
With colorrng.Cells(3, 1).Offset(0, 1)
    .value = "Duplicates Across Tier"
    .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
            .Interior.Color = RGB(0, 32, 96)
            .Font.Color = RGB(255, 255, 255)
End With
With colorrng.Cells(4, 1).Offset(0, 1)
    .value = "Testing for Tag"
    .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
            .Interior.Color = RGB(255, 51, 204)
            .Font.Color = RGB(255, 255, 255)
End With

For Each cell In colorrng
    cell.value = "Color Code"
    cell.BorderAround LineStyle:=xlContinuous, _
            Weight:=xlThin
    cell.Font.Bold = True
    
Next
            If colorrng.MergeCells Then
                Else
                colorrng.Merge
                colorrng.VerticalAlignment = xlCenter
                colorrng.HorizontalAlignment = xlCenter
            End If
            With colorrng.Cells
                 .BorderAround LineStyle:=xlContinuous, _
                    Weight:=xlThin
                    .Font.Bold = True
            End With
            


Set countrng = Range(ActiveSheet.Columns(MainTagTbl.ListColumns(3).Range.Column).Cells(1, 1).Address, ActiveSheet.Columns(MainTagTbl.ListColumns(3).Range.Column).Cells(4, 1))
If countrng.MergeCells Then
            countrng.Cells.UnMerge
        End If
        
With countrng.Cells(1, 1).Offset(0, 1)
    .value = "Format"
    .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
End With
With countrng.Cells(2, 1).Offset(0, 1)
    .value = "count < 3"
    .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
            .Interior.Color = RGB(84, 130, 53)
            .Font.Color = RGB(255, 255, 255)
End With
With countrng.Cells(3, 1).Offset(0, 1)
    .value = "3 <= count <= 6"
    .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
            .Interior.Color = RGB(255, 204, 102)
            .Font.Color = RGB(0, 0, 0)
End With
With countrng.Cells(4, 1).Offset(0, 1)
    .value = "count > 6"
    .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
            .Interior.Color = RGB(255, 124, 128)
            .Font.Color = RGB(0, 0, 0)
End With

For Each cell In countrng
    cell.value = "Counts"
    cell.BorderAround LineStyle:=xlContinuous, _
            Weight:=xlThin
    cell.Font.Bold = True
    
Next
            If countrng.MergeCells Then
                Else
                countrng.Merge
                countrng.VerticalAlignment = xlCenter
                countrng.HorizontalAlignment = xlCenter
            End If
            With countrng.Cells
                 .BorderAround LineStyle:=xlContinuous, _
                    Weight:=xlThin
                    .Font.Bold = True
            End With
End Sub


Sub TagsbyCatTblUpdate(start_date As Date, end_date As Date)


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
Dim CatArr() As Variant
Dim col_val As String

Dim counter_tbl As ListObject
   Set counter_sheet = Worksheets("Countermeasures")
   Set counter_tbl = counter_sheet.ListObjects("Tbl_Counter")
   Set counter_col = counter_tbl.ListColumns

    Set MainTagTable = Worksheets("Tag and Descriptor Tables").ListObjects("MainTagTable")

' if excel table already exists, delete and replace with new one
    SheetName = "Tag and Descriptor Tables"
    MyName = "TagsbyCatTbl"
    With MainTagTable.Range
      Set top_right = .Cells(1, .Columns.Count)
          col_num = .Columns.Count

    End With
    
                entry_count = 0
                    For Each cell In counter_tbl.ListColumns("Category").DataBodyRange
                    CRow = cell.row
                                       'if first entry, redim to hold one spot "(0)"
                                       If entry_count = 0 Then
                                           ReDim Preserve CatArr(0)
                                           CatArr(0) = cell
                                        'For all subsequent entries extend array by 1 and enter contents in cell
                                       Else
                                           ReDim Preserve CatArr(UBound(CatArr) + 1)
                                           CatArr(UBound(CatArr)) = cell
                                       End If
                                       entry_count = entry_count + 1
                Next cell
            
            'MsgBox Join(CatArr, vbCrLf)
            
    CatArr = BlankRemover(CatArr)
                
    Dim CatDict As Scripting.Dictionary
    Set CatDict = New Scripting.Dictionary
    
    'rids duplicates and counts new dictionary
    Set CatDict = DuplicateCountToScript(CatArr)
     cat_num = CatDict.Count
     
    For x = 1 To cat_num
    Next x
    
    
    Set New_Top_Left = top_right.Offset(0, 2)
    Set New_Bottom_Right = New_Top_Left.Offset(1, ((col_num / 2) * (cat_num * 2)) - 1)
    

    MyRangeString = New_Top_Left.Address & ":" & New_Bottom_Right.Address
    MyListExists = False
    For Each ListObj In Sheets(SheetName).ListObjects
        If ListObj.Name = MyName Then
            ListObj.Range.Cells.Interior.Color = RGB(255, 255, 255)
            ListObj.Range.ClearFormats
            ListObj.Range.ClearContents
            MyListExists = False
        End If
    Next ListObj
    
    If Not (MyListExists) Then
        Set TagByCategoryTbl = ActiveSheet.ListObjects.Add(xlSrcRange, Range(MyRangeString), , xlYes)
        TagByCategoryTbl.Name = MyName
    End If
    
    On Error Resume Next
    TagByCategoryTbl.HeaderRowRange.Clear
    For Each cell In TagByCategory.HeaderRowRange
        cell.Offset(-1, 0).Clear
    Next
    
    
    For x = 1 To (col_num / 2)
        If x = 1 Then
        Set tagtitlerange = Range(TagByCategoryTbl.HeaderRowRange(1, 1).Offset(-1, 0).Address, TagByCategoryTbl.HeaderRowRange(1, (TagByCategoryTbl.Range.Columns.Count / (col_num / 2) * x)).Offset(-1, 0).Address)
        'Debug.Print tagtitlerange.Address
        Else
        Set tagtitlerange = Range(TagByCategoryTbl.HeaderRowRange(1, ((cat_num * 2) * (x - 1)) + 1).Offset(-1, 0).Address, TagByCategoryTbl.HeaderRowRange(1, (TagByCategoryTbl.Range.Columns.Count / (col_num / 2) * x)).Offset(-1, 0).Address)
       End If
        If tagtitlerange.MergeCells Then
            tagtitlerange.Cells.UnMerge
        End If
    Next
    
    For Each col In TagByCategoryTbl.Range.Columns
        Column = Column + 1
        Tag_Multiplier = Application.WorksheetFunction.RoundUp(Column / ((col_num + ((cat_num * 2) - col_num))), 0)
        
        If Tag_Multiplier = 1 Then
        TagByCategoryTbl.HeaderRowRange(1, Column).Offset(-1, 0).value = MainTagTable.HeaderRowRange(1, 1).value
        Else
        TagByCategoryTbl.HeaderRowRange(1, Column).Offset(-1, 0).value = MainTagTable.HeaderRowRange(1, 1 + (2 * (Tag_Multiplier - 1))).value
        End If
        
       If Column Mod 2 <> 0 Then
       'Debug.Print Application.WorksheetFunction.RoundUp(Column / 2, 0) Mod cat_num
            If Application.WorksheetFunction.RoundUp(Column / 2, 0) Mod cat_num = 0 Then
                catrow = cat_num
            Else
                catrow = Application.WorksheetFunction.RoundUp(Column / 2, 0) Mod cat_num
            End If
        
        TagByCategoryTbl.HeaderRowRange(1, Column).value = CatDict.Keys(catrow - 1) & " - " & TagByCategoryTbl.HeaderRowRange(1, Column).Offset(-1, 0).value
        'Debug.Print CatDict.Keys(catrow - 1)
        Else
        TagByCategoryTbl.HeaderRowRange(1, Column).value = CatDict.Keys(catrow - 1) & " - Count"
        'Debug.Print CatDict.Keys(catrow - 1)
       End If
       
    Next
    
        StartDay = start_date
    ' Check if valid date but not the first of the month
    ' -- if so, reset StartDay to first day of month.
    If Day(StartDay) <> 1 Then
        StartDay = DateValue(Month(StartDay) & "/1/" & _
            Year(StartDay))
    End If
     
    FinalDay = DateValue(end_date)
    
 Column = 1
 For Each tblcell In TagByCategoryTbl.HeaderRowRange
    If InStr(1, tblcell.value, "Count") Then
    Column = Column + 1
    Else
    tag_cat = Split(tblcell.value, " - ")
    tag_title = tblcell.Offset(-1, 0)
    counter_row_count = 0
    entry_count = 0
                For Each cell In counter_tbl.ListColumns(tag_title).DataBodyRange
                    counter_row_count = counter_row_count + 1
                    CRow = cell.row
                         If counter_tbl.ListColumns("Issue Date").DataBodyRange.Cells(counter_row_count, 1).value >= StartDay _
                            And counter_tbl.ListColumns("Issue Date").DataBodyRange.Cells(counter_row_count, 1).value < FinalDay Then
                                  If counter_tbl.ListColumns("Category").DataBodyRange.Cells(CRow - 1, 1).value = tag_cat(0) Then
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
                                End If
                        End If
                Next cell
                
               ' MsgBox Join(ArrBase, vbCrLf)
               
    If (Not Not ArrBase) = 0 Then GoTo FollowingColumn
                
    Dim ArrDict As Scripting.Dictionary
    Set ArrDict = New Scripting.Dictionary
    
    'rids duplicates and counts new dictionary
    Set ArrDict = DuplicateCountToScript(ArrBase)
    
    Call Functions_M.SortDictionary(ArrDict, False, True, vbBinaryCompare)
    
    TagByCategoryTbl.DataBodyRange.Cells(1, Column).Resize(ArrDict.Count, 1).Value2 = Application.Transpose(ArrDict.Keys)
    TagByCategoryTbl.DataBodyRange.Cells(1, Column + 1).Resize(ArrDict.Count, 1).Value2 = Application.Transpose(ArrDict.Items)
    
    If Column = 1 Then
    resize_num = ArrDict.Items.Count
    Else
       If ArrDict.Count > resize_num Then
            resize_num = ArrDict.Count
        End If
    End If
    
    With TagByCategoryTbl.Range
    TagByCategoryTbl.Resize .Resize(resize_num + 1)
    End With


FollowingColumn:
    Column = Column + 1
    
    End If
    Erase ArrBase
Next tblcell

Application.DisplayAlerts = False
  For x = 1 To (col_num / 2)
        If x = 1 Then
        Set x1left = TagByCategoryTbl.HeaderRowRange(1, 1).Offset(1, 0)
        Set x1right = TagByCategoryTbl.HeaderRowRange(1, (TagByCategoryTbl.Range.Columns.Count / (col_num / 2) * x)).Offset(TagByCategoryTbl.DataBodyRange.Rows.Count, 0)
        Set Tag_Format_Range = Range(x1left.Address, x1right.Address)
        'Tag format range
        
        
        
        For Y = 1 To Tag_Format_Range.Columns.Count
            If Y Mod 2 = 0 Then
                For Each num In Tag_Format_Range.Columns(Y).Cells
                    
                    If num.value >= 0 And num.value <= 3 And num <> Empty And IsNumeric(num) = True Then
                    num.Interior.Color = RGB(84, 130, 53)
                    num.Font.Color = RGB(255, 255, 255)
                    ElseIf num.value >= 4 And num.value <= 6 Then
                    num.Interior.Color = RGB(255, 204, 102)
                    num.Font.Color = RGB(0, 0, 0)
                    ElseIf num.value >= 7 Then
                    num.Interior.Color = RGB(255, 124, 128)
                    num.Font.Color = RGB(0, 0, 0)
                    End If
                    
                    If num = Empty Then
                        If num = Empty And num.Offset(0, -1) <> Empty Then
                        num.value = "(blank)"
                        Else
                        num.Interior.Color = RGB(255, 255, 255)
                        End If
                    End If
                    
                Next num
            
            Else
                
                For Each num In Tag_Format_Range.Columns(Y).Cells
                    
                    'Debug.Print "num: " & num.value
                    'Debug.Print "num: " & num.Address
                    
                    For Each box In Tag_Format_Range
                     '   Debug.Print "box :" & box.value
                      '  Debug.Print "box :" & box.Address
                        If num.value = box.value And box.Address <> num.Address And box <> Empty And IsNumeric(num) = False Then
                        num.Interior.Color = RGB(255, 242, 204)
                        num.Font.Color = RGB(0, 0, 0)
                        End If
                     Next box
                        
                        
                    If Y = 1 And num = Empty And num.Offset(0, 1) = Empty Then
                        num.Interior.Color = RGB(255, 255, 255)
                    End If
                    If Y = 1 And num = Empty And num.Offset(0, 1) <> Empty Then
                        num.value = "(blank)"
                        num.Interior.Color = RGB(255, 255, 255)
                    End If
                    If Y > 1 And num = Empty And num.Offset(0, 1) = Empty Then
                        num.Interior.Color = RGB(255, 255, 255)
                    End If
                    If Y > 1 And num = Empty And num.Offset(0, 1) <> Empty Then
                        num.value = "(blank)"
                        num.Interior.Color = RGB(255, 255, 255)
                        num.Font.Color = RGB(0, 0, 0)
                    End If
                        
                        
                Next num
            End If
        Next
        
        
        Else
        Set xleft = x1right.Offset(0, (x - 1)) 'TagByCategoryTbl.HeaderRowRange(1, (col_num * (x - 1)) + 1).Offset(1, 0)
        Set xright = TagByCategoryTbl.HeaderRowRange(1, (TagByCategoryTbl.Range.Columns.Count / (col_num / 2) * x)).Offset(TagByCategoryTbl.DataBodyRange.Rows.Count, 0)
        Set Tag_Format_Range = Range(xleft.Address, xright.Address)
       
       For Y = 1 To Tag_Format_Range.Columns.Count
            If Y Mod 2 = 0 Then
                For Each num In Tag_Format_Range.Columns(Y).Cells
                    
                    If num.value >= 0 And num.value <= 3 And num <> Empty And IsNumeric(num) = True Then
                    num.Interior.Color = RGB(84, 130, 53)
                    num.Font.Color = RGB(255, 255, 255)
                    ElseIf num.value >= 4 And num.value <= 6 Then
                    num.Interior.Color = RGB(255, 204, 102)
                    num.Font.Color = RGB(0, 0, 0)
                    ElseIf num.value >= 7 Then
                    num.Interior.Color = RGB(255, 124, 128)
                    num.Font.Color = RGB(0, 0, 0)
                    End If
                    
                    If num = Empty Then
                        If num = Empty And num.Offset(0, -1) <> Empty Then
                        num.value = "(blank)"
                        Else
                        num.Interior.Color = RGB(255, 255, 255)
                        End If
                    End If
                    
                Next num
            
            Else
                
                For Each num In Tag_Format_Range.Columns(Y).Cells
                    
                    'Debug.Print "num: " & num.value
                    'Debug.Print "num: " & num.Address
                    
                    For Each box In Tag_Format_Range
                     '   Debug.Print "box :" & box.value
                      '  Debug.Print "box :" & box.Address
                        If num.value = box.value And box.Address <> num.Address And box <> Empty And IsNumeric(num) = False Then
                        num.Interior.Color = RGB(255, 242, 204)
                        num.Font.Color = RGB(0, 0, 0)
                        End If
                     Next box
                        
                        
                    If Y = 1 And num = Empty And num.Offset(0, 1) = Empty Then
                        num.Interior.Color = RGB(255, 255, 255)
                    End If
                    If Y = 1 And num = Empty And num.Offset(0, 1) <> Empty Then
                        num.value = "(blank)"
                        num.Interior.Color = RGB(255, 255, 255)
                    End If
                    If Y > 1 And num = Empty And num.Offset(0, 1) = Empty Then
                        num.Interior.Color = RGB(255, 255, 255)
                    End If
                    If Y > 1 And num = Empty And num.Offset(0, 1) <> Empty Then
                        num.value = "(blank)"
                        num.Interior.Color = RGB(255, 255, 255)
                        num.Font.Color = RGB(0, 0, 0)
                    End If
                        
                        
                Next num
            End If
        Next
       End If
       
        
        If x = 1 Then
          Set tagtitlerange = Range(TagByCategoryTbl.HeaderRowRange(1, 1).Offset(-1, 0).Address, TagByCategoryTbl.HeaderRowRange(1, (TagByCategoryTbl.Range.Columns.Count / (col_num / 2) * x)).Offset(-1, 0).Address)
        'Debug.Print tagtitlerange.Address
        'Debug.Print TagByCategoryTbl.Range.Columns.Count
        Else
        Set tagtitlerange = Range(TagByCategoryTbl.HeaderRowRange(1, ((cat_num * 2) * (x - 1)) + 1).Offset(-1, 0).Address, TagByCategoryTbl.HeaderRowRange(1, (TagByCategoryTbl.Range.Columns.Count / (col_num / 2) * x)).Offset(-1, 0).Address)
        'Debug.Print tagtitlerange.Address
        End If
        
        If tagtitlerange.MergeCells Then
            tagtitlerange.VerticalAlignment = xlCenter
            tagtitlerange.HorizontalAlignment = xlCenter
            Else
            tagtitlerange.Merge Across:=True
            tagtitlerange.VerticalAlignment = xlCenter
            tagtitlerange.HorizontalAlignment = xlCenter
        End If
        With tagtitlerange.Cells
             .BorderAround LineStyle:=xlContinuous, _
                Weight:=xlThin
        End With
    Next
Application.DisplayAlerts = True

With TagByCategoryTbl.Range
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
            
       For Each iCells In TagByCategoryTbl.Range
            iCells.BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThin
        Next iCells
        
        TagByCategoryTbl.TableStyle = "TableStyleMedium9"
        
    End With
    
    TagByCategoryTbl.Range.EntireColumn.AutoFit
    
Set colorrng = Range(ActiveSheet.Columns(TagByCategoryTbl.ListColumns(1).Range.Column).Cells(1, 1).Address, ActiveSheet.Columns(TagByCategoryTbl.ListColumns(1).Range.Column).Cells(4, 1))
If colorrng.MergeCells Then
            colorrng.Cells.UnMerge
        End If
        
With colorrng.Cells(1, 1).Offset(0, 1)
    .value = "Format"
    .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
End With
With colorrng.Cells(2, 1).Offset(0, 1)
    .value = "Duplicates Amongst Tier"
    .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
            .Interior.Color = RGB(255, 242, 204)
            .Font.Color = RGB(0, 0, 0)
End With
With colorrng.Cells(3, 1).Offset(0, 1)
    .value = "Duplicates Across Tiers"
    .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
            .Interior.Color = RGB(0, 32, 96)
            .Font.Color = RGB(255, 255, 255)
End With
With colorrng.Cells(4, 1).Offset(0, 1)
    .value = "Testing for Tag"
    .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
            .Interior.Color = RGB(255, 51, 204)
            .Font.Color = RGB(255, 255, 255)
End With

For Each cell In colorrng
    cell.value = "Color Code"
    cell.BorderAround LineStyle:=xlContinuous, _
            Weight:=xlThin
    cell.Font.Bold = True
    
Next
            Application.DisplayAlerts = False
            If colorrng.MergeCells Then
                Else
                colorrng.Merge
                colorrng.VerticalAlignment = xlCenter
                colorrng.HorizontalAlignment = xlCenter
            End If
            With colorrng.Cells
                 .BorderAround LineStyle:=xlContinuous, _
                    Weight:=xlThin
                    .Font.Bold = True
            End With
            
            


Set countrng = Range(ActiveSheet.Columns(TagByCategoryTbl.ListColumns(3).Range.Column).Cells(1, 1).Address, ActiveSheet.Columns(TagByCategoryTbl.ListColumns(3).Range.Column).Cells(4, 1))

If countrng.MergeCells Then
            countrng.Cells.UnMerge
        End If
        
With countrng.Cells(1, 1).Offset(0, 1)
    .value = "Format"
    .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
End With
With countrng.Cells(2, 1).Offset(0, 1)
    .value = "count < 3"
    .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
            .Interior.Color = RGB(84, 130, 53)
            .Font.Color = RGB(255, 255, 255)
End With
With countrng.Cells(3, 1).Offset(0, 1)
    .value = "3 <= count <= 6"
    .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
            .Interior.Color = RGB(255, 204, 102)
            .Font.Color = RGB(0, 0, 0)
End With
With countrng.Cells(4, 1).Offset(0, 1)
    .value = "count > 6"
    .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
            .Interior.Color = RGB(255, 124, 128)
            .Font.Color = RGB(0, 0, 0)
End With

For Each cell In countrng
    cell.value = "Counts"
    cell.BorderAround LineStyle:=xlContinuous, _
            Weight:=xlThin
    cell.Font.Bold = True
    
Next
            If countrng.MergeCells Then
                Else
                countrng.Merge
                countrng.VerticalAlignment = xlCenter
                countrng.HorizontalAlignment = xlCenter
            End If
            With countrng.Cells
                 .BorderAround LineStyle:=xlContinuous, _
                    Weight:=xlThin
                    .Font.Bold = True
            End With
            Application.DisplayAlerts = True

End Sub

Sub Tbl_CategoriesUpdate(start_date As Date, end_date As Date)

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
Dim Tags() As Variant

Dim ArrBase() As Variant
Dim col_val As String

Dim counter_tbl As ListObject
   Set counter_sheet = Worksheets("Countermeasures")
   Set counter_tbl = counter_sheet.ListObjects("Tbl_Counter")
   Set counter_col = counter_tbl.ListColumns

' if excel table already exists, delete and replace with new one
   SheetName = "Tag and Descriptor Tables"
   MyName = "Tbl_Categories"
   
   
    Set TagsBetweenTierTbl = Worksheets("Tag and Descriptor Tables").ListObjects("TagsbyCatTbl")

' if excel table already exists, delete and replace with new one
    SheetName = "Tag and Descriptor Tables"
    MyName = "Tbl_Categories"
    With TagsBetweenTierTbl.Range
          Set top_right = .Cells(1, .Columns.Count)
    End With
    
    'Call Cat_KPI_Updater_M.UpdateCategory
    Set New_Top_Left = top_right.Offset(0, 2)
    Set Bottom_Right = New_Top_Left.Offset(1, 1)
   
   MyRangeString = New_Top_Left.Address & ":" & Bottom_Right.Address
   
  MyListExists = False
    For Each ListObj In Sheets(SheetName).ListObjects
        If ListObj.Name = MyName Then
            ListObj.Range.ClearFormats
            ListObj.Range.ClearContents
            MyListExists = False
        End If
    Next ListObj
    
    If Not (MyListExists) Then
        Set CatTable = ActiveSheet.ListObjects.Add(xlSrcRange, Range(MyRangeString), , xlYes)
        CatTable.Name = MyName
    End If
    
CatTable.HeaderRowRange(1, 1) = "Category"
CatTable.HeaderRowRange(1, 2) = "Count"

Set tagtitlerange = Range(CatTable.HeaderRowRange(1, 1).Offset(-1, 0), CatTable.HeaderRowRange(1, CatTable.Range.Columns.Count).Offset(-1, 0))
        If tagtitlerange.MergeCells Then
            tagtitlerange.Cells.UnMerge
        End If
        
        
        StartDay = start_date
    ' Check if valid date but not the first of the month
    ' -- if so, reset StartDay to first day of month.
    If Day(StartDay) <> 1 Then
        StartDay = DateValue(Month(StartDay) & "/1/" & _
            Year(StartDay))
    End If
     
    FinalDay = DateValue(end_date)

Column = Column + 1
For Each tblcell In CatTable.HeaderRowRange

    tblcell.Offset(-1, 0) = "Category"

    If InStr(1, tblcell.value, "Count") Then
    Column = Column + 1
    Else
    entry_count = 0
    counter_row_count = 0
                For Each Cell3 In counter_tbl.ListColumns(tblcell.value).DataBodyRange
                    counter_row_count = counter_row_count + 1
                    If counter_tbl.ListColumns("Issue Date").DataBodyRange.Cells(counter_row_count, 1).value >= StartDay _
                            And counter_tbl.ListColumns("Issue Date").DataBodyRange.Cells(counter_row_count, 1).value < FinalDay Then
                                       'if first entry, redim to hold one spot "(0)"
                                       If entry_count = 0 Then
                                           ReDim Preserve ArrBase(0)
                                           ArrBase(0) = Cell3
                                        'For all subsequent entries extend array by 1 and enter contents in cell
                                       Else
                                           ReDim Preserve ArrBase(UBound(ArrBase) + 1)
                                           ArrBase(UBound(ArrBase)) = Cell3
                                       End If
                                       entry_count = entry_count + 1
                    End If
                Next Cell3
                
    If (Not Not ArrBase) = 0 Then GoTo FollowingColumn
                
    ArrBase = BlankRemover(ArrBase)
                   
    Dim ArrDict As Scripting.Dictionary
    Set ArrDict = New Scripting.Dictionary
    
    'rids duplicates and counts new dictionary
    Set ArrDict = DuplicateCountToScript(ArrBase)
    
    Call Functions_M.SortDictionary(ArrDict, False, True, vbBinaryCompare)

    
    CatTable.DataBodyRange.Cells(1, Column).Resize(ArrDict.Count, 1).Value2 = Application.Transpose(ArrDict.Keys)
    CatTable.DataBodyRange.Cells(1, Column + 1).Resize(ArrDict.Count, 1).Value2 = Application.Transpose(ArrDict.Items)
    
    If Column = 1 Then
     resize_num = ArrDict.Count
    Else
        If ArrDict.Count > resize_num Then
            resize_num = ArrDict.Items.Count
        End If
    End If
    
    With CatTable.Range
    CatTable.Resize .Resize(resize_num + 1)
    End With
    
FollowingColumn:
    Column = Column + 1
    
    End If
    Erase ArrBase
Next tblcell
'Formatting
    With CatTable.Range
            .HorizontalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
            
       For Each iCells In CatTable.Range
            iCells.BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThin
        Next iCells
        
        CatTable.TableStyle = "TableStyleMedium9"
        
    End With
    
    For x = 1 To CatTable.Range.Columns.Count
        If x Mod 2 = 0 Then
            For Each num In CatTable.ListColumns(x).DataBodyRange
                
                If num.value >= 0 And num.value <= 3 And num <> Empty Then
                num.Interior.Color = RGB(84, 130, 53)
                num.Font.Color = RGB(255, 255, 255)
                ElseIf num.value >= 4 And num.value <= 6 Then
                num.Interior.Color = RGB(255, 204, 102)
                num.Font.Color = RGB(0, 0, 0)
                ElseIf num.value >= 7 Then
                num.Interior.Color = RGB(255, 124, 128)
                num.Font.Color = RGB(0, 0, 0)
                End If
                
                If num = Empty Then
                    If num = Empty And num.Offset(0, -1) <> Empty Then
                    num.value = "(blank)"
                    Else
                    num.Interior.Color = RGB(255, 255, 255)
                    End If
                End If
                
            Next num
        
        Else
            
            For Each num In CatTable.ListColumns(x).DataBodyRange
                
                For Each box In CatTable.DataBodyRange
                    If num.value = box.value And box.Address <> num.Address And box <> Empty Then
                    num.Interior.Color = RGB(0, 32, 96)
                    num.Font.Color = RGB(255, 255, 255)
                    End If
                 Next box
                    
                    
                If x = 1 And num = Empty And num.Offset(0, 1) = Empty Then
                    num.Interior.Color = RGB(255, 255, 255)
                End If
                If x = 1 And num = Empty And num.Offset(0, 1) <> Empty Then
                    num.value = "(blank)"
                    num.Interior.Color = RGB(255, 255, 255)
                End If
                If x > 1 And num = Empty And num.Offset(0, 1) = Empty Then
                    num.Interior.Color = RGB(255, 255, 255)
                End If
                If x > 1 And num = Empty And num.Offset(0, 1) <> Empty Then
                    num.value = "(blank)"
                    num.Interior.Color = RGB(255, 255, 255)
                    num.Font.Color = RGB(0, 0, 0)
                End If
                    
                    
            Next num
        End If
    Next
    Application.DisplayAlerts = False
    If tagtitlerange.MergeCells Then
                tagtitlerange.HorizontalAlignment = xlCenter
                tagtitlerange.VerticalAlignment = xlCenter
                Else
                tagtitlerange.Merge Across:=True
                tagtitlerange.HorizontalAlignment = xlCenter
                tagtitlerange.VerticalAlignment = xlCenter
            End If
            With tagtitlerange.Cells
                 .BorderAround LineStyle:=xlContinuous, _
                    Weight:=xlThin
            End With
            Application.DisplayAlerts = True
    
    CatTable.Range.EntireColumn.AutoFit
    
Set countrng = Range(ActiveSheet.Columns(CatTable.ListColumns(1).Range.Column).Cells(1, 1).Address, ActiveSheet.Columns(CatTable.ListColumns(1).Range.Column).Cells(4, 1))

If countrng.MergeCells Then
            countrng.Cells.UnMerge
        End If
        
With countrng.Cells(1, 1).Offset(0, 1)
    .value = "Format"
    .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
End With
With countrng.Cells(2, 1).Offset(0, 1)
    .value = "count < 3"
    .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
            .Interior.Color = RGB(84, 130, 53)
            .Font.Color = RGB(0, 0, 0)
End With
With countrng.Cells(3, 1).Offset(0, 1)
    .value = "3 <= count <= 6"
    .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
            .Interior.Color = RGB(255, 204, 102)
            .Font.Color = RGB(0, 0, 0)
End With
With countrng.Cells(4, 1).Offset(0, 1)
    .value = "count > 6"
    .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
            .Interior.Color = RGB(255, 124, 128)
            .Font.Color = RGB(0, 0, 0)
End With

For Each cell In countrng
    cell.value = "Counts"
    cell.BorderAround LineStyle:=xlContinuous, _
            Weight:=xlThin
    cell.Font.Bold = True
    
Next
        Application.DisplayAlerts = False
            If countrng.MergeCells Then
                Else
                countrng.Merge
                countrng.VerticalAlignment = xlCenter
                countrng.HorizontalAlignment = xlCenter
            End If
            With countrng.Cells
                 .BorderAround LineStyle:=xlContinuous, _
                    Weight:=xlThin
                    .Font.Bold = True
            End With
         Application.DisplayAlerts = True
    
End Sub

Sub Tbl_KPIUpdate(start_date As Date, end_date As Date)


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
Dim CatArr() As Variant
Dim col_val As String

Dim counter_tbl As ListObject
   Set counter_sheet = Worksheets("Countermeasures")
   Set counter_tbl = counter_sheet.ListObjects("Tbl_Counter")
   Set counter_col = counter_tbl.ListColumns

    Set CatTable = Worksheets("Tag and Descriptor Tables").ListObjects("Tbl_Categories")

' if excel table already exists, delete and replace with new one
    SheetName = "Tag and Descriptor Tables"
    MyName = "Tbl_KPI"
    With CatTable.Range
      Set top_right = .Cells(1, .Columns.Count)
    End With
                entry_count = 0
                    For Each cell In counter_tbl.ListColumns("Category").DataBodyRange
                    CRow = cell.row
                                       'if first entry, redim to hold one spot "(0)"
                                       If entry_count = 0 Then
                                           ReDim Preserve CatArr(0)
                                           CatArr(0) = cell
                                        'For all subsequent entries extend array by 1 and enter contents in cell
                                       Else
                                           ReDim Preserve CatArr(UBound(CatArr) + 1)
                                           CatArr(UBound(CatArr)) = cell
                                       End If
                                       entry_count = entry_count + 1
                Next cell
            
            'MsgBox Join(CatArr, vbCrLf)
            
    CatArr = BlankRemover(CatArr)
                
    Dim CatDict As Scripting.Dictionary
    Set CatDict = New Scripting.Dictionary
    
    'rids duplicates and counts new dictionary
    Set CatDict = DuplicateCountToScript(CatArr)
     cat_num = CatDict.Count
    
    Set New_Top_Left = top_right.Offset(0, 1)
    Set New_Bottom_Right = New_Top_Left.Offset(1, (cat_num * 2) - 1)
    

    MyRangeString = New_Top_Left.Address & ":" & New_Bottom_Right.Address
    
    MyListExists = False
    For Each ListObj In Sheets(SheetName).ListObjects
        If ListObj.Name = MyName Then
            ListObj.Range.Cells.Interior.Color = RGB(255, 255, 255)
            ListObj.Range.ClearFormats
            ListObj.Range.ClearContents
            MyListExists = False
        End If
    Next ListObj
    
    If Not (MyListExists) Then
        Set KPITable = ActiveSheet.ListObjects.Add(xlSrcRange, Range(MyRangeString), , xlYes)
        KPITable.Name = MyName
    End If
    
    On Error Resume Next
    KPITable.HeaderRowRange.Clear
    For Each cell In KPITable.HeaderRowRange
        cell.Offset(-1, 0).Clear
    Next
    
        Set tagtitlerange = Range(KPITable.HeaderRowRange(1, 1).Offset(-1, 0), KPITable.HeaderRowRange(1, KPITable.Range.Columns.Count).Offset(-1, 0))
        If tagtitlerange.MergeCells Then
            tagtitlerange.Cells.UnMerge
        End If
    
 col = 1
    For col = 1 To KPITable.Range.Columns.Count
        
        KPITable.HeaderRowRange(1, col).Offset(-1, 0) = "KPI"
        
       If col Mod 2 <> 0 Then
       'Debug.Print Application.WorksheetFunction.RoundUp(Column / 2, 0) Mod cat_num
            If Application.WorksheetFunction.RoundUp(col / 2, 0) Mod cat_num = 0 Then
                catrow = cat_num
            Else
                catrow = Application.WorksheetFunction.RoundUp(col / 2, 0) Mod cat_num
            End If
        
        KPITable.HeaderRowRange(1, col).value = CatDict.Keys(catrow - 1)
        'Debug.Print CatDict.Keys(catrow - 1)
        Else
        KPITable.HeaderRowRange(1, col).value = CatDict.Keys(catrow - 1) & " - Count"
        'Debug.Print CatDict.Keys(catrow - 1)
       End If
       
    Next
    
    
     StartDay = start_date
    ' Check if valid date but not the first of the month
    ' -- if so, reset StartDay to first day of month.
    If Day(StartDay) <> 1 Then
        StartDay = DateValue(Month(StartDay) & "/1/" & _
            Year(StartDay))
    End If
     
    FinalDay = DateValue(end_date)
    
 Column = 1
 For Each tblcell In KPITable.HeaderRowRange
    If InStr(1, tblcell.value, "Count") Then
    Column = Column + 1
    Else
    
    tag_title = tblcell.Offset(-1, 0)
    
    entry_count = 0
    counter_row_count = 0
                For Each cell In counter_tbl.ListColumns(tag_title).DataBodyRange
                    'Debug.Print counter_tbl.ListColumns(col_val).DataBodyRange.Address
                    CRow = cell.row
                    counter_row_count = 1 + counter_row_count
                         If counter_tbl.ListColumns("Issue Date").DataBodyRange.Cells(counter_row_count, 1).value >= StartDay _
                            And counter_tbl.ListColumns("Issue Date").DataBodyRange.Cells(counter_row_count, 1).value < FinalDay Then
                                  If counter_tbl.ListColumns("Category").DataBodyRange.Cells(CRow - 1, 1).value = tblcell.value Then
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
                                End If
                                'Debug.Print ArrBase(UBound(ArrBase))
                        End If
                Next cell
                
                'MsgBox Join(ArrBase, vbCrLf)
                
    If (Not Not ArrBase) = 0 Then GoTo FollowingColumn
    
    Dim ArrDict As Scripting.Dictionary
    Set ArrDict = New Scripting.Dictionary
    
    'rids duplicates and counts new dictionary
    Set ArrDict = DuplicateCountToScript(ArrBase)
    
    Call Functions_M.SortDictionary(ArrDict, False, True, vbBinaryCompare)
    
    KPITable.DataBodyRange.Cells(1, Column).Resize(ArrDict.Count, 1).Value2 = Application.Transpose(ArrDict.Keys)
    KPITable.DataBodyRange.Cells(1, Column + 1).Resize(ArrDict.Count, 1).Value2 = Application.Transpose(ArrDict.Items)
    

    If Column = 1 Then
    resize_num = ArrDict.Count
    Else
       If ArrDict.Count > resize_num Then
            resize_num = ArrDict.Count
        End If
    End If
    
    With KPITable.Range
    KPITable.Resize .Resize(resize_num + 1)
    End With

FollowingColumn:
    Column = Column + 1
    
    End If
    Erase ArrBase
Next tblcell

Application.DisplayAlerts = False
  For Y = 1 To KPITable.Range.Columns.Count
        Set Tag_Format_Range = Range(KPITable.HeaderRowRange(1, 1).Offset(1, 0), KPITable.HeaderRowRange(1, KPITable.Range.Columns.Count).Offset(KPITable.DataBodyRange.Rows.Count, 0))
        'Debug.Print Tag_Format_Range.Address
            If Y Mod 2 = 0 Then
                For Each num In Tag_Format_Range.Columns(Y).Cells
                    
                    If num.value >= 0 And num.value <= 3 And num <> Empty And IsNumeric(num) = True Then
                    num.Interior.Color = RGB(84, 130, 53)
                    num.Font.Color = RGB(255, 255, 255)
                    ElseIf num.value >= 4 And num.value <= 6 Then
                    num.Interior.Color = RGB(255, 204, 102)
                    num.Font.Color = RGB(0, 0, 0)
                    ElseIf num.value >= 7 Then
                    num.Interior.Color = RGB(255, 124, 128)
                    num.Font.Color = RGB(0, 0, 0)
                    End If
                    
                   If num = Empty Then
                        If num = Empty And num.Offset(0, -1) <> Empty Then
                        num.value = "(blank)"
                        Else
                        num.Interior.Color = RGB(255, 255, 255)
                        End If
                    End If
                    
                Next num
            
            Else
                
                For Each num In Tag_Format_Range.Columns(Y).Cells
                    
                    For Each box In Tag_Format_Range
                     '   Debug.Print "box :" & box.value
                      '  Debug.Print "box :" & box.Address
                        If num.value = box.value And box.Address <> num.Address And box <> Empty And IsNumeric(num) = False Then
                        num.Interior.Color = RGB(255, 242, 204)
                        num.Font.Color = RGB(0, 0, 0)
                        End If
                     Next box
                        
                    If Y = 1 And num = Empty And num.Offset(0, 1) = Empty Then
                        num.Interior.Color = RGB(255, 255, 255)
                    End If
                    If Y = 1 And num = Empty And num.Offset(0, 1) <> Empty Then
                        num.value = "(blank)"
                        num.Interior.Color = RGB(255, 255, 255)
                    End If
                    If Y > 1 And num = Empty And num.Offset(0, 1) = Empty Then
                        num.Interior.Color = RGB(255, 255, 255)
                    End If
                    If Y > 1 And num = Empty And num.Offset(0, 1) <> Empty Then
                        num.value = "(blank)"
                        num.Interior.Color = RGB(255, 255, 255)
                        num.Font.Color = RGB(0, 0, 0)
                    End If
                Next num
            End If
    Next
            'Debug.Print tagtitlerange.Address
            
            If tagtitlerange.MergeCells Then
                tagtitlerange.HorizontalAlignment = xlCenter
                tagtitlerange.VerticalAlignment = xlCenter
                Else
                tagtitlerange.Merge Across:=True
                tagtitlerange.HorizontalAlignment = xlCenter
                tagtitlerange.VerticalAlignment = xlCenter
            End If
            With tagtitlerange.Cells
                 .BorderAround LineStyle:=xlContinuous, _
                    Weight:=xlThin
            End With
            Application.DisplayAlerts = True

With KPITable.Range
            .HorizontalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
            
       For Each iCells In KPITable.Range
            iCells.BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThin
        Next iCells
        
        KPITable.TableStyle = "TableStyleMedium9"
        
    End With
    
    KPITable.Range.EntireColumn.AutoFit

End Sub

Sub Tbl_IdentifierUpdate(start_date As Date, end_date As Date)

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
Dim identifiers() As Variant

Dim ArrBase() As Variant
Dim col_val As String

Dim counter_tbl As ListObject
   Set counter_sheet = Worksheets("Countermeasures")
   Set counter_tbl = counter_sheet.ListObjects("Tbl_Counter")
   Set counter_col = counter_tbl.ListColumns
            
For Each cell In counter_tbl.HeaderRowRange
    If cell.value = "KPI" Then
    questions_colval = cell.Column
    End If
    If cell.value = "Issue" Then
    Category_ColVal = cell.Column
    End If
Next cell

entry_count = 0
For Each cell2 In counter_tbl.HeaderRowRange
    If cell2.Column > questions_colval And cell2.Column < Category_ColVal Then
    
    If entry_count = 0 Then
        ReDim Preserve identifiers(0)
        identifiers(0) = cell2
     'For all subsequent entries extend array by 1 and enter contents in cell
    Else
        ReDim Preserve identifiers(UBound(identifiers) + 1)
        identifiers(UBound(identifiers)) = cell2
    End If
    entry_count = entry_count + 1
    End If
    
    'MsgBox Join(Tags, vbCrLf)

Next cell2

' if excel table already exists, delete and replace with new one
   SheetName = "Tag and Descriptor Tables"
   MyName = "Tbl_Identifiers"
   
   Set KPITable = Worksheets("Tag and Descriptor Tables").ListObjects("Tbl_KPI")

' if excel table already exists, delete and replace with new one
    SheetName = "Tag and Descriptor Tables"
    MyName = "Tbl_Identifier"
    With KPITable.Range
      Set top_right = .Cells(1, .Columns.Count)

    End With
    Set New_Top_Left = top_right.Offset(0, 2)

   Set Bottom_Right = New_Top_Left.Offset(1, ((UBound(identifiers) + 1) * 2) - 1)
   
   MyRangeString = New_Top_Left.Address & ":" & Bottom_Right.Address
   
  MyListExists = False
    For Each ListObj In Sheets(SheetName).ListObjects
        If ListObj.Name = MyName Then
            ListObj.Range.ClearFormats
            ListObj.Range.ClearContents
            MyListExists = False
        End If
    Next ListObj
    
    If Not (MyListExists) Then
        Set IdentifierTable = ActiveSheet.ListObjects.Add(xlSrcRange, Range(MyRangeString), , xlYes)
        IdentifierTable.Name = MyName
    End If
    
For x = 1 To IdentifierTable.Range.Columns.Count
    'Debug.Print IdentifierTable.Range.Columns.Count
    'Debug.Print x Mod (UBound(Identifiers) + 1)
    'Debug.Print ((x Mod (UBound(Tags))) - 1)
    'Debug.Print IdentifierTable.Range.Columns.Count / (x * 2)
    
    If x Mod 2 = 0 Then
        IdentifierTable.HeaderRowRange(1, x) = IdentifierTable.HeaderRowRange(1, x - 1).value & " Count"
    
    Else
        
        Tag_Multiplier = Application.WorksheetFunction.RoundUp(x / 2, 0)
        
        If Tag_Multiplier = 1 Then
        IdentifierTable.HeaderRowRange(1, x) = identifiers(0)
        Else
        IdentifierTable.HeaderRowRange(1, x) = identifiers(Tag_Multiplier - 1)
        End If
    End If
Next

Set tagtitlerange = Range(IdentifierTable.HeaderRowRange(1, 1).Offset(-1, 0), IdentifierTable.HeaderRowRange(1, IdentifierTable.Range.Columns.Count).Offset(-1, 0))
        If tagtitlerange.MergeCells Then
            tagtitlerange.Cells.UnMerge
        End If

Column = Column + 1


 StartDay = start_date
    ' Check if valid date but not the first of the month
    ' -- if so, reset StartDay to first day of month.
    If Day(StartDay) <> 1 Then
        StartDay = DateValue(Month(StartDay) & "/1/" & _
            Year(StartDay))
    End If
     
    FinalDay = DateValue(end_date)

For Each tblcell In IdentifierTable.HeaderRowRange
    tblcell.Offset(-1, 0) = "Other Identifiers"
    If InStr(1, tblcell.value, "Count") Then
    Column = Column + 1
    Else
    entry_count = 0
       
    cell_counter = 0
    entry_count = 0
    Dim cel_entries() As String
    Dim Entry As Variant
    counter_row_count = 0
        For Each cell In counter_tbl.ListColumns(tblcell.value).DataBodyRange
            cel_entries = Split(cell.value, "; ")
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
             counter_row_count = counter_row_count + 1
                    If cell <> Empty And counter_tbl.ListColumns("Issue Date").DataBodyRange.Cells(counter_row_count, 1).value >= StartDay _
                            And counter_tbl.ListColumns("Issue Date").DataBodyRange.Cells(counter_row_count, 1).value < FinalDay Then
                        cell_counter = cell_counter + 1
                    
                        'for each entry, check if first entry, if so redim array to number of cell entries
                        ' if not, for each entry extend array by one and add entry
                        For Each Entry In cel_entries
                            'MsgBox entry
                            'For first entry
                            If entry_count = 0 Then
                                ReDim Preserve ArrBase(UBound(cel_entries))
                                ArrBase(i) = Entry
                                'For all subsequent entries
                                Else
                                    If ArrBase(UBound(ArrBase)) <> Empty Then
                                    ReDim Preserve ArrBase(UBound(ArrBase) + 1)
                                    ArrBase(UBound(ArrBase)) = Entry
                                    Else
                                    ArrBase(entry_count) = Entry
                    End If
                    End If
                    entry_count = entry_count + 1
                    
                    Next
                    End If
        Next
                
                'MsgBox Join(ArrBase, vbCrLf)
    If (Not Not ArrBase) = 0 Then GoTo FollowingColumn
    
                
    Dim ArrDict As Scripting.Dictionary
    Set ArrDict = New Scripting.Dictionary
    
    'rids duplicates and counts new dictionary
    Set ArrDict = DuplicateCountToScript(ArrBase)
    
    Call Functions_M.SortDictionary(ArrDict, False, True, vbBinaryCompare)
    
    IdentifierTable.DataBodyRange.Cells(1, Column).Resize(ArrDict.Count, 1).Value2 = Application.Transpose(ArrDict.Keys)
    IdentifierTable.DataBodyRange.Cells(1, Column + 1).Resize(ArrDict.Count, 1).Value2 = Application.Transpose(ArrDict.Items)
    
    If Column = 1 Then
    resize_num = ArrDict.Count
    Else
        If ArrDict.Count > resize_num Then
            resize_num = ArrDict.Count + 1
        End If
    End If
    
    With IdentifierTable.Range
    IdentifierTable.Resize .Resize(resize_num)
    End With
FollowingColumn:
    Column = Column + 1
    
    End If
    Erase ArrBase
Next tblcell
'Formatting
    With IdentifierTable.Range
            .HorizontalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
            
       For Each iCells In IdentifierTable.Range
            iCells.BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThin
        Next iCells
        
        IdentifierTable.TableStyle = "TableStyleMedium9"
        
    End With
    
    For x = 1 To IdentifierTable.Range.Columns.Count
        If x Mod 2 = 0 Then
            For Each num In IdentifierTable.ListColumns(x).DataBodyRange
                
                If num.value >= 0 And num.value <= 3 And num <> Empty Then
                num.Interior.Color = RGB(84, 130, 53)
                num.Font.Color = RGB(255, 255, 255)
                ElseIf num.value >= 4 And num.value <= 6 Then
                num.Interior.Color = RGB(255, 204, 102)
                num.Font.Color = RGB(0, 0, 0)
                ElseIf num.value >= 7 Then
                num.Interior.Color = RGB(255, 124, 128)
                num.Font.Color = RGB(0, 0, 0)
                End If
                
                If num = Empty Then
                    If num = Empty And num.Offset(0, -1) <> Empty Then
                    num.value = "(blank)"
                    Else
                    num.Interior.Color = RGB(255, 255, 255)
                    End If
                End If
                
            Next num
        
        Else
            
            For Each num In IdentifierTable.ListColumns(x).DataBodyRange
                
                For Each box In IdentifierTable.DataBodyRange
                    If num.value = box.value And box.Address <> num.Address And box <> Empty Then
                    num.Interior.Color = RGB(0, 32, 96)
                    num.Font.Color = RGB(255, 255, 255)
                    End If
                 Next box
                    
                    
                If x = 1 And num = Empty And num.Offset(0, 1) = Empty Then
                    num.Interior.Color = RGB(255, 255, 255)
                End If
                If x = 1 And num = Empty And num.Offset(0, 1) <> Empty Then
                    num.value = "(blank)"
                    num.Interior.Color = RGB(255, 255, 255)
                End If
                If x > 1 And num = Empty And num.Offset(0, 1) = Empty Then
                    num.Interior.Color = RGB(255, 255, 255)
                End If
                If x > 1 And num = Empty And num.Offset(0, 1) <> Empty Then
                    num.value = "(blank)"
                    num.Interior.Color = RGB(255, 255, 255)
                    num.Font.Color = RGB(0, 0, 0)
                End If
                    
                    
            Next num
        End If
    Next
    
    Application.DisplayAlerts = False
    If tagtitlerange.MergeCells Then
                tagtitlerange.VerticalAlignment = xlCenter
                tagtitlerange.HorizontalAlignment = xlCenter
                Else
                tagtitlerange.Merge Across:=True
                tagtitlerange.VerticalAlignment = xlCenter
                tagtitlerange.HorizontalAlignment = xlCenter
            End If
            With tagtitlerange.Cells
                 .BorderAround LineStyle:=xlContinuous, _
                    Weight:=xlThin
            End With
            Application.DisplayAlerts = True
    
    
    IdentifierTable.Range.EntireColumn.AutoFit
    

Set countrng = Range(ActiveSheet.Columns(IdentifierTable.ListColumns(1).Range.Column).Cells(1, 1).Address, ActiveSheet.Columns(IdentifierTable.ListColumns(1).Range.Column).Cells(4, 1))

If countrng.MergeCells Then
            countrng.Cells.UnMerge
        End If
        
With countrng.Cells(1, 1).Offset(0, 1)
    .value = "Format"
    .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
End With
With countrng.Cells(2, 1).Offset(0, 1)
    .value = "count < 3"
    .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
            .Interior.Color = RGB(84, 130, 53)
            .Font.Color = RGB(0, 0, 0)
End With
With countrng.Cells(3, 1).Offset(0, 1)
    .value = "3 <= count <= 6"
    .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
            .Interior.Color = RGB(255, 204, 102)
            .Font.Color = RGB(0, 0, 0)
End With
With countrng.Cells(4, 1).Offset(0, 1)
    .value = "count > 6"
    .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
            .Font.Name = "Calibri"
            .Interior.Color = RGB(255, 124, 128)
            .Font.Color = RGB(0, 0, 0)
End With

For Each cell In countrng
    cell.value = "Counts"
    cell.BorderAround LineStyle:=xlContinuous, _
            Weight:=xlThin
    cell.Font.Bold = True
    
Next
        Application.DisplayAlerts = False
            If countrng.MergeCells Then
                Else
                countrng.Merge
                countrng.VerticalAlignment = xlCenter
                countrng.HorizontalAlignment = xlCenter
            End If
            With countrng.Cells
                 .BorderAround LineStyle:=xlContinuous, _
                    Weight:=xlThin
                    .Font.Bold = True
            End With
         Application.DisplayAlerts = True
        

End Sub
Sub CreateTagTester()
Dim Tag As String
Dim Obj As Object
Dim TagTables As ListObject
Dim TagSheet As Worksheet
Dim SheetName As String
Dim buttonname As String
Dim buttoncaption As String
Dim buttonheight As Double
Dim buttonwidth As Double
Dim buttontop As Double
Dim buttonleft As Double
Dim subtitle As String
Dim subprocedure As String

SheetName = "Tag and Descriptor Tables"
Set TagSheet = Worksheets("Tag and Descriptor Tables")

For Each TagTables In Worksheets("Tag and Descriptor Tables").ListObjects
    If TagTables.Name = "TagTesterTable" Then
    TagTables.Delete
    End If
Next

    MyRangeString = "A7:A8"
    MyName1 = "TagTesterTable"
    MyListExists = False
    For Each ListObj In Sheets(SheetName).ListObjects
        If ListObj.Name = MyName1 Then
            ListObj.Range.ClearFormats
            ListObj.Range.ClearContents
            MyListExists = False
        End If
    Next ListObj
    
    If Not (MyListExists) Then
        Set TagTesterTable = ActiveSheet.ListObjects.Add(xlSrcRange, Range(MyRangeString), , xlYes)
        TagTesterTable.Name = MyName1
    End If
    
    TagTesterTable.HeaderRowRange(1, 1) = "Tag to Test:"
   
   End Sub



Sub CreateTagTestwithCodeButton()

Dim Tag As String
Dim Obj As Object
Dim TagTables As ListObject
Dim TagSheet As Worksheet
Dim SheetName As String
Dim buttonname As String
Dim buttoncaption As String
Dim buttonheight As Double
Dim buttonwidth As Double
Dim buttontop As Double
Dim buttonleft As Double
Dim subtitle As String
Dim subprocedure As String

SheetName = "Tag and Descriptor Tables"
    
    
    Set TagTesterTable = Worksheets(SheetName).ListObjects("TagTesterTable")
    
    buttonname = "Test for Tag"
    buttoncaption = "Update " & "Tag Tester"
    Set buttonrng = TagTesterTable.HeaderRowRange.Offset(-2, 0)
    buttonwidth = buttonrng.Width
    buttonheight = buttonrng.Height
    buttonleft = buttonrng.left
    buttontop = buttonrng.Top
    
    subtitle = buttonname & "_Click()"
    subprocedure = "TagTesterTable" & "Update"
    
    AddButtonandCode2 SheetName, buttonname, buttoncaption, buttonleft, buttontop, buttonwidth, buttonheight, subprocedure, subtitle

End Sub


Sub TagTesterTableUpdate()
'sub takes the string in the tag tester table and retrieves all the addresses, table titles, and cell contents
'of the cells that contain that string

Dim Tag As String
Dim Obj As Object
Dim TagSheet As Worksheet
Dim TagTables As ListObject
Dim AddressColl As New Collection
Dim TableColl As New Collection
Dim CellValueColl As New Collection
Dim ColorColl As New Collection
Dim FontColl As New Collection
Dim TagAddressTable As ListObject
Dim TagNameTable As ListObject
Dim SheetName As String
Dim buttonname As String
Dim buttoncaption As String
Dim buttonheight As Double
Dim buttonwidth As Double
Dim buttontop As Double
Dim buttonleft As Double
Dim subtitle As String
Dim subprocedure As String

SheetName = "Tag and Descriptor Tables"

Set TagSheet = Worksheets("Tag and Descriptor Tables")      'set sheet var

Set TagNameTable = Worksheets("Tag and Descriptor Tables").ListObjects("TagTesterTable")        'set table var for tag to be tested
    Tag = TagNameTable.DataBodyRange(1, 1).value                                                'set tag var
          TagNameTable.DataBodyRange(1, 1).Interior.Color = RGB(255, 51, 204)
          TagNameTable.DataBodyRange(1, 1).Font.Color = RGB(255, 255, 255)
    
    MyRangeString = "A10:C11"                                                                   'get address for output table
    MyName2 = "TagTesterAddress"                                                                'name output table
    MyListExists = False                                                                        'assume table does not exist
    For Each ListObj In Sheets("Tag and Descriptor Tables").ListObjects                         'if list (tag table) exists, delete
        If ListObj.Name = MyName2 Then
            ListObj.Delete
            MyListExists = False
        End If
    Next ListObj
    
    If Not (MyListExists) Then                                                                  'if output table doesnt exist, create, if does assign table var
        Set TagAddressTable = ActiveSheet.ListObjects.Add(xlSrcRange, Range(MyRangeString), , xlYes)
        TagAddressTable.Name = MyName2
        Else
        Set TagAddressTable = Worksheets("Tag and Descriptor Tables").ListObjects("TagTesterAddress")
    End If
    
        'name output table headers
    TagAddressTable.HeaderRowRange(1, 1) = "Address"
    TagAddressTable.HeaderRowRange(1, 2) = "Table Name"
    TagAddressTable.HeaderRowRange(1, 3) = "Cell Value"
        
        
        'for each cell in each table in the sheet, see if cells contains match the Tag to Test
        'if it does, add cells address, main table, and contents to collections, and change the colorof that cell to purple
    If Tag <> Empty Then
        For Each TagTables In Worksheets("Tag and Descriptor Tables").ListObjects
              For Each cell In TagTables.DataBodyRange
                If cell.value Like "*" & Tag & "*" Then
        
                    AddressColl.Add cell.Address
                    TableColl.Add TagTables.Name
                    CellValueColl.Add cell.value
                    
                    With cell
                        .Interior.Color = RGB(255, 51, 204)
                        .Font.Color = RGB(255, 255, 255)
                    End With
                End If
              Next
              
        Next
        
        'put the collections contents into the table
        row = 1
        For thing = 1 To AddressColl.Count
            If row = 1 Then
                TagAddressTable.DataBodyRange(row, 1).value = AddressColl.item(thing)
                TagAddressTable.DataBodyRange(row, 2).value = TableColl.item(thing)
                TagAddressTable.DataBodyRange(row, 3).value = CellValueColl.item(thing)
            Else
                TagAddressTable.ListRows.Add
                TagAddressTable.DataBodyRange(row, 1).value = AddressColl.item(thing)
                TagAddressTable.DataBodyRange(row, 2).value = TableColl.item(thing)
                TagAddressTable.DataBodyRange(row, 3).value = CellValueColl.item(thing)
            End If
            row = row + 1
        Next
    Else
    End If
    
    For row = 1 To TagAddressTable.ListRows.Count
        If TagAddressTable.DataBodyRange(row, 2) = "TagTesterTable" Then
            TagAddressTable.DataBodyRange(row, 1).value = "End"
            TagAddressTable.DataBodyRange(row, 2).value = "End"
            TagAddressTable.DataBodyRange(row, 3).value = "End"
        End If
    Next
    
End Sub



