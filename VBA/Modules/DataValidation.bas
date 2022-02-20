Attribute VB_Name = "DataValidation"
'Sub routines present in "countermeasures" sheet

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean) 'on doule left click, create data validation for cell based on category of row
    'This code applies to activate in two areas, if a cell underneath a tag column is double clicked, and
'if a cell underneath an identifier column is double clicked. First, the code checks to see if the double click
'happened within the table on the countermeasure sheet. If it didn't, the code ends. Next,
'The code will identifies the tag columns and the identifier columns by seeing what columns
'exist between the "Issue ID" and "Category columns, and then the same between the KPI and Issue columns.
'This allows additions, deletions and changes to those columns and titles and this feature will remain the same and functional.
'After identifying what are the relevant columns, the code loops through the tag columns
'And if it is determined the double click happened in that column, the code continues.
'The code takes the category of the clicked cell and the tag column of the clicked cell and
'creates an array from the tag column where the category of the rows = the category of the selected row
'Then, a sheet is created and a table is created to hold these values.
'Then, data validation is hard coded to pull from the newly created (or updated table)
'The drop down validation stays until the cell is right clicked (see next script)


Dim Validation_Array() As Variant                'array to hold value from tag column
Dim identifiers() As Variant            'array to hold identifier column titles (list of columns of identifiers)
Dim Tags() As Variant                   'array to hold tags column titles (list of columns of tags)
Dim CatArr() As Variant                 'array to holddifferent category values
Dim col_val As String                   'variable to hold column title
Dim TableLocation As Range              'variable to hold current table address
Dim NameofColofTarget As String         'column title of column of target cell
Dim CatofTarget As String               'category in row of target cell
Dim ValidatedTags As New Collection     'variable to hold tag and identifier arrays (organized by category
Dim Tbl_Validation As ListObject        'Tables to hold validation values for different tags/identifiers and categoryes                    'Variable = databodyrange row of target cell
Dim Validation_ID As String                    'List ID to add to collection
Dim TempArray() As Variant              'Temp array in case array generated from table doesn't exist or is blank

Dim counter_tbl As ListObject
   Set counter_sheet = Worksheets("Countermeasures")
   Set counter_tbl = counter_sheet.ListObjects("Tbl_Counter")
   Set counter_col = counter_tbl.ListColumns
    
If Intersect(Target, ThisWorkbook.Worksheets("Countermeasures").ListObjects("Tbl_Counter").DataBodyRange) Is Nothing Then 'if double click is NOT within countermeasures excel table, then do nothing.

    Else 'If double click is within table, then...
    
            For Each cell In counter_tbl.HeaderRowRange    'for each column title in table
                If cell.value = "Issue ID" Then             'if title = "Issue ID", then:
                    issue_id_colval = cell.Column           'assign column number to variable
                End If
                If cell.value = "Category" Then             'if title = "category", then:
                    cat_ColVal = cell.Column                'assign column number to variable
                End If
            Next cell
            
                entry_count = 0
            For Each cell2 In counter_tbl.HeaderRowRange        'for each column
                    If cell2.Column > issue_id_colval And cell2.Column < cat_ColVal Then    'if column is in between "Issue ID" and "Category", then it means its a "Tag" column
                                                                                            'if "tag column", then add title to Tags() array
                            If entry_count = 0 Then
                                ReDim Preserve Tags(0)
                                Tags(0) = cell2
                             'For all subsequent entries extend array by 1 and enter contents in cell
                            Else
                                ReDim Preserve Tags(UBound(Tags) + 1)
                                Tags(UBound(Tags)) = cell2
                            End If
                            
                    entry_count = entry_count + 1
                    
                    End If
            Next cell2
            
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
                    Tags = ArrayRemoveDups(Tags)  'if not empty, remove duplicates
                ElseIf IsTagsEmpty = True Then     'if array is empty,  resize to one entry and create one entry "No List Available"
                    ReDim Preserve Tags(0)
                    Tags(0) = "No List Available"
            End If
            
                'if column title = KPI or Issue, then assign column number to variable
            For Each cell In counter_tbl.HeaderRowRange
                If cell.value = "KPI" Then
                    KPI_ColVal = cell.Column
                End If
                If cell.value = "Issue" Then
                    issue_ColVal = cell.Column
                End If
            Next cell
            
             entry_count = 0
            For Each cell2 In counter_tbl.HeaderRowRange
                    If cell2.Column > KPI_ColVal And cell2.Column < issue_ColVal Then 'if column is in between "KPI" and "Issue", then it means its a "Identifier" column
                                                                                            'if "identifier column", then add title to identifiers() array
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
            Next cell2
        
            If (Not Not identifiers) = 0 Then       'if array never intitialized then it doesn't exist,
                    IsIdentifierEmpty = True         'if list doesn't exist, then it's empty (by default)
                Else
                    IsIdentifierEmpty = True         'if list does exist, assume it is empty
                    For Each item_in_array In identifiers       'test if array is empty. If there is one non-blank cell, change bool value
                        If item_in_array <> Empty Then
                            IsIdentifierEmpty = False
                        End If
                    Next
            End If
                    
            If IsIdentifierEmpty = False Then            'if array isn't empty, remove blanks, remove dups
                    identifiers = BlankRemover(identifiers)     'if not empty, remove blanks
                    identifiers = ArrayRemoveDups(identifiers)  'if not empty, remove duplicates
                ElseIf IsTagsEmpty = True Then     'if array is empty,  resize to one entry and create one entry "No List Available"
                    ReDim Preserve identifiers(0)
                    identifiers(0) = "No List Available"
            End If
    
    
    
    
            
        ColofTarget = Target.Column - Target.ListObject.DataBodyRange.Column + 1        'column of selected cell
        DataBodyRowofTarget = Target.row - Target.ListObject.DataBodyRange.row + 1      'data body row of selected cell
        
        NameofColofTarget = ThisWorkbook.Worksheets("Countermeasures").ListObjects("Tbl_Counter").HeaderRowRange(1, ColofTarget).value          'column title of selected cell
        CatofTarget = ThisWorkbook.Worksheets("Countermeasures").ListObjects("Tbl_Counter").ListColumns("Category").DataBodyRange(DataBodyRowofTarget, 1)   'category of row of selected cell
           
        If CatofTarget = "" Then
            IsCatBlank = True
            Else
            IsCatBlank = False
        End If
           
        For Each Tag In Tags        'for each tag column
            If Intersect(Target, ThisWorkbook.Worksheets("Countermeasures").ListObjects("Tbl_Counter").ListColumns(Tag).DataBodyRange) Is Nothing Then  'if double clicked cell is not within the tag column, do nothing.
       
                Else    'if the double click is within the iterated tag column, then:
                    
                    If IsCatBlank = False Then                  'if category is not blank, select values in tag array where entries cat = target entry cat
                                    entry_count = 0
                                    row_count = 0
                                For Each cell In counter_tbl.ListColumns(Tag).DataBodyRange      'create array of tag where entry rows match target category
                                    row_count = row_count + 1
                                    If counter_tbl.ListColumns("Category").DataBodyRange(row_count, 1).value = CatofTarget Then
                                            'if first entry, redim to hold one spot "(0)"
                                            If entry_count = 0 Then
                                                ReDim Preserve Validation_Array(0)
                                                Validation_Array(0) = cell
                                             'For all subsequent entries extend array by 1 and enter contents in cell
                                            Else
                                                ReDim Preserve Validation_Array(UBound(Validation_Array) + 1)
                                                Validation_Array(UBound(Validation_Array)) = cell
                                            End If
                                            entry_count = entry_count + 1
                                    End If
                                Next cell
                        Else                                'if categiry is blank,  enter all tags from tag column into array
                                    entry_count = 0
                                    row_count = 0
                                For Each cell In counter_tbl.ListColumns(Tag).DataBodyRange      'create array of tag where entry rows match target category
                                            'if first entry, redim to hold one spot "(0)"
                                            If entry_count = 0 Then
                                                ReDim Preserve Validation_Array(0)
                                                Validation_Array(0) = cell
                                             'For all subsequent entries extend array by 1 and enter contents in cell
                                            Else
                                                ReDim Preserve Validation_Array(UBound(Validation_Array) + 1)
                                                Validation_Array(UBound(Validation_Array)) = cell
                                            End If
                                            entry_count = entry_count + 1
                                Next cell
                    End If
                    
                        'MsgBox Join(Validation_Array, vbCrLf)
                    
                    If (Not Not Validation_Array) = 0 Then       'if array never intitialized then it doesn't exist,
                            IsArrayEmpty = True         'if list doesn't exist, then it's empty (by default)
                        Else
                            IsArrayEmpty = True         'if list does exist, assume it is empty
                            For Each item_in_array In Validation_Array       'test if array is empty. If there is one non-blank cell, change bool value
                                If item_in_array <> Empty Then
                                    IsArrayEmpty = False
                                End If
                            Next
                    End If
                            
                    If IsArrayEmpty = False Then            'if array isn't empty, remove blanks, remove dups
                            Validation_Array = BlankRemover(Validation_Array)     'if not empty, remove blanks
                            Validation_Array = ArrayRemoveDups(Validation_Array)  'if not empty, remove duplicates
                        ElseIf IsArrayEmpty = True Then     'if array is empty,  resize to one entry and create one entry "No List Available"
                            ReDim Preserve Validation_Array(0)
                            Validation_Array(0) = "No List Available"
                    End If
                        
                        'MsgBox Join(Validation_Array, vbCrLf)
                        
                    Validation_ID = CatofTarget & " " & NameofColofTarget        'combine category and tag column title for unique call string
                            
            
                    ' if excel table already exists, delete and replace with new one
                    SheetName = "DataValidation"            'Name of sheet that will hold data validation table
                    MyName = Validation_ID
                    
                    DoesSheetExist = False          'assume sheet does not exist
                    
                    For Each sheet In ThisWorkbook.Worksheets
                        If sheet.Name = SheetName Then
                            DoesSheetExist = True               'check if sheet does exist, change bool value
                        End If
                    Next sheet
                        
                        Num_of_Wkst = ThisWorkbook.Worksheets.Count
                        
                        If DoesSheetExist = False Then              'if sheet does not exist, create sheet named "DataValidation" at the end of list of worksheets
                                                                    'if sheet does exist, assign Var Validation_Sheets to sheet
                            Sheets.Add After:=ThisWorkbook.Worksheets(Num_of_Wkst)
                            ThisWorkbook.Worksheets(Num_of_Wkst + 1).Name = SheetName
                            Set Validation_Sheet = ThisWorkbook.Worksheets(SheetName)
                                'turns on alerts after initial deletion/creation
                        End If
                    
                    DoesTableexist = False                          'assume table does not exist
                    For Each tbl In Worksheets(SheetName).ListObjects   'check to see if table exist, if it does assign
                            If tbl.Name = MyName Then
                            DoesTableexist = True
                            End If
                            ListObjectCount = ListObjectCount + 1       'count number of list objects in sheet
                    Next
                    
                        If ListObjectCount = 0 Then 'if first table in sheet, assign range, if not, offset next table location by one column from previous table location
                                Set TableLocation = ThisWorkbook.Worksheets(SheetName).Range("C3:C5")
                            Else
                                Set TableLocation = ThisWorkbook.Worksheets(SheetName).Range("C3:C5").Offset(0, ListObjectCount)
                        End If
                        
                    If DoesTableexist = False Then      'if table doesn't exist, then create table
                         Set Tbl_Validation = ThisWorkbook.Worksheets(SheetName).ListObjects.Add(xlSrcRange, TableLocation, , xlYes)
                             Tbl_Validation.Name = MyName
                             Tbl_Validation.HeaderRowRange(1, 1).value = MyName
                        Set Tbl_Validation = ThisWorkbook.Worksheets(SheetName).ListObjects(MyName)
                        Else                            'if table does exist, assign table to variable
                        Set Tbl_Validation = ThisWorkbook.Worksheets(SheetName).ListObjects(MyName)
                    End If
                   
                    DoesColExist = False                    'assume columns for specific tag title and category does not exist
                    For Each Column In Tbl_Validation.ListColumns
                        If Column.Name = Validation_ID Then
                            DoesColExist = True             'if column does exist change bool value
                        End If
                    Next
                    
                        
                    Tbl_Validation.HeaderRowRange(1, 1).value = Validation_ID     'name column of table
                            
                    Array_count = 0
                    For Each item In Validation_Array            'count how large the array is
                        Array_count = Array_count + 1
                    Next
                    
                    If Tbl_Validation.DataBodyRange.Rows.Count < Array_count Then               'if column length is less than the size of incoming array, then:
                        Cell1 = Tbl_Validation.HeaderRowRange(1, 1).Address                     'Starting address of column
                        cell2 = Tbl_Validation.HeaderRowRange.Offset(Array_count, 0).Address    'Ending address of column
                        Tbl_Validation.Resize Range(Cell1, cell2)        'increase size of column to hold incoming array
                    End If
                    
                    Set PasteRng = Tbl_Validation.ListColumns(Validation_ID).DataBodyRange.Cells(1).Resize(Array_count) 'Set variable for databodyrange of table
                    PasteRng.value = WorksheetFunction.Transpose(Validation_Array)                                      'paste array
                
                
                    ValidationRange = Tbl_Validation.ListColumns(CatofTarget & " " & NameofColofTarget).DataBodyRange.Address   'assign address of new column to variable
                    
                      With Target.Validation      'take column from validation table and assign it to target cell
                          .Delete
                          .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                          xlBetween, Formula1:="='" & SheetName & "'!" & ValidationRange 'Formula1:="=Tbl_DataValidation" & "[" & CatofTarget & " " & NameofColofTarget & "]"
                          .IgnoreBlank = True
                          .InCellDropdown = True
                          .InputTitle = ""
                          .ErrorTitle = ""
                          .InputMessage = ""
                          .ErrorMessage = ""
                          .ShowInput = True
                          .ShowError = True
                      End With
                    
      
            End If
        Next Tag
        
        
       Dim cel_entries() As String
            
       For Each identifier In identifiers
            If Intersect(Target, ThisWorkbook.Worksheets("Countermeasures").ListObjects("Tbl_Counter").ListColumns(identifier).DataBodyRange) Is Nothing Then
                'MsgBox "Target not in Table1"
                Else
                    
                        entry_count = 0
                        For Each cell In counter_tbl.ListColumns(identifier).DataBodyRange
                            cel_entries = Split(cell.value, "; ")
                            counter_row_count = counter_row_count + 1
                            If cell <> Empty Then
                                'for each entry, check if first entry, if so redim array to number of cell entries
                                ' if not, for each entry extend array by one and add entry
                                For Each Entry In cel_entries
                                   ' Debug.Print Entry
                                    'For first entry
                                    If entry_count = 0 Then
                                        Debug.Print UBound(cel_entries)
                                        ReDim Preserve Validation_Array(UBound(cel_entries))
                                        Validation_Array(i) = Entry
                                        'For all subsequent entries
                                        entry_count = entry_count + 1
                                        Else
                                            Debug.Print UBound(Validation_Array)
                                            If Validation_Array(UBound(Validation_Array)) <> Empty Then
                                            ReDim Preserve Validation_Array(UBound(Validation_Array) + 1)
                                            Validation_Array(UBound(Validation_Array)) = Entry
                                            entry_count = entry_count + 1
                                            Else
                                            Validation_Array(entry_count) = Entry
                                            entry_count = entry_count + 1
                                            End If
                                    End If
                            
                                Next
                            End If
                        Next
                            
                        'MsgBox Join(Validation_Array, vbCrLf)
                    
                    If (Not Not Validation_Array) = 0 Then       'if array never intitialized then it doesn't exist,
                            IsArrayEmpty = True         'if list doesn't exist, then it's empty (by default)
                        Else
                            IsArrayEmpty = True         'if list does exist, assume it is empty
                            For Each item_in_array In Validation_Array       'test if array is empty. If there is one non-blank cell, change bool value
                                If item_in_array <> Empty Then
                                    IsArrayEmpty = False
                                End If
                            Next
                    End If
                            
                    If IsArrayEmpty = False Then            'if array isn't empty, remove blanks, remove dups
                            Validation_Array = BlankRemover(Validation_Array)     'if not empty, remove blanks
                            Validation_Array = ArrayRemoveDups(Validation_Array)  'if not empty, remove duplicates
                        ElseIf IsArrayEmpty = True Then     'if array is empty,  resize to one entry and create one entry "No List Available"
                            ReDim Preserve Validation_Array(0)
                            Validation_Array(0) = "No List Available"
                    End If
                        
                        'MsgBox Join(Validation_Array, vbCrLf)
                        
                    Validation_ID = NameofColofTarget        'combine category and tag column title for unique call string
                            
            
                    ' if excel table already exists, delete and replace with new one
                    SheetName = "DataValidation"            'Name of sheet that will hold data validation table
                    MyName = Validation_ID
                    
                    DoesSheetExist = False          'assume sheet does not exist
                    
                    For Each sheet In ThisWorkbook.Worksheets
                        If sheet.Name = SheetName Then
                            DoesSheetExist = True               'check if sheet does exist, change bool value
                        End If
                    Next sheet
                        
                        Num_of_Wkst = ThisWorkbook.Worksheets.Count
                        
                        If DoesSheetExist = False Then              'if sheet does not exist, create sheet named "DataValidation" at the end of list of worksheets
                                                                    'if sheet does exist, assign Var Validation_Sheets to sheet
                            Sheets.Add After:=ThisWorkbook.Worksheets(Num_of_Wkst)
                            ThisWorkbook.Worksheets(Num_of_Wkst + 1).Name = SheetName
                            Set Validation_Sheet = ThisWorkbook.Worksheets(SheetName)
                                'turns on alerts after initial deletion/creation
                        End If
                    
                    DoesTableexist = False                          'assume table does not exist
                    For Each tbl In Worksheets(SheetName).ListObjects   'check to see if table exist, if it does assign
                            If tbl.Name = MyName Then
                            DoesTableexist = True
                            End If
                            ListObjectCount = ListObjectCount + 1       'count number of list objects in sheet
                    Next
                    
                        If ListObjectCount = 0 Then 'if first table in sheet, assign range, if not, offset next table location by one column from previous table location
                                Set TableLocation = ThisWorkbook.Worksheets(SheetName).Range("C3:C5")
                            Else
                                Set TableLocation = ThisWorkbook.Worksheets(SheetName).Range("C3:C5").Offset(0, ListObjectCount)
                        End If
                        
                    If DoesTableexist = False Then      'if table doesn't exist, then create table
                         Set Tbl_Validation = ThisWorkbook.Worksheets(SheetName).ListObjects.Add(xlSrcRange, TableLocation, , xlYes)
                             Tbl_Validation.Name = MyName
                             Tbl_Validation.HeaderRowRange(1, 1).value = MyName
                        Set Tbl_Validation = ThisWorkbook.Worksheets(SheetName).ListObjects(MyName)
                        Else
                        Set Tbl_Validation = ThisWorkbook.Worksheets(SheetName).ListObjects(MyName)
                    End If
                   
                    DoesColExist = False                    'assume columns for specific tag title and category does not exist
                    For Each Column In Tbl_Validation.ListColumns
                        If Column.Name = Validation_ID Then
                            DoesColExist = True             'if column does exist change bool value
                        End If
                    Next
                    
                        
                    Tbl_Validation.HeaderRowRange(1, 1).value = NameofColofTarget     'name new column the string of incoming array
                            
                    Array_count = 0
                    For Each item In Validation_Array            'count how large the array is
                        Array_count = Array_count + 1
                    Next
                    
                    If Tbl_Validation.DataBodyRange.Rows.Count < Array_count Then                               'if column length is less than the size of incoming array, then:
                        Cell1 = Tbl_Validation.HeaderRowRange(1, 1).Address
                        cell2 = Tbl_Validation.HeaderRowRange.Offset(Array_count, 0).Address
                        Tbl_Validation.Resize Range(Cell1, cell2)        'increase size to that of incoming array
                    End If
                        Set PasteRng = Tbl_Validation.ListColumns(Validation_ID).DataBodyRange.Cells(1).Resize(Array_count)
                    PasteRng.value = WorksheetFunction.Transpose(Validation_Array)
                
                
                    ValidationRange = Tbl_Validation.ListColumns(NameofColofTarget).DataBodyRange.Address   'assign address of new column to variable
                    
                      With Target.Validation      'take column from validation table and assign it to target cell
                          .Delete
                          .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                          xlBetween, Formula1:="='" & SheetName & "'!" & ValidationRange 'Formula1:="=Tbl_DataValidation" & "[" & CatofTarget & " " & NameofColofTarget & "]"
                          .IgnoreBlank = True
                          .InCellDropdown = True
                          .InputTitle = ""
                          .ErrorTitle = ""
                          .InputMessage = ""
                          .ErrorMessage = ""
                          .ShowInput = True
                          .ShowError = True
                      End With
                    
                      
      
            End If
        Next
        
        
End If

End Sub




Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)

Dim Validation_Array() As Variant
Dim identifiers() As Variant
Dim Tags() As Variant
Dim CatArr() As Variant
Dim col_val As String
Dim NameofColofTarget As String
Dim CatofTarget As String

Dim counter_tbl As ListObject
   Set counter_sheet = Worksheets("Countermeasures")
   Set counter_tbl = counter_sheet.ListObjects("Tbl_Counter")
   Set counter_col = counter_tbl.ListColumns
    
                    
   For Each cell In counter_tbl.HeaderRowRange    'for each column title in table
                If cell.value = "Issue ID" Then             'if title = "Issue ID", then:
                    issue_id_colval = cell.Column           'assign column number to variable
                End If
                If cell.value = "Category" Then             'if title = "category", then:
                    cat_ColVal = cell.Column                'assign column number to variable
                End If
            Next cell
            
                entry_count = 0
            For Each cell2 In counter_tbl.HeaderRowRange        'for each column
                    If cell2.Column > issue_id_colval And cell2.Column < cat_ColVal Then    'if column is in between "Issue ID" and "Category", then it means its a "Tag" column
                                                                                            'if "tag column", then add title to Tags() array
                            If entry_count = 0 Then
                                ReDim Preserve Tags(0)
                                Tags(0) = cell2
                             'For all subsequent entries extend array by 1 and enter contents in cell
                            Else
                                ReDim Preserve Tags(UBound(Tags) + 1)
                                Tags(UBound(Tags)) = cell2
                            End If
                            
                    entry_count = entry_count + 1
                    
                    End If
            Next cell2
            
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
                    Tags = ArrayRemoveDups(Tags)  'if not empty, remove duplicates
                ElseIf IsTagsEmpty = True Then     'if array is empty,  resize to one entry and create one entry "No List Available"
                    ReDim Preserve Tags(0)
                    Tags(0) = "No List Available"
            End If
            
                'if column title = KPI or Issue, then assign column number to variable
            For Each cell In counter_tbl.HeaderRowRange
                If cell.value = "KPI" Then
                    KPI_ColVal = cell.Column
                End If
                If cell.value = "Issue" Then
                    issue_ColVal = cell.Column
                End If
            Next cell
            
             entry_count = 0
            For Each cell2 In counter_tbl.HeaderRowRange
                    If cell2.Column > KPI_ColVal And cell2.Column < issue_ColVal Then 'if column is in between "KPI" and "Issue", then it means its a "Identifier" column
                                                                                            'if "identifier column", then add title to identifiers() array
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
            Next cell2
        
            If (Not Not identifiers) = 0 Then       'if array never intitialized then it doesn't exist,
                    IsIdentifierEmpty = True         'if list doesn't exist, then it's empty (by default)
                Else
                    IsIdentifierEmpty = True         'if list does exist, assume it is empty
                    For Each item_in_array In identifiers       'test if array is empty. If there is one non-blank cell, change bool value
                        If item_in_array <> Empty Then
                            IsIdentifierEmpty = False
                        End If
                    Next
            End If
                    
            If IsIdentifierEmpty = False Then            'if array isn't empty, remove blanks, remove dups
                    identifiers = BlankRemover(identifiers)     'if not empty, remove blanks
                    identifiers = ArrayRemoveDups(identifiers)  'if not empty, remove duplicates
                ElseIf IsTagsEmpty = True Then     'if array is empty,  resize to one entry and create one entry "No List Available"
                    ReDim Preserve identifiers(0)
                    identifiers(0) = "No List Available"
            End If
    
If Intersect(Target, ThisWorkbook.Worksheets("Countermeasures").ListObjects("Tbl_Counter").DataBodyRange) Is Nothing Then   'if click isn't in countermeasure table, do nothing
            Else        'if it is,
            
    ColofTarget = Target.Column - Target.ListObject.DataBodyRange.Column + 1        'get target column num
    DataBodyRowofTarget = Target.row - Target.ListObject.DataBodyRange.row + 1      'get databodyrow of target cell
    
    NameofColofTarget = ThisWorkbook.Worksheets("Countermeasures").ListObjects("Tbl_Counter").HeaderRowRange(1, ColofTarget).value      'name of title of target column
    CatofTarget = ThisWorkbook.Worksheets("Countermeasures").ListObjects("Tbl_Counter").ListColumns("Category").DataBodyRange(DataBodyRowofTarget, 1)       'category or target cell row
            
        For Each Tag In Tags        'cycle through list of tag columns
            If Intersect(Target, ThisWorkbook.Worksheets("Countermeasures").ListObjects("Tbl_Counter").ListColumns(Tag).DataBodyRange) Is Nothing Then  'if target cell is within this column, continue
            
                Else        'if target cell is within this column, continue
                
                    With Target.Validation  'delete validation list of cell
                        .Delete
                    End With
            End If
        Next


        For Each identifier In identifiers
            If Intersect(Target, ThisWorkbook.Worksheets("Countermeasures").ListObjects("Tbl_Counter").ListColumns(identifier).DataBodyRange) Is Nothing Then
                Else
                
                    With Target.Validation
                        .Delete
                    End With
            End If
        Next
End If
End Sub


