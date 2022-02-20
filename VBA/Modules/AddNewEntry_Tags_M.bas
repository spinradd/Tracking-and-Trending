Attribute VB_Name = "AddNewEntry_Tags_M"
Sub IssueTier1Box() 'sub to populate corresponding drop down box in "Add New Entry" form


 Dim counter_tbl As ListObject
    Set counter_tbl = Worksheets("Countermeasures").ListObjects("Tbl_Counter")
    Dim ArrBase() As Variant
    Dim Cell3 As Variant
    Dim counter_row_count As Variant
    Dim entry_count As Variant
    Dim item As Variant
    
    
    AddNewEntry_Tags.IssueTier1Box.Clear

counter_row_count = 0
    'for loop creates an array of cell values from the ListColumn
For Each Cell3 In counter_tbl.ListColumns("Issue Tier 1 Tag").DataBodyRange
                    counter_row_count = counter_row_count + 1
                    
                    'If condition filters array to include only entries where selected ctageory from userform drop down exists
            If AddNewEntry.CategoryTextBox.value = counter_tbl.ListColumns("Category").DataBodyRange(counter_row_count, 1).value And Cell3 <> Empty Then
                    
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
    
        
    'if all cells are empty or blank, just add "" to the text box, ignore array
    If (Not Not ArrBase) = 0 Then
        AddNewEntry_Tags.IssueTier1Box.AddItem ""
        Exit Sub
    End If
                
    'remove blanks and duplicates from array
    ArrBase = BlankRemover(ArrBase)
    ArrBase = ArrayRemoveDups(ArrBase)
    
    'add each new item to the drop down box
    For Each item In ArrBase
        AddNewEntry_Tags.IssueTier1Box.AddItem item
    Next

End Sub

Sub IssueTier2Box()


 Dim counter_tbl As ListObject
    Set counter_tbl = Worksheets("Countermeasures").ListObjects("Tbl_Counter")
    Dim ArrBase() As Variant
    Dim Cell3 As Variant
    Dim counter_row_count As Variant
    Dim entry_count As Variant
    Dim item As Variant
    
    
    AddNewEntry_Tags.IssueTier2Box.Clear

counter_row_count = 0
For Each Cell3 In counter_tbl.ListColumns("Issue Tier 2 Tag").DataBodyRange
                    counter_row_count = counter_row_count + 1
                    
            If AddNewEntry.CategoryTextBox.value = counter_tbl.ListColumns("Category").DataBodyRange(counter_row_count, 1).value And Cell3 <> Empty Then
                    
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
                
    If (Not Not ArrBase) = 0 Then
        AddNewEntry_Tags.IssueTier2Box.AddItem ""
        Exit Sub
    End If
                
    ArrBase = BlankRemover(ArrBase)
    ArrBase = ArrayRemoveDups(ArrBase)
    
    
    For Each item In ArrBase
        AddNewEntry_Tags.IssueTier2Box.AddItem item
    Next

End Sub
Sub CauseCatBox()

 Dim counter_tbl As ListObject
    Set counter_tbl = Worksheets("Countermeasures").ListObjects("Tbl_Counter")
    Dim ArrBase() As Variant
    Dim Cell3 As Variant
    Dim counter_row_count As Variant
    Dim entry_count As Variant
    Dim item As Variant
    
    
    AddNewEntry_Tags.CauseCatBox.Clear

counter_row_count = 0
For Each Cell3 In counter_tbl.ListColumns("Cause Category").DataBodyRange
                    counter_row_count = counter_row_count + 1
                        
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
            
Next Cell3
                
    If (Not Not ArrBase) = 0 Then
        AddNewEntry_Tags.CauseCatBox.AddItem ""
        Exit Sub
    End If
                
    ArrBase = BlankRemover(ArrBase)
    ArrBase = ArrayRemoveDups(ArrBase)
    
    
    For Each item In ArrBase
        AddNewEntry_Tags.CauseCatBox.AddItem item
    Next

End Sub
Sub CauseDetBox()


 Dim counter_tbl As ListObject
    Set counter_tbl = Worksheets("Countermeasures").ListObjects("Tbl_Counter")
    Dim ArrBase() As Variant
    Dim Cell3 As Variant
    Dim counter_row_count As Variant
    Dim entry_count As Variant
    Dim item As Variant
    
    
    AddNewEntry_Tags.CauseDetBox.Clear

counter_row_count = 0
For Each Cell3 In counter_tbl.ListColumns("Cause Detail").DataBodyRange
                    counter_row_count = counter_row_count + 1
                        
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
            
Next Cell3
                
    If (Not Not ArrBase) = 0 Then
        AddNewEntry_Tags.CauseDetBox.AddItem ""
        Exit Sub
    End If
                
    ArrBase = BlankRemover(ArrBase)
    ArrBase = ArrayRemoveDups(ArrBase)
    
    
    For Each item In ArrBase
        AddNewEntry_Tags.CauseDetBox.AddItem item
    Next

End Sub

Sub BatchBox()

 Dim counter_tbl As ListObject
    Set counter_tbl = Worksheets("Countermeasures").ListObjects("Tbl_Counter")
    Dim ArrBase() As Variant
    Dim Cell3 As Variant
    Dim counter_row_count As Variant
    Dim entry_count As Variant
    Dim item As Variant
    Dim cell_counter As Variant
    Dim cen_entries() As Variant
    
    
    AddNewEntry_Tags.BatchBox.Clear

counter_row_count = 0

'slightly different for loop for creating array for drop down box:
'cells often contain multiple entries in these columns:
' Ex: "batch1; batch2
'Loop seperates entries via determinant: ";"
'Creates array the same way as previous scripts
For Each cell In counter_tbl.ListColumns("Batch").DataBodyRange
                            cel_entries = Split(cell.value, "; ")       'seperate cells contents into individual entries
                            counter_row_count = counter_row_count + 1   'add row count to know to move onto next cell
                            If cell <> Empty Then                       'for each entry, check if first entry, if so redim array to number of cell entries
                                                                            ' if not, for each entry extend array by one and add entry
                                For Each Entry In cel_entries           'for each individual entry in cell
                                    If entry_count = 0 Then             'if this is the first entry in the array, we must dimthe array
                                        ReDim Preserve ArrBase(UBound(cel_entries))
                                        ArrBase(i) = Entry              'add entry to array
                                        'For all subsequent entries
                                        entry_count = entry_count + 1   'increase count of array total
                                        Else                                            'if not first entry:
                                            If ArrBase(UBound(ArrBase)) <> Empty Then   'if end of array is not empty:
                                            ReDim Preserve ArrBase(UBound(ArrBase) + 1)     'increase size
                                            ArrBase(UBound(ArrBase)) = Entry                'add entry to extra slot
                                            entry_count = entry_count + 1                   'increase count of array total
                                            Else                                        'if end of array is empty:
                                            ArrBase(entry_count) = Entry                    'addentry to extra slot
                                            entry_count = entry_count + 1                   'increase count of array total
                                            End If
                                    End If
                            
                                Next
                            End If
Next


                
    If (Not Not ArrBase) = 0 Then
        AddNewEntry_Tags.BatchBox.AddItem ""
        Exit Sub
    End If
                
    ArrBase = BlankRemover(ArrBase)
    ArrBase = ArrayRemoveDups(ArrBase)
    
    
    For Each item In ArrBase
        AddNewEntry_Tags.BatchBox.AddItem item
    Next

End Sub
Sub PrimaryEquiptmentBox()

 Dim counter_tbl As ListObject
    Set counter_tbl = Worksheets("Countermeasures").ListObjects("Tbl_Counter")
    Dim ArrBase() As Variant
    Dim Cell3 As Variant
    Dim counter_row_count As Variant
    Dim entry_count As Variant
    Dim item As Variant
    Dim cell_counter As Variant
    Dim cen_entries() As Variant
    
    
    AddNewEntry_Tags.PrimaryEquiptmentBox.Clear

counter_row_count = 0

For Each cell In counter_tbl.ListColumns("Primary Equipment").DataBodyRange
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
                                        ReDim Preserve ArrBase(UBound(cel_entries))
                                        ArrBase(i) = Entry
                                        'For all subsequent entries
                                        entry_count = entry_count + 1
                                        Else
                                            Debug.Print UBound(ArrBase)
                                            If ArrBase(UBound(ArrBase)) <> Empty Then
                                            ReDim Preserve ArrBase(UBound(ArrBase) + 1)
                                            ArrBase(UBound(ArrBase)) = Entry
                                            entry_count = entry_count + 1
                                            Else
                                            ArrBase(entry_count) = Entry
                                            entry_count = entry_count + 1
                                            End If
                                    End If
                            
                                Next
                            End If
Next


                
    If (Not Not ArrBase) = 0 Then
        AddNewEntry_Tags.PrimaryEquiptmentBox.AddItem ""
        Exit Sub
    End If
                
    ArrBase = BlankRemover(ArrBase)
    ArrBase = ArrayRemoveDups(ArrBase)
    
    
    For Each item In ArrBase
        AddNewEntry_Tags.PrimaryEquiptmentBox.AddItem item
    Next

End Sub
Sub MfgStageBox()

 Dim counter_tbl As ListObject
    Set counter_tbl = Worksheets("Countermeasures").ListObjects("Tbl_Counter")
    Dim ArrBase() As Variant
    Dim Cell3 As Variant
    Dim counter_row_count As Variant
    Dim entry_count As Variant
    Dim item As Variant
    Dim cell_counter As Variant
    Dim cen_entries() As Variant
    
    
    AddNewEntry_Tags.MfgStageBox.Clear

counter_row_count = 0

For Each cell In counter_tbl.ListColumns("Manufacturing Stage").DataBodyRange
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
                                        ReDim Preserve ArrBase(UBound(cel_entries))
                                        ArrBase(i) = Entry
                                        'For all subsequent entries
                                        entry_count = entry_count + 1
                                        Else
                                            Debug.Print UBound(ArrBase)
                                            If ArrBase(UBound(ArrBase)) <> Empty Then
                                            ReDim Preserve ArrBase(UBound(ArrBase) + 1)
                                            ArrBase(UBound(ArrBase)) = Entry
                                            entry_count = entry_count + 1
                                            Else
                                            ArrBase(entry_count) = Entry
                                            entry_count = entry_count + 1
                                            End If
                                    End If
                            
                                Next
                            End If
Next


                
    If (Not Not ArrBase) = 0 Then
        AddNewEntry_Tags.MfgStageBox.AddItem ""
        Exit Sub
    End If
                
    ArrBase = BlankRemover(ArrBase)
    ArrBase = ArrayRemoveDups(ArrBase)
    
    
    For Each item In ArrBase
        AddNewEntry_Tags.MfgStageBox.AddItem item
    Next

End Sub

Sub QualityClassificationBox()

 Dim counter_tbl As ListObject
    Set counter_tbl = Worksheets("Countermeasures").ListObjects("Tbl_Counter")
    Dim ArrBase() As Variant
    Dim Cell3 As Variant
    Dim counter_row_count As Variant
    Dim entry_count As Variant
    Dim item As Variant
    Dim cell_counter As Variant
    Dim cen_entries() As Variant
    
    
    AddNewEntry_Tags.QualityClassificationBox.Clear

counter_row_count = 0

For Each cell In counter_tbl.ListColumns("Quality Classification").DataBodyRange
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
                                        ReDim Preserve ArrBase(UBound(cel_entries))
                                        ArrBase(i) = Entry
                                        'For all subsequent entries
                                        entry_count = entry_count + 1
                                        Else
                                            Debug.Print UBound(ArrBase)
                                            If ArrBase(UBound(ArrBase)) <> Empty Then
                                            ReDim Preserve ArrBase(UBound(ArrBase) + 1)
                                            ArrBase(UBound(ArrBase)) = Entry
                                            entry_count = entry_count + 1
                                            Else
                                            ArrBase(entry_count) = Entry
                                            entry_count = entry_count + 1
                                            End If
                                    End If
                            
                                Next
                            End If
Next


                
    If (Not Not ArrBase) = 0 Then
        AddNewEntry_Tags.QualityClassificationBox.AddItem ""
        Exit Sub
    End If
                
    ArrBase = BlankRemover(ArrBase)
    ArrBase = ArrayRemoveDups(ArrBase)
    
    
    For Each item In ArrBase
        AddNewEntry_Tags.QualityClassificationBox.AddItem item
    Next

End Sub

Sub SafetyTierBox()

 Dim counter_tbl As ListObject
    Set counter_tbl = Worksheets("Countermeasures").ListObjects("Tbl_Counter")
    Dim ArrBase() As Variant
    Dim Cell3 As Variant
    Dim counter_row_count As Variant
    Dim entry_count As Variant
    Dim item As Variant
    Dim cell_counter As Variant
    Dim cen_entries() As Variant
    
    
    AddNewEntry_Tags.SafetyTierBox.Clear

counter_row_count = 0

For Each cell In counter_tbl.ListColumns("Safety Tier").DataBodyRange
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
                                        ReDim Preserve ArrBase(UBound(cel_entries))
                                        ArrBase(i) = Entry
                                        'For all subsequent entries
                                        entry_count = entry_count + 1
                                        Else
                                            Debug.Print UBound(ArrBase)
                                            If ArrBase(UBound(ArrBase)) <> Empty Then
                                            ReDim Preserve ArrBase(UBound(ArrBase) + 1)
                                            ArrBase(UBound(ArrBase)) = Entry
                                            entry_count = entry_count + 1
                                            Else
                                            ArrBase(entry_count) = Entry
                                            entry_count = entry_count + 1
                                            End If
                                    End If
                            
                                Next
                            End If
Next


                
    If (Not Not ArrBase) = 0 Then
        AddNewEntry_Tags.SafetyTierBox.AddItem ""
        Exit Sub
    End If
                
    ArrBase = BlankRemover(ArrBase)
    ArrBase = ArrayRemoveDups(ArrBase)
    
    
    For Each item In ArrBase
        AddNewEntry_Tags.SafetyTierBox.AddItem item
    Next

End Sub

Sub ShowEntriesForm()
AddNewEntry_Tags.Show vbModeless    'show main add new entry form, modeless
    
AddNewEntry_Tags.Top = AddNewEntry.Top + (AddNewEntry_Tags.Height / 8) 'center userform height
AddNewEntry_Tags.left = AddNewEntry.left + (Application.UsableWidth / 3)    'center userform width

AddNewEntry.Show vmodeless      'show supplementary tags userform

AddNewEntry.Top = Application.Top + (AddNewEntry.Height / 8) 'center height
AddNewEntry.left = Application.left + (Application.UsableWidth / 3) - (AddNewEntry.Width / 2) 'center width with minor seperation between main form

End Sub

Sub EntryIdentifierBox()

 Dim counter_tbl As ListObject
    Set counter_tbl = Worksheets("Countermeasures").ListObjects("Tbl_Counter")
    Dim ArrBase() As Variant
    Dim Cell3 As Variant
    Dim counter_row_count As Variant
    Dim entry_count As Variant
    Dim item As Variant
    Dim cell_counter As Variant
    Dim cen_entries() As Variant
    
    
    AddNewEntry_Tags.EntryIdentifierBox.Clear

counter_row_count = 0

For Each cell In counter_tbl.ListColumns("Entry Identifier").DataBodyRange
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
                                        ReDim Preserve ArrBase(UBound(cel_entries))
                                        ArrBase(i) = Entry
                                        'For all subsequent entries
                                        entry_count = entry_count + 1
                                        Else
                                            Debug.Print UBound(ArrBase)
                                            If ArrBase(UBound(ArrBase)) <> Empty Then
                                            ReDim Preserve ArrBase(UBound(ArrBase) + 1)
                                            ArrBase(UBound(ArrBase)) = Entry
                                            entry_count = entry_count + 1
                                            Else
                                            ArrBase(entry_count) = Entry
                                            entry_count = entry_count + 1
                                            End If
                                    End If
                            
                                Next
                            End If
Next


                
    If (Not Not ArrBase) = 0 Then
        AddNewEntry_Tags.EntryIdentifierBox.AddItem ""
        Exit Sub
    End If
                
    ArrBase = BlankRemover(ArrBase)
    ArrBase = ArrayRemoveDups(ArrBase)
    
    
    For Each item In ArrBase
        AddNewEntry_Tags.EntryIdentifierBox.AddItem item
    Next

End Sub






