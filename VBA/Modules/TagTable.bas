Attribute VB_Name = "TagTable"
Sub CreateTagTable()

    Dim ListObj As ListObject
    Dim MyName As String
    Dim MyRangeString As String
    Dim MyListExists As Boolean
    Dim TagTable As ListObject

    MyName = "TagLookupTable"
    MyRangeString = "E3:L3"

    MyListExists = False
    For Each ListObj In Sheets("Lookup Tag").ListObjects
       
        If ListObj.Name = MyName Then MyListExists = True
        
    Next ListObj
    
    If Not (MyListExists) Then
        Set TagTable = ActiveSheet.ListObjects.Add(xlSrcRange, Range(MyRangeString), , xlYes)
        TagTable.Name = MyName
    End If
    
    Dim Tag1_val As String

Tag1_val = Sheets("Lookup Tag").Tag1.value
Tag2_val = Sheets("Lookup Tag").Tag2.value
Tag3_val = Sheets("Lookup Tag").Tag3.value
Tag4_val = Sheets("Lookup Tag").Tag4.value

'Debug.Print ActiveSheet.ListObjects(MyName).HeaderRowRange.Address

Set TagTable = Sheets("Lookup Tag").ListObjects(MyName)

TagTable.HeaderRowRange(1, 1) = Tag1_val
TagTable.HeaderRowRange(1, 2) = "Count 1"
TagTable.HeaderRowRange(1, 3) = Tag2_val
TagTable.HeaderRowRange(1, 4) = "Count 2"
TagTable.HeaderRowRange(1, 5) = Tag3_val
TagTable.HeaderRowRange(1, 6) = "Count 3"
TagTable.HeaderRowRange(1, 7) = Tag4_val
TagTable.HeaderRowRange(1, 8) = "Count 4"


Dim catval As Variant
catval = Sheets("Lookup Tag").Pivot_DD_Box.value
'Debug.Print CatVal


Dim counter_tbl As ListObject
   Set counter_sheet = Worksheets("Countermeasures")
   Set counter_tbl = counter_sheet.ListObjects("Tbl_Counter")
   Set counter_col = counter_tbl.ListColumns
   
'' 4 categories in lookup table

Dim Arr1() As Variant
Dim Arr2() As Variant
Dim Arr3() As Variant
Dim Arr4() As Variant

Dim ArrBase() As Variant
Dim col_val As String
entry_count = 0
 col_val = Tag1_val


    For Each cell In counter_tbl.ListColumns(col_val).DataBodyRange
        CRow = cell.row
        i = 0
        'Debug.Print cRow
        'Debug.Print counter_tbl.ListColumns("Category").DataBodyRange.Cells(cRow, 1).value
        Debug.Print counter_tbl.ListColumns("Category").DataBodyRange.Cells(CRow, 1).value
        
            If counter_tbl.ListColumns("Category").DataBodyRange.Cells(CRow, 1).value = catval And cell <> Empty Then
                If entry_count = 0 Then
                ReDim Preserve ArrBase(1)
                ArrBase(0) = cell
                'For all subsequent entries
                Else
                    If ArrBase(UBound(ArrBase)) <> Empty Then
                    ReDim Preserve ArrBase(UBound(ArrBase) + 1)
                    ArrBase(UBound(ArrBase)) = cell
                    Else
                    ArrBase(entry_count) = cell
                    
                    End If
                End If
                entry_count = entry_count + 1
                MsgBox Join(ArrBase, vbCrLf)
            End If
    
    Next cell
    
    'Select Case x
          '  Case 1
           ' Arr1 = ArrBase
       ' Case 2
        '    Arr2 = ArrBase
      '  Case 3
           ' Arr3 = ArrBase
       ' Case 4
         '   Arr4 = ArrBase
       ' End Select

Debug.Print "ph"

End Sub



