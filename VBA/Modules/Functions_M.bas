
Public Function Count_array(arr_to_count As Variant) As Scripting.Dictionary 'nums = 1D array
'1D array is turned into a dictionary with its Key as the contents of the array,
'and the item as the number of times the key occured in the original array, returns a dictionary
    Dim new_dict As New Scripting.Dictionary
    For Each num In arr_to_count        'for each key in dictionary
        If new_dict.Exists(num) Then    'if it exists, add count to dictionary
            new_dict(num) = new_dict(num) + 1
        Else
            new_dict(num) = 1
        End If
    Next

    Set Count_array = new_dict
End Function

Function Remove_duplicates_array(arr_with_dups As Variant) As Variant 'arr_with_dups= 1 '1D Arrays
'1D function takes 1D array and removes duplicates, returns as 1D array
    
    
    Dim first_item As Long, last_item As Long, i As Long
    Dim item As String
    
    Dim intermediate_arr() As Variant
    Dim coll As New Collection
 
    'find length of current array; resize intermediary array to fit
    first_item = LBound(arr_with_dups)
    last_item = UBound(arr_with_dups)
    ReDim intermediate_arr(first_item To last_item)
 
    'normalize to string type
    For i = first_item To last_item
        intermediate_arr(i) = CStr(arr_with_dups(i))
    Next i
    
    'add items to collection
    On Error Resume Next
    For i = first_item To last_item
        coll.Add intermediate_arr(i), intermediate_arr(i)
    Next i
    Err.Clear
    On Error GoTo 0
 
    'resize new array
    last_item = coll.Count + first_item - 1
    ReDim intermediate_arr(first_item To last_item)
    
    'for new array
    For i = first_item To last_item
        intermediate_arr(i) = coll(i - first_item + 1)
    Next i
    
    Remove_duplicates_array = intermediate_arr
 
End Function

Public Function Create_unique_arr(ByVal original_arr As Variant) As Collection '1D Arrays
    
    'takes a range and outputs the range (original_arr) of distinct unique original_arr
    
    Dim unique_arr As Collection
    Dim raw_val As Variant
    Dim individual_val As String

    Set unique_arr = New Collection
    Set Create_unique_arr = unique_arr

    On Error Resume Next

    For Each raw_val In original_arr
        
        individual_val = Trim(raw_val)
        
        If individual_val = "" Then GoTo NextValue
        
        Else:
            unique_arr.Add individual_val, individual_val

NextValue:
    Next raw_val

    On Error GoTo 0
End Function
Function BlankRemover(ArrayToCondense As Variant) As Variant() '1D Arrays
'takes 1D arrays and removesblanks, returns as 1D array

Dim ArrayWithoutBlanks() As Variant
Dim CellsInArray As Variant
ReDim ArrayWithoutBlanks(0 To 0) As Variant

IsAllBlank = True
For Each CellsInArray In ArrayToCondense
    If CellsInArray <> "" Then
        IsAllBlank = False
        ArrayWithoutBlanks(UBound(ArrayWithoutBlanks)) = CellsInArray
        ReDim Preserve ArrayWithoutBlanks(0 To UBound(ArrayWithoutBlanks) + 1)
    End If
    
    'MsgBox Join(ArrayWithoutBlanks, vbCrLf)

Next CellsInArray


'get rid of extra blank space
If IsAllBank = False Then
        Do While ArrayWithoutBlanks(UBound(ArrayWithoutBlanks)) = Empty
            ReDim Preserve ArrayWithoutBlanks(0 To UBound(ArrayWithoutBlanks) - 1)
        Loop
        BlankRemover = ArrayWithoutBlanks
    
    Else
    
        ReDim ArrayWithoutBlanks(0 To 0)
        BlankRemover = ArrayWithoutBlanks
End If


End Function
Public Sub SortDictionary(Dict As Scripting.Dictionary, _
    SortByKey As Boolean, _
    Optional Descending As Boolean = False, _
    Optional CompareMode As VbCompareMethod = vbTextCompare)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SortDictionary
' This sorts a Dictionary object. If SortByKey is False, the
' the sort is done based on the Items of the Dictionary, and
' these items must be simple data types. They may not be
' Object, Arrays, or User-Defined Types. If SortByKey is True,
' the Dictionary is sorted by Key value, and the Items in the
' Dictionary may be Object as well as simple variables.
'
' If sort by key is True, all element of the Dictionary
' must have a non-blank Key value. If Key is vbNullString
' the procedure will terminate.
'
' By defualt, sorting is done in Ascending order. You can
' sort by Descending order by setting the Descending parameter
' to True.
'
' By default, text comparisons are done case-INSENSITIVE (e.g.,
' "a" = "A"). To use case-SENSITIVE comparisons (e.g., "a" <> "A")
' set CompareMode to vbBinaryCompare.
'
' Note: This procedure requires the
' QSortInPlace function, which is described and available for
' download at www.cpearson.com/excel/qsort.htm .
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Ndx As Long
Dim KeyValue As String
Dim ItemValue As Variant
Dim Arr() As Variant
Dim KeyArr() As String
Dim VTypes() As VbVarType


Dim V As Variant
Dim SplitArr As Variant

Dim TempDict As Scripting.Dictionary
'''''''''''''''''''''''''''''
' Ensure Dict is not Nothing.
'''''''''''''''''''''''''''''
If Dict Is Nothing Then
    Exit Sub
End If
''''''''''''''''''''''''''''
' If the number of elements
' in Dict is 0 or 1, no
' sorting is required.
''''''''''''''''''''''''''''
If (Dict.Count = 0) Or (Dict.Count = 1) Then
    Exit Sub
End If

''''''''''''''''''''''''''''
' Create a new TempDict.
''''''''''''''''''''''''''''
Set TempDict = New Scripting.Dictionary

If SortByKey = True Then
    ''''''''''''''''''''''''''''''''''''''''
    ' We're sorting by key. Redim the Arr
    ' to the number of elements in the
    ' Dict object, and load that array
    ' with the key names.
    ''''''''''''''''''''''''''''''''''''''''
    ReDim Arr(0 To Dict.Count - 1)
    
    For Ndx = 0 To Dict.Count - 1
        Arr(Ndx) = Dict.Keys(Ndx)
    Next Ndx
    
    ''''''''''''''''''''''''''''''''''''''
    ' Sort the key names.
    ''''''''''''''''''''''''''''''''''''''
    QSortInPlace InputArray:=Arr, LB:=-1, UB:=-1, Descending:=Descending, CompareMode:=CompareMode
    ''''''''''''''''''''''''''''''''''''''''''''
    ' Load TempDict. The key value come from
    ' our sorted array of keys Arr, and the
    ' Item comes from the original Dict object.
    ''''''''''''''''''''''''''''''''''''''''''''
    For Ndx = 0 To Dict.Count - 1
        KeyValue = Arr(Ndx)
        TempDict.Add key:=KeyValue, item:=Dict.item(KeyValue)
    Next Ndx
    '''''''''''''''''''''''''''''''''
    ' Set the passed in Dict object
    ' to our TempDict object.
    '''''''''''''''''''''''''''''''''
    Set Dict = TempDict
    ''''''''''''''''''''''''''''''''
    ' This is the end of processing.
    ''''''''''''''''''''''''''''''''
Else
    '''''''''''''''''''''''''''''''''''''''''''''''
    ' Here, we're sorting by items. The Items must
    ' be simple data types. They may NOT be Objects,
    ' arrays, or UserDefineTypes.
    ' First, ReDim Arr and VTypes to the number
    ' of elements in the Dict object. Arr will
    ' hold a string containing
    '   Item & vbNullChar & Key
    ' This keeps the association between the
    ' item and its key.
    '''''''''''''''''''''''''''''''''''''''''''''''
    ReDim Arr(0 To Dict.Count - 1)
    ReDim VTypes(0 To Dict.Count - 1)

    For Ndx = 0 To Dict.Count - 1
        If (IsObject(Dict.Items(Ndx)) = True) Or _
            (IsArray(Dict.Items(Ndx)) = True) Or _
            VarType(Dict.Items(Ndx)) = vbUserDefinedType Then
            Debug.Print "***** ITEM IN DICTIONARY WAS OBJECT OR ARRAY OR UDT"
            Exit Sub
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Here, we create a string containing
        '       Item & vbNullChar & Key
        ' This preserves the associate between an item and its
        ' key. Store the VarType of the Item in the VTypes
        ' array. We'll use these values later to convert
        ' back to the proper data type for Item.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Arr(Ndx) = Dict.Items(Ndx) & vbNullChar & Dict.Keys(Ndx)
            VTypes(Ndx) = VarType(Dict.Items(Ndx))
            
    Next Ndx
    ''''''''''''''''''''''''''''''''''
    ' Sort the array that contains the
    ' items of the Dictionary along
    ' with their associated keys
    ''''''''''''''''''''''''''''''''''
    QSortInPlace InputArray:=Arr, LB:=-1, UB:=-1, Descending:=Descending, CompareMode:=vbTextCompare
    
    For Ndx = LBound(Arr) To UBound(Arr)
        '''''''''''''''''''''''''''''''''''''
        ' Loop trhogh the array of sorted
        ' Items, Split based on vbNullChar
        ' to get the Key from the element
        ' of the array Arr.
        SplitArr = Split(Arr(Ndx), vbNullChar)
        ''''''''''''''''''''''''''''''''''''''''''
        ' It may have been possible that item in
        ' the dictionary contains a vbNullChar.
        ' Therefore, use UBound to get the
        ' key value, which will necessarily
        ' be the last item of SplitArr.
        ' Then Redim Preserve SplitArr
        ' to UBound - 1 to get rid of the
        ' Key element, and use Join
        ' to reassemble to original value
        ' of the Item.
        '''''''''''''''''''''''''''''''''''''''''
        KeyValue = SplitArr(UBound(SplitArr))
        ReDim Preserve SplitArr(LBound(SplitArr) To UBound(SplitArr) - 1)
        ItemValue = Join(SplitArr, vbNullChar)
        '''''''''''''''''''''''''''''''''''''''
        ' Join will set ItemValue to a string
        ' regardless of what the original
        ' data type was. Test the VTypes(Ndx)
        ' value to convert ItemValue back to
        ' the proper data type.
        '''''''''''''''''''''''''''''''''''''''
        Select Case VTypes(Ndx)
            Case vbBoolean
                ItemValue = CBool(ItemValue)
            Case vbByte
                ItemValue = CByte(ItemValue)
            Case vbCurrency
                ItemValue = CCur(ItemValue)
            Case vbDate
                ItemValue = CDate(ItemValue)
            Case vbDecimal
                ItemValue = CDec(ItemValue)
            Case vbDouble
                ItemValue = CDbl(ItemValue)
            Case vbInteger
                ItemValue = CInt(ItemValue)
            Case vbLong
                ItemValue = CLng(ItemValue)
            Case vbSingle
                ItemValue = CSng(ItemValue)
            Case vbString
                ItemValue = CStr(ItemValue)
            Case Else
                ItemValue = ItemValue
        End Select
        ''''''''''''''''''''''''''''''''''''''
        ' Finally, add the Item and Key to
        ' our TempDict dictionary.
        
        TempDict.Add key:=KeyValue, item:=ItemValue
    Next Ndx
End If


'''''''''''''''''''''''''''''''''
' Set the passed in Dict object
' to our TempDict object.
'''''''''''''''''''''''''''''''''
Set Dict = TempDict
End Sub


Public Function Is_form_open(form_name As String) As Boolean

Dim form As Object

Is_form_open = False

For Each form In VBA.UserForms
    If form.Name = form_name Then
        Is_form_open = True
    End If
Next form

End Function

Public Function DoesArrayExist(Arr() As Variant) As Variant

Dim ArrayResults(0 To 1) As Variant

DoesArrExist = True
IsArrEmpty = False


If (Not Not Arr) = 0 Then       'if array never intitialized then it doesn't exist,
        DoesArrExist = False
        IsArrEmpty = True         'if list doesn't exist, then it's empty (by default)
    Else
        For Each item_in_array In Arr       'test if array is empty. If there is one non-blank cell, change bool value
            If item_in_array <> Empty Then
                IsArrEmpty = False
            End If
        Next
End If

ArrayResults(0) = DoesArrExist
ArrayResults(1) = IsArrEmpty

DoesArrayExist = ArrayResults
End Function


Public Function GetArrayfromTable(SheetName As String, ListObject_Name As String, _
                                    Column_Name As String)
'function needs sheetname, listobject name, column header title
'will create a 1D array from table data body range, filled with blanks and duplicate values
    Dim ArrBase() As Variant
    entry_count = 0
    For Each cell In Worksheets(SheetName).ListObjects(ListObject_Name).ListColumns(Column_Name).DataBodyRange
    
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
    Next cell
    
    GetArrayfromTable = ArrBase

End Function

Public Function GenerateRandomInt(upperbound As Long, lowerbound As Long)

    x = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
    
    GenerateRandomInt = x
End Function

Function MonthDays(date_test As Date)
    MonthDays = Day(DateSerial(Year(date_test), Month(date_test) + 1, 1) - 1)
End Function






