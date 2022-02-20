Attribute VB_Name = "Functions_M"

Public Function DuplicateCountToScript(nums As Variant) As Scripting.Dictionary 'nums = 1D array
'1D array is turned into a dictionary with its Key as the contents of the array,
'and the item as the number of times the key occured in the original array, returns a dictionary
    Dim Dict As New Scripting.Dictionary
    For Each num In nums        'for each key in dictionary
        If Dict.Exists(num) Then    'if it exists, add count to dictionary
            Dict(num) = Dict(num) + 1
        Else
            Dict(num) = 1
        End If
    Next

    Set DuplicateCountToScript = Dict
End Function
Function ArrayRemoveDups(MyArray As Variant) As Variant 'MyArray= 1 '1D Arrays
'1D function takes 1D array and removes duplicates, returns as 1D array
    Dim nFirst As Long, nLast As Long, i As Long
    Dim item As String
    
    Dim arrTemp() As Variant
    Dim Coll As New Collection
 
    'Get First and Last Array Positions
    nFirst = LBound(MyArray)
    nLast = UBound(MyArray)
    ReDim arrTemp(nFirst To nLast)
 
    'Convert Array to String
    For i = nFirst To nLast
        arrTemp(i) = CStr(MyArray(i))
    Next i
    
    'Populate Temporary Collection
    On Error Resume Next
    For i = nFirst To nLast
        Coll.Add arrTemp(i), arrTemp(i)
    Next i
    Err.Clear
    On Error GoTo 0
 
    'Resize Array
    nLast = Coll.Count + nFirst - 1
    ReDim arrTemp(nFirst To nLast)
    
    'Populate Array
    For i = nFirst To nLast
        arrTemp(i) = Coll(i - nFirst + 1)
    Next i
    
    'Output Array
    ArrayRemoveDups = arrTemp
 
End Function

Public Function GetUniqueValues(ByVal values As Variant) As Collection '1D Arrays
    
    'takes a range and outputs the range (values) of distinct unique values
    
    Dim result As Collection
    Dim cellValue As Variant
    Dim cellValueTrimmed As String

    Set result = New Collection
    Set GetUniqueValues = result

    On Error Resume Next

    For Each cellValue In values
        cellValueTrimmed = Trim(cellValue)
        If cellValueTrimmed = "" Then GoTo NextValue
        result.Add cellValueTrimmed, cellValueTrimmed
NextValue:
    Next cellValue

    On Error GoTo 0
End Function
Public Function SortDictionaryByValue(Dict As Object _
                    , Optional sortorder As XlSortOrder = xlAscending) As Object '1D Arrays
'function takes a dictionary and sorts dictionary by frequency of occurance, or alphabetical
    
    On Error GoTo eh
    
    Dim arrayList As Object
    Set arrayList = CreateObject("System.Collections.ArrayList")
    
    Dim dictTemp As Object
    Set dictTemp = CreateObject("Scripting.Dictionary")
   
    ' Put values in ArrayList and sort
    ' Store values in tempDict with their keys as a collection
    Dim key As Variant, value As Variant, Coll As Collection
    For Each key In Dict
    
        value = Dict(key)
        
        ' if the value doesn't exist in dict then add
        If dictTemp.Exists(value) = False Then
            ' create collection to hold keys
            ' - needed for duplicate values
            Set Coll = New Collection
            dictTemp.Add value, Coll
            
            ' Add the value
            arrayList.Add value
            
        End If
        
        ' Add the current key to the collection
        dictTemp(value).Add key
    
    Next key
    
    ' Sort the value
    arrayList.Sort
    
    ' Reverse if descending
    If sortorder = xlDescending Then
        arrayList.Reverse
    End If
    
    Dict.RemoveAll
    
    ' Read through the ArrayList and add the values and corresponding
    ' keys from the dictTemp
    Dim item As Variant
    For Each value In arrayList
        Set Coll = dictTemp(value)
        For Each item In Coll
            Dict.Add item, value
        Next item
    Next value
    
    Set arrayList = Nothing
    
    ' Return the new dictionary
    Set SortDictionaryByValue = Dict
        
Done:
    Exit Function
eh:
    If Err.Number = 450 Then
        Err.Raise vbObjectError + 100, "SortDictionaryByValue" _
                , "Cannot sort the dictionary if the value is an object"
    End If
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

Function IsWorkBookOpen(FileName As String)
'check to see if file is open
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open FileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else: Error ErrNo
    End Select
End Function

Public Function GetWorkbook(ByVal sFullName As String) As Workbook
'function to activate workbook
    Dim sFile As String
    Dim wbReturn As Workbook

    sFile = Dir(sFullName)

    On Error Resume Next
        Set wbReturn = Workbooks(sFile)

        If wbReturn Is Nothing Then
            Set wbReturn = Workbooks.Open(sFullName)
        End If
    On Error GoTo 0

    Set GetWorkbook = wbReturn

End Function

Public Function FSOGetFileName(Path As String, Extension As Boolean)
    Dim FileName As String
    Dim FSO As New FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
 
    If Extension = True Then
    'Get File Name
    FileName = FSO.GetFileName(Path)
    Else
    'Get File Name no Extension
    FileName = FSO.GetFileName(Path)
    FileName = left(FileName, InStr(FileName, ".") - 1)
    End If
    
    FSOGetFileName = FileName
    
 
End Function

Public Function IsLoaded(formName As String) As Boolean
Dim frm As Object
For Each frm In VBA.UserForms
    If frm.Name = formName Then
        IsLoaded = True
        Exit Function
    End If
Next frm
IsLoaded = False
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






