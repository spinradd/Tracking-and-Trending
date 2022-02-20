VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PivotTable_Frm 
   Caption         =   "Create Trend Pivot Table & Chart"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4560
   OleObjectBlob   =   "PivotTable_Frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PivotTable_Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private m_Cancelled As Boolean

' Returns the cancelled value to the calling procedure
Public Property Get Cancelled() As Boolean
    Cancelled = m_Cancelled
End Property

Private Sub CancelButton_Click()
    Hide
    m_Cancelled = True
End Sub

Private Sub FinalMonthTxt_Click()

End Sub

Private Sub OKButton_Click()
    Dim field As String
    Dim category As String
    Dim frequency As Long
    
    field = FieldBox.value
    category = CategoryBox.value
    frequency = FrequencyBox.value
    
    If FieldBox = Empty Or CategoryBox = Empty Or FrequencyBox = Empty Then
        MsgBox "Please make sure all values are selected"
    Exit Sub
    End If
    
    If PivotTable_Frm.TrendSummaryRadio.value = False And PivotTable_Frm.RunningTotalRadio.value = False Then
        MsgBox "Please make sure one chart type is selected"
    Exit Sub
    End If
    
    If PivotTable_Frm.TrendSummaryRadio.value = True Then
        Pivot_Table_M.Create_Pivot_Trend category, field, frequency
    End If
    
    If PivotTable_Frm.RunningTotalRadio.value = True Then
        Pivot_Table_M.Create_Pivot_Running_Total category, field, frequency
    End If
    
    
    'Identifier_Updater_M.CreateTagDescriptorSheet start_month, end_month
     
        Hide
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer _
                                       , CloseMode As Integer)
    
    ' Prevent the form being unloaded
    If CloseMode = vbFormControlMenu Then Cancel = True
    
    ' Hide the Userform and set cancelled to true
    Hide
    m_Cancelled = True
    
End Sub

Public Property Get CategoryBoxGet() As String

        CategoryBoxGet = PivotTable_Frm.CategoryBox
        
End Property
Public Property Get FieldBoxGet() As String

        FieldBoxGet = PivotTable_Frm.FieldBox
        
End Property
Public Property Get FrequencyBoxGet() As String

        FrequencyBoxGet = PivotTable_Frm.FrequencyBox
        
End Property

Private Sub FrequencyBox_Change_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
         
    Select Case KeyAscii
        Case Is < vbKey0, Is > vbKey9
            KeyAscii = 0
            Beep
    End Select
End Sub

Private Sub Userform_Initialize()

    FrequencyBox.MaxLength = 5
    
    Dim counter_tbl As ListObject
    Set counter_tbl = Worksheets("Countermeasures").ListObjects("Tbl_Counter")
    Dim ArrBase() As Variant
    Dim NewBase() As Variant
    Dim Cell3 As Variant
    Dim cell As Variant
    Dim counter_row_count As Variant
    Dim entry_count As Variant
    Dim item As Variant
    
    
    entry_count = 0
    counter_row_count = 0
    For Each Cell3 In counter_tbl.ListColumns("Category").DataBodyRange
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
        CategoryBox.AddItem ""
        Exit Sub
    End If
                
    ArrBase = BlankRemover(ArrBase)
    ArrBase = ArrayRemoveDups(ArrBase)
    
    
    For Each item In ArrBase
        CategoryBox.AddItem item
    Next
    
    entry_count = 0
    For Each cell In counter_tbl.HeaderRowRange
                                       'if first entry, redim to hold one spot "(0)"
                                       If entry_count = 0 Then
                                           ReDim Preserve NewBase(0)
                                           NewBase(0) = cell
                                        'For all subsequent entries extend array by 1 and enter contents in cell
                                       Else
                                           ReDim Preserve NewBase(UBound(NewBase) + 1)
                                           NewBase(UBound(NewBase)) = cell
                                       End If
                                       entry_count = entry_count + 1
    Next cell
                
    If (Not Not NewBase) = 0 Then
        FieldBox.AddItem ""
        Exit Sub
    End If
                
    NewBase = BlankRemover(NewBase)
    NewBase = ArrayRemoveDups(NewBase)
    
    
    For Each item In NewBase
        FieldBox.AddItem item
    Next
        
End Sub


