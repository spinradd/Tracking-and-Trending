VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calendar_Frm 
   Caption         =   "Create Calendar"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4560
   OleObjectBlob   =   "Calendar_Frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Calendar_Frm"
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


Private Sub OKButton_Click()

Dim year_val As Long
Dim cat_val As String

    If Len(YearBox.value) < 4 Then
        MsgBox "Please enter 4 digits for the year, ex: '2021'"
        Exit Sub
    End If
    
    If CategoryTextBox.value = Empty Or CategoryTextBox.value = "" Then
        MsgBox "Please enter a category"
        Exit Sub
    End If
    
    year_val = CLng(YearBox.value)
    cat_val = CategoryTextBox.value
    
    Calendar_M.CalendarMakerv cat_val, year_val
        
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
Public Property Get YearBoxGet() As String

        YearBoxGet = Calendar_Frm.YearBox
        
End Property
Private Sub YearBox_Change_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
         
    Select Case KeyAscii
        Case Is < vbKey0, Is > vbKey9
            KeyAscii = 0
            Beep
    End Select
End Sub

Private Sub Userform_Initialize()
    
    YearBox.MaxLength = 4
    
    Dim counter_tbl As ListObject
    Set counter_tbl = Worksheets("Countermeasures").ListObjects("Tbl_Counter")
    Dim ArrBase() As Variant
    Dim Cell3 As Variant
    Dim counter_row_count As Variant
    Dim entry_count As Variant
    Dim item As Variant
    
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
        CategoryTextBox.AddItem ""
        Exit Sub
    End If
                
    ArrBase = BlankRemover(ArrBase)
    ArrBase = ArrayRemoveDups(ArrBase)
    
    For Each item In ArrBase
        CategoryTextBox.AddItem item
    Next
    
End Sub




