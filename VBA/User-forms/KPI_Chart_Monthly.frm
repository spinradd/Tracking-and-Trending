VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} KPI_Chart_Monthly 
   Caption         =   "Create KPI summary for a specific month"
   ClientHeight    =   4185
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   5610
   OleObjectBlob   =   "KPI_Chart_Monthly.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "KPI_Chart_Monthly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Cancelled As Boolean

Private Sub Userform_Initialize()

    Dim x As Variant
    Dim Month_Name As String
    Dim MonthArr(0 To 11) As Variant
    For x = 1 To 12
        MonthArr(x - 1) = MonthName(x)
        MonthBox.AddItem MonthArr(x - 1)
    Next
    
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



Private Sub UserForm_QueryClose(Cancel As Integer _
                                       , CloseMode As Integer)
    
    ' Prevent the form being unloaded
    If CloseMode = vbFormControlMenu Then Cancel = True
    
    ' Hide the Userform and set cancelled to true
    Hide
    m_Cancelled = True
    
End Sub
Private Sub YearEntry_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
         
    Select Case KeyAscii
        Case Is < vbKey0, Is > vbKey9
            KeyAscii = 0
            Beep
    End Select
End Sub
Public Property Get start_year() As String

        start_year = TandDMonthly.YearEntry
        
End Property
Public Property Get start_month() As String

        start_month = TandDMonthly.MonthBox.value
        
End Property

Private Sub EntriesPerSlide_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
         
    Select Case KeyAscii
        Case Is < vbKey0, Is > vbKey9
            KeyAscii = 0
            Beep
    End Select
End Sub
Private Sub FontSize_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
         
    Select Case KeyAscii
        Case Is < vbKey0, Is > vbKey9
            KeyAscii = 0
            Beep
    End Select
End Sub

Private Sub OKButton_Click()
Dim start_month As Date
Dim end_month As Date
Dim item As Variant
Dim x As Variant
Dim Month_Name As String
Dim MonthArr(0 To 11) As Variant
Dim StartMonthNameExists As Variant

    For x = 1 To 12
        MonthArr(x - 1) = MonthName(x)
    Next
    
    StartMonthNameExists = False
    For Each item In MonthArr
        If MonthBox.value = item Then
            StartMonthNameExists = True
        End If
    Next
    
    If YearEntry.value = Empty Or MonthBox.value = Empty Then
        MsgBox "Please select a year."
        Exit Sub
    End If
    
    If StartMonthNameExists = False Then
        MsgBox "Please write the full name of the month in the start month box"
            Exit Sub
    End If
    
    If CategoryTextBox.value = Empty Or CategoryTextBox.value = "" Then
        MsgBox "Please enter a category"
        Exit Sub
    End If
    
        start_month = DateValue(MonthBox.value & " " & YearEntry.value)
        'since report is for one month, just take desired month and add 1 to get first day where report does not apply
        
        If Month(start_month) = 12 Then
        end_month = DateValue(MonthName(1) & " " & YearEntry.value + 1)
        Else
        end_month = DateValue(Month(start_month) + 1 & " " & YearEntry.value)
        End If
        Debug.Print start_month
        Debug.Print end_month

        CreateCharts.CreateKPIMonth start_month, end_month, CategoryTextBox.value
        
        Hide
End Sub
Private Sub CancelButton_Click()
    Hide
    m_Cancelled = True
End Sub


