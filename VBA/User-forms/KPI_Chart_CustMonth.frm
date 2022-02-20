VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} KPI_Chart_CustMonth 
   Caption         =   "KPI Summary for a Custom Monthly Interval"
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   6650
   OleObjectBlob   =   "KPI_Chart_CustMonth.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "KPI_Chart_CustMonth"
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

Private Sub EndYearEntry_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
         
    Select Case KeyAscii
        Case Is < vbKey0, Is > vbKey9
            KeyAscii = 0
            Beep
    End Select
End Sub

Sub Userform_Initialize()
    Dim x As Variant
    Dim Month_Name As String
    Dim MonthArr(0 To 11) As Variant
    For x = 1 To 12
        MonthArr(x - 1) = MonthName(x)
        StartingMonthBox.AddItem MonthArr(x - 1)
        EndMonthBox.AddItem MonthArr(x - 1)
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

Private Sub OKButton_Click()

Dim start_month As Date
Dim end_month As Date
Dim item As Variant
Dim x As Variant
Dim Month_Name As String
Dim MonthArr(0 To 11) As Variant
Dim StartMonthNameExists As Variant
Dim EndMonthNameExists As Variant

    For x = 1 To 12
        MonthArr(x - 1) = MonthName(x)
    Next

    If StartYearEntry.value = Empty Or _
    EndYearEntry.value = Empty Or _
    StartingMonthBox.value = Empty Or _
    EndMonthBox.value = Empty Then
        MsgBox "Please include data for all fields."
        Exit Sub
    End If
    
    StartMonthNameExists = False
    EndMonthNameExists = False
    For Each item In MonthArr
        If StartingMonthBox.value = item Then
            StartMonthNameExists = True
        End If
        If EndMonthBox.value = item Then
            EndMonthNameExists = True
        End If
    Next
    
    If StartMonthNameExists = False Then
        MsgBox "Please write the full name of the month in the start month box"
            Exit Sub
    End If
    If EndMonthNameExists = False Then
        MsgBox "Please write the full name of the month in the end month box"
            Exit Sub
    End If
    
    If StartYearEntry.value > EndYearEntry.value Then
        MsgBox "Please put earlier year in start entry box"
        Exit Sub
    End If
    
     If CategoryTextBox.value = Empty Or CategoryTextBox.value = "" Then
        MsgBox "Please enter a category"
        Exit Sub
    End If
    
    start_month = DateValue(StartingMonthBox & "/1/" & StartYearEntry)
    
    end_month = DateValue(EndMonthBox & "/1/" & EndYearEntry)
    
    If StartYearEntry.value = EndYearEntry.value And _
       start_month > end_month Then
       MsgBox "If start year and end year are the same, put the earlier month in start month"
       Exit Sub
    End If
    
    
    CreateCharts.CreateKPIChartCustMonth start_month, end_month, CategoryTextBox.value
        Hide
End Sub

Private Sub StartYearEntry_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
         
    Select Case KeyAscii
        Case Is < vbKey0, Is > vbKey9
            KeyAscii = 0
            Beep
    End Select
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer _
                                       , CloseMode As Integer)
    
    ' Prevent the form being unloaded
    If CloseMode = vbFormControlMenu Then Cancel = True
    
    ' Hide the Userform and set cancelled to true
    Hide
    m_Cancelled = True
    
End Sub







