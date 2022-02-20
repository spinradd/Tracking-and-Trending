VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} KPI_Chart_Quarterly 
   Caption         =   "Create KPI summary for a quarter"
   ClientHeight    =   3870
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   6320
   OleObjectBlob   =   "KPI_Chart_Quarterly.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "KPI_Chart_Quarterly"
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

Private Sub Label1_Click()

End Sub

Private Sub FinalMonthTxt_Click()

End Sub

Private Sub OKButton_Click()
Dim start_month As Date
Dim end_month As Date
Dim filter_val As String
    
    
    If KPI_Chart_Quarterly.Radio1.value = False And _
    KPI_Chart_Quarterly.Radio2.value = False And _
    KPI_Chart_Quarterly.Radio3.value = False And _
    KPI_Chart_Quarterly.Radio1.value = False And _
    KPI_Chart_Quarterly.YearEntry = Empty Then
        MsgBox "Please make sure both a year and a quarter are selected"
    
    End If
     
    If KPI_Chart_Quarterly.Radio1.value = True Then
        start_month = DateValue("January 1" & CStr(QuarterlyYear))
        end_month = DateValue("April 1 " & CStr(QuarterlyYear))
    ElseIf KPI_Chart_Quarterly.Radio2.value = True Then
        start_month = DateValue("April 1 " & CStr(QuarterlyYear))
        end_month = DateValue("July 1 " & CStr(QuarterlyYear))
    ElseIf KPI_Chart_Quarterly.Radio3.value = True Then
        start_month = DateValue("July 1 " & CStr(QuarterlyYear))
        end_month = DateValue("October 1 " & CStr(QuarterlyYear))
    ElseIf KPI_Chart_Quarterly.Radio4.value = True Then
        start_month = DateValue("October 1 " & CStr(QuarterlyYear))
        end_month = DateValue("January 1 " & CStr(QuarterlyYear + 1))
    End If
        
    If CategoryTextBox.value = Empty Or CategoryTextBox.value = "" Then
        MsgBox "Please enter a category"
        Exit Sub
    End If
     
     CreateCharts.CreateKPIChartCustMonth start_month, end_month, CategoryTextBox.value
     
        Hide
End Sub

Private Sub Userform_Initialize()
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


Public Property Get QuarterlyYear() As String

        QuarterlyYear = KPI_Chart_Quarterly.YearEntry
        
End Property


Private Sub YearEntry_Change_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
         
    Select Case KeyAscii
        Case Is < vbKey0, Is > vbKey9
            KeyAscii = 0
            Beep
    End Select
End Sub

