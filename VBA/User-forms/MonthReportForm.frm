VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MonthReportForm 
   Caption         =   "MDI Monthly Report"
   ClientHeight    =   3980
   ClientLeft      =   130
   ClientTop       =   500
   ClientWidth     =   2910
   OleObjectBlob   =   "MonthReportForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MonthReportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub IssueRadio_Click()

End Sub

Private Sub Userform_Initialize()

    MonthBox.AddItem "January"
    MonthBox.AddItem "February"
    MonthBox.AddItem "March"
    MonthBox.AddItem "April"
    MonthBox.AddItem "May"
    MonthBox.AddItem "June"
    MonthBox.AddItem "July"
    MonthBox.AddItem "August"
    MonthBox.AddItem "September"
    MonthBox.AddItem "October"
    MonthBox.AddItem "November"
    MonthBox.AddItem "December"

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

        start_year = MonthReportForm.YearEntry
        
End Property
Public Property Get start_month() As String

        start_month = MonthReportForm.MonthBox.value
        
End Property

Public Property Get IssueRadioMonth() As String

        If MonthReportForm.IssueRadio.value = True Then
        IssueRadioMonth = 1
        Else
        IssueRadioMonth = 0
        End If
End Property
Public Property Get DateRadioMonth() As String

        If MonthReportForm.DateRadio.value = True Then
        DateRadioMonth = 1
        Else
        DateRadioMonth = 0
        End If
End Property
Public Property Get CatRadioMonth() As String

        If MonthReportForm.CatRadio.value = True Then
        CatRadioMonth = 1
        Else
        CatRadioMonth = 0
        End If
End Property
Public Property Get KPIRadioMonth() As String

        If MonthReportForm.KPIRadio.value = True Then
        KPIRadioMonth = 1
        Else
        KPIRadioMonth = 0
        End If
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
Dim filter_val As String
Dim entries_per_slide As Integer
Dim FontSizeNum As Double
  If MonthReportForm.YearEntry = Empty Or MonthBox.value = Empty Then
        MsgBox "Please select a year."
ElseIf _
    IssueRadio.value = False And _
    DateRadio.value = False And _
    CatRadio.value = False And _
    KPIRadio.value = False Then
    
    MsgBox "Please select a filter method."

ElseIf _
    MonthReportForm.EntriesPerSlide = Empty Or MonthReportForm.FontSize = Empty Then
        MsgBox "Please select a powerpoint format."
    
    Else
    
    Select Case filter_val
    Case IssueRadio.value = True
        filter_val = "Issue ID"
    Case DateRadio.value = True
        filter_val = "Issue Date"
    Case CatRadio.value = True
        filter_val = "Category"
    Case KPIRadio.value = True
        filter_val = "KPI"
    End Select
    
        start_month = DateValue(MonthBox.value & " " & start_year)
        'since report is for one month, just take desired month and add 1 to get first day where report does not apply
        If Month(start_month) = 12 Then
        end_month = DateValue(MonthName(1) & " " & start_year + 1)
        Else
        end_month = DateValue(Month(start_month) + 1 & " " & start_year)
        End If
        
        entries_per_slide = EntriesPerSlide.value
        FontSizeNum = FontSize.value
        
        Hide
        'CreateReports.Create_Pivot_Trend_for_ppt start_month
        MonthlyReports.CreateBasicReport start_month, end_month, filter_val
        Basic_Reports.CreatePPTReport entries_per_slide, FontSizeNum, start_month
        
        Hide
    End If
End Sub
Private Sub CancelButton_Click()
    Hide
    m_Cancelled = True
End Sub


