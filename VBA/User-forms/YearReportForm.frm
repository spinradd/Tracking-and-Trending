VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} YearReportForm 
   Caption         =   "Yearly Report"
   ClientHeight    =   5810
   ClientLeft      =   170
   ClientTop       =   690
   ClientWidth     =   6880
   OleObjectBlob   =   "YearReportForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "YearReportForm"
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
Dim start_month As Date
Dim end_month As Date
Dim filter_val As String
  If YearReportForm.YearEntry = Empty Then
        MsgBox "Please select a year."
    ElseIf _
    IssueRadio.value = False And _
    DateRadio.value = False And _
    IssueRadio.value = False And _
    IssueRadio.value = False Then
    
    MsgBox "Please select a filter method."
    
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
    
        start_month = DateValue("January " & Year_entry)
        end_month = DateValue("December " & Year_entry)
        
        Basic_Reports.CreateBasicReport start_month, end_month, filter_val
        
        Hide
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer _
                                       , CloseMode As Integer)
    
    ' Prevent the form being unloaded
    If CloseMode = vbFormControlMenu Then Cancel = True
    
    ' Hide the Userform and set cancelled to true
    Hide
    m_Cancelled = True
    
End Sub
Public Property Get Year_entry() As String

        Year_entry = YearReportForm.YearEntry
        
End Property
Private Sub YearEntry_Change_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
         
    Select Case KeyAscii
        Case Is < vbKey0, Is > vbKey9
            KeyAscii = 0
            Beep
    End Select
End Sub

Public Property Get IssueRadioYear() As String

        If YearReportForm.IssueRadio.value = True Then
        IssueRadioYear = 1
        Else
        IssueRadioYear = 0
        End If
End Property
Public Property Get DateRadioYear() As String

        If YearReportForm.DateRadio.value = True Then
        DateRadioYear = 1
        Else
        DateRadioYear = 0
        End If
End Property
Public Property Get CatRadioYear() As String

        If YearReportForm.CatRadio.value = True Then
        CatRadioYear = 1
        Else
        CatRadioYear = 0
        End If
End Property
Public Property Get KPIRadioYear() As String

        If YearReportForm.KPIRadio.value = True Then
        KPIRadioYear = 1
        Else
        KPIRadioYear = 0
        End If
End Property
