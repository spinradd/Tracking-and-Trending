VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReportChoiceForm 
   Caption         =   "Year for MDI Data"
   ClientHeight    =   4470
   ClientLeft      =   170
   ClientTop       =   690
   ClientWidth     =   4450
   OleObjectBlob   =   "ReportChoiceForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReportChoiceForm"
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
    If ReportChoiceForm.QuarterRadio.value = False And _
    ReportChoiceForm.YearRadio.value = False And _
    ReportChoiceForm.MonthRadio.value = False And _
    ReportChoiceForm.CustomMonthRadio.value = False Then
        MsgBox "Please select one option."
        Exit Sub
    ElseIf ReportChoiceForm.QuarterRadio.value = True Then
        Hide
        QuarterlyReportForm.Show
    ElseIf ReportChoiceForm.YearRadio.value = True Then
        Hide
         YearReportForm.Show
    ElseIf ReportChoiceForm.MonthRadio.value = True Then
        Hide
         MonthReportForm.Show
    ElseIf ReportChoiceForm.CustomMonthRadio.value = True Then
        Hide
         MonthbyMonthReportForm.Show
    Else
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


