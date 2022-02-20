VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} QuarterlyReportForm 
   Caption         =   "Quarterly Data"
   ClientHeight    =   4755
   ClientLeft      =   170
   ClientTop       =   690
   ClientWidth     =   5970
   OleObjectBlob   =   "QuarterlyReportForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "QuarterlyReportForm"
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

Private Sub Userform_Initialize()

YearEntry.MaxLength = 4
FilterTextBox.AddItem "Issue ID"
FilterTextBox.AddItem "Issue Date"
FilterTextBox.AddItem "Category"
FilterTextBox.AddItem "KPI"


End Sub

Private Sub OKButton_Click()
Dim start_month As Date
Dim end_month As Date
Dim filter_val As String
    
    
    If QuarterlyReportForm.Radio1.value = False And _
    QuarterlyReportForm.Radio2.value = False And _
    QuarterlyReportForm.Radio3.value = False And _
    QuarterlyReportForm.Radio1.value = False And _
    QuarterlyReportForm.YearEntry = Empty Then
        MsgBox "Please make sure both a year and a quarter are selected"
        
    ElseIf _
    FilterTextBox.value = Empty Then
    
    MsgBox "Please select a filter method."
    
    Exit Sub
    End If
    
    filter_val = FilterTextBox.value
     
    If QuarterlyReportForm.Radio1.value = True Then
        start_month = DateValue("January 1 " & YearEntry.value)
        end_month = DateValue("March 1 " & YearEntry.value)
    ElseIf QuarterlyReportForm.Radio2.value = True Then
        start_month = DateValue("April 1 " & YearEntry.value)
        end_month = DateValue("June 1 " & YearEntry.value)
    ElseIf QuarterlyReportForm.Radio3.value = True Then
        start_month = DateValue("July 1 " & YearEntry.value)
        end_month = DateValue("September 1 " & YearEntry.value)
    ElseIf QuarterlyReportForm.Radio4.value = True Then
        start_month = DateValue("October 1 " & YearEntry.value)
        end_month = DateValue("January 1 " & CStr(Int(YearEntry.value) + 1))
     
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


Private Sub YearEntry_Change_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
         
    Select Case KeyAscii
        Case Is < vbKey0, Is > vbKey9
            KeyAscii = 0
            Beep
    End Select
End Sub
