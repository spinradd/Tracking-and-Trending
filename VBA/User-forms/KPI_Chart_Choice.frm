VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} KPI_Chart_Choice 
   Caption         =   "Pick KPI Report Choice"
   ClientHeight    =   4620
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4560
   OleObjectBlob   =   "KPI_Chart_Choice.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "KPI_Chart_Choice"
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
    If QuarterRadio.value = False And _
    YearRadio.value = False And _
    MonthRadio.value = False And _
    AllDataRadio.value = False And _
    CustomMonthRadio.value = False Then
        MsgBox "Please select one option."
        Exit Sub
    ElseIf QuarterRadio.value = True Then
        Hide
        KPI_Chart_Quarterly.Show vbModeless
    ElseIf YearRadio.value = True Then
        Hide
         KPI_Chart_Yearly.Show vbModeless
    ElseIf MonthRadio.value = True Then
        Hide
         KPI_Chart_Monthly.Show vbModeless
    ElseIf CustomMonthRadio.value = True Then
        Hide
         KPI_Chart_CustMonth.Show vbModeless
    ElseIf AllDataRadio.value = True Then
        Hide
         'Identifier_Updater_M.CreateTagDescriptorSheet DateValue(1 / 1 / 200), Date
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

