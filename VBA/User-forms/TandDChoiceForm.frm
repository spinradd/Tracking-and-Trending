VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TandDChoiceForm 
   Caption         =   "Tag and Descriptor Choice Form"
   ClientHeight    =   4455
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4590
   OleObjectBlob   =   "TandDChoiceForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TandDChoiceForm"
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
    If TandDChoiceForm.QuarterRadio.value = False And _
    TandDChoiceForm.YearRadio.value = False And _
    TandDChoiceForm.MonthRadio.value = False And _
    TandDChoiceForm.AllDataRadio.value = False And _
    TandDChoiceForm.CustomMonthRadio.value = False Then
        MsgBox "Please select one option."
        Exit Sub
    ElseIf TandDChoiceForm.QuarterRadio.value = True Then
        Hide
        TandDQuarterly.Show vbModeless
    ElseIf TandDChoiceForm.YearRadio.value = True Then
        Hide
         TandDYear.Show vbModeless
    ElseIf TandDChoiceForm.MonthRadio.value = True Then
        Hide
         TandDMonthly.Show vbModeless
    ElseIf TandDChoiceForm.CustomMonthRadio.value = True Then
        Hide
         TandDCustomMonth.Show vbModeless
    ElseIf TandDChoiceForm.AllDataRadio.value = True Then
        Hide
         Identifier_Updater_M.CreateTagDescriptorSheet DateValue(1 / 1 / 200), Date
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

