VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TandDMonthly 
   Caption         =   "Tag and Descriptor Month"
   ClientHeight    =   2300
   ClientLeft      =   130
   ClientTop       =   520
   ClientWidth     =   4130
   OleObjectBlob   =   "TandDMonthly.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TandDMonthly"
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
    
    If YearEntry = Empty Or MonthBox.value = Empty Then
        MsgBox "Please select a year."
        Exit Sub
    End If
    
    If StartMonthNameExists = False Then
        MsgBox "Please write the full name of the month in the start month box"
            Exit Sub
    End If
    
        start_month = DateValue(MonthBox.value & " " & start_year)
        'since report is for one month, just take desired month and add 1 to get first day where report does not apply
        
        If Month(start_month) = 12 Then
        end_month = DateValue(MonthName(1) & " " & start_year + 1)
        Else
        end_month = DateValue(Month(start_month) + 1 & " " & start_year)
        End If

        Identifier_Updater_M.CreateTagDescriptorSheet start_month, end_month
        
        Hide
End Sub
Private Sub CancelButton_Click()
    Hide
    m_Cancelled = True
End Sub
