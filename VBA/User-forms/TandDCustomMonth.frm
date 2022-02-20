VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TandDCustomMonth 
   Caption         =   "Tag and Descriptor Month by Month"
   ClientHeight    =   3855
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6330
   OleObjectBlob   =   "TandDCustomMonth.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TandDCustomMonth"
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

    If TandDCustomMonth.StartYearEntry.value = Empty Or _
    TandDCustomMonth.EndYearEntry.value = Empty Or _
    TandDCustomMonth.StartingMonthBox.value = Empty Or _
    TandDCustomMonth.EndMonthBox.value = Empty Then
        MsgBox "Please include data for all fields."
        Exit Sub
    End If
    
    StartMonthNameExists = False
    EndMonthNameExists = False
    For Each item In MonthArr
        If TandDCustomMonth.StartingMonthBox.value = item Then
            StartMonthNameExists = True
        End If
        If TandDCustomMonth.EndMonthBox.value = item Then
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
    
    If TandDCustomMonth.StartYearEntry.value > TandDCustomMonth.EndYearEntry.value Then
        MsgBox "Please put earlier year in start entry box"
        Exit Sub
    End If
    
    start_month = DateValue(StartingMonthBox & "/1/" & StartYearEntry)
    
    end_month = DateValue(EndMonthBox & "/1/" & EndYearEntry)
    
    If TandDCustomMonth.StartYearEntry.value = TandDCustomMonth.EndYearEntry.value And _
       start_month > end_month Then
       MsgBox "If start year and end year are the same, put the earlier month in start month"
       Exit Sub
    End If
    
    Identifier_Updater_M.CreateTagDescriptorSheet start_month, end_month
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
Public Property Get Yearly_Start() As String

        Yearly_Start = TandDCustomMonth.StartYearEntry
        
End Property
Public Property Get Yearly_End() As String

        Yearly_End = TandDCustomMonth.EndYearEntry
        
End Property
Public Property Get Monthly_start() As String

        Monthly_start = TandDCustomMonth.StartingMonthBox
        
End Property
Public Property Get Monthly_end() As String

        Monthly_end = TandDCustomMonth.EndMonthBox
        
End Property



