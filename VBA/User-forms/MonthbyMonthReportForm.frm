VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MonthbyMonthReportForm 
   Caption         =   "Date Range Data for MDI"
   ClientHeight    =   5910
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5850
   OleObjectBlob   =   "MonthbyMonthReportForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MonthbyMonthReportForm"
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
Dim filter_val As String
Dim item As Variant
Dim x As Variant
Dim Month_Name As String
Dim MonthArr(0 To 11) As Variant
Dim StartMonthNameExists As Variant
Dim EndMonthNameExists As Variant

    For x = 1 To 12
        MonthArr(x - 1) = MonthName(x)
    Next

    If MonthbyMonthReportForm.StartYearEntry.value = Empty Or _
            MonthbyMonthReportForm.EndYearEntry.value = Empty Or _
            MonthbyMonthReportForm.StartingMonthBox.value = Empty Or _
            MonthbyMonthReportForm.EndMonthBox.value = Empty Then
                MsgBox "Please include data for all fields."
                Exit Sub
    End If
        
    If _
        IssueRadio.value = False And _
        DateRadio.value = False And _
        IssueRadio.value = False And _
        IssueRadio.value = False Then

            MsgBox "Please select a filter method."
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
    
    start_month = DateValue(StartingMonthBox & "/1/" & StartYearEntry)
    
    end_month = DateValue(EndMonthBox & "/1/" & EndYearEntry)
    
    If StartYearEntry.value = EndYearEntry.value And _
       start_month > end_month Then
       MsgBox "If start year and end year are the same, put the earlier month in start month"
       Exit Sub
    End If
    
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

Basic_Reports.CreateBasicReport start_month, end_month, filter_val
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

        Yearly_Start = MonthbyMonthReportForm.StartYearEntry
        
End Property
Public Property Get Yearly_End() As String

        Yearly_End = MonthbyMonthReportForm.EndYearEntry
        
End Property
Public Property Get Monthly_start() As String

        Monthly_start = MonthbyMonthReportForm.StartingMonthBox
        
End Property
Public Property Get Monthly_end() As String

        Monthly_end = MonthbyMonthReportForm.EndMonthBox
        
End Property



Public Property Get IssueRadioMonthbyMonth() As String

        If MonthbyMonthReportForm.IssueRadio.value = True Then
        IssueRadioMonthbyMonth = 1
        Else
        IssueRadioMonthbyMonth = 0
        End If
End Property
Public Property Get DateRadioMonthbyMonth() As String

        If MonthbyMonthReportForm.DateRadio.value = True Then
        DateRadioMonthbyMonth = 1
        Else
        DateRadioMonthbyMonth = 0
        End If
End Property
Public Property Get CatRadioMonthbyMonth() As String

        If MonthbyMonthReportForm.CatRadio.value = True Then
        CatRadioMonthbyMonth = 1
        Else
        CatRadioMonthbyMonth = 0
        End If
End Property
Public Property Get KPIRadioMonthbyMonth() As String

        If MonthbyMonthReportForm.KPIRadio.value = True Then
        KPIRadioMonthbyMonth = 1
        Else
        KPIRadioMonthbyMonth = 0
        End If
End Property

