VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddNewEntry 
   Caption         =   "Add New Entry"
   ClientHeight    =   6490
   ClientLeft      =   70
   ClientTop       =   280
   ClientWidth     =   6040
   OleObjectBlob   =   "AddNewEntry.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "AddNewEntry"
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

Private Sub AddButton_Click()

 Dim counter_tbl As ListObject
 Dim IssueDateString As String
 Dim IssueDate As Date
 Dim DueDateString As String
 Dim DueDate As Date
 Dim Owner As Variant
 Dim row As Variant
 Dim Max As Double
 Dim num As Variant
 
 On Error GoTo ErrorHandler:
 
 
    Set counter_tbl = Worksheets("Countermeasures").ListObjects("Tbl_Counter")

If DateDayTextBox.value = Empty Or _
   DateMonthTextBox.value = Empty Or _
   DateYearTextBox.value = Empty Or _
   DueDayTextBox.value = Empty Or _
   DueMonthTextBox.value = Empty Or _
   DueYearTextBox.value = Empty Or _
   CategoryTextBox.value = Empty Or _
   KPITextBox.value = Empty Or _
   FirstNameTextBox.value = Empty Or _
   LastNameTextBox.value = Empty Then

    MsgBox ("Please ensure the issue date, category, kpi, due date, and owner information is entered. This is the minimum needed to create a new entry.")

    Exit Sub
End If

    IssueDateString = CStr(DateDayTextBox.value) & " " & DateMonthTextBox.value & " " & CStr(DateYearTextBox.value)
    IssueDate = DateValue(IssueDateString)
    DueDateString = CStr(DueDayTextBox.value) & " " & DueMonthTextBox.value & " " & CStr(DueYearTextBox.value)
    DueDate = DateValue(DueDateString)
    
    Owner = FirstNameTextBox.value & " " & LastNameTextBox.value
    
    Max = 1
    For Each num In counter_tbl.ListColumns("Issue ID").DataBodyRange
        If num > Max Then
            Max = num
        End If
    Next
  
    Dim row1 As Double
    Dim IssueIDNum As Variant
    
    row1 = 0
    
    For Each num In counter_tbl.ListColumns("Entry Identifier").DataBodyRange
        row1 = row1 + 1
        If num = AddNewEntry_Tags.EntryIdentifierBox _
                 And AddNewEntry_Tags.EntryIdentifierBox <> Empty _
                 And num <> "N/A" Then
            IssueIDNum = counter_tbl.ListColumns("Issue ID").DataBodyRange(row1, 1).value
            'row1 = counter_tbl.ListColumns("Issue ID").DataBodyRange(row1, 1).row
            
            AddNewEntry_Replace.Show vbModeless
            AddNewEntry_Replace.EntryIdentifierLabel.Caption = AddNewEntry_Tags.EntryIdentifierBox.value
            AddNewEntry_Replace.IssueIDLabel.Caption = IssueIDNum
            AddNewEntry_Replace.RowLabel.Caption = row1
            AddNewEntry_Replace.IssueLabel.Caption = AddNewEntry.IssueTextBox.value
            
            Exit Sub
        End If
    Next
    
    
    counter_tbl.ListRows.Add
        row = counter_tbl.DataBodyRange.Rows.Count
        counter_tbl.ListColumns("Category").DataBodyRange(row, 1).value = CategoryTextBox.value
        counter_tbl.ListColumns("KPI").DataBodyRange(row, 1).value = KPITextBox.value
        counter_tbl.ListColumns("Issue Date").DataBodyRange(row, 1).value = IssueDate
        counter_tbl.ListColumns("Issue Date").DataBodyRange(row, 1).NumberFormat = "d-mmm-yy"
        counter_tbl.ListColumns("Issue").DataBodyRange(row, 1).value = IssueTextBox.value
        counter_tbl.ListColumns("Cause").DataBodyRange(row, 1).value = CauseTextBox.value
        counter_tbl.ListColumns("Countermeasure").DataBodyRange(row, 1).value = CountermeasureTextBox.value
        counter_tbl.ListColumns("Owner").DataBodyRange(row, 1).value = Owner
        counter_tbl.ListColumns("Date Due").DataBodyRange(row, 1).value = DueDate
            counter_tbl.ListColumns("Date Due").DataBodyRange(row, 1).NumberFormat = "d-mmm-yy"
        counter_tbl.ListColumns("Issue ID").DataBodyRange(row, 1).value = Max
        
        If IsLoaded("AddNewEntry_Tags") = True Then
            counter_tbl.ListColumns("Issue Tier 1 Tag").DataBodyRange(row, 1).value = AddNewEntry_Tags.IssueTier1Box
            counter_tbl.ListColumns("Issue Tier 2 Tag").DataBodyRange(row, 1).value = AddNewEntry_Tags.IssueTier2Box
            counter_tbl.ListColumns("Cause Category").DataBodyRange(row, 1).value = AddNewEntry_Tags.CauseCatBox
            counter_tbl.ListColumns("Cause Detail").DataBodyRange(row, 1).value = AddNewEntry_Tags.CauseDetBox
            counter_tbl.ListColumns("Entry Identifier").DataBodyRange(row, 1).value = AddNewEntry_Tags.EntryIdentifierBox
            counter_tbl.ListColumns("Primary Equipment").DataBodyRange(row, 1).value = AddNewEntry_Tags.PrimaryEquiptmentBox
            counter_tbl.ListColumns("Manufacturing Stage").DataBodyRange(row, 1).value = AddNewEntry_Tags.MfgStageBox
            counter_tbl.ListColumns("Batch").DataBodyRange(row, 1).value = AddNewEntry_Tags.BatchBox
            counter_tbl.ListColumns("Quality Classification").DataBodyRange(row, 1).value = AddNewEntry_Tags.QualityClassificationBox
            counter_tbl.ListColumns("Safety Tier").DataBodyRange(row, 1).value = AddNewEntry_Tags.SafetyTierBox
        End If
        
ErrorHandler:
    MsgBox "if you are seeing this error, it is likely because you have deleted/changed the columns with headers: chr(34)Issue Tier 1 Tag, Issue Tier 2 Tag, Cause Category, Cause Detail," _
         & " Entry Identifier, Primary Equipment, Manufacturing Stage, Batch, Quality Classification, or Safety Tier." _
         & " This simply means whatever tags you updated for this entry need to be updated manually within the modified table you created, the ''Tag'' userform will only" _
         & " populate columns that match the tag/descriptor labels on the user-form.chr(34)"
        
   
End Sub


Private Sub CancelButton_Click()
    Hide
    Unload Me

Unload AddNewEntry_Tags
AddNewEntry_Tags.Hide
    m_Cancelled = True
    
    
End Sub


Private Sub ClearButton_Click()
Unload AddNewEntry_Tags
Unload AddNewEntry

AddNewEntry_Tags.Hide
AddNewEntry.Hide

Call AddNewEntry_Tags_M.ShowEntriesForm

End Sub


Private Sub Userform_Initialize()

Me.Top = Application.Top + (Me.Height / 8)
Me.left = Application.left + (Application.UsableWidth / 3) - (Me.Width / 2)

AddNewEntry_Tags.Show vbModeless

AddNewEntry_Tags.Top = Me.Top + (AddNewEntry_Tags.Height / 8)
AddNewEntry_Tags.left = Me.left + (Application.UsableWidth / 3)
    DateYearTextBox.MaxLength = 4
    DateDayTextBox.MaxLength = 2
    
    DueYearTextBox.MaxLength = 4
    DueDayTextBox.MaxLength = 2


    DateMonthTextBox.AddItem "January"
    DateMonthTextBox.AddItem "February"
    DateMonthTextBox.AddItem "March"
    DateMonthTextBox.AddItem "April"
    DateMonthTextBox.AddItem "May"
    DateMonthTextBox.AddItem "June"
    DateMonthTextBox.AddItem "July"
    DateMonthTextBox.AddItem "August"
    DateMonthTextBox.AddItem "September"
    DateMonthTextBox.AddItem "October"
    DateMonthTextBox.AddItem "November"
    DateMonthTextBox.AddItem "December"
    
    DueMonthTextBox.AddItem "January"
    DueMonthTextBox.AddItem "February"
    DueMonthTextBox.AddItem "March"
    DueMonthTextBox.AddItem "April"
    DueMonthTextBox.AddItem "May"
    DueMonthTextBox.AddItem "June"
    DueMonthTextBox.AddItem "July"
    DueMonthTextBox.AddItem "August"
    DueMonthTextBox.AddItem "September"
    DueMonthTextBox.AddItem "October"
    DueMonthTextBox.AddItem "November"
    DueMonthTextBox.AddItem "December"
    
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
    Unload Me
    Unload AddNewEntry_Tags
    AddNewEntry_Tags.Hide
    AddNewEntry_Tags.Hide
    m_Cancelled = True
End Sub
Private Sub DateYearTextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
         
    Select Case KeyAscii
        Case Is < vbKey0, Is > vbKey9
            KeyAscii = 0
            Beep
    End Select
End Sub
Private Sub DateDayTextBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
         
    Select Case KeyAscii
        Case Is < vbKey0, Is > vbKey9
            KeyAscii = 0
            Beep
    End Select
End Sub

Private Sub CategoryTextBox_AfterUpdate()

 Dim counter_tbl As ListObject
    Set counter_tbl = Worksheets("Countermeasures").ListObjects("Tbl_Counter")
    Dim ArrBase() As Variant
    Dim Cell3 As Variant
    Dim counter_row_count As Variant
    Dim entry_count As Variant
    Dim item As Variant
    
    
    KPITextBox.Clear

counter_row_count = 0
For Each Cell3 In counter_tbl.ListColumns("KPI").DataBodyRange
                    counter_row_count = counter_row_count + 1
                    
            If CategoryTextBox.value = counter_tbl.ListColumns("Category").DataBodyRange(counter_row_count, 1).value Then
                    
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
            End If
Next Cell3

'MsgBox Join(arrbase, vbCrLf)

    If (Not Not ArrBase) = 0 Then
        KPITextBox.AddItem ""
        Exit Sub
    End If
                
    ArrBase = BlankRemover(ArrBase)
    ArrBase = ArrayRemoveDups(ArrBase)
    
    
    For Each item In ArrBase
        KPITextBox.AddItem item
    Next
    
Call AddNewEntry_Tags_M.IssueTier1Box
Call AddNewEntry_Tags_M.IssueTier2Box
Call AddNewEntry_Tags_M.CauseCatBox
Call AddNewEntry_Tags_M.CauseDetBox
Call AddNewEntry_Tags_M.PrimaryEquiptmentBox
Call AddNewEntry_Tags_M.MfgStageBox
Call AddNewEntry_Tags_M.BatchBox
Call AddNewEntry_Tags_M.QualityClassificationBox
Call AddNewEntry_Tags_M.SafetyTierBox
Call AddNewEntry_Tags_M.EntryIdentifierBox

    
End Sub



