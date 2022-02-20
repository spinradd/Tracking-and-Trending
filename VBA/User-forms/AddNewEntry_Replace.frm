VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddNewEntry_Replace 
   Caption         =   "UserForm1"
   ClientHeight    =   3570
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "AddNewEntry_Replace.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddNewEntry_Replace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ReplaceButton_Click()

Dim counter_tbl As ListObject
Set counter_tbl = Worksheets("Countermeasures").ListObjects("Tbl_Counter")

 Dim IssueDateString As String
 Dim IssueDate As Date
 Dim DueDateString As String
 Dim DueDate As Date
 Dim Owner As Variant
 Dim row As Variant
 Dim Max As Double
 Dim num As Variant


IssueDateString = CStr(AddNewEntry.DateDayTextBox.value) & " " & AddNewEntry.DateMonthTextBox.value & " " & CStr(AddNewEntry.DateYearTextBox.value)
    IssueDate = DateValue(IssueDateString)
    DueDateString = CStr(AddNewEntry.DueDayTextBox.value) & " " & AddNewEntry.DueMonthTextBox.value & " " & CStr(AddNewEntry.DueYearTextBox.value)
    DueDate = DateValue(DueDateString)
    
    Owner = AddNewEntry.FirstNameTextBox.value & " " & AddNewEntry.LastNameTextBox.value
    
    row = CLng(AddNewEntry_Replace.RowLabel.Caption)

 
    For Each cell In counter_tbl.HeaderRowRange
        counter_tbl.ListColumns(cell.value).DataBodyRange(row, 1).ClearContents
    Next
   
        counter_tbl.ListColumns("Category").DataBodyRange(row, 1).value = AddNewEntry.CategoryTextBox.value
        counter_tbl.ListColumns("KPI").DataBodyRange(row, 1).value = AddNewEntry.KPITextBox.value
        counter_tbl.ListColumns("Issue Date").DataBodyRange(row, 1).value = IssueDate
        counter_tbl.ListColumns("Issue Date").DataBodyRange(row, 1).NumberFormat = "d-mmm-yy"
        counter_tbl.ListColumns("Issue").DataBodyRange(row, 1).value = AddNewEntry.IssueTextBox.value
        counter_tbl.ListColumns("Cause").DataBodyRange(row, 1).value = AddNewEntry.CauseTextBox.value
        counter_tbl.ListColumns("Countermeasure").DataBodyRange(row, 1).value = AddNewEntry.CountermeasureTextBox.value
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

Unload Me
AddNewEntry_Replace.Hide


End Sub


Private Sub CancelButton_Click()
    Hide
    Unload Me
    m_Cancelled = True
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer _
                                       , CloseMode As Integer)
    
    ' Prevent the form being unloaded
    If CloseMode = vbFormControlMenu Then Cancel = True
    
    ' Hide the Userform and set cancelled to true
    Hide
    Unload Me
    m_Cancelled = True
End Sub

