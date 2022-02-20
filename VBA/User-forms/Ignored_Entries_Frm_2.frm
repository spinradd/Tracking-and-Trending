VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Ignored_Entries_Frm_2 
   Caption         =   "Ignore Entries Form 2"
   ClientHeight    =   3960
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "Ignored_Entries_Frm_2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Ignored_Entries_Frm_2"
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
    Call VeevaMacros.AddNewEntries
    m_Cancelled = True
End Sub

Private Sub CancelMacro_Click()
    Hide
    m_Cancelled = True
End Sub

Private Sub OKButton_Click()
Dim IgnoredTable As ListObject
Dim IgnoredSheet As Worksheet
Dim row As Double
Dim item As Variant

Set IgnoredSheet = ThisWorkbook.Worksheets("Entries to Ignore via Import")
    Set IgnoredTable = IgnoredSheet.ListObjects("Tbl_Ignored_Entries")

    If Ignored_Entries_Frm_2.EntryIdentifier = Empty Or Ignored_Entries_Frm_2.ReasoningBox = Empty Then
        MsgBox "Please input values for both text boxes"
        Exit Sub
    End If
    
    For Each item In IgnoredTable.ListColumns("Entry Identifier").DataBodyRange
        If item = ID_entry Then
            MsgBox "This entry is already present, please add a different one or cancel."
            Exit Sub
        End If
    Next
    
    IgnoredTable.ListRows.Add
        row = IgnoredTable.DataBodyRange.Rows.Count
        IgnoredTable.ListColumns("Entry Identifier").DataBodyRange(row, 1).value = ID_entry
        IgnoredTable.ListColumns("Reason to Ignore?").DataBodyRange(row, 1).value = Reasoning
        
    Unload Me
    Ignored_Entries_Frm_2.Show vbModeless
        
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer _
                                       , CloseMode As Integer)
    
    ' Prevent the form being unloaded
    If CloseMode = vbFormControlMenu Then Cancel = True
    
    ' Hide the Userform and set cancelled to true
    Hide
    m_Cancelled = True
    
End Sub
Public Property Get ID_entry() As String
        ID_entry = Ignored_Entries_Frm_2.EntryIdentifier
End Property
Public Property Get Reasoning() As String
        Reasoning = Ignored_Entries_Frm_2.ReasoningBox
End Property



