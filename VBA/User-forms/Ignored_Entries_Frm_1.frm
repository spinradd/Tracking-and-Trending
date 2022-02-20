VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Ignored_Entries_Frm_1 
   Caption         =   "Ignore Entries Form 1"
   ClientHeight    =   5300
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4800
   OleObjectBlob   =   "Ignored_Entries_Frm_1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Ignored_Entries_Frm_1"
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
    Unload Me
    Call VeevaMacros.AddNewEntries
End Sub
Private Sub CancelMacro_Click()
    Hide
    m_Cancelled = True
End Sub


Private Sub OKButton_Click()
        Hide
        Ignored_Entries_Frm_2.Show vbModeless
End Sub

Private Sub Userform_Initialize()

Dim IgnoredTable As ListObject
Dim IgnoredSheet As Worksheet
Dim EntryArray() As Variant
Dim Entry As Variant
    Set IgnoredSheet = ThisWorkbook.Worksheets("Entries to Ignore via Import")
    Set IgnoredTable = IgnoredSheet.ListObjects("Tbl_Ignored_Entries")
    
    EntryArray = IgnoredTable.ListColumns("Entry Identifier").DataBodyRange
    
    For Each Entry In EntryArray
        Ignored_Entries_Frm_1.EntriesListBox.AddItem (Entry)
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

