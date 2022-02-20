VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TandDYear 
   Caption         =   "Tag and Descriptor Yearly"
   ClientHeight    =   2360
   ClientLeft      =   130
   ClientTop       =   520
   ClientWidth     =   4130
   OleObjectBlob   =   "TandDYear.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TandDYear"
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
Dim start_month As Date
Dim end_month As Date
Dim filter_val As String
  If TandDYear.YearEntry = Empty Then
        MsgBox "Please select a year."
        Exit Sub
    End If
    
    
        start_month = DateValue("January " & Year_entry)
        end_month = DateValue("December " & Year_entry)
        
        Identifier_Updater_M.CreateTagDescriptorSheet start_month, end_month
        
        Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer _
                                       , CloseMode As Integer)
    
    ' Prevent the form being unloaded
    If CloseMode = vbFormControlMenu Then Cancel = True
    
    ' Hide the Userform and set cancelled to true
    Hide
    m_Cancelled = True
    
End Sub
Public Property Get Year_entry() As String

        Year_entry = TandDYear.YearEntry
        
End Property
Private Sub YearEntry_Change_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
         
    Select Case KeyAscii
        Case Is < vbKey0, Is > vbKey9
            KeyAscii = 0
            Beep
    End Select
End Sub

