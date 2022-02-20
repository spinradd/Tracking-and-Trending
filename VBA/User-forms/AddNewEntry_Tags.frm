VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddNewEntry_Tags 
   Caption         =   "Add Tags"
   ClientHeight    =   5040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "AddNewEntry_Tags.frx":0000
End
Attribute VB_Name = "AddNewEntry_Tags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Unload Me
        Hide
    End If
End Sub
