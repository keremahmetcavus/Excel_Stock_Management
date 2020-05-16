VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Stok Takip"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4080
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Unload Me
Stokgiris.Show
End Sub

Private Sub CommandButton4_Click()
Unload Me
''Application.Visible = True

End Sub

Private Sub CommandButton7_Click()
Unload Me
''Application.Visible = True
ActiveWorkbook.Save
ActiveWorkbook.Close
End Sub
