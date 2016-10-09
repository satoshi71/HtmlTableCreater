VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Check Form"
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10230
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
   TextBox1.SelStart = 0
   TextBox1.SelLength = Len(TextBox1.Text)
   TextBox1.Copy

End Sub

Private Sub CommandButton2_Click()
   Unload UserForm2
End Sub


Private Sub UserForm_Initialize()
   path_ = ActiveWorkbook.Path
   path_ = "file:///" & Replace(path_, "\", "/") & "/index.html"

   WebBrowser1.Navigate path_
End Sub
