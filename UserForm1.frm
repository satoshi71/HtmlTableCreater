VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Setting"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4380
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CommandButton1_Click()
   If Label1.Caption = "X" Then
      MsgBox ("Border Color is not decided.")
      Exit Sub
   End If
   
   padding = TextBox2.Text
   bcolor = "#" & UCase(TextBox1.Text)
   
   If IsNumeric(padding) = False Then
      MsgBox ("Padding value is not the number.")
      Exit Sub
   End If
   
   
   Call createTable(padding, bcolor)
   Unload UserForm1

End Sub

Private Sub TextBox1_Change()
   Label1.BackColor = RGB(255, 255, 255)
   Label1.Caption = "X"
   
   If Len(TextBox1.Text) < 6 Then Exit Sub

   On Error GoTo ErrHandler
   
   Label1.Caption = ""
   HexCode = TextBox1.Text

   r = CLng("&H" & Mid(HexCode, 1, 2))
   G = CLng("&H" & Mid(HexCode, 3, 2)) * 256
   b = CLng("&H" & Mid(HexCode, 5, 2)) * 256 * 256

   ColorCode = r + G + b

   Label1.BackColor = ColorCode
   Exit Sub

ErrHandler:
   Label1.BackColor = RGB(255, 255, 255)
   Label1.Caption = "X"

End Sub


Private Sub UserForm_Click()
   Label1.BackColor = 6783879
End Sub
