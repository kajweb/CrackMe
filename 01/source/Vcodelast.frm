VERSION 5.00
Begin VB.Form Vcodelast 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1080
      Top             =   1320
   End
End
Attribute VB_Name = "Vcodelast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If Main.tag3 And Main.tag4 Then
    If Int(Mid(Main.Text1.Text, 9, 1)) = Int(Mid(Main.Text1.Text, 1, 1)) / 3 And Int(Mid(Main.Text1.Text, 9, 1)) = Int(Mid(Main.Text1.Text, 3, 1)) - 6 Then
        If Int(Mid(Main.Text1.Text, 12, 1)) = Int(Mid(Main.Code3, 4, 1)) Then
            Main.tag5 = 1
            
            For i = 1 To 6
               If Int(Mid(Main.Text1.Text, i, 1)) < Int(Mid(Main.Text1.Text, 13, 1)) Then Main.tag5 = 0
            Next i
        End If
    End If
Else
    MsgBox "越权操作,认证失败", , "不许开挂"
End If
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub
