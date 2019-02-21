VERSION 5.00
Begin VB.Form VCode 
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
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2640
      Top             =   960
   End
End
Attribute VB_Name = "VCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If Main.tag3 And Int(Mid(Main.Text1.Text, 3, 1)) = Int(Mid(Main.Text1.Text, 9, 1)) + Int(Mid(Main.Text1.Text, 10, 1)) Then
    If Int(Mid(Main.Text1.Text, 4, 1)) = Int(Mid(Main.Text1.Text, 11, 1)) + Int(Mid(Main.Text1.Text, 13, 1)) Then
        If Int(Mid(Main.Text1.Text, 5, 1)) = Int(Mid(Main.Text1.Text, 13, 1)) * 2 Then
            If (Mid(Main.Text1.Text, 6, 1) = Mid(Main.Text1.Text, 8, 1)) Then
                        Main.tag4 = 1
                        Vcodelast.Show 1
            End If
        End If
    End If
End If
'MsgBox Mid(Main.Text1.Text, 6, 1) & Mid(Main.Text1.Text, 7, 1)
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub

