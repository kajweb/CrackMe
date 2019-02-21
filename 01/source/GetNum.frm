VERSION 5.00
Begin VB.Form GetNum 
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
      Left            =   3000
      Top             =   1920
   End
End
Attribute VB_Name = "GetNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If Len(Main.Text1.Text) = 13 Then
    Main.tag1 = 2
Else
    Main.tag1 = 1
End If

End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub

