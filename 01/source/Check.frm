VERSION 5.00
Begin VB.Form Check 
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
      Interval        =   500
      Left            =   1440
      Top             =   1920
   End
End
Attribute VB_Name = "Check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Unload Main
If Main.tag1 = "2" And Main.tag2 And Main.tag3 And Main.tag4 And Main.tag5 Then
CX.Show
Else
Main.APPError = 1
Main.Show
End If
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub
