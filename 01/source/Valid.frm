VERSION 5.00
Begin VB.Form Valid 
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
      Left            =   1800
      Top             =   1320
   End
End
Attribute VB_Name = "Valid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Num(9), tempTag

Private Sub Form_Load()
'On Error Resume Next
For i = 0 To 9
    Num(i) = 0
Next i

For i = 1 To 13
    'MsgBox Mid(Main.Text1.Text, i, 1)
    Num(Mid(Main.Text1.Text, i, 1)) = 1
Next i

tempTag = 1

For i = 0 To 9
    'MsgBox tempTag
    If Num(i) = 0 Then tempTag = 0
    'MsgBox "in" & i
Next i

Main.tag2 = tempTag
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub
