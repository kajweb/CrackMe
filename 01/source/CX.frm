VERSION 5.00
Begin VB.Form CX 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ڶ�����֤"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4140
   Icon            =   "CX.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   0
      Picture         =   "CX.frx":426C2
      ScaleHeight     =   1875
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.Timer Timer2 
         Interval        =   5000
         Left            =   3360
         Top             =   960
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   3360
         Top             =   240
      End
   End
End
Attribute VB_Name = "CX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim times

Private Sub Timer1_Timer()
times = times + 1
If times <= 10 Then
Me.Caption = "���ڶ�����֤..." & 10 - times
Else
Timer1.Enabled = False
Me.Caption = "��ϲ�ƽ�ɹ�"
MsgBox "��ϲ,�ƽ�ɹ�", , "MFC - �ɹ�"
End If
End Sub

Private Sub Timer2_Timer()
Timer2.Enabled = False
If Main.tag1 = 2 And Main.tag2 And Main.tag3 And Main.tag4 And Main.tag5 Then
Timer1.Interval = 500
Else
Timer1.Enabled = False
Me.Caption = "���ź����ƽ�ʧ��"
MsgBox "�ƽ�ʧ��", , "MFC - ʧ��"
End If
End Sub
