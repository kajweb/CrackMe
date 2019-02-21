VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MFC - CrackMe - 吾爱破解"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13515
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":426C2
   ScaleHeight     =   8460
   ScaleWidth      =   13515
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      MaxLength       =   13
      TabIndex        =   1
      Top             =   6480
      Width           =   5535
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   9600
      Picture         =   "Form1.frx":4D34B
      Top             =   5760
      Width           =   1740
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CrackMe - By kajweb"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   2520
      TabIndex        =   0
      Top             =   840
      Width           =   9015
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tag1, tag2, tag3, tag4, tag5, tag6, tag7, tag8
Public Code_T '乱码序组
Dim Code2(2)
Public Code, Code1, Code3, APPError
Const IO = 1 '测试时为1，使用时为0，与程序时序有关

Private Sub Form_Load()
On Error Resume Next
If APPError = 1 Then APPError = 0: MsgBox "注册码错误", , "MFC - 失败"
Code0 = "1597531597531"
checkCode.Show IO '空代码
Code1 = "8949892480891"
Code3 = "1790470937485"
Code2(1) = "4984013068870"
checkCode.Show IO '空代码
Code_T = "1024935716"
Code2(0) = "0689063081710"
checkCode.Show IO '空代码
Code2(2) = "9203278093580"
End Sub

Private Sub Image1_Click()
If Len(Text1.Text) <> 13 Then 失败: Exit Sub
For i = 1 To 13
    If Asc(Mid(Text1.Text, i, 1)) < 48 Or Asc(Mid(Text1.Text, i, 1)) > 57 Then 失败: Exit Sub
Next i
On Error Resume Next
tag1 = 0
'On Error Resume Next
checkCode.Show IO '空代码
'Exit Sub

GetNum.Show 1 '检测是否13位   tag1
Valid.Show 1 '检查是否有0-9   tag2
checkCode.Show IO '空代码
If Left(Text1.Text, 2) = "94" Then tag3 = 1: VCode.Show 1 'tag3前两位，   tag4 3-6    tag5 8-13
checkCode.Show IO '空代码
If tag1 & tag2 & tag3 & tag4 & tag5 Then
    Check.Show
Else
    Check.Show
End If
End Sub

Sub 失败()
MsgBox "注册码应为13位纯数字", , "MFC - 输入错误"
End Sub

