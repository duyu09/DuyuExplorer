VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "权限设置"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   7710
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame2 
      Caption         =   "赋予权限"
      Height          =   2175
      Left            =   120
      TabIndex        =   10
      Top             =   4920
      Width           =   7455
      Begin VB.CommandButton Command2 
         Caption         =   "确定"
         Height          =   495
         Left            =   4560
         TabIndex        =   16
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CheckBox Check10 
         Caption         =   "完全控制"
         Height          =   375
         Left            =   2400
         TabIndex        =   15
         Top             =   480
         Width           =   1695
      End
      Begin VB.CheckBox Check9 
         Caption         =   "读取（包括执行）"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CheckBox Check8 
         Caption         =   "更改（写入）"
         Height          =   315
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox Check7 
         Caption         =   "写入"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "撤销权限"
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   7455
      Begin VB.CheckBox Check4 
         Caption         =   "更改（写入）"
         Height          =   315
         Left            =   2520
         TabIndex        =   4
         Top             =   480
         Width           =   1935
      End
      Begin VB.CheckBox Check3 
         Caption         =   "写入"
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "确定"
         Height          =   495
         Left            =   4560
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CheckBox Check6 
         Caption         =   "拒绝访问"
         Height          =   375
         Left            =   4560
         TabIndex        =   6
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox Check5 
         Caption         =   "完全控制"
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         Caption         =   "读取（包括执行）"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   2415
      End
      Begin VB.CheckBox Check1 
         Caption         =   "无权"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Label Label2 
      Caption         =   "用户名："
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7335
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "请输入用户名", 48
Exit Sub
End If
Dim pa As String
If InStr(1, Form1.Label1.Caption, "\") = 0 Then
pa = Form1.Dir2.Path & "\" & Form1.Label1.Caption
Else
pa = Form1.Label1.Caption
End If

If Check1.Value = 1 Then
Shell "cmd /c y|cacls " & Chr(34) & pa & Chr(34) & " /E /p" & Text1.Text & ":n", vbHide
End If

If Check2.Value = 1 Then
Shell "cmd /c y|cacls " & Chr(34) & pa & Chr(34) & " /E /p" & Text1.Text & ":r", vbHide
End If

If Check3.Value = 1 Then
Shell "cmd /c y|cacls " & Chr(34) & pa & Chr(34) & " /E /p" & Text1.Text & ":w", vbHide
End If

If Check4.Value = 1 Then
Shell "cmd /c y|cacls " & Chr(34) & pa & Chr(34) & " /E /p" & Text1.Text & ":c", vbHide
End If

If Check5.Value = 1 Then
Shell "cmd /c y|cacls " & Chr(34) & pa & Chr(34) & " /E /p" & Text1.Text & ":f", vbHide
End If

If Check6.Value = 1 Then
Shell "cmd /c y|cacls " & Chr(34) & pa & Chr(34) & " /E /d" & Text1.Text, vbHide
End If
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Then
MsgBox "请输入用户名", 48
Exit Sub
End If
Dim pa As String
If InStr(1, Form1.Label1.Caption, "\") = 0 Then
pa = Form1.Dir2.Path & "\" & Form1.Label1.Caption
Else
pa = Form1.Label1.Caption
End If

If Check1.Value = 7 Then
Shell "cmd /c y|cacls " & Chr(34) & pa & Chr(34) & " /E /g" & Text1.Text & ":w", vbHide
End If

If Check2.Value = 8 Then
Shell "cmd /c y|cacls " & Chr(34) & pa & Chr(34) & " /E /g" & Text1.Text & ":c", vbHide
End If

If Check3.Value = 9 Then
Shell "cmd /c y|cacls " & Chr(34) & pa & Chr(34) & " /E /g" & Text1.Text & ":r", vbHide
End If

If Check4.Value = 10 Then
Shell "cmd /c y|cacls " & Chr(34) & pa & Chr(34) & " /E /g" & Text1.Text & ":f", vbHide
End If

End Sub

Private Sub Form_Load()
If InStr(1, Form1.Label1.Caption, "\") = 0 Then
Label1.Caption = "修改" & Form1.Dir2.Path & "\" & Form1.Label1.Caption & "的权限"
Else
Label1.Caption = "修改" & Form1.Label1.Caption & "的权限"
End If
End Sub
