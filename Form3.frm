VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Duyu Explorer - 计算机控制区"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5355
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5355
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command13 
      Caption         =   "合成器"
      Height          =   495
      Left            =   2880
      TabIndex        =   12
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command12 
      Caption         =   "系统音量"
      Height          =   495
      Left            =   1920
      TabIndex        =   11
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Command11 
      Caption         =   "暴力关闭计算机"
      Height          =   615
      Left            =   3720
      TabIndex        =   10
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
      Caption         =   "安全关闭计算机"
      Height          =   495
      Left            =   3600
      TabIndex        =   9
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      Caption         =   "注销"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "重新启动计算机"
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "打开资源管理器"
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "打开网络和共享中心"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   7.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "调整日期时间"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "打开资源监视器"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "打开 Duyu临时任务管理器"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "打开任务管理器"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打开cmd"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ExitWindowsEx Lib "user32" _
(ByVal dwOptions As Long, _
ByVal dwReserved As Long) As Long
Option Explicit
Private Const EWX_LogOff As Long = 0
Private Const EWX_SHUTDOWN As Long = 1
Private Const EWX_REBOOT As Long = 2
Private Const EWX_FORCE As Long = 4
Private Const EWX_POWEROFF As Long = 8


Private Declare Function RtlAdjustPrivilege& Lib "ntdll" (ByVal Privilege&, ByVal Newvalue&, ByVal NewThread&, Oldvalue&)
Private Declare Function NtShutdownSystem& Lib "ntdll" (ByVal ShutdownAction&)
Const SE_SHUTDOWN_PRIVILEGE& = 19
Const SHUTDOWN& = 0
Const RESTART& = 1
Const POWEROFF& = 2

Private Sub Command1_Click()
Shell "cmd.exe", vbNormalFocus
End Sub

Private Sub Command10_Click()
On Error Resume Next
Call ExitWindowsEx(EWX_SHUTDOWN, 0)
On Error Resume Next
Shell ("shutdown /s")
End Sub

Private Sub Command11_Click()
RtlAdjustPrivilege& SE_SHUTDOWN_PRIVILEGE&, 1, 0, 0 '提升权限
NtShutdownSystem& SHUTDOWN& Or POWEROFF& '关机
On Error Resume Next
Shell "shutdown -s"
End Sub

Private Sub Command12_Click()
Form4.Show
End Sub

Private Sub Command13_Click()
Shell "sndvol.exe -r", vbNormalFocus
End Sub

Private Sub Command2_Click()
Shell "taskmgr.exe", vbNormalFocus
End Sub

Private Sub Command3_Click()
Form5.Show
End Sub

Private Sub Command4_Click()
Shell "perfmon.exe /res", vbNormalFocus
End Sub

Private Sub Command5_Click()
Form6.Show
End Sub

Private Sub Command6_Click()
Shell "explorer.exe " & Chr(34) & "控制面板\网络和 Internet\网络和共享中心" & Chr(34), vbNormalFocus
End Sub

Private Sub Command7_Click()
Shell "explorer.exe " & Chr(34) & App.Path & Chr(34), vbNormalFocus
End Sub

Private Sub Command8_Click()
On Error Resume Next
Call ExitWindowsEx(EWX_REBOOT, 0)
On Error Resume Next
Shell ("shutdown /r")
End Sub

Private Sub Command9_Click()
On Error Resume Next
Call ExitWindowsEx(EWX_LogOff, 0)
On Error Resume Next
Shell ("shutdown /l")
End Sub
