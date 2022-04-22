VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "DUYU - 把进度条放到任务栏里"
   ClientHeight    =   2250
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4155
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   4155
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Begin VB.Menu mnuExit 
         Caption         =   "退出DuyuExplorer"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProgress 
         Caption         =   "显示进度条"
      End
      Begin VB.Menu mnuCommand 
         Caption         =   "前置按钮"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Download by http://www.NewXing.com
Private Sub mnuCommand_Click()
    frmMain.ProgressBar1.Visible = False
    frmMain.cmdMain.Visible = True
    frmMain.txtMain.Visible = False
End Sub

Private Sub mnuExit_Click()
Dim a As Integer
a = MsgBox("即将退出软件，是否启动Windows Explorer(资源管理器)?", vbYesNo)
If a = 6 Then Shell "explorer.exe", vbNormalFocus
    Unload frmMain
    Unload Me
    End
End Sub

Private Sub mnuProgress_Click()
    frmMain.ProgressBar1.Visible = True
    frmMain.cmdMain.Visible = False
    frmMain.txtMain.Visible = False
End Sub

