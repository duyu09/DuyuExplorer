VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "DUYU - �ѽ������ŵ���������"
   ClientHeight    =   2250
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4155
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   4155
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Begin VB.Menu mnuExit 
         Caption         =   "�˳�DuyuExplorer"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProgress 
         Caption         =   "��ʾ������"
      End
      Begin VB.Menu mnuCommand 
         Caption         =   "ǰ�ð�ť"
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
a = MsgBox("�����˳�������Ƿ�����Windows Explorer(��Դ������)?", vbYesNo)
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

