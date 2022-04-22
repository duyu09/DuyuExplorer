VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "CPG-WFU"
   ClientHeight    =   2205
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4425
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "单击右键以切换模式"
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.TextBox txtMain 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Text            =   "Type anything in here..."
      ToolTipText     =   "单击右键以切换模式"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cmdMain 
      BackColor       =   &H000000FF&
      Caption         =   "单击前置DuyuExplorer"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      ToolTipText     =   "单击右键以切换模式"
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "DuyuExplorer"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const PBM_SETBARCOLOR = &H409
Private Const PBM_SETBKCOLOR = &H2001

'Download by http://www.NewXing.com
Private Sub Form_Load()
    Me.Show
    Call AttachForm(Me, 100, 30, True)
    frmMain.txtMain.Visible = False
    cmdMain.Visible = False

    PostMessage ProgressBar1.hWnd, PBM_SETBARCOLOR, 0, vbRed
  PostMessage ProgressBar1.hWnd, PBM_SETBKCOLOR, 0, vbGreen
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call DetachForm
End Sub

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'display the popmenu if this control is clicked
    If Button = 2 Then frmMenu.PopupMenu frmMenu.mnuMain
End Sub

Private Sub cmdMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'display the popmenu if this control is right-clicked
    If Button = 2 Then Me.PopupMenu frmMenu.mnuMain
End Sub

Private Sub cmdMain_Click()
    Form1.Show
    Form1.WindowState = 2
End Sub



Private Sub Form_Paint()
    Me.Hide
    With ProgressBar1
        .Top = 0
        .Left = 0
        .Width = Me.Width
        .Height = Me.Height / 2
    End With
    With cmdMain
        .Top = 0
        .Left = 0
        .Width = Me.Width
        .Height = Me.Height
    End With
    With txtMain
        .Top = 0
        .Left = 0
        .Width = Me.Width
        .Height = Me.Height
    End With
    With Label1
        .Top = Me.Height / 2
        .Left = 0
        .Width = Me.Width
        .Height = Me.Height / 2
    End With
    Me.Show
End Sub

Private Sub Form_Resize()
    Call Form_Paint
End Sub



