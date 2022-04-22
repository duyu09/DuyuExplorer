VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "版权信息"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5925
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5925
   StartUpPosition =   2  '屏幕中心
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "DuyuExplorer"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "学校官网：https://www.lcez.cn/"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "在桌面创建快捷方式"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "保留所有权利"
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
      Left            =   3960
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "济南市历城第二中学 55级31班 杜宇 NO.028"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2280
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   240
      Picture         =   "Form8.frx":058A
      Top             =   240
      Width           =   1920
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label3_Click()
Dim os As Object, zhuomian As String
Set os = CreateObject("WScript.Shell")
On Error Resume Next
Call ShortCut(os.SpecialFolders("Desktop") & "\Duyu资源管理器.lnk", App.Path & "\" & App.EXEName & ".exe")
MsgBox "创建完毕。", 48
End Sub

Private Sub Label4_Click()
Shell "explorer.exe https://www.lcez.cn/", vbHide
End Sub
