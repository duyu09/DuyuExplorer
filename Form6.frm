VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Duyu - 设置日期与时间"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5610
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   5610
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      TabIndex        =   16
      Top             =   600
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "时间"
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   4695
      Begin VB.TextBox Text7 
         Height          =   495
         Left            =   3000
         TabIndex        =   11
         Text            =   "00"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   2160
         TabIndex        =   10
         Text            =   "Text6"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   1200
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Text            =   "Text4"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "*10毫秒"
         Height          =   495
         Left            =   3600
         TabIndex        =   15
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "秒"
         Height          =   495
         Left            =   2760
         TabIndex        =   14
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "分"
         Height          =   495
         Left            =   1920
         TabIndex        =   13
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "时"
         Height          =   495
         Left            =   960
         TabIndex        =   12
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "日期"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   1680
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "日"
         Height          =   495
         Left            =   3360
         TabIndex        =   6
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "月"
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
         Left            =   2400
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "年"
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
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   255
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell "cmd /c time " & Val(Text4.Text) & ":" & Val(Text5.Text) & ":" & Val(Text6.Text) & "." & Val(Text7.Text), vbHide
Shell "cmd /c date " & Val(Text1.Text) & "-" & Val(Text2.Text) & "-" & Val(Text3.Text), vbHide
End Sub

Private Sub Form_Load()
Text1.Text = Year(Now)
Text2.Text = Month(Now)
Text3.Text = Day(Now)
Text4.Text = Hour(Now)
Text5.Text = Minute(Now)
Text6.Text = Second(Now)
End Sub
