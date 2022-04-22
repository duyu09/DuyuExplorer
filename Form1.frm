VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl32.ocx"
Begin VB.Form Form1 
   Caption         =   "Duyu - Explorer"
   ClientHeight    =   8895
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13935
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   13935
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   22800
      TabIndex        =   33
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command12 
      Caption         =   "搜索"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   24600
      TabIndex        =   32
      Top             =   240
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   240
      Left            =   14160
      TabIndex        =   31
      Top             =   6000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   8040
      TabIndex        =   30
      Top             =   240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Command11 
      Caption         =   "刷新"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   29
      Top             =   240
      Width           =   855
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   375
      Left            =   19560
      TabIndex        =   27
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      Min             =   5
      Max             =   72
      SelStart        =   10
      Value           =   10
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   16680
      TabIndex        =   26
      Top             =   240
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Max             =   3
      SelStart        =   3
      Value           =   3
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   240
      TabIndex        =   10
      Top             =   4320
      Width           =   14295
      Begin VB.CommandButton Command14 
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12840
         TabIndex        =   36
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command13 
         Caption         =   "新建目录"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11520
         TabIndex        =   35
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11520
         TabIndex        =   34
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command10 
         Caption         =   "计算机控制"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10200
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         Caption         =   "编辑权限"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10200
         TabIndex        =   24
         Top             =   720
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   6960
         TabIndex        =   18
         Top             =   120
         Width           =   3135
         Begin VB.CommandButton Command8 
            Caption         =   "确定"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   23
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox Check4 
            Caption         =   "系统"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            TabIndex        =   22
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox Check3 
            Caption         =   "存档"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   960
            TabIndex        =   21
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox Check2 
            Caption         =   "只读"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox Check1 
            Caption         =   "隐藏"
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CommandButton Command7 
         Caption         =   "复制"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "删除"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "重命名"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4920
         TabIndex        =   14
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         Caption         =   "强制打开"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   615
         Left            =   120
         ScaleHeight     =   555
         ScaleWidth      =   555
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.Line Line3 
         X1              =   11400
         X2              =   11400
         Y1              =   240
         Y2              =   1080
      End
      Begin VB.Line Line1 
         X1              =   3360
         X2              =   3360
         Y1              =   240
         Y2              =   1080
      End
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   840
         TabIndex        =   12
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "转到"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   240
      Width           =   855
   End
   Begin VB.DirListBox Dir2 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   3015
   End
   Begin VB.DirListBox Dir1 
      Height          =   510
      Left            =   11760
      TabIndex        =   6
      Top             =   7080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      Height          =   450
      Hidden          =   -1  'True
      Left            =   12360
      System          =   -1  'True
      TabIndex        =   5
      Top             =   6000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   11880
      TabIndex        =   4
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   12600
      TabIndex        =   3
      Top             =   6600
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3495
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6165
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   240
      Width           =   5055
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   240
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   240
      Width           =   375
   End
   Begin VB.Line Line2 
      X1              =   22560
      X2              =   22560
      Y1              =   0
      Y2              =   960
   End
   Begin VB.Label Label2 
      Caption         =   "加载中......"
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
      Left            =   11880
      TabIndex        =   28
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long





'任务栏 进度条 定义部分
Private W7Task&
Private Enum ITaskbarList3
QueryInterface
AddRef
Release
'IUnknown
HrInit
AddTab
DeleteTab
ActivateTab
SetActiveAlt
MarkFullscreenWindow
SetProgressValue
SetProgressState
RegisterTab
UnregisterTab
SetTabOrder
SetTabActive
ThumbBarAddButtons
ThumbBarUpdateButtons
ThumbBarSetImageList
SetOverlayIcon
SetThumbnailTooltip
SetThumbnailClip
End Enum
Private Type GUID
Data1 As Long
Data2 As Integer
Data3 As Integer
Data4(7) As Byte
End Type
Private Declare Function IIDFromString& Lib "ole32 " (ByVal ID As Long, ByVal IDs As Long)
Private Declare Function CLSIDFromString& Lib "ole32 " (ByVal ID As Long, ByVal IDs As Long)
Private Declare Function CoCreateInstance& Lib "ole32 " (ByVal CLSID As Long, ByVal Outer As Long, ByVal Context As Long, ByVal IID As Long, Obj As Any) '
'任务栏 进度条 函数定义
Private Function CreateW7Task&()
Dim CID As GUID, IID As GUID, objW7Task&
CLSIDFromString StrPtr("{56FDF344-FD6D-11d0-958A-006097C9A090}"), VarPtr(CID)
IIDFromString StrPtr("{EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF}"), VarPtr(IID)
CoCreateInstance VarPtr(CID), 0, 1, VarPtr(IID), objW7Task
CreateW7Task = objW7Task
End Function






Private Sub Command1_Click()
Label2.Caption = "加载中......"
ListView1.Visible = False
Dim bcd As Long
bcd = File1.ListCount + Dir2.ListCount
Dim itmX As ListItem
Dim a As Double, b As Double, c As Integer, gs As Long, qsd As Long

'添加column1的名称。

For a = 1 To File1.ListCount
DoEvents

Dim MyFileSystem As New FileSystemObject
Dim MyFile As File
On Error Resume Next
Set MyFile = MyFileSystem.GetFile(File1.Path & "\" & File1.List(a - 1)) '对应的zd文件标识符
On Error Resume Next
Set itmX = ListView1.ListItems.Add(a, CStr(a) & "z", File1.List(a - 1))
On Error Resume Next
itmX.SubItems(ListView1.ColumnHeaders("DX").SubItemIndex) = FileLen(File1.Path & "\" & File1.List(a - 1)) & "   (" & CLng(FileLen(File1.Path & "\" & File1.List(a - 1)) / 1024) & "KB)"
On Error Resume Next
itmX.SubItems(ListView1.ColumnHeaders("XG").SubItemIndex) = MyFile.DateLastModified
On Error Resume Next
itmX.SubItems(ListView1.ColumnHeaders("CJ").SubItemIndex) = MyFile.DateCreated
On Error Resume Next
itmX.SubItems(ListView1.ColumnHeaders("FW").SubItemIndex) = MyFile.DateLastAccessed
On Error Resume Next
itmX.SubItems(ListView1.ColumnHeaders("LX").SubItemIndex) = MyFile.Type
On Error Resume Next
c = GetAttr(File1.Path & "\" & File1.List(a - 1))
If c = 1 Then
d = "只读"
ElseIf c = 0 Then
d = "常规"
ElseIf c = 2 Then
d = "隐藏"
ElseIf c = 4 Then
d = "系统文件"
ElseIf c = 16 Then
d = "目录"
ElseIf c = 32 Then
d = "存档"
ElseIf c = 3 Then
d = "隐藏 只读"
ElseIf c = 5 Then
d = "只读 系统文件"
ElseIf c = 17 Then
d = "只读目录"
ElseIf c = 33 Then
d = "只读 存档"
ElseIf c = 6 Then
d = "隐藏 系统"
ElseIf c = 18 Then
d = "隐藏目录"
ElseIf c = 34 Then
d = "隐藏 存档"
ElseIf c = 20 Then
d = "系统目录"
ElseIf c = 36 Then
d = "存档 系统"
ElseIf c = 48 Then
d = "存档目录"
ElseIf c = 7 Then
d = "只读 隐藏 系统文件"
ElseIf c = 19 Then
d = "只读 隐藏 目录"
ElseIf c = 35 Then
d = "存档 只读 隐藏"
End If
On Error Resume Next
itmX.SubItems(ListView1.ColumnHeaders("SX").SubItemIndex) = d

gs = gs + 1
qsd = gs
Me.ProgressBar1.Value = CLng(100 * gs / bcd)
frmMain.ProgressBar1.Value = CLng(100 * gs / bcd)
W7Task = CreateW7Task
'然后就可以设置进度了
CallCOMInterface W7Task, SetProgressValue, Me.hWnd, CLng(100 * gs / bcd), 0, 100, 0 '
'用自己的变量替换*进度*和*最大值*。
Next a
Dim fso1 As New FileSystemObject
Dim folder1 As Folder
For b = 1 To Dir2.ListCount
DoEvents
On Error Resume Next
Set itmX = ListView1.ListItems.Add(File1.ListCount + b, CStr(File1.ListCount + b) & "z", Dir2.List(b - 1))
On Error Resume Next
itmX.SubItems(ListView1.ColumnHeaders("LX").SubItemIndex) = "目录"
gs = gs + 1



  Set folder1 = fso1.GetFolder(Dir2.List(b - 1))
  itmX.SubItems(ListView1.ColumnHeaders("DX").SubItemIndex) = folder1.Size & "   (" & CLng(folder1.Size / 1024) & "KB)"
  
DoEvents

W7Task = CreateW7Task
'然后就可以设置进度了
CallCOMInterface W7Task, SetProgressValue, Me.hWnd, CLng(100 * gs / bcd), 0, 100, 0 '
'用自己的变量替换*进度*和*最大值*。
Me.ProgressBar1.Value = CLng(100 * gs / bcd)
frmMain.ProgressBar1.Value = CLng(100 * gs / bcd)
Next b
Label2.Caption = gs & " 个项目 (" & qsd & "个文件 " & CStr(gs - qsd) & "个目录)"
ListView1.Visible = True
Me.ProgressBar1.Value = 0
frmMain.ProgressBar1.Value = 0
W7Task = CreateW7Task
'然后就可以设置进度了
CallCOMInterface W7Task, SetProgressValue, Me.hWnd, 0, 0, 100, 0 '
'用自己的变量替换*进度*和*最大值*。
End Sub

Private Sub Command10_Click()
Form3.Show
End Sub

Private Sub Command11_Click()
 On Error Resume Next
tem = Dir2.Path
On Error Resume Next
Dir2.Path = "c:\"
On Error Resume Next
Dir2.Path = tem
End Sub

Private Sub Command12_Click()
Label2.Caption = "搜索中......"
ListView1.Visible = False
ListView1.ListItems.Clear
List1.Clear
Dim vn As Long, vnr As Long
For vn = 1 To File1.ListCount
If InStr(1, File1.List(vn - 1), Text3.Text) <> 0 Then List1.AddItem (File1.List(vn - 1))
Next vn
kq = vn
For vnr = 1 To Dir2.ListCount
If InStr(1, Dir2.List(vnr - 1), Text3.Text) <> 0 Then List1.AddItem (Dir2.List(vnr - 1))
Next vnr

Dim bcd As Long
bcd = File1.ListCount + Dir2.ListCount
Dim itmX As ListItem
Dim a As Double, b As Double, c As Integer, gs As Long, qsd As Long


For a = 1 To List1.ListCount
DoEvents
If InStr(1, List1.List(a - 1), "\") = 0 Then
Dim MyFileSystem As New FileSystemObject
Dim MyFile As File
On Error Resume Next
Set MyFile = MyFileSystem.GetFile(File1.Path & "\" & List1.List(a - 1)) '对应的zd文件标识符
On Error Resume Next
Set itmX = ListView1.ListItems.Add(a, CStr(a) & "z", List1.List(a - 1))
On Error Resume Next
itmX.SubItems(ListView1.ColumnHeaders("DX").SubItemIndex) = FileLen(File1.Path & "\" & List1.List(a - 1)) & "   (" & CLng(FileLen(File1.Path & "\" & List1.List(a - 1)) / 1024) & "KB)"
On Error Resume Next
itmX.SubItems(ListView1.ColumnHeaders("XG").SubItemIndex) = MyFile.DateLastModified
On Error Resume Next
itmX.SubItems(ListView1.ColumnHeaders("CJ").SubItemIndex) = MyFile.DateCreated
On Error Resume Next
itmX.SubItems(ListView1.ColumnHeaders("FW").SubItemIndex) = MyFile.DateLastAccessed
On Error Resume Next
itmX.SubItems(ListView1.ColumnHeaders("LX").SubItemIndex) = MyFile.Type
On Error Resume Next
c = GetAttr(File1.Path & "\" & List1.List(a - 1))
If c = 1 Then
d = "只读"
ElseIf c = 0 Then
d = "常规"
ElseIf c = 2 Then
d = "隐藏"
ElseIf c = 4 Then
d = "系统文件"
ElseIf c = 16 Then
d = "目录"
ElseIf c = 32 Then
d = "存档"
ElseIf c = 3 Then
d = "隐藏 只读"
ElseIf c = 5 Then
d = "只读 系统文件"
ElseIf c = 17 Then
d = "只读目录"
ElseIf c = 33 Then
d = "只读 存档"
ElseIf c = 6 Then
d = "隐藏 系统"
ElseIf c = 18 Then
d = "隐藏目录"
ElseIf c = 34 Then
d = "隐藏 存档"
ElseIf c = 20 Then
d = "系统目录"
ElseIf c = 36 Then
d = "存档 系统"
ElseIf c = 48 Then
d = "存档目录"
ElseIf c = 7 Then
d = "只读 隐藏 系统文件"
ElseIf c = 19 Then
d = "只读 隐藏 目录"
ElseIf c = 35 Then
d = "存档 只读 隐藏"
End If
On Error Resume Next
itmX.SubItems(ListView1.ColumnHeaders("SX").SubItemIndex) = d

gs = gs + 1
qsd = gs
Me.ProgressBar1.Value = CLng(100 * gs / bcd)


Else
On Error Resume Next
Set itmX = ListView1.ListItems.Add(a, CStr(a) & "z", List1.List(a - 1))
On Error Resume Next
itmX.SubItems(ListView1.ColumnHeaders("LX").SubItemIndex) = "目录"
On Error Resume Next
Dim fso1 As New FileSystemObject
Dim folder1 As Folder
Set folder1 = fso1.GetFolder(List1.List(a - 1))
On Error Resume Next
  itmX.SubItems(ListView1.ColumnHeaders("DX").SubItemIndex) = folder1.Size & "   (" & CLng(folder1.Size / 1024) & "KB)"
End If
Next a

Label2.Caption = "搜索结果：" & List1.ListCount & " 个项目"
ListView1.Visible = True
Me.ProgressBar1.Value = 0
End Sub

Private Sub Command13_Click()
If Text4.Text = "" Then
MsgBox "请输入目录名", 48
Exit Sub
End If
On Error Resume Next
MkDir File1.Path & "\" & Text4.Text
MsgBox "完毕。", vbInformation
Command11_Click
End Sub

Private Sub Command14_Click()
Form8.Show
End Sub

Private Sub Command2_Click()
Dim hImgSmall As Long
   Dim fName As String   '驱动器号、文件夹名、文件名
   Dim r As Long
   Dim hImgLarge As Long
   Dim Info1 As String, Info2 As String
   fName = Text1.Text
   hImgSmall& = SHGetFileInfo(fName$, 0&, shinfo, Len(shinfo), SHGFI_ICON Or SHGFI_SMALLICON Or SHGFI_SYSICONINDEX Or SHGFI_USEFILEATTRIBUTES)
   hImgLarge& = SHGetFileInfo(fName$, 0&, shinfo, Len(shinfo), SHGFI_ICON Or BASIC_SHGFI_FLAGS Or SHGFI_SYSICONINDEX Or SHGFI_USEFILEATTRIBUTES)
   Info1 = Left$(shinfo.szDisplayName, InStr(shinfo.szDisplayName, Chr$(0)) - 1)
   Info2 = Left$(shinfo.szTypeName, InStr(shinfo.szTypeName, Chr$(0)) - 1)
   Debug.Print Info1; Info2
   Picture1.Picture = LoadPicture()
   Picture1.AutoRedraw = True
   Picture2.Picture = LoadPicture()
   Picture2.AutoRedraw = True
   r = ImageList_Draw(hImgSmall&, shinfo.iIcon, Picture1.hDC, 0, 0, ILD_TRANSPARENT)
   r = ImageList_Draw(hImgLarge&, shinfo.iIcon, Picture2.hDC, 3, 3, ILD_TRANSPARENT)
   Set Picture1.Picture = Picture1.Image
   Set Picture2.Picture = Picture2.Image
End Sub



Private Sub Command3_Click()
On Error Resume Next
Dir2.Path = Text1.Text
If Err.Number > 0 Then
asd = MsgBox(Err.Description, 48)
End If
End Sub

Private Sub Command4_Click()
qw = Shell("cmd /c " & Chr(34) & Dir2.Path & "\" & Label1.Caption & Chr(34), vbHide)
End Sub

Private Sub Command5_Click()
Dim ee As Integer, cvb As String
ee = MsgBox("确定重命名？", vbOKCancel)
cvb = Label1.Caption
If ee <> 1 Then Exit Sub
If cvb <> Text2.Text Then
 On Error Resume Next
 Name Dir2.Path & "\" & cvb As Dir2.Path & "\" & Text2.Text
 On Error Resume Next
tem = Dir2.Path
On Error Resume Next
Dir2.Path = "c:\"
On Error Resume Next
Dir2.Path = tem
End If
End Sub

Private Sub Command6_Click()
Dim ee As Integer, cvb As String
ee = MsgBox("确定删除？ " & Label1.Caption, vbOKCancel)
If ee <> 1 Then Exit Sub
 If InStr(1, Label1.Caption, "\") = 0 Then
 On Error Resume Next
 Kill (Dir2.Path & "\" & Label1.Caption)
 Else
 On Error Resume Next
 Kill (Label1.Caption & "\*")
 On Error Resume Next
 RmDir (Label1.Caption)
 End If

On Error Resume Next
tem = Dir2.Path
On Error Resume Next
Dir2.Path = "c:\"
On Error Resume Next
Dir2.Path = tem
End Sub

Private Sub Command7_Click()
Form2.Show
End Sub

Private Sub Command8_Click()
Dim a As Integer
If Check1.Value = 1 Then
a = a + 2
End If
If Check2.Value = 1 Then a = a + 1
If Check3.Value = 1 Then a = a + 32
If Check4.Value = 1 Then a = a + 4
On Error Resume Next
SetAttr Dir2.Path & "\" & Label1.Caption, a
If Err.Number = 0 Then MsgBox ("属性修改成功。") Else MsgBox ("属性修改失败。")
On Error Resume Next
tem = Dir2.Path
On Error Resume Next
Dir2.Path = "c:\"
On Error Resume Next
Dir2.Path = tem
End Sub

Private Sub Command9_Click()
If Form1.Label1.Caption = "" Then
MsgBox "请首先选中文件", 48
Exit Sub
End If
Form7.Show
End Sub

Private Sub Dir2_Change()
 ListView1.ListItems.Clear
 File1.Path = Dir2.Path
 Text1.Text = Dir2.Path
 Drive1.Drive = Left(Dir2.Path, 1)
 Command1_Click
 Me.Caption = "Duyu - Explorer - " & Dir2.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir2.Path = Drive1.Drive
If Err.Number > 0 Then
asd = MsgBox(Err.Description, 48)
End If
End Sub

Private Sub Form_Load()
Me.Show
Me.WindowState = 2
ListView1.ColumnHeaders.Add , "Name", "文件名", 3000
ListView1.ColumnHeaders.Add , "LX", "文件类型", 1777
ListView1.ColumnHeaders.Add , "SX", "文件属性", 2000
ListView1.ColumnHeaders.Add , "DX", "文件大小", 2788
ListView1.ColumnHeaders.Add , "XG", "修改日期", 2222
ListView1.ColumnHeaders.Add , "CJ", "创建日期", 2222
ListView1.ColumnHeaders.Add , "FW", "访问日期", 2222
Text1.Text = App.Path
Command3_Click
Command11_Click
End Sub

Private Sub Form_Resize()
On Error Resume Next
ListView1.Width = Me.Width - 3880
On Error Resume Next
Frame1.Top = Me.Height - 2215
On Error Resume Next
Frame1.Width = Me.Width - 650
On Error Resume Next
ListView1.Height = Me.Height - 3000
On Error Resume Next
Dir2.Height = Me.Height - 3250
On Error Resume Next
If Me.WindowState = 1 Then
frmMain.Show
Me.Hide
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub ListView1_AfterLabelEdit(Cancel As Integer, NewString As String)
 ListView1.ListItems(ListView1.SelectedItem.Index).Text = Text2.Text
 Command5_Click
End Sub


Private Sub ListView1_Click()
On Error Resume Next
Label1.Caption = ListView1.SelectedItem.Text
Text2.Text = ListView1.SelectedItem.Text
Fn = ListView1.SelectedItem.Text
Dim hImgSmall As Long
   Dim fName As String   '驱动器号、文件夹名、文件名
   Dim r As Long
   Dim hImgLarge As Long
   Dim Info1 As String, Info2 As String
   fName = Dir2.Path & "\" & Label1.Caption
   hImgSmall& = SHGetFileInfo(fName$, 0&, shinfo, Len(shinfo), SHGFI_ICON Or SHGFI_SMALLICON Or SHGFI_SYSICONINDEX Or SHGFI_USEFILEATTRIBUTES)
   hImgLarge& = SHGetFileInfo(fName$, 0&, shinfo, Len(shinfo), SHGFI_ICON Or BASIC_SHGFI_FLAGS Or SHGFI_SYSICONINDEX Or SHGFI_USEFILEATTRIBUTES)
   Info1 = Left$(shinfo.szDisplayName, InStr(shinfo.szDisplayName, Chr$(0)) - 1)
   Info2 = Left$(shinfo.szTypeName, InStr(shinfo.szTypeName, Chr$(0)) - 1)
   Debug.Print Info1; Info2
   Picture1.Picture = LoadPicture()
   Picture1.AutoRedraw = True
   Picture2.Picture = LoadPicture()
   Picture2.AutoRedraw = True
   r = ImageList_Draw(hImgSmall&, shinfo.iIcon, Picture1.hDC, 0, 0, ILD_TRANSPARENT)
   r = ImageList_Draw(hImgLarge&, shinfo.iIcon, Picture2.hDC, 3, 3, ILD_TRANSPARENT)
   Set Picture1.Picture = Picture1.Image
   Set Picture2.Picture = Picture2.Image
   On Error Resume Next
   ListView1.ToolTipText = ListView1.ListItems(ListView1.SelectedItem.Index).Text
End Sub



Private Sub ListView1_DblClick()
If ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(1) <> "目录" Then
Text1.Text = Dir2.Path & "\" & ListView1.SelectedItem.Text
Command2_Click
On Error Resume Next
ShellExecute Me.hWnd, "open", Dir2.Path & "\" & ListView1.SelectedItem.Text, "", Dir2.Path, 5
Else
On Error Resume Next
Text1.Text = ListView1.SelectedItem.Text
Command2_Click
On Error Resume Next
Dir2.Path = ListView1.SelectedItem.Text
End If
End Sub


'新建一个窗体，在窗体上添加一个TextBox用来输入文件路径
'和两个picturebox用来显示提取到的图标
'以下是窗体中的代码
Private Sub Picture2_Click()
VB.SavePicture Picture2, App.Path & "\ico.ico"
End Sub
 



Private Sub Slider1_Change()
On Error Resume Next
ListView1.View = Slider1.Value
End Sub



Private Sub Slider2_Change()
On Error Resume Next
ListView1.Font.Size = Slider2.Value
End Sub

