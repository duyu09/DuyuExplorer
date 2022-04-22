VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl32.ocx"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "÷˜…˘ø®“Ù¡øøÿ÷∆"
   ClientHeight    =   3990
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   7080
   Icon            =   "mixer.frx":0000
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   7080
   StartUpPosition =   2  '∆¡ƒª÷––ƒ
   Begin VB.Frame Frame1 
      Caption         =   "…˘ø®…˘“Ùøÿ÷∆"
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin VB.Frame Frame4 
         Caption         =   "Midi“Ù¡ø"
         Height          =   3255
         Left            =   4680
         TabIndex        =   9
         Top             =   240
         Width           =   2055
         Begin VB.CheckBox Check3 
            Caption         =   "æ≤“Ù(&Mute)"
            Height          =   375
            Left            =   240
            TabIndex        =   10
            Top             =   2760
            Width           =   1455
         End
         Begin MSComctlLib.Slider Slider5 
            Height          =   1815
            Left            =   720
            TabIndex        =   11
            Top             =   840
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   3201
            _Version        =   393216
            Orientation     =   1
            Min             =   -100
            Max             =   0
            SelStart        =   -100
            TickStyle       =   2
            TickFrequency   =   10
            Value           =   -100
            TextPosition    =   1
         End
         Begin MSComctlLib.Slider Slider6 
            Height          =   495
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   873
            _Version        =   393216
            Max             =   100
            SelStart        =   50
            TickStyle       =   2
            TickFrequency   =   10
            Value           =   50
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Wave&&Mp3“Ù¡ø"
         Height          =   3255
         Left            =   2400
         TabIndex        =   5
         Top             =   240
         Width           =   2055
         Begin VB.CheckBox Check2 
            Caption         =   "æ≤“Ù(&Mute)"
            Height          =   375
            Left            =   240
            TabIndex        =   6
            Top             =   2760
            Width           =   1455
         End
         Begin MSComctlLib.Slider Slider3 
            Height          =   1815
            Left            =   720
            TabIndex        =   7
            Top             =   840
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   3201
            _Version        =   393216
            Orientation     =   1
            Min             =   -100
            Max             =   0
            SelStart        =   -100
            TickStyle       =   2
            TickFrequency   =   10
            Value           =   -100
            TextPosition    =   1
         End
         Begin MSComctlLib.Slider Slider4 
            Height          =   495
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   873
            _Version        =   393216
            Max             =   100
            SelStart        =   50
            TickStyle       =   2
            TickFrequency   =   10
            Value           =   50
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Master“Ù¡ø"
         Height          =   3255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2055
         Begin VB.CheckBox Check1 
            Caption         =   "æ≤“Ù(&Mute)"
            Height          =   375
            Left            =   240
            TabIndex        =   4
            Top             =   2760
            Width           =   1455
         End
         Begin MSComctlLib.Slider Slider2 
            Height          =   1815
            Left            =   720
            TabIndex        =   3
            Top             =   840
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   3201
            _Version        =   393216
            Orientation     =   1
            Min             =   -100
            Max             =   0
            SelStart        =   -100
            TickStyle       =   2
            TickFrequency   =   10
            Value           =   -100
            TextPosition    =   1
         End
         Begin MSComctlLib.Slider Slider1 
            Height          =   495
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   873
            _Version        =   393216
            Max             =   100
            SelStart        =   50
            TickStyle       =   2
            TickFrequency   =   10
            Value           =   50
         End
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MasterVolume_DeviceID = 0
Private Const WaveVolume_DeviceID = 1
Private Const MidiVolume_DeviceID = 2
Private Const BeMute = True
Private Const UnMute = False
Dim MySetVolume As New Mixer

Private Sub Check2_Click()
  If Check2.Value Then
    MySetVolume.Mute WaveVolume_DeviceID, BeMute
  Else
    MySetVolume.Mute WaveVolume_DeviceID, UnMute
  End If
End Sub
Private Sub Check3_Click()
  If Check3.Value Then
    MySetVolume.Mute MidiVolume_DeviceID, BeMute
  Else
    MySetVolume.Mute MidiVolume_DeviceID, UnMute
  End If
End Sub

Private Sub Check1_Click()
  If Check1.Value Then
    MySetVolume.Mute MasterVolume_DeviceID, BeMute
  Else
    MySetVolume.Mute MasterVolume_DeviceID, UnMute
  End If
End Sub

Private Sub Form_Load()
If MySetVolume.IsControl = False Then
   End
End If
End Sub
Private Sub Slider1_Scroll()
    MySetVolume.SetVolume Slider1.Value, -Slider2.Value, MasterVolume_DeviceID
End Sub
Private Sub Slider2_Scroll()
   MySetVolume.SetVolume Slider1.Value, -Slider2.Value, MasterVolume_DeviceID
End Sub

Private Sub Slider3_Scroll()
   MySetVolume.SetVolume Slider4.Value, -Slider3.Value, WaveVolume_DeviceID
End Sub

Private Sub Slider4_Scroll()
   MySetVolume.SetVolume Slider4.Value, -Slider3.Value, WaveVolume_DeviceID
End Sub

Private Sub Slider5_Scroll()
   MySetVolume.SetVolume Slider6.Value, -Slider5.Value, MidiVolume_DeviceID
End Sub

Private Sub Slider6_Scroll()
   MySetVolume.SetVolume Slider6.Value, -Slider5.Value, MidiVolume_DeviceID
End Sub
