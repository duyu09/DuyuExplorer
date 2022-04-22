VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'声明关于对声卡控制的APIs
Private Declare Function mixerClose& Lib "winmm.dll" (ByVal hmx&)
Private Declare Function mixerGetControlDetails& Lib "winmm.dll" Alias "mixerGetControlDetailsA" (ByVal hmxobj&, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails&)
Private Declare Function mixerGetDevCaps& Lib "winmm.dll" Alias "mixerGetDevCapsA" (ByVal uMxId&, pmxcaps As MIXERCAPS, ByVal cbmxcaps&)
Private Declare Function mixerGetID& Lib "winmm.dll" (ByVal hmxobj&, pumxID&, ByVal fdwId&)
Private Declare Function mixerGetLineControls& Lib "winmm.dll" Alias "mixerGetLineControlsA" (ByVal hmxobj&, pmxlc As MIXERLINECONTROLS, ByVal fdwControls&)
Private Declare Function mixerGetLineInfo& Lib "winmm.dll" Alias "mixerGetLineInfoA" (ByVal hmxobj&, pmxl As MIXERLINE, ByVal fdwInfo&)
Private Declare Function mixerGetNumDevs& Lib "winmm.dll" ()
Private Declare Function mixerMessage& Lib "winmm.dll" (ByVal hmx&, ByVal umsg&, ByVal dwParam1&, ByVal dwParam2&)
Private Declare Function mixerOpen& Lib "winmm.dll" (phmx&, ByVal uMxId&, ByVal dwCallback&, ByVal dwInstance&, ByVal fdwOpen&)
Private Declare Function mixerSetControlDetails& Lib "winmm.dll" (ByVal hmxobj&, pmxcd As MIXERCONTROLDETAILS, ByVal fdwDetails&)
'自定义声卡控制类型
Private Type MIXERCAPS
     wMid As Integer
     wPid As Integer
     vDriverVersion As Long
     szPname As String * 32
     fdwSupport As Long
     cDestinations As Long
End Type
Private Type MIXERCONTROL
     cbStruct As Long
     dwControlID As Long
     dwControlType As Long
     fdwControl As Long
     cMultipleItems As Long
     szShortName As String * 16
     szName As String * 64
     lMinimum As Long
     lMaximum As Long
     Reserved(10) As Long
End Type
Private Type MIXERCONTROLDETAILS
     cbStruct As Long
     dwControlID As Long
     cChannels As Long
     item As Long
     cbDetails As Long
     paDetails As Long
End Type
Private Type MIXERCONTROLDETAILS_BOOLEAN
     fValue As Long
End Type
Private Type MIXERCONTROLDETAILS_LISTTEXT
     dwParam1 As Long
     dwParam2 As Long
     szName As String * 64
End Type
Private Type MIXERCONTROLDETAILS_SIGNED
     lValue As Long
End Type
Private Type MIXERCONTROLDETAILS_UNSIGNED
     dwValue As Long
End Type
Private Type Target
     dwType As Long
     dwDeviceID As Long
     wMid As Integer
     wPid As Integer
     vDriverVersion As Long
     szPname As String * 32
End Type
Private Type MIXERLINE
     cbStruct As Long
     dwDestination As Long
     dwSource As Long
     dwLineID As Long
     fdwLine As Long
     dwUser As Long
     dwComponentType As Long
     cChannels As Long
     cConnections As Long
     cControls As Long
     szShortName As String * 16
     szName As String * 64
     lpTarget As Target
End Type
Private Type MIXERLINECONTROLS
     cbStruct As Long
     dwLineID As Long
     dwControl As Long
     cControls As Long
     cbmxctrl As Long
     pamxctrl As Long
End Type
'声明关于声卡控制的常量
Private Const MM_MIXM_LINE_CHANGE = &H3D0
Private Const MM_MIXM_CONTROL_CHANGE = &H3D1

Private Const MIXER_GETCONTROLDETAILSF_LISTTEXT = &H1&
Private Const MIXER_GETCONTROLDETAILSF_QUERYMASK = &HF&
Private Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&

Private Const MIXER_GETLINECONTROLSF_ALL = &H0&
Private Const MIXER_GETLINECONTROLSF_ONEBYID = &H1&
Private Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Private Const MIXER_GETLINECONTROLSF_QUERYMASK = &HF&

Private Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Private Const MIXER_GETLINEINFOF_DESTINATION = &H0&
Private Const MIXER_GETLINEINFOF_LINEID = &H2&
Private Const MIXER_GETLINEINFOF_QUERYMASK = &HF&
Private Const MIXER_GETLINEINFOF_SOURCE = &H1&
Private Const MIXER_GETLINEINFOF_TARGETTYPE = &H4&

Private Const MIXER_OBJECTF_AUX = &H50000000
Private Const MIXER_OBJECTF_HANDLE = &H80000000
Private Const MIXER_OBJECTF_HMIDIIN = &HC0000000
Private Const MIXER_OBJECTF_HMIDIOUT = &HB0000000
Private Const MIXER_OBJECTF_HMIXER = &H80000000
Private Const MIXER_OBJECTF_HWAVEIN = &HA0000000
Private Const MIXER_OBJECTF_HWAVEOUT = &H90000000
Private Const MIXER_OBJECTF_MIDIIN = &H40000000
Private Const MIXER_OBJECTF_MIDIOUT = &H30000000
Private Const MIXER_OBJECTF_MIXER = &H0&
Private Const MIXER_OBJECTF_WAVEIN = &H20000000
Private Const MIXER_OBJECTF_WAVEOUT = &H10000000

Private Const MIXER_SETCONTROLDETAILSF_CUSTOM = &H1&
Private Const MIXER_SETCONTROLDETAILSF_QUERYMASK = &HF&
Private Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&

Private Const MIXERCONTROL_CONTROLF_DISABLED = &H80000000
Private Const MIXERCONTROL_CONTROLF_MULTIPLE = &H2&
Private Const MIXERCONTROL_CONTROLF_UNIFORM = &H1&

Private Const MIXERCONTROL_CT_CLASS_CUSTOM = &H0&
Private Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Private Const MIXERCONTROL_CT_CLASS_LIST = &H70000000
Private Const MIXERCONTROL_CT_CLASS_MASK = &HF0000000
Private Const MIXERCONTROL_CT_CLASS_METER = &H10000000
Private Const MIXERCONTROL_CT_CLASS_NUMBER = &H30000000
Private Const MIXERCONTROL_CT_CLASS_SLIDER = &H40000000
Private Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Private Const MIXERCONTROL_CT_CLASS_TIME = &H60000000

Private Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Private Const MIXERCONTROL_CT_UNITS_CUSTOM = &H0&
Private Const MIXERCONTROL_CT_UNITS_DECIBELS = &H40000
Private Const MIXERCONTROL_CT_UNITS_MASK = &HFF0000
Private Const MIXERCONTROL_CT_UNITS_PERCENT = &H50000
Private Const MIXERCONTROL_CT_UNITS_SIGNED = &H20000
Private Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000

Private Const MIXERCONTROL_CT_SC_LIST_MULTIPLE = &H1000000
Private Const MIXERCONTROL_CT_SC_LIST_SINGLE = &H0&
Private Const MIXERCONTROL_CT_SC_METER_POLLED = &H0&
Private Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Private Const MIXERCONTROL_CT_SC_SWITCH_BUTTON = &H1000000
Private Const MIXERCONTROL_CT_SC_TIME_MICROSECS = &H0&
Private Const MIXERCONTROL_CT_SC_TIME_MILLISECS = &H1000000
Private Const MIXERCONTROL_CT_SUBCLASS_MASK = &HF000000

Private Const MIXERCONTROL_CONTROLTYPE_BASS = &H50030002
Private Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = &H20010000
Private Const MIXERCONTROL_CONTROLTYPE_BOOLEANMETER = &H10010000
Private Const MIXERCONTROL_CONTROLTYPE_BUTTON = &H21010000
Private Const MIXERCONTROL_CONTROLTYPE_CUSTOM = &H0&
Private Const MIXERCONTROL_CONTROLTYPE_DECIBELS = &H30040000
Private Const MIXERCONTROL_CONTROLTYPE_EQUALIZER = &H50030004
Private Const MIXERCONTROL_CONTROLTYPE_FADER = &H50030000
Private Const MIXERCONTROL_CONTROLTYPE_LOUDNESS = &H20010004
Private Const MIXERCONTROL_CONTROLTYPE_MICROTIME = &H60030000
Private Const MIXERCONTROL_CONTROLTYPE_MILLITIME = &H61030000
Private Const MIXERCONTROL_CONTROLTYPE_MIXER = &H71010001
Private Const MIXERCONTROL_CONTROLTYPE_MONO = &H20010003
Private Const MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT = &H71010000
Private Const MIXERCONTROL_CONTROLTYPE_MUTE = &H20010002
Private Const MIXERCONTROL_CONTROLTYPE_MUX = &H70010001
Private Const MIXERCONTROL_CONTROLTYPE_ONOFF = &H20010001
Private Const MIXERCONTROL_CONTROLTYPE_PAN = &H40020001
Private Const MIXERCONTROL_CONTROLTYPE_PEAKMETER = &H10020001
Private Const MIXERCONTROL_CONTROLTYPE_PERCENT = &H30050000
Private Const MIXERCONTROL_CONTROLTYPE_QSOUNDPAN = &H40020002
Private Const MIXERCONTROL_CONTROLTYPE_SIGNED = &H30020000
Private Const MIXERCONTROL_CONTROLTYPE_SIGNEDMETER = &H10020000
Private Const MIXERCONTROL_CONTROLTYPE_SINGLESELECT = &H70010000
Private Const MIXERCONTROL_CONTROLTYPE_SLIDER = &H40020000
Private Const MIXERCONTROL_CONTROLTYPE_STEREOENH = &H20010005
Private Const MIXERCONTROL_CONTROLTYPE_TREBLE = &H50030003
Private Const MIXERCONTROL_CONTROLTYPE_UNSIGNED = &H30030000
Private Const MIXERCONTROL_CONTROLTYPE_UNSIGNEDMETER = &H10030000
Private Const MIXERCONTROL_CONTROLTYPE_VOLUME = &H50030001

Private Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Private Const MIXERLINE_COMPONENTTYPE_DST_DIGITAL = &H1&
Private Const MIXERLINE_COMPONENTTYPE_DST_HEADPHONES = &H5&
Private Const MIXERLINE_COMPONENTTYPE_DST_LAST = &H8&
Private Const MIXERLINE_COMPONENTTYPE_DST_LINE = &H2&
Private Const MIXERLINE_COMPONENTTYPE_DST_MONITOR = &H3&
Private Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = &H4&
Private Const MIXERLINE_COMPONENTTYPE_DST_TELEPHONE = &H6&
Private Const MIXERLINE_COMPONENTTYPE_DST_UNDEFINED = &H0&
Private Const MIXERLINE_COMPONENTTYPE_DST_VOICEIN = &H8&
Private Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN = &H7&

Private Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Private Const MIXERLINE_COMPONENTTYPE_SRC_ANALOG = &H100A&
Private Const MIXERLINE_COMPONENTTYPE_SRC_AUXILIARY = &H1009&
Private Const MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC = &H1005&
Private Const MIXERLINE_COMPONENTTYPE_SRC_DIGITAL = &H1001&
Private Const MIXERLINE_COMPONENTTYPE_SRC_LAST = &H100A&
Private Const MIXERLINE_COMPONENTTYPE_SRC_LINE = &H1002&
Private Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = &H1003&
Private Const MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER = &H1007&
Private Const MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER = &H1004&
Private Const MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE = &H1006&
Private Const MIXERLINE_COMPONENTTYPE_SRC_UNDEFINED = &H1000&
Private Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = &H1008&

Private Const MIXERLINE_LINEF_ACTIVE = &H1&
Private Const MIXERLINE_LINEF_DISCONNECTED = &H8000&
Private Const MIXERLINE_LINEF_SOURCE = &H80000000

Private Const MIXERLINE_TARGETTYPE_AUX = 5
Private Const MIXERLINE_TARGETTYPE_MIDIIN = 4
Private Const MIXERLINE_TARGETTYPE_MIDIOUT = 3
Private Const MIXERLINE_TARGETTYPE_UNDEFINED = 0
Private Const MIXERLINE_TARGETTYPE_WAVEIN = 2
Private Const MIXERLINE_TARGETTYPE_WAVEOUT = 1

Private Const MIXERR_BASE = 1024
Private Const MIXERR_INVALCONTROL = 1025
Private Const MIXERR_INVALLINE = 1024
Private Const MIXERR_INVALVALUE = 1026
Private Const MIXERR_LASTERROR = 1026

'中间变量
Private hMixer As Long         '存放Mixer设备的句柄
Private MaxSources As Long     '存放当前输出设备可调参数个数
Private ProductName As String  '存放当前默认声卡设备名称
Private Destinations As Long   '存放当前输出设备的数量

' Used For Aquiring Details About Any Given Mixer Control.
' Fader, Mute, PeakMeter...
Private MCD As MIXERCONTROLDETAILS
Private ML As MIXERLINE
Private Type RECT
     rLeft As Long
     rTop As Long
     rRight As Long
     rBottom As Long
End Type
' #########################################################################
' This Is A Type I've Created To Slim Down
' The Coding In The Main Form
Private Type MIXERSETTINGS
     MxrChannels As Long    ' Indicates Whether A Line Is Mono Or Stereo.
     MxrLeftVol As Long     ' Left Volume Value (Balance).
     MxrRightVol As Long    ' Right Volume Value (Balance).
     MxrVol As Long         ' Fader Volume.
     MxrVolID As Long       ' Fader Control ID.
     MxrMute As Long        ' Mute Status.
     MxrMuteID As Long      ' Mute Control ID.
     MxrPeakID As Long      ' Peak Meter ID.
End Type
' A Dynamic Array Of The Aformentioned Type.
Private MixerState() As MIXERSETTINGS
' #########################################################################
' Addition API Subs And Function's.
Private Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC&, ByVal x1&, ByVal y1&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal xSrc&, ByVal ySrc&, ByVal dwRop&)
Private Declare Function DrawEdge& Lib "user32" (ByVal ahDc&, lpRect As RECT, ByVal nEdge&, ByVal nFlags&)
Private Declare Function SetRect& Lib "user32" (lpRect As RECT, ByVal x1&, ByVal y1&, ByVal x2&, ByVal y2&)

Private Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, ByVal ptr&, ByVal cb&)
Private Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal ptr&, struct As Any, ByVal cb&)
'内存操作APIs
Private Declare Function GlobalAlloc& Lib "kernel32" (ByVal wFlags&, ByVal dwBytes&)
Private Declare Function GlobalFree& Lib "kernel32" (ByVal hMem&)
Private Declare Function GlobalLock& Lib "kernel32" (ByVal hMem&)
Private Declare Function GlobalUnlock& Lib "kernel32" (ByVal hMem&)
Public Function IsControl() As Boolean
     IsControl = True
     If Not MixerPresent Then  '判断当前是否有声卡设备可以控制
        IsControl = False   '没有安装声卡
        End
     End If
     If Not OpenMixer Then  '打开mixer设备取得控制ID
        IsControl = False   'Mixer设备打开失败
        End
     End If
     If Not GetDeviceCapabilities Then '取得当前默认声音设备名称，并查询设备参数
        IsControl = False '查询失败
        End
     End If
     GetMixerInfo
End Function
Private Function MixerPresent() As Boolean '判断当前是否有声卡设备可以控制
    If mixerGetNumDevs() > 0 Then  '取得当前声卡数量
       MixerPresent = True      '说明有一个声卡可以控制
    Else
       MixerPresent = False
       MsgBox "没有安装声卡！,不可以对声卡进行控制", vbCritical, "出错"
    End If
End Function
Private Function OpenMixer() As Boolean '打开mixer设备取得控制ID
    If mixerOpen(hMixer, 0, 0, 0, 0) = 0 Then '返回控制句柄
       OpenMixer = True   '打开成功
    Else
       OpenMixer = False
       MsgBox "打开mixer失败！,不可以对声卡进行控制", vbCritical, "出错"
    End If
End Function
Private Function GetDeviceCapabilities() As Boolean
    Dim MxrCaps As MIXERCAPS   'Mixer性能结构
   If mixerGetDevCaps(0, MxrCaps, Len(MxrCaps)) = 0 Then  '查询每一个Mixer设备的性能，将返回参数到结构变量
       Destinations = MxrCaps.cDestinations - 1 '可调对象数量
       '可调默认声卡设备名称
       ProductName = Left(MxrCaps.szPname, InStr(MxrCaps.szPname, vbNullChar) - 1)
       GetDeviceCapabilities = True
    Else
       GetDeviceCapabilities = False
       MsgBox "打开mixer性能查询失败！,不可以对声卡进行控制", vbCritical, "出错"
    End If
End Function

Public Sub SetVolume(ByVal BalanceValue As Integer, ByVal VolumeValue As Long, ByVal Number As Integer)
    Dim FaderVol As Long   '音量调整值
    Dim PanPos As Integer  '左右平衡值
    Dim myBalanceValue As Integer '实际的平衡值(-100,100)
    Dim myVolumeValue As Long '实际音量调整值
    Dim hMem As Long
    Dim MCDMono As MIXERCONTROLDETAILS   '单声道控制结构
    Dim MCDStereo As MIXERCONTROLDETAILS '立体声控制结构
    '将百分比转化为-100~100
    If BalanceValue < 50 Then
       If BalanceValue <= 0 Then
          myBalanceValue = -100
       Else
          myBalanceValue = -(100 - (200 * (BalanceValue / 100)))
       End If
    ElseIf BalanceValue > 50 Then
       If BalanceValue >= 100 Then
          myBalanceValue = 100
       Else
          myBalanceValue = 200 * (BalanceValue / 100) - 100
       End If
    Else
       myBalanceValue = 0
    End If
    '将百分比音量转化为实际音量大小
    If VolumeValue <= 0 Then VolumeValue = 1
    If VolumeValue <= 0 Then '最小
       myVolumeValue = 65535
    ElseIf VolumeValue > 100 Then '最大
       myVolumeValue = 0
    Else
       myVolumeValue = 65535 * (VolumeValue / 100)
    End If
    If MixerState(Number).MxrChannels = 2 Then  '如果是控制立体声
       PanPos = myBalanceValue '取得左右平衡值
       FaderVol = myVolumeValue '
       '设置左右声音大小
       If PanPos >= 0 Then '如果平衡值在中间或右边
          MixerState(Number).MxrRightVol = FaderVol
          MixerState(Number).MxrLeftVol = FaderVol - ((PanPos / 100) * FaderVol)
       Else '如果平衡值在左边
          MixerState(Number).MxrLeftVol = FaderVol
          MixerState(Number).MxrRightVol = FaderVol + ((PanPos / 100) * FaderVol)
       End If
       '预值立体声控制结构
       MCDStereo.cbDetails = Len(MixerState(Number).MxrRightVol)  '4字节
       MCDStereo.cbStruct = Len(MCDStereo)    '控制结构控制块大小24字节
       MCDStereo.dwControlID = MixerState(Number).MxrVolID '控制ID.
       MCDStereo.item = 0
       MCDStereo.cChannels = 2
       '分配内存
       hMem = GlobalAlloc(&H40, 8)
       '锁定内存块
       MCDStereo.paDetails = GlobalLock(hMem)
        '从结构当前结构拷贝到立体声控制结构
       CopyPtrFromStruct MCDStereo.paDetails, MixerState(Number).MxrRightVol, 8
       CopyPtrFromStruct MCDStereo.paDetails, MixerState(Number).MxrLeftVol, 8
       '发送控制命令
       mixerSetControlDetails hMixer, MCDStereo, MIXER_SETCONTROLDETAILSF_VALUE
       '解锁内存块
       GlobalUnlock hMem
       '释放内存
       GlobalFree hMem
    Else  '如果是单声道
       '设置单声道声音大小
       MixerState(Number).MxrVol = myVolumeValue
       '预值单声道控制结构
       MCDMono.cbDetails = Len(MixerState(Number).MxrVol)
       MCDMono.cbStruct = Len(MCDMono)
       MCDMono.dwControlID = MixerState(Number).MxrVolID
       MCDMono.item = 0
       MCDMono.cChannels = 1
       '分配内存
       hMem = GlobalAlloc(&H40, 4)
       '锁定内存块
       MCDMono.paDetails = GlobalLock(hMem)
       '从结构当前结构拷贝到单声道控制结构
       CopyPtrFromStruct MCDMono.paDetails, MixerState(Number).MxrVol, 4
       '发送控制命令
       mixerSetControlDetails hMixer, MCDMono, MIXER_SETCONTROLDETAILSF_VALUE
       '解锁内存块
       GlobalUnlock hMem
       '释放内存
       GlobalFree hMem
    End If
End Sub

Private Sub GetMixerInfo() '获取Mixer原有设置
    '将Mixer原有设置全部保存在MixerState()动态数组中
    Dim Dst As Long, Src As Long   ' Destination And Source Counter's.
    Dim ControlID As Long    '设备的控制ID
    For Dst = 0 To Destinations '打描声卡数
        '预置MIXERLINE结构
        ML.cbStruct = Len(ML)
        ML.dwDestination = Dst
        '返回第一块声卡参数到ML中
        mixerGetLineInfo hMixer, ML, MIXER_GETLINEINFOF_DESTINATION
        '如果此设备的类型是控制声卡，那么读取原有设置
        If ML.dwComponentType = MIXERLINE_COMPONENTTYPE_DST_SPEAKERS Then 'aaaaaaaa
           '如果当前设备可调参数超过10,限制到10以内
           If ML.cConnections > 10 Then
              ML.cConnections = 10
              MaxSources = 10
           Else
              MaxSources = ML.cConnections  '少于10个时返回实际可调参数个数
           End If
           '重定义MixerState()，用于保留原有设置
           ReDim MixerState(MaxSources)
           '保存主音量控制设备的声道数
           MixerState(0).MxrChannels = ML.cChannels
           'FrmMxr.LblName(0).Caption = Left(ML.szName, InStr(ML.szName, vbNullChar) - 1)
          '获取主音量控制设备音量控制ID
           ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_VOLUME)
           If ControlID <> 0 Then '如果返回设备号有效，那么预置MCD结构
               MCD.cbDetails = 4  ' Size Of A Long In Byte's.
               MCD.cbStruct = 24
               MCD.cChannels = ML.cChannels
               MCD.dwControlID = ControlID
               MCD.item = 0
               MCD.paDetails = VarPtr(MixerState(0).MxrVol)
              '取得主音量控制设备音量设置信息
               mixerGetControlDetails hMixer, MCD, MIXER_GETCONTROLDETAILSF_VALUE
               MixerState(0).MxrVol = 65535 - MixerState(0).MxrVol
               MixerState(0).MxrVolID = MCD.dwControlID
           End If
          '获取主音量控制设备静音控制ID
           ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_MUTE)
           If ControlID <> 0 Then '如果返回设备号有效，那么预置MCD结构
              MCD.cbDetails = 4  ' Size Of A Long In Byte's.
              MCD.cbStruct = Len(MCD)
              MCD.cChannels = ML.cChannels
              MCD.dwControlID = ControlID
              MCD.item = 0
              MCD.paDetails = VarPtr(MixerState(0).MxrMute)
              '取得主音量控制设备静音设置信息
              mixerGetControlDetails hMixer, MCD, MIXER_GETCONTROLDETAILSF_VALUE
              MixerState(0).MxrMuteID = MCD.dwControlID
           End If
           '获取主音量控制设备Peak控制ID
           ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_PEAKMETER)
           If ControlID <> 0 Then
              MixerState(0).MxrPeakID = ControlID
           End If
           '找出与主控音量相连系的其它可调参数信息
           For Src = 0 To ML.cConnections - 1
               '预置MIXERLINE结构
               ML.cbStruct = Len(ML)
               ML.dwDestination = Dst
               ML.dwSource = Src
               mixerGetLineInfo hMixer, ML, MIXER_GETLINEINFOF_SOURCE
               '保存控制设备的声道数
               MixerState(Src + 1).MxrChannels = ML.cChannels
               '获取控制设备音量控制ID
               ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_VOLUME)
               If ControlID <> 0 Then
                  '如果返回设备号有效，那么预置MCD结构
                   MCD.cbDetails = 4   ' Size Of A Long In Byte's.
                   MCD.cbStruct = Len(MCD)
                   MCD.cChannels = ML.cChannels
                   MCD.dwControlID = ControlID
                   MCD.item = 0
                   MCD.paDetails = VarPtr(MixerState(Src + 1).MxrVol)
                  '取得控制设备静音设置信息
                  mixerGetControlDetails hMixer, MCD, MIXER_GETCONTROLDETAILSF_VALUE
                  MixerState(Src + 1).MxrVol = 65535 - MixerState(Src + 1).MxrVol
                  MixerState(Src + 1).MxrVolID = MCD.dwControlID
               End If
               '获取控制设备静音控制ID
               ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_MUTE)
               If ControlID <> 0 Then '如果返回设备号有效，那么预置MCD结构
                  MCD.cbDetails = 4   ' Size Of A Long In Byte's.
                  MCD.cbStruct = Len(MCD)
                  MCD.cChannels = ML.cChannels
                  MCD.dwControlID = ControlID
                  MCD.item = 0
                  MCD.paDetails = VarPtr(MixerState(Src + 1).MxrMute)
               '取得控制设备静音设置信息
                  mixerGetControlDetails hMixer, MCD, MIXER_GETCONTROLDETAILSF_VALUE
                  MixerState(Src + 1).MxrMuteID = MCD.dwControlID
               End If
               '获取主音量控制设备Peak控制ID
               ControlID = GetControlID(ML.dwComponentType, MIXERCONTROL_CONTROLTYPE_PEAKMETER)
               If ControlID <> 0 Then
                  MixerState(Src + 1).MxrPeakID = ControlID
               End If
           Next
        End If 'aaaaaaa
    Next
End Sub
Private Function GetControlID(ByVal ComponentType As Long, ByVal ControlType As Long) As Long  '取得查询设备的控制ID
   Dim hMem As Long '分配内存
   Dim MC As MIXERCONTROL
   Dim MxrLine As MIXERLINE
   Dim MLC As MIXERLINECONTROLS

   ' Prep The MxrLine Structure.
   MxrLine.cbStruct = Len(MxrLine)
   MxrLine.dwComponentType = ComponentType  ' This Value Sent In.

   ' Get The Line Info.
   If mixerGetLineInfo(hMixer, MxrLine, MIXER_GETLINEINFOF_COMPONENTTYPE) = 0 Then
      ' Prep The MLC Structure.
      MLC.cbStruct = Len(MLC)
      MLC.dwLineID = ML.dwLineID
      MLC.dwControl = ControlType     ' This Value Sent In.
      MLC.cControls = 1
      MLC.cbmxctrl = Len(MC)

      hMem = GlobalAlloc(&H40, Len(MC))
      MLC.pamxctrl = GlobalLock(hMem)

      MC.cbStruct = Len(MC)

      ' Get The Line Control.
      If mixerGetLineControls(hMixer, MLC, MIXER_GETLINECONTROLSF_ONEBYTYPE) = 0 Then
         ' Copy The Data To The MC Structure.
         CopyStructFromPtr MC, MLC.pamxctrl, Len(MC)
         ' Return The Control ID.
         GetControlID = MC.dwControlID
      End If

      GlobalUnlock hMem
      GlobalFree hMem
   End If
End Function
Public Sub Mute(ByVal Number As Integer, ByVal MuteValue As Boolean)
    Dim hMem As Long '分配内存句柄
    '给出静音参数 1-静音，0-解除
    MixerState(Number).MxrMute = MuteValue
    '预置控制块结构
    MCD.cbStruct = Len(MCD)                      ' Structure Size.
    MCD.dwControlID = MixerState(Number).MxrMuteID  ' Control ID.
    MCD.cbDetails = 4                            ' Size Of A Long Variable, In Byte's.
    MCD.cChannels = 1                            ' Mute Has Only One Channel.
    MCD.item = 0
    '分配内存
    hMem = GlobalAlloc(&H40, 4)
    MCD.paDetails = GlobalLock(hMem)
    CopyPtrFromStruct MCD.paDetails, MixerState(Number).MxrMute, 4
    mixerSetControlDetails hMixer, MCD, MIXER_SETCONTROLDETAILSF_VALUE
    GlobalUnlock hMem
    GlobalFree hMem
End Sub
