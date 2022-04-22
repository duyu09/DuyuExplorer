Attribute VB_Name = "Module3"
'任务栏 进度条
Private Declare Function CallWindowProcW& Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long)
Private Declare Function LocalAlloc& Lib "kernel32" (ByVal f&, ByVal s&)
Private Declare Function LocalFree& Lib "kernel32" (ByVal m&)
Private Declare Sub PutMem1 Lib "msvbvm60" (ByVal Ptr As Long, ByVal NewVal As Byte)
Private Declare Sub PutMem2 Lib "msvbvm60" (ByVal Ptr As Long, ByVal NewVal As Integer)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Ptr As Long, ByVal NewVal As Long)
Private Declare Sub PutMem8 Lib "msvbvm60" (ByVal Ptr As Long, ByVal NewVal As Currency)
Public Function CallCOMInterface&(ByVal CComPtr&, ByVal dwMemberIndex&, ParamArray pParam())
Dim i%, offset&
Dim hMem&
hMem = LocalAlloc(0, ((UBound(pParam) + 2) * 5) + 5 + 6 + 1) '//申请代码内存
offset = hMem
For i = UBound(pParam) To 0 Step -1 '//压入参数
PutMem1 offset, &H68 'push Param
offset = offset + 1
PutMem4 offset, pParam(i)
offset = offset + 4
Next
PutMem1 offset, &H68 'push COM point，压入COM指针
PutMem4 offset + 1, CComPtr
offset = offset + 5
PutMem1 offset, &HA1 'mov eax,dword ptr ds:CComPtr，eax=CComPtr指针第一个函数地址
PutMem4 offset + 1, CComPtr
offset = offset + 5
PutMem1 offset, &HFF 'call dword ptr ds:eax + dwMemberIndex * 4，根据Win32下COM表结构，一个函数地址长度4字节
PutMem1 offset + 1, &H90
PutMem4 offset + 2, dwMemberIndex * 4
offset = offset + 6
PutMem1 offset, &HC3 'retn
PutMem1 offset + 1, &H90 '//nop一行代码
CallCOMInterface = CallWindowProcW(hMem, 0, 0, 0, 0) 'call
LocalFree hMem '//释放内存
End Function

