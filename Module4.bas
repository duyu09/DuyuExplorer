Attribute VB_Name = "Module3"
'������ ������
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
hMem = LocalAlloc(0, ((UBound(pParam) + 2) * 5) + 5 + 6 + 1) '//��������ڴ�
offset = hMem
For i = UBound(pParam) To 0 Step -1 '//ѹ�����
PutMem1 offset, &H68 'push Param
offset = offset + 1
PutMem4 offset, pParam(i)
offset = offset + 4
Next
PutMem1 offset, &H68 'push COM point��ѹ��COMָ��
PutMem4 offset + 1, CComPtr
offset = offset + 5
PutMem1 offset, &HA1 'mov eax,dword ptr ds:CComPtr��eax=CComPtrָ���һ��������ַ
PutMem4 offset + 1, CComPtr
offset = offset + 5
PutMem1 offset, &HFF 'call dword ptr ds:eax + dwMemberIndex * 4������Win32��COM��ṹ��һ��������ַ����4�ֽ�
PutMem1 offset + 1, &H90
PutMem4 offset + 2, dwMemberIndex * 4
offset = offset + 6
PutMem1 offset, &HC3 'retn
PutMem1 offset + 1, &H90 '//nopһ�д���
CallCOMInterface = CallWindowProcW(hMem, 0, 0, 0, 0) 'call
LocalFree hMem '//�ͷ��ڴ�
End Function

