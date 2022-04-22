Attribute VB_Name = "Module2"
Const MAX_PATH = 260
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_DIRECTORY = &H10
'Download by http://www.NewXing.com
Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Type FOLDER_INFO
   curSize As Currency
   lngNumFiles As Long
   lngNumSubFolders As Long
End Type

Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Sub FolderVal(FolderQueue As Collection, lFileNum, lngSize As Currency)
    Dim strTemp As String, strFolder As String
    Dim lRetVal As Long, Fidata As WIN32_FIND_DATA
    Dim lSearchHandle As Long
    strFolder = FolderQueue.item(1)
    '�����ļ�/�ļ��С�
    lSearchHandle = FindFirstFile(strFolder & "*.*", Fidata)
    '�������ļ����ص���Ϣ��
    If lSearchHandle = INVALID_HANDLE_VALUE Then Exit Sub
    '�õ��ļ���Ϣ
    strTemp = TrimNulls(Fidata.cFileName)
    Do While strTemp <> ""
        If (Fidata.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            If strTemp <> "." And strTemp <> ".." Then
                If Right$(strTemp, 1) <> "\" Then strTemp = strTemp & "\"
                FolderQueue.Add strFolder & strTemp
            End If
        Else
            '�����ļ����ֽ�����
            lngSize = lngSize + Fidata.nFileSizeLow
            '�����ļ�����
            lFileNum = lFileNum + 1
        End If
        '������һ���ļ�/�ļ��С�
        lRetVal = FindNextFile(lSearchHandle, Fidata)
        strTemp = ""
        '��ȡ�ļ�����
        If lRetVal <> 0 Then strTemp = TrimNulls(Fidata.cFileName)
    Loop
    'ֹͣ���ҡ�
    lRetVal = FindClose(lSearchHandle)
End Sub

Private Function TrimNulls(strString As String) As String
    Dim l As Long
    l = InStr(1, strString, Chr(0))
    If l = 1 Then
        TrimNulls = ""
    ElseIf l > 0 Then
        TrimNulls = Left$(strString, l - 1)
    Else
        TrimNulls = strString
    End If
End Function
'Download by http://www.NewXing.com
Public Function GetFolderInfo(strFolder As String) As FOLDER_INFO
    Dim lFileNum As Long, lFolderNum As Long
    Dim curSize As Currency, FolderQueue As New Collection
    If Right$(strFolder, 1) <> "\" Then strFolder = strFolder & "\"
    FolderQueue.Add strFolder
    Call FolderVal(FolderQueue, lFileNum, curSize)
    FolderQueue.Remove 1
    Do While FolderQueue.Count > 0
        lFolderNum = lFolderNum + 1
        Call FolderVal(FolderQueue, lFileNum, curSize)
        FolderQueue.Remove 1
        DoEvents
    Loop
    '������Ϣ��
    GetFolderInfo.curSize = curSize
    GetFolderInfo.lngNumFiles = lFileNum
    GetFolderInfo.lngNumSubFolders = lFolderNum
End Function

