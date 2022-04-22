Attribute VB_Name = "Module1"
Option Explicit
Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_LARGEICON = &H0
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_SMALLICON = &H1
Public Const SHGFI_SYSICONINDEX = &H4000
Public Const SHGFI_TYPENAME = &H400
Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
Public Const MAX_PATH = 260
Public Const ILD_TRANSPARENT = &H1
Public Type SHFILEINFO
   hIcon As Long
   iIcon As Long
   dwAttributes As Long
   szDisplayName As String * MAX_PATH
   szTypeName As String * 80
End Type
Public Declare Function SHGetFileInfo Lib _
   "shell32.dll" Alias "SHGetFileInfoA" _
   (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long
Public Declare Function ImageList_Draw Lib "comctl32.dll" _
   (ByVal himl As Long, ByVal i As Long, _
    ByVal hDCDest As Long, ByVal x As Long, _
    ByVal y As Long, ByVal flags As Long) As Long
Public shinfo As SHFILEINFO
Public Const SHGFI_USEFILEATTRIBUTES = &H10
Public Const SHGFI_ICON = &H100
'===================================================

Public Fn As String
