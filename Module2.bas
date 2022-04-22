Attribute VB_Name = "Module1"
'以下添加到模块中．调用很简单，直接看好了．
Private Type LVCOLUMN
    mask As Long
    fmt As Long
    cx As Long
    pszText As String
    cchTextMax As Long
    iSubItem As Long
    iImage As Long
    iOrder As Long
    End Type


Declare Function SendMessageColumn Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long
    Const LVM_FIRST = &H1000
    'ListView Column Header constants
    Const LVCF_FMT = &H1
    Const LVCF_WIDTH = &H2
    Const LVCF_TEXT = &H4
    Const LVCF_SUBITEM = &H8
    Const LVCF_IMAGE = &H10
    Const LVCF_ORDER = &H20
    '
    Const LVCFMT_LEFT = &H0
    Const LVCFMT_RIGHT = &H1
    Const LVCFMT_CENTER = &H2
    Const LVCFMT_JUSTIFYMASK = &H3
    Const LVCFMT_IMAGE = &H800
    Const LVCFMT_BITMAP_ON_RIGHT = &H1000
    Const LVCFMT_COL_HAS_IMAGES = &H8000

Public Sub ColumnHeaderSetIcon(LView As ListView, Column As ColumnHeader, Img As ListImage)
    Dim col As LVCOLUMN
    Dim ret As Long
    col.mask = LVCF_FMT Or LVCF_IMAGE
    col.fmt = LVCFMT_LEFT Or LVCFMT_IMAGE Or LVCFMT_COL_HAS_IMAGES
    col.iImage = Img.Index - 1
    ret = SendMessageColumn(LView.hWnd, LVM_FIRST + 26, Column.Index - 1, col)
End Sub
