Attribute VB_Name = "Modtree"

Private OldProc As Long
Private hTree As Long
Private iNodes As Long

Private Const MAX_LEN = 32
Private Const ID_TREEVIEW = 1000

Private Const ICC_TREEVIEW_CLASSES = &H2

Private Type TvwNode
    hItem As Long
    hParent As Long
    Index As Long
    Key As String
    Text As String
    Image As Long
    Tag As String
End Type

Public NodeX() As TvwNode

Public Enum HitTestInfoConstants
    htAbove = &H100
    htBelow = &H200
    htBelowLast = &H1
    htItemPlusMinus = &H10
    htItemIcon = &H2
    htItemIndent = &H8
    htItemText = &H4
    htItemRight = &H20
    htItemState = &H40
    HTLEFT = &H800
    HTRIGHT = &H400
End Enum

Public Enum RelationConstants
    tvwSort
    tvwFirst
    tvwLast
    tvwChild
End Enum

Private Const GWL_EXSTYLE = (-20)
Private Const GWL_STYLE = (-16)
Private Const GWL_WNDPROC = (-4)

' TreeView messages.
Private Const TV_FIRST = &H1100
Private Const TVM_CREATEDRAGIMAGE = (TV_FIRST + 18)
Private Const TVM_DELETEITEM = (TV_FIRST + 1)
Private Const TVM_EDITLABEL = (TV_FIRST + 14)
Private Const TVM_ENDEDITLABELNOW = (TV_FIRST + 22)
Private Const TVM_ENSUREVISIBLE = (TV_FIRST + 20)
Private Const TVM_EXPAND = (TV_FIRST + 2)
Private Const TVM_GETBKCOLOR = (TV_FIRST + 31)
Private Const TVM_GETBORDER = (TV_FIRST + 36)
Private Const TVM_GETCOUNT = (TV_FIRST + 5)
Private Const TVM_GETEDITCONTROL = (TV_FIRST + 15)
Private Const TVM_GETIMAGELIST = (TV_FIRST + 8)
Private Const TVM_GETINDENT = (TV_FIRST + 6)
Private Const TVM_GETISEARCHSTRINGA = (TV_FIRST + 23)
Private Const TVM_GETITEM = (TV_FIRST + 12)
Private Const TVM_GETITEMHEIGHT = (TV_FIRST + 28)
Private Const TVM_GETITEMRECT = (TV_FIRST + 4)
Private Const TVM_GETNEXTITEM = (TV_FIRST + 10)
Private Const TVM_GETSCROLLTIME = (TV_FIRST + 34)
Private Const TVM_GETTEXTCOLOR = (TV_FIRST + 32)
Private Const TVM_GETTOOLTIPS = (TV_FIRST + 25)
Private Const TVM_GETVISIBLECOUNT = (TV_FIRST + 16)
Private Const TVM_HITTEST = (TV_FIRST + 17)
Private Const TVM_INSERTITEM = (TV_FIRST + 0)
Private Const TVM_SELECTITEM = (TV_FIRST + 11)
Private Const TVM_SETBKCOLOR = (TV_FIRST + 29)
Private Const TVM_SETBORDER = (TV_FIRST + 35)
Private Const TVM_SETIMAGELIST = (TV_FIRST + 9)
Private Const TVM_SETINDENT = (TV_FIRST + 7)
Private Const TVM_SETINSERTMARK = (TV_FIRST + 26)
Private Const TVM_SETITEM = (TV_FIRST + 13)
Private Const TVM_SETITEMHEIGHT = (TV_FIRST + 27)
Private Const TVM_SETSCROLLTIME = (TV_FIRST + 33)
Private Const TVM_SETTEXTCOLOR = (TV_FIRST + 30)
Private Const TVM_SETTOOLTIPS = (TV_FIRST + 24)
Private Const TVM_SORTCHILDREN = (TV_FIRST + 19)
Private Const TVM_SORTCHILDRENCB = (TV_FIRST + 21)
Private Const TVM_SETLINECOLOR = (TV_FIRST + 40)
Private Const TVM_GETLINECOLOR = (TV_FIRST + 41)

' Treeview Notifications
Private Const TVN_FIRST = -400
Private Const TVN_BEGINLABELEDIT = (TVN_FIRST - 10)
Private Const TVN_BEGINDRAG = (TVN_FIRST - 7)
Private Const TVN_BEGINRDRAG = (TVN_FIRST - 8)
Private Const TVN_DELETEITEM = (TVN_FIRST - 9)
Private Const TVN_GETDISPINFO = (TVN_FIRST - 3)
Private Const TVN_GETINFOTIP = (TVN_FIRST - 13)
Private Const TVN_KEYDOWN = (TVN_FIRST - 12)
Private Const TVN_ENDLABELEDIT = (TVN_FIRST - 11)
Private Const TVN_ITEMEXPANDED = (TVN_FIRST - 6)
Private Const TVN_ITEMEXPANDING = (TVN_FIRST - 5)
Private Const TVN_SELCHANGED = (TVN_FIRST - 2)
Private Const TVN_SELCHANGING = (TVN_FIRST - 1)
Private Const TVN_SINGLEEXPAND = (TVN_FIRST - 15)

' TreeView specific styles.
Private Const TVS_CHECKBOXES = &H100
Private Const TVS_DISABLEDRAGDROP = &H10
Private Const TVS_EDITLABELS = &H8
Private Const TVS_FULLROWSELECT = &H1000
Private Const TVS_HASBUTTONS = &H1
Private Const TVS_HASLINES = &H2
Private Const TVS_INFOTIP = &H800
Private Const TVS_LINESATROOT = &H4
Private Const TVS_NOSCROLL = &H2000
Private Const TVS_NOTOOLTIPS = &H80
Private Const TVS_SHOWSELALWAYS = &H20
Private Const TVS_SINGLEEXPAND = &H400
Private Const TVS_TRACKSELECT = &H200

' Notification messages.
Private Const NM_FIRST = 0
Private Const NM_CLICK = (NM_FIRST - 2)
Private Const NM_CUSTOMDRAW = (NM_FIRST - 12)
Private Const NM_DBLCLK = (NM_FIRST - 3)
Private Const NM_KILLFOCUS = (NM_FIRST - 8)
Private Const NM_RCLICK = (NM_FIRST - 5)
Private Const NM_RETURN = (NM_FIRST - 4)

' Inserting stuff.
Private Const TVI_ROOT = &HFFFF0000
Private Const TVI_FIRST = &HFFFF0001
Private Const TVI_LAST = &HFFFF0002
Private Const TVI_SORT = &HFFFF0003

' Mask values.
Private Const TVIF_CHILDREN = &H40
Private Const TVIF_DI_SETITEM = &H1000
Private Const TVIF_HANDLE = &H10
Private Const TVIF_IMAGE = &H2
Private Const TVIF_INTEGRAL = &H80
Private Const TVIF_PARAM = &H4
Private Const TVIF_SELECTEDIMAGE = &H20
Private Const TVIF_STATE = &H8
Private Const TVIF_TEXT = &H1

' More mask values, of the state kind.
Private Const TVIS_BOLD = &H10
Private Const TVIS_CUT = &H4
Private Const TVIS_DROPHILITED = &H8
Private Const TVIS_EXPANDED = &H20
Private Const TVIS_EXPANDEDONCE = &H40
Private Const TVIS_EXPANDPARTIAL = &H80
Private Const TVIS_OVERLAYMASK = &HF00
Private Const TVIS_SELECTED = &H2
Private Const TVIS_STATEIMAGEMASK = &HF000
Private Const TVIS_USERMASK = &HF000

' Expanding stuff.
Private Const TVE_COLLAPSE = &H1
Private Const TVE_COLLAPSERESET = &H8000
Private Const TVE_EXPAND = &H2
Private Const TVE_EXPANDPARTIAL = &H4000
Private Const TVE_TOGGLE = &H3

' ImageList type values.
Private Const TVSIL_NORMAL = 0
Private Const TVSIL_STATE = 2

Private Const WS_BORDER = &H800000
Private Const WS_CHILD = &H40000000
Private Const WS_DISABLED = &H8000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_TABSTOP = &H10000
Private Const WS_EX_CLIENTEDGE = &H200

Private Const WM_SETFOCUS = &H7
Private Const WM_SETREDRAW = &HB
Private Const WM_MOUSEACTIVATE = &H21
Private Const WM_NOTIFY = &H4E
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONUP = &H202
Private Const WM_TIMER = &H113
Private Const WM_VSCROLL = &H115
Private Const SB_LINEDOWN = 1
Private Const SB_LINEUP = 0

' ImageList Declarations
Private Const SM_CXSMICON = 49
Private Const SM_CYSMICON = 50

Private Const ILC_MASK = &H1
Private Const ILC_COLOR = &H0
Private Const ILC_COLORDDB = &HFE
Private Const ILC_COLOR4 = &H4
Private Const ILC_COLOR8 = &H8
Private Const ILC_COLOR16 = &H10
Private Const ILC_COLOR24 = &H18
Private Const ILC_COLOR32 = &H20

Private Const ILD_BLEND25 = &H2
Private Const ILD_BLEND50 = &H4
Private Const ILD_MASK = &H10
Private Const ILD_NORMAL = &H0
Private Const ILD_FOCUS = ILD_BLEND25
Private Const ILD_SELECTED = ILD_BLEND50
Private Const ILD_TRANSPARENT = &H1

Private Const IMAGE_BITMAP = 0
Private Const LR_DEFAULTCOLOR = &H0
Private Const LR_CREATEDIBSECTION = &H2000
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_VGACOLOR = &H80
Private Const LR_LOADFROMFILE = &H10

Private Const CLR_NONE = &HFFFFFFFF
Private Const CLR_DEFAULT = &HFF000000

Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const LANG_NEUTRAL = &H0
Private Const SUBLANG_DEFAULT = &H1

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type DWORD
    LOWORD As Integer
    HIWORD As Integer
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type TVHITTESTINFO
    pt As POINTAPI
    flags As Long
    hItem As Long
End Type

Private Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Private Type TVITEM
    mask As Long
    hItem As Long
    State As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    cChildren As Long
    lParam As Long
End Type

Private Type TVITEMEX
    mask As Long
    hItem As Long
    State As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    cChildren As Long
    lParam As Long
    iIntegral As Long
End Type

Private Type TVDISPINFO
    hdr As NMHDR
    Item As TVITEM
End Type

Private Type TVINSERTSTRUCT
    hParent As Long
    hInsertAfter As Long
    Item As TVITEMEX
End Type

Private Type NMTREEVIEW
    hdr As NMHDR
    action As Long
    itemOld As TVITEM
    itemNew As TVITEM
    ptDrag As POINTAPI
End Type

Private Type iccex
    dwSize As Long          ' size of this structure
    dwICC As Long           ' flags indicating which classes to be initialized
End Type

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (icc As iccex) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Declare Sub ImageList_EndDrag Lib "comctl32.dll" ()
Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal hIml As Long) As Long
Private Declare Function ImageList_ReplaceIcon Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_LoadImage Lib "comctl32.dll" (ByVal hi As Long, ByVal lpbmp As String, ByVal cx As Long, ByVal cGrow As Long, ByVal crMask As Long, ByVal uType As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_AddMasked Lib "comctl32.dll" (ByVal hIml As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Private Declare Function ImageList_BeginDrag Lib "comctl32.dll" (ByVal himlTrack As Long, ByVal iTrack As Long, ByVal dxHotspot As Long, ByVal dyHotspot As Long) As Long
Private Declare Function ImageList_DragEnter Lib "comctl32.dll" (ByVal hwndLock As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ImageList_DragMove Lib "comctl32.dll" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ImageList_DragShowNolock Lib "comctl32.dll" (ByVal fShow As Long) As Long
Private Declare Function ImageList_DragLeave Lib "comctl32.dll" (ByVal hwndLock As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal hIml As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Function ErrorHandler() As Boolean
    
    Dim Buffer As String
    Buffer = Space(200)
    
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, GetLastError, LANG_NEUTRAL, Buffer, 200, ByVal 0&
    
    Debug.Print Buffer

End Function

Public Function TvwCreateEx(ByVal hParent As Long) As Long
    Dim rcl   As RECT
    Dim hIml  As Long
    
    Call InitCommonControls
    
    hCont = CreateWindowEx(0&, "STATIC", "bTreeViewClass", WS_VISIBLE Or WS_CHILD, 0, 0, 250, FrmEditor.PicTree.Height, hParent, 0, App.hInstance, 0)
    hTree = CreateWindowEx(0&, "SysTreeView32", "", WS_VISIBLE Or WS_CHILD Or TVS_HASLINES Or TVS_HASBUTTONS Or TVS_LINESATROOT, 0, 0, 250, FrmEditor.PicTree.Height, hCont, ID_TREEVIEW, App.hInstance, 0)

    If hTree = 0 Then
        MsgBox "TreeView CreateWindow Failed"
        Exit Function
    End If
    
    ReDim Preserve NodeX(0)
    iNodes = 0
    
    SendMessageLong hTree, TVM_SETBKCOLOR, 0, &HFFFFFF ' property BackColor
    SendMessageLong hTree, TVM_SETTEXTCOLOR, 0, &H0 ' property FontColor
    SendMessage hTree, TVM_SETINDENT, 16, 0 ' property Indent
    SendMessage hTree, TVM_SETITEMHEIGHT, 16, 0

    Dim TVIN As TVINSERTSTRUCT, mRoot As Long, mParent As Long, i As Byte
    
    ' design time items
    If bDesign = True Then
        TVIN.hParent = TVI_ROOT
        TVIN.hInsertAfter = TVI_FIRST
        TVIN.Item.pszText = "Root Item" & Chr(0)
        TVIN.Item.cchTextMax = 10
        TVIN.Item.mask = TVIF_TEXT
        mRoot = SendMessage(hTree, TVM_INSERTITEM, 0, TVIN)
        TVIN.hParent = mRoot
        TVIN.Item.pszText = "Parent Item" & Chr(0)
        TVIN.Item.cchTextMax = 12
        mParent = SendMessage(hTree, TVM_INSERTITEM, 0, TVIN)
        SendMessage hTree, TVM_EXPAND, TVE_EXPAND, ByVal mRoot
        For i = 1 To 2
            TVIN.hParent = mParent
            TVIN.Item.pszText = "Child Item" & Chr(0)
            TVIN.Item.cchTextMax = 11
            SendMessage hTree, TVM_INSERTITEM, 0, TVIN
        Next i
        SendMessage hTree, TVM_EXPAND, TVE_EXPAND, ByVal mParent
    End If

    OldProc = GetWindowLong(hCont, GWL_WNDPROC)
    ret = SetWindowLong(hCont, GWL_WNDPROC, AddressOf TvwProc)
    
    CreateTvwEx = hTree

End Function

Private Function TvwProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim RetVal As Long
    Dim tHDR As NMHDR
    Dim tvInsert As TVINSERTSTRUCT, TVDISPINFO As TVDISPINFO
    Dim pt As POINTAPI, Bool As Boolean, H As Long
    Dim rc As RECT
    Dim lLen As Long, iPos As Long
    Dim sText As String
    Dim TVHT As TVHITTESTINFO

    Select Case iMsg
    
    Case WM_NOTIFY
        CopyMemory tHDR, ByVal lParam, Len(tHDR)
        RetVal = 0
        
        Select Case tHDR.code
                
        Case NM_CLICK
            
            GetCursorPos TVHT.pt
            ScreenToClient hTree, TVHT.pt
            SendMessage hTree, TVM_HITTEST, 0, TVHT
            ' If there's an item there, tell the user.
            If TVHT.hItem <> 0 Then
                For ax = 0 To iNodes - 1
                    If NodeX(ax).hItem = TVHT.hItem Then
                        If LCase(NodeX(ax).Text) = "template" Or LCase(NodeX(ax).Text) = "player" Or LCase(NodeX(ax).Text) = "window" Then Exit Function
                        ' call raise function
                        'RaiseEvent ItemClick(NodeX(ax).Key, NodeX(ax).Index)
                         EscribeXY FrmEditor.scivb.GetCaretInLine + 1, FrmEditor.scivb.CurrentLine + 1, NodeX(ax).Text, FrmEditor.scivb
                    End If
                Next
            End If
            ' Otherwise, do a generic click.
            'RaiseEvent Click(TVHT.PT.x, TVHT.PT.y, False)
        End Select

    End Select
    
    TvwProc = CallWindowProc(OldProc, hWnd, iMsg, wParam, lParam)
    
End Function

Public Function ImgAddMasked(ByVal hIml As Long, ByVal hImage As Long, ByVal cMask As Long) As Long
    lRet = ImageList_AddMasked(hIml, hImage, cMask)
    ImgAddMasked = IIf(lRet <> -1, hIml, -1)
End Function

Public Function ImgAddIcon(ByVal hIml As Long, ByVal hIcon As Long) As Long
    ImgAddIcon = ImageList_ReplaceIcon(hIml, -1, hIcon)
End Function

Public Function ImgCreateList(ByVal lColor As Long, ByVal InitialImages As Long, Optional ByVal TotalImages As Variant) As Long
    
    Dim cxSmIcon As Long, cySmIcon As Long
      
    If IccInit = False Then InitCommonControls
      
    If IsNull(TotalImages) = True Then TotalImages = 0
    If lColor = -1 Then lColor = ILC_COLOR24 Or ILC_MASK
       
    cxSmIcon = GetSystemMetrics(SM_CXSMICON)
    cySmIcon = GetSystemMetrics(SM_CYSMICON)
    m_hIml = ImageList_Create(cxSmIcon, cySmIcon, lColor, InitialImages, TotalImages)
    ImgCreateList = m_hIml
    
End Function

Public Function ImgLoadList(ByVal m_Image As String, ByVal m_Color As Long, ByVal m_Res As Long, ByVal m_Width As Long, ByVal m_Height As Long) As Long
    
    Dim hBitmap As Long, hTemp As Long
    m_Images = m_Width / m_Height
    
    hBitmap = LoadImage(App.hInstance, m_Image, IMAGE_BITMAP, m_Width, m_Height, IIf(m_Res = True, LR_CREATEDIBSECTION, LR_LOADFROMFILE))
    hTemp = ImgCreateList(ILC_COLOR24 Or ILC_MASK, m_Images, m_Images)
    ImgLoadList = ImgAddMasked(hTemp, hBitmap, m_Color)
    
End Function

Public Function TvwImageList(ByVal hImg As Long) As Long
    SendMessageLong hTree, TVM_SETIMAGELIST, TVSIL_NORMAL, hImg
End Function

Public Function TvwAddItem(hRelItem, Relation As Long, Key As String, Text As String, Optional Image As Long = -1, Optional SelectedImage As Long = -1) As Long

    Dim TVIN As TVINSERTSTRUCT, hRel As Long, TVI As TVITEMEX
    
    If Relation = 0 Then Relation = tvwLast
    If hRelItem = 0 Then hRelItem = 0&

    If TypeName(hRelItem) = "Long" Then
        hRel = hRelItem
    ElseIf TypeName(hRelItem) = "String" Then
        For ax = 0 To iNodes - 1
            If NodeX(ax).Key = hRelItem Then
                hRel = NodeX(ax).hItem
                Exit For
            Else
                hRel = 0&
            End If
        Next
    End If
    
    TVIN.hParent = hRel
    
    If Image > 0 Then
    
        TVIN.Item.mask = TVIN.Item.mask Or TVIF_IMAGE
        If SelectedImage < 0 Then
            SelectedImage = Image
            TVIN.Item.mask = TVIN.Item.mask Or TVIF_SELECTEDIMAGE
        End If
        
    End If
    
    If SelectedImage > 0 Then
        TVIN.Item.mask = TVIN.Item.mask Or TVIF_SELECTEDIMAGE
    End If
    
    TVIN.Item.mask = TVIN.Item.mask Or TVIF_STATE Or TVIF_TEXT
    TVIN.Item.pszText = Text & Chr(0)
    TVIN.Item.cchTextMax = Len(Text) + 1
    
    If Image >= 0 Then
        TVIN.Item.iImage = Image
    End If
    
    If SelectedImage >= 0 Then
        TVIN.Item.iSelectedImage = SelectedImage
    End If
    
    'TVIN.Item.stateMask = TVIS_BOLD
    'TVIN.Item.state = TVIS_BOLD
    
    If Relation = tvwSort Then
        TVIN.hInsertAfter = TVI_SORT
    ElseIf Relation = tvwFirst Then
        TVIN.hInsertAfter = TVI_FIRST
    ElseIf Relation = tvwLast Then
        TVIN.hInsertAfter = TVI_LAST
    ElseIf Relation = tvwChild Then
        TVIN.hParent = SendMessageLong(hTree, TVM_GETNEXTITEM, TVGN_PARENT, hRel)
        TVIN.hInsertAfter = hRel
    End If
    
    hRel = SendMessage(hTree, TVM_INSERTITEM, 0, TVIN)
    
    If hRel <> 0 Then
        
        SendMessage hTree, TVM_GETITEM, hRel, TVI
        TVI.mask = TVIF_PARAM
        TVI.lParam = hRel
        SendMessage hTree, TVM_SETITEM, hRel, TVI
        
        ReDim Preserve NodeX(iNodes)
        
        NodeX(iNodes).hItem = hRel
        NodeX(iNodes).hParent = SendMessageLong(hTree, TVM_GETNEXTITEM, TVGN_PARENT, hRel)
        NodeX(iNodes).Index = iNodes
        NodeX(iNodes).Text = Text
        NodeX(iNodes).Key = Key
        NodeX(iNodes).Image = Image
        
        iNodes = iNodes + 1

        'If DoSort(hRel) Then
        '    SendMessageL hTree, TVM_SORTCHILDREN, 0, hRel
        'End If
    End If
    
    TvwAddItem = hRel
    
End Function

