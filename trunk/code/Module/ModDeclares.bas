Attribute VB_Name = "ModDeclares"
Option Explicit
'
' ***************************************************************************************
' * Project  | GpTabStrip                                                               *
' *----------|--------------------------------------------------------------------------*
' * Version  | V1.0                                                                     *
' *----------|--------------------------------------------------------------------------*
' * Author   | Genghis Khan(GuangJian Guo)                                              *
' *----------|--------------------------------------------------------------------------*
' * WebSite  | http://www.itkhan.com                                                    *
' *----------|--------------------------------------------------------------------------*
' * MailTo   | webmaster@itkhan.com                                                     *
' *----------|--------------------------------------------------------------------------*
' * Date     | 13 April 2003                                                            *
' ***************************************************************************************

' ======================================================================================
' Constants
' ======================================================================================

#Const DEBUGMODE = 0

'Set of bit flags that indicate which common control classes will be loaded
'from the DLL. The dwICC value of tagINITCOMMONCONTROLSEX can
'be a combination of the following:
Public Const ICC_LISTVIEW_CLASSES = &H1          '/* listview, header
Public Const ICC_TREEVIEW_CLASSES = &H2          '/* treeview, tooltips
Public Const ICC_BAR_CLASSES = &H4               '/* toolbar, statusbar, trackbar, tooltips
Public Const ICC_TAB_CLASSES = &H8               '/* tab, tooltips
Public Const ICC_UPDOWN_CLASS = &H10             '/* updown
Public Const ICC_PROGRESS_CLASS = &H20           '/* progress
Public Const ICC_HOTKEY_CLASS = &H40             '/* hotkey
Public Const ICC_ANIMATE_CLASS = &H80            '/* animate
Public Const ICC_WIN95_CLASSES = &HFF            '/* loads everything above
Public Const ICC_DATE_CLASSES = &H100            '/* month picker, date picker, time picker, updown
Public Const ICC_USEREX_CLASSES = &H200          '/* ComboEx
Public Const ICC_COOL_CLASSES = &H400            '/* Rebar (coolbar) control


' ָ�����ڵĽṹ��ȡ����Ϣ������GetWindowLong��SetWindowLong����
Public Const GWL_EXSTYLE = (-20)                 '/* ��չ������ʽ */
Public Const GWL_HINSTANCE = (-6)                '/* ӵ�д��ڵ�ʵ���ľ�� */
Public Const GWL_HWNDPARENT = (-8)               '/* �ô���֮���ľ������Ҫ��SetWindowWord���ı����ֵ */
Public Const GWL_ID = (-12)                      '/* �Ի�����һ���Ӵ��ڵı�ʶ�� */
Public Const GWL_STYLE = (-16)                   '/* ������ʽ */
Public Const GWL_USERDATA = (-21)                '/* ������Ӧ�ó���涨 */
Public Const GWL_WNDPROC = (-4)                  '/* �ô��ڵĴ��ں����ĵ�ַ */
Public Const DWL_DLGPROC = 4                     '/* ������ڵĶԻ�������ַ */
Public Const DWL_MSGRESULT = 0                   '/* �ڶԻ������д����һ����Ϣ���ص�ֵ */
Public Const DWL_USER = 8                        '/* ������Ӧ�ó���涨 */


' GetDeviceCaps����������GetDeviceCaps����
Public Const DRIVERVERSION = 0                   '/* ����������汾
Public Const BITSPIXEL = 12                      '/*
Public Const LOGPIXELSX = 88                     '/*  Logical pixels/inch in X
Public Const LOGPIXELSY = 90                     '/*  Logical pixels/inch in Y

' Windows������������GetSysColor
Public Const COLOR_ACTIVEBORDER = 10             '/* ����ڵı߿�
Public Const COLOR_ACTIVECAPTION = 2             '/* ����ڵı���
Public Const COLOR_ADJ_MAX = 100                 '/*
Public Const COLOR_ADJ_MIN = -100                '/*
Public Const COLOR_APPWORKSPACE = 12             '/* MDI����ı���
Public Const COLOR_BACKGROUND = 1                '/*
Public Const COLOR_BTNDKSHADOW = 21              '/*
Public Const COLOR_BTNLIGHT = 22                 '/*
Public Const COLOR_BTNFACE = 15                  '/* ��ť
Public Const COLOR_BTNHIGHLIGHT = 20             '/* ��ť��3D������
Public Const COLOR_BTNSHADOW = 16                '/* ��ť��3D��Ӱ
Public Const COLOR_BTNTEXT = 18                  '/* ��ť����
Public Const COLOR_CAPTIONTEXT = 9               '/* ���ڱ����е�����
Public Const COLOR_GRAYTEXT = 17                 '/* ��ɫ���֣���ʹ���˶���������Ϊ��
Public Const COLOR_HIGHLIGHT = 13                '/* ѡ������Ŀ����
Public Const COLOR_HIGHLIGHTTEXT = 14            '/* ѡ������Ŀ����
Public Const COLOR_INACTIVEBORDER = 11           '/* ������ڵı߿�
Public Const COLOR_INACTIVECAPTION = 3           '/* ������ڵı���
Public Const COLOR_INACTIVECAPTIONTEXT = 19      '/* ������ڵ�����
Public Const COLOR_MENU = 4                      '/* �˵�
Public Const COLOR_MENUTEXT = 7                  '/* �˵�����
Public Const COLOR_SCROLLBAR = 0                 '/* ������
Public Const COLOR_WINDOW = 5                    '/* ���ڱ���
Public Const COLOR_WINDOWFRAME = 6               '/* ����
Public Const COLOR_WINDOWTEXT = 8                '/* ��������
Public Const COLORONCOLOR = 3

' ����CombineRgn�ķ���ֵ������Long
Public Const COMPLEXREGION = 3                   '/* �����л��ཻ���ı߽� */
Public Const SIMPLEREGION = 2                    '/* ����߽�û�л��ཻ�� */
Public Const NULLREGION = 1                      '/* ����Ϊ�� */
Public Const ERRORAPI = 0                        '/* ���ܴ���������� */

' ���������ķ���������CombineRgn�ĵĲ���nCombineMode��ʹ�õĳ���
Public Const RGN_AND = 1                         '/* hDestRgn������Ϊ����Դ����Ľ��� */
Public Const RGN_COPY = 5                        '/* hDestRgn������ΪhSrcRgn1�Ŀ��� */
Public Const RGN_DIFF = 4                        '/* hDestRgn������ΪhSrcRgn1����hSrcRgn2���ཻ�Ĳ��� */
Public Const RGN_OR = 2                          '/* hDestRgn������Ϊ��������Ĳ��� */
Public Const RGN_XOR = 3                         '/* hDestRgn������Ϊ������Դ����OR֮��Ĳ��� */

' Missing Draw State constants declarations���ο�DrawState����
'/* Image type */
Public Const DST_COMPLEX = &H0                   '/* ��ͼ����lpDrawStateProc����ָ���Ļص������ڼ�ִ�С�lParam��wParam�ᴫ�ݸ��ص��¼�
Public Const DST_TEXT = &H1                      '/* lParam�������ֵĵ�ַ����ʹ��һ���ִ���������wParam�����ִ��ĳ���
Public Const DST_PREFIXTEXT = &H2                '/* ��DST_TEXT���ƣ�ֻ�� & �ַ�ָ��Ϊ�¸��ַ������»���
Public Const DST_ICON = &H3                      '/* lParam����ͼ����
Public Const DST_BITMAP = &H4                    '/* lParam�еľ��
' /* State type */
Public Const DSS_NORMAL = &H0                    '/* ��ͨͼ��
Public Const DSS_UNION = &H10                    '/* ͼ����ж�������
Public Const DSS_DISABLED = &H20                 '/* ͼ����и���Ч��
Public Const DSS_MONO = &H80                     '/* ��hBrush���ͼ��
Public Const DSS_RIGHT = &H8000                  '/*

' Built in ImageList drawing methods:
Public Const ILD_NORMAL = 0&
Public Const ILD_TRANSPARENT = 1&
Public Const ILD_BLEND25 = 2&
Public Const ILD_SELECTED = 4&
Public Const ILD_FOCUS = 4&
Public Const ILD_MASK = &H10&
Public Const ILD_IMAGE = &H20&
Public Const ILD_ROP = &H40&
Public Const ILD_OVERLAYMASK = 3840&
Public Const ILC_MASK = &H1&
Public Const ILCF_MOVE = &H0&
Public Const ILCF_SWAP = &H1&

Public Const CLR_DEFAULT = -16777216
Public Const CLR_HILIGHT = -16777216
Public Const CLR_NONE = -1

' General windows messages:
Public Const WM_COMMAND = &H111
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102
Public Const WM_SETFOCUS = &H7
Public Const WM_KILLFOCUS = &H8
Public Const WM_SETFONT = &H30
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_SETTEXT = &HC
Public Const WM_NOTIFY = &H4E&

' Show window styles
Public Const SW_SHOWNORMAL = 1
Public Const SW_ERASE = &H4
Public Const SW_HIDE = 0
Public Const SW_INVALIDATE = &H2
Public Const SW_MAX = 10
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_OTHERUNZOOM = 4
Public Const SW_OTHERZOOM = 2
Public Const SW_PARENTCLOSING = 1
Public Const SW_RESTORE = 9
Public Const SW_PARENTOPENING = 3
Public Const SW_SHOW = 5
Public Const SW_SCROLLCHILDREN = &H1
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_SHOWNOACTIVATE = 4

' �����Ĺ�դ��������
Public Const BLACKNESS = &H42                    '/* ��ʾʹ���������ɫ�������0��ص�ɫ�������Ŀ��������򣬣���ȱʡ�������ɫ����ԣ�����ɫΪ��ɫ����
Public Const DSTINVERT = &H550009                '/* ��ʾʹĿ�����������ɫȡ����
Public Const MERGECOPY = &HC000CA                '/* ��ʾʹ�ò����͵�AND���룩��������Դ�����������ɫ���ض�ģʽ���һ��
Public Const MERGEPAINT = &HBB0226               '/* ͨ��ʹ�ò����͵�OR���򣩲������������Դ�����������ɫ��Ŀ������������ɫ�ϲ���
Public Const NOTSRCCOPY = &H330008               '/* ��Դ����������ɫȡ�����ڿ�����Ŀ���������
Public Const NOTSRCERASE = &H1100A6              '/* ʹ�ò������͵�OR���򣩲��������Դ��Ŀ������������ɫֵ��Ȼ�󽫺ϳɵ���ɫȡ����
Public Const PATCOPY = &HF00021                  '/* ���ض���ģʽ������Ŀ��λͼ�ϡ�
Public Const PATINVERT = &H5A0049                '/* ͨ��ʹ�ò���OR���򣩲�������Դ��������ȡ�������ɫֵ���ض�ģʽ����ɫ�ϲ���Ȼ��ʹ��OR���򣩲��������ò����Ľ����Ŀ����������ڵ���ɫ�ϲ���
Public Const PATPAINT = &HFB0A09                 '/* ͨ��ʹ��XOR����򣩲�������Դ��Ŀ����������ڵ���ɫ�ϲ���
Public Const SRCAND = &H8800C6                   '/* ͨ��ʹ��AND���룩����������Դ��Ŀ����������ڵ���ɫ�ϲ�
Public Const SRCCOPY = &HCC0020                  '/* ��Դ��������ֱ�ӿ�����Ŀ���������
Public Const SRCERASE = &H440328                 '/* ͨ��ʹ��AND���룩��������Ŀ�����������ɫȡ������Դ�����������ɫֵ�ϲ���
Public Const SRCINVERT = &H660046                '/* ͨ��ʹ�ò����͵�XOR����򣩲�������Դ��Ŀ������������ɫ�ϲ���
Public Const SRCPAINT = &HEE0086                 '/* ͨ��ʹ�ò����͵�OR���򣩲�������Դ��Ŀ������������ɫ�ϲ���
Public Const WHITENESS = &HFF0062                '/* ʹ���������ɫ��������1�йص���ɫ���Ŀ��������򡣣�����ȱʡ�����ɫ����˵�������ɫ���ǰ�ɫ����

'--- for mouse_event
Public Const MOUSE_MOVED = &H1
Public Const MOUSEEVENTF_ABSOLUTE = &H8000       '/*
Public Const MOUSEEVENTF_LEFTDOWN = &H2          '/* ģ������������
Public Const MOUSEEVENTF_LEFTUP = &H4            '/* ģ��������̧��
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20       '/* ģ������м�����
Public Const MOUSEEVENTF_MIDDLEUP = &H40         '/* ģ������м�����
Public Const MOUSEEVENTF_MOVE = &H1              '/* �ƶ���� */
Public Const MOUSEEVENTF_RIGHTDOWN = &H8         '/* ģ������Ҽ�����
Public Const MOUSEEVENTF_RIGHTUP = &H10          '/* ģ������Ҽ�����
Public Const MOUSETRAILS = 39                    '/*

Public Const BMP_MAGIC_COOKIE = 19778            '/* this is equivalent to ascii string "BM" */
' constants for the biCompression field
Public Const BI_RGB = 0&
Public Const BI_RLE4 = 2&
Public Const BI_RLE8 = 1&
Public Const BI_BITFIELDS = 3&
'Public Const BITSPIXEL = 12                     '/* Number of bits per pixel
' DIB color table identifiers
Public Const DIB_PAL_COLORS = 1                  '/* ����ɫ����װ��һ��16λ�������飬�����뵱ǰѡ���ĵ�ɫ���й� color table in palette indices
Public Const DIB_PAL_INDICES = 2                 '/* No color table indices into surf palette
Public Const DIB_PAL_LOGINDICES = 4              '/* No color table indices into DC palette
Public Const DIB_PAL_PHYSINDICES = 2             '/* No color table indices into surf palette
Public Const DIB_RGB_COLORS = 0                  '/* ����ɫ����װ��RGB��ɫ

' BLENDFUNCTION AlphaFormat-Konstante
Public Const AC_SRC_ALPHA = &H1
' BLENDFUNCTION BlendOp-Konstante
Public Const AC_SRC_OVER = &H0

' ======================================================================================
' Methods
' ======================================================================================
' ����SetBkModen����BkMode
Public Enum KhanBackStyles
    TRANSPARENT = 1                              '/* ͸������������������� */
    OPAQUE = 2                                   '/* �õ�ǰ�ı���ɫ������߻��ʡ���Ӱˢ���Լ��ַ��Ŀ�϶ */
    NEWTRANSPARENT = 3                           '/* NT4: Uses chroma-keying upon BitBlt. Undocumented feature that is not working on Windows 2000/XP.
End Enum

' ����ε����ģʽ
Public Enum KhanPolyFillModeFalgs
    ALTERNATE = 1                                '/* �������
    WINDING = 2                                  '/* ���ݻ�ͼ�������
End Enum

' DrawIconEx
Public Enum KhanDrawIconExFlags
    DI_MASK = &H1                                '/* ��ͼʱʹ��ͼ���MASK���֣��絥��ʹ�ã��ɻ��ͼ�����ģ��
    DI_IMAGE = &H2                               '/* ��ͼʱʹ��ͼ���XOR���֣���ͼ��û��͸������
    DI_NORMAL = &H3                              '/* �ó��淽ʽ��ͼ���ϲ� DI_IMAGE �� DI_MASK��
    DI_COMPAT = &H4                              '/* ����׼��ϵͳָ�룬������ָ����ͼ��
    DI_DEFAULTSIZE = &H8                         '/* ����cxWidth��cyWidth���ã�������ԭʼ��ͼ���С
End Enum

'ָ����װ��ͼ������,LoadImage,CopyImage
Public Enum KhanImageTypes
    IMAGE_BITMAP = 0
    IMAGE_ICON = 1
    IMAGE_CURSOR = 2
    IMAGE_ENHMETAFILE = 3
End Enum

Public Enum KhanImageFalgs
    LR_COLOR = &H2                               '/*
    LR_COPYRETURNORG = &H4                       '/* ��ʾ����һ��ͼ��ľ�ȷ�����������Բ���cxDesired��cyDesired
    LR_COPYDELETEORG = &H8                       '/* ��ʾ����һ��������ɾ��ԭʼͼ��
    LR_CREATEDIBSECTION = &H2000                 '/* ������uTypeָ��ΪIMAGE_BITMAPʱ��ʹ�ú�������һ��DIB����λͼ��������һ�����ݵ�λͼ�������־��װ��һ��λͼ��������ӳ��������ɫ����ʾ�豸ʱ�ǳ����á�
    LR_DEFAULTCOLOR = &H0                        '/* �Գ��淽ʽ����ͼ��
    LR_DEFAULTSIZE = &H40                        '/* �� cxDesired��cyDesiredδ����Ϊ�㣬ʹ��ϵͳָ���Ĺ���ֵ��ʶ����ͼ��Ŀ�͸ߡ���������������������cxDesired��cyDesired����Ϊ�㣬����ʹ��ʵ����Դ�ߴ硣�����Դ�������ͼ����ʹ�õ�һ��ͼ��Ĵ�С��
    LR_LOADFROMFILE = &H10                       '/* ���ݲ���lpszName��ֵװ��ͼ�������δ��������lpszName��ֵΪ��Դ���ơ�
    LR_LOADMAP3DCOLORS = &H1000                  '/* ��ͼ���е����(Dk Gray RGB��128��128��128��)����(Gray RGB��192��192��192��)���Լ�ǳ��(Gray RGB��223��223��223��)���ض��滻��COLOR_3DSHADOW��COLOR_3DFACE�Լ�COLOR_3DLIGHT�ĵ�ǰ����
    LR_LOADTRANSPARENT = &H20                    '/* ��fuLoad����LR_LOADTRANSPARENT��LR_LOADMAP3DCOLORS����ֵ����LRLOADTRANSPARENT���ȡ����ǣ���ɫ��ӿ���COLOR_3DFACE�����������COLOR_WINDOW��
    LR_MONOCHROME = &H1                          '/* ��ͼ��ת���ɵ�ɫ
    LR_SHARED = &H8000                           '/* ��ͼ�񽫱����װ���������LR_SHAREDδ�����ã�������ͬһ����Դ�ڶ��ε������ͼ���Ǿͻ���װ���Ա����ͼ���ҷ��ز�ͬ�ľ����
    LR_COPYFROMRESOURCE = &H4000                 '/*
End Enum

Public Enum KhanDrawTextStyles
    DT_BOTTOM = &H8&                             '/* ����ͬʱָ��DT_SINGLE��ָʾ�ı������ʽ�����εĵױ�
    DT_CALCRECT = &H400&                         '/* ���������������ʽ�����Σ����л�ͼʱ���εĵױ߸�����Ҫ������չ���Ա������������֣����л�ͼʱ����չ���ε��Ҳࡣ��������֡���lpRect����ָ���ľ��λ�������������ֵ
    DT_CENTER = &H1&                             '/* �ı���ֱ����
    DT_EXPANDTABS = &H40&                        '/* ������ֵ�ʱ�򣬶��Ʊ�վ������չ��Ĭ�ϵ��Ʊ�վ�����8���ַ������ǣ�����DT_TABSTOP��־�ı������趨
    DT_EXTERNALLEADING = &H200&                  '/* �����ı��и߶ȵ�ʱ��ʹ�õ�ǰ������ⲿ������ԣ�the external leading attribute��
    DT_INTERNAL = &H1000&                        '/* Uses the system font to calculate text metrics
    DT_LEFT = &H0&                               '/* �ı������
    DT_NOCLIP = &H100&                           '/* �������ʱ�����е�ָ���ľ��Σ�DrawTextEx is somewhat faster when DT_NOCLIP is used.
    DT_NOPREFIX = &H800&                         '/* ͨ����������Ϊ & �ַ���ʾӦΪ��һ���ַ������»��ߡ��ñ�־��ֹ������Ϊ
    DT_RIGHT = &H2&                              '/* �ı��Ҷ���
    DT_SINGLELINE = &H20&                        '/* ֻ������
    DT_TABSTOP = &H80&                           '/* ָ���µ��Ʊ�վ��࣬������������ĸ�8λ
    DT_TOP = &H0&                                '/* ����ͬʱָ��DT_SINGLE��ָʾ�ı������ʽ�����εĵױ�
    DT_VCENTER = &H4&                            '/* ����ͬʱָ��DT_SINGLE��ָʾ�ı������ʽ�����ε��в�
    DT_WORDBREAK = &H10&                         '/* �����Զ����С�����SetTextAlign����������TA_UPDATECP��־���������������Ч
' #if(WINVER >= =&H0400)
    DT_EDITCONTROL = &H2000&                     '/* ��һ�����б༭�ؼ�����ģ�⡣����ʾ���ֿɼ�����
    DT_END_ELLIPSIS = &H8000&                    '/* �����ִ������ھ�����ȫ�����£�����ĩβ��ʾʡ�Ժ�
    DT_PATH_ELLIPSIS = &H4000&                   '/* ���ִ������� \ �ַ�������ʡ�Ժ��滻�ִ����ݣ�ʹ�����ھ�����ȫ�����¡����磬һ���ܳ���·�������ܻ���������ʾ����c:\windows\...\doc\readme.txt
    DT_MODIFYSTRING = &H10000                    '/* ��ָ����DT_ENDELLIPSES �� DT_PATHELLIPSES���ͻ���ִ������޸ģ�ʹ����ʵ����ʾ���ִ����
    DT_RTLREADING = &H20000                      '/* ��ѡ���豸��������������ϣ������������ϵ���ʹ��ҵ����������
    DT_WORD_ELLIPSIS = &H40000                   '/* Truncates any word that does not fit in the rectangle and adds ellipses. Compare with DT_END_ELLIPSIS and DT_PATH_ELLIPSIS.
End Enum

Public Enum KhanDrawFrameControlType
    DFC_CAPTION = 1                              '/* Title bar.
    DFC_MENU = 2                                 '/* Menu bar.
    DFC_SCROLL = 3                               '/* Scroll bar.
    DFC_BUTTON = 4                               '/* Standard button.
    DFC_POPUPMENU = 5                            '/* <b>Windows 98/Me, Windows 2000 or later:</b> Popup menu item.
End Enum

Public Enum KhanDrawFrameControlStyle
    DFCS_BUTTONCHECK = &H0                       '/* Check box.
    DFCS_BUTTONRADIOIMAGE = &H1                  '/* Image for radio button (nonsquare needs image).
    DFCS_BUTTONRADIOMASK = &H2                   '/* Mask for radio button (nonsquare needs mask).
    DFCS_BUTTONRADIO = &H4                       '/* Radio button.
    DFCS_BUTTON3STATE = &H8                      '/* Three-state button.
    DFCS_BUTTONPUSH = &H10                       '/* Push button.
    DFCS_CAPTIONCLOSE = &H0                      '/* <b>Close</b> button.
    DFCS_CAPTIONMIN = &H1                        '/* <b>Minimize</b> button.
    DFCS_CAPTIONMAX = &H2                        '/* <b>Maximize</b> button.
    DFCS_CAPTIONRESTORE = &H3                    '/* <b>Restore</b> button.
    DFCS_CAPTIONHELP = &H4                       '/* <b>Help</b> button.
    DFCS_MENUARROW = &H0                         '/* Submenu arrow.
    DFCS_MENUCHECK = &H1                         '/* Check mark.
    DFCS_MENUBULLET = &H2                        '/* Bullet.
    DFCS_MENUARROWRIGHT = &H4                    '/* Submenu arrow pointing left. This is used for the right-to-left cascading menus used with right-to-left languages such as Arabic or Hebrew.
    DFCS_SCROLLUP = &H0                          '/* Up arrow of scroll bar.
    DFCS_SCROLLDOWN = &H1                        '/* Down arrow of scroll bar.
    DFCS_SCROLLLEFT = &H2                        '/* Left arrow of scroll bar.
    DFCS_SCROLLRIGHT = &H3                       '/* Right arrow of scroll bar.
    DFCS_SCROLLCOMBOBOX = &H5                    '/* Combo box scroll bar.
    DFCS_SCROLLSIZEGRIP = &H8                    '/* Size grip in bottom-right corner of window.
    DFCS_SCROLLSIZEGRIPRIGHT = &H10              '/* Size grip in bottom-left corner of window. This is used with right-to-left languages such as Arabic or Hebrew.
    DFCS_INACTIVE = &H100                        '/* Button is inactive (grayed).
    DFCS_PUSHED = &H200                          '/* Button is pushed.
    DFCS_CHECKED = &H400                         '/* Button is checked.
    DFCS_TRANSPARENT = &H800                     '/* <b>Windows 98/Me, Windows 2000 or later:</b> The background remains untouched.
    DFCS_HOT = &H1000                            '/* <b>Windows 98/Me, Windows 2000 or later:</b> Button is hot-tracked.
    DFCS_ADJUSTRECT = &H2000                     '/* Bounding rectangle is adjusted to exclude the surrounding edge of the push button.
    DFCS_FLAT = &H4000                           '/* Button has a flat border.
    DFCS_MONO = &H8000                           '/* Button has a monochrome border.
End Enum

' ָ��������ʽ������CreatePen�Ĳ���CreatePen��ʹ�õĳ���
Public Enum KhanPenStyles
    ' CreatePen��ExtCreatePen
    ' ���ʵ���ʽ
    PS_SOLID = 0                                 '/* ���ʻ�������ʵ�� */
    PS_DASH = 1                                  '/* ���ʻ����������ߣ�nWidth������1�� */
    PS_DOT = 2                                   '/* ���ʻ������ǵ��ߣ�nWidth������1�� */
    PS_DASHDOT = 3                               '/* ���ʻ������ǵ㻮�ߣ�nWidth������1�� */
    PS_DASHDOTDOT = 4                            '/* ���ʻ������ǵ�-��-���ߣ�nWidth������1�� */
    PS_NULL = 5                                  '/* ���ʲ��ܻ�ͼ */
    PS_INSIDEFRAME = 6                           '/* ����������Բ�����Ρ�Բ�Ǿ��Ρ���ͼ�Լ��ҵ����ɵķ�ն�����л�ͼ����ָ����׼ȷRGB��ɫ�����ڣ��ͽ��ж������� */
    ' ExtCreatePen
    ' ���ʵ���ʽ
    PS_USERSTYLE = 7                             '/* <b>Windows NT/2000:</b> The pen uses a styling array supplied by the user.
    PS_ALTERNATE = 8                             '/* <b>Windows NT/2000:</b> The pen sets every other pixel. (This style is applicable only for cosmetic pens.)
    ' ���ʵıʼ�
    PS_ENDCAP_ROUND = &H0                        '/* End caps are round.
    PS_ENDCAP_SQUARE = &H100                     '/* End caps are square.
    PS_ENDCAP_FLAT = &H200                       '/* End caps are flat.
    PS_ENDCAP_MASK = &HF00                       '/* Mask for previous PS_ENDCAP_XXX values.
    ' ��ͼ���������߶λ���·��������ֱ�ߵķ�ʽ
    PS_JOIN_ROUND = &H0                          '/* Joins are beveled.
    PS_JOIN_BEVEL = &H1000                       '/* Joins are mitered when they are within the current limit set by the SetMiterLimit function. If it exceeds this limit, the join is beveled.
    PS_JOIN_MITER = &H2000                       '/* Joins are round.
    PS_JOIN_MASK = &HF000                        '/* Mask for previous PS_JOIN_XXX values.
    ' ���ʵ�����
    PS_COSMETIC = &H0                            '/* The pen is cosmetic.
    PS_GEOMETRIC = &H10000                       '/* The pen is geometric.
    '
    PS_STYLE_MASK = &HF                          '/* Mask for previous PS_XXX values.
    PS_TYPE_MASK = &HF0000                       '/* Mask for previous PS_XXX (pen type).
End Enum

Public Enum KhanBrushStyle
    BS_SOLID = 0                                 '/* Solid brush.
    BS_HOLLOW = 1                                '/* Hollow brush.
    BS_NULL = 1                                  '/* Same as BS_HOLLOW.
    BS_HATCHED = 2                               '/* Hatched brush.
    BS_PATTERN = 3                               '/* Pattern brush defined by a memory bitmap.
    BS_INDEXED = 4                               '/*
    BS_DIBPATTERN = 5                            '/* A pattern brush defined by a device-independent bitmap (DIB) specification.
    BS_DIBPATTERNPT = 6                          '/* A pattern brush defined by a device-independent bitmap (DIB) specification. If <b>lbStyle</b> is BS_DIBPATTERNPT, the <b>lbHatch</b> member contains a pointer to a packed DIB.
    BS_PATTERN8X8 = 7                            '/* Same as BS_PATTERN.
    BS_DIBPATTERN8X8 = 8                         '/* Same as BS_DIBPATTERN.
    BS_MONOPATTERN = 9                           '/* The brush is a monochrome (black & white) bitmap.
End Enum

Public Enum KhanHatchStyles
    HS_HORIZONTAL = 0                            '/* Horizontal hatch.
    HS_VERTICAL = 1                              '/* Vertical hatch.
    HS_FDIAGONAL = 2                             '/* A 45-degree downward, left-to-right hatch.
    HS_BDIAGONAL = 3                             '/* A 45-degree upward, left-to-right hatch.
    HS_CROSS = 4                                 '/* Horizontal and vertical cross-hatch.
    HS_DIAGCROSS = 5                             '/* A 45-degree crosshatch.
End Enum

' DrawEdge
Public Enum KhanBorderStyles
    BDR_RAISEDOUTER = &H1                        '/* Raised outer edge.
    BDR_SUNKENOUTER = &H2                        '/* Sunken outer edge.
    BDR_RAISEDINNER = &H4                        '/* Raised inner edge.
    BDR_SUNKENINNER = &H8                        '/* Sunken inner edge.
    BDR_OUTER = &H3                              '/* (BDR_RAISEDOUTER Or BDR_SUNKENOUTER)
    BDR_INNER = &HC                              '/* (BDR_RAISEDINNER Or BDR_SUNKENINNER)
    BDR_RAISED = &H5
    BDR_SUNKEN = &HA
    EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
    EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
    EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
    EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
End Enum

Public Enum KhanBorderFlags
    BF_LEFT = &H1                                '/* Left side of border rectangle.
    BF_TOP = &H2                                 '/* Top of border rectangle.
    BF_RIGHT = &H4                               '/* Right side of border rectangle.
    BF_BOTTOM = &H8                              '/* Bottom of border rectangle.
    BF_TOPLEFT = (BF_TOP Or BF_LEFT)
    BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
    BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
    BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
    BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
    BF_DIAGONAL = &H10                           '/* Diagonal border.
    BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
    BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
    BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
    BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
    BF_MIDDLE = &H800                            '/* Fill in the middle.
    BF_SOFT = &H1000                             '/* Use for softer buttons.
    BF_ADJUST = &H2000                           '/* Calculate the space left over.
    BF_FLAT = &H4000                             '/* For flat rather than 3-D borders.
    BF_MONO = &H8000&                            '/* For monochrome borders
End Enum

' ����ָ��һ����λ�ú�״̬������SetWindowPos����
Public Enum KhanSetWindowPosStyles
    HWND_BOTTOM = 1                              '/* ���������ڴ����б�ײ� */
    HWND_NOTOPMOST = -2                          '/* �����������б�������λ���κ�������ڵĺ��� */
    HWND_TOP = 0                                 '/* ����������Z���еĶ�����Z���д����ڷּ��ṹ�У��������һ����������Ĵ�����ʾ��˳�� */
    HWND_TOPMOST = -1                            '/* �����������б�������λ���κ�������ڵ�ǰ�� */
    SWP_SHOWWINDOW = &H40                        '/* ��ʾ���� */
    SWP_HIDEWINDOW = &H80                        '/* ���ش��� */
    SWP_FRAMECHANGED = &H20                      '/* ǿ��һ��WM_NCCALCSIZE��Ϣ���봰�ڣ���ʹ���ڵĴ�Сû�иı� */
    SWP_NOACTIVATE = &H10                        '/* ������� */
    SWP_NOCOPYBITS = &H100                       '
    SWP_NOMOVE = &H2                             '/* ���ֵ�ǰλ�ã�x��y�趨�������ԣ� */
    SWP_NOOWNERZORDER = &H200                    '/* Don't do owner Z ordering */
    SWP_NOREDRAW = &H8                           '/* ���ڲ��Զ��ػ� */
    SWP_NOREPOSITION = SWP_NOOWNERZORDER         '
    SWP_NOSIZE = &H1                             '/* ���ֵ�ǰ��С��cx��cy�ᱻ���ԣ� */
    SWP_NOZORDER = &H4                           '/* ���ִ������б�ĵ�ǰλ�ã�hWndInsertAfter�������ԣ� */
    SWP_DRAWFRAME = SWP_FRAMECHANGED             '/* Χ�ƴ��ڻ�һ���� */
'    HWND_BROADCAST = &HFFFF&
'    HWND_DESKTOP = 0
End Enum

' ָ���������ڵķ��
Public Enum KhanCreateWindowSytles
    ' CreateWindow
    WS_BORDER = &H800000                         '/* ����һ�����߿�Ĵ��ڡ�
    WS_CAPTION = &HC00000                        '/* ����һ���б����Ĵ��ڣ�����WS_BODER��񣩡�
    WS_CHILD = &H40000000                        '/* ����һ���Ӵ��ڡ�����������WS_POPVP�����á�
    WS_CHILDWINDOW = (WS_CHILD)                  '/* ��WS_CHILD��ͬ��
    WS_CLIPCHILDREN = &H2000000                  '/* ���ڸ������ڻ�ͼʱ���ų��Ӵ��������ڴ���������ʱʹ��������
    WS_CLIPSIBLINGS = &H4000000                  '/* �ų��Ӵ���֮����������Ҳ���ǣ���һ���ض��Ĵ��ڽ��յ�WM_PAINT��Ϣʱ��WS_CLIPSIBLINGS ������в�������ų��ڻ�ͼ֮�⣬ֻ�ػ�ָ�����Ӵ��ڡ����δָ��WS_CLIPSIBLINGS��񣬲����Ӵ����ǲ���ģ������ػ��Ӵ��ڵĿͻ���ʱ���ͻ��ػ��ڽ����Ӵ��ڡ�
    WS_DISABLED = &H8000000                      '/* ����һ����ʼ״̬Ϊ��ֹ���Ӵ��ڡ�һ����ֹ״̬�Ĵ��ղ��ܽ��������û���������Ϣ��
    WS_DLGFRAME = &H400000                       '/* ����һ�����Ի���߿���Ĵ��ڡ����ַ��Ĵ��ڲ��ܴ���������
    WS_GROUP = &H20000                           '/* ָ��һ����Ƶĵ�һ�����ơ�����������ɵ�һ�����ƺ������Ŀ�����ɣ��Եڶ������ƿ�ʼÿ�����ƣ�����WS_GROUP���ÿ����ĵ�һ�����ƴ���WS_TABSTOP��񣬴Ӷ�ʹ�û�����������ƶ����û�������ʹ�ù�������ڵĿ��Ƽ�ı���̽��㡣
    WS_HSCROLL = &H100000                        '/* ����һ����ˮƽ�������Ĵ��ڡ�
    WS_MAXIMIZE = &H1000000                      '/* ����һ��������󻯰�ť�Ĵ��ڡ��÷������WS_EX_CONTEXTHELP���ͬʱ���֣�ͬʱ����ָ��WS_SYSMENU���
    WS_MAXIMIZEBOX = &H10000                     '/*
    WS_MINIMIZE = &H20000000                     '/* ����һ����ʼ״̬Ϊ��С��״̬�Ĵ��ڡ�
    WS_ICONIC = WS_MINIMIZE                      '/* ����һ����ʼ״̬Ϊ��С��״̬�Ĵ��ڡ���WS_MINIMIZE�����ͬ��
    WS_MINIMIZEBOX = &H20000                     '/*
    WS_OVERLAPPED = &H0&                         '/* ����һ������Ĵ��ڡ�һ������Ĵ�����һ����������һ���߿���WS_TILED�����ͬ
    WS_POPUP = &H80000000                        '/* ����һ������ʽ���ڡ��÷������WS_CHLD���ͬʱʹ�á�
    WS_SYSMENU = &H80000                         '/* ����һ���ڱ������ϴ��д��ڲ˵��Ĵ��ڣ�����ͬʱ�趨WS_CAPTION���
    WS_TABSTOP = &H10000                         '/* ����һ�����ƣ�����������û�����Tab��ʱ���Ի�ü��̽��㡣����Tab����ʹ���̽���ת�Ƶ���һ����WS_TABSTOP���Ŀ��ơ�
    WS_THICKFRAME = &H40000                      '/* ����һ�����пɵ��߿�Ĵ��ڡ�
    WS_SIZEBOX = WS_THICKFRAME                   '/* ��WS_THICKFRAME�����ͬ
    WS_TILED = WS_OVERLAPPED                     '/* ����һ������Ĵ��ڡ�һ������Ĵ�����һ�������һ���߿���WS_OVERLAPPED�����ͬ��
    WS_VISIBLE = &H10000000                      '/* ����һ����ʼ״̬Ϊ�ɼ��Ĵ��ڡ�
    WS_VSCROLL = &H200000                        '/* ����һ���д�ֱ�������Ĵ��ڡ�
    WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
    WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW         '/* ����һ������WS_OVERLAPPED��WS_CAPTION��WS_SYSMENU MS_THICKFRAME��
    WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU) '/* ����һ������WS_BORDER��WS_POPUP,WS_SYSMENU���Ĵ��ڣ�WS_CAPTION��WS_POPUPWINDOW����ͬʱ�趨����ʹ����ĳ���ɼ���
    ' CreateWindowEx
    WS_EX_ACCEPTFILES = &H10&                    '/* ָ���Ը÷�񴴽��Ĵ��ڽ���һ����ק�ļ���
    WS_EX_APPWINDOW = &H40000                    '/* �����ڿɼ�ʱ����һ�����㴰�ڷ��õ��������ϡ�
    WS_EX_CLIENTEDGE = &H200                     '/* ָ��������һ������Ӱ�ı߽硣
    WS_EX_CONTEXTHELP = &H400                    '/* �ڴ��ڵı���������һ���ʺű�־�����û�������ʺ�ʱ��������Ϊһ���ʺŵ�ָ�롢��������һ���Ӵ��ڣ����Ӵ��ս��յ�WM_HELP��Ϣ���Ӵ���Ӧ�ý������Ϣ���ݸ������ڹ��̣���������ͨ��HELP_WM_HELP�������WinHelp���������HelpӦ�ó�����ʾһ�������Ӵ��ڰ�����Ϣ�ĵ���ʽ���ڡ� WS_EX_CONTEXTHELP������WS_MAXIMIZEBOX��WS_MINIMIZEBOXͬʱʹ�á�
    WS_EX_CONTROLPARENT = &H10000                '/* �����û�ʹ��Tab���ڴ��ڵ��Ӵ��ڼ�������
    WS_EX_DLGMODALFRAME = &H1&                   '/* ����һ����˫�ߵĴ��ڣ��ô��ڿ�����dwStyle��ָ��WS_CAPTION���������һ����������
    WS_EX_LEFT = &H0                             '/* ���ھ�����������ԣ�����ȱʡ���õġ�
    WS_EX_LEFTSCROLLBAR = &H4000                 '/* ��������������Hebrew��Arabic��������֧��reading order alignment�����ԣ����������������ڣ����ڿͻ������󲿷֡������������ԣ��ڸ÷�񱻺��Բ��Ҳ���Ϊ������
    WS_EX_LTRREADING = &H0                       '/* �����ı���LEFT��RIGHT���������ң����Ե�˳����ʾ������ȱʡ���õġ�
    WS_EX_MDICHILD = &H40                        '/* ����һ��MDI�Ӵ��ڡ�
    WS_EX_NOACTIVATE = &H8000000                 '/*
    WS_EX_NOPATARENTNOTIFY = &H4&                '/* ָ���������񴴽��Ĵ����ڱ�����������ʱ���򸸴��ڷ���WM_PARENTNOTFY��Ϣ��
    WS_EX_OVERLAPPEDWINDOW = &H300               '/*
    WS_EX_PALETTEWINDOW = &H188                  '/* WS_EX_WINDOWEDGE, WS_EX_TOOLWINDOW��WS_WX_TOPMOST�������WS_EX_RIGHT:���ھ�����ͨ���Ҷ������ԣ��������ڴ����ࡣֻ���������������Hebrew,Arabic������֧�ֶ�˳����루reading order alignment��������ʱ�÷�����Ч�����򣬺��Ըñ�־���Ҳ���Ϊ������
    WS_EX_RIGHT = &H1000                         '/*
    WS_EX_RIGHTSCROLLBAR = &H0                   '/* ��ֱ�������ڴ��ڵ��ұ߽硣����ȱʡ���õġ�
    WS_EX_RTLREADING = &H2000                    '/* ��������������Hebrew��Arabic��������֧�ֶ�˳����루reading order alignment�������ԣ��򴰿��ı���һ�������ң�RIGHT��LEFT˳��Ķ���˳�������������ԣ��ڸ÷�񱻺��Բ��Ҳ���Ϊ������
    WS_EX_STATICEDGE = &H20000                   '/* Ϊ�������û���������һ��3һά�߽���
    WS_EX_TOOLWINDOW = &H80                      '/*
    WS_EX_TOPMOST = &H8&                         '/* ָ���Ը÷�񴴽��Ĵ���Ӧ���������з���߲㴰�ڵ����沢��ͣ������L����ʹ����δ�����ʹ�ú���SetWindowPos�����ú���ȥ������
    WS_EX_TRANSPARENT = &H20&                    '/* ָ���������񴴽��Ĵ����ڴ����µ�ͬ���������ػ�ʱ���ô��ڲſ����ػ���
    WS_EX_WINDOWEDGE = &H100
End Enum

' Windows�����йص���Ϣ������GetSystemMetrics����
Public Enum KhanSystemMetricsFlags
    SM_CXSCREEN = 0                              '/* ��Ļ��С */
    SM_CYSCREEN = 1                              '/* ��Ļ��С */
    SM_CXVSCROLL = 2                             '/* ��ֱ�������еļ�ͷ��ť�Ĵ�С */
    SM_CYHSCROLL = 3                             '/* ˮƽ�������ϵļ�ͷ��С */
    SM_CYCAPTION = 4                             '/* ���ڱ���ĸ߶� */
    SM_CXBORDER = 5                              '/* �ߴ粻�ɱ�߿�Ĵ�С */
    SM_CYBORDER = 6                              '/* �ߴ粻�ɱ�߿�Ĵ�С */
    SM_CXDLGFRAME = 7                            '/* �Ի���߿�Ĵ�С */
    SM_CYDLGFRAME = 8                            '/* �Ի���߿�Ĵ�С */
    SM_CYVTHUMB = 9                              '/* ��������ˮƽ�������ϵĴ�С */
    SM_CXHTHUMB = 10                             '/* ��������ˮƽ�������ϵĴ�С */
    SM_CXICON = 11                               '/* ��׼ͼ��Ĵ�С */
    SM_CYICON = 12                               '/* ��׼ͼ��Ĵ�С */
    SM_CXCURSOR = 13                             '/* ��׼ָ���С */
    SM_CYCURSOR = 14                             '/* ��׼ָ���С */
    SM_CYMENU = 15                               '/* �˵��߶� */
    SM_CXFULLSCREEN = 16                         '/* ��󻯴��ڿͻ����Ĵ�С */
    SM_CYFULLSCREEN = 17                         '/* ��󻯴��ڿͻ����Ĵ�С */
    SM_CYKANJIWINDOW = 18                        '/* Kanji���ڵĴ�С��Height of Kanji window�� */
    SM_MOUSEPRESENT = 19                         '/* �簲װ�������ΪTRUE */
    SM_CYVSCROLL = 20                            '/* ��ֱ�������еļ�ͷ��ť�Ĵ�С */
    SM_CXHSCROLL = 21                            '/* ˮƽ�������ϵļ�ͷ��С */
    SM_DEBUG = 22                                '/* ��windows�ĵ��԰��������У���ΪTRUE */
    SM_SWAPBUTTON = 23
    SM_RESERVED1 = 24
    SM_RESERVED2 = 25
    SM_RESERVED3 = 26
    SM_RESERVED4 = 27
    SM_CXMIN = 28                                '/* ���ڵ���С�ߴ� */
    SM_CYMIN = 29                                '/* ���ڵ���С�ߴ� */
    SM_CXSIZE = 30                               '/* ������λͼ�Ĵ�С */
    SM_CYSIZE = 31                               '/* ������λͼ�Ĵ�С */
    SM_CXFRAME = 32                              '/* �ߴ�ɱ�߿�Ĵ�С����win95��nt 4.0��ʹ��SM_C?FIXEDFRAME�� */
    SM_CYFRAME = 33                              '/* �ߴ�ɱ�߿�Ĵ�С */
    SM_CXMINTRACK = 34                           '/* ���ڵ���С�켣��� */
    SM_CYMINTRACK = 35                           '/* ���ڵ���С�켣��� */
    SM_CXDOUBLECLK = 36                          '/* ˫������Ĵ�С��ָ����Ļ��һ���ض�����ʾ����ֻ���������������������������굥�������п��ܱ�����˫���¼����� */
    SM_CYDOUBLECLK = 37                          '/* ˫������Ĵ�С */
    SM_CXICONSPACING = 38                        '/* ����ͼ��֮��ļ�����롣��win95��nt 4.0����ָ��ͼ��ļ�� */
    SM_CYICONSPACING = 39                        '/* ����ͼ��֮��ļ�����롣��win95��nt 4.0����ָ��ͼ��ļ�� */
    SM_MENUDROPALIGNMENT = 40                    '/* �絯��ʽ�˵�����˵�����Ŀ����࣬��Ϊ�� */
    SM_PENWINDOWS = 41                           '/* ��װ����֧�ֱʴ��ڵ�DLL�����ʾ�ʴ��ڵľ�� */
    SM_DBCSENABLED = 42                          '/* ��֧��˫�ֽ���ΪTRUE */
    SM_CMOUSEBUTTONS = 43                        '/* ��갴ť������������������û����꣬��Ϊ�� */
    SM_CMETRICS = 44                             '/* ����ϵͳ���������� */
End Enum

' SetMapMode
Public Enum KhanMapModeStyles
    MM_ANISOTROPIC = 8                           '/* �߼���λת���ɾ����������������ⵥλ����SetWindowExtEx��SetViewportExtEx������ָ����λ������ͱ�����
    MM_HIENGLISH = 5                             '/* ÿ���߼���λת��Ϊ0.001inch(Ӣ��)��X�����������ң�Y������������
    MM_HIMETRIC = 3                              '/* ÿ���߼���λת��Ϊ0.01millimeter(����)��X���������ң�Y�����������ϡ�
    MM_ISOTROPIC = 7                             '/* �ӿںʹ��ڷ�Χ���⣬ֻ��x��y�߼���Ԫ�ߴ�Ҫ��ͬ
    MM_LOENGLISH = 4                             '/* ÿ���߼���λת��ΪӢ�磬X���������ң�Y���������ϡ�
    MM_LOMETRIC = 2                              '/* ÿ���߼���λת��Ϊ���ף�X���������ң�Y���������ϡ�
    MM_TEXT = 1                                  '/* ÿ���߼���λת��Ϊһ�����ñ��أ�X���������ң�Y���������¡�
    MM_TWIPS = 6                                 '/* ÿ���߼���λת��Ϊ1 twip (1/1440 inch)��X���������ң�Y�������ϡ�
End Enum

' GetROP2,SetROP2
Public Enum EnumDrawModeFlags
    R2_BLACK = 1                                 '/* ��ɫ
    R2_COPYPEN = 13                              '/* ������ɫ
    R2_LAST = 16
    R2_MASKNOTPEN = 3                            '/* ������ɫ�ķ�ɫ����ʾ��ɫ����AND����
    R2_MASKPEN = 9                               '/* ��ʾ��ɫ�뻭����ɫ����AND����
    R2_MASKPENNOT = 5                            '/* ��ʾ��ɫ�ķ�ɫ�뻭����ɫ����AND����
    R2_MERGENOTPEN = 12                          '/* ������ɫ�ķ�ɫ����ʾ��ɫ����OR����
    R2_MERGEPEN = 15                             '/* ������ɫ����ʾ��ɫ����OR����
    R2_MERGEPENNOT = 14                          '/* ��ʾ��ɫ�ķ�ɫ�뻭����ɫ����OR����
    R2_NOP = 11                                  '/* ����
    R2_NOT = 6                                   '/* ��ǰ��ʾ��ɫ�ķ�ɫ
    R2_NOTCOPYPEN = 4                            '/* R2_COPYPEN�ķ�ɫ
    R2_NOTMASKPEN = 8                            '/* R2_MASKPEN�ķ�ɫ
    R2_NOTMERGEPEN = 2                           '/* R2_MERGEPEN�ķ�ɫ
    R2_NOTXORPEN = 10                            '/* R2_XORPEN�ķ�ɫ
    R2_WHITE = 16                                '/* ��ɫ
    R2_XORPEN = 7                                '/* ��ʾ��ɫ�뻭����ɫ�����������
End Enum

' ======================================================================================
' Types
' ======================================================================================

Public Type tagINITCOMMONCONTROLSEX              '/* icc
   dwSize                   As Long              '/* size of this structure
   dwICC                    As Long              '/* flags indicating which classes to be initialized.
End Type

Public Type POINTAPI
   x                        As Long
   y                        As Long
End Type

Public Type RECT
   Left                     As Long
   Top                      As Long
   Right                    As Long
   Bottom                   As Long
End Type

Public Type LOGPEN
    lopnStyle               As Long
    lopnWidth               As POINTAPI
    lopnColor               As Long
End Type

Public Type LOGBRUSH
   lbStyle                  As Long
   lbColor                  As Long
   lbHatch                  As Long
End Type

' ����ṹ�����˸��ӵĻ�ͼ����������DrawTextEx
Public Type DRAWTEXTPARAMS
    cbSize                  As Long              '/* Specifies the structure size, in bytes */
    iTabLength              As Long              '/* Specifies the size of each tab stop, in units equal to the average character width */
    iLeftMargin             As Long              '/* Specifies the left margin, in units equal to the average character width */
    iRightMargin            As Long              '/* Specifies the right margin, in units equal to the average character width */
    uiLengthDrawn           As Long              '/* Receives the number of characters processed by DrawTextEx, including white-space characters. */
                                                 '/* The number can be the length of the string or the index of the first line that falls below the drawing area. */
                                                 '/* Note that DrawTextEx always processes the entire string if the DT_NOCLIP formatting flag is specified */
End Type

Private Const LF_FACESIZE   As Long = 32
Public Type LOGFONT
   lfHeight                 As Long              '/* The font size (see below) */
   lfWidth                  As Long              '/* Normally you don't set this, just let Windows create the Default */
   lfEscapement             As Long              '/* The angle, in 0.1 degrees, of the font */
   lfOrientation            As Long              '/* Leave as default */
   lfWeight                 As Long              '/* Bold, Extra Bold, Normal etc */
   lfItalic                 As Byte              '/* As it says */
   lfUnderline              As Byte              '/* As it says */
   lfStrikeOut              As Byte              '/* As it says */
   lfCharSet                As Byte              '/* As it says */
   lfOutPrecision           As Byte              '/* Leave for default */
   lfClipPrecision          As Byte              '/* Leave for defaultv
   lfQuality                As Byte              '/* Leave for default */
   lfPitchAndFamily         As Byte              '/* Leave for default */
   lfFaceName(LF_FACESIZE)  As Byte              '/* The font name converted to a byte array */
End Type

Public Type ICONINFO
   fIcon                    As Long
   xHotspot                 As Long
   yHotspot                 As Long
   hBmMask                  As Long
   hbmColor                 As Long
End Type

Public Type IMAGEINFO
    hBitmapImage            As Long
    hBitmapMask             As Long
    cPlanes                 As Long
    cBitsPerPixel           As Long
    rcImage                 As RECT
End Type

'/* DIB ���ļ���С���ܹ�ѶϢ */
Public Type BITMAPFILEHEADER
    bfType                  As Integer           '/* ָ���ļ����ͣ����� BM("magic cookie" - must be "BM" (19778)) */
    bfSize                  As Long              '/* ָ��λͼ�ļ���С����λԪ��Ϊ��λ */
    bfReserved1             As Integer           '/* ������������Ϊ0 */
    bfReserved2             As Integer           '/* ͬ�� */
    bfOffBits               As Long              '/* �Ӵ˼ܹ���λͼ����λ��λԪ��ƫ���� */
End Type

'/* �豸�޹�λͼ (DIB)�Ĵ�С����ɫ��Ϣ  (��λ�� bmp �ļ��Ŀ�ͷ��) 40 bytes */
Public Type BITMAPINFOHEADER
    biSize                  As Long              '/* �ṹ���� */
    biwidth                 As Long              '/* ָ��λͼ�Ŀ�ȣ�������Ϊ��λ */
    biheight                As Long              '/* ָ��λͼ�ĸ߶ȣ�������Ϊ��λ */
    biPlanes                As Integer           '/* ָ��Ŀ���豸�ļ���(����Ϊ 1 ) */
    biBitCount              As Integer           '/* λͼ����ɫλ��,ÿһ�����ص�λ(1��4��8��16��24��32) */
    biCompression           As Long              '/* ָ��ѹ������(BI_RGB Ϊ��ѹ��) */
    biSizeImage             As Long              '/* ͼ��Ĵ�С,���ֽ�Ϊ��λ,����BI_RGB��ʽ��,������Ϊ0 */
    biXPelsPerMeter         As Long              '/* ָ���豸ˮ׼�ֱ��ʣ���ÿ�׵�����Ϊ��λ */
    biYPelsPerMeter         As Long              '/* ��ֱ�ֱ��ʣ�����ͬ�� */
    biClrUsed               As Long              '/* ˵��λͼʵ��ʹ�õĲ�ɫ���е���ɫ������,��Ϊ0�Ļ�,˵��ʹ�����е�ɫ���� */
    biClrImportant          As Long              '/* ˵����ͼ����ʾ����ҪӰ�����ɫ��������Ŀ�������0����ʾ����Ҫ */
End Type

'/* �������ɺ졢�̡�����ɵ���ɫ��� */
Public Type RGBQUAD
    rgbBlue                 As Byte
    rgbGreen                As Byte
    rgbRed                  As Byte
    rgbReserved             As Byte              '/* '����������Ϊ 0 */
End Type

Public Type BITMAPINFO
    bmiHeader               As BITMAPINFOHEADER
    bmiColors               As RGBQUAD
End Type

Public Type BITMAPINFO_1BPP
   bmiHeader                As BITMAPINFOHEADER
   bmiColors(0 To 1)        As RGBQUAD
End Type

Public Type BITMAPINFO_4BPP
   bmiHeader                As BITMAPINFOHEADER
   bmiColors(0 To 15)       As RGBQUAD
End Type

Public Type BITMAPINFO_8BPP
   bmiHeader                As BITMAPINFOHEADER
   bmiColors(0 To 255)      As RGBQUAD
End Type

Public Type BITMAPINFO_ABOVE8
   bmiHeader                As BITMAPINFOHEADER
End Type

Public Type BITMAP
    bmType                  As Long              '/* Type of bitmap */
    bmWidth                 As Long              '/* Pixel width */
    bmHeight                As Long              '/* Pixel height */
    bmWidthBytes            As Long              '/* Byte width = 3 x Pixel width */
    bmPlanes                As Integer           '/* Color depth of bitmap */
    bmBitsPixel             As Integer           '/* Bits per pixel, must be 16 or 24 */
    bmBits                  As Long              '/* This is the pointer to the bitmap data */
End Type

' AlphaBlend
Public Type BLENDFUNCTION
   BlendOp                  As Byte
   BlendFlags               As Byte
   SourceConstantAlpha      As Byte
   AlphaFormat              As Byte
End Type

' ======================================================================================
' API declares:
' ======================================================================================

'����������������������������������������������������������������������������������������
'��-----------------------------��Ϣ��������Ϣ�жӺ���---------------------------------��
'��                                                                                    ��
'
' ����һ�����ڵĴ��ں�������һ����Ϣ�����Ǹ����ڡ�������Ϣ������ϣ�����ú������᷵�ء�
' SendMessageBynum�� SendMessageByString�Ǹú����ġ����Ͱ�ȫ��������ʽ
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' ��һ����ϢͶ�ݵ�ָ�����ڵ���Ϣ���С�Ͷ�ݵ���Ϣ����Windows�¼���������еõ�����
' ���Ǹ�ʱ�򣬻���ͬͶ�ݵ���Ϣ����ָ�����ڵĴ��ں������ر��ʺ���Щ����Ҫ��������Ĵ�����Ϣ�ķ���
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��--------------------------------���ں���(Window)------------------------------------��
'��                                                                                    ��
'
' Creating new windows:
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
' ��С��ָ���Ĵ��ڡ����ڲ�����ڴ������
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
' �ƻ����������ָ���Ĵ����Լ����������Ӵ���
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
' ��ָ���Ĵ�����������ֹ������꼰��������
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
' �ڴ����б���Ѱ����ָ����������ĵ�һ���Ӵ���
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
' �ж�ָ�����ڵĸ�����
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
' ָ��һ�����ڵ��¸�����vb��ʹ�ã��������������vb���Զ�����ʽ֧���Ӵ��ڡ�
' ���磬�ɽ��ؼ���һ���������������е���һ��������������ڴ�����ƶ��ؼ����൱ð�յģ�
' ��ȴ��ʧΪһ����Ч�İ취������������������ڹر��κ�һ������֮ǰ��ע����SetParent���ؼ��ĸ����ԭ�����Ǹ���
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
' ����ָ�����ڣ���ֹ�����¡�ͬʱֻ����һ�����ڴ�������״̬
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
' ǿ���������´��ڣ���������ǰ���ε��������򶼻��ػ�
' ��vb��ʹ�ã���vb�����ؼ����κβ�����Ҫ���£��ɿ���ֱ��ʹ��refresh����
Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
' �ж�һ�����ھ���Ƿ���Ч
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
' ���ƴ��ڵĿɼ���
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
' �ı�ָ�����ڵ�λ�úʹ�С���������ڿ�����������С�ߴ�����ƣ���Щ�ߴ��������������õĲ���
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
' ���������Ϊ����ָ��һ����λ�ú�״̬����Ҳ�ɸı䴰�����ڲ������б��е�λ�á�
' �ú�����DeferWindowPos�������ƣ�ֻ�������������������ֳ�����
' ��vb��ʹ�ã����vb���壬��������win32�����λ���С���������������״̬��
' ���б�Ҫ������һ�����ദ��ģ�����������״̬)
' ����
' hwnd             ����λ�Ĵ���
' hWndInsertAfter  ���ھ�����ڴ����б��У�����hwnd������������ھ���ĺ��棬�ο���ģ��ö��KhanSetWindowPosStyles
' x                �����µ�x���ꡣ��hwnd��һ���Ӵ��ڣ���x�ø����ڵĿͻ��������ʾ
' y                �����µ�y���ꡣ��hwnd��һ���Ӵ��ڣ���y�ø����ڵĿͻ��������ʾ
' cx               ָ���µĴ��ڿ��
' cy               ָ���µĴ��ڸ߶�
' wFlags           ����������һ���������ο���ģ��ö��KhanSetWindowPosStyles
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
' ��ָ�����ڵĽṹ��ȡ����Ϣ��nIndex�����ο���ģ�鳣������
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
' �ڴ��ڽṹ��Ϊָ���Ĵ���������Ϣ��nIndex�����ο���ģ�鳣������
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��------------------------------�����ຯ��(Window Class)------------------------------��
'��                                                                                    ��
'
' Ϊָ���Ĵ���ȡ������
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��-----------------------------������뺯��(Mouse Input)------------------------------��
'
' ���һ�����ڵľ�����������λ�ڵ�ǰ�����̣߳���ӵ����겶������������գ�
Public Declare Function GetCapture Lib "user32" () As Long
' ����겶�����õ�ָ���Ĵ��ڡ�����갴ť���µ�ʱ��������ڻ�Ϊ��ǰӦ�ó��������ϵͳ���������������
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
' Ϊ��ǰ��Ӧ�ó����ͷ���겶��
Public Declare Function ReleaseCapture Lib "user32" () As Long
' ����ģ��һ������¼����������������˫�����Ҽ�������
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
' ��������ж�ָ���ĵ��Ƿ�λ�ھ���lpRect�ڲ�
'Public Declare Function PtInRect Lib "user32" (lpRect As RECT, pt As POINTAPI) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��-----------------------------�������뺯��(Mouse Input)------------------------------��
'
' ���ӵ�����뽹��Ĵ��ڵľ��
Public Declare Function GetFocus Lib "user32" () As Long
' ���뽹���赽ָ���Ĵ���
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��----------------����ռ���任����(Coordinate Space Transtormation)-----------------��
'
' �жϴ������Կͻ��������ʾ��һ�������Ļ����
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
' �ж���Ļ��һ��ָ����Ŀͻ�������
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��---------------------------�豸��������(Device Context)-----------------------------��
'
' ����һ�����ض��豸����һ�µ��ڴ��豸�������ڻ���֮ǰ����ҪΪ���豸����ѡ��һ��λͼ��
' ������Ҫʱ�����豸��������DeleteDC����ɾ����ɾ��ǰ�������ж���Ӧ�ظ���ʼ״̬
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
' Ϊר���豸�����豸����
Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
' ��ȡָ�����ڵ��豸�������ñ�������ȡ���豸����һ��Ҫ��ReleaseDC�����ͷţ�������DeleteDC
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
' �ͷ��ɵ���GetDC��GetWindowDC������ȡ��ָ���豸�������������˽���豸������Ч���������ĵ��ò�������𺦣�
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
' ɾ��ר���豸��������Ϣ�������ͷ�������ش�����Դ����Ҫ��������GetDC����ȡ�ص��豸����
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
' ÿ���豸������������ѡ�����е�ͼ�ζ������а���λͼ��ˢ�ӡ����塢�����Լ�����ȵȡ�
' һ��ѡ���豸������ֻ����һ������ѡ���Ķ�������豸�����Ļ�ͼ������ʹ�á�
' ���磬��ǰѡ���Ļ��ʾ��������豸�����������߶���ɫ����ʽ
' ����ֵͨ�����ڻ��ѡ��DC�Ķ����ԭʼֵ��
' ��ͼ������ɺ�ԭʼ�Ķ���ͨ��ѡ���豸�����������һ���豸����ǰ�����ע��ָ�ԭʼ�Ķ���
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
' ���������ɾ��GDI���󣬱��续�ʡ�ˢ�ӡ����塢λͼ�������Լ���ɫ��ȵȡ�����ʹ�õ�����ϵͳ��Դ���ᱻ�ͷ�
' ��Ҫɾ��һ����ѡ���豸�����Ļ��ʡ�ˢ�ӻ�λͼ����ɾ����λͼΪ��������Ӱ��ͼ����ˢ�ӣ�
' λͼ�������������ɾ������ֻ��ˢ�ӱ�ɾ��
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'����ָ���豸����������豸�Ĺ��ܷ�����Ϣ
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
' ȡ�ö�ָ���������˵����һ���ṹ
' lpObject �κ����ͣ��������ɶ������ݵĽṹ��
' ��Ի��ʣ�ͨ����һ��LOGPEN�ṹ�������չ���ʣ�ͨ����EXTLOGPEN��
' ���������LOGBRUSH�����λͼ��BITMAP�����DIBSectionλͼ��DIBSECTION��
' ��Ե�ɫ�壬Ӧָ��һ�����ͱ����������ɫ���е���Ŀ����
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
' �ڴ��ڣ����豸����������ˮƽ�ͣ��򣩴�ֱ��������
Public Declare Function ScrollDC Lib "user32" (ByVal hDC As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long
' �������������Ϊһ��������
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
' ����һ���ɵ�X1��Y1��X2��Y2�����ľ������򣬲���ʱһ��Ҫ��DeleteObject����ɾ��������
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
' ����һ����lpRectȷ���ľ�������
Public Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
' ����һ��Բ�Ǿ��Σ��þ�����X1��Y1-X2��Y2ȷ��������X3��Y3ȷ������Բ����Բ�ǻ���
' �øú�����������������RoundRect API��������Բ�Ǿ��β���ȫ��ͬ����Ϊ�����ε��ұߺ��±߲�����������֮��
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
' ��ָ��ˢ�����ָ������
Public Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
' ��ָ��ˢ��Χ��ָ������һ�����
Public Declare Function FrameRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetMapMode Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SetMapMode Lib "gdi32" (ByVal hDC As Long, ByVal nMapMode As Long) As Long
' ������Щ��������ע�⵽�ĶԱ������˵�Ǹ��޴�ı��ص�������API�����е�һ�����������������ı䴰�ڵ�����
' ͨ�����д��ڶ��Ǿ��εġ�������һ�����ھͺ���һ���������򡣱���������������������
' ����ζ�������Դ���Բ�ġ����εĴ��ڣ�Ҳ���Խ�����Ϊ��������ಿ�֡���ʵ���Ͽ������κ���״
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
' �ú���ѡ��һ��������Ϊָ���豸�����ĵ�ǰ��������
Public Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��---------------------------------λͼ����(Bitmap)-----------------------------------��
'
' �ú���������ʾ͸�����͸�����ص�λͼ��
Public Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal xDest As Long, ByVal yDest As Long, ByVal WidthDest As Long, ByVal HeightDest As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal Blendfunc As Long) As Long
' ��һ��λͼ��һ���豸�������Ƶ���һ����Դ��Ŀ��DC�໥��������
' ��NT�����£�����һ�����紫����Ҫ����Դ�豸�����н��м��л���ת�������������ִ�л�ʧ��
' ��Ŀ���ԴDC��ӳ���ϵҪ����������صĴ�С�����ڴ�������иı䣬
' ��ô��������������Ҫ�Զ���������ת���۵������жϣ��Ա�������յĴ������
' dwRop��ָ����դ�������롣��Щ���뽫����Դ�����������ɫ���ݣ������Ŀ������������ɫ������������������ɫ��
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
' ����һ�����豸�й�λͼ
Public Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long
' ����һ�����豸�й�λͼ������ָ�����豸��������
' �ڴ��豸���������ɫλͼ���ݣ�Ҳ�뵥ɫλͼ���ݡ���������������Ǵ���һ���뵱ǰѡ��hdc�еĳ������ݡ�
' ��һ���ڴ泡����˵��Ĭ�ϵ�λͼ�ǵ�ɫ�ġ������ڴ��豸������һ��DIBSectionѡ�����У�
' ��������ͻ᷵��DIBSection��һ���������hdc��һ���豸λͼ��
' ��ô������ɵ�λͼ�Ϳ϶��������豸��Ҳ����˵����ɫ�豸���ɵĿ϶��ǲ�ɫλͼ��
' ���nWidth��nHeightΪ�㣬���ص�λͼ����һ��1��1�ĵ�ɫλͼ
' һ��λͼ������Ҫ��һ����DeleteObject�����ͷ���ռ�õ��ڴ漰��Դ
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
' �ú��������豸�޹ص�λͼ��DIB���������豸�йص�λͼ��DDB����������ѡ���Ϊλͼ��λ��
Public Declare Function CreateDIBitmap Lib "gdi32" (ByVal hDC As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO, ByVal wUsage As Long) As Long
' �ú�������Ӧ�ó������ֱ��д��ġ����豸�޹ص�λͼ��DIB����
' �ú����ṩһ��ָ�룬��ָ��ָ��λͼλ����ֵ�ĵط���
' ���Ը��ļ�ӳ������ṩ���������ʹ���ļ�ӳ�����������λͼ��������ϵͳΪλͼ�����ڴ档
Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
' ����λͼ��ͼ���ָ�룬ͬʱ�ڸ��ƹ����н���һЩת������
Public Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
' ����һ��λͼ��ͼ���ָ��
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function LoadImageLong Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��----------------------------------ͼ�꺯��(Icon)------------------------------------��
'
' ����ָ��ͼ������ָ���һ��������������������ڷ������õ�Ӧ�ó���
Public Declare Function CopyIcon Lib "user32" (ByVal hIcon As Long) As Long
' ����һ��ͼ��
Public Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long
' �ú������ͼ����ͷ��κα�ͼ��ռ�õĴ洢�ռ䡣
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
' �ú������޶����豸�����Ĵ��ڵĿͻ��������ͼ��
Public Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
' �ú������޶����豸�����Ĵ��ڵĿͻ��������ͼ�ִ꣬���޶��Ĺ�դ�����������ض�Ҫ���쳤��ѹ��ͼ����ꡣ
Public Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean
' ȡ����ͼ���йص���Ϣ
Public Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��---------------------------------��꺯��(Cursor)-----------------------------------��
'
Public Declare Function CopyCursor Lib "user32" (ByVal hcur As Long) As Long
' ��ָ����ģ���Ӧ�ó���ʵ��������һ�����ָ�롣LoadCursorBynum��LoadCursor���������Ͱ�ȫ����
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
' �ú�������һ����겢�ͷ���ռ�õ��κ��ڴ棬��Ҫʹ�øú���ȥ����һ�������ꡣ
Public Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
' ��ȡ���ָ��ĵ�ǰλ��
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
' �ú����ѹ���Ƶ���Ļ��ָ��λ�á������λ�ò����� ClipCursor�������õ���Ļ��������֮�ڣ�
' ��ϵͳ�Զ��������꣬ʹ�ù���ھ���֮�ڡ�
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��-----------------------------��ˢ����(Pen and Brush)---------------------------------��
'
' ��ָ������ʽ����Ⱥ���ɫ����һ�����ʣ���DeleteObject��������ɾ��
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
' ����ָ����LOGPEN�ṹ����һ������
Public Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long
' ����һ����չ���ʣ�װ�λ򼸺Σ�
Public Declare Function ExtCreatePen Lib "gdi32" (ByVal dwPenStyle As Long, ByVal dwWidth As Long, lplb As LOGBRUSH, ByVal dwStyleCount As Long, lpStyle As Long) As Long
' ��һ��LOGBRUSH���ݽṹ�Ļ����ϴ���һ��ˢ��
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
' �ú������Դ���һ������ָ����Ӱģʽ����ɫ���߼�ˢ�ӡ�
Public Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
' �ú������Դ�������ָ��λͼģʽ���߼�ˢ�ӣ���λͼ������DIB���͵�λͼ��DIBλͼ����CreateDIBSection���������ġ�
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
' �ô�ɫ����һ��ˢ�ӣ�һ��ˢ�Ӳ�����Ҫ������DeleteObject��������ɾ��
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
' Ϊ�κ�һ�ֱ�׼ϵͳ��ɫȡ��һ��ˢ�ӣ���Ҫ��DeleteObject����ɾ����Щˢ�ӡ�
' ��������ϵͳӵ�еĹ��ж��󡣲�Ҫ����Щˢ��ָ����һ�ִ������Ĭ��ˢ��
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��---------------------------��������ĺ���(Font and Text)-----------------------------��
'
' ��ָ�������Դ���һ���߼����壬VB������������ѡ�������ʱ���Եø���Ч
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
' ���ı���浽ָ���ľ����У�wFormat��־�����ο�KhanDrawTextStyles
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
' �ú���ȡ��ָ���豸�����ĵ�ǰ������ɫ��
Public Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
' ���õ�ǰ�ı���ɫ��������ɫҲ��Ϊ��ǰ��ɫ������ı���������ã�ע��ָ�VB�����ؼ�ԭʼ���ı���ɫ
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��------------------------------------��ͼ����----------------------------------------��
'
' �ú�����һ��Բ����Բ������һ����Բ��һ���߶Σ���֮Ϊ���ߣ��ཻ�޶��ıպ�����
' �˻��ɵ�ǰ�Ļ��ʻ��������ɵ�ǰ�Ļ�ˢ��䡣
Public Declare Function Chord Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
' ��ָ������ʽ���һ�����εı߿������������������û�б�Ҫ��ʹ�����3D�߿����塣
' ���Ծ���Դ���ڴ��ռ������˵�����������Ч��Ҫ�ߵöࡣ������һ���̶�����������
' hdc      Ҫ�����л�ͼ���豸����
' qrc      ҪΪ�����߿�ľ���
' edge     ����ǰ׺BDR_��������������ϡ�һ��ָ���ڲ��߿�����͹�����°�����һ����ָ���ⲿ�߿���ʱ�ܻ��ô�EDGE_ǰ׺�ĳ�����
' grfFlags ����BF_ǰ׺�ĳ��������
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
' ��һ��������Ρ�����������ڱ�־�������ʽ��ͨ�����������ɵģ�����ͨ����һ�����߱�ʾ��
' ����ͬ���Ĳ����ٴε�������������ͱ�ʾɾ���������
Public Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
' ��������������һ����׼�ؼ�
Public Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
' ���������Ϊһ��ͼ����ͼ����Ӧ�ø�ʽ������Ч��
Public Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As Long
' �ú������ڻ�һ����Բ����Բ���������޶����ε����ģ�ʹ�õ�ǰ���ʻ���Բ���õ�ǰ�Ļ�ˢ�����Բ��
Public Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
' ��ָ����ˢ�����һ�����Σ����ε��ұߺ͵ױ߲������
Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
' ��ָ����ˢ��Χ��һ�����λ�һ���߿����һ��֡�����߿�Ŀ����һ���߼���λ
Public Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
' ȡ��ָ���豸������ǰ�ı�����ɫ
Public Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
' ���ָ�����豸������ȡ�õ�ǰ�ı������ģʽ
Public Declare Function GetBkMode Lib "gdi32" (ByVal hDC As Long) As Long
' Ϊָ�����豸�������ñ�����ɫ��������ɫ���������Ӱˢ�ӡ����߻����Լ��ַ����米��ģʽΪOPAQUE���еĿ�϶��
' Ҳ��λͼ��ɫת���ڼ�ʹ�á�����ʵ�����豸�ܹ���ʾ����ӽ��� crColor ����ɫ
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
' ָ����Ӱˢ�ӡ����߻����Լ��ַ��еĿ�϶����䷽ʽ������ģʽ����Ӱ������չ������������
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
' ��ָ�����豸������ȡ��һ�����ص�RGBֵ
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
' ��ָ�����豸����������һ�����ص�RGBֵ
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
' ������һ��λͼ�Ķ�����λ���Ƶ�һ�����豸�޹ص�λͼ��
'Public Declare Function GetDIBits Lib "gdi32" ( ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
' ���������豸�޹�λͼ�Ķ�����λ���Ƶ�һ�����豸�йص�λͼ��
Public Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
' ���ָ�����豸��������ö�������ģʽ��
Public Declare Function GetPolyFillMode Lib "gdi32" (ByVal hDC As Long) As Long
' ���ö���ε����ģʽ
Public Declare Function SetPolyFillMode Lib "gdi32" (ByVal hDC As Long, ByVal nPolyFillMode As Long) As Long
' ���ָ�����豸������ȡ�õ�ǰ�Ļ�ͼģʽ�������ɶ����ͼ���������������ʾ��ͼ��ϲ�����
' �������ֻ�Թ�դ�豸��Ч
Public Declare Function GetROP2 Lib "gdi32" (ByVal hDC As Long) As Long
' ����ָ���豸�����Ļ�ͼģʽ��
Public Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
' �õ�ǰ���ʻ�һ���ߣ��ӵ�ǰλ������һ��ָ���ĵ㡣�������������ϣ���ǰλ�ñ��x,y��
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
' Ϊָ�����豸����ָ��һ���µĵ�ǰ����λ�á�ǰһ��λ�ñ�����lpPoint��
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
' �ú�����һ������Բ�������뾶�ཻ�պ϶��ɵı�״Ш��ͼ���˱�ͼ�ɵ�ǰ���ʻ��������ɵ�ǰ��ˢ��䡣
Public Declare Function Pie Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
' �ú�����һ����ֱ�����ŵ��������϶�����ɵĶ���Σ��õ�ǰ���ʻ������������
' �õ�ǰ��ˢ�Ͷ�������ģʽ������Ρ�
Public Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
' �õ�ǰ�������һϵ���߶Ρ�ʹ��PolylineTo����ʱ����ǰλ�û���Ϊ���һ���߶ε��յ㡣
' ��������Polyline�����Ķ�
Public Declare Function Polyline Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function PolyPolygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long
Public Declare Function PolyPolyline Lib "gdi32" (ByVal hDC As Long, lppt As POINTAPI, lpdwPolyPoints As Long, ByVal cCount As Long) As Long
' �ú�����һ�����Σ��õ�ǰ�Ļ��ʻ������������õ�ǰ��ˢ������䡣
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
' ������һ����Բ�ǵľ��Σ��˾����ɵ�ǰ���ʻ����ȣ��ɵ�ǰ��ˢ��䡣
Public Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
' �����������������Сһ�����εĴ�С��
' x�����Ҳ����򣬲�����������ȥ����xΪ��������������εĿ�ȣ���xΪ�������ܼ�С����
' y�Զ�����ײ����������Ӱ���������Ƶ�
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
' �ú���ͨ��Ӧ��һ��ָ����ƫ�ƣ��Ӷ��þ����ƶ�������
' x����ӵ��Ҳ���������y��ӵ������͵ײ�����
' ƫ�Ʒ�����ȡ���ڲ������������Ǹ������Լ����õ���ʲô����ϵͳ
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
' ������windows�����йص���Ϣ��nIndexֵ�ο���ģ��ĳ�������
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
' ����������ڵķ�Χ���Σ����ڵı߿򡢱����������������˵��ȶ������������
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
' ����ָ�����ڿͻ������εĴ�С
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
' �����������һ�����ڿͻ�����ȫ���򲿷�������ᵼ�´������¼��ڼ䲿���ػ�
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
' �ж�ָ��windows��ʾ�������ɫ����ɫ���󿴱�ģ������
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

'��                                                                                    ��
'����������������������������������������������������������������������������������������

'����������������������������������������������������������������������������������������
'��--------------------------------��������(Others)------------------------------------��
'
' ����λͼ��ͼ���ָ�룬ͬʱ�ڸ��ƹ����н���һЩת������
' �������ͨ����ϣ��������ѡ�������豸������һ��λͼʱʹ��
' ���磬�����ѳ�ΪImageList�ؼ�һ���ֵ�ĳ��λͼ��ѡ����λͼ������ʹ�ã���Ϊһ��ֻ�ܽ�λͼѡ��һ���豸����
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long

'Initializes the entire common control dynamic-link library.
'Exported by all versions of Comctl32.dll.
Public Declare Sub InitCommonControls Lib "comctl32" ()
'Initializes specific common controls classes from the common
'control dynamic-link library.
'Returns TRUE (non-zero) if successful, or FALSE otherwise.
'Began being exported with Comctl32.dll version 4.7 (IE3.0 & later).
Public Declare Function InitCommonControlsEx Lib "comctl32" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean
Public Declare Function ImageList_GetBkColor Lib "comctl32" (ByVal hImageList As Long) As Long
Public Declare Function ImageList_ReplaceIcon Lib "comctl32" (ByVal hImageList As Long, ByVal i As Long, ByVal hIcon As Long) As Long
Public Declare Function ImageList_Convert Lib "comctl32" Alias "ImageList_Draw" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hdcDest As Long, ByVal x As Long, ByVal y As Long, ByVal Flags As Long) As Long
Public Declare Function ImageList_Create Lib "comctl32" (ByVal MinCx As Long, ByVal MinCy As Long, ByVal Flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Public Declare Function ImageList_AddMasked Lib "comctl32" (ByVal hImageList As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Public Declare Function ImageList_Replace Lib "comctl32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal hbmImage As Long, ByVal hBmMask As Long) As Long
Public Declare Function ImageList_Add Lib "comctl32" (ByVal hImageList As Long, ByVal hbmImage As Long, hBmMask As Long) As Long
Public Declare Function ImageList_Remove Lib "comctl32" (ByVal hImageList As Long, ByVal ImgIndex As Long) As Long
Public Declare Function ImageList_GetImageInfo Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, pImageInfo As IMAGEINFO) As Long
Public Declare Function ImageList_AddIcon Lib "comctl32" (ByVal hIml As Long, ByVal hIcon As Long) As Long
Public Declare Function ImageList_GetIcon Lib "comctl32" (ByVal hImageList As Long, ByVal ImgIndex As Long, ByVal fuFlags As Long) As Long
Public Declare Function ImageList_SetImageCount Lib "comctl32" (ByVal hImageList As Long, uNewCount As Long)
Public Declare Function ImageList_GetImageCount Lib "comctl32" (ByVal hImageList As Long) As Long
Public Declare Function ImageList_Destroy Lib "comctl32" (ByVal hImageList As Long) As Long
Public Declare Function ImageList_GetIconSize Lib "comctl32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long
Public Declare Function ImageList_SetIconSize Lib "comctl32" (ByVal hImageList As Long, cx As Long, cy As Long) As Long
Public Declare Function ImageList_Draw Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal fStyle As Long) As Long
' Draw an item in an ImageList with more control over positioning and colour:
Public Declare Function ImageList_DrawEx Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal rgbBk As Long, ByVal rgbFg As Long, ByVal fStyle As Long) As Long
Public Declare Function ImageList_GetImageRect Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, prcImage As RECT) As Long
Public Declare Function ImageList_LoadImage Lib "comctl32" Alias "ImageList_LoadImageA" (ByVal hInst As Long, ByVal lpbmp As String, ByVal cx As Long, ByVal cGrow As Long, ByVal crMask As Long, ByVal uType As Long, ByVal uFlags As Long)
Public Declare Function ImageList_SetBkColor Lib "comctl32" (ByVal hImageList As Long, ByVal clrBk As Long) As Long
Public Declare Function ImageList_Copy Lib "comctl32" (ByVal himlDst As Long, ByVal iDst As Long, ByVal himlSrc As Long, ByVal iSrc As Long, ByVal uFlags As Long) As Long
'��                                                                                    ��
'����������������������������������������������������������������������������������������

