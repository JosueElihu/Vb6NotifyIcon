VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNotifyIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================================================================================================
'   Component    : cNotifyIcon 1.5
'   Autor        : J. Elihu
'----------------------------------------------------------------------------------------------------------------------
'   Description  : ShellNotifyIcon VERSION_4 Compatible
'   Credits      : Leandro Ascierto (ClsSysTray, ClsPopUpTray)
'======================================================================================================================
Option Explicit


'-= ModernSubclassing =-
#Const ImplNoIdeProtection = (MST_NO_IDE_PROTECTION <> 0)
#Const ImplSelfContained = True
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetProcByOrdinal Lib "kernel32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcOrdinal As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" Alias "#413" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#If Not ImplNoIdeProtection Then
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
#End If
#If ImplSelfContained Then
Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
#End If


Private Type Guid
  Data1     As Long
  Data2     As Integer
  Data3     As Integer
  Data4(7)  As Byte
End Type

Private Type RECT
  Left      As Long
  Top       As Long
  Right     As Long
  Bottom    As Long
End Type

Private Type POINTAPI
  x         As Long
  y         As Long
End Type

'-= Ole32 =-
Private Declare Sub CoCreateGuid Lib "ole32" (ByRef pguid As Guid)
Private Declare Function StringFromGUID2 Lib "ole32" (ByVal rguid As Long, ByVal lpsz As Long, ByVal cchMax As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal Str As Long, id As Guid) As Long

'-= User32 =-
Private Declare Function CreateWindowExA Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CopyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function CreateIconFromResourceEx Lib "user32" (ByRef presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
Private Declare Function LoadImageW Lib "user32" (ByVal hInst As Long, ByVal lpsz As Long, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function RegisterWindowMessageA Lib "user32" (ByVal lpString As String) As Long
Private Declare Function CalculatePopupWindowPosition Lib "user32" (anchorPoint As POINTAPI, windowSize As POINTAPI, ByVal Flags As Long, excludeRect As RECT, PopupWindowPosition As RECT) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function WaitMessage Lib "user32" () As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Const HWND_TOPMOST      As Long = -1
Private Const SWP_NOMOVE        As Long = &H2
Private Const SWP_NOSIZE        As Long = &H1
Private Const SWP_SHOWWINDOW    As Long = &H40
Private Const SWP_FRAMECHANGED  As Long = &H20
Private Const SWP_HIDEWINDOW    As Long = &H80
Private Const SWP_NOACTIVATE    As Long = &H10
Private Const GWL_EXSTYLE       As Long = -20
Private Const GWL_STYLE         As Long = -16

Private Const WS_POPUP          As Long = &H80000000
Private Const WS_THICKFRAME     As Long = &H40000
Private Const WS_BORDER         As Long = &H800000
Private Const WS_MAXIMIZEBOX    As Long = &H10000
Private Const WS_MINIMIZEBOX    As Long = &H20000
Private Const WS_CAPTION        As Long = &HC00000

Private Const WS_EX_TOOLWINDOW  As Long = 128
Private Const WS_EX_WINDOWEDGE  As Long = &H100&

'-= Shlwapi =-
Private Declare Function PathFileExistsW Lib "shlwapi" (ByVal pszPath As Long) As Long

'-= DPI =-
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

'-= NotifyIcon =-
Private Declare Function Shell_NotifyIconA Lib "Shell32" (ByVal dwMessage As Long, ByRef lpData As Any) As Long
Private Declare Function Shell_NotifyIconGetRect Lib "Shell32" (ByRef Identifier As NOTIFYICONIDENTIFIER, ByRef iconLocation As Any) As Long

Private Type NOTIFYICONDATAA
  cbSize            As Long
  hWnd              As Long
  uId               As Long
  uFlags            As Long
  uCallbackMessage  As Long
  hIcon             As Long
  szTip             As String * 128
  dwState           As Long
  dwStateMask       As Long
  szInfo            As String * 256
  uVersion          As Long
  szInfoTitle       As String * 64
  dwInfoFlags       As Long
  guidItem          As Guid     'lGuid(3) As Long
  hBalloonIcon      As Long
End Type

Private Type NOTIFYICONIDENTIFIER
  cbSize            As Long
  hWnd              As Long
  uId               As Long
  guidItem          As Guid     'lGuid(3) As Long
End Type

'-= Windows messages =-
Private Const WM_APP                As Long = &H8000&
Private Const WM_DESTROY            As Long = &H2
Private Const WM_MOUSEMOVE          As Long = &H200
Private Const WM_RBUTTONDBLCLK      As Long = &H206
Private Const WM_RBUTTONDOWN        As Long = &H204
Private Const WM_RBUTTONUP          As Long = &H205
Private Const WM_MBUTTONDBLCLK      As Long = &H209
Private Const WM_MBUTTONDOWN        As Long = &H207
Private Const WM_MBUTTONUP          As Long = &H208
Private Const WM_LBUTTONDBLCLK      As Long = &H203
Private Const WM_LBUTTONDOWN        As Long = &H201
Private Const WM_LBUTTONUP          As Long = &H202
Private Const WM_USER               As Long = &H400
Private Const WM_NCHITTEST          As Long = &H84
Private Const WM_CONTEXTMENU        As Long = &H7B&
Private Const WM_KILLFOCUS          As Long = &H8
Private Const WM_SIZE               As Long = &H5
Private Const WM_ACTIVATE           As Long = &H6&
Private Const WMAPP_NOTIFYCALLBACK  As Long = WM_APP + 1


'-= Callback Messages =-
Private Const NIN_SELECT            As Long = WM_USER + 0
Private Const NINF_KEY              As Long = &H1
Private Const NIN_BALLOONSHOW       As Long = (WM_USER + 2)
Private Const NIN_BALLOONHIDE       As Long = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT    As Long = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK  As Long = (WM_USER + 5)
Private Const NIN_POPUPOPEN         As Long = WM_USER + 6
Private Const NIN_POPUPCLOSE        As Long = WM_USER + 7
Private Const NIN_KEYSELECT         As Long = NIN_SELECT Or NINF_KEY


'-= Update commands =-
Private Const NIM_ADD               As Long = &H0
Private Const NIM_MODIFY            As Long = &H1
Private Const NIM_DELETE            As Long = &H2
Private Const NIM_SETFOCUS          As Long = &H3
Private Const NIM_SETVERSION        As Long = &H4


'-= Update masks =-
Private Const NIF_ICON              As Long = &H2
Private Const NIF_INFO              As Long = &H10
Private Const NIF_MESSAGE           As Long = &H1
Private Const NIF_STATE             As Long = &H8
Private Const NIF_TIP               As Long = &H4
Private Const NIF_GUID              As Long = &H20
Private Const NIIF_LARGE_ICON       As Long = &H20
Private Const NIF_REALTIME          As Long = &H40
Private Const NIF_SHOWTIP           As Long = &H80
Private Const NIIF_NOSOUND          As Long = &H10
Private Const NIIF_RESPECT_QUIET_TIME As Long = &H80

'-= Version =-
Private Const NOTIFYICON_VERSION    As Long = 3
Private Const NOTIFYICON_VERSION_4  As Long = 4

Public Enum NIIF_ICON_TYPE
  [NIIF_NONE] = 0
  [NIIF_INFO] = 1
  [NIIF_WARNING] = 2
  [NIIF_ERROR] = 3
  [NIIF_USER] = 4
End Enum


'-= Events =-
Event Click()
Event MouseDown(ByVal Button As Integer)
Event MouseUp(ByVal Button As Integer)
Event DblClick(ByVal Button As Integer)
Event ContextMenu(ByVal x As Single, ByVal y As Single)
Event BalloonShow()
Event BalloonClick()
Event BallonTimeOut()
Event PopupOpen(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
Event PopupClose()


Private WM_TASKBAR      As Long

Private m_hWnd          As Long
Private m_pSubclass     As IUnknown
Private m_pScFlyout     As IUnknown
Private m_UseGuid       As Boolean
Private m_Guid          As Guid
Private m_StandardTt    As Boolean
Private m_bShow         As Boolean
Private m_hIcon         As Long
Private m_ToolTip       As String
Private m_bInPopup      As Boolean

'/* Flyout Window */
Private m_hWndPopup     As Long
Private m_Tmlast        As Long

Private Sub Class_Initialize()

    WM_TASKBAR = RegisterWindowMessageA("TaskbarCreated")
    m_hWnd = CreateWindowExA(0, "STATIC", vbNullString, 0, 0, 0, 0, 0, 0, 0, App.hInstance, 0)
    If m_hWnd Then
        Set m_pSubclass = InitSubclassingThunk(m_hWnd, Me, InitAddressOfMethod(Me, 5).WinProc(0, 0, 0, 0, 0))
    End If
End Sub


Private Sub Class_Terminate()

    If m_hIcon Then Call DestroyIcon(m_hIcon): m_hIcon = 0
    If m_bShow Then Call RemoveIcon
    
    If m_hWnd Then
        TerminateSubclassingThunk m_pSubclass, Me
        Call DestroyWindow(m_hWnd)
        m_hWnd = 0
    End If
End Sub


Public Function ShowIcon(Optional Icon As Variant, Optional ByVal Tooltip As String, Optional ShowTooltip As Boolean = True) As Boolean
Dim nid As NOTIFYICONDATAA
    
    If m_bShow Then Call RemoveIcon
    
    nid = ssNid(NIF_ICON Or NIF_MESSAGE Or NIF_TIP) '/* With POPUPOPEN */
    'nid = ssNid(NIF_ICON Or NIF_MESSAGE)        '/* With POPUPOPEN */
    
    nid.hWnd = m_hWnd
    nid.uCallbackMessage = WMAPP_NOTIFYCALLBACK
    If Not IsMissing(Icon) Then nid.hIcon = pvIcon(Icon, True)
    
    nid.szTip = Tooltip & vbNullChar
    If ShowTooltip Then nid.uFlags = nid.uFlags Or NIF_SHOWTIP
    m_StandardTt = ShowTooltip
    
    ShowIcon = Shell_NotifyIconA(NIM_ADD, nid)
    
    '/* NOTIFYICON_VERSION_4 Is Prefered */
    nid.uVersion = NOTIFYICON_VERSION_4
    Call Shell_NotifyIconA(NIM_SETVERSION, nid)
    
    m_bShow = ShowIcon
    
    '/* Save data to use with WM_TASKBARMESSAGE */
    If m_hIcon Then Call DestroyIcon(m_hIcon): m_hIcon = 0
    If nid.hIcon Then m_hIcon = CopyIcon(nid.hIcon)
    m_ToolTip = Tooltip
    
End Function
Public Function RemoveIcon() As Boolean
    If Not m_bShow Then Exit Function
    RemoveIcon = Shell_NotifyIconA(NIM_DELETE, ssNid)
    m_bShow = Not RemoveIcon
End Function

Public Function ShowBallon(ByVal Text As String, ByVal Title As String, Optional ByVal IconType As NIIF_ICON_TYPE, Optional BallonIcon As Variant, _
                           Optional ByVal NoSound As Boolean, Optional ByVal RespectQuietTime As Boolean = True, Optional ByVal RealTime As Boolean) As Boolean
                           
Dim nid     As NOTIFYICONDATAA

    nid = ssNid(NIF_INFO)
    nid.dwInfoFlags = IconType
    
    If Not IsMissing(BallonIcon) Then nid.hBalloonIcon = pvIcon(BallonIcon, False)
    If nid.hBalloonIcon Then nid.dwInfoFlags = NIIF_USER Or NIIF_LARGE_ICON
     
    If RespectQuietTime Then nid.dwInfoFlags = nid.dwInfoFlags Or NIIF_RESPECT_QUIET_TIME
    If NoSound Then nid.dwInfoFlags = nid.dwInfoFlags Or NIIF_NOSOUND
    
    If RealTime Then nid.uFlags = nid.uFlags Or NIF_REALTIME
    If m_StandardTt Then nid.uFlags = nid.uFlags Or NIF_SHOWTIP
    
    nid.szInfoTitle = Title & vbNullChar
    nid.szInfo = Text & vbNullChar
    
    ShowBallon = Shell_NotifyIconA(NIM_MODIFY, nid)
    If nid.hBalloonIcon Then Call DestroyIcon(nid.hBalloonIcon)
    
End Function

Public Function HideBallon() As Boolean
Dim nid     As NOTIFYICONDATAA
    
    nid = ssNid(NIF_INFO)
    nid.szInfoTitle = vbNullChar
    nid.szInfo = vbNullChar
    HideBallon = Shell_NotifyIconA(NIM_MODIFY, nid)
    
End Function
Public Function SetFocus() As Boolean
    SetFocus = (Shell_NotifyIconA(NIM_SETFOCUS, ssNid) = 0)
End Function
Public Sub BeforePopup()
    Call SetForegroundWindow(m_hWnd)
End Sub

Public Function ShowFlyout(Win As Variant, Optional ByVal UnloadIfVbForm As Boolean = True, Optional ChangeStyle As Boolean = True) As Boolean
On Error GoTo e
Dim hWin    As Long

    If m_Tmlast + 500 > GetTickCount Then
        pvUnsubclassPopup
        ShowWindow m_hWndPopup, 0
        m_hWndPopup = 0
        Exit Function
    End If
    
    Select Case VarType(Win)
        Case vbObject: hWin = Win.hWnd
        Case vbLong: hWin = Win
    End Select
    
    Dim lx As Long, ly As Long
    
    m_Tmlast = Time
    
    '// Display the flyout in an appropriate position
    If CalculateFlyout(hWin, lx, ly) Then
        
        '// Position the flyout and make it the foreground window
        If ChangeStyle Then
            SetWindowLongA hWin, GWL_STYLE, WS_THICKFRAME Or WS_BORDER 'Or WS_POPUP
            SetWindowLongA hWin, GWL_EXSTYLE, WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW
        End If
        
        SetWindowPos hWin, HWND_TOPMOST, lx, ly, 0, 0, SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_FRAMECHANGED Or SWP_NOACTIVATE
        
        m_hWndPopup = hWin
        
        pvSubclassPopup hWin
        SetForegroundWindow (hWin)
        
        Do While GetForegroundWindow = hWin
            DoEvents
            WaitMessage
        Loop
        
        ShowWindow hWin, 0
        pvUnsubclassPopup
        
        m_hWndPopup = 0
        
        m_hWndPopup = GetTickCount
    End If
    
e:
    On Error Resume Next
    If VarType(Win) = vbObject Then
        If pvIsVBForm(Win) And UnloadIfVbForm Then Unload Win
    End If
    
End Function

Public Sub CancelFlyout()
    If IsWindow(m_hWndPopup) Then DestroyWindow m_hWndPopup
End Sub


Public Function CalculateFlyout(hWnd As Long, Left As Long, Top As Long, Optional ByVal Flags As Long) As Boolean
Dim Point  As POINTAPI
Dim Size   As POINTAPI
Dim Rct    As RECT
Dim WinRct As RECT

    '/* Flags */
    'Const TPM_VERTICAL = &H40
    'Const TPM_VCENTERALIGN = &H10&
    'Const TPM_CENTERALIGN = &H4&
    'Const TPM_WORKAREA = &H10000
    
    If Flags = 0 Then Flags = &H40 Or &H10& Or &H4& Or &H10000
    
    '/* PositionFlyout */
    Select Case Shell_NotifyIconGetRect(ssNii, Rct)
        Case 0, 1
        
            Point.x = ((Rct.Left + Rct.Right) / 2) / mvScreenDpi
            Point.y = ((Rct.Top + Rct.Bottom) / 2) / mvScreenDpi
            
            GetWindowRect hWnd, WinRct
            Size.x = (WinRct.Right - WinRct.Left)
            Size.y = (WinRct.Bottom - WinRct.Top)
            
            SetRect Rct, Rct.Left / mvScreenDpi, Rct.Top / mvScreenDpi, Rct.Right / mvScreenDpi, Rct.Bottom / mvScreenDpi
            If CalculatePopupWindowPosition(Point, Size, Flags, Rct, WinRct) Then
                Left = WinRct.Left: Top = WinRct.Top
                CalculateFlyout = True
            End If
            
        Case Else
    End Select

End Function

Public Function GetIconRect(ByRef x1 As Long, ByRef y1 As Long, ByRef x2 As Long, ByRef y2 As Long) As Boolean
Dim Rct As RECT
    Select Case Shell_NotifyIconGetRect(ssNii, Rct)
        Case 0, 1
            SetRect Rct, Rct.Left / mvScreenDpi, Rct.Top / mvScreenDpi, Rct.Right / mvScreenDpi, Rct.Bottom / mvScreenDpi
            x1 = Rct.Left: y1 = Rct.Top
            x2 = Rct.Right: y2 = Rct.Bottom
            GetIconRect = True
        Case Else
    End Select
End Function

Property Get Guid() As String
    If m_UseGuid Then Guid = pvGuidToStr(m_Guid) Else Guid = "{00000000-0000-0000-0000-000000000000}"
End Property
Property Let Guid(ByVal Value As String)
Dim mShow   As Boolean
Dim mFlag   As Boolean

    If Value <> vbNullString Then
        
        If pvGuidToStr(m_Guid) = Value Then Exit Property
        mFlag = (CLSIDFromString(StrPtr(Value), m_Guid) = 0)
   
        If Not mFlag Then Debug.Print "ERROR: " & TypeName(Me) & ", Invalid GUID Property": Exit Property
        If mFlag Then If (pvGuidToStr(m_Guid) = "{00000000-0000-0000-0000-000000000000}") Then Debug.Print "ERROR: " & TypeName(Me) & ", Empty GUID Property": Exit Property
        
    Else
        mFlag = False '/* UID */
    End If
    
    If m_UseGuid = mFlag Then Exit Property
    
    mShow = m_bShow
    If mShow Then Call RemoveIcon
    
    m_UseGuid = mFlag
    If mShow Then pvReShow
    
End Property


Property Get InPopupMenu() As Boolean: InPopupMenu = m_bInPopup: End Property
Property Let Icon(Value As Variant)
Dim nid As NOTIFYICONDATAA

    If m_bShow = False Then Exit Property
    nid = ssNid(NIF_ICON)
    nid.hIcon = pvIcon(Value, True)
    
    If m_StandardTt Then nid.uFlags = nid.uFlags Or NIF_SHOWTIP
    If nid.hIcon Then Shell_NotifyIconA NIM_MODIFY, nid
    
End Property
Property Let Tooltip(ByVal Value As String)
Dim nid As NOTIFYICONDATAA

    If m_bShow = False Then Exit Property
    m_ToolTip = Value
    
    If m_StandardTt Then
        nid = ssNid(NIF_TIP Or NIF_SHOWTIP)
    Else
        nid = ssNid(NIF_TIP)
    End If
    
    nid.szTip = Value & vbNullChar
    
    Shell_NotifyIconA NIM_MODIFY, nid
    
End Property

Property Let ShowTooltip(ByVal Value As Boolean)
Dim nid As NOTIFYICONDATAA

    If m_bShow = False Then Exit Property
    m_StandardTt = Value
    If m_StandardTt Then
        nid = ssNid(NIF_TIP Or NIF_SHOWTIP)
    Else
        nid = ssNid(NIF_TIP)
    End If
    
    nid.szTip = m_ToolTip & vbNullChar
    Shell_NotifyIconA NIM_MODIFY, nid
    
End Property

'TODO: Private Subs
'=====================================================================================================================
Private Sub pvSubclassPopup(hWin As Long)
    Set m_pScFlyout = InitSubclassingThunk(hWin, Me, InitAddressOfMethod(Me, 5).WinProcpp(0, 0, 0, 0, 0))
End Sub
Private Sub pvUnsubclassPopup()
    TerminateSubclassingThunk m_pScFlyout, Me
End Sub

Private Function ssNid(Optional uFlags As Long) As NOTIFYICONDATAA
    ssNid.cbSize = Len(ssNid)
    If m_UseGuid Then
        '/* Icon will be identified with the GUID */
        ssNid.guidItem = m_Guid
        ssNid.uFlags = uFlags Or NIF_GUID
    Else
        '/* UID */
        ssNid.uFlags = uFlags
        ssNid.uId = ObjPtr(Me)
        ssNid.hWnd = m_hWnd
    End If
End Function
Private Function ssNii() As NOTIFYICONIDENTIFIER
    ssNii.cbSize = Len(ssNii)
    
    If m_UseGuid Then
    
        '/* Icon will be identified with the GUID */
        ssNii.guidItem = m_Guid
            
    Else
        '/* UID */
        ssNii.uId = ObjPtr(Me)
        ssNii.hWnd = m_hWnd
        
    End If
End Function

Private Function pvReShow() As Boolean
Dim PrevIcon As Long

    If m_hIcon Then
        PrevIcon = CopyIcon(m_hIcon)
        If DestroyIcon(m_hIcon) <> 0 Then m_hIcon = 0
    End If
    
    ShowIcon PrevIcon, m_ToolTip
    If PrevIcon Then Call DestroyIcon(PrevIcon)
        
End Function
Private Function pvGuidToStr(mguid As Guid) As String
Dim ab() As Byte
Dim lm   As Long
    
    ReDim ab(80)
    lm = StringFromGUID2(VarPtr(mguid.Data1), VarPtr(ab(0)), UBound(ab))
    pvGuidToStr = Left(ab, lm - 1)
    
End Function
Private Function pvIcon(Source As Variant, Optional Small As Boolean) As Long
On Error GoTo e

    Select Case VarType(Source)
        Case vbString
            Const SM_CXSMICON       As Long = 49
            Const SM_CYSMICON       As Long = 50
            
            Dim cx As Long: Dim cy As Long
            cx = GetSystemMetrics(SM_CXSMICON)
            cy = GetSystemMetrics(SM_CYSMICON)
            
            If PathFileExistsW(StrPtr(Source)) Then
                If Small Then
                    pvIcon = LoadImageW(0&, StrPtr(Source), 1&, cx, cy, &H10)
                Else
                    pvIcon = LoadImageW(0&, StrPtr(Source), 1&, 0, 0, &H10)
                End If
            Else
                If IsNumeric(Source) And Left(Source, 1) <> "#" Then Source = "#" & Source
                If Small Then
                    pvIcon = LoadImageW(App.hInstance, StrPtr(Source), 1, cx, cy, &H8000& Or &H1000)
                Else
                    pvIcon = LoadImageW(App.hInstance, StrPtr(Source), 1, 0, 0, &H8000& Or &H1000)
                End If
            End If
        Case (vbArray + vbByte)
            Dim Out() As Byte
            Out() = Source
            pvIcon = CreateIconFromResourceEx(Out(LBound(Out)), UBound(Out) - LBound(Out) + 1, 1&, &H30000, 0&, 0&, 0&)
            
        Case (vbArray + vbVariant) '/* Image From Resource {id,type} */
        
            If IsNumeric(Source(0)) Then Out() = LoadResData(CLng(Source(0)), Source(1)) Else Out() = LoadResData(CStr(Source(0)), Source(1))
            pvIcon = CreateIconFromResourceEx(Out(LBound(Out)), UBound(Out) - LBound(Out) + 1, 1&, &H30000, 0&, 0&, 0&)
            
        Case vbLong
            pvIcon = CopyIcon(Source)
    End Select
e:
End Function
Private Function pvIsVBForm(Obj As Variant) As Boolean
Dim Frm As Form

    For Each Frm In Forms
        'If TypeName(Frm) = TypeName(Obj) Then pvIsVBForm = True: Exit Function
        If Obj Is Frm Then pvIsVBForm = True: Exit Function
    Next
End Function



'TODO: ModernSubclassing
'=====================================================================================================================
Private Function InitAddressOfMethod(pObj As Object, ByVal MethodParamCount As Long) As cNotifyIcon
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFV4v6ge9QEMEAgcekEcEAuP9EJAS5+QcAAPOri8LB4AgFuQAAAKuLwsHoGAUAjYEAq7gIAAArq7hEJASLq7hJCIsEq7iBi1Qkq4tEJAzB4AIFCIkCM6uLRCQMweASBcDCCACriTrHQgQBAAAAi0QkCIsAiUIIi0QkEIlCDIHqUBDBAIvCBTwRwQCri8IFUBHBAKuLwgVgEcEAq4vCBYQRwQCri8IFjBHBAKuLwgWUEcEAq4vCBZwRwQCri8IFpBHBALn5BwAAq4PABOL6i8dfgcJQEMEAi0wkEIkRK8LCEAAPHwCLVCQE/0IEi0QkDIkQM8DCDABmkItUJAT/QgSLQgTCBAAPHwCLVCQE/0oEi0IEg/gAfgPCBABZWotCDGgAgAAAagBSUf/gZpC4AUAAgMIIALgBQACAwhAAuAFAAIDCGAC4AUAAgMIkAA==" ' 25.3.2019 14:01:08
    Const THUNK_SIZE    As Long = 16728
    Dim hThunk          As Long
    Dim lSize           As Long
    
    hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
    If hThunk = 0 Then
        Exit Function
    End If
    lSize = CallWindowProc(hThunk, ObjPtr(pObj), MethodParamCount, GetProcAddress(GetModuleHandle("kernel32"), "VirtualFree"), VarPtr(InitAddressOfMethod))
    Debug.Assert lSize = THUNK_SIZE
End Function

Private Function InitSubclassingThunk(ByVal hWnd As Long, pObj As Object, ByVal pfnCallback As Long) As IUnknown
    Dim STR_THUNK       As String: STR_THUNK = "6AAAAABag+oFgepwEBAAV1aLdCQUg8YIgz4AdC+L+oHHKBIQAIvCBQwREACri8IFSBEQAKuLwgVYERAAq4vCBYAREACruQkAAADzpYHCKBIQAFJqHP9SEFqL+IvCq7gBAAAAqzPAq4tEJAyri3QkFKWlM8Crg+8cagBX/3IM/3cM/1IYi0QkGIk4Xl+4XBIQAC1wEBAAwhAADx8Ai0QkCIM4AHUqg3gEAHUkgXgIwAAAAHUbgXgMAAAARnUSi1QkBP9CBItEJAyJEDPAwgwAuAJAAIDCDACQi1QkBP9CBItCBMIEAA8fAItUJAT/SgSLQgR1GIsKUv9xDP9yDP9RHItUJASLClL/URQzwMIEAJBVi+yLVRj/QgT/QhiLQhg7QgR0b4tCEIXAdGiLCotBLIXAdDdS/9BaiUIIg/gBd1OFwHUJgX0MAwIAAHRGiwpS/1EwWoXAdTuLClJq8P9xJP9RKFqpAAAACHUoUjPAUFCNRCQEUI1EJARQ/3UU/3UQ/3UM/3UI/3IQ/1IUWVhahcl1E1KLCv91FP91EP91DP91CP9RIFr/ShhQUug4////WF3CGAAPHwA=" ' 9.6.2020 13:56:03
    Const THUNK_SIZE    As Long = 492
    Static hThunk       As Long
    Dim aParams(0 To 10) As Long
    Dim lSize           As Long
    
    aParams(0) = ObjPtr(pObj)
    aParams(1) = pfnCallback
    #If ImplSelfContained Then
        If hThunk = 0 Then
            hThunk = pvThunkGlobalData("InitSubclassingThunk")
        End If
    #End If
    If hThunk = 0 Then
        hThunk = pvThunkAllocate(STR_THUNK, THUNK_SIZE)
        If hThunk = 0 Then
            Exit Function
        End If
        aParams(2) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemAlloc")
        aParams(3) = GetProcAddress(GetModuleHandle("ole32"), "CoTaskMemFree")
        Call DefSubclassProc(0, 0, 0, 0)                                            '--- load comctl32
        aParams(4) = GetProcByOrdinal(GetModuleHandle("comctl32"), 410)             '--- 410 = SetWindowSubclass ordinal
        aParams(5) = GetProcByOrdinal(GetModuleHandle("comctl32"), 412)             '--- 412 = RemoveWindowSubclass ordinal
        aParams(6) = GetProcByOrdinal(GetModuleHandle("comctl32"), 413)             '--- 413 = DefSubclassProc ordinal
        '--- for IDE protection
        Debug.Assert pvThunkIdeOwner(aParams(7))
        If aParams(7) <> 0 Then
            aParams(8) = GetProcAddress(GetModuleHandle("user32"), "GetWindowLongA")
            aParams(9) = GetProcAddress(GetModuleHandle("vba6"), "EbMode")
            aParams(10) = GetProcAddress(GetModuleHandle("vba6"), "EbIsResetting")
        End If
        #If ImplSelfContained Then
            pvThunkGlobalData("InitSubclassingThunk") = hThunk
        #End If
    End If
    lSize = CallWindowProc(hThunk, hWnd, 0, VarPtr(aParams(0)), VarPtr(InitSubclassingThunk))
    Debug.Assert lSize = THUNK_SIZE
End Function
Private Function TerminateSubclassingThunk(pSubclass As IUnknown, pObj As Object) As IUnknown
    If Not pSubclass Is Nothing Then
        Debug.Assert ThunkPrivateData(pSubclass, 2) = ObjPtr(pObj)
        ThunkPrivateData(pSubclass, 2) = 0
        Set pSubclass = Nothing
    End If
End Function
Private Property Get ThunkPrivateData(pThunk As IUnknown, Optional ByVal Index As Long) As Long
Dim lPtr            As Long
    lPtr = ObjPtr(pThunk)
    If lPtr <> 0 Then
        Call CopyMemory(ThunkPrivateData, ByVal (lPtr Xor &H80000000) + 8 + Index * 4 Xor &H80000000, 4&)
    End If
End Property

Private Property Let ThunkPrivateData(pThunk As IUnknown, Optional ByVal Index As Long, ByVal lValue As Long)
Dim lPtr            As Long
    lPtr = ObjPtr(pThunk)
    If lPtr <> 0 Then
        Call CopyMemory(ByVal (lPtr Xor &H80000000) + 8 + Index * 4 Xor &H80000000, lValue, 4&)
    End If
End Property

Private Function pvThunkIdeOwner(hIdeOwner As Long) As Boolean
    #If Not ImplNoIdeProtection Then
        Dim lProcessId      As Long
        
        Do
            hIdeOwner = FindWindowEx(0, hIdeOwner, "IDEOwner", vbNullString)
            Call GetWindowThreadProcessId(hIdeOwner, lProcessId)
        Loop While hIdeOwner <> 0 And lProcessId <> GetCurrentProcessId()
    #End If
    pvThunkIdeOwner = True
End Function

Private Function pvThunkAllocate(sText As String, Optional ByVal Size As Long) As Long
    Static Map(0 To &H3FF) As Long
    Dim baInput()       As Byte
    Dim lIdx            As Long
    Dim lChar           As Long
    Dim lPtr            As Long
    
    pvThunkAllocate = VirtualAlloc(0, IIf(Size > 0, Size, (Len(sText) \ 4) * 3), &H1000, &H40)
    If pvThunkAllocate = 0 Then
        Exit Function
    End If
    '--- init decoding maps
    If Map(65) = 0 Then
        baInput = StrConv("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", vbFromUnicode)
        For lIdx = 0 To UBound(baInput)
            lChar = baInput(lIdx)
            Map(&H0 + lChar) = lIdx * (2 ^ 2)
            Map(&H100 + lChar) = (lIdx And &H30) \ (2 ^ 4) Or (lIdx And &HF) * (2 ^ 12)
            Map(&H200 + lChar) = (lIdx And &H3) * (2 ^ 22) Or (lIdx And &H3C) * (2 ^ 6)
            Map(&H300 + lChar) = lIdx * (2 ^ 16)
        Next
    End If
    '--- base64 decode loop
    baInput = StrConv(Replace(Replace(sText, vbCr, vbNullString), vbLf, vbNullString), vbFromUnicode)
    lPtr = pvThunkAllocate
    For lIdx = 0 To UBound(baInput) - 3 Step 4
        lChar = Map(baInput(lIdx + 0)) Or Map(&H100 + baInput(lIdx + 1)) Or Map(&H200 + baInput(lIdx + 2)) Or Map(&H300 + baInput(lIdx + 3))
        Call CopyMemory(ByVal lPtr, lChar, 3)
        lPtr = (lPtr Xor &H80000000) + 3 Xor &H80000000
    Next
End Function
#If ImplSelfContained Then
Private Property Get pvThunkGlobalData(sKey As String) As Long
    Dim sBuffer     As String
    
    sBuffer = String$(50, 0)
    Call GetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, sBuffer, Len(sBuffer) - 1)
    pvThunkGlobalData = Val(Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1))
End Property

Private Property Let pvThunkGlobalData(sKey As String, ByVal lValue As Long)
    Call SetEnvironmentVariable("_MST_GLOBAL" & GetCurrentProcessId() & "_" & sKey, lValue)
End Property
#End If

'=====================================================================================================================


Private Function LoWord(ByVal DWord As Long) As Integer
    If DWord And &H8000& Then LoWord = DWord Or &HFFFF0000 Else LoWord = DWord And &HFFFF&
End Function
Private Function HiWord(ByVal DWord As Long) As Integer
    HiWord = (DWord And &HFFFF0000) \ &H10000
End Function
Private Function MakeDWord(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
    MakeDWord = (CLng(HiWord) * &H10000) Or (LoWord And &HFFFF&)
End Function
Private Function Get_X_lParam(ByVal lParam As Long) As Long
    Get_X_lParam = lParam And &H7FFF&
    If lParam And &H8000& Then Get_X_lParam = Get_X_lParam Or &HFFFF8000
End Function
Private Function Get_Y_lParam(ByVal lParam As Long) As Long
    Get_Y_lParam = (lParam And &H7FFF0000) \ &H10000
    If lParam And &H80000000 Then Get_Y_lParam = Get_Y_lParam Or &HFFFF8000
End Function
Private Sub LongToByte(ByVal lLong As Long, ByRef bReturn() As Byte, Optional i As Integer = 0)
    bReturn(i) = lLong And &HFF
    bReturn(i + 1) = (lLong And 65280) / &H100
    bReturn(i + 2) = (lLong And &HFF0000) / &H10000
    bReturn(i + 3) = ((lLong And &HFF000000) \ &H1000000) And &HFF
End Sub

Private Property Get mvScreenDpi() As Single
    Dim hDC As Long
    Static Ret As Single
    If Ret = 0 Then
        hDC = GetDC(0&)
        Ret = GetDeviceCaps(hDC, 118) / (Screen.Width / Screen.TwipsPerPixelX)
        ReleaseDC 0&, hDC
    End If
    mvScreenDpi = Ret
End Property

Public Function WinProcpp(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
Attribute WinProcpp.VB_MemberFlags = "40"
    Select Case uMsg
        Case WM_DESTROY: pvUnsubclassPopup
        Case WM_NCHITTEST
                WinProcpp = 1
                Exit Function
        Case WM_KILLFOCUS
                pvUnsubclassPopup
                ShowWindow hWnd, 0
        Case WM_ACTIVATE
                
                '/* WA_INACTIVE */
                If (LoWord(wParam) = 0) Then
                    pvUnsubclassPopup
                    ShowWindow hWnd, 0
                End If
    End Select
End Function
Public Function WinProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, Handled As Boolean) As Long
Attribute WinProc.VB_MemberFlags = "40"
    Select Case uMsg
        Case WM_DESTROY:
        Case WMAPP_NOTIFYCALLBACK
            'Debug.Print "Here " & Timer
            
            Select Case LoWord(lParam)
                Case NIN_SELECT
                
                    ' /* For NOTIFYICON_VERSION_4
                    '    NIN_SELECT is prerable to listening
                    '    to mouse clicks and key presses directly. */
                    RaiseEvent Click
                Case NIN_KEYSELECT
                    RaiseEvent Click
                Case NIN_BALLOONSHOW
                    RaiseEvent BalloonShow
                Case NIN_BALLOONUSERCLICK
                    RaiseEvent BalloonClick
                Case NIN_BALLOONTIMEOUT
                    RaiseEvent BallonTimeOut
                Case NIN_BALLOONHIDE
                    Debug.Print "NIN_BALLOONHIDE"
                    
                Case WM_RBUTTONDBLCLK: RaiseEvent DblClick(vbRightButton)
                Case WM_LBUTTONDBLCLK: RaiseEvent DblClick(vbLeftButton)
                Case WM_MBUTTONDBLCLK: RaiseEvent DblClick(vbMiddleButton)
                    
                Case WM_RBUTTONDOWN:    RaiseEvent MouseDown(vbRightButton)
                Case WM_LBUTTONDOWN:    RaiseEvent MouseDown(vbLeftButton)
                Case WM_MBUTTONDOWN:    RaiseEvent MouseDown(vbMiddleButton)
                
                Case WM_RBUTTONUP:      RaiseEvent MouseUp(vbRightButton)
                Case WM_LBUTTONUP:      RaiseEvent MouseUp(vbLeftButton)
                Case WM_MBUTTONUP:      RaiseEvent MouseUp(vbMiddleButton)
                
                Case WM_CONTEXTMENU
                
                    Dim lPt(1) As Long
                    lPt(0) = Get_X_lParam(wParam) / mvScreenDpi
                    lPt(1) = Get_Y_lParam(wParam) / mvScreenDpi
                    
                    '/* Our window must be foreground before calling
                    '   TrackPopupMenu or the menu will not disappear
                    '   when the user clicks away */
                    
                    SetForegroundWindow (m_hWnd)
                    m_bInPopup = True
                    RaiseEvent ContextMenu(lPt(0), lPt(1))
                    m_bInPopup = False
                    
                Case NIN_POPUPOPEN
                    
                    Dim Rct As RECT
                    Call Shell_NotifyIconGetRect(ssNii, Rct)
                    SetRect Rct, Rct.Left / mvScreenDpi, Rct.Top / mvScreenDpi, Rct.Right / mvScreenDpi, Rct.Bottom / mvScreenDpi
                    RaiseEvent PopupOpen(Rct.Left, Rct.Top, Rct.Right, Rct.Bottom)
                    
                Case NIN_POPUPCLOSE
                    RaiseEvent PopupClose
                Case Else
                    'Debug.Print lParam, LoWord(lParam), "&H" & Hex(LoWord(lParam))
            End Select
            
        Case WM_TASKBAR '/* WM_TASKBARMESSAGE /*
        
            If m_bShow Then
                m_bShow = False
                ShowIcon m_hIcon, m_ToolTip
            End If
        Case Else: 'Debug.Print uMsg, "&H" & Hex(uMsg)
    End Select
End Function
