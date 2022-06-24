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
'   Component    : cNotifyIcon v1.0
'   Autor        : J. Elihu
'----------------------------------------------------------------------------------------------------------------------
'   Description  : ShellNotifyIcon VERSION_4 Compatible
'   Credits      : Leandro Ascierto (ClsSysTray, ClsPopUpTray)
'======================================================================================================================

Option Explicit

Private Type GUID
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

'/* GUID */
Private Declare Sub CoCreateGuid Lib "ole32" (ByRef pguid As GUID)
Private Declare Function StringFromGUID2 Lib "ole32" (ByVal rguid As Long, ByVal lpsz As Long, ByVal cchMax As Long) As Long

'/* Sub Class */
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'/* Window */
Private Declare Function CreateWindowExA Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

'/* Icon */
Private Declare Function CopyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function CreateIconFromResourceEx Lib "user32" (ByRef presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long
Private Declare Function LoadImageA Lib "user32" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Private Declare Function PathFileExistsA Lib "shlwapi" (ByVal pszPath As String) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

'/* Taskbar && Popup */
Private Declare Function RegisterWindowMessageA Lib "user32" (ByVal lpString As String) As Long
Private Declare Function CalculatePopupWindowPosition Lib "user32" (anchorPoint As POINTAPI, windowSize As POINTAPI, ByVal Flags As Long, excludeRect As RECT, PopupWindowPosition As RECT) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function GetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function WaitMessage Lib "user32" () As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long

'/* Animate */
Private Declare Function DrawAnimatedRects Lib "user32" (ByVal hwnd As Long, ByVal idAni As Long, lprcFrom As RECT, lprcTo As RECT) As Long

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

'//DPI
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal Hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal Hdc As Long) As Long


'// NotifyIcon
Private Declare Function Shell_NotifyIconA Lib "Shell32" (ByVal dwMessage As Long, ByRef lpData As Any) As Long
Private Declare Function Shell_NotifyIconGetRect Lib "Shell32" (ByRef identifier As NOTIFYICONIDENTIFIER, ByRef iconLocation As Any) As Long

Private Type NOTIFYICONDATAA
  cbSize            As Long
  hwnd              As Long
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
  guidItem          As GUID     'lGuid(3) As Long
  hBalloonIcon      As Long
End Type

Private Type NOTIFYICONIDENTIFIER
  cbSize            As Long
  hwnd              As Long
  uId               As Long
  guidItem          As GUID     'lGuid(3) As Long
End Type

'// Windows messages
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


'// Callback Messages
Private Const NIN_SELECT            As Long = WM_USER + 0
Private Const NINF_KEY              As Long = &H1
Private Const NIN_BALLOONSHOW       As Long = (WM_USER + 2)
Private Const NIN_BALLOONHIDE       As Long = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT    As Long = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK  As Long = (WM_USER + 5)
Private Const NIN_POPUPOPEN         As Long = WM_USER + 6
Private Const NIN_POPUPCLOSE        As Long = WM_USER + 7
Private Const NIN_KEYSELECT         As Long = NIN_SELECT Or NINF_KEY


'// Update commands
Private Const NIM_ADD               As Long = &H0
Private Const NIM_MODIFY            As Long = &H1
Private Const NIM_DELETE            As Long = &H2
Private Const NIM_SETFOCUS          As Long = &H3
Private Const NIM_SETVERSION        As Long = &H4


'// Update masks
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

'// Version
Private Const NOTIFYICON_VERSION    As Long = 3
Private Const NOTIFYICON_VERSION_4  As Long = 4

Public Enum NIIF_ICON_TYPE
  [NIIF_NONE] = 0
  [NIIF_INFO] = 1
  [NIIF_WARNING] = 2
  [NIIF_ERROR] = 3
  [NIIF_USER] = 4
End Enum

'// DrawAnimatedRects
Private Const IDANI_OPEN        As Long = &H1
Private Const IDANI_CLOSE       As Long = &H2
Private Const IDANI_CAPTION     As Long = &H3

'/* Sub Class */
Private bvASM(40)               As Byte
Private bvProcs                 As New Collection

'/* Events */
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


'/* -------------- */
Private hWnd_   As Long
Private Guid_   As GUID
Private WmTb_   As Long '/* WM_TASKBARMESSAGE */

Private Icon_   As Long
Private Tltp_   As String
Private Flag_   As Boolean
Private ShwTip_ As Boolean
Private Popup_  As Boolean

'/* Flyout Window */
Private hWndPp_ As Long
Private Tmlast_ As Long


Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    '/* Flyout */
    If hwnd = hWndPp_ Then
        Select Case uMsg
            Case WM_DESTROY
                Call StopSubclassing(hwnd)
            Case WM_NCHITTEST
                WindowProc = 1
                Exit Function
            Case WM_KILLFOCUS
                Call StopSubclassing(hwnd)
                ShowWindow hwnd, 0
                
            Case WM_ACTIVATE
                
                '/* WA_INACTIVE */
                If (LoWord(wParam) = 0) Then
                    Call StopSubclassing(hwnd)
                    ShowWindow hwnd, 0
                End If
            Case Else
        End Select
        GoTo e_
    End If
    
    Select Case uMsg
        Case WM_DESTROY
            Call StopSubclassing(hwnd)
        Case WMAPP_NOTIFYCALLBACK
        
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
                    lPt(0) = Get_X_lParam(wParam) / ssScreenDpi__
                    lPt(1) = Get_Y_lParam(wParam) / ssScreenDpi__
                    
                    '/* Our window must be foreground before calling
                    '   TrackPopupMenu or the menu will not disappear
                    '   when the user clicks away */
                    
                    SetForegroundWindow (hWnd_)
                    Popup_ = True
                    RaiseEvent ContextMenu(lPt(0), lPt(1))
                    Popup_ = False
                    
                Case NIN_POPUPOPEN
                    
                    Dim Rct_ As RECT
                    Call Shell_NotifyIconGetRect(ssNii__, Rct_)
                    SetRect Rct_, Rct_.Left / ssScreenDpi__, Rct_.Top / ssScreenDpi__, Rct_.Right / ssScreenDpi__, Rct_.Bottom / ssScreenDpi__
                    RaiseEvent PopupOpen(Rct_.Left, Rct_.Top, Rct_.Right, Rct_.Bottom)
                    
                Case NIN_POPUPCLOSE
                    RaiseEvent PopupClose
                Case Else
                
            End Select
            
        Case WmTb_ '/* WM_TASKBARMESSAGE /*
        
            If Flag_ Then
                Flag_ = False
                ShowIcon Icon_, Tltp_
            End If
        
        Case Else
            'Debug.Print "Here " & Timer
    End Select
    
e_:
    WindowProc = CallWindowProcA(ssPrevWndProc__(hwnd), hwnd, uMsg, wParam, lParam)
End Function

Private Sub Class_Initialize()

    WmTb_ = RegisterWindowMessageA("TaskbarCreated")
    hWnd_ = CreateWindowExA(0, "STATIC", vbNullString, 0, 0, 0, 0, 0, 0, 0, App.hInstance, 0)
    
    '/* Guid to Uniquely Identify */
    Call CoCreateGuid(Guid_)
    Call SetSubclassing(Me, hWnd_)
    

End Sub
Private Sub Class_Terminate()

    If Icon_ Then Call DestroyIcon(Icon_): Icon_ = 0
    If Flag_ Then Call DeleteIcon
    
    If hWndPp_ Then Call CancelFlyout
    If hWnd_ Then
        Call StopSubclassing(hWnd_)
        Call DestroyWindow(hWnd_)
    End If
    
End Sub

Public Function ShowIcon(Optional Icon As Variant, Optional ByVal Tooltip As String) As Boolean
Dim nid_ As NOTIFYICONDATAA

    If Flag_ Then Call DeleteIcon
    
    nid_ = ssNid__(NIF_ICON Or NIF_MESSAGE Or NIF_TIP) '/* With POPUPOPEN */
    nid_.hwnd = hWnd_
    nid_.uCallbackMessage = WMAPP_NOTIFYCALLBACK
    
    If Not IsMissing(Icon) Then nid_.hIcon = ssIcon__(Icon, True)

    '/* Use the Standard Tooltip */
    ShwTip_ = (Len(Tooltip) > 0)
    If Len(Tooltip) > 0 Then
        nid_.uFlags = nid_.uFlags Or NIF_SHOWTIP
        nid_.szTip = Tooltip & vbNullChar
    Else
        nid_.szTip = " " & vbNullChar
    End If
   
    ShowIcon = Shell_NotifyIconA(NIM_ADD, nid_)
   
    '/* NOTIFYICON_VERSION_4 Is Prefered */
    nid_.uVersion = NOTIFYICON_VERSION_4
    Call Shell_NotifyIconA(NIM_SETVERSION, nid_)
    
    Flag_ = ShowIcon
    
    '/* Save data to use with WM_TASKBARMESSAGE */
    If Icon_ Then Call DestroyIcon(Icon_): Icon_ = 0
    If nid_.hIcon Then Icon_ = CopyIcon(nid_.hIcon)
    Tltp_ = Tooltip
  
End Function

Public Function DeleteIcon() As Boolean
    DeleteIcon = Shell_NotifyIconA(NIM_DELETE, ssNid__)
End Function


Public Function ShowBallon(ByVal Text As String, ByVal Title As String, Optional ByVal IconType As NIIF_ICON_TYPE, Optional BallonIcon As Variant, _
                           Optional ByVal NoSound As Boolean, Optional ByVal RespectQuietTime As Boolean = True, Optional ByVal RealTime As Boolean) As Boolean
                           
Dim nid_     As NOTIFYICONDATAA

    nid_ = ssNid__(NIF_INFO)
    nid_.dwInfoFlags = IconType
    
    If Not IsMissing(BallonIcon) Then nid_.hBalloonIcon = ssIcon__(BallonIcon, False)
    If nid_.hBalloonIcon Then nid_.dwInfoFlags = NIIF_USER Or NIIF_LARGE_ICON
     
    If RespectQuietTime Then nid_.dwInfoFlags = nid_.dwInfoFlags Or NIIF_RESPECT_QUIET_TIME
    If NoSound Then nid_.dwInfoFlags = nid_.dwInfoFlags Or NIIF_NOSOUND
    
    If RealTime Then nid_.uFlags = nid_.uFlags Or NIF_REALTIME
    If ShwTip_ Then nid_.uFlags = nid_.uFlags Or NIF_SHOWTIP
    
    nid_.szInfoTitle = Title & vbNullChar
    nid_.szInfo = Text & vbNullChar
    
    ShowBallon = Shell_NotifyIconA(NIM_MODIFY, nid_)
    If nid_.hBalloonIcon Then Call DestroyIcon(nid_.hBalloonIcon)
    
End Function
Public Function HideBallon() As Boolean
Dim nid_     As NOTIFYICONDATAA
    
    nid_ = ssNid__(NIF_INFO)
    nid_.szInfoTitle = vbNullChar
    nid_.szInfo = vbNullChar
    HideBallon = Shell_NotifyIconA(NIM_MODIFY, nid_)
    
End Function
Public Function SetFocus() As Boolean
    SetFocus = (Shell_NotifyIconA(NIM_SETFOCUS, ssNid__) = 0)
End Function
Public Sub BeforePopup()
    Call SetForegroundWindow(hWnd_)
End Sub

Public Function ShowFlyout(Win As Variant, Optional ByVal UnloadIfVbForm As Boolean = True, Optional ChangeStyle As Boolean = True) As Boolean
On Error GoTo e_
Dim hWin    As Long

    If Tmlast_ + 500 > GetTickCount Then
        Call StopSubclassing(hWndPp_)
        ShowWindow hWndPp_, 0
        hWndPp_ = 0
        Exit Function
    End If
    
    Select Case VarType(Win)
        Case vbObject: hWin = Win.hwnd
        Case vbLong: hWin = Win
    End Select
    
    Dim lx As Long, ly As Long
    
    Tmlast_ = Time
    
    '// Display the flyout in an appropriate position
    If CalculateFlyout(hWin, lx, ly) Then
        
        '// Position the flyout and make it the foreground window
        If ChangeStyle Then
            SetWindowLongA hWin, GWL_STYLE, WS_THICKFRAME Or WS_BORDER 'Or WS_POPUP
            SetWindowLongA hWin, GWL_EXSTYLE, WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW
        End If
        
        SetWindowPos hWin, HWND_TOPMOST, lx, ly, 0, 0, SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_FRAMECHANGED Or SWP_NOACTIVATE
        
        hWndPp_ = hWin
        
        SetSubclassing Me, hWin
        SetForegroundWindow (hWin)
        
        Do While GetForegroundWindow = hWin
            DoEvents
            WaitMessage
        Loop
        
        ShowWindow hWin, 0
        Call StopSubclassing(hWin)
        hWndPp_ = 0
        
        Tmlast_ = GetTickCount
    End If
    
e_:
    On Error Resume Next
    If VarType(Win) = vbObject Then
        If ssIsVBForm__(Win) And UnloadIfVbForm Then Unload Win
    End If
    
End Function
Public Sub CancelFlyout()
    If IsWindow(hWndPp_) Then DestroyWindow hWndPp_
End Sub
Public Function CalculateFlyout(hwnd As Long, Left As Long, Top As Long, Optional ByVal Flags As Long) As Boolean
Dim Point_  As POINTAPI
Dim Size_   As POINTAPI
Dim Rct_    As RECT
Dim WinRct  As RECT

    '/* Flags */
    'Const TPM_VERTICAL = &H40
    'Const TPM_VCENTERALIGN = &H10&
    'Const TPM_CENTERALIGN = &H4&
    'Const TPM_WORKAREA = &H10000
    
    If Flags = 0 Then Flags = &H40 Or &H10& Or &H4& Or &H10000
    
    '/* PositionFlyout */
    Select Case Shell_NotifyIconGetRect(ssNii__, Rct_)
        Case 0, 1
        
            Point_.x = ((Rct_.Left + Rct_.Right) / 2) / ssScreenDpi__
            Point_.y = ((Rct_.Top + Rct_.Bottom) / 2) / ssScreenDpi__
            
            GetWindowRect hwnd, WinRct
            Size_.x = (WinRct.Right - WinRct.Left)
            Size_.y = (WinRct.Bottom - WinRct.Top)
            
            SetRect Rct_, Rct_.Left / ssScreenDpi__, Rct_.Top / ssScreenDpi__, Rct_.Right / ssScreenDpi__, Rct_.Bottom / ssScreenDpi__
            If CalculatePopupWindowPosition(Point_, Size_, Flags, Rct_, WinRct) Then
                Left = WinRct.Left: Top = WinRct.Top
                CalculateFlyout = True
            End If
            
        Case Else
    End Select

End Function

Public Function IconRect(ByRef x1 As Long, ByRef y1 As Long, ByRef x2 As Long, ByRef y2 As Long) As Boolean
Dim Rct_ As RECT
    Select Case Shell_NotifyIconGetRect(ssNii__, Rct_)
        Case 0, 1
            SetRect Rct_, Rct_.Left / ssScreenDpi__, Rct_.Top / ssScreenDpi__, Rct_.Right / ssScreenDpi__, Rct_.Bottom / ssScreenDpi__
            x1 = Rct_.Left: y1 = Rct_.Top
            x2 = Rct_.Right: y2 = Rct_.Bottom
            IconRect = True
        Case Else
    End Select
End Function

Property Get GUID() As String: GUID = ssGuidToStr__(Guid_): End Property
Property Get InPopupMenu() As Boolean: InPopupMenu = Popup_: End Property
Property Let Icon(value As Variant)
Dim nid_ As NOTIFYICONDATAA

    nid_ = ssNid__(NIF_ICON)
    nid_.hIcon = ssIcon__(value, True)
    
    If ShwTip_ Then nid_.uFlags = nid_.uFlags Or NIF_SHOWTIP
    If nid_.hIcon Then Shell_NotifyIconA NIM_MODIFY, nid_
    
End Property
Property Let Tooltip(ByVal value As String)
Dim nid_ As NOTIFYICONDATAA
    
    ShwTip_ = Len(value)
    If Len(value) Then
    
        '/* Use the Standard Tooltip */
        nid_ = ssNid__(NIF_TIP Or NIF_SHOWTIP)
        nid_.szTip = value & vbNullChar
        
    Else
        nid_ = ssNid__(NIF_INFO)
    End If
    Shell_NotifyIconA NIM_MODIFY, nid_
    
End Property


'TODO: Private Subs
'---------------------------------------------------------------------------------------------------------------------

Private Function ssNid__(Optional uFlags As Long) As NOTIFYICONDATAA

    '/* Icon will be identified with the GUID */
    ssNid__.cbSize = Len(ssNid__)
    ssNid__.guidItem = Guid_
    ssNid__.uFlags = uFlags Or NIF_GUID

End Function
Private Function ssNii__() As NOTIFYICONIDENTIFIER
    ssNii__.cbSize = Len(ssNii__)
    ssNii__.guidItem = Guid_
End Function
Private Function ssIcon__(Source As Variant, Optional Small As Boolean) As Long
On Error GoTo e_
    Select Case VarType(Source)
        Case vbString
        
            Const SM_CXSMICON       As Long = 49
            Const SM_CYSMICON       As Long = 50
            
            Dim cx As Long: Dim cy As Long
            cx = GetSystemMetrics(SM_CXSMICON)
            cy = GetSystemMetrics(SM_CYSMICON)
            
            If PathFileExistsA(Source) Then
            
                If Small Then
                    ssIcon__ = LoadImageA(0&, Source, 1&, cx, cy, &H10)
                Else
                    ssIcon__ = LoadImageA(0&, Source, 1&, 0, 0, &H10)
                End If
                
            Else
                If IsNumeric(Source) And Left(Source, 1) <> "#" Then Source = "#" & Source
                If Small Then
                    ssIcon__ = LoadImageA(App.hInstance, Source, 1, cx, cy, &H8000& Or &H1000)
                Else
                    ssIcon__ = LoadImageA(App.hInstance, Source, 1, 0, 0, &H8000& Or &H1000)
                End If
            End If
            
        Case (vbArray + vbByte)
            Dim Out() As Byte
            Out() = Source
            ssIcon__ = CreateIconFromResourceEx(Out(LBound(Out)), UBound(Out) - LBound(Out) + 1, 1&, &H30000, 0&, 0&, 0&)
            
        Case vbLong:
            ssIcon__ = CopyIcon(Source)
    End Select
e_:
End Function
Private Function ssGuidToStr__(mGuid As GUID) As String
Dim ab() As Byte
Dim lm   As Long
    
    ReDim ab(80)
    lm = StringFromGUID2(VarPtr(mGuid.Data1), VarPtr(ab(0)), UBound(ab))
    ssGuidToStr__ = Left(ab, lm)
    
End Function
Private Function ssIsVBForm__(Obj As Variant) As Boolean
Dim Frm As Form

    For Each Frm In Forms
        'If TypeName(Frm) = TypeName(Obj) Then ssIsVBForm__ = True: Exit Function
        If Obj Is Frm Then ssIsVBForm__ = True: Exit Function
    Next
End Function

Private Property Get ssScreenDpi__() As Single
    Dim Hdc As Long
    Static Ret As Single
    If Ret = 0 Then
        Hdc = GetDC(0&)
        Ret = GetDeviceCaps(Hdc, 118) / (Screen.Width / Screen.TwipsPerPixelX)
        ReleaseDC 0&, Hdc
    End If
    ssScreenDpi__ = Ret
End Property


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



Private Sub SetSubclassing(Obj As Object, hwnd As Long)
Dim WindowProcAddress As Long
Dim pObj    As Long
Dim pVar    As Long
Dim lProc   As Long
 
    Dim i As Long
 
    For i = 0 To 40
        bvASM(i) = Choose(i + 1, &H55, &H8B, &HEC, &H83, &HC4, &HFC, &H8D, &H45, &HFC, &H50, &HFF, &H75, &H14, _
                                 &HFF, &H75, &H10, &HFF, &H75, &HC, &HFF, &H75, &H8, &H68, &H0, &H0, &H0, &H0, _
                                 &HB8, &H0, &H0, &H0, &H0, &HFF, &HD0, &H8B, &H45, &HFC, &HC9, &HC2, &H10, &H0)
    Next i
 
    pObj = ObjPtr(Obj)
 
    Call CopyMemory(pVar, ByVal pObj, 4)
    'Call CopyMemory(WindowProcAddress, ByVal (pVar + &H1C + (ProcIndex * 4&)), 4)
    Call CopyMemory(WindowProcAddress, ByVal (pVar + 28), 4)
 
    Call LongToByte(pObj, bvASM, 23)
    Call LongToByte(WindowProcAddress, bvASM, 28)
 
    lProc = SetWindowLongA(hwnd, -4, VarPtr(bvASM(0)))
    bvProcs.Add lProc, CStr(hwnd)
    
End Sub
Private Sub StopSubclassing(hwnd As Long)
On Error GoTo e_
Dim lProc As Long
    
    lProc = ssPrevWndProc__(hwnd)
    If lProc Then Call SetWindowLongA(hwnd, -4, lProc)
    bvProcs.Remove CStr(hwnd)
e_:
End Sub
Private Sub LongToByte(ByVal lLong As Long, ByRef bReturn() As Byte, Optional i As Integer = 0)
    bReturn(i) = lLong And &HFF
    bReturn(i + 1) = (lLong And 65280) / &H100
    bReturn(i + 2) = (lLong And &HFF0000) / &H10000
    bReturn(i + 3) = ((lLong And &HFF000000) \ &H1000000) And &HFF
End Sub
Private Function ssPrevWndProc__(hwnd As Long) As Long
On Error GoTo e_
    ssPrevWndProc__ = bvProcs(CStr(hwnd))
e_:
End Function
