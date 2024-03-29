VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NotifyIcon "
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   229
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   385
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Icon"
      Height          =   1575
      Left            =   3840
      TabIndex        =   11
      Top             =   1560
      Width           =   1815
      Begin VB.OptionButton OpIcon 
         Caption         =   "Resource array"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton OpIcon 
         Caption         =   "Icon File"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton OpIcon 
         Caption         =   "Handle"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton BtnFlyout 
      Caption         =   "Show Flyout"
      Height          =   495
      Left            =   3840
      TabIndex        =   10
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton BtnMain 
      Caption         =   "Show Ballon"
      Height          =   495
      Left            =   3840
      TabIndex        =   9
      Top             =   360
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ballon icon"
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   3495
      Begin VB.OptionButton OpBIcon 
         Caption         =   "Custom icon"
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton OpBIcon 
         Caption         =   "Error"
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton OpBIcon 
         Caption         =   "Warning"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton OpBIcon 
         Caption         =   "Info"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton OpBIcon 
         Caption         =   "None"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tooltip"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3495
      Begin VB.OptionButton OpTooltip 
         Caption         =   "Custom Tooltip"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton OpTooltip 
         Caption         =   "Standar tooltip"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "TrayMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopup 
         Caption         =   "Show Flyout"
         Index           =   0
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "Options"
         Index           =   1
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "Exit"
         Index           =   3
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'/* Api Menu */
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal hWnd As Long, ByVal lptpm As Any) As Long
Private Declare Function AppendMenuA Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long


Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private WithEvents c_Notify As cNotifyIcon
Attribute c_Notify.VB_VarHelpID = -1

Private Sub Form_Load()

    Set c_Notify = New cNotifyIcon
    
    
    ' MOSTRAR ICONO EN EL SYSTRAY
    '--------------------------------------------------------------------------------------
    ' Para mostrar tooltip usando una ventana personalizada establezca el tercer parametro
    ' en true.
    '--------------------------------------------------------------------------------------
    
    'c_Notify.ShowIcon Me.Icon, "App Tooltip"                   '/* Handle Icon     */
    'c_Notify.ShowIcon App.Path & "\000.ico", "App Tooltip"     '/* Icon File       */
    'c_Notify.ShowIcon "#207", "App Tooltip"                    '/* Resource Icon   */
    'c_Notify.ShowIcon LoadResData(103, "PNG"), "App Tooltip"   '/* PNG Array       */
    
    c_Notify.ShowIcon Me.Icon, "App Tooltip"

    
    ' CAMBIAR ICONO EN EL SYSTRAY
    '--------------------------------------------------------------------------------------
    'c_Notify.Icon = Me.Icon                '/* Handle Icon                 */
    'c_Notify.Icon = App.Path & "\001.ico"  '/* Icon File                   */
    'c_Notify.Icon = "#208"                 '/* Resource Icon               */
    'c_Notify.Icon = LoadResData(103, "PNG")'/* PNG Array                   */
    
    ' CAMBIAR TOOLTIP
    '--------------------------------------------------------------------------------------
    ' Puede usar el tooltip estandar del sistema o una ventana personalizada a traves de
    ' los eventos PopupOpen y PopupClose
    '--------------------------------------------------------------------------------------
    
    'c_Notify.ShowTooltip = True  '/* Standart Tooltip                               */
    'c_Notify.ShowTooltip = False '/* Custom Tooltip events (PopupOpen & PopupClose) */
    
    
    ' ELIMINAR ICONO
    '--------------------------------------------------------------------------------------
    'c_Notify.RemoveIcon
    
    
    ' CLSID
    '--------------------------------------------------------------------------------------
    'c_Notify.Guid = "{00000000-0000-0000-0000-000000000000}"
    
End Sub




'TODO: Notify Tooltip Setup
'-----------------------------------------------------------------------------------------------------
Private Sub OpTooltip_Click(Index As Integer)
    Select Case Index
        Case 0: c_Notify.ShowTooltip = True     ' Standart Tooltip
        Case 1: c_Notify.ShowTooltip = False    ' Custom Tooltip events (PopupOpen & PopupClose)
    End Select
End Sub

'TODO: Notify Icon Setup
'-----------------------------------------------------------------------------------------------------
Private Sub OpIcon_Click(Index As Integer)
    Select Case Index
        Case 0: c_Notify.Icon = Me.Icon
        Case 1: c_Notify.Icon = App.Path & "\000.ico"
        Case 2: c_Notify.Icon = LoadResData(101, "PNG")
    End Select
End Sub

'TODO: Ballon Setup && Show
'-----------------------------------------------------------------------------------------------------
Private Sub btnMain_Click()
Dim mbIcon As NIIF_ICON_TYPE
    
    
    If Not OpBIcon(4).Value Then
    
        '/* Show Standar Ballon Icons */
        
        If OpBIcon(0).Value Then mbIcon = NIIF_NONE
        If OpBIcon(1).Value Then mbIcon = NIIF_INFO
        If OpBIcon(2).Value Then mbIcon = NIIF_WARNING
        If OpBIcon(3).Value Then mbIcon = NIIF_ERROR
        
        c_Notify.ShowBallon "Hello World... VB6!", "Important!", mbIcon
        
    Else
        
        '/* Show Custom Ballon Icon (48px * 48px) */
      
        'c_Notify.ShowBallon "Hello World... VB6!", "Important!", , Me.Icon                '/* Handle Icon        */
        'c_Notify.ShowBallon "Hello World... VB6!", "Important!", , App.Path & "\001.ico"  '/* Icon File          */
        'c_Notify.ShowBallon "Hello World... VB6!", "Important!", , "#200"                 '/* Resource Icon      */
        c_Notify.ShowBallon "Hello World... VB6!", "Important!", , LoadResData(102, "PNG")
        
    End If
End Sub
Private Sub BtnFlyout_Click()
    c_Notify.ShowFlyout Form3
End Sub



'TODO: NotifyIcon Events
'-----------------------------------------------------------------------------------------------------
Private Sub c_Notify_Click()

    c_Notify.ShowFlyout Form3
    
End Sub
Private Sub c_Notify_ContextMenu(ByVal x As Single, ByVal y As Single)
Dim lCmd As Long
    lCmd = ApiPopupMenu(Me.hWnd, x, y)
    
    Select Case lCmd
        Case 101: c_Notify.ShowFlyout Form3
        Case 102: Me.Show
        Case 103: Unload Me
    End Select
    
End Sub

Private Sub c_Notify_PopupOpen(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
Dim x As Long, y As Long

    '/* Mostrar Tooltip Personalizado */
    
    Load Form2
    If c_Notify.CalculateFlyout(Form2.hWnd, x, y) Then
        SetWindowPos Form2.hWnd, HWND_TOPMOST, x, y, 0, 0, SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
    End If

End Sub
Private Sub c_Notify_PopupClose()

    '/* Ocultar Tooltip Personalizado */
    Unload Form2
    
End Sub




Private Sub Form_Unload(Cancel As Integer)
    Set c_Notify = Nothing
End Sub




Private Function ApiPopupMenu(hWnd As Long, ByVal x As Long, ByVal y As Long) As Long
Dim hMenu As Long

    hMenu = CreatePopupMenu()
    AppendMenuA hMenu, &H0&, 101, "Show Flyout"
    AppendMenuA hMenu, &H0&, 102, "Main Form"
    AppendMenuA hMenu, &H800&, 0, ByVal 0&
    AppendMenuA hMenu, &H0&, 103, "E&xit"
    ApiPopupMenu = TrackPopupMenuEx(hMenu, &H0& Or &H100& Or &H2&, x, y, hWnd, ByVal 0&)
    DestroyMenu hMenu

End Function

