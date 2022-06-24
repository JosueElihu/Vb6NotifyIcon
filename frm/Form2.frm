VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   630
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2250
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   42
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   150
   StartUpPosition =   3  'Windows Default
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000C0FF&
      FillStyle       =   0  'Solid
      Height          =   210
      Left            =   1860
      Shape           =   3  'Circle
      Top             =   210
      Width           =   210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hello World... VB6!"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   300
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tooltip"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   90
      Width           =   585
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE       As Long = -20
Private Const GWL_STYLE         As Long = -16
Private Const WS_POPUP          As Long = &H80000000
Private Const WS_BORDER         As Long = &H800000

Private Const WS_EX_TOOLWINDOW  As Long = 128
Private Const WS_EX_TOPMOST     As Long = &H8&

Private Sub Form_Load()
Dim lStyle As Long

    Call SetWindowLongA(Me.hwnd, GWL_STYLE, WS_POPUP Or WS_BORDER)
    Call SetWindowLongA(Me.hwnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
    
   
End Sub
