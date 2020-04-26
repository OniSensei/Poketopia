VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LVbutton.ocx"
Begin VB.Form frmOptions 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pnlOptions 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   360
      ScaleHeight     =   2895
      ScaleWidth      =   3255
      TabIndex        =   1
      Top             =   720
      Width           =   3255
      Begin VB.CheckBox Check5 
         BackColor       =   &H00000000&
         Caption         =   "Play Radio"
         BeginProperty Font 
            Name            =   "Eurostar"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00000000&
         Caption         =   "Form transparency"
         BeginProperty Font 
            Name            =   "Eurostar"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00000000&
         Caption         =   "Camera Follow Player"
         BeginProperty Font 
            Name            =   "Eurostar"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Play Audio"
         BeginProperty Font 
            Name            =   "Eurostar"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00000000&
         Caption         =   "Repeat Map Music"
         BeginProperty Font 
            Name            =   "Eurostar"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1935
      End
      Begin lvButton.lvButtons_H lvButtons_H9 
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   2040
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Caption         =   "Save"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   -2147483633
      End
   End
   Begin VB.Label lblPOKE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Eurostar"
         Size            =   15.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'function to make transparent'

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
  ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const G = (-20)
Private Const LWA_COLORKEY = &H1        'to trans'
Private Const LWA_ALPHA = &H2           'to semi trans'
Private Const WS_EX_LAYERED = &H80000

Private Sub Form_Activate()
    Me.BackColor = vbBlack
    If Options.FormTransparency = YES Then
    trans 215
    End If
End Sub

Private Sub trans(Level As Integer)
    Dim Msg As Long
    Msg = GetWindowLong(Me.hwnd, G)
    Msg = Msg Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, G, Msg
    SetLayeredWindowAttributes Me.hwnd, vbBlack, Level, LWA_ALPHA
End Sub
Private Sub Form_Load()
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
Check1 = Options.PlayMusic
Check2 = Options.repeatmusic
Check3 = Options.CameraFollowPlayer
Check4 = Options.FormTransparency
Check5 = Options.PlayRadio
End Sub

Private Sub lvButtons_H9_Click()
If Check1.Value = 1 Then
Options.PlayMusic = YES
PlayMapMusic MapMusic
Else
Options.PlayMusic = NO
StopPlay
End If
If Check2.Value = 1 Then
Options.repeatmusic = YES
Else
Options.repeatmusic = NO
End If
If Check3.Value = 1 Then
Options.CameraFollowPlayer = YES
Else
Options.CameraFollowPlayer = NO
End If
If Check4.Value = 1 Then
Options.FormTransparency = YES
Else
Options.FormTransparency = NO
End If
If Check5.Value = 1 Then
Options.PlayRadio = YES
Else
Options.PlayRadio = NO
End If
SaveOptions
Unload Me
End Sub

