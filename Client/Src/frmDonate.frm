VERSION 5.00
Begin VB.Form frmDonate 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Donate"
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrHotTrack 
      Left            =   9000
      Top             =   5160
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Please wait"
      BeginProperty Font 
         Name            =   "Eurostar"
         Size            =   15.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   1680
      TabIndex        =   1
      Top             =   3720
      Width           =   7215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Press return to merchant after purchase is done."
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   8880
      Width           =   10935
   End
End
Attribute VB_Name = "frmDonate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim iindex As Integer
browser.AddBrowser iindex
browser.navigate iindex, "https://www.sandbox.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=RBKP5TVBNL77E"
End Sub

