VERSION 5.00
Object = "{50347EDF-F3EF-4392-AFDD-71AE67A3A978}#1.0#0"; "LaVolpeAlphaImg2.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "LVbutton.ocx"
Begin VB.Form frmPokemons 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pokemons"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPokemons.frx":0000
   ScaleHeight     =   5760
   ScaleWidth      =   10800
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picItems 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   6240
      ScaleHeight     =   4215
      ScaleWidth      =   6135
      TabIndex        =   62
      Top             =   360
      Visible         =   0   'False
      Width           =   6135
      Begin VB.ListBox lstItems 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Eurostar"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2340
         Left            =   120
         TabIndex        =   66
         Top             =   1080
         Width           =   5895
      End
      Begin lvButton.lvButtons_H lvButtons_H6 
         Height          =   375
         Left            =   120
         TabIndex        =   63
         Top             =   3600
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   661
         Caption         =   "Close"
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
         cFore           =   0
         cFHover         =   0
         cBhover         =   8421504
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin VB.Label lblItemsCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "   Use items"
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
         Height          =   615
         Index           =   0
         Left            =   0
         TabIndex        =   65
         Top             =   120
         Width           =   6135
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Double click to use"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   64
         Top             =   720
         Width           =   2895
      End
   End
   Begin lvButton.lvButtons_H btnRefresh 
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   31
      Top             =   6600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Refresh"
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
      cFore           =   16777215
      cFHover         =   4210752
      cBhover         =   4210752
      LockHover       =   3
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   0
   End
   Begin VB.PictureBox picInformation 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   0
      ScaleHeight     =   5775
      ScaleWidth      =   10815
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      Begin VB.PictureBox picMoves 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   4215
         Left            =   6240
         ScaleHeight     =   4215
         ScaleWidth      =   6135
         TabIndex        =   45
         Top             =   360
         Visible         =   0   'False
         Width           =   6135
         Begin VB.ListBox lstMoves 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Eurostar"
               Size            =   9.75
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   3180
            Left            =   3120
            TabIndex        =   46
            Top             =   600
            Width           =   2895
         End
         Begin lvButton.lvButtons_H lvButtons_H4 
            Height          =   375
            Left            =   120
            TabIndex        =   51
            Top             =   3600
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   661
            Caption         =   "Close"
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
            cFore           =   0
            cFHover         =   0
            cBhover         =   8421504
            LockHover       =   3
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   12632256
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "Double click to learn move."
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3120
            TabIndex        =   52
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label lblMove 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "001"
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
            Height          =   615
            Index           =   4
            Left            =   0
            TabIndex        =   50
            Top             =   2760
            Width           =   3135
         End
         Begin VB.Label lblMove 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "001"
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
            Height          =   615
            Index           =   3
            Left            =   0
            TabIndex        =   49
            Top             =   2040
            Width           =   3135
         End
         Begin VB.Label lblMove 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "001"
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
            Height          =   615
            Index           =   2
            Left            =   0
            TabIndex        =   48
            Top             =   1320
            Width           =   3135
         End
         Begin VB.Label lblMove 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "Fire Spin"
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
            Height          =   615
            Index           =   1
            Left            =   0
            TabIndex        =   47
            Top             =   600
            Width           =   3135
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   9120
         ScaleHeight     =   1695
         ScaleWidth      =   1575
         TabIndex        =   53
         Top             =   600
         Width           =   1575
         Begin VB.Label lblNatureHP 
            BackStyle       =   0  'Transparent
            Caption         =   "HP +1"
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
            Left            =   0
            TabIndex        =   60
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblNatNm 
            BackStyle       =   0  'Transparent
            Caption         =   "Nature boost"
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
            Height          =   375
            Left            =   0
            TabIndex        =   59
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label lblNatureAtk 
            BackStyle       =   0  'Transparent
            Caption         =   "Atk +1"
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
            Left            =   0
            TabIndex        =   58
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblNatureDef 
            BackStyle       =   0  'Transparent
            Caption         =   "Def +1"
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
            Left            =   0
            TabIndex        =   57
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblNatureSpAtk 
            BackStyle       =   0  'Transparent
            Caption         =   "Sp.Atk +1"
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
            Left            =   0
            TabIndex        =   56
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblNatureSpDef 
            BackStyle       =   0  'Transparent
            Caption         =   "Sp.Def +1"
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
            Left            =   0
            TabIndex        =   55
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblNatureSpd 
            BackStyle       =   0  'Transparent
            Caption         =   "Speed +1"
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
            Left            =   0
            TabIndex        =   54
            Top             =   1200
            Width           =   1215
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   4680
         ScaleHeight     =   615
         ScaleWidth      =   1575
         TabIndex        =   43
         Top             =   0
         Width           =   1575
         Begin VB.Label lblnum 
            BackStyle       =   0  'Transparent
            Caption         =   "001"
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
            Left            =   720
            TabIndex        =   44
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   9240
         Top             =   4200
      End
      Begin VB.CommandButton Command7 
         Caption         =   "+"
         Height          =   255
         Left            =   3600
         TabIndex        =   28
         Top             =   4320
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Command6 
         Caption         =   "+"
         Height          =   255
         Left            =   3600
         TabIndex        =   27
         Top             =   3840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Command5 
         Caption         =   "+"
         Height          =   255
         Left            =   3600
         TabIndex        =   26
         Top             =   3360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "+"
         Height          =   255
         Left            =   3600
         TabIndex        =   25
         Top             =   2880
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "+"
         Height          =   255
         Left            =   3600
         TabIndex        =   24
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "+"
         Height          =   255
         Left            =   3600
         TabIndex        =   23
         Top             =   1920
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox lvButton2 
         Height          =   375
         Left            =   6720
         ScaleHeight     =   315
         ScaleWidth      =   1875
         TabIndex        =   29
         Top             =   5760
         Visible         =   0   'False
         Width           =   1935
      End
      Begin lvButton.lvButtons_H lvButtons_H2 
         Height          =   375
         Left            =   6480
         TabIndex        =   30
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Leader"
         CapAlign        =   2
         BackStyle       =   7
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   8421504
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin lvButton.lvButtons_H PokeButton 
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   32
         Top             =   5280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "Set as leader"
         CapAlign        =   2
         BackStyle       =   7
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   8421504
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin lvButton.lvButtons_H PokeButton 
         Height          =   375
         Index           =   6
         Left            =   8160
         TabIndex        =   33
         Top             =   5280
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "Set as leader"
         CapAlign        =   2
         BackStyle       =   7
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   8421504
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin lvButton.lvButtons_H PokeButton 
         Height          =   375
         Index           =   2
         Left            =   1680
         TabIndex        =   34
         Top             =   5280
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "Set as leader"
         CapAlign        =   2
         BackStyle       =   7
         Shape           =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   8421504
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin lvButton.lvButtons_H PokeButton 
         Height          =   375
         Index           =   3
         Left            =   3240
         TabIndex        =   35
         Top             =   5280
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         Caption         =   "Set as leader"
         CapAlign        =   2
         BackStyle       =   7
         Shape           =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   8421504
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin lvButton.lvButtons_H PokeButton 
         Height          =   375
         Index           =   4
         Left            =   4800
         TabIndex        =   36
         Top             =   5280
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Caption         =   "Set as leader"
         CapAlign        =   2
         BackStyle       =   7
         Shape           =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   8421504
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin lvButton.lvButtons_H PokeButton 
         Height          =   375
         Index           =   5
         Left            =   6480
         TabIndex        =   37
         Top             =   5280
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Caption         =   "Set as leader"
         CapAlign        =   2
         BackStyle       =   7
         Shape           =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   8421504
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin lvButton.lvButtons_H lvButtons_H1 
         Height          =   375
         Left            =   8400
         TabIndex        =   38
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Moves"
         CapAlign        =   2
         BackStyle       =   7
         Shape           =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   8421504
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin lvButton.lvButtons_H lvButtons_H3 
         Height          =   375
         Left            =   9360
         TabIndex        =   39
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Evolve"
         CapAlign        =   2
         BackStyle       =   7
         Shape           =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   12632064
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   16776960
      End
      Begin lvButton.lvButtons_H lvButtons_H5 
         Height          =   375
         Left            =   7440
         TabIndex        =   61
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "Items"
         CapAlign        =   2
         BackStyle       =   7
         Shape           =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   8421504
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   12632256
      End
      Begin lvButton.lvButtons_H lvButtons_H7 
         Height          =   300
         Left            =   3600
         TabIndex        =   67
         Top             =   1920
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         Caption         =   "+"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Eurostar"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   12632064
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   16776960
      End
      Begin lvButton.lvButtons_H lvButtons_H8 
         Height          =   300
         Left            =   3600
         TabIndex        =   68
         Top             =   2400
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         Caption         =   "+"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Eurostar"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   12632064
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   16776960
      End
      Begin lvButton.lvButtons_H lvButtons_H9 
         Height          =   300
         Left            =   3600
         TabIndex        =   69
         Top             =   2880
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         Caption         =   "+"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Eurostar"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   12632064
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   16776960
      End
      Begin lvButton.lvButtons_H lvButtons_H10 
         Height          =   300
         Left            =   3600
         TabIndex        =   70
         Top             =   3360
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         Caption         =   "+"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Eurostar"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   12632064
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   16776960
      End
      Begin lvButton.lvButtons_H lvButtons_H11 
         Height          =   300
         Left            =   3600
         TabIndex        =   71
         Top             =   3840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         Caption         =   "+"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Eurostar"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   12632064
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   16776960
      End
      Begin lvButton.lvButtons_H lvButtons_H12 
         Height          =   300
         Left            =   3600
         TabIndex        =   72
         Top             =   4320
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   529
         Caption         =   "+"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Eurostar"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   0
         cFHover         =   0
         cBhover         =   12632064
         LockHover       =   3
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   16776960
      End
      Begin VB.Label lblEXP 
         BackColor       =   &H00000000&
         Caption         =   "0/0"
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   6360
         TabIndex        =   42
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF80&
         Caption         =   "EXP:"
         Height          =   255
         Left            =   5760
         TabIndex        =   40
         Top             =   1680
         Width           =   615
      End
      Begin VB.Image imgType2 
         Height          =   375
         Left            =   6840
         Stretch         =   -1  'True
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Image imgType1 
         Height          =   375
         Left            =   4680
         Stretch         =   -1  'True
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label lblTP 
         BackColor       =   &H00000000&
         Caption         =   "Bulbasaur"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   22
         Top             =   4800
         Width           =   2775
      End
      Begin VB.Label Label12 
         BackColor       =   &H0080FFFF&
         Caption         =   "TP:"
         Height          =   255
         Left            =   0
         TabIndex        =   21
         Top             =   4800
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Hp:"
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblSpeed 
         BackColor       =   &H00000000&
         Caption         =   "Bulbasaur"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   4320
         Width           =   2775
      End
      Begin VB.Label lblSpDef 
         BackColor       =   &H00000000&
         Caption         =   "Bulbasaur"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   18
         Top             =   3840
         Width           =   2775
      End
      Begin VB.Label lblSpAtk 
         BackColor       =   &H00000000&
         Caption         =   "Bulbasaur"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   17
         Top             =   3360
         Width           =   2775
      End
      Begin VB.Label lblDef 
         BackColor       =   &H00000000&
         Caption         =   "Bulbasaur"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   2880
         Width           =   2775
      End
      Begin VB.Label lblAtk 
         BackColor       =   &H00000000&
         Caption         =   "Bulbasaur"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   2400
         Width           =   2775
      End
      Begin VB.Label lblMaxHp 
         BackColor       =   &H00000000&
         Caption         =   "Bulbasaur"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   14
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label lblHp 
         BackColor       =   &H00000000&
         Caption         =   "Bulbasaur"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label lblNature 
         BackColor       =   &H00000000&
         Caption         =   "Bulbasaur"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label lblLevel 
         BackColor       =   &H00000000&
         Caption         =   "Bulbasaur"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label10 
         Caption         =   "HP:"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "SPEED:"
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "SPDEF:"
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "SPATK:"
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "DEF:"
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "ATK:"
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Nature:"
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Level:"
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblName 
         BackColor       =   &H00000000&
         Caption         =   "Bulbasaur"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   120
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   615
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl ShinyImage 
         Height          =   1815
         Left            =   6480
         Top             =   1920
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   3201
         Effects         =   "frmPokemons.frx":906BB
      End
      Begin LaVolpeAlphaImg.AlphaImgCtl imgPokemon 
         Height          =   1815
         Left            =   5760
         Top             =   2160
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   3201
         Attr            =   1536
         Effects         =   "frmPokemons.frx":906D3
      End
   End
   Begin VB.Label Label13 
      BackColor       =   &H00404040&
      Caption         =   "Bulbasaur"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmPokemons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'function to make transparent'

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

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


Private Sub btnRefresh_Click(Index As Integer)
LoadPokemon (selectedpoke)
loadImages
End Sub

Private Sub Command2_Click()
SendAddTP STAT_HP, selectedpoke
End Sub

Private Sub Command3_Click()
SendAddTP STAT_ATK, selectedpoke
End Sub

Private Sub Command4_Click()
SendAddTP STAT_DEF, selectedpoke
End Sub

Private Sub Command5_Click()
SendAddTP STAT_SPATK, selectedpoke
End Sub

Private Sub Command6_Click()
SendAddTP STAT_SPDEF, selectedpoke
End Sub

Private Sub Command7_Click()
SendAddTP STAT_SPEED, selectedpoke
End Sub

Private Sub Command8_Click()
SendSetAsLeader selectedpoke
End Sub

Private Sub Command9_Click()
LoadPokemon (selectedpoke)
loadImages
End Sub

Private Sub Form_Load()
loadImages
selectedpoke = 1
LoadPokemon (selectedpoke)
lblTP.Caption = Val(PokemonInstance(1).TP)
loadImages
LoadPokemon (selectedpoke)

LoadInv
End Sub

Sub loadImages()
Dim i As Long
Dim pn As Long
For i = 1 To 6
If PokemonInstance(i).PokemonNumber > 0 Then
pn = PokemonInstance(i).PokemonNumber
PokeButton(i).Caption = Trim$(Pokemon(pn).Name)
Else
PokeButton(i).Caption = ""
End If
Next
End Sub

Sub LoadPokemon(ByVal num As Long)
If PokemonInstance(num).PokemonNumber > 0 Then
LoadPokeMoves (num)
LoadNatureBoost (num)
Dim str As String
str = PokemonInstance(num).PokemonNumber
If Len(str) = 1 Then
str = "00" & str
End If
If Len(str) = 2 Then
str = "0" & str
End If
If Len(str) = 3 Then
End If
lblnum.Caption = str
If PokemonInstance(num).isShiny = YES Then
imgPokemon.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\Shiny\" & PokemonInstance(num).PokemonNumber & ".gif")
ShinyImage.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\Shiny.gif")

ShinyImage.Animate (lvicAniCmdStart)
Else
imgPokemon.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\" & PokemonInstance(num).PokemonNumber & ".gif")
ShinyImage.Picture = Nothing
End If

lblTP.Caption = Val(PokemonInstance(num).TP)
lblName.Caption = Trim$(Pokemon(PokemonInstance(num).PokemonNumber).Name)
If PokemonInstance(num).nature < 1 Or PokemonInstance(num).nature > MAX_NATURES Then
lblNature.Caption = "None."
Else
lblNature.Caption = Trim$(nature(PokemonInstance(num).nature).Name)
End If
imgType1.Picture = LoadPicture(App.Path & "\Data Files\graphics\types\" & Pokemon(PokemonInstance(num).PokemonNumber).Type & ".bmp")
imgType2.Picture = LoadPicture(App.Path & "\Data Files\graphics\types\" & Pokemon(PokemonInstance(num).PokemonNumber).Type2 & ".bmp")
lblLevel.Caption = Val(PokemonInstance(num).Level)
lblHP.Caption = Val(PokemonInstance(num).HP) & "/" & Val(PokemonInstance(num).MaxHp)
lblMaxHp.Caption = Val(PokemonInstance(num).MaxHp)
lblATK.Caption = Val(PokemonInstance(num).ATK)
lblDEF.Caption = Val(PokemonInstance(num).DEF)
lblSPATK.Caption = Val(PokemonInstance(num).SPATK)
lblSPDEF.Caption = Val(PokemonInstance(num).SPDEF)
lblSPEED.Caption = Val(PokemonInstance(num).SPD)
lblEXP.Caption = Val(PokemonInstance(num).EXP) & "/" & Val(PokemonInstance(num).expNeeded)
Else
End If
LoadInv
End Sub

Sub LoadPokeMoves(ByVal slot As Long)
lstMoves.Clear
Dim i As Long
Dim pokenum As Long
pokenum = PokemonInstance(slot).PokemonNumber
For i = 1 To 4
If PokemonInstance(slot).moves(i).number > 0 Then
lblMove(i).Caption = Trim$(PokemonMove(PokemonInstance(slot).moves(i).number).Name) & vbNewLine & "PP " & PokemonInstance(slot).moves(i).pp
Else
lblMove(i).Caption = "None." & vbNewLine & "PP 0"
End If
Next
For i = 1 To 30
If Pokemon(pokenum).moves(i) > 0 Then
lstMoves.AddItem (Trim$(PokemonMove(Pokemon(pokenum).moves(i)).Name) & " - Lv." & Pokemon(pokenum).movesLV(i))
End If
Next
End Sub

Sub LoadNatureBoost(ByVal slot As Long)
Dim natNum As Long
natNum = PokemonInstance(slot).nature
lblNatureAtk.Caption = "Atk +" & nature(natNum).AddAtk
lblNatureDef.Caption = "Def +" & nature(natNum).AddDef
lblNatureSpAtk.Caption = "Sp.Atk +" & nature(natNum).AddSpAtk
lblNatureSpDef.Caption = "Sp.Def +" & nature(natNum).AddSpDef
lblNatureSpd.Caption = "Speed +" & nature(natNum).AddSpd
lblNatureHP.Caption = "HP +" & nature(natNum).AddHP
End Sub

Private Sub Image2_Click()

End Sub





Private Sub lvButton1_Click()
LoadPokemon (selectedpoke)
loadImages
End Sub

Private Sub lstItems_DblClick()
SendRequest selectedpoke, lstItems.ListIndex + 1, "", "USEITEMONPOKEMON"
End Sub

Private Sub lstMoves_DblClick()
SendRequest selectedpoke, lstMoves.ListIndex + 1, "", "TRYLM"
Unload Me
End Sub

Private Sub lvButton2_Click()
SendSetAsLeader selectedpoke
End Sub

Private Sub lvButtons_H1_Click()
LoadPokeMoves (selectedpoke)
picMoves.Visible = True
picItems.Visible = False
End Sub

Private Sub lvButtons_H10_Click()
SendAddTP STAT_SPATK, selectedpoke
End Sub

Private Sub lvButtons_H11_Click()
SendAddTP STAT_SPDEF, selectedpoke
End Sub

Private Sub lvButtons_H12_Click()
SendAddTP STAT_SPEED, selectedpoke
End Sub

Private Sub lvButtons_H2_Click()
SendSetAsLeader selectedpoke
selectedpoke = 1
LoadPokemon (selectedpoke)
End Sub

Private Sub lvButtons_H3_Click()
Call SendRequest(selectedpoke, 1, "", "TRYEVOLVE")
End Sub

Private Sub lvButtons_H4_Click()
picMoves.Visible = False
End Sub

Private Sub lvButtons_H5_Click()
LoadInv
picItems.Visible = True
picMoves.Visible = False
End Sub

Private Sub lvButtons_H6_Click()
picItems.Visible = False
End Sub

Private Sub lvButtons_H7_Click()
SendAddTP STAT_HP, selectedpoke
End Sub

Private Sub lvButtons_H8_Click()
SendAddTP STAT_ATK, selectedpoke
End Sub

Private Sub lvButtons_H9_Click()
SendAddTP STAT_DEF, selectedpoke
End Sub

Private Sub PokeButton_Click(Index As Integer)
If PokemonInstance(Index).PokemonNumber > 0 Then
selectedpoke = Index
LoadPokemon (selectedpoke)
End If
End Sub

Private Sub Timer1_Timer()
loadImages
End Sub

Public Sub LoadInv()
Dim i As Long
Dim itemnum As Long
Dim itemvalue As Long
lstItems.Clear
For i = 1 To MAX_INV
itemnum = GetPlayerInvItemNum(MyIndex, i)
itemvalue = GetPlayerInvItemValue(MyIndex, i)
If itemvalue = 0 Then itemvalue = 1
If itemnum <= MAX_ITEMS Then
If itemnum = 0 Then
lstItems.AddItem ("Empty")
Else
lstItems.AddItem (Item(itemnum).Name & " x" & itemvalue)
End If
End If
Next
End Sub
