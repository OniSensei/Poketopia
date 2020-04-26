VERSION 5.00
Begin VB.Form frmMapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMapProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   524
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraPoke 
      Caption         =   "Pokemon Data"
      Height          =   7095
      Left            =   6600
      TabIndex        =   60
      Top             =   720
      Width           =   3495
      Begin VB.HScrollBar scrlPokemonDataNum 
         Height          =   255
         Left            =   120
         Max             =   721
         TabIndex        =   81
         Top             =   480
         Width           =   3135
      End
      Begin VB.HScrollBar scrlChance 
         Height          =   255
         Left            =   120
         Max             =   10000
         TabIndex        =   79
         Top             =   6720
         Value           =   1
         Width           =   2295
      End
      Begin VB.TextBox txtHP 
         Height          =   360
         Left            =   1080
         TabIndex        =   77
         Text            =   "100"
         Top             =   5760
         Width           =   1335
      End
      Begin VB.TextBox txtSpd 
         Height          =   360
         Left            =   1080
         TabIndex        =   75
         Text            =   "100"
         Top             =   5280
         Width           =   1335
      End
      Begin VB.TextBox txtSpDef 
         Height          =   360
         Left            =   1080
         TabIndex        =   73
         Text            =   "100"
         Top             =   4800
         Width           =   1335
      End
      Begin VB.TextBox txtSpAtk 
         Height          =   360
         Left            =   1080
         TabIndex        =   71
         Text            =   "100"
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox txtDef 
         Height          =   360
         Left            =   1080
         TabIndex        =   69
         Text            =   "100"
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox txtAtk 
         Height          =   360
         Left            =   1080
         TabIndex        =   67
         Text            =   "100"
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CheckBox CheckCustomStats 
         Caption         =   "Custom stats"
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   3000
         Width           =   2775
      End
      Begin VB.TextBox txtto 
         Height          =   375
         Left            =   2400
         TabIndex        =   64
         Text            =   "To"
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtfrom 
         Height          =   345
         Left            =   960
         TabIndex        =   62
         Text            =   "From"
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label lblPokeName 
         Caption         =   "PokemonName"
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   2040
         Width           =   3135
      End
      Begin VB.Image imgpoke 
         Height          =   1455
         Left            =   120
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label lblpokenum 
         Caption         =   "Pokemon Number:1"
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblChance 
         Caption         =   "Chance: 100%"
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   6360
         Width           =   3135
      End
      Begin VB.Label Label16 
         Caption         =   "HP"
         Height          =   255
         Left            =   240
         TabIndex        =   76
         Top             =   5760
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "SPEED"
         Height          =   255
         Left            =   240
         TabIndex        =   74
         Top             =   5280
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "SP.DEF"
         Height          =   255
         Left            =   240
         TabIndex        =   72
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "SP.ATK"
         Height          =   255
         Left            =   240
         TabIndex        =   70
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "DEF"
         Height          =   255
         Left            =   240
         TabIndex        =   68
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "ATK"
         Height          =   255
         Left            =   240
         TabIndex        =   66
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "to"
         Height          =   255
         Left            =   1920
         TabIndex        =   63
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Level:"
         Height          =   255
         Left            =   240
         TabIndex        =   61
         Top             =   2400
         Width           =   615
      End
   End
   Begin VB.HScrollBar scrlPokeNum 
      Height          =   255
      Left            =   6600
      Max             =   30
      Min             =   1
      TabIndex        =   58
      Top             =   360
      Value           =   1
      Width           =   3495
   End
   Begin VB.Frame frmMaxSizes 
      Caption         =   "Max Sizes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   28
      Top             =   3720
      Width           =   2055
      Begin VB.TextBox txtMaxX 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   30
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtMaxY 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   29
         Text            =   "0"
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Max X:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Max Y:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   630
         Width           =   585
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Map Links"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   21
      Top             =   480
      Width           =   2055
      Begin VB.TextBox txtUp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   720
         TabIndex        =   25
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtDown 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   720
         TabIndex        =   24
         Text            =   "0"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtRight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1320
         TabIndex        =   23
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtLeft 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   22
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblMap 
         BackStyle       =   0  'Transparent
         Caption         =   "Current map: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame fraMapSettings 
      Caption         =   "Map Settings"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2280
      TabIndex        =   16
      Top             =   480
      Width           =   4215
      Begin VB.ComboBox cmbMoral 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMapProperties.frx":020A
         Left            =   960
         List            =   "frmMapProperties.frx":0214
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   360
         Width           =   3135
      End
      Begin VB.HScrollBar scrlMusic 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   17
         Top             =   960
         Width           =   3975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Moral:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   540
      End
      Begin VB.Label lblMusic 
         AutoSize        =   -1  'True
         Caption         =   "Music: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Boot Settings"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   2055
      Begin VB.TextBox txtBootMap 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtBootX 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtBootY 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         Text            =   "0"
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Boot Map:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Boot X:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Boot Y:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   630
      End
   End
   Begin VB.Frame fraNPCs 
      Caption         =   "NPCs"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   2280
      TabIndex        =   4
      Top             =   1920
      Width           =   4215
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   30
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   5400
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   29
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   5400
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   28
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   5040
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   27
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   5040
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   26
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   4680
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   25
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   4680
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   24
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   4320
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   23
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   4320
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   22
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   3960
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   21
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   3960
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   20
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   3600
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   19
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   3600
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   18
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   3240
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   17
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   3240
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   16
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   2880
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   15
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   2880
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   14
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   2520
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   13
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   2520
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   2160
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   2160
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   1800
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   1800
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox cmbNpc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   7440
      Width           =   975
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label3 
      Caption         =   "Pokemon:1"
      Height          =   255
      Left            =   6600
      TabIndex        =   59
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmMapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub CheckCustomStats_Click()
EditPokemons(scrlPokeNum).Custom = CheckCustomStats.Value
End Sub

Private Sub Form_Load()
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    txtName.text = Trim$(map.Name)
    txtUp.text = CStr(map.Up)
    txtDown.text = CStr(map.Down)
    txtLeft.text = CStr(map.Left)
    txtRight.text = CStr(map.Right)
    cmbMoral.ListIndex = map.Moral
    scrlMusic.Value = map.music
    txtBootMap.text = CStr(map.BootMap)
    txtBootX.text = CStr(map.BootX)
    txtBootY.text = CStr(map.BootY)

    For X = 1 To MAX_MAP_NPCS
        cmbNpc(X).AddItem "No NPC"
    Next

    For Y = 1 To MAX_NPCS
        For X = 1 To MAX_MAP_NPCS
            cmbNpc(X).AddItem Y & ": " & Trim$(NPC(Y).Name)
        Next
    Next

    For i = 1 To MAX_MAP_NPCS
        cmbNpc(i).ListIndex = map.NPC(i)
    Next

    lblMap.Caption = "Current map: " & GetPlayerMap(MyIndex)
    txtMaxX.text = map.MaxX
    txtMaxY.text = map.MaxY
    'POKEMON (GOLF)
    loadAllPokemons
    scrlPokeNum.Max = MAX_MAP_POKEMONS
    scrlPokemonDataNum.Max = MAX_POKEMONS
    loadPokemons (1)
End Sub

Private Sub scrlChance_Change()
Dim percent As Double
If scrlChance.Value = 0 Then
Else
percent = 1 / scrlChance.Value
percent = percent * 100
End If

lblChance.Caption = "Chance: 1 out of " & scrlChance.Value & "    (" & percent & "%)"
EditPokemons(scrlPokeNum).Chance = scrlChance.Value
End Sub

Sub loadAllPokemons()
Dim i As Long
For i = 1 To MAX_MAP_POKEMONS
EditPokemons(i) = map.Pokemon(i)
Next
End Sub

Private Sub scrlMusic_Change()
    lblMusic.Caption = CStr(scrlMusic.Value)
    'Call DirectMusic_PlayMidi(scrlMusic.Value, 1)
End Sub

Private Sub cmdOk_Click()
    Dim i As Long
    Dim sTemp As Long
    Dim X As Long, x2 As Long
    Dim Y As Long, y2 As Long
    Dim tempArr() As TileRec
    Dim a As Long
    If Not IsNumeric(txtMaxX.text) Then txtMaxX.text = map.MaxX
    If Val(txtMaxX.text) < MAX_MAPX Then txtMaxX.text = MAX_MAPX
    If Val(txtMaxX.text) > MAX_BYTE Then txtMaxX.text = MAX_BYTE
    If Not IsNumeric(txtMaxY.text) Then txtMaxY.text = map.MaxY
    If Val(txtMaxY.text) < MAX_MAPY Then txtMaxY.text = MAX_MAPY
    If Val(txtMaxY.text) > MAX_BYTE Then txtMaxY.text = MAX_BYTE

    With map
        .Name = Trim$(txtName.text)
        .Up = Val(txtUp.text)
        .Down = Val(txtDown.text)
        .Left = Val(txtLeft.text)
        .Right = Val(txtRight.text)
        .Moral = cmbMoral.ListIndex
        .music = scrlMusic.Value
        .BootMap = Val(txtBootMap.text)
        .BootX = Val(txtBootX.text)
        .BootY = Val(txtBootY.text)
        
        For i = 1 To MAX_MAP_NPCS
            If cmbNpc(i).ListIndex > 0 Then
                sTemp = InStr(1, Trim$(cmbNpc(i).text), ":", vbTextCompare)

                If Len(Trim$(cmbNpc(i).text)) = sTemp Then
                    cmbNpc(i).ListIndex = 0
                End If
            End If
        Next

        For i = 1 To MAX_MAP_NPCS
            .NPC(i) = cmbNpc(i).ListIndex
        Next
        
        For a = 1 To MAX_MAP_POKEMONS
        Select Case EditPokemons(a).Custom
        Case 0
        map.Pokemon(a).PokemonNumber = EditPokemons(a).PokemonNumber
        map.Pokemon(a).LevelFrom = EditPokemons(a).LevelFrom
        map.Pokemon(a).LevelTo = EditPokemons(a).LevelTo
        map.Pokemon(a).Custom = 0
        map.Pokemon(a).ATK = EditPokemons(a).ATK
        map.Pokemon(a).DEF = EditPokemons(a).DEF
        map.Pokemon(a).SPATK = EditPokemons(a).SPATK
        map.Pokemon(a).SPDEF = EditPokemons(a).SPDEF
        map.Pokemon(a).SPD = EditPokemons(a).ATK
        map.Pokemon(a).HP = EditPokemons(a).HP
        map.Pokemon(a).Chance = EditPokemons(a).Chance
        Case 1
        map.Pokemon(a) = EditPokemons(a)
        End Select
        Next
        
        
        
        ' set the data before changing it
        tempArr = map.Tile
        x2 = map.MaxX
        y2 = map.MaxY
        ' change the data
        .MaxX = Val(txtMaxX.text)
        .MaxY = Val(txtMaxY.text)
        ReDim map.Tile(0 To .MaxX, 0 To .MaxY)

        If x2 > .MaxX Then x2 = .MaxX
        If y2 > .MaxY Then y2 = .MaxY

        For X = 0 To x2
            For Y = 0 To y2
                .Tile(X, Y) = tempArr(X, Y)
            Next
        Next

        ClearTempTile
    End With

    Call UpdateDrawMapName
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub scrlPokemonDataNum_Change()
lblpokenum.Caption = "Pokemon Number:" & scrlPokemonDataNum.Value
If scrlPokemonDataNum.Value = 0 Then
lblPokeName.Caption = "None"
Else
lblPokeName.Caption = Trim$(Pokemon(scrlPokemonDataNum.Value).Name)
End If

imgPoke.Picture = LoadPicture(App.Path & "\Data Files\graphics\pokemonsprites\" & scrlPokemonDataNum.Value & ".gif")
EditPokemons(scrlPokeNum.Value).PokemonNumber = scrlPokemonDataNum.Value
End Sub

Private Sub scrlPokeNum_Change()
Label3.Caption = "Pokemon:" & scrlPokeNum.Value
loadPokemons (scrlPokeNum.Value)
End Sub

Sub loadPokemons(ByVal pokeslot As Long)
With EditPokemons(pokeslot)
scrlPokemonDataNum.Value = .PokemonNumber
txtfrom.text = .LevelFrom
txtto.text = .LevelTo
If .Custom = 1 Then
CheckCustomStats.Value = 1
Else
CheckCustomStats.Value = 0
End If
txtAtk.text = .ATK
txtDef.text = .DEF
txtSpAtk.text = .SPATK
txtSpDef.text = .SPDEF
txtSpd.text = .SPD
txtHP.text = .HP
scrlChance.Value = .Chance
imgPoke.Picture = LoadPicture(App.Path & "\Data Files\graphics\pokemonsprites\" & .PokemonNumber & ".gif")

End With


End Sub

Private Sub txtAtk_Change()
EditPokemons(scrlPokeNum).ATK = Val(txtAtk.text)
End Sub

Private Sub txtDef_Change()
EditPokemons(scrlPokeNum).DEF = Val(txtDef.text)
End Sub

Private Sub txtfrom_Change()
EditPokemons(scrlPokeNum).LevelFrom = Val(txtfrom.text)
End Sub

Private Sub txtHP_Change()
EditPokemons(scrlPokeNum).HP = Val(txtHP.text)
End Sub

Private Sub txtSpAtk_Change()
EditPokemons(scrlPokeNum).SPATK = Val(txtSpAtk.text)
End Sub

Private Sub txtSpd_Change()
EditPokemons(scrlPokeNum).SPD = Val(txtSpd.text)
End Sub

Private Sub txtSpDef_Change()
EditPokemons(scrlPokeNum).SPDEF = Val(txtSpDef.text)
End Sub

Private Sub txtto_Change()
EditPokemons(scrlPokeNum).LevelTo = Val(txtto.text)
End Sub
