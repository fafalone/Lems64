VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form fEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lems Level Editor - [No level]"
   ClientHeight    =   8460
   ClientLeft      =   195
   ClientTop       =   615
   ClientWidth     =   9600
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   564
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.StatusBar ucInfo 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   87
      Top             =   8160
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   16881
            MinWidth        =   10584
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdZOrderUp 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8955
      Style           =   1  'Graphical
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   6600
      Width           =   225
   End
   Begin VB.CommandButton cmdZOrderDown 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9180
      Style           =   1  'Graphical
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   6600
      Width           =   225
   End
   Begin VB.CheckBox chkOnTerrain 
      Appearance      =   0  'Flat
      Caption         =   "On terrain"
      Height          =   270
      Left            =   7140
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1425
   End
   Begin ComctlLib.Slider ucScroll 
      Height          =   630
      Left            =   6960
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   5175
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1111
      _Version        =   327682
      BorderStyle     =   1
      LargeChange     =   320
      SmallChange     =   10
      Max             =   1280
      TickStyle       =   2
      TickFrequency   =   0
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   11.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7050
      Style           =   1  'Graphical
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   6255
      Width           =   300
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   11.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7650
      Style           =   1  'Graphical
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   6255
      Width           =   300
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   11.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7350
      Style           =   1  'Graphical
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   6555
      Width           =   300
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   11.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7350
      Style           =   1  'Graphical
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   5955
      Width           =   300
   End
   Begin VB.CheckBox chkBlack 
      Appearance      =   0  'Flat
      Caption         =   "Black"
      Height          =   270
      Left            =   8580
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   7125
      Width           =   855
   End
   Begin VB.CheckBox chkUpsideDown 
      Appearance      =   0  'Flat
      Caption         =   "Upside down"
      Height          =   270
      Left            =   7140
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   7755
      Width           =   1425
   End
   Begin LemsEdit.ucScreen32 ucScreen 
      Height          =   4800
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   8467
      BackColor       =   0
   End
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2550
      Index           =   1
      Left            =   225
      TabIndex        =   1
      Top             =   5400
      Width           =   6450
      Begin VB.TextBox txtSkill 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   7
         Left            =   5205
         MaxLength       =   2
         TabIndex        =   22
         Top             =   2115
         Width           =   480
      End
      Begin VB.TextBox txtSkill 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   6
         Left            =   4485
         MaxLength       =   2
         TabIndex        =   21
         Top             =   2115
         Width           =   480
      End
      Begin VB.TextBox txtSkill 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   5
         Left            =   3765
         MaxLength       =   2
         TabIndex        =   20
         Top             =   2115
         Width           =   480
      End
      Begin VB.TextBox txtSkill 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   4
         Left            =   3045
         MaxLength       =   2
         TabIndex        =   19
         Top             =   2115
         Width           =   480
      End
      Begin VB.TextBox txtSkill 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   3
         Left            =   2325
         MaxLength       =   2
         TabIndex        =   18
         Top             =   2115
         Width           =   480
      End
      Begin VB.TextBox txtSkill 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   2
         Left            =   1605
         MaxLength       =   2
         TabIndex        =   17
         Top             =   2115
         Width           =   480
      End
      Begin VB.TextBox txtSkill 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   1
         Left            =   885
         MaxLength       =   2
         TabIndex        =   16
         Top             =   2115
         Width           =   480
      End
      Begin VB.TextBox txtSkill 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   165
         MaxLength       =   2
         TabIndex        =   15
         Top             =   2115
         Width           =   480
      End
      Begin VB.TextBox txtReleaseRate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3780
         MaxLength       =   2
         TabIndex        =   9
         Top             =   435
         Width           =   510
      End
      Begin VB.TextBox txtPlayingTime 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3780
         MaxLength       =   2
         TabIndex        =   11
         Top             =   855
         Width           =   525
      End
      Begin VB.TextBox txtLemsToBeSaved 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1695
         MaxLength       =   2
         TabIndex        =   7
         Top             =   855
         Width           =   510
      End
      Begin VB.TextBox txtLemsToLetOut 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1695
         MaxLength       =   2
         TabIndex        =   5
         Top             =   435
         Width           =   510
      End
      Begin VB.TextBox txtTitle 
         Height          =   315
         Left            =   585
         MaxLength       =   32
         TabIndex        =   3
         Top             =   15
         Width           =   5655
      End
      Begin VB.TextBox txtScreenStart 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   5640
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   435
         Width           =   600
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title"
         Height          =   255
         Left            =   165
         TabIndex        =   2
         Top             =   60
         Width           =   375
      End
      Begin VB.Label lblSkills 
         Caption         =   "Skills"
         Height          =   180
         Left            =   165
         TabIndex        =   14
         Top             =   1320
         Width           =   1380
      End
      Begin VB.Image imgSkill 
         Height          =   480
         Index           =   7
         Left            =   5205
         Top             =   1605
         Width           =   480
      End
      Begin VB.Image imgSkill 
         Height          =   480
         Index           =   6
         Left            =   4485
         Top             =   1605
         Width           =   480
      End
      Begin VB.Image imgSkill 
         Height          =   480
         Index           =   5
         Left            =   3765
         Top             =   1605
         Width           =   480
      End
      Begin VB.Image imgSkill 
         Height          =   480
         Index           =   4
         Left            =   3045
         Top             =   1605
         Width           =   480
      End
      Begin VB.Image imgSkill 
         Height          =   480
         Index           =   3
         Left            =   2325
         Top             =   1605
         Width           =   480
      End
      Begin VB.Image imgSkill 
         Height          =   480
         Index           =   2
         Left            =   1605
         Top             =   1605
         Width           =   480
      End
      Begin VB.Image imgSkill 
         Height          =   480
         Index           =   1
         Left            =   885
         Top             =   1605
         Width           =   480
      End
      Begin VB.Image imgSkill 
         Height          =   480
         Index           =   0
         Left            =   165
         Top             =   1605
         Width           =   480
      End
      Begin VB.Label lblScreenStart 
         Caption         =   "Screen start"
         Height          =   255
         Left            =   4575
         TabIndex        =   12
         Top             =   480
         Width           =   1065
      End
      Begin VB.Label lblReleaseRate 
         Caption         =   "Release rate"
         Height          =   255
         Left            =   2460
         TabIndex        =   8
         Top             =   480
         Width           =   1380
      End
      Begin VB.Label lblPlayingTime 
         Caption         =   "Playing time "
         Height          =   255
         Left            =   2460
         TabIndex        =   10
         Top             =   900
         Width           =   1395
      End
      Begin VB.Label lblLemsToBeSaved 
         Caption         =   "Lems to be saved"
         Height          =   255
         Left            =   165
         TabIndex        =   6
         Top             =   900
         Width           =   1500
      End
      Begin VB.Label lblLemsToLetOut 
         Caption         =   "Lems to let out"
         Height          =   255
         Left            =   165
         TabIndex        =   4
         Top             =   480
         Width           =   1380
      End
   End
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2550
      Index           =   2
      Left            =   225
      TabIndex        =   23
      Top             =   5400
      Visible         =   0   'False
      Width           =   6450
      Begin LemsEdit.ucScreen32 ucObjectPreview 
         Height          =   2490
         Left            =   1980
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   0
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   4392
         BorderStyle     =   1
         BackColor       =   16711935
      End
      Begin VB.Timer tmrAnimation 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   4515
         Top             =   2070
      End
      Begin VB.ComboBox cbObject 
         Height          =   330
         Left            =   765
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   0
         Width           =   1125
      End
      Begin VB.CommandButton cmdObjectAdd 
         Caption         =   "Add"
         Height          =   450
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1020
         Width           =   1005
      End
      Begin VB.Label lblObjectDescription 
         Height          =   450
         Left            =   150
         TabIndex        =   85
         Top             =   1080
         Width           =   1665
      End
      Begin VB.Label lblObjectWidthVal 
         Height          =   240
         Left            =   795
         TabIndex        =   27
         Top             =   510
         Width           =   810
      End
      Begin VB.Label lblObjectHeightVal 
         Height          =   240
         Left            =   795
         TabIndex        =   29
         Top             =   795
         Width           =   810
      End
      Begin VB.Label lblObject 
         Caption         =   "Object"
         Height          =   270
         Left            =   150
         TabIndex        =   24
         Top             =   60
         Width           =   600
      End
      Begin VB.Label lblObjectWidth 
         Caption         =   "Width:"
         Height          =   240
         Left            =   150
         TabIndex        =   26
         Top             =   510
         Width           =   570
      End
      Begin VB.Label lblObjectHeight 
         Caption         =   "Height:"
         Height          =   240
         Left            =   150
         TabIndex        =   28
         Top             =   795
         Width           =   570
      End
   End
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Index           =   3
      Left            =   225
      TabIndex        =   32
      Top             =   5400
      Visible         =   0   'False
      Width           =   6495
      Begin LemsEdit.ucScreen32 ucTerrainPiecePreview 
         Height          =   2490
         Left            =   1980
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   0
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   4392
         BorderStyle     =   1
         BackColor       =   16711935
      End
      Begin VB.CommandButton cmdTerrainPieceAdd 
         Caption         =   "Add"
         Height          =   450
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1020
         Width           =   1005
      End
      Begin VB.ComboBox cbTerrainPiece 
         Height          =   330
         Left            =   765
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   0
         Width           =   1125
      End
      Begin VB.Label lblTerrainPieceDescription 
         Height          =   450
         Left            =   150
         TabIndex        =   86
         Top             =   1080
         Width           =   1665
      End
      Begin VB.Label lblTerrainPieceHeightVal 
         Height          =   240
         Left            =   795
         TabIndex        =   38
         Top             =   795
         Width           =   810
      End
      Begin VB.Label lblTerrainPieceHeight 
         Caption         =   "Height:"
         Height          =   240
         Left            =   150
         TabIndex        =   37
         Top             =   795
         Width           =   570
      End
      Begin VB.Label lblTerrainPieceWidthVal 
         Height          =   240
         Left            =   795
         TabIndex        =   36
         Top             =   510
         Width           =   810
      End
      Begin VB.Label lblTerrainPieceWidth 
         Caption         =   "Width:"
         Height          =   240
         Left            =   150
         TabIndex        =   35
         Top             =   510
         Width           =   570
      End
      Begin VB.Label lblTerrainPiece 
         Caption         =   "Piece"
         Height          =   270
         Left            =   150
         TabIndex        =   33
         Top             =   60
         Width           =   600
      End
   End
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2550
      Index           =   4
      Left            =   225
      TabIndex        =   41
      Top             =   5400
      Visible         =   0   'False
      Width           =   6450
      Begin VB.CommandButton cmdSteelAreaAdd 
         Caption         =   "Add"
         Height          =   450
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   600
         Width           =   1005
      End
      Begin VB.TextBox txtSteelAreaHeight 
         Height          =   315
         Left            =   795
         MaxLength       =   3
         TabIndex        =   46
         Top             =   840
         Width           =   825
      End
      Begin VB.TextBox txtSteelAreaWidth 
         Height          =   315
         Left            =   795
         MaxLength       =   3
         TabIndex        =   44
         Top             =   465
         Width           =   825
      End
      Begin VB.Label lblSteelAreaHeight 
         Caption         =   "Height"
         Height          =   240
         Left            =   150
         TabIndex        =   45
         Top             =   885
         Width           =   570
      End
      Begin VB.Label lblSteelAreaWidth 
         Caption         =   "Width"
         Height          =   240
         Left            =   150
         TabIndex        =   43
         Top             =   510
         Width           =   570
      End
      Begin VB.Label lblSteelAreaCurrent 
         Caption         =   "Current/New steel area:"
         Height          =   255
         Left            =   150
         TabIndex        =   42
         Top             =   60
         Width           =   1920
      End
   End
   Begin VB.Frame fraOptions 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2550
      Index           =   5
      Left            =   225
      TabIndex        =   48
      Top             =   5400
      Visible         =   0   'False
      Width           =   6450
      Begin VB.CheckBox chkShowTriggerAreas 
         Caption         =   "Show trigger areas"
         Height          =   255
         Left            =   2370
         TabIndex        =   60
         Top             =   1125
         Width           =   1980
      End
      Begin VB.CheckBox chkShowBlackPieces 
         Caption         =   "Show black-pieces"
         Height          =   255
         Left            =   2370
         TabIndex        =   61
         Top             =   1470
         Width           =   1980
      End
      Begin VB.CheckBox chkHighlightSelected 
         Caption         =   "Highlight selected"
         Height          =   255
         Left            =   165
         TabIndex        =   56
         Top             =   1470
         Value           =   1  'Checked
         Width           =   1980
      End
      Begin LemsEdit.ucProgress ucPrgObjects 
         Height          =   315
         Left            =   1095
         Top             =   2160
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         BorderStyle     =   1
         BackColor       =   16777215
         ForeColor       =   0
      End
      Begin VB.CheckBox chkConfirmRemoving 
         Caption         =   "Confirm removing"
         Height          =   255
         Left            =   4545
         TabIndex        =   66
         Top             =   1470
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.ComboBox cbGraphicSet 
         Height          =   330
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Tag             =   "0"
         Top             =   0
         Width           =   915
      End
      Begin VB.CheckBox chkShowSelectionBox 
         Caption         =   "Show selection box"
         Height          =   255
         Left            =   165
         TabIndex        =   55
         Top             =   1125
         Value           =   1  'Checked
         Width           =   1980
      End
      Begin VB.OptionButton optSelectionPreference 
         Caption         =   "Select object"
         Height          =   255
         Index           =   0
         Left            =   4530
         TabIndex        =   63
         Top             =   360
         Value           =   -1  'True
         Width           =   1875
      End
      Begin VB.OptionButton optSelectionPreference 
         Caption         =   "Select terrain"
         Height          =   255
         Index           =   1
         Left            =   4530
         TabIndex        =   64
         Top             =   675
         Width           =   1875
      End
      Begin VB.CheckBox chkShowSteel 
         Caption         =   "Show steel"
         Height          =   255
         Left            =   2370
         TabIndex        =   59
         Top             =   765
         Value           =   1  'Checked
         Width           =   1980
      End
      Begin VB.CheckBox chkShowObjects 
         Caption         =   "Show objects"
         Height          =   255
         Left            =   2370
         TabIndex        =   57
         Top             =   45
         Value           =   1  'Checked
         Width           =   1980
      End
      Begin VB.CheckBox chkShowTerrain 
         Caption         =   "Show terrain"
         Height          =   255
         Left            =   2370
         TabIndex        =   58
         Top             =   405
         Value           =   1  'Checked
         Width           =   1980
      End
      Begin VB.OptionButton optSelectionPreference 
         Caption         =   "Select steel"
         Height          =   255
         Index           =   2
         Left            =   4530
         TabIndex        =   65
         Top             =   990
         Width           =   1875
      End
      Begin LemsEdit.ucProgress ucPrgTerrainPieces 
         Height          =   315
         Left            =   2880
         Top             =   2160
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         BorderStyle     =   1
         BackColor       =   16777215
         ForeColor       =   0
      End
      Begin LemsEdit.ucProgress ucPrgSteelAreas 
         Height          =   315
         Left            =   4680
         Top             =   2160
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         BorderStyle     =   1
         BackColor       =   16777215
         ForeColor       =   0
      End
      Begin VB.Label lblTerrainPieces 
         Caption         =   "Terrain pieces:"
         Height          =   255
         Left            =   165
         TabIndex        =   53
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblSelectionPreference 
         Caption         =   "Selection preference:"
         Height          =   255
         Left            =   4545
         TabIndex        =   62
         Top             =   60
         Width           =   1875
      End
      Begin VB.Label lblTerrainPiecesVal 
         Height          =   255
         Left            =   1425
         TabIndex        =   54
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblObjectsVal 
         Height          =   255
         Left            =   1425
         TabIndex        =   52
         Top             =   405
         Width           =   735
      End
      Begin VB.Label lblObjects 
         Caption         =   "Objects:"
         Height          =   255
         Left            =   165
         TabIndex        =   51
         Top             =   405
         Width           =   1125
      End
      Begin VB.Label lblStatistics 
         Caption         =   "Statistics:"
         Height          =   285
         Left            =   165
         TabIndex        =   67
         Top             =   2205
         Width           =   810
      End
      Begin VB.Line lnSep 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   6450
         Y1              =   1935
         Y2              =   1935
      End
      Begin VB.Line lnSep 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   6450
         Y1              =   1950
         Y2              =   1950
      End
      Begin VB.Label lblGraphicSet 
         Caption         =   "Graphic Set"
         Height          =   255
         Left            =   165
         TabIndex        =   49
         Top             =   60
         Width           =   1800
      End
   End
   Begin ComctlLib.TabStrip tbOptions 
      Height          =   3225
      Left            =   75
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   4875
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   5689
      TabWidthStyle   =   2
      TabFixedWidth   =   2249
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   5
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Level features"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Object"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Terrain"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Steel"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Preferences"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkNotOverlap 
      Appearance      =   0  'Flat
      Caption         =   "Not overlap"
      Height          =   270
      Left            =   7140
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   7125
      Width           =   1425
   End
   Begin VB.Shape shpFrameOptions 
      Height          =   1095
      Left            =   6960
      Top             =   7005
      Width           =   2535
   End
   Begin VB.Image iSelectionSize 
      Height          =   240
      Left            =   8160
      Top             =   6255
      Width           =   240
   End
   Begin VB.Image iSelectionPosition 
      Height          =   240
      Left            =   8160
      Top             =   5910
      Width           =   240
   End
   Begin VB.Image iZOrder 
      Height          =   240
      Left            =   8160
      Top             =   6615
      Width           =   240
   End
   Begin VB.Label lblzOrderVal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8475
      TabIndex        =   73
      Top             =   6600
      Width           =   450
   End
   Begin VB.Label lblSelectionSizeVal 
      Height          =   270
      Left            =   8490
      TabIndex        =   72
      Top             =   6255
      Width           =   915
   End
   Begin VB.Label lblSelectionPositionVal 
      Height          =   270
      Left            =   8490
      TabIndex        =   71
      Top             =   5910
      Width           =   915
   End
   Begin VB.Label lblScreenOriginVal 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   225
      Left            =   8250
      TabIndex        =   69
      Top             =   4935
      Width           =   1065
   End
   Begin VB.Label lblScreenStartSlider 
      Caption         =   "Screen start"
      Height          =   225
      Left            =   7155
      TabIndex        =   68
      Top             =   4935
      Width           =   1065
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Load level..."
         Index           =   0
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save level..."
         Index           =   1
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "C&heck level"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Clear"
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   6
      End
   End
   Begin VB.Menu mnuHelpTop 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&About"
         Index           =   0
      End
   End
   Begin VB.Menu mnuContextSelectionTop 
      Caption         =   "Context selection"
      Visible         =   0   'False
      Begin VB.Menu mnuContextSelection 
         Caption         =   "&Duplicate"
         Index           =   0
      End
      Begin VB.Menu mnuContextSelection 
         Caption         =   "&Remove"
         Index           =   1
      End
      Begin VB.Menu mnuContextSelection 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuContextSelection 
         Caption         =   "Bring to &top"
         Index           =   3
      End
      Begin VB.Menu mnuContextSelection 
         Caption         =   "Bring to &bottom"
         Index           =   4
      End
      Begin VB.Menu mnuContextSelection 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuContextSelection 
         Caption         =   "Not overlap"
         Index           =   6
      End
      Begin VB.Menu mnuContextSelection 
         Caption         =   "On terrain"
         Index           =   7
      End
      Begin VB.Menu mnuContextSelection 
         Caption         =   "Upside down"
         Index           =   8
      End
      Begin VB.Menu mnuContextSelection 
         Caption         =   "Not overlap"
         Index           =   9
      End
      Begin VB.Menu mnuContextSelection 
         Caption         =   "Black"
         Index           =   10
      End
      Begin VB.Menu mnuContextSelection 
         Caption         =   "Upside down"
         Index           =   11
      End
      Begin VB.Menu mnuContextSelection 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuContextSelection 
         Caption         =   "Cancel"
         Index           =   13
      End
   End
End
Attribute VB_Name = "fEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================
' Project:       Lems Level Editor
' Author:        Carles P.V. ©2005-2011
' Dependencies:  -
' First release: 2005.07.08
' Last revision: 2011.04.15
'================================================

Option Explicit

'========================================================================================
' GUI initialization
'========================================================================================

Private Sub Form_Load()
    
    '-- Skill icons
    Set imgSkill(0) = VB.LoadPicture(mMisc.AppPath & "RES\Climber.ico")
    Set imgSkill(1) = VB.LoadPicture(mMisc.AppPath & "RES\Floater.ico")
    Set imgSkill(2) = VB.LoadPicture(mMisc.AppPath & "RES\Bomber.ico")
    Set imgSkill(3) = VB.LoadPicture(mMisc.AppPath & "RES\Blocker.ico")
    Set imgSkill(4) = VB.LoadPicture(mMisc.AppPath & "RES\Builder.ico")
    Set imgSkill(5) = VB.LoadPicture(mMisc.AppPath & "RES\Basher.ico")
    Set imgSkill(6) = VB.LoadPicture(mMisc.AppPath & "RES\Miner.ico")
    Set imgSkill(7) = VB.LoadPicture(mMisc.AppPath & "RES\Digger.ico")
    
    '-- Item info
    Set iSelectionPosition = VB.LoadPicture(mMisc.AppPath & "RES\Position.ico")
    Set iSelectionSize = VB.LoadPicture(mMisc.AppPath & "RES\Size.ico")
    Set iZOrder = VB.LoadPicture(mMisc.AppPath & "RES\ZOrder.ico")
    
    '-- Style changes
    Call mMisc.SetButtonOwnerDraw(cmdObjectAdd, False)
    Call mMisc.SetButtonOwnerDraw(cmdTerrainPieceAdd, False)
    Call mMisc.SetButtonOwnerDraw(cmdSteelAreaAdd, False)
    Call mMisc.SetButtonOwnerDraw(cmdUp, False)
    Call mMisc.SetButtonOwnerDraw(cmdDown, False)
    Call mMisc.SetButtonOwnerDraw(cmdLeft, False)
    Call mMisc.SetButtonOwnerDraw(cmdRight, False)
    Call mMisc.SetButtonOwnerDraw(cmdZOrderUp, False)
    Call mMisc.SetButtonOwnerDraw(cmdZOrderDown, False)
    
    '-- Initialize game screen
    With ucScreen
        .EraseBackground = False
        .WorkMode = [eUserMode]
        .Zoom = 2
        Set .UserIcon = VB.LoadResPicture("CUR_HAND", vbResCursor)
        Call .DIB.Create(320, 160)
        Call .Resize
    End With
    
    '-- Initialize object and terrain piece screens
    ucObjectPreview.Zoom = 2
    ucTerrainPiecePreview.Zoom = 2
    
    '-- Initialize statistics
    ucPrgObjects.Max = MAX_OBJECTS
    ucPrgTerrainPieces.Max = MAX_TERRAINPIECES
    ucPrgSteelAreas.Max = MAX_STEELAREAS
    
    '-- Initialize editor + info
    Call mEdit.Initialize
    Call mEdit.InitializeInfo
    
    '-- Initialize Graphic Set and set default level path (first one)
    Call pvInitializeGraphicSet
    sLEVELPATH = mMisc.AppPath & "Levels\0000.dat"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
  Dim lRet As VBA.VbMsgBoxResult
    
    '-- Ask for saving
    lRet = VBA.MsgBox( _
           "Close Editor?", _
           vbInformation Or vbYesNo _
           )
    
    Select Case lRet
        Case [vbYes]
            '-- Unload
            Call Unload(Me)
        Case [vbNo]
            '-- Cancel action
            Cancel = 1
    End Select
End Sub

'========================================================================================
' Load/Save level
'========================================================================================

Private Sub mnuFile_Click(Index As Integer)
    
  Dim sRet As String
  
    Select Case Index
        
        Case 0 '-- Load level
                
            sRet = mDialogFile.GetFileName( _
                   Me.hWnd, _
                   sLEVELPATH, _
                   "Level files|*.dat", _
                   , _
                   "Load level", _
                   OpenDialog:=True _
                   )
            
            If (sRet <> vbNullString) Then
            
                If (mMain.IsExtendedLevel(sRet)) Then
                    
                    Call VBA.MsgBox( _
                         "Extended level: edition not supported.", _
                         vbInformation _
                         )
                    
                  Else
                    
                    sLEVELPATH = sRet
                    Call pvUpdateAppCaption
                    Call pvUpdateLevelPath
                    Call mMain.LoadLevel(sRet)
                    Call mEdit.InitializeInfo
                    
                    ucScroll.Value = uLEVEL.ScreenStart
                    Call mEdit.DoFrame
                End If
            End If
            
        Case 1 '-- Save (as)
            
            sRet = mDialogFile.GetFileName( _
                   Me.hWnd, _
                   sLEVELPATH, _
                   "Level files|*.dat", _
                   , _
                   "Save level", _
                   OpenDialog:=False _
                   )
           
            If (sRet <> vbNullString) Then
            
                sLEVELPATH = sRet

                Call pvUpdateAppCaption
                Call pvUpdateLevelPath
                Call mMain.SaveLevel(sLEVELPATH)
            End If
            
        Case 3 '-- Check level...
            
            Call VBA.MsgBox( _
                 mMain.CheckLevel(), _
                 vbInformation _
                 )
            
        Case 4 '-- Clear
        
            If (VBA.MsgBox( _
                "Are you sure you want to clear level?", _
                vbExclamation Or vbYesNo _
                ) = vbYes _
                ) Then
                If (VBA.MsgBox( _
                    "Reset level features?", _
                    vbExclamation Or vbYesNo _
                    ) = vbYes _
                    ) Then
                    Call mMain.ClearLevel(ClearLevelFeatures:=True)
                  Else
                    Call mMain.ClearLevel(ClearLevelFeatures:=False)
                End If
                Call mEdit.InitializeInfo
                Call mEdit.DoFrame
            End If
        
        Case 6 '-- Exit
        
            Call Unload(Me)
    End Select
End Sub

Private Sub mnuHelp_Click(Index As Integer)
           
    Call VBA.MsgBox( _
         "Lems Level Editor " & App.Major & "." & App.Minor & "." & App.Revision & Space$(10) & vbCrLf & vbCrLf & _
         "Carles P.V. ©2005-2011", _
         vbInformation _
         )
End Sub

'========================================================================================
' Screen scroll
'========================================================================================

Private Sub ucScreen_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
    Call mEdit.MouseDown(Button, Shift, x, y)
End Sub

Private Sub ucScreen_MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
    Call mEdit.MouseMove(Button, Shift, x, y)
End Sub

Private Sub ucScroll_Change()
    
    Call mEdit.DoScrollTo(ucScroll.Value)
    
    txtScreenStart.Text = ucScroll.Value
    lblScreenOriginVal.Caption = ucScroll.Value
End Sub

Private Sub ucScroll_Scroll()
    Call ucScroll_Change
End Sub

'========================================================================================
' Object/terrain/steel selection/scroll
'========================================================================================

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn
            Call mnuContextSelection_Click(0)
        Case vbKeyDelete
            Call mnuContextSelection_Click(1)
        Case vbKeyPageUp, vbKeyA
            Call mEdit.SelectionFindNextOver
        Case vbKeyPageDown, vbKeyZ
            Call mEdit.SelectionFindNextUnder
    End Select
End Sub

Private Sub ucScreen_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyUp
            Call cmdUp_Click
        Case vbKeyDown
            Call cmdDown_Click
        Case vbKeyLeft
            Call cmdLeft_Click
        Case vbKeyRight
            Call cmdRight_Click
        Case vbKeyEscape
            Call mEdit.RestoreSelectionPosition
    End Select
End Sub

Private Sub cmdUp_Click()
    Call mEdit.SelectionMove(0, -1)
End Sub

Private Sub cmdDown_Click()
    Call mEdit.SelectionMove(0, 1)
End Sub

Private Sub cmdLeft_Click()
    Call mEdit.SelectionMove(-1, 0)
End Sub

Private Sub cmdRight_Click()
    Call mEdit.SelectionMove(1, 0)
End Sub

'========================================================================================
' Selection context menu
'========================================================================================

Private Sub mnuContextSelection_Click(Index As Integer)
    
    Select Case Index
        
        Case 0 '-- Duplicate
            
            If (mEdit.SelectionDuplicate = False) Then
                Call VBA.MsgBox( _
                     "Unable to duplicate current selection.", _
                     vbExclamation _
                     )
            End If
            
        Case 1  '-- Remove
            
            If (mEdit.SelectionExists) Then
                If (chkConfirmRemoving) Then
                    If (VBA.MsgBox( _
                        "Remove current selection?", _
                        vbInformation Or vbYesNo _
                        ) = vbYes _
                        ) Then
                        Call mEdit.SelectionRemove
                    End If
                  Else
                    Call mEdit.SelectionRemove
                End If
            End If
        
        Case 3  '-- Bring to top
            
            Call mEdit.SelectionBringToTop
        
        Case 4  '-- Bring to bottom
            
            Call mEdit.SelectionBringToBottom
            
        Case 6  '-- Not overlap
                
            mnuContextSelection(6).Checked = Not mnuContextSelection(6).Checked
            chkNotOverlap.Value = -mnuContextSelection(6).Checked
        
        Case 7  '-- On terrain
        
            mnuContextSelection(7).Checked = Not mnuContextSelection(7).Checked
            chkOnTerrain.Value = -mnuContextSelection(7).Checked
        
        Case 9  '-- Not overlap
        
            mnuContextSelection(9).Checked = Not mnuContextSelection(9).Checked
            chkNotOverlap.Value = -mnuContextSelection(9).Checked
        
        Case 10 '-- Black
        
            mnuContextSelection(10).Checked = Not mnuContextSelection(10).Checked
            chkBlack.Value = -mnuContextSelection(10).Checked
        
        Case 11  '-- Upside down
        
            mnuContextSelection(11).Checked = Not mnuContextSelection(11).Checked
            chkUpsideDown.Value = -mnuContextSelection(11).Checked
    End Select
End Sub

'========================================================================================
' Options tabs
'========================================================================================

Private Sub tbOptions_Click()
    
  Dim i As Integer

    For i = 1 To fraOptions.Count
        fraOptions(i).Visible = False
    Next i
    fraOptions(tbOptions.SelectedItem.Index).Visible = True
    
    Select Case tbOptions.SelectedItem.Index
        Case 1
            Call txtTitle.SetFocus
        Case 2
            Call cbObject.SetFocus
        Case 3
            Call cbTerrainPiece.SetFocus
        Case 4
            Call txtSteelAreaWidth.SetFocus
        Case 5
            Call cbGraphicSet.SetFocus
    End Select
End Sub

'========================================================================================
' Level features
'========================================================================================

Private Sub txtTitle_Change()
    uLEVEL.Title = txtTitle.Text
End Sub

Private Sub txtLemsToLetOut_Change()
    uLEVEL.LemsToLetOut = Val(txtLemsToLetOut.Text)
End Sub

Private Sub txtLemsToBeSaved_Change()
    uLEVEL.LemsToBeSaved = Val(txtLemsToBeSaved.Text)
End Sub

Private Sub txtReleaseRate_Change()
    uLEVEL.ReleaseRate = Val(txtReleaseRate.Text)
End Sub

Private Sub txtPlayingTime_Change()
    uLEVEL.PlayingTime = Val(txtPlayingTime.Text)
End Sub

Private Sub txtScreenStart_Change()
    uLEVEL.ScreenStart = Val(txtScreenStart.Text)
End Sub

Private Sub txtSkill_Change(Index As Integer)
    
  Dim nVal As Integer
  
    nVal = Val(txtSkill(Index).Text)
    
    With uLEVEL
        Select Case Index
            Case 0: .MaxClimbers = nVal
            Case 1: .MaxFloaters = nVal
            Case 2: .MaxBombers = nVal
            Case 3: .MaxBlockers = nVal
            Case 4: .MaxBuilders = nVal
            Case 5: .MaxBashers = nVal
            Case 6: .MaxMiners = nVal
            Case 7: .MaxDiggers = nVal
        End Select
    End With
End Sub

'========================================================================================
' Object
'========================================================================================

Private Sub cbObject_Click()
    
    With ucObjectPreview

        Call .DIB.Create( _
             uOBJGFX(cbObject.ListIndex).Width, _
             uOBJGFX(cbObject.ListIndex).Height _
             )
        Call .DIB.LoadBlt( _
             uOBJGFX(cbObject.ListIndex).DIB.hDC, _
             0, 0 _
             )
        Call .Resize
        Call .Refresh

        '-- Info
        lblObjectWidthVal.Caption = .DIB.Width
        lblObjectHeightVal.Caption = .DIB.Height
        
        Select Case cbObject.ListIndex
            Case Is = 0
                lblObjectDescription.Caption = "Exit"
            Case Is = 1
                lblObjectDescription.Caption = "Start"
            Case Else
                Select Case uOBJGFX(cbObject.ListIndex).TriggerEffect
                    Case Is = 2
                        lblObjectDescription.Caption = "One-way pointing left"
                    Case Is = 3
                        lblObjectDescription.Caption = "One-way pointing right"
                    Case Is = 32
                        lblObjectDescription.Caption = "Trap"
                    Case Is = 48
                        lblObjectDescription.Caption = "Mortal liquid/ground"
                    Case Is = 64
                        lblObjectDescription.Caption = "Mortal fire/gadget"
                    Case Else
                        lblObjectDescription.Caption = vbNullString
                End Select
        End Select
    End With

    tmrAnimation.Tag = 0
End Sub

Private Sub cmdObjectAdd_Click()
    If (mEdit.AddObject( _
        cbObject.ListIndex _
        ) = False _
        ) Then
        Call VBA.MsgBox( _
             "Error adding object.", _
             vbExclamation _
             )
    End If
End Sub

Private Sub tmrAnimation_Timer()
    
  Static nListIndexOld As Integer

    If (cbObject.ListIndex <> -1) Then

        If (nListIndexOld <> cbObject.ListIndex) Then
            nListIndexOld = cbObject.ListIndex
            Call cbObject_Click
            Exit Sub
        End If

        tmrAnimation.Tag = tmrAnimation.Tag + 1
        If (tmrAnimation.Tag = uOBJGFX(cbObject.ListIndex).EndAnimationFrame) Then
            tmrAnimation.Tag = 0
        End If
        With ucObjectPreview
            Call .DIB.LoadBlt( _
                 uOBJGFX(cbObject.ListIndex).DIB.hDC, _
                 0, tmrAnimation.Tag * .DIB.Height, .DIB.Width, .DIB.Height _
                 )
                 
            Call .Refresh
        End With
    End If
End Sub

'========================================================================================
' Terrain
'========================================================================================

Private Sub cbTerrainPiece_Click()

    With ucTerrainPiecePreview
        
        Call .DIB.Create( _
             uTERGFX(cbTerrainPiece.ListIndex).Width, _
             uTERGFX(cbTerrainPiece.ListIndex).Height _
             )
        Call .DIB.LoadBlt( _
             uTERGFX(cbTerrainPiece.ListIndex).DIB.hDC, _
             0, 0 _
             )
        Call .Resize
        Call .Refresh
        
        '-- Info
        lblTerrainPieceWidthVal.Caption = .DIB.Width
        lblTerrainPieceHeightVal.Caption = .DIB.Height
        lblTerrainPieceDescription.Caption = vbNullString
    End With
End Sub

Private Sub cmdTerrainPieceAdd_Click()
    If (mEdit.AddTerrainPiece( _
        cbTerrainPiece.ListIndex _
        ) = False _
        ) Then
        Call VBA.MsgBox( _
             "Error adding terrain piece.", _
             vbExclamation _
             )
    End If
End Sub

'========================================================================================
' Steel
'========================================================================================

Private Sub txtSteelAreaWidth_Change()
    Call mEdit.SteelAreaSetWidth(Val(txtSteelAreaWidth.Text))
End Sub

Private Sub txtSteelAreaHeight_Change()
    Call mEdit.SteelAreaSetHeight(Val(txtSteelAreaHeight.Text))
End Sub

Private Sub cmdSteelAreaAdd_Click()
    If (mEdit.AddSteelArea( _
        Val(txtSteelAreaWidth.Text), _
        Val(txtSteelAreaHeight.Text) _
        ) = False _
        ) Then
        Call VBA.MsgBox( _
             "Error adding steel area.", _
             vbExclamation _
             )
    End If
End Sub

'========================================================================================
' Preferences
'========================================================================================

Private Sub cbGraphicSet_Click()
    If (cbGraphicSet.Tag = 0) Then
        Call mMain.SetGraphicSet(cbGraphicSet.ListIndex)
        Call mEdit.ResetSelection
        Call mEdit.DoFrame
    End If
End Sub

Private Sub chkShowObjects_Click()
    mEdit.ShowObjects = CBool(chkShowObjects)
End Sub

Private Sub chkShowTerrain_Click()
    mEdit.ShowTerrain = CBool(chkShowTerrain)
End Sub

Private Sub chkShowSteel_Click()
    mEdit.ShowSteel = CBool(chkShowSteel)
End Sub

Private Sub chkShowTriggerAreas_Click()
    mEdit.ShowTriggerAreas = CBool(chkShowTriggerAreas)
End Sub

Private Sub chkHighlightSelected_Click()
    mEdit.HighlightSelected = CBool(chkHighlightSelected)
End Sub

Private Sub chkShowBlackPieces_Click()
    mEdit.ShowBlackPieces = CBool(chkShowBlackPieces)
End Sub

Private Sub chkShowSelectionBox_Click()
    mEdit.ShowSelectionBox = CBool(chkShowSelectionBox)
End Sub

Private Sub optSelectionPreference_Click(Index As Integer)
    mEdit.SelectionPreference = Index
End Sub

'========================================================================================
' Selection Z-Order special and flags
'========================================================================================

Private Sub cmdZOrderUp_Click()
    Call mEdit.SelectionZOrderUp
End Sub

Private Sub cmdZOrderDown_Click()
    Call mEdit.SelectionZOrderDown
End Sub

Private Sub chkNotOverlap_Click()
    Call mEdit.SetNotOverlap(chkNotOverlap)
End Sub

Private Sub chkOnTerrain_Click()
    Call mEdit.ObjectSetOnTerrain(chkOnTerrain)
End Sub

Private Sub chkBlack_Click()
    Call mEdit.TerrainPieceSetBlack(chkBlack)
End Sub

Private Sub chkUpsideDown_Click()
    Call mEdit.TerrainPieceSetUpsideDown(chkUpsideDown)
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub pvInitializeGraphicSet()
  
  Dim i As Long
    
    With Me
    
        '-- Fill combo
        For i = 0 To MAX_GRAPHICSET
            Call .cbGraphicSet.AddItem("# " & i)
        Next i
        
        '-- Select first set
        .cbGraphicSet.ListIndex = 0
    End With
End Sub

Private Sub pvUpdateAppCaption()
    
    Me.Caption = "Lems Level Editor - ["
    Me.Caption = Me.Caption & Mid$(sLEVELPATH, InStrRev(sLEVELPATH, "\") + 1, Len(sLEVELPATH) - InStrRev(sLEVELPATH, "\"))
    Me.Caption = Me.Caption & "]"
End Sub
    
Private Sub pvUpdateLevelPath()

    Me.ucInfo.Panels(1).Text = mMisc.CompactPath(Me.hDC, sLEVELPATH, ucInfo.Panels(1).Width)
End Sub
