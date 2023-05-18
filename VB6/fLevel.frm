VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form fLevel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose level"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   356
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox imgSkill 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   7
      Left            =   4950
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3915
      Width           =   240
   End
   Begin VB.PictureBox imgSkill 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   6
      Left            =   4530
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3915
      Width           =   240
   End
   Begin VB.PictureBox imgSkill 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   5
      Left            =   4185
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3915
      Width           =   240
   End
   Begin VB.PictureBox imgSkill 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   4
      Left            =   3840
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3915
      Width           =   240
   End
   Begin VB.PictureBox imgSkill 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   3
      Left            =   3495
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3915
      Width           =   240
   End
   Begin VB.PictureBox imgSkill 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   2
      Left            =   3150
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3915
      Width           =   240
   End
   Begin VB.PictureBox imgSkill 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   2760
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3915
      Width           =   240
   End
   Begin VB.PictureBox imgSkill 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   2490
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3915
      Width           =   240
   End
   Begin Lems.ucScreen08 ucThumbnail 
      Height          =   630
      Left            =   285
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2970
      Width           =   4830
      _extentx        =   8520
      _extenty        =   1111
      backcolor       =   0
   End
   Begin ComctlLib.TreeView tvLevels 
      Height          =   2625
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   4630
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   423
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ilLevels"
      Appearance      =   1
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4230
      TabIndex        =   18
      Top             =   4785
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3075
      TabIndex        =   17
      Top             =   4785
      Width           =   1050
   End
   Begin VB.Line lnSep1 
      BorderColor     =   &H80000010&
      X1              =   8
      X2              =   352
      Y1              =   309
      Y2              =   309
   End
   Begin VB.Line lnSep2 
      BorderColor     =   &H80000014&
      X1              =   351
      X2              =   7
      Y1              =   310
      Y2              =   310
   End
   Begin ComctlLib.ImageList ilLevels 
      Left            =   0
      Top             =   4770
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      _Version        =   327682
   End
   Begin VB.Label lblSkills 
      Caption         =   "Skills:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2505
      TabIndex        =   8
      Top             =   3705
      Width           =   480
   End
   Begin VB.Label lblSkill 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   4875
      TabIndex        =   16
      Top             =   4245
      Width           =   300
   End
   Begin VB.Label lblSkill 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   4530
      TabIndex        =   15
      Top             =   4245
      Width           =   300
   End
   Begin VB.Label lblSkill 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   4185
      TabIndex        =   14
      Top             =   4245
      Width           =   300
   End
   Begin VB.Label lblSkill 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   13
      Top             =   4245
      Width           =   300
   End
   Begin VB.Label lblSkill 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3495
      TabIndex        =   12
      Top             =   4245
      Width           =   300
   End
   Begin VB.Label lblSkill 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3150
      TabIndex        =   11
      Top             =   4245
      Width           =   300
   End
   Begin VB.Label lblSkill 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2805
      TabIndex        =   10
      Top             =   4245
      Width           =   300
   End
   Begin VB.Label lblSkill 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2460
      TabIndex        =   9
      Top             =   4245
      Width           =   300
   End
   Begin VB.Label lblPlayingTime 
      Caption         =   "Playing time:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   285
      TabIndex        =   6
      Top             =   4245
      Width           =   1110
   End
   Begin VB.Label lblLemsToBeSaved 
      Caption         =   "Lems to be saved:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   285
      TabIndex        =   4
      Top             =   3975
      Width           =   1545
   End
   Begin VB.Label lblLemsToLetOut 
      Caption         =   "Lems to let out:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   285
      TabIndex        =   2
      Top             =   3705
      Width           =   1545
   End
   Begin VB.Label lblPlayingTimeVal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1440
      TabIndex        =   7
      Top             =   4245
      Width           =   660
   End
   Begin VB.Label lblLemsToBeSavedVal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1740
      TabIndex        =   5
      Top             =   3975
      Width           =   360
   End
   Begin VB.Label lblLemsToLetOutVal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1740
      TabIndex        =   3
      Top             =   3705
      Width           =   360
   End
End
Attribute VB_Name = "fLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const TV_FIRST         As Long = &H1100
Private Const TVM_SETBKCOLOR   As Long = TV_FIRST + 29
Private Const TVM_SETTEXTCOLOR As Long = TV_FIRST + 30

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long



'========================================================================================
' Main
'========================================================================================

Private Sub Form_Load()

    '-- No icon
    Set Me.Icon = Nothing
    
    '-- Form cursor
    Set Me.MouseIcon = VB.LoadResPicture("CUR_HAND", vbResCursor)
    
    '-- Image-list images
    Call Me.ilLevels.ListImages.Add(, , VB.LoadPicture(AppPath & "RES\Folder_blue.gif"))
    Call Me.ilLevels.ListImages.Add(, , VB.LoadPicture(AppPath & "RES\Folder_blue.gif"))
    Call Me.ilLevels.ListImages.Add(, , VB.LoadPicture(AppPath & "RES\Checkbox_unchecked.gif"))
    Call Me.ilLevels.ListImages.Add(, , VB.LoadPicture(AppPath & "RES\Checkbox_checked.gif"))
    
    '-- Change treeview colors (masking problems with imagelist)
    Call SendMessage(Me.tvLevels.hwnd, TVM_SETBKCOLOR, 0, vbWhite)
    Call SendMessage(Me.tvLevels.hwnd, TVM_SETTEXTCOLOR, 0, vbBlack)
    
    '-- Level preview drawing offsets
    ucThumbnail.xOffset = 1
    ucThumbnail.yOffset = 1
    
    '-- Skill images
    
    Set imgSkill(0).Picture = VB.LoadPicture(AppPath & "RES\Climber.ico")
    ZoomPicture imgSkill(0), 1
    Set imgSkill(1).Picture = VB.LoadPicture(AppPath & "RES\Floater.ico")
    ZoomPicture imgSkill(1), 1
    Set imgSkill(2).Picture = VB.LoadPicture(AppPath & "RES\Bomber.ico")
    ZoomPicture imgSkill(2), 1
    Set imgSkill(3).Picture = VB.LoadPicture(AppPath & "RES\Blocker.ico")
    ZoomPicture imgSkill(3), 1
    Set imgSkill(4).Picture = VB.LoadPicture(AppPath & "RES\Builder.ico")
    ZoomPicture imgSkill(4), 1
    Set imgSkill(5).Picture = VB.LoadPicture(AppPath & "RES\Basher.ico")
    ZoomPicture imgSkill(5), 1
    Set imgSkill(6).Picture = VB.LoadPicture(AppPath & "RES\Miner.ico")
    ZoomPicture imgSkill(6), 1
    Set imgSkill(7).Picture = VB.LoadPicture(AppPath & "RES\Digger.ico")
    ZoomPicture imgSkill(7), 1
    '-- Fill treeview with all levels
    Screen.MousePointer = vbHourglass
    Call pvShowAllLevels
    Screen.MousePointer = vbDefault
    
    '-- Select current level ID
    Call pvSelectCurrent
End Sub

Private Sub ZoomPicture(pct As PictureBox, zoom As Double)
    With pct
        .AutoRedraw = True
        .Width = .Width * zoom
        .Height = .Height * zoom
        .PaintPicture .Picture, 0, 0, .ScaleWidth, .ScaleHeight
        .Refresh
    End With
End Sub
Private Sub tvLevels_NodeClick(ByVal Node As ComctlLib.Node)
    
    If (Node.Children) Then
    
        '-- Single click expands roots
        If (Node.Expanded = False) Then
            Node.Expanded = True
        End If
        
        '-- Root node: no level selected
        Call ucThumbnail.DIB.Destroy
        Call ucThumbnail.Refresh
        lblLemsToLetOutVal = ""
        lblLemsToBeSavedVal = ""
        lblPlayingTimeVal = ""
        lblSkill(0) = ""
        lblSkill(1) = ""
        lblSkill(2) = ""
        lblSkill(3) = ""
        lblSkill(4) = ""
        lblSkill(5) = ""
        lblSkill(6) = ""
        lblSkill(7) = ""
        
      Else
        '-- Extract level key
        g_nLevelID = Val(Mid$(tvLevels.SelectedItem.Key, 2))
        Debug.Print g_nLevelID
        '-- Load/create level thumbnail
        Call GetLevelThumbnail(ucThumbnail.DIB)
        Call ucThumbnail.Refresh
        
        '-- Get level info
        Call LoadLevelInfo(g_nLevelID)
        With g_uLevel
            lblLemsToLetOutVal = .LemsToLetOut
            lblLemsToBeSavedVal = .LemsToBeSaved
            lblPlayingTimeVal = .PlayingTime & "'"
            lblSkill(0) = IIf(.MaxClimbers, .MaxClimbers, "-")
            lblSkill(1) = IIf(.MaxFloaters, .MaxFloaters, "-")
            lblSkill(2) = IIf(.MaxBombers, .MaxBombers, "-")
            lblSkill(3) = IIf(.MaxBlockers, .MaxBlockers, "-")
            lblSkill(4) = IIf(.MaxBuilders, .MaxBuilders, "-")
            lblSkill(5) = IIf(.MaxBashers, .MaxBashers, "-")
            lblSkill(6) = IIf(.MaxMiners, .MaxMiners, "-")
            lblSkill(7) = IIf(.MaxDiggers, .MaxDiggers, "-")
        End With
    End If
End Sub

Private Sub cmdOK_Click()

    '-- Is a valid node?
    If (tvLevels.SelectedItem.Children = 0) Then
        '-- Yes
        Call VB.Unload(Me)
      Else
        '-- No
        Call VBA.MsgBox( _
             "No level has been selected." & vbCrLf & vbCrLf & "Please, select a valid level.", _
             vbExclamation _
             )
    End If
End Sub

Private Sub cmdCancel_Click()
    
    '-- Just exit
    Call VB.Unload(Me)
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub pvShowAllLevels()
    
  Dim nRatings   As Integer
  Dim sPath      As String
  Dim r          As Integer
  Dim l          As Integer
  Dim lID        As Integer
  Dim s          As String
  Dim bDone      As Boolean
  
    '-- Available ratings
    Select Case g_eGamePack
        Case [ePackLems]
            Call tvLevels.Nodes.Add(, , "Fun", "Fun", 1)
            Call tvLevels.Nodes.Add(, , "Tricky", "Tricky", 1)
            Call tvLevels.Nodes.Add(, , "Taxing", "Taxing", 1)
            Call tvLevels.Nodes.Add(, , "Mayhem", "Mayhem", 1)
            nRatings = 4
        Case [ePackOhNoMoreLems]
            Call tvLevels.Nodes.Add(, , "Tame", "Tame", 1)
            Call tvLevels.Nodes.Add(, , "Crazy", "Crazy", 1)
            Call tvLevels.Nodes.Add(, , "Wild", "Wild", 1)
            Call tvLevels.Nodes.Add(, , "Wicked", "Wicked", 1)
            Call tvLevels.Nodes.Add(, , "Havoc", "Havoc", 1)
            nRatings = 5
        Case [ePackCustom]
            Call tvLevels.Nodes.Add(, , "Custom", "Custom", 2)
    End Select
    
    '-- Load levels...
    
    sPath = AppPath & "LEVELS\"
    
    If (g_eGamePack = [ePackCustom]) Then
        
        '-- Search levels...
        For lID = g_eGamePack * 1000 To g_eGamePack * 1000 + 999
            
            '-- Level ID
            s = Format$(lID, "0000")
            
            '-- Exists?
            If (FileExists(sPath & s & ".dat")) Then
            
                '-- Done?
                bDone = IsLevelDone(Val(s))
                
                '-- Add to list
                Call LoadLevelTitle(Val(s))
                Call tvLevels.Nodes.Add(r + 1, tvwChild, _
                     "k" & s, Trim$(g_uLevel.Title), _
                     IIf(bDone, 4, 3) _
                     )
            End If
        Next lID
    
      Else
    
        For r = 0 To nRatings
            
            '-- Starting level
            lID = g_eGamePack * 1000 + r * 100
            l = 0
            
            '-- Get all levels
            Do While FileExists(sPath & Format$(lID, "0000") & ".dat")
                
                '-- Level ID
                s = Format$(lID, "0000")
                l = l + 1
                
                '-- Done?
                bDone = IsLevelDone(Val(s))
                
                '-- Add to list
                Call LoadLevelTitle(Val(s))
                Call tvLevels.Nodes.Add(r + 1, tvwChild, "k" & s, _
                     l & ". " & Trim$(g_uLevel.Title), _
                     IIf(bDone, 4, 3) _
                     )
                
                '-- Last done?
                If (bDone) Then
                    lID = lID + 1
                  Else
                    Exit Do
                End If
            Loop
        Next r
    End If
End Sub

Private Sub pvSelectCurrent()
    
  Dim sKey As String
    
    On Error GoTo errH
    sKey = "k" & Format$(g_nLevelID, "0000")
    With tvLevels
        .Nodes(sKey).Selected = True
        Call .Nodes(sKey).EnsureVisible
        Call tvLevels_NodeClick(.Nodes(sKey))
    End With

errH:
    On Error GoTo 0
End Sub
