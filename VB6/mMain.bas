Attribute VB_Name = "mMain"
Option Explicit

Public Const MAX_GRAPHICSET    As Byte = 8

Public Const MAX_OBJECTS       As Integer = 32
Public Const MAX_TERRAINPIECES As Integer = 400
Public Const MAX_STEELAREAS    As Integer = 32
Public Const MIN_STEELAREASIZE As Integer = 8

Public Const MIN_LEMSTOLETOUT  As Byte = 1
Public Const MAX_LEMSTOLETOUT  As Byte = 80
Public Const MIN_LEMSTOBESAVED As Byte = MIN_LEMSTOLETOUT
Public Const MIN_RELEASERATE   As Byte = 1
Public Const MAX_RELEASERATE   As Byte = 99
Public Const MAX_PLAYINGTIME   As Byte = 10
Public Const MAX_SKILL         As Byte = 80

Public Type RECTI
    x1                  As Integer
    y1                  As Integer
    x2                  As Integer
    y2                  As Integer
End Type

Public Type tOBJGFX
    DIB                 As New cDIB32
    Width               As Byte
    Height              As Byte
    StartAnimationFrame As Byte
    EndAnimationFrame   As Byte
    lpTriggerRect       As RECTI
    TriggerEffect       As Byte
    SoundEffect         As Byte
End Type

Public Type tTERGFX
    DIB                 As New cDIB32
    Width               As Integer
    Height              As Byte
End Type

Public Type tObject
    ID                  As Byte
    lpRect              As RECTI
    NotOverlap          As Boolean
    OnTerrain           As Boolean
    UpsideDown          As Boolean
End Type

Public Type tTerrainPiece
    ID                  As Byte
    lpRect              As RECTI
    NotOverlap          As Boolean
    Black               As Boolean
    UpsideDown          As Boolean
End Type

Public Type tSteelArea
    lpRect              As RECTI
End Type

Public Type tLevelData
    Title               As String * 32   '  1
    LemsToLetOut        As Byte          ' 33
    LemsToBeSaved       As Byte          ' 34
    ReleaseRate         As Byte          ' 35
    PlayingTime         As Byte          ' 36
    MaxClimbers         As Byte          ' 37
    MaxFloaters         As Byte          ' 38
    MaxBombers          As Byte          ' 39
    MaxBlockers         As Byte          ' 40
    MaxBuilders         As Byte          ' 41
    MaxBashers          As Byte          ' 42
    MaxMiners           As Byte          ' 43
    MaxDiggers          As Byte          ' 44
    ScreenStart         As Integer       ' 45
    GraphicSet          As Byte          ' 47
    GraphicSetEx        As Byte          ' 48
    Objects             As Integer       ' 49
    Object()            As tObject
    TerrainPieces       As Integer
    TerrainPiece()      As tTerrainPiece
    SteelAreas          As Integer
    SteelArea()         As tSteelArea
End Type

Public sLEVELPATH       As String
Public uLEVEL           As tLevelData
Public uOBJGFX()        As tOBJGFX
Public uTERGFX()        As tTERGFX



'========================================================================================
' Methods
'========================================================================================

Public Function IsExtendedLevel( _
                ByVal Filename As String _
                ) As Boolean
                
                
  Dim hFile As Long
  Dim a48   As Byte

    hFile = VBA.FreeFile()
    Open Filename For Binary Access Read As #hFile
      Get #hFile, 48, a48
    Close #hFile
    
    IsExtendedLevel = (a48 > 0)
End Function

Public Sub LoadLevel( _
           ByVal Filename As String _
           )

  Dim hFile As Long
  Dim i     As Long

    hFile = VBA.FreeFile()

    With uLEVEL

        Open Filename For Binary Access Read As #hFile

            Get #hFile, , .Title

            Get #hFile, , .LemsToLetOut
            Get #hFile, , .LemsToBeSaved
            Get #hFile, , .ReleaseRate
            Get #hFile, , .PlayingTime

            Get #hFile, , .MaxClimbers
            Get #hFile, , .MaxFloaters
            Get #hFile, , .MaxBombers
            Get #hFile, , .MaxBlockers
            Get #hFile, , .MaxBuilders
            Get #hFile, , .MaxBashers
            Get #hFile, , .MaxMiners
            Get #hFile, , .MaxDiggers

            Get #hFile, , .ScreenStart

            Get #hFile, , .GraphicSet
            Get #hFile, , .GraphicSetEx
            
            Call pvLoadGraphicSet(.GraphicSet, .GraphicSetEx)

            Get #hFile, , .Objects
            ReDim .Object(0 To .Objects)
            
            For i = 1 To .Objects
                With .Object(i)
                    Get #hFile, , .ID
                    Get #hFile, , .lpRect.x1: .lpRect.x2 = .lpRect.x1 + uOBJGFX(.ID).Width
                    Get #hFile, , .lpRect.y1: .lpRect.y2 = .lpRect.y1 + uOBJGFX(.ID).Height
                    Get #hFile, , .NotOverlap
                    Get #hFile, , .OnTerrain
                    Get #hFile, , .UpsideDown
                End With
            Next i

            Get #hFile, , .TerrainPieces
            ReDim .TerrainPiece(0 To .TerrainPieces)

            For i = 1 To .TerrainPieces
                With .TerrainPiece(i)
                    Get #hFile, , .ID
                    Get #hFile, , .lpRect.x1: .lpRect.x2 = .lpRect.x1 + uTERGFX(.ID).Width
                    Get #hFile, , .lpRect.y1: .lpRect.y2 = .lpRect.y1 + uTERGFX(.ID).Height
                    Get #hFile, , .NotOverlap
                    Get #hFile, , .Black
                    Get #hFile, , .UpsideDown
                End With
            Next i

            Get #hFile, , .SteelAreas
            ReDim .SteelArea(0 To .SteelAreas)

            For i = 1 To .SteelAreas
                With .SteelArea(i)
                    Get #hFile, , .lpRect.x1
                    Get #hFile, , .lpRect.y1
                    Get #hFile, , .lpRect.x2
                    Get #hFile, , .lpRect.y2
                End With
            Next i

        Close #hFile
    End With
End Sub

Public Sub SaveLevel( _
           ByVal Filename As String _
           )
    
  Dim hFile As Long
  Dim i     As Long
  
    On Error Resume Next
    Call VBA.Kill(Filename)
    On Error GoTo 0
  
    hFile = VBA.FreeFile()
  
    With uLEVEL
        
        Open Filename For Binary Access Write As #hFile
        
            Put #hFile, , .Title
            
            Put #hFile, , .LemsToLetOut
            Put #hFile, , .LemsToBeSaved
            Put #hFile, , .ReleaseRate
            Put #hFile, , .PlayingTime
            
            Put #hFile, , .MaxClimbers
            Put #hFile, , .MaxFloaters
            Put #hFile, , .MaxBombers
            Put #hFile, , .MaxBlockers
            Put #hFile, , .MaxBuilders
            Put #hFile, , .MaxBashers
            Put #hFile, , .MaxMiners
            Put #hFile, , .MaxDiggers
            
            Put #hFile, , .ScreenStart
            Put #hFile, , .GraphicSet
            Put #hFile, , .GraphicSetEx
            
            Put #hFile, , .Objects
            If (.Objects) Then
                For i = 1 To .Objects
                    With .Object(i)
                        Put #hFile, , .ID
                        Put #hFile, , .lpRect.x1
                        Put #hFile, , .lpRect.y1
                        Put #hFile, , .NotOverlap
                        Put #hFile, , .OnTerrain
                        Put #hFile, , .UpsideDown
                    End With
                Next i
            End If
    
            Put #hFile, , .TerrainPieces
            If (.TerrainPieces) Then
                For i = 1 To .TerrainPieces
                    With .TerrainPiece(i)
                        Put #hFile, , .ID
                        Put #hFile, , .lpRect.x1
                        Put #hFile, , .lpRect.y1
                        Put #hFile, , .NotOverlap
                        Put #hFile, , .Black
                        Put #hFile, , .UpsideDown
                    End With
                Next i
            End If
                
            Put #hFile, , .SteelAreas
            If (.SteelAreas) Then
                For i = 1 To .SteelAreas
                    With .SteelArea(i)
                        Put #hFile, , .lpRect.x1
                        Put #hFile, , .lpRect.y1
                        Put #hFile, , .lpRect.x2
                        Put #hFile, , .lpRect.y2
                    End With
                Next i
            End If

        Close #hFile
    End With
End Sub

Public Sub ClearLevel( _
           Optional ByVal ClearLevelFeatures As Boolean = True _
           )
    
    With uLEVEL
        
        If (ClearLevelFeatures) Then
        
            .Title = vbNullString
            
            .LemsToLetOut = 0
            .LemsToBeSaved = 0
            .ReleaseRate = 0
            .PlayingTime = 0
            
            .MaxClimbers = 0
            .MaxFloaters = 0
            .MaxBombers = 0
            .MaxBlockers = 0
            .MaxBuilders = 0
            .MaxBashers = 0
            .MaxMiners = 0
            .MaxDiggers = 0
        End If
        
        .Objects = 0
        ReDim .Object(0 To .Objects)

        .TerrainPieces = 0
        ReDim .TerrainPiece(0 To .TerrainPieces)
            
        .SteelAreas = 0
        ReDim .SteelArea(0 To .SteelAreas)
        
        '-- Center view
        fEdit.ucScroll.Value = 640
    End With
End Sub

Public Function CheckLevel( _
                ) As String
    
  Dim sTmp As String
  Dim i    As Long
  Dim c    As Long
  
  Dim bStartExists As Boolean
  Dim bExitExists  As Boolean
  
    sTmp = "Checking results:" & vbCrLf & vbCrLf
  
    With uLEVEL
        
        '-- Main data
        If (.LemsToLetOut < MIN_LEMSTOLETOUT Or .LemsToLetOut > MAX_LEMSTOLETOUT) Then
            sTmp = sTmp & "Invalid 'Lems to let out' value." & vbCrLf
        End If
        If (.LemsToBeSaved < MIN_LEMSTOLETOUT Or .LemsToBeSaved > .LemsToLetOut) Then
            sTmp = sTmp & "Invalid 'Lems to be saved' value." & vbCrLf
        End If
        If (.ReleaseRate < MIN_RELEASERATE Or .ReleaseRate > MAX_RELEASERATE) Then
            sTmp = sTmp & "Invalid 'Release rate' value." & vbCrLf
        End If
        If (.PlayingTime <= 0) Then
            sTmp = sTmp & "Invalid 'Playing time' value." & vbCrLf
        End If
        
        '-- Skills
        If (.MaxClimbers < 0 Or .MaxClimbers > MAX_SKILL) Then
            sTmp = sTmp & "Invalid 'Climbers' value." & vbCrLf
        End If
        If (.MaxFloaters < 0 Or .MaxFloaters > MAX_SKILL) Then
            sTmp = sTmp & "Invalid 'Floaters' value." & vbCrLf
        End If
        If (.MaxBombers < 0 Or .MaxBombers > MAX_SKILL) Then
            sTmp = sTmp & "Invalid 'Bombers' value." & vbCrLf
        End If
        If (.MaxBlockers < 0 Or .MaxBlockers > MAX_SKILL) Then
            sTmp = sTmp & "Invalid 'Blockers' value." & vbCrLf
        End If
        If (.MaxBuilders < 0 Or .MaxBuilders > MAX_SKILL) Then
            sTmp = sTmp & "Invalid 'Builders' value." & vbCrLf
        End If
        If (.MaxBashers < 0 Or .MaxBashers > MAX_SKILL) Then
            sTmp = sTmp & "Invalid 'Bashers' value." & vbCrLf
        End If
        If (.MaxMiners < 0 Or .MaxMiners > MAX_SKILL) Then
            sTmp = sTmp & "Invalid 'Miners' value." & vbCrLf
        End If
        If (.MaxDiggers < 0 Or .MaxDiggers > MAX_SKILL) Then
            sTmp = sTmp & "Invalid 'Diggers' value." & vbCrLf
        End If
        
        '-- Partial results
        If (sTmp = "Checking results:" & vbCrLf & vbCrLf) Then
            sTmp = sTmp & "No invalid 'Level features' values found." & vbCrLf & vbCrLf
          Else
            sTmp = sTmp & vbCrLf
        End If
        
        '-- Start/Exit
        For i = 1 To .Objects
            If (.Object(i).ID = 0) Then
                bExitExists = True
            End If
            If (.Object(i).ID = 1) Then
                bStartExists = True
            End If
        Next i
        If (bExitExists = False) Then
            sTmp = sTmp & "No 'Exit/s' found!." & vbCrLf
        End If
        If (bStartExists = False) Then
            sTmp = sTmp & "No 'Start/s' found!." & vbCrLf
        End If
        
        '-- Out of bound objects
        c = 0
        For i = 1 To .Objects
            If (.Object(i).ID > UBound(uOBJGFX())) Then
                c = c + 1
            End If
        Next i
        If (c = 0) Then
            sTmp = sTmp & "No invalid objects found." & vbCrLf
          Else
            sTmp = sTmp & c & " invalid objects found." & vbCrLf
        End If

        '-- Out of bound terrain pieces
        c = 0
        For i = 1 To .TerrainPieces
            If (.TerrainPiece(i).ID > UBound(uTERGFX())) Then
                c = c + 1
            End If
        Next i
        If (c = 0) Then
            sTmp = sTmp & "No invalid terrain pieces found." & vbCrLf
          Else
            sTmp = sTmp & c & " invalid terrain pieces found." & vbCrLf
        End If
    End With
    
    CheckLevel = sTmp
End Function

Public Sub SetGraphicSet( _
           ByVal Idx As Byte _
           )
  
  Dim i As Long
     
     With uLEVEL
    
        If (Idx <> .GraphicSet) Then
        
            '-- Check first if any object or terrain piece
            If (.TerrainPieces > 0 Or _
                .Objects > 0) Then
                
                Call VBA.MsgBox( _
                     "You need to clear level in order to change set.", _
                     vbInformation _
                     )
                
                '-- Restore
                fEdit.cbGraphicSet.Tag = 1
                fEdit.cbGraphicSet.ListIndex = .GraphicSet
                fEdit.cbGraphicSet.Tag = 0
                
              Else
              
                '-- Load available objects and terrain pieces
                .GraphicSet = Idx
                Call pvLoadGraphicSet(Idx, 0)
                
                '-- Update objects and terrain pieces rectangles
                For i = 1 To .Objects
                   With .Object(i)
                       .lpRect.x2 = .lpRect.x1 + uOBJGFX(.ID).Width
                       .lpRect.y2 = .lpRect.y1 + uOBJGFX(.ID).Height
                   End With
                Next i
                For i = 1 To .TerrainPieces
                   With .TerrainPiece(i)
                       .lpRect.x2 = .lpRect.x1 + uTERGFX(.ID).Width
                       .lpRect.y2 = .lpRect.y1 + uTERGFX(.ID).Height
                   End With
                Next i
            End If
        End If
    End With
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub pvLoadGraphicSet( _
            ByVal GraphicSet As Byte, _
            ByVal GraphicSetEx As Byte _
            )
  
  Dim i    As Long
  Dim sINI As String
  Dim sKEY As String
    
    Screen.MousePointer = vbHourglass
    
    With fEdit
    
        '-- Graphic set INI file
        sINI = mMisc.AppPath & "CONFIG\GS_" & GraphicSet & ".ini"
        
        '-- Clear lists
        Call .cbObject.Clear
        Call .cbTerrainPiece.Clear
        
        '-- Objects collection
        ReDim uOBJGFX(0 To Val( _
            GetINI(sINI, "main", "ObjectCount")) - 1)
        
        '-- Load/add available objects
        For i = 0 To UBound(uOBJGFX())
            
            '-- Create 32bit image
            Call uOBJGFX(i).DIB.CreateFromStdPicture( _
                 VB.LoadPicture(mMisc.AppPath & "GFX\" & _
                                "obj_" & GraphicSet & "_" & Format$(i, "00") & ".bmp" _
                                ) _
                 )
            
            '-- Get animation info
            sKEY = "obj_" & Format$(i, "00")
            With uOBJGFX(i)
                
                '-- Animation frame size
                .Width = Val( _
                    GetINI(sINI, sKEY, "Width"))
                .Height = Val( _
                    GetINI(sINI, sKEY, "Height"))
                
                '-- Number of frames
                .EndAnimationFrame = Val( _
                    GetINI(sINI, sKEY, "EndAnimationFrame"))
                
                '-- Trigger area and related effect
                .TriggerEffect = Val( _
                    GetINI(sINI, sKEY, "TriggerEffect"))
                With .lpTriggerRect
                    .x1 = Val( _
                        GetINI(sINI, sKEY, "TriggerLeft"))
                    .x2 = .x1 + Val( _
                        GetINI(sINI, sKEY, "TriggerWidth"))
                    .y1 = Val( _
                        GetINI(sINI, sKEY, "TriggerTop"))
                    .y2 = .y1 + Val( _
                        GetINI(sINI, sKEY, "TriggerHeight"))
                End With
            End With
            
            '-- Update combo-list
            Call .cbObject.AddItem("# " & i)
        Next i
        
        '-- Terrain pieces collection
        ReDim uTERGFX(0 To Val( _
            GetINI(sINI, "main", "TerrainCount")) - 1)
        
        '-- Load/add available terrain pieces
        For i = 0 To UBound(uTERGFX())
            
            '-- Create 32bit image
            Call uTERGFX(i).DIB.CreateFromStdPicture( _
                 VB.LoadPicture(mMisc.AppPath & "GFX\" & _
                                "ter_" & GraphicSet & "_" & Format$(i, "00") & ".bmp" _
                                ) _
                 )
                
                '-- Set item info
                With uTERGFX(i)
                    .Width = .DIB.Width
                    .Height = .DIB.Height
                End With
                 
            '-- Update combo-list
            Call .cbTerrainPiece.AddItem("# " & i)
        Next i
        
        '-- Update info labels
        .lblObjectsVal.Caption = UBound(uOBJGFX()) + 1
        .lblTerrainPiecesVal.Caption = UBound(uTERGFX()) + 1
        
        '-- Update graphic set
        .cbGraphicSet.Tag = 1
        .cbGraphicSet.ListIndex = uLEVEL.GraphicSet
        .cbGraphicSet.Tag = 0
        
        '-- Select first item
        .cbObject.ListIndex = 0
        .cbTerrainPiece.ListIndex = 0
        
        '-- Enable animation timer
        .tmrAnimation.Enabled = True
    End With
    
    Screen.MousePointer = vbDefault
End Sub
