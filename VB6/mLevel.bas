Attribute VB_Name = "mLevel"
Option Explicit

Public Type RECTI
    x1                  As Integer
    y1                  As Integer
    x2                  As Integer
    y2                  As Integer
End Type

Public Type tOBJGFX
    DIB                 As New cDIB08
    Width               As Byte
    Height              As Byte
    StartAnimationFrame As Byte
    EndAnimationFrame   As Byte
    lpTriggerRect       As RECTI
    TriggerEffect       As Byte
    SoundEffect         As Byte
    SoundEffectAtFrame  As Byte
End Type

Public Type tTERGFX
    DIB                 As New cDIB08
    Width               As Integer
    Height              As Byte
End Type

Public Type tObject
    ID                  As Byte
    lpRect              As RECTI
    NotOverlap          As Boolean
    OnTerrain           As Boolean
    UpsideDown          As Boolean
    pvFrameIdxCur       As Byte
    pvFrameIdxMax       As Byte
    pvLoop              As Boolean
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

Public g_eGamePack      As eGamePackConstants ' 0, 1,..., 9 [custom]
Public g_nLevelID       As Integer            ' #### [pack#|rating#|level##]
Public g_uLevel         As tLevelData         ' level data
Public g_uObjGFX()      As tOBJGFX            ' object item
Public g_uTerGFX()      As tTERGFX            ' terrain item
Public g_oDIBBack       As New cDIB08         ' extended level background image
Public g_nGraphicSet    As Integer            ' current loaded graphic-set
Public g_nGraphicSetEx  As Integer            ' current loaded extended graphic-set (image background)



'========================================================================================
' Initialization / termination
'========================================================================================

Public Sub InitializeLevel()
    
    '-- Reset current loaded graphic-set IDs
    g_nGraphicSet = -1
    g_nGraphicSetEx = -1
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Function LoadLevelTitle( _
                ByVal ID As Integer _
                ) As Boolean

  Dim sPath As String
  Dim hFile As Long
  
    sPath = AppPath & "LEVELS\" & Format$(ID, "0000") & ".dat"
    If (FileExists(sPath)) Then
        With g_uLevel
            hFile = VBA.FreeFile()
            Open sPath For Binary Access Read As #hFile
              Get #hFile, , .Title
            Close #hFile
        End With
        LoadLevelTitle = True
    End If
End Function

Public Function LoadLevelInfo( _
                ByVal ID As Integer _
                ) As Boolean

  Dim sPath As String
  Dim hFile As Long
  
    sPath = AppPath & "LEVELS\" & Format$(ID, "0000") & ".dat"
    If (FileExists(sPath)) Then
        With g_uLevel
            hFile = VBA.FreeFile()
            Open sPath For Binary Access Read As #hFile
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
            Close #hFile
        End With
        LoadLevelInfo = True
    End If
End Function

Public Function LoadLevel( _
                ByVal ID As Integer _
                ) As Boolean

  Dim sPath As String
  Dim hFile As Long
  Dim i     As Long
  
    sPath = AppPath & "LEVELS\" & Format$(ID, "0000") & ".dat"
    
    If (FileExists(sPath)) Then
    
        With g_uLevel
            
            hFile = VBA.FreeFile()
            Open sPath For Binary Access Read As #hFile
            
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
                
                If (g_nGraphicSet <> .GraphicSet Or g_nGraphicSetEx <> .GraphicSetEx) Then
                    g_nGraphicSet = .GraphicSet
                    g_nGraphicSetEx = .GraphicSetEx
                    Call pvLoadGraphicSet(g_nGraphicSet, g_nGraphicSetEx)
                End If
            
                Get #hFile, , .Objects
                ReDim .Object(0 To .Objects)
                
                For i = 1 To .Objects
                    With .Object(i)
                        Get #hFile, , .ID
                        Get #hFile, , .lpRect.x1: .lpRect.x2 = .lpRect.x1 + g_uObjGFX(.ID).Width
                        Get #hFile, , .lpRect.y1: .lpRect.y2 = .lpRect.y1 + g_uObjGFX(.ID).Height
                        Get #hFile, , .NotOverlap
                        Get #hFile, , .OnTerrain
                        Get #hFile, , .UpsideDown
                        .pvFrameIdxCur = g_uObjGFX(.ID).StartAnimationFrame
                        .pvFrameIdxMax = g_uObjGFX(.ID).EndAnimationFrame
                    End With
                Next i
        
                Get #hFile, , .TerrainPieces
                ReDim .TerrainPiece(0 To .TerrainPieces)
    
                For i = 1 To .TerrainPieces
                    With .TerrainPiece(i)
                        Get #hFile, , .ID
                        Get #hFile, , .lpRect.x1: .lpRect.x2 = .lpRect.x1 + g_uTerGFX(.ID).Width
                        Get #hFile, , .lpRect.y1: .lpRect.y2 = .lpRect.y1 + g_uTerGFX(.ID).Height
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
        
        LoadLevel = True
    End If
End Function

Public Function GetNextLevel( _
                ) As Integer
    
  Dim sTmp As String
    
    If (g_eGamePack = [ePackCustom]) Then
        
        '-- Return same level:
        '   Unexpected sequence
        GetNextLevel = g_nLevelID
    
      Else
      
        '-- Get next level (same rating) file path
        sTmp = AppPath & _
               "LEVELS\" & _
               Format$(g_nLevelID + 1, "0000") & ".dat"
        
        '-- Check file
        If (FileExists(sTmp)) Then
            
            '-- Exists: OK
            GetNextLevel = g_nLevelID + 1
          Else
            
            '-- Get first level (next rating)
            sTmp = AppPath & _
                   "LEVELS\" & _
                   Format$((g_nLevelID \ 100 + 1) * 100, "0000") & ".dat"
            
            '-- Check file
            If (FileExists(sTmp)) Then
                
                '-- Exists: OK
                GetNextLevel = (g_nLevelID \ 100 + 1) * 100
              
              Else
                '-- All done: start again
                GetNextLevel = g_eGamePack * 1000
            End If
        End If
    End If
End Function

Public Function GetLevelRatingString( _
                ByVal ID As Integer _
                ) As String
    
  Dim r As Integer
    
    '-- Get rating
    r = Val(Mid$(Format$(ID, "0000"), 2, 1)) + 1
    
    '-- Available ratings
    Select Case g_eGamePack
        Case [ePackLems]
            GetLevelRatingString = Choose(r, "Fun", "Tricky", "Taxing", "Mayhem")
        Case [ePackOhNoMoreLems]
            GetLevelRatingString = Choose(r, "Tame", "Crazy", "Wild", "Wicked", "Havoc")
        Case [ePackCustom]
            GetLevelRatingString = "N/A"
    End Select
End Function

'========================================================================================
' Private
'========================================================================================

Private Sub pvLoadGraphicSet( _
            ByVal GraphicSet As Byte, _
            ByVal GraphicSetEx As Byte _
            )
  
  Dim i    As Long
  Dim sINI As String
  Dim sKey As String
    
    Screen.MousePointer = vbHourglass
    
    With g_uLevel
        
        '-- INI file
        sINI = AppPath & "CONFIG\GS_" & GraphicSet & ".ini"
        
        '-- Objects collection
        ReDim g_uObjGFX(0 To Val( _
            GetINI(sINI, "main", "ObjectCount")) - 1)
        
        '-- Load available objects
        For i = 0 To UBound(g_uObjGFX())
            
            '-- Create 8bit image
            Call g_uObjGFX(i).DIB.CreateFromBitmapFile( _
                 AppPath & "GFX\" & _
                 "obj_" & GraphicSet & "_" & Format$(i, "00") & ".bmp" _
                 )
            
            '-- Get animation info
            sKey = "obj_" & Format$(i, "00")
            With g_uObjGFX(i)
                
                '-- Animation frame size
                .Width = _
                    Val(GetINI(sINI, sKey, "Width"))
                .Height = _
                    Val(GetINI(sINI, sKey, "Height"))
                
                '-- Start and ending frames
                .StartAnimationFrame = _
                    Val(GetINI(sINI, sKey, "StartAnimationFrame"))
                .EndAnimationFrame = _
                    Val(GetINI(sINI, sKey, "EndAnimationFrame"))
                
                '-- Trigger area and related effect
                .TriggerEffect = _
                    Val(GetINI(sINI, sKey, "TriggerEffect"))
                With .lpTriggerRect
                    .x1 = _
                        Val(GetINI(sINI, sKey, "TriggerLeft"))
                    .x2 = .x1 + _
                        Val(GetINI(sINI, sKey, "TriggerWidth"))
                    .y1 = _
                        Val(GetINI(sINI, sKey, "TriggerTop"))
                    .y2 = .y1 + _
                        Val(GetINI(sINI, sKey, "TriggerHeight"))
                End With
                
                '-- Trap sound effect
                .SoundEffect = _
                    Val(GetINI(sINI, sKey, "SoundEffect"))
                    
                '-- Frame to play sound effect
                .SoundEffectAtFrame = _
                    Val(GetINI(sINI, sKey, "SoundEffectAtFrame"))
            End With
        Next i
        
        '-- Terrain
        If (GraphicSetEx > 0) Then
            
            '-- Load background image
             Call g_oDIBBack.CreateFromBitmapFile( _
                  AppPath & "GFX\" & _
                  "back_" & Format$(GraphicSetEx, "0") & "ex.bmp" _
                  )
          Else
            
            '-- Terrain pieces collection
            ReDim g_uTerGFX(0 To Val( _
                GetINI(sINI, "main", "TerrainCount")) - 1)
            
            '-- Load available terrain pieces
            For i = 0 To UBound(g_uTerGFX())
                Call g_uTerGFX(i).DIB.CreateFromBitmapFile( _
                     AppPath & "GFX\" & _
                     "ter_" & Format$(GraphicSet, "0") & "_" & Format$(i, "00") & ".bmp" _
                     )
                     
                '-- Set item info
                With g_uTerGFX(i)
                    .Width = .DIB.Width
                    .Height = .DIB.Height
                End With
            Next i
        End If
    End With

    '-- Finaly, load/merge level palette
    If (GraphicSetEx > 0) Then
        Call MergePaletteEntries( _
            GetINI(sINI, "main", "BrickColorEx" & GraphicSetEx), 7)
        Call MergePaletteEntries( _
            GetINI(sINI, "main", "PaletteEx" & GraphicSetEx), 8)
      Else
        Call MergePaletteEntries( _
            GetINI(sINI, "main", "BrickColorDef"), 7)
        Call MergePaletteEntries( _
            GetINI(sINI, "main", "PaletteDef"), 8)
    End If
    
    '-- Apply to game views
    Call fLems.ucScreen.UpdatePalette(GetGlobalPalette())
    Call fLems.ucPanoramicView.UpdatePalette(GetGlobalPalette())
    
    Screen.MousePointer = vbDefault
End Sub

