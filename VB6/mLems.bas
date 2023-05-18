Attribute VB_Name = "mLems"
Option Explicit

'-- A little bit of API:

Private Type RECT
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound   As Long
End Type

Private Type SAFEARRAY2D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds(1)  As SAFEARRAYBOUND
End Type

Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)

'-- Lems:

Private Type tParticle
    x            As Integer
    y            As Integer
    vx           As Integer
    vy           As Integer
End Type

Private Type tAnimationData
    FrameOffY    As Integer
    FrameIdxMax  As Byte
    FrameHasDir  As Boolean
End Type

Private Const MAX_PARTICLES As Long = 24
Private Type tLem
    Active       As Boolean
    DieNextFrame As Boolean
    Job          As eLemJobConstants
    Ability      As eLemAbilityConstants
    ExplodeCount As Integer
    Frame        As eLemFrameConstants
    FrameSrcY    As Integer
    FrameIdx     As Integer
    FrameIdxMax  As Integer
    FrameOffY    As Integer
    FrameHasDir  As Boolean
    x            As Integer
    y            As Integer
    xs           As Integer
    f            As Integer
    Particles    As Boolean
    Particle(MAX_PARTICLES) As tParticle
End Type

Public Enum eLemFrameConstants
    [eFrameNone] = -1
    [eFrameWalker] = 0
    [eFrameFalling] = 2
    [eFrameSpliting] = 4
    [eFrameDrowning] = 5
    [eFrameBurning] = 6
    [eFrameExploding] = 7
    [eFrameSurviving] = 8
    [eFrameClimber] = 9
    [eFrameClimberEnd] = 11
    [eFrameFloater] = 13
    [eFrameBlocker] = 15
    [eFrameBuilder] = 16
    [eFrameBuilderEnd] = 18
    [eFrameBasher] = 20
    [eFrameMiner] = 22
    [eFrameDigger] = 24
End Enum

Public Enum eLemJobConstants
    [eJobNone] = 0
    [eJobBlocker] = 1
    [eJobBuilder] = 2
    [eJobBasher] = 3
    [eJobMiner] = 4
    [eJobDigger] = 5
End Enum

Public Enum eLemAbilityConstants
    [eAbilityNone] = 0
    [eAbilityClimber] = 1
    [eAbilityFloater] = 2
    [eAbilityBomber] = 4
End Enum

Public Enum eGameStageConstants
    [eStageLetsGo] = 0
    [eStageOpeningDoors]
    [eStagePlaying]
    [eStageEnding]
End Enum

Public Enum eLemsSavedModeConstants
    [eSavedModePercentage] = 0
    [eSavedModePercentageRemaining]
    [eSavedModeCount]
    [eSavedModeCountRemaining]
End Enum

Private Const MIN_FALL_FLOATER          As Long = 20
Private Const MAX_FALL                  As Long = 63
Private Const MIN_OBSTACLE              As Long = 7
Private Const MAX_BRICKS                As Long = 12
Private Const MAX_YCHECK                As Long = 175
Private Const EXPLODE_TICKS             As Long = 75
Private Const DOOR_ID                   As Byte = 1
Private Const DOOR_XOFFSET              As Long = 17
Private Const DOOR_YOFFSET              As Long = -2
Private Const DOOR_FALLCOUNTER          As Long = 4

Private Const LEMS_INFO                 As String = "4361726C657320502E562E207B20323030352D32303131"
Private Const ANIMATION_DATA            As String = "00071000000003100000-11500015000130001500007000071000000007100000000910000000150001510000000071000000031100000022310000002151"
Private Const BASHER_DATA               As String = "12222233345432111222223344543211"
Private Const CHEAT_CODE                As String = "OHNO"

Private m_uAnimationData(24)            As tAnimationData
Private m_aBasherData(31)               As Byte

Private m_eGameStage                    As eGameStageConstants
Private m_lcGameStage                   As Long

Private m_oScreen                       As ucScreen08
Private m_oDIBLems                      As New cDIB08
Private m_oDIBMask                      As New cDIB08
Private m_oDIBScreenBkMask              As New cDIB08
Private m_oDIBScreenBuffer1             As New cDIB08
Private m_oDIBScreenBuffer2             As New cDIB08

Private m_oPanoramicView                As ucScreen08
Private m_uPanoramicViewScanX1          As Integer
Private m_uPanoramicViewScanX2          As Integer
Private m_oDIBLemPoint                  As New cDIB08

Private m_uSAScreenBkMask               As SAFEARRAY2D
Private m_aScreenBkMaskBits()           As Byte
Private m_uScreenBkMaskRect             As RECT

Private m_iCursorPointer                As New StdPicture
Private m_iCursorSelect                 As New StdPicture

Private m_xScreen                       As Long
Private m_xCur                          As Long
Private m_yCur                          As Long

Private m_ePreparedAbility              As eLemAbilityConstants
Private m_ePreparedJob                  As eLemJobConstants
Private m_lLem                          As Long
    
Private m_uPtOut()                      As POINTAPI
Private m_lPtOut                        As Long
Private m_lcPtOut                       As Long

Private m_uLems()                       As tLem
Private m_lLemsOut                      As Long
Private m_lLemsSaved                    As Long
Private m_eLemsSavedMode                As eLemsSavedModeConstants

Private m_bLemsInfo                     As Boolean
Private m_sLemsInfo                     As String
Private m_xInfo                         As Long
Private m_tHold                         As Long

Private m_sCheat                        As String


'========================================================================================
' Initialization / termination
'========================================================================================

Public Sub InitializeLems()

  Dim i As Long
    
    '-- Initialize some data
    For i = 0 To 24
        With m_uAnimationData(i)
            .FrameOffY = _
                Mid$(ANIMATION_DATA, 5 * i + 1, 2)
            .FrameIdxMax = _
                Mid$(ANIMATION_DATA, 5 * i + 3, 2)
            .FrameHasDir = _
                Mid$(ANIMATION_DATA, 5 * i + 5, 1)
        End With
    Next i
    For i = 0 To 31
        m_aBasherData(i) = Mid$(BASHER_DATA, i + 1, 1)
    Next i
    For i = 1 To Len(LEMS_INFO) Step 2
        m_sLemsInfo = m_sLemsInfo & Chr$("&H" & Mid$(LEMS_INFO, i, 2))
    Next i
    
    '-- Private reference to main form 'screens'
    Set m_oScreen = fLems.ucScreen
    Set m_oPanoramicView = fLems.ucPanoramicView
    
    '-- Pre-load selection cursors
    Set m_iCursorPointer = VB.LoadResPicture( _
        "CUR_POINTER", vbResCursor _
        )
    Set m_iCursorSelect = VB.LoadResPicture( _
        "CUR_SELECT", vbResCursor _
        )
    
    '-- Load Lems' frames & masks
    Call m_oDIBLems.CreateFromBitmapFile( _
         AppPath & "GFX\main_1.bmp" _
         )
    Call m_oDIBMask.CreateFromBitmapFile( _
         AppPath & "GFX\main_2.bmp" _
         )
    
    '-- Create back-mask and main buffers (terrain & terrain+objects) DIBs
    Call m_oDIBScreenBkMask.Create( _
         1600, 160 _
         ) ' 250 KB
    Call m_oDIBScreenBuffer1.Create( _
         1600, 160 _
         ) ' 250 KB
    Call m_oDIBScreenBuffer2.Create( _
         1600, 160 _
         ) ' 250 KB
        
    '-- Define back-mask rect.
    Call SetRect( _
         m_uScreenBkMaskRect, _
         0, 0, 1600, 160 _
         )
    
    '-- Map back-mask DIB bytes
    Call pvMapDIB( _
         m_uSAScreenBkMask, _
         m_aScreenBkMaskBits(), _
         m_oDIBScreenBkMask _
         )
    
    '-- Create Lem 'point' (panoramic view)
    Call m_oDIBLemPoint.Create( _
         2, 2, _
         BkColorIdx:=IDX_GREEN216 _
         )
End Sub

Public Sub TerminateLems()

    '-- Unmap 'screen'
    Call pvUnmapDIB(m_aScreenBkMaskBits())
End Sub

'========================================================================================
' Terrain/Mask initialization
'========================================================================================

Public Sub InitializeGame()
    
  Dim i As Long
  
    '-- Reset game state
    Call SetGameStage([eStageLetsGo])
    
    '-- Initialize 'Release rate' and 'Playing time'
    g_lReleaseRateMin = g_uLevel.ReleaseRate
    g_lReleaseRate = g_uLevel.ReleaseRate
    g_lcPlayingTime = g_uLevel.PlayingTime * 60
    
    '-- Initialize Lems' collection and related variables
    ReDim m_uLems(0)
    m_lLem = 0
    m_lLemsOut = 0
    m_lLemsSaved = 0
    m_ePreparedAbility = [eAbilityNone]
    m_ePreparedJob = [eJobNone]
    
    '-- Initialize Exits' collection
    ReDim m_uPtOut(0)
    m_lcPtOut = 0
    m_lPtOut = 1
    
    '-- Reset buffers and back-mask
    Call m_oDIBScreenBuffer1.Reset
    Call m_oDIBScreenBuffer2.Reset
    Call m_oDIBScreenBkMask.Reset
    
    '-- Info + cheat
    m_bLemsInfo = False
    m_xInfo = 0
    m_tHold = 0
    m_sCheat = vbNullString
        
    '-- Prepare level...
    
    '-- Initialize terrain
    Call DrawTerrain(m_oDIBScreenBuffer1)
    
    With g_uLevel
        
        '-- Initialize back-mask
        
        '  Terrain...
        Call MaskBltIdx( _
             m_oDIBScreenBkMask, _
             0, 0, _
             1600, 160, _
             IDX_TERRAIN, _
             m_oDIBScreenBuffer1, _
             0, 0, _
             IDX_NONE, _
             False _
             )
        
        '   Steel...
        For i = 1 To .SteelAreas
            With .SteelArea(i)
                Call MaskRectIdxOverlap( _
                     m_oDIBScreenBkMask, _
                     .lpRect.x1, .lpRect.y1, _
                     .lpRect.x2 - .lpRect.x1, .lpRect.y2 - .lpRect.y1, _
                     IDX_TERRAIN, IDX_STEEL _
                     )
            End With
        Next i

        '   Objects' trigger areas...
        For i = 1 To .Objects
            With .Object(i)
                If (g_uObjGFX(.ID).TriggerEffect > 0) Then
                    If (.OnTerrain) Then
                        Call MaskRectIdxOverlap( _
                             m_oDIBScreenBkMask, _
                             g_uObjGFX(.ID).lpTriggerRect.x1 + .lpRect.x1, _
                             g_uObjGFX(.ID).lpTriggerRect.y1 + .lpRect.y1, _
                             g_uObjGFX(.ID).lpTriggerRect.x2 - g_uObjGFX(.ID).lpTriggerRect.x1, _
                             g_uObjGFX(.ID).lpTriggerRect.y2 - g_uObjGFX(.ID).lpTriggerRect.y1, _
                             IDX_TERRAIN, g_uObjGFX(.ID).TriggerEffect _
                             )
                      Else
                        Call MaskRectIdxBkMask( _
                             m_oDIBScreenBkMask, _
                             g_uObjGFX(.ID).lpTriggerRect.x1 + .lpRect.x1, _
                             g_uObjGFX(.ID).lpTriggerRect.y1 + .lpRect.y1, _
                             g_uObjGFX(.ID).lpTriggerRect.x2 - g_uObjGFX(.ID).lpTriggerRect.x1, _
                             g_uObjGFX(.ID).lpTriggerRect.y2 - g_uObjGFX(.ID).lpTriggerRect.y1, _
                             g_uObjGFX(.ID).TriggerEffect _
                             )
                    End If
                End If
            End With
        Next i
        
        '-- Set loop mode
        For i = 1 To .Objects
            With .Object(i)
                Select Case .ID
                    '-- Door
                    Case Is = DOOR_ID
                        m_lcPtOut = m_lcPtOut + 1
                        ReDim Preserve m_uPtOut(0 To m_lcPtOut)
                        m_uPtOut(m_lcPtOut).x = .lpRect.x1 + DOOR_XOFFSET
                        m_uPtOut(m_lcPtOut).y = .lpRect.y1 + DOOR_YOFFSET
                    '-- Any other
                    Case Else
                        Select Case g_uObjGFX(.ID).TriggerEffect
                            Case Is <> IDX_TRAP
                                .pvLoop = True
                        End Select
                End Select
            End With
        Next i

        '-- Scroll to start position
        Call DoScrollTo(x:=.ScreenStart, ScaleAndCenter:=False)
    End With
End Sub

'========================================================================================
' Scrolling
'========================================================================================

Public Sub DoScroll( _
           ByVal dx As Long _
           )
    
    '-- Add offset
    m_xScreen = m_xScreen + dx
    
    '-- Scroll
    Call DoScrollTo(x:=m_xScreen, ScaleAndCenter:=False)
End Sub

Public Sub DoScrollTo( _
           ByVal x As Long, _
           Optional ByVal ScaleAndCenter As Boolean = True _
           )
    
    '-- Re-scale and center position
    '   (panoramic view scale 1:2.5)
    If (ScaleAndCenter) Then
        m_xScreen = 2.5 * x - 160
      Else
        m_xScreen = x
    End If
    
    '-- Check bounds
    If (m_xScreen < 0) Then
        m_xScreen = 0
      ElseIf (m_xScreen > 1280) Then
        m_xScreen = 1280
    End If
    
    '-- Define new scroll-selection
    m_uPanoramicViewScanX1 = m_xScreen / 2.5
    m_uPanoramicViewScanX2 = m_xScreen / 2.5 + 127
    
    '-- Render frame?
    If (IsTimerPaused) Then
        Call DoFrame
    End If
End Sub

'========================================================================================
' Paint a frame!
'========================================================================================

Public Sub DoFrame()

    '-- Copy first terrain buffer bits to second one to draw objects on
    Call CopyMemory( _
         ByVal m_oDIBScreenBuffer2.lpBits, _
         ByVal m_oDIBScreenBuffer1.lpBits, _
         m_oDIBScreenBuffer1.Size _
         )
         
    '-- Draw all objects
    Call DrawObjects(m_oDIBScreenBuffer2)
       
    '-- Depending on stage...
    Select Case m_eGameStage
        
        Case [eStageLetsGo]
            
            If (IsTimerPaused = False) Then
            
                '-- Play sound
                If (m_lcGameStage = 0) Then
                    Call PlayMidi(RandomTheme:=True)
                    Call PlaySoundFX([eFXLetsGo])
                End If
                
                '-- Stage frame counter
                m_lcGameStage = m_lcGameStage + 1
                
                '-- 10 frames after...
                If (m_lcGameStage = 10) Then
                    Call SetGameStage([eStageOpeningDoors])
                End If
            End If
            
            '-- Panoramic view
            Call pvDrawPanoramicView
        
        Case [eStageOpeningDoors]
            
            If (IsTimerPaused = False) Then
            
                '-- Stage 1st frame: open doors
                If (m_lcGameStage = 0) Then
                    Call pvSetDoorsState(bOpening:=True)
                End If
                
                '-- Stage frame counter
                m_lcGameStage = m_lcGameStage + 1
                
                '-- 10/15 frames after...
                Select Case m_lcGameStage
                    Case 10
                        Call pvSetDoorsState(bOpening:=False)
                    Case 15
                        Call SetGameStage([eStagePlaying])
                End Select
            End If
            
            '-- Panoramic view
            Call pvDrawPanoramicView
            
        Case [eStagePlaying]
            
            If (IsTimerPaused = False) Then
            
                '-- Check/move lems
                Call pvCheckLems
                '-- Panoramic view
                Call pvDrawPanoramicView
                '-- Render lems
                Call pvDrawLems
              
              Else
                '-- Panoramic view
                Call pvDrawPanoramicView
                '-- Only render lems
                Call pvDrawLems
            End If

        Case [eStageEnding]
            
            '-- Stage 1st  frame: stop 'timer' and close doors
            If (m_lcGameStage = 0) Then
                Call pvSetDoorsState(bOpening:=False)
            End If
            
            '-- Stage frame counter
            If (IsTimerPaused = False) Then
                m_lcGameStage = m_lcGameStage + 1
            End If
            
            '-- Restore DIB palette
            Call m_oScreen.UpdatePalette( _
                 GetGlobalPalette() _
                 )
            
            '-- Panoramic view
            Call pvDrawPanoramicView
            
            '-- 10 frames...
            If (m_lcGameStage < 10) Then
                '-- Fade out
                Call m_oScreen.UpdatePalette( _
                     GetFadedOutGlobalPalette(Amount:=m_lcGameStage * 25) _
                     )
              Else
                '-- Level finished: stop timer and music
                Call StopTimer
                Call CloseMidi
                '-- 'Report'
                Call fLems.LevelDone
                Exit Sub
            End If
    End Select
    
    '-- Paint visible area (main view)
    Call BltFast( _
         m_oScreen.DIB, _
         0, 0, 320, 160, _
         m_oDIBScreenBuffer2, _
         m_xScreen, 0 _
         )
     
    '-- Extra info...
    Call pvShowExtraInfo
    
    '-- Update!
    Call m_oScreen.Refresh
End Sub

'========================================================================================
' Lem disposing, job, ability...
'========================================================================================

Public Sub AddLem()
    
    '-- Lems out counter
    m_lLemsOut = m_lLemsOut + 1
    
    '-- Resize Lems' array
    ReDim Preserve m_uLems(m_lLemsOut)
    
    '-- Define new Lem position and properties
    With m_uLems(m_lLemsOut)
        .x = m_uPtOut(m_lPtOut).x
        .y = m_uPtOut(m_lPtOut).y
        .xs = 1
        .Active = True
    End With
    
    '-- Initially, falling
    Call pvSetLemAnimation( _
         LemID:=m_lLemsOut, _
         Frame:=[eFrameFalling], _
         FrameIdx:=0, _
         f:=DOOR_FALLCOUNTER _
         )
    
    '-- Next Lem, next door (if any)
    m_lPtOut = m_lPtOut + 1
    If (m_lPtOut > m_lcPtOut) Then
        m_lPtOut = 1
    End If
    
    '-- Show info
    fLems.ucInfo.PanelText(3) = "Out " & m_lLemsOut & "/" & g_uLevel.LemsToLetOut
End Sub

Public Sub SetArmageddonLem( _
           ByRef LemID As Long _
           )
    
    With m_uLems(LemID)
        
        '-- Activate 'Armageddon Lem'
        If (.Active) Then
            
            '-- Only if it's not yet activated
            If (.ExplodeCount = 0) Then
                .Ability = .Ability Or [eAbilityBomber]
                .ExplodeCount = EXPLODE_TICKS
            End If
          
          Else
            '-- Not active: find next
            If (LemID < m_lLemsOut) Then
                LemID = LemID + 1
                Call SetArmageddonLem(LemID)
            End If
        End If
    End With
End Sub

Public Function HitTest( _
                ) As String
    
  Dim uPt As POINTAPI
    
    '-- Get cursor position (translate it to game screen coordinates)
    Call GetCursorPos(uPt)
    Call ScreenToClient(fLems.ucScreen.hwnd, uPt)
    
    '-- Screen is 2x zoomed (so apply offset)
    m_xCur = uPt.x \ 2 + m_xScreen
    m_yCur = uPt.y \ 2
    
    '-- Hit-Test now (return Lem's description and number)
    HitTest = pvHitTest()
End Function

Public Sub PrepareAbility( _
           ByVal Ability As eLemAbilityConstants _
           )
    
    '-- Prepare ability to be applied (reset job)
    m_ePreparedAbility = Ability
    m_ePreparedJob = [eJobNone]
End Sub

Public Sub PrepareJob( _
           ByVal Job As eLemJobConstants _
           )
    
    '-- Prepare job to be applied (reset ability)
    m_ePreparedJob = Job
    m_ePreparedAbility = [eAbilityNone]
End Sub

Public Sub ApplyPrepared()
     
  Dim nIdx As Integer
    
    If (IsTimerPaused = False And IsFastForward = False) Then
    
        '-- Apply prepared ability/job
        
        If (m_lLem > 0) Then
            
            With m_uLems(m_lLem)
                
                '-- Almost dead or out?
                If Not (.Frame >= [eFrameSpliting] And .Frame <= [eFrameSurviving]) Then
                    
                    '-- Ability prepared
                    If (m_ePreparedAbility) Then
                        
                        '-- Not yet?
                        If ((.Ability And m_ePreparedAbility) = 0) Then
                        
                            '-- Set ability
                            .Ability = .Ability Or m_ePreparedAbility
                            If (m_ePreparedAbility = [eAbilityBomber]) Then
                                .ExplodeCount = EXPLODE_TICKS
                            End If
                            
                            '-- Update remaining
                            nIdx = (m_ePreparedAbility \ 2) + 1
                            With fLems
                                .lblButton(nIdx).Caption = Val(.lblButton(nIdx).Caption) - 1
                                If (Val(.lblButton(nIdx).Caption) = 0) Then
                                    Call .ucToolbar.CheckButton(nIdx, False)
                                    Call .ucToolbar.EnableButton(nIdx, False)
                                    m_ePreparedAbility = [eAbilityNone]
                                End If
                            End With
                            
                            '-- Play sound
                            Call PlaySoundFX([eFXMousePre])
                        End If
                        
                    '-- Job prepared
                    ElseIf (m_ePreparedJob) Then
                        
                        '-- Can do job?
                        If (pvCanDoJobNow(m_lLem, m_ePreparedJob)) Then
                            
                            '-- A different job?
                            If (.Job <> m_ePreparedJob) Or _
                               (.Frame = [eFrameBuilderEnd]) Then
                                
                                '-- Set job
                                .Job = m_ePreparedJob
                                Select Case .Job
                                    Case [eJobNone]
                                        Call pvSetLemAnimation(m_lLem, [eFrameWalker])
                                    Case [eJobBlocker]
                                        Call pvSetLemAnimation(m_lLem, [eFrameBlocker])
                                    Case [eJobBuilder]
                                        Call pvSetLemAnimation(m_lLem, [eFrameBuilder])
                                    Case [eJobBasher]
                                        Call pvSetLemAnimation(m_lLem, [eFrameBasher])
                                    Case [eJobMiner]
                                        Call pvSetLemAnimation(m_lLem, [eFrameMiner])
                                    Case [eJobDigger]
                                        Call pvSetLemAnimation(m_lLem, [eFrameDigger])
                                End Select
                                
                                '-- Update remaining
                                nIdx = .Job + 3
                                With fLems
                                    .lblButton(nIdx).Caption = Val(.lblButton(nIdx).Caption) - 1
                                    If (Val(.lblButton(nIdx).Caption) = 0) Then
                                        Call .ucToolbar.CheckButton(nIdx, False)
                                        Call .ucToolbar.EnableButton(nIdx, False)
                                        m_ePreparedJob = [eJobNone]
                                    End If
                                End With
                                
                                '-- Play sound
                                Call PlaySoundFX([eFXMousePre])
                            End If
                        End If
                    End If
                End If
            End With
        End If
    End If
End Sub

Public Function GetLemsOut( _
                ) As Long

    '-- All lems let out (+died)
    GetLemsOut = m_lLemsOut
End Function

Public Function GetLemsSaved( _
                ) As Long
  
    '-- Get all saved lems
    GetLemsSaved = m_lLemsSaved
End Function

Public Property Get LemsSavedMode( _
                    ) As eLemsSavedModeConstants
  
    '-- Get 'Lems saved' mode
    LemsSavedMode = m_eLemsSavedMode
End Property

Public Property Let LemsSavedMode( _
                    ByVal Mode As eLemsSavedModeConstants _
                    )
  
    '-- Set 'Lems saved' mode
    m_eLemsSavedMode = Mode
End Property

Public Sub CycleLemsSavedModes()
    
    '-- Cycle through modes
    m_eLemsSavedMode = m_eLemsSavedMode + 1
    If (m_eLemsSavedMode > 3) Then
        m_eLemsSavedMode = 0
    End If
    
    '-- Refresh panel
    fLems.ucInfo.PanelText(4) = pvGetSavedLemsString()
End Sub

Public Property Get InfoState( _
                    ) As Boolean
  
    '-- Get info state
    InfoState = m_bLemsInfo
End Property

Public Property Let InfoState( _
                    ByVal Enable As Boolean _
                    )

    '-- Set info state
    m_bLemsInfo = Enable
    If (IsTimerPaused) Then
        Call DoFrame
    End If
End Property

'========================================================================================
' Menu screens
'========================================================================================

Public Sub ShowLevelFeatures()
    
    With g_uLevel
    
        '-- Disable not available
        Call fLems.ucToolbar.EnableButton(1, .MaxClimbers)
        Call fLems.ucToolbar.EnableButton(2, .MaxFloaters)
        Call fLems.ucToolbar.EnableButton(3, .MaxBombers)
        Call fLems.ucToolbar.EnableButton(4, .MaxBlockers)
        Call fLems.ucToolbar.EnableButton(5, .MaxBuilders)
        Call fLems.ucToolbar.EnableButton(6, .MaxBashers)
        Call fLems.ucToolbar.EnableButton(7, .MaxMiners)
        Call fLems.ucToolbar.EnableButton(8, .MaxDiggers)
        
        '-- Show skills
        fLems.lblButton(1).Caption = .MaxClimbers
        fLems.lblButton(2).Caption = .MaxFloaters
        fLems.lblButton(3).Caption = .MaxBombers
        fLems.lblButton(4).Caption = .MaxBlockers
        fLems.lblButton(5).Caption = .MaxBuilders
        fLems.lblButton(6).Caption = .MaxBashers
        fLems.lblButton(7).Caption = .MaxMiners
        fLems.lblButton(8).Caption = .MaxDiggers
        fLems.lblButton(9).Caption = .ReleaseRate
        fLems.lblButton(10).Caption = .ReleaseRate
        
        '-- Enable Plus/Minus, Pause...
        Call fLems.ucToolbar.EnableButton(9, True)
        Call fLems.ucToolbar.EnableButton(10, True)
        Call fLems.ucToolbar.EnableButton(11, True)
        Call fLems.ucToolbar.EnableButton(12, True)
        Call fLems.ucToolbar.EnableButton(13, True)
        
        '-- Show level title
        fLems.ucInfo.PanelText(1) = RTrim$(.Title)
      
        '-- Show out and saved
        fLems.ucInfo.PanelText(3) = "Out 0/" & .LemsToLetOut
        fLems.ucInfo.PanelText(4) = pvGetSavedLemsString()
        
        '-- Show playing time
        fLems.ucInfo.PanelText(5) = GetMinSecString(.PlayingTime * 60)
    End With
End Sub

Public Sub ShowMenuScreen()

    '-- State flag/menus
    fLems.Tag = [eModeMenuScreen]
    
    With fLems.ucScreen
        
        '-- Tile background
        Call FXTile( _
             .DIB, 0, 0, .DIB.Width, .DIB.Height, _
             [eTileLem] _
             )
        
        '-- Paint menu info
        Call FXText( _
             .DIB, 64, 35, _
             "Press left mouse button to play", _
             IDX_YELLOW _
             )
             
        Call FXText( _
             .DIB, 38, 50, _
             "Press right mouse button to choose level", _
             IDX_YELLOW _
             )
        
        '-- Paint options info
        If (fLems.mnuOptions(0).Checked) Then
            Call FXText( _
                 .DIB, 101, 100, _
                 "[S]ound effects  on", _
                 IDX_GREEN _
                 )
          Else
            Call FXText( _
                 .DIB, 101, 100, _
                 "[S]ound effects  off", _
                 IDX_RED _
                 )
        End If
        If (fLems.mnuOptions(1).Checked) Then
            Call FXText( _
                 .DIB, 101, 115, _
                 "[M]usic          on", _
                 IDX_GREEN _
                 )
          Else
            Call FXText( _
                 .DIB, 101, 115, _
                 "[M]usic          off", _
                 IDX_RED _
                 )
        End If
    
        '-- Refresh
        Call .Refresh
    End With
End Sub

Public Sub ShowLevelScreen()

    '-- State flag/menus
    fLems.Tag = [eModeLevelScreen]
    
    With fLems.ucScreen
        
        '-- Tile background
        Call FXTile( _
             .DIB, 0, 0, .DIB.Width, .DIB.Height, _
             [eTileGround] _
             )
        
        '-- Paint title
        Call FXText( _
             .DIB, 10, 10, _
             "Level " & g_nLevelID Mod 100 + 1, _
             IDX_YELLOW _
             )
        Call FXText( _
             .DIB, 160 - 3 * Len(RTrim$(g_uLevel.Title)), 25, _
             g_uLevel.Title, _
             IDX_YELLOW _
             )
        
        '-- Paint level info
        Call FXText( _
             .DIB, 110, 50, _
             "Number of lems " & g_uLevel.LemsToLetOut, _
             IDX_YELLOW _
             )
        Call FXText( _
             .DIB, 110, 65, _
             pvGetSavedLemsPercentage(g_uLevel.LemsToBeSaved) & " to be saved", _
             IDX_YELLOW _
             )
        Call FXText( _
             .DIB, 110, 80, _
             "Release rate " & g_uLevel.ReleaseRate, _
             IDX_YELLOW _
             )
        Call FXText( _
             .DIB, 110, 95, _
             GetMinSecString(g_uLevel.PlayingTime * 60), _
             IDX_YELLOW _
             )
        Call FXText( _
             .DIB, 110, 110, _
             "Rating " & GetLevelRatingString(g_nLevelID), _
             IDX_YELLOW _
             )
        
        '-- Paint 'continue' message
        Call FXText( _
             .DIB, 70, 140, _
             "Press mouse button to continue", _
             IDX_YELLOW _
             )
    
        '-- Refresh
        Call .Refresh
    End With
End Sub

Public Sub ShowGameScreen()

    '-- State flag
    fLems.Tag = [eModeGameScreen]
    
    With fLems.ucScreen
        
        '-- Tile background
        Call FXTile( _
             .DIB, 0, 0, .DIB.Width, .DIB.Height, _
             [eTileGround] _
             )
        
        '-- Paint game info
        Call FXText( _
             .DIB, 94, 15, _
             "All lems accounted for", _
             IDX_YELLOW _
             )
        Call FXText( _
             .DIB, 115, 50, _
             "You rescued " & pvGetSavedLemsPercentage(m_lLemsSaved), _
             IDX_YELLOW _
             )
        Call FXText( _
             .DIB, 115, 65, _
             "You needed  " & pvGetSavedLemsPercentage(g_uLevel.LemsToBeSaved), _
             IDX_YELLOW _
             )
        
        '-- Paint 'retry level' or 'continue' message
        If (m_lLemsSaved < g_uLevel.LemsToBeSaved) Then
            Call FXText( _
                 .DIB, 46, 125, _
                 "Press left mouse button to retry level", _
                 IDX_YELLOW _
                 )
            Call FXText( _
                 .DIB, 61, 140, _
                 "Press right mouse button for menu", _
                 IDX_YELLOW _
                 )
          Else
            Call FXText( _
                 .DIB, 70, 140, _
                 "Press mouse button to continue", _
                 IDX_YELLOW _
                 )
            Call SetLevelDone( _
                 g_nLevelID, Done:=True _
                 )
        End If
    
        '-- Refresh
        Call .Refresh
    End With
End Sub

Public Sub ProcessCheatCode( _
           ByVal KeyCode As KeyCodeConstants _
           )
    
    '-- Add to cheat string
    m_sCheat = m_sCheat & Chr$(KeyCode)
    
    '-- Process last cheat-lenght chars
    If (Len(m_sCheat) > Len(CHEAT_CODE)) Then
        m_sCheat = Mid$(m_sCheat, 2)
    End If
    
    '-- Is that our cheat code
    If (UCase$(m_sCheat) = CHEAT_CODE) Then
        Call SetLevelDone(g_nLevelID, Done:=True)
    End If
End Sub

'----------------------------------------------------------------------------------------
' Game stage [0: 'Lets go', 1: 'Opening doors', 2: Playing, 3: 'Ending']
'----------------------------------------------------------------------------------------

Public Sub SetGameStage( _
           ByVal New_GameStage As eGameStageConstants _
           )
    
    '-- Set current stage
    If (New_GameStage <> m_eGameStage) Then
        m_eGameStage = New_GameStage
        m_lcGameStage = 0
    End If
End Sub

Public Function GetGameStage( _
                ) As eGameStageConstants
    
    '-- Return current stage
    GetGameStage = m_eGameStage
End Function

'----------------------------------------------------------------------------------------
' Copy whole screen
'----------------------------------------------------------------------------------------

Public Sub CopyWholeScreenToClipboard()

    '-- Get current palette and copy to clipboard
    Call m_oDIBScreenBuffer2.SetPalette(GetGlobalPalette())
    Call m_oDIBScreenBuffer2.CopyToClipboard
End Sub

'----------------------------------------------------------------------------------------
' Rendering terrain, objects
'----------------------------------------------------------------------------------------

Public Sub DrawTerrain( _
           oDIB As cDIB08 _
           )
            
  Dim i As Long
  
    With g_uLevel
    
        If (.GraphicSetEx > 0) Then
            Call MaskBlt( _
                 oDIB, _
                 304, 0, _
                 960, 160, _
                 g_oDIBBack, _
                 0, 0, _
                 IDX_TRANS _
                 )
          Else
            For i = 1 To .TerrainPieces
                With .TerrainPiece(i)
                    If (.Black) Then
                        If (.NotOverlap) Then
                            Call MaskBltIdxOverlap( _
                                 oDIB, _
                                 .lpRect.x1, .lpRect.y1, _
                                 g_uTerGFX(.ID).Width, g_uTerGFX(.ID).Height, _
                                 IDX_NONE, IDX_NONE, _
                                 g_uTerGFX(.ID).DIB, _
                                 0, 0, _
                                 IDX_TRANS, _
                                 .UpsideDown _
                                 )
                          Else
                            Call MaskBltIdx( _
                                 oDIB, _
                                 .lpRect.x1, .lpRect.y1, _
                                 g_uTerGFX(.ID).Width, g_uTerGFX(.ID).Height, _
                                 IDX_NONE, _
                                 g_uTerGFX(.ID).DIB, _
                                 0, 0, _
                                 IDX_TRANS, _
                                 .UpsideDown _
                                 )
                        End If
                      Else
                        If (.NotOverlap) Then
                            Call MaskBltOverlap( _
                                 oDIB, _
                                 .lpRect.x1, .lpRect.y1, _
                                 g_uTerGFX(.ID).Width, g_uTerGFX(.ID).Height, _
                                 IDX_NONE, _
                                 g_uTerGFX(.ID).DIB, _
                                 0, 0, _
                                 IDX_TRANS, _
                                 .UpsideDown _
                                 )
                          Else
                            Call MaskBlt( _
                                 oDIB, _
                                 .lpRect.x1, .lpRect.y1, _
                                 g_uTerGFX(.ID).Width, g_uTerGFX(.ID).Height, _
                                 g_uTerGFX(.ID).DIB, _
                                 0, 0, _
                                 IDX_TRANS, _
                                 .UpsideDown _
                                 )
                        End If
                    End If
                End With
            Next i
        End If
    End With
End Sub

Public Sub DrawObjects( _
           oDIB As cDIB08 _
           )
    
  Dim i As Long
    
    With g_uLevel
    
        '-- First draw 'on-terrain' objects
        For i = 1 To .Objects
            If (.Object(i).OnTerrain) Then
                Call DrawObject(i, oDIB)
            End If
        Next i
        
        '-- Then, any other object
        For i = 1 To .Objects
            If (.Object(i).OnTerrain = False) Then
                Call DrawObject(i, oDIB)
            End If
        Next i
    End With
End Sub

Public Sub DrawObject( _
           ByVal Idx As Integer, _
           oDIB As cDIB08 _
           )
        
    With g_uLevel.Object(Idx)
        
        '-- Animate object?
        If ((IsTimerPaused = False) And .pvLoop) Then
            
            '-- Trap? Play FX now?
            If (g_uObjGFX(.ID).TriggerEffect = IDX_TRAP) Then
                If (.pvFrameIdxCur = g_uObjGFX(.ID).SoundEffectAtFrame) Then
                    Call PlaySoundFX(g_uObjGFX(.ID).SoundEffect)
                End If
            End If
            
            '-- Animate
            .pvFrameIdxCur = .pvFrameIdxCur + 1
            If (.pvFrameIdxCur = .pvFrameIdxMax) Then
                .pvFrameIdxCur = 0
                
                '-- Trap? (single loop)
                If (g_uObjGFX(.ID).TriggerEffect = IDX_TRAP) Then
                    .pvLoop = False
                End If
            End If
        End If
        
        '-- Rendering object
        If (.OnTerrain) Then
            Call MaskBltOverlapNot( _
                 oDIB, _
                 .lpRect.x1, .lpRect.y1, _
                 g_uObjGFX(.ID).Width, g_uObjGFX(.ID).Height, _
                 IDX_NONE, _
                 g_uObjGFX(.ID).DIB, _
                 0, 1& * .pvFrameIdxCur * g_uObjGFX(.ID).Height, _
                 IDX_TRANS _
                 )
          Else
            If (.NotOverlap) Then
                Call MaskBltOverlap( _
                     oDIB, _
                     .lpRect.x1, .lpRect.y1, _
                     g_uObjGFX(.ID).Width, g_uObjGFX(.ID).Height, _
                     IDX_NONE, _
                     g_uObjGFX(.ID).DIB, _
                     0, 1& * .pvFrameIdxCur * g_uObjGFX(.ID).Height, _
                     IDX_TRANS _
                     )
              Else
                Call MaskBltOverlapNot( _
                     oDIB, _
                     .lpRect.x1, .lpRect.y1, _
                     g_uObjGFX(.ID).Width, g_uObjGFX(.ID).Height, _
                     IDX_BRICK, _
                     g_uObjGFX(.ID).DIB, _
                     0, 1& * .pvFrameIdxCur * g_uObjGFX(.ID).Height, _
                     IDX_TRANS _
                     )
            End If
        End If
    End With
End Sub

Public Function GetLevelThumbnail( _
                oDIBThumbnail As cDIB08 _
                ) As Boolean

  Dim sPath   As String
  Dim bCreate As Boolean
  Dim oDIB    As New cDIB08
  
    '-- First check thumbnail-file exists
    sPath = AppPath & "LEVELS\" & Format$(g_nLevelID, "0000")
    If (Not FileExists(sPath & ".bmp")) Then
        bCreate = True
      Else
        '-- Now check file date-time stamp
        If (VBA.FileDateTime(sPath & ".dat") <> GetLevelThumbnailDateTimeStamp(g_nLevelID)) Then
            bCreate = True
        End If
    End If
        
    If (bCreate) Then

        '-- Create it now
        Call mLevel.LoadLevel(g_nLevelID)
        Call oDIB.Create(1600, 160)
        Call oDIBThumbnail.Create(320, 40)
        Call oDIBThumbnail.SetPalette(GetGlobalPalette())
        Call DrawTerrain(oDIB)
        Call DrawObjects(oDIB)
        Call FXStretch(oDIBThumbnail, oDIB)

        '-- And save it
        Call oDIBThumbnail.Save(sPath & ".bmp")
        
        '-- Set date-time stamp
        Call SetLevelThumbnailDateTimeStamp(g_nLevelID, VBA.FileDateTime(sPath & ".dat"))
        
        '-- Created
        GetLevelThumbnail = True
        
      Else
        '-- From cache
        Call oDIBThumbnail.CreateFromBitmapFile(sPath & ".bmp")
   End If
End Function

'========================================================================================
' Private
'========================================================================================

'----------------------------------------------------------------------------------------
' Checking all lems (most important routine... and longest)
'----------------------------------------------------------------------------------------
'
' Lem frames are 16x16: left-top (0,0) to right-bottom (15,15)
' Lems base-line are located at bottom-line (y = 15) and horizontaly centered.
' Offsets are applied on animation changing.
' In-animation offsets (adjusts) are applied here.

Private Sub pvCheckLems()
  
  Dim l As Long, bOneActive As Boolean
  Dim i As Long, j As Long, ix As Long, iy As Long, px As Byte
  Dim bSkip As Boolean
    
    '-- Check all lems
    For l = 1 To m_lLemsOut
        
        With m_uLems(l)
            
            '-- Is active?
            If (.Active) Then
                
                '-- Reset 'skip code' flag
                bSkip = False
                
                '-- At least, one is active
                bOneActive = True
                
                '-- Next frame now
                .FrameIdx = .FrameIdx + 1
                If (.FrameIdx > .FrameIdxMax) Then
                    .FrameIdx = 0
                End If
                
                '-- Start checking...
                Select Case .Job
                    
                    Case [eJobNone]
                
                        Select Case .Frame
                        
                            Case [eFrameWalker]
                                
                                '-- One step forward
                                .x = .x + .xs
                                
                                '-- Check feet
                                ix = 7 - (.xs < 0)
                                iy = 16
                                
                                If (pvCheckPixel(.x + ix, .y + iy, l, True)) Then

                                    '-- Can go up?
                                    For i = 0 To MIN_OBSTACLE - 1
                                        Select Case pvCheckPixel(.x + ix, .y + iy - 1, l, True)
                                            Case IDX_NONE
                                                Exit For
                                            Case IDX_NULL
                                                bSkip = True
                                                Exit For
                                            Case IDX_BLOCKER
                                                .xs = -.xs
                                                .x = .x + 2 * .xs
                                                .y = .y + i
                                                bSkip = True
                                                Exit For
                                            Case Else
                                                .y = .y - 1
                                        End Select
                                    Next i
                                    If (bSkip) Then GoTo lblSkip
                                        
                                    '-- Obstacle: can climb?
                                    If (i = MIN_OBSTACLE) Then
                                        .y = .y + i
                                        .x = .x - .xs
                                        If (.Ability And [eAbilityClimber]) Then
                                            Call pvSetLemAnimation(l, [eFrameClimber])
                                          Else
                                            .xs = -.xs
                                        End If
                                    End If
                                  
                                  Else
                                    
                                    '-- Almost falling?
                                    For i = 0 To MIN_OBSTACLE - 1
                                        Select Case pvCheckPixel(.x + ix, .y + iy, l, True)
                                            Case IDX_NONE, IDX_BLOCKER
                                                .y = .y + 1
                                            Case IDX_NULL
                                                bSkip = True
                                                Exit For
                                            Case Else
                                                Exit For
                                        End Select
                                    Next i
                                    If (bSkip) Then GoTo lblSkip
                                    
                                    '-- Yes, just falling
                                    If (i = MIN_OBSTACLE) Then
                                        Call pvSetLemAnimation(l, [eFrameFalling])
                                        .x = .x - .xs
                                        .y = .y - i + 3
                                        .f = 3
                                        GoTo lblSkip
                                    End If
                                End If
                                
                                '-- Check ahead (blocker)
                                ix = 9 + 3 * (.xs < 0)
                                iy = 10
                                If (pvGetPixelLo(.x + .xs + ix, .y + iy, l) = IDX_BLOCKER) Then
                                    .xs = -.xs
                                    .x = .x + .xs
                                End If
                                
                            Case [eFrameFalling]
                                
                                '-- 3-pixel loop
                                ix = 8 + (.xs < 0)
                                iy = 15
                                
                                For i = 1 To 3
                                    
                                    '-- One step down
                                    .y = .y + 1
                                    
                                    '-- Floater?
                                    If (.f >= MIN_FALL_FLOATER) Then
                                        If (.Ability And [eAbilityFloater]) Then
                                            Call pvSetLemAnimation(l, [eFrameFloater])
                                            Exit For
                                        End If
                                    End If
                                    
                                    '-- A soft landing?
                                    Select Case pvCheckPixel(.x + ix, .y + iy, l, True)
                                        Case IDX_NULL
                                            Exit For
                                        Case IDX_NONE, IDX_BLOCKER
                                        Case Else
                                            If (.f > MAX_FALL) Then
                                                Call pvSetLemAnimation(l, [eFrameSpliting])
                                                Exit For
                                              Else
                                                Call pvSetLemAnimation(l, [eFrameWalker])
                                                Exit For
                                            End If
                                    End Select
                                      
                                    '-- Animation counter (max fall)
                                    .f = .f + 1
                              Next i
                                
                            Case [eFrameSpliting]
                                                                      
                                Select Case .FrameIdx
                                    Case 1
                                        Call PlaySoundFX([eFXSplat])
                                    Case .FrameIdxMax
                                        .Active = False
                                End Select
                                
                            Case [eFrameDrowning]
                                    
                                '-- Move?
                                ix = 13 + 11 * (.xs < 0)
                                iy = 16
                                If (pvGetPixelHi(.x + ix, .y + iy, l) = IDX_LIQUID) Then
                                    .x = .x + .xs
                                End If
                                 
                                Select Case .FrameIdx
                                    Case 1
                                        Call PlaySoundFX([eFXGlug])
                                    Case .FrameIdxMax
                                        .Active = False
                                End Select
                            
                            Case [eFrameBurning]
                                
                                Select Case .FrameIdx
                                    Case 1
                                        Call PlaySoundFX([eFXFire])
                                    Case .FrameIdxMax
                                        .Active = False
                                End Select
                                
                            Case [eFrameExploding]
                            
                                '-- Play sound
                                If (.FrameIdx = 11 And IsArmageddonActivated = False) Then
                                    Call PlaySoundFX([eFXOhNo])
                                End If
                                
                                '-- Check feet
                                iy = 16
                                For j = 1 To 3
                                    i = 0
                                    For ix = 8 To 9
                                        Select Case pvCheckPixel(.x + ix, .y + iy, l, True)
                                            Case IDX_NULL
                                                bSkip = True
                                                Exit For
                                            Case IDX_NONE, IDX_BLOCKER
                                                i = i + 1
                                        End Select
                                    Next ix
                                    If (bSkip) Then
                                        Exit For
                                      ElseIf (i = 2) Then
                                        If (.f = 0) Then
                                            .f = 1
                                            Call pvRemoveBlockerMask(l)
                                        End If
                                        .y = .y + 1
                                    End If
                                Next j
                                
                                '-- Almost...
                                If (.FrameIdx = .FrameIdxMax) Then
                                    
                                    '-- Explode
                                    Call PlaySoundFX([eFXExplode])
                                    If (.f = 0) Then
                                        .f = 1
                                        Call pvRemoveBlockerMask(l)
                                    End If
                                    
                                    '-- Apply mask
                                    Call pvDrawMask(l)
                                    
                                    '-- One less
                                    .Active = False
                                End If
                            
                            Case [eFrameSurviving]
                                
                                Select Case .FrameIdx
                                    
                                    Case 1
                                    
                                        '-- Play sound
                                        Call PlaySoundFX([eFXYipee])
                                        
                                        '-- Save it now
                                        m_lLemsSaved = m_lLemsSaved + 1
                                        fLems.ucInfo.PanelText(4) = pvGetSavedLemsString()
                                 
                                    Case .FrameIdxMax
                                    
                                        '-- One less (but, saved)
                                        .Active = False
                                End Select
                                
                            Case [eFrameClimber]
                                
                                '-- Adjust animation (small obstacle)
                                If (.f = 0) Then
                                    
                                    '-- 4-pixel loop
                                    ix = 8 + (.xs < 0)
                                    iy = 7
                                    If (pvGetPixelLo(.x + ix, .y + iy, l) = IDX_NONE) Then
                                        For i = 0 To 4
                                            Select Case pvCheckPixel(.x + ix, .y + iy + i, l)
                                                Case IDX_NULL
                                                    bSkip = True
                                                    Exit For
                                                Case Is <> IDX_NONE
                                                    .y = .y - 1
                                            End Select
                                        Next i
                                        If (bSkip) Then
                                            GoTo lblSkip
                                          Else
                                            .y = .y + 3
                                        End If
                                        
                                        '-- Animation end
                                        Call pvSetLemAnimation(l, [eFrameClimberEnd], 1)
                                        GoTo lblSkip
                                    End If
                                End If
                                
                                Select Case .FrameIdx
                                
                                    Case 4 To .FrameIdxMax
                                    
                                        '-- Climbing
                                        ix = 8 + (.xs < 0)
                                        iy = 6
                                        
                                        '-- Insurmountable obstacle?
                                        Select Case pvCheckPixel(.x + ix, .y + iy + i, l)
                                            Case IDX_NULL
                                                GoTo lblSkip
                                            Case IDX_NONE
                                                .y = .y - 1
                                                Call pvSetLemAnimation(l, [eFrameClimberEnd])
                                                GoTo lblSkip
                                        End Select
                                        
                                        '-- Can't climb
                                        Select Case pvCheckPixel(.x + ix - .xs, .y + iy + i, l)
                                            Case IDX_NULL
                                                GoTo lblSkip
                                            Case Is <> IDX_NONE
                                                .x = .x - .xs
                                                .xs = -.xs
                                                Call pvSetLemAnimation(l, [eFrameFalling])
                                                GoTo lblSkip
                                        End Select
                                        
                                        '-- Must fall?
                                        For iy = 9 To 13
                                            Select Case pvCheckPixel(.x + ix, .y + iy + i, l)
                                                Case IDX_NULL
                                                    bSkip = True
                                                    Exit For
                                                Case IDX_NONE
                                                    Call pvSetLemAnimation(l, [eFrameFalling])
                                                    bSkip = True
                                                    Exit For
                                            End Select
                                        Next iy
                                        If (bSkip) Then GoTo lblSkip
                                        
                                        .y = .y - 1
                                End Select
                                
                                '-- Animation counter (loop)
                                If (.FrameIdx = .FrameIdxMax) Then
                                    .f = .f + 1
                                End If
                            
                            Case [eFrameClimberEnd]
                            
                                If (.FrameIdx = .FrameIdxMax) Then
                                
                                    '-- Adjust animation
                                    .y = .y - 7
                                    
                                    '-- Walk again
                                    Call pvSetLemAnimation(l, [eFrameWalker])
                                End If
                                
                            Case [eFrameFloater]
                                
                                Select Case .FrameIdx
                                
                                    Case 0
                                    
                                        '-- Adjust animation
                                        .FrameIdx = 4
                                        
                                    Case 4, 5
                                    
                                        '-- Decelerate when open
                                        If (.f < 8) Then
                                            .y = .y - 1 + 2 * (.FrameIdx = 5 And .f = 4)
                                        End If
                                End Select
                                
                                Select Case .f
                                    
                                    Case 4 To 6
                                        
                                        '-- Hold frame 4 on first loop
                                        .FrameIdx = 4
                                    
                                    Case Is > 6
                                    
                                        '-- Slow down animation
                                        If (.f Mod 2 = 0 And .FrameIdx <> .FrameIdxMax) Then
                                            .FrameIdx = .FrameIdx - 1
                                        End If
                                End Select
                                
                                '-- A soft landing?
                                ix = 8 + (.xs < 0)
                                iy = 14
                                For i = 1 To 2
                                    .y = .y + 1
                                    Select Case pvCheckPixel(.x + ix, .y + iy, l, True)
                                        Case IDX_NULL
                                            Exit For
                                        Case Is <> IDX_NONE
                                            .y = .y - 1
                                            Call pvSetLemAnimation(l, [eFrameWalker])
                                            Exit For
                                    End Select
                                Next i
                                
                                '-- Animation counter
                                .f = .f + 1
                             
                             Case Else
                                
                                '-- Something wrong... ?!
                                Call pvSetLemAnimation(l, [eFrameWalker])
                        End Select
                    
                    Case [eJobBlocker]
                        
                        '-- Draw mask
                        If (.f = 0) Then
                            .f = 1
                            Call pvDrawMask(l)
                        End If
                        
                        '-- 'Un-blocked'?
                        Select Case .FrameIdx
                        
                            Case 0, 4, 8, 12
                            
                                '-- Check feet line
                                iy = 16
                                i = 0
                                For ix = 7 To 10
                                    Select Case pvCheckPixel(.x + ix, .y + iy, l, True)
                                        Case IDX_NULL
                                            Call pvRemoveBlockerMask(l)
                                            Exit For
                                        Case IDX_NONE, IDX_BLOCKER
                                            i = i + 1
                                    End Select
                                Next ix
                                If (i = 4) Then
                                    .Job = [eJobNone]
                                    Call pvRemoveBlockerMask(l)
                                    Call pvSetLemAnimation(l, [eFrameFalling])
                                End If
                        End Select

                    Case [eJobBuilder]
                        
                        Select Case .Frame
                        
                            Case [eFrameBuilder]
                                
                                '-- Last three bricks?
                                If (.f >= MAX_BRICKS - 3) Then
                                    '-- Play sound
                                    If (.FrameIdx = 9) Then
                                        Call PlaySoundFX([eFXChink])
                                    End If
                                End If
                                
                                Select Case .FrameIdx
                                  
                                    Case 0
                                        
                                        '-- One brick up
                                        .x = .x + 2 * .xs
                                        .y = .y - 1
                                                                        
                                        '-- Work finished?
                                        If (.f = MAX_BRICKS) Then
                                            Call pvSetLemAnimation(l, [eFrameBuilderEnd])
                                            GoTo lblSkip
                                        End If
                                        
                                        '-- Can continue (feet)?
                                        ix = 8 + (.xs < 0)
                                        iy = 15
                                        Select Case pvCheckPixel(.x + ix, .y + iy, l, True)
                                            Case IDX_NULL
                                                GoTo lblSkip
                                            Case IDX_BLOCKER
                                                .xs = -.xs
                                                GoTo lblSkip
                                            Case Is <> IDX_NONE
                                                .xs = -.xs
                                                .x = .x + 2 * .xs
                                                .Job = [eJobNone]
                                                Call pvSetLemAnimation(l, [eFrameWalker])
                                                GoTo lblSkip
                                        End Select
                                        
                                        '-- Can continue (ahead-blocker)?
                                        ix = 12 + 9 * (.xs < 0)
                                        iy = 11
                                        Select Case pvCheckPixel(.x + ix, .y + iy, l)
                                            Case IDX_NULL
                                                GoTo lblSkip
                                            Case IDX_BLOCKER
                                                .xs = -.xs
                                                GoTo lblSkip
                                        End Select

                                        '-- Can continue (ahead-all)?
                                        ix = 9 + 3 * (.xs < 0)
                                        iy = 11
                                        Select Case pvCheckPixel(.x + ix, .y + iy, l)
                                            Case IDX_NULL
                                                GoTo lblSkip
                                            Case IDX_BLOCKER
                                                .xs = -.xs
                                                GoTo lblSkip
                                            Case Is <> IDX_NONE
                                                .xs = -.xs
                                                .Job = [eJobNone]
                                                Call pvSetLemAnimation(l, [eFrameWalker])
                                                GoTo lblSkip
                                        End Select

                                        '-- Can continue (head)?
                                        ix = 10 + 5 * (.xs < 0)
                                        iy = 7
                                        Select Case pvCheckPixel(.x + ix, .y + iy, l)
                                            Case IDX_NULL
                                                GoTo lblSkip
                                            Case Is <> IDX_NONE
                                                .xs = -.xs
                                                .Job = [eJobNone]
                                                Call pvSetLemAnimation(l, [eFrameWalker])
                                        End Select

                                    Case 9
                                    
                                        Call pvDrawMask(l)
                                  
                                    Case .FrameIdxMax
                                    
                                        '-- Animation counter
                                        .f = .f + 1
                                End Select
                            
                            Case [eFrameBuilderEnd]
                            
                                '-- Hey! Nothing to do?
                                If (.FrameIdx = .FrameIdxMax) Then
                                    .Job = [eJobNone]
                                    Call pvSetLemAnimation(l, [eFrameWalker])
                                End If
                        End Select
                       
                    Case [eJobBasher]
                        
                        Select Case .FrameIdx
                          
                            Case 1, 17

                                '-- Can continue?
                                ix = 15 + 15 * (.xs < 0)
                                iy = 10
                                px = pvCheckPixel(.x + ix, .y + iy, l)
                                If Not ((px = IDX_NULL) Or _
                                        (px = IDX_TERRAIN) Or _
                                        (px = IDX_BASHRIGHT And .xs > 0) Or _
                                        (px = IDX_BASHLEFT And .xs < 0)) Then
                                        .f = 1
                                End If
                        End Select
                                               
                        Select Case .FrameIdx
                          
                            Case 1 To 4
                            
                                '-- Draw mask now
                                Call pvDrawMask(l, .FrameIdx - 1)
                          
                            Case 17 To 20
                            
                                '-- Draw mask now
                                Call pvDrawMask(l, .FrameIdx - 17)
                            
                            Case 11 To 15, 27 To .FrameIdx
                            
                                '-- Adjust animation
                                .x = .x + .xs
                                  
                                '-- Check feet (exit, trap...)
                                ix = 8 + (.xs < 0)
                                iy = 15
                                If (pvCheckPixel(.x + ix, .y + iy, l, True) = IDX_NULL) Then
                                    GoTo lblSkip
                                End If
                        End Select
                        
                        Select Case .FrameIdx
                          
                            Case 4, 20
                                
                                '-- Stop
                                If (.f = 1) Then
                                    .Job = [eJobNone]
                                    Call pvSetLemAnimation(l, [eFrameWalker])
                                End If
                            
                            Case Else
                            
                                '-- Falling?
                                ix = 8 + (.xs < 0)
                                For iy = 16 To 17
                                    If (pvGetPixelLo(.x + ix, .y + iy, l) <> IDX_NONE) Then
                                        Exit For
                                    End If
                                Next iy
                                Select Case iy
                                    Case 17 ' 1 pixel
                                        .y = .y + 1
                                    Case 18 ' 2+ pixel
                                        .Job = [eJobNone]
                                        Call pvSetLemAnimation(l, [eFrameFalling])
                                End Select
                        End Select

                    Case [eJobMiner]
                        
                        Select Case .FrameIdx
                          
                            Case 0
                            
                                '-- Advance
                                .x = .x + .xs
                                .y = .y + 2
                            
                            Case 1
                              
                                '-- Adjust animation
                                .x = .x + .xs
                                
                                '-- Check ahead (blocker)
                                iy = 8
                                ix = 10 + 5 * (.xs < 0)
                                If (pvGetPixelLo(.x + ix, .y + iy, l) = IDX_BLOCKER) Then
                                    .xs = -.xs
                                End If
                               
                                '-- Draw mask now
                                Call pvDrawMask(l)
                                              
                            Case 3
                            
                                '-- Can continue (feet)?
                                ix = 12 + 9 * (.xs < 0)
                                iy = 16
                                px = pvCheckPixel(.x + ix, .y + iy, l)
                                If (px = IDX_NULL) Then
                                    GoTo lblSkip
                                  Else
                                    If ((px = IDX_TERRAIN) Or _
                                        (px = IDX_BASHRIGHT And .xs > 0) Or _
                                        (px = IDX_BASHLEFT And .xs < 0)) Then
                                      Else
                                        .Job = [eJobNone]
                                        Call pvSetLemAnimation(l, [eFrameWalker])
                                        GoTo lblSkip
                                    End If
                                End If
                                
                                '-- Can continue (ahead)?
                                ix = 14 + 13 * (.xs < 0)
                                iy = 10
                                If (px = IDX_NULL) Then
                                    GoTo lblSkip
                                  Else
                                    If ((px = IDX_TERRAIN) Or _
                                        (px = IDX_BASHRIGHT And .xs > 0) Or _
                                        (px = IDX_BASHLEFT And .xs < 0)) Then
                                      Else
                                        .Job = [eJobNone]
                                        Call pvSetLemAnimation(l, [eFrameWalker])
                                    End If
                                End If
                                
                            Case 6 To 14
                            
                                '-- Can continue (feet)?
                                ix = 9 + 3 * (.xs < 0)
                                iy = 15
                                px = pvCheckPixel(.x + ix, .y + iy, l, True)
                                If (px = IDX_NULL) Then
                                    GoTo lblSkip
                                  Else
                                    If ((px = IDX_TERRAIN) Or _
                                        (px = IDX_BASHRIGHT And .xs > 0) Or _
                                        (px = IDX_BASHLEFT And .xs < 0)) Then
                                      Else
                                        .Job = [eJobNone]
                                        Call pvSetLemAnimation(l, [eFrameWalker])
                                        GoTo lblSkip
                                    End If
                                End If
                                
                            Case 15 To .FrameIdxMax
                                
                                '-- Adjust animation
                                If (.FrameIdx = 15) Then
                                    .x = .x + 2 * .xs
                                End If
                                
                                '-- Can continue (feet)?
                                ix = 9 + 3 * (.xs < 0)
                                iy = 16
                                px = pvCheckPixel(.x + ix, .y + iy, l, True)
                                If (px = IDX_NULL) Then
                                    GoTo lblSkip
                                  Else
                                    If ((px = IDX_TERRAIN) Or _
                                        (px = IDX_BASHRIGHT And .xs > 0) Or _
                                        (px = IDX_BASHLEFT And .xs < 0)) Then
                                      Else
                                        .Job = [eJobNone]
                                        Call pvSetLemAnimation(l, [eFrameWalker])
                                        GoTo lblSkip
                                    End If
                                End If
                        End Select
                    
                    Case [eJobDigger]
                            
                        Select Case .FrameIdx
                          
                            Case 1, 9
                            
                                '-- Check feet
                                iy = 14
                                For ix = 7 + (.xs < 0) To 9 + (.xs < 0)
                                    px = pvCheckPixel(.x + ix, .y + iy, l, True)
                                    If (px <> IDX_NONE) Then
                                        Exit For
                                    End If
                                Next ix
                                If (px = IDX_NULL) Then
                                    GoTo lblSkip
                                  Else
                                    If ((px = IDX_TERRAIN) Or _
                                        (px = IDX_BASHRIGHT) Or _
                                        (px = IDX_BASHLEFT)) Then
                                        '-- Draw mask now
                                        Call pvDrawMask(l)
                                        .y = .y + 1
                                      ElseIf (px = IDX_STEEL) Then
                                        .Job = [eJobNone]
                                        Call pvSetLemAnimation(l, [eFrameWalker])
                                      Else
                                        '-- Falling?
                                        .Job = [eJobNone]
                                        Call pvSetLemAnimation(l, [eFrameFalling])
                                    End If
                                End If
                        End Select
                End Select
                
                '-- Bomber ability?
lblSkip:        If (.Ability And [eAbilityBomber]) Then
                    
                    '-- Already counting-down?
                    If (.ExplodeCount > 0) Then
                        
                        '-- Count-down
                        .ExplodeCount = .ExplodeCount - 1
                      
                      Else
                        
                        '-- Remove ability and job
                        .Ability = .Ability And Not [eAbilityBomber]
                        .Job = [eJobNone]
                                        
                        '-- Particles rendering activated
                        .Particles = True
                        
                        '-- Initialize particles
                        For i = 0 To MAX_PARTICLES
                            .Particle(i).x = .x + 8 + (VBA.Rnd * 8 - 4)
                            .Particle(i).y = .y + 12 + VBA.Rnd * 2
                            .Particle(i).vx = VBA.Rnd * 16 - 8
                            .Particle(i).vy = -VBA.Rnd * 8 - 4
                        Next i
                        
                        '-- Avoid animation?
                        If (.Frame = [eFrameFalling] Or _
                            .Frame = [eFrameFloater] Or _
                            .Frame = [eFrameClimber] Or _
                            .Frame = [eFrameClimberEnd]) Then
                            
                            '-- Explode
                            Call PlaySoundFX([eFXExplode])
                            Call pvSetLemAnimation(l, [eFrameExploding])
                            Call pvDrawMask(l)
                            
                            '-- Sorry
                            .Active = False
                            
                          Else
                            '-- Pre-explosion animation
                            Call pvSetLemAnimation(l, [eFrameExploding])
                        End If
                    End If
                End If
                
              Else
                    
                '-- Render particles?
                If (.Particles) Then
                    '-- Active in case remaining particles
                    bOneActive = True
                End If
            End If
        End With
    Next l
    
    '-- None active?
    If (bOneActive = False) Then
        If ((m_lLemsOut = g_uLevel.LemsToLetOut) Or IsArmageddonActivated) Then
            Call SetGameStage([eStageEnding])
        End If
    End If
End Sub

'----------------------------------------------------------------------------------------
' Render all lems
'----------------------------------------------------------------------------------------

Private Sub pvDrawLems()
   
  Dim l As Long
  
    For l = 1 To m_lLemsOut
        
        With m_uLems(l)
            
            If (.Active) Then

                '-- Draw frame
                Call MaskBlt( _
                     m_oDIBScreenBuffer2, _
                     .x, .y, _
                     16, 16, _
                     m_oDIBLems, _
                     .FrameIdx * 17, .FrameSrcY - 17 * (.xs < 0 And .FrameHasDir), _
                     IDX_TRANS _
                     )
                     
                '-- Count-down?
                If (.Ability And [eAbilityBomber]) Then
                    Call pvDrawCountDown(l)
                End If
                
                '-- Preserving a last frame (traps)
                If (.DieNextFrame) Then
                    .Active = False
                End If
            
              Else
              
                '-- Explosion/Particles?
                If (.Frame = [eFrameExploding]) Then
                    .Frame = [eFrameNone]
                    Call pvDrawExplosion(l)
                End If
                If (.Particles) Then
                    Call pvDrawParticles(l)
                End If
            End If
        End With
    Next l
End Sub

'----------------------------------------------------------------------------------------
' Set current lem animation
'----------------------------------------------------------------------------------------

Private Sub pvSetLemAnimation( _
            ByVal LemID As Long, _
            ByVal Frame As eLemFrameConstants, _
            Optional ByVal FrameIdx As Long = 0, _
            Optional ByVal f As Long = 0 _
            )
                                        
    With m_uLems(LemID)
    
        '-- Special cases...
        
        '   y offset...
        Select Case .Frame
            Case [eFrameFalling], [eFrameFloater]
                If (Frame = [eFrameWalker]) Then
                    .y = .y - 1
                End If
            Case [eFrameBuilder]
                If (.FrameIdx > 13) Then
                    .y = .y - 1
                End If
            Case [eFrameMiner]
                If (Frame <> [eFrameBasher]) Then
                    Select Case .FrameIdx
                        Case 0
                            .y = .y - 2
                        Case 1 To 8
                            .y = .y - 1
                    End Select
                End If
            Case [eFrameDigger]
                If (Frame <> [eFrameFalling]) Then
                    .y = .y - 2
                End If
        End Select
        
        '   x offset...
        Select Case .Frame
            Case [eFrameBuilder]
                If (.FrameIdx > 13) Then
                    .x = .x + .xs
                End If
            Case [eFrameBasher]
                .x = .x + .xs * m_aBasherData(.FrameIdx)
            Case [eFrameMiner]
                If (.FrameIdx > 2) Then
                    .x = .x + .xs
                End If
            Case [eFrameDigger]
                .x = .x + .xs
        End Select
        
        '-- Get frame (animation) data...
        .Frame = Frame
        .FrameIdx = FrameIdx
        .FrameSrcY = .Frame * 17
        .FrameOffY = m_uAnimationData(Frame).FrameOffY
        .FrameIdxMax = m_uAnimationData(Frame).FrameIdxMax
        .FrameHasDir = m_uAnimationData(Frame).FrameHasDir
        .f = f
        
        '-- Offset?
        .y = .y + .FrameOffY
    End With
End Sub

'----------------------------------------------------------------------------------------
' Can lem do that job now?
'----------------------------------------------------------------------------------------

Private Function pvCanDoJobNow( _
                 ByVal LemID As Long, _
                 ByVal Job As eLemJobConstants _
                 ) As Boolean
    
  Dim ix As Long
  Dim iy As Long
  Dim px As Long
    
    With m_uLems(LemID)
        
        If (.Frame <> [eFrameFalling]) Then
        
            Select Case Job
            
                Case [eJobBasher]
                    
                    '-- Check ahead
                    ix = -1 - 17 * (.xs > 0)
                    iy = 9
                    px = pvGetPixelLo(.x + ix, .y + iy, LemID)
                    pvCanDoJobNow = ( _
                                    (px = IDX_NONE) Or _
                                    (px = IDX_TERRAIN) Or _
                                    (px = IDX_BASHRIGHT And .xs > 0) Or _
                                    (px = IDX_BASHLEFT And .xs < 0) _
                                    )
                
                Case [eJobMiner]
                    
                    '-- Check feet
                    ix = 11 + 7 * (.xs < 0)
                    iy = 16
                    px = pvGetPixelLo(.x + ix, .y + iy, LemID)
                    pvCanDoJobNow = ( _
                                    (px = IDX_STEEL) Or _
                                    (px = IDX_BASHRIGHT And .xs < 0) Or _
                                    (px = IDX_BASHLEFT And .xs > 0) _
                                    )
                                      
                    '-- Check ahead
                    ix = 15 + 15 * (.xs < 0)
                    iy = 11
                    px = pvGetPixelLo(.x + ix, .y + iy, LemID)
                    pvCanDoJobNow = pvCanDoJobNow Or _
                                    ( _
                                    (px = IDX_STEEL) Or _
                                    (px = IDX_BASHRIGHT And .xs < 0) Or _
                                    (px = IDX_BASHLEFT And .xs > 0) _
                                    )
                                    
                    pvCanDoJobNow = Not pvCanDoJobNow
                    
                Case [eJobDigger]
                    
                    '-- Check feet
                    ix = 7 - (.xs < 0)
                    iy = 16
                    px = pvGetPixelLo(.x + ix, .y + iy, LemID)
                    pvCanDoJobNow = pvCanDoJobNow Or _
                                    ( _
                                    (px = IDX_TERRAIN) Or _
                                    (px = IDX_BASHRIGHT) Or _
                                    (px = IDX_BASHLEFT) _
                                    )
                                    
                    ix = 8 + (.xs < 0)
                    px = pvGetPixelLo(.x + ix, .y + iy, LemID)
                    pvCanDoJobNow = pvCanDoJobNow And Not (px = IDX_STEEL)
                                    
               Case Else
                    
                    '-- OK
                    pvCanDoJobNow = True
            End Select
        End If
    End With
End Function

'----------------------------------------------------------------------------------------
' Get pixel idx at (x,y)
'----------------------------------------------------------------------------------------

Private Function pvGetPixelLo( _
                 ByVal x As Long, _
                 ByVal y As Long, _
                 ByVal LemID As Long _
                 ) As Byte
    
    With m_uScreenBkMaskRect
    
        '-- Valid coordinate?
        If (x >= .x1 And x < .x2 And _
            y >= .y1 And y < .y2 _
            ) Then
                
            '-- Return pixel lo-idx (terrain-related idx)
            pvGetPixelLo = m_aScreenBkMaskBits(x, y) And MSK_LO
          
          Else
            '-- Where are you going?!
            If (y > MAX_YCHECK) Then
                m_uLems(LemID).Active = False
            End If
        End If
    End With
End Function

Private Function pvGetPixelHi( _
                 ByVal x As Long, _
                 ByVal y As Long, _
                 ByVal LemID As Long _
                 ) As Byte
    
    With m_uScreenBkMaskRect
    
        '-- Valid coordinate?
        If (x >= .x1 And x < .x2 And _
            y >= .y1 And y < .y2 _
            ) Then
                
            '-- Return pixel hi-idx (trigger-related idx)
            pvGetPixelHi = m_aScreenBkMaskBits(x, y) And MSK_HI
          
          Else
            '-- Where do you go?!
            If (y > MAX_YCHECK) Then
                m_uLems(LemID).Active = False
            End If
        End If
    End With
End Function

'----------------------------------------------------------------------------------------
' Get (and check) pixel idx at (x,y)
'----------------------------------------------------------------------------------------

Private Function pvCheckPixel( _
                 ByVal x As Long, _
                 ByVal y As Long, _
                 ByVal LemID As Long, _
                 Optional ByVal CheckTrap As Boolean = False _
                 ) As Byte
    
    With m_uScreenBkMaskRect
    
        '-- Valid coordinate?
        If (x >= .x1 And x < .x2 And _
            y >= .y1 And y < .y2 _
            ) Then
        
            '-- Check trigger ID
            Select Case m_aScreenBkMaskBits(x, y) And MSK_HI
            
                Case IDX_EXIT
                    
                    m_uLems(LemID).Job = [eJobNone]
                    Call pvSetLemAnimation(LemID, [eFrameSurviving])
                    pvCheckPixel = IDX_NULL
            
                Case IDX_TRAP
                    
                    If (CheckTrap) Then
                        m_uLems(LemID).Job = [eJobNone]
                        If (m_uLems(LemID).DieNextFrame = False) Then
                            m_uLems(LemID).DieNextFrame = pvFindAndActivateTrap(x, y)
                        End If
                        pvCheckPixel = IDX_NULL
                    End If
            
                Case IDX_LIQUID
                    
                    m_uLems(LemID).Job = [eJobNone]
                    Call pvSetLemAnimation(LemID, [eFrameDrowning])
                    pvCheckPixel = IDX_NULL
                
                Case IDX_FIRE
                
                    m_uLems(LemID).Job = [eJobNone]
                    Call pvSetLemAnimation(LemID, [eFrameBurning])
                    pvCheckPixel = IDX_NULL
                    
                Case Else
                
                    '-- Return pixel terrain ID
                    pvCheckPixel = m_aScreenBkMaskBits(x, y)
            End Select
          
          Else
            '-- Where do you go?!
            If (y > MAX_YCHECK) Then
                m_uLems(LemID).Active = False
            End If
        End If
    End With
End Function

'----------------------------------------------------------------------------------------
' Render mask (onto main buffer and back-mask)
'----------------------------------------------------------------------------------------

Private Sub pvDrawMask( _
            ByVal LemID As Long, _
            Optional ByVal lStep As Long = 0 _
            )
    
    With m_uLems(LemID)
    
        Select Case .Frame
        
            Case [eFrameExploding]
                
                '-- Draw exploding hole mask
               Call MaskBltIdxBkMask( _
                     m_oDIBScreenBkMask, m_oDIBScreenBuffer1, _
                     .x, .y + 2, 16, 22, _
                     IDX_TERRAIN + IDX_BASHLEFT + IDX_BASHRIGHT, IDX_NONE, IDX_NONE, _
                     m_oDIBMask, _
                     0, 0, _
                     IDX_TRANS _
                     )
                     
'                If (pvGetPixelLo(.x + 7, .y + 16, LemID) <> IDX_STEEL And _
'                    pvGetPixelLo(.x + 8, .y + 16, LemID) <> IDX_STEEL _
'                    ) Then
'                    Call MaskBltIdxBkMask( _
'                         m_oDIBScreenBkMask, m_oDIBScreenBuffer1, _
'                         .x, .y + 2, 16, 22, _
'                         IDX_TERRAIN + IDX_BASHLEFT + IDX_BASHRIGHT + IDX_STEEL, IDX_NONE, IDX_NONE, _
'                         m_oDIBMask, _
'                         0, 0, _
'                         IDX_TRANS _
'                         )
'                End If
            
            Case [eFrameBlocker]
          
                '-- Draw blocker mask
                Call MaskBltIdxBkMask( _
                     m_oDIBScreenBkMask, m_oDIBScreenBuffer1, _
                     .x, .y + 2, 16, 16, _
                     IDX_NONE, IDX_BLOCKER, IDX_NONE, _
                     m_oDIBMask, _
                     0, 74, _
                     IDX_TRANS _
                     )
          
            Case [eFrameBuilder]

                '-- Draw brick
                Call MaskBltIdxBkMask( _
                     m_oDIBScreenBkMask, m_oDIBScreenBuffer1, _
                     .x, .y, 16, 16, _
                     IDX_NONE, IDX_TERRAIN, IDX_BRICK, _
                     m_oDIBMask, _
                     -(.xs < 0) * 17, 91, _
                     IDX_TRANS _
                     )
                     
            Case [eFrameBasher]
            
                '-- Draw basher hole mask (sequence)
                Call MaskBltIdxBkMask( _
                     m_oDIBScreenBkMask, m_oDIBScreenBuffer1, _
                     .x, .y + 7 + 2 * lStep, 16, 3, _
                     IDX_TERRAIN + IIf(.xs > 0, IDX_BASHRIGHT, IDX_BASHLEFT), IDX_NONE, IDX_NONE, _
                     m_oDIBMask, _
                     -(.xs < 0) * 17, 27 + 3 * lStep, _
                     IDX_TRANS _
                     )
            
            Case [eFrameMiner]
          
                '-- Draw miner hole mask
                Call MaskBltIdxBkMask( _
                     m_oDIBScreenBkMask, m_oDIBScreenBuffer1, _
                     .x, .y, 16, 16, _
                     IDX_TERRAIN + IIf(.xs > 0, IDX_BASHRIGHT, IDX_BASHLEFT), IDX_NONE, IDX_NONE, _
                     m_oDIBMask, _
                     -(.xs < 0) * 17, 40, _
                     IDX_TRANS _
                     )
                
            Case [eFrameDigger]
                
                '-- Draw digger hole mask
                Call MaskBltIdxBkMask( _
                     m_oDIBScreenBkMask, m_oDIBScreenBuffer1, _
                     .x, .y, 16, 16, _
                     IDX_TERRAIN + IDX_BASHLEFT + IDX_BASHRIGHT, IDX_NONE, IDX_NONE, _
                     m_oDIBMask, _
                     -(.xs < 0) * 17, 57, _
                     IDX_TRANS _
                     )
        End Select
    End With
End Sub

'----------------------------------------------------------------------------------------
' Render mask (special case: blocker un-blocked)
'----------------------------------------------------------------------------------------

Private Sub pvRemoveBlockerMask( _
            ByVal LemID As Long _
            )

    With m_uLems(LemID)
    
        '-- Remove blocker pixels from mask buffer
        Call MaskBltIdxBkMask( _
             m_oDIBScreenBkMask, m_oDIBScreenBuffer1, _
             .x, .y + 2, 16, 16, _
             IDX_BLOCKER, IDX_NONE, IDX_NONE, _
             m_oDIBMask, _
             0, 74, _
             IDX_TRANS _
             )
    End With
End Sub

'----------------------------------------------------------------------------------------
' Paint count-down, explosion, particles
'----------------------------------------------------------------------------------------

Private Sub pvDrawCountDown( _
            ByVal LemID As Long _
            )
    
    With m_uLems(LemID)
        
        '-- Need to paint count-down?
        If (.Frame = [eFrameBurning] Or _
            .Frame = [eFrameDrowning] Or _
            .Frame = [eFrameSpliting] Or _
            .Frame = [eFrameSurviving]) Then
            
            '-- Not needed...
            .Ability = .Ability And Not [eAbilityBomber]
          
          Else
            
            '-- Draw
            Call MaskBlt( _
                 m_oDIBScreenBuffer2, _
                 .x, .y + 6 * (.Frame = [eFrameFloater]), 16, 5, _
                 m_oDIBMask, _
                 (.ExplodeCount \ 15) * 17, 108, _
                 IDX_TRANS _
                 )
        End If
    End With
End Sub

Private Sub pvDrawExplosion( _
            ByVal LemID As Long _
            )

    With m_uLems(LemID)
    
        Call MaskBlt( _
             m_oDIBScreenBuffer2, _
             .x - 5, .y - 6, 26, 32, _
             m_oDIBMask, _
             58, 0, _
             IDX_TRANS _
             )
    End With
End Sub

Private Sub pvDrawParticles( _
            ByVal LemID As Long _
            )
            
  Dim i As Long
  Dim c As Long
  
    With m_uLems(LemID)
    
        For i = 0 To MAX_PARTICLES
            
            With .Particle(i)
                
                '-- No movement if paused
                If (Not IsTimerPaused) Then
                
                    '-- Next position...
                    .x = .x + .vx
                    .y = .y + .vy
                    .vy = .vy + 1
                    
                    '-- Out of screen particles count
                    If (.y > MAX_YCHECK) Then
                        c = c + 1
                    End If
                End If
                
                '-- Render particle
                Call m_oDIBScreenBuffer2.SetPixelIdx(.x, .y, i \ (MAX_PARTICLES \ 3 + 1) + 1)
            End With
        Next i
        
        '-- Still particles
        .Particles = (c <= MAX_PARTICLES)
    End With
End Sub

'----------------------------------------------------------------------------------------
' Set doors state (open/closed: animation enabled/disabled)
'----------------------------------------------------------------------------------------

Private Sub pvSetDoorsState( _
            ByVal bOpening As Boolean _
            )
    
  Dim i As Long
    
    '-- Play sound
    If (bOpening) Then
        Call PlaySoundFX([eFXDoor])
    End If
    
    '-- Open/close all doors (start/stop animation)
    With g_uLevel
        For i = 1 To .Objects
            With .Object(i)
                '-- Door ID?
                If (.ID = DOOR_ID) Then
                    .pvLoop = bOpening
                End If
            End With
        Next i
    End With
End Sub

'----------------------------------------------------------------------------------------
' Find trap lem is just over and activate it
'----------------------------------------------------------------------------------------

Private Function pvFindAndActivateTrap( _
                 ByVal x As Long, _
                 ByVal y As Long _
                 ) As Boolean
  
  Dim i As Long
    
    With g_uLevel
        
        For i = 1 To .Objects
            
            With .Object(i)
                
                '-- Traps' trigger effect?
                If (g_uObjGFX(.ID).TriggerEffect = IDX_TRAP And .pvLoop = False) Then
                    
                    '-- Our coordinate is in trap rectangle?
                    If (x >= .lpRect.x1 And x < .lpRect.x2 And _
                        y >= .lpRect.y1 And y < .lpRect.y2) Then
                        
                        '-- Start animation
                        .pvLoop = True
                        .pvFrameIdxCur = 0
                        
                        '-- Activated
                        pvFindAndActivateTrap = True
                    End If
                End If
            End With
        Next i
    End With
End Function

'----------------------------------------------------------------------------------------
' Hit-test: check if we can apply prepared ability or job just now; also get description
'----------------------------------------------------------------------------------------

Private Function pvHitTest( _
                 ) As String
    
  Dim l      As Long
  Dim lLemID As Long
  Dim lMinx  As Long
  Dim lMiny  As Long
  Dim bSet   As Boolean
    
    '-- Don't check if nothing prepared
    If (m_ePreparedAbility Or m_ePreparedJob) Then
    
        '-- Minimum x/y distances to frame center (16x16 -> ~[8,8] rel.)
        lMinx = 16
        lMiny = 16
        
        '-- Check all lems
        For l = 1 To m_lLemsOut
           
            With m_uLems(l)
                
                '-- Only check active
                If (.Active) Then
                    
                    '-- Over a lem?
                    If (m_xCur >= .x And m_xCur < .x + 16 And _
                        m_yCur >= .y And m_yCur < .y + 16) Then
                        
                        '-- Reset
                        bSet = False
                    
                        '-- Is there a prepared ability? Check if we can apply
                        If (m_ePreparedAbility) Then
                            bSet = ((.Ability And m_ePreparedAbility) = [eAbilityNone])
                        
                        '-- Or, is there a prepared job? Check if we can apply
                        ElseIf (m_ePreparedJob) Then
                            If ((.Job <> m_ePreparedJob) Or (.Frame = [eFrameBuilderEnd])) Then
                                bSet = (.Frame = [eFrameWalker] Or _
                                        .Frame = [eFrameBuilder] Or _
                                        .Frame = [eFrameBuilderEnd] Or _
                                        .Frame = [eFrameBasher] Or _
                                        .Frame = [eFrameMiner] Or _
                                        .Frame = [eFrameDigger] _
                                        )
                            End If
                        End If
                        
                        If (bSet) Then
                            '-- Nearest to center...
                            If (Abs(m_xCur - .x - 8) < lMinx) Then
                                lMinx = Abs(m_xCur - .x - 8)
                                lLemID = l
                            End If
                            If (Abs(m_yCur - .y - 8) < lMiny) Then
                                lMiny = Abs(m_yCur - .y - 8)
                                lLemID = l
                            End If
                        End If
                    End If
                End If
            End With
        Next l
        
        '-- One found
        If (lLemID <> 0) Then
        
            '-- Something has been prepared: update cursor and return lem description
            Set m_oScreen.UserIcon = m_iCursorSelect
            m_lLem = lLemID
            pvHitTest = pvLemDescription(lLemID)
            Exit Function
        End If
    End If
                    
    '-- No lem found: update cursor and 'reset' current lem idx.
    Set m_oScreen.UserIcon = m_iCursorPointer
    m_lLem = -1
End Function

Private Function pvLemDescription( _
                 ByVal LemID As Long _
                 ) As String

    With m_uLems(LemID)
        
        '-- Return lem description depending on current lem frame
        Select Case .Frame
            Case [eFrameFalling]
                pvLemDescription = "Faller"
            Case [eFrameBlocker]
                pvLemDescription = "Blocker"
            Case [eFrameBuilder]
                pvLemDescription = "Builder"
            Case [eFrameDigger]
                pvLemDescription = "Digger"
            Case [eFrameBasher]
                pvLemDescription = "Basher"
            Case [eFrameMiner]
                pvLemDescription = "Miner"
            Case Else
                Select Case .Ability
                    Case [eAbilityClimber] Or [eAbilityFloater]
                        pvLemDescription = "Athlete"
                    Case [eAbilityClimber]
                        pvLemDescription = "Climber"
                    Case [eAbilityFloater]
                        pvLemDescription = "Floater"
                    Case Else
                        pvLemDescription = "Walker"
                End Select
        End Select
        
        '-- Also add lem idx.
        pvLemDescription = pvLemDescription & Space$(1) & m_lLem
    End With
End Function

Private Function pvGetSavedLemsString( _
                 ) As String
    
    If (g_uLevel.LemsToLetOut = 0 Or _
        g_uLevel.LemsToBeSaved = 0 _
        ) Then
        pvGetSavedLemsString = "-"
      
      Else
        If (m_lLemsSaved >= g_uLevel.LemsToBeSaved) Then
            pvGetSavedLemsString = "Done!"
          
          Else
            Select Case m_eLemsSavedMode
                Case [eSavedModePercentage]
                    pvGetSavedLemsString = Format$(m_lLemsSaved / g_uLevel.LemsToLetOut, "0%")
                Case [eSavedModePercentageRemaining]
                    pvGetSavedLemsString = Format$((g_uLevel.LemsToBeSaved - m_lLemsSaved) / g_uLevel.LemsToLetOut, "0%")
                Case [eSavedModeCount]
                    pvGetSavedLemsString = Format$(m_lLemsSaved, "0")
                Case [eSavedModeCountRemaining]
                    pvGetSavedLemsString = Format$(g_uLevel.LemsToBeSaved - m_lLemsSaved, "0")
            End Select
            pvGetSavedLemsString = "Saved " & pvGetSavedLemsString
        End If
    End If
End Function

Private Function pvGetSavedLemsPercentage( _
                 ByVal LemsRef As Integer _
                 ) As String
    
    If (g_uLevel.LemsToLetOut = 0) Then
        pvGetSavedLemsPercentage = "0%"
      Else
        pvGetSavedLemsPercentage = Format$(LemsRef / g_uLevel.LemsToLetOut, "0%")
    End If
End Function

'----------------------------------------------------------------------------------------
' Extra info
'----------------------------------------------------------------------------------------
 
Private Sub pvShowExtraInfo()

    If (m_bLemsInfo) Then
        m_tHold = m_tHold + 1
        m_xInfo = m_xInfo - -(m_tHold > 100) * -(Not IsTimerPaused)
        If (m_xInfo < -230) Then m_xInfo = 230
        Call mLemsRenderer.FXText( _
             m_oScreen.DIB, 90 + m_xInfo, 1, _
             m_sLemsInfo, IDX_RED - IDX_BLUE _
             )
    End If
End Sub

'----------------------------------------------------------------------------------------
' Panoramic view
'----------------------------------------------------------------------------------------

Private Sub pvDrawPanoramicView()
  
  Dim l As Long
    
    '-- Stretch our current buffer, and normalize color
    Call FXPanoramicView( _
         m_oPanoramicView.DIB, m_oDIBScreenBuffer2, _
         m_uPanoramicViewScanX1, m_uPanoramicViewScanX2 _
         )
    
    '-- Draw lems
    If (m_eGameStage = [eStagePlaying]) Then
        For l = 1 To m_lLemsOut
            With m_uLems(l)
                If (.Active) Then
                    Call BltFast( _
                         m_oPanoramicView.DIB, _
                         (.x + 7) / 2.5, (.y + 15) \ 2 - 1, _
                         2, 2, _
                         m_oDIBLemPoint, _
                         0, 0 _
                         )
                End If
            End With
        Next l
    End If
    
    '-- Refresh view
    Call m_oPanoramicView.Refresh
End Sub

'----------------------------------------------------------------------------------------
' Mapping and unmapping DIBs
'----------------------------------------------------------------------------------------

Private Sub pvMapDIB( _
            uSA As SAFEARRAY2D, _
            aBits() As Byte, _
            oDIB As cDIB08 _
            )

    With uSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = oDIB.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = oDIB.BytesPerScanline
        .pvData = oDIB.lpBits
    End With
    Call CopyMemory(ByVal VarPtrArray(aBits()), VarPtr(uSA), 4)
End Sub

Private Sub pvUnmapDIB( _
            aBits() As Byte _
            )

    Call CopyMemory(ByVal VarPtrArray(aBits()), 0&, 4)
End Sub
