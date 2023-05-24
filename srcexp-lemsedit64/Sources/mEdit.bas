Attribute VB_Name = "mEdit"
Option Explicit

'-- A little bit of API:

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const PS_SOLID As Long = 0

Private Declare PtrSafe Function ScreenToClient Lib "user32" (ByVal hWnd As LongPtr, lpPoint As Any) As Long  ' lpPoint As POINT) As Long
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare PtrSafe Function CreatePen Lib "gdi32" (ByVal iStyle As Long, ByVal cWidth As Long, ByVal crColor As Long) As LongPtr
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hDC As LongPtr, ByVal x As Long, ByVal y As Long) As Long

Private Declare PtrSafe Function MoveToEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
Private Declare PtrSafe Function LineTo Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long


'-- LemsEdit:

Public Enum eSelectionConstants
    [eObject] = 0
    [eTerrain] = 1
    [eSteel] = 2
End Enum

Private Const MAX_SCROLL       As Long = 1280
Private Const MAX_POS          As Long = 1600
Private Const MAX_OFFSET       As Long = 1
Private Const VIEW_WIDTH       As Long = 320
Private Const VIEW_HEIGHT      As Long = 160

Private m_oScreen              As ucScreen32
Private m_hDC                  As LongPtr
Private m_x                    As Integer
Private m_y                    As Integer
Private m_x0                   As Integer
Private m_y0                   As Integer
Private m_xScreen              As Integer

Private m_lScreenBackcolor     As Long
Private m_lScreenBackcolorRev  As Long
Private m_bShowObjects         As Boolean
Private m_bShowTerrain         As Boolean
Private m_bShowSteel           As Boolean
Private m_bShowTriggerAreas    As Boolean
Private m_bHighlightSelected   As Boolean
Private m_bShowBlackPieces     As Boolean
Private m_bShowSelectionBox    As Boolean

Private m_eSelectionPreference As eSelectionConstants
Private m_eSelectionType       As eSelectionConstants
Private m_nSelectionIdx        As Integer
Private m_uSelectionRect       As RECTI



'========================================================================================
' Main initialization
'========================================================================================

Public Sub Initialize()
    
    '-- Short references...
    Set m_oScreen = fEdit.ucScreen
    m_hDC = fEdit.ucScreen.DIB.hDC
    
    '-- Set default values
    m_lScreenBackcolor = &H10000
    m_lScreenBackcolorRev = &H1
    m_bShowSelectionBox = True
    m_bHighlightSelected = True
    m_bShowObjects = True
    m_bShowTerrain = True
    m_bShowSteel = True
    m_bShowTriggerAreas = False
    m_bShowBlackPieces = False
    m_eSelectionPreference = [eObject]
    
    '-- Reset
    uLEVEL.GraphicSet = &HFF
End Sub

'========================================================================================
' Info initialization
'========================================================================================

Public Sub InitializeInfo()
    
    '-- Initialize and show info
    Call pvUpdateLevelInfo
    Call pvUpdateStatistics
    Call pvResetSelection
    Call pvUpdateSelectionInfo
End Sub

'========================================================================================
' Render screen
'========================================================================================

Public Sub DoFrame()
    
    '-- Reset (clear) buffer
    Call m_oScreen.DIB.Cls(m_lScreenBackcolor)
    
    '-- Render terrain pieces / objects / steel areas
    If (m_bShowTerrain) Then
        Call pvDrawTerrain
    End If
    If (m_bShowObjects) Then
        Call pvDrawObjects
    End If
    If (m_bShowSteel) Then
        Call pvDrawSteel
    End If
    If (m_bShowTriggerAreas) Then
        Call pvDrawTriggerAreas
    End If
    
    '-- Draw selection box
    If (m_bShowSelectionBox) Then
        Call pvDrawSelectionBox
    End If
    
    '-- Refresh from buffer
    Call m_oScreen.Refresh
    
    '-- Update info
    Call pvUpdateSelectionInfo
End Sub
 
'========================================================================================
' Scroll screen
'========================================================================================

Public Sub DoScrollTo( _
           ByVal x As Long _
           )
    
  Dim lxScreenPrev As Long
    
    lxScreenPrev = m_xScreen
    m_xScreen = x
    
    If (m_xScreen < 0) Then
        m_xScreen = 0
    ElseIf (m_xScreen > MAX_SCROLL) Then
        m_xScreen = MAX_SCROLL
    End If
    
    If (GetAsyncKeyState(vbKeySpace) < 0) Then
        Call SelectionMove(m_xScreen - lxScreenPrev, 0)
      Else
        Call DoFrame
    End If
End Sub

'========================================================================================
' Adding objects / terrain pieces / steel areas
'========================================================================================

Public Function AddObject( _
       ByVal ID As Integer _
       ) As Boolean
    
    If (uLEVEL.Objects < MAX_OBJECTS) Then
    
        With uLEVEL
        
            .Objects = .Objects + 1
            ReDim Preserve .Object(0 To .Objects)
            
            With .Object(.Objects)
                .ID = ID
                With .lpRect
                    .x1 = m_xScreen + (VIEW_WIDTH - uOBJGFX(ID).Width) \ 2
                    .y1 = (VIEW_HEIGHT - uOBJGFX(ID).Height) \ 2
                    .x2 = .x1 + uOBJGFX(ID).Width
                    .y2 = .y1 + uOBJGFX(ID).Height
                End With
                Call pvCopyRect(m_uSelectionRect, .lpRect)
                m_eSelectionType = [eObject]
            End With
            m_nSelectionIdx = .Objects
        End With
        AddObject = True
        
        Call pvUpdateStatistics
        Call pvUpdateSelectionFlags
        Call DoFrame
    End If
End Function

Public Function AddTerrainPiece( _
       ByVal ID As Integer _
       ) As Boolean
    
    If (uLEVEL.TerrainPieces < MAX_TERRAINPIECES) Then
    
        With uLEVEL
        
            .TerrainPieces = .TerrainPieces + 1
            ReDim Preserve .TerrainPiece(0 To .TerrainPieces)
            
            With .TerrainPiece(.TerrainPieces)
                .ID = ID
                With .lpRect
                    .x1 = m_xScreen + (VIEW_WIDTH - uTERGFX(ID).Width) \ 2
                    .y1 = (VIEW_HEIGHT - uTERGFX(ID).Height) \ 2
                    .x2 = .x1 + uTERGFX(ID).Width
                    .y2 = .y1 + uTERGFX(ID).Height
                End With
                Call pvCopyRect(m_uSelectionRect, .lpRect)
                m_eSelectionType = [eTerrain]
            End With
            m_nSelectionIdx = .TerrainPieces
        End With
        AddTerrainPiece = True
        
        Call pvUpdateStatistics
        Call pvUpdateSelectionFlags
        Call DoFrame
    End If
End Function

Public Function AddSteelArea( _
       ByVal Width As Integer, _
       ByVal Height As Integer _
       ) As Boolean

    If (uLEVEL.SteelAreas < MAX_STEELAREAS) Then
    
        If (Width >= MIN_STEELAREASIZE And Height >= MIN_STEELAREASIZE) Then
            
            With uLEVEL
                
                .SteelAreas = .SteelAreas + 1
                ReDim Preserve .SteelArea(0 To .SteelAreas)
            
                With .SteelArea(.SteelAreas)
                    With .lpRect
                        .x1 = m_xScreen + (VIEW_WIDTH - Width) \ 2
                        .y1 = (VIEW_HEIGHT - Height) \ 2
                        .x2 = .x1 + Width
                        .y2 = .y1 + Height
                    End With
                    Call pvCopyRect(m_uSelectionRect, .lpRect)
                    m_eSelectionType = [eSteel]
                End With
                m_nSelectionIdx = .SteelAreas
            End With
            AddSteelArea = True
        
            Call pvUpdateStatistics
            Call pvUpdateSelectionFlags
            Call DoFrame
        End If
    End If
End Function

Public Sub UpdateStatistics()
    
    Call pvUpdateStatistics
End Sub

'========================================================================================
' Steel area size
'========================================================================================

Public Sub SteelAreaSetWidth( _
           ByVal Width As Integer _
           )

    If (SelectionExists And Width >= MIN_STEELAREASIZE) Then
        With uLEVEL
            With .SteelArea(m_nSelectionIdx)
                With .lpRect
                    .x2 = .x1 + Width
                End With
                Call pvCopyRect(m_uSelectionRect, .lpRect)
            End With
        End With
    
        Call DoFrame
    End If
End Sub

Public Sub SteelAreaSetHeight( _
           ByVal Height As Integer _
           )

    If (SelectionExists And Height >= MIN_STEELAREASIZE) Then
        With uLEVEL
            With .SteelArea(m_nSelectionIdx)
                With .lpRect
                    .y2 = .y1 + Height
                End With
                Call pvCopyRect(m_uSelectionRect, .lpRect)
            End With
        End With
    
        Call DoFrame
    End If
End Sub

'========================================================================================
' Dragging selection
'========================================================================================

Public Sub MouseDown( _
           Button As Integer, _
           Shift As Integer, _
           x As Long, _
           y As Long _
           )
    
    '-- Store current position
    m_x = x
    m_y = y
    
    '-- Hit-test...
    If (pvHitTest(Button, x, y)) Then
        
        '-- Show context menu?
        If (Button = vbRightButton) Then
            Call fEdit.PopUpMenu(fEdit.mnuContextSelectionTop)
        End If
    End If
End Sub

Public Sub MouseMove( _
           Button As Integer, _
           Shift As Integer, _
           x As Long, _
           y As Long _
           )
    
    If (Button = vbLeftButton) Then
        
        '-- Clip position
        If (x < 0) Then x = 0 Else If (x >= VIEW_WIDTH) Then x = VIEW_WIDTH
        If (y < 0) Then y = 0 Else If (y >= VIEW_HEIGHT) Then y = VIEW_HEIGHT
        
        '-- Scroll screen / move selection
        If (GetAsyncKeyState(vbKeyShift) < 0) Then
            fEdit.ucScroll.Value = fEdit.ucScroll.Value - (x - m_x)
          Else
            Call SelectionMove(x - m_x, y - m_y)
        End If
        
        '-- Store current position
        m_x = x
        m_y = y
        
      Else
        If (GetAsyncKeyState(vbKeyControl) < 0) Then
            Call pvHitTest(0, x, y)
        End If
    End If
End Sub

'========================================================================================
' Manipulating selection
'========================================================================================

Public Sub RestoreSelectionPosition()
    
    If (SelectionExists) Then
        Call SelectionMove( _
             m_x0 - m_uSelectionRect.x1, _
             m_y0 - m_uSelectionRect.y1, _
             redraw:=False _
             )
        Call ResetSelection
        Call DoFrame
    End If
End Sub

Public Sub ResetSelection()
    
    Call pvResetSelection
    Call pvUpdateSelectionInfo
End Sub

Public Function SelectionExists( _
                ) As Boolean

    SelectionExists = (m_nSelectionIdx > 0)
End Function

Public Sub SelectionMove( _
           ByVal dx As Long, _
           ByVal dy As Long, _
           Optional ByVal redraw As Boolean = True _
           )
    
    If (SelectionExists) Then
        
        Select Case m_eSelectionType
            
            Case [eObject]
                With uLEVEL.Object(m_nSelectionIdx)
                    Call pvMoveRect(.lpRect, dx, dy)
                    Call pvCopyRect(m_uSelectionRect, .lpRect)
                End With
            
            Case [eTerrain]
                With uLEVEL.TerrainPiece(m_nSelectionIdx)
                    Call pvMoveRect(.lpRect, dx, dy)
                    Call pvCopyRect(m_uSelectionRect, .lpRect)
                End With
            
            Case [eSteel]
                With uLEVEL.SteelArea(m_nSelectionIdx)
                    Call pvMoveRect(.lpRect, dx, dy)
                    Call pvCopyRect(m_uSelectionRect, .lpRect)
                End With
        End Select
        
        If (redraw) Then Call DoFrame
    End If
End Sub

Public Function SelectionDuplicate( _
                ) As Boolean
    
    If (SelectionExists) Then
    
        With uLEVEL
        
            Select Case m_eSelectionType
                
                Case [eObject]
                    
                    If (.Objects < MAX_OBJECTS) Then
                        
                        .Objects = .Objects + 1
                        ReDim Preserve .Object(0 To .Objects)
                        
                        .Object(.Objects) = .Object(m_nSelectionIdx)
                        m_nSelectionIdx = .Objects
                        
                        With .Object(.Objects)
                            Call pvMoveRect(.lpRect, 2, 2)
                            Call pvCopyRect(m_uSelectionRect, .lpRect)
                        End With
                        
                        SelectionDuplicate = True
                    End If
                        
                Case [eTerrain]
                
                    If (.TerrainPieces < MAX_TERRAINPIECES) Then
                        
                        .TerrainPieces = .TerrainPieces + 1
                        ReDim Preserve .TerrainPiece(0 To .TerrainPieces)
                        
                        .TerrainPiece(.TerrainPieces) = .TerrainPiece(m_nSelectionIdx)
                        m_nSelectionIdx = .TerrainPieces
                        
                        With .TerrainPiece(.TerrainPieces)
                            Call pvMoveRect(.lpRect, 2, 2)
                            Call pvCopyRect(m_uSelectionRect, .lpRect)
                        End With
                        
                        SelectionDuplicate = True
                    End If
                    
                Case [eSteel]
            
                    If (.SteelAreas < MAX_STEELAREAS) Then
                    
                        .SteelAreas = .SteelAreas + 1
                        ReDim Preserve .SteelArea(0 To .SteelAreas)
                        
                        .SteelArea(.SteelAreas) = .SteelArea(m_nSelectionIdx)
                        m_nSelectionIdx = .SteelAreas
                        
                        With .SteelArea(.SteelAreas)
                            Call pvMoveRect(.lpRect, 2, 2)
                            Call pvCopyRect(m_uSelectionRect, .lpRect)
                        End With
                        
                        SelectionDuplicate = True
                    End If
            End Select
        End With
        
        If (SelectionDuplicate) Then
            Call pvUpdateSelectionInfo
            Call pvUpdateStatistics
            Call DoFrame
        End If
        
      Else
        SelectionDuplicate = True
    End If
End Function

Public Sub SelectionRemove()
    
  Dim i As Long
  
    If (SelectionExists) Then
    
        With uLEVEL
        
            Select Case m_eSelectionType
                
                Case [eObject]
                    
                    For i = m_nSelectionIdx To .Objects - 1
                        .Object(i) = .Object(i + 1)
                    Next i
                    .Objects = .Objects - 1
                    ReDim Preserve .Object(0 To .Objects)
                        
                Case [eTerrain]
                
                    For i = m_nSelectionIdx To .TerrainPieces - 1
                        .TerrainPiece(i) = .TerrainPiece(i + 1)
                    Next i
                    .TerrainPieces = .TerrainPieces - 1
                    ReDim Preserve .TerrainPiece(0 To .TerrainPieces)
                
                Case [eSteel]
            
                    For i = m_nSelectionIdx To .SteelAreas - 1
                        .SteelArea(i) = .SteelArea(i + 1)
                    Next i
                    .SteelAreas = .SteelAreas - 1
                    ReDim Preserve .SteelArea(0 To .SteelAreas)
            End Select
        End With
        
        Call pvResetSelection
        Call pvUpdateStatistics
        Call DoFrame
    End If
End Sub

Public Sub SelectionBringToTop()
    
  Dim uTmpObject       As tObject
  Dim uTmpTerrainPiece As tTerrainPiece
  Dim uTmpSteelArea    As tSteelArea
  Dim i                As Long
  
    If (SelectionExists) Then
    
        With uLEVEL
        
            Select Case m_eSelectionType
                
                Case [eObject]
                    
                    uTmpObject = .Object(m_nSelectionIdx)
                    For i = m_nSelectionIdx To .Objects - 1
                        .Object(i) = .Object(i + 1)
                    Next i
                    .Object(.Objects) = uTmpObject
                    m_nSelectionIdx = .Objects
    
                Case [eTerrain]
                
                    uTmpTerrainPiece = .TerrainPiece(m_nSelectionIdx)
                    For i = m_nSelectionIdx To .TerrainPieces - 1
                        .TerrainPiece(i) = .TerrainPiece(i + 1)
                    Next i
                    .TerrainPiece(.TerrainPieces) = uTmpTerrainPiece
                    m_nSelectionIdx = .TerrainPieces
                
                Case [eSteel]
            
                    uTmpSteelArea = .SteelArea(m_nSelectionIdx)
                    For i = m_nSelectionIdx To .SteelAreas - 1
                        .SteelArea(i) = .SteelArea(i + 1)
                    Next i
                    .SteelArea(.SteelAreas) = uTmpSteelArea
                    m_nSelectionIdx = .SteelAreas
            End Select
        End With
        
        Call pvUpdateSelectionInfo
        Call DoFrame
    End If
End Sub

Public Sub SelectionBringToBottom()
    
  Dim uTmpObject       As tObject
  Dim uTmpTerrainPiece As tTerrainPiece
  Dim uTmpSteelArea    As tSteelArea
  Dim i                As Long
  
    If (SelectionExists) Then
    
        With uLEVEL
        
            Select Case m_eSelectionType
                
                Case [eObject]
                    
                    uTmpObject = .Object(m_nSelectionIdx)
                    For i = m_nSelectionIdx To 2 Step -1
                        .Object(i) = .Object(i - 1)
                    Next i
                    .Object(1) = uTmpObject
                    m_nSelectionIdx = 1
                    
                Case [eTerrain]
                
                    uTmpTerrainPiece = .TerrainPiece(m_nSelectionIdx)
                    For i = m_nSelectionIdx To 2 Step -1
                        .TerrainPiece(i) = .TerrainPiece(i - 1)
                    Next i
                    .TerrainPiece(1) = uTmpTerrainPiece
                    m_nSelectionIdx = 1
                
                Case [eSteel]
            
                    uTmpSteelArea = .SteelArea(m_nSelectionIdx)
                    For i = m_nSelectionIdx To 2 Step -1
                        .SteelArea(i) = .SteelArea(i - 1)
                    Next i
                    .SteelArea(1) = uTmpSteelArea
                    m_nSelectionIdx = 1
            End Select
        End With
        
        Call pvUpdateSelectionInfo
        Call DoFrame
    End If
End Sub

Public Sub SelectionZOrderUp()
    
  Dim uTmpObject       As tObject
  Dim uTmpTerrainPiece As tTerrainPiece
  Dim uTmpSteelArea    As tSteelArea
  
    If (SelectionExists) Then
        
        With uLEVEL
        
            Select Case m_eSelectionType
                
                Case [eObject]
                    
                    If (m_nSelectionIdx < .Objects) Then
                        uTmpObject = .Object(m_nSelectionIdx)
                        .Object(m_nSelectionIdx) = .Object(m_nSelectionIdx + 1)
                        .Object(m_nSelectionIdx + 1) = uTmpObject
                        m_nSelectionIdx = m_nSelectionIdx + 1
                    End If
                    
                Case [eTerrain]
    
                    If (m_nSelectionIdx < .TerrainPieces) Then
                        uTmpTerrainPiece = .TerrainPiece(m_nSelectionIdx)
                        .TerrainPiece(m_nSelectionIdx) = .TerrainPiece(m_nSelectionIdx + 1)
                        .TerrainPiece(m_nSelectionIdx + 1) = uTmpTerrainPiece
                        m_nSelectionIdx = m_nSelectionIdx + 1
                    End If
                
                Case [eSteel]
            
                    If (m_nSelectionIdx < .SteelAreas) Then
                        uTmpSteelArea = .SteelArea(m_nSelectionIdx)
                        .SteelArea(m_nSelectionIdx) = .SteelArea(m_nSelectionIdx + 1)
                        .SteelArea(m_nSelectionIdx + 1) = uTmpSteelArea
                        m_nSelectionIdx = m_nSelectionIdx + 1
                    End If
            End Select
        End With
        
        Call pvUpdateSelectionInfo
        Call DoFrame
    End If
End Sub

Public Sub SelectionZOrderDown()
    
  Dim uTmpObject       As tObject
  Dim uTmpTerrainPiece As tTerrainPiece
  Dim uTmpSteelArea    As tSteelArea
   
    If (SelectionExists) Then

        With uLEVEL
        
            Select Case m_eSelectionType
                
                Case [eObject]
                    
                    If (m_nSelectionIdx > 1) Then
                        uTmpObject = .Object(m_nSelectionIdx)
                        .Object(m_nSelectionIdx) = .Object(m_nSelectionIdx - 1)
                        .Object(m_nSelectionIdx - 1) = uTmpObject
                        m_nSelectionIdx = m_nSelectionIdx - 1
                    End If
                    
                Case [eTerrain]
                
                    If (m_nSelectionIdx > 1) Then
                        uTmpTerrainPiece = .TerrainPiece(m_nSelectionIdx)
                        .TerrainPiece(m_nSelectionIdx) = .TerrainPiece(m_nSelectionIdx - 1)
                        .TerrainPiece(m_nSelectionIdx - 1) = uTmpTerrainPiece
                        m_nSelectionIdx = m_nSelectionIdx - 1
                    End If
                    
                Case [eSteel]
            
                    If (m_nSelectionIdx > 1) Then
                        uTmpSteelArea = .SteelArea(m_nSelectionIdx)
                        .SteelArea(m_nSelectionIdx) = .SteelArea(m_nSelectionIdx - 1)
                        .SteelArea(m_nSelectionIdx - 1) = uTmpSteelArea
                        m_nSelectionIdx = m_nSelectionIdx - 1
                    End If
            End Select
        End With
        
        Call pvUpdateSelectionInfo
        Call DoFrame
    End If
End Sub

Public Sub SelectionFindNextOver()
  
  Dim uPt As POINTAPI
  Dim i   As Long
  
    Call GetCursorPos(uPt)
    Call ScreenToClient(fEdit.ucScreen.hWnd, uPt)
    
    uPt.x = uPt.x \ 2 + m_xScreen
    uPt.y = uPt.y \ 2
    
    If (SelectionExists) Then

        With uLEVEL
        
            Select Case m_eSelectionType
                
                Case [eObject]
                    
                    For i = m_nSelectionIdx + 1 To .Objects
                        With .Object(i)
                            If (pvPtInRect(.lpRect, uPt.x, uPt.y)) Then
                                Call pvCopyRect(m_uSelectionRect, .lpRect)
                                m_nSelectionIdx = i
                                fEdit.cbObject.ListIndex = .ID
                                Exit For
                            End If
                        End With
                    Next i
                    
                Case [eTerrain]
                
                    For i = m_nSelectionIdx + 1 To .TerrainPieces
                        With .TerrainPiece(i)
                            If (pvPtInRect(.lpRect, uPt.x, uPt.y)) Then
                                Call pvCopyRect(m_uSelectionRect, .lpRect)
                                m_nSelectionIdx = i
                                fEdit.cbTerrainPiece.ListIndex = .ID
                                Exit For
                            End If
                        End With
                    Next i
                    
                Case [eSteel]
            
                    For i = m_nSelectionIdx + 1 To .SteelAreas
                        With .SteelArea(i)
                            If (pvPtInRect(.lpRect, uPt.x, uPt.y)) Then
                                Call pvCopyRect(m_uSelectionRect, .lpRect)
                                m_nSelectionIdx = i
                                Exit For
                            End If
                        End With
                    Next i
            End Select
        End With
        
        Call pvUpdateSelectionInfo
        Call pvUpdateSelectionFlags
    Call DoFrame
    End If
End Sub

Public Sub SelectionFindNextUnder()

  Dim uPt As POINTAPI
  Dim i   As Long
  
    Call GetCursorPos(uPt)
    Call ScreenToClient(fEdit.ucScreen.hWnd, uPt)
    
    uPt.x = uPt.x \ 2 + m_xScreen
    uPt.y = uPt.y \ 2
    
    If (SelectionExists) Then

        With uLEVEL
        
            Select Case m_eSelectionType
                
                Case [eObject]
                    
                    For i = m_nSelectionIdx - 1 To 1 Step -1
                        With .Object(i)
                            If (pvPtInRect(.lpRect, uPt.x, uPt.y)) Then
                                Call pvCopyRect(m_uSelectionRect, .lpRect)
                                fEdit.cbObject.ListIndex = .ID
                                m_nSelectionIdx = i
                                Exit For
                            End If
                        End With
                    Next i
                    
                Case [eTerrain]
                
                    For i = m_nSelectionIdx - 1 To 1 Step -1
                        With .TerrainPiece(i)
                            If (pvPtInRect(.lpRect, uPt.x, uPt.y)) Then
                                Call pvCopyRect(m_uSelectionRect, .lpRect)
                                m_nSelectionIdx = i
                                fEdit.cbTerrainPiece.ListIndex = .ID
                                Exit For
                            End If
                        End With
                    Next i
                    
                Case [eSteel]
            
                    For i = m_nSelectionIdx - 1 To 1 Step -1
                        With .SteelArea(i)
                            If (pvPtInRect(.lpRect, uPt.x, uPt.y)) Then
                                Call pvCopyRect(m_uSelectionRect, .lpRect)
                                m_nSelectionIdx = i
                                Exit For
                            End If
                        End With
                    Next i
            End Select
        End With
        
        Call pvUpdateSelectionInfo
        Call pvUpdateSelectionFlags
        Call DoFrame
    End If
End Sub

'========================================================================================
' Setting property values (preferences)
'========================================================================================

Public Property Let ShowSelectionBox(ByVal New_ShowSelectionBox As Boolean)
    
    m_bShowSelectionBox = New_ShowSelectionBox
    Call DoFrame
End Property

Public Property Let HighlightSelected(ByVal New_HighlightSelected As Boolean)
    
    m_bHighlightSelected = New_HighlightSelected
    Call DoFrame
End Property

Public Property Let ShowObjects(ByVal New_ShowObjects As Boolean)
    
    m_bShowObjects = New_ShowObjects
    Call pvResetSelection
    Call DoFrame
End Property

Public Property Let ShowTerrain(ByVal New_ShowTerrain As Boolean)
    
    m_bShowTerrain = New_ShowTerrain
    Call pvResetSelection
    Call DoFrame
End Property

Public Property Let ShowSteel(ByVal New_ShowSteel As Boolean)
    
    m_bShowSteel = New_ShowSteel
    Call DoFrame
End Property

Public Property Let ShowTriggerAreas(ByVal New_ShowTriggerAreas As Boolean)
    
    m_bShowTriggerAreas = New_ShowTriggerAreas
    Call DoFrame
End Property

Public Property Let ShowBlackPieces(ByVal New_ShowBlackPieces As Boolean)
    
    m_bShowBlackPieces = New_ShowBlackPieces
    Call DoFrame
End Property

Public Property Let SelectionPreference(ByVal New_SelectionPreference As eSelectionConstants)
    
    m_eSelectionPreference = New_SelectionPreference
    Call pvResetSelection
    Call DoFrame
End Property

'========================================================================================
' Special flags
'========================================================================================

Public Sub SetNotOverlap( _
           ByVal New_NotOverlap As Boolean _
           )
    
    If (SelectionExists) Then
        Select Case m_eSelectionType
            Case [eObject]
                uLEVEL.Object(m_nSelectionIdx).NotOverlap = New_NotOverlap
                Call pvUpdateSelectionFlags
                Call DoFrame
            Case [eTerrain]
                uLEVEL.TerrainPiece(m_nSelectionIdx).NotOverlap = New_NotOverlap
                Call pvUpdateSelectionFlags
                Call DoFrame
        End Select
    End If
End Sub

Public Sub ObjectSetOnTerrain( _
           ByVal New_OnTerrain As Boolean _
           )
    
    If (SelectionExists) Then
        uLEVEL.Object(m_nSelectionIdx).OnTerrain = New_OnTerrain
        Call pvUpdateSelectionFlags
        Call DoFrame
    End If
End Sub

Public Sub TerrainPieceSetBlack( _
           ByVal New_Black As Boolean _
           )
    
    If (SelectionExists) Then
        uLEVEL.TerrainPiece(m_nSelectionIdx).Black = New_Black
        Call pvUpdateSelectionFlags
        Call DoFrame
    End If
End Sub

Public Sub TerrainPieceSetUpsideDown( _
           ByVal New_UpsideDown As Boolean _
           )
    
    If (SelectionExists) Then
        uLEVEL.TerrainPiece(m_nSelectionIdx).UpsideDown = New_UpsideDown
        Call pvUpdateSelectionFlags
        Call DoFrame
    End If
End Sub

'========================================================================================
' Private
'========================================================================================

Private Function pvHitTest( _
                 ByVal Button As Integer, _
                 ByVal x As Long, _
                 ByVal y As Long _
                 ) As Boolean
    
    x = x + m_xScreen
    
    If (Button = vbRightButton And pvPtInRect(m_uSelectionRect, x, y)) Then
    
        pvHitTest = True
        
      Else
        
        m_nSelectionIdx = 0
        Call pvSetRectEmpty(m_uSelectionRect)
        
        Select Case m_eSelectionPreference
            
            Case [eObject]
                
                pvHitTest = pvHitTestObject(x, y)
                If (pvHitTest = False) Then
                    pvHitTest = pvHitTestTerrainPiece(x, y)
                    If (pvHitTest = False) Then
                        pvHitTest = pvHitTestSteelArea(x, y)
                    End If
                End If
                
            Case [eTerrain]
                
                pvHitTest = pvHitTestTerrainPiece(x, y)
                If (pvHitTest = False) Then
                    pvHitTest = pvHitTestObject(x, y)
                    If (pvHitTest = False) Then
                        pvHitTest = pvHitTestSteelArea(x, y)
                    End If
                End If
            
            Case [eSteel]
            
                pvHitTest = pvHitTestSteelArea(x, y)
                If (pvHitTest = False) Then
                    pvHitTest = pvHitTestObject(x, y)
                    If (pvHitTest = False) Then
                        pvHitTest = pvHitTestTerrainPiece(x, y)
                    End If
                End If
        End Select
        
        m_x0 = m_uSelectionRect.x1
        m_y0 = m_uSelectionRect.y1
        
        Call pvUpdateSelectionFlags
        Call DoFrame
    End If
End Function

Private Function pvHitTestObject( _
                 ByVal x As Long, _
                 ByVal y As Long _
                 ) As Boolean
  
  Dim i As Long
                
    If (m_bShowObjects) Then
        For i = uLEVEL.Objects To 1 Step -1
            With uLEVEL.Object(i)
                If (pvPtInRect(.lpRect, x, y)) Then
'                    If (GetPixel(uOBJGFX(.ID).DIB.hDC, x - .lpRect.x1, y - .lpRect.y1) <> CLR_TRANS) Then
                        Call pvCopyRect(m_uSelectionRect, .lpRect)
                        m_nSelectionIdx = i
                        m_eSelectionType = [eObject]
                        fEdit.cbObject.ListIndex = .ID
                        pvHitTestObject = True
                        Exit For
'                    End If
                End If
            End With
        Next i
    End If
End Function

Private Function pvHitTestTerrainPiece( _
                 ByVal x As Long, _
                 ByVal y As Long _
                 ) As Boolean
  
  Dim i  As Long
  Dim px As Long

    If (m_bShowTerrain) Then
        For i = uLEVEL.TerrainPieces To 1 Step -1
            With uLEVEL.TerrainPiece(i)
                If (pvPtInRect(.lpRect, x, y)) Then
                    If (.UpsideDown) Then
                        px = GetPixel(uTERGFX(.ID).DIB.hDC, x - .lpRect.x1, uTERGFX(.ID).Height - (y - .lpRect.y1) - 1)
                      Else
                        px = GetPixel(uTERGFX(.ID).DIB.hDC, x - .lpRect.x1, y - .lpRect.y1)
                    End If
                    If (px <> CLR_TRANS) Then
                        Call pvCopyRect(m_uSelectionRect, .lpRect)
                        m_nSelectionIdx = i
                        m_eSelectionType = [eTerrain]
                        fEdit.cbTerrainPiece.ListIndex = .ID
                        pvHitTestTerrainPiece = True
                        Exit For
                    End If
                End If
            End With
        Next i
    End If
End Function

Private Function pvHitTestSteelArea( _
                 ByVal x As Long, _
                 ByVal y As Long _
                 ) As Boolean
  
  Dim i As Long

    If (m_bShowSteel) Then
        For i = uLEVEL.SteelAreas To 1 Step -1
            With uLEVEL.SteelArea(i)
                If (pvPtInRect(.lpRect, x, y)) Then
                    Call pvCopyRect(m_uSelectionRect, .lpRect)
                    m_nSelectionIdx = i
                    m_eSelectionType = [eSteel]
                    pvHitTestSteelArea = True
                    Exit For
                End If
            End With
        Next i
    End If
End Function

Private Function pvCanMoveRect( _
                 lpRect As RECTI, _
                 ByVal dx As Long, _
                 ByVal dy As Long _
                 ) As Boolean

    With lpRect
        If ((.x1 > MAX_POS - MAX_OFFSET) Or _
            (.x2 < MAX_OFFSET) Or _
            (.y1 > VIEW_HEIGHT - MAX_OFFSET) Or _
            (.y2 < MAX_OFFSET)) Then
            pvCanMoveRect = False
          Else
            pvCanMoveRect = True
        End If
    End With
End Function

Private Sub pvMoveRect( _
            lpRect As RECTI, _
            ByVal dx As Long, _
            ByVal dy As Long _
            )
    
    Call pvOffsetRect(lpRect, dx, dy)
    
    With lpRect
        If (.x1 > MAX_POS - MAX_OFFSET) Then
            Call pvOffsetRect(lpRect, (MAX_POS - MAX_OFFSET) - .x1, 0)
        End If
        If (.x2 < MAX_OFFSET) Then
            Call pvOffsetRect(lpRect, MAX_OFFSET - .x2, 0)
        End If
        If (.y1 > VIEW_HEIGHT - MAX_OFFSET) Then
            Call pvOffsetRect(lpRect, 0, (VIEW_HEIGHT - MAX_OFFSET) - .y1)
        End If
        If (.y2 < MAX_OFFSET) Then
            Call pvOffsetRect(lpRect, 0, MAX_OFFSET - .y2)
        End If
    End With
End Sub

Private Sub pvCopyRect( _
            lpDestRect As RECTI, _
            lpSourceRect As RECTI _
            )

    With lpSourceRect
        lpDestRect.x1 = .x1
        lpDestRect.y1 = .y1
        lpDestRect.x2 = .x2
        lpDestRect.y2 = .y2
    End With
End Sub

Private Sub pvSetRectEmpty( _
            lpRect As RECTI _
            )
    
    With lpRect
        .x2 = .x1
        .y2 = .y1
    End With
End Sub

Private Sub pvOffsetRect( _
            lpRect As RECTI, _
            ByVal x As Long, _
            ByVal y As Long _
            )

    With lpRect
        .x1 = .x1 + x
        .y1 = .y1 + y
        .x2 = .x2 + x
        .y2 = .y2 + y
    End With
End Sub

Private Function pvPtInRect( _
                 lpRect As RECTI, _
                 ByVal x As Long, _
                 ByVal y As Long _
                 ) As Boolean

    With lpRect
        pvPtInRect = (x >= .x1 And x < .x2) And (y >= .y1 And y < .y2)
    End With
End Function

Private Function pvIsRectEmpty( _
                 lpRect As RECTI _
                 ) As Boolean

    With lpRect
        pvIsRectEmpty = (.x1 = .x2) Or (.y1 = .y2)
    End With
End Function

Private Sub pvDrawTerrain()
 
  Dim i As Long
    
    For i = 1 To uLEVEL.TerrainPieces
        With uLEVEL.TerrainPiece(i)
            If (.lpRect.x1 >= m_xScreen And .lpRect.x1 < m_xScreen + VIEW_WIDTH) Or _
               (.lpRect.x2 >= m_xScreen And .lpRect.x2 < m_xScreen + VIEW_WIDTH) Then
                If (m_bHighlightSelected And _
                   (m_eSelectionType = [eTerrain]) And _
                   (m_nSelectionIdx = i)) Then
                    If (.Black) Then
                        If (.NotOverlap) Then
                            Call mEditRenderer.MaskBltColorOverlap( _
                                 m_oScreen.DIB, _
                                 -m_xScreen + .lpRect.x1, .lpRect.y1, _
                                 uTERGFX(.ID).Width, uTERGFX(.ID).Height, _
                                 m_lScreenBackcolorRev, _
                                 CLR_LIGHTEN, _
                                 uTERGFX(.ID).DIB, _
                                 0, 0, _
                                 CLR_TRANS, _
                                 .UpsideDown _
                                 )
                          Else
                            Call mEditRenderer.MaskBltColor( _
                                 m_oScreen.DIB, _
                                 -m_xScreen + .lpRect.x1, .lpRect.y1, _
                                 uTERGFX(.ID).Width, uTERGFX(.ID).Height, _
                                 CLR_LIGHTEN, _
                                 uTERGFX(.ID).DIB, _
                                 0, 0, _
                                 CLR_TRANS, _
                                 .UpsideDown _
                                 )
                        End If
                      Else
                        If (.NotOverlap) Then
                            Call mEditRenderer.MaskBltLightenOverlap( _
                                 m_oScreen.DIB, _
                                 -m_xScreen + .lpRect.x1, .lpRect.y1, _
                                 uTERGFX(.ID).Width, uTERGFX(.ID).Height, _
                                 m_lScreenBackcolorRev, _
                                 uTERGFX(.ID).DIB, _
                                 0, 0, _
                                 CLR_TRANS, _
                                 .UpsideDown _
                                 )
                          Else
                            Call mEditRenderer.MaskBltLighten( _
                                 m_oScreen.DIB, _
                                 -m_xScreen + .lpRect.x1, .lpRect.y1, _
                                 uTERGFX(.ID).Width, uTERGFX(.ID).Height, _
                                 uTERGFX(.ID).DIB, _
                                 0, 0, _
                                 CLR_TRANS, _
                                 .UpsideDown _
                                 )
                        End If
                    End If
                  Else
                    If (.Black) Then
                        If (m_bShowBlackPieces) Then
                            If (.NotOverlap) Then
                                Call mEditRenderer.MaskBltColorOverlap( _
                                     m_oScreen.DIB, _
                                     -m_xScreen + .lpRect.x1, .lpRect.y1, _
                                     uTERGFX(.ID).Width, uTERGFX(.ID).Height, _
                                     m_lScreenBackcolorRev, _
                                     CLR_RED, _
                                     uTERGFX(.ID).DIB, _
                                     0, 0, _
                                     CLR_TRANS, _
                                     .UpsideDown _
                                     )
                              Else
                                Call mEditRenderer.MaskBltColor( _
                                     m_oScreen.DIB, _
                                     -m_xScreen + .lpRect.x1, .lpRect.y1, _
                                     uTERGFX(.ID).Width, uTERGFX(.ID).Height, _
                                     CLR_RED, _
                                     uTERGFX(.ID).DIB, _
                                     0, 0, _
                                     CLR_TRANS, _
                                     .UpsideDown _
                                     )
                            End If
                          Else
                            If (.NotOverlap) Then
                                Call mEditRenderer.MaskBltColorOverlap( _
                                     m_oScreen.DIB, _
                                     -m_xScreen + .lpRect.x1, .lpRect.y1, _
                                     uTERGFX(.ID).Width, uTERGFX(.ID).Height, _
                                     m_lScreenBackcolorRev, m_lScreenBackcolorRev, _
                                     uTERGFX(.ID).DIB, _
                                     0, 0, _
                                     CLR_TRANS, _
                                     .UpsideDown _
                                     )
                              Else
                                Call mEditRenderer.MaskBltColor( _
                                     m_oScreen.DIB, _
                                     -m_xScreen + .lpRect.x1, .lpRect.y1, _
                                     uTERGFX(.ID).Width, uTERGFX(.ID).Height, _
                                     m_lScreenBackcolorRev, _
                                     uTERGFX(.ID).DIB, _
                                     0, 0, _
                                     CLR_TRANS, _
                                     .UpsideDown _
                                     )
                            End If
                        End If
                      Else
                        If (.NotOverlap) Then
                            Call mEditRenderer.MaskBltOverlap( _
                                 m_oScreen.DIB, _
                                 -m_xScreen + .lpRect.x1, .lpRect.y1, _
                                 uTERGFX(.ID).Width, uTERGFX(.ID).Height, _
                                 m_lScreenBackcolorRev, _
                                 uTERGFX(.ID).DIB, _
                                 0, 0, _
                                 CLR_TRANS, _
                                 .UpsideDown _
                                 )
                          Else
                            Call mEditRenderer.MaskBlt( _
                                 m_oScreen.DIB, _
                                 -m_xScreen + .lpRect.x1, .lpRect.y1, _
                                 uTERGFX(.ID).Width, uTERGFX(.ID).Height, _
                                 uTERGFX(.ID).DIB, _
                                 0, 0, _
                                 CLR_TRANS, _
                                 .UpsideDown _
                                 )
                        End If
                    End If
                End If
            End If
        End With
    Next i
End Sub

Private Sub pvDrawObjects()

  Dim i As Long
    
    For i = 1 To uLEVEL.Objects
        If (uLEVEL.Object(i).OnTerrain) Then
            Call pvDrawObject(i)
        End If
    Next i
    For i = 1 To uLEVEL.Objects
        If (uLEVEL.Object(i).OnTerrain = False) Then
            Call pvDrawObject(i)
        End If
    Next i
End Sub

Private Sub pvDrawObject( _
            ByVal Idx As Integer _
            )
  
    With uLEVEL.Object(Idx)
        If (.lpRect.x1 >= m_xScreen And .lpRect.x1 < m_xScreen + VIEW_WIDTH) Or _
           (.lpRect.x2 >= m_xScreen And .lpRect.x2 < m_xScreen + VIEW_WIDTH) Then
            If (m_bHighlightSelected And _
               (m_eSelectionType = [eObject]) And _
               (m_nSelectionIdx = Idx)) Then
                If (.OnTerrain) Then
                    Call mEditRenderer.MaskBltLightenOverlapNot( _
                         m_oScreen.DIB, _
                         -m_xScreen + .lpRect.x1, .lpRect.y1, _
                         uOBJGFX(.ID).Width, uOBJGFX(.ID).Height, _
                         m_lScreenBackcolorRev, _
                         uOBJGFX(.ID).DIB, _
                         0, 0, _
                         CLR_TRANS _
                         )
                  Else
                    If (.NotOverlap) Then
                        Call mEditRenderer.MaskBltLightenOverlap( _
                             m_oScreen.DIB, _
                             -m_xScreen + .lpRect.x1, .lpRect.y1, _
                             uOBJGFX(.ID).Width, uOBJGFX(.ID).Height, _
                             m_lScreenBackcolorRev, _
                             uOBJGFX(.ID).DIB, _
                             0, 0, _
                             CLR_TRANS _
                             )
                      Else
                        Call mEditRenderer.MaskBltLighten( _
                             m_oScreen.DIB, _
                             -m_xScreen + .lpRect.x1, .lpRect.y1, _
                             uOBJGFX(.ID).Width, uOBJGFX(.ID).Height, _
                             uOBJGFX(.ID).DIB, _
                             0, 0, _
                             CLR_TRANS _
                             )
                    End If
                End If
              Else
                If (.OnTerrain) Then
                    Call mEditRenderer.MaskBltOverlapNot( _
                         m_oScreen.DIB, _
                         -m_xScreen + .lpRect.x1, .lpRect.y1, _
                         uOBJGFX(.ID).Width, uOBJGFX(.ID).Height, _
                         m_lScreenBackcolorRev, _
                         uOBJGFX(.ID).DIB, _
                         0, 0, _
                         CLR_TRANS _
                         )
                  Else
                    If (.NotOverlap) Then
                        Call mEditRenderer.MaskBltOverlap( _
                             m_oScreen.DIB, _
                             -m_xScreen + .lpRect.x1, .lpRect.y1, _
                             uOBJGFX(.ID).Width, uOBJGFX(.ID).Height, _
                             m_lScreenBackcolorRev, _
                             uOBJGFX(.ID).DIB, _
                             0, 0, _
                             CLR_TRANS _
                             )
                      Else
                        Call mEditRenderer.MaskBlt( _
                             m_oScreen.DIB, _
                             -m_xScreen + .lpRect.x1, .lpRect.y1, _
                             uOBJGFX(.ID).Width, uOBJGFX(.ID).Height, _
                             uOBJGFX(.ID).DIB, _
                             0, 0, _
                             CLR_TRANS _
                             )
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub pvDrawSteel()

  Dim i     As Long
  Dim uRect As RECTI
  
    For i = 1 To uLEVEL.SteelAreas
        With uLEVEL.SteelArea(i)
            If (.lpRect.x1 >= m_xScreen And .lpRect.x1 < m_xScreen + VIEW_WIDTH) Or _
               (.lpRect.x2 >= m_xScreen And .lpRect.x2 < m_xScreen + VIEW_WIDTH) Then
                Call pvCopyRect(uRect, .lpRect)
                Call pvOffsetRect(uRect, -m_xScreen, 0)
                With uRect
                    Call mEditRenderer.MaskRectOr( _
                         m_oScreen.DIB, _
                         .x1, .y1, _
                         .x2 - .x1, .y2 - .y1, _
                         CLR_BLUE _
                         )
                End With
            End If
        End With
    Next i
End Sub
        
Private Sub pvDrawTriggerAreas()
  
  Dim i     As Long
  Dim uRect As RECTI
    
    For i = 1 To uLEVEL.Objects
        With uLEVEL.Object(i)
            If (uOBJGFX(.ID).TriggerEffect > 0) Then
                Call pvCopyRect(uRect, uOBJGFX(.ID).lpTriggerRect)
                Call pvOffsetRect(uRect, -m_xScreen + .lpRect.x1, .lpRect.y1)
                With uRect
                    Call mEditRenderer.MaskRectOr( _
                         m_oScreen.DIB, _
                         .x1, .y1, _
                         .x2 - .x1, .y2 - .y1, _
                         CLR_GREEN _
                         )
                End With
            End If
        End With
    Next i
End Sub
        
Private Sub pvResetSelection()
    
    m_nSelectionIdx = 0
    Call pvSetRectEmpty(m_uSelectionRect)
    Call pvUpdateSelectionFlags
End Sub

Private Sub pvDrawSelectionBox()

  Dim Clr     As Long
  Dim hPen    As LongPtr
  Dim hOldPen As LongPtr
  Dim uPt     As POINTAPI

    If (pvIsRectEmpty(m_uSelectionRect) = False) Then
        
        '-- Set color
        Select Case m_eSelectionType
            Case [eObject]
                Clr = vbGreen
            Case [eTerrain]
                Clr = vbYellow
            Case [eSteel]
                Clr = vbCyan
        End Select
            
        '-- Create selection box pen
        hPen = CreatePen(PS_SOLID, 1, Clr)
        hOldPen = SelectObject(m_hDC, hPen)
        
        '-- Draw box
        With m_uSelectionRect
            Call MoveToEx(m_hDC, -m_xScreen + .x1, .y1, uPt)
            Call LineTo(m_hDC, -m_xScreen + .x2 - 1, .y1)
            Call LineTo(m_hDC, -m_xScreen + .x2 - 1, .y2 - 1)
            Call LineTo(m_hDC, -m_xScreen + .x1, .y2 - 1)
            Call LineTo(m_hDC, -m_xScreen + .x1, .y1)
        End With
        
        '-- Unselect and destroy pen
        Call SelectObject(m_hDC, hOldPen)
        Call DeleteObject(hPen)
    End If
End Sub

Private Sub pvUpdateLevelInfo()
    
    With uLEVEL
        fEdit.txtTitle = RTrim$(.Title)
        fEdit.txtLemsToLetOut = .LemsToLetOut
        fEdit.txtLemsToBeSaved = .LemsToBeSaved
        fEdit.txtReleaseRate = .ReleaseRate
        fEdit.txtPlayingTime = .PlayingTime
        fEdit.txtScreenStart = .ScreenStart
        fEdit.txtSkill(0).Text = .MaxClimbers
        fEdit.txtSkill(1).Text = .MaxFloaters
        fEdit.txtSkill(2).Text = .MaxBombers
        fEdit.txtSkill(3).Text = .MaxBlockers
        fEdit.txtSkill(4).Text = .MaxBuilders
        fEdit.txtSkill(5).Text = .MaxBashers
        fEdit.txtSkill(6).Text = .MaxMiners
        fEdit.txtSkill(7).Text = .MaxDiggers
    End With
End Sub

Private Sub pvUpdateSelectionInfo()
    
    If (SelectionExists) Then
        With m_uSelectionRect
            fEdit.lblSelectionPositionVal.Caption = .x1 & "," & .y1
            fEdit.lblSelectionSizeVal.Caption = .x2 - .x1 & "x" & .y2 - .y1
            fEdit.lblzOrderVal.Caption = Format$(m_nSelectionIdx, "000")
            If (m_eSelectionType = [eSteel]) Then
                fEdit.txtSteelAreaWidth = .x2 - .x1
                fEdit.txtSteelAreaHeight = .y2 - .y1
            End If
        End With
      Else
        fEdit.lblSelectionPositionVal.Caption = vbNullString
        fEdit.lblSelectionSizeVal.Caption = vbNullString
        fEdit.lblzOrderVal.Caption = vbNullString
        fEdit.txtSteelAreaWidth = MIN_STEELAREASIZE
        fEdit.txtSteelAreaHeight = MIN_STEELAREASIZE
    End If
End Sub

Private Sub pvUpdateSelectionFlags()
    
    fEdit.cmdUp.Enabled = SelectionExists
    fEdit.cmdDown.Enabled = SelectionExists
    fEdit.cmdLeft.Enabled = SelectionExists
    fEdit.cmdRight.Enabled = SelectionExists
    
    fEdit.cmdZOrderUp.Enabled = SelectionExists
    fEdit.cmdZOrderDown.Enabled = SelectionExists
    
    fEdit.chkNotOverlap.Enabled = (SelectionExists And (m_eSelectionType = [eObject] Or m_eSelectionType = [eTerrain]))
    fEdit.chkOnTerrain.Enabled = (SelectionExists And m_eSelectionType = [eObject])
    fEdit.chkUpsideDown.Enabled = (SelectionExists And m_eSelectionType = [eTerrain])
    fEdit.chkBlack.Enabled = (SelectionExists And m_eSelectionType = [eTerrain])
    
    fEdit.mnuContextSelection(3).Enabled = (m_eSelectionType <> [eSteel])
    fEdit.mnuContextSelection(4).Enabled = (m_eSelectionType <> [eSteel])
    fEdit.mnuContextSelection(5).Visible = (m_eSelectionType <> [eSteel])
    fEdit.mnuContextSelection(6).Visible = (m_eSelectionType = [eObject])
    fEdit.mnuContextSelection(7).Visible = (m_eSelectionType = [eObject])
    fEdit.mnuContextSelection(8).Visible = False
    fEdit.mnuContextSelection(9).Visible = (m_eSelectionType = [eTerrain])
    fEdit.mnuContextSelection(10).Visible = (m_eSelectionType = [eTerrain])
    fEdit.mnuContextSelection(11).Visible = (m_eSelectionType = [eTerrain])
    
    With uLEVEL
        If (SelectionExists) Then
            Select Case m_eSelectionType
                Case [eObject]
                    fEdit.chkNotOverlap = -.Object(m_nSelectionIdx).NotOverlap
                    fEdit.mnuContextSelection(6).Checked = .Object(m_nSelectionIdx).NotOverlap
                    fEdit.chkOnTerrain = -.Object(m_nSelectionIdx).OnTerrain
                    fEdit.mnuContextSelection(7).Checked = .Object(m_nSelectionIdx).OnTerrain
                Case [eTerrain]
                    fEdit.chkNotOverlap = -.TerrainPiece(m_nSelectionIdx).NotOverlap
                    fEdit.mnuContextSelection(9).Checked = .TerrainPiece(m_nSelectionIdx).NotOverlap
                    fEdit.chkBlack = -.TerrainPiece(m_nSelectionIdx).Black
                    fEdit.mnuContextSelection(10).Checked = .TerrainPiece(m_nSelectionIdx).Black
                    fEdit.chkUpsideDown = -.TerrainPiece(m_nSelectionIdx).UpsideDown
                    fEdit.mnuContextSelection(11).Checked = .TerrainPiece(m_nSelectionIdx).UpsideDown
            End Select
        End If
    End With
End Sub

Private Sub pvUpdateStatistics()
    
    With uLEVEL
        fEdit.ucPrgObjects.Value = .Objects
        fEdit.ucPrgObjects.Caption = "Object: " & .Objects & "/" & MAX_OBJECTS
        fEdit.ucPrgTerrainPieces.Value = .TerrainPieces
        fEdit.ucPrgTerrainPieces.Caption = "Terrain: " & .TerrainPieces & "/" & MAX_TERRAINPIECES
        fEdit.ucPrgSteelAreas.Value = .SteelAreas
        fEdit.ucPrgSteelAreas.Caption = "Steel: " & .SteelAreas & "/" & MAX_STEELAREAS
    End With
End Sub
