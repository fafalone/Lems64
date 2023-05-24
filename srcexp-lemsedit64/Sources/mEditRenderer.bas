Attribute VB_Name = "mEditRenderer"
Option Explicit

'-- API:

#If Win64 Then
	Private Const cbPtr = 8
#Else
    Private Const cbPtr = 4
#End If

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound   As Long
End Type

Private Type SAFEARRAY1D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As LongPtr
    Bounds     As SAFEARRAYBOUND
End Type

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As LongPtr

Private Type RECT
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Declare PtrSafe Function SetRect Lib "user32" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare PtrSafe Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Private Declare PtrSafe Function IsRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare PtrSafe Function OffsetRect Lib "user32" (lprc As RECT, ByVal dx As Long, ByVal dy As Long) As Long

'//

Public Const CLR_TRANS   As Long = &HFF00FF
Public Const CLR_RED     As Long = &HFF0000
Public Const CLR_GREEN   As Long = &HFF00&
Public Const CLR_BLUE    As Long = &HFF&
Public Const CLR_LIGHTEN As Long = &H7F7F7F
Public Const CLR_FIXMASK As Long = &HFFFFFF



'========================================================================================
' Methods
'========================================================================================

Public Sub MaskBlt( _
           oDstDIB As cDIB32, _
           ByVal xDst As Long, ByVal yDst As Long, _
           ByVal wDst As Long, ByVal hDst As Long, _
           oSrcDIB As cDIB32, _
           ByVal xSrc As Long, ByVal ySrc As Long, _
           ByVal SrcMaskColor As Long, _
           Optional ByVal UpsideDown As Boolean = False _
           )

  Dim lDstBits() As Long
  Dim lSrcBits() As Long
  Dim uDstSA     As SAFEARRAY1D
  Dim uSrcSA     As SAFEARRAY1D
  
  Dim i As Long, i1 As Long, j1 As Long, k1 As Long, r1 As Long
  Dim j As Long, i2 As Long, j2 As Long, k2 As Long, r2 As Long
    
    If (pvCheckDIBDIBRects( _
        oDstDIB, xDst, yDst, wDst, hDst, _
        oSrcDIB, xSrc, ySrc, UpsideDown, _
        i1, i2, j1, j2, k1, k2, r1, r2 _
        )) Then
        
        Call pvMapDIB(uDstSA, lDstBits(), oDstDIB)
        Call pvMapDIB(uSrcSA, lSrcBits(), oSrcDIB)
        
        For j = j1 To j2
            k2 = k1 - i1
            For i = i1 To i1 + i2
                If ((lSrcBits(i) And CLR_FIXMASK) <> SrcMaskColor) Then
                    lDstBits(i + k2) = lSrcBits(i)
                End If
            Next i
            i1 = i1 + r2
            k1 = k1 + r1
        Next j
        
        Call pvUnmapDIB(lDstBits())
        Call pvUnmapDIB(lSrcBits())
    End If
End Sub

Public Sub MaskBltOverlap( _
           oDstDIB As cDIB32, _
           ByVal xDst As Long, ByVal yDst As Long, _
           ByVal wDst As Long, ByVal hDst As Long, _
           ByVal DstMaskColor As Long, _
           oSrcDIB As cDIB32, _
           ByVal xSrc As Long, ByVal ySrc As Long, _
           ByVal SrcMaskColor As Long, _
           Optional ByVal UpsideDown As Boolean = False _
           )

  Dim lDstBits() As Long
  Dim lSrcBits() As Long
  Dim uDstSA     As SAFEARRAY1D
  Dim uSrcSA     As SAFEARRAY1D
  
  Dim i As Long, i1 As Long, j1 As Long, k1 As Long, r1 As Long
  Dim j As Long, i2 As Long, j2 As Long, k2 As Long, r2 As Long
    
    If (pvCheckDIBDIBRects( _
        oDstDIB, xDst, yDst, wDst, hDst, _
        oSrcDIB, xSrc, ySrc, UpsideDown, _
        i1, i2, j1, j2, k1, k2, r1, r2 _
        )) Then
        
        Call pvMapDIB(uDstSA, lDstBits(), oDstDIB)
        Call pvMapDIB(uSrcSA, lSrcBits(), oSrcDIB)
        
        For j = j1 To j2
            k2 = k1 - i1
            For i = i1 To i1 + i2
                If ((lSrcBits(i) And CLR_FIXMASK) <> SrcMaskColor) Then
                    If ((lDstBits(i + k2) And CLR_FIXMASK) = DstMaskColor) Then
                        lDstBits(i + k2) = lSrcBits(i)
                    End If
                End If
            Next i
            i1 = i1 + r2
            k1 = k1 + r1
        Next j
        
        Call pvUnmapDIB(lDstBits())
        Call pvUnmapDIB(lSrcBits())
    End If
End Sub

Public Sub MaskBltOverlapNot( _
           oDstDIB As cDIB32, _
           ByVal xDst As Long, ByVal yDst As Long, _
           ByVal wDst As Long, ByVal hDst As Long, _
           ByVal DstMaskColor As Long, _
           oSrcDIB As cDIB32, _
           ByVal xSrc As Long, ByVal ySrc As Long, _
           ByVal SrcMaskColor As Long, _
           Optional ByVal UpsideDown As Boolean = False _
           )

  Dim lDstBits() As Long
  Dim lSrcBits() As Long
  Dim uDstSA     As SAFEARRAY1D
  Dim uSrcSA     As SAFEARRAY1D
  
  Dim i As Long, i1 As Long, j1 As Long, k1 As Long, r1 As Long
  Dim j As Long, i2 As Long, j2 As Long, k2 As Long, r2 As Long
    
    If (pvCheckDIBDIBRects( _
        oDstDIB, xDst, yDst, wDst, hDst, _
        oSrcDIB, xSrc, ySrc, UpsideDown, _
        i1, i2, j1, j2, k1, k2, r1, r2 _
        )) Then
        
        Call pvMapDIB(uDstSA, lDstBits(), oDstDIB)
        Call pvMapDIB(uSrcSA, lSrcBits(), oSrcDIB)
        
        For j = j1 To j2
            k2 = k1 - i1
            For i = i1 To i1 + i2
                If ((lSrcBits(i) And CLR_FIXMASK) <> SrcMaskColor) Then
                    If ((lDstBits(i + k2) And CLR_FIXMASK) <> DstMaskColor) Then
                        lDstBits(i + k2) = lSrcBits(i)
                    End If
                End If
            Next i
            i1 = i1 + r2
            k1 = k1 + r1
        Next j
        
        Call pvUnmapDIB(lDstBits())
        Call pvUnmapDIB(lSrcBits())
    End If
End Sub

Public Sub MaskBltColor( _
           oDstDIB As cDIB32, _
           ByVal xDst As Long, ByVal yDst As Long, _
           ByVal wDst As Long, ByVal hDst As Long, _
           ByVal DstColor As Long, _
           oSrcDIB As cDIB32, _
           ByVal xSrc As Long, ByVal ySrc As Long, _
           ByVal SrcMaskColor As Long, _
           Optional ByVal UpsideDown As Boolean = False _
           )

  Dim lDstBits() As Long
  Dim lSrcBits() As Long
  Dim uDstSA     As SAFEARRAY1D
  Dim uSrcSA     As SAFEARRAY1D
  
  Dim i As Long, i1 As Long, j1 As Long, k1 As Long, r1 As Long
  Dim j As Long, i2 As Long, j2 As Long, k2 As Long, r2 As Long
    
    If (pvCheckDIBDIBRects( _
        oDstDIB, xDst, yDst, wDst, hDst, _
        oSrcDIB, xSrc, ySrc, UpsideDown, _
        i1, i2, j1, j2, k1, k2, r1, r2 _
        )) Then
        
        Call pvMapDIB(uDstSA, lDstBits(), oDstDIB)
        Call pvMapDIB(uSrcSA, lSrcBits(), oSrcDIB)
        
        For j = j1 To j2
            k2 = k1 - i1
            For i = i1 To i1 + i2
                If ((lSrcBits(i) And CLR_FIXMASK) <> SrcMaskColor) Then
                    lDstBits(i + k2) = DstColor
                End If
            Next i
            i1 = i1 + r2
            k1 = k1 + r1
        Next j
        
        Call pvUnmapDIB(lDstBits())
        Call pvUnmapDIB(lSrcBits())
    End If
End Sub

Public Sub MaskBltColorOverlap( _
           oDstDIB As cDIB32, _
           ByVal xDst As Long, ByVal yDst As Long, _
           ByVal wDst As Long, ByVal hDst As Long, _
           ByVal DstMaskColor As Long, _
           ByVal DstColor As Long, _
           oSrcDIB As cDIB32, _
           ByVal xSrc As Long, ByVal ySrc As Long, _
           ByVal SrcMaskColor As Long, _
           Optional ByVal UpsideDown As Boolean = False _
           )

  Dim lDstBits() As Long
  Dim lSrcBits() As Long
  Dim uDstSA     As SAFEARRAY1D
  Dim uSrcSA     As SAFEARRAY1D
  
  Dim i As Long, i1 As Long, j1 As Long, k1 As Long, r1 As Long
  Dim j As Long, i2 As Long, j2 As Long, k2 As Long, r2 As Long
    
    If (pvCheckDIBDIBRects( _
        oDstDIB, xDst, yDst, wDst, hDst, _
        oSrcDIB, xSrc, ySrc, UpsideDown, _
        i1, i2, j1, j2, k1, k2, r1, r2 _
        )) Then
        
        Call pvMapDIB(uDstSA, lDstBits(), oDstDIB)
        Call pvMapDIB(uSrcSA, lSrcBits(), oSrcDIB)
        
        For j = j1 To j2
            k2 = k1 - i1
            For i = i1 To i1 + i2
                If ((lSrcBits(i) And CLR_FIXMASK) <> SrcMaskColor) Then
                    If ((lDstBits(i + k2) And CLR_FIXMASK) = DstMaskColor) Then
                        lDstBits(i + k2) = DstColor
                    End If
                End If
            Next i
            i1 = i1 + r2
            k1 = k1 + r1
        Next j
        
        Call pvUnmapDIB(lDstBits())
        Call pvUnmapDIB(lSrcBits())
    End If
End Sub

Public Sub MaskBltLighten( _
           oDstDIB As cDIB32, _
           ByVal xDst As Long, ByVal yDst As Long, _
           ByVal wDst As Long, ByVal hDst As Long, _
           oSrcDIB As cDIB32, _
           ByVal xSrc As Long, ByVal ySrc As Long, _
           ByVal SrcMaskColor As Long, _
           Optional ByVal UpsideDown As Boolean = False _
           )
           
  Dim lDstBits() As Long
  Dim lSrcBits() As Long
  Dim uDstSA     As SAFEARRAY1D
  Dim uSrcSA     As SAFEARRAY1D
  
  Dim i As Long, i1 As Long, j1 As Long, k1 As Long, r1 As Long
  Dim j As Long, i2 As Long, j2 As Long, k2 As Long, r2 As Long
    
    If (pvCheckDIBDIBRects( _
        oDstDIB, xDst, yDst, wDst, hDst, _
        oSrcDIB, xSrc, ySrc, UpsideDown, _
        i1, i2, j1, j2, k1, k2, r1, r2 _
        )) Then
        
        Call pvMapDIB(uDstSA, lDstBits(), oDstDIB)
        Call pvMapDIB(uSrcSA, lSrcBits(), oSrcDIB)
        
        For j = j1 To j2
            k2 = k1 - i1
            For i = i1 To i1 + i2
                If ((lSrcBits(i) And CLR_FIXMASK) <> SrcMaskColor) Then
                    lDstBits(i + k2) = lSrcBits(i) \ 2 + CLR_LIGHTEN
                End If
            Next i
            i1 = i1 + r2
            k1 = k1 + r1
        Next j
        
        Call pvUnmapDIB(lDstBits())
        Call pvUnmapDIB(lSrcBits())
    End If
End Sub

Public Sub MaskBltLightenOverlap( _
           oDstDIB As cDIB32, _
           ByVal xDst As Long, ByVal yDst As Long, _
           ByVal wDst As Long, ByVal hDst As Long, _
           ByVal DstMaskColor As Long, _
           oSrcDIB As cDIB32, _
           ByVal xSrc As Long, ByVal ySrc As Long, _
           ByVal SrcMaskColor As Long, _
           Optional ByVal UpsideDown As Boolean = False _
           )
           
  Dim lDstBits() As Long
  Dim lSrcBits() As Long
  Dim uDstSA     As SAFEARRAY1D
  Dim uSrcSA     As SAFEARRAY1D
  
  Dim i As Long, i1 As Long, j1 As Long, k1 As Long, r1 As Long
  Dim j As Long, i2 As Long, j2 As Long, k2 As Long, r2 As Long
    
    If (pvCheckDIBDIBRects( _
        oDstDIB, xDst, yDst, wDst, hDst, _
        oSrcDIB, xSrc, ySrc, UpsideDown, _
        i1, i2, j1, j2, k1, k2, r1, r2 _
        )) Then
        
        Call pvMapDIB(uDstSA, lDstBits(), oDstDIB)
        Call pvMapDIB(uSrcSA, lSrcBits(), oSrcDIB)
        
        For j = j1 To j2
            k2 = k1 - i1
            For i = i1 To i1 + i2
                If ((lSrcBits(i) And CLR_FIXMASK) <> SrcMaskColor) Then
                    If ((lDstBits(i + k2) And CLR_FIXMASK) = DstMaskColor) Then
                        lDstBits(i + k2) = lSrcBits(i) \ 2 + CLR_LIGHTEN
                    End If
                End If
            Next i
            i1 = i1 + r2
            k1 = k1 + r1
        Next j
        
        Call pvUnmapDIB(lDstBits())
        Call pvUnmapDIB(lSrcBits())
    End If
End Sub

Public Sub MaskBltLightenOverlapNot( _
           oDstDIB As cDIB32, _
           ByVal xDst As Long, ByVal yDst As Long, _
           ByVal wDst As Long, ByVal hDst As Long, _
           ByVal DstMaskColor As Long, _
           oSrcDIB As cDIB32, _
           ByVal xSrc As Long, ByVal ySrc As Long, _
           ByVal SrcMaskColor As Long, _
           Optional ByVal UpsideDown As Boolean = False _
           )
           
  Dim lDstBits() As Long
  Dim lSrcBits() As Long
  Dim uDstSA     As SAFEARRAY1D
  Dim uSrcSA     As SAFEARRAY1D
  
  Dim i As Long, i1 As Long, j1 As Long, k1 As Long, r1 As Long
  Dim j As Long, i2 As Long, j2 As Long, k2 As Long, r2 As Long
    
    If (pvCheckDIBDIBRects( _
        oDstDIB, xDst, yDst, wDst, hDst, _
        oSrcDIB, xSrc, ySrc, UpsideDown, _
        i1, i2, j1, j2, k1, k2, r1, r2 _
        )) Then
        
        Call pvMapDIB(uDstSA, lDstBits(), oDstDIB)
        Call pvMapDIB(uSrcSA, lSrcBits(), oSrcDIB)
        
        For j = j1 To j2
            k2 = k1 - i1
            For i = i1 To i1 + i2
                If ((lSrcBits(i) And CLR_FIXMASK) <> SrcMaskColor) Then
                    If ((lDstBits(i + k2) And CLR_FIXMASK) <> DstMaskColor) Then
                        lDstBits(i + k2) = lSrcBits(i) \ 2 + CLR_LIGHTEN
                    End If
                End If
            Next i
            i1 = i1 + r2
            k1 = k1 + r1
        Next j
        
        Call pvUnmapDIB(lDstBits())
        Call pvUnmapDIB(lSrcBits())
    End If
End Sub

Public Sub MaskRectOr( _
           oDstDIB As cDIB32, _
           ByVal xDst As Long, ByVal yDst As Long, _
           ByVal wDst As Long, ByVal hDst As Long, _
           ByVal DstOrColor As Long _
           )

  Dim lDstBits() As Long
  Dim uDstSA     As SAFEARRAY1D
  
  Dim i As Long, i1 As Long, j1 As Long, r As Long
  Dim j As Long, i2 As Long, j2 As Long
    
    If (pvCheckDIBRectRects( _
        oDstDIB, xDst, yDst, wDst, hDst, _
        i1, i2, j1, j2, r _
        )) Then
        
        Call pvMapDIB(uDstSA, lDstBits(), oDstDIB)
        
        For j = j1 To j2
            For i = i1 To i1 + i2
                lDstBits(i) = lDstBits(i) Or DstOrColor
            Next i
            i1 = i1 + r
        Next j
        
        Call pvUnmapDIB(lDstBits())
    End If
End Sub

'========================================================================================
' Private
'========================================================================================

Private Function pvCheckDIBDIBRects( _
                 oDstDIB As cDIB32, _
                 ByVal xDst As Long, ByVal yDst As Long, _
                 ByVal wDst As Long, ByVal hDst As Long, _
                 oSrcDIB As cDIB32, _
                 ByVal xSrc As Long, ByVal ySrc As Long, _
                 ByVal UpsideDown As Boolean, _
                 i1 As Long, i2 As Long, _
                 j1 As Long, j2 As Long, _
                 k1 As Long, k2 As Long, _
                 r1 As Long, r2 As Long _
                 ) As Boolean

  Dim uDstRect As RECT
  Dim uSrcRect As RECT

    Call SetRect(uDstRect, 0, 0, oDstDIB.Width, oDstDIB.Height)
    Call SetRect(uSrcRect, xDst, yDst, xDst + wDst, yDst + hDst)
    Call IntersectRect(uDstRect, uDstRect, uSrcRect)
    
    If (IsRectEmpty(uDstRect) = 0) Then
        
        Call OffsetRect(uDstRect, -xDst, -yDst)
        
        With uDstRect
            i1 = .x1
            i2 = .x2 - .x1 - 1
            If (UpsideDown) Then
                j1 = oSrcDIB.Height - .y2
                j2 = oSrcDIB.Height - .y1 - 1
              Else
                j1 = .y1
                j2 = .y2 - 1
            End If
        End With
        
        If (UpsideDown) Then
            r1 = -oDstDIB.Width
            k1 = (i1 + xDst) - (uDstRect.y2 - 1 + yDst) * r1
          Else
            r1 = oDstDIB.Width
            k1 = (i1 + xDst) + (uDstRect.y1 - 0 + yDst) * r1
        End If
        r2 = oSrcDIB.Width
        i1 = (i1 + xSrc) + (j1 + ySrc) * r2
        
        pvCheckDIBDIBRects = True
    End If
End Function

Private Function pvCheckDIBRectRects( _
                 oDstDIB As cDIB32, _
                 ByVal xDst As Long, ByVal yDst As Long, _
                 ByVal wDst As Long, ByVal hDst As Long, _
                 i1 As Long, i2 As Long, _
                 j1 As Long, j2 As Long, _
                 r As Long _
                 ) As Boolean

  Dim uDstRect As RECT
  Dim uSrcRect As RECT

    Call SetRect(uDstRect, 0, 0, oDstDIB.Width, oDstDIB.Height)
    Call SetRect(uSrcRect, xDst, yDst, xDst + wDst, yDst + hDst)
    Call IntersectRect(uDstRect, uDstRect, uSrcRect)
    
    If (IsRectEmpty(uDstRect) = 0) Then
        
        r = oDstDIB.Width
        
        With uDstRect
            i1 = .y1 * r + .x1
            i2 = .x2 - .x1 - 1
            j1 = .y1
            j2 = .y2 - 1
        End With
        
        pvCheckDIBRectRects = True
    End If
End Function

Private Sub pvMapDIB( _
            uSA As SAFEARRAY1D, _
            lBits() As Long, _
            oDIB As cDIB32 _
            )
    
    With uSA
        .cbElements = 4
        .cDims = 1
        .Bounds.lLbound = 0
        .Bounds.cElements = oDIB.Width * oDIB.Height
        .pvData = oDIB.lpBits
    End With
    Call CopyMemory(ByVal VarPtrArray(lBits()), VarPtr(uSA), cbPtr)
End Sub

Private Sub pvUnmapDIB( _
            lBits() As Long _
            )
    
    Call CopyMemory(ByVal VarPtrArray(lBits()), 0&, cbPtr)
End Sub
