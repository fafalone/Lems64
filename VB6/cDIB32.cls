VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIB32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' Class:         cDIB32.cls (cDIB simplified)
' Author:        Carles P.V.
' Dependencies:
' Last revision: 2004.09.12
'================================================

Option Explicit

'-- API:

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Type RECT
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Const DIB_RGB_COLORS As Long = 0
Private Const COLORONCOLOR   As Long = 3
Private Const HALFTONE       As Long = 4
Private Const OBJ_BITMAP     As Long = 7

Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, lpBitsInfo As Any, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Private Declare Function SetDIBColorTable Lib "gdi32" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, lpRGBQuad As Any) As Long
Private Declare Function GetDIBColorTable Lib "gdi32" (ByVal hDC As Long, ByVal un1 As Long, ByVal un2 As Long, lpRGBQuad As Any) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ColorRef As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDst As Any, ByVal Length As Long)

'-- Public enums.:

Public Enum eDIB32StretchBltModeConstants
    [eColorOnColor] = COLORONCOLOR ' Nearest
    [eHalftone] = HALFTONE         ' NT-based systems (biliniar interpolation + ordered dithering on palette based bitmaps)
End Enum

'-- Private variables:

Private m_uBIH    As BITMAPINFOHEADER
Private m_hDC     As Long
Private m_hDIB    As Long
Private m_hOldDIB As Long
Private m_lpBits  As Long



'========================================================================================
' Class
'========================================================================================

Private Sub Class_Terminate()

    '-- Destroy current DIB, if any
    Call Me.Destroy
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Function Create( _
                ByVal Width As Long, _
                ByVal Height As Long _
                ) As Long
    
    If (Width > 0 And Height > 0) Then
    
        '-- Destroy previous
        Call Me.Destroy
        
        '-- Define DIB header
        With m_uBIH
            .biSize = Len(m_uBIH)
            .biPlanes = 1
            .biBitCount = 32
            .biWidth = Width
            .biHeight = -Height
            .biSizeImage = (4 * .biWidth) * -.biHeight
        End With
        
        '-- Create DIB and select into a DC
        m_hDC = CreateCompatibleDC(0)
        If (m_hDC <> 0) Then
            m_hDIB = CreateDIBSection(m_hDC, m_uBIH, DIB_RGB_COLORS, m_lpBits, 0, 0)
            If (m_hDIB <> 0) Then
                m_hOldDIB = SelectObject(m_hDC, m_hDIB)
              Else
                Call Me.Destroy
            End If
        End If
        
        '-- Success
        Create = m_lpBits
    End If
End Function

Public Function CreateFromStdPicture( _
                Image As StdPicture, _
                Optional ByVal Flip As Boolean = False _
                ) As Long

  Dim hDC     As Long
  Dim hOldBmp As Long
  Dim uBI     As BITMAP
    
    If (Not Image Is Nothing) Then
        
        '-- Check object type
        If (GetObjectType(Image.handle) = OBJ_BITMAP) Then
        
            '-- Get object info
            Call GetObject(Image.handle, Len(uBI), uBI)
            
            '-- Create DIB
            If (Create(uBI.bmWidth, uBI.bmHeight)) Then
                hDC = CreateCompatibleDC(m_hDC)
                If (hDC <> 0) Then
                    hOldBmp = SelectObject(hDC, Image.handle)
                    
                    With m_uBIH
                        If (Flip) Then
                            Call StretchBlt( _
                                 m_hDC, _
                                 0, -.biHeight - 1, .biWidth, .biHeight, _
                                 hDC, _
                                 0, 0, .biWidth, -.biHeight, _
                                 vbSrcCopy _
                                 )
                          Else
                            Call BitBlt( _
                                 m_hDC, _
                                 0, 0, .biWidth, -.biHeight, _
                                 hDC, _
                                 0, 0, _
                                 vbSrcCopy _
                                 )
                        End If
                    End With
                    
                    '-- Destroy temp. DC
                    Call SelectObject(hDC, hOldBmp)
                    Call DeleteDC(hDC)
                    
                    '-- Success
                    CreateFromStdPicture = m_lpBits
                End If
            End If
        End If
    End If
End Function

Public Function CloneTo( _
                oDIB As cDIB32 _
                ) As Long
    
    If (m_hDIB <> 0) Then
    
        '-- Create dest. DIB
        If (oDIB Is Nothing) Then
            Set oDIB = New cDIB32
        End If
        Call oDIB.Create(m_uBIH.biWidth, -m_uBIH.biHeight)
        
        '-- Copy color data
        Call CopyMemory(ByVal oDIB.lpBits, ByVal m_lpBits, m_uBIH.biSizeImage)
        
        '-- Success
        CloneTo = oDIB.lpBits
    End If
End Function

Public Sub Destroy()

    '-- Destroy DIB
    If (m_hDC <> 0) Then
        If (m_hDIB <> 0) Then
            Call SelectObject(m_hDC, m_hOldDIB)
            Call DeleteObject(m_hDIB)
        End If
        Call DeleteDC(m_hDC)
    End If
    
    '-- Reset BIH structure
    Call ZeroMemory(m_uBIH, Len(m_uBIH))
    
    '-- Reset DIB vars.
    m_hDC = 0
    m_hDIB = 0
    m_hOldDIB = 0
    m_lpBits = 0
End Sub

Public Function LoadBlt( _
                ByVal hSrcDC As Long, _
                Optional ByVal x As Long = 0, _
                Optional ByVal y As Long = 0, _
                Optional ByVal nWidth As Long, _
                Optional ByVal nHeight As Long _
                ) As Long
    
    If (m_hDIB <> 0) Then
        
        If (nWidth = 0) Then
            nWidth = m_uBIH.biWidth
        End If
        If (nHeight = 0) Then
            nHeight = -m_uBIH.biHeight
        End If
        
        LoadBlt = BitBlt( _
                  m_hDC, _
                  0, 0, nWidth, nHeight, _
                  hSrcDC, _
                  x, y, _
                  vbSrcCopy _
                  )
    End If
End Function

Public Function Paint( _
                ByVal hDC As Long, _
                Optional ByVal x As Long = 0, Optional ByVal y As Long = 0, _
                Optional ByVal ROP As RasterOpConstants = vbSrcCopy, _
                Optional ByVal StretchBltMode As eDIB32StretchBltModeConstants = [eColorOnColor] _
                ) As Long
    
    Paint = Me.Stretch( _
            hDC, _
            x, y, _
            m_uBIH.biWidth, -m_uBIH.biHeight, _
            , , , , _
            ROP, _
            StretchBltMode _
            )
End Function

Public Function Stretch(ByVal hDC As Long, _
                ByVal x As Long, ByVal y As Long, _
                ByVal nWidth As Long, ByVal nHeight As Long, _
                Optional ByVal xSrc As Long, Optional ByVal ySrc As Long, _
                Optional ByVal nSrcWidth As Long, Optional ByVal nSrcHeight As Long, _
                Optional ByVal lROP As RasterOpConstants = vbSrcCopy, _
                Optional ByVal StretchBltMode As eDIB32StretchBltModeConstants = [eColorOnColor] _
                ) As Long
  
  Dim lOldMode As Long
    
    If (m_hDIB <> 0) Then
        
        If (nSrcWidth = 0) Then
            nSrcWidth = m_uBIH.biWidth
        End If
        If (nSrcHeight = 0) Then
            nSrcHeight = -m_uBIH.biHeight
        End If
        
        lOldMode = SetStretchBltMode(hDC, StretchBltMode)
        
        Stretch = StretchBlt( _
                  hDC, _
                  x, y, _
                  nWidth, nHeight, _
                  m_hDC, _
                  xSrc, ySrc, _
                  nSrcWidth, nSrcHeight, _
                  lROP _
                  )
                  
        Call SetStretchBltMode(hDC, lOldMode)
    End If
End Function

Public Sub Cls(Optional ByVal ColorRef As Long = &H0&)
  
  Dim uRect  As RECT
  Dim hBrush As Long
   
    If (m_hDIB <> 0) Then
        Call SetRect(uRect, 0, 0, m_uBIH.biWidth, -m_uBIH.biHeight)
        hBrush = CreateSolidBrush(ColorRef)
        Call FillRect(m_hDC, uRect, hBrush)
        Call DeleteObject(hBrush)
    End If
End Sub

Public Sub Reset()
    
    If (m_hDIB <> 0) Then
        Call ZeroMemory(ByVal m_lpBits, m_uBIH.biSizeImage)
    End If
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Get hDC() As Long
    hDC = m_hDC
End Property

Public Property Get hDIB() As Long
    hDIB = m_hDIB
End Property

Public Property Get lpBits() As Long
    lpBits = m_lpBits
End Property

Public Property Get Width() As Long
    Width = m_uBIH.biWidth
End Property

Public Property Get Height() As Long
    Height = -m_uBIH.biHeight
End Property

Public Property Get BytesPerScanline() As Long
    BytesPerScanline = 4 * m_uBIH.biWidth
End Property

Public Property Get Size() As Long
    Size = m_uBIH.biSizeImage
End Property
