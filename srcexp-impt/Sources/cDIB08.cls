VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIB08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================================
' Class:         cDIB08.cls (cDIB simplified)
' Author:        Carles P.V.
' Dependencies:
' Last revision: 2005.08.16
'========================================================================================

Option Explicit

'-- API:

Private Type BITMAPFILEHEADER
    bfType      As Integer
    bfSize      As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits   As Long
End Type

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

Private Type BITMAPINFO
    bmiHeader       As BITMAPINFOHEADER
    bmiColors(1023) As Byte
End Type

Private Const BITMAPTYPE     As Integer = &H4D42
Private Const DIB_RGB_COLORS As Long = 0
Private Const COLORONCOLOR   As Long = 3
Private Const HALFTONE       As Long = 4

Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDst As Any, ByVal Length As Long)

Private Const CF_DIB As Long = 8

Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long

Private Const GMEM_FIXED    As Long = &H0
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_DDESHARE As Long = &H2000
Private Const GMEM_ZEROINIT As Long = &H40

Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long

'-- Public enums.:

Public Enum eDIB08StretchBltModeConstants
    [eColorOnColor] = COLORONCOLOR ' Nearest
    [eHalftone] = HALFTONE         ' NT-based systems (biliniar interpolation + ordered dithering on palette based bitmaps)
End Enum

'-- Private variables:

Private m_uBI        As BITMAPINFO
Private m_aBits()    As Byte
Private m_lScanWidth As Long



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
                ByVal Height As Long, _
                Optional ByVal BkColorIdx As Byte = 0 _
                ) As Boolean
    
    
    On Error GoTo errH
    
    If (Width > 0 And Height > 0) Then
    
        '-- Scan width
        m_lScanWidth = 4 * ((Width * 8 + 31) \ 32)
        
        '-- Define DIB header
        With m_uBI
            With .bmiHeader
                .biSize = 40
                .biPlanes = 1
                .biBitCount = 8
                .biWidth = Width
                .biHeight = -Height
                .biSizeImage = m_lScanWidth * -.biHeight
            End With
        End With
        
        '-- Initialize DIB data array
        ReDim m_aBits(0 To m_uBI.bmiHeader.biSizeImage - 1)
        
        '-- Fill?
        If (BkColorIdx > 0) Then
            Call Me.Cls(BkColorIdx)
        End If
        
        '-- Success
        Create = True
    End If
    
errH:
    On Error GoTo 0
End Function

Public Function CreateFromBitmapFile( _
                ByVal Filename As String _
                ) As Boolean
                
  Dim hFile As Long
  Dim uBFH  As BITMAPFILEHEADER
 
    On Error GoTo errH
    
    '-- File exists ?
    If (VBA.FileLen(Filename)) Then
    
        hFile = VBA.FreeFile()
        
        Open Filename For Binary Access Read As #hFile
            
            '-- Get file header
            Get #hFile, , uBFH
            
            '-- Minimum check
            If (uBFH.bfType = BITMAPTYPE) Then
                
                With m_uBI
                    
                    '-- Get bitmap header
                    Get #hFile, , .bmiHeader
                    
                    '-- 8-bpp ?
                    If (.bmiHeader.biBitCount = 8) Then
                        
                        '-- Get palette
                        Get #hFile, , .bmiColors()
                    
                        With .bmiHeader
                            
                            '-- Set upside-down DIB and get scanline width
                            .biHeight = -Abs(.biHeight)
                            m_lScanWidth = 4 * ((.biWidth * 8 + 31) \ 32)
                            
                            '-- Resize color data and read
                            ReDim m_aBits(.biSizeImage - 1)
                            Get #hFile, , m_aBits()
                        End With
                        
                        '-- Success
                        CreateFromBitmapFile = True
                    End If
                End With
            End If
        Close #hFile
    End If

errH:
    On Error GoTo 0
End Function

Public Function CloneTo( _
                oDIB As cDIB08 _
                ) As Boolean

    On Error GoTo errH
    
    If (m_uBI.bmiHeader.biSizeImage) Then
        
        With m_uBI
        
            '-- Check if initialized
            If (oDIB Is Nothing) Then
                Set oDIB = New cDIB08
            End If
            
            '-- Create dest. DIB
            Call oDIB.Create(.bmiHeader.biWidth, -.bmiHeader.biHeight)
    
            '-- Copy palette
            Call oDIB.SetPalette(.bmiColors())
    
            '-- Copy color data
            Call CopyMemory(ByVal oDIB.lpBits, m_aBits(0), .bmiHeader.biSizeImage)
    
            '-- Success
            CloneTo = True
        End With
    End If
    
errH:
    On Error GoTo 0
End Function

Public Sub Destroy()

    If (m_uBI.bmiHeader.biSizeImage) Then
    
        '-- Reset BI structure
        With m_uBI
            Call ZeroMemory(.bmiHeader, 40)
            Call ZeroMemory(.bmiColors(0), 1024)
        End With
        
        '-- Delete DIB data
        Erase m_aBits()
    End If
End Sub

Public Sub Cls( _
           ByVal ColorIdx As Byte _
           )
    
  Dim i As Long, i2 As Long
  Dim j As Long, j2 As Long
  Dim lOffset As Long
  
    With m_uBI.bmiHeader
    
        If (.biSizeImage) Then
        
            i2 = .biWidth - 1
            j2 = -.biHeight - 1
            
            For j = 0 To j2
                For i = lOffset To lOffset + i2
                    m_aBits(i) = ColorIdx
                Next i
                lOffset = lOffset + m_lScanWidth
            Next j
        End If
    End With
End Sub

Public Function Paint( _
                ByVal hDC As Long, _
                Optional ByVal x As Long = 0, Optional ByVal y As Long = 0, _
                Optional ByVal ROP As RasterOpConstants = vbSrcCopy, _
                Optional ByVal StretchBltMode As eDIB08StretchBltModeConstants = [eColorOnColor] _
                ) As Long
    
    With m_uBI.bmiHeader
    
        Paint = Me.Stretch( _
                hDC, _
                x, y, _
                .biWidth, -.biHeight, _
                , , , , _
                ROP, _
                StretchBltMode _
                )
    End With
End Function

Public Function Stretch(ByVal hDC As Long, _
                ByVal x As Long, ByVal y As Long, _
                ByVal Width As Long, ByVal Height As Long, _
                Optional ByVal xSrc As Long, Optional ByVal ySrc As Long, _
                Optional ByVal SrcWidth As Long, Optional ByVal SrcHeight As Long, _
                Optional ByVal lROP As RasterOpConstants = vbSrcCopy, _
                Optional ByVal StretchBltMode As eDIB08StretchBltModeConstants = [eColorOnColor] _
                ) As Long
  
  Dim lOldMode As Long
    
    With m_uBI.bmiHeader
    
        If (.biSizeImage) Then
            
            If (SrcWidth = 0) Then
                SrcWidth = .biWidth
            End If
            If (SrcHeight = 0) Then
                SrcHeight = -.biHeight
            End If
            
            lOldMode = SetStretchBltMode(hDC, StretchBltMode)
            
            Stretch = StretchDIBits( _
                      hDC, _
                      x, y, _
                      Width, Height, _
                      xSrc, ySrc, _
                      SrcWidth, SrcHeight, _
                      m_aBits(0), m_uBI, DIB_RGB_COLORS, _
                      lROP _
                      )
            
            Call SetStretchBltMode(hDC, lOldMode)
        End If
    End With
End Function

Public Sub Reset()
    
    With m_uBI.bmiHeader
        If (.biSizeImage) Then
            Call ZeroMemory(m_aBits(0), .biSizeImage)
        End If
    End With
End Sub

Public Sub SetPalette( _
           Palette() As Byte _
           )
    
    With m_uBI
        If (.bmiHeader.biSizeImage) Then
            Call CopyMemory(.bmiColors(0), Palette(0), 1024)
        End If
    End With
    
' Note: Palette array must be initialized before.
End Sub

Public Sub GetPalette( _
           Palette() As Byte _
           )
    
    With m_uBI
        If (.bmiHeader.biSizeImage) Then
            Call CopyMemory(Palette(0), .bmiColors(0), 1024)
        End If
    End With
    
' Note: Palette array must be initialized before.
End Sub

Public Function CopyToClipboard( _
                ) As Long

  Dim hMem  As Long
  Dim lpMem As Long
  Dim r     As Long
    
    With m_uBI.bmiHeader
    
        If (.biSizeImage) Then
        
            If (OpenClipboard(0)) Then
                
                '-- Get mem. handle and its pointer
                hMem = GlobalAlloc(GMEM_MOVEABLE, LenB(m_uBI) + .biSizeImage)
                If (hMem <> 0) Then
                    lpMem = GlobalLock(hMem)
                    If (lpMem <> 0) Then
                        
                        '-- Copy header (bitmap header + palette)
                        .biHeight = -.biHeight                           ' fix height
                        Call CopyMemory(ByVal lpMem, m_uBI, LenB(m_uBI)) ' copy header
                        .biHeight = -.biHeight                           ' restore height
                            
                        '-- Copy image data (up-down)
                        lpMem = lpMem + LenB(m_uBI)
                        For r = -.biHeight - 1 To 0 Step -1
                            Call CopyMemory(ByVal lpMem, m_aBits(r * m_lScanWidth), m_lScanWidth)
                            lpMem = lpMem + m_lScanWidth
                        Next r
                        
                        '-- Unlock
                        Call GlobalUnlock(hMem)
                        
                        '-- Paste to Clipboard
                        Call EmptyClipboard
                        CopyToClipboard = SetClipboardData(CF_DIB, hMem)
                    End If
                End If
                Call CloseClipboard
            End If
        End If
    End With
End Function

'========================================================================================
' Properties
'========================================================================================

Public Property Get HasDIB() As Boolean
    HasDIB = (m_uBI.bmiHeader.biSizeImage > 0)
End Property

Public Property Get lpBits() As Long
    lpBits = VarPtr(m_aBits(0))
End Property

Public Property Get Width() As Long
    Width = m_uBI.bmiHeader.biWidth
End Property

Public Property Get Height() As Long
    Height = -m_uBI.bmiHeader.biHeight
End Property

Public Property Get BytesPerScanline() As Long
    BytesPerScanline = m_lScanWidth
End Property

Public Property Get Size() As Long
    Size = m_uBI.bmiHeader.biSizeImage
End Property

'========================================================================================
' Helpers
'========================================================================================

Public Function GetPixelIdx( _
                ByVal x As Long, _
                ByVal y As Long _
                ) As Byte
    
    With m_uBI.bmiHeader
        If (.biSizeImage) Then
            If (x >= 0 And y >= 0 And x < .biWidth And y < -.biHeight) Then
                GetPixelIdx = m_aBits(y * m_lScanWidth + x)
            End If
        End If
    End With
End Function

Public Sub SetPixelIdx( _
           ByVal x As Long, _
           ByVal y As Long, _
           ByVal Idx As Byte _
           )
    
    With m_uBI.bmiHeader
        If (.biSizeImage) Then
            If (x >= 0 And y >= 0 And x < .biWidth And y < -.biHeight) Then
                m_aBits(y * m_lScanWidth + x) = Idx
            End If
        End If
    End With
End Sub

Public Function Save( _
                ByVal Filename As String _
                ) As Boolean
    
  Dim uBFH  As BITMAPFILEHEADER
  Dim hFile As Integer
  
    With m_uBI.bmiHeader
    
        If (.biSizeImage) Then
    
            On Error Resume Next
            Call VBA.Kill(Filename)
            On Error GoTo 0
            
            On Error GoTo errH
           
            hFile = VBA.FreeFile()
            Open Filename For Binary Access Write As #hFile
            
            uBFH.bfType = BITMAPTYPE
            uBFH.bfOffBits = Len(uBFH) + Len(m_uBI)
            uBFH.bfSize = uBFH.bfOffBits + .biSizeImage
            
            Put #hFile, , uBFH      ' file header
            Put #hFile, , m_uBI     ' bitmap header
            Put #hFile, , m_aBits() ' bitmap data
                            
            Close #hFile
            
            '-- Success
            Save = True
        End If
    End With
    
errH:
    On Error GoTo 0
End Function

