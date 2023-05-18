VERSION 5.00
Begin VB.UserControl ucProgress 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3000
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   20
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   200
End
Attribute VB_Name = "ucProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' User control:  ucProgress.ctl (modified)
' Author:        Carles P.V.
' Dependencies:  None
' Last revision: 2003.05.25
'================================================

Option Explicit

'-- API:

Private Type RECT
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Declare Function TranslateColor Lib "olepro32" Alias "OleTranslateColor" (ByVal Clr As OLE_COLOR, ByVal Palette As Long, Col As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_FRAMECHANGED  As Long = &H20
Private Const SWP_NOACTIVATE    As Long = &H10
Private Const SWP_NOMOVE        As Long = &H2
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOREDRAW      As Long = &H8
Private Const SWP_NOSIZE        As Long = &H1
Private Const SWP_NOZORDER      As Long = &H4
Private Const SWP_SHOWWINDOW    As Long = &H40

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_STYLE         As Long = (-16)
Private Const WS_THICKFRAME     As Long = &H40000
Private Const WS_BORDER         As Long = &H800000
Private Const GWL_EXSTYLE       As Long = (-20)
Private Const WS_EX_WINDOWEDGE  As Long = &H100&
Private Const WS_EX_CLIENTEDGE  As Long = &H200&
Private Const WS_EX_STATICEDGE  As Long = &H20000

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Private Type LOGFONT
   lfHeight            As Long
   lfWidth             As Long
   lfEscapement        As Long
   lfOrientation       As Long
   lfWeight            As Long
   lfItalic            As Byte
   lfUnderline         As Byte
   lfStrikeOut         As Byte
   lfCharSet           As Byte
   lfOutPrecision      As Byte
   lfClipPrecision     As Byte
   lfQuality           As Byte
   lfPitchAndFamily    As Byte
   lfFaceName(1 To 32) As Byte
End Type

Private Const LOGPIXELSY             As Long = 90
Private Const FW_NORMAL              As Long = 400
Private Const FW_BOLD                As Long = 700
Private Const FF_DONTCARE            As Long = 0
Private Const DEFAULT_QUALITY        As Long = 0
Private Const DEFAULT_PITCH          As Long = 0
Private Const DEFAULT_CHARSET        As Long = 1
Private Const ANTIALIASED_QUALITY    As Long = 2
Private Const NONANTIALIASED_QUALITY As Long = 3

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_CENTER     As Long = &H1
Private Const DT_NOCLIP     As Long = &H100
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_VCENTER    As Long = &H4

Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Const TRANSPARENT As Long = 1

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)

'-- Public enums.:

Public Enum eBorderStyleConstantsEx
    [eNone] = 0
    [eThin]
    [eThick]
End Enum

'-- Default property values:

Private Const m_def_BorderStyle = [eThick]
Private Const m_def_BackColor = vbButtonFace
Private Const m_def_ForeColor = vbHighlight
Private Const m_def_Max = 100

'-- Property variables:

Private m_eBorderStyle As eBorderStyleConstantsEx
Private m_oleBackColor As OLE_COLOR
Private m_oleForeColor As OLE_COLOR
Private m_lMax         As Long
Private m_sCaption     As String

'-- Private variables:

Private m_lValue       As Long
Private m_uControlRect As RECT
Private m_uForeRect    As RECT
Private m_uBackRect    As RECT
Private m_lPos         As Long
Private m_lLastPos     As Long
Private m_hForeBrush   As Long
Private m_hBackBrush   As Long

'-- Event declarations:

Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)



'========================================================================================
' UserControl
'========================================================================================

Private Sub UserControl_Initialize()
    If (m_lMax = 0) Then m_lMax = 1
End Sub

Private Sub UserControl_Terminate()
    If (m_hForeBrush <> 0) Then Call DeleteObject(m_hForeBrush)
    If (m_hBackBrush <> 0) Then Call DeleteObject(m_hBackBrush)
End Sub

'//

Private Sub UserControl_Resize()
    Call pvGetProgress
    Call pvCalcRects
    Call UserControl_Paint
End Sub

Private Sub UserControl_Paint()

  Dim hTmpDC     As Long
  Dim hTmpBmp    As Long
  Dim hOldTmpBmp As Long
  Dim hFont      As Long
  Dim hOldFont   As Long
  Dim hRgn       As Long
  Dim lClr       As Long

    '-- Buffer
    hTmpDC = CreateCompatibleDC(hDC)
    hTmpBmp = CreateCompatibleBitmap(hDC, ScaleWidth, ScaleHeight)
    hOldTmpBmp = SelectObject(hTmpDC, hTmpBmp)
    
    '-- Font and text mode
    hFont = CreateFontIndirect(pvOLEFontToLogFont(Font, hTmpDC))
    hOldFont = SelectObject(hTmpDC, hFont)
    Call SetBkMode(hTmpDC, TRANSPARENT)
    
    '-- Background
    Call FillRect(hTmpDC, m_uForeRect, m_hForeBrush)
    Call FillRect(hTmpDC, m_uBackRect, m_hBackBrush)
    
    '-- 0 -> %
    Call TranslateColor(m_oleBackColor, 0, lClr)
    Call SetTextColor(hTmpDC, lClr)
    hRgn = CreateRectRgn(m_uForeRect.x1, m_uForeRect.y1, m_uForeRect.x2, m_uForeRect.y2)
    Call SelectClipRgn(hTmpDC, hRgn)
    Call DrawText(hTmpDC, m_sCaption, -1, m_uControlRect, DT_SINGLELINE Or DT_CENTER Or DT_NOCLIP Or DT_VCENTER)
    Call DeleteObject(hRgn)
    
    '-- % -> 100
    Call TranslateColor(m_oleForeColor, 0, lClr)
    Call SetTextColor(hTmpDC, lClr)
    hRgn = CreateRectRgn(m_uBackRect.x1, m_uBackRect.y1, m_uBackRect.x2, m_uBackRect.y2)
    Call SelectClipRgn(hTmpDC, hRgn)
    Call DrawText(hTmpDC, m_sCaption, -1, m_uControlRect, DT_SINGLELINE Or DT_CENTER Or DT_NOCLIP Or DT_VCENTER)
    Call DeleteObject(hRgn)
    
    '-- Paint from buffer
    Call BitBlt(hDC, 0, 0, ScaleWidth, ScaleHeight, hTmpDC, 0, 0, vbSrcCopy)

    '-- Clean up
    Call SelectObject(hTmpDC, hOldFont)
    Call DeleteObject(hFont)
    Call SelectObject(hTmpDC, hOldTmpBmp)
    Call DeleteObject(hTmpBmp)
    Call DeleteDC(hTmpDC)
End Sub

'========================================================================================
' Events
'========================================================================================

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Get BorderStyle() As eBorderStyleConstants
    BorderStyle = m_eBorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As eBorderStyleConstants)
    m_eBorderStyle = New_BorderStyle
    Call pvSetBorder
    Call pvGetProgress
    Call pvCalcRects
    Call UserControl_Paint
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_oleBackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_oleBackColor = New_BackColor
    Call pvCreateBackBrush
    Call UserControl_Paint
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_oleForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_oleForeColor = New_ForeColor
    Call pvCreateForeBrush
    Call UserControl_Paint
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
End Property

Public Property Get Max() As Long
    Max = m_lMax
End Property

Public Property Let Max(ByVal New_Max As Long)
    If (New_Max < 1) Then New_Max = 1
    m_lMax = New_Max
    Call UserControl_Paint
End Property

Public Property Get Caption() As String
    Caption = m_sCaption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_sCaption = New_Caption
    Call UserControl_Paint
End Property

Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "400"
    Value = m_lValue
End Property

Public Property Let Value(ByVal New_Value As Long)

    m_lValue = New_Value
    
    Call pvGetProgress
    If (m_lPos <> m_lLastPos) Then
        m_lLastPos = m_lPos
        Call pvCalcRects
        Call UserControl_Paint
    End If
End Property

'//

Private Sub UserControl_InitProperties()

    m_eBorderStyle = m_def_BorderStyle
    m_oleBackColor = m_def_BackColor
    m_oleForeColor = m_def_ForeColor
    m_lMax = m_def_Max
    
    Call pvSetBorder
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    With PropBag
        m_eBorderStyle = .ReadProperty("BorderStyle", m_def_BorderStyle)
        m_oleBackColor = .ReadProperty("BackColor", m_def_BackColor)
        m_oleForeColor = .ReadProperty("ForeColor", m_def_ForeColor)
        m_lMax = .ReadProperty("Max", m_def_Max)
        UserControl.Enabled = .ReadProperty("Enabled", True)
    End With

    Call pvSetBorder
    Call pvCalcRects
    Call pvCreateForeBrush
    Call pvCreateBackBrush
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty("BorderStyle", m_eBorderStyle, m_def_BorderStyle)
        Call .WriteProperty("BackColor", m_oleBackColor, m_def_BackColor)
        Call .WriteProperty("ForeColor", m_oleForeColor, m_def_ForeColor)
        Call .WriteProperty("Max", m_lMax, m_def_Max)
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
    End With
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub pvCreateForeBrush()
    
  Dim lClr As Long
    
    If (m_hForeBrush <> 0) Then
        Call DeleteObject(m_hForeBrush)
        m_hForeBrush = 0
    End If
    Call TranslateColor(ForeColor, 0, lClr)
    m_hForeBrush = CreateSolidBrush(lClr)
End Sub

Private Sub pvCreateBackBrush()

  Dim lClr As Long
  
    If (m_hBackBrush <> 0) Then
        Call DeleteObject(m_hBackBrush)
        m_hBackBrush = 0
    End If
    Call TranslateColor(BackColor, 0, lClr)
    m_hBackBrush = CreateSolidBrush(lClr)
End Sub

Private Sub pvGetProgress()
    
    m_lPos = (m_lValue * ScaleWidth) \ m_lMax
End Sub

Private Sub pvCalcRects()
    
    Call SetRect(m_uControlRect, 0, 0, ScaleWidth, ScaleHeight)
    Call SetRect(m_uForeRect, 0, 0, m_lPos, ScaleHeight)
    Call SetRect(m_uBackRect, m_lPos, 0, ScaleWidth, ScaleHeight)
End Sub

Private Sub pvSetBorder()

    Select Case m_eBorderStyle
        Case [eNone]
            Call pvSetWinStyle(GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME)
            Call pvSetWinStyle(GWL_EXSTYLE, 0, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE)
        Case [eThin]
            Call pvSetWinStyle(GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME)
            Call pvSetWinStyle(GWL_EXSTYLE, WS_EX_STATICEDGE, WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE)
        Case [eThick]
            Call pvSetWinStyle(GWL_STYLE, 0, WS_BORDER Or WS_THICKFRAME)
            Call pvSetWinStyle(GWL_EXSTYLE, WS_EX_CLIENTEDGE, WS_EX_STATICEDGE Or WS_EX_WINDOWEDGE)
    End Select
End Sub

Private Sub pvSetWinStyle(ByVal lType As Long, ByVal lStyle As Long, ByVal lStyleNot As Long)

  Dim lS As Long
    
    lS = GetWindowLong(hWnd, lType)
    lS = (lS And Not lStyleNot) Or lStyle
    Call SetWindowLong(hWnd, lType, lS)
    Call SetWindowPos(hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED)
End Sub

Private Function pvOLEFontToLogFont(oFont As StdFont, ByVal hDCRef As Long) As LOGFONT

    With pvOLEFontToLogFont
        
        Call CopyMemory(.lfFaceName(1), ByVal oFont.Name, Len(oFont.Name) + 1)
        .lfCharSet = oFont.Charset
        .lfItalic = -oFont.Italic
        .lfUnderline = -oFont.Underline
        .lfStrikeOut = -oFont.Strikethrough
        .lfWeight = oFont.Weight
        .lfHeight = -(oFont.Size * GetDeviceCaps(hDCRef, LOGPIXELSY) / 72)
        .lfQuality = ANTIALIASED_QUALITY
    End With
End Function
