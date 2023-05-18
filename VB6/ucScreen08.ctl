VERSION 5.00
Begin VB.UserControl ucScreen08 
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MousePointer    =   99  'Custom
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   100
End
Attribute VB_Name = "ucScreen08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================================
' User control:  ucScreen08.ctl (simplified)
'
' Author:        Carles P.V.
' Dependencies:  cDIB08.cls, mLemsRenderer.bas
' First release: 2006.11.17
' Last revision: 2011.05.04
'========================================================================================

Option Explicit

'-- Public enums.:

Public Enum eBorderStyleConstants
    [eNone] = 0
    [eFixedSingle]
End Enum

'-- Private variables:

Private m_oDIBActual As cDIB08 ' DIB section actual size
Private m_oDIBScaled As cDIB08 ' DIB section scaled size
Private m_lxOffset   As Long    'run-time only
Private m_lyOffset   As Long    'run-time only
Private m_lZoom      As Long    'run-time only

'-- Event declarations:

Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)



'========================================================================================
' UserControl
'========================================================================================

Private Sub UserControl_Initialize()

    '-- Initialize DIB
    Set m_oDIBActual = New cDIB08
    Set m_oDIBScaled = New cDIB08
End Sub

Private Sub UserControl_Terminate()

    '-- Destroy DIB
    Set m_oDIBActual = Nothing
    Set m_oDIBScaled = Nothing
End Sub

Private Sub UserControl_Paint()

    '-- Refresh Canvas
    Call pvRefresh
End Sub

'========================================================================================
' Events + Scrolling
'========================================================================================

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (m_lZoom) Then
        RaiseEvent MouseDown(Button, Shift, x \ m_lZoom, y \ m_lZoom)
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (m_lZoom) Then
        RaiseEvent MouseMove(Button, Shift, x \ m_lZoom, y \ m_lZoom)
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (m_lZoom) Then
        RaiseEvent MouseUp(Button, Shift, x \ m_lZoom, y \ m_lZoom)
    End If
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Sub Initialize( _
           ByVal Width As Long, _
           ByVal Height As Long, _
           Optional ByVal Zoom As Long = 1 _
           )
    
    '-- DIB actual size
    Call m_oDIBActual.Create(Width, Height)
    
    '-- Create scaled DIB if necessary
    m_lZoom = Zoom
    If (m_lZoom > 1) Then
        Call m_oDIBScaled.Create(m_lZoom * Width, m_lZoom * Height)
    End If
End Sub

Public Sub InitializeFromFile( _
           ByVal Filename As String, _
           Optional ByVal Zoom As Long = 1 _
           )
    
    '-- DIB actual size
    If (m_oDIBActual.CreateFromBitmapFile(Filename)) Then
    
        '-- Create scaled DIB if necessary
        m_lZoom = Zoom
        If (m_lZoom > 1) Then
            Call m_oDIBScaled.Create(m_lZoom * m_oDIBActual.Width, m_lZoom * m_oDIBActual.Height)
        End If
    End If
End Sub

Public Sub Destroy()
    
    Call m_oDIBActual.Destroy
    Call m_oDIBScaled.Destroy
End Sub

Public Sub UpdatePalette( _
           ByRef Palette() As Byte _
           )
    
    Call m_oDIBActual.SetPalette(Palette())
    Call m_oDIBScaled.SetPalette(Palette())
End Sub

Public Sub Refresh()
    Call pvRefresh
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Get DIB() As cDIB08
    Set DIB = m_oDIBActual
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Get BorderStyle() As eBorderStyleConstants
    BorderStyle = UserControl.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As eBorderStyleConstants)
    UserControl.BorderStyle() = New_BorderStyle
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled = New_Enabled
End Property
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Get UserIcon() As StdPicture
Attribute UserIcon.VB_MemberFlags = "400"
    Set UserIcon = UserControl.MouseIcon
End Property
Public Property Set UserIcon(ByVal New_MouseIcon As StdPicture)
    Set UserControl.MouseIcon = New_MouseIcon
    Call pvUpdatePointer
End Property

Public Property Let xOffset(ByVal New_xOffset As Long)
    m_lxOffset = New_xOffset
End Property
Public Property Get xOffset() As Long
    xOffset = m_lxOffset
End Property

Public Property Let yOffset(ByVal New_yOffset As Long)
    m_lyOffset = New_yOffset
End Property
Public Property Get yOffset() As Long
    yOffset = m_lyOffset
End Property

Public Property Get ScaleWidth() As Long
Attribute ScaleWidth.VB_MemberFlags = "400"
    ScaleWidth = UserControl.ScaleWidth
End Property
Public Property Get ScaleHeight() As Long
Attribute ScaleHeight.VB_MemberFlags = "400"
    ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_MemberFlags = "400"
    hWnd = UserControl.hWnd
End Property

Private Sub UserControl_InitProperties()
    UserControl.BorderStyle = [eNone]
    UserControl.BackColor = vbApplicationWorkspace
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", [eNone])
    UserControl.BackColor = PropBag.ReadProperty("BackColor", vbApplicationWorkspace)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, [eNone])
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, vbApplicationWorkspace)
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub pvRefresh()
  
    If (m_oDIBActual.HasDIB) Then
        '-- Paint
        If (m_oDIBScaled.HasDIB) Then
            Call FXStretch(m_oDIBScaled, m_oDIBActual)
            Call m_oDIBScaled.Paint(UserControl.hDC, m_lxOffset, m_lyOffset)
          Else
            Call m_oDIBActual.Paint(UserControl.hDC, m_lxOffset, m_lyOffset)
        End If
      Else
        '-- Erase background
        Call Cls
    End If
End Sub

Private Sub pvUpdatePointer()

    If (Not UserControl.MouseIcon Is Nothing) Then
         UserControl.MousePointer = vbCustom
      Else
         UserControl.MousePointer = vbDefault
    End If
End Sub
