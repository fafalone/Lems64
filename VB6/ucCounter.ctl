VERSION 5.00
Begin VB.UserControl ucCounter 
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   960
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   64
End
Attribute VB_Name = "ucCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================================
' User control:  ucCounter.ctl (0-99)
'
' Author:        Carles P.V.
' Dependencies:  Project
' Last revision: 2011.04.29
'========================================================================================

Option Explicit

'-- Private constants and variables:

Private Const CONTROL_WIDTH  As Long = 21
Private Const CONTROL_HEIGHT As Long = 11

'-- Property variables:

Private m_sCaption As String



'========================================================================================
' UserControl
'========================================================================================

Private Sub UserControl_Paint()
    
  Dim bIsSystem As Boolean
  Dim bCaption  As Boolean
  Dim c         As Long
  
    bIsSystem = (BackColor = vbButtonFace)
    bCaption = Val(m_sCaption) <> 0
    
    c = TranslateColor(BackColor)
    c = IIf(bCaption, _
        IIf(bIsSystem, vb3DHighlight, ShiftColor(c, &H40)), _
        IIf(bIsSystem, vb3DShadow, ShiftColor(c, -&H40)) _
        )
        
    Line (0, 0)-(CONTROL_WIDTH, CONTROL_HEIGHT), c, BF
    PSet (0, 0), BackColor
    PSet (CONTROL_WIDTH - 1, 0), BackColor
    PSet (0, CONTROL_HEIGHT - 1), BackColor
    PSet (CONTROL_WIDTH - 1, CONTROL_HEIGHT - 1), BackColor
    
    If (bCaption) Then
        CurrentX = (CONTROL_WIDTH - TextWidth(m_sCaption)) \ 2 + 1
        CurrentY = (CONTROL_HEIGHT - TextHeight(vbNullString)) \ 2 - 1
        Print m_sCaption;
    End If
End Sub

Private Sub UserControl_Resize()
    
    Call Cls
    Call UserControl.Size( _
         CONTROL_WIDTH * Screen.TwipsPerPixelX, _
         CONTROL_HEIGHT * Screen.TwipsPerPixelY _
         )
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    Call UserControl_Paint
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let Caption(ByVal New_Caption As String)
    If (m_sCaption <> New_Caption) Then
        m_sCaption = New_Caption
        Call UserControl_Paint
    End If
End Property
Public Property Get Caption() As String
    Caption = m_sCaption
End Property

Private Sub UserControl_InitProperties()
    UserControl.BackColor = vbButtonFace
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, vbButtonFace)
End Sub
