Attribute VB_Name = "mFullScreen"
Option Explicit

'-- API:

Private Const CCDEVICENAME As Long = 32
Private Const CCFORMNAME   As Long = 32

Private Type Point
    x As Long
    y As Long
End Type

Private Type DEVMODE
    dmDeviceName(CCDEVICENAME - 1) As Byte
    dmSpecVersion                   As Integer
    dmDriverVersion                 As Integer
    dmSize                          As Integer
    dmDriverExtra                   As Integer
    dmFields                        As Long
    dmPosition                      As Point
    dmDisplayOrientation            As Long
    dmDisplayFixedOutput            As Long
    dmColor                         As Integer
    dmDuplex                        As Integer
    dmYResolution                   As Integer
    dmTTOption                      As Integer
    dmCollate                       As Integer
    dmFormName(0 To CCFORMNAME - 1) As Byte
    dmLogPixels                     As Integer
    dmBitsPerPel                    As Long
    dmPelsWidth                     As Long
    dmPelsHeight                    As Long
    dmDisplayFlags                  As Long
    dmDisplayFrequency              As Long
End Type
Private Const CDS_TEST               As Long = 2
Private Const CDS_FULLSCREEN         As Long = 4
Private Const DM_PELSWIDTH           As Long = &H80000
Private Const DM_PELSHEIGHT          As Long = &H100000
Private Const DISP_CHANGE_SUCCESSFUL As Long = 0

Private Declare PtrSafe Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As LongPtr, ByVal iModeNum As Long, lpDevMode As Any) As Long
Private Declare PtrSafe Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long

Public Enum WindowZOrderDefaults
    HWND_DESKTOP = 0&
    HWND_TOP = 0&
    HWND_BOTTOM = 1&
    HWND_TOPMOST = -1
    HWND_NOTOPMOST = -2
End Enum
Public Enum SWP_Flags
    SWP_NOSIZE = &H1
    SWP_NOMOVE = &H2
    SWP_NOZORDER = &H4
    SWP_NOREDRAW = &H8
    SWP_NOACTIVATE = &H10
    SWP_FRAMECHANGED = &H20
    SWP_DRAWFRAME = &H20
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_NOCOPYBITS = &H100
    SWP_NOREPOSITION = &H200
    SWP_NOSENDCHANGING = &H400
    
    SWP_DEFERERASE = &H2000
    SWP_ASYNCWINDOWPOS = &H4000
End Enum

Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As SWP_Flags) As Long

Private Enum GWL_INDEX
    GWL_WNDPROC = (-4)
    GWL_HINSTANCE = (-6)
    GWL_HWNDPARENT = (-8)
    GWL_ID = (-12)
    GWL_STYLE = (-16)
    GWL_EXSTYLE = (-20)
    GWL_USERDATA = (-21)
End Enum
Private Const WS_VISIBLE As Long = &H10000000

#If Win64 Then
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As GWL_INDEX) As LongPtr
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As GWL_INDEX, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As GWL_INDEX) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As GWL_INDEX, ByVal dwNewLong As Long) As Long
#End If
'-- Private constants and variables:

Private Const FS_HEIGHT  As Long = 480
Private m_oForm          As Form
Private m_lStyle         As Long
Private m_bIsFullScreen  As Boolean
Private m_lBestModeWidth As Long
Private m_lLeft          As Long
Private m_lTop           As Long
Private m_lWidth         As Long
Private m_lHeight        As Long



'========================================================================================
' Methods
'========================================================================================

Public Sub InitializeFullScreen(oForm As Form)
    
    '-- Reference full-screen window
    Set m_oForm = oForm
    
    '-- Get best mode width assuming <height = 480>
    m_lBestModeWidth = pvGetBestWidth()
End Sub

Public Function ToggleFullScreen( _
                ) As Boolean
 
  Dim uDM  As DEVMODE
  Dim lRet As Long
  
    '-- No full screen mode available for this game layout
    If (m_lBestModeWidth = 0) Then
        Call VBA.MsgBox( _
             "Unable to switch to full-screen mode.", _
             vbExclamation _
             )
        Exit Function
    End If
    
    '-- Initialize DEVMODE structure
    '   (only width and height will be passed)
    uDM.dmSize = LenB(uDM)
    uDM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
    
    '-- Full/default mode...
    If (m_bIsFullScreen = False) Then
        
        '-- Set new <width x height> as specified
        uDM.dmPelsWidth = m_lBestModeWidth
        uDM.dmPelsHeight = FS_HEIGHT
        
        '-- First, test mode
        lRet = ChangeDisplaySettings(uDM, CDS_FULLSCREEN Or CDS_TEST)
        If (lRet <> DISP_CHANGE_SUCCESSFUL) Then
            Call VBA.MsgBox( _
                 "An error ocurred while trying to change to full-screen mode", _
                 vbExclamation _
                 )
          Else
            
            '-- Store window's current style bits
            '   as well as current position.
            m_lStyle = CLng(GetWindowLong(m_oForm.hWnd, GWL_STYLE))
            m_lLeft = m_oForm.Left \ Screen.TwipsPerPixelX
            m_lTop = m_oForm.Top \ Screen.TwipsPerPixelY
            m_lWidth = m_oForm.Width \ Screen.TwipsPerPixelX
            m_lHeight = m_oForm.Height \ Screen.TwipsPerPixelY
            
            '-- No border, no caption, ... nothing.
            '   Remove WS_VISIBLE flag: SetWindowPos will
            '   make it visible after changing display settings
            Call SetWindowLong(m_oForm.hWnd, GWL_STYLE, 0)
            
            '-- Change width and height (passed uDM structure).
            '   CDS_FULLSCREEN tells system we are changing to/from
            '   full-screen mode, so not to update other windows
            '   position
            If (ChangeDisplaySettings(uDM, CDS_FULLSCREEN) = DISP_CHANGE_SUCCESSFUL) Then
            
                '-- Full-screen
                m_bIsFullScreen = True
                
                '-- Relocate controls
                Call pvOffsetControls(1)
                
                '-- Full-screen now: move window to (0,0) and make it
                '   visible again. Force redraw in order to update
                '   window's new style and place it as top-most
                Call SetWindowPos( _
                     m_oForm.hWnd, _
                     0, _
                     0, 0, m_lBestModeWidth, FS_HEIGHT, _
                     SWP_SHOWWINDOW Or SWP_FRAMECHANGED _
                     )
                
                '-- OK, success
                ToggleFullScreen = True
            End If
        End If

      Else
        
        '-- Restore window's defaults (except visibility)
        Call SetWindowLong(m_oForm.hWnd, GWL_STYLE, m_lStyle And Not WS_VISIBLE)
        
        '-- Passing NULL changes mode back!
        If (ChangeDisplaySettings(ByVal 0, CDS_FULLSCREEN) = DISP_CHANGE_SUCCESSFUL) Then
        
            '-- No full-screen
            m_bIsFullScreen = False
            
            '-- Relocate controls
            Call pvOffsetControls(-1)
            
            '-- Finally, restore position and visibility
            Call SetWindowPos( _
                 m_oForm.hWnd, _
                 HWND_TOP, _
                 m_lLeft, m_lTop, m_lWidth, m_lHeight, _
                 SWP_SHOWWINDOW Or SWP_FRAMECHANGED _
                 )
        
            '-- OK, suposed success
            ToggleFullScreen = True
        End If
    End If
End Function

'========================================================================================
' Properties
'========================================================================================

Public Property Get IsFullScreen( _
                    ) As Boolean

    IsFullScreen = m_bIsFullScreen
End Property

'========================================================================================
' Private
'========================================================================================

Private Function pvGetBestWidth() As Long

  Dim uDM   As DEVMODE
  Dim lMode As Long
  Dim lMax  As Long
   
    '-- Set DEVMODE flags
    uDM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
    uDM.dmSize = LenB(uDM)
      
    '-- Get [w]x480 where [w] is maximum available width
    lMax = 0
    lMode = 0
    Do While EnumDisplaySettings(0&, lMode, uDM) > 0
        If (uDM.dmPelsHeight = FS_HEIGHT) Then
            If (uDM.dmPelsWidth > lMax) Then
                lMax = uDM.dmPelsWidth
            End If
        End If
        lMode = lMode + 1
    Loop
   
    pvGetBestWidth = lMax
End Function

Private Sub pvOffsetControls( _
            ByVal Sign As Integer _
            )
    
  Dim o As Control
    
    On Error Resume Next
    For Each o In m_oForm.Controls
        o.Left = o.Left + Sign * (m_lBestModeWidth - 640) \ 2
    Next o
    On Error GoTo 0
End Sub
