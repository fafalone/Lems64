Attribute VB_Name = "mHook"
Option Explicit

'-- API:

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const GWL_WNDPROC    As Long = -4
Private Const WM_ACTIVATE    As Long = &H6
Private Const WM_MOUSEWHEEL  As Long = &H20A

'-- Private variables:

Private m_lpPrevWndProcLems  As Long



'========================================================================================
' Methods
'========================================================================================

Public Sub InitializeHook(oForm As Form)

    m_lpPrevWndProcLems = SetWindowLong(fLems.hwnd, GWL_WNDPROC, AddressOf pvWindowProcLems)
End Sub

'========================================================================================
' Private
'========================================================================================

Private Function pvWindowProcLems( _
                 ByVal hwnd As Long, _
                 ByVal uMsg As Long, _
                 ByVal wParam As Long, _
                 ByVal lParam As Long _
                 ) As Long
    
    Select Case uMsg
    
        Case WM_ACTIVATE
            
            Call mTiming.SetAppActive(CBool(wParam And &HFFFF&))
        
        Case WM_MOUSEWHEEL
        
            If (wParam > 0) Then
                Call SendKeysB("Z")
              Else
                Call SendKeysB("A")
            End If
    End Select
    
    pvWindowProcLems = CallWindowProc(m_lpPrevWndProcLems, hwnd, uMsg, wParam, lParam)
End Function
Public Sub SendKeysB(ByRef Text As String, Optional ByRef Wait As Boolean)
    Static wsh As Object
    If wsh Is Nothing Then
        Set wsh = CreateObject("WScript.Shell")
    End If
    wsh.SendKeys Text, Wait
End Sub
'Public Sub SendKeysB(ByRef Text As String, Optional ByRef Wait As Boolean)
'    CreateObject("WScript.Shell").SendKeys Text, Wait
'End Sub
