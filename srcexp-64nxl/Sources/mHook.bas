Attribute VB_Name = "mHook"
Option Explicit

'-- API:

#If Win64 Then
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As GWL_INDEX) As LongPtr
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As GWL_INDEX, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As GWL_INDEX) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As GWL_INDEX, ByVal dwNewLong As Long) As Long
#End If
Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Private Enum GWL_INDEX
    GWL_WNDPROC = (-4)
    GWL_HINSTANCE = (-6)
    GWL_HWNDPARENT = (-8)
    GWL_ID = (-12)
    GWL_STYLE = (-16)
    GWL_EXSTYLE = (-20)
    GWL_USERDATA = (-21)
End Enum

Private Const WM_ACTIVATE    As Long = &H6
Private Const WM_MOUSEWHEEL  As Long = &H20A

'-- Private variables:

Private m_lpPrevWndProcLems  As LongPtr



'========================================================================================
' Methods
'========================================================================================

Public Sub InitializeHook(oForm As Form)

    m_lpPrevWndProcLems = SetWindowLong(fLems.hWnd, GWL_WNDPROC, AddressOf pvWindowProcLems)
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub SendKeysB(ByRef Text As String, Optional ByRef Wait As Boolean)
    Static wsh As Object
    If wsh Is Nothing Then
        Set wsh = CreateObject("WScript.Shell")
    End If
    wsh.SendKeys Text, Wait
End Sub

Private Function pvWindowProcLems( _
                 ByVal hwnd As LongPtr, _
                 ByVal uMsg As Long, _
                 ByVal wParam As LongPtr, _
                 ByVal lParam As LongPtr _
                 ) As LongPtr
    
    Select Case uMsg
    
        Case WM_ACTIVATE
            
            Call mTiming.SetAppActive(CBool(wParam And &HFFFF&))
        
        Case WM_MOUSEWHEEL
            Dim noExt As Long
            CopyMemory noExt, wParam, 4
            If (noExt > 0) Then
                'Call VBA.SendKeys("Z")
                SendKeysB "Z"
              Else
                'Call VBA.SendKeys("A")
                SendKeysB "A"
            End If
    End Select
    
    pvWindowProcLems = CallWindowProc(m_lpPrevWndProcLems, hwnd, uMsg, wParam, lParam)
End Function
