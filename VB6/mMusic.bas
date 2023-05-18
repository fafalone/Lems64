Attribute VB_Name = "mMusic"
Option Explicit

'-- API:

Private Declare Function mciSendString Lib "winmm" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const GWL_WNDPROC           As Long = (-4)
Private Const MM_MCINOTIFY          As Long = &H3B9
Private Const MCI_NOTIFY_SUCCESSFUL As Long = &H1

'-- Private variables:

Private m_bMusic        As Boolean
Private m_bPlaying      As Boolean
Private m_sFile         As String
Private m_lpPrevWndProc As Long
Private m_hWnd          As Long



'========================================================================================
' Methods
'========================================================================================

Public Sub InitializeMusic()
    
    '-- New generator seed
    Call VBA.Randomize(Timer)
    
    '-- Default music state
    m_bMusic = False
    
    '-- Hook for notifications (midi loop)
    m_hWnd = fLems.hWnd
    m_lpPrevWndProc = SetWindowLong(m_hWnd, GWL_WNDPROC, AddressOf pvMidiNotify)
End Sub

Public Sub TerminateMusic()

    '-- Unhook
    Call SetWindowLong(m_hWnd, GWL_WNDPROC, m_lpPrevWndProc)
End Sub

Public Sub PlayMidi( _
           Optional ByVal RandomTheme As Boolean = True _
           )

  Dim sFile      As String
  Dim sShortFile As String * 67
  Dim lRet       As Long
    
    If (RandomTheme) Then
        
        '-- Get random theme
        sFile = AppPath & "MUSIC\Theme" & Format$(Rnd * 14, "00") & ".mid"
        lRet = GetShortPathName(sFile, sShortFile, Len(sShortFile))
        m_sFile = Left$(sShortFile, lRet)
    End If
    
    If (m_bMusic) Then
        
        '-- Start playing
        If (Not m_bPlaying) Then
            Screen.MousePointer = vbHourglass
            Call mciSendString("open " & m_sFile & " type sequencer", vbNullString, 0, m_hWnd)
            Screen.MousePointer = vbDefault
        End If
        Call mciSendString("play " & m_sFile & " from 0 notify", vbNullString, 0, m_hWnd)
        
        m_bPlaying = True
    End If
End Sub

Public Sub CloseMidi()
    
    '-- Close midi
    Call mciSendString("close " & m_sFile, vbNullString, 0, 0)
    
    m_bPlaying = False
End Sub

Public Sub SetMusicState( _
           ByVal Enable As Boolean _
           )
    
    '-- Enable/disable sound effects
    m_bMusic = Enable
    
    '-- Stop/Continue?
    If (fLems.Tag = [eModePlaying]) Then
        If (m_bMusic) Then
            If (m_bPlaying) Then
                Call PlayMidi(RandomTheme:=False)
              Else
                Call PlayMidi(RandomTheme:=True)
            End If
          Else
            If (m_bPlaying) Then
                Call mciSendString("stop " & m_sFile, vbNullString, 0, 0)
            End If
        End If
    End If
End Sub

'========================================================================================
' Private
'========================================================================================

Private Function pvMidiNotify( _
                 ByVal hWnd As Long, _
                 ByVal uMsg As Long, _
                 ByVal wParam As Long, _
                 ByVal lParam As Long _
                 ) As Long
    
    If (uMsg = MM_MCINOTIFY) Then
        If (wParam = MCI_NOTIFY_SUCCESSFUL) Then
            Call PlayMidi(RandomTheme:=False)
        End If
    End If
    
    pvMidiNotify = CallWindowProc(m_lpPrevWndProc, hWnd, uMsg, wParam, lParam)
End Function
