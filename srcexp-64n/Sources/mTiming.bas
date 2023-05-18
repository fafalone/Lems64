Attribute VB_Name = "mTiming"
Option Explicit

'-- API:

Private Declare PtrSafe Function timeGetTime Lib "winmm" () As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'-- Private variables:

Private m_bAppActive            As Boolean
Private m_bTimerActivated       As Boolean
Private m_bTimerPaused          As Boolean
Private m_bArmageddonActivated  As Boolean
Private m_bExit                 As Boolean

Private m_lcTimer               As Long
Private m_lcNewLem              As Long
Private m_lcArmageddonLem       As Long

'-- Public variables and constants:

Public Const TMR_DT_FRAME       As Long = 75 ' ~ 13.3 fps (supposed)
Public Const TMR_TIMESEC        As Long = 18 ' ~ 1.35 s at 13.3 fps
Public Const RELEASE_RATE_MAX   As Long = 99

Public g_lcPlayingTime          As Long
Public g_lReleaseRate           As Long
Public g_lReleaseRateMin        As Long


Public Function UnsignedSub(ByVal Start As Long, ByVal Decr As Long) As Long
UnsignedSub = ((Start And &H7FFFFFFF) - (Decr And &H7FFFFFFF)) Xor ((Start Xor Decr) And &H80000000)
End Function


'========================================================================================
' Methods
'========================================================================================

Public Sub SetAppActive( _
           ByVal Active As Boolean _
           )
    
    '-- This application is active
    m_bAppActive = Active
End Sub

Public Sub StartTimer()
    
  Dim t As Long
  
    If (m_bTimerActivated = False) Then
        
        '-- Frame timer activated
        m_bTimerActivated = True
        
        '-- Start it
        m_bExit = False
        m_lcTimer = 0
        
        Do
            '-- Frame dt
            If Abs((UnsignedSub(timeGetTime(), t))) >= TMR_DT_FRAME Then '*
                t = timeGetTime()
                Call pvTimerTick
            End If
            
            '-- Keep a low CPU usage!!!
            Call Sleep(1)
            
            '-- Allow multi-tasking
            Call VBA.DoEvents
            
        Loop Until m_bExit
        
        '-- Frame timer and Armageddon deactivated
        m_bTimerActivated = False
        m_bArmageddonActivated = False
    End If
    
'* Modify only TMR_DT_FRAME value in order to
'  speed up/down fps
End Sub

Public Sub StopTimer()
    
    If (m_bTimerActivated) Then
        '-- Stop it
        m_bExit = True
        m_bTimerPaused = False
    End If
End Sub

Public Sub PauseTimer( _
           ByVal bPause As Boolean _
           )
    
    '-- Frame timer paused
    m_bTimerPaused = bPause
End Sub

Public Sub StartArmageddon()
    
    '-- Activate Armageddon
    If (GetGameStage = [eStagePlaying]) Then
        m_bArmageddonActivated = True
        m_lcArmageddonLem = 0
    End If
End Sub

Public Function IsTimerPaused( _
                ) As Boolean
    
    '-- Is paused?
    IsTimerPaused = m_bTimerPaused
End Function

Public Function IsArmageddonActivated( _
                ) As Boolean
    
    '-- Is activated?
    IsArmageddonActivated = m_bArmageddonActivated
End Function

Public Function IsFastForward( _
                ) As Boolean
    
    '-- Going fast-forward?
    IsFastForward = Not (GetAsyncKeyState(vbKeyF) = 0 And _
                         fLems.ucToolbar.IsButtonPressed(12) = False _
                         ) _
                    And m_bAppActive
End Function

Public Function GetMinSecString( _
                ByVal Seconds As Long _
                ) As String
    
    '-- Return "Time mm:ss"
    GetMinSecString = "Time " & _
                      Format$(Seconds \ 60, "0") & _
                      ":" & _
                      Format$(Seconds Mod 60, "00")
End Function

'========================================================================================
' Private
'========================================================================================

Private Sub pvTimerTick()
  
    If (m_bTimerPaused = False) Then
        
        '-- Hit-test and render frame
        If (GethWndFromPoint <> fLems.picFullScreenOff.hWnd) Then
            fLems.ucInfo.PanelText(2) = HitTest()
        End If
        
        Do  ' Fast-forward!
            
            '-- Perform a frame
            Call VBA.DoEvents
            Call DoFrame
            
            '-- Can release?
            If (GetGameStage = [eStagePlaying] And m_bArmageddonActivated = False) Then
                
                '-- Can release one more?
                If (GetLemsOut < g_uLevel.LemsToLetOut) Then
                    
                    '-- 'New Lem' release delay counter
                    m_lcNewLem = m_lcNewLem + 1
                    
                    '-- Release it now?
                    If (m_lcNewLem > 52.5 - g_lReleaseRate / 2 Or _
                        GetLemsOut() = 0 _
                        ) Then ' minimum 4 frames (rate = 99)
                        
                        '-- Reset counter
                        m_lcNewLem = 0
                        
                        '-- New Lem
                        Call AddLem
                    End If
                End If
              
              Else
                
                '-- Can activate one more?
                If (m_lcArmageddonLem < GetLemsOut()) Then
                    
                    '-- Oh no!
                    m_lcArmageddonLem = m_lcArmageddonLem + 1
                    Call SetArmageddonLem(m_lcArmageddonLem)
                End If
            End If
            
            '-- Time...
            If (GetGameStage() = [eStagePlaying]) Then
                
                '-- *Second* accumulator
                m_lcTimer = m_lcTimer + 1
                If (m_lcTimer = TMR_TIMESEC) Then
                    m_lcTimer = 0
                    
                    '-- One *second* less
                    g_lcPlayingTime = g_lcPlayingTime - 1
                    
                    '-- Update time
                    fLems.ucInfo.PanelText(5) = GetMinSecString(g_lcPlayingTime)
                    
                    '-- Sorry, time over
                    If (g_lcPlayingTime = 0) Then
                        Call SetGameStage([eStageEnding])
                    End If
                End If
            End If
            
        Loop Until (GetGameStage() = [eStageEnding] Or _
                    IsFastForward = False Or _
                    m_bTimerActivated = False Or _
                    m_bTimerPaused = True _
                    )
    End If
End Sub
