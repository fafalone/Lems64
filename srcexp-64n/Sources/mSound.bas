Attribute VB_Name = "mSound"
Option Explicit

'-- API:

Private Declare PtrSafe Function PlaySound Lib "winmm" Alias "PlaySoundA" (lpData As Any, ByVal hModule As LongPtr, ByVal dwFlags As Long) As Long

Private Const SND_ASYNC     As Long = &H1
Private Const SND_NODEFAULT As Long = &H2
Private Const SND_MEMORY    As Long = &H4
Private Const SND_NOWAIT    As Long = &H2000

'-- Private types, variables:

Private Type tSFXData
    Data() As Byte
End Type

Private m_uSFX(24)      As tSFXData
Private m_bSoundEffects As Boolean

'-- Public enums.:

Public Enum eSoundFXConstants
    [eFXBang] = 0
    [eFXChain]
    [eFXChangeOp]
    [eFXChink]
    [eFXDie]
    [eFXDoor]
    [eFXElectric]
    [eFXExplode]
    [eFXFire]
    [eFXGlug]
    [eFXLetsGo]
    [eFXManTrap]
    [eFXMousePre]
    [eFXOhNo]
    [eFXOing]
    [eFXScrape]
    [eFXSlicer]
    [eFXSplash]
    [eFXSplat]
    [eFXTenton1x]
    [eFXTenton5x]
    [eFXThud]
    [eFXThunk]
    [eFXTing]
    [eFXYipee] = 24
End Enum



'========================================================================================
' Methods
'========================================================================================

Public Sub InitializeSound()
  
  Dim sPath As String
  
    sPath = AppPath & "SOUND\"
  
    Call pvLoadSoundStream(sPath & "Bang.wav", m_uSFX(0))
    Call pvLoadSoundStream(sPath & "Chain.wav", m_uSFX(1))
    Call pvLoadSoundStream(sPath & "ChangeOp.wav", m_uSFX(2))
    Call pvLoadSoundStream(sPath & "Chink.wav", m_uSFX(3))
    Call pvLoadSoundStream(sPath & "Die.wav", m_uSFX(4))
    Call pvLoadSoundStream(sPath & "Door.wav", m_uSFX(5))
    Call pvLoadSoundStream(sPath & "Electric.wav", m_uSFX(6))
    Call pvLoadSoundStream(sPath & "Explode.wav", m_uSFX(7))
    Call pvLoadSoundStream(sPath & "Fire.wav", m_uSFX(8))
    Call pvLoadSoundStream(sPath & "Glug.wav", m_uSFX(9))
    Call pvLoadSoundStream(sPath & "LetsGo.wav", m_uSFX(10))
    Call pvLoadSoundStream(sPath & "ManTrap.wav", m_uSFX(11))
    Call pvLoadSoundStream(sPath & "MousePre.wav", m_uSFX(12))
    Call pvLoadSoundStream(sPath & "OhNo.wav", m_uSFX(13))
    Call pvLoadSoundStream(sPath & "Oing.wav", m_uSFX(14))
    Call pvLoadSoundStream(sPath & "Scrape.wav", m_uSFX(15))
    Call pvLoadSoundStream(sPath & "Slicer.wav", m_uSFX(16))
    Call pvLoadSoundStream(sPath & "Splash.wav", m_uSFX(17))
    Call pvLoadSoundStream(sPath & "Splat.wav", m_uSFX(18))
    Call pvLoadSoundStream(sPath & "Tenton1x.wav", m_uSFX(19))
    Call pvLoadSoundStream(sPath & "Tenton5x.wav", m_uSFX(20))
    Call pvLoadSoundStream(sPath & "Thud.wav", m_uSFX(21))
    Call pvLoadSoundStream(sPath & "Thunk.wav", m_uSFX(22))
    Call pvLoadSoundStream(sPath & "Ting.wav", m_uSFX(23))
    Call pvLoadSoundStream(sPath & "Yipee.wav", m_uSFX(24))
    
    '-- By default, sound effects
    m_bSoundEffects = True
End Sub

Public Sub PlaySoundFX( _
           ByVal SoundFX As eSoundFXConstants _
           )
    
    '-- Play sound FX
    If (m_bSoundEffects) Then
        Call PlaySound( _
             m_uSFX(SoundFX).Data(0), _
             0, _
             SND_ASYNC Or SND_MEMORY Or SND_NOWAIT _
             )
    End If
End Sub

Public Sub SetSoundEffectsState( _
           ByVal Enable As Boolean _
           )
    
    '-- Enable/disable sound effects
    m_bSoundEffects = Enable
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub pvLoadSoundStream( _
            ByVal Filename As String, _
            uFX As tSFXData)

  Dim hFile As Long
    
    '-- Get a free file handle
    hFile = VBA.FreeFile()
    
    '-- Open file
    Open Filename For Binary Access Read As #hFile
        
    '-- Resize array and get sound data
    ReDim uFX.Data(FileLen(Filename) - 1)
    Get #hFile,, uFX.Data
    
    '-- Close
    Close #hFile
End Sub

