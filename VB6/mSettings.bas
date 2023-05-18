Attribute VB_Name = "mSettings"
Option Explicit
Option Compare Text

'========================================================================================
' Methods
'========================================================================================

Public Sub SetLevelDone( _
           ByVal ID As Integer, _
           ByVal Done As Boolean _
           )
    
    Call PutINI( _
         AppPath & "CONFIG\GAME.ini", _
         "Levels", "Level_" & Format$(ID, "0000"), _
         IIf(Done, "1", "0") _
         )
End Sub

Public Function IsLevelDone( _
                ByVal ID As Integer _
                ) As Boolean
    
    IsLevelDone = Val( _
         GetINI( _
         AppPath & "CONFIG\GAME.ini", _
         "Levels", "Level_" & Format$(ID, "0000")) _
         )
End Function

Public Sub SetLastLevel( _
           ByVal ID As Integer _
           )
    
    Call PutINI( _
         AppPath & "CONFIG\GAME.ini", _
         "Settings", "Last", _
         Format$(ID, "0000") _
         )
End Sub

Public Function GetLastLevel( _
                ) As Integer
    
    GetLastLevel = Val( _
         GetINI( _
         AppPath & "CONFIG\GAME.ini", _
         "Settings", "Last") _
         )
End Function

Public Sub SetLevelThumbnailDateTimeStamp( _
           ByVal ID As Integer, _
           ByVal Stamp As String _
           )
    
    Call PutINI( _
         AppPath & "CONFIG\GAME.ini", _
         "Thumbnails", "Thumbnail_" & Format$(ID, "0000"), _
         Stamp _
         )
End Sub

Public Function GetLevelThumbnailDateTimeStamp( _
                ByVal ID As Integer _
                ) As String
    
    GetLevelThumbnailDateTimeStamp = GetINI( _
         AppPath & "CONFIG\GAME.ini", _
         "Thumbnails", "Thumbnail_" & Format$(ID, "0000") _
         )
End Function

Public Sub SetWindowColor( _
           ByVal Clr As Long _
           )

    Call PutINI( _
         AppPath & "CONFIG\GAME.ini", _
         "Settings", "Color", _
         CStr(Clr) _
         )
End Sub

Public Function GetWindowColor( _
                ) As Long
  
  Dim sRet As String
  
    sRet = GetINI( _
         AppPath & "CONFIG\GAME.ini", _
         "Settings", "Color" _
         )
    
    If (sRet = vbNullString) Then
        GetWindowColor = vbButtonFace
      Else
        GetWindowColor = Val(sRet)
    End If
End Function

Public Sub SetSavedLemsMode( _
           ByVal Mode As eLemsSavedModeConstants _
           )
    
    Call PutINI( _
         AppPath & "CONFIG\GAME.ini", _
         "Settings", "SavedLemsMode", _
         CStr(Mode) _
         )
End Sub

Public Function GetSavedLemsMode( _
                ) As eLemsSavedModeConstants
    
    GetSavedLemsMode = Val( _
         GetINI( _
         AppPath & "CONFIG\GAME.ini", _
         "Settings", "SavedLemsMode") _
         )
End Function


