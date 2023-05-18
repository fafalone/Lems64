Attribute VB_Name = "mINI"
Option Explicit
Option Compare Text

'-- API:

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'========================================================================================
' Methods
'========================================================================================

Public Sub PutINI( _
           ByVal INIFileName As String, _
           ByVal INIHead As String, _
           ByVal INIKey As String, _
           ByVal INIVal As String _
           )
  
    Call WritePrivateProfileString( _
         INIHead, _
         INIKey, _
         INIVal, _
         INIFileName _
         )
End Sub

Public Function GetINI( _
                ByVal INIFileName As String, _
                ByVal INIHead As String, _
                ByVal INIKey As String _
                ) As String
  
  Dim lc As Long
  Dim s  As String * 260
    
    lc = GetPrivateProfileString( _
         INIHead, _
         INIKey, _
         vbNullString, _
         s, _
         Len(s), _
         INIFileName _
         )
         
    GetINI = Left$(s, lc)
End Function
