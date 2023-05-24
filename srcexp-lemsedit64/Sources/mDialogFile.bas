Attribute VB_Name = "mDialogFile"
'================================================
' Module:        mDialogFile.bas
' Author:        -
' Dependencies:  None
' Last revision: 2003.03.28
'================================================

Option Explicit

'-- API:

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As OFNFlags
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
    ' pvReserved As LongPtr
    ' dwReserved As Long
    ' FlagsEx As OFNFlagsEx
End Type

Private Enum OFNFlags
    OFN_READONLY = &H00000001
    OFN_OVERWRITEPROMPT = &H00000002
    OFN_HIDEREADONLY = &H00000004
    OFN_NOCHANGEDIR = &H00000008
    OFN_SHOWHELP = &H00000010
    OFN_ENABLEHOOK = &H00000020
    OFN_ENABLETEMPLATE = &H00000040
    OFN_ENABLETEMPLATEHANDLE = &H00000080
    OFN_NOVALIDATE = &H00000100
    OFN_ALLOWMULTISELECT = &H00000200
    OFN_EXTENSIONDIFFERENT = &H00000400
    OFN_PATHMUSTEXIST = &H00000800
    OFN_FILEMUSTEXIST = &H00001000
    OFN_CREATEPROMPT = &H00002000
    OFN_SHAREAWARE = &H00004000
    OFN_NOREADONLYRETURN = &H00008000&
    OFN_NOTESTFILECREATE = &H00010000
    OFN_NONETWORKBUTTON = &H00020000
    OFN_NOLONGNAMES = &H00040000  ' force no long names for 4.x modules
    OFN_EXPLORER = &H00080000  ' new look commdlg
    OFN_NODEREFERENCELINKS = &H00100000
    OFN_LONGNAMES = &H00200000  ' force long names for 3.x modules
'  OFN_ENABLEINCLUDENOTIFY and OFN_ENABLESIZING require
'  Windows 2000 or higher to have any effect.
    OFN_ENABLEINCLUDENOTIFY = &H00400000  ' send include message to callback
    OFN_ENABLESIZING = &H00800000
    OFN_DONTADDTORECENT = &H02000000
    OFN_FORCESHOWHIDDEN = &H10000000  ' Show All files including System and hidden files
End Enum
Private Enum OFNFlagsEx
    OFN_EX_NOPLACESBAR = &H1
End Enum
Private Const OFN_OPENFLAGS       As Long = &H881024
Private Const OFN_SAVEFLAGS       As Long = &H880026
Private Const MAX_PATH            As Long = 260

Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOFN As OPENFILENAME) As Long
Private Declare PtrSafe Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOFN As OPENFILENAME) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As LongPtr)



'========================================================================================
' Methods
'========================================================================================

Public Function GetFileName( _
                ByVal hWndOwner As LongPtr, _
                Optional Path As String, _
                Optional Filter As String, _
                Optional FilterIndex As Long = 1, _
                Optional Title As String, _
                Optional OpenDialog As Boolean = True _
                ) As String
   
 Dim uOFN As OPENFILENAME
 Dim lRet As Long
 Dim i    As Long
 
    For i = 1 To Len(Filter)
        If (Mid$(Filter, i, 1) = "|") Then
            Mid$(Filter, i, 1) = vbNullChar
        End If
    Next i
    
    If (Len(Filter) < MAX_PATH) Then
        Filter = Filter & String$(MAX_PATH - Len(Filter), 0)
      Else
        Filter = Filter & vbNullChar & vbNullChar
    End If

    With uOFN
        .lStructSize = LenB(uOFN)
        .hInstance = App.hInstance
        .hwndOwner = hWndOwner
        .lpstrTitle = Title
        .lpstrFilter = Filter
        .nFilterIndex = FilterIndex
        .nMaxFile = MAX_PATH
        .lpstrFile = Path & String(MAX_PATH - Len(Path), 0)
    End With
    
    If (OpenDialog) Then
        uOFN.Flags = uOFN.Flags Or OFN_OPENFLAGS
        lRet = GetOpenFileName(uOFN)
      Else
        uOFN.Flags = uOFN.Flags Or OFN_SAVEFLAGS
        lRet = GetSaveFileName(uOFN)
    End If
    Debug.Print "FD LastDllError=0x" & Err.LastDllError
    If (lRet) Then
        GetFileName = pvTrimNull(uOFN.lpstrFile)
    End If
End Function

'========================================================================================
' Private
'========================================================================================

Private Function pvTrimNull(StartString As String) As String
  
  Dim lPos As Long
  
    lPos = InStr(StartString, vbNullChar)
    If (lPos) Then
        pvTrimNull = Left$(StartString, lPos - 1)
      Else
        pvTrimNull = StartString
    End If
End Function
