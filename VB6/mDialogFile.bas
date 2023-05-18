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
    lStructSize       As Long
    hWndOwner         As Long
    hInstance         As Long
    lpstrFilter       As String
    lpstrCustomFilter As String
    nMaxCustFilter    As Long
    nFilterIndex      As Long
    lpstrFile         As String
    nMaxFile          As Long
    lpstrFileTitle    As String
    nMaxFileTitle     As Long
    lpstrInitialDir   As String
    lpstrTitle        As String
    Flags             As Long
    nFileOffset       As Integer
    nFileExtension    As Integer
    lpstrDefExt       As String
    lCustData         As Long
    lpfnHook          As Long
    lpTemplateName    As String
End Type

Private Const OFN_HELPBUTTON      As Long = &H10
Private Const OFN_HIDEREADONLY    As Long = &H4
Private Const OFN_ENABLEHOOK      As Long = &H20
Private Const OFN_ENABLETEMPLATE  As Long = &H40
Private Const OFN_EXPLORER        As Long = &H80000
Private Const OFN_OVERWRITEPROMPT As Long = &H2
Private Const OFN_PATHMUSTEXIST   As Long = &H800
Private Const OFN_FILEMUSTEXISTS  As Long = &H1000
Private Const OFN_ENABLESIZING    As Long = &H800000
Private Const OFN_OPENFLAGS       As Long = &H881024
Private Const OFN_SAVEFLAGS       As Long = &H880026
Private Const MAX_PATH            As Long = 260

Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)



'========================================================================================
' Methods
'========================================================================================

Public Function GetFileName( _
                ByVal hWndOwner As Long, _
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
        .lStructSize = Len(uOFN)
        .hInstance = App.hInstance
        .hWndOwner = hWndOwner
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
