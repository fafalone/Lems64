Attribute VB_Name = "mMisc"
Option Explicit

'-- API:

Private Const BITSPIXEL As Long = 12

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Private Declare PtrSafe Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As LongPtr
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long

Private Const SW_SHOW As Long = 5

Private Declare PtrSafe Function PathCompactPath Lib "shlwapi" Alias "PathCompactPathA" (ByVal hDC As LongPtr, ByVal pszPath As String, ByVal dx As Long) As Long
Private Declare PtrSafe Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr

Private Const CLR_INVALID As Long = &HFFFF&

Private Declare PtrSafe Function OleTranslateColor Lib "oleaut32" (ByVal clr As stdole.OLE_COLOR, ByVal hpal As LongPtr, lpcolorref As Long) As Long

Private Enum GWL_INDEX
    GWL_WNDPROC = (-4)
    GWL_HINSTANCE = (-6)
    GWL_HWNDPARENT = (-8)
    GWL_ID = (-12)
    GWL_STYLE = (-16)
    GWL_EXSTYLE = (-20)
    GWL_USERDATA = (-21)
End Enum

Private Const BS_OWNERDRAW As Long = &HB

#If Win64 Then
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As GWL_INDEX) As LongPtr
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As GWL_INDEX, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As GWL_INDEX) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As GWL_INDEX, ByVal dwNewLong As Long) As Long
#End If

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal xy As LongLong) As LongPtr



'========================================================================================
' Methods
'========================================================================================

Public Function InIDE( _
                Optional c As Boolean = False _
                ) As Boolean
  
  Static b As Boolean
  
    b = c
    If (b = False) Then
        Debug.Assert InIDE(True)
    End If
    InIDE = b
    
' by ULLI
' http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=64778&lngWId=1
End Function

Public Function ScreenColourDepth( _
                ) As Long
 
 Dim hTmpDC As LongPtr
   
    hTmpDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    ScreenColourDepth = GetDeviceCaps(hTmpDC, BITSPIXEL)
    Call DeleteDC(hTmpDC)
End Function

Public Function FileExists( _
                ByVal Filename As String _
                ) As Boolean
    
    If (Len(Filename)) Then
        FileExists = (VBA.Dir$(Filename) <> vbNullString)
    End If
End Function

Public Function AppPath( _
                ) As String

    If (Right$(App.Path, 1) <> "\") Then
        AppPath = App.Path & "\"
      Else
        AppPath = App.Path
    End If
End Function

Public Function CompactPath( _
                ByVal hDC As LongPtr, _
                ByVal FullPath As String, _
                ByVal Width As Long _
                ) As String

  Dim lPos As Long

    '-- Compact
    Call PathCompactPath(hDC, FullPath, Width)

    '-- Remove all trailing Chr$(0)'s
    lPos = InStr(1, FullPath, Chr$(0))
    If (lPos > 0) Then
        CompactPath = Left$(FullPath, lPos - 1)
      Else
        CompactPath = FullPath
    End If
    
' from:
' KPD-Team 2000
' URL: http://www.allapi.net/
' e-mail: KPDTeam@Allapi.net
End Function

Public Sub Navigate( _
           ByVal hOwnerWnd As LongPtr, _
           ByVal URL As String _
           )
    
    Call ShellExecute(hOwnerWnd, "open", URL, vbNullString, vbNullString, SW_SHOW)
End Sub

Public Function TranslateColor( _
                ByVal Clr As OLE_COLOR, _
                Optional hPal As LongPtr = 0 _
                ) As Long
    
    '-- OLE/RGB color to RGB color
    If (OleTranslateColor(Clr, hPal, TranslateColor)) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Public Function ShiftColor( _
                ByVal Clr As Long, _
                ByVal Amount As Long _
                ) As Long

  Dim lR As Long
  Dim lB As Long
  Dim lG As Long
    
    '-- Add amount
    lR = (Clr And &HFF) + Amount
    lG = ((Clr \ &H100) Mod &H100) + Amount
    lB = ((Clr \ &H10000) Mod &H100) + Amount
    
    '-- Check byte bounds
    If (lR < 0) Then lR = 0 Else If (lR > 255) Then lR = 255
    If (lG < 0) Then lG = 0 Else If (lG > 255) Then lG = 255
    If (lB < 0) Then lB = 0 Else If (lB > 255) Then lB = 255
    
    '-- Return shifted color
    ShiftColor = lR + 256& * lG + 65536 * lB
End Function

Public Function BestConstrastColor( _
                ByVal BackColor As Long _
                ) As Long
                
  Dim lR As Long
  Dim lB As Long
  Dim lG As Long
  Dim lV As Long
    
    '-- Get components
    lR = (BackColor And &HFF)
    lG = (BackColor \ &H100) Mod &H100
    lB = (BackColor \ &H10000) Mod &H100
    
    '-- Brightness
    lV = (299 * lR + 587 * lG + 114 * lB) \ 1000
    
    '-- Black or white
    If (lV > 128) Then
        BestConstrastColor = vbBlack
      Else
        BestConstrastColor = vbWhite
    End If
End Function

Public Sub SetButtonOwnerDraw( _
           oButton As CommandButton, _
           bEnable As Boolean _
           )
  
  Dim lRet As Long
  
    lRet = CLng(GetWindowLong(oButton.hWnd, GWL_STYLE))
    If (bEnable) Then
        Call SetWindowLong(oButton.hWnd, GWL_STYLE, lRet Or BS_OWNERDRAW)
      Else
        Call SetWindowLong(oButton.hWnd, GWL_STYLE, lRet And Not BS_OWNERDRAW)
    End If
End Sub

Public Function GetCursorYPos( _
                ) As Long
      
  Dim uPt As POINTAPI
    
    Call GetCursorPos(uPt)
    GetCursorYPos = uPt.y
End Function

Public Function GethWndFromPoint( _
                ) As LongPtr
 
  Dim uPt As POINTAPI
  Dim pt As LongLong
    Call GetCursorPos(uPt)
    CopyMemory pt, uPt, LenB(pt)
    GethWndFromPoint = WindowFromPoint(pt)
End Function

