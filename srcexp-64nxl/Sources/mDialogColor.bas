Attribute VB_Name = "mDialogColor"
'================================================
' Module:        mChooseColor.bas
' Author:        -
' Dependencies:  -
' Last revision: 2003.11.02
'================================================

Option Explicit

'-- API:

Private Type tCHOOSECOLOR
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    rgbResult As Long
    lpCustColors As LongPtr
    Flags As ChooseColorFlags
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
End Type

Private Enum ChooseColorFlags
    CC_RGBINIT = &H00000001
    CC_FULLOPEN = &H00000002
    CC_PREVENTFULLOPEN = &H00000004
    CC_SHOWHELP = &H00000008
    CC_ENABLEHOOK = &H00000010
    CC_ENABLETEMPLATE = &H00000020
    CC_ENABLETEMPLATEHANDLE = &H00000040
    CC_SOLIDCOLOR = &H00000080
    CC_ANYCOLOR = &H00000100
End Enum

Private Const CC_NORMAL    As Long = CC_ANYCOLOR Or CC_RGBINIT
Private Const CC_EXTENDED  As Long = CC_ANYCOLOR Or CC_RGBINIT Or CC_FULLOPEN

Private Declare PtrSafe Function ChooseColor Lib "comdlg32" Alias "ChooseColorA" (CCOLOR As tCHOOSECOLOR) As Long

'-- Private variables:

Private m_lCustomColors(15) As Long
Private m_bInitialized      As Boolean



'========================================================================================
' Methods
'========================================================================================

Public Function SelectColor( _
                ByVal hWndParent As LongPtr, _
                ByVal DefaultColor As Long, _
                Optional ByVal Extended As Boolean = False _
                ) As Long
 
  Dim uCC  As tCHOOSECOLOR
  Dim lRet As Long
  Dim lIdx As Long
 
    With uCC
        
        '-- Initiliaze custom colors (16 greys)
        If (m_bInitialized = False) Then
            m_bInitialized = True
            For lIdx = 0 To 15
                m_lCustomColors(lIdx) = RGB(lIdx * 17, lIdx * 17, lIdx * 17)
            Next lIdx
        End If
        
        '-- Prepare struct.
        .lStructSize = Len(uCC)
        .hwndOwner = hWndParent
        .rgbResult = DefaultColor
        .lpCustColors = VarPtr(m_lCustomColors(0))
        .Flags = IIf(Extended, CC_EXTENDED, CC_NORMAL)
        
        '-- Show Color dialog
        lRet = ChooseColor(uCC)
         
        '-- Get color / Cancel
        If (lRet) Then
            SelectColor = .rgbResult
          Else
            SelectColor = True
        End If
    End With
End Function
