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
    lStructSize    As Long
    hwndOwner      As Long
    hInstance      As Long
    rgbResult      As Long
    lpCustColors   As Long
    Flags          As Long
    lCustData      As Long
    lpfnHook       As Long
    lpTemplateName As String
End Type

Private Const CC_RGBINIT   As Long = &H1
Private Const CC_FULLOPEN  As Long = &H2
Private Const CC_ANYCOLOR  As Long = &H100

Private Const CC_NORMAL    As Long = CC_ANYCOLOR Or CC_RGBINIT
Private Const CC_EXTENDED  As Long = CC_ANYCOLOR Or CC_RGBINIT Or CC_FULLOPEN

Private Declare Function ChooseColor Lib "comdlg32" Alias "ChooseColorA" (CCOLOR As tCHOOSECOLOR) As Long

'-- Private variables:

Private m_lCustomColors(15) As Long
Private m_bInitialized      As Boolean



'========================================================================================
' Methods
'========================================================================================

Public Function SelectColor( _
                ByVal hWndParent As Long, _
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
