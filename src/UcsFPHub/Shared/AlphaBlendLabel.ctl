VERSION 5.00
Begin VB.UserControl AlphaBlendLabel 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ClipBehavior    =   0  'None
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Windowless      =   -1  'True
End
Attribute VB_Name = "AlphaBlendLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
'
' AlphaBlendLabel (c) 2020-2023 by wqweto@gmail.com
'
' Poor Man's Label Control
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "AlphaBlendLabel"

#Const ImplUseShared = (DebugMode <> 0)

'=========================================================================
' Public enums
'=========================================================================

Public Enum UcsLabelTextAlignEnum
    ucsLtaHorLeft = 0
    ucsLtaHorCenter = 1
    ucsLtaHorRight = 2
    ucsLtaVertTop = 0
    ucsLtaVertCenter = 4
    ucsLtaVertBottom = 8
    ucsLtaCenter = ucsLtaHorCenter Or ucsLtaVertCenter
End Enum

Public Enum UcsLabelTextFlagsEnum
    ucsLtfNone = 0
    ucsLtfDirectionRightToLeft = &H1 * 16
    ucsLtfDirectionVertical = &H2 * 16
    ucsLtfNoFitBlackBox = &H4 * 16
    ucsLtfDisplayFormatControl = &H20 * 16
    ucsLtfNoFontFallback = &H400 * 16
    ucsLtfMeasureTrailingSpaces = &H800& * 16
    ucsLtfNoWrap = &H1000& * 16
    ucsLtfLineLimit = &H2000& * 16
    ucsLtfNoClip = &H4000& * 16
End Enum

'=========================================================================
' Public events
'=========================================================================

Event Click()
Attribute Click.VB_UserMemId = -600
Event OwnerDraw(ByVal hGraphics As Long, ByVal hFont As Long, sCaption As String, sngLeft As Single, sngTop As Single, sngWidth As Single, sngHeight As Single)
Event DblClick()
Attribute DblClick.VB_UserMemId = -601
Event ContextMenu()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_UserMemId = -607

'=========================================================================
' API
'=========================================================================

'--- DIB Section constants
Private Const DIB_RGB_COLORS                As Long = 0
'--- for AlphaBlend
Private Const AC_SRC_ALPHA                  As Long = 1
'--- for GdipSetTextRenderingHint
Private Const TextRenderingHintAntiAlias    As Long = 4
Private Const TextRenderingHintClearTypeGridFit As Long = 5
'--- for GdipSetSmoothingMode
Private Const SmoothingModeAntiAlias        As Long = 4
Private Const DT_CALCRECT                   As Long = &H400

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, lpBits As Long, ByVal Handle As Long, ByVal dw As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, ByVal lX As Long, ByVal lY As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As Any, ByVal wFormat As Long) As Long
'--- GDI+
Private Declare Function GdiplusStartup Lib "gdiplus" (hToken As Long, pInputBuf As Any, Optional ByVal pOutputBuf As Long = 0) As Long
Private Declare Function GdipCreateFontFromDC Lib "gdiplus" (ByVal hDC As Long, hCreatedFont As Long) As Long
Private Declare Function GdipDeleteFont Lib "gdiplus" (ByVal hFont As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, hBrush As Long) As Long
Private Declare Function GdipSetSolidFillColor Lib "gdiplus" (ByVal hBrush As Long, ByVal argb As Long) As Long
Private Declare Function GdipSetTextRenderingHint Lib "gdiplus" (ByVal hGraphics As Long, ByVal lMode As Long) As Long
Private Declare Function GdipDrawString Lib "gdiplus" (ByVal hGraphics As Long, ByVal lStrPtr As Long, ByVal lLength As Long, ByVal hFont As Long, uRect As RECTF, ByVal hStringFormat As Long, ByVal hBrush As Long) As Long
Private Declare Function GdipMeasureString Lib "gdiplus" (ByVal hGraphics As Long, ByVal lStrPtr As Long, ByVal lLength As Long, ByVal hFont As Long, uRect As RECTF, ByVal hStringFormat As Long, uBoundingBox As RECTF, lCodepointsFitted As Long, lLinesFilled As Long) As Long
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal hFormatAttributes As Long, ByVal nLanguage As Integer, hStringFormat As Long) As Long
Private Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal hStringFormat As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal hBrush As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipSetStringFormatFlags Lib "gdiplus" (ByVal hStringFormat As Long, ByVal lFlags As Long) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal hStringFormat As Long, ByVal eAlign As StringAlignment) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "gdiplus" (ByVal hStringFormat As Long, ByVal eAlign As StringAlignment) As Long
Private Declare Function GdipFillRectangle Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal sngX As Single, ByVal sngY As Single, ByVal sngWidth As Single, ByVal sngHeight As Single) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal lSmoothingMd As Long) As Long

Private Type BITMAPINFOHEADER
    biSize              As Long
    biWidth             As Long
    biHeight            As Long
    biPlanes            As Integer
    biBitCount          As Integer
    biCompression       As Long
    biSizeImage         As Long
    biXPelsPerMeter     As Long
    biYPelsPerMeter     As Long
    biClrUsed           As Long
    biClrImportant      As Long
End Type

Public Enum StringAlignment
   StringAlignmentNear = 0
   StringAlignmentCenter = 1
   StringAlignmentFar = 2
End Enum

Private Type RECTF
   Left                 As Single
   Top                  As Single
   Right                As Single
   Bottom               As Single
End Type

Private Type UcsRgbQuad
    R                   As Byte
    G                   As Byte
    B                   As Byte
    A                   As Byte
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const DEF_AUTOREDRAW        As Boolean = False
Private Const DEF_AUTOSIZE          As Boolean = False
Private Const DEF_TEXTOFFSETX       As Single = 0
Private Const DEF_TEXTOFFSETY       As Single = 0
Private Const DEF_FORECOLOR         As Long = vbButtonText
Private Const DEF_FOREOPACITY       As Single = 1
Private Const DEF_BACKCOLOR         As Long = vbButtonFace
Private Const DEF_BACKOPACITY       As Single = 1
Private Const DEF_SHADOWOFFSETX     As Single = 1
Private Const DEF_SHADOWOFFSETY     As Single = 1
Private Const DEF_SHADOWCOLOR       As Long = vbBlack
Private Const DEF_SHADOWOPACITY     As Single = 0
Private Const DEF_TEXTALIGN         As Long = ucsLtaHorLeft
Private Const DEF_TEXTFLAGS         As Long = 0
Private Const DEF_WORDWRAP          As Boolean = False
Private Const DEF_BORDERSTYLE       As Long = 0

Private m_bAutoRedraw           As Boolean
Private m_bAutoSize             As Boolean
Private m_sCaption              As String
Private WithEvents m_oFont      As StdFont
Attribute m_oFont.VB_VarHelpID = -1
Private m_sngTextOffsetX        As Single
Private m_sngTextOffsetY        As Single
Private m_clrFore               As OLE_COLOR
Private m_sngForeOpacity        As Single
Private m_clrBack               As OLE_COLOR
Private m_sngBackOpacity        As Single
Private m_sngShadowOffsetX      As Single
Private m_sngShadowOffsetY      As Single
Private m_clrShadow             As OLE_COLOR
Private m_sngShadowOpacity      As Single
Private m_eTextAlign            As UcsLabelTextAlignEnum
Private m_eTextFlags            As UcsLabelTextFlagsEnum
Private m_lBorderStyle          As Long
'--- run-time
Private m_bShown                As Boolean
Private m_eContainerScaleMode   As ScaleModeConstants
Private m_hFont                 As Long
Private m_hRedrawDib            As Long
Private m_nDownButton           As Integer
Private m_nDownShift            As Integer
Private m_sngDownX              As Single
Private m_sngDownY              As Single
Private m_sLastError            As String

'=========================================================================
' Error handling
'=========================================================================

Private Function PrintError(sFunction As String) As VbMsgBoxResult
    m_sLastError = Err.Description
    #If USE_DEBUG_LOG <> 0 Then
        DebugLog MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description & " &H" & Hex$(Err.Number), vbLogEventTypeError
    #Else
        Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    #End If
End Function

'=========================================================================
' Properties
'=========================================================================

Property Get AutoRedraw() As Boolean
    AutoRedraw = m_bAutoRedraw
End Property

Property Let AutoRedraw(ByVal bValue As Boolean)
    If m_bAutoRedraw <> bValue Then
        m_bAutoRedraw = bValue
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get AutoSize() As Boolean
Attribute AutoSize.VB_UserMemId = -500
    AutoSize = m_bAutoSize
End Property

Property Let AutoSize(ByVal bValue As Boolean)
    If m_bAutoSize <> bValue Then
        m_bAutoSize = bValue
        If m_bAutoSize And TypeOf Extender Is VBControlExtender Then
            pvSizeExtender Extender
        End If
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
    Caption = m_sCaption
End Property

Property Let Caption(sValue As String)
    If m_sCaption <> sValue Then
        m_sCaption = sValue
        If m_bAutoSize And TypeOf Extender Is VBControlExtender Then
            pvSizeExtender Extender
        End If
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get Font() As StdFont
Attribute Font.VB_UserMemId = -512
    Set Font = m_oFont
End Property

Property Set Font(oValue As StdFont)
    If Not m_oFont Is oValue Then
        Set m_oFont = oValue
        pvPrepareFont m_oFont, m_hFont
        If m_bAutoSize And TypeOf Extender Is VBControlExtender Then
            pvSizeExtender Extender
        End If
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get TextOffsetX() As Single
    TextOffsetX = m_sngTextOffsetX
End Property

Property Let TextOffsetX(ByVal sngValue As Single)
    If m_sngTextOffsetX <> sngValue Then
        m_sngTextOffsetX = sngValue
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get TextOffsetY() As Single
    TextOffsetY = m_sngTextOffsetY
End Property

Property Let TextOffsetY(ByVal sngValue As Single)
    If m_sngTextOffsetY <> sngValue Then
        m_sngTextOffsetY = sngValue
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = m_clrFore
End Property

Property Let ForeColor(ByVal clrValue As OLE_COLOR)
    If m_clrFore <> clrValue Then
        m_clrFore = clrValue
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get ForeOpacity() As Single
    ForeOpacity = m_sngForeOpacity
End Property

Property Let ForeOpacity(ByVal sngValue As Single)
    If m_sngForeOpacity <> sngValue Then
        m_sngForeOpacity = IIf(sngValue > 1, 1, IIf(sngValue < 0, 0, sngValue))
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_UserMemId = -501
    BackColor = m_clrBack
End Property

Property Let BackColor(ByVal clrValue As OLE_COLOR)
    If m_clrBack <> clrValue Then
        m_clrBack = clrValue
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get BackOpacity() As Single
    BackOpacity = m_sngBackOpacity
End Property

Property Let BackOpacity(ByVal sngValue As Single)
    If m_sngBackOpacity <> sngValue Then
        m_sngBackOpacity = IIf(sngValue > 1, 1, IIf(sngValue < 0, 0, sngValue))
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get ShadowOffsetX() As Single
    ShadowOffsetX = m_sngShadowOffsetX
End Property

Property Let ShadowOffsetX(ByVal sngValue As Single)
    If m_sngShadowOffsetX <> sngValue Then
        m_sngShadowOffsetX = sngValue
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get ShadowOffsetY() As Single
    ShadowOffsetY = m_sngShadowOffsetY
End Property

Property Let ShadowOffsetY(ByVal sngValue As Single)
    If m_sngShadowOffsetY <> sngValue Then
        m_sngShadowOffsetY = sngValue
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get ShadowColor() As OLE_COLOR
    ShadowColor = m_clrShadow
End Property

Property Let ShadowColor(ByVal clrValue As OLE_COLOR)
    If m_clrShadow <> clrValue Then
        m_clrShadow = clrValue
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get ShadowOpacity() As Single
    ShadowOpacity = m_sngShadowOpacity
End Property

Property Let ShadowOpacity(ByVal sngValue As Single)
    If m_sngShadowOpacity <> sngValue Then
        m_sngShadowOpacity = IIf(sngValue > 1, 1, IIf(sngValue < 0, 0, sngValue))
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get TextAlign() As UcsLabelTextAlignEnum
    TextAlign = m_eTextAlign
End Property

Property Let TextAlign(ByVal eValue As UcsLabelTextAlignEnum)
    If m_eTextAlign <> eValue Then
        m_eTextAlign = eValue
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get TextFlags() As UcsLabelTextFlagsEnum
    TextFlags = m_eTextFlags
End Property

Property Let TextFlags(ByVal eValue As UcsLabelTextFlagsEnum)
    If m_eTextFlags <> eValue Then
        m_eTextFlags = eValue
        If m_bAutoSize And TypeOf Extender Is VBControlExtender Then
            pvSizeExtender Extender
        End If
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get BorderStyle() As Long
Attribute BorderStyle.VB_UserMemId = -504
    BorderStyle = m_lBorderStyle
End Property

Property Let BorderStyle(ByVal lValue As Long)
    If m_lBorderStyle <> lValue Then
        m_lBorderStyle = lValue
        pvRefresh
        PropertyChanged
    End If
End Property

Property Get WordWrap() As Boolean
    WordWrap = (m_eTextFlags And ucsLtfNoWrap) = 0
End Property

Property Let WordWrap(ByVal bValue As Boolean)
    If bValue Then
        TextFlags = m_eTextFlags And Not ucsLtfNoWrap
    Else
        TextFlags = m_eTextFlags Or ucsLtfNoWrap
    End If
End Property

Property Get LastError() As String
     LastError = m_sLastError
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function MeasureString(sCaption As String, sngWidth As Single, sngHeight As Single) As Boolean
    Dim hDC             As Long
    Dim hGraphics       As Long
    Dim uRect           As RECTF
    Dim uBounds         As RECTF
    Dim rcRect(0 To 3)  As Long
    Dim pFont           As IFont
    Dim hPrevFont       As Long
    Dim hStringFormat   As Long
    
    hDC = GetDC(ContainerHwnd)
    If hDC = 0 Then
        GoTo QH
    End If
    If m_hFont <> 0 Then
        If GdipCreateFromHDC(hDC, hGraphics) <> 0 Then
            GoTo QH
        End If
        If GdipSetSmoothingMode(hGraphics, SmoothingModeAntiAlias) <> 0 Then
            GoTo QH
        End If
        If GdipSetTextRenderingHint(hGraphics, IIf(m_bAutoRedraw, TextRenderingHintAntiAlias, TextRenderingHintClearTypeGridFit)) <> 0 Then
            GoTo QH
        End If
        If Not pvPrepareStringFormat(m_eTextAlign Or m_eTextFlags, hStringFormat) Then
            GoTo QH
        End If
        uRect.Right = ScaleX(sngWidth, m_eContainerScaleMode, vbPixels)
        uRect.Bottom = ScaleY(sngHeight, m_eContainerScaleMode, vbPixels)
        If GdipMeasureString(hGraphics, StrPtr(sCaption), Len(sCaption), m_hFont, uRect, hStringFormat, uBounds, 0, 0) <> 0 Then
            GoTo QH
        End If
        '--- ceil
        sngWidth = -Int(-uBounds.Right)
        sngHeight = -Int(-uBounds.Bottom)
    Else
        Set pFont = m_oFont
        hPrevFont = SelectObject(hDC, pFont.hFont)
        rcRect(2) = ScaleX(sngWidth, m_eContainerScaleMode, vbPixels)
        rcRect(3) = ScaleY(sngHeight, m_eContainerScaleMode, vbPixels)
        Call DrawText(hDC, StrPtr(sCaption), Len(sCaption), rcRect(0), DT_CALCRECT)
        Call SelectObject(hDC, hPrevFont)
        sngWidth = rcRect(2)
        sngHeight = rcRect(3)
    End If
    sngWidth = ScaleY(sngWidth, vbPixels, m_eContainerScaleMode)
    sngHeight = ScaleY(sngHeight, vbPixels, m_eContainerScaleMode)
    '--- success
    MeasureString = True
QH:
    If hStringFormat <> 0 Then
        Call GdipDeleteStringFormat(hStringFormat)
        hStringFormat = 0
    End If
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
        hGraphics = 0
    End If
    If hDC <> 0 Then
        Call ReleaseDC(ContainerHwnd, hDC)
        hDC = 0
    End If
End Function

Public Sub Refresh()
Attribute Refresh.VB_UserMemId = -550
    UserControl.Refresh
End Sub

'= private ===============================================================

Private Function pvPaintControl(ByVal hDC As Long) As Boolean
    Const FUNC_NAME     As String = "pvPaintControl"
    Dim hGraphics       As Long
    Dim hFont           As Long
    Dim sCaption        As String
    Dim hStringFormat   As Long
    Dim hBrush          As Long
    Dim uRect           As RECTF
    Dim sngLeft         As Single
    Dim sngTop          As Single
    Dim sngWidth        As Single
    Dim sngHeight       As Single
    Dim rcRect(0 To 3)  As Long
    Dim pFont           As IFont
    Dim hPrevFont       As Long
    Dim lTemp           As Long
    
    On Error GoTo EH
    If GetModuleHandle(StrPtr("gdiplus")) = 0 Then
        rcRect(2) = ScaleX(ScaleWidth, ScaleMode, vbPixels)
        rcRect(3) = ScaleY(ScaleHeight, ScaleMode, vbPixels)
        Set pFont = m_oFont
        hPrevFont = SelectObject(hDC, pFont.hFont)
        Call DrawText(hDC, StrPtr(m_sCaption), Len(m_sCaption), rcRect(0), 0)
        Call SelectObject(hDC, hPrevFont)
        GoTo QH
    End If
    If GdipCreateFromHDC(hDC, hGraphics) <> 0 Then
        GoTo QH
    End If
    If GdipSetSmoothingMode(hGraphics, SmoothingModeAntiAlias) <> 0 Then
        GoTo QH
    End If
    If GdipSetTextRenderingHint(hGraphics, IIf(m_bAutoRedraw, TextRenderingHintAntiAlias, TextRenderingHintClearTypeGridFit)) <> 0 Then
        GoTo QH
    End If
    hFont = m_hFont
    sCaption = m_sCaption
    sngWidth = ScaleWidth
    sngHeight = ScaleHeight
    RaiseEvent OwnerDraw(hGraphics, hFont, sCaption, sngLeft, sngTop, sngWidth, sngHeight)
    If sngWidth > 0 Then
        If GdipCreateSolidFill(pvTranslateColor(m_clrBack, m_sngBackOpacity), hBrush) <> 0 Then
            GoTo QH
        End If
        If GdipFillRectangle(hGraphics, hBrush, sngLeft + 0.5, sngTop + 0.5, sngWidth - 1, sngHeight - 1) <> 0 Then
            GoTo QH
        End If
        If Not pvPrepareStringFormat(m_eTextAlign Or m_eTextFlags, hStringFormat) Then
            GoTo QH
        End If
        uRect.Left = sngLeft + m_sngTextOffsetX
        uRect.Top = sngTop + m_sngTextOffsetY
        uRect.Right = sngLeft + sngWidth
        uRect.Bottom = sngTop + sngHeight
        If m_sngShadowOpacity <> 0 Then
            If GdipSetSolidFillColor(hBrush, pvTranslateColor(m_clrShadow, m_sngShadowOpacity)) <> 0 Then
                GoTo QH
            End If
            uRect.Left = uRect.Left + m_sngShadowOffsetX
            uRect.Top = uRect.Top + m_sngShadowOffsetY
            If GdipDrawString(hGraphics, StrPtr(sCaption), -1, hFont, uRect, hStringFormat, hBrush) <> 0 Then
                GoTo QH
            End If
            uRect.Left = uRect.Left - m_sngShadowOffsetX
            uRect.Top = uRect.Top - m_sngShadowOffsetY
        End If
        If GdipSetSolidFillColor(hBrush, pvTranslateColor(m_clrFore, m_sngForeOpacity)) <> 0 Then
            GoTo QH
        End If
        If GdipDrawString(hGraphics, StrPtr(sCaption), -1, hFont, uRect, hStringFormat, hBrush) <> 0 Then
            GoTo QH
        End If
        If m_lBorderStyle > 0 Then
            lTemp = pvTranslateColor(m_clrFore, m_sngForeOpacity)
            If Not pvDrawRect(hGraphics, sngLeft, sngTop, sngWidth, sngHeight, lTemp, lTemp, lTemp, lTemp) Then
                GoTo QH
            End If
        End If
    End If
    '-- success
    pvPaintControl = True
QH:
    On Error Resume Next
    If hFont <> 0 And hFont <> m_hFont Then
        Call GdipDeleteFont(hFont)
        hFont = 0
    End If
    If hStringFormat <> 0 Then
        Call GdipDeleteStringFormat(hStringFormat)
        hStringFormat = 0
    End If
    If hBrush <> 0 Then
        Call GdipDeleteBrush(hBrush)
        hBrush = 0
    End If
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
        hGraphics = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function pvPrepareFont(oFont As StdFont, hFont As Long) As Boolean
    pvPrepareFont = GdipPrepareFont(ContainerHwnd, oFont, hFont)
End Function

Private Function pvPrepareStringFormat(ByVal lFlags As Long, hStringFormat As Long) As Boolean
    pvPrepareStringFormat = GdipPrepareStringFormat(lFlags, hStringFormat)
End Function

Private Function pvTranslateColor(ByVal clrValue As OLE_COLOR, Optional ByVal Alpha As Single = 1) As Long
    Dim uQuad           As UcsRgbQuad
    Dim lTemp           As Long
    
    Call OleTranslateColor(clrValue, 0, VarPtr(uQuad))
    lTemp = uQuad.R
    uQuad.R = uQuad.B
    uQuad.B = lTemp
    lTemp = Alpha * &HFF
    If lTemp > 255 Then
        uQuad.A = 255
    ElseIf lTemp < 0 Then
        uQuad.A = 0
    Else
        uQuad.A = lTemp
    End If
    Call CopyMemory(pvTranslateColor, uQuad, 4)
End Function

Private Sub pvRefresh()
    m_bShown = False
    If m_hRedrawDib <> 0 Then
        Call DeleteObject(m_hRedrawDib)
        m_hRedrawDib = 0
    End If
    UserControl.Refresh
End Sub

Private Sub pvHandleMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_nDownButton = Button
    m_nDownShift = Shift
    m_sngDownX = X
    m_sngDownY = Y
End Sub

Private Sub pvSizeExtender(oExt As VBControlExtender)
    Dim sngLeft         As Single
    Dim sngTop          As Single
    Dim sngWidth        As Single
    Dim sngHeight       As Single
    
    sngLeft = oExt.Left
    sngTop = oExt.Top
    If WordWrap Then
        sngWidth = oExt.Width
    End If
    If MeasureString(IIf(LenB(m_sCaption) <> 0, m_sCaption, " "), sngWidth, sngHeight) Then
        If WordWrap Then
            sngWidth = oExt.Width
        Else
            Select Case (TextAlign And 3)
            Case ucsLtaHorRight
                sngLeft = sngLeft + oExt.Width - sngWidth
            Case ucsLtaHorCenter
                sngLeft = sngLeft + (oExt.Width - sngWidth) / 2
            End Select
        End If
        Select Case (TextAlign And 12)
        Case ucsLtaVertBottom
            sngTop = sngTop + oExt.Height - sngHeight
        Case ucsLtaVertCenter
            sngTop = sngTop + (oExt.Height - sngHeight) / 2
        End Select
        oExt.Move sngLeft, sngTop, sngWidth, sngHeight
    End If
End Sub

Private Function pvDrawRect(ByVal hGraphics As Long, _
            ByVal sngLeft As Single, ByVal sngTop As Single, ByVal sngWidth As Single, ByVal sngHeight As Single, _
            ByVal clrLeft As Long, ByVal clrTop As Long, ByVal clrRight As Long, ByVal clrBottom As Long) As Boolean
    Const FUNC_NAME     As String = "pvDrawRect"
    Dim hBrush          As Long
    Dim sngPixel        As Single
    
    On Error GoTo EH
    If GetModuleHandle(StrPtr("gdiplus")) = 0 Then
        GoTo QH
    End If
    sngPixel = IconScale(16!) / 16!
    If GdipCreateSolidFill(clrLeft, hBrush) <> 0 Then
        GoTo QH
    End If
    If GdipFillRectangle(hGraphics, hBrush, sngLeft + 0.5, sngTop + 0.5, sngPixel, sngHeight) <> 0 Then
        GoTo QH
    End If
    If GdipSetSolidFillColor(hBrush, clrTop) <> 0 Then
        GoTo QH
    End If
    If GdipFillRectangle(hGraphics, hBrush, sngLeft + 0.5 + sngPixel, sngTop + 0.5, sngWidth, sngPixel) <> 0 Then
        GoTo QH
    End If
    If GdipSetSolidFillColor(hBrush, clrRight) <> 0 Then
        GoTo QH
    End If
    If GdipFillRectangle(hGraphics, hBrush, sngLeft + sngWidth - 0.5 - sngPixel, sngTop + 0.5 + sngPixel, sngPixel, sngHeight - sngPixel) <> 0 Then
        GoTo QH
    End If
    If GdipSetSolidFillColor(hBrush, clrBottom) <> 0 Then
        GoTo QH
    End If
    If GdipFillRectangle(hGraphics, hBrush, sngLeft + 0.5 + sngPixel, sngTop + sngHeight - 0.5 - sngPixel, sngWidth - 2 * sngPixel - 1, sngPixel) <> 0 Then
        GoTo QH
    End If
    pvDrawRect = True
QH:
    If hBrush <> 0 Then
        Call GdipDeleteBrush(hBrush)
        hBrush = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

'= common ================================================================

Private Function pvCreateDib(ByVal hMemDC As Long, ByVal lWidth As Long, ByVal lHeight As Long, hDib As Long, Optional lpBits As Long) As Boolean
    Const FUNC_NAME     As String = "pvCreateDib"
    Dim uHdr            As BITMAPINFOHEADER
    
    On Error GoTo EH
    With uHdr
        .biSize = LenB(uHdr)
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = lWidth
        .biHeight = -lHeight
        .biSizeImage = 4 * lWidth * lHeight
    End With
    hDib = CreateDIBSection(hMemDC, uHdr, DIB_RGB_COLORS, lpBits, 0, 0)
    If hDib = 0 Then
        GoTo QH
    End If
    '--- success
    pvCreateDib = True
QH:
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function ToScaleMode(sScaleUnits As String) As ScaleModeConstants
    Select Case sScaleUnits
    Case "Twip"
        ToScaleMode = vbTwips
    Case "Point"
        ToScaleMode = vbPoints
    Case "Pixel"
        ToScaleMode = vbPixels
    Case "Character"
        ToScaleMode = vbCharacters
    Case "Centimeter"
        ToScaleMode = vbCentimeters
    Case "Millimeter"
        ToScaleMode = vbMillimeters
    Case "Inch"
        ToScaleMode = vbInches
    Case Else
        ToScaleMode = vbTwips
    End Select
End Function

Private Function FontToString(oFont As StdFont) As String
    FontToString = oFont.Name & oFont.Size & oFont.Bold & oFont.Italic & oFont.Underline & oFont.Strikethrough & oFont.Weight & oFont.Charset
End Function

#If Not ImplUseShared Then
Private Function GdipPrepareFont(ByVal hWnd As Long, oFont As StdFont, hFont As Long) As Boolean
    Const FUNC_NAME     As String = "GdipPrepareFont"
    Dim hDC             As Long
    Dim pFont           As IFont
    Dim hPrevFont       As Long
    Dim hNewFont        As Long
    
    On Error GoTo EH
    Set pFont = oFont
    If pFont Is Nothing Then
        GoTo QH
    End If
    hDC = GetDC(hWnd)
    If hDC = 0 Then
        GoTo QH
    End If
    hPrevFont = SelectObject(hDC, pFont.hFont)
    If hPrevFont = 0 Then
        GoTo QH
    End If
    If GdipCreateFontFromDC(hDC, hNewFont) <> 0 Then
        GoTo QH
    End If
    '--- commit
    If hFont <> 0 Then
        Call GdipDeleteFont(hFont)
    End If
    hFont = hNewFont
    hNewFont = 0
    '--- success
    GdipPrepareFont = True
QH:
    If hNewFont <> 0 Then
        Call GdipDeleteFont(hNewFont)
        hNewFont = 0
    End If
    If hPrevFont <> 0 Then
        Call SelectObject(hDC, hPrevFont)
        hPrevFont = 0
    End If
    If hDC <> 0 Then
        Call ReleaseDC(hWnd, hDC)
        hDC = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Private Function GdipPrepareStringFormat(ByVal lFlags As UcsLabelTextFlagsEnum, hStringFormat As Long) As Boolean
    Const FUNC_NAME     As String = "GdipPrepareStringFormat"
    Dim hNewFormat      As Long
    
    On Error GoTo EH
    If GdipCreateStringFormat(0, 0, hNewFormat) <> 0 Then
        GoTo QH
    End If
    If GdipSetStringFormatAlign(hNewFormat, lFlags And 3) <> 0 Then
        GoTo QH
    End If
    If GdipSetStringFormatLineAlign(hNewFormat, (lFlags \ 4) And 3) <> 0 Then
        GoTo QH
    End If
    If GdipSetStringFormatFlags(hNewFormat, lFlags \ 16) <> 0 Then
        GoTo QH
    End If
    '--- commit
    If hStringFormat <> 0 Then
        Call GdipDeleteStringFormat(hStringFormat)
    End If
    hStringFormat = hNewFormat
    hNewFormat = 0
    '--- success
    GdipPrepareStringFormat = True
QH:
    If hNewFormat <> 0 Then
        Call GdipDeleteStringFormat(hNewFormat)
        hNewFormat = 0
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Private Property Get ScreenTwipsPerPixelX() As Single
    ScreenTwipsPerPixelX = Screen.TwipsPerPixelX
End Property

Private Function IconScale(ByVal sngSize As Single) As Long
    If ScreenTwipsPerPixelX < 6.5 Then
        IconScale = Int(sngSize * 3)
    ElseIf ScreenTwipsPerPixelX < 9.5 Then
        IconScale = Int(sngSize * 2)
    ElseIf ScreenTwipsPerPixelX < 11.5 Then
        IconScale = Int(sngSize * 3 \ 2)
    Else
        IconScale = Int(sngSize * 1)
    End If
End Function
#End If

'=========================================================================
' Events
'=========================================================================

Private Sub m_oFont_FontChanged(ByVal PropertyName As String)
    pvPrepareFont m_oFont, m_hFont
    If m_bAutoSize And TypeOf Extender Is VBControlExtender Then
        pvSizeExtender Extender
    End If
    pvRefresh
    PropertyChanged
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, ScaleX(X, ScaleMode, m_eContainerScaleMode), ScaleY(Y, ScaleMode, m_eContainerScaleMode))
    pvHandleMouseDown Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, ScaleX(X, ScaleMode, m_eContainerScaleMode), ScaleY(Y, ScaleMode, m_eContainerScaleMode))
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Const FUNC_NAME     As String = "UserControl_MouseUp"
    
    On Error GoTo EH
    RaiseEvent MouseUp(Button, Shift, ScaleX(X, ScaleMode, m_eContainerScaleMode), ScaleY(Y, ScaleMode, m_eContainerScaleMode))
    If Button = -1 Then
        GoTo QH
    End If
    If Button <> 0 And X >= 0 And X < ScaleWidth And Y >= 0 And Y < ScaleHeight Then
        If (m_nDownButton And Button And vbLeftButton) <> 0 Then
            RaiseEvent Click
        ElseIf (m_nDownButton And Button And vbRightButton) <> 0 Then
            RaiseEvent ContextMenu
        End If
    End If
    m_nDownButton = 0
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub UserControl_DblClick()
    pvHandleMouseDown vbLeftButton, m_nDownShift, m_sngDownX, m_sngDownY
    RaiseEvent DblClick
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    HitResult = vbHitResultHit
End Sub

Private Sub UserControl_Resize()
    pvRefresh
End Sub

Private Sub UserControl_Hide()
    m_bShown = False
End Sub

Private Sub UserControl_Paint()
    Const FUNC_NAME     As String = "UserControl_Paint"
    Const Opacity       As Long = &HFF
    Dim hMemDC          As Long
    Dim hPrevDib        As Long
    
    On Error GoTo EH
    If AutoRedraw Then
        hMemDC = CreateCompatibleDC(hDC)
        If hMemDC = 0 Then
            GoTo DefPaint
        End If
        If m_hRedrawDib = 0 Then
            If Not pvCreateDib(hMemDC, ScaleWidth, ScaleHeight, m_hRedrawDib) Then
                GoTo DefPaint
            End If
            hPrevDib = SelectObject(hMemDC, m_hRedrawDib)
            If Not pvPaintControl(hMemDC) Then
                GoTo DefPaint
            End If
        Else
            hPrevDib = SelectObject(hMemDC, m_hRedrawDib)
        End If
        If AlphaBlend(hDC, 0, 0, ScaleWidth, ScaleHeight, hMemDC, 0, 0, ScaleWidth, ScaleHeight, AC_SRC_ALPHA * &H1000000 + Opacity * &H10000) = 0 Then
            GoTo DefPaint
        End If
    Else
        If Not pvPaintControl(hDC) Then
            GoTo DefPaint
        End If
    End If
    If False Then
DefPaint:
        If m_hRedrawDib <> 0 Then
            '--- note: before deleting DIB try de-selecting from dc
            Call SelectObject(hMemDC, hPrevDib)
            Call DeleteObject(m_hRedrawDib)
            m_hRedrawDib = 0
        End If
    End If
QH:
    On Error Resume Next
    If hMemDC <> 0 Then
        Call SelectObject(hMemDC, hPrevDib)
        Call DeleteDC(hMemDC)
        hMemDC = 0
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub UserControl_InitProperties()
    Const FUNC_NAME     As String = "UserControl_InitProperties"
    
    On Error GoTo EH
    m_eContainerScaleMode = ToScaleMode(Ambient.ScaleUnits)
    m_bAutoRedraw = DEF_AUTOREDRAW
    m_bAutoSize = DEF_AUTOSIZE
    m_sCaption = Ambient.DisplayName
    Set m_oFont = Ambient.Font
    m_sngTextOffsetX = DEF_TEXTOFFSETX
    m_sngTextOffsetY = DEF_TEXTOFFSETY
    m_clrFore = DEF_FORECOLOR
    m_sngForeOpacity = DEF_FOREOPACITY
    m_clrBack = DEF_BACKCOLOR
    m_sngBackOpacity = DEF_BACKOPACITY
    m_sngShadowOffsetX = DEF_SHADOWOFFSETX
    m_sngShadowOffsetY = DEF_SHADOWOFFSETY
    m_clrShadow = DEF_SHADOWCOLOR
    m_sngShadowOpacity = DEF_SHADOWOPACITY
    m_eTextAlign = DEF_TEXTALIGN
    m_eTextFlags = DEF_TEXTFLAGS
    pvPrepareFont m_oFont, m_hFont
    Exit Sub
EH:
    PrintError FUNC_NAME
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Const FUNC_NAME     As String = "UserControl_ReadProperties"
    
    On Error GoTo EH
    m_eContainerScaleMode = ToScaleMode(Ambient.ScaleUnits)
    With PropBag
        m_bAutoRedraw = .ReadProperty("AutoRedraw", DEF_AUTOREDRAW)
        m_bAutoSize = .ReadProperty("AutoSize", DEF_AUTOSIZE)
        m_sCaption = .ReadProperty("Caption", vbNullString)
        Set m_oFont = .ReadProperty("Font", Ambient.Font)
        m_sngTextOffsetX = .ReadProperty("TextOffsetX", DEF_TEXTOFFSETX)
        m_sngTextOffsetY = .ReadProperty("TextOffsetY", DEF_TEXTOFFSETY)
        m_clrFore = .ReadProperty("ForeColor", DEF_FORECOLOR)
        m_sngForeOpacity = .ReadProperty("ForeOpacity", DEF_FOREOPACITY)
        m_clrBack = .ReadProperty("BackColor", DEF_BACKCOLOR)
        m_sngBackOpacity = .ReadProperty("BackOpacity", DEF_BACKOPACITY)
        m_sngShadowOffsetX = .ReadProperty("ShadowOffsetX", DEF_SHADOWOFFSETX)
        m_sngShadowOffsetY = .ReadProperty("ShadowOffsetY", DEF_SHADOWOFFSETY)
        m_clrShadow = .ReadProperty("ShadowColor", DEF_SHADOWCOLOR)
        m_sngShadowOpacity = .ReadProperty("ShadowOpacity", DEF_SHADOWOPACITY)
        m_eTextAlign = .ReadProperty("TextAlign", DEF_TEXTALIGN)
        m_eTextFlags = .ReadProperty("TextFlags", DEF_TEXTFLAGS)
        WordWrap = .ReadProperty("WordWrap", DEF_WORDWRAP)
        m_lBorderStyle = Abs(.ReadProperty("BorderStyle", DEF_BORDERSTYLE))
    End With
    pvPrepareFont m_oFont, m_hFont
    If m_bAutoSize And TypeOf Extender Is VBControlExtender Then
        pvSizeExtender Extender
    End If
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Const FUNC_NAME     As String = "UserControl_ReadProperties"
    
    On Error GoTo EH
    With PropBag
        .WriteProperty "AutoRedraw", m_bAutoRedraw, DEF_AUTOREDRAW
        .WriteProperty "AutoSize", m_bAutoSize, DEF_AUTOSIZE
        .WriteProperty "BackColor", m_clrBack, DEF_BACKCOLOR
        .WriteProperty "BackOpacity", m_sngBackOpacity, DEF_BACKOPACITY
        .WriteProperty "BorderStyle", m_lBorderStyle, DEF_BORDERSTYLE
        .WriteProperty "Caption", m_sCaption, vbNullString
        If FontToString(m_oFont) <> FontToString(Ambient.Font) Then
            .WriteProperty "Font", m_oFont, Ambient.Font
        End If
        .WriteProperty "ForeColor", m_clrFore, DEF_FORECOLOR
        .WriteProperty "ForeOpacity", m_sngForeOpacity, DEF_FOREOPACITY
        .WriteProperty "ShadowColor", m_clrShadow, DEF_SHADOWCOLOR
        .WriteProperty "ShadowOffsetX", m_sngShadowOffsetX, DEF_SHADOWOFFSETX
        .WriteProperty "ShadowOffsetY", m_sngShadowOffsetY, DEF_SHADOWOFFSETY
        .WriteProperty "ShadowOpacity", m_sngShadowOpacity, DEF_SHADOWOPACITY
        .WriteProperty "TextAlign", m_eTextAlign, DEF_TEXTALIGN
        .WriteProperty "TextFlags", m_eTextFlags, DEF_TEXTFLAGS
        .WriteProperty "TextOffsetX", m_sngTextOffsetX, DEF_TEXTOFFSETX
        .WriteProperty "TextOffsetY", m_sngTextOffsetY, DEF_TEXTOFFSETY
        .WriteProperty "WordWrap", WordWrap, DEF_WORDWRAP
    End With
QH:
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

'Private Sub UserControl_AmbientChanged(PropertyName As String)
'    If PropertyName = "ScaleUnits" Then
'        m_eContainerScaleMode = ToScaleMode(Ambient.ScaleUnits)
'    End If
'End Sub

'=========================================================================
' Base class events
'=========================================================================

Private Sub UserControl_Initialize()
    Dim aInput(0 To 3)  As Long
    
    If GetModuleHandle(StrPtr("gdiplus")) = 0 Then
        If LoadLibrary(StrPtr("gdiplus")) <> 0 Then
            aInput(0) = 1
            Call GdiplusStartup(0, aInput(0))
        End If
    End If
    m_eContainerScaleMode = vbTwips
End Sub

Private Sub UserControl_Terminate()
    If m_hFont <> 0 Then
        Call GdipDeleteFont(m_hFont)
        m_hFont = 0
    End If
    If m_hRedrawDib <> 0 Then
        Call DeleteObject(m_hRedrawDib)
        m_hRedrawDib = 0
    End If
End Sub
