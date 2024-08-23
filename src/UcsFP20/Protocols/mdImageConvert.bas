Attribute VB_Name = "mdImageConvert"
'=========================================================================
'
' UcsFP20 (c) 2008-2024 by Unicontsoft
'
' Unicontsoft Fiscal Printers Component 2.0
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
'
' PDF and PNG/image to ZPL convertion functions
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const MODULE_NAME As String = "mdImageConvert"

'=========================================================================
' API
'=========================================================================

'--- for GdipCreateBitmapFromScan0
Private Const PixelFormat32bppPARGB         As Long = &HE200B
Private Const PixelFormat32bppARGB          As Long = &H26200A
Private Const PixelFormat1bppIndexed        As Long = &H30101
'--- for GdipDrawImageXxx
Private Const UnitPixel                     As Long = 2
'--- for GdipBitmapConvertFormat
Private Const DitherTypeSolid               As Long = 1
Private Const DitherTypeErrorDiffusion      As Long = 9
'--- for GdipInitializePalette
Private Const PaletteTypeCustom             As Long = 0
Private Const PaletteTypeFixedBW            As Long = 2
'--- for GdipBitmapLockBits
Private Const ImageLockModeRead             As Long = 1
'--- for GdipSetInterpolationMode
Private Const InterpolationModeHighQualityBicubic As Long = 7

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function SHCreateMemStream Lib "shlwapi" Alias "#12" (pInit As Any, ByVal cbInit As Long) As stdole.IUnknown
'---- pdfium
Private Declare Sub FPDF_InitLibrary Lib "pdfium" Alias "_FPDF_InitLibrary@0" ()
Private Declare Function FPDF_GetLastError Lib "pdfium" Alias "_FPDF_GetLastError@0" () As Long
Private Declare Function FPDF_LoadMemDocument Lib "pdfium" Alias "_FPDF_LoadMemDocument@12" (pData As Any, ByVal lSize As Long, ByVal sPassword As String) As Long
Private Declare Sub FPDF_CloseDocument Lib "pdfium" Alias "_FPDF_CloseDocument@4" (ByVal hDoc As Long)
Private Declare Function FPDF_LoadPage Lib "pdfium" Alias "_FPDF_LoadPage@8" (ByVal hDoc As Long, ByVal PageIdx As Long) As Long
Private Declare Sub FPDF_ClosePage Lib "pdfium" Alias "_FPDF_ClosePage@4" (ByVal hPage As Long)
Private Declare Function FPDF_GetPageWidth Lib "pdfium" Alias "_FPDF_GetPageWidth@4" (ByVal hPage As Long) As Double
Private Declare Function FPDF_GetPageHeight Lib "pdfium" Alias "_FPDF_GetPageHeight@4" (ByVal hPage As Long) As Double
Private Declare Function FPDFBitmap_Create Lib "pdfium" Alias "_FPDFBitmap_Create@12" (ByVal lWidth As Long, ByVal lHeight As Long, ByVal lAlpha As Long) As Long
Private Declare Sub FPDFBitmap_Destroy Lib "pdfium" Alias "_FPDFBitmap_Destroy@4" (ByVal hBM As Long)
Private Declare Sub FPDFBitmap_FillRect Lib "pdfium" Alias "_FPDFBitmap_FillRect@24" (ByVal hBM As Long, ByVal lLeft As Long, ByVal lTop As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByVal clrFill As Long)
Private Declare Sub FPDF_RenderPageBitmap Lib "pdfium" Alias "_FPDF_RenderPageBitmap@32" (ByVal hBM As Long, ByVal hPage As Long, ByVal lLeft As Long, ByVal lTop As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByVal lRotation As Long, ByVal lFlags As Long)
Private Declare Function FPDFBitmap_GetBuffer Lib "pdfium" Alias "_FPDFBitmap_GetBuffer@4" (ByVal hBM As Long) As Long
Private Declare Function FPDFBitmap_GetStride Lib "pdfium" Alias "_FPDFBitmap_GetStride@4" (ByVal hBM As Long) As Long
'--- gdiplus
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal lWidth As Long, ByVal lHeight As Long, ByVal lStride As Long, ByVal lPixelFormat As Long, ByVal pScanData As Long, hImage As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal hImage As Long) As Long
Private Declare Function GdipBitmapConvertFormat Lib "gdiplus" (ByVal hImage As Long, ByVal lFormat As Long, ByVal lDitherType As Long, ByVal lPaletteType As Long, pPalette As Any, ByVal AlphaThresholdPercent As Single) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, Optional ByVal srcUnit As Long = UnitPixel, Optional ByVal hImageAttributes As Long, Optional ByVal pfnCallback As Long, Optional ByVal lCallbackData As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipInitializePalette Lib "gdiplus" (pPalette As Any, ByVal lPaletteType As Long, ByVal lOptimalColors As Long, ByVal fUseTransparentColor As Long, ByVal hBitmap As Long) As Long
Private Declare Function GdipCloneImage Lib "gdiplus" (ByVal hImage As Long, hCloneImage As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal hImage As Long, nWidth As Single, nHeight As Single) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal pStream As stdole.IUnknown, mImage As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal hImage As Long, hGraphics As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal clrFill As Long, hBrush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal hBrush As Long) As Long
Private Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hBrush As Long, ByVal lX As Long, ByVal lY As Long, ByVal lWidth As Long, ByVal lHeight As Long) As Long
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal hBitmap As Long, lpRect As Any, ByVal lFlags As Long, ByVal lPixelFormat As Long, uLockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal hBitmap As Long, uLockedBitmapData As BitmapData) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal lMode As Long) As Long

Private Type ColorPalette
    Flags               As Long
    Count               As Long
    Entries(0 To 255)   As Long
End Type

Private Type BitmapData
    Width               As Long
    Height              As Long
    Stride              As Long
    PixelFormat         As Long
    Scan0               As Long
    reserved            As Long
End Type

Private m_hPdfiumLib                As Long
Private m_sLastError                As String

'=========================================================================
' Error handling
'=========================================================================

Private Function PrintError(sFunction As String) As VbMsgBoxResult
    #If USE_DEBUG_LOG <> 0 Then
        DebugLog MODULE_NAME, sFunction & "(" & Erl & ")", Err.Description & " &H" & Hex$(Err.Number), vbLogEventTypeError
    #Else
        Debug.Print "Critical error: " & Err.Description & " [" & MODULE_NAME & "." & sFunction & "]"
    #End If
End Function

'=========================================================================
' Functions
'=========================================================================

Public Function GetImageConvertLastError() As String
    GetImageConvertLastError = m_sLastError
End Function

Public Function LoadPdfPageToBitmap(baPdf() As Byte, _
            Optional TargetWidth As Long, _
            Optional TargetHeight As Long, _
            Optional ByVal PdfPage As Long, _
            Optional ByVal PdfRotation As Long, _
            Optional ByVal PdfFlags As Long) As Long
    Const FUNC_NAME     As String = "LoadPdfPageToBitmap"
    Dim hDoc            As Long
    Dim hPage           As Long
    Dim hBM             As Long
    Dim pData           As Long
    Dim hTempImg        As Long
    Dim hNewImg         As Long
    Dim lWidth          As Long
    Dim lHeight         As Long
    Dim hGraphics       As Long
    Dim sApiSource      As String
    
    On Error GoTo EH
    pvInit
    hDoc = FPDF_LoadMemDocument(baPdf(0), UBound(baPdf) + 1, vbNullString)
    If hDoc = 0 Then
        pvSetPdfError FPDF_GetLastError
        GoTo QH
    End If
    hPage = FPDF_LoadPage(hDoc, PdfPage)
    If hPage = 0 Then
        pvSetPdfError FPDF_GetLastError
        GoTo QH
    End If
    '--- convert width/height from points to pixels at 300 DPI
    lWidth = Int(FPDF_GetPageWidth(hPage) * 300# / 72# + 0.5)
    lHeight = Int(FPDF_GetPageHeight(hPage) * 300# / 72# + 0.5)
    pvSetAspect lWidth, lHeight, TargetWidth, TargetHeight
    hBM = FPDFBitmap_Create(lWidth, lHeight, 1)
    If hBM = 0 Then
        pvSetPdfError FPDF_GetLastError
        GoTo QH
    End If
    Call FPDFBitmap_FillRect(hBM, 0, 0, lWidth, lHeight, -1)
    Call FPDF_RenderPageBitmap(hBM, hPage, 0, 0, lWidth, lHeight, PdfRotation, PdfFlags)
    pData = FPDFBitmap_GetBuffer(hBM)
    If pData = 0 Then
        pvSetPdfError FPDF_GetLastError
        GoTo QH
    End If
    If pvCheckGdipError(GdipCreateBitmapFromScan0(lWidth, lHeight, FPDFBitmap_GetStride(hBM), PixelFormat32bppPARGB, pData, hTempImg)) Then
        sApiSource = "GdipCreateBitmapFromScan0"
        GoTo QH
    End If
    If pvCheckGdipError(GdipCreateBitmapFromScan0(TargetWidth, TargetHeight, 0, PixelFormat32bppPARGB, 0, hNewImg)) Then
        sApiSource = "GdipCreateBitmapFromScan0"
        GoTo QH
    End If
    If pvCheckGdipError(GdipGetImageGraphicsContext(hNewImg, hGraphics)) Then
        sApiSource = "GdipGetImageGraphicsContext"
        GoTo QH
    End If
    If pvCheckGdipError(GdipSetInterpolationMode(hGraphics, InterpolationModeHighQualityBicubic)) Then
        sApiSource = "GdipSetInterpolationMode"
        GoTo QH
    End If
    If pvCheckGdipError(GdipDrawImageRectRectI(hGraphics, hTempImg, 0, 0, TargetWidth, TargetHeight, 0, 0, lWidth, lHeight)) Then
        sApiSource = "GdipDrawImageRectRectI"
        GoTo QH
    End If
    '--- commit
    LoadPdfPageToBitmap = hNewImg
    hNewImg = 0
QH:
    If hBM <> 0 Then
        Call FPDFBitmap_Destroy(hBM)
    End If
    If hPage <> 0 Then
        Call FPDF_ClosePage(hPage)
    End If
    If hDoc <> 0 Then
        Call FPDF_CloseDocument(hDoc)
    End If
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
    End If
    If hTempImg <> 0 Then
        Call GdipDisposeImage(hTempImg)
    End If
    If hNewImg <> 0 Then
        Call GdipDisposeImage(hNewImg)
    End If
    If LenB(sApiSource) <> 0 Then
        m_sLastError = m_sLastError & " [" & FUNC_NAME & "." & sApiSource & "]"
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Function LoadPngToBitmap(baPng() As Byte, Optional TargetWidth As Long, Optional TargetHeight As Long) As Long
    Const FUNC_NAME     As String = "LoadPdfPageToBitmap"
    Dim pStream         As stdole.IUnknown
    Dim hImg            As Long
    Dim sngWidth        As Single
    Dim sngHeight       As Single
    Dim lWidth          As Long
    Dim lHeight         As Long
    Dim hNewImg         As Long
    Dim hGraphics       As Long
    Dim hBrush          As Long
    Dim sApiSource      As String
    
    On Error GoTo EH
    pvInit
    Set pStream = SHCreateMemStream(baPng(0), UBound(baPng) + 1)
    If pvCheckGdipError(GdipLoadImageFromStream(pStream, hImg)) Then
        sApiSource = "GdipLoadImageFromStream"
        GoTo QH
    End If
    If pvCheckGdipError(GdipGetImageDimension(hImg, sngWidth, sngHeight)) Then
        sApiSource = "GdipGetImageDimension"
        GoTo QH
    End If
    lWidth = Int(sngWidth + 0.5)
    lHeight = Int(sngHeight + 0.5)
    pvSetAspect lWidth, lHeight, TargetWidth, TargetHeight
    If pvCheckGdipError(GdipCreateBitmapFromScan0(TargetWidth, TargetHeight, 0, PixelFormat32bppARGB, 0, hNewImg)) Then
        sApiSource = "GdipCreateBitmapFromScan0"
        GoTo QH
    End If
    If pvCheckGdipError(GdipGetImageGraphicsContext(hNewImg, hGraphics)) Then
        sApiSource = "GdipGetImageGraphicsContext"
        GoTo QH
    End If
    If pvCheckGdipError(GdipSetInterpolationMode(hGraphics, InterpolationModeHighQualityBicubic)) Then
        sApiSource = "GdipSetInterpolationMode"
        GoTo QH
    End If
    If pvCheckGdipError(GdipCreateSolidFill(-1, hBrush)) Then
        sApiSource = "GdipCreateSolidFill"
        GoTo QH
    End If
    If pvCheckGdipError(GdipFillRectangleI(hGraphics, hBrush, 0, 0, TargetWidth, TargetHeight)) Then
        sApiSource = "GdipFillRectangleI"
        GoTo QH
    End If
    If pvCheckGdipError(GdipDrawImageRectRectI(hGraphics, hImg, 0, 0, TargetWidth, TargetHeight, 0, 0, lWidth, lHeight)) Then
        sApiSource = "GdipDrawImageRectRectI"
        GoTo QH
    End If
    '--- commit
    LoadPngToBitmap = hNewImg
    hNewImg = 0
QH:
    If hBrush <> 0 Then
        Call GdipDeleteBrush(hBrush)
    End If
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
    End If
    If hImg <> 0 Then
        Call GdipDisposeImage(hImg)
    End If
    If hNewImg <> 0 Then
        Call GdipDisposeImage(hNewImg)
    End If
    If LenB(sApiSource) <> 0 Then
        m_sLastError = m_sLastError & " [" & FUNC_NAME & "." & sApiSource & "]"
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Function ConvertBitmapToMonochrome(ByVal hImg As Long, Optional ByVal ErrorDiffusion As Boolean) As Long
    Const FUNC_NAME     As String = "ConvertBitmapToMonochrome"
    Dim uPal            As ColorPalette
    Dim lTemp           As Long
    Dim lDither         As Long
    Dim hNewImg         As Long
    Dim sApiSource      As String

    On Error GoTo EH
    pvInit
    uPal.Count = 2
    If pvCheckGdipError(GdipInitializePalette(uPal, PaletteTypeFixedBW, 0, 0, 0)) Then
        sApiSource = "GdipInitializePalette"
        GoTo QH
    End If
    '--- swap palette entries so white is 0 and black is 1
    lTemp = uPal.Entries(0)
    uPal.Entries(0) = uPal.Entries(1)
    uPal.Entries(1) = lTemp
    If pvCheckGdipError(GdipCloneImage(hImg, hNewImg)) Then
        sApiSource = "GdipCloneImage"
        GoTo QH
    End If
    lDither = IIf(ErrorDiffusion, DitherTypeErrorDiffusion, DitherTypeSolid)
    If pvCheckGdipError(GdipBitmapConvertFormat(hNewImg, PixelFormat1bppIndexed, lDither, PaletteTypeCustom, uPal, 0)) Then
        sApiSource = "GdipBitmapConvertFormat"
        GoTo QH
    End If
    '--- commit
    ConvertBitmapToMonochrome = hNewImg
    hNewImg = 0
QH:
    If hImg <> 0 Then
        Call GdipDisposeImage(hImg)
    End If
    If hNewImg <> 0 Then
        Call GdipDisposeImage(hNewImg)
    End If
    If LenB(sApiSource) <> 0 Then
        m_sLastError = m_sLastError & " [" & FUNC_NAME & "." & sApiSource & "]"
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Function ConvertBitmapToZplGraphics(ByVal hImg As Long, Optional ByVal Invert As Boolean) As String
    Const FUNC_NAME     As String = "ConvertBitmapToZplGraphics"
    Dim uData           As BitmapData
    Dim baData()        As Byte
    Dim lSize           As Long
    Dim lIdx            As Long
    Dim lStride         As Long
    Dim cOutput         As Collection
    Dim sApiSource      As String
    
    On Error GoTo EH
    pvInit
    If pvCheckGdipError(GdipBitmapLockBits(hImg, ByVal 0, ImageLockModeRead, PixelFormat1bppIndexed, uData)) Then
        sApiSource = "GdipBitmapLockBits"
        GoTo QH
    End If
    lStride = (uData.Width + 7) \ 8
    If uData.Stride < lStride Then
        lStride = uData.Stride
    End If
    lSize = lStride * uData.Height
    ReDim baData(0 To lSize - 1) As Byte
    For lIdx = 0 To uData.Height - 1
        Call CopyMemory(baData(lIdx * lStride), ByVal uData.Scan0 + lIdx * uData.Stride, lStride)
    Next
    If Invert Then
        For lIdx = 0 To UBound(baData)
            baData(lIdx) = baData(lIdx) Xor &HFF
        Next
    End If
    Set cOutput = New Collection
    cOutput.Add "A," & lSize & "," & lSize & "," & lStride & ","
    If Not pvToCompressedHex(baData, lStride, cOutput) Then
        GoTo QH
    End If
    ConvertBitmapToZplGraphics = ConcatCollection(cOutput, vbNullString)
QH:
    If uData.Scan0 <> 0 Then
        Call GdipBitmapUnlockBits(hImg, uData)
    End If
    If hImg <> 0 Then
        Call GdipDisposeImage(hImg)
    End If
    If LenB(sApiSource) <> 0 Then
        m_sLastError = m_sLastError & " [" & FUNC_NAME & "." & sApiSource & "]"
    End If
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume QH
End Function

Public Sub DrawBitmapToHDC(ByVal hDC As Long, ByVal hImg As Long, ByVal lLeft As Long, ByVal lTop As Long)
    Const FUNC_NAME     As String = "DrawBitmapToHDC"
    Dim hGraphics       As Long
    Dim sngWidth        As Single
    Dim sngHeight       As Single
    Dim lWidth          As Long
    Dim lHeight         As Long
    Dim sApiSource      As String
    
    On Error GoTo EH
    pvInit
    If pvCheckGdipError(GdipGetImageDimension(hImg, sngWidth, sngHeight)) Then
        sApiSource = "GdipGetImageDimension"
        GoTo QH
    End If
    lWidth = Int(sngWidth + 0.5)
    lHeight = Int(sngHeight + 0.5)
    If pvCheckGdipError(GdipCreateFromHDC(hDC, hGraphics)) Then
        sApiSource = "GdipCreateFromHDC"
        GoTo QH
    End If
    If pvCheckGdipError(GdipDrawImageRectRectI(hGraphics, hImg, lLeft, lTop, lWidth, lHeight, 0, 0, lWidth, lHeight)) Then
        sApiSource = "GdipDrawImageRectRectI"
        GoTo QH
    End If
QH:
    If hGraphics <> 0 Then
        Call GdipDeleteGraphics(hGraphics)
    End If
    If LenB(sApiSource) <> 0 Then
        m_sLastError = m_sLastError & " [" & FUNC_NAME & "." & sApiSource & "]"
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume QH
End Sub

'= private ===============================================================

Private Sub pvInit()
    Dim sFile           As String
    
    If m_hPdfiumLib = 0 Then
        sFile = LocateFile(App.Path & "\pdfium.dll")
        If LenB(sFile) <> 0 Then
            m_hPdfiumLib = LoadLibrary(StrPtr(sFile))
        End If
        Call FPDF_InitLibrary
        If m_hPdfiumLib = 0 Then
            m_hPdfiumLib = 1
        End If
    End If
    m_sLastError = vbNullString
End Sub

Private Sub pvSetAspect(ByVal lWidth As Long, ByVal lHeight As Long, lTargetWidth As Long, lTargetHeight As Long)
    If lTargetWidth = 0 And lTargetHeight = 0 Then
        lTargetWidth = lWidth
        lTargetHeight = lHeight
    ElseIf lTargetWidth <> 0 And lTargetHeight = 0 Then
        lTargetHeight = lHeight * lTargetWidth / lWidth
    ElseIf lTargetWidth = 0 And lTargetHeight <> 0 Then
        lTargetWidth = lWidth * lTargetHeight / lHeight
    Else
        If lHeight * lTargetWidth < lWidth * lTargetHeight Then
            lTargetHeight = lHeight * lTargetWidth / lWidth
        Else
            lTargetWidth = lWidth * lTargetHeight / lHeight
        End If
    End If
End Sub

'--- https://www.zebra.com/content/dam/support-dam/en/documentation/unrestricted/guide/software/zplii-pm-vol2-en.pdf
'--- Page 52: Alternative Data Compression Scheme for ~DG and ~DB Commands
Private Function pvToCompressedHex(baData() As Byte, ByVal lStride As Long, cOutput As Collection) As Boolean
    Dim sText           As String
    Dim lIdx            As Long
    Dim sPrevLine       As String
    Dim sLine           As String
    Dim lCount          As Long
    
    sText = ToHex(baData)
    lStride = lStride * 2
    For lIdx = 0 To (Len(sText) + lStride - 1) \ lStride
        sLine = Mid$(sText, lIdx * lStride + 1, lStride)
        If sLine = sPrevLine Then
            cOutput.Add ":"
        Else
            sPrevLine = sLine
            lCount = pvCountLastChar(sLine, "0") \ 2
            If lCount > 1 Then
                pvCompressLineData Left$(sLine, Len(sLine) - lCount * 2), cOutput
                cOutput.Add ","
            Else
                lCount = pvCountLastChar(sLine, "f") \ 2
                If lCount > 1 Then
                    pvCompressLineData Left$(sLine, Len(sLine) - lCount * 2), cOutput
                    cOutput.Add "!"
                Else
                    pvCompressLineData sLine, cOutput
                End If
            End If
        End If
    Next
    pvToCompressedHex = True
End Function

Private Sub pvCompressLineData(sText As String, cOutput As Collection)
    Const STR_MAP       As String = "_ G H I J K L M N O P Q R S T U V W X Y"
    Static bInitMap     As Boolean
    Static aLowerMap()  As String
    Static aUpperMap()  As String
    Dim oMatch          As Object
    Dim lOffset         As Long
    Dim lSize           As Long
    Dim sEncode         As String
    
    If Not bInitMap Then
        bInitMap = True
        aLowerMap = Split(LCase$(STR_MAP))
        aUpperMap = Split(UCase$(STR_MAP))
    End If
    For Each oMatch In InitRegExp("([0-9a-fA-F])\1{2,}").Execute(sText)
        If lOffset < oMatch.FirstIndex Then
            cOutput.Add Mid$(sText, lOffset + 1, oMatch.FirstIndex - lOffset)
        End If
        sEncode = vbNullString
        lSize = oMatch.Length
        Do While lSize >= 400
            sEncode = sEncode & "z"
            lSize = lSize - 400
        Loop
        If lSize >= 20 Then
            sEncode = sEncode & aLowerMap(lSize \ 20)
            lSize = lSize Mod 20
        End If
        If lSize > 0 Then
            sEncode = sEncode & aUpperMap(lSize)
        End If
        cOutput.Add sEncode & oMatch.SubMatches(0)
        lOffset = oMatch.FirstIndex + oMatch.Length
    Next
    If lOffset < Len(sText) Then
        cOutput.Add Mid$(sText, lOffset + 1)
    End If
End Sub

Private Function pvCountLastChar(sText As String, sChar As String) As Long
    Dim lIdx            As Long
    
    For lIdx = Len(sText) To 1 Step -1
        If Mid$(sText, lIdx, 1) <> sChar Then
            Exit For
        End If
    Next
    pvCountLastChar = Len(sText) - lIdx
End Function

Private Sub pvSetPdfError(ByVal lError As Long)
    m_sLastError = vbNullString
    Select Case lError
    Case 0: Exit Sub
    Case 1: m_sLastError = "Unknown error"
    Case 2: m_sLastError = "File not found or could not be opened"
    Case 3: m_sLastError = "File not in PDF format or corrupted"
    Case 4: m_sLastError = "Password required or incorrect password"
    Case 5: m_sLastError = "Unsupported security scheme"
    Case 6: m_sLastError = "Page not found or content error"
    Case 7: m_sLastError = "Load XFA error"
    Case 8: m_sLastError = "Layout XFA error"
    Case Else: m_sLastError = "FPDF error"
    End Select
    m_sLastError = m_sLastError & " (" & lError & ")"
End Sub

Private Function pvCheckGdipError(ByVal lStatus As Long) As Boolean
    m_sLastError = vbNullString
    Select Case lStatus
    Case 0: Exit Function
    Case 1: m_sLastError = "Generic error"
    Case 2: m_sLastError = "Invalid parameter"
    Case 3: m_sLastError = "Out of memory"
    Case 4: m_sLastError = "Object busy"
    Case 5: m_sLastError = "Insufficient buffer"
    Case 6: m_sLastError = "Not implemented"
    Case 7: m_sLastError = "Win32 error " & Err.LastDllError
    Case 8: m_sLastError = "Wrong state"
    Case 9: m_sLastError = "Aborted"
    Case 10: m_sLastError = "File not found"
    Case 11: m_sLastError = "Value overflow"
    Case 12: m_sLastError = "Access denied"
    Case 13: m_sLastError = "Unknown image format"
    Case 14: m_sLastError = "Font family not found"
    Case 15: m_sLastError = "Font style not found"
    Case 16: m_sLastError = "Not True Type font"
    Case 17: m_sLastError = "Unsupported Gdiplus version"
    Case 18: m_sLastError = "Gdiplus not initialized"
    Case 19: m_sLastError = "Property not found"
    Case 20: m_sLastError = "Property not supported"
    Case 21: m_sLastError = "Profile not found"
    Case Else: m_sLastError = "GDI+ error"
    End Select
    m_sLastError = m_sLastError & " (" & lStatus & ")"
    '--- failed
    pvCheckGdipError = True
End Function

