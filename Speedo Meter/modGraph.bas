Attribute VB_Name = "modGraph"
' Collection of all kind of general graphic functions
' Parts by me and parts by others.

Option Explicit

Public Const RGN_AND = 1
Public Const RGN_OR = 2
Public Const RGN_XOR = 3
Public Const RGN_DIFF = 4
Public Const RGN_COPY = 5

Public Const FW_DONTCARE = 0                ' takes the default font weight
Public Const FW_ULTRALIGHT = 200
Public Const FW_LIGHT = 300
Public Const FW_NORMAL = 400                ' Normal font weight.
Public Const FW_BOLD = 700                  ' Bold font weight.
Public Const FW_EXTRABOLD = 800
Public Const FW_HEAVY = 900

Public Const CLIP_LH_ANGLES = 16            ' Needed for tilted fonts.

Public Const PS_SOLID = 0
Public Const PS_DOT = 2                     '  .......
Public Const PS_DASH = 1                    '  -------
Public Const PS_DASHDOT = 3                 '  _._._._
Public Const PS_DASHDOTDOT = 4              '  _.._.._
Public Const PS_NULL = 5

' Fill Types
Public Const VT_NONE = 0         ' VT_ from dutch: Vul Type
Public Const VT_SOLID = 1
Public Const VT_FLOOD = 2
Public Const VT_FLOODRAS = 3
Public Const VT_CLIPBRAS = 4

' Flood Types
Public Const FT_LEFTRIGHT = 0
Public Const FT_LEFTRIGHT2 = 1
Public Const FT_UPDOWN = 2
Public Const FT_UPDOWN2 = 3
Public Const FT_ULDR = 4         ' upper-left to bottom-right, etc.
Public Const FT_ULDR2 = 5
Public Const FT_DLUR = 6
Public Const FT_DLUR2 = 7
Public Const FT_CIRCLE = 8
Public Const FT_SQUARE = 9

Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNHIGHLIGHT = 20
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_MENU = 4
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_WINDOWTEXT = 8

Public ColorSet(256) As Long     ' pallet colors
Public PenWidth As Long

Public Const Pi = 3.1415928
'--------- API's ------------
Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Declare Function SelectClipPath Lib "gdi32" (ByVal hDC As Long, ByVal iMode As Long) As Long
Public Declare Function BeginPath Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function EndPath Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function AbortPath Lib "gdi32" (ByVal hDC As Long) As Long

Declare Function Polyline Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Declare Function PtVisible Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Declare Function Pie Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Declare Function Arc Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Declare Function Chord Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Declare Function FloodFill Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Declare Function SetPolyFillMode Lib "gdi32" (ByVal hDC As Long, ByVal nPolyFillMode As Long) As Long

Public Const FLOODFILLBORDER = 0
Public Const FLOODFILLSURFACE = 1

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCPAINT = &HEE0086    ' (DWORD) dest = source OR dest
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCINVERT = &H660046 ' (DWORD) dest = source XOR dest

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public RC As RECT

Type SIZE
    cx As Long
    cy As Long
End Type

Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Declare Function GetTextExtentPoint Lib "gdi32" Alias "GetTextExtentPointA" (ByVal hDC As Long, ByVal lpszString As String, ByVal cbString As Long, lpSize As SIZE) As Long
Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal wFlags As Long) As Long

Public Const TA_LEFT = 0
Public Const TA_RIGHT = 2
Public Const TA_CENTER = 6
Public Const TA_UPDATECP = 1
Public Const TA_BASELINE = 24

Public Const DT_TOP = &H0
Public Const DT_LEFT = &H0
Public Const DT_BOTTOM = &H8
Public Const DT_CALCRECT = &H400
Public Const DT_CENTER = &H1
Public Const DT_EXPANDTABS = &H40
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_INTERNAL = &H1000
Public Const DT_NOCLIP = &H100
Public Const DT_NOPREFIX = &H800
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_TABSTOP = &H80
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10

''''''''
Type TEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type

Public TM As TEXTMETRIC

Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal E As Long, ByVal o As Long, ByVal w As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal q As Long, ByVal PAF As Long, ByVal F As String) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hDC As Long, lpMetrics As TEXTMETRIC) As Long
Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hDC As Long, ByVal nCount As Long, ByVal lpFacename As String) As Long

'BITMAP

Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type
Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type
Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type
Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As String * 1024 ' Array length is arbitrary; may be changed
End Type

Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dX As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_TOP = 0
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOMOVE = &H2
Public Const wFlags = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const WM_USER = &H400
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1
Public Const EM_LINESCROLL = &HB6
Public Const EM_LINEFROMCHAR = &HC9


Public Const CF_PALETTE = 9
Public Const APICells = 256
Type PALETTEENTRY    '4 Bytes
        peRed As String * 1
        peGreen As String * 1
        peBlue As String * 1
        peFlags As String * 1
End Type
Type LOGPALETTE
  palVersion As Integer 'Windows 3.0 version or higher
  palNumEntries As Integer 'number of color in palette
  palPalEntry(APICells) As PALETTEENTRY 'array of element colors
End Type
Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public CustPal As LOGPALETTE

Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Declare Function SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal HPALETTE As Long, ByVal bForceBackground As Long) As Long
Declare Function RealizePalette Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function CloseClipboard Lib "user32" () As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetCapture Lib "user32" () As Long

Public Const BF_ADJUST = &H2000    ' Calculate the space left over.
Public Const BF_LEFT = &H1
Public Const BF_BOTTOM = &H8
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_DIAGONAL = &H10
Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Public Const BF_FLAT = &H4000      ' For flat rather than 3-D borders.
Public Const BF_MIDDLE = &H800     ' Fill in the middle.
Public Const BF_MONO = &H8000      ' For monochrome borders.
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000      ' Use for softer buttons.
Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)

Public Const BDR_INNER = &HC
Public Const BDR_OUTER = &H3
Public Const BDR_RAISED = &H5
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKEN = &HA
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2


Function StayOnTop(Form As Form)
   Dim lFlags As Long
   Dim lStay As Long

   lFlags = SWP_NOSIZE Or SWP_NOMOVE
   lStay = SetWindowPos(Form.hwnd, HWND_TOPMOST, 0, 0, 0, 0, lFlags)
End Function

Private Sub CheckRGBByte(ByRef RGBb)
   If RGBb < 0 Then RGBb = 0
   If RGBb > 255 Then RGBb = 255
End Sub

Public Sub APIcls(pic As Control, Kl As Long)
   APIrect pic.hDC, 0, 0, Kl, Kl, 0, 0, pic.ScaleWidth, pic.ScaleHeight
End Sub

Public Sub DrawEllipse(hDC As Long, iX1 As Long, iY1 As Long, iX2 As Long, iY2 As Long, ByVal linewidth As Integer, ByVal clrline As Long, ByVal clrfill As Long)
   Dim hPn As Long, hPnOld As Long
   Dim hBr As Long, hBrOld As Long
   
   hPn = CreatePen(0, linewidth, clrline)
   hPnOld = SelectObject(hDC, hPn)
   hBr = CreateSolidBrush(clrfill)
   hBrOld = SelectObject(hDC, hBr)

   Ellipse hDC, iX1, iY1, iX2, iY2
   
   SelectObject hDC, hBrOld
   DeleteObject hBr
   SelectObject hDC, hPnOld
   DeleteObject hPn
End Sub

Public Sub DrawEllipseNoFill(hDC As Long, iX1 As Long, iY1 As Long, iX2 As Long, iY2 As Long, ByVal linewidth As Integer, ByVal clrline As Long, Optional ByVal penstyle As Integer = PS_SOLID)
   Dim hPn As Long, hPnOld As Long
   Dim hBr As Long, hBrOld As Long
   
   hPn = CreatePen(penstyle, linewidth, clrline)
   hPnOld = SelectObject(hDC, hPn)

   Arc hDC, iX1, iY1, iX2, iY2, iX1, iY1, iX1, iY1
   
   SelectObject hDC, hPnOld
   DeleteObject hPn
End Sub

Public Sub DrawRectangle(hDC As Long, ByVal iX1 As Long, ByVal iY1 As Long, ByVal iX2 As Long, ByVal iY2 As Long, ByVal linewidth As Integer, ByVal clrline As Long, ByVal clrfill As Long, Optional ByVal penstyle As Long = 0)
   Dim hPn As Long, hPnOld As Long
   Dim hBr As Long, hBrOld As Long
   
   hPn = CreatePen(penstyle, linewidth, clrline)
   hPnOld = SelectObject(hDC, hPn)
   hBr = CreateSolidBrush(clrfill)
   hBrOld = SelectObject(hDC, hBr)

   Rectangle hDC, iX1, iY1, iX2, iY2
   
   SelectObject hDC, hBrOld
   DeleteObject hBr
   SelectObject hDC, hPnOld
   DeleteObject hPn
End Sub

Public Sub DrawRectangleNoFill(hDC As Long, ByVal iX1 As Long, ByVal iY1 As Long, ByVal iX2 As Long, ByVal iY2 As Long, ByVal linewidth As Integer, ByVal clrline As Long)
   Dim hPn As Long, hPnOld As Long
   Dim hBr As Long, hBrOld As Long
   
   DrawLine hDC, iX1, iY1, iX2, iY1, linewidth, clrline
   DrawLine hDC, iX2, iY1, iX2, iY2, linewidth, clrline
   DrawLine hDC, iX2, iY2, iX1, iY2, linewidth, clrline
   DrawLine hDC, iX1, iY2, iX1, iY1, linewidth, clrline
End Sub

Public Sub DrawLine(ByVal hDC As Long, ByVal iX1 As Long, ByVal iY1 As Long, ByVal iX2 As Long, ByVal iY2 As Long, ByVal linewidth As Integer, ByVal clrline As Long, Optional ByVal penstyle As Integer = PS_SOLID)
   Dim hPn As Long, hPnOld As Long
   Dim hBr As Long, hBrOld As Long
   Dim Pt As POINTAPI

   hPn = CreatePen(penstyle, linewidth, clrline)
   hPnOld = SelectObject(hDC, hPn)

   MoveToEx hDC, iX1, iY1, Pt
   LineTo hDC, iX2, iY2
   
   SelectObject hDC, hPnOld
   DeleteObject hPn
End Sub

Public Sub APIline(hDC As Long, _
                  pentype As Long, _
                  BorderW As Long, BorderKl As Long, _
                  iX1 As Long, iY1 As Long, _
                  iX2 As Long, iY2 As Long)
   Dim Pt As POINTAPI
   Dim hPn As Long, hPnOld As Long
   
   hPn = CreatePen(pentype, BorderW, BorderKl Xor &H1000000)
   hPnOld = SelectObject(hDC, hPn)

   MoveToEx hDC, iX1, iY1, Pt
   LineTo hDC, iX2, iY2

   SelectObject hDC, hPnOld
   DeleteObject hPn
End Sub

Public Sub APIrect(hDC As Long, _
                  pentype As Long, BorderW As Long, _
                  BorderKl As Long, FillKl As Long, _
                  iX1 As Long, iY1 As Long, _
                  iX2 As Long, iY2 As Long)
   Dim hPn As Long, hPnOld As Long
   Dim hBr As Long, hBrOld As Long

   hPn = CreatePen(pentype, BorderW, BorderKl) ' Xor &H1000000)
   hPnOld = SelectObject(hDC, hPn)
   hBr = CreateSolidBrush(FillKl) ' Xor &H1000000)
   hBrOld = SelectObject(hDC, hBr)
   
   Rectangle hDC, iX1, iY1, iX2, iY2
   
   SelectObject hDC, hBrOld
   DeleteObject hBr
   SelectObject hDC, hPnOld
   DeleteObject hPn
End Sub

Public Sub APIrrect(hDC As Long, _
                    iPenWidth As Long, _
                    BorderKl As Long, FillKl As Long, _
                    iX1 As Long, iY1 As Long, _
                    iX2 As Long, iY2 As Long, _
                    iRounding As Long)
   Dim hPn As Long, hPnOld As Long
   Dim hBr As Long, hBrOld As Long

   If FillKl > -1 Then
      hBr = CreateSolidBrush(FillKl Xor &H1000000)
      hBrOld = SelectObject(hDC, hBr)
      End If
   If BorderKl > -1 Then
      hPn = CreatePen(0, iPenWidth, BorderKl Xor &H1000000)
      hPnOld = SelectObject(hDC, hPn)
      End If
      
   RoundRect hDC, iX1, iY1, iX2, iY2, (iX2 - iX1) * iRounding / 100, (iY2 - iY1) * iRounding / 100
   
   If BorderKl > -1 Then
      SelectObject hDC, hPnOld
      DeleteObject hPn
      End If
   If FillKl > -1 Then
      SelectObject hDC, hBrOld
      DeleteObject hBr
      End If
End Sub

Public Sub APIText(pic As Control, ByVal mx As Long, ByVal my As Long, txt As String, Align As Long, Hk As Long, Kleur As Long)
    Dim h As Long, Wt As Long, i As Long, E As Long
    Dim w, o, u, s, c, OP, CP, q, PAF
    Dim F As String
    Dim hFnt As Long, hFntOld As Long
    Dim tdX As Long, tdY As Long
    Dim mX1 As Long, mY1 As Long
   
    SetTextColor pic.hDC, Kleur&
    
    If InStr(txt, vbCrLf) Then
        RC.Left = 0: RC.Top = 0
        DrawText pic.hDC, txt, Len(txt), RC, Align Or DT_CALCRECT
        OffsetRect RC, mx - RC.Right \ 2, my - RC.Bottom \ 2
        DrawText pic.hDC, txt, Len(txt), RC, Align
    Else
        Dim TM As TEXTMETRIC
        Dim sz As SIZE
        
        GetTextMetrics pic.hDC, TM
        
        h = TM.tmHeight
        Wt = TM.tmWeight
        i = TM.tmItalic 'Asc(TM.tmItalic)
        F$ = String(128, " ")
        
        GetTextFace pic.hDC, 128, F$
        
        E = Hk
        
        hFnt = CreateFont(h, w, E, o, Wt, i, u, s, c, OP, CP, q, PAF, F$)
        hFntOld = SelectObject(pic.hDC, hFnt)
        
        GetTextExtentPoint pic.hDC, txt, Len(txt), sz
        
        Select Case Hk
        Case 0
            tdX = sz.cx: tdY = sz.cy
            mX1 = mx - tdX \ 2: mY1 = my - tdY \ 2
        Case 900
            tdX = sz.cy: tdY = sz.cx
            mX1 = mx - tdX \ 2: mY1 = my + tdY \ 2: tdY = -tdY
        Case -900
            tdX = sz.cy: tdY = sz.cx
            mX1 = mx + tdX \ 2: tdX = -tdX: mY1 = my - tdY \ 2
        Case 1800
            tdX = sz.cx: tdY = sz.cy
            mX1 = mx + tdX \ 2: tdX = -tdX: mY1 = my + tdY \ 2: tdY = -tdY
        End Select
        
        TextOut pic.hDC, mx, my, txt, Len(txt)
        
        SelectObject pic.hDC, hFntOld
        DeleteObject hFnt
    End If
End Sub

Public Sub PlotText(hDC As Long, ByVal mx As Long, ByVal my As Long, txt As String, clr As Long, Optional textalignment As Long = TA_LEFT)
    Dim uAlignPrev As Long
    
    SetTextColor hDC, clr&
    
    uAlignPrev = SetTextAlign(hDC, textalignment)
    TextOut hDC, mx, my, txt, Len(txt)
    
    SetTextAlign hDC, uAlignPrev
End Sub

Private Sub HandleTextPlotBoudingBoxMath(tx1 As Long, ty1 As Long, tx2 As Long, ty2 As Long, ByVal escapement As Long, ByVal textalignment As Long, sz As SIZE)
    Select Case escapement
        Case 0
            If textalignment = TA_CENTER Then
                tx1 = tx1 - ((sz.cx + 1) \ 2)
            End If
            If textalignment = TA_RIGHT Then
                tx1 = tx1 - sz.cx
            End If

            tx2 = tx1 + sz.cx
            ty2 = ty1 + sz.cy

        Case 900

        Case 1800

        Case 2700
            tx2 = tx1 - sz.cy
            ty2 = ty1 - sz.cx
    End Select
End Sub

Public Sub PlotRotatedText(ByVal hDC As Long, ByVal txt As String, ByVal x As Single, ByVal y As Single, ByVal txtclr As Long, ByVal lineclr As Long, ByVal fillclr As Long, ByVal drawbox As Boolean, ByVal fillbox As Boolean, ByVal font_name As String, ByVal textsize As Long, ByVal weight As Long, ByVal escapement As Long, ByVal use_italic As Boolean, ByVal use_underline As Boolean, ByVal use_strikethrough As Boolean, Optional textalignment As Long = TA_LEFT)
    Dim uAlignPrev As Long
    Dim newfont As Long
    Dim oldfont As Long
    Dim sz As SIZE
    Dim tx1 As Long, ty1 As Long
    Dim tx2 As Long, ty2 As Long

    newfont = CreateFont(textsize, 0, escapement, escapement, weight, use_italic, use_underline, use_strikethrough, 0, 0, CLIP_LH_ANGLES, 0, 0, font_name)
    oldfont = SelectObject(hDC, newfont)
    
    tx1 = x
    ty1 = y
    tx2 = tx1
    ty2 = ty1
    
    GetTextExtentPoint32 hDC, txt, Len(txt), sz
    
    HandleTextPlotBoudingBoxMath tx1, ty1, tx2, ty2, escapement, textalignment, sz
    
    If drawbox Then
        If fillbox Then
            DrawRectangle hDC, tx1, ty1, tx2, ty2, 1, lineclr, fillclr
        Else
            DrawRectangleNoFill hDC, tx1, ty1, tx2, ty2, 1, lineclr
        End If
    End If
    
    uAlignPrev = SetTextAlign(hDC, textalignment)

    SetTextColor hDC, txtclr

    TextOut hDC, x, y, txt, Len(txt)

    SetTextAlign hDC, uAlignPrev

    newfont = SelectObject(hDC, oldfont)
    DeleteObject newfont
End Sub

Public Sub GetTextPlotBoudingBoxCoords(tx1 As Long, ty1 As Long, tx2 As Long, ty2 As Long, ByVal hDC As Long, ByVal txt As String, ByVal x As Single, ByVal y As Single, ByVal font_name As String, ByVal textsize As Long, ByVal weight As Long, ByVal escapement As Long, ByVal use_italic As Boolean, ByVal use_underline As Boolean, ByVal use_strikethrough As Boolean, Optional textalignment As Long = TA_LEFT)
    Dim newfont As Long
    Dim oldfont As Long
    Dim sz As SIZE

    newfont = CreateFont(textsize, 0, escapement, escapement, weight, use_italic, use_underline, use_strikethrough, 0, 0, CLIP_LH_ANGLES, 0, 0, font_name)
    oldfont = SelectObject(hDC, newfont)
    
    tx1 = x
    ty1 = y
    tx2 = tx1
    ty2 = ty1
    
    GetTextExtentPoint32 hDC, txt, Len(txt), sz
    
    HandleTextPlotBoudingBoxMath tx1, ty1, tx2, ty2, escapement, textalignment, sz
    
    newfont = SelectObject(hDC, oldfont)
    DeleteObject newfont
End Sub


' the kad (frame) pictures must be set to twips mode for this
' routine to work
Public Sub CheckScrolls(pic As Control, _
                        kadI As Control, kadO As Control, _
                        HS As Control, VS As Control)
   kadI.Width = kadO.Width - 60
   VS.Left = kadI.Width - VS.Width + 15
   HS.Width = kadI.Width - VS.Width
   
   kadI.Height = kadO.Height - 60
   HS.Top = kadI.Height - HS.Height + 15
   VS.Height = kadI.Height - HS.Height
   
    If pic.Width > kadI.Width - 90 Then
        kadI.Height = kadO.Height - HS.Height - 90: VS.Height = kadI.Height
        HS.Max = pic.Width - kadI.Width + 45: HS.Visible = True
    Else
        VS.Height = kadI.Height
        HS.Value = HS.Min: HS.Visible = False
    End If
    
    If pic.Height > kadI.Height - 90 Then
        kadI.Width = kadO.Width - VS.Width - 90: HS.Width = kadI.Width
        VS.Max = pic.Height - kadI.Height + 45: VS.Visible = True
    Else
        HS.Width = kadI.Width
        VS.Value = VS.Min: VS.Visible = False
    End If
    
    If pic.Width > kadI.Width - 90 And pic.Height > kadI.Height - 90 Then
        kadI.Height = kadO.Height - HS.Height - 90: VS.Height = kadI.Height
        HS.Max = pic.Width - kadI.Width + 45: HS.Visible = True
        kadI.Width = kadO.Width - VS.Width - 90: HS.Width = kadI.Width
        VS.Max = pic.Height - kadI.Height + 45: VS.Visible = True
    End If
End Sub

Function FixPath(ByVal p As Variant) As String
   If Right(p, 1) = "\" Then FixPath = p Else FixPath = p & "\"
End Function

Public Sub Flood8b(pic As Control, _
                   StartColor As Long, EndColor As Long, _
                   ByVal Stijl As Long)
   Dim Pt As POINTAPI
   Dim KlDis As Long          ' EndColor(Index)-StartColor(Index)
   Dim id As Long             ' color ID counter
   Dim i As Long              ' counter
   Dim St As Single           ' after how much of I (pixels) a color ID has to change
   Dim StK As Single          ' with which amount colorID's are changed
   Dim w As Long, h As Long   ' Width-Height
   Dim d As Long              ' distance in pixels
   Dim x As Long, y As Long
   Dim XX1 As Long, YY1 As Long
   Dim XX2 As Long, YY2 As Long
   
   Dim hPn As Long, hPnOld As Long
   Dim hBr As Long, hBrOld As Long
   
   If StartColor < 16 Then StartColor = 16
   If EndColor < 16 Then EndColor = 255
   KlDis = EndColor - StartColor
   id = StartColor - 16
   APIcls pic, 7
   '
   Select Case Stijl
   
   Case FT_LEFTRIGHT, FT_LEFTRIGHT2  ' West-East
   x = 0: y = -1
   w = pic.ScaleWidth: h = pic.ScaleHeight + 1
   If Stijl = FT_LEFTRIGHT2 Then d = w / 2 Else d = w
   Select Case d
     Case Is > Abs(KlDis): St = (Abs(d / KlDis)): StK = 1
     Case Is < Abs(KlDis): St = 1: StK = (Abs(KlDis / d))
     Case Else: St = 1: StK = 1
   End Select
   If St < 1 Then St = 1
   For i = 0 To w - 1
     hPn = CreatePen(0, 0, ColorSet(16 + id) Xor &H2000000) ' Solid(0), Standard width(0) so 1 pixel
     hPnOld = SelectObject(pic.hDC, hPn)  ' select pen
     MoveToEx pic.hDC, x + i, y, Pt       '
     LineTo pic.hDC, x + i, y + h         ' use pen
     SelectObject pic.hDC, hPnOld         ' (re)set to prev. pen
     DeleteObject hPn                     ' remove new used pen
     If i >= w \ 2 - 1 And Stijl = FT_LEFTRIGHT2 Then
        KlDis = -KlDis: Stijl = FT_LEFTRIGHT
        If w \ 2 <> w / 2 Then id = (240 + id + StK * Sgn(KlDis)) Mod 240
        Else
        If i Mod St = 0 Then id = (240 + id + StK * Sgn(KlDis)) Mod 240
        End If
   Next i
   
   Case FT_UPDOWN, FT_UPDOWN2  'Nord-South
   x = 0: y = 0
   w = pic.ScaleWidth: h = pic.ScaleHeight
   If Stijl = FT_UPDOWN2 Then d = h / 2 Else d = h
   Select Case d
     Case Is > Abs(KlDis): St = (Abs(d / KlDis)): StK = 1
     Case Is < Abs(KlDis): St = 1: StK = (Abs(KlDis / d))
     Case Else: St = 1: StK = 1
   End Select
   If St < 1 Then St = 1
   For i = 0 To h - 1
     hPn = CreatePen(0, 0, ColorSet(16 + id) Xor &H2000000)
     hPnOld = SelectObject(pic.hDC, hPn)
     MoveToEx pic.hDC, x, y + i, Pt
     LineTo pic.hDC, x + w, y + i
     SelectObject pic.hDC, hPnOld
     DeleteObject hPn
     If i >= h \ 2 And Stijl = FT_UPDOWN2 Then
        KlDis = -KlDis: Stijl = FT_UPDOWN
        If h \ 2 <> h / 2 Then id = (240 + id + StK * Sgn(KlDis)) Mod 240
        Else
        If i Mod St = 0 Then id = (240 + id + StK * Sgn(KlDis)) Mod 240
        End If
   Next i
   
   Case FT_ULDR, FT_ULDR2  ' diagonal
   x = 0: y = 0
   w = pic.ScaleWidth: h = pic.ScaleHeight
   If Stijl = FT_ULDR2 Then d = (w + h) \ 2 Else d = w + h
   Select Case d
     Case Is > Abs(KlDis): St = (Abs(d / KlDis)): StK = 1
     Case Is < Abs(KlDis): St = 1: StK = (Abs(KlDis / d))
     Case Else: St = 1: StK = 1
   End Select
   If St < 1 Then St = 1
   i = 0
   Do While i < w + h
     
     If i < w Then XX2 = x + i: YY2 = y Else YY2 = YY2 + 1
     If i < h Then YY1 = y + i: XX1 = x Else XX1 = XX1 + 1
     hPn = CreatePen(0, 2, ColorSet(16 + id) Xor &H2000000)
     hPnOld = SelectObject(pic.hDC, hPn)
     MoveToEx pic.hDC, XX1, YY1, Pt
     LineTo pic.hDC, XX2, YY2
     SelectObject pic.hDC, hPnOld
     DeleteObject hPn
     i = i + 1
     If i = d And Stijl = FT_ULDR2 Then
        KlDis = -KlDis: Stijl = FT_ULDR
        Else
        If i Mod St = 0 Then id = (240 + id + StK * Sgn(KlDis)) Mod 240
        End If
   Loop
   
   Case FT_DLUR, FT_DLUR2  'diagonal2
   x = 0: y = -1
   w = pic.ScaleWidth: h = pic.ScaleHeight + 1
   If Stijl = FT_DLUR2 Then d = (w + h) \ 2 Else d = w + h
   Select Case d
     Case Is > Abs(KlDis): St = (Abs(d / KlDis)): StK = 1
     Case Is < Abs(KlDis): St = 1: StK = (Abs(KlDis / d))
     Case Else: St = 1: StK = 1
   End Select
   If St < 1 Then St = 1
   i = 0
   Do While i < w + h
     If i < w + 1 Then XX2 = x + (w - i): YY2 = y Else YY2 = YY2 + 1
     If i < h Then YY1 = y + i: XX1 = x + w Else XX1 = XX1 - 1
     hPn = CreatePen(0, 0, ColorSet(16 + id) Xor &H2000000)
     hPnOld = SelectObject(pic.hDC, hPn)
     MoveToEx pic.hDC, XX1 - 1, YY1, Pt
     LineTo pic.hDC, XX2 - 1, YY2
     SelectObject pic.hDC, hPnOld
     DeleteObject hPn
     i = i + 1
     If i = d And Stijl = FT_DLUR2 Then
        KlDis = -KlDis: Stijl = FT_DLUR
        Else
        If i Mod St = 0 Then id = (240 + id + StK * Sgn(KlDis)) Mod 240
        End If
   Loop
   
   Case FT_CIRCLE  ' circles
   w = pic.ScaleWidth: h = pic.ScaleHeight
   x = w / 2: y = h / 2
   d = Sqr(x * x + y * y)
   Select Case d
     Case Is > Abs(KlDis): St = (Abs(d / KlDis)): StK = 1
     Case Is < Abs(KlDis): St = 1: StK = (Abs(KlDis / d))
     Case Else: St = 1: StK = 1
   End Select
   If St < 1 Then St = 1
   XX1 = x + d * Cos(Pi)
   YY1 = y + d * Sin(Pi * 3 / 2)
   XX2 = x + d * Cos(0)
   YY2 = y + d * Sin(Pi * 1 / 2)
   While XX1 < XX2 And YY1 < YY2
     hPn = CreatePen(0, 2, ColorSet(16 + id) Xor &H2000000)
     hPnOld = SelectObject(pic.hDC, hPn)
     Ellipse pic.hDC, XX1, YY1, XX2, YY2
     SelectObject pic.hDC, hPnOld
     DeleteObject hPn
     XX1 = XX1 + 1: YY1 = YY1 + 1
     XX2 = XX2 - 1: YY2 = YY2 - 1
     i = i + 1
     If i Mod St = 0 Then id = (240 + id + StK * Sgn(KlDis)) Mod 240
   Wend
   
   Case FT_SQUARE ' rectangle
   x = 0: y = 0
   w = pic.ScaleWidth: h = pic.ScaleHeight
   
   If h > d Then d = h / 2 Else d = w / 2
   Select Case d
     Case Is > Abs(KlDis): St = (Abs(d / KlDis)): StK = 1
     Case Is < Abs(KlDis): St = 1: StK = (Abs(KlDis / d))
     Case Else: St = 1: StK = 1
   End Select
   If St < 1 Then St = 1
   i = 0
   XX1 = 0: YY1 = 0: XX2 = w: YY2 = h
   While XX1 < XX2 And YY1 < YY2
     hPn = CreatePen(0, 2, ColorSet(16 + id) Xor &H2000000)
     hPnOld = SelectObject(pic.hDC, hPn)
     Rectangle pic.hDC, XX1, YY1, XX2, YY2
     SelectObject pic.hDC, hPnOld
     DeleteObject hPn
     XX1 = XX1 + 1: YY1 = YY1 + 1
     XX2 = XX2 - 1: YY2 = YY2 - 1
     i = i + 1
     If i Mod St = 0 Then id = (240 + id + StK * Sgn(KlDis)) Mod 240
   Wend
   
   End Select
   If pic.AutoRedraw = True Then pic.Refresh

End Sub

' obj must be one who knows the .Line method
Public Sub BevelObject(obj As Object, _
                     x1 As Long, y1 As Long, _
                     x2 As Long, y2 As Long, _
                     ByRef Bevel As Long)
   Dim ColorUpper As Long
   Dim ColorLower As Long
   
   If Bevel = 0 Then
      ColorUpper = &H80000014
      ColorLower = &H80000010
      Else
      ColorUpper = &H80000010
      ColorLower = &H80000014
      End If
   obj.Line (x1, y1)-(x2, y1), ColorUpper
   obj.Line (x1, y1)-(x1, y2), ColorUpper
   obj.Line (x2, y1)-(x2, y2), ColorLower
   obj.Line (x1, y2)-(x2, y2), ColorLower
End Sub

Public Function IsColorID(Color As Long, StartID As Long) As Integer
   Dim i As Long
   
   For i = StartID To 255
   If ColorSet(i) = Color Then Exit For
   Next i
   IsColorID = i
End Function

Public Sub Pause(ByVal Pze As Long)
   Dim mTime As Variant
   mTime = Timer
   While Timer - mTime < Pze / 1000: DoEvents: Wend
End Sub

Public Sub DrawFormula(pic As Control, _
                      x1, y1, dX, dy, _
                      Pts, GPlus, _
                      angle, Kl&, Vul)
   Dim hDC As Long
   Dim hPn As Long, hPnOld As Long
   Dim hBr As Long, hBrOld As Long
   Dim xs As Long, ys As Long
   Dim xc As Long, yc As Long
   ReDim Pt(Pts) As POINTAPI
   Dim p As Long, g As Long
   
   hDC = pic.hDC
   
   On Error GoTo TFrmFout:
   dX = dX - 1: dy = dy - 1
   xs = dX / 2: xc = x1 + xs
   ys = dy / 2: yc = y1 + ys
   
   For p = 0 To Pts - 1
      g = (g + GPlus) Mod 360
      Pt(p).x = xc + xs * Cos((g + angle) * Pi / 180)
      Pt(p).y = yc + ys * Sin((g + angle) * Pi / 180)
   Next p
   hPn = CreatePen(0, PenWidth, Kl& Xor &H2000000)
   hPnOld = SelectObject(hDC, hPn)
   hBr = CreateSolidBrush(Kl& Xor &H20000000)
   hBrOld = SelectObject(hDC, hBr)
   If Vul = 1 Then
      Polygon hDC, Pt(0), Pts
      Else
      Polyline hDC, Pt(0), Pts
      End If
   SelectObject hDC, hBrOld
   DeleteObject hBr
   SelectObject hDC, hPnOld
   DeleteObject hPn
   
   If pic.AutoRedraw = True Then pic.Refresh
TFrmEinde:
   Exit Sub
TFrmFout:
   MsgBox "Fout:" & Str(Err) & vbCrLf & Error$
   Resume TFrmEinde:
End Sub

Public Function ptx(ByVal x As Integer) As Integer
    ptx = x * Screen.TwipsPerPixelX
End Function

Public Function pty(ByVal y As Integer) As Integer
    pty = y * Screen.TwipsPerPixelY
End Function

Public Function pxt(ByVal x As Variant) As Integer
    pxt = x / Screen.TwipsPerPixelX
End Function

Public Function pyt(ByVal y As Variant) As Integer
    pyt = y / Screen.TwipsPerPixelY
End Function

Public Sub DrawArrow(ByVal hDC As Long, ByVal cx As Long, ByVal cy As Long, _
                        ByVal angle As Long, ByVal dX As Long, ByVal dy As Long, ByVal ArrWidth As Long, _
                        ByVal HeadW As Single, ByVal DoubleArr As Boolean, ByVal linewidth As Integer, ByVal clrline As Long, ByVal clrfill As Long)

   Dim hPn As Long, hPnOld As Long
   Dim hBr As Long, hBrOld As Long
   
   Dim Xc1 As Long, Yc1 As Long
   Dim Xc2 As Long, Yc2 As Long
   Dim Xc3 As Long, Yc3 As Long
   Dim AngleXA As Single, AngleYA As Single  ' back-side
   Dim AngleXL As Single, AngleYL As Single  ' turn to the left
   Dim AngleXR As Single, AngleYR As Single  ' turn to the right
   Dim AngleXV As Single, AngleYV As Single  ' fore-side
   Dim StraalAx As Long, StraalAy As Long
   Dim Pt(12) As POINTAPI, AaPt As Long
   Dim Deg As Single
   
   Deg = (Atn(1) * 4) / 180
   angle = (360 + 180 - angle) Mod 360
   AngleXA = Cos(angle * Deg):         AngleYA = Sin(angle * Deg)
   AngleXV = Cos((angle + 180) * Deg): AngleYV = Sin((angle + 180) * Deg)
   AngleXL = Cos((angle - 90) * Deg):  AngleYL = Sin((angle - 90) * Deg)
   AngleXR = Cos((angle + 90) * Deg):  AngleYR = Sin((angle + 90) * Deg)
   StraalAx = dX: StraalAy = dy
   If ArrWidth > dX Then ArrWidth = dX
   If ArrWidth > dy Then ArrWidth = dy
   If HeadW < ArrWidth Then HeadW = ArrWidth
   If ArrWidth > HeadW Then ArrWidth = HeadW
   If dX < HeadW Then
      HeadW = dX
      dX = 0
      Else
      dX = dX - HeadW
      End If
   If dy < HeadW Then
      HeadW = dy
      dy = 0
      Else
      dy = dy - HeadW
      End If
      
   If DoubleArr = True Then StraalAx = dX: StraalAy = dy
   
   Xc1 = cx + dX * AngleXV
   Yc1 = cy + dy * AngleYV
   Xc2 = cx + StraalAx * AngleXA
   Yc2 = cy + StraalAy * AngleYA
   Xc3 = cx - dX * AngleXV
   Yc3 = cy - dy * AngleYV
   
   Pt(0).x = Xc2 + ArrWidth * AngleXL
   Pt(0).y = Yc2 + ArrWidth * AngleYL
   Pt(1).x = Xc1 + ArrWidth * AngleXL
   Pt(1).y = Yc1 + ArrWidth * AngleYL
   Pt(2).x = Xc1 + HeadW * AngleXL
   Pt(2).y = Yc1 + HeadW * AngleYL
   Pt(3).x = Xc1 + HeadW * AngleXV
   Pt(3).y = Yc1 + HeadW * AngleYV
   Pt(4).x = Xc1 + HeadW * AngleXR
   Pt(4).y = Yc1 + HeadW * AngleYR
   Pt(5).x = Xc1 + ArrWidth * AngleXR
   Pt(5).y = Yc1 + ArrWidth * AngleYR
   Pt(6).x = Xc2 + ArrWidth * AngleXR
   Pt(6).y = Yc2 + ArrWidth * AngleYR
   If DoubleArr = False Then
      AaPt = 8
      Pt(7).x = Xc2 + ArrWidth * AngleXL
      Pt(7).y = Yc2 + ArrWidth * AngleYL
      Else
      AaPt = 12
      Pt(7).x = Xc3 - ArrWidth * AngleXL
      Pt(7).y = Yc3 - ArrWidth * AngleYL
      Pt(8).x = Xc3 - HeadW * AngleXL
      Pt(8).y = Yc3 - HeadW * AngleYL
      Pt(9).x = Xc3 - HeadW * AngleXV
      Pt(9).y = Yc3 - HeadW * AngleYV
      Pt(10).x = Xc3 - HeadW * AngleXR
      Pt(10).y = Yc3 - HeadW * AngleYR
      Pt(11).x = Xc3 - ArrWidth * AngleXR
      Pt(11).y = Yc3 - ArrWidth * AngleYR
   End If
   
   hPn = CreatePen(0, linewidth, clrline)
   hPnOld = SelectObject(hDC, hPn)
   hBr = CreateSolidBrush(clrfill)
   hBrOld = SelectObject(hDC, hBr)
   
   SetPolyFillMode hDC, FLOODFILLSURFACE
   Polygon hDC, Pt(0), AaPt
   
   SelectObject hDC, hBrOld
   DeleteObject hBr
   SelectObject hDC, hPnOld
   DeleteObject hPn
End Sub

Public Sub DrawFilledPolygon(ByVal hDC As Long, ByVal Total As Integer, points() As POINTAPI, linewidth As Integer, clrline As Long, clrfill As Long)
   Dim hPn As Long, hPnOld As Long
   Dim hBr As Long, hBrOld As Long

   hPn = CreatePen(0, linewidth, clrline)
   hPnOld = SelectObject(hDC, hPn)
   hBr = CreateSolidBrush(clrfill)
   hBrOld = SelectObject(hDC, hBr)

   SetPolyFillMode hDC, FLOODFILLSURFACE
   Polygon hDC, points(0), Total

   SelectObject hDC, hBrOld
   DeleteObject hBr
   SelectObject hDC, hPnOld
   DeleteObject hPn
End Sub

Public Sub DrawPolyline(ByVal hDC As Long, ByVal Total As Integer, points() As POINTAPI, linewidth As Integer, clrline As Long)
   Dim hPn As Long, hPnOld As Long

   hPn = CreatePen(0, linewidth, clrline)
   hPnOld = SelectObject(hDC, hPn)

   SetPolyFillMode hDC, FLOODFILLSURFACE
   Polyline hDC, points(0), Total

   SelectObject hDC, hPnOld
   DeleteObject hPn
End Sub


