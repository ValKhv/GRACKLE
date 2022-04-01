VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************************************
'*********************************************************************************************************************
'*********************************************************************************************************************
' DESCRIPTION: API Wrapper to draw Text Directly on Desktop or Form
' VERSION    : v 0.2 30.01.2015
' COPYRIGHT  : V.Khvatov DigitalXpert
'*********************************************************************************************************************
'*********************************************************************************************************************
'*********************************************************************************************************************
Option Compare Database
Option Explicit




'######################################################################################################################################################


#If VBA7 Then
        '************************************************************************
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As LongPtr
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetWindowDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Boolean
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
Private Declare PtrSafe Function BitBlt Lib "gdi32" (ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare PtrSafe Function InvalidateRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT, ByVal bErase As Long) As Long

'----------------------------------
Private Declare PtrSafe Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As LongPtr, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
Private Declare PtrSafe Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As LongPtr
Private Declare PtrSafe Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As LongPtr, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare PtrSafe Function DrawText2 Lib "user32" Alias "DrawTextW" (ByVal hdc As LongPtr, ByVal lpStr As LongPtr, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long


Private Declare PtrSafe Function GetTextColor Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetTextColor Lib "gdi32" (ByVal hdc As LongPtr, ByVal crColor As Long) As Long
Private Declare PtrSafe Function GetBkColor Lib "gdi32" (ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function SetBkColor Lib "gdi32" (ByVal hdc As LongPtr, ByVal crColor As Long) As Long
Private Declare PtrSafe Function SetBkMode Lib "gdi32" (ByVal hdc As LongPtr, ByVal nBkMode As Long) As Long

Private Declare PtrSafe Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long, ByVal x3 As Long, ByVal Y3 As Long) As LongPtr
Private Declare PtrSafe Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As LongPtr
Private Declare PtrSafe Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As LongPtr
Private Declare PtrSafe Function CombineRgn Lib "gdi32" (ByVal hDestRgn As LongPtr, ByVal hSrcRgn1 As LongPtr, ByVal hSrcRgn2 As LongPtr, ByVal nCombineMode As Long) As Long
Private Declare PtrSafe Function FillRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr, ByVal hBrush As LongPtr) As Long
Private Declare PtrSafe Function FrameRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr, ByVal hBrush As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare PtrSafe Function RoundRect Lib "gdi32" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long, ByVal x3 As Long, ByVal Y3 As Long) As Long
Private Declare PtrSafe Function Ellipse Lib "gdi32" (ByVal hdc As LongPtr, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long

Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
Private Declare PtrSafe Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As LongPtr
Private Declare PtrSafe Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As LongPtr
        '-------------------------------------------------
        Private hWnd As LongPtr                                             ' Указатель на окно
        Private m_hDC As LongPtr                                            ' Графический контекст
        Private m_NewFont As LongPtr                                        ' Указатель на новый шрифт
        Private m_OldFont As LongPtr                                        ' Указатель на старый шрифт
        Private m_bufferDC As LongPtr                                       ' Указатель на гр контекст буфера
        Private m_BMP As LongPtr                                            ' Указатель на bitmap в памяти
        Private m_BMPOLD As LongPtr                                         ' Указатель на bitmap в памяти
        Private m_hBrsh As LongPtr                                          ' Указатель на кисть
        Private m_hPen As LongPtr                                           ' Указатель на перо
        
        Private hFnt As LongPtr
        Private hFntOld As LongPtr

        '************************************************************************
#Else

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Boolean
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long

'----------------------------------
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawText2 Lib "user32" Alias "DrawTextW" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long


Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal x3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal x3 As Long, ByVal Y3 As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
        
        '--------------------------------------------------
        Private hWnd As Long                                                 ' Указатель на окно
        Private m_hDC As Long                                                ' Графический контекст
        Private m_NewFont As Long                                            ' Указатель на новый шрифт
        Private m_OldFont As Long                                            ' Указатель на старый шрифт
        Private m_bufferDC As Long                                           ' Указатель на гр контекст буфера
        Private m_BMP As Long                                                ' Указатель на bitmap в памяти
        Private m_BMPOLD As Long                                             ' Указатель на bitmap в памяти
        
        Private hFnt As Long
        Private hFntOld As Long
        '************************************************************************
#End If

'****************************************************************************************************
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90
Private Const TwipsPerInch = 1440

Private Const SM_CXFULLSCREEN = 16
Private Const SM_CYFULLSCREEN = 17
' КОНСТАНТЫ, УПРАВЛЯЮЩИЕ ПРОЗРАЧНОСТЬЮ
Private Const OPAQUE = 2
Private Const TRANSPARENT = 1
'***********************************************************************************************
' КОНСТАНТЫ ШРИФТА
'Private Const DESIREDFONTSIZE = 12
Private Const LF_FACESIZE = 32
Private Const RGN_OR = 2
Private Const GWL_WNDPROC = -4
Private Const GWL_STYLE = (-16)
Private Const WM_USER = &H400
Private Const WM_PAINT = &HF
Private Const WM_PRINT = &H317

Private Const ANSI_CHARSET = 0                                              ' ANSI character set
Private Const RUSSIAN_CHARSET = 204                                         ' Для русского
Private Const CLIP_CHARACTER_PRECIS = 1                                     ' clip one character at a time, OR'd
Private Const CLIP_LH_ANGLES = 16                                           ' Needed for tilted fonts, OR'd
Private Const FONT_CLIP_PRECISION = CLIP_LH_ANGLES
Private Const FONT_ITALIC = 0                                               ' Italics T/F
Private Const FONT_NAME = "Arial"                                           ' typeface/font name
Private Const FONT_OUTPUT_PRECISION = 0                                     ' don't care
Private Const FONT_ORIENTATION = 0                                          ' don't care, Windows assumes same as escapement
Private Const FONT_PITCH_N_FAMILY = 0                                       ' don't care, let font mapper pick easily
Private Const FONT_QUALITY = 0                                              ' don't care, let font mapper pick easily
Private Const FONT_STRIKEOUT = 0                                            ' Strikeout T/F
Private Const FONT_UNDERLINE = 0                                            ' Underline T/F
Private Const FONT_SIZE As Integer = 18
Private Const FONT_BOLD As Boolean = False

Private Const TTS_ALWAYSTIP = 1
Private Const TTF_IDISHWND = 1
Private Const TTF_CENTERTIP = 2
Private Const TTF_RTLREADING = 4
Private Const TTF_SUBCLASS = &H10
Private Const TTF_TRACK = &H20
Private Const TTF_ABSOLUTE = &H80
Private Const TTF_TRANSPARENT = &H100
Private Const TTF_DI_SETITEM = &H8000

Private Const WS_BORDER = &H800000
Private Const SW_SHOWNA = 8

Private Const ALTERNATE = 1
Private Const BLACK_BRUSH = 4
Private Const DKGRAY_BRUSH = 3
 
Private Const CW_USEDEFAULT = &H80000000
Private Const HORZRES = 8                                                   '  Horizontal width in pixels
Private Const VERTRES = 10                                                  '  Vertical width in pixels
Private Const ASPECTX        As Long = 40    'Length of the X leg
Private Const ASPECTY        As Long = 42    'Length of the Y leg

Private Const FW_BOLD = 700                                                 ' Bold weight
Private Const FW_DONTCARE = 0                                               ' don't care, let font mapper use default weight
Private Const FW_NORMAL = 400                                               ' Normal weight
Private Const FW_HEAVY = 900
Private Const FW_SEMIBOLD = 600
Private Const FW_BLACK = FW_HEAVY
Private Const FW_DEMIBOLD = FW_SEMIBOLD
Private Const FW_EXTRABOLD = 800
Private Const FW_EXTRALIGHT = 200
Private Const FW_LIGHT = 300
Private Const FW_MEDIUM = 500
Private Const FW_REGULAR = FW_NORMAL
Private Const FW_THIN = 100
Private Const FW_ULTRABOLD = FW_EXTRABOLD
Private Const FW_ULTRALIGHT = FW_EXTRALIGHT

Private Const FONT_WEIGHT = FW_NORMAL
Private Const FONT_WIDTH = 0                                                ' don't care, let font mapper pick according to height

Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private Const iOffSet = 8                                                   ' Сдвиг при отрисовки контура
Private Const HWND_DESKTOP As Long = 0                                      ' Дескриптор экрана

Private Const TTM_ACTIVATE = WM_USER + 1
Private Const TTM_SETDELAYTIME = WM_USER + 3
Private Const TTM_ADDTOOL = WM_USER + 4
Private Const TTM_DELTOOL = WM_USER + 5
Private Const TTM_NEWTOOLRECT = WM_USER + 6
Private Const TTM_RELAYEVENT = WM_USER + 7

' DRAWTEXT CONSTANT
Private Const DT_TOP = &H0
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_VCENTER = &H4
Private Const DT_BOTTOM = &H8
Private Const DT_WORDBREAK = &H10
Private Const DT_SINGLELINE = &H20
Private Const DT_EXPANDTABS = &H40
Private Const DT_TABSTOP = &H80
Private Const DT_NOCLIP = &H100
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_CALCRECT = &H400
Private Const DT_NOPREFIX = &H800
Private Const DT_INTERNAL = &H1000
Private Const DT_EDITCONTROL = &H2000
Private Const DT_END_ELLIPSIS = &H8000
Private Const DT_PATH_ELLIPSIS = &H4000
Private Const DT_WORD_ELLIPSIS = &H40000

' СТИЛИ КИСТИ И ПЕРА
Private Const BS_SOLID = 0
Private Const BS_HATCHED = 2

Private Const HS_HORIZONTAL = 0
Private Const HS_VERTICAL = 1
Private Const HS_FDIAGONAL = 2
Private Const HS_BDIAGONAL = 3
Private Const HS_CROSS = 4
Private Const HS_DIAGCROSS = 5

Private Const PS_SOLID = 0
Private Const PS_DASH = 1
Private Const PS_DOT = 2
Private Const PS_DASHDOT = 3
Private Const PS_DASHDOTDOT = 4

Private Const FF_DONTCARE = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const DEFAULT_CHARSET = 1

'******************************************************************************************************************

'******************************************************************************************************************************************************
Private Type RECT                                                             ' Прямоугольник для отрисовки
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'------------------------------------
Private Type POINTAPI                                                          ' Указатель - точка
        x As Long
        y As Long
End Type
'-------------------------------------
Private Type LOGFONT                                                           ' Структура логического шрифта
       lfHeight As Long
       lfWidth As Long
       lfEscapement As Long
       lfOrientation As Long
       lfWeight As Long
       lfItalic As Byte
       lfUnderline As Byte
       lfStrikeOut As Byte
       lfCharSet As Byte
       lfOutPrecision As Byte
       lfClipPrecision As Byte
       lfQuality As Byte
       lfPitchAndFamily As Byte
       lfFaceName(1 To LF_FACESIZE) As Byte        'lfFaceName As String * LF_FACESIZE
End Type
Private Type SIZE                                  ' Структура для расчета размера строки
    CX As Long
    CY As Long
End Type
Private Type LOGBRUSH                              ' Структура - кисть
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type
Private Type LOGPEN                                ' Структура - перо
        lopnStyle As Long
        lopnWidth As POINTAPI
        lopnColor As Long
End Type

'######################################################################################################################################################
Private m_Rect As RECT                          ' Прямоугольник, ограничивающий метку (заданный)
Private m_EffRect As RECT                       ' Действительный прямоугольник, в котором происходит отрисовка
Private m_Font As StdFont                       ' Стандартный шрифт
Private m_Holst As HolstType                    ' Флаг, определяющий, какое из устройств стоит использовать
                                                ' 0 - экран, 1 - активное окно, 2 - заданную форму
Private m_Text As String                        ' Значение
Private m_ForeColor As Long                     ' Цвет текста
Private m_BackColor As Long                     ' Цвет подложки, если она не прозрачна
Private m_Transparency As Boolean               ' Прозрачный фон
Private m_RecreateBuffer As Integer             ' Переменная управляет перерисовкой буфера
Private m_Angle As Long                         ' Угол поворота текста
Private m_RecreateFont As Boolean               ' Флаг пересоздания шрифта
Private m_ALLCENTER As Boolean                  ' Флаг, указывающий, что весь прямоугольник метки надо центрировать по контексту
Private m_AUTOSIZE As Boolean                   ' Флаг, указывающий, что весь прямоугольник может быть подобран на основании заданной строки
Private m_BorderColor As Long                   ' Цвет границы, если она не прозрачна
Private m_BorderTransparent As Boolean          ' Прозрачность границы
Private m_BorderWidth As Integer                ' Ширина границы
Private m_BackStyle As Integer                  ' Стиль заднего фона
Private m_BorderStyle As Integer                ' Стиль границы (сплошная, штрих..)
Private m_PaintStyle As Integer                 ' Стиль закраски
Private m_TextBgrd As Boolean                   ' Флаг позволяет управлять подложкой самого текста (отдельно от прямоугольника
                                                ' По умолчанию - выключено, то есть сам текст - прозрачный)
Private m_SingleLine As Boolean                 ' Переменная управляет выводом текста в одну линию
Private m_CHARSET As Integer                    ' Default Charset
Private m_TextAligne As Integer                 ' Выравнивание по горизонтале (-1 - влево, 0 - центр, 1 - вправо), по умолчанию - 0
Private m_TextVAligne As Integer                ' Выравнивание по вертикале (-1 - вниз, 0 - центр, 1 - на верх), по умолчанию - 0
Private m_TextTuncated As Boolean               ' Обрезание текста
Private m_TextDrawMode As Integer               ' Рисование текста (0 - станд DrawText, 2 - TextDraw2 и по ссылке)
Private bAllocateDC As Boolean                  ' DC определено
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'============================================================================================================================================
' Charset
'===========================================================================================================================================
Public Property Get CHARSET() As Integer
    CHARSET = m_CHARSET
End Property
Public Property Let CHARSET(iCharset As Integer)
    If iCharset > 255 Then Err.Raise 1000, , "Wrong charset number"
    m_CHARSET = iCharset
End Property

'============================================================================================================================================
' Text DrawMode
'============================================================================================================================================
Public Property Get TextDrawMode() As Integer
    TextDrawMode = m_TextDrawMode   ' 0 by default
End Property
Public Property Let TextDrawMode(iTextDrawMode As Integer)
    m_TextDrawMode = iTextDrawMode  ' Should be 0 or 2
End Property

'============================================================================================================================================
' Свойство: Наименование шрифта
'===========================================================================================================================================
Public Property Get FontName() As String
     FontName = m_Font.Name
End Property
Public Property Let FontName(sFontName As String)
     m_Font.Name = sFontName
End Property
'============================================================================================================================================
' Свойство: Размер шрифта
'===========================================================================================================================================
Public Property Get FontSize() As Integer
     FontSize = m_Font.SIZE
End Property
Public Property Let FontSize(iFontSize As Integer)
     m_Font.SIZE = iFontSize
End Property
'============================================================================================================================================
' Свойство: Жирный шрифта
'===========================================================================================================================================
Public Property Get FontBold() As Boolean
     FontBold = m_Font.Bold
End Property
Public Property Let FontBold(bBOLD As Boolean)
     m_Font.Bold = bBOLD
End Property
'============================================================================================================================================
' Свойство: Наклонный шрифта
'===========================================================================================================================================
Public Property Get FontItalic() As Boolean
     FontItalic = m_Font.Italic
End Property
Public Property Let FontItalic(bITALIC As Boolean)
     m_Font.Italic = bITALIC
End Property
'============================================================================================================================================
' Свойство: холст (0 - экран, 1 - активное окно, 2 - заданная форма)
'===========================================================================================================================================
Public Property Get Holst() As HolstType
     Holst = m_Holst
End Property
Public Property Let Holst(iHolst As HolstType)
     m_Holst = iHolst
End Property
'============================================================================================================================================
' Свойство: граф контекст DC (только для внешнего, иначе - активное окно)
'===========================================================================================================================================
Public Property Get aDC() As Variant
     aDC = CVar(m_hDC)
End Property
Public Property Let aDC(iDC As Variant)
#If Win64 Then
    m_hDC = CLngPtr(iDC)
#Else
    m_hDC = CLng(iDC)
#End If
    m_Holst = 2            ' Раз установили DC - сразу переключаем
End Property
'============================================================================================================================================
' Свойство: дескриптор окна
'===========================================================================================================================================
Public Property Get ahWnd() As Variant
     ahWnd = CVar(hWnd)
End Property
Public Property Let ahWnd(iHwnd As Variant)
#If Win64 Then
    hWnd = CLngPtr(iHwnd)
#Else
    hWnd = CLng(iHwnd)
#End If
End Property
'============================================================================================================================================
' Свойство: текст
'===========================================================================================================================================
Public Property Get Text() As String
     Text = m_Text
End Property
Public Property Let Text(sText As String)
     m_Text = sText
End Property
'====================================================================================================================
' Свойство: прозрачность
'====================================================================================================================
Public Property Get Transparency() As Boolean
   Transparency = m_Transparency
End Property
Public Property Let Transparency(ByVal bTransparency As Boolean)
   m_Transparency = bTransparency
End Property
'============================================================================================================================================
' Свойство: цвет переднего плана
'===========================================================================================================================================
Public Property Get ForeColor() As Long
     ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(iColor As Long)
     m_ForeColor = iColor
End Property
'=======================================================================================================================================================
' Свойство: цвет подложки (если не прозрачна)
'=======================================================================================================================================================
Public Property Get BackColor() As Long
   BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal iColor As Long)
   m_BackColor = iColor
End Property
'=======================================================================================================================================================
' Свойство: X позиция левого верхнего угла
'=======================================================================================================================================================
Public Property Get Left() As Long
   Left = m_Rect.Left
End Property
Public Property Let Left(ByVal lLeft As Long)
Dim lWidth As Long
   lWidth = m_Rect.Right - m_Rect.Left
   m_Rect.Left = lLeft
   m_Rect.Right = lLeft + lWidth
End Property
'=======================================================================================================================================================
' Свойство: Y позиция левого верхнего угла
'=======================================================================================================================================================
Public Property Get Top() As Long
   Top = m_Rect.Top
End Property
Public Property Let Top(ByVal lTop As Long)
Dim lHeigh As Long
   lHeigh = m_Rect.Bottom - m_Rect.Top
   m_Rect.Top = lTop
   m_Rect.Bottom = lTop + lHeigh
End Property
'=======================================================================================================================================================
' Свойство: Ширина прямоугольника
'=======================================================================================================================================================
Public Property Get Width() As Long
   Width = m_Rect.Right - m_Rect.Left
End Property
Public Property Let Width(ByVal lWidth As Long)
   m_Rect.Right = m_Rect.Left + lWidth
End Property
'=======================================================================================================================================================
' Свойство: Высота прямоугольника
'=======================================================================================================================================================
Public Property Get height() As Long
   height = m_Rect.Bottom - m_Rect.Top
End Property
Public Property Let height(ByVal lHeigh As Long)
   m_Rect.Bottom = m_Rect.Top + lHeigh
End Property
'=======================================================================================================================================================
' Свойство: центрирование относительно контекста
'=======================================================================================================================================================
Public Property Get Center() As Boolean
   Center = m_ALLCENTER
End Property
Public Property Let Center(ByVal bCenter As Boolean)
   m_ALLCENTER = bCenter
End Property
'=======================================================================================================================================================
' Свойство: центрирование относительно контекста
'=======================================================================================================================================================
Public Property Get AutoSize() As Boolean
   AutoSize = m_AUTOSIZE
End Property
Public Property Let AutoSize(ByVal bAutosize As Boolean)
   m_AUTOSIZE = bAutosize
End Property
'=======================================================================================================================================================
' Свойство: цвет границы
'=======================================================================================================================================================
Public Property Get BorderColor() As Long
   BorderColor = m_BorderColor
End Property
Public Property Let BorderColor(ByVal iColor As Long)
   m_BorderColor = iColor
End Property
'=======================================================================================================================================================
' Свойство: прозрачность границы
'=======================================================================================================================================================
Public Property Get BorderTransparent() As Boolean
   BorderTransparent = m_BorderTransparent
End Property
Public Property Let BorderTransparent(ByVal bTransparent As Boolean)
   m_BorderTransparent = bTransparent
End Property
'=======================================================================================================================================================
' Свойство: Ширина границы
'=======================================================================================================================================================
Public Property Get BorderWidth() As Integer
   BorderWidth = m_BorderWidth
End Property
Public Property Let BorderWidth(ByVal IWidth As Integer)
   m_BorderWidth = IWidth
End Property
'=======================================================================================================================================================
' Свойство: Стиль заднего фона
'=======================================================================================================================================================
Public Property Get BackStyle() As Integer
   BackStyle = m_BackStyle
End Property
Public Property Let BackStyle(ByVal iStyle As Integer)
   m_BackStyle = iStyle
End Property
'=======================================================================================================================================================
' Свойство: Стиль границы (1-4)
'=======================================================================================================================================================
Public Property Get BorderStyle() As Integer
   BorderStyle = m_BorderStyle
End Property
Public Property Let BorderStyle(ByVal iStyle As Integer)
   m_BorderStyle = iStyle
   If m_BorderStyle < 0 Then m_BorderStyle = 0
   If m_BorderStyle > 4 Then m_BorderStyle = 4
End Property
'=======================================================================================================================================================
' Свойство: Стиль закраски
'=======================================================================================================================================================
Public Property Get PaintStyle() As Integer
   PaintStyle = m_PaintStyle
End Property
Public Property Let PaintStyle(ByVal iStyle As Integer)
   m_PaintStyle = iStyle
   If m_PaintStyle < 0 Then m_PaintStyle = 0
   If m_PaintStyle > 8 Then m_PaintStyle = 8
End Property
'=======================================================================================================================================================
' Свойство: подложка текста (выводится независимо от прямоугольника, ограничивающего метку)
'=======================================================================================================================================================
Public Property Get TextBackground() As Boolean
   TextBackground = m_TextBgrd
End Property
Public Property Let TextBackground(ByVal bBckgrd As Boolean)
   m_TextBgrd = bBckgrd
End Property
'=======================================================================================================================================================
' Свойство: Текст выводится в одну строку
'=======================================================================================================================================================
Public Property Get SingleLine() As Boolean
   SingleLine = m_SingleLine
End Property
Public Property Let SingleLine(ByVal bSingleLine As Boolean)
   m_SingleLine = bSingleLine
End Property
'=======================================================================================================================================================
' Свойство: Выравнивание текста по горизонтале (-1 влево, 0 - середина, 1 - вправо)
'=======================================================================================================================================================
Public Property Get TextAligne() As Integer
   TextAligne = m_TextAligne
End Property
Public Property Let TextAligne(ByVal iAligne As Integer)
   m_TextAligne = iAligne
   If m_TextAligne < -1 Then m_TextAligne = -1
   If m_TextAligne > 1 Then m_TextAligne = 1
End Property
'=======================================================================================================================================================
' Свойство: Выравнивание текста по вертикале (-1 вниз, 0 - середина, 1 - вверх)
'=======================================================================================================================================================
Public Property Get TextVAligne() As Integer
   TextVAligne = m_TextVAligne
End Property
Public Property Let TextVAligne(ByVal iVAligne As Integer)
   m_TextVAligne = iVAligne
   If m_TextVAligne < -1 Then m_TextVAligne = -1
   If m_TextVAligne > 1 Then m_TextVAligne = 1
End Property
'=======================================================================================================================================================
' Свойство: Обрезать текст
'=======================================================================================================================================================
Public Property Get TextTruncate() As Boolean
   TextTruncate = m_TextTuncated
End Property
Public Property Let TextTruncate(ByVal bTruncate As Boolean)
   m_TextTuncated = bTruncate
End Property

'*********************************************************************************************************************************************
'*********************************************************************************************************************************************
'*********************************************************************************************************************************************
'*********************************************************************************************************************************************
'*********************************************************************************************************************************************
'*********************************************************************************************************************************************
'*********************************************************************************************************************************************
'*********************************************************************************************************************************************
'=======================================================================================================================================================
' Отрисовка метки на экране с установленным
'=======================================================================================================================================================
Public Sub Draw()
Dim bRes As Boolean                 ' Возвращаемый результат
     
'----------------------------------------------------------------
' ПОДГОТОВКА ПОДГОТОВКА ПОДГОТОВКА ПОДГОТОВКА ПОДГОТОВКА ПОДГОТОВКА ПОДГОТОВКА ПОДГОТОВКА ПОДГОТОВКА ПОДГОТОВКА
     If Not bAllocateDC Then                                               ' ПОЛУЧАЕМ hDC
       Call AllocateDC
     End If
     '------------------------------
     m_EffRect = CalculateEffRect(m_Text, m_hDC, m_Rect, m_Font, _
                                                m_AUTOSIZE, m_ALLCENTER) ' Высчитали эффектиный прямоугольник (куда пишем)
     '------------------------------
     bRes = ProcessBUfer(m_hDC, m_EffRect)       ' Сохранили (или наоборот восстановили буфер)
     If Not bRes Then m_Transparency = False     ' Не удалось скопировать буфер, поэтому рисуем только с цветом подлжки
'---------------------------------------------------------------------------
' ОТРИСОВКА ОТРИСОВКА ОТРИСОВКА ОТРИСОВКА ОТРИСОВКА ОТРИСОВКА ОТРИСОВКА ОТРИСОВКА ОТРИСОВКА ОТРИСОВКА
     bRes = DrawBack(CVar(m_hDC), m_EffRect, m_BorderColor, m_BackColor, m_BorderTransparent, m_Transparency, m_BorderWidth, _
                   m_BorderStyle, m_PaintStyle, m_BackStyle)                           '!!!!!!!!!!!!!!!!!!!!!!!!!!!!! ВЫВОДИМ ПОВЕРХНОСТЬ
     bRes = DRAWSTRING(m_Text, CVar(m_hDC), m_EffRect, m_Font, m_ForeColor, m_BackColor, _
                                              m_Transparency, m_TextBgrd, GetTextStyle())
End Sub
'=======================================================================================================================================================
' Функция восстанавливает подложку (то, что было сохранено в начале)
'=======================================================================================================================================================
Public Sub Delete()
Dim bRes As Boolean
    'InvalidateRect hwnd, m_EffRect, True
    bRes = RestoreBGFromBuffer(CVar(m_hDC), m_EffRect)
End Sub
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
'-------------------------------------------
Set m_Font = New StdFont                 ' Шрифт
       m_RecreateBuffer = 0              ' Текст еще не нанесен, под меткой чистая область
       m_Rect.Left = 0                   ' ПРямоугольник отрисовки метки
       m_Rect.Bottom = 40
       m_Rect.Right = 100
       m_Rect.Top = 0
       '--------------------------------------------------------------
       m_Font.Name = FONT_NAME:  m_Font.SIZE = FONT_SIZE:   m_Font.Bold = FONT_BOLD:   m_Font.Italic = FONT_ITALIC
       m_CHARSET = RUSSIAN_CHARSET       ' (default charset)
       m_Holst = 1                       ' По умолчанию - берем активное окно
       m_ForeColor = vbBlack             ' Цвет переднего плана по умолчанию
       m_BackColor = vbWhite             ' Цвет заднего плана по умолчанию
       m_BorderColor = vbRed             ' Граница по умолчанию
       m_BorderTransparent = False       ' Граница по умолчанию НЕПРОЗРАЧНА
       m_Transparency = True             ' По умолчанию - прозрачно
       m_BorderWidth = 1                 ' Толщина границы по умолчанию
       m_BackStyle = 0                   ' Стиль подложки по умолчанию
       m_BorderStyle = 0                 ' Сплошная линия
       m_PaintStyle = 0                  ' Сплошная
       bAllocateDC = False               ' DC не определено
       m_TextBgrd = False                ' Текст выводится прозрачно
       m_SingleLine = True               ' Текст выводится в одну строку
       m_TextAligne = 0                  ' Текст по горизонтале по середине
       m_TextVAligne = 0                 ' Текст по вертикале по середине
       m_TextTuncated = False            ' Не обрезать текст
       m_TextDrawMode = 0                ' Use DrawText by default (see DRAWSTRING)
       '---------------------------------------------------------------
       m_ALLCENTER = False               ' Изначально центрирование метки относительно контекста не определено
       m_AUTOSIZE = False                ' Изначально изменение ширины и высоты в зависимости от текста - не предусмотрено
End Sub
Private Sub Class_Terminate()
       ReleaseDC hWnd, m_hDC
End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------
' ОТРИСОВКА ПОВЕРХНОСТИ
'-------------------------------------------------------------------------------------------------------------------------------------------
Private Function DrawBack(aDC As Variant, aRECT As RECT, Optional BorderColor As Long = 0, _
                             Optional BackColor As Long = 16777215, Optional bBorderTransparent As Boolean = False, _
                             Optional bBackTransparent As Boolean = False, Optional iBorderWidth As Integer = 1, _
                             Optional iBorderStyle As Integer = 0, Optional iPaintStyle As Integer = 0, Optional iStyle As Integer = 0) As Boolean
    Const iOffSet As Integer = 6
    If bBorderTransparent And bBackTransparent Then Exit Function
#If Win64 Then
    Dim hRgn As LongPtr
    Dim hdc As LongPtr
    hdc = CLngPtr(aDC)
#Else
    Dim hRgn As Long
    Dim hdc As Long
    hdc = CLng(aDC)
#End If
'-----------------------------------------------------------------------------------------------------------------------
Select Case iStyle
Case 0:                     ' СТИЛЬ ПО УМОЛЧАНИЮ - ЗАПОЛНЕННЫЙ ПРЯМОУГОЛЬНИК С РАМКОЙ
     hRgn = CreateRectRgn(aRECT.Left, aRECT.Top, aRECT.Right, aRECT.Bottom)         ' Создали регион
     If Not bBackTransparent Then
        SetupBrush False, 1, BackColor
        FillRgn hdc, hRgn, m_hBrsh                                                     ' ЗАЛИЛИ ЦВЕТОМ
        SetupBrush True                                                                ' Удалили кисть
     End If
     If Not bBorderTransparent Then
        SetupBrush False, 1, BorderColor, PaintStyle
        FrameRgn hdc, hRgn, m_hBrsh, BorderWidth, BorderWidth                            ' Отрисовали бордюр
        SetupBrush True                                                                  ' Удалили кисть
     End If
     DeleteObject hRgn
'---------------------------------------------------------------------------
Case 1:                     ' ЗАКРУГЛЕННЫЙ ПРЯМОУГОЛЬНИК C ОПРЕДЕЛЕННЫМ СТИЛЕМ ЛИНИИ
     SetupBrush False, 2, BackColor, PaintStyle
     SetupPen False, BorderColor, BorderStyle
     RoundRect hdc, aRECT.Left, aRECT.Top, aRECT.Right, aRECT.Bottom, 5, 5          ' Рисуем
     SetupBrush True: SetupPen True                                                 ' Освобождаем ресурсы
Case 2:                     ' КРУГ ПОЗАДИ
     SetupBrush False, 2, BackColor, 0                                              ' Устанавливаем основной цвет
     SetupPen False, vbWhite, BorderStyle                                           ' Границу делаем белой
          Ellipse hdc, aRECT.Left, aRECT.Top, aRECT.Right, aRECT.Bottom             ' Рисуем большой элипс (цветной)
     SetupBrush False, 2, vbWhite, 0                                                ' Устанавливаем основной цвет
          Ellipse hdc, aRECT.Left + iOffSet, aRECT.Top + iOffSet, _
                            aRECT.Right - iOffSet, aRECT.Bottom - iOffSet           ' Рисуем меньший элипс (белый)
     SetupBrush False, 2, BackColor, 0                                              ' Переставляем в основной цвет
          Ellipse hdc, aRECT.Left + 2 * iOffSet, aRECT.Top + 2 * iOffSet, _
                             aRECT.Right - 2 * iOffSet, aRECT.Bottom - 2 * iOffSet  ' Рисуем внутренний элипс
     SetupBrush True: SetupPen True                                                 ' Освобождаем ресурсы
Case 3:                     ' BALOON
     Call DrawBalloon(aDC, aRECT)
Case Else:                  ' СТИЛЬ НЕ ОПРЕДЕЛЕН
          Exit Function
End Select
'------------------------------------------------
ExitHere:
    DrawBack = True '!!!!!!!!!!!!!!!!!
    Exit Function
'----------------------------
ErrHandle:
    MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "DRAW BACK ERROR"
    Err.Clear
End Function
'------------------------------------------------------------------------------------------------------------------------------
' Функция создает кисть и присваивает ее контексту, а также удаляет ее в дальнейшем
' iWay = 1 - SolidBrush, iWay = 2  - InDirect
'------------------------------------------------------------------------------------------------------------------------------
Private Sub SetupBrush(Optional bDestroy As Boolean = False, Optional iWay As Integer = 1, _
                                          Optional bColor As Long = 14811135, Optional bPaintStyle As Integer = BS_SOLID)
Dim p As LOGPEN, b As LOGBRUSH

If m_hBrsh <> 0 Then DeleteObject m_hBrsh
If bDestroy Then Exit Sub
'----------------------------------------------------------------
    Select Case iWay
    Case 1:                           ' Сплошная обычная кисть
         m_hBrsh = CreateSolidBrush(bColor)
    Case 2:                           ' Кисть с заданными параметрами
        b.lbColor = bColor
        Select Case bPaintStyle
        Case 0:
            b.lbStyle = BS_SOLID: b.lbHatch = HS_HORIZONTAL
        Case 1:
            b.lbStyle = BS_HATCHED: b.lbHatch = HS_HORIZONTAL
        Case 2:
            b.lbStyle = BS_HATCHED: b.lbHatch = HS_VERTICAL
        Case 3:
            b.lbStyle = BS_HATCHED: b.lbHatch = HS_FDIAGONAL
        Case 4:
            b.lbStyle = BS_HATCHED: b.lbHatch = HS_BDIAGONAL
        Case 5:
            b.lbStyle = BS_HATCHED: b.lbHatch = HS_CROSS
        Case 6:
            b.lbStyle = BS_HATCHED: b.lbHatch = HS_DIAGCROSS
        Case Else
        End Select
        m_hBrsh = CreateBrushIndirect(b)                                                 ' Создаем кисть
        SelectObject m_hDC, m_hBrsh
    Case Else
    End Select
End Sub
'------------------------------------------------------------------------------------------------------------------------------
' Функция создает перо и присваивает его контексту, а также удаляет его в дальнейшем
'------------------------------------------------------------------------------------------------------------------------------
Private Sub SetupPen(Optional bDestroy As Boolean = False, Optional pColor As Long = 0, Optional pStyle As Integer = PS_SOLID)
Dim p As LOGPEN, b As LOGBRUSH

If m_hPen <> 0 Then DeleteObject m_hPen
If bDestroy Then Exit Sub
'----------------------------------------------------------------
     p.lopnColor = pColor: p.lopnStyle = pStyle '(PS_SOLID = 0,PS_DASH = 1,PS_DOT = 2,PS_DASHDOT = 3,PS_DASHDOTDOT = 4)
     m_hPen = CreatePenIndirect(p)                                                    ' Создаем перо
     SelectObject m_hDC, m_hPen
End Sub

'------------------------------------------------------------------------------------------------------------------------------
' ОТРИСОВКА СТРОКИ
'------------------------------------------------------------------------------------------------------------------------------
Private Function DRAWSTRING(sText As String, aDC As Variant, aRECT As RECT, fnt As StdFont, TextColor As Long, BackColor As Long, _
                             Optional bTransparent As Boolean = True, Optional bTextBckgrd As Boolean = False, _
                                                                            Optional iFormat As Long = DT_SINGLELINE) As Boolean
Dim lRes As Long

If sText = "" Then Exit Function
'---------------------------------------------
#If Win64 Then
    Dim hdc As LongPtr
    hdc = CLngPtr(aDC)
#Else
    Dim hdc As LongPtr
    hdc = CLngPtr(aDC)
#End If
'------------------------------------------------------------------------------------------------------------
   SetFontToDC aDC, fnt, True                     ' Установили шрифт
   SetTextColor hdc, TextColor                    ' Установили  цвет выводимого текста
   If Not m_TextBgrd Then                         ' ПРОЗРАЧНОСТЬ И ЦВЕТ ПОДЛОЖКИ
      SetBkMode hdc, TRANSPARENT
   Else                                           ' ПОДЛОЖКА ТЕКСТА (определяется независимо от прямоугольника метки)
      SetBkMode hdc, OPAQUE
      SetBkColor hdc, m_BackColor
   End If
'------------------------------------------------------------------------------------------------------------
    Select Case m_TextDrawMode
    Case 0:    ' Standard DrawText
         lRes = DrawText(hdc, sText, Len(sText), aRECT, iFormat)
    Case 1:    ' Nothing to do right now
    
    Case 2:    ' DrawTextW
         lRes = DrawText2(hdc, StrPtr(sText), Len(sText), aRECT, iFormat)
    Case Else
    End Select

    '------------------------------------------------------------------------------------
ExitHere:
    SetFontToDC aDC, fnt, False                   ' Удалили шрифт
    DRAWSTRING = True   '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'------------------------
ErrHandle:
    ErrPrint "DRAWSTRING", Err.Number, Err.Description
    Err.Clear
End Function
'--------------------------------------------------------------------------------------------------------------------------
' Функция получает графический контекст в зависимости от установок холста
'--------------------------------------------------------------------------------------------------------------------------
Private Sub AllocateDC()
       If m_Holst = 0 Then
          Call SetDCtoScreen
       ElseIf m_Holst = 1 Then
          Call SetDCtoActiveWindow
       Else
          If m_hDC = 0 Then Exit Sub
       End If
'----------------------------
       bAllocateDC = True
'----------------------------
End Sub
'--------------------------------------------------------------------------------------------------------------------------
' Установка графического контекста на активное окно
'--------------------------------------------------------------------------------------------------------------------------
Private Sub SetDCtoActiveWindow()
        hWnd = GetActiveWindow
        m_hDC = GetWindowDC(hWnd)
End Sub
'--------------------------------------------------------------------------------------------------------------------------
' Установка графического контекста на экран
'--------------------------------------------------------------------------------------------------------------------------
Private Sub SetDCtoScreen()
        hWnd = HWND_DESKTOP
        m_hDC = GetDC(HWND_DESKTOP)
End Sub
'------------------------------------------------------------------------------------------------------------------------------
' Функция назначает новый шрифт на контекст, восстанавливает старый
'------------------------------------------------------------------------------------------------------------------------------
Private Sub SetFontToDC(aDC As Variant, Optional fnt As StdFont, Optional bSetNew As Boolean = True)
Dim LF As LOGFONT
     On Error GoTo ErrHandle
'If Not m_RecreateFont Then Exit Sub              ' ЕСЛИ НЕТ НЕОБХОДИМОСТИ ПЕРЕНАЗНАЧАТЬ ШРИФТ - ВЫХОДИМ
'------------------------------------------------------------------------------------------------------
#If Win64 Then
     Dim hdc As LongPtr
     hdc = CLngPtr(aDC)
#Else
     Dim hdc As Long
     hdc = CLng(aDC)
#End If
'-----------------------------------------
    Select Case bSetNew
    Case True:                                   ' НОВЫЙ ШРИФТ
        LF = ConvertFontToLF(CVar(hdc), fnt)
        m_NewFont = CreateFontIndirect(LF)
        m_RecreateFont = False                                               ' Сбрасываем необходимость пересоздавать шрифт
        If m_NewFont <> 0 Then m_OldFont = SelectObject(hdc, m_NewFont)      ' Установили новый шрифт, сохранили старый
    Case False:                                  ' ВОССТАНОВЛЕНИЕ СТАРОГО ШРИФТА
        If m_OldFont <> 0 Then SelectObject hdc, m_OldFont                   ' Восстановили старый шрифт
        If m_NewFont <> 0 Then DeleteObject m_NewFont                        ' Уничтожаем предыдущий шрифт
        m_RecreateFont = True
    End Select
'-------------------------------------------
ExitHere:
    Exit Sub
'--------------------------
ErrHandle:
    MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "FONT ERROR"
    Err.Clear
End Sub
'------------------------------------------------------------------------------------------------------------------------------
' Функция конвертирует стандартный шрифт в структуру LOGFONT
'------------------------------------------------------------------------------------------------------------------------------
Private Function ConvertFontToLF(aDC As Variant, fnt As StdFont) As LOGFONT
Dim LF As LOGFONT                              ' Структура, описывающая шрифт'
Dim sFontName As String                        ' Подготовленная строка с наименованием шрифта
Dim I As Long                                  ' Циклическая переменная

On Error GoTo ErrHandle

#If Win64 Then
     Dim hdc As LongPtr
     hdc = CLngPtr(aDC)
#Else
     Dim hdc As Long
     hdc = CLng(aDC)
#End If
'-----------------------------------------
        sFontName = fnt.Name & vbNullChar
'--------------------------------------------------------------------------------
' Устанавливаем параметры шрифта
        LF.lfCharSet = m_CHARSET 'RUSSIAN_CHARSET
        If m_Angle <> 0 Then LF.lfEscapement = 1800
        If fnt.Italic Then LF.lfItalic = 1
        LF.lfHeight = -MulDiv(fnt.SIZE, GetDeviceCaps(hdc, LOGPIXELSY), 72)
        If fnt.Bold Then
              LF.lfWeight = FW_BOLD
        Else
              LF.lfWeight = FW_NORMAL
        End If
        For I = 1 To Len(fnt.Name)
            LF.lfFaceName(I) = Asc(Mid(fnt.Name, I, 1))
        Next
'-------------------------------------------------------------
ExitHere:
        ConvertFontToLF = LF  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        Exit Function
'-------------------------
ErrHandle:
    ErrPrint "ConvertFontToLF", Err.Number, Err.Description
    Err.Clear
End Function
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Обрабатываем буфер в зависимости от команды, устанавливаемой меткой m_RecreateBuffer
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function ProcessBUfer(aDC As Variant, aRCT As RECT) As Boolean
Dim bRes As Boolean
     Select Case m_RecreateBuffer
     Case 0:                        ' Первый раз отрисовываем метку, сохраняем буфер под ней
         bRes = CreateBuffer(aDC, aRCT)                   ' Создаем буфер из текущей положки
         m_RecreateBuffer = 1                             ' Устанавливаем регулярный режим
     Case 1:                        ' Обычный режим, имеем сохраненный буфер
         If m_Transparency Then                           ' Проверяем необходимость восстановления
            bRes = RestoreBGFromBuffer(aDC, aRCT)         ' Восстанавливаем буфер
         End If
     Case 2:                        ' Восстанавливаемся и создаем новый буфер
            bRes = RestoreBGFromBuffer(aDC, aRCT)         ' Восстанавливаем буфер
            If bRes Then bRes = CreateBuffer(aDC, aRCT)   ' Создаем буфер из новой положки
     End Select
'------------------------------------------
     ProcessBUfer = bRes '!!!!!!!!!!!!!!!!!!!
End Function
'---------------------------------------------------------------------------------------------------------------------------
' Функция создает DC в памяти и копирует туда область под меткой
'---------------------------------------------------------------------------------------------------------------------------
Private Function CreateBuffer(aDC As Variant, aRCT As RECT) As Boolean
Dim lWidthPixels As Long, lHeightPixels As Long                ' Длина и ширина копируемой области
    On Error GoTo ErrHandle
#If Win64 Then
    Dim hdc As LongPtr
    hdc = CLngPtr(aDC)
#Else
    Dim hdc As Long
    hdc = CLng(aDC)
#End If

If hdc = 0 Then Exit Function
'----------------------------------------------------------------------------------------------------------
' Рассчитываем ширину и высоту копируемой области
lWidthPixels = aRCT.Right - aRCT.Left: lHeightPixels = aRCT.Bottom - aRCT.Top
If lWidthPixels = 0 Or lHeightPixels = 0 Then Exit Function
'-------------------------------------------------------
' Создаем графический контекст и его BitMap в памяти
If m_bufferDC <> 0 Or m_BMP <> 0 Then Call RemoveBuffer
'-----------------------------------------------------------------------------------------------------------
m_bufferDC = CreateCompatibleDC(hdc)                                         ' Создали в памяти совместимый DC
    If (m_bufferDC <> 0) Then                                                ' Если создали удачно
      m_BMP = CreateCompatibleBitmap(hdc, lWidthPixels, lHeightPixels)       ' Создаем BMP
        '--------------------------------------------------
        If (m_BMP <> 0) Then
           m_BMPOLD = SelectObject(m_bufferDC, m_BMP)                        ' Положили на DC в памяти рисунок
           ' КОПИРУЕМ В ПАМЯТЬ РЕГИОН ИЗ ИСХОДНОЙ ФОРМЫ
           BitBlt m_bufferDC, 0, 0, lWidthPixels, lHeightPixels, hdc, aRCT.Left, aRCT.Top, SRCCOPY
        End If
    End If
'------------------------------------------------------------------------------------------------
ExitHere:
    CreateBuffer = True  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'----------------------------------------------------------
ErrHandle:
    MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Function
'------------------------------------------------------------------------------------------------------------------------------
' Функция восстанавливает подложку из буфера
'------------------------------------------------------------------------------------------------------------------------------
Private Function RestoreBGFromBuffer(aDC As Variant, aRCT As RECT) As Boolean
Dim lWidthPixels As Long, lHeightPixels As Long            ' Длина и ширина копируемой области
    On Error GoTo ErrHandle
#If Win64 Then
    Dim hdc As LongPtr
    hdc = CLngPtr(aDC)
#Else
    Dim hdc As Long
    hdc = CLng(aDC)
#End If

If hdc = 0 Then Exit Function                              ' не задано устройство
If m_bufferDC = 0 Then Exit Function                       ' Не существует соответствующего контекста
'----------------------------------------------------------------------------------------------------------
' Рассчитываем ширину и высоту копируемой области
lWidthPixels = aRCT.Right - aRCT.Left
lHeightPixels = aRCT.Bottom - aRCT.Top
If lWidthPixels = 0 Or lHeightPixels = 0 Then Exit Function
'-----------------------------------------
           ' КОПИРУЕМ РЕГИОН ИЗ ПАМЯТИ В ИСХОДНУЮ ФОРМУ
           BitBlt hdc, aRCT.Left, aRCT.Top, lWidthPixels, lHeightPixels, m_bufferDC, 0, 0, SRCCOPY
'------------------------------------------------------------------------------------------------
ExitHere:
    RestoreBGFromBuffer = True  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'----------------------------------------------------------
ErrHandle:
    MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description
    Err.Clear
End Function
'------------------------------------------------------------------------------------------------------------------------------
' Функция удаляет буфер
'------------------------------------------------------------------------------------------------------------------------------
Private Sub RemoveBuffer()
           SelectObject m_bufferDC, m_BMPOLD
           DeleteObject m_BMP
           DeleteObject m_bufferDC
End Sub
'------------------------------------------------------------------------------------------------------------------------------
' Функция позволяет рисовать прямоугольный контур для метки (НЕИСПОЛЬЗУЕТСЯ)
'------------------------------------------------------------------------------------------------------------------------------
Private Sub CreateBorder(cDC As Variant, rc As RECT, Optional BorderColor As Long = 6579400)
   
   
#If Win64 Then
    Dim hRgn As LongPtr
    Dim hBrsh As LongPtr
    Dim hdc As LongPtr
    hdc = CLngPtr(cDC)
#Else
    Dim hRgn As Long
    Dim hBrsh As LongPtr
    Dim hdc As Long
    hdc = CLng(cDC)
#End If
'---------------------------------------------------------------------------
hRgn = CreateRectRgn(rc.Left, rc.Top, rc.Right, rc.Bottom)        ' Создали регион
hBrsh = CreateSolidBrush(BorderColor)                             ' Создали кисть
FrameRgn hdc, hRgn, hBrsh, 1, 1                                   ' Отрисовали бордюр


DeleteObject hBrsh                                                ' Удалили кисть
End Sub
'-----------------------------------------------------------------------------------------------------------------------
' Функция получает размеры строки для установленного шрифта,контекста
' Работает ТОШЛЬКО ПОСЛЕ УСТАНОВКИ ШРИФТА НА КОНТЕКСТ
'-----------------------------------------------------------------------------------------------------------------------
Private Function GetTextSize(aDC As Variant, str As String) As SIZE
Dim TextSize As SIZE
'----------------------------------------------------
#If Win64 Then
    Dim hdc As LongPtr
    hdc = CLngPtr(aDC)
#Else
    Dim hdc As Long
    hdc = CLng(aDC)
#End If
    GetTextExtentPoint32 hdc, str, Len(str), TextSize
'----------------------------------------------------
ExitHere:
    GetTextSize = TextSize '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------
' ФУНКЦИЯ РАССЧИТЫВАЕТ ЭФФЕКТИВНЫЙ ПРЯМОУГОЛЬНИК (тот который будет рисован)
' По отношению к заданному - он может быть смещен из-за центрирования, или обрезан из-за нарушения границ или
' потсроен свободно на основании анализа текста (его высоты и ширины)
'------------------------------------------------------------------------------------------------------------------------------
Private Function CalculateEffRect(sText As String, aDC As Variant, _
                                       LogRect As RECT, fnt As StdFont, AutoSize As Boolean, ALLCENTER As Boolean) As RECT
Dim RectEFF As RECT                                                     ' Возвращаемый результат
Dim mySZ As SIZE
Dim ContextECT As RECT, lWidth As Long, lHeight As Long                 ' Прямоугольник контекста (его длина и ширина в пикселах)
Const iOffSet As Long = 2                                               ' Количество пикселей, отнимаемых от границы

     RectEFF = LogRect                                                  ' Копируем исходный прямоугольник
'----------------------------------------------------------------------------------
     If m_AUTOSIZE Then                                                 ' Изменяем длину и ширину в зависимости от строки
         SetFontToDC aDC, fnt, True                                     '    Создаем шрифт временно
         mySZ = GetTextSize(aDC, sText)                                 ' Получаем длинуи ширину в пределах контекста и шрифта
         RectEFF.Right = RectEFF.Left + mySZ.CX + iOffSet               ' Рассчитываем авторазмеры
         RectEFF.Bottom = RectEFF.Top + mySZ.CY + iOffSet
         SetFontToDC aDC, fnt, False                                    '    Уничтожаем шрифт - предварительный расчет
     End If
'----------------------------------------------------------------------------------
     ContextECT = GetContextRectangle(aDC, iOffSet)                     ' ОБЩИЙ КОНТУР КОНТЕКСТА (его длина и ширина)
     '----------------------------------------------------
     If ALLCENTER Then                                                  ' Если метка центрируется - пересчитываем прямоугольник
          lWidth = RectEFF.Right - RectEFF.Left: lHeight = RectEFF.Bottom - RectEFF.Top
          RectEFF.Left = ContextECT.Left + (ContextECT.Right - ContextECT.Left - lWidth) / 2
          RectEFF.Right = RectEFF.Left + lWidth
          RectEFF.Top = ContextECT.Top + (ContextECT.Bottom - ContextECT.Top - lHeight) / 2
          RectEFF.Bottom = RectEFF.Top + lHeight
     End If
     ' ПРОИЗВОДИМ ОБРЕЗКУ В СЛУЧАЕ ВЫХОДА ЗА ГРАНИЦУ
     If RectEFF.Right > ContextECT.Right Then RectEFF.Right = ContextECT.Right
     If RectEFF.Bottom > ContextECT.Bottom Then RectEFF.Bottom = ContextECT.Bottom
'-----------------------------------------------------------
ExitHere:
     CalculateEffRect = RectEFF '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'---------------------------------------------------------------------------------------------------------------------------------
' Функция вычисляет стиль текста
'---------------------------------------------------------------------------------------------------------------------------------
Private Function GetTextStyle() As Long
Dim iRes As Long

If m_TextAligne = -1 Then
    iRes = DT_LEFT
ElseIf m_TextAligne = 0 Then
    iRes = DT_CENTER
Else
    iRes = DT_RIGHT
End If
'------------------------------------------------------------------------------
If m_SingleLine Then iRes = iRes Or DT_SINGLELINE
'------------------------------------------------------------------------------
If m_TextVAligne = -1 Then
    iRes = iRes Or DT_BOTTOM
ElseIf m_TextVAligne = 0 Then
    iRes = iRes Or DT_VCENTER
Else
    iRes = iRes Or DT_TOP
End If
'------------------------------------------------------------------------------
If m_TextTuncated Then iRes = iRes & DT_END_ELLIPSIS
'iRes = DT_CENTER Or DT_SINGLELINE Or DT_VCENTER
'----------------------------------------
ExitHere:
     GetTextStyle = iRes
End Function


'---------------------------------------------------------------------------------------------------------------------------------
' Функция отрисовывает балун
'---------------------------------------------------------------------------------------------------------------------------------
Private Sub DrawBalloon(aHdc As Variant, rc As RECT)
    Dim pts(0 To 2) As POINTAPI
Const COLOR_INFOTEXT = 23
Const COLOR_INFOBK = 24
Const COLOR_GRAYTEXT = 17
Const COLOR_3DLIGHT = 22
Const iOffSet = 8
'---------------------------------------------------------
 #If Win64 Then
     Dim hRgn As LongPtr, hrgn1 As LongPtr, hrgn2 As LongPtr
     Dim hdc As LongPtr
     hdc = CLngPtr(aHdc)
 #Else
     Dim hRgn As Long, hrgn1 As Long, hrgn2 As Long
     Dim hdc As Long
     hdc = CLng(aHdc)
 #End If
 '---------------------------------------------------------
    pts(0).x = rc.Left + iOffSet
    pts(0).y = rc.Top
    pts(1).x = pts(0).x
    pts(1).y = pts(0).y + iOffSet
    pts(2).x = pts(1).x + iOffSet
    pts(2).y = pts(1).y
    hRgn = CreateRectRgn(0, 0, 0, 0)
    hrgn1 = CreateRoundRectRgn(rc.Left, rc.Top + iOffSet, rc.Right, rc.Bottom, 15, 15)
    hrgn2 = CreatePolygonRgn(pts(0), 3, ALTERNATE)
    CombineRgn hRgn, hrgn1, hrgn2, RGN_OR
        FillRgn hdc, hRgn, GetSysColorBrush(COLOR_INFOBK)
        FrameRgn hdc, hRgn, GetStockObject(DKGRAY_BRUSH), 1, 1
    'RC.Top = RC.Top + iOffset * 2
    'RC.Bottom = RC.Bottom - iOffset
    'RC.Left = RC.Left + iOffset
    'RC.Right = RC.Right - iOffset
    ' Draw the Shadow Text
    'SetTextColor hdc, GetSysColor(COLOR_3DLIGHT)
    'DrawText hdc, lpszText, Len(lpszText), rc, DT_VCENTER + DT_NOCLIP
    'rc.Left = rc.Left - 1
    'rc.Top = rc.Top - 1
    ' Draw the Text
    'SetTextColor hdc, GetSysColor(COLOR_INFOTEXT)
    'DrawText hdc, lpszText, Len(lpszText), rc, DT_VCENTER + DT_NOCLIP
End Sub
'-----------------------------------------------------------------------------------------------------------------------
' Функция получает размеры контекста
'-----------------------------------------------------------------------------------------------------------------------
Private Function GetContextRectangle(cDC As Variant, Optional iOffSet As Long = 0, Optional ByHWND As Boolean = True) As RECT
Dim myRect As RECT
#If Win64 Then
    Dim hdc As LongPtr
    hdc = CLngPtr(cDC)
#Else
    Dim hdc As Long
    hdc = CLng(cDC)
#End If
    If ByHWND Then   ' Больше походит для VBA
        GetClientRect hWnd, myRect
    Else              ' Больше походит для экрана и внешнего окна
        myRect.Left = iOffSet:   myRect.Top = iOffSet
        myRect.Right = GetDeviceCaps(hdc, HORZRES) - 2 * iOffSet
        myRect.Bottom = GetDeviceCaps(hdc, VERTRES) - 2 * iOffSet
    End If
'--------------------------------------------------
    GetContextRectangle = myRect '!!!!!!!!!!!!!!!
End Function

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Error Handler
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function ErrPrint(FuncName As String, ErrNumber As Long, ErrDescription As String, Optional bDebug As Boolean = True, _
                                                                                                  Optional sModName As String = "cLabel") As String
Dim sRes As String
Const ERR_CHAR As String = "#"
Const ERR_REPEAT As Integer = 60

sRes = String(ERR_REPEAT, ERR_CHAR) & vbCrLf & "ERROR OF [" & sModName & ": " & FuncName & "]" & vbTab & "ERR#" & ErrNumber & vbTab & Now() & _
       vbCrLf & ErrDescription & vbCrLf & String(ERR_REPEAT, ERR_CHAR)
If bDebug Then Debug.Print sRes
'----------------------------------------------------------
ExitHere:
       Beep
       ErrPrint = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function

