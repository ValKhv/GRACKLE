VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************************************************************************
'*********************************************************************************************************************************************************
' DESCRIPTION: The class class implements Shape Object
' VERSION    : v 0.2 20180703
' COPYRIGHT  : V.Khvatov DigitalXpert
'*********************************************************************************************************************************************************
'*********************************************************************************************************************************************************

Option Explicit

'*********************************************************************************************************************************************************
Private Const ANSI_CHARSET = 0                                              ' ANSI character set
Private Const RUSSIAN_CHARSET = 204                                         ' Для русского

Private Const FW_BOLD = 700
Private Const FW_NORMAL = 400


Private Const LF_FACESIZE = 32
Private Const OPAQUE = 2
Private Const TRANSPARENT = 1

Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90
Private Const TwipsPerInch = 1440

Public Enum ShapeType
    GC_POINT = 0
    GC_LINE = 1
    GC_CIRCLE = 2
    GC_ELLIPS = 3
    GC_TRIANGLE = 4
    GC_RECTANGLE = 5
    GC_POLIGON = 6
    GC_CURVE = 7
End Enum

Public Enum LineType               ' COMPATIBLE WITH WIN API
    GC_LSOLID = 0                       '  ------
    GC_DASH = 1                         '  - - - -
    GC_DOT = 2                          '  .......
    GC_DASHDOT = 3                      '  _._._._
    GC_DASHDOTDOT = 4                   '  _.._.._
End Enum

Public Enum AlignType
    GC_BYCOORDINATES = -1
    GC_TOPLEFT = 0
    GC_TOPCENTER = 1
    GC_TOPRIGHT = 2
    GC_CENTERLEFT = 3
    GC_CENTERCENTER = 4
    GC_CENTERRIGHT = 5
    GC_BOTTOMLEFT = 6
    GC_BOTTOMCENTER = 7
    GC_BOTTOMRIGHT = 8
End Enum

'*********************************************************************************************************************************************************
Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type LOGFONT
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
       lfFaceName(LF_FACESIZE) As Byte        'lfFaceName As String * LF_FACESIZE
End Type
Private Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type
Private Type LOGPEN
        lopnStyle As Long
        lopnWidth As POINTAPI
        lopnColor As Long
End Type

Private Type SIZE
         CX As Long
         CY As Long
End Type

#If VBA7 Then
        Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
        Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
        
        Private Declare PtrSafe Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
        Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
        
        
        Private Declare PtrSafe Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As LongPtr
        Private Declare PtrSafe Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As LongPtr
        Private Declare PtrSafe Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As LongPtr
        
        Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
        Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Boolean
        
        Private Declare PtrSafe Function GetRgnBox Lib "gdi32" (ByVal hRgn As LongPtr, lpRect As RECT) As Long
        Private Declare PtrSafe Function FillRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr, ByVal hBrush As LongPtr) As Long
        Private Declare PtrSafe Function FrameRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr, ByVal hBrush As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As Long
        
        Private Declare PtrSafe Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As LongPtr
        Private Declare PtrSafe Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As LongPtr
        Private Declare PtrSafe Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As LongPtr
        Private Declare PtrSafe Function MoveToEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
        Private Declare PtrSafe Function LineTo Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
        Private Declare PtrSafe Function PolyBezier Lib "gdi32" (ByVal hdc As LongPtr, lppt As POINTAPI, ByVal cPoints As Long) As Long
        Private Declare PtrSafe Function SetPixel Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
        
        Private Declare PtrSafe Function SetTextColor Lib "gdi32" (ByVal hdc As LongPtr, ByVal crColor As Long) As Long
        Private Declare PtrSafe Function GetBkColor Lib "gdi32" (ByVal hdc As LongPtr) As Long
        Private Declare PtrSafe Function SetBkColor Lib "gdi32" (ByVal hdc As LongPtr, ByVal crColor As Long) As Long
        Private Declare PtrSafe Function SetBkMode Lib "gdi32" (ByVal hdc As LongPtr, ByVal nBkMode As Long) As Long
        Private Declare PtrSafe Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As LongPtr, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

#Else
        Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
        Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
        
        Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
        Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
        
        
        Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
        Private Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long
        Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
        
        Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
        Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Boolean
        
        Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
        Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
        Private Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
        
        Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
        Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
        Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
        Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
        Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
        Private Declare Function PolyBezier Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, ByVal cPoints As Long) As Long
        Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
        
        Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
        Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
        Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
        Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
        Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

#End If

'*********************************************************************************************************************************************************
Private m_Left As Long
Private m_Top As Long
Private m_Width As Long
Private m_Height As Long

Private m_Text As String
Private m_TextAutoSize As Boolean
Private m_TextAlign As AlignType

Private m_BackColor As Long
Private m_BackTransparent As Boolean

Private m_BorderColor As Long
Private m_BorderWidth As Integer
Private m_BorderStyle As LineType
Private m_BorderTransparent As Boolean

Private m_FontName As String
Private m_FontSize As Integer
Private m_FontColor As Long
Private m_FontBold As Boolean
Private m_FontItalic As Boolean
Private m_FontUnderline As Boolean
Private m_FontCharSet As Integer

Private m_ShapeType As ShapeType
Private m_POINTS() As POINTAPI
Private m_Count As Integer
Private m_HasDrawed As Boolean

'*********************************************************************************************************************************************************
' DEFAULT CONSTANTS
Private Const DEFAULT_LEFT As Long = 100
Private Const DEFAULT_TOP As Long = 100
Private Const DEFAULT_WIDTH As Long = 120
Private Const DEFAULT_HEIGHT As Long = 80

Private Const DEFAULT_FONTNAME As String = "Tahoma"                                '& vbNullChar
Private Const DEFAULT_FONTSIZE As Long = 12
Private Const DEFAULT_FONTBOLD As Boolean = False
Private Const DEFAULT_FONTITALIC As Boolean = False
Private Const DEFAULT_FONTUNDERLINE As Boolean = False

Private Const DEFAULT_CHARSET As Integer = RUSSIAN_CHARSET

Private Const DEFAULT_BORDERWIDTH As Integer = 1
Private Const DEFAULT_BORDERCOLOR As Long = vbRed
Private Const DEFAULT_BORDERSTYLE As Integer = GC_LSOLID
Private Const DEFAULT_BORDERTRANSPARENT As Boolean = False

Private Const DEFAULT_BACKCOLOR As Long = vbBlue
Private Const DEFAULT_BACKSTYLE As Integer = 0
Private Const DEFAULT_BACKTRASPARENT As Boolean = False

Private Const DEFAULT_CAPTION As String = "SHAPE"
Private Const DEFAULT_TEXT_ALIGN As Integer = GC_CENTERCENTER ' CENTERCENTER
Private Const DEFAULT_TEXT_SINGLELINE As Boolean = True
Private Const DEFAULT_TEXT_TRUNCATED As Boolean = True
Private Const DEFAULT_TEXT_COLOR As Long = vbWhite
Private Const DEFAULT_TEXT_AUTOSIZE As Boolean = False

Private Const DEFAULT_SHAPE As Integer = GC_TRIANGLE           ' TRIANGLE

'*********************************************************************************************************************************************************
Public Property Get ShapeType() As ShapeType
    ShapeType = m_ShapeType
End Property
Public Property Let ShapeType(iShape As ShapeType)
    m_ShapeType = iShape
    SetupShape (iShape)
End Property

Public Property Get HasDrawed() As Boolean
       HasDrawed = m_HasDrawed
End Property

Public Property Get Left() As Long
       Left = m_Left
End Property
Public Property Let Left(iLeft As Long)
       m_Left = iLeft
End Property

Public Property Get Top() As Long
       Top = m_Top
End Property
Public Property Let Top(iTop As Long)
       m_Top = iTop
End Property

Public Property Get Width() As Long
       Width = m_Width
End Property
Public Property Let Width(IWidth As Long)
       m_Width = IWidth
End Property

Public Property Get height() As Long
       height = m_Height
End Property
Public Property Let height(iHeight As Long)
       m_Height = iHeight
End Property

Public Property Get Right() As Long
       Right = m_Left + m_Width
End Property
Public Property Let Right(iRight As Long)
       m_Width = iRight - m_Left
End Property

Public Property Get Bottom() As Long
       Bottom = m_Top + m_Height
End Property
Public Property Let Bottom(iBottom As Long)
       m_Height = iBottom - m_Top
End Property

Public Property Get Caption() As String
       Caption = m_Text
End Property
Public Property Let Caption(sCaption As String)
       m_Text = sCaption
End Property

Public Property Get BackColor() As Long
       BackColor = m_BackColor
End Property
Public Property Let BackColor(iBackColor As Long)
       m_BackColor = iBackColor
End Property

Public Property Get BorderColor() As Long
       BorderColor = m_BorderColor
End Property
Public Property Let BorderColor(iBorderColor As Long)
       m_BorderColor = iBorderColor
End Property

Public Property Get BorderWidth() As Integer
       BorderWidth = m_BorderWidth
End Property
Public Property Let BorderWidth(iBorderWidth As Integer)
       m_BorderWidth = iBorderWidth
End Property

Public Property Get BorderType() As LineType
       BorderType = m_BorderStyle
End Property
Public Property Let BorderType(iBorderType As LineType)
       m_BorderStyle = iBorderType
End Property

Public Property Get BorderTransparent() As Boolean
       BorderTransparent = m_BorderTransparent
End Property
Public Property Let BorderTransparent(iBorderTransparent As Boolean)
       m_BorderTransparent = iBorderTransparent
End Property

Public Property Get TextAlign() As AlignType
       TextAlign = m_TextAlign
End Property
Public Property Let TextAlign(iTextAlign As AlignType)
       m_TextAlign = iTextAlign
End Property

Public Property Get BackTransparent() As Boolean
       BackTransparent = m_BackTransparent
End Property
Public Property Let BackTransparent(iBackTransparent As Boolean)
       m_BackTransparent = iBackTransparent
End Property

Public Property Get FontName() As String
       FontName = m_FontName
End Property
Public Property Let FontName(sFontName As String)
       m_FontName = sFontName
End Property

Public Property Get FontSize() As Integer
       FontSize = m_FontSize
End Property
Public Property Let FontSize(iFontSize As Integer)
       m_FontSize = iFontSize
End Property

Public Property Get FontColor() As Long
       FontColor = m_FontColor
End Property
Public Property Let FontColor(iFontColor As Long)
       m_FontColor = iFontColor
End Property

Public Property Get TextAutoSize() As Boolean
       TextAutoSize = m_TextAutoSize
End Property
Public Property Let TextAutoSize(iTextAutoSize As Boolean)
       m_TextAutoSize = iTextAutoSize
End Property

Public Property Get FontBold() As Boolean
       FontBold = m_FontBold
End Property
Public Property Let FontBold(iFontBold As Boolean)
       m_FontBold = iFontBold
End Property

Public Property Get FontItalic() As Boolean
       FontItalic = m_FontItalic
End Property
Public Property Let FontItalic(iFontItalic As Boolean)
       m_FontItalic = iFontItalic
End Property

Public Property Get FontUnderline() As Boolean
       FontUnderline = m_FontUnderline
End Property
Public Property Let FontUnderline(iFontUnderline As Boolean)
       m_FontUnderline = iFontUnderline
End Property

Public Property Get FontCharset() As Integer
       FontCharset = m_FontCharSet
End Property
Public Property Let FontCharset(iFontCharset As Integer)
       m_FontCharSet = iFontCharset
End Property

Public Property Get PointsCount() As Integer
    PointsCount = m_Count
End Property

Public Property Get POINT_X(indx As Integer) As Long

On Error GoTo ErrHandle
'------------------------------
   If indx > m_Count Then Err.Raise 1000, , "Wrong Index = " & indx & " Can't be more than Count = " & m_Count
'-------------------------------
ExitHere:
   POINT_X = m_POINTS(indx).x '!!!!!!!!!!!!!
   Exit Property
'-------------
ErrHandle:
   ErrPrint "POINT_X", Err.Number, Err.Description
   Err.Clear
End Property
Public Property Let POINT_X(indx As Integer, x As Long)
    If indx > m_Count Then
        m_Count = indx: ReDim Preserve m_POINTS(m_Count)
    End If
    m_POINTS(m_Count).x = x
End Property

Public Property Get POINT_Y(indx As Integer) As Long

On Error GoTo ErrHandle
'------------------------------
   If indx > m_Count Then Err.Raise 1000, , "Wrong Index = " & indx & " Can't be more than Count = " & m_Count
'-------------------------------
ExitHere:
   POINT_Y = m_POINTS(indx).y  '!!!!!!!!!!!!!
   Exit Property
'-------------
ErrHandle:
   ErrPrint "POINT_Y", Err.Number, Err.Description
   Err.Clear
End Property
Public Property Let POINT_Y(indx As Integer, y As Long)
    If indx > m_Count Then
        m_Count = indx: ReDim Preserve m_POINTS(m_Count)
    End If
    m_POINTS(m_Count).y = y
End Property

Public Sub AddPoint(x As Long, y As Long)
    m_Count = m_Count + 1: ReDim Preserve m_POINTS(m_Count)
    m_POINTS(m_Count).x = x: m_POINTS(m_Count).y = y
End Sub

Public Sub SetCircle(x As Long, y As Long, Radious As Long)
   m_ShapeType = GC_CIRCLE:  m_Count = 1: ReDim m_POINTS(m_Count)
   m_POINTS(0).x = x: m_POINTS(0).y = y
   m_POINTS(1).x = Radious
End Sub

Public Sub SetEllipse(iLeft As Long, iTop As Long, IWidth As Long, iHeight As Long)
   m_ShapeType = GC_ELLIPS: m_Count = 1: ReDim m_POINTS(m_Count)
   m_POINTS(0).x = iLeft: m_POINTS(0).y = iTop
   m_POINTS(1).x = iLeft = IWidth: m_POINTS(1).y = iTop + iHeight
End Sub

Public Sub SetRectangle(iLeft As Long, iTop As Long, IWidth As Long, iHeight As Long)
   m_ShapeType = GC_RECTANGLE: m_Count = 1: ReDim m_POINTS(m_Count)
   m_POINTS(0).x = iLeft: m_POINTS(0).y = iTop
   m_POINTS(1).x = iLeft + IWidth: m_POINTS(1).y = iTop + iHeight
End Sub

Public Sub SetTriangle(X1 As Long, Y1 As Long, x2 As Long, Y2 As Long, x3 As Long, Y3 As Long)
   m_ShapeType = GC_TRIANGLE: m_Count = 2: ReDim m_POINTS(m_Count)
   
   m_POINTS(0).x = X1: m_POINTS(0).y = Y1
   m_POINTS(1).x = x2: m_POINTS(1).y = Y2
   m_POINTS(2).x = x3: m_POINTS(2).y = Y3
End Sub

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub SetLine(X1 As Long, Y1 As Long, x2 As Long, Y2 As Long, Optional IWidth As Long = -1, Optional iColor As Long = -1)
   m_ShapeType = GC_LINE:  m_Count = 1: ReDim m_POINTS(m_Count)
   m_POINTS(0).x = X1: m_POINTS(0).y = Y1
   m_POINTS(1).x = x2: m_POINTS(1).y = Y2
   If IWidth > 0 Then m_BorderWidth = IWidth
   If iColor <> -1 Then m_BorderColor = iColor
End Sub

Public Sub SetFont(zFont As StdFont)
    If zFont Is Nothing Then Exit Sub
    
    m_FontName = zFont.Name
    m_FontSize = zFont.SIZE
    m_FontBold = zFont.Bold
    m_FontItalic = zFont.Italic
    m_FontUnderline = zFont.Underline
    m_FontCharSet = zFont.CHARSET
    
End Sub
Public Function GetFont() As StdFont
Dim zFont As New StdFont
    zFont.Name = m_FontName
    zFont.SIZE = m_FontSize
    zFont.Bold = m_FontBold
    zFont.Italic = m_FontItalic
    zFont.Underline = m_FontUnderline
    zFont.CHARSET = m_FontCharSet
'--------------------------------------
    Set GetFont = zFont '!!!!!!!!!!!
End Function
Public Sub FromRect(iLeft As Long, iTop As Long, iRight As Long, iBottom As Long)
    m_Left = iLeft: m_Top = iTop
    m_Width = iRight - iLeft: m_Height = iBottom - iRight
End Sub
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub Draw(hWnd As Variant)
Dim ShapeRect As RECT

#If VBA7 Then
       Dim hdc As LongPtr
#Else
       Dim hdc As Long
#End If

On Error GoTo ErrHandle
'-------------------------
    If IsEmpty(hWnd) Then Exit Sub
    
    hdc = GetDC(hWnd)
    If hdc = 0 Then Err.Raise 1000, , "Wrong Graphical Context"
    
    '----------------------------
    Select Case m_ShapeType
        Case GC_POINT, GC_LINE, GC_CURVE:
             ShapeRect = DRAWFIGURE(hdc)
             
        Case GC_CIRCLE, GC_ELLIPS, GC_TRIANGLE, GC_RECTANGLE, GC_POLIGON:
             ShapeRect = DRAWREGION(hdc)
             
        Case Else
             Err.Raise 1000, , "Wrong Shape type"
    End Select
    
    If m_Text <> "" Then Call DRAWSTRING(hdc, ShapeRect)
'-------------------------
ExitHere:
    If hdc <> 0 Then ReleaseDC hWnd, hdc
    m_HasDrawed = True
    Exit Sub
'-----------
ErrHandle:
    ErrPrint "Draw", Err.Number, Err.Description
    Err.Clear
End Sub
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' This function drwa region based function
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function DRAWREGION(hdc As Variant) As RECT
Dim FrameBrush As LOGBRUSH, FillBrush As LOGBRUSH, ShapeRect As RECT

#If VBA7 Then
       Dim hPen As LongPtr
       Dim hBrush As LongPtr
       Dim hRgn As LongPtr
#Else
       Dim hPen As Long
       Dim hBrush As Long
       Dim hRgn As Long
#End If


Const ALTERNATE As Integer = 1

On Error GoTo ErrHandle
'------------------------------------------
If hdc = 0 Then Exit Function

    Select Case m_ShapeType
        Case GC_CIRCLE:      ' = 2
            hRgn = CreateEllipticRgn(m_POINTS(0).x - m_POINTS(1).x, _
                  m_POINTS(0).y - m_POINTS(1).x, _
                  m_POINTS(0).x + m_POINTS(1).x, m_POINTS(0).y + m_POINTS(1).x)
        Case GC_ELLIPS:      ' = 3
            hRgn = CreateEllipticRgn(m_POINTS(0).x, m_POINTS(0).y, m_POINTS(1).x, m_POINTS(1).y)
        Case GC_TRIANGLE:    ' = 4
            hRgn = CreatePolygonRgn(m_POINTS(0), m_Count + 1, ALTERNATE)
        Case GC_RECTANGLE:   ' = 5
            hRgn = CreateRectRgn(m_POINTS(0).x, m_POINTS(0).y, m_POINTS(1).x, m_POINTS(1).y)
        Case GC_POLIGON:     ' = 6
            hRgn = CreatePolygonRgn(m_POINTS(0), m_Count + 1, ALTERNATE)
        Case Else
            Err.Raise 1000, "Unproper DRAWREGION using"
    End Select
'-------------------------------------------
If Not m_BackTransparent Then
    FillBrush.lbColor = m_BackColor
    hBrush = CreateBrushIndirect(FillBrush)
    FillRgn hdc, hRgn, hBrush
End If

If Not m_BorderTransparent Then
    FrameBrush.lbColor = m_BorderColor
    hPen = CreateBrushIndirect(FrameBrush)
    FrameRgn hdc, hRgn, hPen, m_BorderWidth, m_BorderWidth
End If
'------------------------------------------
ExitHere:
    If hRgn <> 0 Then
        GetRgnBox hRgn, ShapeRect
        DeleteObject hRgn
    End If
    
    DRAWREGION = ShapeRect '!!!!!!!!!!!!!!!!!!
    
    If hBrush <> 0 Then DeleteObject hBrush
    If hPen <> 0 Then DeleteObject hPen
    Exit Function
'-----------------
ErrHandle:
    ErrPrint "DRAWREGION", Err.Number, Err.Description
    Err.Clear
End Function
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Draw Line -based pattern
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function DRAWFIGURE(hdc As Variant) As RECT
Dim LinePen As LOGPEN
#If VBA7 Then
       Dim hPen As LongPtr
       Dim hPen_Old As LongPtr
#Else
       Dim hPen As Long
       Dim hPen_Old As Long
#End If
On Error GoTo ErrHandle
'---------------------------------------
    If hdc = 0 Then Exit Function
    If m_BorderTransparent Then Exit Function
    
    LinePen.lopnColor = m_BorderColor: LinePen.lopnStyle = m_BorderStyle
    LinePen.lopnWidth.x = m_BorderStyle
    
    hPen = CreatePenIndirect(LinePen)
    hPen_Old = SelectObject(hdc, hPen)

'------------------------------------------
Select Case m_ShapeType
    Case GC_POINT:  ' = 0
            SetPixel hdc, m_POINTS(0).x, m_POINTS(0).y, m_BorderColor
 
    Case GC_LINE:   ' = 1
            MoveToEx hdc, m_POINTS(0).x, m_POINTS(0).y, m_POINTS(0)
            LineTo hdc, m_POINTS(1).x, m_POINTS(1).y
            
    Case GC_CURVE:  '= 7
            PolyBezier hdc, m_POINTS(0), m_Count + 1
    Case Else
            Err.Raise 1000, , "Unproper using of DRAWFIGURE"
End Select
'------------------------------------------
ExitHere:
    If hPen_Old <> 0 Then
        SelectObject hdc, hPen_Old
        DeleteObject hPen
    End If

    DRAWFIGURE = GetShapeRect() '!!!!!!!!!!!!!!!!!!
    Exit Function
'-----------------
ErrHandle:
    ErrPrint "DRAWFIGURE", Err.Number, Err.Description
    Err.Clear
End Function
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Draw String
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub DRAWSTRING(hdc As Variant, TextRect As RECT)
Dim lRes As Long, LF As LOGFONT
#If VBA7 Then
       Dim hFont As LongPtr
       Dim hFont_Old As LongPtr
#Else
       hFont As Long
       hFont_Old As Long
#End If

On Error GoTo ErrHandle
'---------------------------------------------
If m_Text = "" Then Exit Sub
If hdc = 0 Then Exit Sub

If (m_Width = 0) Or (m_Height = 0) Then Exit Sub

    LF = GetLogFONT2(hdc, m_FontName, m_FontSize, m_FontBold, _
             m_FontItalic, m_FontUnderline, m_FontCharSet, False, 0)
    
    hFont = CreateFontIndirect(LF)
    hFont_Old = SelectObject(hdc, hFont)
    
    SetTextColor hdc, m_FontColor
    SetBkMode hdc, TRANSPARENT
    
    DrawText hdc, m_Text, Len(m_Text), TextRect, AlignStyle(m_TextAlign)
'----------------------------------------------
    If hFont_Old <> 0 Then SelectObject hdc, hFont_Old
'------------------------------------------------------------------------------------
ExitHere:
    If hFont <> 0 Then DeleteObject hFont
    Exit Sub
'------------------------
ErrHandle:
    ErrPrint "DRAWSTRING", Err.Number, Err.Description
    Err.Clear
End Sub

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
    
             m_Left = DEFAULT_LEFT:  m_Top = DEFAULT_TOP
             m_Width = DEFAULT_WIDTH: m_Height = DEFAULT_HEIGHT

             m_Text = DEFAULT_CAPTION: m_TextAutoSize = DEFAULT_TEXT_AUTOSIZE
             m_TextAlign = DEFAULT_TEXT_ALIGN

             m_BackColor = DEFAULT_BACKCOLOR: m_BackTransparent = DEFAULT_BACKTRASPARENT

             m_BorderColor = DEFAULT_BORDERCOLOR: m_BorderWidth = DEFAULT_BORDERWIDTH
             m_BorderStyle = DEFAULT_BORDERSTYLE: m_BorderTransparent = DEFAULT_BORDERTRANSPARENT

             m_FontName = DEFAULT_FONTNAME: m_FontSize = DEFAULT_FONTSIZE
             m_FontColor = DEFAULT_TEXT_COLOR: m_FontBold = DEFAULT_FONTBOLD
             m_FontItalic = DEFAULT_FONTITALIC: m_FontUnderline = DEFAULT_FONTUNDERLINE
             m_FontCharSet = DEFAULT_CHARSET

             m_ShapeType = DEFAULT_SHAPE
             Call SetupShape(m_ShapeType)         ' --> m_POINTS(); m_Count

End Sub
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Error Handler
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function ErrPrint(FuncName As String, ErrNumber As Long, ErrDescription As String, Optional bDebug As Boolean = True, _
                                                                                                  Optional sModName As String = "cShape") As String
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
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Setup Shape
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub SetupShape(iShape As ShapeType)
    Select Case iShape
        Case GC_POINT:
             m_Count = 0: ReDim m_POINTS(0)
             m_POINTS(0).x = m_Left + m_Width / 2: m_POINTS(0).y = m_Top + m_Height / 2
             
        Case GC_LINE:
             m_Count = 1: ReDim m_POINTS(1)
             m_POINTS(0).x = m_Left: m_POINTS(0).y = m_Top
             m_POINTS(1).x = m_Left + m_Width: m_POINTS(1).y = m_Top + m_Height
             
        Case GC_CIRCLE:
             m_Count = 1: ReDim m_POINTS(1)
             m_POINTS(0).x = m_Left + m_Width / 2: m_POINTS(0).y = m_Top + m_Height / 2                      ' Center
             m_POINTS(1).x = IIf(m_Width <= m_Height, m_Width, m_Height): m_POINTS(1).y = m_POINTS(1).x      ' Raious
             
        Case GC_ELLIPS:
             m_Count = 1: ReDim m_POINTS(1)                                                                  ' Rectangle
             m_POINTS(0).x = m_Left: m_POINTS(0).y = m_Top
             m_POINTS(1).x = m_Left + m_Width: m_POINTS(1).y = m_Top + m_Height
        
        Case GC_TRIANGLE:
             m_Count = 2: ReDim m_POINTS(2)
             m_POINTS(0).x = m_Left: m_POINTS(0).y = m_Top + m_Width
             m_POINTS(1).x = m_Left + m_Width: m_POINTS(1).y = m_Top + m_Height
             m_POINTS(1).x = m_Left + m_Width / 2: m_POINTS(1).y = m_Top
        
        Case GC_RECTANGLE:
             m_Count = 1: ReDim m_POINTS(1)
             m_POINTS(0).x = m_Left: m_POINTS(0).y = m_Top                                                      ' Left   Top
             m_POINTS(1).x = m_Left: m_POINTS(1).y = m_Top + m_Width                                            ' Right  Bottom
        
        Case GC_POLIGON:
             m_Count = 5: ReDim m_POINTS(5)
             m_POINTS(0).x = m_Left + m_Width / 4: m_POINTS(0).y = m_Top
             m_POINTS(1).x = m_Left: m_POINTS(1).y = m_Top + m_Height / 2
             m_POINTS(2).x = m_POINTS(0).x: m_POINTS(2).y = m_Top + m_Height
             m_POINTS(3).x = m_Left + 2 * m_Width / 3: m_POINTS(3).y = m_Top + m_Height
             m_POINTS(4).x = m_Left + m_Width: m_POINTS(4).y = m_Top + m_Height / 2
             m_POINTS(5).x = m_POINTS(3).x: m_POINTS(5).y = m_Top
             
        Case GC_CURVE:
             m_Count = 5: ReDim m_POINTS(5)
             m_POINTS(0).x = m_Left: m_POINTS(0).y = m_Top
             m_POINTS(1).x = m_Left + m_Width / 4: m_POINTS(1).y = m_Top + m_Height / 4
             m_POINTS(2).x = m_Left + m_Width / 3: m_POINTS(2).y = m_Top + m_Height / 2
             m_POINTS(3).x = m_Left + m_Width / 2: m_POINTS(3).y = m_Top + m_Height
             m_POINTS(4).x = m_Left + 2 * m_Width / 3: m_POINTS(4).y = m_Top + m_Height / 2
             m_POINTS(5).x = m_Left + m_Width: m_POINTS(5).y = m_Top
             
    End Select
  
End Sub
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Calculate ShapeRect
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetShapeRect() As RECT
Dim RCT As RECT, I As Integer
Dim min_X As Long, max_X As Long, min_Y As Long, max_Y As Long

Const MIN_SIZE As Long = 5

    min_X = m_POINTS(0).x: max_X = m_POINTS(0).x
    min_Y = m_POINTS(0).y: max_Y = m_POINTS(0).y
    
    For I = 0 To m_Count
        If m_POINTS(I).x < min_X Then min_X = m_POINTS(I).x
        If m_POINTS(I).x > max_X Then max_X = m_POINTS(I).x
        If m_POINTS(I).y < min_Y Then min_Y = m_POINTS(I).y
        If m_POINTS(I).y > max_Y Then max_Y = m_POINTS(I).y
    Next I
'------------------------------------------
    RCT.Left = min_X: RCT.Top = min_Y
    If max_X - min_X < MIN_SIZE Then
       RCT.Right = RCT.Left + MIN_SIZE
    Else
       RCT.Right = max_X
    End If
    
    If max_Y - min_Y < MIN_SIZE Then
       RCT.Bottom = RCT.Top + MIN_SIZE
    Else
       RCT.Top = max_Y
    End If
'------------------------------------------
ExitHere:
    GetShapeRect = RCT '!!!!!!!!!!!!!!
End Function



Private Function GetLogFONT2(hdc As Variant, FontName As String, FontSize As Integer, FontBold As Boolean, FontItalic As Boolean, _
        FontUnderline As Boolean, CHARSET As Integer, Optional FontStrikethrough As Boolean, Optional iAngle As Long) As LOGFONT
Dim LF As LOGFONT
Dim iChar As Integer
Dim b() As Byte

On Error GoTo ErrHandle
'------------------------
   
   With LF
     b = StrConv(FontName, vbFromUnicode)
     
     For iChar = 1 To Len(FontName)
       .lfFaceName(iChar - 1) = b(iChar - 1)
     Next iChar
          
     .lfEscapement = iAngle
     .lfHeight = -MulDiv((FontSize), (GetDeviceCaps(hdc, LOGPIXELSY)), 72)
     .lfItalic = FontItalic
     
     If (FontBold) Then
       .lfWeight = FW_BOLD
     Else
       .lfWeight = FW_NORMAL
     End If
     .lfUnderline = FontUnderline
     .lfStrikeOut = FontStrikethrough
     .lfCharSet = CHARSET
   End With
'------------------------
ExitHere:
   GetLogFONT2 = LF '!!!!!!
   Exit Function
'---------------
ErrHandle:
   ErrPrint "ConvertFontToLF", Err.Number, Err.Description
   Err.Clear
End Function


Private Property Get m_Rect() As RECT
Dim zRect As RECT

   zRect.Left = m_Left: zRect.Top = m_Top
   zRect.Right = m_Left + m_Width: zRect.Bottom = m_Top + m_Height
 '----------------------------
   m_Rect = zRect '!!!!!!!
End Property

Private Function AlignStyle(iAlign As AlignType) As Long
Dim iRes As Long

Const DT_BOTTOM = &H8
Const DT_SINGLELINE = &H20
Const DT_CENTER = &H1
Const DT_LEFT = &H0

Const DT_NOCLIP = &H100
Const DT_RIGHT = &H2
Const DT_TOP = &H0
Const DT_VCENTER = &H4


Select Case iAlign
    Case GC_BYCOORDINATES:
         iRes = -1
    Case GC_TOPLEFT:
         iRes = DT_SINGLELINE Or DT_TOP Or DT_LEFT
    Case GC_TOPCENTER:
         iRes = DT_SINGLELINE Or DT_TOP Or DT_CENTER
    Case GC_TOPRIGHT:
         iRes = DT_SINGLELINE Or DT_TOP Or DT_RIGHT
    Case GC_CENTERLEFT:
         iRes = DT_SINGLELINE Or DT_VCENTER Or DT_LEFT
    Case GC_CENTERCENTER:
         iRes = DT_SINGLELINE Or DT_VCENTER Or DT_CENTER
    Case GC_CENTERRIGHT:
         iRes = DT_SINGLELINE Or DT_VCENTER Or DT_RIGHT
    Case GC_BOTTOMLEFT:
         iRes = DT_SINGLELINE Or DT_BOTTOM Or DT_LEFT
    Case GC_BOTTOMCENTER:
         iRes = DT_SINGLELINE Or DT_BOTTOM Or DT_CENTER
    Case GC_BOTTOMRIGHT:
         iRes = DT_SINGLELINE Or DT_BOTTOM Or DT_RIGHT
End Select
'--------------------------------------------------------
     AlignStyle = iRes '!!!!!!!!!!!!!
End Function
