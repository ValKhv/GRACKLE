VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cHolst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************************************************************************
'*********************************************************************************************************************************************************
' DESCRIPTION: The class implements the graphical context used by other classes within the GUI
'              As a canvas (i.e. Holst) is any suitable Windows object, which has hWnd and hDC, allowing you to output to it and other manipulations
' VERSION    : v 0.1 30.01.2018
' COPYRIGHT  : V.Khvatov DigitalXpert
'*********************************************************************************************************************************************************
'*********************************************************************************************************************************************************
Option Explicit

Implements ISubclass




'Windows and WinClass Styles
Private Const WIN_CLASS_STATIC = "Static"
Private Const CW_USEDEFAULT = &H80000000

Private Const WS_CHILD = &H40000000
Private Const WS_EX_TOOLWINDOW = &H80
Private Const WS_OVERLAPPEDWINDOW = &HCF0000
Private Const WS_EX_APPWINDOW = &H40000
Private Const WS_EX_WINDOWEDGE = &H100
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_EX_STATICEDGE = &H20000
Private Const WS_EX_TOPMOST As Long = &H8&
Private Const WS_EX_LAYERED = &H80000
Private Const WS_POPUP = &H80000000
Private Const WS_VISIBLE = &H10000000
Private Const SS_BITMAP = &HE


Private Const CS_HREDRAW = &H2
Private Const CS_VREDRAW = &H1
Private Const CS_PARENTDC = &H80

' Window Message
Private Const WM_SETFOCUS = &H7
Private Const WM_KILLFOCUS = &H8
Private Const WM_ENABLE = &HA
Private Const WM_SETREDRAW = &HB
Private Const WM_SETTEXT = &HC
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_PAINT = &HF
Private Const WM_CLOSE = &H10
Private Const WM_QUERYENDSESSION = &H11
Private Const WM_QUIT = &H12
Private Const WM_QUERYOPEN = &H13
Private Const WM_ERASEBKGND = &H14
Private Const WM_SYSCOLORCHANGE = &H15
Private Const WM_ENDSESSION = &H16
Private Const WM_SHOWWINDOW = &H18
Private Const WM_WININICHANGE = &H1A
Private Const WM_DEVMODECHANGE = &H1B
Private Const WM_ACTIVATEAPP = &H1C
Private Const WM_FONTCHANGE = &H1D
Private Const WM_TIMECHANGE = &H1E
Private Const WM_CANCELMODE = &H1F
Private Const WM_SETCURSOR = &H20
Private Const WM_MOUSEACTIVATE = &H21
Private Const WM_CHILDACTIVATE = &H22
Private Const WM_QUEUESYNC = &H23


' Standard Cursor IDs
Const IDC_ARROW = 32512&

' Color Types
Const COLOR_WINDOW = 5
Const COLOR_WINDOWFRAME = 6
Const COLOR_MENUTEXT = 7
Const COLOR_WINDOWTEXT = 8

'Show Window Commands
Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const SW_NORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOW = 5
Private Const SW_MINIMIZE = 6
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_RESTORE = 9
Private Const SW_SHOWDEFAULT = 10
Private Const SW_MAX = 10

Private Const LF_FACESIZE = 32

Private Const DT_SINGLELINE = &H20
Private Const DT_CENTER = &H1
Private Const DT_VCENTER = &H4

Private Const OPAQUE = 2
Private Const TRANSPARENT = 1

Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Private Const BS_SOLID = 0
Private Const BS_HATCHED = 2

Private Const HS_HORIZONTAL = 0     '-------------
Private Const HS_VERTICAL = 1       '|||||||||||||
Private Const HS_FDIAGONAL = 2      '\\\\\\\\\\\\\
Private Const HS_BDIAGONAL = 3      '/////////////
Private Const HS_CROSS = 4          '+++++++++++++
Private Const HS_DIAGCROSS = 5      'XXXXXXXXXXXXXX

Private Const ANSI_CHARSET = 0                                              ' ANSI character set
Private Const RUSSIAN_CHARSET = 204                                         ' ??? ????????
Private Const FW_BOLD = 700
Private Const FW_NORMAL = 400

Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90
Private Const TwipsPerInch = 1440

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

Private Type MSG
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

'*********************************************************************************************************************************************************
#If VBA7 Then
        Private Type WNDCLASSEX
            cbSize As Long
            style As Long
            lpfnWndProc As LongPtr
            cbClsExtra As Long
            cbWndExtra As Long
            hInstance As LongPtr
            hIcon As LongPtr
            hCursor As LongPtr
            hbrBackground As LongPtr
            lpszMenuName As String
            lpszClassName As String
            hIconSm As LongPtr
        End Type
    
        Private Type PAINTSTRUCT
            hdc As LongPtr
            fErase As Long
            rcPaint As RECT
            fRestore As Long
            fIncUpdate As Long
            rgbReserved(0 To 31) As Byte
        End Type
        
        
        Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
        Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
        Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
        Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
        Private Declare PtrSafe Function DestroyWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
        Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
        Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
        Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hWndChild As LongPtr, ByVal hWndNewParent As LongPtr) As LongPtr
        
        Private Declare PtrSafe Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (pcWndClassEx As WNDCLASSEX) As Integer
        Private Declare PtrSafe Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As LongPtr) As Long
        Private Declare PtrSafe Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
        Private Declare PtrSafe Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As LongPtr
        Private Declare PtrSafe Function InvalidateRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT, ByVal bErase As Long) As Long

        Private Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
        Private Declare PtrSafe Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hWnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
        Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
        Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
        Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As LongPtr) As LongPtr
        Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
        Private Declare PtrSafe Function BitBlt Lib "gdi32" (ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
        Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As LongPtr, ByVal lpCursorName As String) As LongPtr
        Private Declare PtrSafe Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As LongPtr, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
        
        Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
        Private Declare PtrSafe Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As LongPtr
        Private Declare PtrSafe Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As LongPtr
        Private Declare PtrSafe Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As LongPtr
        Private Declare PtrSafe Function SetTextColor Lib "gdi32" (ByVal hdc As LongPtr, ByVal crColor As Long) As Long
                
                
        Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
        Private Declare PtrSafe Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As LongPtr
        Private Declare PtrSafe Function Polygon Lib "gdi32" (ByVal hdc As LongPtr, lpPoint As POINTAPI, ByVal nCount As Long) As Long
        Private Declare PtrSafe Function FillRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr, ByVal hBrush As LongPtr) As Long
        Private Declare PtrSafe Function FrameRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr, ByVal hBrush As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As Long
        Private Declare PtrSafe Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As LongPtr
        Private Declare PtrSafe Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As LongPtr
        Private Declare PtrSafe Function GetRgnBox Lib "gdi32" (ByVal hRgn As LongPtr, lpRect As RECT) As Long
        
        
        Private Declare PtrSafe Function FillRect Lib "user32" (ByVal hdc As LongPtr, lpRect As RECT, ByVal hBrush As LongPtr) As Long
        Private Declare PtrSafe Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As LongPtr, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
        Private Declare PtrSafe Function MoveToEx Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
        Private Declare PtrSafe Function LineTo Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long


        Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal hObject As LongPtr) As LongPtr
        Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Boolean
        
        Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long


        Private Declare PtrSafe Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
        Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
        
        Private Declare PtrSafe Function SetBkColor Lib "gdi32" (ByVal hdc As LongPtr, ByVal crColor As Long) As Long
        Private Declare PtrSafe Function SetBkMode Lib "gdi32" (ByVal hdc As LongPtr, ByVal nBkMode As Long) As Long
        
        Private Declare PtrSafe Function BeginPaint Lib "user32" (ByVal hWnd As LongPtr, lpPaint As PAINTSTRUCT) As LongPtr
        Private Declare PtrSafe Function EndPaint Lib "user32" (ByVal hWnd As LongPtr, lpPaint As PAINTSTRUCT) As Long


        Private Type Holst
            iType As HolstType
            hWnd_Parent As LongPtr
            hWnd As LongPtr
            hdc As LongPtr
            hFont As LongPtr
            hFont_Old As LongPtr
            hBrush As LongPtr
            hBrush_Frame As LongPtr
            hPen As LongPtr
            hBrush_Old As LongPtr
            hPen_Old As LongPtr
            holstRECT As RECT
            HasDrawn As Boolean
            hProc As LongPtr
            ClassName As String
            WinTitle As String
            hInstance As LongPtr
        End Type
        
#Else
        Private Type WNDCLASSEX
            cbSize As Long
            style As Long
            lpfnWndProc As Long
            cbClsExtra As Long
            cbWndExtra As Long
            hInstance As Long
            hIcon As Long
            hCursor As Long
            hbrBackground As Long
            lpszMenuName As String
            lpszClassName As String
            hIconSm As Long
        End Type
    
        Private Type PAINTSTRUCT
            hdc As Long
            fErase As Long
            rcPaint As RECT
            fRestore As Long
            fIncUpdate As Long
            rgbReserved(0 To 31) As Byte
        End Type
        
        
        Private Declare Function GetDesktopWindow Lib "user32" () As Long
        Private Declare Function GetActiveWindow Lib "user32" () As Long
        Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
        Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
        Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
        Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
        Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
        Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
        
        Private Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (pcWndClassEx As WNDCLASSEX) As Integer
        Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
        Private Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
        Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long
        Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long

        Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
        Private Declare PtrSafe Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As MSG, ByVal hWnd As LongPtr, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
        Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
        Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
        Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
        Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
        Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
        Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As String) As Long
        Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
        
        Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
        Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
        Private Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long
        Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
        Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
        
        Private Declare PtrSafe Function FillRect Lib "user32" (ByVal hdc As LongPtr, lpRect As RECT, ByVal hBrush As LongPtr) As Long
        Private Declare PtrSafe Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As LongPtr
        Private Declare PtrSafe Function FrameRgn Lib "gdi32" (ByVal hdc As LongPtr, ByVal hRgn As LongPtr, ByVal hBrush As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As Long
        
        Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpstring As String, ByVal nCount As Long) As Long

        Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
        Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Boolean
        
        Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long


        Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
        Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
        
        Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
        Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long

        Private Type Holst
            iType As HolstType
            hWnd_Parent As Long
            hWnd As Long
            hdc As Long
            hFont As Long
            hFont_Old As Long
            hBrush As Long
            hPen As Long
            hBrush_Old As Long
            hPen_Old As Long
            holstRECT As RECT
            HasDrawn As Boolean
            hProc As Long
            ClassName As String
        End Type
        
#End If
'*********************************************************************************************************************************************************

Private m_Holst As Holst
Private m_Ready As Boolean
Private m_Font As StdFont
Private m_TextColor As Long
Private m_Test As Boolean


Private Const NEWWIN_LEFT As Long = 1000
Private Const NEWWIN_TOP As Long = 500
Private Const NEWWIN_WIDTH As Long = 500
Private Const NEWWIN_HEIGHT As Long = 400

'*********************************************************************************************************************************************************
Public Property Get isReady() As Boolean
        isReady = m_Ready
End Property

Public Property Get hWnd() As Variant
        hWnd = m_Holst.hWnd
End Property
Public Property Get hdc() As Variant
    If m_Holst.hWnd = 0 Then
       hdc = 0
    Else
       hdc = GetDC(m_Holst.hWnd)
    End If
End Property
Public Property Get iHolstType() As HolstType
        iHolstType = m_Holst.iType
End Property
Public Property Get hFont() As Variant
        hFont = m_Holst.hFont
End Property
Public Property Get hBrush() As Variant
        hBrush = m_Holst.hBrush
End Property

Public Property Get hPen() As Variant
        hPen = m_Holst.hPen
End Property

Public Property Get Left() As Long
        Left = m_Holst.holstRECT.Left
End Property
Public Property Get Top() As Long
        Top = m_Holst.holstRECT.Top
End Property
Public Property Get Right() As Long
        Right = m_Holst.holstRECT.Right
End Property
Public Property Get Bottom() As Long
        Bottom = m_Holst.holstRECT.Bottom
End Property
Public Property Get Width() As Long
        Width = m_Holst.holstRECT.Right - m_Holst.holstRECT.Left
End Property
Public Property Get height() As Long
        height = m_Holst.holstRECT.Bottom - m_Holst.holstRECT.Top
End Property


'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'=======================================================================================================================================================
' Set Holst
'=======================================================================================================================================================
Public Sub SetHolst(iHolst As HolstType, Optional iWinStyle As WinStyle = GC_WIN_TRASPARENT, Optional iHwnd As Variant, _
                                                             Optional sWinClass As String, Optional sWinTitle As String, Optional iWinProc As Variant, _
                                  Optional iLeft As Long = -1, Optional iTop As Long = -1, Optional IWidth As Long = -1, Optional iHeight As Long = -1, _
                                                                                                       Optional WinAlign As AlignType = GC_BYCOORDINATES)
On Error GoTo ErrHandle
'----------------------------------------
   If (m_Holst.hWnd <> 0) Then Err.Raise 1000, , "Once installed, the holst/canvas can not be changed"
   
   m_Holst.iType = iHolst
   '-----------------------------------------------
   Select Case m_Holst.iType
       Case GC_DESKTOP:
            m_Holst.hWnd = GetDesktopWindow()
            m_Holst.hdc = GetDC(m_Holst.hWnd)
            GetClientRect m_Holst.hWnd, m_Holst.holstRECT

       Case GC_ACTIVE_WINDOW:
            m_Holst.hWnd = GetActiveWindow()
            m_Holst.hdc = GetDC(m_Holst.hWnd)
            GetClientRect m_Holst.hWnd, m_Holst.holstRECT
            
       Case GC_SPECIFIC_FORM, GC_SPECIFIC_CONTROL:
            If IsMissing(iHwnd) Then Err.Raise 1000, , "To set a client form as a HOLST you need provide its hWnd"
            #If VBA7 Then
                m_Holst.hWnd = CLngPtr(iHwnd)
            #Else
                m_Holst.hWnd = CLng(iHwnd)
            #End If
    
            m_Holst.hdc = GetDC(m_Holst.hWnd)
            GetClientRect m_Holst.hWnd, m_Holst.holstRECT

       Case GC_NEW_WINDOW:
            If Not IsMissing(iHwnd) Then
                m_Holst.hWnd_Parent = iHwnd
            Else
                m_Holst.hWnd_Parent = GetDesktopWindow()  ' Couldbe RESET by CreateNewWindow
            End If
            If Not IsMissing(iWinProc) Then m_Holst.hProc = iWinProc
            m_Holst.WinTitle = sWinTitle: m_Holst.ClassName = sWinClass
            SetNewWinSizeAndPosition iLeft, iTop, IWidth, iHeight, WinAlign
            
            Call CreateNewWindow(iWinStyle, sWinTitle, sWinClass)
   End Select
'----------------------------------------
ExitHere:
   If m_Holst.hWnd <> 0 Then m_Ready = True
   Exit Sub
'-----------
ErrHandle:
   ErrPrint "SetHolst", Err.Number, Err.Description
   Err.Clear
End Sub

Public Sub Clean()
         
End Sub

'=======================================================================================================================================================
' Setup Pen
'=======================================================================================================================================================
Public Function GetPen(Optional iColor As Long = vbRed, Optional IWidth As Long = 1, Optional iStyle As LineType = GC_LSOLID) As Variant
Dim lp As LOGPEN

On Error GoTo ErrHandle
'-------------------------------
If Not m_Ready Then Exit Function
If m_Holst.hdc = 0 Then Err.Raise 1000, , "Wrong Context to setup the pen"

    If m_Holst.hPen <> 0 Then DeleteObject m_Holst.hPen
    
    lp = Get_LogPEN(iColor, IWidth, iStyle)
    m_Holst.hPen = CreatePenIndirect(lp)
    
    If m_Holst.hPen = 0 Then Err.Raise 1000, , "Something wrong with pen creation"
    
    If m_Holst.hPen_Old = 0 Then
        m_Holst.hPen_Old = SelectObject(m_Holst.hdc, m_Holst.hPen)
    Else
        SelectObject m_Holst.hdc, m_Holst.hPen
    End If
'-----------------------------
ExitHere:
    GetPen = m_Holst.hPen '!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint "GetPen", Err.Number, "DLLERR:" & Err.LastDllError & " " & Err.Description
    Err.Clear
End Function

'=======================================================================================================================================================
' Setup Brush (FillBrush)
'=======================================================================================================================================================
Public Function GetFillBrush(Optional iColor As Long = vbBlue, Optional bSolid As Boolean = True, Optional iStyle As BackType = GC_SOLID) As Variant
Dim lB As LOGBRUSH

On Error GoTo ErrHandle
'-------------------------------
If Not m_Ready Then Exit Function
If m_Holst.hdc = 0 Then Err.Raise 1000, , "Wrong Context to setup the brush"

    If m_Holst.hBrush <> 0 Then DeleteObject m_Holst.hBrush

    If bSolid Then
       m_Holst.hBrush = CreateSolidBrush(iColor)
    Else
       lB = Get_LogBrush(iColor, iStyle)
       m_Holst.hBrush = CreateBrushIndirect(lB)
    End If

    If m_Holst.hBrush = 0 Then Err.Raise 1000, , "Something wrong with brush creation"
    
    If m_Holst.hBrush_Old = 0 Then
        m_Holst.hBrush_Old = SelectObject(m_Holst.hdc, m_Holst.hBrush)
    Else
        SelectObject m_Holst.hdc, m_Holst.hBrush
    End If
'-----------------------------
ExitHere:
    GetFillBrush = m_Holst.hBrush  '!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint "GetFillBrush", Err.Number, "DLLERR:" & Err.LastDllError & " " & Err.Description
    Err.Clear
End Function
'=======================================================================================================================================================
' Setup Brush (FrameBrush)
'=======================================================================================================================================================
Public Function GetFrameBrush(Optional iColor As Long = vbBlue, Optional bSolid As Boolean = True, Optional iStyle As BackType = GC_SOLID) As Variant
Dim lB As LOGBRUSH

On Error GoTo ErrHandle
'-------------------------------
If Not m_Ready Then Exit Function
If m_Holst.hdc = 0 Then Err.Raise 1000, , "Wrong Context to setup the brush"

    If m_Holst.hBrush_Frame <> 0 Then DeleteObject m_Holst.hBrush_Frame

    If bSolid Then
       m_Holst.hBrush_Frame = CreateSolidBrush(iColor)
    Else
       lB = Get_LogBrush(iColor, iStyle)
       m_Holst.hBrush_Frame = CreateBrushIndirect(lB)
    End If

    If m_Holst.hBrush_Frame = 0 Then Err.Raise 1000, , "Something wrong with brush creation"
    
'-----------------------------
ExitHere:
    GetFrameBrush = m_Holst.hBrush_Frame   '!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint "GetFrameBrush", Err.Number, "DLLERR:" & Err.LastDllError & " " & Err.Description
    Err.Clear
End Function
'=======================================================================================================================================================
' Calculate TextSize fitted to rect
'=======================================================================================================================================================
Public Function CalculateFontSizeToFitRect(sText As String, fnt As StdFont, IWidth As Long, iHeight As Long, _
                                                                          Optional minSize As Integer = 5, Optional MAXSIZE As Integer = 60) As Integer
Dim fitRect As RECT

If m_Holst.hdc = 0 Then Exit Function
If sText = "" Then Exit Function
    fitRect.Right = IWidth: fitRect.Bottom = iHeight
    CalculateFontSizeToFitRect = FontSizeToRect(sText, fitRect, fnt, m_Holst.hdc, minSize, MAXSIZE)
End Function
'=======================================================================================================================================================
' Setup Font
'=======================================================================================================================================================
Public Function GetFont(Optional FontName As String = "Arial", Optional iTextColor As Long = vbBlack, Optional FontSize As Integer = 18, _
                                       Optional iBold As Boolean = False, Optional iItalic As Boolean = False, Optional iUnderline As Boolean = False, _
                                                            Optional iCharset As Integer = 204, Optional iAngle As Long, Optional AutoSize As Boolean, _
                                                                  Optional IWidth As Long, Optional iHeight As Long, Optional Text As String) As Variant
Dim LF As LOGFONT

On Error GoTo ErrHandle
'-------------------------------
If Not m_Ready Then GoTo ExitHere
If m_Holst.hdc = 0 Then Err.Raise 1000, , "Wrong Context to setup the brush"

If m_Holst.hFont <> 0 Then DeleteObject m_Holst.hFont

    m_Font.Name = FontName: m_Font.SIZE = FontSize: m_Font.Bold = iBold: m_Font.Italic = iItalic
    m_Font.Underline = iUnderline: m_Font.CHARSET = iCharset
    
    If AutoSize Then
       If (iHeight * IWidth <> 0) And (Text <> vbNullString) Then
                m_Font.SIZE = CalculateFontSizeToFitRect(Text, m_Font, IWidth, iHeight)
       End If
    End If
    

    LF = Get_LogFONT(m_Font, m_Holst.hdc, iAngle)
    m_Holst.hFont = CreateFontIndirect(LF)
   

    If m_Holst.hFont = 0 Then Err.Raise 1000, , "Something wrong with font creation"
    
    If m_Holst.hFont_Old = 0 Then
        m_Holst.hFont_Old = SelectObject(m_Holst.hdc, m_Holst.hFont)
    Else
        SelectObject m_Holst.hdc, m_Holst.hFont
    End If
    
    If iTextColor <> m_TextColor Then
        m_TextColor = iTextColor
        SetTextColor m_Holst.hdc, m_TextColor
    End If
'-----------------------------
ExitHere:
    GetFont = m_Holst.hFont   '!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint "GetFont", Err.Number, "DLLERR:" & Err.LastDllError & " " & Err.Description
    Err.Clear: Resume ExitHere
End Function


'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
         m_Holst.iType = GC_UNDEFINED
         Set m_Font = New StdFont
End Sub
Private Sub Class_Terminate()
         Set m_Font = Nothing
    With m_Holst
         If .hBrush <> 0 Then DeleteObject .hBrush
         If .hBrush_Frame <> 0 Then DeleteObject .hBrush_Frame
         If .hPen <> 0 Then DeleteObject .hPen
         If .hFont <> 0 Then DeleteObject .hFont
         
         If .HasDrawn Then Clean
         
         If .hBrush_Old <> 0 Then SelectObject .hdc, .hBrush_Old
         If .hBrush <> 0 Then DeleteObject .hBrush
         If .hPen_Old <> 0 Then SelectObject .hdc, .hPen_Old
         If .hPen <> 0 Then DeleteObject .hPen
         If .hFont_Old <> 0 Then SelectObject .hdc, .hFont_Old
         If .hFont <> 0 Then DeleteObject .hFont
         
         If .hdc <> 0 And .hWnd <> 0 Then
            If .hBrush_Old <> 0 Then SelectObject .hdc, .hBrush_Old
            If .hFont_Old <> 0 Then SelectObject .hdc, .hFont_Old
            If .hPen_Old <> 0 Then SelectObject .hdc, .hPen_Old
            ReleaseDC .hWnd, .hdc
         End If
         
         If .iType = GC_NEW_WINDOW Then DestroyWindow .hWnd
         If (.ClassName <> "") And (.ClassName <> WIN_CLASS_STATIC) Then UnregisterClass .ClassName, .hInstance
   End With
   
End Sub
'=======================================================================================================================================================
' Test By Default
'=======================================================================================================================================================
Public Sub DrawBox(iLeft As Long, iTop As Long, IWidth As Long, iHeight As Long, Optional bCenter As Boolean, _
                                              Optional BackColor As Long = -1, Optional BorderColor As Long = -1, Optional BorderType As LineType = 0, _
                                                   Optional BorderWidth As Long = 1, Optional sText As String, Optional FontColor As Long = vbBlack, _
                                                                    Optional FontSize As Integer = 12, Optional AlignText As AlignType = GC_CENTERCENTER)
Dim RectBox As RECT

On Error GoTo ErrHandle
'-----------------------------------
If Not m_Ready Then Exit Sub

If m_Holst.hdc <> 0 Then ReleaseDC m_Holst.hWnd, m_Holst.hdc
m_Holst.hdc = GetDC(m_Holst.hWnd)
If m_Holst.hdc = 0 Then Err.Raise 1000, , "Wrong DC handler"
    
    RectBox.Left = iLeft: RectBox.Top = iTop
    RectBox.Right = iLeft + IWidth: RectBox.Bottom = iTop + iHeight

    If bCenter Then Call CentrRect(RectBox, m_Holst.holstRECT)
    
    If BackColor >= 0 Then
        GetFillBrush BackColor
        FillRect m_Holst.hdc, RectBox, m_Holst.hBrush
    End If

    If BorderColor > 0 Then
        GetPen BorderColor, BorderWidth, BorderType
   
        DrawLine RectBox.Left, RectBox.Top, RectBox.Right, RectBox.Top
        DrawLine RectBox.Right, RectBox.Top, RectBox.Right, RectBox.Bottom
        DrawLine RectBox.Right, RectBox.Bottom, RectBox.Left, RectBox.Bottom
        DrawLine RectBox.Left, RectBox.Bottom, RectBox.Left, RectBox.Top
   
    End If

    If sText <> "" Then
            Call GetFont(, FontColor, FontSize)
           'SetBkColor m_Holst.hdc, pBox.BackColor
           SetBkMode m_Holst.hdc, TRANSPARENT
           DrawText m_Holst.hdc, sText, Len(sText), RectBox, DT_SINGLELINE Or DT_CENTER Or DT_VCENTER 'AlignText
     End If
'----------------------------------------------------------------------
ExitHere:
    ReleaseDC m_Holst.hWnd, m_Holst.hdc
    m_Holst.hdc = 0
    Exit Sub
'---------------
ErrHandle:
    ErrPrint "DrawLabel", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Sub
'=======================================================================================================================================================
' DrawLine
'=======================================================================================================================================================
Public Sub DrawLine(X1 As Long, Y1 As Long, x2 As Long, Y2 As Long)
Dim mPoint As POINTAPI

On Error GoTo ErrHandle
'----------------------------
If Not m_Ready Then Exit Sub
If m_Holst.hdc = 0 Then Err.Raise 1000, , "Wronng Context Descriptor"

            mPoint.x = X1: mPoint.y = Y1
            MoveToEx m_Holst.hdc, X1, Y1, mPoint
            LineTo m_Holst.hdc, x2, Y2
ExitHere:
     Exit Sub
ErrHandle:
     ErrPrint "DrawLine", Err.Number, Err.Description
     Err.Clear
End Sub

'=======================================================================================================================================================
' Start Event Handling
'=======================================================================================================================================================
Public Sub GO(Optional WinMsgProxy As Integer = 1, Optional bTest As Boolean)
Dim myMsg As MSG

On Error GoTo ErrHandle
'------------------------------------------------
    If Not m_Ready Then Exit Sub
    If m_Holst.hWnd = 0 Then Exit Sub
    
    If m_Holst.iType = GC_NEW_WINDOW Then
        If WinMsgProxy = 1 Then
            SubWnd m_Holst.hWnd, Me           ' External Interface
        ElseIf WinMsgProxy = 2 Then   ' Message Loop
            Call MessageLoop(m_Holst.hWnd)
        ElseIf WinMsgProxy = 3 Then   ' Both - Message Loop and External Interface
            SubWnd m_Holst.hWnd, Me
            Call MessageLoop(m_Holst.hWnd)
        End If
    Else
        Call Draw
    End If
'------------------------------
ExitHere:
    Exit Sub
'-------------
ErrHandle:
   ErrPrint "GO", Err.Number, Err.Description
   Err.Clear
End Sub


'=======================================================================================================================================================
' External Interface Message Management
'=======================================================================================================================================================
#If VBA7 Then
    Private Function ISubclass_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef bHandled As Boolean) As Long '

    Select Case uMsg
       Case WM_PAINT:
           'Debug.Print "PAINT"
           Call Draw
           DoEvents
           'Call TestHolst(m_Holst.hWND)

       Case WM_CLOSE, WM_QUIT:
                     MsgBox "QUIT"
           If m_Holst.hWnd <> 0 Then DestroyWindow m_Holst.hWnd
           If m_Holst.ClassName <> "" Then UnregisterClass m_Holst.ClassName, Application.hWndAccessApp
       Case Else

      End Select


        '-- Do some stuff with the message


    End Function
#Else
    Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long
        '''''''''''''''''''''''''''''''''''''''''''''''''
        ' The code that runs based on the values of iMsg,
        ' WParam, and LParam.
        '''''''''''''''''''''''''''''''''''''''''''''''''
    End Function
#End If

Private Sub Draw()
        Dim rt As RECT
        Dim hdc As LongPtr
        Dim PS As PAINTSTRUCT, ShapeRect As RECT
        
        Debug.Print "DRAW"
        GetClientRect m_Holst.hWnd, rt
        
        hdc = BeginPaint(m_Holst.hWnd, PS)
        'DrawText hdc, "DRAW!", Len("DRAW!"), rt, DT_CENTER
        
 '           iShape = 1: TEXT = "TEST": FillColor = &HFFFF00
 '           ihWnd = CLngPtr(hwnd)
 '           ihDC = GetDC(ihWnd)
            GetClientRect m_Holst.hWnd, ShapeRect
            Call DrawShape(m_Holst.hWnd, hdc, ShapeRect, 1, "TEST", &HFFFF00)



        
        EndPaint m_Holst.hWnd, PS
        
        ' since we handled this message, return 0. dont let the
        ' DefWindowProc handle it
    
End Sub
'=======================================================================================================================================================
' MESSAGE LOOP
'=======================================================================================================================================================
Private Function MessageLoop(hWnd As Variant) As Long
Dim myMsg As MSG

On Error GoTo ErrHandle
'------------------------------------------------
    While GetMessage(myMsg, hWnd, 0, 0) <> 0      ' waiting for PostQuitMessage to be called to break out
       Select Case myMsg.message
       Case WM_PAINT:
       
       Case Else

      End Select
        
        '----------------------------------------
        TranslateMessage myMsg
        DispatchMessage myMsg
      Wend
'------------------------------------------------
ExitHere:
    MessageLoop = myMsg.wParam
    Exit Function
'----------------
ErrHandle:
    ErrPrint "MessageLoop", Err.Number, Err.Description
    Err.Clear
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
                                                                                                     Optional sModName As String = "cHolst") As String
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
'--------------------------------------------------------------------------------------------------------------------------------------------------------
' Create New Window !!!!
'---------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function CreateNewWindow(Optional iWinStyle As WinStyle = GC_WIN_TRASPARENT, Optional ByVal szWindowTitle As String, _
                                                                                                      Optional ByVal szWindowClass As String) As Boolean
Dim bClassReg As Boolean, iexWStyle As Long, iWStyle As Long, sWinTitle As String
Dim IWidth As Long, iHeight As Long, iLeft As Long, iTop As Long

Const DirectCommand As Boolean = True  'DEBUGGING ONLY

On Error GoTo ErrHandle
'----------------------------------------
If m_Holst.iType <> GC_NEW_WINDOW Then Exit Function


' PRE-WINDOW
With m_Holst
    iLeft = .holstRECT.Left:   IWidth = .holstRECT.Right - .holstRECT.Left
    iTop = .holstRECT.Top:  iHeight = .holstRECT.Bottom - .holstRECT.Top
    If (IWidth <= 0) Or (iHeight <= 0) Then Err.Raise 1000, , "Wrong Size Parameters for new window"
    If (iTop < 0) Or (iLeft < 0) Then Err.Raise 1000, , "Wrong Size Parameters for new window"

    Select Case iWinStyle
        Case GC_WIN_TRASPARENT:
            .ClassName = WIN_CLASS_STATIC
            .WinTitle = vbNullString
            If .hWnd_Parent = 0 Then .hWnd_Parent = GetDesktopWindow()
            .hInstance = 0&
          
            iexWStyle = WS_EX_TOOLWINDOW
            iWStyle = WS_CHILD
            
            If DirectCommand Then

'                     m_Holst.hwnd = CreateWindowEx(iexWStyle, .ClassName, _
'                           vbNullString, iWStyle, iLeft, iTop, _
'                           iWidth, iHeight, m_Holst.hWnd_Parent, 0, .hInstance, 0)
'
'
'                            m_Holst.hdc = GetDC(hwnd)
    
'                            SetParent m_Holst.hwnd, m_Holst.hWnd_Parent
'                            SetBkMode hdc, 1
'
                
                 
            End If
        
        
        Case GC_WIN_SPLASH:
            .ClassName = WIN_CLASS_STATIC
            .WinTitle = vbNullString
            
            .hInstance = GetDesktopWindow
          
            iexWStyle = WS_EX_STATICEDGE Or WS_EX_TOPMOST
            iWStyle = WS_POPUP Or WS_VISIBLE
            
            If DirectCommand Then
                .hWnd = CreateWindowEx(WS_EX_STATICEDGE Or WS_EX_TOPMOST, "static", "", _
                                   WS_POPUP Or WS_VISIBLE, iLeft, iTop, _
                                  IWidth, iHeight, .hWnd_Parent, 0&, .hWnd_Parent, 0&)
            End If
        Case GC_WIN_FULLREDRAW:
            .ClassName = IIf(szWindowClass <> "", szWindowClass, "GRACLEWIN")
            .WinTitle = IIf(szWindowTitle <> "", szWindowTitle, "New Window")
            .hInstance = hWndApplication
            .hWnd_Parent = 0&
          
            '.bNewClassReg = CreateWinClass(.hInstance, .ClassName)
            'If Not .bNewClassReg Then Exit Function
          
            iexWStyle = WS_EX_APPWINDOW Or WS_EX_WINDOWEDGE
            iWStyle = WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or WS_OVERLAPPEDWINDOW
            
            If DirectCommand Then
                  m_Holst.hWnd = CreateNewWindow2(GetWinHook, szWindowClass, szWindowTitle, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT)
                        
                        
                        GoTo ExitHere
            End If
          
    End Select
'---------------------------------------------
'ACTUAL WINDOW CREATION
     If Not DirectCommand Then
         .hWnd = CreateWindowEx(iexWStyle, .ClassName, .WinTitle, iWStyle, _
                              .holstRECT.Left, .holstRECT.Top, .holstRECT.Right - .holstRECT.Left, _
                              .holstRECT.Bottom - .holstRECT.Top, .hWnd_Parent, 0&, .hInstance, 0)
     End If
                    
        If .hWnd = 0 Then Err.Raise 1000, , "Failed to create the window!"
'-----------------------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------------------
' POST-WINDOW
        If .hWnd_Parent <> 0 Then SetParent .hWnd, .hWnd_Parent
        .hdc = GetDC(.hWnd)

        Select Case iWinStyle
            Case GC_WIN_TRASPARENT:
                  SetBkMode .hdc, 1
                  UpdateWindow .hWnd
            Case GC_WIN_SPLASH:
            Case GC_WIN_FULLREDRAW:
        End Select
        '---------------------------------
        ShowWindow .hWnd, SW_SHOW
        
End With
         'If m_Holst.bTestDRAW Then Call TestBox
'----------------------------
ExitHere:
    CreateNewWindow = True
    Exit Function
'-----------------
ErrHandle:
    ErrPrint "CreateNewWindow", Err.Number, "DLLERR: " & Err.LastDllError & " " & Err.Description
    Err.Clear
    If bClassReg Then UnregisterClass m_Holst.ClassName, m_Holst.hWnd_Parent
End Function

Private Function CreateNewWindow2(ByVal MyWndProc As Long, ByVal szWindowClass As String, ByVal szWindowTitle As String, _
                  ByVal x As Long, ByVal y As Long, ByVal CX As Long, ByVal CY As Long) As Variant
Dim I As Long

    
    If Not RegisterWinClass(Application.hWndAccessApp, szWindowClass, MyWndProc) Then
        MsgBox "Failed to register window!"
        CreateNewWindow2 = -1
        Exit Function
    End If
    
    m_Holst.hWnd = CreateWindowEx(WS_EX_APPWINDOW Or WS_EX_WINDOWEDGE, _
                              szWindowClass, _
                              szWindowTitle, _
                              WS_CLIPSIBLINGS Or WS_CLIPCHILDREN Or WS_OVERLAPPEDWINDOW, _
                              x, y, CX, CY, 0, 0, Application.hWndAccessApp, 0)
                              
    If m_Holst.hWnd = 0 Then
        MsgBox "Failed to create the window!"
        UnregisterClass szWindowClass, Application.hWndAccessApp
        CreateNewWindow2 = -1
        Exit Function
    End If
    
    ' show the window
    UpdateWindow m_Holst.hWnd
    ShowWindow m_Holst.hWnd, SW_SHOW
    
    CreateNewWindow2 = m_Holst.hWnd
End Function

'--------------------------------------------------------------------------------------------------------------------------------------------------------
' Create Win Class
'---------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function RegisterWinClass(ihInstance As Variant, Optional ByVal szWindowClass As String, Optional ByVal hWndProc As Variant) As Boolean
Dim bRes As Boolean, wcex As WNDCLASSEX
Dim sClassName As String

On Error GoTo ErrHandle
'--------------------------------------------
#If VB7 Then
    Dim ihWnndProc As LongPtr
    If Not IsMissing(hWndProc) Then ihWndProc = CLngPtr(hWndProc)
#Else
    Dim ihWndProc As Long
    If Not IsMissing(hWndProc) Then ihWndProc = CLng(hWndProc)
#End If
    If ihWndProc = 0 Then Err.Raise 1000, , "Can't registrate class without a proper handler "
    sClassName = IIf(szWindowClass <> "", szWindowClass, "GRC_WINCLASS")


With wcex
    .cbSize = LenB(wcex)
    .style = CS_HREDRAW Or CS_VREDRAW Or CS_PARENTDC
    .lpfnWndProc = ihWndProc
    .cbClsExtra = 0
    .cbWndExtra = 0
    .hInstance = ihInstance
    .hIcon = 0
    .hCursor = LoadCursor(0, IDC_ARROW)
    .hbrBackground = COLOR_WINDOW + 1
    .lpszMenuName = vbNullString
    .lpszClassName = sClassName
    .hIconSm = 0
End With

If RegisterClassEx(wcex) = 0 Then Err.Raise 1000, , "Failed to register class window"
bRes = True
'--------------------------------------------
ExitHere:
    RegisterWinClass = bRes '!!!!!!!!!!!
    Exit Function
'---------------
ErrHandle:
    ErrPrint "RegisterWinClass", Err.Number, "DLLERR:" & Err.LastDllError & "  " & Err.Description
    Err.Clear
End Function



'--------------------------------------------------------------------------------------------------------------------------------------------------------
' Only for New Windows
'---------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub SetNewWinSizeAndPosition(iLeft As Long, iTop As Long, IWidth As Long, iHeight As Long, Optional WinAlign As AlignType = GC_BYCOORDINATES)
Dim ParentREct As RECT, aWidth As Long, aHeight As Long, pWidth As Long, pHeight As Long

On Error GoTo ErrHandle
'----------------------------------
With m_Holst
   If .hWnd_Parent = 0 Then Err.Raise 1000, "Can't validate size and position"
   GetClientRect .hWnd_Parent, ParentREct
   
   '---------------------------------------------------------------------------
   aWidth = IIf(IWidth <> -1, IWidth, NEWWIN_WIDTH): aHeight = IIf(iHeight <> -1, iHeight, NEWWIN_HEIGHT)
   pWidth = ParentREct.Right - ParentREct.Left: pHeight = ParentREct.Bottom - ParentREct.Top
  
  
  Select Case WinAlign
                Case GC_BYCOORDINATES:
                    .holstRECT.Left = IIf(iLeft <> -1, iLeft, NEWWIN_LEFT)
                    .holstRECT.Top = IIf(iTop <> -1, iTop, NEWWIN_TOP)
                Case GC_TOPLEFT:
                    .holstRECT.Left = ParentREct.Left:
                    .holstRECT.Top = ParentREct.Top
                Case GC_TOPCENTER:
                    .holstRECT.Left = ParentREct.Left + (pWidth - aWidth) / 2
                    .holstRECT.Top = ParentREct.Top
                Case GC_TOPRIGHT:
                    .holstRECT.Left = ParentREct.Right - aWidth
                    .holstRECT.Top = ParentREct.Top
                Case GC_CENTERLEFT:
                    .holstRECT.Left = ParentREct.Left
                    .holstRECT.Top = ParentREct.Top + (pHeight - aHeight) / 2
                Case GC_CENTERCENTER:
                    .holstRECT.Left = ParentREct.Left + (pWidth - aWidth) / 2
                    .holstRECT.Top = ParentREct.Top + (pHeight - aHeight) / 2
                Case GC_CENTERRIGHT:
                    .holstRECT.Left = ParentREct.Right - aWidth
                    .holstRECT.Top = ParentREct.Top + (pHeight - aHeight) / 2
                Case GC_BOTTOMLEFT:
                    .holstRECT.Left = ParentREct.Left
                    .holstRECT.Top = ParentREct.Bottom - aHeight
                Case GC_BOTTOMCENTER:
                    .holstRECT.Left = ParentREct.Left + (pWidth - aWidth) / 2
                    .holstRECT.Top = ParentREct.Bottom - aHeight
                Case GC_BOTTOMRIGHT:
                    .holstRECT.Left = ParentREct.Right - aWidth
                    .holstRECT.Top = ParentREct.Bottom - aHeight
  End Select
 
                .holstRECT.Right = .holstRECT.Left + aWidth
                .holstRECT.Bottom = .holstRECT.Top + aHeight
 
End With
'---------------------------------
ExitHere:
    Exit Sub
'--------------
ErrHandle:
    ErrPrint "SetNewWinSizeAndPosition", Err.Number, Err.Description
    Err.Clear
End Sub
'--------------------------------------------------------------------------------------------------------------------------------------------------------
' Place SmallRect in center of BigRect (recalculate position of small rect)
'---------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub CentrRect(smallRect As RECT, bigRect As RECT, Optional bRelative As Boolean = True)
Dim IWidth As Long, iHeight As Long, iTop As Long, iLeft As Long

    IWidth = smallRect.Right - smallRect.Left: iHeight = smallRect.Bottom - smallRect.Top
    iTop = IIf(bRelative, 0, bigRect.Top): iLeft = IIf(bRelative, 0, bigRect.Left)
    
    smallRect.Left = iLeft + (bigRect.Right - bigRect.Left - IWidth) / 2
    smallRect.Top = iTop + (bigRect.Bottom - bigRect.Top - iHeight) / 2
    smallRect.Right = smallRect.Left + IWidth: smallRect.Bottom = smallRect.Top + iHeight

End Sub
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Get Log Brush
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function Get_LogBrush(iColor As Long, iStyle As BackType) As LOGBRUSH
Dim b As LOGBRUSH

On Error GoTo ExitHere
'-----------------------------------------
    b.lbColor = iColor
Select Case iStyle
    Case GC_SOLID:                      ' = 0
         b.lbStyle = 0: b.lbHatch = HS_HORIZONTAL
    Case GC_GRADIENT:                   ' = -2
         b.lbStyle = BS_SOLID: b.lbHatch = HS_HORIZONTAL
    Case GC_HATCHED_HORIZONTAL:         ' = -1     '-------------
         b.lbStyle = BS_HATCHED: b.lbHatch = HS_HORIZONTAL
    Case GC_HATCHED_VERTICAL, GC_HATCHED_FDIAGONAL, _
         GC_HATCHED__BDIAGONAL, GC_HATCHED_CROSS, GC_HATCHED_DIAGCROSS:
         b.lbStyle = BS_HATCHED: b.lbHatch = iStyle
End Select

'---------------------------
ExitHere:
    Get_LogBrush = b '!!!!!!!!!!!!!!
    Exit Function
'------------
ErrHandle:
    ErrPrint "Get_LogBrush", Err.Number, Err.Describtion
    Err.Clear
End Function
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Get Log Pen
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function Get_LogPEN(iColor As Long, IWidth As Long, iStyle As LineType) As LOGPEN
Dim p As LOGPEN

On Error GoTo ExitHere
'-----------------------------------------


    p.lopnColor = iColor
    p.lopnStyle = iStyle
    p.lopnWidth.x = IWidth
    
'---------------------------
ExitHere:
    Get_LogPEN = p '!!!!!!!!!!!!!!
    Exit Function
'------------
ErrHandle:
    ErrPrint "Get_LogPEN", Err.Number, Err.Describtion
    Err.Clear
End Function

Private Function Get_LogFONT(fnt As StdFont, hdc As Variant, Optional iAngle As Long) As LOGFONT
Dim sFontName As String, LF As LOGFONT
Dim iChar As Integer
Dim b() As Byte

On Error GoTo ErrHandle
'------------------------
   
   With LF
     b = StrConv(fnt.Name, vbFromUnicode)   ' FNT.Name & vbNullChar
     
     For iChar = 1 To Len(fnt.Name)
       .lfFaceName(iChar - 1) = b(iChar - 1)
     Next iChar
          
     .lfEscapement = iAngle
     .lfHeight = -MulDiv((fnt.SIZE), (GetDeviceCaps(hdc, LOGPIXELSY)), 72)
     .lfItalic = fnt.Italic
     
     If (fnt.Bold) Then
       .lfWeight = FW_BOLD
     Else
       .lfWeight = FW_NORMAL
     End If
     .lfUnderline = fnt.Underline
     .lfStrikeOut = fnt.Strikethrough
     .lfCharSet = fnt.CHARSET
   End With
'------------------------
ExitHere:
   Get_LogFONT = LF '!!!!!!
   Exit Function
'---------------
ErrHandle:
   ErrPrint "ConvertFontToLF", Err.Number, Err.Description
   Err.Clear
End Function


'----------------------------------------------------------------------------------------------------------------------------------------------------------
' Calculate Font Size to fit Rectangle
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function FontSizeToRect(sText As String, R As RECT, fnt As StdFont, DC As Variant, Optional minSize As Integer = 5, Optional MAXSIZE As Integer = 58) As Integer
Dim iFontSize As Long, LF As LOGFONT, FontBox As SIZE, RectSize As SIZE, ISTEP As Integer, iStepInitial

On Error GoTo ErrHandle
'-----------------------------------------------
If fnt Is Nothing Then Err.Raise 1000, , "Wrong Font to determinate the size"
LF = Get_LogFONT(fnt, DC): iFontSize = fnt.SIZE

RectSize.CY = R.Bottom - R.Top: RectSize.CX = R.Right - R.Left
If RectSize.CY <= 0 Or RectSize.CX <= 0 Then Err.Raise 1000, , "Wrong Rectangle to fit it"
'------------------------------------------------
    iStepInitial = SizeCompare(GetSizeOfText(sText, LF, DC), RectSize)
    ISTEP = iStepInitial
    Do
        iFontSize = iFontSize + ISTEP
        LF.lfHeight = -MulDiv((iFontSize), (GetDeviceCaps(DC, LOGPIXELSY)), 72)
        ISTEP = SizeCompare(GetSizeOfText(sText, LF, DC), RectSize)
     Loop Until (ISTEP * iStepInitial <= 0 Or iFontSize <= minSize Or iFontSize >= MAXSIZE)
'------------------------------------------------
ExitHere:
    FontSizeToRect = iFontSize
    Exit Function
ErrHandle:
    ErrPrint "FonSizetToRect", Err.Number, Err.Description
    Err.Clear
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------
' SET & DELETE TEMP FONT
'--------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetSizeOfText(sText As String, LF As LOGFONT, DC As Variant) As SIZE
Dim sz As SIZE

On Error GoTo ErrHandle
'-----------------------------------
#If VBA7 Then
    Dim h_Fnt As LongPtr
    Dim h_FntOld As LongPtr
#Else
    Dim h_Fnt As Long
    Dim h_FntOld As Long
#End If
'---------------------------------------
        h_Fnt = CreateFontIndirect(LF)
        h_FntOld = SelectObject(DC, h_Fnt)
        
        GetTextExtentPoint32 DC, sText, Len(sText), sz
        
        SelectObject DC, h_FntOld
        DeleteObject h_Fnt
'----------------------------------------
ExitHere:
        GetSizeOfText = sz '!!!!!!!!!!
        Exit Function
'-------------
ErrHandle:
        ErrPrint "GetSizeOfText", Err.Number, Err.Description
        Err.Clear
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------
' Function Return + 1 if TrySize <= FixSize and -1 if TrySize > FixSize
'--------------------------------------------------------------------------------------------------------------------------------------------
Private Function SizeCompare(TrySize As SIZE, FixSize As SIZE) As Integer
Dim iRes As Integer

    If (TrySize.CX > FixSize.CX) Or (TrySize.CY > FixSize.CY) Then
             iRes = -1
    Else
             iRes = 1
    End If

'-------------------
    SizeCompare = iRes '!!!!!!!!!!!!!!!
End Function


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

Private Sub TestHolst()
Dim myBOX As BOX, ShapeRect As RECT
Dim iShape As Integer, Text As String, FillColor As Long
                     
                     

On Error GoTo ErrHandle
'---------------------------------
iShape = 1: Text = "TEST": FillColor = &HFFFF00

    GetClientRect m_Holst.hWnd, ShapeRect
    Call DrawShape(hWnd, hdc, ShapeRect, iShape, Text, FillColor)






'-----------------------------------------------------
ExitHere:
    Exit Sub
'------------
ErrHandle:
    ErrPrint "TestBox", Err.Number, Err.Description
    Err.Clear
End Sub

Private Sub DrawShape(iHwnd As LongPtr, ihDc As LongPtr, ShapeRect As RECT, Optional iShapeType As Integer = 1, Optional sText As String, Optional FillColor As Long = -1)

#If Win64 Then
    Dim hFont As LongPtr, hRgn As LongPtr, hWnd As LongPtr, hdc As LongPtr
    Dim hFillBrush As LongPtr, hFrameBrush As LongPtr, lNumCoords As Long
    
    hWnd = CLngPtr(iHwnd)
    hdc = CLngPtr(ihDc)
#Else
    Dim hFont As Long, hRgn As Long, hWnd As Long, hdc As Long
    Dim hFillBrush As Long, hFrameBrush As Long, lNumCoords As Long
    
    hWnd = CLng(iHwnd)
    hdc = CLng(ihDc)
#End If

Dim poly(1 To 3) As POINTAPI, LogFont_Text As LOGFONT
Dim IWidth As Long, iHeight As Long, iTop As Long, iLeft As Long
Dim LogBrush_Frame As LOGBRUSH, LogBrush_Fill As LOGBRUSH
 
Const FontHeight As Long = 14
Const FontWidth As Long = 9
Const lFrameWidth As Long = 1
Const lFrameHeight As Long = 1

iTop = ShapeRect.Top: iLeft = ShapeRect.Left: IWidth = ShapeRect.Right - ShapeRect.Left
iHeight = ShapeRect.Bottom - ShapeRect.Top

LogFont_Text.lfHeight = FontHeight
LogFont_Text.lfWidth = FontWidth
    

LogBrush_Frame.lbColor = vbBlack
hFrameBrush = CreateBrushIndirect(LogBrush_Frame)
    
If FillColor < 0 Then
    LogBrush_Fill.lbColor = vbWhite
    FillColor = vbWhite
Else
    LogBrush_Fill.lbColor = FillColor
End If

hFillBrush = CreateBrushIndirect(LogBrush_Fill)
    
DoEvents
    
With ShapeRect
   Select Case iShapeType
          Case 1: 'Triangle
                lNumCoords = 3
                poly(1).x = IWidth / 2
                poly(1).y = iTop
                poly(2).x = IWidth
                poly(2).y = iHeight
                poly(3).x = 0
                poly(3).y = iHeight
                Polygon hdc, poly(1), lNumCoords
                'create the triangle region.
                hRgn = CreatePolygonRgn(poly(1), lNumCoords, 1)
                'fill the region.
                FillRgn hdc, hRgn, hFillBrush
                'draw the region frame.
                FrameRgn hdc, hRgn, hFrameBrush, _
                lFrameWidth, lFrameHeight
            Case 2: ' Ellipse
                hRgn = CreateEllipticRgn _
                (.Left, .Top, .Right, .Bottom)
                'fill the region.
                FillRgn hdc, hRgn, hFillBrush
                'draw the region frame.
                FrameRgn hdc, hRgn, hFrameBrush, _
                lFrameWidth, lFrameHeight
            Case 3: ' Rectangle
                hRgn = CreateRectRgn(.Left, .Top, .Right, .Bottom)
                'fill the region.
                FillRgn hdc, hRgn, hFillBrush
                'draw the region frame.
                FrameRgn hdc, hRgn, hFrameBrush, _
                lFrameWidth, lFrameHeight
        End Select
    End With
    
    'create Font and select it in the Shapes DCs.
    hFont = CreateFontIndirect(LogFont_Text)
    Call SelectObject(hdc, hFont)
    GetRgnBox hRgn, ShapeRect
    
    'add shape text.
    DrawText hdc, sText, Len(sText), ShapeRect, _
    DT_CENTER + DT_VCENTER + DT_SINGLELINE
    
    'release resources.
    DeleteObject hRgn
    DeleteObject hFillBrush
    DeleteObject hFrameBrush
    DeleteObject hFont
End Sub

