VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************************************
'**********************************************************************************************************************************************
'**********************************************************************************************************************************************
' DESCRIPTION: The class calculates the sizes of various objects in a Windows environment. Used mainly in the GUI module
' VERSION    : v 0.3 29.05.2018
' COPYRIGHT  : V.Khvatov
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
Option Explicit

Public Enum LengthUnit
    GC_PIXELS = 0
    GC_TWIPS = 1
End Enum


Private Const LF_FACESIZE = 32
Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const FF_DONTCARE = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const DEFAULT_CHARSET = 1
Private Const DT_CALCRECT = &H400

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
   lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90
Private Const TwipsPerInch = 1440

Private Const DEFAULT_FONT_NAME As String = "MS Sans Serif"
Private Const DEFAULT_FONT_SIZE As Integer = 8

Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const ACC_FORM_CLIENT_CLASS = "OFormSub"
Private Const ACC_FORM_CLIENT_CHILD_CLASS = "OFEDT"

Const ERROR_INVALID_WINDOW_HANDLE As Long = 1400
Const ERROR_INVALID_WINDOW_HANDLE_DESCR As String = "Invalid window handle."
'******************************************************************************************************************


#If Win64 Then
      Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
      Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
      Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
      Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As LongPtr, _
                                                                        ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
      Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal wCmd As Long) As LongPtr
      
      Private Declare PtrSafe Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As LongPtr
      Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
      Private Declare PtrSafe Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
      Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
      Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal m_hDC As LongPtr, ByVal nIndex As Long) As Long
      Private Declare PtrSafe Function DrawText Lib "user32" Alias "DrawTextA" (ByVal m_hDC As LongPtr, ByVal lpStr As String, _
                                                ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
      Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal m_hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
      Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal m_hDC As LongPtr) As Long
      
    
    Private m_hDC As LongPtr
    Private hFnt As LongPtr
    Private hFntOld As LongPtr
#Else
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function GetDesktopWindow Lib "user32" () As Long
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, _
                                                                        ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
    
    Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
    Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
    Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal m_hDC As Long, ByVal nIndex As Long) As Long
    Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal m_hDC As Long, ByVal lpStr As String, ByVal nCount As Long, _
                                                                     lpRect As RECT, ByVal wFormat As Long) As Long
    Private Declare Function SelectObject Lib "gdi32" (ByVal m_hDC As Long, ByVal hObject As Long) As Long
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal m_hDC As Long) As Long
    
    
    Private m_hDC As Long
    Private hFnt As Long
    Private hFntOld As Long
#End If
'******************************************************************************************************************

Private m_Font As StdFont
'******************************************************************************************************************

Public Property Get FontName() As String
      FontName = m_Font.Name
End Property
Public Property Let FontName(sFontName As String)
      m_Font.Name = sFontName
End Property

Public Property Get FontSize() As Integer
      FontSize = m_Font.SIZE
End Property
Public Property Let FontSize(iFontSize As Integer)
      m_Font.SIZE = iFontSize
End Property

Public Property Get FontBold() As Boolean
      FontBold = m_Font.Bold
End Property
Public Property Let FontBold(iBold As Boolean)
      m_Font.Bold = iBold
End Property

Public Property Get FontItalic() As Boolean
      FontBold = m_Font.Italic
End Property
Public Property Let FontItalic(iItalic As Boolean)
      m_Font.Italic = iItalic
End Property



Public Function TextWidth(ByVal sText As String, Optional sFontName As String, Optional iFontSize As Integer) As Long
Dim R As RECT
    R = GetFontSize(sText, sFontName, iFontSize)
    TextWidth = R.Right
End Function

Public Function TextHeight(ByVal sText As String, Optional sFontName As String, Optional iFontSize As Integer) As Long
Dim R As RECT
    R = GetFontSize(sText, sFontName, iFontSize)
    TextHeight = R.Bottom
End Function

Public Function PixelsToTwipsWidth(ByVal x As Long) As Long

On Error GoTo ErrHandle
'---------------------------------------------
If x <= 0 Then Exit Function
m_hDC = GetDC(0)
  PixelsToTwipsWidth = _
          x / GetDeviceCaps(m_hDC, LOGPIXELSX) * TwipsPerInch
  ReleaseDC 0, m_hDC
'------------------------------------
ExitHere:
     Exit Function
'------------------
ErrHandle:
     ErrPrint "PixelsToTwipsWidth", Err.Number, Err.Description
     Err.Clear
End Function
Public Function PixelsToTwipsHeight(ByVal y As Long) As Long

On Error GoTo ErrHandle
'---------------------------------------------
If y <= 0 Then Exit Function
m_hDC = GetDC(0)
  PixelsToTwipsHeight = _
          y / GetDeviceCaps(m_hDC, LOGPIXELSY) * TwipsPerInch
  ReleaseDC 0, m_hDC
'------------------------------------
ExitHere:
     Exit Function
'------------------
ErrHandle:
     ErrPrint "PixelsToTwipsHeight", Err.Number, Err.Description
     Err.Clear
End Function

Public Function TwipsToPixelsWidth(ByVal x As Long) As Long

On Error GoTo ErrHandle
'---------------------------------------------
If x <= 0 Then Exit Function
  m_hDC = GetDC(0)
  TwipsToPixelsWidth = x / TwipsPerInch * GetDeviceCaps(m_hDC, LOGPIXELSX)
  ReleaseDC 0, m_hDC
'------------------------------------
ExitHere:
     Exit Function
'------------------
ErrHandle:
     ErrPrint "TwipsToPixelsWidth", Err.Number, Err.Description
     Err.Clear
End Function

Public Function TwipsToPixelsHeight(ByVal y As Long) As Long

On Error GoTo ErrHandle
'---------------------------------------------
If y <= 0 Then Exit Function
  m_hDC = GetDC(0)
  TwipsToPixelsHeight = y / TwipsPerInch * GetDeviceCaps(m_hDC, LOGPIXELSY)
  ReleaseDC 0, m_hDC
'------------------------------------
ExitHere:
     Exit Function
'------------------
ErrHandle:
     ErrPrint "TwipsToPixelsHeight", Err.Number, Err.Description
     Err.Clear
End Function

'==============================================================================================================================================
' Get Size of Window: left;top;width;height
'==============================================================================================================================================
Public Function GetWindowCoordinates(hWnd As Variant) As Long()
Dim rectWindow As RECT, WinDim(3) As Long

On Error GoTo ErrHandle
'----------------------------------------
If GetWindowRect(hWnd, rectWindow) = 0 Then
    If Err.LastDllError = ERROR_INVALID_WINDOW_HANDLE Then ' Check LastDLLError and display a dialog box if the error
        Err.Raise ERROR_INVALID_WINDOW_HANDLE, , ERROR_INVALID_WINDOW_HANDLE_DESCR
    End If
Else
    WinDim(0) = rectWindow.Left: WinDim(1) = rectWindow.Top
    WinDim(2) = rectWindow.Right - rectWindow.Left: WinDim(3) = rectWindow.Bottom - rectWindow.Top
End If
'-----------------------------------
ExitHere:
    GetWindowCoordinates = WinDim '!!!!!!!!!!!!!!
    Exit Function
'-------------
ErrHandle:
    ErrPrint "GetWindowCoordinates", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'==============================================================================================================================================
' Get Desktop Handle
'==============================================================================================================================================
Public Function GetDeskTopHandler() As Variant
    GetDeskTopHandler = GetDesktopWindow()
End Function
Public Function GetActiveWindowHandler() As Variant
    GetActiveWindowHandler = GetActiveWindow()
End Function
'==============================================================================================================================================
' Get Inside Form Client For Access Form only
'   An Access form's hWnd is actually bound to the recordselector "window"
'==============================================================================================================================================
Public Function GetInsideAccessFormHandler(frm As Form) As Variant

#If Win64 Then
    Dim hWnd As LongPtr
#Else
    Dim hWnd As Long
#End If
'------------------------------------------------
On Error GoTo ErrHandle
    
    hWnd = GetWindow(frm.hWnd, GW_CHILD)
 
    Do While hWnd    '   iterate through all child windows of the form
        If FindClass(hWnd) = ACC_FORM_CLIENT_CLASS Then   '   if we locate the client area whose class name is "OFormSub"
            
            If FindClass(GetWindow(hWnd, GW_CHILD)) = ACC_FORM_CLIENT_CHILD_CLASS Then
                            GetInsideAccessFormHandler = hWnd   '   if we found a match, then return the handle and we're outta here.
                            Exit Do
            End If
        End If
        
        hWnd = GetWindow(hWnd, GW_HWNDNEXT)           '   get a handle to the next child window
    Loop
'----------------------------------------------
ExitHere:
    Exit Function
'-----------------
ErrHandle:
    ErrPrint "", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function


'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
    Set m_Font = New StdFont
    
    m_Font.Name = DEFAULT_FONT_NAME
    m_Font.SIZE = DEFAULT_FONT_SIZE
End Sub
Private Sub Class_Terminate()
    Set m_Font = Nothing
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
                                                                                                  Optional sModName As String = "cSize") As String
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
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Function GetFontSize(ByVal sText As String, Optional sFontName As String, Optional iFontSize As Integer) As RECT
Dim tLF As LOGFONT, tR As RECT
    
    If sFontName <> "" Then m_Font.Name = sFontName
    If iFontSize > 0 Then m_Font.SIZE = iFontSize
    
    m_hDC = GetDC(0):
    Convert_StdFont_To_LogFont m_Font, m_hDC, tLF
    
    hFnt = CreateFontIndirect(tLF)
    hFntOld = SelectObject(m_hDC, hFnt)
    
    DrawText m_hDC, sText, -1, tR, DT_CALCRECT
    
    SelectObject m_hDC, hFntOld
    DeleteObject hFnt
'----------------------------------------------------
ExitHere:
    GetFontSize = tR '!!!!!!!!!!!!!!!!
    Exit Function
'-----------------
ErrHandle:
    ErrPrint "GetFontSize", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function

Private Sub Convert_StdFont_To_LogFont(fntThis As StdFont, mm_hDC As Variant, tLF As LOGFONT)
    Dim sFont As String
    Dim iChar As Integer
    Dim b() As Byte
    
    ' Convert an OLE StdFont to a LOGFONT structure:
    With tLF
        sFont = fntThis.Name
        b = StrConv(sFont, vbFromUnicode)
        For iChar = 1 To Len(sFont)
            .lfFaceName(iChar - 1) = b(iChar - 1)
        Next iChar
        ' Based on the Win32SDK documentation:
        .lfHeight = -MulDiv((fntThis.SIZE), (GetDeviceCaps(m_hDC, LOGPIXELSY)), 72)
        .lfItalic = fntThis.Italic
        If (fntThis.Bold) Then
            .lfWeight = FW_BOLD
        Else
            .lfWeight = FW_NORMAL
        End If
        .lfUnderline = fntThis.Underline
        .lfStrikeOut = fntThis.Strikethrough
        .lfCharSet = fntThis.CHARSET
    End With
End Sub

 
Private Function FindClass(iHwnd As Variant) As String
Dim strbuffer As String, lngCount As Long
Const MAX_LEN = 255
On Error Resume Next

#If Win64 Then
    Dim hWnd As LongPtr
    hWnd = CLngPtr(iHwnd)
#Else
    Dim hWnd As Long
    hWnd = CLng(iHwnd)
#End If

    strbuffer = String$(MAX_LEN - 1, 0)
    lngCount = GetClassName(hWnd, strbuffer, MAX_LEN)
    If lngCount > 0 Then
        FindClass = Left$(strbuffer, lngCount)
    End If
End Function


