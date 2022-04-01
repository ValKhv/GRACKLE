VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*************************************************************************************************************************************************
'*************************************************************************************************************************************************
' SPLITTER MANAGEMENT
' @V.Khvatov      21/03/2018  (c) DigitalXpert
'*************************************************************************************************************************************************
'*************************************************************************************************************************************************

Option Compare Database
Option Explicit
'************************************************************************************************************************************************
Private Type RECT            ' Структура - прямоугольник
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Type POINTAPI       ' Структура - точка
   x As Long
   y As Long
End Type
'************************************************************************************************************************************************
Private Const DSTINVERT = &H550009
Private Const PATINVERT = &H5A0049
Private Const SRCCOPY = &HCC0020
Private Const PATCOPY = &HF00021
Private Const BLACKNESS = &H42
'************************************************************************************************************************************************
' ВЫЗОВЫ API
#If Win64 Or VB7 Then                          ' НОВАЯ СРЕДА
'------------------------------------------------------------------------------------------------------------
    Private Declare PtrSafe Function ClientToScreen Lib "user32" (ByVal hWnd As LongPtr, lpPoint As POINTAPI) As Long
    Private Declare PtrSafe Function PatBlt Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
                                                                                           ByVal nHeight As Long, ByVal dwRop As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
    Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    '-----------------------------------------------------------------
    Private lngHdcDesktop As LongPtr                                 ' Дескриптор десктопа
    
'------------------------------------------------------------------------------------------------------------
#Else                                            ' СТАРАЯ СРЕДА
    Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As Any) As Long  ' lpPoint As POINTAPI) As Long
    Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
                                                                                           ByVal nHeight As Long, ByVal dwRop As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
    '-----------------------------------------------------------------
    Private lngHdcDesktop As Long                                   ' Дескриптор десктопа

'------------------------------------------------------------------------------------------------------------
#End If
'*********************************************************************************************************************************************
Private blnDraw As Boolean
Private lngInitialPosition As Long


Private rctClient As RECT                                       ' Отрисовываемый прямоугольник
Private Redess As Boolean                                       ' Отбрасывание тени при перетаскивании
Private m_VERT As Boolean                                       ' Вертикальный сплиттер (по-умолчанию горизонтальный)


Public TopControls As New Collection
Public BottomControls As New Collection
Public LeftControls As New Collection
Public RightControls As New Collection
Public WithEvents VSplit As Label
Attribute VSplit.VB_VarHelpID = -1

Public HeaderHeight As Long
Public FooterHeight As Long
Public ctl As Control
Public WithEvents HSplit As Label
Attribute HSplit.VB_VarHelpID = -1
Public WithRedess As Boolean

Public BorderTop As Long
Public BorderBottom As Long
Public BorderLeft As Long
Public BorderRight As Long

Public height As Long
Public Width As Long


'*********************************************************************************************************************************************
'*********************************************************************************************************************************************
'=============================================================================================================================================
' Свойство: ВЕРТИКАЛЬНЫЙ ИЛИ ГОРИЗОНТАЛЬНЫЙ БАР
'=============================================================================================================================================
Public Property Get isVertical() As Boolean
   isVertical = m_VERT
End Property
Public Property Let isVertical(ByVal bVert As Boolean)
   m_VERT = bVert
End Property

'================================================================================================================================================
' ИНИЦИАЛИЗАЦИЯ СПЛИТТЕРА
'================================================================================================================================================
Public Sub Init()
    If m_VERT Then                            ' ВЕРТИКАЛЬНЫЙ СПЛИТТЕР
    '------------------------------------------------------------------
        With VSplit
            .OnMouseDown = "[event procedure]"
            .OnMouseMove = "[event procedure]"
            .OnMouseUp = "[event procedure]"
            '.HyperlinkAddress = Chr(32) ' Pour afficher le curseur main
        End With
        VSplit.Width = Width
    Else                                      ' ГОРИЗОНТАЛЬНЫЙ СПЛИТТЕР
    '------------------------------------------------------------------
        With HSplit
            .OnMouseDown = "[event procedure]"
            .OnMouseMove = "[event procedure]"
            .OnMouseUp = "[event procedure]"
            '.HyperlinkAddress = Chr(32) ' Pour afficher le curseur main
        End With
        HSplit.height = height
    End If
    '------------------------------------------------------------------
End Sub
'================================================================================================================================================
' Изменение размеров родителя
'================================================================================================================================================
Public Sub OnParentRedim()
    Dim lngHeight As Long
    Dim lngWidth As Long
    Dim lngWidthSplit As Long
    '-----------------------------------------------------------------------------------
    If m_VERT Then                                       ' ВЕРТИКАЛЬНЫЙ СПЛИТТЕР
        lngHeight = VSplit.Parent.InsideHeight
        lngWidth = VSplit.Parent.InsideWidth
    
        If lngHeight <= VSplit.Top Or lngWidth <= (VSplit.Left + VSplit.Width) Then Exit Sub
    
        VSplit.height = lngHeight - VSplit.Top
    
        For Each ctl In LeftControls
            ctl.Width = VSplit.Left - ctl.Left
        Next ctl
    
        For Each ctl In RightControls
            ctl.Left = VSplit.Left + VSplit.Width
            ctl.Width = lngWidth - (VSplit.Left + VSplit.Width)
        Next ctl
    '-----------------------------------------------------------------------------------
    Else                                                 ' ГОРИЗОНТАЛЬНЫЙ СПЛИТТЕР
        lngHeight = HSplit.Parent.InsideHeight
        lngWidth = HSplit.Parent.InsideWidth
    
        If lngHeight - (HeaderHeight + FooterHeight) <= (HSplit.Top + HSplit.height) Or lngWidth <= HSplit.Left Then Exit Sub
    
        HSplit.Width = lngWidth - HSplit.Left
    
        For Each ctl In TopControls
            ctl.height = HSplit.Top - ctl.Top
        Next ctl
    
        For Each ctl In BottomControls
            ctl.Top = HSplit.Top + HSplit.height
            ctl.height = lngHeight - (HSplit.Top + HSplit.height) - (HeaderHeight + FooterHeight)
        Next ctl
    End If
    '-----------------------------------------------------------------------------------
End Sub



'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'-----------------------------------------------------------------------------------------------------------------------------------------------
' ИНИЦИАЛИЗАЦИЯ КЛАССА
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    m_VERT = False                                   ' По умолчанию сплиттер горизонтальный
    lngHdcDesktop = GetDC(0)
    BorderTop = 300
    BorderBottom = 300
    BorderLeft = 300
    BorderRight = 300
    
    WithRedess = False
    height = 100
    Width = 100
End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
' РЕЛИЗ КЛАССА
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
    If lngHdcDesktop Then
        Call ReleaseDC(0, lngHdcDesktop)
    End If
    Set TopControls = Nothing
    Set BottomControls = Nothing
    Set HSplit = Nothing
    
    Set LeftControls = Nothing
    Set RightControls = Nothing
    Set VSplit = Nothing

    Set ctl = Nothing
End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------
' Отработка нажатия на горизонтальном сплиттере
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub HSplit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    lngInitialPosition = HSplit.Top
    blnDraw = True
    Screen.MousePointer = 7
    Redess = (Shift = 1) Or WithRedess
    If Not Redess Then
        Call GetWindowRect(GetParent(HSplit.Parent.hWnd), rctClient)
        Call DrawSplit(HSplit.Left + 15, HSplit.Top + HeaderHeight, HSplit.Width - 15, HSplit.height, DSTINVERT)
    End If
End Sub
Private Sub VSplit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    lngInitialPosition = VSplit.Left
    blnDraw = True
    Screen.MousePointer = 7
    Redess = (Shift = 1) Or WithRedess
    If Not Redess Then
        Call GetWindowRect(GetParent(VSplit.Parent.hWnd), rctClient)
        Call DrawSplit(VSplit.Left, HeaderHeight + 15, VSplit.Width, (VSplit.height - VSplit.Top - (HeaderHeight + FooterHeight + 15)), DSTINVERT)
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
' Отработка перемещения на горизонтальном сплиттере
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub HSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngNewPos As Long
    If blnDraw Then
        lngNewPos = lngInitialPosition + y
        If lngNewPos > BorderTop And lngNewPos < (HSplit.Parent.InsideHeight - (HeaderHeight + FooterHeight) - BorderBottom) Then
            'HSplit.Top = lngNewPos
            If Redess Then
                HSplit.Top = lngNewPos
                MoveControls lngNewPos
            Else
                Call DrawSplit(HSplit.Left + 15, HSplit.Top + HeaderHeight, HSplit.Width - 15, HSplit.height, DSTINVERT)
                HSplit.Top = lngNewPos
                Call DrawSplit(HSplit.Left + 15, HSplit.Top + HeaderHeight, HSplit.Width - 15, HSplit.height, DSTINVERT)
            End If
        End If
        lngInitialPosition = HSplit.Top
    End If
End Sub
Private Sub VSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngNewPos As Long
    If blnDraw Then
        lngNewPos = lngInitialPosition + x
        If lngNewPos > BorderLeft And lngNewPos < (VSplit.Parent.InsideWidth - BorderRight) Then
            If Redess Then
                VSplit.Left = lngNewPos
                MoveControls lngNewPos
            Else
                Call DrawSplit(VSplit.Left, HeaderHeight + 15, VSplit.Width, (VSplit.height - VSplit.Top - (HeaderHeight + FooterHeight + 15)), DSTINVERT)
                VSplit.Left = lngNewPos
                Call DrawSplit(VSplit.Left, HeaderHeight + 15, VSplit.Width, (VSplit.height - VSplit.Top - (HeaderHeight + FooterHeight + 15)), DSTINVERT)
            End If
        End If
        lngInitialPosition = VSplit.Left
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
' Отпустили мышь на горизонтальном сплиттере
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub HSplit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    blnDraw = False
    Screen.MousePointer = 0
    HSplit.Top = lngInitialPosition
    If Not Redess Then
        Call DrawSplit(HSplit.Left + 15, HSplit.Top + HeaderHeight, HSplit.Width - 15, HSplit.height, DSTINVERT)
        MoveControls lngInitialPosition
    End If
End Sub
Private Sub VSplit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    blnDraw = False
    Screen.MousePointer = 0
    VSplit.Left = lngInitialPosition
    If Not Redess Then
        Call DrawSplit(lngInitialPosition, HeaderHeight + 15, VSplit.Width, (VSplit.height - VSplit.Top - (HeaderHeight + FooterHeight + 15)), DSTINVERT)
        MoveControls VSplit.Left
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
' ПЕРЕМЕЩЕНИЕ КОНТРОЛА
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub MoveControls(lngPos As Long)

On Error Resume Next
    If m_VERT Then                                  ' ВЕРТИКАЛЬНЫЙ СПЛИТТЕР
        For Each ctl In LeftControls
            ctl.Width = lngPos - ctl.Left
        Next ctl
        For Each ctl In RightControls
            ctl.Left = lngPos + VSplit.Width
            ctl.Width = VSplit.Parent.InsideWidth - (lngPos + VSplit.Width)
        Next ctl
    Else                                            ' ГОРИЗОНТАЛЬНЫЙ СПЛИТТЕР
        For Each ctl In TopControls
            ctl.height = lngPos - ctl.Top
        Next ctl
        For Each ctl In BottomControls
            ctl.height = HSplit.Parent.InsideHeight - (lngPos + HSplit.height) - (HeaderHeight + FooterHeight)
            ctl.Top = lngPos + HSplit.height
        Next ctl
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------
' ОТРИСОВКА КОНТРОЛА
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub DrawSplit(lngX As Long, lngY As Long, lngWidth As Long, lngHeight As Long, Mask As Long)
    
    Dim pt As POINTAPI
    Dim RCT As RECT

    
    pt.x = Int(lngX / 15)
    pt.y = Int(lngY / 15)
    lngHeight = Int(lngHeight / 15)
    lngWidth = Int(lngWidth / 15)
    '-------------------------------------------------------------------------------
    If m_VERT Then                                  ' ВЕРТИКАЛЬНЫЙ СПЛИТТЕР
        Call ClientToScreen(VSplit.Parent.hWnd, pt)
    
        ' Prise en compte des coordonnйes d'Access "ClientArea"
        If pt.x < rctClient.Left Then pt.x = RCT.Left
        If pt.y < rctClient.Top Then pt.y = RCT.Top
        If pt.x > rctClient.Right Then pt.x = rctClient.Right - 2
        If pt.y + lngHeight > rctClient.Bottom Then lngHeight = rctClient.Bottom - pt.y - 2
        
        Call PatBlt(lngHdcDesktop, pt.x, pt.y, lngWidth, lngHeight, Mask)
    '-------------------------------------------------------------------------------
    Else                                            ' ГОРИЗОНТАЛЬНЫЙ СПЛИТТЕР
    '-------------------------------------------------------------------------------
        Call ClientToScreen(HSplit.Parent.hWnd, pt)
    
        ' Prise en compte des coordonnйes d'Access "ClientArea"
        If pt.x < rctClient.Left Then
            lngWidth = lngWidth + (pt.x - rctClient.Left)
            pt.x = rctClient.Left
        End If
        If pt.y < rctClient.Top Then pt.y = rctClient.Top
        If (pt.x + lngWidth) > (rctClient.Left + rctClient.Right) Then lngWidth = rctClient.Right - pt.x - 2
        If (pt.y + lngHeight) > rctClient.Bottom Then pt.y = rctClient.Bottom - 2
        
        Call PatBlt(lngHdcDesktop, pt.x, pt.y, lngWidth, lngHeight, Mask)
    End If
    '-------------------------------------------------------------------------------
End Sub


