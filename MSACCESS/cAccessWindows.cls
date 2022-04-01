VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAccessWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************************************************************
'*******************************************************************************************************************************
' Incapsulation of Access Windows manipulatot
' The Event Catcher was developed by  (с) Dev Ashish   http://www.mvps.org/access/api/api0035.htm
' and was modify to manage of general function
' Because of the operator  AddressOF, this instruction is placed into external module #_GUI:
'              Public AWin As cAccessWindows
'              '==============================================================================================================
'              ' THIS FUNCTION IS REDIRECT FOR EMBEDDED AddressOf operator TO cAccessWindows class
'              '==============================================================================================================
'              Public Function WndProc(ByVal hWND As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'                     WndProc = AWin.ClientWndProc(hWND, Msg, wParam, lParam)
'              End Function  ' See Hook Function
'*******************************************************************************************************************************
'*******************************************************************************************************************************
Option Compare Database
Option Explicit

Public Event AWinREPAINT(ByVal hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long, ByRef Cancel As Boolean)
'****************************************************************
Private Type RECT            ' Rectangle Structure
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Type POINTAPI       ' Point Structure
   x As Long
   y As Long
End Type
'******************************************************************************************************************
#If Win64 Then
    Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal wCmd As Long) As Long
    Private Declare PtrSafe Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hWnd As LongPtr, _
                ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
    Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
    Private Declare PtrSafe Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
    Private Declare PtrSafe Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
    Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, lParam As Any) As Long
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
                                                                ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, _
                                         ByVal y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal MSG As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, _
                                                                                              ByVal wNewWord As Long) As Long
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
                                                                                              
    Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As LongPtr, ByVal lpstring As String, ByVal cch As Long) As Long
    '*****************************************************************
    Private Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, _
            ByVal bRevert As Long) As Long

    Private Declare PtrSafe Function EnableMenuItem Lib "user32" (ByVal hMenu As _
            Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long

    Private Declare PtrSafe Function GetMenuItemInfo Lib "user32" Alias _
            "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal Un As Long, ByVal b As _
            Long, lpMenuItemInfo As MENUITEMINFO) As Long
            Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
                     ByVal nHeight As Long, ByVal bRepaint As Long) As Long

    Private Declare PtrSafe Function MapWindowPoints Lib "user32" (ByVal hWndFrom As Long, ByVal hwndTo As Long, lppt As Any, _
                    ByVal cPoints As Long) As Long
    '*****************************************************************
    Private m_WNDMDIClient As LongPtr                         ' Access Interface MDI Client
    Private m_PrevWndProc  As LongPtr                         ' Pointer for an independent processing of messages in the current active window
    Private m_AhWnd As LongPtr                                ' Active Window Pointer (Form)
    Private m_AhDC As LongPtr                                 ' Pointer for the graphic context of the active window (form)

                    
#Else
    Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
    Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As Long, _
                ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
    Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
    Private Declare Function GetDesktopWindow Lib "user32" () As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
                                                                ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, _
                                         ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                   ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
                                                                                              ByVal wNewWord As Long) As Long
    Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, _
                                                                           ByVal lpstring As String, _
                                                                           ByVal cch As Long) As Long
    '*****************************************************************
    Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, _
        ByVal bRevert As Long) As Long

    Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As _
        Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long

    Private Declare Function GetMenuItemInfo Lib "user32" Alias _
        "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal Un As Long, ByVal b As _
        Long, lpMenuItemInfo As MENUITEMINFO) As Long
    Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, _
                     ByVal nHeight As Long, ByVal bRepaint As Long) As Long

    Private Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, _
                    ByVal cPoints As Long) As Long

    '*****************************************************************
    Private m_WNDMDIClient As Long                         ' Access Interface MDI Client
    Private m_PrevWndProc  As Long                         ' Pointer for an independent processing of messages in the current active window
    Private m_AhWnd As Long                                ' Active Window Pointer (Form)
    Private m_AhDC As Long                                 ' Pointer for the graphic context of the active window (form)

#End If
'****************************************************************
Private Const GW_HWNDFIRST = 0          ' Secondary window, displays the first upper level window in the list
Private Const GW_HWNDLAST = 1           ' Calls up the display unit of the next upper level in the list
Private Const GW_HWNDNEXT = 2           ' Calls up the window from the window below this window
Private Const GW_HWNDPREV = 3           ' Calls up the window from the window from the outside window
Private Const GW_CHILD = 5              ' For the revival on the window, adjusted by reference to the current
'****************************************************************
Private Const GWL_STYLE = (-16)                   ' Getting a window style
Private Const GWL_HINSTANCE = (-6)
Private Const GWL_WNDPROC  As Long = (-4)         ' Controlling the timing
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_MDICHILD = &H40

'****************************************************************
Private Const WM_PAINT = &HF
Private Const WM_CLOSE = &H10
Private Const WM_CHILDACTIVATE = &H22
Private Const WM_CLEAR = &H303
Private Const WM_GETTEXT = &HD
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203

Private Const WS_OVERLAPPED = &H0&
Private Const WS_POPUP = &H80000000
Private Const WS_CHILD = &H40000000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_DISABLED = &H8000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_CAPTION = &HC00000
Private Const WS_BORDER = &H800000
Private Const WS_DLGFRAME = &H400000
Private Const WS_VSCROLL = &H200000
Private Const WS_HSCROLL = &H100000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_GROUP = &H20000
Private Const WS_TABSTOP = &H10000

Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000

Private Const WS_TILED = WS_OVERLAPPED
Private Const WS_ICONIC = WS_MINIMIZE
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or _
                                     WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Private Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW

Private Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Private Const WS_CHILDWINDOW = (WS_CHILD)

Private Const WS_EX_DLGMODALFRAME = &H1&
Private Const WS_EX_NOPARENTNOTIFY = &H4&
Private Const WS_EX_TOPMOST = &H8&
Private Const WS_EX_ACCEPTFILES = &H10&
Private Const WS_EX_TRANSPARENT = &H20&
'***********************************************************************************
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Const SM_CXVSCROLL = 2
Private Const SM_CYHSCROLL = 3
Private Const SM_CXBORDER = 5
Private Const SM_CYBORDER = 6
'**********************************************************************************
' Флаги функции SetWindowPos
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Private Const SW_RESTORE = 9

Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

' SetWindowPos() hWndInsertAfter values
Private Const HWND_TOP = 0
Private Const HWND_BOTTOM = 1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const Title = ""

'******************************************
Private Const ACC_MAIN_CLASS = "OMain"                 ' Renaming the Access Window Class
Private Const ACC_FORM_CLIENT_CLASS = "OFormSub"       ' Name of the class for the client window of the Access form
Private Const ACC_CBX_LISTBOX_CLASS = "OGrid"          ' Name of the window class for the Access switch in combo
Private Const ACC_CBX_LISTBOX_PARENT_CLASS = "ODCombo" ' Name of the class of the list combo Assu
Private Const ACC_CBX_EDIT_CLASS = "OKttbx"            ' Naming the Edit controls in Access
Private Const ACC_FORM_CLIENT_CHILD_CLASS = "OFEDT"    ' Working window inside the form
Private Const ACC_TAB_CONTROLL_CLASS = "OTabControl"   ' Closed window switch
Private Const WC_MDICLIENT = "MDIClient"

Private Const HWND_DESKTOP As Long = 0                 ' Desktop Pointer
Private Const LOGPIXELSX As Long = 88                  ' Logically pixels/inch by X
Private Const LOGPIXELSY As Long = 90                  ' Logically pixels/inch  by Y
Private Const TwipsPerInch = 1440&                     ' Twips per inch
Private Const BITSPIXEL = 12                           ' Bits per Pixel

'****************************************************************
' Horizontal and vertical resolution
Private m_ScreenXdpi As Long
Private m_ScreenYdpi As Long
'*******************************************
Private lRet As Long                                  ' Returning
'********************************************
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wid As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Const MF_GRAYED = &H1&
Const MF_BYCOMMAND = &H0&
Const SC_CLOSE = &HF060&


'=====================================================================================================================================
' Get HWND for Active Window
'=====================================================================================================================================
Public Function ActiveWinHwnd() As Variant
   ActiveWinHwnd = GetActiveWindow()   '!!!!!!!!!!!!!!!
End Function

'=====================================================================================================================================
' Get Window Caption
'=====================================================================================================================================
Public Function GetWinCaption(vHwnd As Variant) As String
Dim lenTxt As Long, retText As String, sRes As String

On Error Resume Next
'----------------------------------
If IsZero(vHwnd) Then Exit Function
If vHwnd = 0 Then Exit Function
'---------------
#If Win64 Then
     Dim iHwnd As LongPtr
         iHwnd = CLngPtr(vHwnd)
#Else
     Dim iHwnd As Long
         iHwnd = CLng(vHwnd)
#End If
'-----------------------------------
     lenTxt = GetWindowTextLength(iHwnd) + 1
     GetWindowText iHwnd, retText, lenTxt
     sRes = Trim(retText)
'-----------------------------------
     GetWinCaption = sRes '!!!!!!!!!!
End Function

'=====================================================================================================================================
' Check If Form is Modal
'=====================================================================================================================================
Public Function IsModal(f As Form) As Boolean

Dim lWinstyle As Long

  lWinstyle = GetWindowLong(f.hWnd, GWL_EXSTYLE)
  If (lWinstyle And WS_EX_MDICHILD) Then
    IsModal = False
  Else
    IsModal = True
  End If
End Function

'=====================================================================================================================================
' Функция центрирует диалоговое окно
'=====================================================================================================================================
Public Function CenterForm(parForm As Form, Optional bAccessWindow As Boolean = False) As Boolean
    Dim PlainHWND As Long                           ' Указатель на окно, относительно которого центрируем
    Dim varAccess As RECT, varForm As RECT
    Dim varX As Long, varY As Long
    Dim bRes As Boolean
    
    On Error GoTo ErrHandle
    '-------------------------------------------------------------------------------------------
    If bAccessWindow Then
       PlainHWND = Application.hWndAccessApp
    Else
       PlainHWND = GetDesktopWindow()
    End If
    '-------------------------------------------------------------------------------------------
    Call GetClientRect(PlainHWND, varAccess)                ' Получаем координаты клиентской части, относительно которой центрируемся
    Call GetWindowRect(parForm.hWnd, varForm)               ' Получаем координаты окна формы
    varX = CLng((varAccess.Left + varAccess.Right) / 2) - CLng((varForm.Right - varForm.Left) / 2) ' Вычмсляем новый левый угол формы
    varY = CLng((varAccess.Top + varAccess.Bottom) / 2) - CLng((varForm.Bottom - varForm.Top) / 2) ' Вычисляем новый верх формы
    varY = varY - 45 ' Подправляем смещения
    varY = varY - 20
    '---------------------------------------------------------------------------------------------
    Call ShowWindow(parForm.hWnd, SW_RESTORE)               ' Позиционируемся
    Call SetWindowPos(parForm.hWnd, 0, varX, varY, (varForm.Right - varForm.Left), (varForm.Bottom - varForm.Top), SWP_NOZORDER Or SWP_SHOWWINDOW Or SWP_NOSIZE) ' Set new form coordinates
     
    bRes = True
'-----------------------------------------------------------------------------------------------------------
ExitHere:
    CenterForm = bRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'-----------------------------------------------------------------------------------
ErrHandle:
    ErrPrint "CenterForm", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'================================================================================================================================
' Свойство: Доступность кнопки закрытия главного окна
'================================================================================================================================
Public Property Get EnabledMainWinClose() As Boolean
    Dim hWnd As Long
    Dim hMenu As Long
    Dim result As Long
    Dim MI As MENUITEMINFO
    
    MI.cbSize = Len(MI)
    MI.dwTypeData = String(80, 0)
    MI.cch = Len(MI.dwTypeData)
    MI.fMask = MF_GRAYED
    MI.wid = SC_CLOSE
    hWnd = Application.hWndAccessApp
    hMenu = GetSystemMenu(hWnd, 0)
    result = GetMenuItemInfo(hMenu, MI.wid, 0, MI)
    EnabledMainWinClose = (MI.fState And MF_GRAYED) = 0
End Property
Public Property Let EnabledMainWinClose(boolClose As Boolean)
    Dim hWnd As Long
    Dim wFlags As Long
    Dim hMenu As Long
    Dim result As Long
    
    hWnd = Application.hWndAccessApp
    hMenu = GetSystemMenu(hWnd, 0)
    If Not boolClose Then
        wFlags = MF_BYCOMMAND Or MF_GRAYED
    Else
        wFlags = MF_BYCOMMAND And Not MF_GRAYED
    End If
    result = EnableMenuItem(hMenu, SC_CLOSE, wFlags)
End Property

'#####################################################################################################################################
'#####################################################################################################################################
'#####################################################################################################################################
'#####################################################################################################################################
#If Win64 Then

    '================================================================================================================================
    ' Active Window Pointer
    '================================================================================================================================
    Public Property Get ahWnd() As LongPtr
        ahWnd = m_AhWnd
    End Property
    Public Property Let ahWnd(ByVal iHwnd As LongPtr)
        m_AhWnd = iHwnd
    End Property

    '================================================================================================================================
    ' Descriptor of the Active Window
    '================================================================================================================================
    Public Property Get aHdc() As LongPtr
        aHdc = m_AhDC
    End Property
    Public Property Let aHdc(ByVal ihDc As LongPtr)
        m_AhDC = ihDc
    End Property
    '================================================================================================================================
    ' hWND Access MDI Client Access
    '================================================================================================================================
    Public Property Get hWNDMDIClient() As LongPtr
        If m_WNDMDIClient = 0 Then m_WNDMDIClient = fMDIClienthWnd()
        hWNDMDIClient = m_WNDMDIClient
    End Property

    Public Function ClientWndProc(ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, _
                                                                                                  ByVal lParam As Long) As LongPtr
    Dim bCancel As Boolean

        bCancel = False
        ' ------------------------------------------------------------------------------------
        Select Case MSG
            Case WM_PAINT:
            
                RaiseEvent AWinREPAINT(hWnd, MSG, wParam, lParam, bCancel)
                If bCancel Then Exit Function  ' Клиент отменил событие, не возвращаем управление
        End Select
        ' -------------------------------------------------------------------------------
        ' Return Management to Windows
            ClientWndProc = CallWindowProc(m_PrevWndProc, hWnd, MSG, wParam, lParam)
    End Function
    
    '===============================================================================================================================
    ' The function toggles off Windows messages for a given window (only one window at a time)
    '===============================================================================================================================
    Public Function Unhook() As Boolean
    
    On Error GoTo Unhook_ERR
    '--------------------------------
        If m_AhWnd <= 0 Then Exit Function
    '------------------------------------------------------------------------------------
    ' Restore the front end of the processor
        Call SetWindowLongPtr(m_AhWnd, GWL_WNDPROC, m_PrevWndProc)
    '----------------------------------------------------------------
        Unhook = True      '!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        m_PrevWndProc = 0
        Exit Function
    '-----------------------------------
Unhook_ERR:
   ErrPrint2 "Unhook", Err.Number, Err.Description, "cAccessWindows"
   Err.Clear
End Function

'#####################################################################################################################################
#Else
    '================================================================================================================================
    ' Active Window Pointer
    '================================================================================================================================
    Public Property Get ahWnd() As Long
        ahWnd = m_AhWnd
    End Property
    Public Property Let ahWnd(ByVal iHwnd As Long)
        m_AhWnd = iHwnd
    End Property

    '================================================================================================================================
    ' Descriptor of the Active Window
    '================================================================================================================================
    Public Property Get aHdc() As Long
        aHdc = m_AhDC
    End Property
    Public Property Let aHdc(ByVal ihDc As Long)
        m_AhDC = ihDc
    End Property

    '================================================================================================================================
    ' hWND Access MDI Client Access
    '================================================================================================================================
    Public Property Get hWNDMDIClient() As Long
        If m_WNDMDIClient = 0 Then m_WNDMDIClient = fMDIClienthWnd()
        hWNDMDIClient = m_WNDMDIClient
    End Property

    Public Function ClientWndProc(ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, _
                                                                                                  ByVal lParam As Long) As Long
    Dim bCancel As Boolean

        bCancel = False
        ' ------------------------------------------------------------------------------------
        Select Case MSG
            Case WM_PAINT:
            
                RaiseEvent AWinREPAINT(hWnd, MSG, wParam, lParam, bCancel)
                If bCancel Then Exit Function  ' Клиент отменил событие, не возвращаем управление
        End Select
        ' -------------------------------------------------------------------------------
        ' Return Management to Windows
            ClientWndProc = CallWindowProc(m_PrevWndProc, hWnd, MSG, wParam, lParam)
    End Function
    '===============================================================================================================================
    ' The function toggles off Windows messages for a given window (only one window at a time)
    '===============================================================================================================================
    Public Function Unhook() As Boolean
    
    On Error GoTo Unhook_ERR
    '--------------------------------
        If m_AhWnd <= 0 Then Exit Function
    '------------------------------------------------------------------------------------
    ' Restore the front end of the processor
        Call SetWindowLong(m_AhWnd, GWL_WNDPROC, m_PrevWndProc)
    '----------------------------------------------------------------
        Unhook = True      '!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        m_PrevWndProc = 0
        Exit Function
    '-----------------------------------
Unhook_ERR:
   ErrPrint2 "Unhook", Err.Number, Err.Description, "cAccessWindows"
   Err.Clear
End Function
#End If
'##############################################################################################################################
'##############################################################################################################################
'##############################################################################################################################

'===============================================================================================================================
' МЕТОД: Скрыть / показать главное коно
'===============================================================================================================================
Public Function ShowHideMainWindow(bShow As Boolean) As Boolean
Dim HWND_TOPMOST, SWP_SHOWWINDOW, SWP_NOACTIVATE
On Error GoTo ShowHideMainWindow_ERR
If bShow Then
SetWindowPos Application.hWndAccessApp, HWND_TOPMOST, 0, 0, 1200, 1200, SWP_SHOWWINDOW
' 0, 0, 1200, 1200, SWP_SHOWWINDOW ' С этими значениями можно поиграть
' зависит от размера экрана
Else
    SetWindowPos Application.hWndAccessApp, HWND_TOPMOST, -30, -30, 0, 0, SWP_SHOWWINDOW
End If
'----------------------------------------------------------------------------------
    ShowHideMainWindow = True '!!!!!!!!!!!!!!!!!
    Exit Function
'----------------------------------------------------------
ShowHideMainWindow_ERR:
    MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, Err.Source
End Function
'===============================================================================================================================
' МЕТОД: Установка в качестве активного окна клиентской (рабочей части) окна формы
'===============================================================================================================================
Public Function SetFormClientAsAWin(aForm As Access.Form) As Boolean

On Error GoTo SetFormClientAsAWin_ERR
        m_AhWnd = GetHwndClientArea(aForm)
        If m_AhWnd <> 0 Then
           m_AhDC = GetDCClientArea(aForm)
        Else
           Exit Function
        End If
'---------------------------------------------------------
        SetFormClientAsAWin = True     '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        Exit Function
'---------------------------------
SetFormClientAsAWin_ERR:
        MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, Err.Source
End Function
'===============================================================================================================================
' МЕТОД: Функция безопасно меняет свойства базы данных
'===============================================================================================================================
Public Function ChangeProperty(strPropName As String, varPropType As Variant, varPropValue As Variant) As Integer
Dim dbs As DAO.Database, prp As DAO.Property

Const conPropNotFoundError = 3270
Set dbs = CurrentDb

On Error GoTo Change_Err

    dbs.Properties(strPropName) = varPropValue
    ChangeProperty = True
Change_Bye:
    Exit Function
Change_Err:
    If Err = conPropNotFoundError Then ' Свойство не найдено.
        Set prp = dbs.CreateProperty(strPropName, varPropType, varPropValue)
        dbs.Properties.Append prp
        Resume Next
    Else
        ' Неизвестная ошибка.
        ChangeProperty = False
        Resume Change_Bye
    End If
End Function
'===============================================================================================================================
' МЕТОД: Очистка класса от установок (перевод их в состояние Default)  (TBD: TEXT)
'===============================================================================================================================
Public Sub ClearObjects()
    If m_PrevWndProc <> 0 Then Unhook
    m_PrevWndProc = 0                  ' Дескриптор функции - обработчика
    m_AhWnd = 0                        ' Дексриптор активного окна
    m_AhDC = 0                         ' Дескриптор графического контекста активного окна
End Sub

'===============================================================================================================================
' МЕТОД: Функция включает перехват сообщений Windows для активного окна (только одно окно в одно время) (TBD: TEXT)
'===============================================================================================================================
Public Function Hook() As Boolean
Dim AdrrWar As Variant
On Error GoTo ErrHandle
    If Not IsWindow(m_AhWnd) Then Exit Function

#If Win64 Then
    If m_AhWnd <= 0 Then m_PrevWndProc = SetWindowLongPtr(m_AhWnd, GWL_WNDPROC, AddressOf WndProc)   ' TBD: SetWindowLongPtr - not for 32
#Else
    If m_AhWnd <= 0 Then m_PrevWndProc = SetWindowLong(m_AhWnd, GWL_WNDPROC, AddressOf WndProc)   ' TBD: SetWindowLongPtr - not for 32
#End If
'-----------------------------------
ExitHere:
   Hook = True '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
   Exit Function
'-----------------
ErrHandle:
  ErrPrint "Hook", Err.Number, Err.Description
  Err.Clear
End Function
'==========================================================================================================================
' // Функция возвращает для формы Access дескриптор графического контекста клиентской области
' Тип Variant. Для win64 - LongPtr, для win32 - Long
'==========================================================================================================================
Public Function GetDCClientArea(aForm As Access.Form) As Variant
On Error GoTo ErrHandle

If aForm Is Nothing Then Exit Function
'------------------------------------------------------------
#If Win64 Then
    Dim mhWND As LongPtr
    mhWND = CLngPtr(GetHwndClientArea(aForm))
#Else
    Dim mhWND As Long
    mhWND = CLng(GetHwndClientArea(aForm))
#End If
'------------------------------------------------------------
ExitHere:
   GetDCClientArea = GetDC(mhWND)  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
   Exit Function
'-------------------------------------
ErrHandle:
    MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "GetDCClientArea ERROR"
    Err.Clear
End Function
'==========================================================================================================================
' //Функция возвращает для формы Access дескриптор окна клиентской области
' Для win64 - LongPtr, для win32 - Long
'==========================================================================================================================
Public Function GetHwndClientArea(aForm As Access.Form) As Variant
Dim fHWND As Variant                        ' дескриптор окна формы (главного окна)
Dim ResHWND As Variant                      ' возвращаемый результат

On Error GoTo ErrHandle
'---------------------------------------------
If aForm Is Nothing Then Exit Function
fHWND = aForm.hWnd                 ' Получаем верхнеуровневый дескриптор
'------------------------------------------------------------
#If Win64 Then
    ResHWND = GetRightHwnd4AccessForm(CLngPtr(fHWND))
#Else
    ResHWND = GetRightHwnd4AccessForm(CLng(fHWND))
    
#End If
'------------------------------------------------------------
ExitHere:
   GetHwndClientArea = ResHWND '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
   Exit Function
'-------------------------------------
ErrHandle:
    MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "GetHwndClientArea ERROR"
    Err.Clear
End Function
'===========================================================================================================================
' Определяем максимальную высоту объекта
'===========================================================================================================================
Public Function GetClientHeight(lhWnd) As Long
Dim mmRECT As RECT                      ' Определяем прямоугольник

GetClientRect lhWnd, mmRECT
'------------------------------------
GetClientHeight = mmRECT.Bottom - mmRECT.Top '!!!!!!!!!!!!!!!!!!!!!!
End Function
'===================================================================================================================
' Персечет пикселов в Twips по оси X
'===================================================================================================================
Public Function TwipsPerPixelX() As Single
  Dim lngDC As Long
  lngDC = GetDC(HWND_DESKTOP)
  TwipsPerPixelX = 1440& / GetDeviceCaps(lngDC, LOGPIXELSX)
  ReleaseDC HWND_DESKTOP, lngDC
End Function
'===================================================================================================================
' Персечет пикселов в Twips по оси Y
'===================================================================================================================
Public Function TwipsPerPixelY() As Single
'--------------------------------------------------
'Returns the height of a pixel, in twips.
'--------------------------------------------------
  Dim lngDC As Long
  lngDC = GetDC(HWND_DESKTOP)
  TwipsPerPixelY = 1440& / GetDeviceCaps(lngDC, LOGPIXELSY)
  ReleaseDC HWND_DESKTOP, lngDC
End Function


'-----------------------------------------------------------------------------------------------------------------------------
' Функция отыскивает необходимый hWnd для оперирование. Обычный hWnd - только область данных, контекст формы по нему не
' получить. Поэтому перечисляем все классы связанных окон для поиска контекста
' Возвращает тип Variant (если win64 - LongPtr, win32 - Long)
'-----------------------------------------------------------------------------------------------------------------------------
Private Function GetRightHwnd4AccessForm(frmHWND As Variant) As Variant
On Error GoTo ErrHandle
#If Win64 Then
    Dim whWnd As LongPtr, H As LongPtr
    whWnd = GetWindow(CLngPtr(frmHWND), GW_CHILD)           'Получаем первое окно
#Else
    Dim whWnd As Long, H As Long
    whWnd = GetWindow(CLng(frmHWND), GW_CHILD)              'Получаем первое окно
#End If
'----------------------------------------------------
'Начинаем проверку всех окон формы
Do While whWnd
    Select Case GetClassWin(whWnd)                                              'Определяем клас окна
            Case ACC_FORM_CLIENT_CLASS, ACC_TAB_CONTROLL_CLASS                  'Приемлемые классы оконо
                    'Рекурсивный вызов
                    If GetClassWin(GetRightHwnd4AccessForm(whWnd)) = ACC_FORM_CLIENT_CHILD_CLASS Then        'Нашли класс
                        Exit Do
                    End If
            Case "OFEDT"                                                        'Окно для работы
                       Exit Do
    End Select
    '----------------------------------------
    'Получаем заголовок следующего окна
     whWnd = GetWindow(whWnd, GW_HWNDNEXT)
Loop
'----------------------------------------------------------------
ExitHere:
      GetRightHwnd4AccessForm = whWnd '!!!!!!!!!!!!!!!!!!!!
      Exit Function
'---------------------------------------------------------
ErrHandle:
     MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description
     Err.Clear
End Function
'================================================================================================================================
' Функция позволяет позиционировать вторичную форму FRM2 относительно имеющегося контрола
' PARAMETERS:  frmName  - наименование формы
'              ctl      - контрол, относительно которого позиционируем форму
'              Position - позиция вторичной формы
'                       0 = Underneath
'                       1 = On Top
'                       2 = Right side
'                       3 = Left side
'                       4 = Bottom Right Hand Corner
'================================================================================================================================
Public Function PositionFormRelativeToControl(FrmName As String, ctl As Access.Control, _
                                                                                       Optional Position As Long = 0) As Boolean
Dim m_hWndSection As Long   ' Window handle to our Form's Detail Section
Dim frm As Access.Form      ' Form we will position under the control
Dim hWndMDI As Long         ' Access MDI document window
Dim MDIborderX As Long      ' MDI borders
Dim MDIborderY As Long

Dim rc As RECT              ' For positioning window
Dim rcWin As RECT
Dim pt As POINTAPI
Dim lOffsetX As Long, lOffsetY As Long


Dim m_ScreenWidth As Long   ' Screen dimensions
Dim m_ScreenHeight As Long
Dim lStyle As Long          ' Window Style var

' Since we are turning off screen redraw ignore all errors
On Error Resume Next

' Turn off redraw
' Leave this alone util you are done debugging
'Application.Echo False

' Open the Form
DoCmd.OpenForm FrmName
' Does form exist?
Set frm = Forms.Item(FrmName)

If Not frm Is Nothing Then

' Get the Window handle for the form Section containing this control
m_hWndSection = fFindSectionhWnd(ctl)
' Calculate the LEFT offset for this control from the edge of the Section
' First calc our screen resolution
GetScreenDPI
' Now get our screen dimensions
m_ScreenWidth = GetSystemMetrics(SM_CXSCREEN)
m_ScreenHeight = GetSystemMetrics(SM_CYSCREEN)

Select Case Position

Case 0
lOffsetX = ctl.Left / (TwipsPerInch / m_ScreenXdpi)
lOffsetY = (ctl.Top + ctl.height) / (TwipsPerInch / m_ScreenYdpi)

Case 1
lOffsetX = ctl.Left / (TwipsPerInch / m_ScreenXdpi)
lOffsetY = (-frm.WindowHeight + ctl.Top) / (TwipsPerInch / m_ScreenYdpi)

Case 2
lOffsetX = (ctl.Left + ctl.Width) / (TwipsPerInch / m_ScreenXdpi)
lOffsetY = (ctl.Top) / (TwipsPerInch / m_ScreenYdpi)

Case 3
' Nov-2004 Logic error
' Must use Form's width NOT COntrol's width
'lOffsetX = (ctl.Left - ctl.Width) / (TWIPSPERINCH / m_ScreenXdpi)
lOffsetX = (ctl.Left - frm.WindowWidth) / (TwipsPerInch / m_ScreenXdpi)
lOffsetY = (ctl.Top) / (TwipsPerInch / m_ScreenYdpi)

Case 4
lOffsetX = (ctl.Left + ctl.Width) / (TwipsPerInch / m_ScreenXdpi)
lOffsetY = (ctl.Top + ctl.height) / (TwipsPerInch / m_ScreenYdpi)

Case Else
' Default to Underneath
lOffsetX = ctl.Left / (TwipsPerInch / m_ScreenXdpi)
lOffsetY = (ctl.Top + ctl.height) / (TwipsPerInch / m_ScreenYdpi)

End Select

' Get window rectangle of the Section
lRet = GetWindowRect(m_hWndSection, rc)
' Add in offsets for the calling control
pt.x = lOffsetX + rc.Left&
pt.y = lOffsetY + rc.Top

' Bounds checking to ensure form will be completely visible on screen

lRet = GetWindowRect(frm.hWnd, rcWin)

With rcWin
If m_ScreenWidth - pt.x < .Right - .Left Then
pt.x = m_ScreenWidth - (.Right - .Left)
ElseIf pt.x < 2 Then 'm_S
pt.x = 2
End If

If m_ScreenHeight - pt.y < .Bottom - .Top Then
pt.y = m_ScreenHeight - (.Bottom - .Top)
ElseIf pt.y < 2 Then
pt.y = 2
End If

End With


' If the Form's POPUP property is True then skip the
' ScreenToClient stuff
If Not frm.PopUp = True Then

' find MDIClient window
hWndMDI = FindWindowEx(Application.hWndAccessApp, 0&, "MDIClient", Title)
' COnvert to Client coordinates of our MDI window
lRet = ScreenToClient(hWndMDI, pt)
' We have to subtract our MDI window's Border
lRet = GetWindowRect(hWndMDI, rcWin)
lRet = GetClientRect(hWndMDI, rc)
MDIborderX = ((rcWin.Right - rcWin.Left) - (rc.Right - rc.Left))
MDIborderY = ((rcWin.Bottom - rcWin.Top) - (rc.Bottom - rc.Top))

' See if ScrollBar is visible in the MDI window
lStyle = GetWindowLong(hWndMDI, GWL_STYLE)

If lStyle And WS_HSCROLL Then
MDIborderY = MDIborderY - GetSystemMetrics(SM_CYHSCROLL)
End If

If lStyle And WS_VSCROLL Then
MDIborderX = MDIborderX - GetSystemMetrics(SM_CXVSCROLL)
End If
' Remainder is our Border thickness
MDIborderX = MDIborderX / 2
MDIborderY = MDIborderY / 2
Else
' POPUP = True
' We need to subtract the Border thickness of the main Access
'Aplication Window
MDIborderX = GetSystemMetrics(SM_CXBORDER)
MDIborderY = GetSystemMetrics(SM_CYBORDER)
End If

' Position our Form underneath of the calling control
Call SetWindowPos(frm.hWnd, 0&, pt.x - MDIborderX, pt.y - MDIborderY, 0, 0, SWP_NOSIZE)

End If

' Turn on redraw
'Application.Echo True
' Cleanup
Set frm = Nothing

' Return Success
PositionFormRelativeToControl = True

End Function
'=================================================================================================================================================
' Функция получает статус (координаты ) окна формы. Возвращает строку вида: Left;Top;Right;Bottom
'=================================================================================================================================================
Public Function GetWindowState(currForm As Form, Optional sDelim As String = ";") As String
Dim WinProp As RECT                               ' Определяющий прямоугольник
Dim I As Long                                     ' Результат
Dim lpMyCoord As POINTAPI                         ' Точки пересчета
Dim lngOffsetX As Long, lngOffsetY As Long        ' Смещение
Dim iLeft As Long, iTop As Long, iRight As Long, iBottom As Long
Dim sRes As String                                ' Возвращаемый результат
On Error GoTo Err_GetWindowState

' Получаем абсолютные координаты текущей формы
I = GetWindowRect(currForm.hWnd, WinProp)
    If I <> 0 Then
        lpMyCoord.x = WinProp.Left
        lpMyCoord.y = WinProp.Top
            ' Преобразуем в относительные
            I = MapWindowPoints(GetDesktopWindow(), GetParent(currForm.hWnd), lpMyCoord, 1)
            lngOffsetY = WinProp.Top - lpMyCoord.y
            lngOffsetX = WinProp.Left - lpMyCoord.x
                    iLeft = WinProp.Left - lngOffsetX
                    iTop = WinProp.Top - lngOffsetY
                    iRight = WinProp.Right - lngOffsetX
                    iBottom = WinProp.Bottom - lngOffsetY
    End If
'--------------------------------------------------------------------------------
Exit_GetWindowState:
   If (iLeft + iTop + iRight + iBottom) <> 0 Then
      sRes = iLeft & sDelim & iTop & sDelim & iRight & sDelim & iBottom
   Else
      sRes = ""
   End If
   GetWindowState = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Exit Function
'---------------------------------------
Err_GetWindowState:
    MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "GetWindowState ERROR"
Resume Exit_GetWindowState
End Function
'=================================================================================================================================================
' Функция устанавливает положение окна формы, принимая на вход строку вида Left;Top;Right;Bottom
'=================================================================================================================================================
Public Sub SetWindowState(currForm As Form, sState As String, Optional sDelim As String = ";")
Dim I As Long           ' Флаг результата
Dim iLeft As Long, iTop As Long, iRight As Long, iBottom As Long
Dim MyArr() As String

On Error GoTo Err_SetWindowState
If sState = "" Then Exit Sub

'-------------------------------------------------------------------------------
' Считываем переменные координаты
    MyArr = Split(sState, sDelim)
    iLeft = MyArr(0): iTop = MyArr(1): iRight = MyArr(2): iBottom = MyArr(3)
'-------------------------------------------------------------------------------
' Перемещаем окно
If (iRight - iLeft) > 0 And (iBottom - iTop) > 0 Then
    I = MoveWindow&(currForm.hWnd, iLeft, iTop, iRight - iLeft, iBottom - iTop, 0)
End If
'--------------------------------------------------------------------------------
Exit_SetWindowState:
    Exit Sub
Err_SetWindowState:
    MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "SetWindowState ERROR"
    Resume Exit_SetWindowState
End Sub
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
On Error GoTo ErrHandler
    'Set UW.AWin = Me                   ' Инициирование редиректора на модуль для использования AddressOf в классе
    Call ClearObjects
'---------------------------------------------------------
ExitHere:
    Exit Sub
'-----------------------------------
ErrHandler:
    MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, Err.Source
End Sub
Private Sub Class_Terminate()
    If m_PrevWndProc <> 0 Then Unhook
End Sub
'-----------------------------------------------------------------------------------------------------------------------
' Получение разрешения экрана
'-----------------------------------------------------------------------------------------------------------------------
Private Sub GetScreenDPI()
Dim lngDC As Long


lngDC = GetDC(0)                                  ' Grab any DC

m_ScreenXdpi = GetDeviceCaps(lngDC, LOGPIXELSX)   'Horizontal
m_ScreenYdpi = GetDeviceCaps(lngDC, LOGPIXELSY)   'Vertical

lngDC = ReleaseDC(0, lngDC)
End Sub

Private Function fFindSectionhWnd(ctl As Access.Control) As Long
On Error GoTo Err_fFindSectionhWnd
' Get ListBox's hWnd
Dim hWnd_LSB As Long
Dim hWnd_Temp As Long

' Window RECT vars
Dim rc As RECT
Dim pt As POINTAPI

' Loop Counters
Dim SectionCounter As Long
Dim ctr As Long

' Which Section contains the Control?
Select Case ctl.Section
Case acDetail   '0
SectionCounter = 2
Case acHeader   '1
SectionCounter = 1
Case acFooter   '2
SectionCounter = 3
Case Else
'  ****   NEED ERROR HANDLING! ****
End Select

' Setup SectionCounter
' Form Header, Detail and then Footer
ctr = 1

' Nov -2004 - Modification by Onno Willems
' Let's get first Child Window of the FORM
If TypeOf ctl.Parent Is Access.Page Then
' If the control is on a page on a tab, we can't get a window handle
If TypeOf ctl.Parent.Parent Is Access.TabControl Then
' As expected, the page is on a tab
If TypeOf ctl.Parent.Parent.Parent Is Access.Form Then
' And the parent of the tab is the form we wanted
hWnd_LSB = GetWindow(ctl.Parent.Parent.Parent.hWnd, GW_CHILD)
End If
End If
Else
' Normal control directly on form
hWnd_LSB = GetWindow(ctl.Parent.hWnd, GW_CHILD)
End If

' Let's walk through every sibling window of the Form
Do
If GetClassWin(hWnd_LSB) = "OFormSub" Then
' First OFormSub is the Form's Header. We want the next next one
' which is the detail section
If ctr = SectionCounter Then
fFindSectionhWnd = hWnd_LSB
Exit Function
End If

' Increment our Section Counter
ctr = ctr + 1

End If

' Let's get the NEXT SIBLING Window
hWnd_LSB = GetWindow(hWnd_LSB, GW_HWNDNEXT)

' Let's Start the process from the Top again
' Really just an error check
Loop While hWnd_LSB <> 0

' SORRY - NO ListBox hWnd is available
fFindSectionhWnd = 0

Exit_fFindSectionhWnd:
Exit Function

Err_fFindSectionhWnd:
MsgBox Err.Description
Resume Exit_fFindSectionhWnd

End Function

'---------------------------------------------------------------------------------------------------------------------------
' Получение дескриптора клиентской части MDI MS Access
'---------------------------------------------------------------------------------------------------------------------------
Private Function fMDIClienthWnd() As Long
    fMDIClienthWnd = FindWindowEx(hWndAccessApp, 0, WC_MDICLIENT, vbNullString)
End Function
'-------------------------------------------------------------------------------------------------------------------------------
' Определяем класс окна,например,"OFormSub"
'-------------------------------------------------------------------------------------------------------------------------------
Private Function GetClassWin(iHwnd As Variant) As String
Dim s As String, cnt As Long

#If Win64 Then
    Dim hWnd As LongPtr
    hWnd = CLngPtr(iHwnd)
#Else
    Dim hWnd As Long
    hWnd = CLng(iHwnd)
#End If


    s = String$(255, " ") 'Заполняем строку
    cnt = GetClassName(hWnd, s, 255) 'Записываем класс в s
    If cnt > 0 Then
        GetClassWin = Left$(s, cnt) 'Получаем класс
    End If
End Function
'-------------------------------------------------------------------------------------------------------------------------------
' //Получение графического контекста для формы
' Возвращает тип Variant (win64 - LongPtr, win32 - Long)
'-------------------------------------------------------------------------------------------------------------------------------
Private Function GetRightDC4AccessForm(frmHWND As Variant) As Variant

#If Win64 Then
    Dim fHWND As LongPtr
    fHWND = CLngPtr(GetRightHwnd4AccessForm(CLngPtr(frmHWND)))
#Else
    Dim fHWND As Long
    fHWND = CLng(GetRightHwnd4AccessForm(CLng(frmHWND)))
#End If
'-------------------------------------------
'возвращаем результат
    GetRightDC4AccessForm = GetDC(fHWND)    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'======================================================================================================================================================
' Позиционирование формы
'======================================================================================================================================================
Public Sub sGetFormDimensions(frm As Form, ByRef lngLeft As Long, ByRef lngTop As Long, ByRef lngWidth As Long, ByRef lngHeight As Long)
Dim lngRet As Long, tRect As RECT
    lngRet = GetWindowRect(frm.hWnd, tRect)
    With tRect
        lngLeft = .Left
        lngTop = .Top
        lngWidth = .Right - .Left
        lngHeight = .Bottom - .Top
        ConvertPixelsToTwips lngLeft, lngTop
        ConvertPixelsToTwips lngWidth, lngHeight
    End With
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Конвертирование пикселей в твипы
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ConvertPixelsToTwips(x As Long, y As Long)
'From the KB
Dim hdc As Long, hWnd As Long, RetVal As Long
Dim XPixelsPerInch, YPixelsPerInch
Const LOGPIXELSX = 88
Const LOGPIXELSY = 90
    hdc = GetDC(0)
    XPixelsPerInch = GetDeviceCaps(hdc, LOGPIXELSX)
    YPixelsPerInch = GetDeviceCaps(hdc, LOGPIXELSY)
    RetVal = ReleaseDC(0, hdc)
    x = (x / XPixelsPerInch) * TwipsPerInch
    y = (y / YPixelsPerInch) * TwipsPerInch
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
                                                                                                  Optional sModName As String = "cAccessWindows") As String
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

