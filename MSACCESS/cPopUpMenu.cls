VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPopUpMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*************************************************************************************************************
'*************************************************************************************************************
' КОНТЕКСТНОЕ МЕНЮ. Класс создает контекстное меню и размещает его в нужной позиции.
' Для работы требуется hwnd формы и положение курсора
' ////////////////////////////////////////////////////////////////////////////////////////////////////////////
' ||                         ПРИМЕР ИСПОЛЬЗОВАНИЯ
' ||  Dim mnu As cPopUpMenu: Set mnu = New cPopUpMenu
' ||    With mnu
' ||
' ||        .AddItem 1, "Default", True
' ||        .AddItem 0, "-"
' ||        .AddItem 2, "Новое меню 2"
' ||        .AddItem 3, "Новое меню 3"
' ||
' ||        .AddItem 2, "Child Menus", , , True         ' Disabled menu item must has ID other than 0
' ||        Dim i As Long
' ||        For i = 1 To 10
' ||            Dim submnu As cPopUpMenu: Set submnu = New cPopUpMenu
' ||            With submnu
' ||                .Caption = "Child Menu " & i
' ||                .AddItem i * 10 + 1, "Child Menu " & i
' ||            End With
' ||            .AddItem i * 10, submnu
' ||        Next
' ||        .AddItem 0, "-"
' ||        .AddItem 3, "Checked", , True
' ||        .AddItem 4, "Grayed", , , , True
' ||        .AddItem 5, "New Column", , , , , True
' ||
' ||        MsgBox .ShowPopup
' ||    End With
' ////////////////////////////////////////////////////////////////////////////////////////////////////////////
'
'
'************************************************************************************************************
'*************************************************************************************************************
'*************************************************************************************************************

Option Compare Database
Option Explicit

Public Caption As String        ' Когда это работает как подменю, то название должно быть добавлено для родительского


' ###########################################################################################################
'      ДЕКЛАРАЦИИ ДЕКЛАРАЦИИ ДЕКЛАРАЦИИ ДЕКЛАРАЦИИ ДЕКЛАРАЦИИ ДЕКЛАРАЦИИ ДЕКЛАРАЦИИ ДЕКЛАРАЦИИ ДЕКЛАРАЦИИ
' ###########################################################################################################
#If Win64 Then

Private Declare PtrSafe Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As LongPtr, ByVal uItem As Long, ByVal fByPos As Long) As Long
    Private Declare PtrSafe Function CreatePopupMenu Lib "user32" () As LongPtr
    Private Declare PtrSafe Function DestroyMenu Lib "user32" (ByVal hMenu As LongPtr) As Long
    Private Declare PtrSafe Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As LongPtr, ByVal wFlags As Long, ByVal wIDNewItem As LongPtr, ByVal lpNewItem As String) As Long
    Private Declare PtrSafe Function AppendMenuW Lib "user32" (ByVal hMenu As LongPtr, ByVal wFlags As Long, ByVal wIDNewItem As LongPtr, ByVal lpNewItem As LongPtr) As Long
    Private Declare PtrSafe Function EnableMenuItem Lib "user32" (ByVal hMenu As LongPtr, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long

    Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTL) As Long
    Private Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function TrackPopupMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal wFlags As Long, _
          ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hWnd As LongPtr, ByVal lprc As Long) As Long
    
    Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal wCmd As Long) As Long
    Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As LongPtr, lpdwProcessId As Long) As Long
    Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long

    Private m_hMenu     As LongPtr

#Else

    Private Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
    Private Declare Function CreatePopupMenu Lib "user32" () As Long
    Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
    Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, lpNewItem As String) As Long
    Private Declare Function AppendMenuW Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
    Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long

    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTL) As Long
    Private Declare Function GetDesktopWindow Lib "user32" () As Long
    Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, _
          ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal lprc As Long) As Long
    Private Declare Function GetFocus Lib "user32" () As Long
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
    Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
    Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    
    Private m_hMenu     As Long
#End If
'####################################################################
Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDFIRST = 0
Private Const MF_BYCOMMAND = &H0&
Private Const MF_BYPOSITION = &H400
Private Const MF_CHECKED = &H8&
Private Const MF_DISABLED = &H2&
Private Const MF_GRAYED = &H1&
Private Const MF_MENUBARBREAK = &H20&
Private Const MF_MENUBREAK = &H40&
Private Const MF_POPUP = &H10&
Private Const MF_SEPARATOR = &H800&
Private Const MF_STRING = &H0&
Private Const MIIM_ID = &H2
Private Const MIIM_SUBMENU = &H4
Private Const MIIM_TYPE = &H10
Private Const TPM_LEFTALIGN = &H0&
Private Const TPM_RETURNCMD = &H100&
Private Const TPM_RIGHTBUTTON = &H2
Private Const TPM_LEFTBUTTON = &H0&



Private Const APIFALSE As Long = 0
Private Const APITRUE As Long = 1


Private Type POINTL
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type




'###########################################################################################################
'###########################################################################################################
'###########################################################################################################
'###########################################################################################################


#If Win64 Then

    Friend Property Get hMenu() As LongPtr
        hMenu = m_hMenu
    End Property

    Public Sub AddItem(ByVal nID As Long, varItem As Variant, Optional bDefault As Boolean = False, _
        Optional bChecked As Boolean = False, Optional bDisabled As Boolean = False, Optional bGrayed As Boolean = False, _
                                                                              Optional bNewColumn As Boolean = False)
    
        If TypeName(varItem) = "String" Then
            If varItem = "-" Then
                AppendMenu m_hMenu, MF_STRING Or MF_SEPARATOR, nID, ByVal vbNullString
            Else
                If IsUnicode(CStr(varItem)) Then
                    AppendMenuW m_hMenu, MF_STRING Or IIf(bNewColumn, MF_MENUBARBREAK, 0) Or IIf(bChecked, MF_CHECKED, 0), nID, StrPtr(varItem)
                Else
                    AppendMenu m_hMenu, MF_STRING Or IIf(bNewColumn, MF_MENUBARBREAK, 0) Or IIf(bChecked, MF_CHECKED, 0), nID, ByVal varItem
                End If
            End If
        ElseIf TypeOf varItem Is cPopUpMenu Then
            Dim cSubMenu As cPopUpMenu
            Set cSubMenu = varItem
            AppendMenu m_hMenu, MF_STRING Or MF_POPUP Or IIf(bNewColumn, MF_MENUBARBREAK, 0), cSubMenu.hMenu, ByVal cSubMenu.Caption
        End If
    
        If bDefault Then SetMenuDefaultItem m_hMenu, nID, APIFALSE
        If bGrayed Then EnableMenuItem m_hMenu, nID, MF_BYCOMMAND Or MF_GRAYED
        If bDisabled Then EnableMenuItem m_hMenu, nID, MF_BYCOMMAND Or MF_DISABLED
    End Sub


    Public Function ShowPopup(Optional ByVal hWnd As LongPtr, Optional ByVal PopX As Long = -1, _
                     Optional ByVal PopY As Long = -1, Optional ByVal hWndOfBeneathControl As Long = -1) As Long
        Dim H As LongPtr, x As Long, y As Long
        Dim idCurrent As Long
        Dim hDesktop As LongPtr, hChild As LongPtr
    
        If hWnd = 0 Then
            ' Find the top window of current process
            hDesktop = GetDesktopWindow()
            hChild = GetWindow(hDesktop, GW_CHILD)
            idCurrent = GetCurrentProcessId()
        
            Do While hChild
                Dim idChild As Long
                GetWindowThreadProcessId hChild, idChild
                If idChild = idCurrent Then Exit Do
                hChild = GetWindow(hChild, GW_HWNDNEXT)
            Loop
        
            If hChild = 0 Then Err.Raise -1, "cMenu.TrackPopup", "Cannot find top window of current process!"
            H = hChild
        Else
            H = hWnd
        End If
    
        If hWndOfBeneathControl <> -1 Then
            Dim rt As RECT
            GetWindowRect hWndOfBeneathControl, rt
            x = rt.Left
            y = rt.Bottom
        Else
            Dim pt As POINTL
            GetCursorPos pt
            If PopX = -1 Then x = pt.x Else: x = PopX
            If PopY = -1 Then y = pt.y Else: y = PopY
        End If
    
        ShowPopup = TrackPopupMenu(m_hMenu, TPM_RETURNCMD + TPM_LEFTALIGN + TPM_LEFTBUTTON, x, y, 0, H, 0)
    End Function

#Else

    Friend Property Get hMenu() As Long
        hMenu = m_hMenu
    End Property

    Public Sub AddItem(ByVal nID As Long, varItem As Variant, Optional bDefault As Boolean = False, _
        Optional bChecked As Boolean = False, Optional bDisabled As Boolean = False, Optional bGrayed As Boolean = False, _
                                                                              Optional bNewColumn As Boolean = False)
    
        If TypeName(varItem) = "String" Then
            If varItem = "-" Then
                AppendMenu m_hMenu, MF_STRING Or MF_SEPARATOR, nID, ByVal vbNullString
            Else
                If IsUnicode(CStr(varItem)) Then
                    AppendMenuW m_hMenu, MF_STRING Or IIf(bNewColumn, MF_MENUBARBREAK, 0) Or IIf(bChecked, MF_CHECKED, 0), nID, StrPtr(varItem)
                Else
                    AppendMenu m_hMenu, MF_STRING Or IIf(bNewColumn, MF_MENUBARBREAK, 0) Or IIf(bChecked, MF_CHECKED, 0), nID, ByVal varItem
                End If
            End If
        ElseIf TypeOf varItem Is cPopUpMenu Then
            Dim cSubMenu As cPopUpMenu
            Set cSubMenu = varItem
            AppendMenu m_hMenu, MF_STRING Or MF_POPUP Or IIf(bNewColumn, MF_MENUBARBREAK, 0), cSubMenu.hMenu, ByVal cSubMenu.Caption
        End If
    
        If bDefault Then SetMenuDefaultItem m_hMenu, nID, APIFALSE
        If bGrayed Then EnableMenuItem m_hMenu, nID, MF_BYCOMMAND Or MF_GRAYED
        If bDisabled Then EnableMenuItem m_hMenu, nID, MF_BYCOMMAND Or MF_DISABLED
    End Sub

    Public Function ShowPopup(Optional ByVal hWnd As Long = -1, Optional ByVal PopX As Long = -1, _
                 Optional ByVal PopY As Long = -1, Optional ByVal hWndOfBeneathControl As Long = -1) As Long
        Dim H As Long, x As Long, y As Long
        Dim idCurrent As Long
        Dim hDesktop As Long, hChild As Long
    
        If hWnd = -1 Or hWnd = 0 Then
            ' Find the top window of current process
            hDesktop = GetDesktopWindow()
            hChild = GetWindow(hDesktop, GW_CHILD)
            idCurrent = GetCurrentProcessId()
        
            Do While hChild
                Dim idChild As Long
                GetWindowThreadProcessId hChild, idChild
                If idChild = idCurrent Then Exit Do
                hChild = GetWindow(hChild, GW_HWNDNEXT)
            Loop
        
            If hChild = 0 Then Err.Raise -1, "cMenu.TrackPopup", "Cannot find top window of current process!"
            H = hChild
        Else
            H = hWnd
        End If
    
        If hWndOfBeneathControl <> -1 Then
            Dim rt As RECT
            GetWindowRect hWndOfBeneathControl, rt
            x = rt.Left
            y = rt.Bottom
        Else
            Dim pt As POINTL
            GetCursorPos pt
            If PopX = -1 Then x = pt.x Else: x = PopX
            If PopY = -1 Then y = pt.y Else: y = PopY
        End If
    
        ShowPopup = TrackPopupMenu(m_hMenu, TPM_RETURNCMD + TPM_LEFTALIGN + TPM_LEFTBUTTON, x, y, 0, H, 0)
    End Function

#End If
'************************************************************************************************************
'------#########---##########---#########----##-----------##-------###---------##########---#########--------
'------##-----##---##------##------###--------##---------##-------##-##------------##-------##---------------
'------##-----##---##------##------###---------##-------##-------##---##-----------##-------##---------------
'------#########---##########------###----------##-----##-------##-----##----------##-------#####------------
'------##----------##---##---------###-----------##---##-------###########---------##-------##---------------
'------##----------##----##--------###------------##-##-------##---------##--------##-------##---------------
'------##----------##-----###---#########----------###-------##-----------##-------##-------#########--------
'************************************************************************************************************
Private Sub Class_Initialize()
    m_hMenu = CreatePopupMenu()
End Sub

Private Sub Class_Terminate()
    DestroyMenu m_hMenu
End Sub





