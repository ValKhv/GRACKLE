Attribute VB_Name = "#_HELPER"
'*********************************.ze$$e. **********************************************************************************************************
'              .ed$$$eee..      .$$$$$$$P""              ########  #######       #### ####### ##   ##  ##     #######
'           z$$$$$$$$$$$$$$$$$ee$$$$$$"                  ##        ##    ##     ## ## ##      ##  ##   ##     ##
'        .d$$$$$$$$$$$$$$$$$$$$$$$$$"                    ##        ##   ##     ##  ## ##      ####     ##     ####
'      .$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$e..                ##   #### ######     ####### ##      ## ##    ##     ##
'    .$$****""""***$$$$$$$$$$$$$$$$$$$$$$$$$$$be.        ##     ## ##   ##   ##    ## ##      ##   ##  ##     ##
'                     ""**$$$$$$$$$$$$$$$$$$$$$$$L       ######### ##    ## ##     ## ####### ##    ## #####  #######
'                       z$$$$$$$$$$$$$$$$$$$$$$$$$
'                     .$$$$$$$$P**$$$$$$$$$$$$$$$$              ##     ##  #####       ####
'                    d$$$$$$$"              4$$$$$               ##    ##  ##  ##     ## ##
'                  z$$$$$$$$$                $$$P"                ##   ##  ####      ##  ##
'                 d$$$$$$$$$F                $P"                   ##  ##  ##  ##   #######
'                 $$$$$$$$$$F                                       ## ##  ##  ##  ##    ##   ## ### #### ##   ###### #### ######
'                  *$$$$$$$$"                                        ####  #####  ##     ##   ##  ## ##   ##   ##  ## ##   ##  ##
'                    "***""               _____________                                       ###### ###  ##   ####   #### #####
' STANDARD MODULE WITH FILE FUNCTIONS    |v 2017/03/19 |                                      ##  ## ##   ##   ##     ##   ##  ##
' The module contains callback and sublassing functions and is part of the G-VBA library      ##  ## #### #### ##     #### ##   ##
'****************************************************************************************************************************************************
' Due to the peculiarities of the VBA, this module is sensitive to changes and requires compiling and reloading the entire project before using it,
' otherwise there is a high risk of some functions crashing during their execution
'****************************************************************************************************************************************************
'****************************************************************************************************************************************************
Option Explicit

' SOME WIN API CONSTANTS AND API CALLS
Public Const GWL_WNDPROC As Long = -4

Private Const DT_CENTER = &H1
Private Const WM_PAINT = &HF
Private Const WM_DESTROY = &H2
Private Const HCBT_ACTIVATE = 5

' BUTTON IDs FOR MESSAGEBOX
Private Const IDOK = 1
Private Const IDCANCEL = 2
Private Const IDABORT = 3
Private Const IDRETRY = 4
Private Const IDIGNORE = 5
Private Const IDYES = 6
Private Const IDNO = 7


Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'********************************************************************************************************
#If Win64 Then
       Private Type PAINTSTRUCT
            hdc As LongPtr
            fErase As Long
            rcPaint As RECT
            fRestore As Long
            fIncUpdate As Long
            rgbReserved(0 To 31) As Byte
        End Type
        
        Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal MSG As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
        Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

        Private Declare PtrSafe Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As LongPtr, ByVal lpstring As String, ByVal hData As LongPtr) As Long
        Private Declare PtrSafe Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As LongPtr, ByVal lpstring As String) As LongPtr
        Private Declare PtrSafe Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As LongPtr, ByVal lpstring As String) As LongPtr

        Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
        Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
        Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr

        Private Declare PtrSafe Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As LongPtr, ByVal lpstring As String) As LongPtr
        Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As LongPtr
        
        
        Private Declare PtrSafe Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
        
        Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As LongPtr)
        Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
        Private Declare PtrSafe Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
        
        
        Private Declare PtrSafe Function BeginPaint Lib "user32" (ByVal hWnd As LongPtr, lpPaint As PAINTSTRUCT) As LongPtr
        Private Declare PtrSafe Function EndPaint Lib "user32" (ByVal hWnd As LongPtr, lpPaint As PAINTSTRUCT) As Long
        Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
        Private Declare PtrSafe Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As LongPtr, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long


        Public Function FnPtrToLong(ByVal lngFnPtr As LongPtr) As LongPtr
            FnPtrToLong = lngFnPtr
        End Function

        Public Function GetWinHook() As Variant
            GetWinHook = FnPtrToLong(AddressOf MainWndProc)
        End Function

        Private Function MainWndProc(ByVal hWnd As LongPtr, ByVal message As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Variant

         Dim rt As RECT, hdc As LongPtr, PS As PAINTSTRUCT

                        ' i put this here so that you can see how to process a message.
                        ' this is the WM_PAINT message where it repaints the window.
                        ' lets put "Hello World!" at the top of it like they do on
                        ' the win32 C++ pre-made projects
            
            If message = WM_PAINT Then
                GetClientRect hWnd, rt
                hdc = BeginPaint(hWnd, PS)
                DrawText hdc, "Hello World!", Len("Hello World!"), rt, DT_CENTER
                EndPaint hWnd, PS
        
                        ' since we handled this message, return 0. dont let the
                        ' DefWindowProc handle it
                MainWndProc = 0
                
                Exit Function
            End If

                        ' watch for WM_DESTROY message, if its sent, then let the GetMessage loop in
                        ' CreateNewWindow2 know so it breaks out of the GetMessage loop
            If message = WM_DESTROY Then
                PostQuitMessage 0
                MainWndProc = 0
                Exit Function
            End If
        
            MainWndProc = DefWindowProc(hWnd, message, wParam, lParam)
        End Function


        Private Function GetAddr(ByVal iAddr As LongPtr) As LongPtr
                GetAddr = iAddr '!!!!!!!!!!
        End Function


        Public Sub UnSubWnd(ByVal hWnd As LongPtr)
            Dim pOldWndProc As LongPtr

            If IsWindow(hWnd) <> 0 Then
                If GetWindowLong(hWnd, GWL_WNDPROC) = GetAddr(AddressOf fnWndProc) Then
                    pOldWndProc = GetProp(hWnd, "pOldWndProc")
                    If pOldWndProc <> 0 Then
                        SetWindowLongPtr hWnd, GWL_WNDPROC, pOldWndProc
                
                        RemoveProp hWnd, "pOldWndProc"
                        RemoveProp hWnd, "pHandler"
                    End If
                End If
            End If
        End Sub

        Public Function CallWindProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As LongPtr
        Dim pOldWndProc As LongPtr
    
            pOldWndProc = GetProp(hWnd, "pOldWndProc")
        
            If pOldWndProc <> 0 Then
                CallWindProc = CallWindowProc(pOldWndProc, hWnd, uMsg, wParam, lParam)
            Else
                CallWindProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
            End If
        End Function

        Private Function fnWndProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As LongPtr
        Dim oHandler As ISubclass, bHandled As Boolean, ret As Variant, pOldWndProc As LongPtr, pHandler As LongPtr
    
            pOldWndProc = GetProp(hWnd, "pOldWndProc")
            pHandler = GetProp(hWnd, "pHandler")
    
            If pHandler <> 0 Then
                CopyMemory oHandler, pHandler, 4&
                ret = oHandler.WndProc(CLng(hWnd), uMsg, wParam, lParam, bHandled)
                ZeroMemory oHandler, 4&
            End If
    
            If Not bHandled Then
                If pOldWndProc <> 0 Then
                    fnWndProc = CallWindowProc(pOldWndProc, hWnd, uMsg, wParam, lParam)
                Else
                    fnWndProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
                End If
            Else
                    fnWndProc = ret
            End If
        End Function

        Public Sub SubWnd(ByVal hWnd As LongPtr, oHandler As ISubclass)
         Dim pOldWndProc As LongPtr, pHandler As LongPtr
            
            If IsWindow(hWnd) <> 0 Then
                If Not GetWindowLong(hWnd, GWL_WNDPROC) = GetAddr(AddressOf fnWndProc) Then
                        If Not oHandler Is Nothing Then
                            pOldWndProc = SetWindowLongPtr(hWnd, GWL_WNDPROC, AddressOf fnWndProc)
                            pHandler = ObjPtr(oHandler)

                            SetProp hWnd, "pOldWndProc", pOldWndProc
                            SetProp hWnd, "pHandler", pHandler
                        End If
        
                End If
            End If
        End Sub
'********************************************************************************************************
#Else
       Private Type PAINTSTRUCT
            hdc As Long
            fErase As Long
            rcPaint As RECT
            fRestore As Long
            fIncUpdate As Long
            rgbReserved(0 To 31) As Byte
        End Type
        
        Private Type DRAWTEXTPARAMS
            cbSize As Long
            iTabLength As Long
            iLeftMargin As Long
            iRightMargin As Long
            uiLengthDrawn As Long
        End Type

        Private Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpstring As String) As Long
        Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long


        Private Declare Function CallWindowProc Lib "C:\user32.dll" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
            (lpDest As Any, lpSource As Any, ByVal cbCopy As Long)
        Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpstring As String, ByVal hData As Long) As Long
        Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpstring As String) As Long
        Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpstring As String) As Long
        Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
        Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
        Private Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
        Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
        Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
        Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
        Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
        Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
        Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal Un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
        Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
        
        Public Function FnPtrToLong(ByVal lngFnPtr As Long) As Long
            FnPtrToLong = lngFnPtr
        End Function

        Public Function GetWinHook() As Long
            GetWinHook = FnPtrToLong(AddressOf MainWndProc)
        End Function

        Private Function MainWndProc(ByVal hWnd As Long, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

         Dim rt As RECT, hdc As Long, PS As PAINTSTRUCT

                        ' i put this here so that you can see how to process a message.
                        ' this is the WM_PAINT message where it repaints the window.
                        ' lets put "Hello World!" at the top of it like they do on
                        ' the win32 C++ pre-made projects
            
            If message = WM_PAINT Then
                GetClientRect hWnd, rt
                hdc = BeginPaint(hWnd, PS)
                DrawText hdc, "Hello World!", Len("Hello World!"), rt, DT_CENTER
                EndPaint hWnd, PS
        
                        ' since we handled this message, return 0. dont let the
                        ' DefWindowProc handle it
                MainWndProc = 0
                
                Exit Function
            End If

                        ' watch for WM_DESTROY message, if its sent, then let the GetMessage loop in
                        ' CreateNewWindow2 know so it breaks out of the GetMessage loop
            If message = WM_DESTROY Then
                PostQuitMessage 0
                MainWndProc = 0
                Exit Function
            End If
        
            MainWndProc = DefWindowProc(hWnd, message, wParam, lParam)
        End Function


    Private Function GetAddr(ByVal iAddr As Long) As Long
                GetAddr = iAddr '!!!!!!!!!!
    End Function


    Public Sub UnSubWnd(ByVal hWnd As Long)
            Dim pOldWndProc As Long
            

        If IsWindow(hWnd) <> 0 Then
            If GetWindowLong(hWnd, GWL_WNDPROC) = GetAddr(AddressOf fnWndProc) Then
                pOldWndProc = GetProp(hWnd, "pOldWndProc")
                If pOldWndProc <> 0 Then
                    SetWindowLong hWnd, GWL_WNDPROC, pOldWndProc
                
                    RemoveProp hWnd, "pOldWndProc"
                    RemoveProp hWnd, "pHandler"
                End If
            End If
        End If
    End Sub

        Public Function CallWindProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        Dim pOldWndProc As Long
    
            pOldWndProc = GetProp(hWnd, "pOldWndProc")
        
            If pOldWndProc <> 0 Then
                CallWindProc = CallWindowProc(pOldWndProc, hWnd, uMsg, wParam, lParam)
            Else
                CallWindProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
            End If
        End Function

        Private Function fnWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        Dim oHandler As ISubclass, bHandled As Boolean, ret As Long
        Dim pOldWndProc As Long, pHandler  As Long
        
            pOldWndProc = GetProp(hWnd, "pOldWndProc")
            pHandler = GetProp(hWnd, "pHandler")
    
            If pHandler <> 0 Then
                CopyMemory oHandler, pHandler, 4&
                ret = oHandler.WndProc(hWnd, uMsg, wParam, lParam, bHandled)
                ZeroMemory oHandler, 4&
            End If
    
            If Not bHandled Then
                If pOldWndProc <> 0 Then
                    fnWndProc = CallWindowProc(pOldWndProc, hWnd, uMsg, wParam, lParam)
                Else
                    fnWndProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
                End If
            Else
                    fnWndProc = ret
            End If
        End Function

        Public Sub SubWnd(ByVal hWnd As Long, oHandler As ISubclass)
         Dim pOldWndProc As LongPtr, pHandler As LongPtr
            
            If IsWindow(hWnd) <> 0 Then
                If Not GetWindowLong(hWnd, GWL_WNDPROC) = GetAddr(AddressOf fnWndProc) Then
                        If Not oHandler Is Nothing Then
                            pOldWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf fnWndProc)
                            pHandler = ObjPtr(oHandler)

                            SetProp hWnd, "pOldWndProc", pOldWndProc
                            SetProp hWnd, "pHandler", pHandler
                        End If
        
                End If
            End If
        End Sub
#End If
'********************************************************************************************************



'-------------------------------------------------------------------------------------------------------------------------------------------------
' Get high part of long number
'-------------------------------------------------------------------------------------------------------------------------------------------------
Public Function hiword(ByVal DWord As Long) As Integer
    hiword = (DWord And &HFFFF0000) \ &H10000
End Function
'-------------------------------------------------------------------------------------------------------------------------------------------------
' Get low part of long number
'-------------------------------------------------------------------------------------------------------------------------------------------------
Public Function loword(ByVal DWord As Long) As Integer
    If (DWord And &H8000&) = 0 Then
        loword = DWord And &HFFFF&
    Else
        loword = DWord Or &HFFFF0000
    End If
End Function


'------------------------------------------------------------------------------------------------------------------------------------------------------
' Helper function, used by the MessageBox class for the reverse hook of the MsgBox dialog box (so it is public)
' The function is used in the dialog class #_Dialog
'------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function MsgBoxHookProc(ByVal uMsg As Long, _
                                ByVal wParam As Long, _
                                ByVal lParam As Long) As Long

    #If Win64 Then
        Dim lPtr As LongPtr
        Dim lProcHook As LongPtr
    #Else
        Dim lPtr As Long
        Dim lProcHook As Long
    #End If
    
    Dim cM As cMessageBox
    
    Select Case uMsg
        Case HCBT_ACTIVATE
            lPtr = GetProp(hWndApplication, "ObjPtr")
            If (lPtr <> 0) Then
                Set cM = ObjectFromPtr(lPtr)
                If Not cM Is Nothing Then
                    If Len(cM.ButtonText1) > 0 And Len(cM.ButtonText2) > 0 And Len(cM.ButtonText3) > 0 Then
                        If cM.UseCancel Then
                            SetDlgItemText wParam, IDYES, cM.ButtonText1
                            SetDlgItemText wParam, IDNO, cM.ButtonText2
                            SetDlgItemText wParam, IDCANCEL, cM.ButtonText3
                        Else
                            SetDlgItemText wParam, IDABORT, cM.ButtonText1
                            SetDlgItemText wParam, IDRETRY, cM.ButtonText2
                            SetDlgItemText wParam, IDIGNORE, cM.ButtonText3
                        End If
                        
                    ElseIf Len(cM.ButtonText1) > 0 And Len(cM.ButtonText2) Then
                        If cM.UseCancel Then
                            SetDlgItemText wParam, IDOK, cM.ButtonText1
                            SetDlgItemText wParam, IDCANCEL, cM.ButtonText2
                        Else
                            SetDlgItemText wParam, IDYES, cM.ButtonText1
                            SetDlgItemText wParam, IDNO, cM.ButtonText2
                        End If
                    Else
                        SetDlgItemText wParam, IDOK, cM.ButtonText1
                    End If
                    lProcHook = cM.ProcHook
                End If
            End If
            RemovePropPointer
            If lProcHook <> 0 Then UnhookWindowsHookEx lProcHook
    End Select
    
    MsgBoxHookProc = False
End Function

#If Win64 Then
    Private Property Get ObjectFromPtr(ByVal lPtr As LongPtr) As Object
        Dim obj As Object
        
        CopyMemory obj, lPtr, 4
        Set ObjectFromPtr = obj
        CopyMemory obj, 0&, 4
    End Property
#Else
    Private Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
        Dim obj As Object
        
        CopyMemory obj, lPtr, 4
        Set ObjectFromPtr = obj
        CopyMemory obj, 0&, 4
    End Property
#End If

Public Sub RemovePropPointer()
    #If Win64 Then
        Dim lPtr As LongPtr
    #Else
        Dim lPtr As Long
    #End If
    
    lPtr = GetProp(hWndApplication, "ObjPtr")
    If lPtr <> 0 Then RemoveProp hWndApplication, "ObjPtr"
End Sub

'mTaskDialogHelper
'The following three lines must be in a standard module (.bas) of any project using cTaskDialog:
Public Function TaskDialogCallbackProc(ByVal hWnd As Long, ByVal uNotification As Long, ByVal wParam As Long, ByVal lParam As LongPtr, ByVal lpRefData As cTaskDialog) As LongPtr
    TaskDialogCallbackProc = lpRefData.ProcessCallback(hWnd, uNotification, wParam, lParam)
End Function
Public Function TaskDialogEnumChildProc(ByVal hWnd As LongPtr, ByVal lParam As cTaskDialog) As Long
    TaskDialogEnumChildProc = lParam.ProcessEnumCallback(hWnd)
End Function
Public Function TaskDialogSubclassProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As LongPtr, ByVal uIdSubclass As Long, ByVal dwRefData As cTaskDialog) As LongPtr
    TaskDialogSubclassProc = dwRefData.ProcessSubclass(hWnd, uMsg, wParam, lParam, uIdSubclass)
End Function
