Attribute VB_Name = "#_DIALOG"
'******************************************************************************************************************************************************
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
'                 $$$$$$$$$$F                                       ## ##  ##  ##  ##    ##   #####  ####     ### ##    ####   #####
'                  *$$$$$$$$"                                        ####  #####  ##     ##   ##  ##  ##     #### ##   ##  ## ##
'                    "***""                   _____________                                   ##  ##  ##    ## ## ##   ##  ## ## ###
' STANDARD MODULE WITH DEFAULT GUI/DIALOG FUNCTIONS |v 2021/03/15 |                           ##  ##  ##   ###### ##   ##  ## ##  ##
' The module contains frequently used functions and is part of the G-VBA library              #####  #### ##   ## ####  ####   #####
'****************************************************************************************************************************************************
'******************************************************************************************************************************************************
' Description: This Module was generated in 31-May-20 to provide dialog option
'              The class defines the dialog capabilities of the application that provide various forms of user interaction.
'   @ Valery Khvatov (valery.khvatov@gmail.com), [01/20200531]
' VBA (Access and Excel) contains many built-in features for dialogue with the user.
' Some functions (such as calling a password entry form or displaying a message with Unicode require additional effort.
'--------------------------------------------------------------------------------------------------------------------------
' This module contains additional features:
'           - [#_DIALOG].DIALOG_ABOUT: Return Version and description this module
'           - [#_DIALOG].OpenDialog: select file, directory, save file. Works with the built-in Office dialog
'                   without declarations and directly through the API. Supports saving the user's last selection (GetLastFolder )
'           - [#_DIALOG].InputBoxW: replacing the standard input dialog using additional flags (to display with an icon or replace buttons,
'                   as well as enter a hidden password)
'           - [#_DIALOG].AskUser: The simplest wrapper function over the standard MsgBox that allows you to ask the user a question and
'                   return true if hi/shee agree
'           - [#_DIALOG].MsgBoxW: A function that extends the standard MsgBox (allows you to display Unicode, set your own names for buttons,
'                   display a message with a timer, messages with hypertext, and also design a dialog box - subtitles, icons and sizes)
'           - [#_DIALOG].TaskDialogMessage: an additional version of MsgBox that allows you to customize the appearance of the message,
'                   including footer and header
'           - [#_DIALOG].SetFormIcon: set your own icon for any user form
'           - [#_DIALOG].ColorDialog: color dialog
'           - [#_DIALOG].GetOptions: a quick output from the drop-down list or the launch of a dynamic form
'                    (imported in GFORM, if the form is specified, then such a form is called))
'           - [#_DIALOG].PlaySound: plays a sound file, can be stopped via StopSound
'           - [#_DIALOG].GDialogList: alternative list selection based on GFORM (part of GetOptions )
'           - [#_DIALOG].GDialogWeb: launching the built-in browser window, in which you can display a page, graph, picture or
'                     media file (extended multimedia messages)
'
'                    ------------------------------------------
'           Attention! Some features are built using the subclass technique. It is not recommended to place breakpoints in this file,
'  / \      after changes in this module, as well as modules #_GUI , #_HELPER, and cTaskDialog, you must compile the project
' / ! \     (Debu>Compile ) and reload the base / project
' -----              ------------------------------------------
'                                                                      Additions and comments are welcome.
' ||    THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ||    ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' ||    WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' ||    DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
' ||    DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
' ||    (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
' ||    LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
' ||    ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' ||    (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
' ||    SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
Option Explicit

'************************************************************
' CONSTANT AND DECLARATIONS
Public Const APP_TITLE As String = "> GRACKLE LIBRARY"
Private Const MOD_NAME As String = "#_DIALOG"
Private Const MOD_VERSION As String = "20220401"

Private Const MAX_PATH = 260
Private Const WM_USER = &H400
Private Const WM_GETICON = &H7F
Private Const WM_SETICON = &H80

Private Const BFFM_INITIALIZED As Long = 1
Private Const BFFM_SETSELECTION As Long = WM_USER + 102

Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1

' Flags for the options parameter for folder dialog
Private Const BIF_returnonlyfsdirs = &H1
Private Const BIF_dontgobelowdomain = &H2
Private Const BIF_statustext = &H4
Private Const BIF_returnfsancestors = &H8
Private Const BIF_editbox = &H10
Private Const BIF_validate = &H20
Private Const BIF_browseforcomputer = &H1000
Private Const BIF_browseforprinter = &H2000
Private Const BIF_browseincludefiles = &H4000
Private Const BIF_nonewfolder = &H200

'Constants to be used in our API functions (windows actions)
    Private Const EM_SETPASSWORDCHAR = &HCC
    Private Const WH_CBT = 5
    Private Const HCBT_ACTIVATE = 5
    Private Const HC_ACTION = 0

' LoadImage() image types
    Private Const IMAGE_ICON = 1

' LoadImage() flags
    Private Const LR_LOADFROMFILE = &H10

Private Const MINUS_LIMIT As Integer = -20000
Private Const ERR_WRONG_FORMAT As Long = 10777

Private sMsgBoxDefaultLabel(1 To 7) As String
Private sMsgBoxCustomLabel(1 To 7) As String
Private bMsgBoxCustomInit As Boolean

'************************************************************
' MAIN API AND DEPENDENT STRUCTURES DECLARATION
#If VBA7 Then

    Private Type OPENFILENAME
        lStructSize As Long
        hWndOwner As LongPtr
        hInstance As LongPtr
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        FLAGS As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As LongPtr
        lpfnHook As LongPtr
        lpTemplateName As String
    End Type
    
    Private Type BrowseInfo
        hOwner As LongPtr
        pIDLRoot As Long
        pszDisplayName As String
        lpszTitle As String
        ulFlags As Long
        lpfn As LongPtr
        lParam As LongPtr
        iImage As Long
    End Type
    
    Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
                                                  lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
    Private Declare PtrSafe Sub wlib_AccColorDialog Lib "msaccess.exe" Alias "#53" (ByVal hWnd As LongPtr, lngRGB As Long)
    
    Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BrowseInfo) As LongPtr
    Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidList As LongPtr, ByVal lpBuffer As String) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
    Private Declare PtrSafe Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As LongPtr, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As LongPtr
    Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
    
    Private Declare PtrSafe Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As LongPtr, ByVal lpstring As String) As LongPtr
    Private Declare PtrSafe Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As LongPtr, ByVal lpstring As String) As Long
    
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, _
        ByVal ncode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
        (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hMod As LongPtr, ByVal dwThreadId As Long) As LongPtr
    Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
    Private Declare PtrSafe Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" _
        (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As LongPtr, _
         ByVal nIDDlgItem As LongPtr, ByVal lpstring As String) As LongPtr
                
    Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias _
        "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
    Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" _
        (ByVal hWnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    
    
        
    Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long
    
    Private hHook As LongPtr
'******************************************************
#Else
    Public Type OPENFILENAME
        lStructSize As Long
        hWndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        FLAGS As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
    End Type
    
    Private Type BrowseInfo
        hOwner As Long
        pIDLRoot As Long
        pszDisplayName As String
        lpszTitle As String
        ulFlags As Long
        lpfn As Long
        lParam As Long
        iImage As Long
    End Type
         
    Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
                                                  lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
    Private Declare Sub wlib_AccColorDialog Lib "msaccess.exe" Alias "#53" (ByVal Hwnd As Long, lngRGB As Long)

    Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BrowseInfo) As Long
    Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
    Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
    
    Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
    Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long

    
    Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, _
        ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
        (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
    
    Private Declare Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" _
        (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As Long, _
        ByVal nIDDlgItem As Long, ByVal lpString As String) As Long

    Private Declare Function GetModuleHandle Lib "kernel32" Alias _
        "GetModuleHandleA" (ByVal lpModuleName As String) As Long
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
        (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

    Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long,ByVal wType As Long) As Long
    
    Private hHook As Long
#End If


'**************************************************
' DIALOG SPECIFIC ENUMS
Public Enum grcDialogType
        GC_OPEN_FILE = 1
        GC_SAVE_AS = 2
        GC_FILE_PICKER = 3
        GC_FOLDER_PICKER = 4
End Enum

Public Enum grcOptionType
        GC_POPUP_STYLE = 1
        GC_RADIO_STYLE = 2
End Enum

Public Enum grcInputType
        GC_EASY_STYLE = 1
        GC_EXTENDED_STYLE = 2
        GC_PASSWORD_STYLE = 3
End Enum

Public Enum grcICONS         ' GARMONIZE WITH TASK DIALOG
    GC_WARNING_ICON = -1                    'exclamation point in a yellow 'yield' triangle (same image as IDI_EXCLAMATION)
    GC_ERROR_ICON = -2                      'round red circle containg 'X' (same as IDI_HAND)
    GC_INFORMATION_ICON = -3                'round blue circle containing 'i' (same image as IDI_ASTERISK)
    GC_SHIELD_ICON = -4                     'Vista's security shield
    GC_APPLICATION = 32512&                 'miniature picture of an application window
    GC_ERROR = 32513&                       'error
    GC_QUESTION = 32514&                    'round blue circle containing '?'
    GC_WINLOGO = 32517&
    GC_SHIELD_GRADIENT_ICON = -5            'same image as TD_SHIELD_ICON; main message text on gradient blue background
    GC_SHIELD_WARNING_ICON = -6             'exclamation point in yellow Shield shape; main message text on gradient orange background
    GC_SHIELD_ERROR_ICON = -7               'X contained within Shield shape; main message text on gradient red background
    GC_SHIELD_OK_ICON = -8                  'Shield shape containing green checkmark; main message text on gradient green background
    GC_SHIELD_GRAY_ICON = -9                'same image as TD_SHIELD_ICON; main message text on medium gray background
    GC_NO_ICON = 0                          'no icon; text on white background
End Enum

Public Enum grcMsgBoxType
        GC_UNICODE = 1
        GC_SHELL = 2
        GC_CUSTOM_LBL = 3
        GC_YESNO_MSG = 4
        GC_WIDE_MSG = 5
        GC_SUCCESS_MSG = 6
        GC_WARNING_MSG = 7
End Enum

'**********************
' GLOBAL VARS
    Private slRootFolder As String

'******************************************************************************************************
'======================================================================================================================================================
' About this module
'======================================================================================================================================================
Public Function DIALOG_ABOUT()
Dim sRes As String

Const PRE_DEF_TEXT As String = "Part of the GRACKLE library for displaying custom dialogs." & _
                               " It is an extension of the #_GUI module and contains a file dialog, a color control dialog, an input dialog," & _
                               " and text output dialogs that extend the built-in VBA capabilities."

      DIALOG_ABOUT = PRE_DEF_TEXT & vbCrLf & "VERS: " & MOD_VERSION
End Function

'======================================================================================================================================================
' This function return True If user press YES for question
'======================================================================================================================================================
Public Function AskUser(Optional Question As String = "Do you want to continue?", Optional Title As String = APP_TITLE) As Boolean
   AskUser = MsgBox(Question, vbYesNo + vbQuestion, Title) = vbYes '!!!!!!!!!!!!!!!
End Function
'======================================================================================================================================================
' Advanced getting options in the form of a dialog: mode = 1 is minimal context dialog (based on cPopUpMenuClass);
'                                                   mode = 2 - extended window,
' The icon is only shown if Prompt is not empty. Requre: class cPopUpMenu and cTaskDialog
' Return Number starting from 1
'======================================================================================================================================================
Public Function GetOptions(Optional sList As String = "Option1;Option2;Option3", Optional DLM As String = ";", _
                                                        Optional sTitle As String = "GetOptions", Optional iMode As grcOptionType = GC_POPUP_STYLE, _
                 Optional sPrompt As String = "SELECT THE OPTIONS", Optional sDescription As String, Optional iIcon As grcICONS = GC_QUESTION) As Integer
                 
Dim sMenu As String, OPTS() As String, nDim As Integer, I As Integer
Dim iRes As Integer, MNU As cPopUpMenu, TaskDialog As cTaskDialog

    On Error GoTo ErrHandle
'------------------------
    iRes = -1: If sList = "" Then GoTo ExitHere
    OPTS = Split(sList, DLM): nDim = UBound(OPTS)
    
    Select Case iMode
    Case GC_POPUP_STYLE:                            ' Context Menu Based
           Set MNU = New cPopUpMenu
           With MNU
                .Caption = sTitle
                For I = 0 To nDim
                     .AddItem I + 1, OPTS(I)
                Next I
                
                iRes = .ShowPopup()
           End With
           
    Case GC_RADIO_STYLE:                             ' Separate Option Based on TaskDialog
           Set TaskDialog = New cTaskDialog
           With TaskDialog
                .Init: .Title = sTitle
                
                
                If Not IsBlank(sPrompt) Then
                        .MainInstruction = sPrompt
                        .IconMain = iIcon
                End If
                
                If Not IsBlank(sDescription) Then .Content = sDescription
                        
                .CommonButtons = TDCBF_CANCEL_BUTTON _
                                             Or TDCBF_OK_BUTTON
                                             
                For I = 0 To nDim
                        .AddRadioButton I + 1, OPTS(I)
                Next I
                
                .ShowDialog
                If .ResultMain = TD_OK Then iRes = .ResultRad
            End With
    Case Else
    End Select
'------------------------
ExitHere:
    GetOptions = iRes '!!!!!!!!!!!!
    Set MNU = Nothing: Set TaskDialog = Nothing
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "GetOptions", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Resume ExitHere
End Function

'======================================================================================================================================================
' The function customizes the dialog for text input, expanding the capabilities of the standard InputBox.
' To switch to advanced mode, which allows you to set icons, buttons - you must specify iMode = 2
'======================================================================================================================================================
Public Function InputBoxW(Optional sPrompt As String = "Please, Input Text", Optional sTitle As String = APP_TITLE, _
                                                             Optional sDefault As String, Optional XPos As Single = -1, Optional YPos As Single = -1, _
                                    Optional iMode As grcInputType = GC_EASY_STYLE, Optional BUTTONS As Long = TDCBF_OK_BUTTON Or TDCBF_CANCEL_BUTTON, _
                                                                                                Optional iIcon As Long = TD_INFORMATION_ICON) As String
Dim sRes As String, TaskDialog As cTaskDialog, iHwnd As Variant

    On Error GoTo ErrHandle
'---------------------------
Select Case iMode
Case GC_EASY_STYLE:
    If XPos >= 0 And YPos >= 0 Then
        sRes = InputBox(sPrompt, sTitle, sDefault, XPos, YPos)
    Else
        sRes = InputBox(sPrompt, sTitle, sDefault)
    End If

Case GC_EXTENDED_STYLE:
    Set TaskDialog = New cTaskDialog
    iHwnd = GetActiveHwnd()

    With TaskDialog
        .Init
        .Content = sPrompt
        .FLAGS = TDF_INPUT_BOX
        .CommonButtons = BUTTONS
        .IconMain = iIcon
        .Title = sTitle
        .ParenthWnd = GetActiveHwnd 'Get active hWnd (Access.hWndAccessApp from CLI)
        .ShowDialog
       
        If .ResultMain = TD_OK Then
            sRes = .ResultInput
        End If
    End With
Case GC_PASSWORD_STYLE:
    sRes = InputBoxDK(sPrompt, sTitle)
End Select
'---------------------------
ExitHere:
    InputBoxW = sRes '!!!!!!!!!!!!!!!!!!!!!
    Set TaskDialog = Nothing
    Exit Function
'--------
ErrHandle:
    ErrPrint2 "InputBoxW", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'======================================================================================================================================================
' Open ColorDialog
'======================================================================================================================================================
Public Function ColorDialog(Optional StartColor As Variant, Optional iColorFormat As Integer = 0) As Variant
Dim lngColor As Long, strColor, vRes As Variant, iBColor As Long

Const DEFAULT_COLOR As String = "#FFFFFF"

    On Error GoTo ErrHandle
'--------------------------
If Not IsMissing(StartColor) Then
     If IsNumeric(StartColor) Then
            lngColor = CLng(StartColor)
     Else
            lngColor = CLng("&H" & Right("000000" + _
                  Replace(Nz(StartColor, ""), "#", ""), 6))
     End If
Else
            lngColor = CLng("&H" & Right("000000" + _
                  Replace(Nz(DEFAULT_COLOR, ""), "#", ""), 6))
End If
  
  iBColor = lngColor
  wlib_AccColorDialog GetActiveHwnd(), lngColor
  
  If lngColor = iBColor Then
        vRes = -1: GoTo ExitHere
  End If
  
  If iColorFormat = 0 Then
        vRes = lngColor
  Else
        vRes = Right("000000" & Hex(lngColor), 6)
  End If
'--------------------------
ExitHere:
    ColorDialog = vRes '!!!!!!!!!!!!!!!!!!!!
    Exit Function
'------------
ErrHandle:
    ErrPrint2 "ColorDialog", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Resume ExitHere
End Function

'======================================================================================================================================================
' Play Music file
'======================================================================================================================================================
Public Function PlaySound(ByVal sPath As String, Optional Duration As Long = -1) As Boolean
Dim iPlay As Long, sFile As String, bRes As Boolean
    
    On Error GoTo ErrHandle
'-------------------------
'[1] Check Path
    If sPath = "" Then Exit Function
    If InStr(1, sPath, ":") > 0 Then   ' This is real path
        If Dir(sPath) = "" Then Exit Function
        sFile = GetShortPath(sPath)
    Else                               ' This is system sound file
        sFile = GetMediaPath(sPath): If sFile = "" Then Exit Function
    End If
    
'[2] Start Paying
iPlay = mciSendString("play " & sFile, 0&, 0, 0)
If iPlay = 0 Then
    bRes = True
    If Duration > 0 Then
        Wait Duration
        StopSound sPath
    End If
End If
'-------------------------
ExitHere:
    PlaySound = bRes '!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "PlaySound", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'======================================================================================================================================================
' Stop Music file (if it is playing)
'======================================================================================================================================================
Public Sub StopSound(Optional sPath As String)
Dim iPlay As Long, sFile As String

    On Error Resume Next
'----------------------
    sFile = GetShortPath(sPath)
    iPlay = mciSendString("close " & sFile, 0&, 0, 0)
'----------------------
    
End Sub
'======================================================================================================================================================
' Open Dialog with Web Form
'======================================================================================================================================================
Public Function GDialogWeb(Optional sLink As String = "https://www.google.com/", Optional sTitle As String = "Web Dialog", _
                                                                                       Optional sPrompt As String = "Plase, review dynamic content", _
                                                                                      Optional iZoom As Long, Optional DLM As String = ";") As Boolean
Dim bRes As Boolean, sFormName As String, sRef As String

Const FORM_NAME As String = "g_Web"
Const DIALOGWEB As String = "DIALOGWEB"

On Error GoTo ErrHandle
'----------------------------
    sFormName = GFORM(FORM_NAME): If sFormName = "" Then Err.Raise 11107, , "Can't found the form " & FORM_NAME

    TempVars(DIALOGWEB).value = ""
    sRef = sLink & DLM & sTitle & DLM & sPrompt & DLM & iZoom
    
    DoCmd.OpenForm FORM_NAME, , , , , , sRef
    
    sRef = Nz(TempVars(DIALOGWEB).value): If sRef <> "" Then bRes = True
'----------------------------
ExitHere:
    GDialogWeb = bRes '!!!!!!!!!!!!!!!!!
    Exit Function
'---------
ErrHandle:
    ErrPrint2 "GDialogWeb", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'======================================================================================================================================================
' Open Dialog with List (GFORMS are required - see #_PLUGINS)
'======================================================================================================================================================
Public Function GDialogList(Optional sTitle As String = "Dialog List", Optional sOptions As String = "Option1;Option2;Option3;Option4;Option5", _
                                                                                                               Optional DLM As String = ";") As String
Dim sRef As String, sFormName As String

Const FORM_NAME As String = "g_DialogList"
Const DIALOGLIST As String = "DIALOGLIST"

      On Error GoTo ErrHandle
'-------------------
      sFormName = GFORM(FORM_NAME): If sFormName = "" Then Err.Raise 11107, , "Can't found the form " & FORM_NAME

      TempVars(DIALOGLIST).value = ""
      sRef = sTitle & DLM & sOptions
      
      DoCmd.OpenForm FORM_NAME, , , , , acDialog, sRef
'-------------------
ExitHere:
      GDialogList = Nz(TempVars(DIALOGLIST).value)
      Exit Function
'---------
ErrHandle:
      ErrPrint2 "GDialogList", Err.Number, Err.Description, MOD_NAME
      Err.Clear
End Function

'======================================================================================================================================================
' The Folder Dialof with options
'======================================================================================================================================================
Public Function Browse4Folder(Optional strPrompt As String = "Pick a folder", Optional strRoot As String, _
                                                             Optional bEditBox As Boolean = True, Optional bNoNewFolder As Boolean = False) As String
Dim objFolder As Object, objFolderItem As Object, objShell As Object
Dim intOptions As Integer, sRes As String, sInitialPath As String
    
    On Error GoTo ErrHandle
'--------------------------------
    intOptions = BIF_returnonlyfsdirs
    If bEditBox Then intOptions = intOptions + BIF_editbox
    If bNoNewFolder Then intOptions = intOptions + BIF_nonewfolder
    
    sInitialPath = strRoot: If sInitialPath = "" Then sInitialPath = CurrentProject.Path
    
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(0, strPrompt, intOptions, strRoot)
    
    If Not (objFolder Is Nothing) Then
        Set objFolderItem = objFolder.Self
        sRes = objFolderItem.Path
        Set objFolderItem = Nothing
        Set objFolder = Nothing
    End If
'--------------------------------
ExitHere:
    Browse4Folder = sRes '!!!!!!!!!!
    Set objShell = Nothing
    Exit Function
'----------------
ErrHandle:
    ErrPrint2 "Browse4Folder", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'====================================================================================================================================================
' Open File Dialog:
' 1 - OpenFile;2 - SaveAs; 4 = OpenFolder
' If InitialFolder is empty then open folder via LastFolder
' If LastFolder is empty also,  then it will pen current project folder and reset LastFolder
'====================================================================================================================================================
Public Function OpenDialog(Optional nDialogType As grcDialogType = GC_OPEN_FILE, Optional sTitle As String, _
                    Optional sFilters As String = "All Files,*.*;Text Files,*.txt", Optional bAllowMultiSelect As Boolean = True, _
                    Optional InitialFolder As String = "", Optional bUseAPI As Boolean)
Dim fDialog As Object
Dim varfile As Variant
Dim sRes As String
Dim sWork() As String, nDim As Integer, I As Integer
Dim sInitial As String
    
On Error GoTo ErrHandle
'---------------------------------------------------------------
If nDialogType = GC_FOLDER_PICKER And bUseAPI Then
    'sRes = GetFolderDialog(, InitialFolder, Left(sTitle, 60))
    sRes = Browse4Folder(sTitle, InitialFolder)
ElseIf nDialogType = GC_FILE_PICKER And bUseAPI Then
    sRes = OpenFileDialog(InitialFolder, sTitle, , sFilters)
Else
    Set fDialog = Application.FileDialog(nDialogType)
    
    With fDialog
          sInitial = IIf(InitialFolder <> "", InitialFolder, GetLastFolder)
          .InitialFileName = IIf(sInitial <> "", sInitial, CurrentProject.Path & "\")
          .AllowMultiSelect = bAllowMultiSelect
            
            Select Case nDialogType
                Case GC_OPEN_FILE:     ' = 1
                        .Title = IIf(sTitle <> "", sTitle, "Open File")
                        .Filters.Clear
                                sWork = Split(sFilters, ";"): nDim = UBound(sWork)
                        For I = 0 To nDim
                            .Filters.Add Split(sWork(I), ",")(0), Split(sWork(I), ",")(1)
                        Next I
                Case GC_SAVE_AS        ' = 2
                        .Title = IIf(sTitle <> "", sTitle, "Save File As")
                Case GC_FILE_PICKER    ' = 3
                        .Title = IIf(sTitle <> "", sTitle, "Pick the File")
                        .Filters.Clear
                                sWork = Split(sFilters, ";"): nDim = UBound(sWork)
                        For I = 0 To nDim
                            .Filters.Add Split(sWork(I), ",")(0), Split(sWork(I), ",")(1)
                        Next I
                Case GC_FOLDER_PICKER  ' = 4
                        .Title = IIf(sTitle <> "", sTitle, "Pick the Folder")
                Case Else
                        Err.Raise 1000, , "WRONG DIALOG_TYPE"
                End Select
                
          If .Show = True Then
             For Each varfile In .SelectedItems
                sRes = IIf(sRes <> "", sRes & ";" & varfile, varfile)
             Next
          End If
    End With
End If
'-----------------------------------------------------------------------------
ExitHere:
           OpenDialog = sRes '!!!!!!!!!!!!!!!!
           If sRes <> "" Then SetLastFolder FolderNameOnly(sRes)
           
           Set fDialog = Nothing
           Exit Function
'-------------------------------------------
ErrHandle:
           ErrPrint2 "OpenDialog", Err.Number, Err.Description, MOD_NAME
           Err.Clear:            Resume ExitHere
End Function


'======================================================================================================================================================
' Set Form Icon
'======================================================================================================================================================
Public Function SetFormIcon(hWnd, IconPath As String) As Boolean
#If VBA7 Then
    Dim hIcon As LongPtr, iHwnd As LongPtr
    iHwnd = CLngPtr(hWnd)
#Else
    Dim hIcon As Long, iHwnd As Long
    iHwnd = CLng(hWnd)
#End If

   hIcon = LoadImage(0&, IconPath, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)

   '// wParam = 0; Setting small icon. wParam = 1; setting large icon
   If hIcon <> 0 Then
      Call SendMessage(iHwnd, WM_SETICON, ICON_SMALL, ByVal hIcon)
      SetFormIcon = True
   End If
End Function


'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////







'******************************************************************************************
'******************************************************************************************
'******************************************************************************************

'======================================================================================================================================================
' Window MessageBox Wrapper - to show Unicode Character
'======================================================================================================================================================
Public Function MsgBoxW(Prompt As String, Optional BUTTONS As VbMsgBoxStyle = vbOKOnly, Optional Title As String = APP_TITLE, _
                                                        Optional iMode As grcMsgBoxType = GC_SHELL, Optional AckTime As Integer = -1, _
                                                Optional CustomButtonText As String = "Custom1;Custom2", Optional Instruction As String, Optional DLM As String = ";") As VbMsgBoxResult

Dim iRes As VbMsgBoxResult, sPrompt As String, sTitle As String
Dim SH As Object, LBLS() As String, nLBLS As Integer, iSyle As VbMsgBoxStyle, sInstruction As String

    On Error GoTo ErrHandle
'------------------------
    Select Case iMode
    Case GC_UNICODE:
            sPrompt = Prompt & vbNullChar 'Add null terminators
            sTitle = Title & vbNullChar
            iRes = MessageBoxW(GetActiveHwnd, StrPtr(sPrompt), StrPtr(sTitle), BUTTONS)
    Case GC_SHELL:
            Set SH = CreateObject("WScript.Shell")
            iRes = SH.PopUp(Prompt, AckTime, Title, BUTTONS)
    Case GC_CUSTOM_LBL:                    '2-3 custom Label
            If IsBlank(CustomButtonText) Then Exit Function
            LBLS = Split(CustomButtonText, DLM): nLBLS = UBound(LBLS)
            If nLBLS = 0 Then
                  MsgBoxCustom_Set vbOK, LBLS(0): iSyle = vbOKOnly
            ElseIf nLBLS = 1 Then
                  MsgBoxCustom_Set vbOK, LBLS(0): MsgBoxCustom_Set vbCancel, LBLS(1)
                  iSyle = vbOKCancel
            Else
                  MsgBoxCustom_Set vbYes, LBLS(0): MsgBoxCustom_Set vbNo, LBLS(1): MsgBoxCustom_Set vbCancel, LBLS(2)
                  iSyle = vbYesNoCancel
            End If
                        
            Call MsgBoxCustom(iRes, Prompt, iSyle, Title)
            iRes = ConvertCustomResults(iSyle, iRes)
    Case GC_WIDE_MSG:
             sInstruction = Instruction: If IsBlank(sInstruction) Then sInstruction = "Please read carefully"
             iRes = TaskDialogMessage(Prompt, , sInstruction, BUTTONS, TD_SHIELD_GRADIENT_ICON, , , 400)
    Case GC_YESNO_MSG:
             sInstruction = Instruction: If IsBlank(sInstruction) Then sInstruction = "Ñonfirmation required"
             iRes = CLng(MsgYesNo(sInstruction, Prompt))
    Case GC_SUCCESS_MSG:
             sInstruction = Instruction: If IsBlank(sInstruction) Then sInstruction = "Congratulations, the process was completed successfully"
             Call MsgInfo(sInstruction, Prompt, TD_SHIELD_OK_ICON): iRes = vbOK
    Case GC_WARNING_MSG:
             sInstruction = Instruction: If IsBlank(sInstruction) Then sInstruction = "Please have information"
             Call MsgInfo(sInstruction, Prompt, TD_SHIELD_WARNING_ICON): iRes = vbOK
                        
    Case Else:
            iRes = MsgBox(Prompt, BUTTONS, Title)
    End Select
'------------------------
ExitHere:
    MsgBoxW = iRes '!!!!!!!!!
    Set SH = Nothing
    Exit Function
'------------
ErrHandle:
    ErrPrint2 "MsgBoxW", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'======================================================================================================================================================
' TEST Custom MessageBox
' The function is optional, but allows you to check the specifics of MsgBox calls
'======================================================================================================================================================
Public Sub Custom_MsgBox_Test()
Dim iRes As VbMsgBoxResult, iCorrect As Integer

    
'[1] Simple MessageBox based on TaskDialogMessage Class
    iRes = TaskDialogMessage("TD_MsgBox: Test Default Param Call"): If iRes = 8 Then iCorrect = iCorrect + 1
    
'[2] MessageBox based on TaskDialogMessage Class with HyperLink
'    iRes = TaskDialogMessage("TD_MsgBox: Test Message with Hyperlink (press OK): " & vbNewLine & _
'                                         "This is a link to my website <a href=""https://www.AccessUI.com"">AccessUI.com</a>", , , _
'                                         TDCBF_CANCEL_BUTTON Or TDCBF_OK_BUTTON, , TDF_ENABLE_HYPERLINKS)
'    If iRes = vbOK Then iCorrect = iCorrect + 1
'[3] MessageBox based on TaskDialogMessage Class with Expanded Footer
'    iRes = TaskDialogMessage("TD_MsgBox: Test Message with Expanded Foot (press OK): ", , , _
'                                         TDCBF_CANCEL_BUTTON Or TDCBF_OK_BUTTON, IDI_QUESTION, TDF_EXPAND_FOOTER_AREA, _
'                                         "More Information...;Additional information by expanded footer", 250)
'    If iRes = vbOK Then iCorrect = iCorrect + 1
'[4] MessageBox based on TaskDialogMessage Class with Extra Width
'    iRes = TaskDialogMessage("TD_MsgBox: Test Message with Extra Wide (press YES): ", , , _
'                                         TDCBF_YES_BUTTON Or TDCBF_NO_BUTTON, TD_SHIELD_GRADIENT_ICON, , , 400)
'    If iRes = vbYes Then iCorrect = iCorrect + 1
'[5] MessageBox based on TaskDialogMessage Class with Custom Icon (!!!!! MUST HAVE THE ICON FILE ON HDD)
'     iRes = TaskDialogMessage("TD_MsgBox: Test Message with Custom Icon (press YES): ", , "Look for Custom Icon", _
'                                         TDCBF_YES_BUTTON Or TDCBF_NO_BUTTON, , TDF_USE_HICON_MAIN, , , , "\DataBase80.png")
'     If iRes = vbYes Then iCorrect = iCorrect + 1
    
'[6]   MsgBoxW  With Shell & Timing
'      iRes = MsgBoxW("Test MsgBoxW with WinShell Call (Press Yes)", vbYesNoCancel, , GC_SHELL)
'      If iRes = vbYes Then iCorrect = iCorrect + 1
      
'[7]   MsgBoxW  With Unicode
'      iRes = MsgBoxW("Test MsgBoxW with Unicode (Press Yes)" & vbCrLf & ChrW(670), vbYesNoCancel, , GC_UNICODE)
'      If iRes = vbYes Then iCorrect = iCorrect + 1
      
'[8]   MsgBoxW  With Custom Labels
'      iRes = MsgBoxW("Test MsgBoxW with Custom Labels (Press [Oh ya...])", vbYesNoCancel, , GC_CUSTOM_LBL, , "Oh ya...;Nooo")
'      If iRes = 1 Then iCorrect = iCorrect + 1
      
'[9]   MsgBoxW  With Warning
'      iRes = MsgBoxW("Test MsgBoxW with WARNINGS (Press OK)", vbOKOnly, , GC_WARNING_MSG, , , "WARNINGS!!!!")
'      If iRes = vbOK Then iCorrect = iCorrect + 1
      
'[10]  MsgBoxW  Extra WIDE
'      iRes = MsgBoxW("Extra wide message (Press OK)", vbOKOnly, , GC_WIDE_MSG, , , "EXYTA WIDE!!!!")
'      If iRes = vbOK Then iCorrect = iCorrect + 1
      
'[11]  MsgBoxW  With Warning
'       iRes = MsgBoxW("Extra wide message (Press YES)", vbYesNo, , GC_YESNO_MSG)
'       If iRes = vbYes Then iCorrect = iCorrect + 1
      
'If iCorrect = 11 Then Debug.Print "PASS TEST"
End Sub
'======================================================================================================================================================
' Message Box, Based On TaskDialog Class with extra features
'                 ---------------------------------------------------------------------------
' [!!!!] Attention, the function works with a classifier and requires prior compilation and saving in terms of dependent modules #_GUI and #_HELPER
'                 ----------------------------------------------------------------------------
' The method makes extensive use of subclassing and depends of the cTaskDialog class and standard modules (primarily #_GUI and #_HELPER ).
' Due to VBA features, changes in these modules require the project to be recompiled and restarted, otherwise the project may crash.
' To perform the necessary actions in the IDE/VBE, select Debug -> Compile, then close the database and reopen it. The crash problem can go away.
' Use the MsgBoxW function with safer flags otherwise.
'
' EXAMPLES:
'       MsgBox with HyperLinks:              THEBUTTONS = TDCBF_CLOSE_BUTTON; bHyperLinkAllowed  = True
'                                            Prompt = "Try this with a standard MsgBox!" & vbNewLine & _
'                                                   "This is a link to my website <a href=""https://www.AccessUI.com"">AccessUI.com</a>"
'                                            THEFLAGS = TDF_ENABLE_HYPERLINKS
'
'       MsgBox with Expanded Text in Footer: THEBUTTONS = TDCBF_CANCEL_BUTTON Or TDCBF_OK_BUTTON
'                                            THEFLAGS = TDF_EXPAND_FOOTER_AREA
'                                            THEICON = IDI_QUESTION
'                                            THEWIDTH = 250
'
'       MsgBox with Extra Width:             THEBUTTONS = TDCBF_YES_BUTTON Or TDCBF_NO_BUTTON
'                                            THEICON = TD_SHIELD_GRADIENT_ICON
'                                            THEWIDTH = 400
'
'      MsgBox with Custom Icon:              THEBUTTONS = TDCBF_YES_BUTTON Or TDCBF_NO_BUTTON
'                                            Prompt = "Test Message"
'                                            Instruction = "Main Instruction"
'                                            THEFLAG = TDF_USE_HICON_MAIN
'                                            CustomIconSrc  = "\DataBase80.png" (Could be File or Name Of Local resource)
'======================================================================================================================================================
Public Function TaskDialogMessage(Prompt As String, Optional Title As String = APP_TITLE, Optional Instruction As String, _
      Optional THEBUTTONS As Long = TDCBF_CLOSE_BUTTON, Optional THEICON As TDICONS = MINUS_LIMIT, Optional THEFLAGS As TASKDIALOG_FLAGS = MINUS_LIMIT, _
          Optional FOOTER_MSG As String = "More information...;Additional information  by expanded footer", Optional THEWIDTH As Integer = MINUS_LIMIT, _
                    Optional FOOTER_ICON As TDICONS = IDI_APPLICATION, Optional CustomIconSrc As String, Optional DLM As String = ";") As VbMsgBoxResult
Dim iRes As VbMsgBoxResult, sInstruction As String
Dim TaskDialog As cTaskDialog, UEXP() As String, hIcon As Long



    On Error GoTo ErrHandle
'-----------------------------
    If Prompt = "" Then Exit Function
    
    sInstruction = Instruction: If IsBlank(sInstruction) Then sInstruction = "Please read important information"
    Set TaskDialog = New cTaskDialog

    With TaskDialog
        .Init
        
         '[1]   SET FLAGS
         If THEFLAGS > MINUS_LIMIT Then .FLAGS = THEFLAGS
         
         '[2]   FOOTER PROCESSING
         If THEFLAGS = TDF_EXPAND_FOOTER_AREA And Not IsBlank(FOOTER_MSG) Then
               UEXP = Split(FOOTER_MSG, DLM): If UBound(UEXP) < 1 Then Err.Raise ERR_WRONG_FORMAT, , "Wrong Text Format"
               
               .ExpandedControlText = UEXP(0): .ExpandedInfo = UEXP(1)
               .IconFooter = FOOTER_ICON
         End If
        
        '[3]    SET MAJORS PARAMS
        .Title = Title
        .MainInstruction = sInstruction
        .Content = Prompt
        .CommonButtons = THEBUTTONS
        
        '[4]    SET WIDTH
        If THEWIDTH > MINUS_LIMIT Then .Width = THEWIDTH
        
        '[3]    SET ICONS
        If IsBlank(CustomIconSrc) Then
            If THEICON > MINUS_LIMIT Then .IconMain = THEICON
        Else
            If THEFLAGS = TDF_USE_HICON_MAIN Then
                    hIcon = CLng(RESOURCES.LoadIcon(CustomIconSrc))
                    If hIcon <> -1 Then .IconMain = hIcon
            End If
        End If
        
        
        '[4]    SHOW DIALOG AND GET RESULT
        .ShowDialog
        iRes = CLng(.ResultMain)
    End With
    
    
'-----------------------------
ExitHere:
    TaskDialogMessage = iRes '!!!!!!!!!!
    If hIcon <> -1 Then Call DestroyIconG(hIcon)
    Set TaskDialog = Nothing
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "TaskDialogMessage", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Resume ExitHere
End Function



'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' INPUT DIALOG HELPERS

'------------------------------------------------------------------------------------------------------------------------------------------------------
'Password masked inputbox
'Allows you to hide characters entered in a VBA Inputbox.
'
'Code written by Daniel Klann
'March 2003
'64-bit modifications developed by Alexey Tseluiko
'and Ryan Wells (wellsr.com)
'February 2019
'------------------------------------------------------------------------------------------------------------------------------------------------------
#If VBA7 Then
    Private Function NewProc(ByVal lngCode As Long, ByVal wParam As Long, ByVal lParam As Long) As LongPtr
#Else
    Private Function NewProc(ByVal lngCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If

    Dim RetVal
    Dim strClassName As String, lngBuffer As Long
    If lngCode < HC_ACTION Then
        NewProc = CallNextHookEx(hHook, lngCode, wParam, lParam)
        Exit Function
    End If

    strClassName = String$(256, " ")
    lngBuffer = 255
    If lngCode = HCBT_ACTIVATE Then 'A window has been activated
        RetVal = GetClassName(wParam, strClassName, lngBuffer)
        If Left$(strClassName, RetVal) = "#32770" Then
            'This changes the edit control so that it display the password character *.
            'You can change the Asc("*") as you please.
            SendDlgItemMessage wParam, &H1324, EM_SETPASSWORDCHAR, Asc("*"), &H0
        End If
    End If
    'This line will ensure that any other hooks that may be in place are
    'called correctly.
    CallNextHookEx hHook, lngCode, wParam, lParam
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------
' Function to get password
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function InputBoxDK(Prompt, Title) As String
#If VBA7 Then
    Dim lngModHwnd As LongPtr
#Else
    Dim lngModHwnd As Long
#End If

    Dim lngThreadID As Long
    lngThreadID = GetCurrentThreadId
    lngModHwnd = GetModuleHandle(vbNullString)
    hHook = SetWindowsHookEx(WH_CBT, AddressOf NewProc, lngModHwnd, lngThreadID)
    InputBoxDK = InputBox(Prompt, Title)
    UnhookWindowsHookEx hHook
End Function


'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' MSGBOX HELPERS

Private Sub MsgBoxCustom_Set(ByVal nID As Integer, Optional ByVal vLabel As Variant)
' Set button nID label to CStr(vLabel) for Public Sub MsgBoxCustom
' vbOK = 1, vbCancel = 2, vbAbort = 3, vbRetry = 4, vbIgnore = 5, vbYes = 6, vbNo = 7
' If nID is zero, all button labels will be set to default
' If vLabel is missing, button nID label will be set to default
' vLabel should not have more than 10 characters (approximately)
    If nID = 0 Then Call MsgBoxCustom_Init
    If nID < 1 Or nID > 7 Then Exit Sub
    If Not bMsgBoxCustomInit Then Call MsgBoxCustom_Init
    If IsMissing(vLabel) Then
        sMsgBoxCustomLabel(nID) = sMsgBoxDefaultLabel(nID)
    Else
        sMsgBoxCustomLabel(nID) = CStr(vLabel)
    End If
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Display standard VBA MsgBox with custom button labels
' Return vID as result from MsgBox corresponding to clicked button (ByRef...Variant is compatible with any type)
' vbOK = 1, vbCancel = 2, vbAbort = 3, vbRetry = 4, vbIgnore = 5, vbYes = 6, vbNo = 7
' Arguments sPrompt, vButtons, vTitle, vHelpfile, and vContext match arguments of standard VBA MsgBox function
' This is Public Sub instead of Public Function so it will not be listed as a user-defined function (UDF)
'Sub Custom_MsgBox_Demo1()
'    MsgBoxCustom_Set vbOK, "Open"
'    MsgBoxCustom_Set vbCancel, "Close"
'    MsgBoxCustom ans, "Click a button.", vbOKCancel
'End Sub
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub MsgBoxCustom(ByRef vID As Variant, ByVal sPrompt As String, Optional ByVal vButtons As Variant = 0, Optional ByVal vTitle As Variant, _
    Optional ByVal vHelpfile As Variant, Optional ByVal vContext As Variant = 0)
    
    hHook = SetWindowsHookEx(WH_CBT, AddressOf MsgBoxCustom_Proc, 0, GetCurrentThreadId)
    
    If IsMissing(vHelpfile) And IsMissing(vTitle) Then
        vID = MsgBox(sPrompt, vButtons)
    ElseIf IsMissing(vHelpfile) Then
        vID = MsgBox(sPrompt, vButtons, vTitle)
    ElseIf IsMissing(vTitle) Then
        vID = MsgBox(sPrompt, vButtons, , vHelpfile, vContext)
    Else
        vID = MsgBox(sPrompt, vButtons, vTitle, vHelpfile, vContext)
    End If
    If hHook <> 0 Then UnhookWindowsHookEx hHook
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Init MessageBox with Custom Label
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub MsgBoxCustom_Init()
' Initialize default button labels for Public Sub MsgBoxCustom
    Dim nID As Integer
    Dim vA As Variant               ' base 0 array populated by Array function (must be Variant)
    vA = VBA.Array(vbNullString, "OK", "Cancel", "Abort", "Retry", "Ignore", "Yes", "No")
    For nID = 1 To 7
        sMsgBoxDefaultLabel(nID) = vA(nID)
        sMsgBoxCustomLabel(nID) = sMsgBoxDefaultLabel(nID)
    Next nID
    bMsgBoxCustomInit = True
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Reset Labels
' Reset button nID to default label for Public Sub MsgBoxCustom
' vbOK = 1, vbCancel = 2, vbAbort = 3, vbRetry = 4, vbIgnore = 5, vbYes = 6, vbNo = 7
' If nID is zero, all button labels will be set to default
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub MsgBoxCustom_Reset(ByVal nID As Integer)
    Call MsgBoxCustom_Set(nID)
End Sub

Private Sub RemovePropPointer()
    #If Win64 Then
        Dim lPtr As LongPtr
    #Else
        Dim lPtr As Long
    #End If
    
    lPtr = GetProp(hWndApplication, "ObjPtr")
    If lPtr <> 0 Then RemoveProp hWndApplication, "ObjPtr"
End Sub

#If VBA7 Then
    Private Function MsgBoxCustom_Proc(ByVal lMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
    Private Function MsgBoxCustom_Proc(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
' Hook callback function for Public Function MsgBoxCustom
    Dim nID As Integer
    If lMsg = HCBT_ACTIVATE And bMsgBoxCustomInit Then
        For nID = 1 To 7
            SetDlgItemText wParam, nID, sMsgBoxCustomLabel(nID)
        Next nID
    End If
    MsgBoxCustom_Proc = CallNextHookEx(hHook, lMsg, wParam, lParam)
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Convert Custom Labels Results
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function ConvertCustomResults(iSyle As VbMsgBoxStyle, iLbl As Long) As Long
Dim iRes As Long

    On Error Resume Next
'----------------
Select Case iSyle:
Case vbOKOnly:
                iRes = iLbl
Case vbOKCancel:
                iRes = iLbl
Case vbYesNoCancel:
            If iLbl = 2 Then
                iRes = 3
            ElseIf iLbl = 6 Then
                iRes = 1
            ElseIf iLbl = 7 Then
                iRes = 2
            End If
Case Else:
     iRes = -1
End Select
'----------------
ExitHere:
    ConvertCustomResults = iRes '!!!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Simple Wrapper for Yes/No Result
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function MsgYesNo(ByVal Instruction As String, ByVal Content As String, Optional defaultBtn As TDRESULT) As TDRESULT
    Dim TaskDialog                      As cTaskDialog
    Set TaskDialog = New cTaskDialog

    If defaultBtn = 0 Then
        defaultBtn = TD_NO
    End If

    With TaskDialog
        .Init
        .MainInstruction = Instruction
        .Content = Content
        .CommonButtons = TDCBF_YES_BUTTON Or TDCBF_NO_BUTTON
        .DefaultButton = defaultBtn
        .IconMain = IDI_QUESTION
        .ShowDialog
        MsgYesNo = .ResultMain
    End With
    Set TaskDialog = Nothing
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------
' Simple Wrapper for Information Only
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub MsgInfo(ByVal Instruction As String, ByVal Content As String, Optional ByVal mainIcon As TDICONS)
    Dim TaskDialog                      As cTaskDialog
    Set TaskDialog = New cTaskDialog

    With TaskDialog
        .Init
        .MainInstruction = Instruction
        .Content = Content
        .CommonButtons = TDCBF_OK_BUTTON
        .IconMain = mainIcon
        .ShowDialog
    End With
    Set TaskDialog = Nothing
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' COLOR DIALOG HELPERS

'------------------------------------------------------------------------------------------------------------------------------------------------------
' Convert color
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function Color_Hex_To_Long(strColor As String) As Long
    Dim iRed As Integer
    Dim iGreen As Integer
    Dim iBlue As Integer

    strColor = Replace(strColor, "#", "")
    strColor = Right("000000" & strColor, 6)
    iBlue = val("&H" & Mid(strColor, 1, 2))
    iGreen = val("&H" & Mid(strColor, 3, 2))
    iRed = val("&H" & Mid(strColor, 5, 2))

    Color_Hex_To_Long = RGB(iRed, iGreen, iBlue)
End Function

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' FILE/FOLDER DIALOG HELPERS

'------------------------------------------------------------------------------------------------------------------------------------------------------
' CALLBACK FOR FOLDER DIALOG
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function BrowseCallbackProc(hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
#If VBA7 Then
    Dim iHwnd As LongPtr
    iHwnd = CLngPtr(hWnd)
#Else
    Dim iHwnd As Long
    iHwnd = CLng(hWnd)
#End If
    If uMsg = BFFM_INITIALIZED Then
        SendMessage iHwnd, BFFM_SETSELECTION, 1, slRootFolder
    End If
    BrowseCallbackProc = 0
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Format Filter
' "All Files,*.*;Text Files,*.txt"
' "All Files (*.*); *.*"
' "Text Files (*.txt),*.txt,Add-In Files (*.xla),*.xla"
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function FormatFilter(sFilter As String, Optional DLM As String = vbNullChar) As String
Dim sWork As String

If sFilter = "" Then
   sWork = "All Files (*.*); *.*"
Else
   sWork = Replace(sFilter, "  ", " ")              ' Remove double space
End If
  
  sWork = Replace(sWork, vbCrLf, DLM)
  sWork = Replace(sWork, ",", DLM)
  sWork = Replace(sWork, ";", DLM)
  sWork = Replace(sWork, "|", DLM)
'------------------------------------
ExitHere:
  FormatFilter = sWork & DLM & DLM '!!!!!!!!!!!!!!!!!
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------
' Get Folder Dialog
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetFolderDialog(Optional ByVal hWnd As Variant = 0, Optional ByVal strRootFolder As String = "", _
            Optional ByVal strTitle As String = "") As String
Dim sBuffer As String
Dim lBrowseInfo As BrowseInfo
Dim lngRet As Long

#If VBA7 Then
    Dim iHwnd As LongPtr, ilpIDList As LongPtr
    iHwnd = CLngPtr(hWnd)
#Else
    Dim iHwnd As Long, ilpIDList As Long
    iHwnd = CLng(hWnd)
#End If

On Error GoTo ErrHandle
'--------------------------------------------------------------------
    With lBrowseInfo
        .hOwner = iHwnd
        .lpszTitle = strTitle
        .pIDLRoot = 0
        .ulFlags = BIF_returnonlyfsdirs
        .lParam = 0
    End With

    If strRootFolder <> "" Then
        slRootFolder = strRootFolder
        CopyMemory lBrowseInfo.lpfn, AddressOf BrowseCallbackProc, 4
    End If


    sBuffer = String$(MAX_PATH, vbNullChar)
    ilpIDList = SHBrowseForFolder(lBrowseInfo)

    If ilpIDList Then
        lngRet = SHGetPathFromIDList(ilpIDList, sBuffer)
        If lngRet Then
            GetFolderDialog = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        Else
            GetFolderDialog = ""
        End If
    End If
'-----------------------------
ExitHere:
    Exit Function
'--------------
ErrHandle:
    ErrPrint2 "GetFolderDialog", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------
' Get File Dialog
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function OpenFileDialog(Optional sInitDir As String, Optional sTitle As String = "Open file", Optional sFileNameOrMask As String, _
                Optional sPairFilter As String = "All Files (*.*); *.*") As String

Dim OpenFile    As OPENFILENAME, sRes As String
Dim lReturn     As Long, sFilesFilter As String, sInitialFolder As String
  
    On Error GoTo ErrHandle
'------------------------------------------
    sInitialFolder = sInitDir
    If sInitialFolder = "" Then sInitialFolder = CurrentProject.Path
      
    sFilesFilter = Replace(sPairFilter, "  ", " ")              ' Remove double space
    sFilesFilter = Replace(sFilesFilter, "; ", ";")             ' DLM with space
    sFilesFilter = Replace(sFilesFilter, ";", Chr$(0))          ' Lets GO

  
    With OpenFile
        .lpstrFilter = ""
        .nFilterIndex = 1
        
        
        .hWndOwner = 0
        .lpstrFile = String(257, 0)
    
        #If VBA7 Then
            .nMaxFile = LenB(OpenFile.lpstrFile) - 1
            .lStructSize = LenB(OpenFile)
        #Else
            .nMaxFile = Len(OpenFile.lpstrFile) - 1
            .lStructSize = Len(OpenFile)
        #End If
    
        .lpstrFileTitle = OpenFile.lpstrFile
        .nMaxFileTitle = OpenFile.nMaxFile
    
        .lpstrInitialDir = sInitialFolder
        .lpstrTitle = sTitle
    
        .lpstrFilter = FormatFilter(sFilesFilter)
        .nFilterIndex = 1 ' All Files (*.*) - set first template
         
         
         
        If sFileNameOrMask <> "" Then .lpstrFile = sFileNameOrMask & String$(512 - Len(sFileNameOrMask), 0)
        .FLAGS = 0
        
    End With
    
'------------------------------------------
    lReturn = GetOpenFileName(OpenFile)
    If lReturn <> 0 Then sRes = Trim(Left(OpenFile.lpstrFile, InStr(1, OpenFile.lpstrFile, vbNullChar) - 1))
'------------------------------------------
ExitHere:
    OpenFileDialog = sRes '!!!!!!!!!!!!!!!!!!!
    Exit Function
'-----------------
ErrHandle:
    ErrPrint2 "OpenFileDialog", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

