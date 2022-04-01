VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTaskDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'                                       ____    ____ .______        ___                                       '
'                                       \   \  /   / |   _  \      /   \                                      '
'                                        \   \/   /  |  |_)  |    /  ^  \                                     '
'                                         \      /   |   _  <    /  /_\  \                                    '
'                                          \    /    |  |_)  |  /  _____  \                                   '
'   .___________.     ___           _______.\__/ ___ |_______  /__/     \___       __        ______     _______
'   |           |    /   \         /       ||  |/  / |       \ |  |     /   \     |  |      /  __  \   /  _____|
'   `---|  |----`   /  ^  \       |   (----`|  '  /  |  .--.  ||  |    /  ^  \    |  |     |  |  |  | |  |  __
'       |  |       /  /_\  \       \   \    |    <   |  |  |  ||  |   /  /_\  \   |  |     |  |  |  | |  | |_ |
'       |  |      /  _____  \  .----)   |   |  .  \  |  '--'  ||  |  /  _____  \  |  `----.|  `--'  | |  |__| |
'       |__|     /__/     \__\ |_______/    |__|\__\ |_______/ |__| /__/     \__\ |_______| \______/   \______|
                                                 
Private Const vTdVersion As String = "20.10.15"
'***********************************************
'VBATaskDialog fork from v1.0 of VB6 code below
'by Kevin Bell (kevin@accessui.com)
'assistance from Rob Cooper and Shane Groff
'***********************************************

'You may integrate, modify and deploy this class in your own applications as long as the credits are left intact.

'***********************************************
'VERSION HISTORY
'20.10.15
'-- Initial public release for the Access Cascade Conference 10/17/2020

'*********************************
'cTaskDialog v1.0
'by Jon Johnson (aka fafalone)
'fafalone@gmail.com
'*********************************

'
'Provides a full implementation of TaskDialogIndirect and extensive enhancements.
'
'Code may be modified, reused, republished.. anything as long as credit is given
'
'Some enums and declarations originally written by Randy Birch at VBNet and Julius Laslo at vb@rchiv
'
'*********************************
'USAGE
'For full details on usage and support, see the development thread at VBForums:
'http://www.vbforums.com/showthread.php?777021
'The thread contains detailed samples complete with images for all features.
'Submit any bug reports in that thread as well.
'
'*********************************
'REQUIREMENTS
'--Common Controls 6.0 are required. Your application, and the VB IDE if running from there,
'  must include a manifest specifying 6.0. The cTaskDialog sample project includes one in its
'  resource file.
'  Your application must call InitCommonControlsEx from Sub Main and start with that.
'
'--mTaskDialogHelper contains 3 lines that must be included in any project that uses cTaskDialog.
'  You can include that module, or take the lines from in and place in a regular module (they
'  can't be in cTaskDialog.cls). Here's the 3 lines in case you don't have mTaskDialogHelper...
'  Public Function TaskDialogCallbackProc(ByVal hwnd As Long, ByVal uNotification As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal lpRefData As cTaskDialog) As Long: TaskDialogCallbackProc = lpRefData.ProcessCallback(hwnd, uNotification, wParam, lParam): End Function
'  Public Function TaskDialogEnumChildProc(ByVal hwnd As Long, ByVal lParam As cTaskDialog) As Long: TaskDialogEnumChildProc = lParam.ProcessEnumCallback(hwnd): End Function
'  Public Function TaskDialogSubclassProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As cTaskDialog) As Long: TaskDialogSubclassProc = dwRefData.ProcessSubclass(hwnd, uMsg, wParam, lParam, uIdSubclass): End Function
'
'--The TaskDialogIndirect API is only available on Windows Vista and newer.
'
'*********************************
'VERSION HISTORY
'
'NEW IN VERSION 1.0
'--You can now make a custom button into a Split Button (dropdown arrow). Use .SetSplitButton
'  with the ID of the custom button. This fires a DropdownButtonClicked event. Cannot be used
'  with Command Links. Note there's one more function prototype in mTaskDialogHelper you'll
'  need to copy if you're placing those in your own module (they've been condensed too so the
'  whole thing is only 3 lines now).
'--A logo type image can now be placed with .SetLogoImage. The current placement options are
'  the top right corner, or next to the buttons (only if no controls, either custom or the
'  expando/verify are present). The image is passed as an HBITMAP or HICON, so you can load
'  in whichever way you want. No copy is made, so you must destroy the image yourself, but
'  only after the dialog closes (or it will not appear).
'--Autoclose has been implemented. Set the .AutocloseTime property to a value in seconds; when
'  you get the property, it returns the current time remaining.
'--Events for custom controls: ComboItemChanged, ComboDropdown, InputBoxChange, DateTimeChange
'--Custom controls now have the option to manually set their width (and height for combo).
'--Custom controls in the Buttons position now have their width adjusted to fill the space.
'  Custom controls in the Footer position still use a standard size, however for InputBox,
'  ComboBox, and Slider if you specify a width of -1 the width will scale to the full width
'--For custom controls in the footer area, and the datetime control in the content area, there's
'  an additional alignment option to choose between left, center, and right. In the footer area
'  this could be used to show footer text and have the control off to the right.
'--Focus is no longer always on an InputBox: Focus is set on whatever custom control is in the
'  content area, if none button area, last footer area. Or, there's now a .DefaultCustomControl
'  option to manually specify which one, or not have focus set on any of them.
'--Bug fix: Several fixes relating to alignment and events for custom controls used on dialogs
'           that are loaded as a new page (NavigatePage/TDN_NAVIGATED)
'--Bug fix: .DefaultButton used TDBUTTONS instead of TDRESULT enum to identify IDs
'--Bug fix: .DefaultRadioButton was incorrectly associated with the TDBUTTONS enum
'--Alignment: -Adjusted Content area custom controls slightly higher when expanded
'              info is used.
'             -Custom controls in content area now have an appropriate X-position and width when
'              no icon appears, shifting everything to the left.
'             -Made sure that content text actually had a link before making the link adjustment,
'              since the flag would still be used if there's a link only in the footer area.
'             -When there's a link but no command links or radio buttons, content area controls
'              needed to be adjusted upwards a bit.
'--Alignment: -Custom controls size and position is now automatically adjust for the current DPI.
'             -You can also specify a scale factor manually with the DPIScaleX/DPIScaleY properties.
'KNOWN ISSUES: -If there's a custom control in the content area, when an expando control is expanded
'               and then collapsed, the excess whitespace isn't removed despite there not being any
'               extra vbCrLf's to account for it. No way to correct this has been discovered yet.
'              -If the dialog has to be resized in response to an expando control, the width is
'               reduced by several pixels for an unknown reason- but only on the first time. If the
'               numbers returned by getting the client RECT were wrong, you'd expect to lose more
'               pixels every time, but that's not the case. No solution has been discovered.
'              -Some alignment scenarios remain unsupported. If expando with the expand-footer w/
'               expand-default, have text with multiple lines, alignment will be off. This remains
'               a significant challenge to address as the DirectUIHWND reports incorrect information
'               to GetPixel, so figuring out where things are will require font height analysis and
'               reverse engineering line break length determination. Will work on in next version.
'              -Sliders with ticks on top and bottom simply don't fit on higher DPIs in the footer
'               position. Height available is too small no matter how it's positioned.
'----------------------------------------------------------------------------------------------------
'NEW IN VERSION 0.8
'--Added ability to use comboxes (normal edit dropdown and dropdown list), date/time pickers,
'  and sliders like the inputbox controls.
'--There's still the 3 alignment positions from the inputbox, so you can have controls in each
'  of these places (however, only 1 of each type is allowed)
'--When TDF_KILL_SHIELD_ICON is used, there's now an option to replace it with a different icon:
'   Set .IconReplaceGradient...
'    1) To an ID in your app's resource file (only works when compiled)
'    2) Add the TDF_USE_SHELL32_ICONID or TDF_USE_IMAGERES_ICONID flag and use an ID from those
'    3) Set .hInst to a custom module and use an icon ID from that
'   Using an hIcon is not supported at this time.
'--Bug fix: Using TDF_KILL_SHIELD_ICON and calling .Init crashed the app
'--Bug fix: Using an InputBox at the footer position, the bottom half was cut off if no
'           footer icon was set.
'--Bug fix: Input box position adjusted for expando only if in the content position, now adjusts
'           if in footer position too. Prevented this bug in all other controls.
'--Bug fix: If input box was in footer position, and expanded info was expanded automatically
'           and into the footer area, the input box would be in the wrong position. Fixed and
'           prevented for new controls.
'KNOWN ISSUE: If the slider control is placed in the footer position and an expando control is
'             set to expand into the footer area, the transparent background on the slider is
'             lost (turns white) until it's clicked. It's a very unusual and specific situation,
'             so not going to delay this version while I come up with a fix.
'TODO:
'The next version will be 1.0. There's a couple major features planned, including images and more
'events (for the custom controls). Maybe more controls; like dropdown menus on the buttons.
'I'm very aware this class needs substantial cleanup, that's also being saved for 1.0
' So please go easy on the sloppiness!!! :)
'
'---------------------------------------------------------------------------------------------------
'NEW IN VERSION 0.7
'--Input boxes! See details below for the TDF_INPUT_BOX custom flag.
'--Added flag TDF_KILL_SHIELD_ICON - kill it but keep that colored background.
'--Fixed bug where updating the footer icon in an open dialog instead updated the main icon
'NEW IN VERSION 0.6.1
'--Can't believe I forgot the TDF_SIZE_TO_CONTENTS flag.
'NEW IN VERSION 0.6
'--Support for custom icons on the buttons themselves
'0.5.2: critical bugfixes
'NEW IN VERSION 0.5.1
'--Added a force case function for the enums; they'll be automatically cased like normal now
'  while in the IDE
'--Bugfix: If a dialog was called as a new page, setting holds/disabled/elevated automatically
'          did not work.
'--Bugfix: When navigating to a new page, the old hWnd was not cleared, leading to some calls
'          not working properly.
'--Customized the progressbarstate enum since it's likely to conflict; if you use a PBST_ var,
'  it needs to be changed to ePBST, if it wasn't publicly declared outside this class
'NEW IN VERSION 0.5
'--Added support for multiple pages
'NEW IN VERSION 0.4.1
'--Added queueing for EnableButton, EnableRadioButton, and SetButtonElevated: now you can
'  disable or elevate a button before the dialog is shown and it will be done automatically
'NEW IN VERSION 0.4
'--Added support for TDM_SET_BUTTON_ELEVATION_REQUIRED_STATE; which adds the security shield
'  to a button indicated elevated user permission is required to perform the action
'--Added a SimpleDialog function to show the original TaskDialog() for very simple messages;
'  it can be a drop-in replacement for MsgBox, taking the common Message,Buttons,Title in
'  the same order
'--Set default dialog title to App.Title. If NULL is passed, it shows the exe name rather
'  than a blank string; and vb6.exe when in the IDE.
'--Misc minor bugfixes
'NEW IN VERSION 0.3.2/3
'Some internal improvements suggested by Bonnie West on VBForums. Also minor bugfixes.
'NEW IN VERSION 0.3.1
'Bugfixes; now use LoadLibrary instead of GetModuleHandle, which may fail in some circumstances
'NEW IN VERSION 0.3
'Automated the process of not closing the dialog when a button is clicked
'Uses .SetButtonHold and .ReleaseButtonHold
'NEW IN VERSION 0.2
'Schmidt gave me the brilliant idea to make using icons from shell32.dll and imageres.dll
'possible simply by setting .hInstance to their handle and specifying the index.
'See the new custom flags, and the readme for full details

'    _______   _______   ______  __           ___      .______           ___      .___________. __    ______   .__   __.      _______.
'   |       \ |   ____| /      ||  |         /   \     |   _  \         /   \     |           ||  |  /  __  \  |  \ |  |     /       |
'   |  .--.  ||  |__   |  ,----'|  |        /  ^  \    |  |_)  |       /  ^  \    `---|  |----`|  | |  |  |  | |   \|  |    |   (----`
'   |  |  |  ||   __|  |  |     |  |       /  /_\  \   |      /       /  /_\  \       |  |     |  | |  |  |  | |  . `  |     \   \
'   |  '--'  ||  |____ |  `----.|  `----. /  _____  \  |  |\  \----. /  _____  \      |  |     |  | |  `--'  | |  |\   | .----)   |
'   |_______/ |_______| \______||_______|/__/     \__\ | _| `._____|/__/     \__\     |__|     |__|  \______/  |__| \__| |_______/

            
'#If Win64 Then
    ' // TaskDialog APIs
    Private Declare PtrSafe Function TaskDialog Lib "comctl32.dll" ( _
            ByVal hwndParent As LongPtr, _
            ByVal hInstance As LongPtr, _
            ByVal pszWindowTitle As LongPtr, _
            ByVal pszMainInstruction As LongPtr, _
            ByVal pszContent As LongPtr, _
            ByVal dwCommonButtons As Long, _
            ByVal pszIcon As LongPtr, _
            pnButton As Long) As Long

    Private Declare PtrSafe Function TaskDialogIndirect Lib "comctl32.dll" ( _
            pTaskConfig As TASKDIALOGCONFIG, _
            pnButton As Long, _
            pnRadioButton As Long, _
            pfVerificationFlagChecked As Boolean) As Long
'#Else

'    Private Declare Function TaskDialog Lib "comctl32.dll" _
'                                        (ByVal hwndParent As Long, _
'                                         ByVal hInstance As Long, _
'                                         ByVal pszWindowTitle As Long, _
'                                         ByVal pszMainInstruction As Long, _
'                                         ByVal pszContent As Long, _
'                                         ByVal dwCommonButtons As Long, _
'                                         ByVal pszIcon As Long, _
'                                         pnButton As Long) As Long
'    'WINCOMMCTRLAPI HRESULT WINAPI TaskDialogIndirect(const TASKDIALOGCONFIG *pTaskConfig, __out_opt int *pnButton, __out_opt int *pnRadioButton, __out_opt BOOL *pfVerificationFlagChecked);
'    Private Declare Function TaskDialogIndirect Lib "comctl32.dll" (pTaskConfig As TASKDIALOGCONFIG, _
'                                                                    pnButton As Long, _
'                                                                    pnRadioButton As Long, _
'                                                                    pfVerificationFlagChecked As Boolean) As Long
'#End If

Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
Private Declare PtrSafe Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, lpParam As Any) As LongPtr
Private Declare PtrSafe Function DefSubclassProc Lib "comctl32" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function DestroyIcon Lib "user32.dll" (ByVal hIcon As LongPtr) As Long
Private Declare PtrSafe Function EnumChildWindows Lib "user32" (ByVal hwndParent As LongPtr, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As LongPtr, ByVal nCount As Long, lpObject As Any) As Long
Private Declare PtrSafe Function GetParent Lib "user32.dll" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetPixel Lib "gdi32" (ByVal hdc As LongPtr, ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As GWL_nIndex) As LongPtr
Private Declare PtrSafe Function GetWindowRect Lib "user32.dll" (ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function InvalidateRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As Any, ByVal bErase As Long) As Long
Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
Private Declare PtrSafe Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As LongPtr, ByVal lpsz As Long, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As LongPtr
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function RemoveWindowSubclass Lib "comctl32" (ByVal hWnd As LongPtr, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare PtrSafe Function ScreenToClient Lib "user32" (ByVal hWnd As LongPtr, ByRef lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function SendMessageW Lib "user32" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Private Declare PtrSafe Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As GWL_nIndex, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function SetWindowSubclass Lib "comctl32" (ByVal hWnd As LongPtr, ByVal pfnSubclass As LongPtr, ByVal uIdSubclass As LongPtr, ByVal dwRefData As LongPtr) As Long
Private Declare PtrSafe Function ShellExecuteW Lib "shell32.dll" (ByVal hWnd As LongPtr, ByVal lpOperation As LongPtr, ByVal lpFile As LongPtr, ByVal lpParameters As LongPtr, ByVal lpDirectory As LongPtr, ByVal nShowCmd As Long) As LongPtr
Private Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'     ______   ______   .__   __.      _______..___________.     ___      .__   __. .___________.     _______.
'    /      | /  __  \  |  \ |  |     /       ||           |    /   \     |  \ |  | |           |    /       |
'   |  ,----'|  |  |  | |   \|  |    |   (----``---|  |----`   /  ^  \    |   \|  | `---|  |----`   |   (----`
'   |  |     |  |  |  | |  . `  |     \   \        |  |       /  /_\  \   |  . `  |     |  |         \   \
'   |  `----.|  `--'  | |  |\   | .----)   |       |  |      /  _____  \  |  |\   |     |  |     .----)   |
'    \______| \______/  |__| \__| |_______/        |__|     /__/     \__\ |__| \__|     |__|     |_______/

Private Const WM_USER                   As Long = &H400&
Private Const S_OK                      As Long = 0
Private Const S_FALSE                   As Long = 1&
Private Const CSIDL_SYSTEM              As Long = &H25
Private Const IDD_EDIT1                 As Long = 101&
Private Const IDD_DT1                   As Long = 102&
Private Const IDD_DT2                   As Long = 103&
Private Const IDD_COMBO                 As Long = 104&
Private Const IDD_SLIDER                As Long = 105&

Private Const WS_EX_CLIENTEDGE          As Long = &H200
Private Const WS_EX_LEFT                As Long = &H0&
Private Const WS_EX_LTRREADING          As Long = &H0&
Private Const WS_EX_RIGHTSCROLLBAR      As Long = &H0&
Private Const WS_EX_TRANSPARENT         As Long = &H20
Private Const WS_EX_LAYERED             As Long = &H80000

Private Const DEFAULT_GUI_FONT          As Long = 17
Private Const WM_SETFONT                As Long = &H30
Private Const WM_GETTEXTLENGTH          As Long = &HE
Private Const WM_GETTEXT                As Long = &HD
Private Const WM_SETTEXT                As Long = &HC
Private Const WM_COMMAND                As Long = &H111
Private Const WM_NOTIFY                 As Long = &H4E
Private Const WM_HSCROLL                As Long = &H114
Private Const WM_DESTROY                As Long = &H2
Private Const WM_CTLCOLORSTATIC         As Long = &H138
Private Const BM_SETIMAGE               As Long = &HF7
Private Const EM_SETSEL                 As Long = &HB1
Private Const ECM_FIRST                 As Long = &H1500
Private Const EM_SETCUEBANNER           As Long = (ECM_FIRST + 1)
Private Const ES_LEFT                   As Long = &H0
Private Const ES_PASSWORD               As Long = &H20
Private Const BS_SPLITBUTTON            As Long = &HC
Private Const CBN_SELCHANGE             As Long = 1
Private Const CBN_DROPDOWN              As Long = 7
Private Const EN_CHANGE                 As Long = &H300
Private Const EN_UPDATE                 As Long = &H400
Private Const BCN_FIRST                 As Long = -1250
Private Const BCN_DROPDOWN              As Long = (BCN_FIRST + 2)
Private Const H_MAX                     As Long = &HFFFF + 1
Private Const DTN_FIRST                 As Long = (H_MAX - 760&)
Private Const DTN_LAST                  As Long = (H_MAX - 799&)
Private Const DTN_DATETIMECHANGE        As Long = (DTN_FIRST + 1)
Private Const TRBN_FIRST                As Long = -1501
Private Const TRBN_THUMBPOSCHANGING     As Long = TRBN_FIRST - 1
Private Const TBS_HORZ                  As Long = &H0
Private Const TBS_BOTH                  As Long = &H8
Private Const TBS_BOTTOM                As Long = &H0
Private Const TBS_TOP                   As Long = &H4
Private Const TBS_NOTICKS               As Long = &H10
Private Const TBS_AUTOTICKS             As Long = &H1
Private Const TBS_TRANSPARENTBKGND      As Long = &H1000
Private Const TBM_SETRANGEMIN           As Long = (WM_USER + 7)
Private Const TBM_SETRANGEMAX           As Long = (WM_USER + 8)
Private Const TBM_SETTICFREQ            As Long = (WM_USER + 20)
Private Const TBM_SETLINESIZE           As Long = (WM_USER + 23)
Private Const TBM_SETPAGESIZE           As Long = (WM_USER + 21)
Private Const TBM_GETPOS                As Long = (WM_USER)
Private Const TBM_SETPOS                As Long = (WM_USER + 5)
Private Const BCM_FIRST                 As Long = &H1600
Private Const BCM_SETDROPDOWNSTATE      As Long = (BCM_FIRST + &H6)
Private Const DTM_FIRST                 As Long = &H1000
Private Const DTM_GETSYSTEMTIME         As Long = (DTM_FIRST + 1)
Private Const DTM_SETSYSTEMTIME         As Long = (DTM_FIRST + 2)
Private Const DTM_SETRANGE              As Long = (DTM_FIRST + 4)
Private Const STM_SETICON               As Long = &H170
Private Const STM_GETICON               As Long = &H171
Private Const STM_SETIMAGE              As Long = &H172
Private Const STM_GETIMAGE              As Long = &H173
Private Const SS_LEFT                   As Long = &H0
Private Const SS_CENTER                 As Long = &H1
Private Const SS_RIGHT                  As Long = &H2
Private Const SS_ICON                   As Long = &H3
Private Const SS_BLACKRECT              As Long = &H4
Private Const SS_GRAYRECT               As Long = &H5
Private Const SS_WHITERECT              As Long = &H6
Private Const SS_BLACKFRAME             As Long = &H7
Private Const SS_GRAYFRAME              As Long = &H8
Private Const SS_WHITEFRAME             As Long = &H9
Private Const SS_SIMPLE                 As Long = &HB
Private Const SS_LEFTNOWORDWRAP         As Long = &HC
Private Const SS_BITMAP                 As Long = &HE
Private Const SS_ENHMETAFILE            As Long = &HF
Private Const SS_ETCHEDHORZ             As Long = &H10
Private Const SS_ETCHEDVERT             As Long = &H11
Private Const SS_ETCHEDFRAME            As Long = &H12
Private Const SS_REALSIZECONTROL        As Long = &H40
Private Const SS_NOPREFIX               As Long = &H80
Private Const SS_NOTIFY                 As Long = &H100
Private Const SS_CENTERIMAGE            As Long = &H200
Private Const SS_RIGHTJUST              As Long = &H400
Private Const SS_SUNKEN                 As Long = &H1000

Private Const GDT_ERROR                 As Long = -1
Private Const GDT_VALID                 As Long = 0
Private Const GDT_NONE                  As Long = 1
Private Const GDTR_MAX                  As Long = &H2
Private Const GDTR_MIN                  As Long = &H1

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_FRAMECHANGED = &H20

Private Const WC_COMBOBOXEX = "ComboBoxEx32"
Private Const CBEM_SETIMAGELIST = (WM_USER + 2)
Private Const CBEM_GETCOMBOCONTROL = (WM_USER + 6)
Private Const CBEM_GETEDITCONTROL = (WM_USER + 7)
Private Const CBEM_SETUNICODEFORMAT = 8192 + 5
Private Const CBEM_INSERTITEMW = (WM_USER + 11)
Private Const CBEM_SETITEMW = (WM_USER + 12)
Private Const CBS_DROPDOWN = &H2&
Private Const CBS_DROPDOWNLIST = &H3&
Private Const CBS_AUTOHSCROLL = &H40
Private Const CB_GETCURSEL = &H147
Private Const CB_SETCURSEL = &H14E


Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90


'   .___________.____    ____ .______    _______      _______.
'   |           |\   \  /   / |   _  \  |   ____|    /       |
'   `---|  |----` \   \/   /  |  |_)  | |  |__      |   (----`
'       |  |       \_    _/   |   ___/  |   __|      \   \
'       |  |         |  |     |  |      |  |____ .----)   |
'       |__|         |__|     | _|      |_______||_______/

#If Win64 Then
Private Type TASKDIALOG_BUTTON
    nButtonID                           As Long
    pszButtonText                       As Long
    pszButtonText2                      As Long
End Type

Private Type TASKDIALOGCONFIG
    cbSize                              As Long
    hwndParent                          As Long
    hwndParent2                         As Long
    hInstance                           As Long
    hInstance2                          As Long
    dwFlags                             As TASKDIALOG_FLAGS
    dwCommonButtons                     As TDBUTTONS
    pszWindowTitle                      As Long
    pszWindowTitle2                     As Long
    pszMainIcon                         As Long
    pszMainIcon2                        As Long
    pszMainInstruction                  As Long
    pszMainInstruction2                 As Long
    pszContent                          As Long
    pszContent2                         As Long
    cButtons                            As Long
    pButtons                            As Long
    pButtons2                           As Long
    nDefaultButton                      As Long
    cRadioButtons                       As Long
    pRadioButtons                       As Long
    pRadioButtons2                      As Long
    nDefaultRadioButton                 As Long
    pszVerificationText                 As Long
    pszVerificationText2                As Long
    pszExpandedInformation              As Long
    pszExpandedInformation2             As Long
    pszExpandedControlText              As Long
    pszExpandedControlText2             As Long
    pszCollapsedControlText             As Long
    pszCollapsedControlText2            As Long
    pszFooterIcon                       As Long
    pszFooterIcon2                      As Long
    pszFooter                           As Long
    pszFooter2                          As Long
    pfCallback                          As Long
    pfCallback2                         As Long
    lpCallbackData                      As Long
    lpCallbackData2                     As Long
    cxWidth                             As Long
End Type

Private Type COMBOBOXEXITEMW
    Mask                                As COMBOBOXEXITEM_Mask
    iItem                               As LongPtr
    pszText                             As Long    '// LPCSTR
    pszText2                            As Long    '// LPCSTR
    cchTextMax                          As Long
    iImage                              As Long
    iSelectedImage                      As Long
    iOverlay                            As Long
    iIndent                             As Long
    lParam                              As LongPtr
End Type
#Else

Private Type TASKDIALOG_BUTTON
    nButtonID                           As Long
    pszButtonText                       As Long
End Type

Private Type TASKDIALOGCONFIG
    cbSize                              As Long
    hwndParent                          As Long
    hInstance                           As Long
    dwFlags                             As TASKDIALOG_FLAGS
    dwCommonButtons                     As TDBUTTONS
    pszWindowTitle                      As Long
    'hMainIcon As Long
    pszMainIcon                         As TDICONS
    pszMainInstruction                  As Long
    pszContent                          As Long
    cButtons                            As Long
    pButtons                            As Long
    nDefaultButton                      As Long
    cRadioButtons                       As Long
    pRadioButtons                       As Long
    nDefaultRadioButton                 As Long
    pszVerificationText                 As Long
    pszExpandedInformation              As Long
    pszExpandedControlText              As Long
    pszCollapsedControlText             As Long
    pszFooterIcon                       As TDICONS
    'hFooterIcon As Long
    pszFooter                           As Long
    pfCallback                          As Long
    lpCallbackData                      As Long
    cxWidth                             As Long
End Type

Private Type COMBOBOXEXITEMW
    Mask                                As COMBOBOXEXITEM_Mask
    iItem                               As Long
    pszText                             As Long    '// LPCSTR
    cchTextMax                          As Long
    iImage                              As Long
    iSelectedImage                      As Long
    iOverlay                            As Long
    iIndent                             As Long
    lParam                              As LongPtr
End Type
#End If

Private Type RECT
    Left                                As Long
    Top                                 As Long
    Right                               As Long
    Bottom                              As Long
End Type

Private Type POINTAPI
    x                                   As Long
    y                                   As Long
End Type

Private Type NMHDR
    hWndFrom                            As Long   ' Window handle of control sending message
    IDFrom                              As Long        ' Identifier of control sending message
    code                                As Long          ' Specifies the notification code
End Type

Private Type BITMAP
    BMType                              As Long
    BMWidth                             As Long
    BMHeight                            As Long
    BMWidthBytes                        As Long
    BMPlanes                            As Integer
    BMBitsPixel                         As Integer
    BMBits                              As Long
End Type

Private Type SYSTEMTIME
    wYear                               As Integer
    wMonth                              As Integer
    wDayOfWeek                          As Integer
    wDay                                As Integer
    wHour                               As Integer
    wMinute                             As Integer
    wSecond                             As Integer
    wMilliseconds                       As Integer
End Type

Private Type ctdComboItem
    sText                               As String
    iImage                              As Long
    iOverlay                            As Long
End Type


'    _______ .__   __.  __    __  .___  ___.      _______.
'   |   ____||  \ |  | |  |  |  | |   \/   |     /       |
'   |  |__   |   \|  | |  |  |  | |  \  /  |    |   (----`
'   |   __|  |  . `  | |  |  |  | |  |\/|  |     \   \
'   |  |____ |  |\   | |  `--'  | |  |  |  | .----)   |
'   |_______||__| \__|  \______/  |__|  |__| |_______/

Private Enum ImageTypes
    IMAGE_BITMAP = 0
    IMAGE_ICON = 1
    IMAGE_CURSOR = 2
    IMAGE_ENHMETAFILE = 3
End Enum

Private Enum WinStyles
    WS_OVERLAPPED = &H0
    WS_TABSTOP = &H10000
    WS_MAXIMIZEBOX = &H10000
    WS_MINIMIZEBOX = &H20000
    WS_GROUP = &H20000
    WS_THICKFRAME = &H40000
    WS_SYSMENU = &H80000
    WS_HSCROLL = &H100000
    WS_VSCROLL = &H200000
    WS_DLGFRAME = &H400000
    WS_BORDER = &H800000
    WS_CAPTION = (WS_BORDER Or WS_DLGFRAME)
    WS_MAXIMIZE = &H1000000
    WS_CLIPCHILDREN = &H2000000
    WS_CLIPSIBLINGS = &H4000000
    WS_DISABLED = &H8000000
    WS_VISIBLE = &H10000000
    WS_MINIMIZE = &H20000000
    WS_CHILD = &H40000000
    WS_POPUP = &H80000000

    WS_TILED = WS_OVERLAPPED
    WS_ICONIC = WS_MINIMIZE
    WS_SIZEBOX = WS_THICKFRAME

    ' Common Window Styles
    WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
    WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
    WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
    WS_CHILDWINDOW = WS_CHILD
End Enum

Private Enum DTSTYLES
    DTS_SHORTDATEFORMAT = &H0
    DTS_UPDOWN = &H1
    DTS_SHOWNONE = &H2
    DTS_LONGDATEFORMAT = &H4
    DTS_TIMEFORMAT = &H9
    DTS_APPCANPARSE = &H10
    DTS_RIGHTALIGN = &H20
    DTS_SHORTDATECENTURYFORMAT = &HC
End Enum

Public Enum SldTickStyle
    SldTickStyleBottom = 0
    SldTickStyleTop = 1
    SldTickStyleBoth = 2
    SldTickStyleNone = 3
End Enum

Public Enum eProgressBarStates
    ePBST_NORMAL = 1
    ePBST_ERROR = 2
    ePBST_PAUSED = 3
End Enum

Public Enum TDInputBoxAlign
    TDIBA_Content = 0
    TDIBA_Buttons = 1
    TDIBA_Footer = 2
End Enum

Public Enum eShowWindowTypes
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_NORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_MAXIMIZE = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
End Enum

Private Enum COMBOBOXEXITEM_Mask
    CBEIF_TEXT = &H1
    CBEIF_IMAGE = &H2
    CBEIF_SELECTEDIMAGE = &H4
    CBEIF_OVERLAY = &H8
    CBEIF_INDENT = &H10
    CBEIF_LPARAM = &H20
    CBEIF_DI_SETITEM = &H10000000
End Enum

Public Enum TDICONS
    TD_WARNING_ICON = -1    'exclamation point in a yellow 'yield' triangle (same image as IDI_EXCLAMATION)
    TD_ERROR_ICON = -2    'round red circle containg 'X' (same as IDI_HAND)
    TD_INFORMATION_ICON = -3    'round blue circle containing 'i' (same image as IDI_ASTERISK)
    TD_SHIELD_ICON = -4    'Vista's security shield
    IDI_APPLICATION = 32512&    'miniature picture of an application window
    IDI_ERROR = 32513&
    IDI_QUESTION = 32514&    'round blue circle containing '?'
    IDI_WINLOGO = 32517&
    TD_SHIELD_GRADIENT_ICON = -5    'same image as TD_SHIELD_ICON; main message text on gradient blue background
    TD_SHIELD_WARNING_ICON = -6    'exclamation point in yellow Shield shape; main message text on gradient orange background
    TD_SHIELD_ERROR_ICON = -7    'X contained within Shield shape; main message text on gradient red background
    TD_SHIELD_OK_ICON = -8    'Shield shape containing green checkmark; main message text on gradient green background
    TD_SHIELD_GRAY_ICON = -9    'same image as TD_SHIELD_ICON; main message text on medium gray background
    TD_NO_ICON = 0    'no icon; text on white background
End Enum

'taskdialog common button flags
Public Enum TDBUTTONS
    TDCBF_OK_BUTTON = &H1&      'return value 1 (IDOK)
    TDCBF_YES_BUTTON = &H2&     'return value 6 (IDYES)
    TDCBF_NO_BUTTON = &H4&      'return value 7 (IDNO)
    TDCBF_CANCEL_BUTTON = &H8&  'return value 2 (IDCANCEL)
    TDCBF_RETRY_BUTTON = &H10&   'return value 4 (IDRETRY)
    TDCBF_CLOSE_BUTTON = &H20&   'return value 8 (IDCLOSE)

    TDCBF_ABORT_BUTTON = 65536
    TDCBF_IGNORE_BUTTON = 131072
    TDCBF_TRYAGAIN_BUTTON = 262144
    TDCBF_CONTINUE_BUTTON = 524288
    '// Note: Clicking the "Help" button will not close the dialog, but will
    '// raise the TaskDialogPage.Help event.
    'Crashes Access
    TDCBF_HELP_BUTTON = 16384    '104857
End Enum

Public Enum TDRESULT
    TD_OK = 1
    TD_YES = 6
    TD_NO = 7
    TD_CANCEL = 2
    TD_RETRY = 4
    TD_CLOSE = 8
    TD_ABORT = 3
    TD_IGNORE = 5
    TD_TRYAGAIN = 10
    TD_CONTINUE = 11
End Enum

Public Enum TASKDIALOG_FLAGS
    TDF_ENABLE_HYPERLINKS = &H1
    TDF_USE_HICON_MAIN = &H2
    TDF_USE_HICON_FOOTER = &H4
    TDF_ALLOW_DIALOG_CANCELLATION = &H8
    TDF_USE_COMMAND_LINKS = &H10
    TDF_USE_COMMAND_LINKS_NO_ICON = &H20
    TDF_EXPAND_FOOTER_AREA = &H40
    TDF_EXPANDED_BY_DEFAULT = &H80
    TDF_VERIFICATION_FLAG_CHECKED = &H100
    TDF_SHOW_PROGRESS_BAR = &H200
    TDF_SHOW_MARQUEE_PROGRESS_BAR = &H400
    TDF_CALLBACK_TIMER = &H800
    TDF_POSITION_RELATIVE_TO_WINDOW = &H1000
    TDF_RTL_LAYOUT = &H2000
    TDF_NO_DEFAULT_RADIO_BUTTON = &H4000
    TDF_CAN_BE_MINIMIZED = &H8000&
    TDF_NO_SET_FOREGROUND = &H10000
    TDF_SIZE_TO_CONTENTS = &H1000000
    '    TDF_USE_SHELL32_ICONID = &H10000    'CUSTOM FLAG
    '    TDF_USE_IMAGERES_ICONID = &H20000    'CUSTOM FLAG
    '    TDF_EXEC_HYPERLINKS = &H40000    'CUSTOM FLAG
    '    TDF_USE_SHELL32_ICONID_BUTTON = &H80000    'CUSTOM FLAG
    '    TDF_USE_IMAGERES_ICONID_BUTTON = &H100000    'CUSTOM FLAG
    '    TDF_KILL_SHIELD_ICON = &H200000    'CUSTOM FLAG: Allows the colored background from the TD_SHIELD_x icons to be used with just the text and no shield icon
    '    TDF_INPUT_BOX = &H400000        'CUSTOM FLAG: Input box. See below.
    '    TDF_COMBO_BOX = &H800000        'CUSTOM FLAG: Combo box; .ComboType controls edit or list
    '    TDF_DATETIME = &H1000000        'CUSTOM FLAG: Add calendar control
    '    TDF_SLIDER = &H2000000          'CUSTOM FLAG: Add Slider control
    TDF_USE_SHELL32_ICONID = &H400000    'CUSTOM FLAG
    TDF_USE_IMAGERES_ICONID = &H800000    'CUSTOM FLAG
    TDF_DATETIME = &H2000000         'CUSTOM FLAG: Add calendar control
    TDF_SLIDER = &H2000000          'CUSTOM FLAG: Add Slider control
    TDF_INPUT_BOX = &H4000000        'CUSTOM FLAG: Input box. See below.
    TDF_COMBO_BOX = &H8000000               'CUSTOM FLAG: Combo box; .ComboType controls edit or list
    TDF_KILL_SHIELD_ICON = &H10000000    'CUSTOM FLAG: Allows the colored background from the TD_SHIELD_x icons to be used with just the text and no shield icon
    TDF_EXEC_HYPERLINKS = &H20000000           'CUSTOM FLAG
    TDF_USE_SHELL32_ICONID_BUTTON = &H40000000    'CUSTOM FLAG
    TDF_USE_IMAGERES_ICONID_BUTTON = &H80000000    'CUSTOM FLAG
End Enum


'TDF_INPUT_BOX:
'Creates a textbox for input. There are multiple ways to set its position:
'1) Default, at the bottom of the content in the whitespace before the gray area
'2) To the left of the buttons. Not compatible with the expanded-info or verify-checkbox styles.
'3) As a footer. Aligned to retain the footer icon, but have an inputbox next to it.
'
'Additional inputbox items:
'-Initial text can be set, and current text from an open dialog, with .InputText
'-Cue banner can be set, .InputCueBanner
'-Input box can be set as a password input with * masks, .InputIsPassword


Public Enum TASKDIALOG_MESSAGES
    TDM_NAVIGATE_PAGE = WM_USER + 101&
    TDM_CLICK_BUTTON = WM_USER + 102&                  '// wParam = Button ID
    TDM_SET_MARQUEE_PROGRESS_BAR = WM_USER + 103&      '// wParam = 0 (nonMarque) wParam != 0 (Marquee)
    TDM_SET_PROGRESS_BAR_STATE = WM_USER + 104&        '// wParam = new progress state
    TDM_SET_PROGRESS_BAR_RANGE = WM_USER + 105&        '// lParam = tdMAKELPARAM(nMinRange& nMaxRange)
    TDM_SET_PROGRESS_BAR_POS = WM_USER + 106&          '// wParam = new position
    TDM_SET_PROGRESS_BAR_MARQUEE = WM_USER + 107&      '// wParam = 0 (stop marquee), wParam != 0 (start marquee), lparam = speed (milliseconds between repaints)
    TDM_SET_ELEMENT_TEXT = WM_USER + 108&              '// wParam = element (TASKDIALOG_ELEMENTS), lParam = new element text (LPCWSTR)
    TDM_CLICK_RADIO_BUTTON = WM_USER + 110&            '// wParam = Radio Button ID
    TDM_ENABLE_BUTTON = WM_USER + 111&                 '// lParam = 0 (disable), lParam != 0 (enable), wParam = Button ID
    TDM_ENABLE_RADIO_BUTTON = WM_USER + 112&           '// lParam = 0 (disable), lParam != 0 (enable), wParam = Radio Button ID
    TDM_CLICK_VERIFICATION = WM_USER + 113&            '// wParam = 0 (unchecked), 1 (checked), lParam = 1 (set key focus)
    TDM_UPDATE_ELEMENT_TEXT = WM_USER + 114&           '// wParam = element (TASKDIALOG_ELEMENTS), lParam = new element text (LPCWSTR)
    TDM_SET_BUTTON_ELEVATION_REQUIRED_STATE = WM_USER + 115&    '// wParam = Button ID, lParam = 0 (elevation not required)& lParam != 0 (elevation required)
    TDM_UPDATE_ICON = WM_USER + 116&                   '// wParam = icon element (TASKDIALOG_ICON_ELEMENTS), lParam = new icon (hIcon if TDF_USE_HICON_* was set, PCWSTR otherwise)
End Enum    'TASKDIALOG_MESSAGES;

Public Enum TASKDIALOG_NOTIFICATIONS
    TDN_CREATED = 0
    TDN_NAVIGATED = 1
    TDN_BUTTON_CLICKED = 2           '// wParam = Button ID
    TDN_HYPERLINK_CLICKED = 3        '// lParam = (LPCWSTR)pszHREF
    TDN_TIMER = 4                    '// wParam = Milliseconds since dialog created or timer reset
    TDN_DESTROYED = 5
    TDN_RADIO_BUTTON_CLICKED = 6     '// wParam = Radio Button ID
    TDN_DIALOG_CONSTRUCTED = 7
    TDN_VERIFICATION_CLICKED = 8     '// wParam = 1 if checkbox checked, 0 if not, lParam is unused and always 0
    TDN_HELP = 9
    TDN_EXPANDO_BUTTON_CLICKED = 10  '// wParam = 0 (dialog is now collapsed), wParam != 0 (dialog is now expanded)
End Enum    'TASKDIALOG_NOTIFICATIONS;

Public Enum TASKDIALOG_ELEMENTS
    TDE_CONTENT = 0
    TDE_EXPANDED_INFORMATION = 1
    TDE_FOOTER = 2
    TDE_MAIN_INSTRUCTION = 3
End Enum

Public Enum TASKDIALOG_ICON_ELEMENTS
    TDIE_ICON_MAIN = 0
    TDIE_ICON_FOOTER = 1
End Enum

Private Enum GWL_nIndex
    GWL_WNDPROC = (-4)
    GWL_HINSTANCE = (-6)
    GWL_HWNDPARENT = (-8)
    GWL_ID = (-12)
    GWL_STYLE = (-16)
    GWL_EXSTYLE = (-20)
    GWL_USERDATA = (-21)
End Enum

Public Enum TDCustomControlType
    tdccInputBox = 1
    tdccComboBox = 2
    tdccDateTime = 3
    tdccSlider = 4
    tdccNothing = 99
End Enum


Public Enum TDControlAlign
    tdcaLeft = 0
    tdcaCenter = 1
    tdcaRight = 2
End Enum

Public Enum LogoImageType
    LogoBitmap = 0
    LogoIcon = 1
End Enum

Public Enum LogoPosition
    LogoTopRight = 0
    LogoButtons = 1
End Enum

Public Enum DateTimeType
    dttDate = 0
    dttDateWithCheck = 1
    dttTime = 2
    dttTimeWithCheck = 3
    dttDateTime = 4
    dttDateTimeWithCheck = 5
    dttDateTimeWithCheckTimeOnly = 6
    dttDateTimeWithCheckDateOnly = 7
End Enum

Public Enum ComboType
    cbtDropdownEdit = 0
    cbtDropdownList = 1
End Enum

Public Enum eWinIcoSize
    ICO_16 = 16
    ICO_24 = 24
    ICO_32 = 32
    ICO_48 = 48
    ICO_64 = 64
End Enum


' _______ ____    ____  _______ .__   __. .___________.     _______.
'|   ____|\   \  /   / |   ____||  \ |  | |           |    /       |
'|  |__    \   \/   /  |  |__   |   \|  | `---|  |----`   |   (----`
'|   __|    \      /   |   __|  |  . `  |     |  |         \   \
'|  |____    \    /    |  |____ |  |\   |     |  |     .----)   |
'|_______|    \__/     |_______||__| \__|     |__|     |_______/

Public Event DialogCreated(ByVal hWnd As Long)
Public Event ButtonClick(ByVal ButtonID As Long)
Public Event HyperlinkClick(ByVal lPtr As LongPtr)
Public Event Timer(ByVal TimerValue As Long)
Public Event DialogDestroyed()
Public Event RadioButtonClick(ByVal ButtonID As Long)
Public Event DialogConstucted(ByVal hWnd As Long)
Public Event VerificationClicked(ByVal value As Long)
Public Event ExpandButtonClicked(ByVal value As Long)
Public Event Navigated()
Public Event Help()
'----------------------
'Custom Events:
Public Event DropdownButtonClicked(ByVal hWnd As LongPtr)
Public Event ComboItemChanged(ByVal iNewItem As LongPtr)
Public Event ComboDropdown()
Public Event InputBoxChange(sText As String)
Public Event DateTimeChange(ByVal dtNew As Date, ByVal lCheckStatus As Long)
Public Event SliderChange(ByVal lNewValue As LongPtr)
Public Event AutoClose()


' ____    ____      ___      .______       __       ___      .______    __       _______      _______.
' \   \  /   /     /   \     |   _  \     |  |     /   \     |   _  \  |  |     |   ____|    /       |
'  \   \/   /     /  ^  \    |  |_)  |    |  |    /  ^  \    |  |_)  | |  |     |  |__      |   (----`
'   \      /     /  /_\  \   |      /     |  |   /  /_\  \   |   _  <  |  |     |   __|      \   \
'    \    /     /  _____  \  |  |\  \----.|  |  /  _____  \  |  |_)  | |  `----.|  |____ .----)   |
'     \__/     /__/     \__\ | _| `._____||__| /__/     \__\ |______/  |_______||_______||_______/
                                                                                                    
Private nSliderTickStyle                As SldTickStyle

Private m_ScaleX                        As Single
Private m_ScaleY                        As Single
Public lngTimerID                       As Long

Private lpAdress                        As Long
Private bvASM(40)                       As Byte
Private m_TDhWnd                        As LongPtr
Private bResetTimer                     As Boolean
Private lBtnHold()                      As Long
Private lBtnElv()                       As Long
Private bElv                            As Boolean
Private lBtnDis()                       As Long
Private bDis                            As Boolean
Private lRadDis()                       As Long
Private bRadDis                         As Boolean
Private btnCaps()                       As String    'weird things happen if these are set but not stored
Private rbtnCaps()                      As String
Private icoDataC()                      As Long    'icon data for custom buttons
Private icoDataN(5)                     As Long    'icon data for normal buttons
Private bAutoclose                      As Boolean
Private lAutoclose                      As Long    'in seconds
Private lAutocloseTrack                 As Long
Private dtAC1 As Date, dtAC2            As Date
Private bACTmrSup                       As Boolean
Private hEditBox                        As LongPtr
Private hEditCombo                      As LongPtr
Private hComboCombo                     As LongPtr
Private hStatic                         As LongPtr
Private hLogoImage                      As LongPtr
Private nLogoType                       As LogoImageType
Private nLogoPos                        As LogoPosition
Private lLogoX As Long, lLogoY          As Long
Private hSlider                         As LongPtr
Private lSliderMin                      As Long
Private lSliderMax                      As Long
Private lSliderValue                    As LongPtr
Private lSliderNTicks                   As Long
Private lSliderChangeS                  As Long
Private lSliderChangeL                  As Long
Private hUI                             As LongPtr
Private sEditText                       As String
Private sComboText                      As String
Private nComboIdx                       As LongPtr
Private lButtonY                        As Long
Private lButtonX                        As Long
Private lButtonY2                       As Long
Private lButtonX2                       As Long
Private lButtonY3                       As Long
Private lButtonX3                       As Long
Private lButtonX4                       As Long
Private lButtonY4                       As Long
Private arBtnX()                        As Long
Private bLinkAdjust                     As Boolean
Private bEditFooter                     As Boolean
Private bEditPassword                   As Boolean
Private bReposTrigger                   As Boolean
Private bPublicTimer                    As Boolean
Private tDTCur                          As SYSTEMTIME
Private tDTCur2                         As SYSTEMTIME
Private bDTCheck                        As Boolean
Private bDTCheck2                       As Boolean
Private lLastCY                         As Long
Private nTmrCnt                         As Long
Private himlCombo                       As LongPtr
Private nSplitButtonID                  As Long
Private m_sTitle                        As String
Private m_sMainInstr                    As String
Private m_sContent                      As String
Private m_sFooter                       As String
Private m_sExpandedInfo                 As String
Private m_sExpControl                   As String
Private m_sColControl                   As String
Private m_sVerify                       As String
Private m_sInput                        As String
Private m_sComboCue                     As String
Private m_uButtons()                    As TASKDIALOG_BUTTON
Private m_uRadioButtons()               As TASKDIALOG_BUTTON
Private m_Icon As Long, m_IconFooter    As Long
Private m_HasCustomButtons As Boolean, m_HasRadioButtons As Boolean
Private m_hButton()                     As LongPtr
Private m_BtnLRSize                     As Long
Private nInputAlign                     As TDInputBoxAlign
Private nComboAlign                     As TDInputBoxAlign
Private nDateTimeAlign                  As TDInputBoxAlign
Private nSliderAlign                    As TDInputBoxAlign
Private cxInputBox                      As Long
Private cxSlider                        As Long
Private cxCombo                         As Long
Private cyCombo                         As Long
Private hCombo                          As LongPtr
Private hDateTime As LongPtr
Private hDateTime2   As LongPtr
Private bPosUsed(2)                     As Boolean
Private m_HasCustomControl              As Boolean
Private m_lNewIcon                      As Long
Private nBreaksAdded                    As Integer
Private uTDC                            As TASKDIALOGCONFIG

Private hShell32                        As LongPtr
Private hImageRes                       As LongPtr
Private lgpX                            As Long
Private lgpY                            As Long


Private nDefaultCustCtrl                As TDCustomControlType
Private nDTContentAlign                 As TDControlAlign
Private nDTFooterAlign                  As TDControlAlign
Private nCBFooterAlign                  As TDControlAlign
Private nIBFooterAlign                  As TDControlAlign
Private nSLFooterAlign                  As TDControlAlign
Private nDTType                         As DateTimeType

Private tDTInitial                      As SYSTEMTIME
Private bDTSetInit                      As Boolean

Private tDTRange(1)                     As SYSTEMTIME
Private bDTSetMin                       As Boolean
Private bDTSetMax                       As Boolean
Private nCbType                         As ComboType

Private aComboItems()                   As ctdComboItem
Private tComboInit                      As ctdComboItem
Private nComboInit                      As Long
Private bCBInitValid                    As Boolean

Private m_ResMain                       As Long
Private m_ResRad                        As Long
Private m_ResVerify                     As Long



'     ______  __           ___           _______.     _______.
'    /      ||  |         /   \         /       |    /       |
'   |  ,----'|  |        /  ^  \       |   (----`   |   (----`
'   |  |     |  |       /  /_\  \       \   \        \   \
'   |  `----.|  `----. /  _____  \  .----)   |   .----)   |
'    \______||_______|/__/     \__\ |_______/    |_______/
                       
Private Sub Class_Initialize()
    Dim hdc                             As LongPtr
    
    hdc = GetDC(0&)
    m_ScaleX = GetDeviceCaps(hdc, LOGPIXELSX) / 96
    m_ScaleY = GetDeviceCaps(hdc, LOGPIXELSY) / 96
    ReleaseDC 0&, hdc
    ReDim m_uButtons(0)
    ReDim m_uRadioButtons(0)
    ReDim btnCaps(0)
    ReDim rbtnCaps(0)
    ReDim lBtnHold(0)
    ReDim lBtnElv(0)
    ReDim lBtnDis(0)
    ReDim lRadDis(0)
    ReDim icoDataC(0)
    ReDim aComboItems(0)
    lSliderMin = 0
    lSliderMax = 10
    lSliderNTicks = 1
    lSliderValue = 0
    lSliderChangeS = 1
    lSliderChangeL = 2
    nSliderTickStyle = 0
    m_BtnLRSize = 16
    
    '#If Win64 = False Then
        '// LoadLibrary is slow on O365 x64 version 2009
        hImageRes = LoadLibrary("imageres.dll")
        hShell32 = LoadLibrary("shell32.dll")
    '#End If
    
    uTDC.cbSize = Len(uTDC)
    uTDC.hInstance = 0     'App.hInstance
    
    '//Use the app default title
    m_sTitle = APP_TITLE
    
    #If Win64 Then
        AssignLongPtr uTDC.pszWindowTitle, uTDC.pszWindowTitle2, StrPtr(m_sTitle)
    #Else
        uTDC.pszWindowTitle = StrPtr(m_sTitle)
    #End If

End Sub

Private Sub Class_Terminate()
    If hShell32 Then FreeLibrary hShell32
    If hImageRes Then FreeLibrary hImageRes
End Sub

'   .___  ___.  _______ .___________. __    __    ______    _______       _______.
'   |   \/   | |   ____||           ||  |  |  |  /  __  \  |       \     /       |
'   |  \  /  | |  |__   `---|  |----`|  |__|  | |  |  |  | |  .--.  |   |   (----`
'   |  |\/|  | |   __|      |  |     |   __   | |  |  |  | |  |  |  |    \   \
'   |  |  |  | |  |____     |  |     |  |  |  | |  `--'  | |  '--'  |.----)   |
'   |__|  |__| |_______|    |__|     |__|  |__|  \______/  |_______/ |_______/
                                    
#If Win64 Then
'// TASKDIALOGCONFIG does NOT use 8-byte alignment so we have to use two Longs
'// Code From Shane Groff 2019-12-8
Private Sub AssignLongPtr(ByRef x As Long, ByRef x2 As Long, ByVal value As LongLong)
    Dim y As LongLong

    y = value And 4294967295#
    If (y > &H7FFFFFFF) Then
        x = CLng(y - 4294967296#)
        x2 = value / 4294967296# - 1
    Else
        x = CLng(y)
        x2 = value / 4294967296#
    End If

End Sub
#End If

Private Sub AddComboBox(nType As ComboType)
    Dim lEditX                          As Long
    Dim lEditY                          As Long
    Dim cxEdit                          As Long
    Dim cyEdit                          As Long
    Dim lOffX                           As Long
    Dim tRC                             As RECT
    Dim dwStyle                         As Long
    Dim hFont                           As Long
    Dim I                               As Long
    
    m_HasCustomControl = True
    GetClientRect m_TDhWnd, tRC
    'Debug.Print "Combo y=" & lButtonY2
    If nComboAlign = TDIBA_Footer Then
        If bPosUsed(2) Then Exit Sub
        bPosUsed(2) = True
        If cxCombo = -1 Then
            cxEdit = tRC.Right - ((16 * m_ScaleX) + 32)    '48
        Else
            cxEdit = 180 * m_ScaleX
        End If
        If cxCombo > 0 Then cxEdit = cxCombo
        Select Case nCBFooterAlign
            Case tdcaLeft
                lEditX = (16 * m_ScaleX) + 16    '32
            Case tdcaCenter
                lEditX = ((16 * m_ScaleX) + 16) + (((tRC.Right - ((16 * m_ScaleX) + 16) - 10) - (cxEdit)) \ 2)
            Case tdcaRight
                lEditX = tRC.Right - cxEdit - 10
        End Select
        If (m_sVerify <> "") And (m_sExpandedInfo <> "") Then
            lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4)    'lButtonY2 + 55
        Else
            lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4)    'lButtonY2 + 37
        End If
        If ((uTDC.dwFlags And TDF_EXPAND_FOOTER_AREA) = TDF_EXPAND_FOOTER_AREA) And ((uTDC.dwFlags And TDF_EXPANDED_BY_DEFAULT) = TDF_EXPANDED_BY_DEFAULT) And (m_sExpandedInfo <> "") Then
            lEditY = tRC.Bottom - ((58 * m_ScaleY) + 4)
            '        Debug.Print "lEditY=" & lEditY
        End If
        If (m_sFooter = "") Or (m_sFooter = "$input") Then
            m_sFooter = " "
            Call SendMessageW(m_TDhWnd, TDM_UPDATE_ELEMENT_TEXT, TDE_FOOTER, ByVal StrPtr(m_sFooter))
        End If
    ElseIf nComboAlign = TDIBA_Content Then
        If bPosUsed(0) Then Exit Sub
        bPosUsed(0) = True
        If m_Icon Then
            lOffX = (32 * m_ScaleX) + 20    '52
        Else
            lOffX = 12
        End If
        If (uTDC.dwFlags And TDF_KILL_SHIELD_ICON) And (m_lNewIcon = 0) Then
            lOffX = 12
        End If
        lEditX = lOffX
        If (uTDC.dwFlags And TDF_USE_COMMAND_LINKS) Or (uTDC.dwFlags And TDF_USE_COMMAND_LINKS_NO_ICON) Then
            lEditY = lButtonY2 - (32 * m_ScaleY)
        Else
            If ((Right$(m_sContent, 4) = vbCrLf & vbCrLf)) Then
                lEditY = lButtonY2 - (40 * m_ScaleY)
            Else
                lEditY = lButtonY2 - (34 * m_ScaleY)
            End If
            If bLinkAdjust Then
                lEditY = lEditY - (6 * m_ScaleY)
            End If
        End If
        If (uTDC.dwFlags And TDF_SHOW_MARQUEE_PROGRESS_BAR) Or (uTDC.dwFlags And TDF_SHOW_PROGRESS_BAR) Then
            lEditY = lEditY - (24 * m_ScaleY)
        End If
        If (m_sExpandedInfo <> "") Then
            lEditY = lEditY - (6 * m_ScaleY)
        End If
        cxEdit = tRC.Right - lOffX - 12
        If cxCombo > 0 Then cxEdit = cxCombo
    Else
        If bPosUsed(1) Then Exit Sub
        bPosUsed(1) = True
        lEditX = 10
        lEditY = lButtonY2
        '    cxEdit = 160
        cxEdit = GetLeftmostCommonButtonX()
        '    Debug.Print "Got Leftmost x=" & cxEdit
        cxEdit = cxEdit - lEditX - 10
        If cxEdit < 10 Then cxEdit = 10
        If cxCombo > 0 Then cxEdit = cxCombo
    End If
    If cyCombo > 0 Then
        cyEdit = cyCombo
    Else
        cyEdit = 115 * m_ScaleY
    End If


    dwStyle = WS_CHILD Or WS_VISIBLE Or CBS_AUTOHSCROLL Or WS_TABSTOP
    If nType = cbtDropdownEdit Then
        dwStyle = dwStyle Or CBS_DROPDOWN
    Else
        dwStyle = dwStyle Or CBS_DROPDOWNLIST
    End If

    hCombo = CreateWindowEx(0&, StrPtr(WC_COMBOBOXEX), StrPtr(""), _
                            dwStyle, _
                            lEditX, lEditY, cxEdit, cyEdit, _
                            hUI, IDD_COMBO, Access.hWndAccessApp, ByVal 0&)
    '        Debug.Print "hCombo=" & hCombo & ",x=" & lEditX & ",y=" & lEditY & ",cx=" & cxEdit
    Call SendMessageW(hCombo, CBEM_SETUNICODEFORMAT, 1, ByVal 0&)

    hEditCombo = SendMessageW(hCombo, CBEM_GETEDITCONTROL, 0, ByVal 0&)
    hComboCombo = SendMessageW(hCombo, CBEM_GETCOMBOCONTROL, 0, ByVal 0&)

    If m_sComboCue <> "" Then
        SendMessageW hEditCombo, EM_SETCUEBANNER, 0&, ByVal StrPtr(m_sComboCue)
    End If
    If himlCombo Then
        Call SendMessageW(hCombo, CBEM_SETIMAGELIST, 0, ByVal himlCombo)
    End If
    For I = 0 To UBound(aComboItems)
        CBX_InsertItem hCombo, aComboItems(I).sText, aComboItems(I).iImage, aComboItems(I).iOverlay
    Next I
    If bCBInitValid Then
        Dim cbxi                        As COMBOBOXEXITEMW
        cbxi.Mask = CBEIF_IMAGE Or CBEIF_OVERLAY
        cbxi.iItem = -1
        cbxi.iImage = tComboInit.iImage
        cbxi.iOverlay = tComboInit.iOverlay
        If tComboInit.sText <> "" Then
            cbxi.Mask = cbxi.Mask Or CBEIF_TEXT
            #If Win64 Then
                AssignLongPtr cbxi.pszText, cbxi.pszText2, StrPtr(tComboInit.sText)
            #Else
                cbxi.pszText = StrPtr(tComboInit.sText)
            #End If
        End If
        Call SendMessageW(hCombo, CBEM_SETITEMW, 0, cbxi)
    End If
    If nComboInit <> -1 Then
        Call SendMessageW(hCombo, CB_SETCURSEL, nComboInit, ByVal 0&)
    End If
    #If Win64 Then
    #Else
        Subclass hCombo, AddressOf TaskDialogSubclassProc, m_TDhWnd, ObjPtr(Me)
    #End If
End Sub

Public Sub AddCustomButton(nID As Long, sText As String, Optional hIcon As Long = 0&)
    If m_HasCustomButtons Then
        ReDim Preserve m_uButtons(UBound(m_uButtons) + 1)
        ReDim Preserve btnCaps(UBound(btnCaps) + 1)
        btnCaps(UBound(btnCaps)) = sText
        ReDim Preserve icoDataC(UBound(m_uButtons))
        If hIcon Then icoDataC(UBound(icoDataC)) = hIcon
        m_uButtons(UBound(m_uButtons)).nButtonID = nID
        #If Win64 Then
            AssignLongPtr m_uButtons(UBound(m_uButtons)).pszButtonText, m_uButtons(UBound(m_uButtons)).pszButtonText2, StrPtr(btnCaps(UBound(btnCaps)))
        #Else
            m_uButtons(UBound(m_uButtons)).pszButtonText = StrPtr(btnCaps(UBound(btnCaps)))
        #End If
    Else
        btnCaps(0) = sText
        m_uButtons(0).nButtonID = nID
        #If Win64 Then
            AssignLongPtr m_uButtons(0).pszButtonText, m_uButtons(0).pszButtonText2, StrPtr(btnCaps(0))
        #Else
            m_uButtons(0).pszButtonText = StrPtr(btnCaps(0))
        #End If
        If hIcon Then icoDataC(0) = hIcon
        m_HasCustomButtons = True
    End If
End Sub

Private Sub AddDateTime()
    Dim lEditX                          As Long
    Dim lEditY                          As Long
    Dim cxEdit                          As Long
    Dim cxEdit2                         As Long
    Dim lOffX                           As Long
    Dim tRC                             As RECT
    Dim dwUnits                         As Long
    Dim baseX                           As Long

    Dim dwStyle                         As Long
    Dim hFont                           As LongPtr
    Dim I                               As Long

    m_HasCustomControl = True
    dwUnits = GetDeviceCaps(GetDC(0), LOGPIXELSX)

    baseX = tdLoWord(dwUnits)

    GetClientRect m_TDhWnd, tRC

    Select Case nDTType

        Case dttDate
            cxEdit = 100
        Case dttDateWithCheck
            cxEdit = 116
        Case dttTime
            cxEdit = 90
        Case dttTimeWithCheck
            cxEdit = 106
        Case dttDateTime
            cxEdit = 100
            cxEdit2 = 90
        Case dttDateTimeWithCheck
            cxEdit = 116
            cxEdit2 = 106
        Case dttDateTimeWithCheckDateOnly
            cxEdit = 116
            cxEdit2 = 90
        Case dttDateTimeWithCheckTimeOnly
            cxEdit = 100
            cxEdit2 = 106
    End Select
    'cxEdit = cxEdit * m_ScaleX
    'cxEdit2 = cxEdit2 * m_ScaleX
    If nDateTimeAlign = TDIBA_Footer Then
        If bPosUsed(2) Then Exit Sub
        bPosUsed(2) = True
        Select Case nDTFooterAlign

            Case tdcaLeft
                lEditX = ((16 * m_ScaleX) + 16)    '32
            Case tdcaCenter
                lEditX = ((16 * m_ScaleX) + 16) + (((tRC.Right - ((16 * m_ScaleX) + 16) - 10) - ((cxEdit * m_ScaleX) + (cxEdit2 * m_ScaleX) + 5)) \ 2)
            Case tdcaRight
                lEditX = tRC.Right - (cxEdit * m_ScaleX) - (cxEdit2 * m_ScaleX) - 5 - 10
        End Select
        If lEditX < ((16 * m_ScaleX) + 16) Then lEditX = ((16 * m_ScaleX) + 16)
        If (m_sVerify <> "") And (m_sExpandedInfo <> "") Then
            'TODO: ADJUST FOR EXPAND-TO-FOOTER
            lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4)    'lButtonY3 + (23 * m_ScaleY) + 22 '58
        Else
            lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4)    'lButtonY3 + (23 * m_ScaleY) + 17 'lButtonY3 + (40 * m_ScaleY)
        End If
        If ((uTDC.dwFlags And TDF_EXPAND_FOOTER_AREA) = TDF_EXPAND_FOOTER_AREA) And ((uTDC.dwFlags And TDF_EXPANDED_BY_DEFAULT) = TDF_EXPANDED_BY_DEFAULT) And (m_sExpandedInfo <> "") Then
            lEditY = tRC.Bottom - ((58 * m_ScaleY) + 4)
            '        Debug.Print "lEditY=" & lEditY
        End If
        If (m_sFooter = "") Or (m_sFooter = "$input") Then
            m_sFooter = " "
            Call SendMessageW(m_TDhWnd, TDM_UPDATE_ELEMENT_TEXT, TDE_FOOTER, ByVal StrPtr(m_sFooter))
        End If
    ElseIf nDateTimeAlign = TDIBA_Content Then
        If bPosUsed(0) Then Exit Sub
        bPosUsed(0) = True
        If m_Icon Then
            lOffX = (32 * m_ScaleX) + 20    '52 + 16 '* m_ScaleX
        Else
            lOffX = 12
        End If
        If (uTDC.dwFlags And TDF_KILL_SHIELD_ICON) And (m_lNewIcon = 0) Then
            lOffX = 12
        End If
        Select Case nDTContentAlign

            Case tdcaLeft
                lEditX = lOffX
            Case tdcaCenter
                lEditX = lOffX + (((tRC.Right - lOffX - 12) - ((cxEdit * m_ScaleX) + (cxEdit2 * m_ScaleX) + 5)) \ 2)
            Case tdcaRight
                lEditX = (tRC.Right) - (cxEdit * m_ScaleX) - (cxEdit2 * m_ScaleX) - 5 - 12

        End Select
        If lEditX < lOffX Then lEditX = lOffX
        If (uTDC.dwFlags And TDF_USE_COMMAND_LINKS) Or (uTDC.dwFlags And TDF_USE_COMMAND_LINKS_NO_ICON) Or (m_HasRadioButtons = True) Then
            lEditY = lButtonY3 - (34 * m_ScaleY)
            '        Debug.Print "Set DT Y=" & lEditY
        Else
            If ((Right$(m_sContent, 4) = vbCrLf & vbCrLf)) Then
                lEditY = lButtonY3 - (44 * m_ScaleY)
            Else
                lEditY = lButtonY3 - (34 * m_ScaleY)
            End If
            If bLinkAdjust Then
                lEditY = lEditY - (6 * m_ScaleY)
            End If
        End If
        If (uTDC.dwFlags And TDF_SHOW_MARQUEE_PROGRESS_BAR) Or (uTDC.dwFlags And TDF_SHOW_PROGRESS_BAR) Then
            lEditY = lEditY - (22 * m_ScaleY)
        End If
        If (m_sExpandedInfo <> "") Then
            lEditY = lEditY - (6 * m_ScaleY)
        End If

    Else    'Button align
        If bPosUsed(1) Then Exit Sub
        bPosUsed(1) = True
        lEditX = 10
        lEditY = lButtonY3
    End If



    dwStyle = WS_CHILD Or WS_OVERLAPPED Or WS_VISIBLE
    If (nDTType = dttDateWithCheck) Or (nDTType = dttTimeWithCheck) Or (nDTType = dttDateTimeWithCheck) Or (nDTType = dttDateTimeWithCheckDateOnly) Then dwStyle = dwStyle Or DTS_SHOWNONE

    If (nDTType = dttDate) Or (nDTType = dttDateWithCheck) Or (nDTType >= dttDateTime) Then
        dwStyle = dwStyle Or DTS_SHORTDATEFORMAT
    Else
        dwStyle = dwStyle Or DTS_TIMEFORMAT
    End If
    hDateTime = CreateWindowEx(WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR, _
                               StrPtr("SysDateTimePick32"), StrPtr(""), _
                               dwStyle, _
                               lEditX, lEditY, cxEdit * m_ScaleX, 24 * m_ScaleY, _
                               hUI, IDD_DT1, Access.hWndAccessApp, ByVal 0&)

    If (nDTType = dttDateTime) Or (nDTType >= dttDateTimeWithCheck) Then
        dwStyle = WS_CHILD Or WS_OVERLAPPED Or WS_VISIBLE Or DTS_TIMEFORMAT
        If (nDTType = dttDateTimeWithCheck) Or (nDTType = dttDateTimeWithCheckTimeOnly) Then dwStyle = dwStyle Or DTS_SHOWNONE

        hDateTime2 = CreateWindowEx(WS_EX_CLIENTEDGE Or WS_EX_LEFT Or WS_EX_LTRREADING Or WS_EX_RIGHTSCROLLBAR, _
                                    StrPtr("SysDateTimePick32"), StrPtr(""), _
                                    dwStyle, _
                                    (lEditX + ((cxEdit + 5) * m_ScaleX)), lEditY, cxEdit2 * m_ScaleX, 24 * m_ScaleY, _
                                    hUI, IDD_DT2, Access.hWndAccessApp, ByVal 0&)
    End If

    hFont = GetStockObject(DEFAULT_GUI_FONT)
    If hFont Then
        SendMessageW hDateTime, WM_SETFONT, hFont, ByVal 0&
        If hDateTime2 Then
            SendMessageW hDateTime2, WM_SETFONT, hFont, ByVal 0&
        End If
    End If

    If (bDTSetMin = True) Or (bDTSetMax = True) Then
        SetDTRange
    End If
    If bDTSetInit Then
        SendMessageW hDateTime, DTM_SETSYSTEMTIME, GDT_VALID, tDTInitial
        If hDateTime2 Then
            SendMessageW hDateTime2, DTM_SETSYSTEMTIME, GDT_VALID, tDTInitial
        End If
    End If
    If (nDTType = dttDateWithCheck) Or (nDTType = dttTimeWithCheck) Or (nDTType = dttDateTimeWithCheck) Or (nDTType = dttDateTimeWithCheckDateOnly) Then
        If bDTCheck = False Then
            SendMessageW hDateTime, DTM_SETSYSTEMTIME, GDT_NONE, ByVal 0&
        End If
    End If
    If (nDTType = dttDateTimeWithCheck) Then
        If bDTCheck2 = False Then
            SendMessageW hDateTime2, DTM_SETSYSTEMTIME, GDT_NONE, ByVal 0&
        End If
    End If
    If (nDTType = dttDateTimeWithCheckTimeOnly) Then
        If bDTCheck = False Then
            SendMessageW hDateTime2, DTM_SETSYSTEMTIME, GDT_NONE, ByVal 0&
        End If
    End If
End Sub

Private Sub AddInputBox()
    Dim lEditX                          As Long
    Dim lEditY                          As Long
    Dim cxEdit                          As Long
    Dim lOffX                           As Long
    Dim tRC                             As RECT
    Dim dwStyle                         As Long
    Dim hFont                           As LongPtr
    Dim I                               As Long

    m_HasCustomControl = True
    GetClientRect m_TDhWnd, tRC


    'Debug.Print "lButtonY=" & lButtonY & ",btm=" & tRC.Bottom
    If nInputAlign = TDIBA_Footer Then
        If bPosUsed(2) Then Exit Sub
        bPosUsed(2) = True
        If cxInputBox = -1 Then
            cxEdit = tRC.Right - ((16 * m_ScaleX) + 32)    '48
        Else
            cxEdit = 180 * m_ScaleX
        End If
        If cxInputBox > 0 Then cxEdit = cxInputBox
        Select Case nIBFooterAlign

            Case tdcaLeft
                lEditX = (16 * m_ScaleX) + 16    '32
            Case tdcaCenter
                lEditX = ((16 * m_ScaleX) + 16) + (((tRC.Right - ((16 * m_ScaleX) + 16) - 10) - (cxEdit)) \ 2)
            Case tdcaRight
                lEditX = tRC.Right - (cxEdit) - 10
        End Select    'lEditY = tRC.Bottom - 27
        If (m_sVerify <> "") And (m_sExpandedInfo <> "") Then
            lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4)    'lButtonY3 + (23 * m_ScaleY) + 22 '58
        Else
            lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4)    'lButtonY3 + (23 * m_ScaleY) + 17 'lButtonY3 + (40 * m_ScaleY)
        End If
        If ((uTDC.dwFlags And TDF_EXPAND_FOOTER_AREA) = TDF_EXPAND_FOOTER_AREA) And ((uTDC.dwFlags And TDF_EXPANDED_BY_DEFAULT) = TDF_EXPANDED_BY_DEFAULT) And (m_sExpandedInfo <> "") Then
            lEditY = tRC.Bottom - ((58 * m_ScaleY) + 4)
        End If
        If (m_sFooter = "") Or (m_sFooter = "$input") Then
            m_sFooter = " "
            Call SendMessageW(m_TDhWnd, TDM_UPDATE_ELEMENT_TEXT, TDE_FOOTER, ByVal StrPtr(m_sFooter))
        End If
    ElseIf nInputAlign = TDIBA_Content Then
        If bPosUsed(0) Then Exit Sub
        bPosUsed(0) = True
        If m_Icon Then
            lOffX = (32 * m_ScaleX) + 20    '52
        Else
            lOffX = 12
        End If
        If (uTDC.dwFlags And TDF_KILL_SHIELD_ICON) And (m_lNewIcon = 0) Then
            lOffX = 12
        End If
        lEditX = lOffX
        If (uTDC.dwFlags And TDF_USE_COMMAND_LINKS) Or (uTDC.dwFlags And TDF_USE_COMMAND_LINKS_NO_ICON) Or (m_HasRadioButtons = True) Then
            lEditY = lButtonY - (34 * m_ScaleY)
        Else
            If ((Right$(m_sContent, 4) = vbCrLf & vbCrLf)) Then
                lEditY = lButtonY - (44 * m_ScaleY)
            Else
                lEditY = lButtonY - (34 * m_ScaleY)
            End If
            If bLinkAdjust Then
                lEditY = lEditY - (6 * m_ScaleY)
            End If
        End If
        If (uTDC.dwFlags And TDF_SHOW_MARQUEE_PROGRESS_BAR) Or (uTDC.dwFlags And TDF_SHOW_PROGRESS_BAR) Then
            lEditY = lEditY - (22 * m_ScaleY)
        End If
        If (m_sExpandedInfo <> "") Then
            lEditY = lEditY - (6 * m_ScaleY)
        End If
        cxEdit = tRC.Right - lOffX - 12
        If cxInputBox > 0 Then cxEdit = cxInputBox
    Else    'TDIBA_Buttons
        If bPosUsed(1) Then Exit Sub
        bPosUsed(1) = True
        lEditX = 10
        lEditY = lButtonY
        '    cxEdit = 160
        cxEdit = GetLeftmostCommonButtonX()
        '    Debug.Print "Got Leftmost x=" & cxEdit
        cxEdit = cxEdit - lEditX - 10
        If cxEdit < 10 Then cxEdit = 10
        If cxInputBox > 0 Then cxEdit = cxInputBox
    End If
    dwStyle = WS_VISIBLE Or WS_CHILD Or WS_TABSTOP Or ES_LEFT
    If bEditPassword Then
        dwStyle = dwStyle Or ES_PASSWORD
    End If


    hEditBox = CreateWindowEx(WS_EX_CLIENTEDGE, StrPtr("Edit"), StrPtr(""), _
                              dwStyle, _
                              lEditX, lEditY, cxEdit, 22 * m_ScaleY, _
                              hUI, IDD_EDIT1, Access.hWndAccessApp, ByVal 0&)



    hFont = GetStockObject(DEFAULT_GUI_FONT)
    If hFont Then
        SendMessageW hEditBox, WM_SETFONT, hFont, ByVal 0&
    End If
    If m_sInput <> "" Then
        SendMessageW hEditBox, EM_SETCUEBANNER, 0&, ByVal StrPtr(m_sInput)
    End If
    If sEditText <> "" Then
        SendMessageW hEditBox, WM_SETTEXT, 0&, ByVal StrPtr(sEditText)
        SendMessageW hEditBox, EM_SETSEL, 0&, ByVal (-1&)
    End If


End Sub

Private Sub AddLogo()
    Dim tBMP                            As BITMAP
    Dim CX                              As Long
    Dim CY                              As Long
    Dim px                              As Long
    Dim py                              As Long
    Dim tRC                             As RECT

    GetClientRect m_TDhWnd, tRC
    Call GetObject(hLogoImage, LenB(tBMP), tBMP)
    CX = tBMP.BMWidth
    CY = tBMP.BMHeight

    If nLogoPos = LogoTopRight Then
        px = tRC.Right - CX - lLogoX
        py = lLogoY
    ElseIf nLogoPos = LogoButtons Then
        px = lLogoX
        py = (lButtonY - 6) + lLogoY
    End If
    lgpX = px
    lgpY = py
    hStatic = CreateWindowEx(0, StrPtr("STATIC"), StrPtr(""), _
                             WS_VISIBLE Or WS_CHILD Or SS_BITMAP, _
                             px, py, CX, CY, hUI, 110, Access.hWndAccessApp, ByVal 0&)
    If hStatic Then
        SendMessageW hStatic, STM_SETIMAGE, nLogoType, ByVal hLogoImage
    End If
End Sub

Public Sub AddRadioButton(nID As Long, sText As String)
    If m_HasRadioButtons Then
        ReDim Preserve m_uRadioButtons(UBound(m_uRadioButtons) + 1)
        ReDim Preserve rbtnCaps(UBound(rbtnCaps) + 1)
        rbtnCaps(UBound(rbtnCaps)) = sText
        m_uRadioButtons(UBound(m_uRadioButtons)).nButtonID = nID
        #If Win64 Then
            AssignLongPtr m_uRadioButtons(UBound(m_uRadioButtons)).pszButtonText, m_uRadioButtons(UBound(m_uRadioButtons)).pszButtonText2, StrPtr(rbtnCaps(UBound(rbtnCaps)))
        #Else
            m_uRadioButtons(UBound(m_uRadioButtons)).pszButtonText = StrPtr(rbtnCaps(UBound(rbtnCaps)))
        #End If
    Else
        rbtnCaps(0) = sText
        m_uRadioButtons(0).nButtonID = nID
        #If Win64 Then
            AssignLongPtr m_uRadioButtons(0).pszButtonText, m_uRadioButtons(0).pszButtonText2, StrPtr(rbtnCaps(0))
        #Else
            m_uRadioButtons(0).pszButtonText = StrPtr(rbtnCaps(0))
        #End If
        
        m_HasRadioButtons = True
    End If
End Sub

Private Sub AddSlider()
    Dim lEditX                          As Long
    Dim lEditY                          As Long
    Dim cxEdit                          As Long
    Dim tRC                             As RECT
    Dim lOffX                           As Long
    Dim dwStyle                         As Long
    Dim hFont                           As Long
    Dim I                               As Long

    m_HasCustomControl = True
    GetClientRect m_TDhWnd, tRC
    If nSliderAlign = TDIBA_Footer Then
        If bPosUsed(2) Then Exit Sub
        bPosUsed(2) = True
        If cxSlider = -1 Then
            cxEdit = tRC.Right - ((16 * m_ScaleX) + 32)    '48
        Else
            cxEdit = 180 * m_ScaleX
        End If
        If cxSlider > 0 Then cxEdit = cxSlider
        Select Case nSLFooterAlign

            Case tdcaLeft
                lEditX = (16 * m_ScaleX) + 16    '32
            Case tdcaCenter
                lEditX = ((16 * m_ScaleX) + 16) + (((tRC.Right - ((16 * m_ScaleX) + 16) - 10) - (cxEdit)) \ 2)
            Case tdcaRight
                lEditX = tRC.Right - cxEdit - 10
        End Select
        If nSliderTickStyle = SldTickStyleBoth Then
            If (m_sVerify <> "") And (m_sExpandedInfo <> "") Then
                lEditY = tRC.Bottom - ((30 * m_ScaleY) + 5)    'lButtonY4 + 47
            Else
                lEditY = tRC.Bottom - ((30 * m_ScaleY))    'lButtonY4 + 29 tRC.Bottom - 38
            End If
        Else
            If (m_sVerify <> "") And (m_sExpandedInfo <> "") Then
                lEditY = tRC.Bottom - ((30 * m_ScaleY) + 5)    'lButtonY4 + 55
            Else
                lEditY = tRC.Bottom - ((30 * m_ScaleY))    ' lButtonY4 + 37 'tRC.Bottom - 30
            End If
        End If
        If ((uTDC.dwFlags And TDF_EXPAND_FOOTER_AREA) = TDF_EXPAND_FOOTER_AREA) And ((uTDC.dwFlags And TDF_EXPANDED_BY_DEFAULT) = TDF_EXPANDED_BY_DEFAULT) And (m_sExpandedInfo <> "") Then
            lEditY = tRC.Bottom - ((64 * m_ScaleY) + 5)
            '        Debug.Print "lEditY=" & lEditY
        End If
        If (m_sFooter = "") Or (m_sFooter = "$input") Then
            m_sFooter = " "
            '        Call SendMessageW(m_TDhWnd, TDM_UPDATE_ELEMENT_TEXT, TDE_FOOTER, ByVal StrPtr(m_sFooter))
        End If

        If (m_sExpandedInfo <> "") Then
            lEditY = lEditY - (6 * m_ScaleY)
        End If
        Call SendMessageW(m_TDhWnd, TDM_UPDATE_ELEMENT_TEXT, TDE_FOOTER, ByVal StrPtr(m_sFooter))
    ElseIf nSliderAlign = TDIBA_Content Then
        If bPosUsed(0) Then Exit Sub
        bPosUsed(0) = True
        If m_Icon Then
            lOffX = (32 * m_ScaleX) + 20    '52
        Else
            lOffX = 12
        End If
        If (uTDC.dwFlags And TDF_KILL_SHIELD_ICON) And (m_lNewIcon = 0) Then
            lOffX = 12
        End If
        lEditX = lOffX
        If (uTDC.dwFlags And TDF_USE_COMMAND_LINKS) Or (uTDC.dwFlags And TDF_USE_COMMAND_LINKS_NO_ICON) Or (m_HasRadioButtons = True) Then
            lEditY = lButtonY4 - (42 * m_ScaleY)
        Else
            If nSliderTickStyle = SldTickStyleBoth Then
                lEditY = lButtonY4 - (54 * m_ScaleY)
            Else
                lEditY = lButtonY4 - (44 * m_ScaleY)
            End If
            If bLinkAdjust Then
                lEditY = lEditY - (6 * m_ScaleY)
            End If
        End If
        If (uTDC.dwFlags And TDF_SHOW_MARQUEE_PROGRESS_BAR) Or (uTDC.dwFlags And TDF_SHOW_PROGRESS_BAR) Then
            lEditY = lEditY - (26 * m_ScaleY)
        End If

        cxEdit = tRC.Right - lOffX - 12
        If cxSlider > 0 Then cxEdit = cxSlider
    Else
        If bPosUsed(1) Then Exit Sub
        bPosUsed(1) = True
        lEditX = 10
        If nSliderTickStyle = SldTickStyleBoth Then
            lEditY = lButtonY4 - 10
        Else
            lEditY = lButtonY4
        End If
        '    cxEdit = 160
        cxEdit = GetLeftmostCommonButtonX()
        '    Debug.Print "Got Leftmost x=" & cxEdit
        cxEdit = cxEdit - lEditX - 10
        If cxSlider > 0 Then cxEdit = cxSlider
        If cxEdit < 10 Then cxEdit = 10
    End If

    dwStyle = WS_CHILD Or WS_VISIBLE Or TBS_AUTOTICKS Or TBS_HORZ Or TBS_TRANSPARENTBKGND
    Select Case nSliderTickStyle
        Case SldTickStyleBottom
            dwStyle = dwStyle Or TBS_BOTTOM
        Case SldTickStyleTop
            dwStyle = dwStyle Or TBS_TOP
        Case SldTickStyleBoth
            dwStyle = dwStyle Or TBS_BOTH
        Case SldTickStyleNone
            dwStyle = dwStyle Or TBS_NOTICKS
    End Select

    hSlider = CreateWindowEx(0&, StrPtr("msctls_trackbar32"), StrPtr("Slider"), dwStyle, _
                             lEditX, lEditY, cxEdit, 36 * m_ScaleY, _
                             hUI, IDD_SLIDER, Access.hWndAccessApp, ByVal 0&)
    If hSlider Then
        SendMessageW hSlider, TBM_SETRANGEMIN, 0, ByVal lSliderMin
        SendMessageW hSlider, TBM_SETRANGEMAX, 1, ByVal lSliderMax
        SendMessageW hSlider, TBM_SETTICFREQ, lSliderNTicks, ByVal 0&
        SendMessageW hSlider, TBM_SETLINESIZE, 0, ByVal lSliderChangeS
        SendMessageW hSlider, TBM_SETPAGESIZE, 0, ByVal lSliderChangeL
        SendMessageW hSlider, TBM_SETPOS, 1&, ByVal lSliderValue
    End If
    'Subclass hSlider, AddressOf TaskDialogSubclassProc, hSlider, ObjPtr(Me)

End Sub

Private Function ButtonIsHeld(id As Long) As Boolean
    Dim I                               As Long
    For I = 0 To UBound(lBtnHold)
        If lBtnHold(I) = id Then
            ButtonIsHeld = True
            Exit Function
        End If
    Next I
End Function

Private Function CBX_InsertItem(hCB As LongPtr, sText As String, Optional iImage As Long = -1, Optional iOverlay As Long = -1, Optional lParam As LongPtr = 0, Optional iItem As Long = -1) As LongPtr
'Debug.Print "CBX_InsertItem " & sText & ", " & iImage
    Dim cbxi                            As COMBOBOXEXITEMW

    With cbxi
        .Mask = CBEIF_TEXT _
                Or CBEIF_LPARAM _
                Or CBEIF_IMAGE _
                Or CBEIF_SELECTEDIMAGE _
                Or CBEIF_INDENT _
                Or CBEIF_LPARAM
        '.Mask = CBEIF_TEXT
        .cchTextMax = Len(sText)
        #If Win64 Then
            AssignLongPtr .pszText, .pszText2, StrPtr(sText)
        #Else
            .pszText = StrPtr(sText)
        #End If
        .iItem = iItem
        .iImage = iImage
        .iSelectedImage = iImage
        .iOverlay = iOverlay
        .lParam = lParam
    End With
    
    #If Win64 Then
        'Crashes calling SendMessageW
        Debug.Print sText
        CBX_InsertItem = SendMessageW(hCB, CBEM_INSERTITEMW, 0, cbxi)
    #Else
        CBX_InsertItem = SendMessageW(hCB, CBEM_INSERTITEMW, 0, cbxi)
    #End If
End Function

Public Sub ClearCustomButtons()
    ReDim m_uButtons(0)
    ReDim btnCaps(0)
    ReDim icoDataC(0)
    nSplitButtonID = 0
    uTDC.cButtons = 0
    #If Win64 Then
        AssignLongPtr uTDC.pButtons, uTDC.pButtons2, VarPtr(m_uButtons(0).nButtonID)
    #Else
        uTDC.pButtons = VarPtr(m_uButtons(0).nButtonID)
    #End If
    m_HasCustomButtons = False
End Sub

Public Sub ClearRadioButtons()
    ReDim m_uRadioButtons(0)
    ReDim rbtnCaps(0)
    uTDC.cRadioButtons = 0
    #If Win64 Then
        AssignLongPtr uTDC.pRadioButtons, uTDC.pRadioButtons2, VarPtr(m_uRadioButtons(0).nButtonID)
    #Else
         uTDC.pRadioButtons = VarPtr(m_uRadioButtons(0).nButtonID)
    #End If
    m_HasRadioButtons = False
End Sub

Public Sub ClickButton(ButtonID As Long)
    If m_TDhWnd <> 0& Then
        Call SendMessageW(m_TDhWnd, TDM_CLICK_BUTTON, ButtonID, ByVal 0&)
    End If
End Sub

Public Sub ClickRadioButton(ButtonID As Long)
    If m_TDhWnd <> 0& Then
        Call SendMessageW(m_TDhWnd, TDM_CLICK_RADIO_BUTTON, ButtonID, ByVal 0&)
    End If
End Sub

Public Sub ClickVerification(lChecked As Long, Optional lSetFocus As Long = 1)
'0=unchecked,1=checked
'lSetFocus=1, sets keyboard focus on box. 0 to not set focus
    If m_TDhWnd <> 0& Then
        Call SendMessageW(m_TDhWnd, TDM_CLICK_BUTTON, lChecked, ByVal lSetFocus)
    End If
End Sub

Public Sub CloseDialog()
'closes the dialog if open
    If m_TDhWnd <> 0& Then
        Call ClickButton(TDCBF_CANCEL_BUTTON)
    End If

End Sub

Public Sub ComboAddItem(sText As String, Optional iImage As Long = -1, Optional iOverlay As Long = -1)
    If aComboItems(0).sText = "" Then
        aComboItems(0).sText = sText
        aComboItems(0).iImage = iImage
        aComboItems(0).iOverlay = iOverlay
        Exit Sub
    End If
    ReDim Preserve aComboItems(UBound(aComboItems) + 1)
    aComboItems(UBound(aComboItems)).sText = sText
    aComboItems(UBound(aComboItems)).iImage = iImage
    aComboItems(UBound(aComboItems)).iOverlay = iOverlay
    If hCombo Then
        CBX_InsertItem hCombo, sText, iImage, iOverlay
    End If
End Sub

Private Sub ComboReposition()
    Dim lOffX                           As Long
    Dim tRC                             As RECT
    Dim lEditX                          As Long
    Dim lEditY                          As Long
    Dim cxEdit                          As Long

    If hCombo Then

        lButtonY2 = 0
        ReDim arBtnX(0)

        Call EnumChildWindows(m_TDhWnd, AddressOf TaskDialogEnumChildProc, ObjPtr(Me))

        GetClientRect hUI, tRC

        If nComboAlign = TDIBA_Content Then
            If m_Icon Then
                lOffX = (32 * m_ScaleX) + 20    '52
            Else
                lOffX = 12
            End If
            If (uTDC.dwFlags And TDF_KILL_SHIELD_ICON) And (m_lNewIcon = 0) Then
                lOffX = 12
            End If
            lEditX = lOffX
            If (uTDC.dwFlags And TDF_USE_COMMAND_LINKS) Or (uTDC.dwFlags And TDF_USE_COMMAND_LINKS_NO_ICON) Or (m_HasRadioButtons = True) Then
                lEditY = lButtonY2 - (28 * m_ScaleY)
            Else
                '            If ((Right$(m_sContent, 4) = vbCrLf & vbCrLf)) Then
                lEditY = lButtonY2 - (44 * m_ScaleY)
                '            Else
                '                lEditY = lButtonY2 - 34
                '            End If
            End If
            If (uTDC.dwFlags And TDF_SHOW_MARQUEE_PROGRESS_BAR) Or (uTDC.dwFlags And TDF_SHOW_PROGRESS_BAR) Then
                lEditY = lEditY - (24 * m_ScaleY)
            End If
            If (m_sExpandedInfo <> "") Then
                lEditY = lEditY - (6 * m_ScaleY)
            End If
            cxEdit = tRC.Right - lOffX - 12
        ElseIf nComboAlign = TDIBA_Footer Then
            '        lEditX = 32
            '        lEditY = tRC.Bottom - 30
            '        cxEdit = 180
            If cxCombo = -1 Then
                cxEdit = tRC.Right - ((16 * m_ScaleX) + 32)    '48
            Else
                cxEdit = 180 * m_ScaleX
            End If
            If cxCombo > 0 Then cxEdit = cxCombo
            Select Case nCBFooterAlign
                Case tdcaLeft
                    lEditX = (16 * m_ScaleX) + 16    '32
                Case tdcaCenter
                    lEditX = ((16 * m_ScaleX) + 16) + (((tRC.Right - ((16 * m_ScaleX) + 16) - 10) - (cxEdit)) \ 2)
                Case tdcaRight
                    lEditX = tRC.Right - cxEdit - 10
            End Select
            If (m_sVerify <> "") And (m_sExpandedInfo <> "") Then
                lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4)    'lButtonY2 + 55
            Else
                lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4)    'lButtonY2 + 37
            End If

        End If
        If (nComboAlign = TDIBA_Content) Or (nComboAlign = TDIBA_Footer) Then
            SetWindowPos hCombo, 0&, lEditX, lEditY, cxEdit, 115&, SWP_NOSIZE
            RepaintWindow hCombo
        End If
        SetWindowPos m_TDhWnd, 0, 0, 0, tRC.Right, tRC.Bottom + (28 * m_ScaleY), SWP_NOMOVE
    End If
End Sub

Public Sub ComboSetCurrentState(sText As String, Optional iImage As Long = -1, Optional iOverlay As Long = -1)
    Dim cbxi                            As COMBOBOXEXITEMW
    cbxi.Mask = CBEIF_IMAGE Or CBEIF_OVERLAY
    cbxi.iItem = -1
    cbxi.iImage = iImage
    cbxi.iOverlay = iOverlay
    If sText <> "" Then
        cbxi.Mask = cbxi.Mask Or CBEIF_TEXT
        #If Win64 Then
            AssignLongPtr cbxi.pszText, cbxi.pszText2, StrPtr(sText)
        #Else
            cbxi.pszText = StrPtr(sText)
        #End If
    End If
    Call SendMessageW(hCombo, CBEM_SETITEMW, 0, cbxi)

End Sub

Public Sub ComboSetInitialItem(nIndex As Long)
    nComboInit = nIndex
End Sub

Public Sub ComboSetInitialState(sText As String, Optional iImage As Long = -1, Optional iOverlay As Long = -1)
    bCBInitValid = True
    tComboInit.sText = sText
    tComboInit.iImage = iImage
    tComboInit.iOverlay = iOverlay
End Sub

Private Sub DateTimeReposition()
    Dim tRC                             As RECT
    Dim lEditX                          As Long
    Dim lEditY                          As Long
    Dim cxEdit                          As Long
    Dim cxEdit2                         As Long

    If hDateTime Then
        lButtonY3 = 0
        ReDim arBtnX(0)

        Call EnumChildWindows(m_TDhWnd, AddressOf TaskDialogEnumChildProc, ObjPtr(Me))

        GetClientRect hUI, tRC

        Select Case nDTType

            Case dttDate
                cxEdit = 100
            Case dttDateWithCheck
                cxEdit = 116
            Case dttTime
                cxEdit = 90
            Case dttTimeWithCheck
                cxEdit = 106
            Case dttDateTime
                cxEdit = 100
                cxEdit2 = 90
            Case dttDateTimeWithCheck
                cxEdit = 116
                cxEdit2 = 106
            Case dttDateTimeWithCheckDateOnly
                cxEdit = 116
                cxEdit2 = 90
            Case dttDateTimeWithCheckTimeOnly
                cxEdit = 100
                cxEdit2 = 106
        End Select
        Dim lOffX                       As Long
        If nDateTimeAlign = TDIBA_Content Then
            If m_Icon Then
                lOffX = (32 * m_ScaleX) + 20
            Else
                lOffX = 12
            End If
            Select Case nDTContentAlign

                Case tdcaLeft
                    lEditX = lOffX
                Case tdcaCenter
                    '                Debug.Print "repos width=" & tRC.Right & ",cx1=" & cxEdit & ",cx2=" & cxEdit2
                    '                Debug.Print "repos x=" & (((tRC.Right - 52 - 12) - (cxEdit + cxEdit2 + 5)) \ 2)
                    lEditX = lOffX + (((tRC.Right - lOffX - 12) - ((cxEdit * m_ScaleX) + (cxEdit2 * m_ScaleX) + 5)) \ 2)
                Case tdcaRight
                    lEditX = (tRC.Right) - (cxEdit * m_ScaleX) - (cxEdit2 * m_ScaleX) - 5 - 12

            End Select
            If lEditX < lOffX Then lEditX = lOffX
            If (uTDC.dwFlags And TDF_USE_COMMAND_LINKS) Or (uTDC.dwFlags And TDF_USE_COMMAND_LINKS_NO_ICON) Or (m_HasRadioButtons = True) Then
                lEditY = lButtonY3 - (34 * m_ScaleY)
                '            Debug.Print "Set DT Y=" & lEditY
            Else
                '            If ((Right$(m_sContent, 4) = vbCrLf & vbCrLf)) Then
                lEditY = lButtonY3 - (44 * m_ScaleY)
                '            Else
                '                lEditY = lButtonY3 - 34
                '            End If
                If bLinkAdjust Then
                    lEditY = lEditY - (6 * m_ScaleY)
                End If
            End If
            If (uTDC.dwFlags And TDF_SHOW_MARQUEE_PROGRESS_BAR) Or (uTDC.dwFlags And TDF_SHOW_PROGRESS_BAR) Then
                lEditY = lEditY - (22 * m_ScaleY)
            End If
            If (m_sExpandedInfo <> "") Then
                lEditY = lEditY - (6 * m_ScaleY)
            End If

        ElseIf nDateTimeAlign = TDIBA_Footer Then
            '        lEditX = 32
            '        lEditY = tRC.Bottom - 30
            '        m_sFooter = " "
            Select Case nDTFooterAlign

                Case tdcaLeft
                    lEditX = ((16 * m_ScaleX) + 16)    '32
                Case tdcaCenter
                    lEditX = ((16 * m_ScaleX) + 16) + (((tRC.Right - ((16 * m_ScaleX) + 16) - 10) - ((cxEdit * m_ScaleX) + (cxEdit2 * m_ScaleX) + 5)) \ 2)
                Case tdcaRight
                    lEditX = tRC.Right - (cxEdit * m_ScaleX) - (cxEdit2 * m_ScaleX) - 5 - 10
            End Select
            If lEditX < ((16 * m_ScaleX) + 16) Then lEditX = ((16 * m_ScaleX) + 16)
            If (m_sVerify <> "") And (m_sExpandedInfo <> "") Then
                'TODO: ADJUST FOR EXPAND-TO-FOOTER
                lEditY = tRC.Bottom - ((24 * m_ScaleY) + 6)    'lButtonY3 + (23 * m_ScaleY) + 22 '58
            Else
                lEditY = tRC.Bottom - ((24 * m_ScaleY) + 6)    'lButtonY3 + (23 * m_ScaleY) + 17 'lButtonY3 + (40 * m_ScaleY)
            End If

        End If
        'If (nDTType = dttTime) Then
        '    cxEdit = 90
        'Else
        '    cxEdit = 100
        'End If


        If (nDateTimeAlign = TDIBA_Content) Or (nDateTimeAlign = TDIBA_Footer) Then
            SetWindowPos hDateTime, 0&, lEditX, lEditY, cxEdit * m_ScaleX, 24 * m_ScaleY, SWP_NOSIZE
            RepaintWindow hDateTime
            If hDateTime2 Then
                SetWindowPos hDateTime2, 0&, (lEditX + ((cxEdit + 5) * m_ScaleX)), lEditY, cxEdit2 * m_ScaleX, 24 * m_ScaleY, SWP_NOSIZE
                RepaintWindow hDateTime2
            End If
        End If
        SetWindowPos m_TDhWnd, 0, 0, 0, tRC.Right, tRC.Bottom + 30 * m_ScaleY, SWP_NOMOVE
    End If
End Sub

Public Sub DateTimeSetInitial(dt As Date)
    bDTSetInit = True
    With tDTInitial
        .wDay = VBA.DateTime.Day(dt)
        .wMonth = VBA.DateTime.Month(dt)
        .wYear = VBA.DateTime.Year(dt)
        .wHour = VBA.DateTime.Hour(dt)
        .wMinute = VBA.DateTime.Minute(dt)
        .wSecond = VBA.DateTime.Second(dt)
    End With
End Sub

Public Sub DateTimeSetRange(bEnableMin As Boolean, bEnableMax As Boolean, dtMin As Date, dtMax As Date)
    bDTSetMin = bEnableMin
    bDTSetMax = bEnableMax
    With tDTRange(0)
        .wDay = VBA.DateTime.Day(dtMin)
        .wMonth = VBA.DateTime.Month(dtMin)
        .wYear = VBA.DateTime.Year(dtMin)
        .wHour = VBA.DateTime.Hour(dtMin)
        .wMinute = VBA.DateTime.Minute(dtMin)
        .wSecond = VBA.DateTime.Second(dtMin)
    End With
    With tDTRange(1)
        .wDay = VBA.DateTime.Day(dtMax)
        .wMonth = VBA.DateTime.Month(dtMax)
        .wYear = VBA.DateTime.Year(dtMax)
        .wHour = VBA.DateTime.Hour(dtMax)
        .wMinute = VBA.DateTime.Minute(dtMax)
        .wSecond = VBA.DateTime.Second(dtMax)
    End With
    SetDTRange
End Sub

Public Sub EnableButton(ButtonID As Long, lEnable As Long)
'lEnable=0 disable; <>0 enable
    If m_TDhWnd <> 0& Then
        Call SendMessageW(m_TDhWnd, TDM_ENABLE_BUTTON, ButtonID, ByVal lEnable)
    Else
        If lEnable = 0 Then
            If (lBtnDis(0) = 0) And (UBound(lBtnDis) = 0) Then
                lBtnDis(0) = ButtonID
            Else
                ReDim Preserve lBtnDis(UBound(lBtnDis) + 1)
                lBtnDis(UBound(lBtnDis)) = ButtonID
            End If
        Else
            Dim lnew()                  As Long
            ReDim lnew(0)
            Dim I As Long, k            As Long
            For I = 0 To UBound(lBtnDis)
                If lBtnDis(I) <> ButtonID Then
                    ReDim Preserve lnew(k)
                    lnew(k) = lBtnDis(I)
                End If
            Next I
            lBtnDis = lnew
        End If
        If (UBound(lBtnDis) = 0) And (lBtnDis(0) = 0) Then
            bDis = False
        Else
            bDis = True
        End If
    End If
End Sub

Public Sub EnableRadioButton(ButtonID As Long, lEnable As Long)
'lEnable=0 disable; <>0 enable
    Dim lnew()                          As Long
    Dim I                               As Long
    Dim k                               As Long

    If m_TDhWnd <> 0& Then
        Call SendMessageW(m_TDhWnd, TDM_ENABLE_RADIO_BUTTON, ButtonID, ByVal lEnable)
    Else
        If lEnable = 0 Then
            If (lRadDis(0) = 0) And (UBound(lRadDis) = 0) Then
                lRadDis(0) = ButtonID
            Else
                ReDim Preserve lRadDis(UBound(lRadDis) + 1)
                lRadDis(UBound(lRadDis)) = ButtonID
            End If
        Else
            ReDim lnew(0)
            For I = 0 To UBound(lRadDis)
                If lRadDis(I) <> ButtonID Then
                    ReDim Preserve lnew(k)
                    lnew(k) = lRadDis(I)
                End If
            Next I
            lRadDis = lnew
        End If
        If (UBound(lRadDis) = 0) And (lRadDis(0) = 0) Then
            bRadDis = False
        Else
            bRadDis = True
        End If
    End If
End Sub

Private Sub FocusMainControl()
'when custom controls are added, determines which one should get focus
'Priorities:
'Main=manually specified; else 1st=content, 2nd=buttons, 3rd=footer
    If nDefaultCustCtrl > 0 Then
        Select Case nDefaultCustCtrl
            Case tdccInputBox: SetFocusAPI hEditBox
            Case tdccComboBox: SetFocusAPI hCombo
            Case tdccDateTime: SetFocusAPI hDateTime
            Case tdccSlider: SetFocusAPI hSlider
            Case tdccNothing: Exit Sub
        End Select
        Exit Sub
    End If

    If ((uTDC.dwFlags And TDF_INPUT_BOX) And (nInputAlign = TDIBA_Content)) Then
        SetFocusAPI hEditBox
        Exit Sub
    End If
    If ((uTDC.dwFlags And TDF_COMBO_BOX) And (nComboAlign = TDIBA_Content)) Then
        SetFocusAPI hCombo
        Exit Sub
    End If
    If ((uTDC.dwFlags And TDF_DATETIME) And (nDateTimeAlign = TDIBA_Content)) Then
        SetFocusAPI hDateTime
        Exit Sub
    End If
    If ((uTDC.dwFlags And TDF_SLIDER) And (nSliderAlign = TDIBA_Content)) Then
        SetFocusAPI hSlider
        Exit Sub
    End If
    If ((uTDC.dwFlags And TDF_INPUT_BOX) And (nInputAlign = TDIBA_Buttons)) Then
        SetFocusAPI hEditBox
        Exit Sub
    End If
    If ((uTDC.dwFlags And TDF_COMBO_BOX) And (nComboAlign = TDIBA_Buttons)) Then
        SetFocusAPI hCombo
        Exit Sub
    End If
    If ((uTDC.dwFlags And TDF_DATETIME) And (nDateTimeAlign = TDIBA_Buttons)) Then
        SetFocusAPI hDateTime
        Exit Sub
    End If
    If ((uTDC.dwFlags And TDF_SLIDER) And (nSliderAlign = TDIBA_Buttons)) Then
        SetFocusAPI hSlider
        Exit Sub
    End If
    If ((uTDC.dwFlags And TDF_INPUT_BOX) And (nInputAlign = TDIBA_Footer)) Then
        SetFocusAPI hEditBox
        Exit Sub
    End If
    If ((uTDC.dwFlags And TDF_COMBO_BOX) And (nComboAlign = TDIBA_Footer)) Then
        SetFocusAPI hCombo
        Exit Sub
    End If
    If ((uTDC.dwFlags And TDF_DATETIME) And (nDateTimeAlign = TDIBA_Footer)) Then
        SetFocusAPI hDateTime
        Exit Sub
    End If
    If ((uTDC.dwFlags And TDF_SLIDER) And (nSliderAlign = TDIBA_Footer)) Then
        SetFocusAPI hSlider
        Exit Sub
    End If
End Sub


Private Function GetComboTextW(hEdit As LongPtr) As String
    Dim bytS()                          As Byte
    Dim ch                              As LongPtr
    
    ch = SendMessageW(hEdit, WM_GETTEXTLENGTH, 0, ByVal 0&) * 2 + 1
    ReDim Preserve bytS(CByte(ch)) '//KCB added CByte
    ch = SendMessageW(hEdit, WM_GETTEXT, ch, ByVal VarPtr(bytS(0))) * 2 - 1
    If ch >= 0 Then
        ReDim Preserve bytS(CByte(ch)) '//KCB added CByte
    End If
    GetComboTextW = CStr(bytS)

End Function

Public Sub GetCustomButtons(nID() As Long, sText() As String)
    Dim I                               As Long
    Dim lID()                           As Long
    Dim c                               As Long
    
    c = UBound(m_uButtons)
    ReDim lID(c)
    For I = 0 To c
        lID(I) = m_uButtons(I).nButtonID
    Next I
    nID = lID
    sText = btnCaps

End Sub

Private Function GetLeftmostCommonButtonX() As Long
'arBtnX stores all button X positions, but this includes command links
'and radio buttons. To size button-area controls, we need the furthest
'left position of only the buttons that appear in the button area.
    Dim nButtonsNA                      As Long    'number of n/a buttons (radio+cmdlink)
    Dim I As Long, J                    As Long
    Dim arValid()                       As Long
    Dim CX                              As Long

    If m_HasRadioButtons Then
        nButtonsNA = UBound(m_uRadioButtons) + 1
    End If
    If (uTDC.dwFlags And TDF_USE_COMMAND_LINKS) Or (uTDC.dwFlags And TDF_USE_COMMAND_LINKS_NO_ICON) Then
        If m_HasCustomButtons Then
            nButtonsNA = nButtonsNA + UBound(m_uButtons) + 1
        Else
            'if there's NO buttons on the bottom, return -1 to indicate that
            'the caller should then use the dialog width or a standard width
            GetLeftmostCommonButtonX = -1
            Exit Function
        End If
    End If
Debug.Print "Got nButtonsNA=" & nButtonsNA
    If nButtonsNA > 0 Then
        SortLongArray arBtnX    'radio buttons and command links will be the smallest
        'X values (for scenarios where there's even room in the
        'button area for the control), so we can eliminate the
        'nButtonsNA smallest values and this translates to
        'eliminating the N/A buttons.
        ReDim arValid(UBound(arBtnX) - nButtonsNA)
        For I = nButtonsNA To UBound(arBtnX)
            arValid(J) = arBtnX(I)
            J = J + 1
        Next I
    Else
        arValid = arBtnX
    End If
    If UBound(arValid) = 0 Then
        GetLeftmostCommonButtonX = arValid(0)
        Exit Function
    End If
    CX = arValid(0)
    For I = 1 To UBound(arValid)
        If arValid(I) < CX Then
            CX = arValid(I)
        End If
    Next I
    GetLeftmostCommonButtonX = CX
End Function

Public Sub GetRadioButtons(nID() As Long, sText() As String)
    Dim I                               As Long
    Dim lID()                           As Long
    Dim c                               As Long
    
    c = UBound(m_uRadioButtons)
    ReDim lID(c)
    For I = 0 To c
        lID(I) = m_uRadioButtons(I).nButtonID
    Next I
    nID = lID
    sText = rbtnCaps

End Sub

Public Sub Init()
'resets the entire module
    hLogoImage = 0
    nLogoPos = 0
    nLogoType = 0
    lLogoX = 0
    lLogoY = 0
    lAutoclose = 0
    bLinkAdjust = False
    ClearCustomButtons
    ClearRadioButtons
    m_HasCustomControl = False
    ReDim m_uButtons(0)
    ReDim m_uRadioButtons(0)
    ReDim btnCaps(0)
    ReDim rbtnCaps(0)
    ReDim lBtnElv(0)
    ReDim lBtnHold(0)
    ReDim lBtnDis(0)
    ReDim lRadDis(0)
    ReDim icoDataC(0)
    ReDim aComboItems(0)
    tComboInit.sText = ""
    tComboInit.iImage = 0
    tComboInit.iOverlay = 0
    nCbType = 0
    bCBInitValid = False
    nDTType = 0
    bDTSetInit = False
    bDTCheck = False
    bDTSetMin = False
    bDTSetMax = False
    nComboInit = -1
    himlCombo = 0
    hDateTime = 0
    hDateTime2 = 0
    hCombo = 0
    hSlider = 0
    hEditBox = 0
    icoDataN(5) = 0
    icoDataN(4) = 0
    icoDataN(3) = 0
    icoDataN(2) = 0
    icoDataN(1) = 0
    icoDataN(0) = 0
    m_lNewIcon = 0
    m_BtnLRSize = 16
    bDis = False
    bRadDis = False
    bElv = False
    nInputAlign = TDIBA_Content
    nComboAlign = TDIBA_Content
    nDateTimeAlign = TDIBA_Content
    nDTContentAlign = 0
    nDTFooterAlign = 0
    nCBFooterAlign = 0
    nIBFooterAlign = 0
    nSLFooterAlign = 0
    nSliderAlign = TDIBA_Content
    nDefaultCustCtrl = 0
    bPosUsed(0) = False
    bPosUsed(1) = False
    bPosUsed(2) = False
    cxInputBox = 0
    cxCombo = 0
    cyCombo = 0
    cxSlider = 0
    lSliderMin = 0
    lSliderMax = 10
    lSliderNTicks = 1
    lSliderValue = 0
    lSliderChangeS = 1
    lSliderChangeL = 2
    nSliderTickStyle = 0
    nBreaksAdded = 0
        With uTDC
        .cxWidth = 0
        .dwCommonButtons = 0
        .dwFlags = 0
        .hInstance = 0    ' App.hInstance
        .hwndParent = 0
        .nDefaultButton = 0
        .nDefaultRadioButton = 0
        m_sExpandedInfo = vbNullString
        .pszExpandedInformation = 0    'StrPtr(m_sExpandedInfo)

        m_sInput = vbNullString
        sEditText = vbNullString
        m_sColControl = vbNullString
        m_sContent = vbNullString
        m_sExpControl = vbNullString
        m_sFooter = vbNullString
        m_IconFooter = 0
        m_Icon = 0
        m_sMainInstr = vbNullString
        m_sVerify = vbNullString
        m_sTitle = APP_TITLE
        .pszFooterIcon = 0
        .pszMainIcon = 0
        
        #If Win64 Then
            AssignLongPtr .pszCollapsedControlText, .pszCollapsedControlText2, StrPtr(m_sColControl)
            AssignLongPtr .pszContent, .pszContent2, StrPtr(m_sContent)
            AssignLongPtr .pszExpandedControlText, .pszExpandedControlText2, StrPtr(m_sExpControl)
            AssignLongPtr .pszFooter, .pszFooter2, StrPtr(m_sFooter)
            AssignLongPtr .pszMainInstruction, .pszMainInstruction2, StrPtr(m_sMainInstr)
            AssignLongPtr .pszVerificationText, .pszVerificationText2, StrPtr(m_sVerify)
            AssignLongPtr .pszWindowTitle, .pszWindowTitle2, StrPtr(m_sTitle)
        #Else
            .pszCollapsedControlText = StrPtr(m_sColControl)
            .pszContent = StrPtr(m_sContent)
            .pszExpandedControlText = StrPtr(m_sExpControl)
            .pszFooter = StrPtr(m_sFooter)
            .pszMainInstruction = StrPtr(m_sMainInstr)
            .pszVerificationText = StrPtr(m_sVerify)
            .pszWindowTitle = StrPtr(m_sTitle)
        #End If
        
    End With

'    With uTDC
'        .cxWidth = 0
'        .dwCommonButtons = 0
'        .dwFlags = 0
'        .hInstance = 0    ' App.hInstance
'        .hwndParent = 0
'        .nDefaultButton = 0
'        .nDefaultRadioButton = 0
'        m_sExpandedInfo = ""
'        .pszExpandedInformation = 0    'StrPtr(m_sExpandedInfo)
'
'        m_sColControl = ""
'        uTDC.pszCollapsedControlText = StrPtr(m_sColControl)
'        '
'        m_sContent = ""
'        .pszContent = StrPtr(m_sContent)
'        '
'        m_sExpControl = ""
'        .pszExpandedControlText = StrPtr(m_sExpControl)
'
'        '
'        m_sFooter = ""
'        .pszFooter = StrPtr(m_sFooter)
'        '
'        m_IconFooter = 0
'        .pszFooterIcon = 0
'        '
'        m_Icon = 0
'        .pszMainIcon = 0
'        '
'        m_sMainInstr = ""
'        .pszMainInstruction = StrPtr(m_sMainInstr)
'        '
'        m_sVerify = ""
'        .pszVerificationText = StrPtr(m_sVerify)
'        '
'        m_sTitle = APP_TITLE
'        .pszWindowTitle = StrPtr(m_sTitle)
'        '
'        m_sInput = ""
'        sEditText = ""
'
'    End With
    bEditPassword = False
    bEditFooter = False
End Sub

Private Sub InputReposition()
    Dim tRC                             As RECT
    Dim lOffX                           As Long
    Dim lEditX                          As Long
    Dim lEditY                          As Long
    Dim cxEdit As Long

    If hEditBox Then

        lButtonY = 0
        ReDim arBtnX(0)

        Call EnumChildWindows(m_TDhWnd, AddressOf TaskDialogEnumChildProc, ObjPtr(Me))

        GetClientRect hUI, tRC
        If nInputAlign = TDIBA_Content Then
            If m_Icon Then
                lOffX = (32 * m_ScaleX) + 20    '52
            Else
                lOffX = 12
            End If
            If (uTDC.dwFlags And TDF_KILL_SHIELD_ICON) And (m_lNewIcon = 0) Then
                lOffX = 12
            End If
            lEditX = lOffX
            If (uTDC.dwFlags And TDF_USE_COMMAND_LINKS) Or (uTDC.dwFlags And TDF_USE_COMMAND_LINKS_NO_ICON) Or (m_HasRadioButtons = True) Then
                lEditY = lButtonY - (34 * m_ScaleY)
            Else
                'If ((Right$(m_sContent, 4) = vbCrLf & vbCrLf)) Then
                lEditY = lButtonY - (44 * m_ScaleY)
                'Else
                '    lEditY = lButtonY - (34 * m_ScaleY)
                'End If
                If bLinkAdjust Then
                    lEditY = lEditY - (6 * m_ScaleY)
                End If
            End If
            If (uTDC.dwFlags And TDF_SHOW_MARQUEE_PROGRESS_BAR) Or (uTDC.dwFlags And TDF_SHOW_PROGRESS_BAR) Then
                lEditY = lEditY - (22 * m_ScaleY)
            End If
            If (m_sExpandedInfo <> "") Then
                lEditY = lEditY - (6 * m_ScaleY)
            End If
            cxEdit = tRC.Right - lOffX - 12
        ElseIf nInputAlign = TDIBA_Footer Then
            '        lEditX = 32
            '        lEditY = tRC.Bottom - 27
            '        cxEdit = 180
            If cxInputBox = -1 Then
                cxEdit = tRC.Right - ((16 * m_ScaleX) + 32)    '48
            Else
                cxEdit = 180 * m_ScaleX
            End If
            If cxInputBox > 0 Then cxEdit = cxInputBox
            Select Case nIBFooterAlign

                Case tdcaLeft
                    lEditX = (16 * m_ScaleX) + 16    '32
                Case tdcaCenter
                    lEditX = ((16 * m_ScaleX) + 16) + (((tRC.Right - ((16 * m_ScaleX) + 16) - 10) - (cxEdit)) \ 2)
                Case tdcaRight
                    lEditX = tRC.Right - (cxEdit) - 10
            End Select    'lEditY = tRC.Bottom - 27
            If (m_sVerify <> "") And (m_sExpandedInfo <> "") Then
                'TODO: ADJUST FOR EXPAND-TO-FOOTER
                lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4)    'lButtonY3 + (23 * m_ScaleY) + 22 '58
            Else
                lEditY = tRC.Bottom - ((24 * m_ScaleY) + 4)    'lButtonY3 + (23 * m_ScaleY) + 17 'lButtonY3 + (40 * m_ScaleY)
            End If

        End If
        If (nInputAlign = TDIBA_Content) Or (nInputAlign = TDIBA_Footer) Then
            SetWindowPos hEditBox, 0&, lEditX, lEditY, cxEdit, 22&, SWP_NOSIZE
            RepaintWindow hEditBox
        End If
        SetWindowPos m_TDhWnd, 0, 0, 0, tRC.Right, tRC.Bottom + (28 * m_ScaleY), SWP_NOMOVE
    End If
End Sub

Private Function MakeSplitButton(hWnd As LongPtr) As Long
'adds the drop down menu button to a command button
'Subclassing is required to detect a click
'However the PARENT received the BCN_DROPDOWN message

    Dim dwStyle                         As LongPtr

    dwStyle = GetWindowLong(hWnd, GWL_STYLE)

    dwStyle = dwStyle Or BS_SPLITBUTTON

    Call SetWindowLong(hWnd, GWL_STYLE, dwStyle)
    Call UpdateWindow(hWnd)

End Function

Public Function NavigatePage(cTD As cTaskDialog) As Long

    Dim I                               As Long
    Dim uTDC2                           As TASKDIALOGCONFIG
    Dim lBTN()                          As Long
    Dim lRBN()                          As Long
    Dim sBTN()                          As String
    Dim sRBN()                          As String
    Dim uBTN()                          As TASKDIALOG_BUTTON
    Dim uRBN()                          As TASKDIALOG_BUTTON
    Dim sz()                            As String

    If m_TDhWnd <> 0& Then
        ReDim sz(7)
        cTD.GetCustomButtons lBTN, sBTN
        cTD.GetRadioButtons lRBN, sRBN
        ReDim uBTN(0)
        ReDim uRBN(0)
        With uTDC2
            .cbSize = Len(uTDC2)
            If (UBound(lBTN) = 0) And (lBTN(0) = 0) Then
                .cButtons = 0
            Else
                .cButtons = UBound(lBTN) + 1
                ReDim uBTN(UBound(lBTN))
                For I = 0 To UBound(lBTN)
                    uBTN(I).nButtonID = lBTN(I)
                    #If Win64 Then
                        AssignLongPtr uBTN(I).pszButtonText, uBTN(I).pszButtonText2, StrPtr(sBTN(I))
                    #Else
                        uBTN(I).pszButtonText = StrPtr(sBTN(I))
                    #End If

                Next I
            End If
            #If Win64 Then
                AssignLongPtr .pButtons, .pButtons2, VarPtr(uBTN(0).nButtonID)
            #Else
                .pButtons = VarPtr(uBTN(0).nButtonID)
            #End If
            If (UBound(lRBN) = 0) And (lRBN(0) = 0) Then
                .cRadioButtons = 0
            Else
                .cRadioButtons = UBound(lRBN) + 1
                ReDim uRBN(UBound(lRBN))
                For I = 0 To UBound(lRBN)
                    uRBN(I).nButtonID = lRBN(I)
                    #If Win64 Then
                        AssignLongPtr uRBN(I).pszButtonText, uRBN(I).pszButtonText2, StrPtr(sRBN(I))
                    #Else
                        uRBN(I).pszButtonText = StrPtr(sRBN(I))
                    #End If
                Next I
            End If
            #If Win64 Then
                AssignLongPtr .pRadioButtons, .pRadioButtons2, VarPtr(uRBN(0).nButtonID)
            #Else
                .pRadioButtons = VarPtr(uRBN(0).nButtonID)
            #End If
            .cxWidth = cTD.Width
            .dwCommonButtons = cTD.CommonButtons
            .dwFlags = cTD.FLAGS
            .hInstance = cTD.hInst
            .hwndParent = cTD.ParenthWnd
            .nDefaultButton = cTD.DefaultButton
            .nDefaultRadioButton = cTD.DefaultRadioButton

            sz(0) = cTD.CollapsedControlText
            sz(1) = cTD.Content
            sz(2) = cTD.ExpandedControlText
            sz(3) = cTD.ExpandedInfo
            sz(4) = cTD.Footer
            sz(5) = cTD.MainInstruction
            sz(6) = cTD.VerifyText
            sz(7) = cTD.Title


            If ((.dwFlags And TDF_INPUT_BOX) And (cTD.InputAlign = TDIBA_Footer)) Or _
               ((.dwFlags And TDF_COMBO_BOX) And (cTD.ComboAlign = TDIBA_Footer)) Or _
               ((.dwFlags And TDF_DATETIME) And (cTD.DateTimeAlign = TDIBA_Footer)) Or _
               ((.dwFlags And TDF_SLIDER) And (cTD.SliderAlign = TDIBA_Footer)) Then

                'ensure footer text isn't blank, or it won't appear
                If sz(4) = "" Then
                    sz(4) = "$input"
                End If
            End If
            bEditFooter = False
            If sz(4) = "$input" Then
                bEditFooter = True
                sz(4) = " "
                '    uTDC.pszFooter = StrPtr(m_sFooter)
            End If
            If ((.dwFlags And TDF_INPUT_BOX) And (cTD.InputAlign = TDIBA_Content)) Or _
               ((.dwFlags And TDF_COMBO_BOX) And (cTD.ComboAlign = TDIBA_Content)) Or _
               ((.dwFlags And TDF_DATETIME) And (cTD.DateTimeAlign = TDIBA_Content)) Or _
               ((.dwFlags And TDF_SLIDER) And (cTD.SliderAlign = TDIBA_Content)) Then
                If sz(1) <> "" Then
                    'If Right$(m_sContent, 3) <> (vbCrLf & " ") Then
                    '                m_sContent = m_sContent & vbCrLf & " "
                    '                nBreaksAdded = 1
                    'If (uTDC.dwFlags And TDF_SLIDER) Then 'double-break is mandatory for slider
                    If Right$(sz(1), 5) <> (vbCrLf & vbCrLf & " ") Then
                        sz(1) = sz(1) & vbCrLf & vbCrLf
                        If (.dwFlags And TDF_ENABLE_HYPERLINKS) Then
                            sz(1) = sz(1) & " "  '(link control requires a space or it ignores the double-break)
                            'Meanwhile, everything else gets thrown out of alignment, so ONLY add to link
                        End If
                        nBreaksAdded = 2
                        '                        uTDC.pszContent = StrPtr(m_sContent)
                    End If
                    'End If
                    'End If
                Else    'No content, main instruction only: add one break
                    If Right$(sz(5), 2) <> vbCrLf Then
                        sz(5) = sz(5) & vbCrLf
                        '                uTDC.pszMainInstruction = StrPtr(m_sMainInstr)
                    End If

                End If
                If sz(3) <> "" Then
                    If Right$(sz(3), 2) <> vbCrLf Then
                        sz(3) = sz(3) & vbCrLf
                        'If (uTDC.dwFlags And TDF_SLIDER) Then 'double-break is mandatory for slider
                        If Right$(sz(3), 4) <> (vbCrLf & vbCrLf) Then
                            sz(3) = sz(3) & vbCrLf  '(it's already sure to have 1)
                        End If
                        'End If
                        '                uTDC.pszExpandedInformation = StrPtr(m_sExpandedInfo)
                    End If

                End If
            End If

            If (.dwFlags And TDF_CALLBACK_TIMER) Then
                bPublicTimer = True
            Else
                If (.dwFlags And TDF_INPUT_BOX) Or (.dwFlags And TDF_COMBO_BOX) Or (.dwFlags And TDF_DATETIME) Or (.dwFlags And TDF_SLIDER) Then
                    If m_sExpandedInfo <> "" Then
                        If (.dwFlags And TDF_EXPAND_FOOTER_AREA) <> TDF_EXPAND_FOOTER_AREA Then
                            .dwFlags = .dwFlags Or TDF_CALLBACK_TIMER
                        End If
                    End If
                End If
                If (lAutoclose > 0) Then
                    .dwFlags = .dwFlags Or TDF_CALLBACK_TIMER
                End If
            End If

            #If Win64 Then
                AssignLongPtr .pszCollapsedControlText, .pszCollapsedControlText2, StrPtr(sz(0))
                AssignLongPtr .pszContent, .pszContent2, StrPtr(sz(1))
                AssignLongPtr .pszExpandedControlText, .pszExpandedControlText2, StrPtr(sz(2))
                AssignLongPtr .pszExpandedInformation, .pszExpandedInformation2, StrPtr(sz(3))
                AssignLongPtr .pszFooter, .pszFooter2, StrPtr(sz(4))
                AssignLongPtr .pszMainInstruction, .pszMainInstruction2, StrPtr(sz(5))
                AssignLongPtr .pszVerificationText, .pszVerificationText2, StrPtr(sz(6))
                AssignLongPtr .pszWindowTitle, .pszWindowTitle2, StrPtr(sz(7))

            #Else
                .pszCollapsedControlText = StrPtr(sz(0))
                .pszContent = StrPtr(sz(1))
                .pszExpandedControlText = StrPtr(sz(2))
                .pszExpandedInformation = StrPtr(sz(3))
                .pszFooter = StrPtr(sz(4))
                .pszMainInstruction = StrPtr(sz(5))
                .pszVerificationText = StrPtr(sz(6))
                .pszWindowTitle = StrPtr(sz(7))
            #End If

            If (.dwFlags And TDF_USE_HICON_MAIN) Then
                .pszMainIcon = cTD.IconMain
            Else
                .pszMainIcon = tdMakeIntResource(cTD.IconMain)
            End If
            If (.dwFlags And TDF_USE_HICON_FOOTER) Then
                .pszFooterIcon = cTD.IconFooter
            Else
                .pszFooterIcon = tdMakeIntResource(cTD.IconFooter)
            End If

            #If Win64 Then
                AssignLongPtr uTDC.lpCallbackData, uTDC.lpCallbackData2, ObjPtr(Me)
            #Else
                .pfCallback = tdFARPROC(AddressOf TaskDialogCallbackProc)
                .lpCallbackData = ObjPtr(cTD)
            #End If
        End With

        Call SendMessageW(m_TDhWnd, TDM_NAVIGATE_PAGE, 0, uTDC2)
        m_TDhWnd = 0    'This TaskDialog is now the old page and is gone. if recreated, the hwnd would be reassigned
    End If
End Function

Public Function ProcessCallback(hWnd As Long, uNotification As Long, wParam As Long, lParam As LongPtr) As Long
'***DO NOT CALL THIS FUNCTION***
    Dim bSC                             As Boolean
    Dim lLen                            As LongPtr
    Dim sText                           As String
    Dim RS                              As LongPtr

    Select Case uNotification
        Case TDN_DIALOG_CONSTRUCTED
            If (uTDC.dwFlags And TDF_KILL_SHIELD_ICON) Then
                SendMessageW hWnd, TDM_UPDATE_ICON, TDIE_ICON_MAIN, ByVal m_lNewIcon
            End If
            RaiseEvent DialogConstucted(hWnd)
        Case TDN_CREATED
            m_TDhWnd = hWnd
            dtAC1 = Now
            ReDim m_hButton(0)
            lButtonX = 0
            lButtonY = 0
            lButtonX2 = 0
            lButtonY2 = 0
            lButtonX3 = 0
            lButtonY3 = 0
            lButtonX4 = 0
            lButtonY4 = 0
            ReDim arBtnX(0)

            Call EnumChildWindows(hWnd, AddressOf TaskDialogEnumChildProc, ObjPtr(Me))
            SetButtonIcons
            SetInitialButtonStates
            If (m_sExpandedInfo <> "") Then bPosUsed(1) = True    'expando control taking up the button position, so we can't add ours
            If (m_sVerify <> "") Then bPosUsed(1) = True    'same for verify
            If (uTDC.dwFlags And TDF_INPUT_BOX) Then AddInputBox: bSC = True
            If (uTDC.dwFlags And TDF_COMBO_BOX) Then AddComboBox nCbType
            If (uTDC.dwFlags And TDF_DATETIME) Then AddDateTime: bSC = True
            If (uTDC.dwFlags And TDF_SLIDER) Then AddSlider: bSC = True
            If (hLogoImage) Then AddLogo: bSC = True
            If m_HasCustomControl = True Then FocusMainControl
            If bSC Then    'combo box subclasses different hWnds in the Add func
                Subclass hUI, AddressOf TaskDialogSubclassProc, m_TDhWnd, ObjPtr(Me)
            End If
            RaiseEvent DialogCreated(hWnd)
        Case TDN_HYPERLINK_CLICKED
            If uTDC.dwFlags And TDF_EXEC_HYPERLINKS Then
                Call ShellExecuteW(0, 0, lParam, 0, 0, 1)
            End If
            ProcessCallback = S_OK
            RaiseEvent HyperlinkClick(lParam)
            Exit Function
        Case TDN_BUTTON_CLICKED
            If ButtonIsHeld(wParam) Then
                ProcessCallback = S_FALSE
            Else
                ProcessCallback = S_OK
            End If
            RaiseEvent ButtonClick(wParam)
            Exit Function
        Case TDN_TIMER
            dtAC2 = Now
            If (lAutoclose > 0) And (DateDiff("s", dtAC1, dtAC2) > lAutoclose) Then
                RaiseEvent AutoClose
                CloseDialog
            Else
                lAutocloseTrack = lAutoclose - DateDiff("s", dtAC1, dtAC2)
            End If
            If bReposTrigger Then
                If nTmrCnt = 5 Then
                    nTmrCnt = 0
                    Dim t3              As RECT
                    GetClientRect m_TDhWnd, t3
                    If t3.Bottom <> lLastCY Then
                        bReposTrigger = False
                        InputReposition
                        ComboReposition
                        DateTimeReposition
                        SliderReposition
                    End If
                Else
                    nTmrCnt = nTmrCnt + 1
                End If
            End If

            If bPublicTimer Then
                RaiseEvent Timer(wParam)
                If bResetTimer Then
                    bResetTimer = False
                    ProcessCallback = S_FALSE
                Else
                    ProcessCallback = S_OK
                End If
                Exit Function
            End If
        Case TDN_RADIO_BUTTON_CLICKED
            m_ResRad = wParam    'for when we need result before dialog is closed
            RaiseEvent RadioButtonClick(wParam)
            Exit Function
        Case TDN_VERIFICATION_CLICKED
            m_ResVerify = wParam
            RaiseEvent VerificationClicked(wParam)
            Exit Function
        Case TDN_HELP
            RaiseEvent Help
            Exit Function
        Case TDN_EXPANDO_BUTTON_CLICKED
            If (uTDC.dwFlags And TDF_INPUT_BOX) Or (uTDC.dwFlags And TDF_COMBO_BOX) Or (uTDC.dwFlags And TDF_DATETIME) Or (uTDC.dwFlags And TDF_SLIDER) Then
                If (uTDC.dwFlags And TDF_EXPAND_FOOTER_AREA) <> TDF_EXPAND_FOOTER_AREA Then
                    Dim t2              As RECT
                    GetClientRect m_TDhWnd, t2
                    lLastCY = t2.Bottom
                    bReposTrigger = True
                    If wParam Then
                        '                    Debug.Print "Expando wParam 1; nba=" & nBreaksAdded
                        If nBreaksAdded = 1 Then
                            '                        Debug.Print "Remove single"
                            Me.Content = Left$(m_sContent, Len(m_sContent) - 2)
                        ElseIf nBreaksAdded = 2 Then
                            '                        Debug.Print "Remove double"
                            Me.Content = Left$(m_sContent, Len(m_sContent) - 4)
                        End If
                    Else
                        '                    Debug.Print "Expando wParam 0; nba=" & nBreaksAdded
                        If nBreaksAdded = 1 Then
                            Me.Content = m_sContent & vbCrLf
                        ElseIf nBreaksAdded = 2 Then
                            '                        Debug.Print "add double"
                            If (uTDC.dwFlags And TDF_ENABLE_HYPERLINKS) Then
                                Me.Content = m_sContent & vbCrLf & vbCrLf & " "  '(link control requires a space or it ignores the double-break)
                            Else
                                Me.Content = m_sContent & vbCrLf & vbCrLf
                            End If


                        End If
                    End If
                End If
            End If

            RaiseEvent ExpandButtonClicked(wParam)
            Exit Function
        Case TDN_NAVIGATED
            m_TDhWnd = hWnd    'new dialog treated as page means no TDN_CREATED
            ReDim m_hButton(0)
            dtAC1 = Now
            lButtonX = 0
            lButtonY = 0
            lButtonX2 = 0
            lButtonY2 = 0
            lButtonX3 = 0
            lButtonY3 = 0
            lButtonX4 = 0
            lButtonY4 = 0
            ReDim arBtnX(0)

            Call EnumChildWindows(hWnd, AddressOf TaskDialogEnumChildProc, ObjPtr(Me))


            SetButtonIcons
            SetInitialButtonStates
            If (m_sExpandedInfo <> "") Then bPosUsed(1) = True    'expando control taking up the button position, so we can't add ours
            If (m_sVerify <> "") Then bPosUsed(1) = True    'same for verify
            If (uTDC.dwFlags And TDF_INPUT_BOX) Then AddInputBox: bSC = True
            If (uTDC.dwFlags And TDF_COMBO_BOX) Then AddComboBox nCbType
            If (uTDC.dwFlags And TDF_DATETIME) Then AddDateTime: bSC = True
            If (uTDC.dwFlags And TDF_SLIDER) Then AddSlider: bSC = True
            If (hLogoImage) Then AddLogo
            If m_HasCustomControl = True Then FocusMainControl
            If bSC Then    'combo box subclasses different hWnds in the Add func

                Subclass hUI, AddressOf TaskDialogSubclassProc, m_TDhWnd, ObjPtr(Me)
            End If
            RaiseEvent Navigated
            Exit Function
        Case TDN_DESTROYED
            If hEditBox Then
                lLen = SendMessageW(hEditBox, WM_GETTEXTLENGTH, 0, ByVal 0&) * 2
                If lLen Then
                    sText = Space$(CLng(lLen))    '//KCB added CLng
                    Call SendMessageW(hEditBox, WM_GETTEXT, lLen, ByVal sText)
                    sEditText = StrConv(sText, vbFromUnicode)
                End If
            End If
            If hCombo Then
                sComboText = GetComboTextW(hEditCombo)
                nComboIdx = SendMessageW(hCombo, CB_GETCURSEL, 0, ByVal 0&)
            End If
            If hDateTime Then
                RS = SendMessageW(hDateTime, DTM_GETSYSTEMTIME, 0&, tDTCur)
                If RS = GDT_VALID Then
                    bDTCheck = True
                Else
                    bDTCheck = False
                End If
            End If
            If hDateTime2 Then
                RS = SendMessageW(hDateTime2, DTM_GETSYSTEMTIME, 0&, tDTCur2)
                If RS = GDT_VALID Then
                    bDTCheck2 = True
                Else
                    bDTCheck2 = False
                End If
            End If
            If hSlider Then
                lSliderValue = SendMessageW(hSlider, TBM_GETPOS, 0, ByVal 0&)
            End If
            RaiseEvent DialogDestroyed
            m_TDhWnd = 0
    End Select
End Function

Public Function ProcessEnumCallback(hWnd As LongPtr) As Long
    Dim ub                              As Long
    Dim sClass                          As String
    Dim np                              As Long
    Dim tRC                             As RECT
    Dim pt                              As POINTAPI

    sClass = Space$(64)
    Call GetClassName(hWnd, sClass, 64)
    np = InStr(sClass, Chr$(0))
    If np > 1 Then
        sClass = Left(sClass, np - 1)
    End If
    If sClass = "DirectUIHWND" Then
        hUI = hWnd
    End If
    If sClass = "Button" Then
        GetWindowRect hWnd, tRC
        pt.x = tRC.Left
        pt.y = tRC.Top
        ScreenToClient m_TDhWnd, pt
        If arBtnX(0) = 0 Then
            arBtnX(0) = pt.x
        Else
            ReDim Preserve arBtnX(UBound(arBtnX) + 1)
            arBtnX(UBound(arBtnX)) = pt.x
        End If
        '        Debug.Print "Button top=" & pt.Y
        If True Then
            'If (uTDC.dwFlags And TDF_USE_COMMAND_LINKS) Or (uTDC.dwFlags And TDF_USE_COMMAND_LINKS_NO_ICON) Then
            If nInputAlign = TDIBA_Content Then
                If lButtonY = 0 Then
                    lButtonY = pt.y
                    lButtonX = pt.x
                End If
                If (pt.y > 10) And ((pt.y < lButtonY) And (lButtonY <> 0)) Then
                    lButtonY = pt.y
                    lButtonX = pt.x
                End If
            ElseIf (nInputAlign = TDIBA_Buttons) Or (nInputAlign = TDIBA_Footer) Then
                If (pt.y > lButtonY) Then
                    lButtonX = pt.x
                    lButtonY = pt.y
                End If
            End If
            If nComboAlign = TDIBA_Content Then
                If lButtonY2 = 0 Then
                    lButtonY2 = pt.y
                    lButtonX2 = pt.x
                End If
                If (pt.y > 10) And ((pt.y < lButtonY2) And (lButtonY2 <> 0)) Then
                    lButtonY2 = pt.y
                    lButtonX2 = pt.x
                End If
            ElseIf (nComboAlign = TDIBA_Buttons) Or (nComboAlign = TDIBA_Footer) Then
                If (pt.y > lButtonY2) Then
                    lButtonX2 = pt.x
                    lButtonY2 = pt.y
                End If
            End If
            If nDateTimeAlign = TDIBA_Content Then
                If lButtonY3 = 0 Then
                    lButtonY3 = pt.y
                    lButtonX3 = pt.x
                End If
                If (pt.y > 10) And ((pt.y < lButtonY3) And (lButtonY3 <> 0)) Then
                    lButtonY3 = pt.y
                    lButtonX3 = pt.x
                End If
            ElseIf (nDateTimeAlign = TDIBA_Buttons) Or (nDateTimeAlign = TDIBA_Footer) Then
                If (pt.y > lButtonY3) Then
                    lButtonX3 = pt.x
                    lButtonY3 = pt.y
                End If
            End If
            If nSliderAlign = TDIBA_Content Then
                If lButtonY4 = 0 Then
                    lButtonY4 = pt.y
                    lButtonX4 = pt.x
                End If
                If (pt.y > 10) And ((pt.y < lButtonY4) And (lButtonY4 <> 0)) Then
                    lButtonY4 = pt.y
                    lButtonX4 = pt.x
                End If
            ElseIf (nSliderAlign = TDIBA_Buttons) Or (nSliderAlign = TDIBA_Footer) Then
                If (pt.y > lButtonY4) Then
                    lButtonX4 = pt.x
                    lButtonY4 = pt.y
                End If
            End If
        Else
            If pt.y > lButtonY Then
                lButtonX = pt.x
                lButtonY = pt.y
            End If
            If pt.y > lButtonY2 Then
                lButtonX2 = pt.x
                lButtonY2 = pt.y
            End If
            If pt.y > lButtonY3 Then
                lButtonX3 = pt.x
                lButtonY3 = pt.y
            End If
            If pt.y > lButtonY4 Then
                lButtonX4 = pt.x
                lButtonY4 = pt.y
            End If
        End If
        ub = UBound(m_hButton)
        If (ub = 0) And (m_hButton(0) = 0) Then
            m_hButton(0) = hWnd
        Else
            ub = ub + 1
            ReDim Preserve m_hButton(ub)
            m_hButton(ub) = hWnd
        End If
    End If

    ProcessEnumCallback = 1
End Function

Public Function ProcessSubclass(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As LongPtr, ByVal uIdSubclass As Long) As LongPtr
    Dim hdc1                            As LongPtr
    Dim pxl                             As Long


    On Error GoTo e0
    'Debug.Print "TDSC uMsg=" & CStr(uMsg)
    Select Case uMsg
        Case WM_CTLCOLORSTATIC
            If (lParam = hStatic) And (hStatic <> 0&) Then
                hdc1 = GetDC(hUI)
                If nLogoPos = LogoTopRight Then
                    pxl = GetPixel(hdc1, lgpX - 1, 0)
                ElseIf nLogoPos = LogoButtons Then
                    pxl = GetPixel(hdc1, lgpX, lgpY - 1)    '109)
                End If
                '            Debug.Print "pxl(" & lgpX & "," & lgpY & ")=0x" & Hex$(pxl) & ",y=" & lButtonY

                ProcessSubclass = CreateSolidBrush(pxl)


                Exit Function
            End If

        Case BCM_SETDROPDOWNSTATE
            If wParam = 1 Then
                RaiseEvent DropdownButtonClicked(hWnd)
            End If
        Case WM_HSCROLL
            If lParam = hSlider Then RaiseEvent SliderChange(SliderValue)
        Case WM_COMMAND
            '        Debug.Print "WM_COMMAND " & tdHiWord(wParam)
            Select Case tdHiWord(wParam)

                Case CBN_SELCHANGE
                    RaiseEvent ComboItemChanged(ComboIndex)
                Case CBN_DROPDOWN
                    RaiseEvent ComboDropdown
                Case EN_CHANGE
                    RaiseEvent InputBoxChange(InputText)
            End Select
        Case WM_NOTIFY
            Dim NM                      As NMHDR
            CopyMemory NM, ByVal lParam, LenB(NM)
            '            Debug.Print "WM_NOTIFY code=" & NM.Code
            Select Case NM.code
                Case DTN_DATETIMECHANGE
                    RaiseEvent DateTimeChange(DateTimeValue, DateTimeChecked)
            End Select
            '      Exit Function
        Case WM_DESTROY
            Call UnSubclass(hWnd, PtrTDWndProc, uIdSubclass)

    End Select

    ProcessSubclass = DefSubclassProc(hWnd, uMsg, wParam, lParam)
    Exit Function
e0:
Debug.Print "ProcessSubclass Error=" & Err.Description

End Function

Public Sub ProgressSetRange(lMin As Long, lMax As Long)
    Dim dwRange                         As Long
    dwRange = tdMAKELPARAM(lMin, lMax)
    If m_TDhWnd <> 0& Then
        Call SendMessageW(m_TDhWnd, TDM_SET_PROGRESS_BAR_RANGE, 0, ByVal dwRange)
    End If
End Sub

Public Sub ProgressSetState(lState As eProgressBarStates)
    If m_TDhWnd <> 0& Then
        Call SendMessageW(m_TDhWnd, TDM_SET_PROGRESS_BAR_STATE, lState, ByVal 0&)
    End If
End Sub

Public Sub ProgressSetType(lType As Long)
'0=regular
'1=marquee
    If m_TDhWnd <> 0& Then
        Call SendMessageW(m_TDhWnd, TDM_SET_MARQUEE_PROGRESS_BAR, lType, ByVal 0&)
    End If
End Sub

Public Sub ProgressSetValue(lValue As Long)
    If m_TDhWnd <> 0& Then
        Call SendMessageW(m_TDhWnd, TDM_SET_PROGRESS_BAR_POS, lValue, ByVal 0&)
    End If
End Sub

Public Sub ProgressStartMarquee(Optional dwSpeed As Long = 0)
    If m_TDhWnd <> 0& Then
        Call SendMessageW(m_TDhWnd, TDM_SET_PROGRESS_BAR_MARQUEE, 1, ByVal dwSpeed)
    End If
End Sub

Public Sub ProgressStopMarquee()
    If m_TDhWnd <> 0& Then
        Call SendMessageW(m_TDhWnd, TDM_SET_PROGRESS_BAR_MARQUEE, 0, ByVal 0&)
    End If
End Sub

Private Function PtrTDWndProc() As Long
    #If Win64 Then
        'ToDo: What to do with x64
    #Else
        PtrTDWndProc = tdFARPROC(AddressOf TaskDialogSubclassProc)
    #End If
End Function

Public Sub ReleaseButtonHold(ButtonID As Long)
    Dim lnew()                          As Long
    Dim I                               As Long
    Dim J                               As Long
    ReDim lnew(0)
    
    For I = 0 To UBound(lBtnHold)
        If lBtnHold(I) <> ButtonID Then
            ReDim Preserve lnew(J)
            lnew(J) = lBtnHold(I)
            J = J + 1
        End If
    Next I
    lBtnHold = lnew
End Sub

Private Sub RepaintWindow(hWnd As LongPtr)
    Dim tR                              As RECT
    Dim tP                              As POINTAPI
    GetClientRect hWnd, tR
    InvalidateRect hWnd, tR, 1
End Sub

Public Sub ResetTimer()
    bResetTimer = True
End Sub

Public Sub SetButtonElevated(ButtonID As Long, lState As Long)
    Dim lnew()                          As Long
    Dim I                               As Long
    Dim k                               As Long

    If m_TDhWnd <> 0& Then
        Call SendMessageW(m_TDhWnd, TDM_SET_BUTTON_ELEVATION_REQUIRED_STATE, ButtonID, ByVal lState)
    Else
        If lState Then
            If (lBtnElv(0) = 0) And (UBound(lBtnElv) = 0) Then
                lBtnElv(0) = ButtonID
            Else
                ReDim Preserve lBtnElv(UBound(lBtnElv) + 1)
                lBtnElv(UBound(lBtnElv)) = ButtonID
            End If
        Else
            ReDim lnew(0)
            For I = 0 To UBound(lBtnElv)
                If lBtnElv(I) <> ButtonID Then
                    ReDim Preserve lnew(k)
                    lnew(k) = lBtnElv(I)
                End If
            Next I
            lBtnElv = lnew
        End If
        If (UBound(lBtnElv) = 0) And (lBtnElv(0) = 0) Then
            bElv = False
        Else
            bElv = True
        End If
    End If
End Sub

Public Sub SetButtonHold(ButtonID As Long)
    If (lBtnHold(0) = 0) And (UBound(lBtnHold) = 0) Then
        lBtnHold(0) = ButtonID
    Else
        ReDim Preserve lBtnHold(UBound(lBtnHold) + 1)
        lBtnHold(UBound(lBtnHold)) = ButtonID
    End If
End Sub

Private Function SetButtonIcons()
    Dim I                               As Long
    Dim J                               As Long
    Dim k                               As Long
    Dim lOffset                         As Long
    Dim hButton                         As LongPtr
    Dim stay                            As Boolean
    Dim hIco                            As LongPtr
    Dim sID                             As String

    sID = "23"
    On Error Resume Next
    If m_HasRadioButtons Then
        lOffset = UBound(m_uRadioButtons) + 1
    End If

    If m_HasCustomButtons Then
        For I = 0 To UBound(icoDataC)
            hButton = m_hButton(I + lOffset)
            If icoDataC(I) Then
                If (uTDC.dwFlags And TDF_USE_SHELL32_ICONID_BUTTON) = TDF_USE_SHELL32_ICONID_BUTTON Then
                    hIco = LoadImage(hShell32, icoDataC(I), 1, m_BtnLRSize, m_BtnLRSize, 0)
                    Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
                    Call DestroyIcon(hIco)
                ElseIf (uTDC.dwFlags And TDF_USE_IMAGERES_ICONID_BUTTON) = TDF_USE_IMAGERES_ICONID_BUTTON Then
                    hIco = LoadImage(hImageRes, icoDataC(I), 1, m_BtnLRSize, m_BtnLRSize, 0)
                    Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
                    Call DestroyIcon(hIco)
                Else
                    hIco = icoDataC(I)
                    Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
                End If

            End If
            If m_uButtons(I).nButtonID = nSplitButtonID Then

                Call Subclass(hButton, AddressOf TaskDialogSubclassProc, m_TDhWnd, ObjPtr(Me))

                '                SetParent hButton, m_TDhWnd
                MakeSplitButton hButton
            End If
        Next I
        J = UBound(m_uButtons) + 1
    End If

    For I = 0 To 5
        If icoDataN(I) Then
            stay = True
        End If
    Next I
    If (stay = False) Or (uTDC.dwCommonButtons = 0) Then Exit Function
    I = 0
    If (uTDC.dwCommonButtons And TDCBF_OK_BUTTON) Then
        hButton = m_hButton(J + lOffset)
        If icoDataN(0) Then
            If (uTDC.dwFlags And TDF_USE_SHELL32_ICONID_BUTTON) = TDF_USE_SHELL32_ICONID_BUTTON Then
                hIco = LoadImage(hShell32, icoDataN(0), 1, m_BtnLRSize, m_BtnLRSize, 0)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
                Call DestroyIcon(hIco)
            ElseIf (uTDC.dwFlags And TDF_USE_IMAGERES_ICONID_BUTTON) = TDF_USE_IMAGERES_ICONID_BUTTON Then
                hIco = LoadImage(hImageRes, icoDataN(0), 1, m_BtnLRSize, m_BtnLRSize, 0)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
                Call DestroyIcon(hIco)
            Else
                hIco = icoDataN(0)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
            End If
            '        Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal icoDataN(0))
        End If
        k = k + 1
    End If
    If (uTDC.dwCommonButtons And TDCBF_YES_BUTTON) Then
        hButton = m_hButton(J + lOffset + k)
        If icoDataN(1) Then
            If (uTDC.dwFlags And TDF_USE_SHELL32_ICONID_BUTTON) = TDF_USE_SHELL32_ICONID_BUTTON Then
                hIco = LoadImage(hShell32, icoDataN(1), 1, m_BtnLRSize, m_BtnLRSize, 0)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
                Call DestroyIcon(hIco)
            ElseIf (uTDC.dwFlags And TDF_USE_IMAGERES_ICONID_BUTTON) = TDF_USE_IMAGERES_ICONID_BUTTON Then
                hIco = LoadImage(hImageRes, icoDataN(1), 1, m_BtnLRSize, m_BtnLRSize, 0)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
                Call DestroyIcon(hIco)
            Else
                hIco = icoDataN(1)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
            End If
        End If
        k = k + 1
    End If
    If (uTDC.dwCommonButtons And TDCBF_NO_BUTTON) Then
        hButton = m_hButton(J + lOffset + k)
        If icoDataN(2) Then
            If (uTDC.dwFlags And TDF_USE_SHELL32_ICONID_BUTTON) = TDF_USE_SHELL32_ICONID_BUTTON Then
                hIco = LoadImage(hShell32, icoDataN(2), 1, m_BtnLRSize, m_BtnLRSize, 0)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
                Call DestroyIcon(hIco)
            ElseIf (uTDC.dwFlags And TDF_USE_IMAGERES_ICONID_BUTTON) = TDF_USE_IMAGERES_ICONID_BUTTON Then
                hIco = LoadImage(hImageRes, icoDataN(2), 1, m_BtnLRSize, m_BtnLRSize, 0)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
                Call DestroyIcon(hIco)
            Else
                hIco = icoDataN(2)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
            End If
        End If
        k = k + 1
    End If
    If (uTDC.dwCommonButtons And TDCBF_CANCEL_BUTTON) Then
        hButton = m_hButton(J + lOffset + k)
        If icoDataN(3) Then
            If (uTDC.dwFlags And TDF_USE_SHELL32_ICONID_BUTTON) = TDF_USE_SHELL32_ICONID_BUTTON Then
                hIco = LoadImage(hShell32, icoDataN(3), 1, m_BtnLRSize, m_BtnLRSize, 0)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
                Call DestroyIcon(hIco)
            ElseIf (uTDC.dwFlags And TDF_USE_IMAGERES_ICONID_BUTTON) = TDF_USE_IMAGERES_ICONID_BUTTON Then
                hIco = LoadImage(hImageRes, icoDataN(3), 1, m_BtnLRSize, m_BtnLRSize, 0)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
                Call DestroyIcon(hIco)
            Else
                hIco = icoDataN(3)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
            End If
        End If
        k = k + 1
    End If
    If (uTDC.dwCommonButtons And TDCBF_RETRY_BUTTON) Then
        hButton = m_hButton(J + lOffset + k)
        If icoDataN(4) Then
            If (uTDC.dwFlags And TDF_USE_SHELL32_ICONID_BUTTON) = TDF_USE_SHELL32_ICONID_BUTTON Then
                hIco = LoadImage(hShell32, icoDataN(4), 1, m_BtnLRSize, m_BtnLRSize, 0)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
                Call DestroyIcon(hIco)
            ElseIf (uTDC.dwFlags And TDF_USE_IMAGERES_ICONID_BUTTON) = TDF_USE_IMAGERES_ICONID_BUTTON Then
                hIco = LoadImage(hImageRes, icoDataN(4), 1, m_BtnLRSize, m_BtnLRSize, 0)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
                Call DestroyIcon(hIco)
            Else
                hIco = icoDataN(4)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
            End If
        End If
        k = k + 1
    End If
    If (uTDC.dwCommonButtons And TDCBF_CLOSE_BUTTON) Then
        hButton = m_hButton(J + lOffset + k)
        If icoDataN(5) Then
            If (uTDC.dwFlags And TDF_USE_SHELL32_ICONID_BUTTON) = TDF_USE_SHELL32_ICONID_BUTTON Then
                hIco = LoadImage(hShell32, icoDataN(5), 1, m_BtnLRSize, m_BtnLRSize, 0)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
                Call DestroyIcon(hIco)
            ElseIf (uTDC.dwFlags And TDF_USE_IMAGERES_ICONID_BUTTON) = TDF_USE_IMAGERES_ICONID_BUTTON Then
                hIco = LoadImage(hImageRes, icoDataN(5), 1, m_BtnLRSize, m_BtnLRSize, 0)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
                Call DestroyIcon(hIco)
            Else
                hIco = icoDataN(5)
                Call SendMessageW(hButton, BM_SETIMAGE, 1, ByVal hIco)
            End If
        End If
        k = k + 1
    End If
End Function

Public Sub SetCommonButtonIcon(id As TDBUTTONS, hIcon As Long)
    Select Case id
        Case TDCBF_OK_BUTTON: icoDataN(0) = hIcon
        Case TDCBF_YES_BUTTON: icoDataN(1) = hIcon
        Case TDCBF_NO_BUTTON: icoDataN(2) = hIcon
        Case TDCBF_CANCEL_BUTTON: icoDataN(3) = hIcon
        Case TDCBF_RETRY_BUTTON: icoDataN(4) = hIcon
        Case TDCBF_CLOSE_BUTTON: icoDataN(5) = hIcon
    End Select
End Sub

Private Sub SetDTRange()
    Dim dwValid                         As Long
    If bDTSetMin Then dwValid = GDTR_MIN
    If bDTSetMax Then dwValid = dwValid Or GDTR_MAX
    If hDateTime Then
        If (nDTType = dttTime) Or (nDTType = dttTimeWithCheck) Then
            tDTRange(1).wDay = tDTRange(0).wDay
            tDTRange(1).wMonth = tDTRange(0).wMonth
            tDTRange(1).wYear = tDTRange(0).wYear
        End If
        Call SendMessageW(hDateTime, DTM_SETRANGE, dwValid, ByVal VarPtr(tDTRange(0)))
    End If
    If hDateTime2 Then
        tDTRange(1).wDay = tDTRange(0).wDay
        tDTRange(1).wMonth = tDTRange(0).wMonth
        tDTRange(1).wYear = tDTRange(0).wYear
        'fixes a bug.. if the date isn't set like that, applying it to a time control will make the max ignored
        Call SendMessageW(hDateTime2, DTM_SETRANGE, dwValid, ByVal VarPtr(tDTRange(0)))
    End If
End Sub

Private Sub SetInitialButtonStates()
    Dim I                               As Long
    If bElv Then
        For I = 0 To UBound(lBtnElv)
            If lBtnElv(I) <> 0& Then Call SendMessageW(m_TDhWnd, TDM_SET_BUTTON_ELEVATION_REQUIRED_STATE, lBtnElv(I), ByVal 1&)
        Next I
    End If
    If bDis Then
        For I = 0 To UBound(lBtnDis)
            If lBtnDis(I) <> 0& Then Call SendMessageW(m_TDhWnd, TDM_ENABLE_BUTTON, lBtnDis(I), ByVal 0&)
        Next I
    End If
    If bRadDis Then
        For I = 0 To UBound(lRadDis)
            If lRadDis(I) <> 0& Then Call SendMessageW(m_TDhWnd, TDM_ENABLE_RADIO_BUTTON, lRadDis(I), ByVal 0&)
        Next I
    End If
End Sub

Public Sub SetLogoImage(hImage As Long, ImageType As LogoImageType, ImagePosition As LogoPosition, Optional OffsetX As Long = 0, Optional OffsetY As Long = 0)
    hLogoImage = hImage
    nLogoType = ImageType
    nLogoPos = ImagePosition
    lLogoX = OffsetX
    lLogoY = OffsetY
End Sub

Public Sub SetSplitButton(idCustomButton As Long)
    nSplitButtonID = idCustomButton
End Sub

Public Sub SetWindowsButtonIconSize(SIZE As eWinIcoSize)
    m_BtnLRSize = SIZE
End Sub

Public Function ShowDialog() As Boolean
'main function; returns True if an option was selected, False if cancelled
'then use .

    Dim pnButton                        As Long
    Dim pnRadButton                     As Long
    Dim pfVerify                        As Boolean
    Dim hr                              As Long
    Dim sErr                            As String
    
    lAutoclose = lAutocloseTrack
    'sEditText = ""
    '//Set the Access App hWnd if not set
    If uTDC.hwndParent = 0 Then
        uTDC.hwndParent = Application.hWndAccessApp
    End If

    If ((uTDC.dwFlags And TDF_INPUT_BOX) And (nInputAlign = TDIBA_Footer)) Or _
       ((uTDC.dwFlags And TDF_COMBO_BOX) And (nComboAlign = TDIBA_Footer)) Or _
       ((uTDC.dwFlags And TDF_DATETIME) And (nDateTimeAlign = TDIBA_Footer)) Or _
       ((uTDC.dwFlags And TDF_SLIDER) And (nSliderAlign = TDIBA_Footer)) Then

        'ensure footer text isn't blank, or it won't appear
        If m_sFooter = "" Then
            m_sFooter = "$input"
        End If
    End If
    bEditFooter = False
    If m_sFooter = "$input" Then
        bEditFooter = True
        m_sFooter = " "
        #If Win64 Then
            AssignLongPtr uTDC.pszFooter, uTDC.pszFooter2, StrPtr(m_sFooter)
        #Else
            uTDC.pszFooter = StrPtr(m_sFooter)
        #End If
    End If
    If ((uTDC.dwFlags And TDF_INPUT_BOX) And (nInputAlign = TDIBA_Content)) Or _
       ((uTDC.dwFlags And TDF_COMBO_BOX) And (nComboAlign = TDIBA_Content)) Or _
       ((uTDC.dwFlags And TDF_DATETIME) And (nDateTimeAlign = TDIBA_Content)) Or _
       ((uTDC.dwFlags And TDF_SLIDER) And (nSliderAlign = TDIBA_Content)) Then
        If m_sContent <> "" Then
            'If Right$(m_sContent, 3) <> (vbCrLf & " ") Then
            '                m_sContent = m_sContent & vbCrLf & " "
            '                nBreaksAdded = 1
            'If (uTDC.dwFlags And TDF_SLIDER) Then 'double-break is mandatory for slider
            If Right$(m_sContent, 5) <> (vbCrLf & vbCrLf & " ") Then
                m_sContent = m_sContent & vbCrLf & vbCrLf
                If (uTDC.dwFlags And TDF_ENABLE_HYPERLINKS) Then
                    Dim lnkPos          As Long
                    lnkPos = (InStr(LCase$(m_sContent), "<a href"))
                    '                        Debug.Print "lnkPos=" & lnkPos
                    If lnkPos Then
                        bLinkAdjust = True
                        m_sContent = m_sContent & " "    ' (link control requires a space or it ignores the double-break)
                        'Meanwhile, everything else gets thrown out of alignment, so ONLY add to link
                    End If
                End If
                nBreaksAdded = 2
                #If Win64 Then
                    AssignLongPtr uTDC.pszContent, uTDC.pszContent2, StrPtr(m_sContent)
                #Else
                    uTDC.pszContent = StrPtr(m_sContent)
                #End If
            End If
            'End If
            'End If
        Else    'No content, main instruction only: add one break
            If Right$(m_sMainInstr, 2) <> vbCrLf Then
                m_sMainInstr = m_sMainInstr & vbCrLf
                #If Win64 Then
                    AssignLongPtr uTDC.pszMainInstruction, uTDC.pszMainInstruction2, StrPtr(m_sMainInstr)
                #Else
                    uTDC.pszMainInstruction = StrPtr(m_sMainInstr)
                #End If
            End If

        End If
        If m_sExpandedInfo <> "" Then
            If Right$(m_sExpandedInfo, 2) <> vbCrLf Then
                m_sExpandedInfo = m_sExpandedInfo & vbCrLf
                'If (uTDC.dwFlags And TDF_SLIDER) Then 'double-break is mandatory for slider
                If Right$(m_sExpandedInfo, 4) <> (vbCrLf & vbCrLf) Then
                    m_sExpandedInfo = m_sExpandedInfo & vbCrLf  '(it's already sure to have 1)
                End If
                'End If
                #If Win64 Then
                    AssignLongPtr uTDC.pszExpandedInformation, uTDC.pszExpandedInformation2, StrPtr(m_sExpandedInfo)
                #Else
                    uTDC.pszExpandedInformation = StrPtr(m_sExpandedInfo)
                #End If
            End If

        End If
    End If
    If (uTDC.dwFlags And TDF_CALLBACK_TIMER) Then
        bPublicTimer = True
    Else
        If (uTDC.dwFlags And TDF_INPUT_BOX) Or (uTDC.dwFlags And TDF_COMBO_BOX) Or (uTDC.dwFlags And TDF_DATETIME) Or (uTDC.dwFlags And TDF_SLIDER) Then
            If m_sExpandedInfo <> "" Then
                If (uTDC.dwFlags And TDF_EXPAND_FOOTER_AREA) <> TDF_EXPAND_FOOTER_AREA Then
                    uTDC.dwFlags = uTDC.dwFlags Or TDF_CALLBACK_TIMER
                End If
            End If
        End If
        If (lAutoclose > 0) Then
            uTDC.dwFlags = uTDC.dwFlags Or TDF_CALLBACK_TIMER
        End If
    End If
    If (uTDC.dwFlags And TDF_USE_HICON_MAIN) Then
        uTDC.pszMainIcon = m_Icon
    Else
        uTDC.pszMainIcon = tdMakeIntResource(m_Icon)
    End If
    If (uTDC.dwFlags And TDF_USE_HICON_FOOTER) Then
        uTDC.pszFooterIcon = m_IconFooter
    Else
        uTDC.pszFooterIcon = tdMakeIntResource(m_IconFooter)
    End If

    If m_HasCustomButtons Then
        uTDC.cButtons = UBound(m_uButtons) + 1
    Else
        uTDC.cButtons = 0
    End If
    #If Win64 Then
        AssignLongPtr uTDC.pButtons, uTDC.pButtons2, VarPtr(m_uButtons(0).nButtonID)
    #Else
        uTDC.pButtons = VarPtr(m_uButtons(0).nButtonID)
    #End If

    If m_HasRadioButtons Then
        uTDC.cRadioButtons = UBound(m_uRadioButtons) + 1
    Else
        uTDC.cRadioButtons = 0
    End If

    #If Win64 Then
        AssignLongPtr uTDC.pRadioButtons, uTDC.pRadioButtons2, VarPtr(m_uRadioButtons(0).nButtonID)
    #Else
        uTDC.pRadioButtons = VarPtr(m_uRadioButtons(0).nButtonID)
    #End If

    If (uTDC.dwFlags And TDF_USE_SHELL32_ICONID) Then
        #If Win64 Then
            AssignLongPtr uTDC.hInstance, uTDC.hInstance2, VarPtr(hShell32)
        #Else
            uTDC.hInstance = hShell32
        #End If
        uTDC.pszMainIcon = tdMakeIntResource(m_Icon)
        uTDC.pszFooterIcon = tdMakeIntResource(m_IconFooter)

    ElseIf (uTDC.dwFlags And TDF_USE_IMAGERES_ICONID) Then
        #If Win64 Then
            AssignLongPtr uTDC.hInstance, uTDC.hInstance2, VarPtr(hImageRes)
        #Else
            uTDC.hInstance = hImageRes
        #End If
        uTDC.pszMainIcon = tdMakeIntResource(m_Icon)
        uTDC.pszFooterIcon = tdMakeIntResource(m_IconFooter)
    End If


    #If Win64 Then
        AssignLongPtr uTDC.pfCallback, uTDC.pfCallback2, tdFARPROC(AddressOf TaskDialogCallbackProc)
        AssignLongPtr uTDC.lpCallbackData, uTDC.lpCallbackData2, ObjPtr(Me)
    #Else
        uTDC.pfCallback = tdFARPROC(AddressOf TaskDialogCallbackProc)
        uTDC.lpCallbackData = ObjPtr(Me)
    #End If

    sComboText = ""
    nComboIdx = -1

    hr = TaskDialogIndirect(uTDC, pnButton, pnRadButton, pfVerify)

    'Call GetLastError_Msg(hr, , sErr)
    'Debug.Print "TaskDialogIndirect ret=0x" & Hex$(hr) & ": " & sErr
    m_ResMain = pnButton
    m_ResRad = pnRadButton
    m_ResVerify = pfVerify
    If pnButton <> 0 Then
        ShowDialog = True
    End If

End Function

Public Function SimpleDialog(sMessage As String, Optional dwBtn As TDBUTTONS = TDCBF_OK_BUTTON, Optional sTitle As String, Optional sMainText As String, Optional dwIco As TDICONS, Optional hWndOwner As Long, Optional hInst As Long) As TDBUTTONS
    Dim dwIcon                          As Long
    Dim pnButton                        As Long
    Dim Success                         As Long

    Dim pszTitle                        As LongPtr
    Dim pszMain                         As LongPtr
    Dim pszContent                      As LongPtr

    
    If sTitle = "" Then
        sTitle = APP_TITLE
    End If

        pszTitle = StrPtr(sTitle)
        If sMainText <> vbNullString Then pszMain = StrPtr(sMainText)
        If sMessage <> vbNullString Then pszContent = StrPtr(sMessage)


    If dwIco Then
        dwIcon = tdMakeIntResource(dwIco)
    End If
    Call TaskDialog(hWndOwner, _
                    hInst, _
                    pszTitle, _
                    pszMain, _
                    pszContent, _
                    dwBtn, _
                    dwIcon, _
                    pnButton)

    SimpleDialog = pnButton
End Function

Private Sub SliderReposition()
    If hSlider Then
        Dim tRC                         As RECT
        Dim lOffX                       As Long
        Dim lEditX                      As Long
        Dim lEditY                      As Long
        Dim cxEdit                      As Long
        lButtonY4 = 0
        ReDim arBtnX(0)

        Call EnumChildWindows(m_TDhWnd, AddressOf TaskDialogEnumChildProc, ObjPtr(Me))

        GetClientRect hUI, tRC

        If nSliderAlign = TDIBA_Content Then
            If m_Icon Then
                lOffX = (32 * m_ScaleX) + 20    '52
            Else
                lOffX = 12
            End If
            If (uTDC.dwFlags And TDF_KILL_SHIELD_ICON) And (m_lNewIcon = 0) Then
                lOffX = 12
            End If
            lEditX = lOffX
            If (uTDC.dwFlags And TDF_USE_COMMAND_LINKS) Or (uTDC.dwFlags And TDF_USE_COMMAND_LINKS_NO_ICON) Or (m_HasRadioButtons = True) Then
                lEditY = lButtonY4 - (42 * m_ScaleY)
            Else
                If nSliderTickStyle = SldTickStyleBoth Then
                    lEditY = lButtonY4 - (54 * m_ScaleY)
                Else
                    lEditY = lButtonY4 - (44 * m_ScaleY)
                End If
                If bLinkAdjust Then
                    lEditY = lEditY - (6 * m_ScaleY)
                End If
            End If
            If (uTDC.dwFlags And TDF_SHOW_MARQUEE_PROGRESS_BAR) Or (uTDC.dwFlags And TDF_SHOW_PROGRESS_BAR) Then
                lEditY = lEditY - (22 * m_ScaleY)
            End If
            If (m_sExpandedInfo <> "") Then
                lEditY = lEditY - (6 * m_ScaleY)
            End If
            cxEdit = tRC.Right - lOffX - 12
            If cxSlider > 0 Then cxEdit = cxSlider
        ElseIf nSliderAlign = TDIBA_Footer Then
            '        lEditX = 32
            '        If nSliderTickStyle = SldTickStyleBoth Then
            '            lEditY = tRC.Bottom - 38
            '        Else
            '            lEditY = tRC.Bottom - 30
            '        End If
            '        cxEdit = 180
            '        If cxSlider > 0 Then cxEdit = cxSlider
            If cxSlider = -1 Then
                cxEdit = tRC.Right - ((16 * m_ScaleX) + 32)    '48
            Else
                cxEdit = 180 * m_ScaleX
            End If
            If cxSlider > 0 Then cxEdit = cxSlider
            Select Case nSLFooterAlign

                Case tdcaLeft
                    lEditX = (16 * m_ScaleX) + 16    '32
                Case tdcaCenter
                    lEditX = ((16 * m_ScaleX) + 16) + (((tRC.Right - ((16 * m_ScaleX) + 16) - 10) - (cxEdit)) \ 2)
                Case tdcaRight
                    lEditX = tRC.Right - cxEdit - 10
            End Select
            If nSliderTickStyle = SldTickStyleBoth Then
                If (m_sVerify <> "") And (m_sExpandedInfo <> "") Then
                    lEditY = tRC.Bottom - ((30 * m_ScaleY) + 5)    'lButtonY4 + 47
                Else
                    lEditY = tRC.Bottom - ((30 * m_ScaleY))    'lButtonY4 + 29 tRC.Bottom - 38
                End If
            Else
                If (m_sVerify <> "") And (m_sExpandedInfo <> "") Then
                    lEditY = tRC.Bottom - ((30 * m_ScaleY) + 5)    'lButtonY4 + 55
                Else
                    lEditY = tRC.Bottom - ((30 * m_ScaleY))    ' lButtonY4 + 37 'tRC.Bottom - 30
                End If
            End If
            If (m_sFooter = "") Or (m_sFooter = "$input") Then
                m_sFooter = " "
                '        Call SendMessageW(m_TDhWnd, TDM_UPDATE_ELEMENT_TEXT, TDE_FOOTER, ByVal StrPtr(m_sFooter))
            End If

            If (m_sExpandedInfo <> "") Then
                lEditY = lEditY - (6 * m_ScaleY)
            End If

        End If
        If (nSliderAlign = TDIBA_Content) Or (nSliderAlign = TDIBA_Footer) Then
            SetWindowPos hSlider, 0&, lEditX, lEditY, cxEdit, 36&, SWP_NOSIZE
            RepaintWindow hSlider

        End If
        SetWindowPos m_TDhWnd, 0, 0, 0, tRC.Right, tRC.Bottom + (42 * m_ScaleY), SWP_NOMOVE

    End If
End Sub

Public Sub SliderSetChangeValues(lSmallChange As Long, lLargeChange As Long)
    lSliderChangeS = lSmallChange
    lSliderChangeL = lLargeChange
    If hSlider Then
        SendMessageW hSlider, TBM_SETLINESIZE, 0, ByVal lSliderChangeS
        SendMessageW hSlider, TBM_SETPAGESIZE, 0, ByVal lSliderChangeL
    End If
End Sub

Public Sub SliderSetRange(lMin As Long, lMax As Long, Optional lTickFreq As Long)
    lSliderMin = lMin
    lSliderMax = lMax
    lSliderNTicks = lTickFreq
    If hSlider Then
        SendMessageW hSlider, TBM_SETRANGEMIN, 0, ByVal lMin
        SendMessageW hSlider, TBM_SETRANGEMAX, 1, ByVal lMax
        If lTickFreq Then
            SendMessageW hSlider, TBM_SETTICFREQ, lSliderNTicks, ByVal 0&
        End If
    End If
End Sub

Private Sub SortLongArray(paintArray() As Long)
' This sub uses the Bubble Sort algorithm to sort an array of integers.
    Dim lngY                            As Long
    Dim intTemp                         As Long
    Dim blnExchangeMade                 As Boolean

    blnExchangeMade = True

    Do While blnExchangeMade
        blnExchangeMade = False

        For lngY = LBound(paintArray) To (UBound(paintArray) - 1)
            If paintArray(lngY) > paintArray(lngY + 1) Then
                intTemp = paintArray(lngY)
                paintArray(lngY) = paintArray(lngY + 1)
                paintArray(lngY + 1) = intTemp
                blnExchangeMade = True
            End If
        Next
    Loop
End Sub

Private Function Subclass(hWnd As LongPtr, lpfn As LongPtr, Optional uId As LongPtr = 0&, Optional dwRefData As LongPtr = 0&) As Boolean
    If uId = 0 Then uId = hWnd
    Subclass = SetWindowSubclass(hWnd, lpfn, uId, dwRefData): Debug.Assert Subclass
End Function

Private Function tdFARPROC(pfn As LongPtr) As LongPtr
    tdFARPROC = pfn
End Function

Private Function tdHiWord(ByVal dwValue As Long) As Integer
    CopyMemory tdHiWord, ByVal VarPtr(dwValue) + 2, 2
End Function

Private Function tdLoWord(dwValue As Long) As Integer
    CopyMemory tdLoWord, dwValue, 2
End Function

Private Function tdMakeIntResource(ByVal dwVal As Long) As Long
    tdMakeIntResource = &HFFFF& And dwVal
End Function

Private Function tdMAKELPARAM(wLow As Long, wHigh As Long) As Long
    tdMAKELPARAM = tdLoWord(wLow) Or (&H10000 * tdLoWord(wHigh))
End Function

Private Function UnSubclass(hWnd As LongPtr, ByVal lpfn As Long, pid As Long) As Boolean
    UnSubclass = RemoveWindowSubclass(hWnd, lpfn, pid)
End Function


'   .______   .______        ______   .______    _______ .______      .___________. __   _______      _______.
'   |   _  \  |   _  \      /  __  \  |   _  \  |   ____||   _  \     |           ||  | |   ____|    /       |
'   |  |_)  | |  |_)  |    |  |  |  | |  |_)  | |  |__   |  |_)  |    `---|  |----`|  | |  |__      |   (----`
'   |   ___/  |      /     |  |  |  | |   ___/  |   __|  |      /         |  |     |  | |   __|      \   \
'   |  |      |  |\  \----.|  `--'  | |  |      |  |____ |  |\  \----.    |  |     |  | |  |____ .----)   |
'   | _|      | _| `._____| \______/  | _|      |_______|| _| `._____|    |__|     |__| |_______||_______/
                                               
Public Property Get AutocloseTime() As Long: AutocloseTime = lAutocloseTrack: End Property

Public Property Let AutocloseTime(value As Long): lAutocloseTrack = value: End Property

Public Property Get CollapsedControlText() As String: CollapsedControlText = m_sColControl: End Property

Public Property Let CollapsedControlText(sText As String)
    m_sColControl = sText
    
    #If Win64 Then
        AssignLongPtr uTDC.pszCollapsedControlText, uTDC.pszCollapsedControlText2, StrPtr(m_sColControl)
    #Else
        uTDC.pszCollapsedControlText = StrPtr(m_sColControl)
    #End If
End Property

Public Property Get ComboAlign() As TDInputBoxAlign: ComboAlign = nComboAlign: End Property

Public Property Let ComboAlign(nAlign As TDInputBoxAlign): nComboAlign = nAlign: End Property

Public Property Get ComboAlignInFooter() As TDControlAlign: ComboAlignInFooter = nCBFooterAlign: End Property

Public Property Let ComboAlignInFooter(nAlign As TDControlAlign): nCBFooterAlign = nAlign: End Property

Public Property Get ComboCueBanner() As String: ComboCueBanner = m_sComboCue: End Property

Public Property Let ComboCueBanner(sText As String)
    m_sComboCue = sText
    If hEditCombo Then
        SendMessageW hEditCombo, EM_SETCUEBANNER, 0&, ByVal StrPtr(m_sComboCue)
    End If
End Property

Public Property Get ComboHeight() As Long: ComboHeight = cyCombo: End Property

Public Property Let ComboHeight(CY As Long): cyCombo = CY: End Property

Public Property Get ComboImageList() As LongPtr: ComboImageList = himlCombo: End Property

Public Property Let ComboImageList(himl As LongPtr): himlCombo = himl: End Property

Public Property Get ComboIndex() As LongPtr
    nComboIdx = -1
    If hCombo Then
        nComboIdx = SendMessageW(hCombo, CB_GETCURSEL, 0&, ByVal 0&)
    End If
    ComboIndex = nComboIdx
End Property

Public Property Let ComboIndex(nIndex As LongPtr)
    nComboIdx = nIndex
    If hCombo Then
        Call SendMessageW(hCombo, CB_SETCURSEL, nIndex, ByVal 0&)
    End If
End Property

Public Property Get ComboStyle() As ComboType: ComboStyle = nCbType: End Property

Public Property Let ComboStyle(nStyle As ComboType): nCbType = nStyle: End Property

Public Property Get ComboText() As String
    If hCombo Then
        ComboText = GetComboTextW(hEditCombo)
    End If
End Property

Public Property Get ComboWidth() As Long: ComboWidth = cxCombo: End Property

Public Property Let ComboWidth(CX As Long): cxCombo = CX: End Property

Public Property Get CommonButtons() As TDBUTTONS: CommonButtons = uTDC.dwCommonButtons: End Property

Public Property Let CommonButtons(dwBtns As TDBUTTONS): uTDC.dwCommonButtons = dwBtns: End Property

Public Property Get Content() As String: Content = m_sContent: End Property

Public Property Let Content(sText As String)
    m_sContent = sText
    If (uTDC.dwFlags And TDF_INPUT_BOX) Then
        If nInputAlign = TDIBA_Content Then
            If Right$(m_sContent, 2) <> vbCrLf Then
                m_sContent = m_sContent & vbCrLf
            End If
        End If
    End If

    #If Win64 Then
        AssignLongPtr uTDC.pszContent, uTDC.pszContent2, StrPtr(m_sContent)
    #Else
        uTDC.pszContent = StrPtr(m_sContent)
    #End If
    
    If m_TDhWnd <> 0& Then
        Call SendMessageW(m_TDhWnd, TDM_UPDATE_ELEMENT_TEXT, TDE_CONTENT, ByVal StrPtr(m_sContent))
    End If
End Property

Public Property Get DateTimeAlign() As TDInputBoxAlign: ComboAlign = nDateTimeAlign: End Property

Public Property Let DateTimeAlign(nAlign As TDInputBoxAlign): nDateTimeAlign = nAlign: End Property

Public Property Get DateTimeAlignInContent() As TDControlAlign: DateTimeAlignInContent = nDTContentAlign: End Property

Public Property Let DateTimeAlignInContent(nAlign As TDControlAlign): nDTContentAlign = nAlign: End Property

Public Property Get DateTimeAlignInFooter() As TDControlAlign: DateTimeAlignInFooter = nDTFooterAlign: End Property

Public Property Let DateTimeAlignInFooter(nAlign As TDControlAlign): nDTFooterAlign = nAlign: End Property

Public Property Get DateTimeChecked() As Long
    Dim RS As LongPtr, rs2                 As LongPtr
    RS = SendMessageW(hDateTime, DTM_GETSYSTEMTIME, 0&, tDTCur)
    If hDateTime2 Then
        rs2 = SendMessageW(hDateTime2, DTM_GETSYSTEMTIME, 0&, tDTCur2)
    End If
    If RS = GDT_VALID Then
        bDTCheck = True
    Else
        bDTCheck = False
    End If
    If rs2 = GDT_VALID Then
        bDTCheck2 = True
    Else
        bDTCheck2 = False
    End If
    If (nDTType = dttDateTimeWithCheck) Then    'don't return these values if there's only 1 check
        If bDTCheck And bDTCheck2 Then
            DateTimeChecked = 4
        ElseIf (bDTCheck = True) And (bDTCheck2 = False) Then
            DateTimeChecked = 3
        ElseIf (bDTCheck = False) And (bDTCheck2 = True) Then
            DateTimeChecked = 2
        ElseIf (bDTCheck = False) And (bDTCheck2 = False) Then
            DateTimeChecked = 0
        End If
    Else
        If (nDTType = dttDateWithCheck) Or (nDTType = dttTimeWithCheck) Or (nDTType = dttDateTimeWithCheckDateOnly) Then
            If bDTCheck = True Then
                DateTimeChecked = 1
            Else
                DateTimeChecked = 0
            End If
        Else    'date time with a single check, on the time
            If bDTCheck2 = True Then
                DateTimeChecked = 1
            Else
                DateTimeChecked = 0
            End If
        End If
    End If
End Property

Public Property Let DateTimeChecked(NewValue As Long)
    If hDateTime Then
        Call SendMessageW(hDateTime, DTM_GETSYSTEMTIME, 0&, tDTCur)
    End If
    If hDateTime2 Then
        Call SendMessageW(hDateTime2, DTM_GETSYSTEMTIME, 0&, tDTCur2)
    End If
    If nDTType = dttDateTimeWithCheck Then
        Select Case NewValue
            Case 4
                bDTCheck = True
                bDTCheck2 = True
            Case 3
                bDTCheck = True
                bDTCheck2 = False
            Case 2
                bDTCheck = False
                bDTCheck2 = True
            Case 0
                bDTCheck = False
                bDTCheck2 = False
        End Select
        If hDateTime Then
            If bDTCheck Then
                Call SendMessageW(hDateTime, DTM_SETSYSTEMTIME, GDT_VALID, tDTCur)
            Else
                Call SendMessageW(hDateTime, DTM_SETSYSTEMTIME, GDT_NONE, tDTCur)
            End If
            If bDTCheck2 Then
                Call SendMessageW(hDateTime2, DTM_SETSYSTEMTIME, GDT_VALID, tDTCur2)
            Else
                Call SendMessageW(hDateTime2, DTM_SETSYSTEMTIME, GDT_NONE, tDTCur2)
            End If
        End If

    Else
        If NewValue = 1 Then
            bDTCheck = True
        ElseIf NewValue = 0 Then
            bDTCheck = False
        End If

        If nDTType = dttDateTimeWithCheckTimeOnly Then
            If bDTCheck Then
                Call SendMessageW(hDateTime2, DTM_SETSYSTEMTIME, GDT_VALID, tDTCur)
            Else
                Call SendMessageW(hDateTime2, DTM_SETSYSTEMTIME, GDT_NONE, tDTCur)
            End If
        Else    'only checkbox is the first control
            If bDTCheck Then
                Call SendMessageW(hDateTime, DTM_SETSYSTEMTIME, GDT_VALID, tDTCur)
            Else
                Call SendMessageW(hDateTime, DTM_SETSYSTEMTIME, GDT_NONE, tDTCur)
            End If
        End If
    End If

End Property

Public Property Get DateTimeType() As DateTimeType: DateTimeType = nDTType: End Property

Public Property Let DateTimeType(nType As DateTimeType): nDTType = nType: End Property

Public Property Get DateTimeValue() As Date
    Call SendMessageW(hDateTime, DTM_GETSYSTEMTIME, 0&, tDTCur)
    Call SendMessageW(hDateTime2, DTM_GETSYSTEMTIME, 0&, tDTCur2)
    If hDateTime2 Then
        DateTimeValue = DateSerial(tDTCur.wYear, tDTCur.wMonth, tDTCur.wDay) + TimeSerial(tDTCur2.wHour, tDTCur2.wMinute, tDTCur2.wSecond)
    Else
        DateTimeValue = DateSerial(tDTCur.wYear, tDTCur.wMonth, tDTCur.wDay) + TimeSerial(tDTCur.wHour, tDTCur.wMinute, tDTCur.wSecond)
    End If
End Property

Public Property Let DateTimeValue(dt As Date)
    With tDTCur
        .wDay = VBA.DateTime.Day(dt)
        .wMonth = VBA.DateTime.Month(dt)
        .wYear = VBA.DateTime.Year(dt)
        .wHour = VBA.DateTime.Hour(dt)
        .wMinute = VBA.DateTime.Minute(dt)
        .wSecond = VBA.DateTime.Second(dt)
    End With
    With tDTCur2
        .wHour = VBA.DateTime.Hour(dt)
        .wMinute = VBA.DateTime.Minute(dt)
        .wSecond = VBA.DateTime.Second(dt)
    End With
    If hDateTime2 Then
        Call SendMessageW(hDateTime2, DTM_SETSYSTEMTIME, GDT_VALID, tDTCur2)
    End If
    If hDateTime Then
        Call SendMessageW(hDateTime, DTM_SETSYSTEMTIME, GDT_VALID, tDTCur)
    End If
End Property

Public Property Get DefaultButton() As TDRESULT: DefaultButton = uTDC.nDefaultButton: End Property

Public Property Let DefaultButton(ButtonID As TDRESULT): uTDC.nDefaultButton = ButtonID: End Property

Public Property Get DefaultCustomControl() As TDCustomControlType: DefaultCustomControl = nDefaultCustCtrl: End Property

Public Property Let DefaultCustomControl(iControl As TDCustomControlType): nDefaultCustCtrl = iControl: End Property

Public Property Get DefaultRadioButton() As Long: DefaultButton = uTDC.nDefaultRadioButton: End Property

Public Property Let DefaultRadioButton(ButtonID As Long): uTDC.nDefaultRadioButton = ButtonID: End Property

Public Property Get DPIScaleX() As Single: DPIScaleX = m_ScaleX: End Property

Public Property Let DPIScaleX(value As Single): m_ScaleX = value: End Property

Public Property Get DPIScaleY() As Single: DPIScaleY = m_ScaleY: End Property

Public Property Let DPIScaleY(value As Single): m_ScaleY = value: End Property

Public Property Get ExpandedControlText() As String: ExpandedControlText = m_sExpControl: End Property

Public Property Let ExpandedControlText(sText As String)
    m_sExpControl = sText
    
    #If Win64 Then
        AssignLongPtr uTDC.pszExpandedControlText, uTDC.pszExpandedControlText2, StrPtr(m_sExpControl)
    #Else
        uTDC.pszExpandedControlText = StrPtr(m_sExpControl)
    #End If
End Property

Public Property Get ExpandedInfo() As String: ExpandedInfo = m_sExpandedInfo: End Property

Public Property Let ExpandedInfo(sText As String)
    m_sExpandedInfo = sText
    
    #If Win64 Then
        AssignLongPtr uTDC.pszExpandedInformation, uTDC.pszExpandedInformation2, StrPtr(m_sExpandedInfo)
    #Else
        uTDC.pszExpandedInformation = StrPtr(m_sExpandedInfo)
    #End If
    
    If m_TDhWnd <> 0& Then
        Call SendMessageW(m_TDhWnd, TDM_UPDATE_ELEMENT_TEXT, TDE_EXPANDED_INFORMATION, ByVal StrPtr(m_sExpandedInfo))
    End If
End Property

Public Property Get FLAGS() As TASKDIALOG_FLAGS: FLAGS = uTDC.dwFlags: End Property

Public Property Let FLAGS(dwFlags As TASKDIALOG_FLAGS): uTDC.dwFlags = dwFlags: End Property

Public Property Get Footer() As String: Footer = m_sFooter: End Property

Public Property Let Footer(sText As String)
    m_sFooter = sText
    #If Win64 Then
        AssignLongPtr uTDC.pszFooter, uTDC.pszFooter2, StrPtr(m_sFooter)
    #Else
        uTDC.pszFooter = StrPtr(m_sFooter)
    #End If
    If m_TDhWnd <> 0& Then
        Call SendMessageW(m_TDhWnd, TDM_UPDATE_ELEMENT_TEXT, TDE_FOOTER, ByVal StrPtr(m_sFooter))
    End If
End Property

Public Property Get hInst() As Long: hInst = uTDC.hInstance: End Property

Public Property Let hInst(hInstance As Long): uTDC.hInstance = hInstance: End Property

Public Property Get hWndCombo() As LongPtr: hWndCombo = hCombo: End Property

Public Property Get hWndComboEdit() As LongPtr: hWndComboEdit = hEditCombo: End Property

Public Property Get hWndDateTime() As LongPtr: hWndDateTime = hDateTime: End Property

Public Property Get hWndDlg() As LongPtr: hWndDlg = m_TDhWnd: End Property

Public Property Get hWndDUI() As LongPtr: hWndDUI = hUI: End Property              'the DirectUIHWND class that the controls are on; and what is used as a parent for added controls

Public Property Get hWndInput() As LongPtr: hWndInput = hEditBox: End Property

Public Property Get hWndSlider() As LongPtr: hWndSlider = hSlider: End Property

Public Property Get IconFooter() As TDICONS: IconFooter = m_IconFooter: End Property

Public Property Let IconFooter(lIcon As TDICONS)
    m_IconFooter = lIcon

    If m_TDhWnd <> 0& Then
        If (uTDC.dwFlags And TDF_USE_HICON_FOOTER) Then
            Call SendMessageW(m_TDhWnd, TDM_UPDATE_ICON, TDIE_ICON_FOOTER, ByVal m_IconFooter)
        Else
            Dim dwIcon                  As Long
            dwIcon = tdMakeIntResource(m_IconFooter)
            Call SendMessageW(m_TDhWnd, TDM_UPDATE_ICON, TDIE_ICON_FOOTER, ByVal dwIcon)
        End If
    End If
End Property

Public Property Get IconMain() As TDICONS: IconMain = m_Icon: End Property

Public Property Let IconMain(lIcon As TDICONS)
    m_Icon = lIcon

    If m_TDhWnd <> 0& Then
        If (uTDC.dwFlags And TDF_USE_HICON_MAIN) Then
            Call SendMessageW(m_TDhWnd, TDM_UPDATE_ICON, TDIE_ICON_MAIN, ByVal m_Icon)
        Else
            Dim dwIcon                  As Long
            dwIcon = tdMakeIntResource(m_Icon)
            Call SendMessageW(m_TDhWnd, TDM_UPDATE_ICON, TDIE_ICON_MAIN, ByVal dwIcon)
        End If
    End If
End Property

Public Property Get IconReplaceGradient() As Long: IconReplaceGradient = m_lNewIcon: End Property

Public Property Let IconReplaceGradient(nIndex As Long)
    m_lNewIcon = nIndex
    If m_TDhWnd <> 0& Then
        If (uTDC.dwFlags And TDF_KILL_SHIELD_ICON) Then
            SendMessageW m_TDhWnd, TDM_UPDATE_ICON, TDIE_ICON_MAIN, ByVal m_lNewIcon
        End If
    End If
End Property

Public Property Get InputAlign() As TDInputBoxAlign: InputAlign = nInputAlign: End Property

Public Property Let InputAlign(nAlign As TDInputBoxAlign): nInputAlign = nAlign: End Property

Public Property Get InputAlignInFooter() As TDControlAlign: InputAlignInFooter = nIBFooterAlign: End Property

Public Property Let InputAlignInFooter(nAlign As TDControlAlign): nIBFooterAlign = nAlign: End Property

Public Property Get InputCueBanner() As String: InputCueBanner = m_sInput: End Property

Public Property Let InputCueBanner(sText As String)
    m_sInput = sText
    If hEditBox Then
        SendMessageW hEditBox, EM_SETCUEBANNER, 0&, ByVal StrPtr(m_sInput)
    End If
End Property

Public Property Get InputIsPassword() As Boolean: InputIsPassword = bEditPassword: End Property

Public Property Let InputIsPassword(bVal As Boolean)
    If hEditBox Then
        If bEditPassword <> bVal Then
            Dim dwStyle                 As LongPtr
            dwStyle = GetWindowLong(hEditBox, GWL_STYLE)
            If bVal Then
                dwStyle = dwStyle Or ES_PASSWORD
            Else
                dwStyle = dwStyle And Not ES_PASSWORD
            End If
            Call SetWindowLong(hEditBox, GWL_STYLE, dwStyle)

        End If
    End If
    bEditPassword = bVal
End Property

Public Property Get InputText() As String
    If hEditBox Then
        Dim lLen As LongPtr, sText         As String
        lLen = SendMessageW(hEditBox, WM_GETTEXTLENGTH, 0, ByVal 0&) * 2
        If lLen Then
            sText = Space$(CLng(lLen)) '//KCB added CLng
            Call SendMessageW(hEditBox, WM_GETTEXT, lLen, ByVal sText)
            sEditText = StrConv(sText, vbFromUnicode)
        End If
    End If
    InputText = sEditText
End Property

Public Property Let InputText(sText As String)
    sEditText = sText
    If hEditBox Then
        SendMessageW hEditBox, WM_SETTEXT, 0&, ByVal StrPtr(sEditText)
    End If
End Property

Public Property Get InputWidth() As Long: InputWidth = cxInputBox: End Property

Public Property Let InputWidth(CX As Long): cxInputBox = CX: End Property

Public Property Get MainInstruction() As String: MainInstruction = m_sMainInstr: End Property

Public Property Let MainInstruction(sText As String)
    m_sMainInstr = sText
    #If Win64 Then
        AssignLongPtr uTDC.pszMainInstruction, uTDC.pszMainInstruction2, StrPtr(m_sMainInstr)
    #Else
        uTDC.pszMainInstruction = StrPtr(m_sMainInstr)
    #End If
    
    If m_TDhWnd <> 0& Then
        Call SendMessageW(m_TDhWnd, TDM_UPDATE_ELEMENT_TEXT, TDE_MAIN_INSTRUCTION, ByVal StrPtr(m_sMainInstr))
    End If
End Property

Public Property Get ParenthWnd() As Long: ParenthWnd = uTDC.hwndParent: End Property

Public Property Let ParenthWnd(hWnd As Long): uTDC.hwndParent = hWnd: End Property

Public Property Get ResultComboIndex() As LongPtr: ResultComboIndex = nComboIdx: End Property

Public Property Get ResultComboText() As String: ResultComboText = sComboText: End Property

Public Property Get ResultDateTime() As Date
    If (nDTType = dttDate) Or (nDTType = dttDateWithCheck) Then
        ResultDateTime = DateSerial(tDTCur.wYear, tDTCur.wMonth, tDTCur.wDay)
    End If
    If (nDTType = dttTime) Or (nDTType = dttTimeWithCheck) Then
        ResultDateTime = TimeSerial(tDTCur.wHour, tDTCur.wMinute, tDTCur.wSecond)
    End If
    If (nDTType = dttDateTime) Or (nDTType = dttDateTimeWithCheck) Then
        ResultDateTime = DateSerial(tDTCur.wYear, tDTCur.wMonth, tDTCur.wDay) + TimeSerial(tDTCur2.wHour, tDTCur2.wMinute, tDTCur2.wSecond)
    End If
End Property

Public Property Get ResultDateTimeChecked() As Long
    If (nDTType = dttDateTimeWithCheck) Then    'don't return these values if there's only 1 check
        If bDTCheck And bDTCheck2 Then
            ResultDateTimeChecked = 4
        ElseIf (bDTCheck = True) And (bDTCheck2 = False) Then
            ResultDateTimeChecked = 3
        ElseIf (bDTCheck = False) And (bDTCheck2 = True) Then
            ResultDateTimeChecked = 2
        ElseIf (bDTCheck = False) And (bDTCheck2 = False) Then
            ResultDateTimeChecked = 0
        End If
    Else
        If (nDTType = dttDateWithCheck) Or (nDTType = dttTimeWithCheck) Or (nDTType = dttDateTimeWithCheckDateOnly) Then
            If bDTCheck = True Then
                ResultDateTimeChecked = 1
            Else
                ResultDateTimeChecked = 0
            End If
        Else    'date time with a single check, on the time
            If bDTCheck2 = True Then
                ResultDateTimeChecked = 1
            Else
                ResultDateTimeChecked = 0
            End If
        End If
    End If

End Property

Public Property Get ResultInput() As String: ResultInput = sEditText: End Property

Public Property Get ResultMain() As TDRESULT: ResultMain = m_ResMain: End Property

Public Property Get ResultRad() As Long: ResultRad = m_ResRad: End Property

Public Property Get ResultSlider() As LongPtr: ResultSlider = lSliderValue: End Property

Public Property Get ResultVerify() As Long: ResultVerify = m_ResVerify: End Property

Public Property Get SliderAlign() As TDInputBoxAlign: SliderAlign = nSliderAlign: End Property

Public Property Let SliderAlign(nAlign As TDInputBoxAlign): nSliderAlign = nAlign: End Property

Public Property Get SliderAlignInFooter() As TDControlAlign: SliderAlignInFooter = nSLFooterAlign: End Property

Public Property Let SliderAlignInFooter(nAlign As TDControlAlign): nSLFooterAlign = nAlign: End Property

Public Property Get SliderTickStyle() As SldTickStyle: SliderTickStyle = nSliderTickStyle: End Property

Public Property Let SliderTickStyle(nStyle As SldTickStyle): nSliderTickStyle = nStyle: End Property

Public Property Get SliderValue() As LongPtr
    If hSlider Then
        lSliderValue = SendMessageW(hSlider, TBM_GETPOS, 0, ByVal 0&)
    End If
    SliderValue = lSliderValue
    Debug.Print lSliderValue
End Property

Public Property Let SliderValue(nValue As LongPtr)
    lSliderValue = nValue
    If hSlider Then
        Call SendMessageW(hSlider, TBM_SETPOS, 1&, ByVal lSliderValue)
    End If

End Property

Public Property Get SliderWidth() As Long: SliderWidth = cxSlider: End Property

Public Property Let SliderWidth(CX As Long): cxSlider = CX: End Property

Public Property Get Title() As String: Title = m_sTitle: End Property

Public Property Let Title(sTitle As String)
    m_sTitle = sTitle
    #If Win64 Then
        AssignLongPtr uTDC.pszWindowTitle, uTDC.pszWindowTitle2, StrPtr(m_sTitle)
    #Else
        uTDC.pszWindowTitle = StrPtr(m_sTitle)
    #End If
End Property

Public Property Get VerifyText() As String: VerifyText = m_sVerify: End Property

Public Property Let VerifyText(sText As String)
    m_sVerify = sText
    
    #If Win64 Then
        AssignLongPtr uTDC.pszVerificationText, uTDC.pszVerificationText2, StrPtr(m_sVerify)
    #Else
        uTDC.pszVerificationText = StrPtr(m_sVerify)
    #End If
End Property

Public Property Get Width() As Long: Width = uTDC.cxWidth: End Property

Public Property Let Width(value As Long): uTDC.cxWidth = value: End Property

Public Property Get Version() As String
    Version = vTdVersion
End Property

