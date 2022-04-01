Attribute VB_Name = "#_GUI"
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
'                 $$$$$$$$$$F                                       ## ##  ##  ##  ##    ##    ####  ##   ## ######
'                  *$$$$$$$$"                                        ####  #####  ##     ##   ##     ##   ##   ##
'                    "***""                   _____________                                   ## ### ##   ##   ##
' STANDARD MODULE WITH DEFAULT GUI FUNCTIONS |v 2017/03/19 |                                  ##  ## ##   ##   ##
' The module contains frequently used functions and is part of the G-VBA library               ####   #####  ######
'****************************************************************************************************************************************************
'****************************************************************************************************************************************************
' While G-VBA contains basic methods, the concrete implementation of working with users is presented in the #_DIALOG module, _
' which contains methods for calling dialog boxes for working with files, selecting lists, message boxes, etc.
'****************************************************************************************************************************************************
Option Explicit

Private Const MOD_NAME As String = "#_GUI"

Public Type GUID
    Data1                               As Long
    Data2                               As Integer
    Data3                               As Integer
    Data4(7)                            As Byte
End Type

Public Type BOX
    Left As Long
    Top As Long
    Width As Long
    height As Long
End Type



#If Win64 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
        
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
            
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As LongPtr, ByVal lpstring As String, ByVal cch As Long) As Long
    
    Private Declare PtrSafe Function DestroyIcon Lib "user32" (ByVal hIcon As LongPtr) As Long  '#####
 

#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpstring As String, ByVal aint As Long) As Long
     
    Private Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
       
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
        (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
    
    Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpstring As String) As Long
    Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpstring As String) As Long

    Private Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpstring As String) As Long
    
    Private Declare Function CallNextHookEx Lib "user32" _
        (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

    
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
    
    Private Declare Function DestroyIcon Lib "user32" Alias "DestroyIcon" (ByVal hIcon As Long) As Long  '##############

    Private hHook As Long           ' handle to the Hook procedure (global variable)
     
#End If

Private m_hWnd As Long
Public AWin As cAccessWindows




Public Enum HolstType
    GC_UNDEFINED = 0
    GC_DESKTOP = 1
    GC_ACTIVE_WINDOW = 2
    GC_SPECIFIC_FORM = 3
    GC_SPECIFIC_CONTROL = 4
    GC_NEW_WINDOW = 5
End Enum

Public Enum BackType
    GC_SOLID = 0
    GC_GRADIENT = -2
    GC_HATCHED_HORIZONTAL = -1     '-------------
    GC_HATCHED_VERTICAL = 1       '|||||||||||||
    GC_HATCHED_FDIAGONAL = 2      '\\\\\\\\\\\\\
    GC_HATCHED__BDIAGONAL = 3      '/////////////
    GC_HATCHED_CROSS = 4          '+++++++++++++
    GC_HATCHED_DIAGCROSS = 5      'XXXXXXXXXXXXXX
End Enum

Public Enum WinStyle
     GC_WIN_TRASPARENT = 0
     GC_WIN_SPLASH = 1
     GC_WIN_FULLREDRAW = 2
End Enum
'**************************************************************************************************************************************************


'======================================================================================================================================================
' Get active form hWnd
'======================================================================================================================================================
Public Function GetActiveHwnd() As Variant
Dim vRes As Variant

    On Error GoTo ErrHandle
'-----------------------------
    vRes = Screen.ActiveForm.hWnd
'-----------------------------
ExitHere:
    GetActiveHwnd = vRes '!!!!!!!!!!!!
    Exit Function
'--------
ErrHandle:
    If Err.Number = 2475 Then      ' No Any Active Form
        Err.Clear: vRes = Application.hWndAccessApp: Resume ExitHere
    Else
         ErrPrint2 "GetActiveHwnd", Err.Number, Err.Description, MOD_NAME
        Err.Clear
    End If
End Function

'======================================================================================================================================================
' Destroy Icon
'======================================================================================================================================================
Public Function DestroyIconG(hIconVar As Variant) As Long
#If Win64 Then
    Dim hIcon As LongPtr
    hIcon = CLngPtr(hIconVar)
#Else
    Dim hIcon As LongPtr
    hIcon = CLng(hIconVar)
#End If

    DestroyIconG = DestroyIcon(hIcon) '

End Function



'===================================================================================================================================================
' THIS FUNCTION IS REDIRECT FOR EMBEDDED AddressOf operator TO cAccessWindows class
'===================================================================================================================================================
Public Function WndProc(ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Variant
  WndProc = AWin.ClientWndProc(hWnd, MSG, wParam, lParam)
End Function   ' See Hook Function
'===================================================================================================================================================
' Get ActiveWindow Capture
'===================================================================================================================================================
Public Function GetActiveWindowCapture() As String
#If Win64 Then
  Dim hWnd As LongPtr
#Else
  Dim hWnd As Long
#End If
Dim charcount As Long, lpstring As String, strbuffer As Long, sRes As String

   On Error Resume Next
'---------------------------------
strbuffer = 300
lpstring = String$(strbuffer, Chr$(0))

hWnd = GetActiveWindow()
If hWnd = 0 Then Exit Function

charcount = GetWindowText(hWnd, lpstring, strbuffer)
If charcount > 0 Then
   sRes = Left$(lpstring, charcount)
End If
'-----------------------------------
ExitHere:
   GetActiveWindowCapture = sRes '!!!!!!!!!!
End Function
'===================================================================================================================================================
' Get String Actual Width
'===================================================================================================================================================
Public Function GetTextWidth(sText As String, Optional FontName As String = "MS Sans Serif", Optional FontSize As Integer = 8, _
                                                                                                 Optional Measurement As LengthUnit = GC_TWIPS) As Long
Dim TW As cSize, iRes As Long

    On Error GoTo ErrHandle
'-----------------------------
    If sText = "" Then Exit Function
    Set TW = New cSize
    iRes = TW.TextWidth(sText, FontName, FontSize)
    
    If Measurement = GC_TWIPS Then iRes = TW.PixelsToTwipsWidth(iRes)
'--------------------
ExitHere:
    GetTextWidth = iRes '!!!!!!!!!!
    Set TW = Nothing
    Exit Function
'----------------
ErrHandle:
    ErrPrint "GetTextWidth", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'===================================================================================================================================================
' Get String Actual Height
'===================================================================================================================================================
Public Function GetTextHeight(sText As String, Optional FontName As String = "MS Sans Serif", Optional FontSize As Integer = 8, _
                                                                                                 Optional Measurement As LengthUnit = GC_TWIPS) As Long
Dim TW As cSize, iRes As Long

    On Error GoTo ErrHandle
'-----------------------------
    If sText = "" Then Exit Function
    Set TW = New cSize
    iRes = TW.TextHeight(sText, FontName, FontSize)
    
    If Measurement = GC_TWIPS Then iRes = TW.PixelsToTwipsWidth(iRes)
'--------------------
ExitHere:
    GetTextHeight = iRes '!!!!!!!!!!
    Set TW = Nothing
    Exit Function
'----------------
ErrHandle:
    ErrPrint "GetTextHeighth", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'=====================================================================================================================================================
' Get Desktop hWnd
'=====================================================================================================================================================
Public Function GetDesktopHWND() As Variant
Dim PRPGR As New cSize
    GetDesktopHWND = PRPGR.GetDeskTopHandler()
'-------------------
    Set PRPGR = Nothing
End Function




'=====================================================================================================================================================
' Get Window Size and Position in a Pixels
'=====================================================================================================================================================
Public Function GetWindowSize(hWnd As Variant) As BOX
Dim PRPGR As New cSize, iBOX As BOX, WinDim() As Long

On Error Resume Next

    WinDim = PRPGR.GetWindowCoordinates(hWnd)
    iBOX.Left = WinDim(0): iBOX.Top = WinDim(1)
    iBOX.Width = WinDim(2): iBOX.height = WinDim(3)
    
    GetWindowSize = iBOX
    Set PRPGR = Nothing
End Function
'=====================================================================================================================================================
' Get Access Form Innner Client hWND
'=====================================================================================================================================================
Public Function AccessFormInnerHWND(f As Form) As Variant
Dim PRPGR As New cSize, fHWND As Variant

On Error GoTo ErrHandle
'-----------------------------
    If f Is Nothing Then Exit Function
    fHWND = PRPGR.GetInsideAccessFormHandler(f)
'------------------------------
ExitHere:
    AccessFormInnerHWND = fHWND
    Set PRPGR = Nothing
    Exit Function
'----------
ErrHandle:
    ErrPrint "AccessFormInnerHWND", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'=====================================================================================================================================================
' Get Access Form Innner Client Size
' gives the same values as the form methods f.InsideWidth & f.InsideHeight
'=====================================================================================================================================================
Public Function AccessFromInnerSise(f As Form, Optional Units As LengthUnit = GC_TWIPS) As BOX
Dim PRPGR As New cSize, fHWND As Variant, WinDim() As Long, BB As BOX

On Error GoTo ErrHandle
'-----------------------------
    If f Is Nothing Then Exit Function
    fHWND = PRPGR.GetInsideAccessFormHandler(f)
    BB = ArrayToBOX(PRPGR.GetWindowCoordinates(fHWND))
    
    If (BB.Width <> 0) Or (BB.Width <> 0) Then
        If Units = GC_TWIPS Then
            BB.Left = PRPGR.PixelsToTwipsWidth(BB.Left): BB.Width = PRPGR.PixelsToTwipsWidth(BB.Width)
            BB.Top = PRPGR.PixelsToTwipsHeight(BB.Top): BB.height = PRPGR.PixelsToTwipsHeight(BB.height)
        End If
     End If
'------------------------------
ExitHere:
    AccessFromInnerSise = BB '!!!!!!!!!!!!!!
    Set PRPGR = Nothing
    Exit Function
'----------
ErrHandle:
    ErrPrint "AccessFromInnerSise", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'=====================================================================================================================================================
' This Function Center some chldBox regarding prntBox, return new coordinate in chldBox
'=====================================================================================================================================================
Public Sub CenterBox(ByRef chldBox As BOX, ByRef prntBox As BOX)
    chldBox.Top = prntBox.Top + (prntBox.height - chldBox.height) \ 2
    chldBox.Left = prntBox.Left + (prntBox.Width - chldBox.Width) \ 2
End Sub
'=====================================================================================================================================================
' Set Bgrd for Data Sheet Forms
'=====================================================================================================================================================
Public Sub SetBgrdDataSheet(DataSheetFrm As Object, Optional BgrdColor As Long = 13152973)
        SetTableProperty DataSheetFrm, "DatasheetBackColor", 4, BgrdColor
End Sub









'===================================================================================================================================================
' Create Dynamic Form with Label on the fly
'===================================================================================================================================================
Public Function CreateForm(Prompt As String, Optional Title As String) As String
Dim f As Form, LBL As Label, dblWidth As Double

On Error GoTo ErrHandle
'------------------------------
'Application.Echo False
    
'Set f = CreateForm("")
'    myName = f.Name
'    f.RecordSelectors = False
'    f.NavigationButtons = False
'    f.DividingLines = False
'    f.ScrollBars = 0  ' none
'    f.PopUp = True
'    f.BorderStyle = acDialog
'    f.Modal = True
'    f.ControlBox = False
'    f.AutoResize = True
'    f.AutoCenter = True
'
    ' set the title
    '
'    If IsMissing(Title) Then
'        f.Caption = "Info"
'    Else
'        f.Caption = Title
'    End If
    
    ' add a label for the Prompt
    '
'    Set lbl = CreateControl(f.Name, acLabel)
'    lbl.Caption = Prompt
'    lbl.BackColor = 0 ' transparent
'    lbl.BorderColor = 0
'    lbl.Left = 100
'    lbl.Top = 100
'    If strFontName <> "" Then lbl.FontName = strFontName
'    If intFontSize > 0 Then lbl.FontSize = intFontSize
'    lbl.SizeToFit
'    dblWidth = lbl.Width + 200
'    f.Width = dblWidth - 200
'    f.Section(acDetail).Height = lbl.Height + 200
    
    ' display the form (first close and save it so that when
    ' it is reopened it will auto-centre itself)
    '
    
'    DoCmd.Close acForm, myName, acSaveYes
    
'     DoCmd.OpenForm myName
'    DoCmd.MoveSize , , dblWidth
'    DoCmd.RepaintObject acForm, myName

    ' turn screen repainting back on again
    '
'    Application.Echo True

    ' display form for specifed number of seconds
    '
'    If duration <= 0 Then duration = 2
'------------------------------
ExitHere:
    Exit Function
'--------------------
ErrHandle:
    ErrPrint "CreateForm", Err.Number, Err.Description
    Err.Clear
End Function




Public Property Get hWndApplication() As Long
    If m_hWnd = 0 Then
        If Application.Name = "Microsoft Access" Then
             m_hWnd = FindWindow("OMain", vbNullString)
        ElseIf Application.Name = "Microsoft Word" Then
            m_hWnd = FindWindow("OpusApp", vbNullString)
        ElseIf Application.Name = "Microsoft Excel" Then
            m_hWnd = FindWindow("XLMAIN", vbNullString)
        End If
    End If
    hWndApplication = m_hWnd
End Property



'===================================================================================================================================================
' Convert BOX Structure To Array
'===================================================================================================================================================
Public Function BOXToArray(zBOX As BOX) As Long()
Dim Arr(3) As Long
On Error GoTo ErrHandle
'----------------------------------------
   Arr(0) = zBOX.Left: Arr(1) = zBOX.Top
   Arr(2) = zBOX.Width: Arr(3) = zBOX.height
'----------------------------------------
ExitHere:
    BOXToArray = Arr '!!!!!!!!!!!!!!!
    Exit Function
'----------------
ErrHandle:
    ErrPrint "BOXToArray", Err.Number, Err.Description
    Err.Clear
End Function
'=================================================================================================================================================
' Show waiting message
'=================================================================================================================================================
Public Sub PleaseWait(Optional sText As String = "Please Wait", Optional nRepeat As Integer = 8, Optional WAIT_MILLESEC As Long = 300, _
                              Optional FONT_SIZE As Integer = 32, Optional WIN_HEIGHT As Long = 400, Optional WIN_WIDTH As Long = 800, _
                                                                                                                  Optional iSeries As Integer = 0)
Dim iSpec As Integer, sCode As String, I As Integer

Const C_Lowerbound As Integer = 1
Const C_Upperbound As Integer = 4

    On Error Resume Next
'------------------------------------------------
    Randomize
    
    For I = 0 To nRepeat
        iSpec = Int((C_Upperbound - C_Lowerbound + 1) * Rnd + C_Lowerbound)
        sCode = "Please Wait " & vbCrLf & GetSpectactor(iSpec, iSeries)
        Call ShowMsg(sCode, , , , , 2, , False, , , WIN_WIDTH, WIN_HEIGHT, WAIT_MILLESEC, FONT_SIZE)
    Next I
    
End Sub
'=================================================================================================================================================
' Функция выводит сообщение в заданный контекст
' EXAMPLES:  ShowMsg CHrW$(&H25A3),,,,,2
'            ShowMsg "Ti" & ChrW$(&H1EBF) & "ng Vi" & ChrW$(&H1EC7) & "t " & vbCRLF & "Unicode",,,,,2,,False
'            ShowMsg "TEST",,,,,,,,,,,,1000
'==================================================================================================================================================
Public Sub ShowMsg(sText As String, Optional Holst As Long = 1, Optional TextColor As Long = 16777215, Optional BgrdColor As Long = 16711680, _
                                  Optional iCharset As Integer = 204, Optional DrawTextMode As Integer = 0, Optional bCenter As Boolean = True, _
                                                  Optional bSingleLine As Boolean = True, Optional iLeft As Long = 0, Optional iTop As Long = 0, _
                                          Optional IWidth As Long = 400, Optional iHeight As Long = 200, Optional WaitForDissapiere As Long = 0, _
                                                                                                                Optional iFontSize As Integer)
Dim MyLabel As New cLabel

If sText = "" Then GoTo ExitHere
MyLabel.Text = sText: MyLabel.Holst = Holst
'---------------------------------------
MyLabel.ForeColor = TextColor: MyLabel.BackColor = BgrdColor: MyLabel.BorderTransparent = True: MyLabel.Transparency = False
MyLabel.Left = iLeft: MyLabel.Top = iTop: MyLabel.Width = IWidth: MyLabel.height = iHeight
MyLabel.Center = bCenter: MyLabel.SingleLine = bSingleLine: MyLabel.TextDrawMode = DrawTextMode
If iFontSize > 0 Then MyLabel.FontSize = iFontSize
'---------------------------------
        MyLabel.Draw
        
        If WaitForDissapiere > 0 Then
           Wait WaitForDissapiere
           MyLabel.Delete
        End If
'---------------------------------
ExitHere:
    Set MyLabel = Nothing
End Sub

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


'========================================================================================================================
' Context Menu
'Public Sub TestShowPopUpMenu()
'Dim sMenu As String
'    sMenu = "Default Menu;True" & vbCrLf & "-" & vbCrLf & "New Menu1" & vbCrLf & _
'           "New Menu2" & vbCrLf & "Disabled Menu;;;True" & vbCrLf & "Checked Menu;;True" & vbCrLf & _
'           "-" & vbCrLf & "Grayed Menu;;;;True"
'
'    MsgBox "User select case: " & ShowPopUpMenu(sMenu)
'End Sub
'========================================================================================================================
Public Function ShowPopUpMenu(smnu As String, Optional DLM As String = ";", Optional SPLT As String = vbCrLf) As String
Dim MNU As cPopUpMenu, MNUITEMS() As String, nDim As Integer
Dim I As Integer, sWork As String, CMNDS() As String, nCMNDS As Integer, J As Integer
Dim bDefault As Boolean, bChecked As Boolean, bDisabled As Boolean, bGrayed As Boolean, iRes As Long
Dim bSubMenu As Boolean

On Error GoTo ErrHandle
'----------------------------------------------------------
If smnu = "" Then Err.Raise "Wrong Menu Format"
   
MNUITEMS = Split(smnu, SPLT): nDim = UBound(MNUITEMS)
Set MNU = New cPopUpMenu
With MNU
    For I = 0 To nDim
        bDefault = False: bChecked = False: bDisabled = False: bGrayed = False
        CMNDS = Split(MNUITEMS(I), DLM): nCMNDS = UBound(CMNDS)
        For J = 0 To nCMNDS
            CMNDS(J) = Trim(CMNDS(J))
            Select Case J
            Case 0:
                 sWork = Trim(CMNDS(J))
                 MNUITEMS(I) = sWork
            Case 1:
                 If CMNDS(J) <> "" Then bDefault = CBool(CMNDS(J))
            Case 2:
                 If CMNDS(J) <> "" Then bChecked = CBool(CMNDS(J))
            Case 3:
                 If CMNDS(J) <> "" Then bDisabled = CBool(CMNDS(J))
            Case 4:
                 If CMNDS(J) <> "" Then bGrayed = CBool(CMNDS(J))
            End Select
        Next J
        .AddItem I, sWork, bDefault, bChecked, bDisabled, bGrayed
    Next I
    iRes = .ShowPopup()
End With
'----------------------------------------------------------
ExitHere:
    ShowPopUpMenu = MNUITEMS(iRes) '!!!!!!!!!!!!!!!
    Exit Function
'----------------
ErrHandle:
    ErrPrint "ShowPopUpMenu", Err.Number, Err.Description
    Err.Clear
End Function
'==========================================================================================================================================================
' Create the virtual var and add menu item
'    If sMenuIetem = "" Then re-initi Menu Builder
'==========================================================================================================================================================
Public Function PopUpAddItem(Optional sMenuIetem As String, Optional bDefault As Boolean, Optional bChecked As Boolean, Optional bDisabled As Boolean, _
                                         Optional bGrayed As Boolean, Optional DLM As String = ";", Optional MenuItemSplitter As String = vbCrLf, _
                                                                                                                   Optional iDLM As Integer = 29) As String
Dim sMenu As String, sTimer As String, sWork As String

Const LATENCY As Long = 1.2             ' Life Time in sec

On Error GoTo ErrHandle
'--------------------------
Const POP_MENU As String = "POPMENU"

If sMenuIetem = "" Then
    TempVars(POP_MENU).value = ""
    Exit Function
End If
'-----------------------------
sMenu = Nz(TempVars(POP_MENU).value, "")
If sMenu <> "" Then
   Call SplitTimeStamp(sMenu, sTimer, sMenu)
   If sTimer <> "" Then
       If Timer - CDbl(sTimer) > LATENCY Then sMenu = ""
   End If
End If
'-----------------------------
    sWork = sMenuIetem
    If bDefault Then sWork = sWork & DLM & "True"
    If bChecked Or bDisabled Or bGrayed Then
       If Not bDefault Then sWork = sWork & DLM
       sWork = sWork & DLM & TrueOrNothing(bChecked) & DLM & TrueOrNothing(bDisabled) & DLM & TrueOrNothing(bGrayed)
    End If
    
    sMenu = IIf(sMenu <> "", sMenu & MenuItemSplitter & sWork, sWork)
'-----------------------------
ExitHere:
    TempVars(POP_MENU).value = sMenu & Chr(iDLM) & Timer
    PopUpAddItem = sMenu '!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'------------
ErrHandle:
    ErrPrint "PopUpAddItem", Err.Number, Err.Description
End Function
Private Function TrueOrNothing(bVal As Boolean) As String
      TrueOrNothing = IIf(bVal, "True", "") '!!!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' The Function extract timestamp part from string
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub SplitTimeStamp(sInput As String, ByRef sTimer As String, ByRef sValue As String, Optional iDLM As Integer = 29)
Dim sRes() As String

On Error Resume Next
'-------------------------------
    sRes = Split(sInput, Chr(iDLM))
    sValue = sRes(0): sTimer = sRes(1)
End Sub

Public Sub RecheckFiles()
Dim sCheckSum As String, sPath As String
Dim RS As DAO.Recordset, bOnline As Boolean

Set RS = CurrentDb.OpenRecordset("q_FILES")
With RS
     If Not .EOF Then
         .MoveLast: .MoveFirst
         Do While Not .EOF
             sPath = .FIELDS("FullPath")
             bOnline = Dir(sPath) <> ""
             
             .Edit
                    If Not bOnline Then
                          .FIELDS("Status").value = 2
                    Else
                          .FIELDS("Status").value = 1
                    End If
                           .FIELDS("CRC32") = GetFileHash(sPath)
                           .FIELDS("FileSize") = GetFileSize(sPath)
                           .FIELDS("FileDate") = GetFileDate(sPath)
             .Update
             
             .MoveNext
         Loop
     End If
End With

ExitHere:
     Exit Sub
'--------------
End Sub


'======================================================================================================================================================
' PROGRESSBAR
'             EXAMPLE:
'               Public Sub TestProgressBar()
'               Dim I As Integer, DI As Double
'                   For I = 1 To 10
'                           DI = I / 10
'                           ShowProgressBar DI, 3, , "I = " & I
'                           DoEvents
'                           Wait 1000
'                    Next I
'                           DoEvents
'                           ShowProgressBar , 4, , "THE END"
'               End Sub
'======================================================================================================================================================
Public Sub ShowProgressBar(Optional CurValue As Double = 0, Optional iCommand As Integer = 1, _
                                                 Optional iFullValue As Integer = 100, Optional sTitle As String = "", _
                                           Optional bHourGlass As Boolean = True, Optional WaitForDissapiere As Long = 0)
Dim sProgressLine As String, iWait As Long
  
On Error GoTo ErrHandle
'------------------------------------------------------------
    Select Case iCommand
    Case 0:                  ' Start Standard ProgressBar
         If bHourGlass Then DoCmd.Hourglass True
         SysCmd acSysCmdInitMeter, sTitle, iFullValue
    Case 1:                  ' Show Progress via Standard ProgressBar
         SysCmd acSysCmdUpdateMeter, CInt(CurValue)
    Case 2:                  ' Remove Standard Progressbar
         If bHourGlass Then DoCmd.Hourglass False
         SysCmd acSysCmdRemoveMeter
    Case 3:                  ' Show Graphical Progress Bar (on Screen)
    
         sProgressLine = " " & GetProgressBarLine(CurValue) & " "
         If sTitle <> "" Then sProgressLine = sTitle & vbCrLf & sProgressLine
         
         Call ShowMsg(sProgressLine, , , , , 2, , False, , , , , WaitForDissapiere)
    
    Case 4:                  ' Wait and Remove Graphical Progressbar (on Screen)
         iWait = IIf(WaitForDissapiere = 0, 2000, WaitForDissapiere)
         Call ShowMsg(sTitle, , , , , 2, , False, , , , , iWait)
    Case Else
    End Select
'-----------------------------------------
ExitHere:

    Exit Sub
'-----------------------------------------
ErrHandle:
     Err.Clear
End Sub

'======================================================================================================================================================
' Draw ProgressBar via HOLST
'=====================================================================================================================================================
Public Function DrawProgressBar(dProgress As Double, Optional sText As String, Optional hWnd As Variant, Optional iHeight As Long = 100, _
                               Optional IWidth As Long = 400, Optional iBackColor As Long = 14702384, Optional iBorderColor As Long = vbBlue, _
                          Optional iForeColor As Long = 6684927, Optional iFontColor As Long = vbWhite, Optional bShowProc As Boolean = True) As Double
Dim PRGBOX As BOX, iLengt As Long, dRes As Double
Dim myHolst As New cHolst, ihDc As Variant, iHwnd As Variant

Const EDGE_LEN As Long = 5
'----------------------------------------------------------
    PRGBOX.height = iHeight: PRGBOX.Width = IWidth
    
    If Not IsMissing(hWnd) Then
           If hWnd <> 0 Then
              Call myHolst.SetHolst(GC_SPECIFIC_FORM, , iHwnd)
           End If
    Else
            myHolst.SetHolst (GC_DESKTOP)
            ihDc = myHolst.hdc: iHwnd = myHolst.hWnd
    End If
    
    
       If ihDc = 0 Then Err.Raise 1000, , "Can't get graphical context"
       PRGBOX.Left = myHolst.Left + (myHolst.Width - PRGBOX.Width) / 2: If PRGBOX.Left <= 0 Then Err.Raise 1000, , "Wrong Geometry"
       PRGBOX.Top = myHolst.Top + (myHolst.height - PRGBOX.height) / 2: If PRGBOX.Top <= 0 Then Err.Raise 1000, , "Wrong Geometry"
       iLengt = PRGBOX.Width * dProgress
       
       ' DRAW THE BACK
       Call myHolst.DrawBox(PRGBOX.Left, PRGBOX.Top, PRGBOX.Width, PRGBOX.height, , iBackColor, iBorderColor, , 2)
       ' DRAW THE FORE
       
       If dProgress = 0 Then GoTo ExitHere
       
       Call myHolst.DrawBox(PRGBOX.Left + EDGE_LEN, PRGBOX.Top + EDGE_LEN, iLengt, PRGBOX.height - 2 * EDGE_LEN, , iForeColor)
       ' DRAW THE TEXT
       If sText <> "" Then
        Call myHolst.DrawBox(PRGBOX.Left, PRGBOX.Top + EDGE_LEN, PRGBOX.Width, PRGBOX.height, , , , , , sText, iFontColor)
       Else
          If bShowProc Then
             Call myHolst.DrawBox(PRGBOX.Left, PRGBOX.Top + EDGE_LEN, PRGBOX.Width, PRGBOX.height, , , , , , Round(dProgress * 100, 2) & "%", iFontColor)
          End If
       End If
       dRes = dProgress
'----------------------------------------------------------
ExitHere:
    DrawProgressBar = dRes '!!!!!!!!!!!!!!!
    Set myHolst = Nothing
    Exit Function
'-------------------
ErrHandle:
    ErrPrint "DrawProgressBar", Err.Number, Err.Description
    Err.Clear: dRes = -1: Resume ExitHere
End Function
'===============================================================================================================================================
' Maximize Form
' Maximized Form should read TempVars!MaxForm = Me.Name in OnLoad  event and clear it on Unload
'===============================================================================================================================================
Public Function MaxForm(formName As String) As Boolean
Dim bRes As Boolean, sWork As String, iFormMode As Integer

On Error GoTo ErrHandle
'----------------------------------------------------------------------------------------------------------------
    If Nz(TempVars!MaxForm) = "" Then                                   ' No any form in Max Mode, so Maximize It
        TempVars!MaxForm = formName                                     ' Now going to reopen form
        iFormMode = acDialog
    ElseIf TempVars!MaxForm <> formName Then                            ' Some form is openning in Max, can't open two form simulat
        sWork = TempVars!MaxForm
        If MsgBox("Now the form " & sWork & " is open in Max Mode. Only One Maximized form is allowed" & _
           vbCrLf & "Should we close the form " & sWork & " and maximize the form " & _
           formName & "?", vbYesNoCancel + vbQuestion, "MaxForm") = vbYes Then
           TempVars!MaxForm = formName
           DoCmd.Close acForm, sWork, acSaveYes
           Sleep 1000: DoCmd.OpenForm formName, , , , , acDialog
        Else
           Exit Function
        End If
    Else                                                                ' Our Form in MAXX MODE, so Minimize it
        TempVars!MaxForm = "": iFormMode = acWindowNormal
    End If
'---------------------------------------------------------------------------------------------------------------------
    If IsFormLoaded(formName) Then
        DoCmd.Close acForm, formName, acSaveYes
        Sleep 1000
    End If
    DoCmd.OpenForm formName, , , , , iFormMode
'---------------------------------------------------------------------------
ExitHere:
        MaxForm = True '!!!!!!!!!!!
        Exit Function
    '--------------------------------
ErrHandle:
        ErrPrint "MaxForm", Err.Number, Err.Description
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
                                                                                                  Optional sModName As String = "#_GUI") As String
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


'--------------------------------------------------------------------------------------------------
' Function creates sime special string to imitate progress-bar
' Test for example: GetProgressBarLine(0.7, -1)
'-------------------------------------------------------------------------------------------------
Public Function GetProgressBarLine(PERC As Double, Optional RectLen As Integer = 20, _
                      Optional iSpec As Integer = -1, Optional SpecSeries As Integer = 0) As String
Dim FullChar As String, EmptyChar As String
Dim RealNum As Long, RestNum As Long, sRes As String

Const PERCMIN As Double = 0.05, PERCMAX As Double = 0.95
Const IEMPTY As Long = 9617, IFULL As Long = 9608

On Error Resume Next
'----------------------------------------------
      FullChar = ChrW$(IFULL): EmptyChar = ChrW$(IEMPTY)
      If PERC < PERCMIN Then
           sRes = String(RectLen, EmptyChar)
      ElseIf PERC > PERCMAX Then
           sRes = String(RectLen, FullChar)
      Else
            RealNum = CLng(PERC * RectLen): RestNum = RectLen - RealNum
            sRes = String(RealNum, FullChar) & String(RestNum, EmptyChar)
      End If
'---------------------------------
ExitHere:
      GetProgressBarLine = GetSpectactor(iSpec, SpecSeries) & sRes '!!!!!!!!!!!!!!!!
End Function
'---------------------------------------------------------------------------------------------------------------
' Function written to generate some special charts (UNICODE)
'---------------------------------------------------------------------------------------------------------------
Public Function GetSpectactor(Optional iSpec As Integer = -1, Optional SpecSeries As Integer = 0) As String
Dim SPECCODES(3, 3) As Long, Spectactor As String
Const upperBound As Long = 3, lowerBound As Long = 0

On Error GoTo ErrHandle
'------------------------------------------------------------------------------------
' SPECTACTOR MANAGEMENT (for more see http://www.codetable.net/ )

SPECCODES(0, 0) = 9626: SPECCODES(0, 1) = 9624: SPECCODES(0, 2) = 9630: SPECCODES(0, 3) = 9623           '  boxes
SPECCODES(1, 0) = 8987: SPECCODES(1, 1) = 10710: SPECCODES(1, 2) = 10711: SPECCODES(1, 3) = 10705        '  hourglasses
SPECCODES(2, 0) = 128337: SPECCODES(2, 1) = 128341: SPECCODES(2, 2) = 128349: SPECCODES(2, 3) = 128357   '  clocks
SPECCODES(3, 0) = 9286: SPECCODES(3, 1) = 9287: SPECCODES(3, 2) = 9288:: SPECCODES(3, 3) = 9289          '  ocr symbols
'-------------------------------------------------------------------
      Select Case iSpec
        Case 0:    Spectactor = ""                                                                                          ' Empty Spectacors
        Case 1:    Spectactor = ChrW$(SPECCODES(SpecSeries, 0)) & " "                                                       ' Symbol-0
        Case 2:    Spectactor = ChrW$(SPECCODES(SpecSeries, 1)) & " "                                                       ' Symbol-1
        Case 3:    Spectactor = ChrW$(SPECCODES(SpecSeries, 2)) & " "                                                       ' Symbol-2
        Case 4:    Spectactor = ChrW$(SPECCODES(SpecSeries, 3)) & " "                                                       ' Symbol-3
        Case -1:   Spectactor = ChrW$(SPECCODES(SpecSeries, Int((upperBound - lowerBound + 1) * Rnd + lowerBound))) & " "   ' Random Symbol
        Case Else:
      End Select
'----------------------------------------------------------------------------------------------
ExitHere:
    GetSpectactor = Spectactor '!!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'------------------------------
ErrHandle:
    ErrPrint "GetSpectactor", Err.Number, Err.Description
    Err.Clear
End Function




'====================================================================================================================================================
' ImageMagic Take screenshot
' -crop <width>{%}x<height>{%}{+-}<x offset>{+-}<y offset> (see http://kirste.userpage.fu-berlin.de/chemnet/use/suppl/imagemagick/www/import.html)
'====================================================================================================================================================
Public Function TakeScreenshot(Optional sPath As String, Optional bCrop As Boolean, Optional x As String = "+20", _
                                         Optional y As String = "+20", Optional Width As String = "400", Optional height As String = "300") As String
Dim sMagick As String, sCommand As String, sOutPut As String, sCrop As String


On Error GoTo ErrHandle
'-----------------------------------------------------------------------------
      sOutPut = IIf(sPath <> "", sPath, CurrentProject.Path & "\" & GenRandomStr(8, , , True) & ".png")
      If Dir(sOutPut) <> "" Then Kill sOutPut

      sMagick = GetExecutor("IMAGEMAGICK"): If sMagick = "" Then Exit Function
      sCrop = IIf(bCrop, "-crop " & Width & "x" & height & x & y, "")
      
      sCommand = "convert screenshot:" & sCrop
      Shell sMagick & " " & sCommand & " " & sOutPut
'-----------------------------------------------------------------------------
ExitHere:
      TakeScreenshot = sOutPut '!!!!!!!!!!!!!!!
      Exit Function
'--------------------
ErrHandle:
      ErrPrint "TakeScreenshot", Err.Number, Err.Description
      Err.Clear
End Function


'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Convert Array(3) To BOX Structure
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function ArrayToBOX(zARR As Variant) As BOX
Dim zBOX As BOX, nDim As Integer

On Error GoTo ErrHandle
'----------------------------------------
If Not IsArray(zARR) Then GoTo ExitHere
nDim = UBound(zARR): If nDim <> 3 Then Err.Raise 1000, , "Wrong Array Dimension"
    zBOX.Left = zARR(0): zBOX.Top = zARR(1)
    zBOX.Width = zARR(2): zBOX.height = zARR(3)
'----------------------------------------
ExitHere:
    ArrayToBOX = zBOX '!!!!!!!!!!!!!!!
    Exit Function
'----------------
ErrHandle:
    ErrPrint2 "ArrayToBox", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function



