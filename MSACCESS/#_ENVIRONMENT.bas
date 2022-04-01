Attribute VB_Name = "#_ENVIRONMENT"
'*********************************.ze$$e. **********************************************************************************************************
'              .ed$$$eee..      .$$$$$$$P""              ########  #######       #### ####### ##   ##  ##     #######            [MS ACCESS VERSION]
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
'                 $$$$$$$$$$F                                       ## ##  ##  ##  ##    ##   #####  ##   ## ##   ## ####  ####
'                  *$$$$$$$$"                                        ####  #####  ##     ##   ##     ###  ##  ##  ##  ##   ##  #
'                    "***""               _____________                                       ####   ## # ##   ## ##  ##   ####
' STANDARD MODULE WITH DEFAULT FUNCTIONS |v 2017/03/19 |                                      ##     ##  ###    ####  ##   ##  #
' The module contains frequently used functions and is part of the G-VBA library              #####  ##   ##     ### ####  ##  ##
'****************************************************************************************************************************************************
'****************************************************************************************************************************************************
Option Explicit

Private Const MOD_NAME As String = "#_ENVIRONMENT"
'********************************
Private Const NTM_REGULAR = &H40&
Private Const NTM_BOLD = &H20&
Private Const NTM_ITALIC = &H1&
Private Const TMPF_FIXED_PITCH = &H1
Private Const TMPF_VECTOR = &H2
Private Const TMPF_DEVICE = &H8
Private Const TMPF_TRUETYPE = &H4
Private Const ELF_VERSION = 0
Private Const ELF_CULTURE_LATIN = 0
Private Const RASTER_FONTTYPE = &H1
Private Const DEVICE_FONTTYPE = &H2
Private Const TRUETYPE_FONTTYPE = &H4
Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64

Private Const SM_CMONITORS              As Long = 80    ' number of display monitors
Private Const SM_CXVIRTUALSCREEN = 78
Private Const SM_CYVIRTUALSCREEN = 79

Private Const MONITOR_CCHDEVICENAME     As Long = 32    ' device name fixed length
Private Const MONITOR_PRIMARY           As Long = 1
Private Const MONITOR_DEFAULTTONULL     As Long = 0
Private Const MONITOR_DEFAULTTOPRIMARY  As Long = 1
Private Const MONITOR_DEFAULTTONEAREST  As Long = 2

Private Const LOCALE_SLIST = &HC         '  list item separator

Private Const SYS_OUT_OF_MEM        As Long = &H0
Private Const ERROR_FILE_NOT_FOUND  As Long = &H2
Private Const ERROR_PATH_NOT_FOUND  As Long = &H3
Private Const ERROR_BAD_FORMAT      As Long = &HB
Private Const NO_ASSOC_FILE         As Long = &H1F
Private Const MIN_SUCCESS_LNG       As Long = &H20
Private Const MAX_PATH              As Long = &H104

Private Const USR_NULL              As String = "NULL"
Private Const S_DIR                 As String = "C:\" '// Change as required (drive that .exe will be on)

Public Type SOFTWARE
    Caption As String
    Description As String
    Version As String
    InstallDate As Date
    IdentifyingNumber As String
    InstallLocation As String
    InstallState As String
    Name As String
    PackageCache As String
    SKUNumber As String
    Vendor As String
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
   lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type NEWTEXTMETRIC
   tmHeight As Long
   tmAscent As Long
   tmDescent As Long
   tmInternalLeading As Long
   tmExternalLeading As Long
   tmAveCharWidth As Long
   tmMaxCharWidth As Long
   tmWeight As Long
   tmOverhang As Long
   tmDigitizedAspectX As Long
   tmDigitizedAspectY As Long
   tmFirstChar As Byte
   tmLastChar As Byte
   tmDefaultChar As Byte
   tmBreakChar As Byte
   tmItalic As Byte
   tmUnderlined As Byte
   tmStruckOut As Byte
   tmPitchAndFamily As Byte
   tmCharSet As Byte
   ntmFlags As Long
   ntmSizeEM As Long
   ntmCellHeight As Long
   ntmAveWidth As Long
End Type

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type MONITORINFOEX
   cbSize As Long
   rcMonitor As RECT
   rcWork As RECT
   dwFlags As Long
   szDevice As String * MONITOR_CCHDEVICENAME
End Type

Private Enum DevCap     ' GetDeviceCaps nIndex (video displays)
    HORZSIZE = 4        ' width in millimeters
    VERTSIZE = 6        ' height in millimeters
    HORZRES = 8         ' width in pixels
    VERTRES = 10        ' height in pixels
    BITSPIXEL = 12      ' color bits per pixel
    LOGPIXELSX = 88     ' horizontal DPI (assumed by Windows)
    LOGPIXELSY = 90     ' vertical DPI (assumed by Windows)
    COLORRES = 108      ' actual color resolution (bits per pixel)
    VREFRESH = 116      ' vertical refresh rate (Hz)
End Enum

Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion(0 To 127) As Byte      '  Maintenance string for PSS usage
End Type

'********************

#If Win64 Then
    Private Declare PtrSafe Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExA" (ByVal hdc As LongPtr, _
                                lpLogFont As LOGFONT, ByVal lpEnumFontProc As LongPtr, ByVal lParam As LongPtr, ByVal dW As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
        
    Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (pguid As GUID) As Long
    Private Declare PtrSafe Function StringFromGUID2 Lib "ole32" ( _
        rguid As GUID, ByVal lpsz As LongPtr, ByVal cchMax As Long) As Long
        
    Private Declare PtrSafe Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, _
               ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
    Private Declare PtrSafe Function GetUserDefaultLCID Lib "kernel32" () As Long
    Private Declare PtrSafe Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
    Private Declare PtrSafe Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Private Declare PtrSafe Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
    
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function MonitorFromWindow Lib "user32" _
                            (ByVal hWnd As LongPtr, ByVal dwFlags As Long) As LongPtr
    Private Declare PtrSafe Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" _
            (ByVal hMonitor As LongPtr, ByRef lpMI As MONITORINFOEX) As Boolean
    Private Declare PtrSafe Function CreateDC Lib "gdi32" Alias "CreateDCA" _
                    (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As LongPtr) As LongPtr
    Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hdc As LongPtr) As Long
    
    Private Declare PtrSafe Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" _
        (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
    
    Private hWnd As LongPtr
    Private hdc As LongPtr
    
#Else
    Private Declare Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExA" (ByVal hdc As Long, lpLogFont As LOGFONT, _
                                                           ByVal lpEnumFontProc As Long, ByVal lParam As Long, ByVal dW As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

    Private Declare Function CoCreateGuid Lib "ole32" (pguid As GUID) As Long
    Private Declare Function StringFromGUID2 Lib "ole32" (rguid As GUID, ByVal lpsz As Long, ByVal cchMax As Long) As Long

    Private Declare Function GetLocaleInfo Lib "kernel32" Alias _
                "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, _
                       ByVal lpLCData As String, ByVal cchData As Long) As Long
    Private Declare Function GetUserDefaultLCID% Lib "kernel32" ()
    Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
    Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
    Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByVal lpVersionInformation As OSVERSIONINFO) As Long
    
    Private Declare Function GetSystemMetrics Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Private Declare Function GetActiveWindow Lib "user32" Alias "GetActiveWindow" () As Long
    Private Declare Function MonitorFromWindow Lib "user32" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long
    Private Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" ( _
                        ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
    Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As Long
    Private Declare Function DeleteDC Lib "gdi32" Alias "DeleteDC" (ByVal hdc As Long) As Long
    
    Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" _
        (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
    
    Private hWnd As Long
    Private hdc As Long
#End If



'****************************************************************************************************************************************************
Private FontArray() As String   'The Array that will hold all the Fonts (needed for sorting)
Private FntInc As Integer       'The FontArray element incremental counter.
'****************************************************************************************************************************************************
'====================================================================================================================================================
' GRACLE VERSION
'====================================================================================================================================================
Public Function GetGrackleVersion(Optional GrackleName As String = "_GRACKLE") As String
Dim sFile As String, sDate As String

On Error Resume Next
'------------------------
   sFile = Application.VBE.VBProjects("_GRACKLE").fileName
   If sFile = "" Then Exit Function
   sDate = Format(FileDateTime(sFile), "YYYYMMDDHHNNSS")
'------------------------
ExitHere:
   GetGrackleVersion = sDate '!!!!!!!!!!!!!
End Function

'======================================================================================================================================================
' Get GracklePath
'======================================================================================================================================================
Public Function GetGracklePath() As String
Dim sRes As String, sTest As String, iL As Integer
     
Const GRACKLE_NAME As String = "GRACKLE.addb"
Const GRACKLE_PARAM As String = "GRACKLE_PATH"
    
    On Error GoTo ErrHandle
'----------------------
sTest = GetEnviron(GRACKLE_PARAM)
If sTest <> "" Then
   sRes = sTest: GoTo ExitHere
End If

sTest = Environ("HOMEDRIVE") & Environ("HOMEPATH") & "\Google Drive\_ZWORKS\_DATABASES\VBALIB\" & GRACKLE_NAME
If Dir(sTest) <> "" Then
   sRes = sTest: GoTo ExitHere
End If

sTest = SearchFiles("VBALIB", Environ("HOMEDRIVE") & Environ("HOMEPATH"))(0)
If sTest <> "" Then
   iL = InStr(1, sTest, "VBALIB")
   If iL > 0 Then sRes = Left(sTest, iL + Len("VBALIB")) & GRACKLE_NAME
End If
'----------------------
ExitHere:
    GetGracklePath = sRes '!!!!!!!!!!!
    Exit Function
'---------------
ErrHandle:
    ErrPrint2 "", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'======================================================================================================================================================
' Set Path to Gratis in Environment Path (Require restart Access after this action)
'======================================================================================================================================================
Public Sub SetGRACKLEPath()
Dim sPath As String, sMsg As String
   
Const GRACKLE_PARAM As String = "GRACKLE_PATH"

   On Error GoTo ErrHandle
'--------------------------
   sMsg = "Do you want set the path for Gratis VB Lib? Be Carefull it could lead to broke current macros." & vbCrLf & _
             "Please, be aware that this command start from GRATIS Lib. Should we strt right now?"
             
   If MsgBox(sMsg, vbYesNoCancel + vbQuestion, "Set Grackle Path") <> vbYes Then Exit Sub
   
   sPath = CurrentDb.Name
   If UCase(FilenameWithoutExtension(sPath)) <> "GRACKLE" Then Err.Raise 10005, , "This code works only inside in Gracle DB"
   
   Call SetEnvironWin(GRACKLE_PARAM, sPath)
'--------------------------
ExitHere:
   Exit Sub
'----------
ErrHandle:
   ErrPrint2 "SetGratisPath", Err.Number, Err.Description, MOD_NAME
   Err.Clear
End Sub

'======================================================================================================================================================
' Get Media Path for standard sounds
'======================================================================================================================================================
Public Function GetMediaPath(Optional SysSound As String = "notify.wav") As String
Dim sRes As String, sPath As String

Const MEDIA_FLDR As String = "Media"

    On Error Resume Next
'-------------------
    sPath = Environ("windir") & "\" & MEDIA_FLDR
    If Dir(sPath, vbDirectory) = "" Then Exit Function
    
    sPath = sPath & "\" & SysSound
    If Dir(sPath) = "" Then Exit Function
    sRes = sPath
'-------------------
ExitHere:
    GetMediaPath = sRes '!!!!!!!!!!!!!!!
End Function

'======================================================================================================================================================
' Create GUID
'======================================================================================================================================================
Public Function CreateGUID() As String
 Dim NewGUID As GUID
 CoCreateGuid NewGUID
 CreateGUID = Space$(38)
 StringFromGUID2 NewGUID, StrPtr(CreateGUID), 39
End Function
'========================================================================================================================================================
' Check if input is GUID, then return normilized string for success and empty for false
'=======================================================================================================================================================
Public Function IsGuid(sID As Variant) As String
Dim mySID As String, sRes As String

On Error GoTo ErrHandle
'------------------------------------------------
If varType(sID) = vbLong Then Exit Function

    mySID = StringFromGUID(sID)
    '-------------------------------------------------------------------------------
    If Left(mySID, 6) = "{guid " Then mySID = Right(mySID, Len(mySID) - 6)
    If Right(mySID, 2) = "}}" Then mySID = Left(mySID, Len(mySID) - 1)
    '-------------------------------------------------------------------------------
        If (Len(mySID) = 38) Then
            If (Mid(mySID, 10, 1) = "-") Then
                If (Mid(mySID, 15, 1) = "-") Then
                    If (Mid(mySID, 20, 1) = "-") Then
                        If (Mid(mySID, 25, 1) = "-") Then
                            If (Left(mySID, 1) = "{") Then
                                If (Right(mySID, 1) = "}") Then
                                        sRes = mySID
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
'--------------------------------------------------------------------------------
ExitHere:
        IsGuid = sRes '!!!!!!!!!!!!!!!!!!!!!
        Exit Function
'-----------------------
ErrHandle:
        Select Case Err.Number
        Case 13:
            Err.Clear
        Case Else:
            ErrPrint "IsGuid", Err.Number, Err.Description
            Err.Clear
        End Select
End Function
'======================================================================================================================================================
' Get Screen resolution
'======================================================================================================================================================
Public Function GetScreenResolution() As String
Dim Sezam As New cSize, hWnd As Variant, WinDim() As Long
    hWnd = Sezam.GetDeskTopHandler()
    WinDim = Sezam.GetWindowCoordinates(hWnd)
'-------------------------
    GetScreenResolution = WinDim(2) & " x " & WinDim(3)  '!!!!!!!!!!
    Set Sezam = Nothing
End Function
'======================================================================================================================================================
' Get Access Version
'======================================================================================================================================================
Public Function GetAccessEXEVersion() As String
Dim sAccessVerNo As String

    On Error Resume Next
'-----------------------
    sAccessVerNo = SysCmd(acSysCmdAccessVer) & "." & SysCmd(715)
    Select Case sAccessVerNo
        'Access 2000
        Case "9.0.0.0000" To "9.0.0.2999": GetAccessEXEVersion = "Microsoft Access 2000 - Build:" & sAccessVerNo
        Case "9.0.0.3000" To "9.0.0.3999": GetAccessEXEVersion = "Microsoft Access 2000 SP1 - Build:" & sAccessVerNo
        Case "9.0.0.4000" To "9.0.0.4999": GetAccessEXEVersion = "Microsoft Access 2000 SP2 - Build:" & sAccessVerNo
        Case "9.0.0.6000" To "9.0.0.6999": GetAccessEXEVersion = "Microsoft Access 2000 SP3 - Build:" & sAccessVerNo
        'Access 2002
        Case "10.0.2000.0" To "10.0.2999.9": GetAccessEXEVersion = "Microsoft Access 2002 - Build:" & sAccessVerNo
        Case "10.0.3000.0" To "10.0.3999.9": GetAccessEXEVersion = "Microsoft Access 2002 SP1 - Build:" & sAccessVerNo
        Case "10.0.4000.0" To "10.0.4999.9": GetAccessEXEVersion = "Microsoft Access 2002 SP2 - Build:" & sAccessVerNo
        'Access 2003
        Case "11.0.0000.0" To "11.0.5999.9999": GetAccessEXEVersion = "Microsoft Access 2003 - Build:" & sAccessVerNo
        Case "11.0.6000.0" To "11.0.6999.9999": GetAccessEXEVersion = "Microsoft Access 2003 SP1 - Build:" & sAccessVerNo
        Case "11.0.7000.0" To "11.0.7999.9999": GetAccessEXEVersion = "Microsoft Access 2003 SP2 - Build:" & sAccessVerNo
        Case "11.0.8000.0" To "11.0.8999.9999": GetAccessEXEVersion = "Microsoft Access 2003 SP3 - Build:" & sAccessVerNo
        'Access 2007
        Case "12.0.0000.0" To "12.0.5999.9999": GetAccessEXEVersion = "Microsoft Access 2007 - Build:" & sAccessVerNo
        Case "12.0.6000.0" To "12.0.6422.9999": GetAccessEXEVersion = "Microsoft Access 2007 SP1 - Build:" & sAccessVerNo
        Case "12.0.6423.0" To "12.0.5999.9999": GetAccessEXEVersion = "Microsoft Access 2007 SP2 - Build:" & sAccessVerNo
        'Unable to locate specific build versioning for SP3 - to be validated at a later date.
        '  Hopefully MS will eventually post the info on their website?!
        Case "12.0.6000.0" To "12.0.9999.9999": GetAccessEXEVersion = "Microsoft Access 2007 SP3 - Build:" & sAccessVerNo
        'Access 2010
        Case "14.0.0000.0000" To "14.0.6022.1000": GetAccessEXEVersion = "Microsoft Access 2010 - Build:" & sAccessVerNo
        Case "14.0.6023.1000" To "14.0.7014.9999": GetAccessEXEVersion = "Microsoft Access 2010 SP1 - Build:" & sAccessVerNo
        Case "14.0.7015.1000" To "14.0.9999.9999": GetAccessEXEVersion = "Microsoft Access 2010 SP2 - Build:" & sAccessVerNo
        'Access 2013
        Case "15.0.0000.0000" To "15.0.4569.1505": GetAccessEXEVersion = "Microsoft Access 2013 - Build:" & sAccessVerNo
        Case "15.0.4569.1506" To "15.0.9999.9999": GetAccessEXEVersion = "Microsoft Access 2013 SP1 - Build:" & sAccessVerNo
        'Access 2016
        '   See: https://support.office.com/en-us/article/Version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7#bkmk_byversion
        '   To build a proper sequence for all the 2016 versions!
        Case "16.0.0000.0000" To "16.0.9999.9999": GetAccessEXEVersion = "Microsoft Access 2016 - Build:" & sAccessVerNo
        Case Else: GetAccessEXEVersion = "Unknown Version"
    End Select
    If SysCmd(acSysCmdRuntime) Then GetAccessEXEVersion = GetAccessEXEVersion & " Run-time"
End Function
'======================================================================================================================================================
' Getting Workstation Name
'======================================================================================================================================================
Public Function ComputerName() As String
    ComputerName = Environ$("ComputerName")
End Function

'======================================================================================================================================================
' Getting UserName Name
'======================================================================================================================================================
Public Function UserName() As String
    UserName = Environ$("UserName")
End Function

'======================================================================================================================================================
' How Many Monitirs
'======================================================================================================================================================
Public Function GetMonitors() As Long
    GetMonitors = GetSystemMetrics(SM_CMONITORS)
End Function
'======================================================================================================================================================
' Max of Horisontal Resolution in Points
'======================================================================================================================================================
Public Function GetHorizontalResolution() As Long
    GetHorizontalResolution = GetSystemMetrics(SM_CXVIRTUALSCREEN)
End Function
'======================================================================================================================================================
' Max of Vertical Resolution in Points
'=====================================================================================================================================================
Public Function GetVerticalResolution() As Long
    GetVerticalResolution = GetSystemMetrics(SM_CYVIRTUALSCREEN)
End Function
'======================================================================================================================================================
' Display Info
'=====================================================================================================================================================
Public Function DisplayInfo(Optional DLM As String = ";") As String
Dim sRes As String

    On Error Resume Next
'-----------------------
    sRes = "Total monitors: " & GetMonitors & DLM & " Max Screen Resolution: " & GetScreenResolution
'-----------------------
ExitHere:
    DisplayInfo = sRes '!!!!!!!!!!!
End Function

'========================================================================================================================================================
' All IP ddresses
'========================================================================================================================================================
Public Function GetIPAddress(Optional DLM As String = ";", Optional FLDSEP As String = ",") As String
    Const StrComputer As String = "."   ' Computer name. Dot means local computer
    Dim oWMI, IPConfigSet, IPConfig, IPAddress, I
    Dim strIPAddress As String

On Error GoTo ErrHandle
'------------------------------
    ' Connect to the WMI service
    Set oWMI = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\" & StrComputer & "\root\cimv2")

    ' Get all TCP/IP-enabled network adapters
    Set IPConfigSet = oWMI.ExecQuery _
        ("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")

    ' Get all IP addresses associated with these adapters
    For Each IPConfig In IPConfigSet
        IPAddress = IPConfig.IPAddress
        If Not IsNull(IPAddress) Then
            strIPAddress = IIf(strIPAddress <> "", strIPAddress & DLM, "")
            strIPAddress = strIPAddress & Join(IPAddress, FLDSEP)
        End If
    Next
'------------------------------
ExitHere:
    GetIPAddress = strIPAddress '!!!!!!!!!
    Exit Function
'----------------
ErrHandle:
    ErrPrint "GetIPAddress", Err.Number, Err.Description
    Err.Clear
End Function
'========================================================================================================================================================
' Function Print Environment Info
'========================================================================================================================================================
Public Function EnvInfo(Optional DLM As String = ";") As String
Dim sRes As String, Arr(6) As String, bMac As Boolean

    On Error Resume Next
'--------------------

bMac = IsMac()                                                 ' Check is Mac or Windows
If Not bMac Then                  ' WINDOWS ENVIRONMENT
    Arr(0) = Application.Name
    #If Win64 Then                                              ' BITs
            Arr(0) = Arr(0) & " 64-bit"
    #Else
            Arr(0) = Arr(0) & " 32-bit"
    #End If
    
    Select Case Application.Name
    Case "Microsoft Access":
        Arr(1) = GetAccessEXEVersion                            ' ACCESS VERSION
        Arr(2) = GetOSName                                      ' OS NAME
        If Arr(2) <> "" Then
                 Arr(2) = Split(Arr(2), "|")(0)
        Else
                 Arr(2) = "Microsoft Windows"
        End If
    Case "Microsoft Excel":
        Arr(1) = Application.Version                            ' EXCEL VERSION
        'ARR(2) = Application.OperatingSystem                   ' OS NAME
    End Select
    
    Arr(3) = "Language ID: " & GetLanguageID()                  ' Get Language ID
    Arr(3) = "Processor: " & ProcessorInfo()                   ' PROCESSOR INFO
    Arr(4) = "RAM: " & MemoryInfo()                             ' AVAILABLE MEMORY
    Arr(5) = DisplayInfo()
    
    sRes = GetExecutor("IRFANVIEW"): If sRes <> "" Then Arr(6) = "IRFANVIEW"
    sRes = GetExecutor("IMAGEMAGICK"): If sRes <> "" Then Arr(6) = ConcateString(Arr(6), "IMAGEMAGICK", ",")
    sRes = GetExecutor("GHOSTSCRIPT"): If sRes <> "" Then Arr(6) = ConcateString(Arr(6), "GHOSTSCRIPT", ",")
    sRes = GetExecutor("OPENSSL"): If sRes <> "" Then Arr(6) = ConcateString(Arr(6), "OPENSSL", ",")
    
    Arr(6) = "External Libs: " & Arr(6)
   
    sRes = Join(Arr, DLM)
Else                               ' MAC ENVIRONMENT
    sRes = "MAC OFFICE"
End If
'--------------------
ExitHere:
    EnvInfo = sRes '!!!!!!!!!!!!!!!
End Function
'======================================================================================================================================================
' Get Processor Information
'======================================================================================================================================================
Public Function ProcessorInfo() As String
Dim sRes As String, oWinMng As Object, oCPU As Object

    On Error GoTo ErrHandle
'---------------------
      Set oWinMng = GetObject("WinMgmts:").instancesof("Win32_Processor")
      
      For Each oCPU In oWinMng
            sRes = IIf(sRes <> "", sRes & vbCrLf, "") & oCPU.Name & " " & oCPU.CurrentClockSpeed & " Mhz"
      Next
'---------------------
ExitHere:
    ProcessorInfo = sRes '!!!!!!!!
    Set oCPU = Nothing: Set oWinMng = Nothing
    Exit Function
'------------
ErrHandle:
    ErrPrint2 "ProcessorInfo", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Resume ExitHere
    
End Function

'======================================================================================================================================================
' Get Processor Information
'======================================================================================================================================================
Public Function MemoryInfo() As String
Dim oMem As Object, oMEMS As Object, dRam As Double
Dim sRes As String

    On Error GoTo ErrHandle
'--------------------------
Set oMEMS = GetObject("winmgmts:").ExecQuery("SELECT * FROM Win32_PhysicalMemory")
    For Each oMem In oMEMS
        dRam = dRam + oMem.Capacity
    Next
    sRes = Format(dRam / 1024 / 1024, "## ##0") & " MB"
'--------------------------
ExitHere:
    MemoryInfo = sRes '!!!!!!!!
    Set oMem = Nothing: Set oMEMS = Nothing
    Exit Function
'---------------
ErrHandle:
    ErrPrint2 "MemoryInfo", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Resume ExitHere
End Function

'========================================================================================================================================================
' Get Language ID (LanguageSet ={msoLanguageIDInstall =1;msoLanguageIDUI =2;msoLanguageIDHelp=3;msoLanguageIDExeMode=4})
' RETURN: 1049 - RU; 1033 - EN-US; 2057 - EN-UK; 3084 - French-Canada; 1031 - German (см. https://msdn.microsoft.com/en-us/goglobal/bb964664.aspx)
'========================================================================================================================================================
Public Function GetLanguageID(Optional LanguageSet As Integer = 4) As String
Dim mli As Variant
        
        mli = Application.LanguageSettings.LanguageID(LanguageSet)
        GetLanguageID = CStr(mli)
End Function

'========================================================================================================================================================
' Check Mac or PC
'========================================================================================================================================================
Public Function IsMac() As Boolean
Dim bRes As Boolean

    On Error Resume Next
'-------------
#If Mac Then
    bRes = True
#End If
'-------------
ExitHere:
    IsMac = bRes '!!!!!!!!
End Function
'========================================================================================================================================================
' Return Serach result as delimited string
'========================================================================================================================================================
Public Function SearchFilesAll(sWord As String, Optional sFolder As String = "C:\Users\valer\_REF", Optional iTopOnly As Integer = 5, _
                            Optional bDataModified As Boolean, Optional bIncludeSubFolders As Boolean = True, Optional DLM As String = ";") As String
Dim sFiles() As String

On Error Resume Next
'------------------------------------
    sFiles = SearchFiles(sWord, sFolder, iTopOnly, bDataModified, bIncludeSubFolders)
    If sFiles(0) <> "" Then SearchFilesAll = Join(sFiles, DLM) '!!!!!!!!!!!!!
End Function

'========================================================================================================================================================
' Search Windows Serach Indexes
'========================================================================================================================================================
Public Function SearchFiles(sWord As String, sFolder As String, Optional iTopOnly As Integer = 5, Optional bDataModified As Boolean, _
                                                                                               Optional bIncludeSubFolders As Boolean = True) As String()

Dim objConnection As Object, objRecordSet As Object
Dim sRes() As String, nDim As Integer, sSQL As String, sWhere As String, sOrderBy As String

On Error GoTo ErrHandle
'------------------------------------------------------------
If sWord = "" Then Exit Function
If sFolder = "" Then Exit Function
If Dir(sFolder, vbDirectory) = "" Then Exit Function

Set objConnection = CreateObject("ADODB.Connection"): Set objRecordSet = CreateObject("ADODB.Recordset")

nDim = -1: ReDim sRes(0)
sSQL = "SELECT " & IIf(iTopOnly > 0, "Top " & iTopOnly & " ", "") & "System.ItemPathDisplay" & _
        IIf(bDataModified, ", System.DateModified", "") & " FROM SYSTEMINDEX "
sWhere = "FREETEXT(" & sCH(sWord) & ") AND " & IIf(bIncludeSubFolders, "SCOPE", "DIRECTORY") & " = " & sCH(sFolder)
sOrderBy = "System.FileName"
sSQL = sSQL & " WHERE " & sWhere & " ORDER BY " & sOrderBy


objConnection.Open "Provider=Search.CollatorDSO;Extended Properties='Application=Windows';"
objRecordSet.Open sSQL, objConnection

With objRecordSet
   If Not .EOF Then
            .MoveFirst
        Do Until objRecordSet.EOF
            nDim = nDim + 1: ReDim Preserve sRes(nDim)
            sRes(nDim) = .FIELDS.Item("System.ItemPathDisplay")
            .MoveNext
        Loop
   End If
End With
'------------------------------------------------------------
ExitHere:
    SearchFiles = sRes '!!!!!!!!!!!!!!!!!!!!
    Set objRecordSet = Nothing: Set objConnection = Nothing:
    Exit Function
'----------------
ErrHandle:
    ErrPrint "SearchFiles", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function

'========================================================================================================================================================
' Getting Serial of HardDrive
'========================================================================================================================================================
Public Function HDDSerial() As String
On Error Resume Next
   HDDSerial = Hex$(CreateObject("Scripting.FileSystemObject").GetDrive("C").SerialNumber)
End Function

'====================================================================================================================================================
' Copy Text to Clipboard
'====================================================================================================================================================
Public Sub ToClipBoard(ByVal sText As String)
Dim objClipboard As Object

On Error Resume Next
'--------------------------------------
  Set objClipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
  objClipboard.SetText sText
  objClipboard.PutInClipboard
 
  Set objClipboard = Nothing
End Sub
'====================================================================================================================================================
' Extract Text From Clipboard
'====================================================================================================================================================
Public Function FromClipboard() As String
Dim objClipboard As Object
  
On Error Resume Next
'----------------------
  Set objClipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
 
  objClipboard.GetFromClipboard
  FromClipboard = objClipboard.GetText   '!!!!!!!!!!!!!!!!!!!!
 
  Set objClipboard = Nothing
End Function
'====================================================================================================================================================
' File to Clipboard
'====================================================================================================================================================
Public Function FileToClipBoard(sPath As String) As Boolean
Dim MyClipboard As New cClipBoard, bRes As Boolean

On Error GoTo ErrHandle
'---------------------------
    If sPath = "" Then Exit Function
    If Dir(sPath) = "" Then Exit Function
    
    bRes = MyClipboard.ClipboardCopySingleFile(sPath)
'---------------------------
ExitHere:
    FileToClipBoard = True '!!!!!!!!!!!!!!!!!!
    Exit Function
'-----------------
ErrHandle:
    ErrPrint "FileToClipBoard", Err.Number, Err.Description
    Err.Clear
End Function
'====================================================================================================================================================
' Text From Clipboard To Text File
'====================================================================================================================================================
Public Function TextFromClipboardToFile(sFile As String) As String
Dim sRes As String, sText As String

On Error GoTo ErrHandle
'---------------------------------
    If sFile = "" Then Exit Function
    If Not IsTextInClipboard() Then Exit Function
    
    sText = FromClipboard()
    If sText = "" Then Exit Function
    If Dir(sFile) <> "" Then Kill sFile
    WriteStringToFile sFile, sText
    sRes = Dir(sFile)
'---------------------------------
ExitHere:
    TextFromClipboardToFile = sRes '!!!!!!!!!!
    Exit Function
'------------
ErrHandle:
    ErrPrint2 "", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'====================================================================================================================================================
' Check if Text is in Clipboard
'====================================================================================================================================================
Public Function IsTextInClipboard() As Boolean
Dim MyClipboard As New cClipBoard

    On Error Resume Next
'--------------------
    IsTextInClipboard = MyClipboard.Is_Txt_in_Clipboard()  '!!!!!!!
    Set MyClipboard = Nothing
End Function

'====================================================================================================================================================
' Check if Imafe is in Clipboard
'====================================================================================================================================================
Public Function IsImageInClipboard() As Boolean
Dim MyClipboard As New cClipBoard

    On Error Resume Next
'--------------------
    IsImageInClipboard = MyClipboard.Is_Pic_in_Clipboard() '!!!!!!!
    Set MyClipboard = Nothing
End Function
'====================================================================================================================================================
' Save Image From Clipboard
'====================================================================================================================================================
Public Function ImageFromClipboard(Optional sFile As String, Optional sAddOption As String) As String
Dim MyImage As New cImage, sFolder As String, sFileName As String, sRes As String
  
On Error Resume Next
'----------------------
  If sFile <> "" Then
      sFileName = FileNameOnly(sFile)
      sFolder = FolderNameOnly(sFile)
  End If
  
  sRes = MyImage.SaveClipboardToImage(sFileName, sFolder, sAddOption)
'----------------------
ExitHere:
  ImageFromClipboard = sRes '!!!!!!!!!!!!!
  Set MyImage = Nothing
  Exit Function
'--------------------
ErrHandle:
  ErrPrint "ImageFromClipboard", Err.Number, Err.Description
  Err.Clear
End Function
'====================================================================================================================================================
' Add Registry Key. The Prefix should be start from:
' HKEY_CURRENT_USER (or HKCU); HKEY_LOCAL_MACHINE (or HKLM); HKEY_CLASSES_ROOT (or HKCR); HKEY_USERS (or HKEY_USERS)
' HKEY_CURRENT_CONFIG; HKEY_CURRENT_CONFIG
'====================================================================================================================================================
Public Sub RegKeySave(i_RegKey As String, i_Value As String, Optional i_Type As String = "REG_SZ", Optional sKeyPrefix As String = "HKCU\Software\")
Dim myWS As Object
   
   On Error GoTo ErrHandle
'---------------------
  Set myWS = CreateObject("WScript.Shell")
  myWS.RegWrite i_RegKey, i_Value, i_Type
'---------------------
ExitHere:
  Exit Sub
'----------
ErrHandle:
  ErrPrint "RegKeySave", Err.Number, Err.Description, MOD_NAME
  Err.Clear
End Sub
'====================================================================================================================================================
' Read Specific registry Key
'====================================================================================================================================================
Public Function RegKeyRead(s_RegKey As String) As String
Dim myWS As Object
  
  On Error GoTo ErrHandle
'-------------------------------------
  
  Set myWS = CreateObject("WScript.Shell")
  RegKeyRead = myWS.RegRead(s_RegKey)
'-------------------------------------
ExitHere:
  
  Set myWS = Nothing
  Exit Function
'----------------
ErrHandle:
  ErrPrint "RegKeyRead", Err.Number, Err.Description
  Err.Clear: Resume ExitHere
End Function

'====================================================================================================================================================
' Checking if a Registry key exists
'====================================================================================================================================================
Public Function RegKeyExists(s_RegKey As String) As Boolean
Dim myWS As Object

On Error GoTo ErrHandle
'----------------------------
Set myWS = CreateObject("WScript.Shell")
  myWS.RegRead s_RegKey
  RegKeyExists = True
'----------------------------------
ExitHere:
    Set myWS = Nothing
    Exit Function
'-----------------
ErrHandle:
    Err.Clear: Resume ExitHere
End Function


'======================================================================================================================================================
' List All Software in local environment
' Return String Array with format: CAPTION; VERSION;INSTALL_DATE;INSTALL_LOCATION
'======================================================================================================================================================
'==========================================================================================================
' Creae all software list
'==========================================================================================================
Public Function GetSoftwareList(Optional ProcInBgrd As Boolean = True) As SOFTWARE()
Dim oWMI As Object, colSoftware As Object, oSoftware As Object
Dim ALLSOFT() As SOFTWARE, nDim As Integer, sDate As String
Dim TEMPSOFT As SOFTWARE, I As Integer, J As Integer, bError As Boolean

Const THIS_COMPUTER As String = "."

On Error GoTo ErrHandle
'-------------------------------------------------
nDim = -1: ReDim ALLSOFT(0)


    Set oWMI = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\" & THIS_COMPUTER & "\root\cimv2")
    Set colSoftware = oWMI.ExecQuery _
        ("Select * from Win32_Product")
If ProcInBgrd Then
    If MsgBox("THIS OPERATION MIGHT TAKE SEVERAL MINUTES!" & vbCrLf & _
          "Should we proceed it?", _
          vbYesNoCancel + vbInformation, "Software List") <> vbYes Then GoTo ExitHere
End If

DoCmd.Hourglass True  ' LONG TIME PROCESS

For Each oSoftware In colSoftware
  sDate = ""
  If Nz(oSoftware.Caption, "") <> "" Then
     nDim = nDim + 1: ReDim Preserve ALLSOFT(nDim)
  
    ALLSOFT(nDim).Caption = oSoftware.Caption
    ALLSOFT(nDim).Description = Nz(oSoftware.Description, "")
    
    sDate = Nz(oSoftware.InstallDate, "")
    If sDate = "" Then sDate = Nz(oSoftware.InstallDate2, "")
    If sDate <> "" Then
       ALLSOFT(nDim).InstallDate = DateSerial(CInt(Left(sDate, 4)), CInt(Mid(sDate, 5, 2)), Right(sDate, 2))
    End If
    
    ALLSOFT(nDim).IdentifyingNumber = Nz(oSoftware.IdentifyingNumber, "")
    ALLSOFT(nDim).InstallLocation = Nz(oSoftware.InstallLocation, "")
    ALLSOFT(nDim).InstallState = Nz(oSoftware.InstallState, "")
    ALLSOFT(nDim).Name = Nz(oSoftware.Name, "")
    ALLSOFT(nDim).PackageCache = Nz(oSoftware.PackageCache, "")
    ALLSOFT(nDim).SKUNumber = Nz(oSoftware.SKUNumber, "")
    ALLSOFT(nDim).Vendor = Nz(oSoftware.Vendor, "")
    ALLSOFT(nDim).Version = Nz(oSoftware.Version, "")
 End If
Next
'-----------------------------------------------------------------------
'make bubble sort
    For I = 0 To nDim - 1
        For J = I + 1 To nDim
            If LCase(ALLSOFT(I).Caption) > LCase(ALLSOFT(J).Caption) Then
               TEMPSOFT = ALLSOFT(J)
               ALLSOFT(J) = ALLSOFT(I)
               ALLSOFT(I) = TEMPSOFT
            End If
        Next J
    Next I
'-----------------------------------------------------------------------
ExitHere:
  If ProcInBgrd Then
       MsgBox "The Function Work is complete " & IIf(bError, " with some errors unfortunatly", " with success"), vbOK, "Software List"
  End If
  
  DoCmd.Hourglass False
  GetSoftwareList = ALLSOFT '!!!!!!!
  Set oSoftware = Nothing: Set colSoftware = Nothing: Set oWMI = Nothing
  Exit Function
'---------------
ErrHandle:
  ErrPrint "GetSoftwareList", Err.Number, Err.Description
  bError = True: Err.Clear: Resume ExitHere
End Function
'======================================================================================================================================================
' Return List of HardDrives. Require the HardDrive Structure (see in #_FILE)
'======================================================================================================================================================
Public Function DriveInfo(Optional SpecificDriveLetter As String) As HardDrive()
Dim Drv As Object, Letter As String, Total As Variant, Free As Variant, FreePercent As Variant, TotalPercent As Variant
Dim fs As Object, HDS() As HardDrive, nDim As Integer

On Error GoTo ErrHandle
'-----------------------------------------
nDim = -1: ReDim HDS(0)
Set fs = CreateObject("Scripting.FileSystemObject")

    For Each Drv In fs.drives
        If (SpecificDriveLetter <> "") And _
             SpecificDriveLetter <> Drv.DriveLetter Then GoTo NextDRV
        
        nDim = nDim + 1: ReDim Preserve HDS(nDim)
        HDS(nDim).DriveType = Drv.DriveType
        HDS(nDim).isReady = Drv.isReady
        If HDS(nDim).isReady Then
          HDS(nDim).DriveLetter = Drv.DriveLetter
          HDS(nDim).SerialNumber = Hex$(Drv.SerialNumber)
          HDS(nDim).VolumeName = Drv.VolumeName
          HDS(nDim).ShareName = Drv.ShareName
        
          HDS(nDim).Path = Drv.Path
             
          HDS(nDim).TotalSize = Drv.TotalSize
          HDS(nDim).FreeSpace = Drv.FreeSpace
          
          If (SpecificDriveLetter <> "") And SpecificDriveLetter = HDS(nDim).DriveLetter Then Exit For
        End If
NextDRV:
    Next Drv
'----------------------------------------
ExitHere:
    DriveInfo = HDS '!!!!!!!!!!!!!
    Set fs = Nothing
    Exit Function
'-----------------
ErrHandle:
    ErrPrint "DriveInfo", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'======================================================================================================================================================
' Return List of Installed Fonts
'======================================================================================================================================================
Public Function EnumFonts() As String()
Dim LF As LOGFONT
 
On Error GoTo ErrHandle
'------------------------------------------
   FntInc = 0: Erase FontArray: ReDim FontArray(0)
   
   hdc = GetDC(0)
   EnumFontFamiliesEx hdc, LF, AddressOf EnumFontFamProc, ByVal 0&, 0
   ReleaseDC 0, hdc
   '----------------------------------------
   'Sort the FontArray string array.
   Call Quicksort(FontArray(), 0, UBound(FontArray))
'-------------------------------------------
ExitHere:
   EnumFonts = FontArray '!!!!!!!!!!!!!!!!!!!!
   Exit Function
'---------------
ErrHandle:
   ErrPrint "EnumFonts", Err.Number, Err.Description
   Err.Clear
End Function
Private Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, ByVal FontType As Long, lParam As Long) As Long
   Dim FaceName As String
   
   FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
  
   ReDim Preserve FontArray(FntInc)
   FontArray(FntInc) = Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
  
   EnumFontFamProc = 1
   
   FntInc = UBound(FontArray) + 1
End Function
'======================================================================================================================================================
' Import external reg file (i.e. sRegFile = "C:\WINDOWS\DESKTOP\ENTRY.REG")
'======================================================================================================================================================
Public Function ImportRegFile(sRegFile As String) As Boolean
Dim bRes As Boolean, oShell As Object, oFile As Object, iReturn As Long
    
    On Error GoTo ErrHandle
'-------------------------------
If sRegFile = "" Then Exit Function

Set oFile = CreateObject("scripting.FileSystemObject")
If Not oFile.FileExists(sRegFile) Then Err.Raise 10006, , "Can't find the reg file " & FileNameOnly(sRegFile)

Set oShell = CreateObject("wscript.shell")
           
        iReturn = oShell.Run("regedit.exe /s" & sRegFile)
        bRes = True
'-------------------------------
ExitHere:
        ImportRegFile = bRes '!!!!!!!!!!!
        Set oFile = Nothing: Set oShell = Nothing
        Exit Function
'-----------------
ErrHandle:
        ErrPrint2 "ImportRegFile", Err.Number, Err.Description, MOD_NAME
        Err.Clear
End Function
'======================================================================================================================================================
' Getting OS Name
'======================================================================================================================================================
Public Function GetOSName() As String
Dim oWMI As Object, oITEMS As Object, oItem As Object
Dim sRes As String, SQL As String

Const CURRENT_COMPUTER As String = "."

    On Error GoTo ErrHandle
'-------------------------------
    SQL = "SELECT * FROM Win32_OperatingSystem"
    
    
    Set oWMI = GetObject("winmgmts:\\" & CURRENT_COMPUTER & "\root\cimv2")
    Set oITEMS = oWMI.ExecQuery(SQL, , 48)

    For Each oItem In oITEMS
        sRes = oItem.Name
    Next
'-------------------------------
ExitHere:
    GetOSName = sRes '!!!!!!!!!!!!
    Set oItem = Nothing: Set oITEMS = Nothing: Set oWMI = Nothing
    Exit Function
'-----------
ErrHandle:
    ErrPrint2 "GetOSName", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Resume ExitHere
End Function
'======================================================================================================================================================
' Get All Environment List
'======================================================================================================================================================
Public Function GetEnvironmentList(Optional DLM As String = ";") As String
Dim sRes As String, I  As Integer
    On Error Resume Next
'------------------------
    I = 1
    Do
       If I > 1000 Then Exit Do
       If Environ(I) = "" Then Exit Do
       sRes = sRes & DLM & Environ(I)
       I = I + 1
    Loop
If sRes <> "" Then sRes = Right(sRes, Len(sRes) - Len(DLM))
'------------------------
ExitHere:
    GetEnvironmentList = sRes '!!!!!!!!!!!
End Function
'======================================================================================================================================================
' Set Environment Var for Windows; sCope = User|System
'======================================================================================================================================================
Public Sub SetEnvironWin(sVarName As String, sVarVal As String, Optional sScope As String = "User")
    Dim objShell As Object, ScopeVars As Object
   On Error GoTo ErrHandle
'------------------
    If sVarName = "" Then Exit Sub
    
    Set objShell = CreateObject("WScript.Shell")
    Set ScopeVars = objShell.Environment(sScope)
    
    ScopeVars(sVarName) = sVarVal
'------------------
ExitHere:
    Set ScopeVars = Nothing: Set objShell = Nothing
    Exit Sub
'------------
ErrHandle:
    ErrPrint2 "SetEnvironWin", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Sub

'======================================================================================================================================================
' Get Environment Var From Windows
'======================================================================================================================================================
Public Function GetEnviron(sVarName As String) As String
Dim sRes As String
   On Error GoTo ErrHandle
'------------------
    sRes = String(255, vbNullChar)
    If sVarName = "" Then Exit Function
    Call GetEnvironmentVariable(sVarName, sRes, Len(sRes))
'------------------
ExitHere:
    GetEnviron = TrimNull(sRes) '!!!!!!!!!!!
    Exit Function
'------------
ErrHandle:
    ErrPrint2 "GetEnviron", Err.Number, Err.Description, MOD_NAME
    Err.Clear

End Function
'----------------------------------------------------------------------------------------------------------------------------------------------------
' Trim API String
'----------------------------------------------------------------------------------------------------------------------------------------------------
Private Function TrimNull(Item As String)
Dim iPos As Long
    iPos = InStr(Item, vbNullChar)
    TrimNull = IIf(iPos > 0, Left$(Item, iPos - 1), Item)
End Function
'======================================================================================================================================================
' Create Environment Param (!!! The function Environ read all variables only with restarting Access)
'======================================================================================================================================================
Public Sub SetEnviron(sVarName As String, sVarVal As String)
   On Error GoTo ErrHandle
'------------------
    If sVarName = "" Then Exit Sub
    SetEnvironmentVariable sVarName, sVarVal
'------------------
ExitHere:
    Exit Sub
'------------
ErrHandle:
    ErrPrint2 "SetEnviron", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Sub

'======================================================================================================================================================
' Get Information About CURRENT Screen (specific value depending of Item): horizontalresolution, verticalresolution, widthinches, heightinches
'                                                   diagonalinches, pixelsperinchx, pixelsperinchy, pixelsperinch, ppidiag, windotsperinchx, dpix
'                                                   windotsperinchy, dpiy, windotsperinch, dpiwin, adjustmentfactor, zoomfac, isprimary, update
'                                                   help
' @J. Woolley, https://wellsr.com/vba/2019/excel/calculate-screen-size-and-other-display-details-with-vba/
'======================================================================================================================================================
Public Function ScreenInfo(Item As String) As Variant
#If Win64 Then
        Dim hMonitor As LongPtr
#Else
        Dim hMonitor As Long
#End If
Dim xHSizeSq As Double, xVSizeSq As Double, xPix As Double, xDot As Double
Dim tMonitorInfo As MONITORINFOEX, nMonitors As Integer, vResult As Variant, sItem As String

    On Error GoTo ErrHandle
'----------------------------
'If Application.Name = "Microsoft Excel" Then Application.Volatile

nMonitors = GetSystemMetrics(SM_CMONITORS)
    If nMonitors < 2 Then
        nMonitors = 1                                       ' in case GetSystemMetrics failed
        hWnd = 0
    Else
        hWnd = GetActiveWindow()
        hMonitor = MonitorFromWindow(hWnd, MONITOR_DEFAULTTONULL)
        If hMonitor = 0 Then
            Debug.Print "ActiveWindow does not intersect a monitor"
            hWnd = 0
        Else
            tMonitorInfo.cbSize = Len(tMonitorInfo)
            If GetMonitorInfo(hMonitor, tMonitorInfo) = False Then
                Debug.Print "GetMonitorInfo failed"
                hWnd = 0
            Else
                hdc = CreateDC(tMonitorInfo.szDevice, 0, 0, 0)
                If hdc = 0 Then
                    Debug.Print "CreateDC failed"
                    hWnd = 0
                End If
            End If
        End If
    End If
    
    If hWnd = 0 Then
        hdc = GetDC(hWnd)
        tMonitorInfo.dwFlags = MONITOR_PRIMARY
        tMonitorInfo.szDevice = "PRIMARY" & vbNullChar
    End If
    
    sItem = Trim(LCase(Item))
    Select Case sItem
    Case "horizontalresolution", "pixelsx"                  ' HorizontalResolution (pixelsX)
        vResult = GetDeviceCaps(hdc, DevCap.HORZRES)
    Case "verticalresolution", "pixelsy"                    ' VerticalResolution (pixelsY)
        vResult = GetDeviceCaps(hdc, DevCap.VERTRES)
    Case "widthinches", "inchesx"                           ' WidthInches (inchesX)
        vResult = GetDeviceCaps(hdc, DevCap.HORZSIZE) / 25.4
    Case "heightinches", "inchesy"                          ' HeightInches (inchesY)
        vResult = GetDeviceCaps(hdc, DevCap.VERTSIZE) / 25.4
    Case "diagonalinches", "inchesdiag"                     ' DiagonalInches (inchesDiag)
        vResult = Sqr(GetDeviceCaps(hdc, DevCap.HORZSIZE) ^ 2 + GetDeviceCaps(hdc, DevCap.VERTSIZE) ^ 2) / 25.4
    Case "pixelsperinchx", "ppix"                           ' PixelsPerInchX (ppiX)
        vResult = 25.4 * GetDeviceCaps(hdc, DevCap.HORZRES) / GetDeviceCaps(hdc, DevCap.HORZSIZE)
    Case "pixelsperinchy", "ppiy"                           ' PixelsPerInchY (ppiY)
        vResult = 25.4 * GetDeviceCaps(hdc, DevCap.VERTRES) / GetDeviceCaps(hdc, DevCap.VERTSIZE)
    Case "pixelsperinch", "ppidiag"                         ' PixelsPerInch (ppiDiag)
        xHSizeSq = GetDeviceCaps(hdc, DevCap.HORZSIZE) ^ 2
        xVSizeSq = GetDeviceCaps(hdc, DevCap.VERTSIZE) ^ 2
        xPix = GetDeviceCaps(hdc, DevCap.HORZRES) ^ 2 + GetDeviceCaps(hdc, DevCap.VERTRES) ^ 2
        vResult = 25.4 * Sqr(xPix / (xHSizeSq + xVSizeSq))
    Case "windotsperinchx", "dpix"                          ' WinDotsPerInchX (dpiX)
        vResult = GetDeviceCaps(hdc, DevCap.LOGPIXELSX)
    Case "windotsperinchy", "dpiy"                          ' WinDotsPerInchY (dpiY)
        vResult = GetDeviceCaps(hdc, DevCap.LOGPIXELSY)
    Case "windotsperinch", "dpiwin"                         ' WinDotsPerInch (dpiWin)
        xHSizeSq = GetDeviceCaps(hdc, DevCap.HORZSIZE) ^ 2
        xVSizeSq = GetDeviceCaps(hdc, DevCap.VERTSIZE) ^ 2
        xDot = GetDeviceCaps(hdc, DevCap.LOGPIXELSX) ^ 2 * xHSizeSq + GetDeviceCaps(hdc, DevCap.LOGPIXELSY) ^ 2 * xVSizeSq
        vResult = Sqr(xDot / (xHSizeSq + xVSizeSq))
    Case "adjustmentfactor", "zoomfac"                      ' AdjustmentFactor (zoomFac)
        xHSizeSq = GetDeviceCaps(hdc, DevCap.HORZSIZE) ^ 2
        xVSizeSq = GetDeviceCaps(hdc, DevCap.VERTSIZE) ^ 2
        xPix = GetDeviceCaps(hdc, DevCap.HORZRES) ^ 2 + GetDeviceCaps(hdc, DevCap.VERTRES) ^ 2
        xDot = GetDeviceCaps(hdc, DevCap.LOGPIXELSX) ^ 2 * xHSizeSq + GetDeviceCaps(hdc, DevCap.LOGPIXELSY) ^ 2 * xVSizeSq
        vResult = 25.4 * Sqr(xPix / xDot)
    Case "isprimary"                                        ' IsPrimary
        vResult = CBool(tMonitorInfo.dwFlags And MONITOR_PRIMARY)
    Case "displayname"                                      ' DisplayName
        vResult = tMonitorInfo.szDevice & vbNullChar
        vResult = Left(vResult, (InStr(1, vResult, vbNullChar) - 1))
    Case "update"                                           ' Update
        vResult = Now
    Case "help"                                             ' Help
        vResult = "HorizontalResolution (pixelsX), VerticalResolution (pixelsY), " _
            & "WidthInches (inchesX), HeightInches (inchesY), DiagonalInches (inchesDiag), " _
            & "PixelsPerInchX (ppiX), PixelsPerInchY (ppiY), PixelsPerInch (ppiDiag), " _
            & "WinDotsPerInchX (dpiX), WinDotsPerInchY (dpiY), WinDotsPerInch (dpiWin), " _
            & "AdjustmentFactor (zoomFac), IsPrimary, DisplayName, Update, Help"
    Case Else                                               ' Else
        vResult = CVErr(2015)                               ' return #VALUE! error (2015)
    End Select
    
    If hWnd = 0 Then
        ReleaseDC hWnd, hdc
    Else
        DeleteDC hdc
    End If
'--------------------
ExitHere:
    ScreenInfo = vResult '!!!!!!!!!
    Exit Function
'---------
ErrHandle:
    ErrPrint2 "ScreenIndo", Err.Number, Err.Description, MOD_NAME
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
                                                                                                Optional sModName As String = "#_ENVIRONMENT") As String
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

'=====================================================================================================================================================
' Create Thumbnail Fore MediaFile
'=====================================================================================================================================================
Public Function CreateThumbnail(sPathSrc As String, Optional sThumbFolder As String, Optional sPrefix As String = "th_", _
                                                                           Optional IWidth As Long = 200, Optional iHeight As Long = 100, _
                                                                           Optional iWait As Long = 0, Optional bFromClipBoard As Boolean) As String
Dim MyImage As cImage, sExt As String, sRes As String, sWork As String

On Error GoTo ErrHandle
'---------------------------------------------------------
If bFromClipBoard Then
     sWork = IIf(sThumbFolder <> "", sThumbFolder, CurrentProject.Path & "\" & "THUMB")
     If Dir(sWork, vbDirectory) = "" Then MkDir (sWork)
     sWork = sWork & "\" & GenRandomStr(, False) & Int((91) * Rnd + 10) & ".png"
     If Dir(sWork) <> "" Then Kill sWork
     sRes = ImageFromClipboard(sWork, "/resize=(" & IWidth & "," & iHeight & ")")
     GoTo ExitHere
End If
If sPathSrc = "" Then Exit Function
sExt = FileExt(sPathSrc)
Select Case UCase(sExt):
    Case "DOC", "DOCX":
        Exit Function
    Case "PPT", "PPTX":
        Exit Function
    Case "PDF":
        sRes = PDFThumbnail(sPathSrc, sThumbFolder, sPrefix, IWidth, iHeight, iWait)
        GoTo ExitHere
    Case "JPEG", "JPG", "PNG", "BMP", "GIF":
        Set MyImage = New cImage
        sRes = MyImage.ThumbnailFile(sPathSrc, sThumbFolder, sPrefix, IWidth, iHeight, iWait)
    Case "MP4", "MPEG", "MPG", "AVI", "FLV", "M4V":
        Set MyImage = New cImage
        sRes = MyImage.ThumbnailFile(sPathSrc, sThumbFolder, sPrefix, IWidth, iHeight, iWait)
    Case Else

End Select
    If iWait > 0 Then Wait iWait

'---------------------------------
ExitHere:
    Set MyImage = Nothing
    CreateThumbnail = sRes '!!!!!!!!!!!!!
    Exit Function
'----------------
ErrHandle:
    ErrPrint "CreateThumbnail", Err.Number, Err.Description
    Err.Clear: Set MyImage = Nothing
End Function


'====================================================================================================================================================
' PDF To Thumbnail
'====================================================================================================================================================
Public Function PDFThumbnail(sSourceFile As String, Optional sFolder As String, Optional sPrefix As String = "th_", _
                                               Optional IWidth As Long = 100, Optional iHeight As Long = 200, Optional iWait As Long = 0) As String
Dim sOutPut As String, THMBFolder As String
Dim bRes As Boolean, sParam As String

Const GHOST_PARAM As String = "-q -sDEVICE=png16m -dNOPAUSE -dFirstPage=1 -dLastPage=1 -dPDFFitPage=true -dBATCH"


On Error GoTo ErrHandle
'---------------------------------
      If sSourceFile = "" Then Exit Function
      If Dir(sSourceFile) = "" Then Err.Raise 1000, , "Wrong File Name: " & sSourceFile
      If UCase(FileExt(sSourceFile)) <> "PDF" Then Exit Function
      
      sParam = GHOST_PARAM & " -g" & IWidth & "x" & iHeight
      THMBFolder = IIf(sFolder <> "", sFolder, CurrentProject.Path & "\THUMB")
      sOutPut = THMBFolder & "\" & sPrefix & Split(FileNameOnly(sSourceFile), ".")(0) & ".png"

      bRes = GSExecute(sParam, sSourceFile, sOutPut)
      If Not bRes Then sOutPut = ""
'---------------------------------
ExitHere:
      PDFThumbnail = sOutPut '!!!!!!!!!!!!!!!
      Exit Function
'--------------------
ErrHandle:
      ErrPrint "PDFThumbnail", Err.Number, Err.Description
      Err.Clear
End Function

'======================================================================================================================================================
' Find Installation Directory
'======================================================================================================================================================
Public Function GetInstallDirectory(ByVal usProgName As String) As String

    Dim fRetPath As String * MAX_PATH
    Dim fRetLng As Long

    fRetLng = FindExecutable(usProgName, S_DIR, fRetPath)

    If fRetLng >= MIN_SUCCESS_LNG Then
        GetInstallDirectory = Left$(Trim$(fRetPath), InStrRev(Trim$(fRetPath), "\"))
    End If

End Function



'======================================================================================================================================================================
' Search File in Folder and sub folder
'======================================================================================================================================================================
Public Function RecurseSearch(sPath As String, sFileNameOrPattern As String, Optional bFirstResult As Boolean = True, Optional DLM As String = ";") As String
Dim FSO As Object, myFolder As Object, mySubFolder As Object, myFile As Object
Dim sRes As String

On Error GoTo ErrHandle
'--------------------------------------------------------
If (sPath = "") Or (sFileNameOrPattern = "") Then Exit Function

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set myFolder = FSO.GetFolder(sPath)
    
    For Each myFile In myFolder.Files
        If FileMatch(myFile.Name, sFileNameOrPattern) Then sRes = ConcateString(sRes, myFile.Path, DLM)
        If sRes <> "" And bFirstResult Then GoTo ExitHere
        
    Next myFile

    For Each mySubFolder In myFolder.SubFolders
        For Each myFile In mySubFolder.Files
            If FileMatch(myFile.Name, sFileNameOrPattern) Then sRes = ConcateString(sRes, myFile.Path, DLM)
                    If sRes <> "" And bFirstResult Then GoTo ExitHere
        Next
        sRes = IIf(sRes <> "", sRes & DLM, "") & RecurseSearch(mySubFolder.Path, sFileNameOrPattern, bFirstResult, DLM)
        If sRes <> "" And bFirstResult Then GoTo ExitHere
    Next
'--------------------------------------
ExitHere:
    RecurseSearch = sRes '!!!!!!!!!!!!
    Set myFile = Nothing: Set myFolder = Nothing: Set FSO = Nothing
    Exit Function
'-----------------
ErrHandle:
    ErrPrint "RecurseSearch", Err.Number, Err.Description
    Err.Clear
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
' Is File match with template
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function FileMatch(sFileName As String, Optional sPattern As String = "*") As Boolean
Dim bRes As Boolean
'---------------------------------
If sFileName = "" Then Exit Function
If (sFileName Like sPattern) Then bRes = True
'---------------------------------
        FileMatch = bRes '!!!!!!!!!!!!
End Function

'====================================================================================================================================================================
' Get List Separator
'====================================================================================================================================================================
Public Function GetListSeparator() As String
Dim ListSeparator As String
Dim iRetVal1 As Long
Dim iRetVal2 As Long
Dim lpLCDataVar As String

Dim Position As Integer
Dim Locale As Long

Locale = GetUserDefaultLCID()

iRetVal1 = GetLocaleInfo(Locale, LOCALE_SLIST, lpLCDataVar, 0)

ListSeparator = String$(iRetVal1, 0)

iRetVal2 = GetLocaleInfo(Locale, LOCALE_SLIST, ListSeparator, iRetVal1)

Position = InStr(ListSeparator, Chr$(0))
If Position > 0 Then
ListSeparator = Left$(ListSeparator, Position - 1)
'----------------------------------------
    GetListSeparator = ListSeparator  '!!!!!!!!!!!!!!
End If

End Function

'======================================================================================================================================================
' Get All installed software list
'======================================================================================================================================================
Public Function ListAllSoftware(Optional bWithVersion As Boolean, Optional DLM As String = ";") As String
Dim oWMI As Object, oSoftware As Object, oListing As Object
Dim sRes As String, Arr() As String, nDim As Long

Const CURRENT_COMP As String = "."

    On Error GoTo ErrHandle
'---------------------------------
    Debug.Print "ATTENTION! Long wirking time of ListAllSoftware, wait"
    DoCmd.Hourglass True
    
    ReDim Arr(0): nDim = -1
    
    Set oWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & CURRENT_COMP & "\root\cimv2")
    Set oListing = oWMI.ExecQuery("Select * from Win32_Product")

    If oListing.Count > 0 Then
        For Each oSoftware In oListing
            With oSoftware
                sRes = Nz(.Caption, "")
                If sRes <> "" Then
                        nDim = nDim + 1: ReDim Preserve Arr(nDim)
                        Arr(nDim) = sRes: If bWithVersion Then Arr(nDim) = Arr(nDim) & "(" & .Version & ")"
                 End If
            End With
        Next
    End If
    
    If nDim > 1 Then BubbleSort Arr
    sRes = Join(Arr, DLM)
'-----------------------------
ExitHere:
    ListAllSoftware = sRes '!!!!!!!!
    DoCmd.Hourglass False
    Set oSoftware = Nothing: Set oSoftware = Nothing: Set oWMI = Nothing
    Exit Function
'-----------
ErrHandle:
    ErrPrint2 "ListAllSoftware", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Resume ExitHere
End Function

'======================================================================================================================================================
' Print Debug Information
'======================================================================================================================================================
Public Sub DebugPrint(sMsg As String, Optional bDebug As Boolean = True, Optional bErr As Boolean, Optional FuncName As String, _
                                                                    Optional ModName As String, Optional AltPrefix As String, Optional dTime As Double)
Dim sTxt As String, sTitle As String, sPrefix As String

Const MSG_PREFIX As String = "     ---> "
Const ERR_PREFIX As String = "     #### "

      On Error Resume Next
'---------------
      If Not bDebug Then Exit Sub                   ' Exit when no Debug Mode
      If sMsg = vbNullString Then Exit Sub          ' Can't message without msg body
      
      If IsBlank(AltPrefix) Then
            sPrefix = IIf(bErr, ERR_PREFIX, MSG_PREFIX)
      Else
            sPrefix = AltPrefix
      End If
      
      sTitle = sPrefix & IIf(Not IsBlank(ModName), ModName & ".", "") & IIf(Not IsBlank(FuncName), FuncName & ": ", "")
      sTxt = sTitle & " " & sMsg & IIf(dTime > 0, " |Time: " & dTime, "")
'---------------
ExitHere:
      Debug.Print sTxt
End Sub
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

