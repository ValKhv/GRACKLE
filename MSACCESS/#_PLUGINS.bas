Attribute VB_Name = "#_PLUGINS"
'*********************************.ze$$e. ************************************************************************************************************
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
'                 d$$$$$$$$$F                $P"                   ##  ##  ##  ##   #######   #####  ##    ##  ##  ####  #### ###   ##  #####
'                 $$$$$$$$$$F                                       ## ##  ##  ##  ##    ##   ##  ## ##    ##  ## ##  ##  ##  ## #  ## ##
'                  *$$$$$$$$"                                        ####  #####  ##     ##   ##  ## ##    ##  ## ##      ##  ## ## ## ###
'                    "***""               _____________                                       #####  ##    ##  ## ## ###  ##  ##  # ##   ##
' STANDARD MODULE WITH DEFAULT FUNCTIONS |v 2021/08/20 |                                      ##     ##    ##  ## ##  ##  ##  ##   ###     ##
' The module manages the connection of external libraries and is part of the G-VBA library    ##     #####  ####   ####  #### ##    ## #####
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
Option Explicit


Private Const MOD_NAME As String = "#_PLUGINS"
'************************
'====================================================================================================================================================
' Get The Path to External Library
'====================================================================================================================================================
Public Function GetExecutor(sExecName As String) As String
Dim sRes As String, sMsg As String, sLink As String

On Error GoTo ErrHandle
'---------------------------
Select Case UCase(sExecName)
Case "IRFANVIEW":
    sRes = GetIrfanViewPath()
Case "IMAGEMAGICK":
    sRes = GetImageMagick()
Case "GHOSTSCRIPT":
    sRes = GetGhostScript()
Case "OPENSSL":
    sRes = GetOpenSSL()
Case Else
End Select
'---------------------------
ExitHere:
    GetExecutor = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'-------------
ErrHandle:
    ErrPrint2 "GetExecutor", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'======================================================================================================================================================
' This function import form on the fly and start it
'======================================================================================================================================================
Public Function GFORM(sFormName As String) As String
Dim sPath As String, sRes As String

Const GFORM_PARAM As String = "GFORMS"

    On Error GoTo ErrHandle
'-------------------------------------
If sFormName = "" Then Exit Function
If IsForm(sFormName) Then
    sRes = sFormName
    GoTo ExitHere
End If
'------------
    sPath = GetLocal(GFORM_PARAM)
    If sPath = "" Then
            sPath = FolderNameOnlyAbstract(GetGracklePath())
            If sPath <> "" Then sPath = sPath & GFORM_PARAM & ".accdb"
    End If
    If sPath = "" Then sPath = OpenDialog(GC_FILE_PICKER, "GFORMS Library", "Access Database,*.accdb", False, GetLastFolder())
    If sPath = "" Then Exit Function
    If Dir(sPath) = "" Then Err.Raise 10008, , "Can't find the GFORMS Library"

    If ImportForm(sFormName, sPath) Then sRes = sFormName
'-------------------------------------
ExitHere:
    GFORM = sRes '!!!!!!!!!!!!!!!!!!!
    Exit Function
'--------------
ErrHandle:
    ErrPrint2 "GFORM", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function


'====================================================================================================================================================
' Run GhostScript to command
'====================================================================================================================================================
Public Function GSExecute(Optional sParam As String, Optional sSrcFile As String, Optional sOutPutFile As String) As Boolean
Dim sCommand As String, sEXE As String, sOutPut As String

On Error GoTo ErrHandle
'-------------------------------------
If sParam = "" Then Err.Raise 1000, , " No any command for GhostScript"
sEXE = GetExecutor("GHOSTSCRIPT"): If sEXE = "" Then Err.Raise 1000, , "No any GhostScript instance on this Workstation, please install"
sOutPut = IIf(sOutPutFile = "", sSrcFile, sOutPutFile)

sCommand = SH(sEXE) & " " & sParam & " " & "-sOutputFile=" & sOutPut & " " & sSrcFile
'----------------------------------------
ExitHere:
     GSExecute = SyncShell(sCommand) '!!!!!!!!!!!
     Exit Function
'-----------
ErrHandle:
     ErrPrint2 "GSExecute", Err.Number, Err.Description, MOD_NAME
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
' Get BinPath for ImageMagick (image processor, is needed fo some specific operation)
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetImageMagick() As String
Dim sRes As String, iRes As Long

Const DLM As String = ";"
Const LOCALPARAM As String = "IMAGEMAGICK"
Const EXEFILES As String = "magick.exe"                                            ' Add some new exec name to proper work
Const EXEFOLDERS As String = "ImageMagick-7.0.8-Q16;ImageMagick-7.1.0-Q16-HDRI"    ' Add some new folder to proper work

    sRes = FindProg(EXEFILES, EXEFOLDERS, LOCALPARAM, DLM)
    
    
    If sRes = "" Then
                   iRes = MsgBox("Unable to find ImageMagick executable file automatically." & vbCrLf & _
                   "Do you want to specify it manually (YES) or download the installation (NO), exit - CANCEL?", vbYesNoCancel + vbQuestion, "GetImageMagick")
                   
                   Select Case iRes
                   Case vbYes:
                             sRes = OpenDialog(GC_FILE_PICKER, "Pick the magick.exe", , , CurrentProject.Path)
                   Case vbNo:
                             Application.FollowHyperlink "https://imagemagick.org"
                   Case vbCancel:
                            Exit Function
                   End Select
                   
                   If sRes <> "" Then Call SetLocal(LOCALPARAM, sRes, "The Path to imagemagick server")
    End If
'----------------------------------
ExitHere:
    GetImageMagick = sRes '!!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' the function calculates the path to the given program for special extensions (external libraries)
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function FindProg(sExecNames As String, sPossibleFolders As String, Optional sLocalParam As String, Optional DLM As String = ";") As String
Dim sPath As String, sProgramFiles As String, sProgramFiles86 As String, sWork As String
Dim FOLDERS() As String, nFolders As Integer, sEXECs() As String, nEXECS As Integer, I As Integer, J As Integer

    On Error GoTo ErrHandle
'---------------------
    If sLocalParam <> "" Then                 ' TRY TO GET PATH FROM LOCAL
       sPath = GetLocal(sLocalParam)
       If sPath <> "" Then
           If Dir(sPath) <> "" Then
                 GoTo ExitHere
           Else
                 sPath = "": Call SetLocal(sLocalParam, "")
           End If
       Else
                 Call SetLocal(sLocalParam, "")
       End If
    End If
    sProgramFiles = Environ("ProgramFiles"): sProgramFiles86 = Environ("PROGRAMFILES(X86)")
    FOLDERS = Split(sPossibleFolders, DLM): nFolders = UBound(FOLDERS)
    sEXECs = Split(sExecNames, DLM): nEXECS = UBound(sEXECs)
    
    For I = 0 To nFolders
        For J = 0 To nEXECS
            sPath = sProgramFiles & "\" & FOLDERS(I) & "\" & sEXECs(J)
            If Dir(sPath) <> "" Then GoTo ExitHere
            sPath = sProgramFiles86 & "\" & FOLDERS(I) & "\" & sEXECs(J)
            If Dir(sPath) <> "" Then GoTo ExitHere
        Next J
    Next I
'---------------------
ExitHere:
    If Dir(sPath) = "" Then
           Exit Function
    Else
           If sLocalParam <> "" Then Call SetLocal(sLocalParam, sPath)
    End If
    FindProg = sPath '!!!!!!!!!!!!
    Exit Function
'-------
ErrHandle:
    ErrPrint2 "FindProg", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------
' Get BinPath for OpenSSL
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetOpenSSL() As String
Dim sRes As String, iRes As Long

Const DLM As String = ";"
Const LOCALPARAM As String = "OPENSSL"
Const EXEFILES As String = "openssl.exe"                      ' Add some new exec name to proper work
Const EXEFOLDERS As String = "OpenSSL\x64\bin"                ' Add some new folder to proper work

    sRes = FindProg(EXEFILES, EXEFOLDERS, LOCALPARAM, DLM)
    
    If sRes = "" Then
                   iRes = MsgBox("Unable to find GHOSTSCRIPT executable file automatically." & vbCrLf & _
                   "Do you want to specify it manually (YES) or download the installation (NO), exit - CANCEL?", vbYesNoCancel + vbQuestion, "GetOpenSSL")
                   
                   Select Case iRes
                   Case vbYes:
                             sRes = OpenDialog(GC_FILE_PICKER, "Pick the OpenSSL", , , CurrentProject.Path)
                   Case vbNo:
                             Application.FollowHyperlink "https://www.openssl.org/"
                   Case vbCancel:
                            Exit Function
                   End Select
                   
                   If sRes <> "" Then Call SetLocal(LOCALPARAM, sRes, "The Path to OpenSSL server")
    End If
'----------------------------------
ExitHere:
    GetOpenSSL = sRes '!!!!!!!!!!!!!
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------
' Get BinPath for GhostScript
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetGhostScript() As String
Dim sRes As String, iRes As Long

Const DLM As String = ";"
Const LOCALPARAM As String = "GHOSTSCRIPT"
Const EXEFILES As String = "gswin64c.exe"          ' Add some new exec name to proper work
Const EXEFOLDERS As String = "gs\gs9.26\bin;gs\gs9.54.0\bin"                        ' Add some new folder to proper work

    sRes = FindProg(EXEFILES, EXEFOLDERS, LOCALPARAM, DLM)
    
    If sRes = "" Then
                   iRes = MsgBox("Unable to find GHOSTSCRIPT executable file automatically." & vbCrLf & _
                   "Do you want to specify it manually (YES) or download the installation (NO), exit - CANCEL?", vbYesNoCancel + vbQuestion, "GetGhostScript")
                   
                   Select Case iRes
                   Case vbYes:
                             sRes = OpenDialog(GC_FILE_PICKER, "Pick the ghostscript", , , CurrentProject.Path)
                   Case vbNo:
                             Application.FollowHyperlink "https://www.ghostscript.com/download.html"
                   Case vbCancel:
                            Exit Function
                   End Select
                   
                   If sRes <> "" Then Call SetLocal(LOCALPARAM, sRes, "The Path to ghostscript server")
    End If
'----------------------------------
ExitHere:
    GetGhostScript = sRes '!!!!!!!!!!!!!
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------
' Get BinPath To IrfanView
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetIrfanViewPath() As String
Dim oWScript As Object, sPath As String, iRes As Integer

Const DLM As String = ";"
Const LOCALPARAM As String = "IrfanView"
Const EXEFILES As String = "i_view64.exe;i_view32.exe"          ' Add some new exec name to proper work
Const EXEFOLDERS As String = "IrfanView"                        ' Add some new folder to proper work

Const REG_PATH As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\IrfanView\shell\open\command\"


On Error GoTo ErrHandle
'-----------------------------------------------
  sPath = FindProg(EXEFILES, EXEFOLDERS, LOCALPARAM, DLM)
  If sPath = "" Then
        Set oWScript = CreateObject("WScript.Shell")  ' TRY TO GET IRFANVIEW FROM REGISTRY
        sPath = oWScript.RegRead(REG_PATH)
        If sPath <> "" Then
                sPath = Replace(sPath, "%1", ""): sPath = Trim(Replace(sPath, Chr(34), ""))
                SetLocal "IrfanView", sPath
                GoTo ExitHere
        End If
  End If
  
  If sPath = "" Then                                 ' TRY TO GET IRFANVIEW PATH MNUALLY
                   iRes = MsgBox("Unable to find IrfanView executable file automatically." & vbCrLf & _
                   "Do you want to specify it manually (YES) or download the installation (NO), exit - CANCEL?", vbYesNoCancel + vbQuestion, "GetIrfanViewPat")
                   
                   Select Case iRes
                   Case vbYes:
                             sPath = OpenDialog(GC_FILE_PICKER, "Pick the irfanview", , , CurrentProject.Path)
                   Case vbNo:
                             Application.FollowHyperlink "https://www.irfanview.com/main_download_engl.htm"
                   Case vbCancel:
                            Exit Function
                   End Select
                   
                   If sPath <> "" Then Call SetLocal(LOCALPARAM, sPath, "The Path to irfanview")
  End If
  
'------------------------------------------------
ExitHere:
  GetIrfanViewPath = sPath '!!!!!!!!!!!!!!!
  Set oWScript = Nothing
  Exit Function
'---------------
ErrHandle:
  ErrPrint2 "GetIrfanViewPath", Err.Number, Err.Description, MOD_NAME
  Err.Clear: Resume ExitHere
End Function




