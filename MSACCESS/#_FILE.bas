Attribute VB_Name = "#_FILE"
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
'                 $$$$$$$$$$F                                       ## ##  ##  ##  ##    ##   ###### ####  ##    ######
'                  *$$$$$$$$"                                        ####  #####  ##     ##   ##      ##   ##    ##
'                    "***""               _____________                                       ####    ##   ##    ####
' STANDARD MODULE WITH FILE FUNCTIONS    |v 2017/03/19 |                                      ##      ##   ##    ##
' The module contains frequently used  files functions and is part of the G-VBA library       ##     ####  ##### ######
'****************************************************************************************************************************************************
'****************************************************************************************************************************************************
Option Explicit

Private Const MOD_NAME As String = "#_FILE"
'*********************************************************************************************
Public Type FileInfo
    fileName As String
    FilePath As String
    FileExt As String
    ParentFolder As String
    FullPath As String
    FileSize As Long
    FileHash As String
    FileProps As String
    FileDigest As String
    FileDate As Date
    FileCRC32 As String
End Type

Public Enum DrvType
    UNKNOW = 0
    REMOVABLE = 1
    FIXED = 2
    NETWORK = 3
    CDROM = 4
    RAMDRIVE = 5
End Enum

Public Type HardDrive
    DriveLetter As String
    DriveType As DrvType
    FileSystem As String
    FreeSpace As Variant
    TotalSize As Variant
    SerialNumber As String
    Path As String
    VolumeName As String
    AvailableSize As Variant
    isReady As Boolean
    ShareName As String
End Type

Public Type G_Doc
    Title As String
    Authors As String
    KeyWords As String
    Pages As Integer
    Subject As String
End Type
'*********************************************************************************************
Private Type MD5_CTX                  ' Structure to generate MD5 HASH
  I(1 To 2) As Long
  buf(1 To 4) As Long
  inp(1 To 64) As Byte
  digest(1 To 16) As Byte
End Type


Private Enum FCOPY
    FOF_CREATEPROGRESSDLG = &H0&
    FOF_NOCREATEPROGRESSDLG = &H4&
    FOF_CREATENEWFLDERIFEXISTS = &H8&
    FOF_COPYALLNOPROMPT = &H10&
    FOF_ALLOWUNDO = &H40&
    FOF_CREATEPROGRESSDLGNOTEXT = &H100&
    FOF_FILESONLY = &H80                  '  on *.*, do only files
    FOF_SIMPLEPROGRESS = &H100            '  means don't show names of files
    FOF_NOCONFIRMMKDIR = &H200            '  don't confirm making any needed dirs
End Enum

Private Declare PtrSafe Function PathIsRelative Lib "Shlwapi" _
        Alias "PathIsRelativeA" (ByVal Path As String) As Long

    Public Enum EMakeDirStatus
        ErrSuccess = 0
        ErrRelativePath
        ErrInvalidPathSpecification
        ErrDirectoryCreateError
        ErrSpecIsFileName
        ErrInvalidCharactersInPath
    End Enum
    Const MAX_PATH = 260
'*********************************************************************************************
#If VBA7 Then
    Type SHFILEOPSTRUCT
        hWnd As LongPtr
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Integer
        fAborted As Boolean
        hNameMaps As LongPtr
        sProgress As String
    End Type
    
    Private Declare PtrSafe Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
                                     (lpFileOp As SHFILEOPSTRUCT) As LongPtr
    Private Declare PtrSafe Function GetTempPath Lib "kernel32" _
                             Alias "GetTempPathA" (ByVal nBufferLength As LongPtr, _
                                                   ByVal lpBuffer As String) As Long

    Private Declare PtrSafe Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, _
                 ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

    Private Declare PtrSafe Sub MD5Init Lib "cryptdll" (Context As MD5_CTX)
    Private Declare PtrSafe Sub MD5Update Lib "cryptdll" (Context As MD5_CTX, ByVal strInput As String, ByVal lLen As Long)
    Private Declare PtrSafe Sub MD5Final Lib "cryptdll" (Context As MD5_CTX)

#Else
    Type SHFILEOPSTRUCT
        hWnd As Long
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Integer
        fAborted As Boolean
        hNameMaps As Long
        sProgress As String
    End Type
    
    Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
                                     (lpFileOp As SHFILEOPSTRUCT) As Long
    Private Declare Function GetTempPath Lib "kernel32" _
                             Alias "GetTempPathA" (ByVal nBufferLength As LongPtr, _
                                                   ByVal lpBuffer As String) As Long

    Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, _
                 ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

    Private Declare Sub MD5Init Lib "cryptdll" (Context As MD5_CTX)
    Private Declare Sub MD5Update Lib "cryptdll" (Context As MD5_CTX, ByVal strInput As String, ByVal lLen As Long)
    Private Declare Sub MD5Final Lib "cryptdll" (Context As MD5_CTX)

#End If

Private Const FO_MOVE = &H1
Private Const FO_COPY = &H2
Private Const FO_DELETE = &H3
Private Const FO_RENAME = &H4
Private Const FOF_MULTIDESTFILES = &H1
Private Const FOF_CONFIRMMOUSE = &H2
Private Const FOF_SILENT = &H4                      '  don't create progress/report
Private Const FOF_RENAMEONCOLLISION = &H8
Private Const FOF_NOCONFIRMATION = &H10             '  Don't prompt the user.
Private Const FOF_WANTMAPPINGHANDLE = &H20          '  Fill in SHFILEOPSTRUCT.hNameMappings
                                      '  Must be freed using SHFreeNameMappings

'*****************************************
'======================================================================================================================================================
' Get short path for filename
'======================================================================================================================================================
Public Function GetShortPath(ByVal sPath As String) As String
Dim sRes As String, FSO As Object

    On Error Resume Next
'--------------------
    If sPath = "" Then Exit Function
    If Dir(sPath) = "" Then Exit Function
    Set FSO = CreateObject("Scripting.FileSystemObject")
    sRes = FSO.GetFile(sPath).ShortPath
'--------------------
ExitHere:
    GetShortPath = sRes '!!!!!!!!!!!
    Set FSO = Nothing
End Function
'======================================================================================================================================================
' Copy whole folder and wait
'======================================================================================================================================================
Public Function CopyFolderAndWait(sCopyFromFolder As String, sDestinationFolder As String) As Boolean
Dim objShell As Object, objFolder As Object

    On Error GoTo ErrHandle
'------------------------------
If sCopyFromFolder = "" Or sDestinationFolder = "" Then Exit Function
If Dir(sCopyFromFolder, vbDirectory) = "" Then Err.Raise 10007, , "Can't find the source folder " & sCopyFromFolder
If Dir(sDestinationFolder, vbDirectory) = "" Then Err.Raise 10007, , "Can't find the destination folder " & sDestinationFolder

Set objShell = CreateObject("Shell.Application"): Set objFolder = objShell.NameSpace(sDestinationFolder)
objFolder.CopyHere sCopyFromFolder, FCOPY.FOF_CREATEPROGRESSDLG Or FCOPY.FOF_COPYALLNOPROMPT
'------------------------------
ExitHere:
    CopyFolderAndWait = True '!!!!!!!!!!!!
    Set objFolder = Nothing: Set objShell = Nothing
    Exit Function
'----------------
ErrHandle:
    ErrPrint2 "CopyFolderAndWait", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'======================================================================================================================================================
' Open Folder in Explorer
'=====================================================================================================================================================
Public Sub OpenFolderExplorer(sFolder As String)
On Error Resume Next

  Application.FollowHyperlink sFolder, , True
End Sub

'======================================================================================================================================================
' Create Multistep Folder
'======================================================================================================================================================
Public Function CreateFolderRecursive(destDir As String) As Boolean
Dim I As Long, prevDir As String

On Error GoTo ErrHandle
'----------------------------
For I = Len(destDir) To 1 Step -1
       If Mid(destDir, I, 1) = "\" Then
           prevDir = Left(destDir, I - 1)
           Exit For
       End If
Next I
   
   If prevDir = "" Then CreateFolderRecursive = False: Exit Function
   If Not Len(Dir(prevDir & "\", vbDirectory)) > 0 Then
       If Not CreateFolderRecursive(prevDir) Then CreateFolderRecursive = False: Exit Function
   End If

   MkDir destDir
'----------------------------
ExitHere:
   CreateFolderRecursive = True '!!!!!!!!!!!!!
   Exit Function
'-----------
ErrHandle:
   ErrPrint2 "CreateFolderRecursive", Err.Number, Err.Description, MOD_NAME
   Err.Clear
End Function
'======================================================================================================================================================
'  Build Path
'======================================================================================================================================================
Public Function BuildPath(sFolder As String, sFile As String) As String
Dim sRes As String
sRes = sFolder
If sRes <> "" Then
    If Right(sRes, 1) <> "\" Then sRes = sRes & "\"
End If
If sFile <> "" Then
    sRes = sRes & IIf(Left(sFile, 1) = "\", Right(sFile, Len(sFile) - 1), sFile)
End If
    BuildPath = sRes '!!!!!!!!!!!!!
End Function
'=====================================================================================================================================================
' Check if path is relative
'=====================================================================================================================================================
Public Function IsPathRelation(sPath As String, sROOT As String) As Boolean
Dim bRelation As Boolean                      ' Относительный путь (если выбираем относительно корня)
    
    On Error Resume Next
'------------------------------------------------------
    If sPath = "" Then Exit Function
    If Not IsFolderExists(sPath) Then Exit Function
    
    If sROOT = "" Then
       If MsgBox("Undefined Root Path. Pick the Root Folder Now? ", vbYesNoCancel, "Root Folder") = vbYes Then
          sROOT = OpenDialog(GC_FOLDER_PICKER, "Pict the Root Folder")
          If sROOT <> "" Then SetRoot sROOT
       End If
    ElseIf Not IsFolderExists(sROOT) Then
       If MsgBox("The Root Folder " & sROOT & " is not Accessible. Change the Root Now? ", _
                                                         vbYesNoCancel, "Root Path") = vbYes Then
          sROOT = OpenDialog(GC_FOLDER_PICKER, "Pict the Root Folder")
          If sROOT <> "" Then SetRoot sROOT
       End If
    End If
    If sROOT = "" Then Exit Function
'----------------------------------------------------------------------------------------------------
       If InStr(1, sPath, sROOT) > 0 Then bRelation = True
'------------------------
ExitHere:
       IsPathRelation = bRelation '!!!!!!!!!!!!!!!!!!!!!!!!!
End Function

'======================================================================================================================================================
'  Get Temp File Name
'======================================================================================================================================================
Public Function TempFileName() As String
Dim FSO As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")
    TempFileName = FSO.GetTempName()
    Set FSO = Nothing
End Function

'================================================================================================================================================
' Get ParentFolder
'================================================================================================================================================
Public Function GetParentFolder(Path As String, Optional nDeep As Integer = 1) As String
Dim TPATH() As String, nDim As Integer, sPath As String, I As Integer
Dim sRes As String

On Error GoTo ErrHandle
'-----------------------------------
If Path = "" Then Err.Raise 1000, , "No Correct Path"
sPath = Path: If Right(sPath, 1) = "\" Then sPath = Left(sPath, Len(sPath) - 1)
TPATH = Split(sPath, "\"): nDim = UBound(TPATH)

For I = 1 To nDeep
     If I > nDim Then Exit For
     If sRes <> "" Then sRes = "\" & sRes
     sRes = TPATH(nDim - I) & sRes
Next I

'---------------------------------
ExitHere:
    GetParentFolder = sRes '!!!!!!!!!!!!!!!
    Exit Function
'------------------
ErrHandle:
    ErrPrint "GetParentFolder", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function

'================================================================================================================================================
' Function Calculate String (Main Part of Path) to identify has file been loaded or not.
' The Root could be changed, the essential part will ot be changed
'================================================================================================================================================
Public Function GetFileDigest(sPath As String, Optional sROOT As String) As String
Dim sRes As String, sFolder As String

        sRes = Replace(sPath, sROOT, "")                  ' PATH WITHOUT ROOT
        sRes = Replace(sRes, Environ("USERPROFILE"), "")  ' PATH WITHOUT GENERAL PATH
'--------------------
ExitHere:
    sRes = Right(sRes, 100)                               ' 100 symbol limit
    GetFileDigest = sRes '!!!!!!!!!!!!!!!!
End Function
'================================================================================================================================================
' Create FileInsertSpec
'================================================================================================================================================
Public Function GenerateFileInsertSpec() As String
    GenerateFileInsertSpec = "TBL_FILE=;FLD_PATH=;FLD_ESSENTIAL=;" & _
            "FLD_SIZE=;FLD_DATE=;FLD_EXT=;FLD_TITLE=;FLD_PAGES=;FLD_AUTHORS=;FLD_DURATION=;FLD_FLDR="
End Function

'================================================================================================================================================
' Add File(s) To Folder
'================================================================================================================================================
Public Function AddFiles(Files As Variant, FSPEC As String, Optional TBL As String = "ITEMS", Optional FLD_DIGEST As String = "DIGEST") As String
Dim nDim As Long, sRes As String, I As Integer, sDigest As String, sROOT As String
Dim FIOM As FileInfo, sPath As String


For I = 1 To nDim
    sPath = CStr(Files(I))
    sDigest = GetFileDigest(sPath, sROOT)
    If IsFileInTBL(sDigest, FLD_DIGEST, TBL) = -1 Then  ' TIME TO ADD FILE INTO TBL
          FIOM = GetFileInfo(sPath, sROOT, sDigest, False, False, 2)
       
       '||||||||>>>>>
            ' TBDF
       '<|||||||||||||
       
       
    End If
Next I
'-----------------------------
ExitHere:
    Exit Function
'--------------
ErrHandle:
    ErrPrint "AddFiles", Err.Number, Err.Description
    Err.Clear
End Function
'----------------------------------------------------------------------------------------------------------------------------------------------
' Check if File is placed in TBL already. If not - returns "-1"
'----------------------------------------------------------------------------------------------------------------------------------------------
Private Function IsFileInTBL(sDigest As String, Optional FLD_DIGEST As String = "DIGEST", Optional TBL As String = "ITEMS") As Long
       IsFileInTBL = Nz(DLookup("ID", SHT(TBL), FLD_DIGEST & "=" & sCH(sDigest)), -1) '!!!!!!!!!!!!!!!!
End Function

'================================================================================================================================================
' Form FileInfo Array to add to TBL
'================================================================================================================================================
Public Function CollectFileToAdd(Files As Variant, Optional TBL As String = "ITEMS", Optional FLD_DIGEST As String = "DIGEST", _
                Optional bGetExtendedPop As Boolean, Optional bCRC32 As Boolean = False, Optional nParentFolderDeep As Integer = 2, _
                                                                                          Optional bShowProgress As Boolean = True) As FileInfo()
Dim FINO() As FileInfo, nDim As Integer, nFiles As Integer, I As Integer
Dim sROOT As String, sDigest As String, dProgress As Double, dPrevProgress As Double
      
      
Const dSHAG As Double = 0.1

On Error GoTo ErrHandle
'-----------------------------------------------------
      If Not IsArray(Files) Then Err.Raise 1000, , "Not qualified array"
      sROOT = GetRoot(): nFiles = UBound(Files): nDim = -1: ReDim FINO(0)
      '-----------------
      
         
      For I = 0 To nFiles
            '------------------------------------------------------------
            ' ESC CATCH (to cancel very long process)
            If IsKeyPress(27) Then
               If MsgBox("Do you want to cancel fileinfo gathering now?", _
                     vbYesNoCancel + vbQuestion, "CollectFileToAdd") = vbYes Then GoTo ExitHere
            End If
            dProgress = I / nFiles
            If (dProgress - dPrevProgress) > dSHAG Then
                dPrevProgress = dProgress
                ShowProgressBar dProgress, 3, , "Calculate File " & I & " From " & nFiles
            End If
            
            sDigest = GetFileDigest(CStr(Files(I)), sROOT)
            
            If sDigest = "" Then GoTo NextFile
            If IsFileInTBL(sDigest, FLD_DIGEST, TBL) = -1 Then  ' TIME TO ADD FILE INTO TBL
                    nDim = nDim + 1: ReDim Preserve FINO(nDim)
                    FINO(nDim) = GetFileInfo(CStr(Files(I)), sROOT, sDigest, bGetExtendedPop, bCRC32, nParentFolderDeep)
            End If
NextFile:
      Next I
'-----------------------------------------------------
ExitHere:

      Exit Function
'-------------
ErrHandle:
      ErrPrint "GetFileInfoToAdd", Err.Number, Err.Description
      Err.Clear
End Function

'================================================================================================================================================
' Process FileInfo as string
'================================================================================================================================================
Public Function GetFileInfo2(sPath As String, sROOT As String, FileDigest As String, Optional bGetExtendedPop As Boolean, _
                            Optional bCRC32 As Boolean = False, Optional nParentFolderDeep As Integer = 1, Optional DLM As String = ";") As String
Dim sRes As String, fi As FileInfo

fi = GetFileInfo(sPath, GetRoot(), GetFileDigest(sPath, GetRoot()), True, True, 1)


sRes = fi.FullPath & DLM & fi.FilePath & DLM & fi.ParentFolder & DLM & fi.fileName & DLM & fi.FileExt & DLM
sRes = sRes & fi.FileSize & DLM & fi.FileDate & DLM & fi.FileHash & DLM & fi.FileCRC32
sRes = fi.FileDigest & DLM & sRes & DLM & Replace(fi.FileProps, DLM, UFDELIM)
'----------------------------------------------
      GetFileInfo2 = sRes '!!!!!!!!!!!!!!!!!
End Function
'================================================================================================================================================
' Process FileInfo
'================================================================================================================================================
Public Function GetFileInfo(sPath As String, sROOT As String, FileDigest As String, Optional bGetExtendedPop As Boolean, _
                                                      Optional bCRC32 As Boolean = False, Optional nParentFolderDeep As Integer = 1) As FileInfo
Dim TFILE As FileInfo, sWork As String, KWS() As String
Dim FPROPS As cArrayList
Const DLM As String = ";"
Const SEQV As String = "="

On Error GoTo ErrHandle
'-------------------------------
        TFILE.FullPath = sPath
        TFILE.FilePath = Replace(sPath, sROOT, "")
        TFILE.FileDigest = FileDigest
        
        TFILE.fileName = FileNameOnly(sPath)
        TFILE.FileSize = GetFileSize(sPath)
        TFILE.FileExt = UCase(FileExt(sPath))
        TFILE.ParentFolder = GetParentFolder(sPath, nParentFolderDeep)
        TFILE.FileDate = GetFileDate(sPath)
        
        If bCRC32 Then TFILE.FileCRC32 = GetFileHash(sPath)
'-------------------------------
If bGetExtendedPop Then
        Set FPROPS = New cArrayList
        Select Case TFILE.FileExt
        Case "DOC", "DOCX", "PDF", "XLS", "XLSX", "PPT", "PPTX":
            sWork = GetFileAuthors(sPath)
            If sWork <> "" Then
               FPROPS.Add "AUTHORS" & SEQV & sWork
               sWork = ""
            End If
            sWork = GetFilePages(sPath)
            If sWork <> "" Then
               FPROPS.Add "PAGES" & SEQV & sWork
               sWork = ""
            End If
            
            sWork = FPROPS.ToString(DLM)
            Set FPROPS = Nothing
                        
        Case "MP4", "AVI", "MPEG":
            sWork = GetMediaFileDuration(sPath)
            If sWork <> "" Then
               FPROPS.Add "CRC32" & SEQV & sWork
            End If
        End Select
   
        If sWork <> "" Then TFILE.FileProps = FPROPS.ToString(DLM)
End If
           
'-------------------------------
ExitHere:
    GetFileInfo = TFILE '!!!!!!!!!!!!!!!
    Set FPROPS = Nothing
    Exit Function
'----------
ErrHandle:
    ErrPrint "GetFileInfo", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'===================================================================================================================
' Get PDF Indormation
'===================================================================================================================
Public Function GetPDFMetaData(ByVal sFile As String, Optional DLM As String = ";", _
                                                                          Optional SEQV As String = "=") As String
Dim oApp As Object, sExt As String
Dim oDoc As Object, sRes As String
Dim sVal As String
  
Set oApp = CreateObject("AcroExch.App")
Set oDoc = CreateObject("AcroExch.PDDoc")


On Error GoTo ErrHandle
'-------------------------------------------------------------
    sExt = FileExt(sFile): If Not UCase(sExt) = "PDF" Then Exit Function

    With oDoc
        If .Open(sFile) Then
             sVal = .GetInfo("Title"): If sVal <> "" Then sRes = IIf(sRes <> "", sRes & DLM, "") & "Title" & SEQV & sVal
             sVal = .GetInfo("Producer"): If sVal <> "" Then sRes = IIf(sRes <> "", sRes & DLM, "") & "Producer" & SEQV & sVal
             sVal = .GetInfo("Subject"): If sVal <> "" Then sRes = IIf(sRes <> "", sRes & DLM, "") & "Subject" & SEQV & sVal
             sVal = .GetInfo("Author"): If sVal <> "" Then sRes = IIf(sRes <> "", sRes & DLM, "") & "Author" & SEQV & sVal
             sVal = .GetInfo("Keywords"): If sVal <> "" Then sRes = IIf(sRes <> "", sRes & DLM, "") & "KeyWords" & SEQV & sVal
            .Close
        End If
    End With
'------------------------------
ExitHere:
    GetPDFMetaData = sRes '!!!!!!!!!!!!!!!!!!!!
    Set oDoc = Nothing: Set oApp = Nothing
    Exit Function
'---------------------
ErrHandle:
    ErrPrint "GetPDFMetaData", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'===================================================================================================================
' Get WORD Indormation
'===================================================================================================================
Public Function GetOfficeMetaData(ByVal sFile As String, Optional bCountPages As Boolean = False, _
                                              Optional DLM As String = ";", Optional SEQV As String = "=") As String
Dim sExt As String, bRes As Boolean, sRes As String
Dim sVal As String

On Error GoTo ErrHandle
'-------------------------------------------------------------
    If Dir(sFile) = "" Then Exit Function
    sExt = FileExt(sFile)
    
    Select Case UCase(sExt)
    Case "DOC", "DOCX", "XLS", "XLSX", "PPT", "PPTX":
         bRes = True
    Case Else
    End Select
    If Not bRes Then Exit Function
             sVal = GetFileProperties(sFile, 21): If sVal <> "" Then sRes = IIf(sRes <> "", sRes & DLM, "") & "Title" & SEQV & sVal
             sVal = GetFileProperties(sFile, 33): If sVal <> "" Then sRes = IIf(sRes <> "", sRes & DLM, "") & "Company" & SEQV & sVal
             sVal = GetFileProperties(sFile, 22): If sVal <> "" Then sRes = IIf(sRes <> "", sRes & DLM, "") & "Subject" & SEQV & sVal
             sVal = GetFileProperties(sFile, 20): If sVal <> "" Then sRes = IIf(sRes <> "", sRes & DLM, "") & "Author" & SEQV & sVal
             sVal = GetFileProperties(sFile, 18): If sVal <> "" Then sRes = IIf(sRes <> "", sRes & DLM, "") & "KeyWords" & SEQV & sVal
             
             If bCountPages Then
                 sVal = CStr(GetFilePages(sFile))
                 If sVal <> "0" Then sRes = IIf(sRes <> "", sRes & DLM, "") & "Pages" & SEQV & sVal
             End If
'------------------------------
ExitHere:
    GetOfficeMetaData = sRes '!!!!!!!!!!!!!!!!!!!!
    Exit Function
'---------------------
ErrHandle:
    ErrPrint "GetOfficeMetaData", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function


'===================================================================================================================
' Get File HASH - CRC32 (by VBA - slow) or MD5  (by cripto dll - fast)
'===================================================================================================================
Public Function GetFileHash(sPath As String, Optional bMD5 As Boolean = True) As String
Dim CRPT As New cCRYPTO, sRes As String
    
    On Error GoTo ErrHandle
'-------------------------
If bMD5 Then
       sRes = CalcMD5(sPath)
Else
    Set CRPT = New cCRYPTO
        sRes = CRPT.CRC32_File(sPath)
    Set CRPT = Nothing
End If
'----------------------------
ExitHere:
    GetFileHash = sRes '!!!!!!!!!!
    Exit Function
'------------
ErrHandle:
    ErrPrint "GetFileHash", Err.Number, Err.Description
    Err.Clear
End Function
'====================================================================================================================
' Execute File Command
'====================================================================================================================
Public Sub Execute(ExecPath As String, Optional ParamLine As String, Optional iFocus As Long = vbNormalFocus)
Dim RetVal As Variant, ExecLine As String
On Error GoTo ErrHandle
     ExecLine = ExecPath & " " & ParamLine
     RetVal = Shell(ExecLine, iFocus)
'-------------------------------
ExitHere:
     Exit Sub
'--------------------
ErrHandle:
     ErrPrint "Execute", Err.Number, Err.Description
     Err.Clear
End Sub
'==================================================================================================================================
' This Function create process and wait it to finish
'==================================================================================================================================
Public Function SyncShell(sCommand As String, Optional waitOnReturn As Boolean = True, _
                                                                        Optional WindowStyle As VbAppWinStyle = vbHide) As Boolean
Dim wsh As Object
Dim lngErrorCode As Long

On Error GoTo ErrHandle
'-------------------------------------------
    If sCommand = "" Then Exit Function
    Set wsh = VBA.CreateObject("WScript.Shell")
    lngErrorCode = wsh.Run(sCommand, WindowStyle, waitOnReturn)
'------------------------
ExitHere:
    SyncShell = (lngErrorCode = 0)
    Exit Function
'----------
ErrHandle:
    ErrPrint "SyncShell", Err.Number, Err.Description
    Err.Clear
End Function

'=======================================================================================================================================================
' Copy Or Move File
'=======================================================================================================================================================
Public Function CopyFile(FromPath As String, ToPath As String, Optional bOverwrite As Boolean = True, Optional bMove As Boolean = False) As Boolean
Dim bRes As Boolean
Dim FSO As Object
Set FSO = VBA.CreateObject("Scripting.FileSystemObject")

On Error GoTo ErrHandle
'------------------------------
    If Not FSO.FileExists(FromPath) Then
        MsgBox FromPath & " does not exist!", vbExclamation, "Source File Missing"
        GoTo ExitHere
    ElseIf FSO.FileExists(ToPath) And Not bOverwrite Then
        MsgBox ToPath & " already exists!", vbExclamation, "Destination File Exists"
        GoTo ExitHere
    End If
'-------------------------------
If bMove Then
   Call FSO.MoveFile(FromPath, ToPath)
   bRes = True
Else
   Call FSO.CopyFile(FromPath, ToPath, bOverwrite)
   bRes = True
End If
'-------------------------------
ExitHere:
    CopyFile = bRes '!!!!!!!!!!!
    Set FSO = Nothing
    Exit Function
'-----------
ErrHandle:
    ErrPrint "CopyFile", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function

'=======================================================================================================================================================
' Rename File
'=======================================================================================================================================================
Public Function RenameFile(OldPath As String, NewPath As String) As Boolean
Dim bRes As Boolean
On Error GoTo ErrHandle
'-----------------------------
If Dir(OldPath) = "" Then
        MsgBox OldPath & " does not exist!", vbExclamation, "Source File Missing"
        GoTo ExitHere
ElseIf Dir(NewPath) <> "" Then
        MsgBox NewPath & " already exists!", vbExclamation, "Destination File Exists"
        GoTo ExitHere
End If
If IsFileOpen(OldPath) Then
        MsgBox "The file " & FileNameOnly(OldPath) & "is opened. Can Produce operation", vbCritical, "RenameFile"
        Exit Function
End If
'-----------------------------
    Name OldPath As NewPath
    bRes = True
'--------------------------------
ExitHere:
    RenameFile = bRes '!!!!!!!!!!!!!!
    Exit Function
'------------------
ErrHandle:
    ErrPrint "RenameFile", Err.Number, Err.Description
    Err.Clear
End Function

'=======================================================================================================================================================
' Check if file open
'=======================================================================================================================================================
Public Function IsFileOpen(fileName As String) As Boolean
    Dim FileNum As Integer, errnum As Integer

    On Error Resume Next   ' Turn error checking off.
    FileNum = FreeFile()   ' Get a free file number.
    ' Attempt to open the file and lock it.
    Open fileName For Input Lock Read As #FileNum
    Close FileNum          ' Close the file.
    errnum = Err           ' Save the error number that occurred.
    On Error GoTo 0        ' Turn error checking back on.

    ' Check to see which error occurred.
    Select Case errnum

        ' No error occurred.
        ' File is NOT already open by another user.
        Case 0
         IsFileOpen = False

        ' Error number for "Permission Denied."
        ' File is already opened by another user.
        Case 70
            IsFileOpen = True

        ' Another error occurred.
        Case Else
            Error errnum
    End Select

End Function


'=======================================================================================================================================================
' Get File Duration VIA GetFileProperties
'=======================================================================================================================================================
Public Function GetMediaFileDuration(sFile As String) As String
       GetMediaFileDuration = GetFileProperties(sFile, 27)
End Function

'=======================================================================================================================================================
' Get File Author
'=======================================================================================================================================================
Public Function GetFileAuthors(sFile As String) As String
       GetFileAuthors = GetFileProperties(sFile, 20)
End Function


'=======================================================================================================================================================
' Get File Size
'=======================================================================================================================================================
Public Function GetFileSize(sFile As String) As Long
    If sFile = "" Then Exit Function
    If Dir(sFile) = "" Then Exit Function

      GetFileSize = FileLen(sFile) '!!!!!!!!!!!!
End Function

'=======================================================================================================================================================
' Get File DateTime
'=======================================================================================================================================================
Public Function GetFileDate(sFile As String) As String
    If sFile = "" Then Exit Function
    If Dir(sFile) = "" Then Exit Function

      GetFileDate = FileDateTime(sFile) '!!!!!!!!!!!!
End Function

'=======================================================================================================================================================
' Get File Pages
'=======================================================================================================================================================
Public Function GetFilePages(sFile As String) As Integer
Dim sExt As String, nPages As Integer
    
    If sFile = "" Then Exit Function
    If Dir(sFile) = "" Then Exit Function
    sExt = UCase(FileExt(sFile))
    
    Select Case sExt
    Case "PDF":
        nPages = PDFpageCount(sFile)
    Case "DOC", "DOCX":
        nPages = WordPageCount(sFile)
    Case "PPT", "PPTX":
        nPages = GetFileProperties(sFile, 151)
    End Select
'---------------------------
ExitHere:
    GetFilePages = nPages '!!!!!!!!!!!
End Function

Private Function WordPageCount(sFilePathName As String) As Integer
Dim oApp As Object, oDoc As Object, nPages As Integer

Const wdStatisticPages = 2

On Error GoTo ErrHandle
'-----------------------------------
    Set oApp = CreateObject("Word.Application")
    Set oDoc = oApp.Documents.Open(sFilePathName)
    
    oDoc.Repaginate
    nPages = oDoc.ComputeStatistics(wdStatisticPages)
    
    oDoc.Close False: oApp.Quit
'-------------------
ExitHere:
    Set oDoc = Nothing: Set oApp = Nothing
    WordPageCount = nPages '!!!!!!!!
    Exit Function
'--------------
ErrHandle:
    Beep
    ErrPrint "", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
Private Function PDFpageCount(sFilePathName As String) As Integer

Dim nFileNum As Integer
Dim sInput As String
Dim sNumPages As String
Dim iPosN1 As Integer, iPosN2 As Integer
Dim iPosCount1 As Integer, iPosCount2 As Integer
Dim iEndsearch As Integer

On Error GoTo ErrHandle
'------------------------------
nFileNum = FreeFile

Open sFilePathName For Binary Lock Read Write As #nFileNum
  
  Do Until EOF(nFileNum)
      Input #1, sInput
      sInput = UCase(sInput)
      iPosN1 = InStr(1, sInput, "/N ") + 3
      iPosN2 = InStr(iPosN1, sInput, "/")
      iPosCount1 = InStr(1, sInput, "/COUNT ") + 7
      iPosCount2 = InStr(iPosCount1, sInput, "/")
      
   If iPosN1 > 3 Then
      sNumPages = Mid(sInput, iPosN1, iPosN2 - iPosN1)
      Exit Do
   ElseIf iPosCount1 > 7 Then
      sNumPages = Mid(sInput, iPosCount1, iPosCount2 - iPosCount1)
      Exit Do
   ' Prevent overflow and assigns 0 to number of pages if strings are not in binary
   ElseIf iEndsearch > 1001 Then
      sNumPages = "0"
      Exit Do
   End If
      iEndsearch = iEndsearch + 1
   Loop
'-----------------------------
ExitHere:
  Close #nFileNum
  
  PDFpageCount = CInt(sNumPages)
  Exit Function
'--------------
ErrHandle:
  ErrPrint "PDFpageCount", Err.Number, Err.Description
  Err.Clear
End Function
'========================================================================================================================================================
' Get File Extended Property
' This function receives extended information about the file. The list of properties is wide and differs depending on the file type.
' Some examples:
'           0 - Name           1 - Size             2 - Item type   3 - Date modified   4 - Date created    5 - Date accessed
'           6 - Attributes     7 - Offline status   8 - Availability 9 - Perceived type
'           10 - Owner     11 - Kind   12 - Date taken 13 - Contributing artists   14 - Album  15 - Year   16 - Genre
'           18 - Tags  19 - Rating 20 - Authors    21 - Title  22 - Subject    23 - Categories
'           24 - Comments  25 - Copyright  26 - #  27 - Length 28 - Bit rate
'           33 - Company   34 - File description   35 - Program name   36 - Duration   37 - Is online
'           39 - Location 126 - Web page 127 - Content status 128 - Content type 129 - Date acquired
'           130 - Date archived    131 - Date completed 157 - File extension
'           158 - Filename 159 - File version 160 - Flag colour
'           168 - Horizontal resolution 169 - Width 170 - Vertical resolution 171 - Height
'           172 - Importance 175 - Encryption status 183 - Folder name 184 - Folder path 185 - Folder
'           186 - Participants 187 - Path 192 - Language 193 - Date visited
'           208 - Subtitle 209 - User web URL 210 - Writers
'           230 - Album artist 231 - Sort album artist 281 - Summary 303 - Video compression
'           306 - Frame height 307 - Frame rate 308 - Frame width 309 - Video orientation
'           310 - Total bitrate
' More information see here: https://technet.microsoft.com/en-us/library/ee176615.aspx,
'                            https://www.access-programmers.co.uk/forums/attachment.php?attachmentid=66654&d=1498350162
'========================================================================================================================================================
Public Function GetFileProperties(file As String, propertyVal As Integer) As Variant
Dim varfolder, varfile
    
On Error Resume Next
'----------------------------------------------------------
    With CreateObject("Shell.Application")
        Set varfolder = .NameSpace(Left(file, InStrRev(file, "\") - 1))
        Set varfile = varfolder.ParseName(Right(file, Len(file) - InStrRev(file, "\")))
        GetFileProperties = varfolder.GetDetailsOf(varfile, propertyVal)
    End With
    
End Function

'=======================================================================================================================================================
' CREATE TEMPORY FILENAME(PATH
' RETURN STRING VAR AS PATH TO FILE
'=======================================================================================================================================================
Public Function GetTempFile(sPrefix As String, Optional EXT As String) As String
Dim sTmpPath As String * 512
Dim sTmpName As String * 576
Dim nRet As Long
Dim sRes As String

On Error GoTo ErrHandle
'--------------------------------------------------------------
         nRet = GetTempPath(512, sTmpPath)
         If (nRet > 0 And nRet < 512) Then
            nRet = GetTempFileName(sTmpPath, sPrefix, 0, sTmpName)
            If nRet <> 0 Then
               sRes = Left$(sTmpName, _
                  InStr(sTmpName, vbNullChar) - 1)
            End If
         End If
'--------------------------------------------------------------
If EXT <> "" And sRes <> "" Then
      sRes = Left(sRes, InStrRev(sRes, ".")) & EXT
End If
'-------------------------------
ExitHere:
      GetTempFile = sRes '!!!!!!!!!!!!!!
      Exit Function
'------------
ErrHandle:
      ErrPrint "GetTempFile", Err.Number, Err.Description
      Err.Clear: Resume ExitHere
End Function

'=====================================================================================================================
' Open File with associated app
'=====================================================================================================================
Public Sub OpenFile(sPath As String, Optional bFileLocation As Boolean = False)
Dim MyFLS As New cFileSystem

On Error GoTo ErrHandle
'-------------------------------------------
   If Dir(sPath) = "" Then Exit Sub
       If Not bFileLocation Then
           Call MyFLS.bOpenFileShell(sPath)
       Else
           Call OpenFileLocation(sPath)
       End If
'-------------------------------------------
ExitHere:
      Set MyFLS = Nothing
      Exit Sub
'---------------
ErrHandle:
      ErrPrint "OpenFile", Err.Number, Err.Description
      Err.Clear
End Sub

'=====================================================================================================================
' Open Folder and Locate File
'=====================================================================================================================
Public Sub OpenFileLocation(sPath As String)
   Call Shell("explorer.exe /select ," & sPath, vbNormalFocus)
End Sub

'=======================================================================================================================================================
' LIST ALL FILES IN FOLDER AND SUBFOLDERS
' RETURN STRING ARRAY
'=======================================================================================================================================================
Public Function LISTFILES(Optional StartFolder As String = "C:\") As String()
Dim objFSO As Object                                   ' Основной объект - манипулятор файловой системой
Dim objTopFolder As Object                             ' Корневой каталог
Dim FLS() As String, nFiles As Integer                 ' Основной массив файловых указателей
Dim sRes As String

Const DLM As String = ";"

On Error GoTo ErrHandle
'--------------------------------------------------------------
nFiles = -1: ReDim FLS(0)
If StartFolder = "" Then GoTo ExitHere
If Dir(StartFolder, vbDirectory) = "" Then GoTo ExitHere
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTopFolder = objFSO.GetFolder(StartFolder)
'--------------------------------------------------------------
' Запускаем рекурсивную процедуру
    sRes = RecursiveFolder(objTopFolder, objFSO, True) '!!!!!!!!!!!!!
    If Left(sRes, Len(DLM)) = DLM Then sRes = Right(sRes, Len(sRes) - Len(DLM))
    FLS = Split(sRes, DLM): nFiles = UBound(FLS)
'-----------------------------------------
ExitHere:
    LISTFILES = FLS '!!!!!!!!!!!!!
    Set objFSO = Nothing: Set objTopFolder = Nothing
    Exit Function
'--------------------------
ErrHandle:
   ErrPrint "LISTFILES", Err.Number, Err.Description
   Err.Clear: Resume ExitHere
End Function
'---------------------------------------------------------------------------------------------------------------------
' Рекурсивная функция, вычисляющая файлы в директории
'---------------------------------------------------------------------------------------------------------------------
Private Function RecursiveFolder(ByVal objFolder As Object, _
                     ByRef objFSO, Optional IncludeSubFolders As Boolean = True, Optional DLM As String = ";") As String
Dim objFile As Object                        '  Scripting.File
Dim objSubFolder As Object                   '  Scripting.folder
Dim sPath As String, nDim As Integer
Dim ProgressStep As Double
Dim sRes As String, sWork As String
On Error GoTo ErrHandle

'------------------------------------------------------------------------------------
    For Each objFile In objFolder.Files
                  sRes = sRes & ";" & objFile.Path
    Next objFile
'------------------------------------------------------------------------------------
    If IncludeSubFolders Then
        For Each objSubFolder In objFolder.SubFolders
          '----------------------------------------------------------------------------
          sRes = sRes & RecursiveFolder(objSubFolder, objFSO, IncludeSubFolders, DLM)
          '----------------------------------------------------------------------------
        Next objSubFolder
    End If
'-----------------------------------------------------------------------------------------
ExitHere:
          RecursiveFolder = sRes '!!!!!!!!!!!!!!
          Exit Function
'----------------------------------------
ErrHandle:
          ErrPrint "RecursiveFolder", Err.Number, Err.Description
          Err.Clear: Resume ExitHere
End Function

'=====================================================================================================================================================
' GetFileName From Path
'=====================================================================================================================================================
Public Function FileNameOnly(strFullPath As String) As String
Dim DLM As String
    DLM = IIf(InStr(1, strFullPath, "\") > 0, "\", "/")
     FileNameOnly = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, DLM))
End Function

'=====================================================================================================================================================
' Get File Extension
'=====================================================================================================================================================
Public Function FileExt(sFile As String) As String
Dim iL As Integer
     If sFile = "" Then Exit Function
     iL = InStrRev(sFile, "."): If iL = 0 Then Exit Function
     iL = Len(sFile) - iL: If iL > 6 Then Exit Function
     
     FileExt = Right(sFile, iL) '!!!!!!!!!!!!!!!!!!
End Function

'=====================================================================================================================================================
' GetFileName Without Extension
'=====================================================================================================================================================
Public Function FilenameWithoutExtension(ByVal sPath As String) As String
Dim FSO As Object

Set FSO = CreateObject("Scripting.FileSystemObject")
 
    FilenameWithoutExtension = FSO.GetBaseName(sPath)  '!!!!!
    Set FSO = Nothing
End Function
'=====================================================================================================================================================
' Get FolderName for abstract path
'=====================================================================================================================================================
Public Function FolderNameOnlyAbstract(sPath As String) As String
Dim sRes As String, iL As Integer
  
If sPath = "" Then Exit Function
'------------------------------------
  iL = InStrRev(sPath, ".")
  
  If iL > 0 Then
       iL = Len(sPath) - iL
       If iL > 1 Or iL < 7 Then
           iL = InStrRev(sPath, "\")
           If iL > 0 Then sRes = Left(sPath, iL)
       End If
  Else
       sRes = sPath
  End If
'--------------------------------------------------------
  FolderNameOnlyAbstract = sRes '!!!!!!!!!!!!!!!!!!!
End Function
'=====================================================================================================================================================
' Parent Folder
'=====================================================================================================================================================
Public Function ParentFolder(sFullPath As String) As String
Dim PARTS() As String, nParts As Integer
Const DLM As String = "\"

On Error Resume Next
'------------------------------
If sFullPath = "" Then Exit Function
PARTS = Split(sFullPath, DLM): nParts = UBound(PARTS)
If nParts = 0 Then Exit Function
'------------------------------
ExitHere:
       ParentFolder = PARTS(nParts - 1) '!!!!!!!!!!!!!
End Function
'=====================================================================================================================================================
' Convert size in bytes to b,kb,mb,gb
'=====================================================================================================================================================
Public Function FormatFileSize(FileBytes As Variant) As String
Dim sRes As String

On Error GoTo ErrHandle
'---------------------------------------
If IsEmpty(FileBytes) Then
   sRes = "ERR"
   GoTo ExitHere
End If
If Not IsNumeric(FileBytes) Then
   sRes = "ERR"
   GoTo ExitHere
End If
'---------------------------------------
Select Case FileBytes
        Case 0 To 1023
            sRes = Format(FileBytes, "0") & "B"
        Case 1024 To 104875
            sRes = Format(FileBytes / 1024, "0") & "KB"
        Case 104876 To 1073741823
            sRes = Format(FileBytes / 104876, "0") & "MB"
        Case 1073741824 To 1.11111111111074E+20
            sRes = Format(FileBytes / 1073741823, "0.00") & "GB"
    End Select
'---------------------------------------
ExitHere:
     FormatFileSize = sRes '!!!!!!!!!!!!!!!!
     Exit Function
'-----------
ErrHandle:
     ErrPrint "FormatFileSize", Err.Number, Err.Description
     Err.Clear
     Resume ExitHere
End Function
'=====================================================================================================================================================
' CREATE FOLDERS AND SUB FOLDERS
'=====================================================================================================================================================
Public Function FolderCreate(ByVal sPath As String) As Boolean
Dim FSO As Object, bRes As Boolean

On Error GoTo ErrHandle
'------------------------------
    bRes = True: If IsFolderExists(sPath) Then GoTo ExitHere
    Set FSO = CreateObject("scripting.filesystemobject")

    FSO.CreateFolder sPath
'------------------------------
ExitHere:
    FolderCreate = bRes '!!!!!!!!!!!!!!
    Set FSO = Nothing
    Exit Function
'----------------
ErrHandle:
    ErrPrint "FolderCreate", Err.Number, Err.Description
    Err.Clear: bRes = False: Resume ExitHere
End Function
'=====================================================================================================================================================
' Change File Xtension
'=====================================================================================================================================================
Public Function ChangeXtension(sFile As String, sNewExt As String) As String
Dim iL As Long, sRes As String
    iL = InStrRev(sFile, ".")
    If iL > 0 Then
         sRes = Left(sFile, iL) & sNewExt
    Else
         sRes = sFile & "." & sNewExt
    End If
'--------------------------
ExitHere:
    ChangeXtension = sRes '!!!!!!!!!!!!!
End Function

'=====================================================================================================================================================
' Copy/Move Folder, on Success Return new location
'=====================================================================================================================================================
Public Function CopyFolder(FromPath As String, ToPath As String, Optional bMove As Boolean = False, Optional bShowProgress As Boolean = False) As String
Dim OP As SHFILEOPSTRUCT
Dim sRes As String
Dim FSO As Object

On Error GoTo ErrHandle
'---------------------------------------
    If Right(FromPath, 1) = "\" Then
        FromPath = Left(FromPath, Len(FromPath) - 1)
    End If

    If Right(ToPath, 1) = "\" Then
        ToPath = Left(ToPath, Len(ToPath) - 1)
    End If
    Set FSO = CreateObject("scripting.filesystemobject")
    If FSO.folderexists(FromPath) = False Then
        ErrPrint "CopyFolder", 0, FromPath & " doesn't exist"
        GoTo ExitHere
    End If
'----------------------------------------------------------------------------------------
    If bShowProgress Then   ' SHFileOperation
        With OP
            .wFunc = IIf(bMove, FO_MOVE, FO_COPY)
            .pTo = ToPath
            .pFrom = FromPath
            .fFlags = FOF_SIMPLEPROGRESS
       End With
        '~~> Perform operation
       SHFileOperation OP
    Else                    ' FSO ONLY
            DoCmd.Hourglass True
        If bMove Then
            If FSO.folderexists(ToPath) = True Then
                ErrPrint "CopyFolder", 0, ToPath & " exist, not possible to move to a existing folder"
                GoTo ExitHere
            End If
            FSO.MoveFolder Source:=FromPath, Destination:=ToPath
            sRes = ToPath
        Else
            FSO.CopyFolder Source:=FromPath, Destination:=ToPath
            sRes = ToPath
        End If
    End If
'---------------------------------------
ExitHere:
     DoCmd.Hourglass False
     CopyFolder = sRes '!!!!!!!!!!!
     Set FSO = Nothing
     Exit Function
'--------------------------
ErrHandle:
     ErrPrint "CopyFolder", Err.Number, Err.Description
     Err.Clear: Resume ExitHere
End Function

'=====================================================================================================================================================
' Функция выбирает файл
'=====================================================================================================================================================
Public Function PICKFILE(Optional MSG As String = "Выберите файл", Optional sFiltre As String = "Все файлы,*.*", _
                           Optional bMultiSelect As Boolean = False, Optional sInitialFolder As String = "") As String
      PICKFILE = OpenDialog(1, MSG, sFiltre, bMultiSelect, IIf(sInitialFolder = "", GetRoot(), sInitialFolder))
End Function

'=====================================================================================================================================================
' Функция выбирает каталог
'=====================================================================================================================================================
Public Function PICKFOLDER(Optional MSG As String = "Выберите каталог", Optional sInitialFolder As String = "") As String
      PICKFOLDER = OpenDialog(4, MSG, , , IIf(sInitialFolder = "", GetRoot(), sInitialFolder))
End Function

'=====================================================================================================================================================
' Функция проверяет существование файла
'=====================================================================================================================================================
Public Function IsFileExists(ByVal strFile As String, Optional bFindFolders As Boolean = False) As Boolean
    Dim lngAttributes As Long
    On Error Resume Next
'------------------------------------------------------------
' Учитываем атрибуты ДЛЯ ЧТЕНИЯ, СКРЫТЫЕ, СИСТЕМНЫЕ
    lngAttributes = (vbReadOnly Or vbHidden Or vbSystem)
    If bFindFolders Then
        lngAttributes = (lngAttributes Or vbDirectory)       ' Проверяем, включаем ли символ
    Else
        Do While Right$(strFile, 1) = "\"                    ' Удаляем последний символ, так чтобы не смотреть внутрь каталога
            strFile = Left$(strFile, Len(strFile) - 1)
        Loop
    End If
'-----------------------------------------------------------
    IsFileExists = (Len(Dir(strFile, lngAttributes)) > 0)    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function

'======================================================================================================================================================
' Get list of parents folders
'======================================================================================================================================================
Public Function GetAllParentFolders(sFolder, Optional DLM As String = ";") As String
Dim a() As String, nDim As Integer
Dim sParent As String, sWork As String

    ReDim a(0): nDim = -1: sWork = sFolder
    Do While sWork <> ""
       sParent = GetPathForParentFolder(sWork)
       If sParent = sWork Then Exit Do
            nDim = nDim + 1: ReDim Preserve a(nDim)
            a(nDim) = sParent: sWork = sParent
     Loop
'-----------------------------
ExitHere:
    GetAllParentFolders = Join(a, DLM) '!!!!!!!!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Get full path for parent folder
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetPathForParentFolder(Folder As String) As String
Dim sFolder As String, sRes As String
Dim iL As Integer
    
If Folder = "" Then Exit Function
If Right(Folder, 2) = ":\" Or Right(Folder, 1) = ":" Then
    sRes = Folder: GoTo ExitHere
End If

If Right(Folder, 1) = "\" Then
    sFolder = Left(Folder, Len(Folder) - 1)
Else
    sFolder = Folder
End If
   iL = InStrRev(sFolder, "\"):  sRes = Left(sFolder, iL - 1)
'---------------------------
ExitHere:
    GetPathForParentFolder = sRes '!!!!!!!!!!!!
End Function

'======================================================================================================================================================
' Get list of child folders
'======================================================================================================================================================
Public Function GetAllChildFolders(sFolder As String, Optional DLM As String = ";") As String
Dim sRes As String, sWork As String
Dim fs As Object, f As Object, sf As Object
    
    On Error GoTo ErrHandle
'--------------------------
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set sf = fs.GetFolder(sFolder).SubFolders
    
    For Each f In sf
        sRes = sRes & DLM & f.Path
        sWork = GetAllChildFolders(f.Path)
        If sWork <> "" Then sRes = sRes & sWork
    Next
'--------------------------
ExitHere:
    GetAllChildFolders = sRes '!!!!!!!!!!!!!!
    Set f = Nothing: Set sf = Nothing: Set fs = Nothing
    Exit Function
'----------------
ErrHandle:
    ErrPrint2 "GetAllChildFolders", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'=====================================================================================================================
' Return List of all files with recursive search
'=====================================================================================================================
Public Function LISTALLFILES(StartFolder As String, Optional FileExt As String = "") As String()
Dim FSO As Object                                       ' Основной объект - манипулятор файловой системой
Dim TopFolder As Object                                 ' Корневой каталог
Dim sRes() As String                                    ' Возвращаемый массив файлов
    On Error GoTo ErrHandle
    ReDim sRes(0)                                       ' Начальное положение
'----------------------------------------------------------------
    If StartFolder = "" Or Not IsFolderExists(StartFolder) Then Err.Raise 1000, , "No such Folder " & StartFolder
    If Right(StartFolder, 1) <> "\" Then StartFolder = StartFolder & "\"
    DoCmd.Hourglass True
'--------------------------------------------------------------
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set TopFolder = FSO.GetFolder(StartFolder)
    Call LookSubFolders(TopFolder, sRes, FileExt)
'-----------------------------------------
ExitHere:
     LISTALLFILES = sRes '!!!!!!!!!!!!!!!!!!!!!!
     Set FSO = Nothing
     Set TopFolder = Nothing
     DoCmd.Hourglass False
     Exit Function
'-------------------------------
ErrHandle:
     Beep
     ErrPrint "LISTALLFILES", Err.Number, Err.Description
     Err.Clear: Resume ExitHere
End Function
'----------------------------------------------------------------------------------------------------------------------
' Get files in subfolders
'----------------------------------------------------------------------------------------------------------------------
Private Sub LookSubFolders(ByRef Folder As Object, ByRef sRes() As String, FileExt As String)
Dim FSO As Object, objFile As Object                            '  Scripting.File
Dim objSubFolder As Object                                      '  Scripting.folder
Dim nDim As Long, sWork As String, myList As String             ' Размерность массива и рабочая переменная
Dim sExt As String                                              ' Текущее расширение

On Error GoTo ErrHandle

If FileExt <> "" Then myList = UCase(FileExt)
                 Set FSO = CreateObject("Scripting.FileSystemObject")
'------------------------------------------------------------------
    nDim = UBound(sRes)
    If nDim = 0 And sRes(0) = "" Then nDim = -1
'------------------------------------------------------------------
    For Each objFile In Folder.Files
         sWork = ""
         If FileExt = "" Then         ' Считываем все файлы
                 sWork = objFile.Path
         Else                         ' Только файлы с совпадающими расширениями
                 sExt = UCase(FSO.GetExtensionName(objFile.Path))
                 If InList(myList, sExt, ";") Then ' Расширение в списке
                    sWork = objFile.Path
                 End If
         End If
         If sWork <> "" Then
            nDim = nDim + 1: ReDim Preserve sRes(nDim)
            sRes(nDim) = sWork
         End If
    Next objFile
'------------------------------------------------------------------
    For Each objSubFolder In Folder.SubFolders
            Call LookSubFolders(objSubFolder, sRes, FileExt)
    Next objSubFolder
'------------------------------------------------------------------
ExitHere:
       Set FSO = Nothing
       Exit Sub
'-----------------------------
ErrHandle:
       Debug.Print "ERR#" & Err.Number & vbCrLf & Err.Description
       Err.Clear
       Resume ExitHere
End Sub

'=====================================================================================================================================================
' Is folder Exists // DEPRECATED
'=====================================================================================================================================================
Public Function IsFolderExists(ByVal strPath As String) As Boolean
'    On Error Resume Next
'    IsFolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
     
     IsFolderExists = IsFolder(strPath)   ' redirect to new function
End Function

'=====================================================================================================================================================
' IsFile Exists //CORRECTED
'=====================================================================================================================================================
Public Function IsFile(sPath As String) As Boolean
Dim fs As Object, bRes As Boolean

    On Error Resume Next
'---------------------
    Set fs = CreateObject("Scripting.FileSystemObject")
    bRes = fs.FileExists(sPath)
'---------------------
ExitHere:
    IsFile = bRes '!!!!!!!!!!
    Set fs = Nothing
End Function

'=====================================================================================================================================================
' IsFolder Exists //CORRECTED
'=====================================================================================================================================================
Public Function IsFolder(sPath As String) As Boolean
Dim fs As Object, bRes As Boolean
    
    On Error Resume Next
'---------------------
    Set fs = CreateObject("Scripting.FileSystemObject")
    bRes = fs.folderexists(sPath)
'---------------------
ExitHere:
    IsFolder = bRes '!!!!!!!!!!
    Set fs = Nothing
End Function

'=====================================================================================================================================================
' PROVE FILE NAME (Remove Illegal Characters)
'=====================================================================================================================================================
Public Function ProveFileName(sFile As String) As String
Dim I As Integer, nDim As Integer, sWork As String
Dim sRes As String

Const Illegalls As String = "?$<>|/\"""
On Error Resume Next
'------------------
If sFile = "" Then Exit Function


nDim = Len(Illegalls): sRes = Trim(sFile)
For I = 0 To nDim - 1
    sWork = Mid(Illegalls, I + 1, 1)
    sRes = Replace(sRes, sWork, "")
Next I
'---------------------
ExitHere:
    ProveFileName = sRes '!!!!!!!!!!!!!!!!!
End Function
'=====================================================================================================================================================
' Calculate Folder Name from File Path
'=====================================================================================================================================================
Public Function FolderNameOnly(sFullPath As String) As String
Dim FSO As Object, sWork As String
Dim sRes As String

On Error Resume Next
    If sFullPath = "" Then Exit Function
'--------------------------------------------
    sWork = IIf(Right(sFullPath, 1) = "\", Left(sFullPath, Len(sFullPath) - 1), sFullPath)
    If Dir(sWork) <> "" Then                  ' Представленный путь - файл
        Set FSO = CreateObject("Scripting.FileSystemObject") ' Создаем файловый объект
        sRes = FSO.GetParentFolderName(sFullPath)
    ElseIf Dir(sWork, vbDirectory) <> "" Then ' Представленный путь - директория
        sRes = sWork
    Else                                      ' Файл не удалось найти
        sRes = ""
    End If
    If sRes <> "" And Right(sRes, 1) <> "\" Then sRes = sRes & "\"
'-------------------------------------------
ExitHere:
    FolderNameOnly = sRes '!!!!!!!!!!!!!!!
    Set FSO = Nothing
End Function

Public Function GetFileOwner(ByVal xPath As String, ByVal xName As String) As String
Dim xFolder As Object
Dim xFolderItem As Object
Dim xShell As Object

On Error Resume Next
'-----------------------------------------
xName = StrConv(xName, vbUnicode)
xPath = StrConv(xPath, vbUnicode)
Set xShell = CreateObject("Shell.Application")
Set xFolder = xShell.NameSpace(StrConv(xPath, vbFromUnicode))
If Not xFolder Is Nothing Then
  Set xFolderItem = xFolder.ParseName(StrConv(xName, vbFromUnicode))
End If
If Not xFolderItem Is Nothing Then
  GetFileOwner = xFolder.GetDetailsOf(xFolderItem, 8)
Else
  GetFileOwner = ""
End If
Set xShell = Nothing
Set xFolder = Nothing
Set xFolderItem = Nothing
End Function

'=======================================================================================================================================
' Функция возвращает список файлов в ZIP архиве в виде массива
'=======================================================================================================================================
Public Function GetZipList(ByVal sPath As String) As String()
Dim sRes() As String, nDim As Integer, I As Integer, mDim As Integer
Dim MyZIP As New cZIP
    
Const DLM As String = ";"

On Error GoTo ErrHandle
'----------------------------------------------------------------------
    ReDim sRes(0)
If sPath = "" Then GoTo ExitHere
If Dir(sPath) = "" Then GoTo ExitHere
If UCase(FileExt(sPath)) <> "ZIP" Then GoTo ExitHere

    sRes = Split(MyZIP.LISTFILES(sPath, DLM), DLM)
'----------------------------------------------------------------------
ExitHere:
            GetZipList = sRes '!!!!!!!!!!!!!
            Set MyZIP = Nothing
            Exit Function
'--------------------------
ErrHandle:
            ErrPrint "GetZipList", Err.Number, Err.Description
            Resume ExitHere
End Function

'=======================================================================================================================================================
' READ TEXT FROM FILE
'=======================================================================================================================================================
Public Function ReadTextFile(strPath As String) As String
Dim FSO As Object, FSTR As Object
Dim OpenAsUnicode As Boolean, intAsc1Chr As Integer, intAsc2Chr As Integer

Dim sRes As String

On Error GoTo ErrHandle
'------------------------------
Set FSO = CreateObject("Scripting.FileSystemObject")
If Not FSO.FileExists(strPath) Then Err.Raise 1000, , "Can't find the file " & strPath
'------------------------------------------------------------------
' AUTODETECT UNICODE
    Set FSTR = FSO.OpenTextFile(strPath, 1, False)
    intAsc1Chr = Asc(FSTR.Read(1))
    intAsc2Chr = Asc(FSTR.Read(1))
        FSTR.Close
        
        If intAsc1Chr = 255 And intAsc2Chr = 254 Then
            OpenAsUnicode = True
        Else
            OpenAsUnicode = False
        End If
'----------------------------------------------------------------------
'GET THE CONTENT
    Set FSTR = FSO.OpenTextFile(strPath, 1, 0, OpenAsUnicode)
    sRes = FSTR.ReadAll()
    FSTR.Close
'----------------------------------------------------------------------
ExitHere:
            ReadTextFile = sRes '!!!!!!!!!!!!!
            If Not FSTR Is Nothing Then FSTR.Close
            Set FSTR = Nothing: Set FSO = Nothing
            Exit Function
'--------------------------
ErrHandle:
            ErrPrint "ReadTextFile", Err.Number, Err.Description
            Resume ExitHere
End Function

'=======================================================================================================================================================
' READ TEXT FROM FILE WITH CHARSET-CODING
'=======================================================================================================================================================
Public Function ReadTextFileUTF8(strPath As String, Optional CHARSET As String = "utf-8") As String
Dim objStream As Object, sRes As String

On Error GoTo ErrHandle
'-----------------------------
Set objStream = CreateObject("ADODB.Stream")

objStream.CHARSET = CHARSET
objStream.Open

objStream.LoadFromFile (strPath)
sRes = objStream.ReadText()
'-----------------------------
ExitHere:
     ReadTextFileUTF8 = sRes '!!!!!!!!!!!
     Set objStream = Nothing
     Exit Function
'-----------
ErrHandle:
     ErrPrint "ReadTextFileUTF8", Err.Number, Err.Description
     Err.Clear: Resume ExitHere
End Function

'=======================================================================================================================================================
' WRITE TEXT TO FILE WITH CHARSET-CODING
'=======================================================================================================================================================
Public Function WriteStringToFileUTF8(strText As String, strPath As String, Optional CHARSET As String = "utf-8") As Boolean
Dim objStream As Object, bRes As Boolean

On Error GoTo ErrHandle
'-----------------------------
Set objStream = CreateObject("ADODB.Stream")
objStream.CHARSET = CHARSET


objStream.Open
objStream.WriteText strText
objStream.SaveToFile strPath, 2

bRes = True
'-----------------------------
ExitHere:
     WriteStringToFileUTF8 = bRes '!!!!!!!!!!!!
     Set objStream = Nothing
     Exit Function
'-----------
ErrHandle:
     ErrPrint "WriteStringToFileUTF8", Err.Number, Err.Description
     Err.Clear: Resume ExitHere
End Function
'================================================================================================================================
' The finction write string to file
'================================================================================================================================
Public Function WriteStringToFile(sPath As String, sOUT As String) As Boolean
Dim FSO As Object, Fileout As Object, bRes As Boolean
    
On Error GoTo ErrHandle
'-------------------------------------------------------
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set Fileout = FSO.CreateTextFile(sPath, True, True)
        
        Fileout.Write sOUT
        Fileout.Close: bRes = True
    '---------------------------------------------
ExitHere:
        WriteStringToFile = bRes '!!!!!!!!!!
        Set Fileout = Nothing: Set FSO = Nothing
        Exit Function
    '------------------
ErrHandle:
        ErrPrint "WriteStringToFile", Err.Number, Err.Description
        Err.Clear:  Resume ExitHere
End Function

'================================================================================================================================
' Quick Store string to local file, return file name
'================================================================================================================================
Public Function StoreString(str As String) As String
Dim sFile As String
    sFile = CurrentProject.Path & "\" & FileNameOnly(GetTempFile("str", "txt"))
    If WriteStringToFile(sFile, str) Then StoreString = sFile '!!!!!!!!!!!!!
End Function

'================================================================================================================================
' File Copiing or Move with progress: iOP = 0 - COPY, iOP = 1 - MOVE
'================================================================================================================================
Public Function FileCopyMove(src As String, dest As String, Optional iOp As Integer = 0, _
                                                                                Optional NoConfirm As Boolean = False) As Boolean
          
Dim WinType_SFO As SHFILEOPSTRUCT, lflags As Long
   
#If Win64 Then
    Dim lRet As LongPtr
#Else
    Dim lRet As Long
#End If
   
On Error GoTo ErrHandle
'-------------------------------------
lflags = FOF_ALLOWUNDO
If NoConfirm Then lflags = lflags & FOF_NOCONFIRMATION
   
With WinType_SFO
       .wFunc = IIf(iOp = 0, FO_COPY, FO_MOVE)
       .pFrom = src
       .pTo = dest
       .fFlags = lflags
End With
   
   lRet = SHFileOperation(WinType_SFO)
'-----------------------------------------
ExitHere:
   FileCopyMove = (lRet = 0)
   Exit Function
'----------------
ErrHandle:
   ErrPrint "FileCopyMove", Err.Number, Err.Description
   Err.Clear
End Function

'=====================================================================================================================================================
' Get Special Folder
'=====================================================================================================================================================
Public Function GetSpecialFolder(SpecFolderName As String) As String
'Special folders are : AllUsersDesktop, AllUsersStartMenu
'AllUsersPrograms, AllUsersStartup, Desktop, Favorites
'Fonts, MyDocuments, NetHood, PrintHood, Programs, Recent
'SendTo, StartMenu, Startup, Templates
 
'Get Favorites folder and open it
    Dim WshShell As Object
    Dim SpecialPath As String

    Set WshShell = CreateObject("WScript.Shell")
    SpecialPath = WshShell.SpecialFolders("Favorites")
    MsgBox SpecialPath
    'Open folder in Explorer
    Shell "explorer.exe " & SpecialPath, vbNormalFocus
End Function
'----------------------------------------------------------------------------------------------------------------------------------------------------------
' Check sum fast generation
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function CalcMD5(strFileName As String) As String
    Dim strbuffer As String
    Dim myContext As MD5_CTX
    Dim result As String
    Dim lp As Long
    Dim MD5 As String

    strbuffer = Space(FileLen(strFileName))

    Open strFileName For Binary Access Read As #1
        Get #1, , strbuffer
    Close #1

    MD5Init myContext
    MD5Update myContext, strbuffer, Len(strbuffer)
    MD5Final myContext

    result = StrConv(myContext.digest, vbUnicode)
    
    For lp = 1 To Len(result)
            CalcMD5 = CalcMD5 & Right("00" & Hex(Asc(Mid(result, lp, 1))), 2)
    Next
    
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
                                                                                                  Optional sModName As String = "#_FILE") As String
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



'======================================================================================================================================================
' Import Binary File to DB (RS = > Set objRS = CreateObject("ADODB.Recordset"))
'======================================================================================================================================================
Public Function ImportBinaryFile(ByVal fileName As String, RS As Object, FieldName As String) As Boolean
Dim iFileNum As Integer, lFileLength As Long
Dim abBytes() As Byte, iCtr As Integer

On Error GoTo ErrHandle
'------------------------------------
If Dir(fileName) = "" Then Exit Function

'read file contents to byte array
iFileNum = FreeFile
Open fileName For Binary Access Read As #iFileNum
lFileLength = LOF(iFileNum)
ReDim abBytes(lFileLength)
Get #iFileNum, , abBytes()

'put byte array contents into db field
RS.FIELDS(FieldName).AppendChunk abBytes()
Close #iFileNum

'------------------------------------
ExitHere:
     ImportBinaryFile = True '!!!!!!!!!!!!!
     Exit Function
'---------------------
ErrHandle:
     ErrPrint2 "ImportBinaryFile", Err.Number, Err.Description, MOD_NAME
     Err.Clear
End Function

'======================================================================================================================================================
'  Export Binary File From DB to Local File System (RS = > Set objRS = CreateObject("ADODB.Recordset"))
'======================================================================================================================================================
Public Function ExportBinaryFile(fileName As String, RS As Object, FieldName As String) As Boolean
'************************************************
'SAMPLE USAGE
'Dim sConn As String
'Dim oConn As New ADODB.Connection
'Dim oRs As New ADODB.Recordset
'
'
'sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MyDb.MDB;Persist Security Info=False"
'
'oConn.Open sConn
'oRs.Open "SELECT * FROM MyTable", oConn, adOpenKeyset,
' adLockOptimistic
'LoadFileFromDB "C:\MyDocuments\MyDoc.Doc",  oRs, "MyFieldName"
'oRs.Close
'************************************************
Dim iFileNum As Integer, lFileLength As Long
Dim abBytes() As Byte, iCtr As Integer

On Error GoTo ErrHandle
'------------------------------

iFileNum = FreeFile
Open fileName For Binary As #iFileNum
lFileLength = LenB(RS(FieldName))

abBytes = RS(FieldName).GetChunk(lFileLength)
Put #iFileNum, , abBytes()
Close #iFileNum
'------------------------------
ExitHere:
    ExportBinaryFile = True '!!!!!!!!!!!!!!!!!
    Exit Function
'-----------------
ErrHandle:
    ErrPrint2 "ExportBinaryFile", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function


