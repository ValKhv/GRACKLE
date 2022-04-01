VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************************************************
'*********************************************************************************************************************************
'*********************************************************************************************************************************
' DESCRIPTION: Класс инкапсулирует элементарные файловые операции и осуществляет работу с файловой системой
'              под управлением Windows /NTFS
' VERSION    : v 0.1 16.03.2012
' DEPENDENCY :
' COPYRIGHT  : V.Khvatov / DigitalXpert
'*********************************************************************************************************************************
'*********************************************************************************************************************************
'*********************************************************************************************************************************
Option Explicit
Option Compare Database
'******************************************************************************************************

Private Type SHFILEOPSTRUCT
   hWnd As Long
   wFunc As Long
   pFrom As String
   pTo As String
   fFlags As Integer
   fAnyOperationsAborted As Long
   hNameMappings As Long
   lpszProgressTitle As String ' можно использовать только
                               '    FOF_SIMPLEPROGRESS
End Type

Private Const SW_SHOWMAXIMIZED = 3
Private Const ERROR_SUCCESS = 0&
Private Const FO_COPY = &H2
Private Const FO_DELETE = &H3
Private Const FO_MOVE = &H1
Private Const FO_RENAME = &H4
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_CONFIRMMOUSE = &H2
Private Const FOF_FILESONLY = &H80
Private Const FOF_NO_CONNECTED_ELEMENTS As Long = &H2000
Private Const FOF_MULTIDESTFILES = &H1
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_NOCONFIRMMKDIR = &H200
Private Const FOF_RENAMEONCOLLISION = &H8
Private Const FOF_SILENT = &H4
Private Const FOF_SIMPLEPROGRESS = &H100
Private Const FOF_WANTMAPPINGHANDLE = &H20

Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_NORMAL = &H80&

Private Const FOF_NOCOPYSECURITYATTRIBS As Long = &H800
Private Const FOF_NOERRORUI As Long = &H400
Private Const FOF_NORECURSION As Long = &H1000
Private Const FOF_WANTNUKEWARNING As Long = &H4000
Private Const MAX_PATH = 260
Private Const OF_READ = &H0&

Private Type FILETIME
    dwLowDate As Long
    dwHighDate As Long
End Type
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMillisecs As Integer
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const GENERIC_WRITE = &H40000000


'#############################################################################################################################
' ДЕКЛАРАЦИИ
#If Win64 Then                                               ' НОВАЯ СРЕДА
        Private Declare PtrSafe Function RemoveDirectory Lib "kernel32.dll" Alias "RemoveDirectoryA" ( _
                            ByVal lpPathName As String) As Long

        Private Declare PtrSafe Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, _
                                                                  lpFileTime As FILETIME) As Long
        Private Declare PtrSafe Function SetFileTime Lib "kernel32" (ByVal hFile As Long, ByVal MullP As Long, ByVal NullP2 As Long, _
                                                                  lpLastWriteTime As FILETIME) As Long
        Private Declare PtrSafe Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, _
                                                                  lpFileTime As FILETIME) As Long
        Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
        Private Declare PtrSafe Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, _
                                                                      lpFindFileData As WIN32_FIND_DATA) As Long
        Private Declare PtrSafe Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, _
                                                                      lpFindFileData As WIN32_FIND_DATA) As Long
        Private Declare PtrSafe Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

        Private Declare PtrSafe Function SHFileOperation Lib "shell32.dll" (ByRef lpFileOp As SHFILEOPSTRUCT) As Long
        Private Declare PtrSafe Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

        Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias _
                        "ShellExecuteA" (ByVal hWnd As Long, _
                         ByVal lpOperation As String, _
                         ByVal lpFile As String, _
                         ByVal lpParameters As String, _
                         ByVal lpDirectory As String, _
                         ByVal nShowCmd As Long) As Long

        Private Declare PtrSafe Function CreateFile Lib "kernel32" _
            Alias "CreateFileA" _
            (ByVal lpFileName As String, _
            ByVal dwDesiredAccess As Long, _
            ByVal dwShareMode As Long, _
            ByVal lpSecurityAttributes As Long, _
            ByVal dwCreationDisposition As Long, _
            ByVal dwFlagsAndAttributes As Long, _
            ByVal hTemplateFile As Long) As Long


        Private Declare PtrSafe Function lOpen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, _
                                                              ByVal iReadWrite As Long) As Long
        Private Declare PtrSafe Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
        Private Declare PtrSafe Function GetFileSize Lib "kernel32" (ByVal hFile As Long, _
                                                             lpFileSizeHigh As Long) As Long
        
        Private Declare PtrSafe Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, _
                                                                                               ByVal lpBuffer As String) As Long
        Private Declare PtrSafe Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, _
                                  ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

        Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
        Private Declare PtrSafe Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
        Private Declare PtrSafe Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" ( _
                                     ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
'----------------------------------------------------------------------------------------------------------------------------
#Else                                                       ' СТАРАЯ СРЕДА
    Private Declare Function RemoveDirectory Lib "kernel32.dll" Alias "RemoveDirectoryA" ( _
                            ByVal lpPathName As String) As Long

    Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, _
                                                                  lpFileTime As FILETIME) As Long
    Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, ByVal MullP As Long, ByVal NullP2 As Long, _
                                                                  lpLastWriteTime As FILETIME) As Long
    Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, _
                                                                  lpFileTime As FILETIME) As Long
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, _
                                                                      lpFindFileData As WIN32_FIND_DATA) As Long
    Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, _
                                                                      lpFindFileData As WIN32_FIND_DATA) As Long
    Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

    Private Declare Function SHFileOperation Lib _
                "shell32.dll" (ByRef lpFileOp As _
                    SHFILEOPSTRUCT) As Long
    Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long


    Private Declare Function ShellExecute Lib "shell32.dll" Alias _
                        "ShellExecuteA" (ByVal hwnd As Long, _
                         ByVal lpOperation As String, _
                         ByVal lpFile As String, _
                         ByVal lpParameters As String, _
                         ByVal lpDirectory As String, _
                         ByVal nShowCmd As Long) As Long

    Private Declare Function CreateFile Lib "kernel32" _
            Alias "CreateFileA" _
            (ByVal lpFileName As String, _
            ByVal dwDesiredAccess As Long, _
            ByVal dwShareMode As Long, _
            ByVal lpSecurityAttributes As Long, _
            ByVal dwCreationDisposition As Long, _
            ByVal dwFlagsAndAttributes As Long, _
            ByVal hTemplateFile As Long) As Long


    Private Declare Function lOpen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, _
                                                              ByVal iReadWrite As Long) As Long
    Private Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
    Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, _
                                                             lpFileSizeHigh As Long) As Long
                                                             
    Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, _
                                                                        ByVal lpBuffer As String) As Long
    Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, _
                                  ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
    Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
    Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
    Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal _
                        lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
#End If
'#############################################################################################################################
'



'********************************************************************
'======================================================================================================================================================
' Short File Name
'======================================================================================================================================================
Public Function GetShortFileName(sPath As String) As String
Dim lAns As Long
Dim sAns As String
Dim iLen As Integer
   
On Error Resume Next

'this function doesn't work if the file doesn't exist
If Dir(sPath) = "" Then Exit Function

sAns = Space(255)
lAns = GetShortPathName(sPath, sAns, 255)
GetShortFileName = Left(sAns, lAns) '!!!!!!!!!!

End Function
'==========================================================================================================================
' ФУНКЦИЯ КОРРЕКТНОГО УДАЛЕНИЯ ФАЙЛА В КОРЗИНУ (требуется lhwnd, например lhwnd=Application.hWndAccessApp )
'===========================================================================================================================
Public Function MoveFileToRecycle(ByVal stFilePath As String, ByVal lhWnd As Long) As Boolean
  Dim CFileStruct As SHFILEOPSTRUCT
'----------------------------------------
' Создаем структуру файла
  With CFileStruct
    .hWnd = lhWnd
    .fFlags = FOF_ALLOWUNDO
    .pFrom = stFilePath
    .wFunc = FO_DELETE
  End With

  If SHFileOperation(CFileStruct) <> ERROR_SUCCESS Then  ' Возникла ошибка
     MoveFileToRecycle = False
  Else
     MoveFileToRecycle = True
  End If
End Function
'=================================================================================================
' Функция проверяет, существует ли заданная директория на компьютере пользователя
'================================================================================================
Public Function CheckDirExisting(stPath As String) As Boolean
On Error Resume Next
 Dim strDir As String               ' Вспомогательная строка
 strDir = Dir(stPath, vbDirectory)
    If (Len(strDir) = 0 Or Err = 76) Then
        CheckDirExisting = False
    Else
        CheckDirExisting = True
    End If
End Function


'=================================================================================================
' Функция создания новой директории (в том числе многоступенчатой)
'================================================================================================
Public Function CreateDir(stPath As String) As Boolean
Dim boResult As Boolean                                           ' Булевский результат
Dim I As Long                                                     ' Циклическая переменная
Dim stlocPath As String                                           ' Локальная поддиректория

On Error GoTo CreateDirERR
'-------------------------------------------------------------------------
' Еще раз проверяем, не существует ли директория
     boResult = CheckDirExisting(stPath)
      If Not boResult Then             ' Пробуем создать директорию
         ' НАЧИНАЕМ ИТЕРАЦИЮ ПО ПУТИ, НА КАЖДОМ ЭТАПЕ СОЗДВАЯ ПОДПАПКУ
            Do
                I = InStr(I + 1, stPath & "\", "\")              ' Ищем вхождение наклонной черты
                stlocPath = Left$(stPath, I - 1)                 ' выделяем локальный кусок подпапку
                If Right$(stlocPath, 1) <> ":" And Dir$(stlocPath, vbDirectory) = "" Then
                    MkDir stlocPath                              ' СОЗДАЕМ ПОДПАПКУ
                End If
            Loop Until I >= Len(stPath)
      Else                                                        ' Нашли директорию
         MsgBox "Директория существеут"
         CreateDir = False
         Exit Function
      End If
'------------------------------------------------------------------------------------------------------------------------------------
' Успешно создали диреткорию
      CreateDir = True
      Exit Function
'-------------------------------------------------------------------------------------
' Ошибка
CreateDirERR:
  MsgBox "Ошибка CreateDir: " + CStr(Err.Number) + Err.Description
  CreateDir = False
End Function

'=================================================================================================
' Функция проверки существования файла
'================================================================================================
Public Function CheckFileExisting(ByVal stFilePath As String) As Boolean
On Error GoTo CheckFileExistingERR
        CheckFileExisting = (Dir(stFilePath) <> "")
        Exit Function
CheckFileExistingERR:
    MsgBox " Ошибка CheckFileExisting: " & CStr(Err.Number) + "; " + Err.Description
    CheckFileExisting = False
End Function
'=================================================================================================
' Функция пытается получить исключительный доступ к файлу, если не получается - возвр false
'================================================================================================
 Public Function IsFreeFile(ByVal sLongFileName As String) As Boolean
        Dim hFile As Long
        IsFreeFile = False
        hFile = CreateFile(sLongFileName, 0, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
        If hFile <> -1 Then
            IsFreeFile = True
        End If
        CloseHandle hFile
   End Function
Public Function IsFileLocked(strFileName As String) As Boolean
Dim nFreeFile As Integer
   On Error Resume Next
   ' If the file is already opened by another process,
   ' and the specified type of access is not allowed,
   ' the Open operation fails and an error occurs.
   nFreeFile = FreeFile
   Open strFileName For Binary Access Read Write Lock Read Write As #nFreeFile
   Close #nFreeFile
   ' If an error occurs, the document is currently open.
   If Err.Number <> 0 Then
      ' Display the error number and description.
      Debug.Print "Error #" & str(Err.Number) & " - " & Err.Description
      IsFileLocked = True
      Err.Clear
      Close #nFreeFile
   End If
End Function
'=================================================================================================
' Функция копирует файл
' По умолчанию, если путь отсутсвует, то он создается (bCreatePath = True)
'================================================================================================
Public Function bFileCope(stFilePathFrom As String, stFilePathTo As String, _
                                            Optional bCreatePath As Boolean = True) As Boolean
Dim sWork As String
On Error GoTo bFileCopeERR
    '------------------------------------------------------------------------------
    If bCreatePath Then                    ' Защищеное копирование (если путь не существует - он создается)
      sWork = GetFilePathPart(stFilePathTo)
      If Not Me.CheckDirExisting(sWork) Then
         CreateDir sWork
      End If
    End If
    '------------------------------------------------------------------------------
    FileCopy stFilePathFrom, stFilePathTo
    bFileCope = True
    Exit Function
bFileCopeERR:
    MsgBox "Ошибка функции bFileCope: " + CStr(Err.Number) + " ;" + Err.Description
    bFileCope = False
End Function
'==========================================================================================================================
' Функция перемещает файл на основе FSO
'==========================================================================================================================
Public Sub MoveFileFSO(sSource As String, sDestination As String, Optional bOverwrite As Boolean = True)
Dim FSO As Object
     
    On Error GoTo Oops
    Set FSO = CreateObject("Scripting.FileSystemObject")
     
'------------------------------------------------------------------------------------
' Проверяем существование файла
    If Not Me.FileExists(sSource) Then
        MsgBox "Исходный файл: " & sSource & " не существует", vbCritical
        GoTo ExitOops
    End If
'-----------------------------------------------------------------------------------
' Проверяем, существует ли
    If FileExists(sDestination) And bOverwrite Then
         ' Сначала удаляем целевой файл, затем перемещаем
        With FSO
            .DeleteFile (sDestination)
            .MoveFile sSource, sDestination
        End With
    ElseIf FileExists(sDestination) And Not bOverwrite Then
         ' Не позволено перезаписывать файл, выходим
        GoTo ExitOops
    Else ' Целевой файл не существует - перемещаем
        FSO.MoveFile sSource, sDestination
    End If
' ---------------------------------------------------------------------------------
ExitOops:
    Set FSO = Nothing
    Exit Sub
'----------------------------------------
Oops:
    MsgBox "Error # " & str(Err.Number) & " " & Err.Description
    Resume ExitOops
End Sub
'==========================================================================================================================
' Функция копирует файл на основе FSO
'==========================================================================================================================
Public Sub CopyFilesFSO(sSource As String, sDestination As String, Optional bOverwrite As Boolean = True)
Dim FSO As Object
     
    On Error GoTo Oops
    Set FSO = CreateObject("Scripting.FileSystemObject")
     
'------------------------------------------------------------------------------------
' Проверяем существование файла
    If Not Me.FileExists(sSource) Then
        MsgBox "Исходный файл: " & sSource & " не существует", vbCritical
        GoTo ExitOops
    End If
'-----------------------------------------------------------------------------------
' Проверяем, существует ли
    If FileExists(sDestination) And bOverwrite Then
         ' Сначала удаляем целевой файл, затем перемещаем
        With FSO
            .DeleteFile (sDestination)
            .CopyFile Source:=sSource, Destination:=sDestination
        End With
    ElseIf FileExists(sDestination) And Not bOverwrite Then
         ' Не позволено перезаписывать файл, выходим
        GoTo ExitOops
    Else ' Целевой файл не существует - перемещаем
        FSO.CopyFile Source:=sSource, Destination:=sDestination
    End If
' ---------------------------------------------------------------------------------
ExitOops:
    Set FSO = Nothing
    Exit Sub
'----------------------------------------
Oops:
    MsgBox "Error # " & str(Err.Number) & " " & Err.Description
    Resume ExitOops
End Sub
'==========================================================================================================================
' Функция копирует целый каталог, включая все покаталоги на основе FSO
'==========================================================================================================================
Public Function CopyFolder(FromPath As String, ToPath As String) As Boolean
    Dim FSO As Object
    Dim bRes As Boolean

On Error GoTo CopyFolder_ERR
'----------------------------------------------------------------------------------------
    If Right(FromPath, 1) = "\" Then
        FromPath = Left(FromPath, Len(FromPath) - 1)
    End If

    If Right(ToPath, 1) = "\" Then
        ToPath = Left(ToPath, Len(ToPath) - 1)
    End If
'----------------------------------------------------------------------------------------
    Set FSO = CreateObject("scripting.filesystemobject")
'----------------------------------------------------------------------------------------
    If FSO.folderexists(FromPath) = False Then
        MsgBox "ERR: " & FromPath & " не найден", vbCritical, "CopyFolder ERROR"
        Exit Function
    End If
'----------------------------------------------------------------------------------------
    DoCmd.Hourglass True
        FSO.CopyFolder Source:=FromPath, Destination:=ToPath
    DoCmd.Hourglass False
        bRes = True
        MsgBox "Каталог скопирован " & FromPath & " в " & ToPath
'----------------------------------------------------------------------------------------
CopyFolder_EXIT:
      CopyFolder = bRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
      Exit Function
'--------------------
CopyFolder_ERR:
      MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "CopyFolder ERROR"
      Resume CopyFolder_EXIT
End Function
'==========================================================================================================================
' Функция перемещает целый каталог, включая все покаталоги на основе FSO
'==========================================================================================================================
Public Function MoveFolder(FromPath As String, ToPath As String) As Boolean
    Dim FSO As Object
    Dim bRes As Boolean
    
On Error GoTo MoveFolder_ERR
'----------------------------------------------------------------------------------
    If Right(FromPath, 1) = "\" Then
        FromPath = Left(FromPath, Len(FromPath) - 1)
    End If

    If Right(ToPath, 1) = "\" Then
        ToPath = Left(ToPath, Len(ToPath) - 1)
    End If
    '----------------------------------------------------
    Set FSO = CreateObject("scripting.filesystemobject")
    '----------------------------------------------------
    If FSO.folderexists(FromPath) = False Then
        MsgBox FromPath & " не найден"
        Exit Function
    End If

    If FSO.folderexists(ToPath) = True Then
        MsgBox ToPath & " существует, не возможно переместить в имеющийся каталог"
        Exit Function
    End If
'----------------------------------------------------------------------------------------

    FSO.MoveFolder Source:=FromPath, Destination:=ToPath
    MsgBox "The folder is moved from " & FromPath & " to " & ToPath
'----------------------------------------------------------------------------------------
    DoCmd.Hourglass True
        FSO.MoveFolder Source:=FromPath, Destination:=ToPath
    DoCmd.Hourglass False
        bRes = True
        MsgBox "Каталог " & FromPath & " перемещен в " & ToPath
'----------------------------------------------------------------------------------------
MoveFolder_EXIT:
      MoveFolder = bRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
      Exit Function
'--------------------
MoveFolder_ERR:
      MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "MoveFolder ERROR"
      Resume MoveFolder_EXIT
End Function
'=================================================================================================
' Функция открывает файл ассоциированным приложением
'================================================================================================
Public Function bOpenFileShell(stFilePath As String) As Boolean
'On Error GoTo bOpenFileShellERR
On Error Resume Next
    Call ShellExecute(0, "open", stFilePath, "", "", SW_SHOWMAXIMIZED)
    bOpenFileShell = True
    Exit Function
bOpenFileShellERR:
    MsgBox "Ошибка функции bOpenFileShell: " + CStr(Err.Number) + " ;" + Err.Description
    bOpenFileShell = False
End Function

'=================================================================================================
' Функция универсальной проверки директории
'================================================================================================
Public Function PathExistX(sPath As String, Optional bDir As Boolean = False, _
                Optional lOUT_Error As Long = 0, Optional bShowError As Boolean = True) As Boolean

On Error GoTo errPE
    PathExistX = False

    If Trim$(sPath) = "" Then Exit Function

    PathExistX = IIf(Len(Dir(sPath, IIf(bDir, vbDirectory, vbNormal))) > 0, True, False)

    Exit Function
errPE:
    PathExistX = False
    lOUT_Error = Err.Number

    If bShowError Then
        MsgBox "Error: " & Error, vbCritical, "Error " & Err
    End If

    Err.Clear
End Function

'=================================================================================================
' Функция расширенного копирования файла
'================================================================================================
Public Function CopyX(hWind As Long, sSource As String, sTarget As String, _
                 Optional bConfirmOverwrite As Boolean = False, Optional bShowProgress As Boolean = True, _
                 Optional lOUT_Abort As Long = 0, Optional lOUT_Error As Long = 0) As Long

On Error GoTo errCopy
    Dim SHFO As SHFILEOPSTRUCT

    With SHFO
        .hWnd = hWind

        If bConfirmOverwrite And bShowProgress Then
            .fFlags = FOF_NOCONFIRMATION Or FOF_SIMPLEPROGRESS
        ElseIf bConfirmOverwrite And Not bShowProgress Then
            .fFlags = FOF_NOCONFIRMATION
        ElseIf Not bConfirmOverwrite And bShowProgress Then
            .fFlags = FOF_SIMPLEPROGRESS
        Else
            .fFlags = FOF_SILENT
        End If

        .pFrom = sSource
        .pTo = sTarget

        .wFunc = FO_COPY
    End With

    CopyX = SHFileOperation(SHFO)

    lOUT_Abort = SHFO.fAnyOperationsAborted
    lOUT_Error = 0

    Exit Function
errCopy:
    lOUT_Abort = 1
    lOUT_Error = Err.Number

    Err.Clear

    CopyX = 0
End Function

'=================================================================================================
' Функция перемещения файла с дополнительными параметрами
'================================================================================================
Public Function MoveX(hWind As Long, sSource As String, sTarget As String, _
                Optional bConfirmOverwrite As Boolean = False, Optional bShowProgress As Boolean = True, _
                Optional lOUT_Abort As Long = 0, Optional lOUT_Error As Long = 0) As Long

On Error GoTo errMoveX
    Dim SHFO As SHFILEOPSTRUCT

    With SHFO
        .hWnd = hWind

        If bConfirmOverwrite And bShowProgress Then
            .fFlags = FOF_NOCONFIRMATION Or FOF_SIMPLEPROGRESS
        ElseIf bConfirmOverwrite And Not bShowProgress Then
            .fFlags = FOF_NOCONFIRMATION
        ElseIf Not bConfirmOverwrite And bShowProgress Then
            .fFlags = FOF_SIMPLEPROGRESS
        Else
            .fFlags = FOF_SILENT
        End If

        .pFrom = sSource
        .pTo = sTarget

        .wFunc = FO_MOVE
    End With

    MoveX = SHFileOperation(SHFO)

    lOUT_Abort = SHFO.fAnyOperationsAborted
    lOUT_Error = 0

    Exit Function
errMoveX:
    lOUT_Abort = 1
    lOUT_Error = Err.Number

    Err.Clear

    MoveX = 0
End Function
'=================================================================================================
' Функция удаления файла с дополнительными параметрами
'================================================================================================
Public Function DeleteX(hWind As Long, sSource As String, Optional bToRecycleBin As Boolean = True, _
               Optional bConfirm As Boolean = True, Optional bShowProgress As Boolean = True, _
               Optional lOUT_Abort As Long = 0, Optional lOUT_Error As Long = 0) As Long

On Error GoTo errDelete

    Dim SHFO As SHFILEOPSTRUCT

    With SHFO
        .hWnd = hWind

        If bToRecycleBin And bShowProgress And bConfirm Then
            .fFlags = FOF_ALLOWUNDO Or FOF_SIMPLEPROGRESS
        ElseIf bToRecycleBin And bShowProgress And Not bConfirm Then
            .fFlags = FOF_ALLOWUNDO Or FOF_SIMPLEPROGRESS Or FOF_NOCONFIRMATION
        ElseIf bToRecycleBin And Not bShowProgress And bConfirm Then
            .fFlags = FOF_ALLOWUNDO
        ElseIf bToRecycleBin And Not bShowProgress And Not bConfirm Then
            .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION
        ElseIf Not bToRecycleBin And bShowProgress And bConfirm Then
            .fFlags = FOF_SIMPLEPROGRESS
        ElseIf Not bToRecycleBin And bShowProgress And Not bConfirm Then
            .fFlags = FOF_SIMPLEPROGRESS Or FOF_NOCONFIRMATION
        ElseIf Not bToRecycleBin And Not bShowProgress And bConfirm Then
            .fFlags = FOF_SILENT
        ElseIf Not (bToRecycleBin Or bShowProgress Or bConfirm) Then
            .fFlags = FOF_NOCONFIRMATION
        Else
            .fFlags = FOF_SILENT
        End If

        .pFrom = sSource

        .wFunc = FO_DELETE
    End With

    DeleteX = SHFileOperation(SHFO)

    lOUT_Abort = SHFO.fAnyOperationsAborted
    lOUT_Error = 0

    Exit Function
errDelete:
    lOUT_Abort = 1
    lOUT_Error = Err.Number

    Err.Clear

    DeleteX = 0
End Function

'=========================================================================================================================
' Функция удаляет все файлы
'=========================================================================================================================
Public Sub DeleteAllFiles(sFolder As String)
Dim MyPath   As String                   ' Путь к файлу
Dim FSO As Object, fldr As Object        ' Файловый манипулятор и каталог
Dim fileName As Object                   ' Указатель на файл

Set FSO = CreateObject("Scripting.FileSystemObject") ' Создаем файловый объект
Set fldr = FSO.GetFolder(sFolder)

For Each fileName In fldr.Files
    fileName.Delete True ' delete all files
Next
'----------------------------------------------
DeleteAllFiles_EXIT:
    Exit Sub
'--------------
DeleteAllFiles_ERROR:
   MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "DeleteAllFiles ERROR"
   Resume DeleteAllFiles_EXIT
End Sub

'=================================================================================================
' Функция принудительной установки даты модификации (даты файла)
'  FileName - полный путь
'  TheDate - дата последней модицикации, например, "21/01/1998"
' ВЫЗОВ: a = SetFileDateTime("c:\1.txt", Now) или  Call SetFileDateTime("c:\1.txt", "21/01/1998")
' Возвращает true в случае успеха
'================================================================================================
Public Function SetFileDateTime(ByVal fileName As String, ByVal TheDate As String) As Boolean
    If Dir(fileName) = "" Then Exit Function
    If Not IsDate(TheDate) Then Exit Function
Dim lFileHnd As Long
Dim lRet As Long
Dim typFileTime As FILETIME
Dim typLocalTime As FILETIME
Dim typSystemTime As SYSTEMTIME

With typSystemTime
    .wYear = Year(TheDate)
    .wMonth = Month(TheDate)
    .wDay = Day(TheDate)
    .wDayOfWeek = Weekday(TheDate) - 1
    .wHour = Hour(TheDate)
    .wMinute = Minute(TheDate)
    .wSecond = Second(TheDate)
End With
lRet = SystemTimeToFileTime(typSystemTime, typLocalTime)
lRet = LocalFileTimeToFileTime(typLocalTime, typFileTime)
lFileHnd = CreateFile(fileName, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
lRet = SetFileTime(lFileHnd, ByVal 0&, ByVal 0&, typFileTime)
CloseHandle lFileHnd
    SetFileDateTime = lRet > 0
End Function

'=================================================================================================
' Функция удаления каталога
'================================================================================================
Public Function DeleteDir(sPath As String) As Boolean
Dim RetVal As Long

    RetVal = RemoveDirectory(EndSlash(sPath))
    If RetVal = 1 Then
        DeleteDir = True
    Else
        DeleteDir = False
    End If
End Function

'=================================================================================================
' Функция удаления каталога вместе со всем содержимым
'================================================================================================
Public Function DelTree(ByVal strDir As String) As Long
Dim x As Long
Dim intAttr As Integer
Dim strAllDirs As String
Dim strFile As String

        DelTree = -1
On Error Resume Next
strDir = Trim$(strDir)
If Len(strDir) = 0 Then Exit Function
If Right$(strDir, 1) = "\" Then strDir = Left$(strDir, Len(strDir) - 1)
If InStr(strDir, "\") = 0 Then Exit Function
intAttr = GetAttr(strDir)
If (intAttr And vbDirectory) = 0 Then Exit Function
strFile = Dir$(strDir & "\*.*", vbSystem Or vbDirectory Or vbHidden)
Do While Len(strFile)
If strFile <> "." And strFile <> ".." Then
intAttr = GetAttr(strDir & "\" & strFile)
If (intAttr And vbDirectory) Then
strAllDirs = strAllDirs & strFile & Chr$(0)
Else
If intAttr <> vbNormal Then
SetAttr strDir & "\" & strFile, vbNormal
If Err Then DelTree = Err: Exit Function
End If
Kill strDir & "\" & strFile
If Err Then DelTree = Err: Exit Function
End If
End If
strFile = Dir$
Loop
Do While Len(strAllDirs)
    x = InStr(strAllDirs, Chr$(0))
    strFile = Left$(strAllDirs, x - 1)
    strAllDirs = Mid$(strAllDirs, x + 1)
    x = DelTree(strDir & "\" & strFile)
    If x Then DelTree = x: Exit Function
Loop

RmDir strDir
If Err Then
    DelTree = Err
Else
    DelTree = 0
End If
End Function

'=================================================================================================
' Функция получения размера всего каталога
' Возвращает размер в байтах
'================================================================================================
Public Function SizeOfDir(ByVal DirPath As String) As Double
Dim hFind As Long
Dim fdata As WIN32_FIND_DATA
Dim dblSize As Double
Dim sName As String
Dim x As Long
On Error Resume Next
x = GetAttr(DirPath)
If Err Then SizeOfDir = 0: Exit Function

If (x And vbDirectory) = vbDirectory Then
    dblSize = 0
    Err.Clear
    sName = Dir$(EndSlash(DirPath) & "*.*", vbSystem Or vbHidden Or vbDirectory)
        If Err.Number = 0 Then
            hFind = FindFirstFile(EndSlash(DirPath) & "*.*", fdata)
            If hFind = 0 Then Exit Function
            Do
                If (fdata.dwFileAttributes And vbDirectory) = vbDirectory Then
                    sName = Left$(fdata.cFileName, InStr(fdata.cFileName, vbNullChar) - 1)
                    If sName <> "." And sName <> ".." Then
                        dblSize = dblSize + SizeOfDir(EndSlash(DirPath) & sName)
                    End If
                Else
                    dblSize = dblSize + fdata.nFileSizeHigh * 65536 + fdata.nFileSizeLow
                End If
                DoEvents
            Loop While FindNextFile(hFind, fdata) <> 0
            hFind = FindClose(hFind)
        End If
Else
    On Error Resume Next
    dblSize = FileLen(DirPath)
End If
    SizeOfDir = dblSize  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function

'=================================================================================================
' Функция проверяет наличие слэша в конце
'================================================================================================
Private Function EndSlash(ByVal PathIn As String) As String
    If Right$(PathIn, 1) = "\" Then
        EndSlash = PathIn
    Else
        EndSlash = PathIn & "\"
    End If
End Function
'=================================================================================================
' Функция сравнивает два файла на идентичность,
' проверки существования файлов не осуществляется
'================================================================================================
Public Function Compare2Files(sFile1 As String, sFile2 As String) As Boolean
Dim nF1 As Integer, nF2 As Integer                  ' номера открытых файлов
Dim issame As Boolean                               ' Флаг сравнения
Dim Start As Long, whole As Long, part As Long      ' Текущие указатели
Dim x As Long
Dim buffer1 As String, buffer2 As String            ' Буферы
On Error GoTo Compare2Files_ERR

    nF1 = FreeFile: Open sFile1 For Binary As #nF1
    nF2 = FreeFile: Open sFile2 For Binary As #nF2
' Инициализируем
        issame = True
' Начинаем сравнение
    If LOF(nF1) <> LOF(nF2) Then
            issame = False
    Else
        whole = LOF(nF1) \ 10000
        part = LOF(nF1) Mod 10000
        buffer1 = String$(10000, 0)
        buffer2 = String$(10000, 0)
        Start = 1
            For x = 1 To whole
                Get #nF1, Start, buffer1
                Get #nF2, Start, buffer2
                If buffer1 <> buffer2 Then
                    issame = False
                    Exit For
                End If
                    Start = Start + 10000
            Next
        buffer1 = String$(part, 0)
        buffer2 = String$(part, 0)
        Get #nF1, Start, buffer1
        Get #nF2, Start, buffer2
            If buffer1 <> buffer2 Then issame = False
    End If
Close

' Выводим результаты
  Compare2Files = issame '!!!!!!!!!!!!!!!!!!!!!!!!!!!
  Exit Function
  
' Ловушка ошибок
Compare2Files_ERR:
    MsgBox "ERR#" & Err.Number & vbCrLf & "ERR DESCRIPTION: " & Err.Description, vbCritical
End Function
'=================================================================================================
' Функция считывает размер файла на диске
'================================================================================================
Public Function GetFileSizeX(FilePath As String) As Double
    Dim lpFSHigh As Long
    Dim Pointer As Long, sizeofFile As Long
    Pointer = lOpen(FilePath, OF_READ)
    'size of the file
    sizeofFile = GetFileSize(Pointer, lpFSHigh)
    GetFileSizeX = sizeofFile
    lclose Pointer
End Function

'==============================================================================================================================
' Функция добавляет номера строк к файлу (если он представим строками)
' Parameters: strSource - file to read
'             strDestination - file to write to
'             intLineStart - value to start numbering at
'             intLineIncrement - value to increment line numbers by
' Returns   : nothing
'==============================================================================================================================
Public Sub AddLineNumbersToFile(strSource As String, strDestination As String, intLineStart As Integer, _
                                                                                                intLineIncrement As Integer)
  
  Dim lngSourceFile As Long
  Dim lngDestinationFile As Long
  Dim strOneLine As String
  Dim intCounter As Integer
 
  On Error GoTo PROC_ERR
  
  ' Start the counter at the appropriate place
  intCounter = intLineStart
  
  ' Open the source file
  lngSourceFile = FreeFile
  Open strSource For Input Access Read As lngSourceFile
  
  ' Open the destination file
  lngDestinationFile = FreeFile
  Open strDestination For Output Access Write As lngDestinationFile
 
  ' Scan the source file, adding line numbers as necessary
  Do Until EOF(lngSourceFile)
    intCounter = intCounter + intLineIncrement
    Line Input #lngSourceFile, strOneLine
    Print #lngDestinationFile, intCounter & ":", strOneLine
  Loop
  
  ' Close both files
  Close lngSourceFile
  Close lngDestinationFile
    
PROC_EXIT:
  Exit Sub
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "AddLineNumbersToFile"
  Resume PROC_EXIT
   
End Sub
'==============================================================================================================================
' Функция добавляет строку к файлу
' Parameters: strFile - Path and name of the file
'             strText - Text to append
' Returns   : True if successful
'==============================================================================================================================
Public Function AppendToTextFile( _
  strFile As String, _
  strText As String) _
  As Boolean
  Dim intFile As Integer

  On Error GoTo PROC_ERR

  ' Get a free file handle
  intFile = FreeFile

  ' Open the file in append mode, write to it, and close it
  Open strFile For Append As intFile
  Print #intFile, strText
  Close #intFile

  AppendToTextFile = True

PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "AppendToTextFile"
  Resume PROC_EXIT

End Function
'==============================================================================================================================
' Копирование директорий (не поддерживает рекурсий в под-директории)
' Parameters: strSource - Source directory
'             strDestination - Destination directory
' Returns   : Number of files copied
'==============================================================================================================================
Public Function CopyDirectory(strSource As String, strDestination As String) As Long
  Dim strFileName As String
  Dim intCount As Integer
  Dim fOK As Boolean

  On Error GoTo PROC_ERR
  
  ' Add wildcards to the name
  strFileName = Dir(strSource & "\*.*")
           
  ' Loop through the directory entries
  Do While strFileName <> ""
    ' Call the CopyFile routine
    fOK = CopyFile(strSource & "\" & strFileName, strDestination & _
      "\" & strFileName)
    
    ' Check to see that the call worked
    If Not fOK Then
      ' Copy failed, so exit the loop
      CopyDirectory = False
      Exit Do
    End If

    ' Increment the number of files copied
    intCount = intCount + 1
    strFileName = Dir
  Loop
  
  ' Return the count
  CopyDirectory = intCount

PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "CopyDirectory"
  Resume PROC_EXIT
  
End Function
'==============================================================================================================================
' Копирование файлов
' Parameters: strSource - Source file
'             strDestination - Destination file
' Returns   : True if successful, False otherwise
'==============================================================================================================================
Public Function CopyFile(strSource As String, strDestination As String) As Boolean
  Const cintBufferSize = 4096
  Dim strbuffer As String * cintBufferSize
  Dim strTempBuffer As String
  Dim intSourceFile As Integer
  Dim intDestinationFile As Integer
  Dim lngCounter As Long

  On Error GoTo PROC_ERR
  
  ' Open the input file, use readonly mode to avoid placing locks.
  intSourceFile = FreeFile
  Open strSource For Binary Access Read As #intSourceFile
  
  ' Open the output file
  intDestinationFile = FreeFile
  Open strDestination For Binary As #intDestinationFile
  
  ' Go through the file, a buffer at a time
  For lngCounter = 1 To LOF(intSourceFile) \ cintBufferSize
    Get #intSourceFile, , strbuffer
    Put #intDestinationFile, , strbuffer
  Next lngCounter
  
  ' Check for buffer overruns
  lngCounter = LOF(intSourceFile) Mod cintBufferSize
  
  ' Do the remainder
  If lngCounter > 0 Then
    Get #intSourceFile, , strbuffer
    strTempBuffer = Left$(strbuffer, lngCounter)
    Put #intDestinationFile, , strTempBuffer
  End If

  ' Close and clean up
  Close #intSourceFile
  Close #intDestinationFile
  CopyFile = True
  
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "CopyFile"
  Resume PROC_EXIT
  
End Function

Public Function DirExists(strDir As String) As Boolean
  ' Comments  : Determines if the named directory exists
  ' Parameters: strDir - Directory to check
  ' Returns   : True if the directory exists, False otherwise
  ' Source    : Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR

  DirExists = Len(Dir$(strDir & "\.", vbDirectory)) > 0
  
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "DirExists"
  Resume PROC_EXIT
  
End Function
'========================================================================================================================
' Функция проверяет наличие файла (включая скрытые, системные и только для чтения)
' PARAMS:  strFile - путь к файлу
'          bFindFolders - если strFile - путь к каталогу, то возвращается False, пока bFindFolders не устанавливается в True
'========================================================================================================================
Public Function FileExists(ByVal strFile As String, Optional bFindFolders As Boolean) As Boolean
    Dim lngAttributes As Long

    'Включаем read-only files, hidden files, system files
    lngAttributes = (vbReadOnly Or vbHidden Or vbSystem)

    If bFindFolders Then
        lngAttributes = (lngAttributes Or vbDirectory) 'Include folders as well.
    Else
        'Убираем лишний слэш, если передали путь к каталогу
        Do While Right$(strFile, 1) = "\"
            strFile = Left$(strFile, Len(strFile) - 1)
        Loop
    End If

    'If Dir() returns something, the file exists.
    On Error Resume Next
    FileExists = (Len(Dir(strFile, lngAttributes)) > 0)  '!!!!!!!!!!!!!!!!!!!!
End Function
'========================================================================================================================
' Файл считывается в строковый массив
'========================================================================================================================
Public Function FilesToArray( _
  strDirectory As String, _
  fIncludeHidden As Boolean, _
  fIncludeSystem As Boolean, _
  astrIn() As String) _
  As Integer
  ' Comments  : Populates the passed array with a list of files in the
  '             specified directory
  ' Parameters: strDirectory - Directory to read
  '             fIncludeHidden - True to include hidden files
  '             fIncludeSystem - True to include system files
  '             astrIn() - Array to hold file names (0-based)
  ' Returns   : Number of files in the directory
  ' Source    : Total VB SourceBook 6
  '
  Dim intCounter As Integer
  Dim StrTmp As String
  Dim lngAttr As Long
  
  On Error GoTo PROC_ERR

  ' Build up the options flag
  lngAttr = VBA.vbNormal
  
  If fIncludeHidden Then
    lngAttr = lngAttr + VBA.vbHidden
  End If
  
  If fIncludeSystem Then
    lngAttr = lngAttr + VBA.vbSystem
  End If
  
  ' Make sure there is a trailing slash
  If Right$(strDirectory, 1) <> "\" Then
    strDirectory = strDirectory & "\"
  End If
  
  ' Get the first entry
  StrTmp = Dir(strDirectory, lngAttr)
  
  ' Loop through each entry
  Do Until StrTmp = ""

    ' Are there entries left?
    If StrTmp <> "" Then
      ' Grow the array and add the item
      ReDim Preserve astrIn(intCounter)
      astrIn(intCounter) = StrTmp
      intCounter = intCounter + 1
    End If
    StrTmp = Dir
  Loop

  ' Return the count
  FilesToArray = intCounter

PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "FilesToArray"
  Resume PROC_EXIT

End Function
'======================================================================================================================================================
' Функция получает дерево поддиректориев и файлов для заданной директории. Фозвращается стандартное дерево.
'======================================================================================================================================================
Public Function FolderToTree(SourceFolderName As String, Optional IncludeSubFolders As Boolean = True, Optional lParent As Long = -1, _
                                                        Optional bGetProperties As Boolean = True, Optional iLevel As Long = 0) As String
Dim FSO As Object                                                                ' Используется объект FSO
Dim SourceFolder As Object, SubFolder As Object                                  ' Рабочие объекты
Dim FileItem As Object                                                           ' Рабочий указатель на файл
Dim iParent As Long, sProp As String, iCount As Long                             ' Циклическая переменная
Dim MyTree As New cTree, SubTree As New cTree                                    ' Структура - типа дерева
Dim sRes As String, sWork As String                                              ' Возвращаемый результат
    
On Error GoTo FolderToTree_ERR
'-----------------------------------------------------------------------------------------------
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set SourceFolder = FSO.GetFolder(SourceFolderName)
'-----------------------------------------------------------------------------------------------
    iParent = MyTree.Add("FOLDER", SourceFolderName, GetFolderName(SourceFolderName), lParent, , "Level=" & _
                         CStr(iLevel) & ";Files=" & SourceFolder.Files.Count)
'------------------------------------------------------------------------------------------------
    
    '--------------------------------------------------------------------------------------------
    For Each FileItem In SourceFolder.Files
        If bGetProperties Then                      ' Если флаг bGetProperties - выводим свойства
            sProp = "FullPath=" & FileItem.Path & FileItem.Name & ";Size=" & FileItem.SIZE & _
                    ";Type=" & FileItem.Type & ";DataCreated=" & FileItem.DateCreated & ";DataAccessed=" & FileItem.DateLastAccessed & _
                    ";DataModified=" & FileItem.DateLastModified & ";Attributes=" & FileItem.Attributes
        Else
            sProp = "FullPath=" & FileItem.Path & FileItem.Name
        End If
        '-------------------------------------------------
        MyTree.Add , sProp, FileItem.Name, iParent   ' Добавляем файл к дереву
    Next FileItem
    '--------------------------------------------------------------------------------------------
    If IncludeSubFolders Then                        ' Проходим все подкаталоги
        For Each SubFolder In SourceFolder.SubFolders
            sWork = FolderToTree(SubFolder.Path, True, iParent, bGetProperties, "Level=" & _
                         CStr(iLevel + 1) & ";Files=" & SourceFolder.Files.Count)
            iCount = SubTree.StringToTree(sWork)
            iCount = MyTree.InsertSubTree(iParent, SubTree)
        Next SubFolder
    End If
'---------------------------------------------------------------------------------------------
FolderToTree_EXIT:
    sRes = MyTree.TreeToString()
    Set MyTree = Nothing
    Set FileItem = Nothing
    Set SourceFolder = Nothing
    Set FSO = Nothing
    FolderToTree = sRes  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'---------------------------------------------
FolderToTree_ERR:
    MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "Ошибка при получении дерева файлов"
    Err.Clear
    Resume FolderToTree_EXIT
End Function
'=======================================================================================================================================================
' Получает файл по его пути и возвращает содержимое в переменной VARIANT
'=======================================================================================================================================================
Public Function GetFile(strFile As String) As Variant
  ' Comments  : Opens the named file and returns its contents as a variant
  ' Parameters: strFile - Path and name of file to open
  ' Returns   : Variant holding file contents or <error> if an error occurs
  ' Source    : Total VB SourceBook 6
  '
  Dim intFile As Integer
  Dim varTmp As Variant
  Dim varOut As Variant

  On Error GoTo PROC_ERR
  
  ' Get a handle an open the file
  intFile = FreeFile
  Open strFile For Input As intFile
  
  ' Walk through contents
  Do Until EOF(intFile)
    ' Get the line and put it into the variant
    Input #intFile, varTmp
    varOut = varOut & varTmp
  Loop

  ' Close the file and return the contents
  Close #intFile
  GetFile = varOut
  
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetFile"
  Resume PROC_EXIT

End Function
'===============================================================================================================================
' Возвращает имя файла с его расширением из полного пути
' Parameters: strIn - path and name to parse
' Returns   : file name + extension
'===============================================================================================================================
Public Function GetFileNamePart(strIn As String) As String
  Dim intCounter As Integer
  Dim StrTmp As String

  On Error GoTo PROC_ERR
  
  ' Parse the string
  For intCounter = Len(strIn) To 1 Step -1
    ' It its a slash, grab the sub string
    If Mid$(strIn, intCounter, 1) <> "\" Then
      StrTmp = Mid$(strIn, intCounter, 1) & StrTmp
    Else
      Exit For
    End If
  Next intCounter

  ' Return the value
  GetFileNamePart = StrTmp  '!!!!!!!!!!!!!!!!!!!!!!!!!
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetNamePart"
  Resume PROC_EXIT

End Function
'===============================================================================================================================
' Функция получает имя каталога из его пути
'===============================================================================================================================
Public Function GetFolderName(FolderPath As String)
Dim sRes As String, iL As Long
    sRes = IIf(Right(FolderPath, 1) = "\", Left(FolderPath, Len(FolderPath) - 1), FolderPath)
    '---------------------------------------------------------------------------------------------------
    iL = InStrRev(sRes, "\")
    sRes = Right(sRes, Len(sRes) - iL)
'-------------------------------------------------------------------------------------------------------
    GetFolderName = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'===============================================================================================================================
' Функция из полного имени возвращает только путь к файлу
' Parameters: strPath - string to parse
' Returns   : path part
'===============================================================================================================================
Public Function GetFilePathPart(strPath As String) As String
  Dim intCounter As Integer

  On Error GoTo PROC_ERR
  
  ' Parse the string backwards
  For intCounter = Len(strPath) To 1 Step -1
    ' Short-circuit when we reach the slash
    If Mid$(strPath, intCounter, 1) = "\" Then
      Exit For
    End If
  Next intCounter

  ' Return the value
  GetFilePathPart = Left$(strPath, intCounter)

PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetPathPart"
  Resume PROC_EXIT
  
End Function

Public Function IsAttrSet( _
  strFileName As String, _
  lngAttr As Long) _
  As Boolean
  ' Comments  : Determines if the specified attribute is set on the
  '             specified file.
  ' Parameters: strFileName - fully qualified path and name of the file
  '             to check.
  '             lngAttr - one of vbNormal, vbReadOnly, vbHidden, vbSystem
  ' Returns   : True if the specified attribute is set
  ' Source    : Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR
  
  If (VBA.GetAttr(strFileName) And lngAttr) = lngAttr Then
    IsAttrSet = True
  Else
    IsAttrSet = False
  End If
  
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "IsAttrSet"
  Resume PROC_EXIT
  
End Function

Public Function KillFile(strFile As String) As Boolean
  ' Comments  : Deletes the named file, handling errors if the file does
  '             not exist
  ' Parameters: strFile - path and name of the file to delete
  ' Returns   : True if the file was deleted, False otherwise
  ' Source    : Total VB SourceBook
  '
  
  ' Disable error handling in case the file doesn't exist
  On Error Resume Next
  Kill strFile
  On Error GoTo PROC_ERR

  KillFile = True

PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "KillFile"
  Resume PROC_EXIT

End Function
'========================================================================================================================
' Открытие заданной папки (Explorer)
'========================================================================================================================
Public Sub OpenExplorerWindow(sFolderPath As String)
  If DirExists(sFolderPath) Then
        ShellExecute 0, "open", sFolderPath, 0, 0, 1
  End If
End Sub
'========================================================================================================================
' Открытие файла
'========================================================================================================================
Public Function MoveFile(strFile As String, strDestination As String) As Boolean
  ' Comments  : Moves the specified file to a new destination
  ' Parameters: strFile - fully qualified path and name of the file to move
  '             strDestination - path of the directory to move to
  Dim bRes As Boolean
  Dim strFileOnly As String
  
  On Error GoTo PROC_ERR
  
  ' Get the source file name with no path
  strFileOnly = FileNameOnly(strFile)
  
  ' Use name to move it. (Name with the same file name moves the file)
  Name strFile As strDestination & "\" & strFileOnly
  bRes = True
PROC_EXIT:
  MoveFile = bRes '!!!!!!!!!!!!!!!!!!!!!!!
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "MoveFile"
  Resume PROC_EXIT

End Function

Public Function ParsePath( _
  strIn As String, _
  strDrive As String, _
  strPath As String, _
  strFileName As String, _
  strExtension As String) _
  As Boolean
  ' Comments  : Parses the passed full path into its component parts
  ' Parameters: strIn - Fully qualified path to parse
  ' Sets      : strDrive - Set to the drive letter
  '             strPath - Set to the path
  '             strFileName - Set to the file name (without the extension)
  '             strExtension - Set to the file's extension
  ' Returns   : True if successful, False otherwise
  ' Source    : Total VB SourceBook 6
  '
  Dim StrTmp As String
  Dim intPathLen As Integer
  Dim intPos As Integer
  Dim intTmpLen As Integer
  Dim fFileNameFound As Integer

  On Error GoTo PROC_ERR

  StrTmp = strIn

  ' Look for the colon to find the drive
  If Mid$(StrTmp, 2, 1) = ":" Then
    strDrive = Left$(StrTmp, 2)
    StrTmp = Mid$(StrTmp, 3)
  End If

  intPathLen = Len(StrTmp)

  ' Walk through the rest of the string
  For intPos = intPathLen To 1 Step -1

    Select Case Mid$(StrTmp, intPos, 1)
    
      Case "."
        ' The file's extension
        intTmpLen = Len(StrTmp) - intPos
        If intTmpLen >= 1 Then
          strExtension = Mid$(StrTmp, intPos, intTmpLen + 1)
        End If
        StrTmp = Left$(StrTmp, intPos - 1)

      Case "\"
        ' The filename
        intTmpLen = Len(StrTmp) - intPos
        If intTmpLen >= 1 Then
          strFileName = Mid$(StrTmp, intPos + 1, intTmpLen)
          StrTmp = Left$(StrTmp, intPos)
          fFileNameFound = True
          Exit For
        End If

    End Select
  Next intPos

  If Not fFileNameFound Then
    strFileName = StrTmp
  Else
    ' Path is what's left over
    strPath = StrTmp
  End If

  ParsePath = True
  
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ParsePath"
  Resume PROC_EXIT

End Function

Public Function PurgeFile(strFile As String) As Boolean
  ' Comments  : Replaces a file's contents with space characters for security
  ' Parameters: strFile - path and name of file to purge
  ' Returns   : True if successful, False otherwise
  ' Source    : Total VB SourceBook 6
  '
  Dim strPurge1 As String
  Dim strPurge2 As String
  Dim lngCount As Long
  Dim lngFile As Long
  Dim lngCounter As Long
  Dim lngPos As Long

  Const cintBlockSize = 2048

  On Error GoTo PROC_ERR

  ' Build up a string of padding characters
  strPurge1 = String(cintBlockSize, "X")
  strPurge2 = String(cintBlockSize, " ")

  ' Get a handle and open the file
  lngFile = FreeFile
  Open strFile For Binary As lngFile

  ' Get the buffer boundary
  lngCount = (LOF(lngFile) \ cintBlockSize) + 1

  ' Walk through all characters
  For lngCounter = 1 To lngCount
    ' Go to the position
    lngPos = Seek(lngFile)
    ' Place the pad characters
    Put lngFile, , strPurge1
    Put lngFile, lngPos, strPurge2

  Next lngCounter

  ' Close the file and delete it
  Close lngFile
  Kill strFile

  PurgeFile = True
  
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "PurgeFile"
  Resume PROC_EXIT

End Function

Public Function ReadFile( _
  ByVal intFile As Integer, _
  ByRef abytBuffer() As Byte, _
  ByVal lngNumberOfBytes As Long) _
  As Long
  ' Comments  : This function attempts to read the specified number of
  '             bytes from the file.
  ' Parameters: intFile - The file to read from
  '             abytBuffer - The buffer to read the bytes into
  '             lngNumberOfBytes - The number of bytes to read
  ' Returns   : The actual number of bytes read.
  ' Source    : Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR
  
  Dim lngLen As Long
  Dim lngActualBytesRead As Long
  Dim lngStart As Long
    
  ' Get the starting position of the next read
  lngStart = Loc(intFile) + 1
  ' Get the length of the file
  lngLen = LOF(intFile)

  ' Check to see if there is more data to read from the file
  If lngStart < lngLen Then
    ' Check to see if we are attempting to read past the end of the file
    If lngStart + lngNumberOfBytes < lngLen Then
      lngActualBytesRead = lngNumberOfBytes
    Else
      ' If we are attempting to read more data than is left in the file,
      ' calculate
      ' how much data we should read
      lngActualBytesRead = lngLen - (lngStart - 1)
    End If
    
    ' Create the buffer to hold the data
    ReDim abytBuffer(lngActualBytesRead - 1) As Byte
    ' Do the read
    Get intFile, lngStart, abytBuffer
  Else
    ' If we attempted to read past the end of file, return zero bytes read
    lngActualBytesRead = 0
  End If
  
  ' Return the number of bytes read
  ReadFile = lngActualBytesRead
  
PROC_EXIT:
  Exit Function

PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ReadFile"
  Resume PROC_EXIT
  
End Function

Public Function SetFileAttr( _
  strFileName As String, _
  fNormal As Boolean, _
  fReadOnly As Boolean, _
  fHidden As Boolean, _
  fSystem As Boolean) _
  As Long
  ' Comments  : Sets the specified attributes on the specified file
  ' Parameters: strFileName - fully qualified path and name of the file
  '             fNormal - True to attribute the file as Normal
  '             fReadOnly - True to attribute the file as ReadOnly
  '             fHidden - True to attribute the file as Hidden
  '             fSystem - True to attribute the file as System
  ' Returns   : new attributes of the file as a long
  ' Source    : Total VB SourceBook 6
  '
  Dim lngAttr As Long
  
  On Error GoTo PROC_ERR
  
  ' Build up the attributes integer
  If fNormal Then
    lngAttr = lngAttr + VBA.vbNormal
  End If
  
  If fReadOnly Then
    lngAttr = lngAttr + VBA.vbReadOnly
  End If
  
  If fHidden Then
    lngAttr = lngAttr + VBA.vbHidden
  End If
  
  If fSystem Then
    lngAttr = lngAttr + VBA.vbSystem
  End If
  
  ' Call the VBA SetAttr command to do the actual work
  SetAttr strFileName, lngAttr
  
  ' Return the new attributes
  lngAttr = GetAttr(strFileName)
  
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "SetFileAttr"
  Resume PROC_EXIT

End Function

Public Function SubDirsToArray( _
  strPath As String, _
  astrIn() As String, _
  fIncludeReadonly As Boolean, _
  fIncludeHidden As Boolean) _
  As Integer
  ' Comments  : Populates the passed array with a list of subdirectories in
  '             the specified directory
  ' Parameters: strPath - Directory to read
  '             astrIn() - Array to hold file names (0-based)
  '             fIncludeReadonly - True to include read-only directories
  '             fIncludeHidden - True to include hidden directories
  ' Returns   : Number of files in the directory
  ' Source    : Total VB SourceBook 6
  '
  Dim intCounter As Integer
  Dim StrTmp As String
  Dim lngAttr As Long
  Dim fInclude As Boolean
  
  On Error GoTo PROC_ERR

  ' Get the correct attributes
  lngAttr = VBA.vbDirectory
  
  If fIncludeReadonly Then
    lngAttr = lngAttr + VBA.vbReadOnly
  End If
  
  If fIncludeHidden Then
    lngAttr = lngAttr + VBA.vbHidden
  End If
  
  ' Make sure there is a trailing slash
  If Right$(strPath, 1) <> "\" Then
    strPath = strPath & "\"
  End If
  
  ' Get the first directory entry
  StrTmp = Dir(strPath, lngAttr)
  
  ' Loop through all entries
  Do Until StrTmp = ""
  
    fInclude = True
    
    ' Are the correct attributes set?
    If IsAttrSet(strPath & StrTmp, vbDirectory) Then
      If IsAttrSet(strPath, VBA.vbHidden) Then
        fInclude = fIncludeHidden
      End If
      If IsAttrSet(strPath & StrTmp, VBA.vbReadOnly) Then
        fInclude = fIncludeReadonly
      End If
    Else
      fInclude = False
    End If
      
    ' Don't add the . and .. entries
    If StrTmp = "." Or StrTmp = ".." Then
      fInclude = False
    End If
    
    If fInclude Then
      ' Grow the array to hold the new entry
      ReDim Preserve astrIn(intCounter)
      ' Add the item to the array
      astrIn(intCounter) = StrTmp
      ' Increment the array size counter
      intCounter = intCounter + 1
    End If
      
    ' Get the next entry in the directory node
    StrTmp = Dir
  Loop
  
  ' Return the number found
  SubDirsToArray = intCounter

PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "SubDirsToArray"
  Resume PROC_EXIT

End Function
'================================================================================================================================
' Функция проверяет, существует ли заданный путь для копирования, если задан путь с именем файла, то он принимается как новое
' полное имя файла, если без имени файла, образуется полное имя с использованием унаследованного имени
'================================================================================================================================
Public Function CreateCopyPath(sNewPath As String, sFileName As String) As String
Dim sRes As String, sWork As String, iL As Long
'-------------------------------------
' Проверяем, существует ли новая директория, как директория
If DirExists(sNewPath) Then           ' Задана существующая директория, обьразуем новый путь
   sRes = IIf(Right(sNewPath, 1) = "\", sNewPath, sNewPath & "\") & sFileName
   GoTo CreateCopyPath_EXIT
Else                                  ' Или директория не существует, или задана путь сразу с файлом
   sWork = GetFileNamePart(sNewPath)             ' пробуем получить имя файла
   iL = Len(sWork) - InStrRev(sWork, ".")        ' позиция точки в имени файла справа
   If iL > 0 And iL < 5 Then                     ' в заданном пути содержится файл
      sWork = GetFilePathPart(sNewPath)                          ' Получаем путь
      If Not DirExists(sWork) Then CreateDir sWork               ' Если директория в составе пути не существует - создаем ее
      sRes = sNewPath                                            ' Возвращаем входящий путь с именем файла
   Else                                          ' в заданном пути нет имени файла, зажана не существующая директория
      CreateDir sNewPath                                         ' создаем ранее отсутсвующую директорию
      sRes = IIf(Right(sNewPath, 1) = "\", sNewPath, sNewPath & "\") & sFileName
   End If
End If
'--------------------------------------------------
CreateCopyPath_EXIT:
        CreateCopyPath = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function

