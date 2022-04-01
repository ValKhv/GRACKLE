VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cZIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************************************************************
'*******************************************************************************************************************************
' Класс размещает файлы в архиве и извлекает их оттуда
'*******************************************************************************************************************************
'*******************************************************************************************************************************
Option Explicit
Option Compare Database

'#############################################################################################################################
' ДЕКЛАРАЦИИ
#If Win64 Then                                               ' НОВАЯ СРЕДА
     '-----------------------------------------------------------------------------------------------------------------------
        Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'----------------------------------------------------------------------------------------------------------------------------
#Else                                                       ' СТАРАЯ СРЕДА
        Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
'#############################################################################################################################
Private Type FileInZip
       fileName As String
       FilePath As String
       FileTYPE As String
       LastModified As String
       UnpackedSize As String
       PackedSize As String
       Ratio As String
       CRC As String
End Type

'=====================================================================================================================================
' Функция по заданному пути к ZIP- файлу выдает массив строк-описаний (реально - одну строку с разделением FileDelim )
'                  всех содержащихся внутри файлов
'=====================================================================================================================================
Public Function GetFilesInZip(ZipFile As String, Optional sPairDelim As String = "=", _
                                      Optional sKVDelim As String = ";", Optional FileDelim As String = "#", _
                                      Optional bFirst As Boolean = True) As String
Dim sRes As String, sWork As String                ' Возвращаемый результат, рабочая строка
Dim objShell As Object, varFileName As Object      ' Системный объект, указатель на файл в архиве
Dim myFileInZip As FileInZip                       ' Структура с описание файла
Dim varNameOfZipFile As Variant
    varNameOfZipFile = ZipFile
'------------------------------------------------------------------------
' Проверяем наличие файла
    If bFirst Then
       If Len(Dir(ZipFile)) = 0 Then GoTo ExitHere
    End If
'------------------------------------------------------------------------------------------------------------------
    Set objShell = CreateObject("Shell.Application")
    For Each varFileName In objShell.NameSpace(varNameOfZipFile).Items   ' НАЧИНАЕМ ПРОЦЕССИНГ ВСЕХ ВНУТР. ФАЙЛОВ
                With objShell.NameSpace(varNameOfZipFile)
                    myFileInZip.fileName = .GetDetailsOf(varFileName, 0)
                    myFileInZip.FileTYPE = .GetDetailsOf(varFileName, 1)
                    myFileInZip.LastModified = FilterUnicode(.GetDetailsOf(varFileName, 7))
                    myFileInZip.UnpackedSize = .GetDetailsOf(varFileName, 5)
                    myFileInZip.PackedSize = .GetDetailsOf(varFileName, 2)
                    myFileInZip.Ratio = FilterUnicode(.GetDetailsOf(varFileName, 6))
                    myFileInZip.CRC = .GetDetailsOf(varFileName, 8)
                    myFileInZip.FilePath = ZipFile & "\" & myFileInZip.fileName
                End With
        '------------------------------------------------------------------------------
        sRes = sRes & FileInZipToKV(myFileInZip) & FileDelim
        If myFileInZip.PackedSize = "" Then                    ' ОБРАБАТЫВАЕМ ПАПКУ (внутри могут быть файлы => рекурсия)
               sWork = GetFilesInZip(myFileInZip.FilePath, sPairDelim, sKVDelim, FileDelim, False)
               If sWork <> "" Then sRes = sRes & sWork & FileDelim
        End If
    Next varFileName
'----------------------------------------------------------------------------------------------
ExitHere:
  GetFilesInZip = sRes  '!!!!!!!!!!!!!!!!!!!
  Set objShell = Nothing
End Function
'=====================================================================================================================================
' Функция архивирует в ZIP
'       InputFile - единственный файл или набор файлов, разделенных ";"
'=====================================================================================================================================
Public Sub Zip(ZipFile As String, InputFile As String)
Dim myFileVar() As String

On Error GoTo ErrHandler
'----------------------------------------------
    Call AddFilesToZip(ZipFile, CreateVar(InputFile))
'------------------------------------------------
ExitProc:
    Exit Sub
'-----------------------------
ErrHandler:
    Select Case Err.Number
        Case Else
            MsgBox "Error " & Err.Number & _
                   ": " & Err.Description, _
                   vbCritical, "Unexpected error"
    End Select
    Resume ExitProc
    Resume
End Sub
'=====================================================================================================================================
' Функция выдает список всех файлов в архиве
'=====================================================================================================================================
Public Function LISTFILES(sZipPath As String, Optional DLM As String = ";") As String
Dim sRes As String
    sRes = Join(ListFilesInZip(sZipPath), DLM)
'-----------------------------------
    LISTFILES = sRes '!!!!!!!!!!!!
End Function
'=====================================================================================================================================
' Функция разархивирует zip - файл
'=====================================================================================================================================
Public Sub UnZip(ZipFile As String, Optional TargetFolderPath As String = vbNullString, Optional OverwriteFile As Boolean = False)
On Error GoTo ErrHandler
    Dim oApp As Object
    Dim FSO As Object
    Dim fil As Object
    Dim DefPath As String
    Dim strDate As String

    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Len(TargetFolderPath) = 0 Then
        DefPath = CurrentProject.Path & ""
    Else
        If FSO.folderexists(TargetFolderPath) Then
            DefPath = TargetFolderPath & ""
        Else
            Err.Raise 53, , "Folder not found"
        End If
    End If

    If FSO.FileExists(ZipFile) = False Then
        MsgBox "System could not find " & ZipFile _
            & " upgrade cancelled.", _
            vbInformation, "Error Unziping File"
        Exit Sub
    Else
        'Extract the files into the newly created folder
        Set oApp = CreateObject("Shell.Application")

        With oApp.NameSpace(ZipFile & "")
            If OverwriteFile Then
                For Each fil In .Items
                    If FSO.FileExists(DefPath & fil.Name) Then
                        Kill DefPath & fil.Name
                    End If
                Next
            End If
            oApp.NameSpace(CVar(DefPath)).CopyHere .Items
        End With

        On Error Resume Next
        Kill Environ("Temp") & "Temporary Directory*"

        'Kill zip file
        Kill ZipFile
    End If
'---------------------------------------------------------------------------
ExitProc:
    On Error Resume Next
    Set oApp = Nothing
    Exit Sub
'------------------------------------------
ErrHandler:
    Select Case Err.Number
        Case Else
            MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Unexpected error"
    End Select
    Resume ExitProc
    Resume
End Sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub GoToSleep(TimeInMilliSeconds As Long)
' This code was originally written by
' Doug Steele, MVP  AccessHelp@rogers.com
' http://I.Am/DougSteele
' You are free to use it in any application
' provided the copyright notice is left unchanged.
'
' Description:  Pause execution for a given number of millseconds.
'
' Inputs:       TimeInMilliSeconds  Long  The number of milliseconds for which
'                                         to pause execution.

    If TimeInMilliSeconds > 0 Then
        Call Sleep(TimeInMilliSeconds)
    End If

End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------
' Добавление файлов к архиву Zip
' Inputs:       ZipFile          String   Full path to the zip file.
'                                         (Note that the file does not have to
'                                         exist already)
'               colFilesToAdd    Variant  An array containing full paths to each
'                                         of the files to add.
'-----------------------------------------------------------------------------------------------------------------------------------------
Private Sub AddFilesToZip(ZipFile As String, colFilesToAdd As Variant)


On Error GoTo EH

Dim objShell As Object
Dim lngFilesAdded As Long
Dim lngFilesInZip As Long
Dim lngLoop As Long
Dim varNameOfZipFile As Variant

    DoCmd.Hourglass True

    If IsArray(colFilesToAdd) = False Then
    
' No files were passed to add. Do nothing

    Else
    
' Check whether the zip file already exists.
' If it does, make sure that it isn't an empty file.
' If it is an empty file (or if it doesn't already exist), initialize it.

        If Len(Dir(ZipFile)) > 0 Then
            If FileLen(ZipFile) = 0 Then
                Call InitializeZipFile(ZipFile)
            End If
        Else
            Call InitializeZipFile(ZipFile)
        End If

' Instantiate an instant of the Shell Application object

        Set objShell = CreateObject("Shell.Application")

' For some reason I haven't bothered to determine, the name of the zip file
' must be stored in a variant in order to use it as the NameSpace.

        varNameOfZipFile = ZipFile

' Determine how many files are already in the zip file.

        lngFilesInZip = objShell.NameSpace(varNameOfZipFile).Items.Count

' Add each file in the array to the zip file.

        For lngLoop = LBound(colFilesToAdd) To UBound(colFilesToAdd)
            
            lngFilesAdded = 0

' Don't add a file unless it actually exists...

            If Len(Dir(colFilesToAdd(lngLoop))) > 0 Then
                objShell.NameSpace(varNameOfZipFile).CopyHere (colFilesToAdd(lngLoop))
                lngFilesAdded = lngFilesAdded + 1
'                Do Until objShell.NameSpace(varNameOfZipFile).Items.Count >= (lngFilesInZip + lngFilesAdded)
'                    Call GoToSleep(100)
'                Loop
            End If
        
        Next lngLoop
 
' Depending on the size of the files being added, compressing can take an appreciable time.
' This is a bit of a kludge to ensure that we wait until compressing is done.

        If lngFilesAdded > 0 Then
            lngLoop = 0
            Do Until objShell.NameSpace(varNameOfZipFile).Items.Count >= (lngFilesInZip + lngFilesAdded) _
                Or lngLoop > 20

                Call GoToSleep(100)
                lngLoop = lngLoop + 1

            Loop
        End If
 
    End If

EndIt:
    Set objShell = Nothing
    DoCmd.Hourglass False
    Exit Sub

EH:
    MsgBox Err.Number & ": " & Err.Description & vbCrLf & _
        "Occurred in AddFilesToZip"
    Resume EndIt

End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------
' Функция возвращает список файлов в ZIP архиве в виде массива строк
'-----------------------------------------------------------------------------------------------------------------------------------------
Private Function ListFilesInZip(ZipFile As String) As Variant

On Error GoTo EH

Dim objShell As Object
Dim lngFilesInZip As Long
Dim lngLoop As Long
Dim strFiles() As String
Dim varFileName As Object
Dim varNameOfZipFile As Variant

    DoCmd.Hourglass True

    If Len(Dir(ZipFile)) = 0 Then
    
' Zip file doesn't exist.

    Else
    
' Instantiate an instant of the Shell Application object

        Set objShell = CreateObject("Shell.Application")

' For some reason I haven't bothered to determine, the name of the zip file
' must be stored in a variant in order to use it as the NameSpace.

        varNameOfZipFile = ZipFile

' Determine how many files are already in the zip file.

        lngFilesInZip = objShell.NameSpace(varNameOfZipFile).Items.Count

        ReDim strFiles(0 To (lngFilesInZip - 1))
        lngLoop = 0

        Debug.Print "Details contained in " & ZipFile & vbCrLf
        
        For Each varFileName In objShell.NameSpace(varNameOfZipFile).Items
            strFiles(lngLoop) = varFileName
' http://msdn.microsoft.com/en-us/library/bb787870(VS.85).aspx says that
'  0 Retrieves the name of the item.
'  1 Retrieves the size of the item.
'  2 Retrieves the type of the item.
'  3 Retrieves the date and time that the item was last modified.
'  4 Retrieves the attributes of the item.
' -1 Retrieves the info tip information for the item.
'
' I found
'  0 Retrieves the name of the item.
'  1 Retrieves the type of the item.
'  7 Retrieves the date and time that the item was last modified.
'  5 Retrieves the unpacked size of the item.
'  6 Retrieves the Ratio packed/unpacked size of the item.
'  2 Retrieves the packed size of the item.
'  8 Retrieves the CRC of the item.

            With objShell.NameSpace(varNameOfZipFile)
                Debug.Print "Name = " & .GetDetailsOf(varFileName, 0)
                Debug.Print "Type = " & .GetDetailsOf(varFileName, 1)
                Debug.Print "Last Modified = " & .GetDetailsOf(varFileName, 7)
                Debug.Print "Unpacked Size = " & .GetDetailsOf(varFileName, 5)
                Debug.Print "Packed Size = " & .GetDetailsOf(varFileName, 2)
                Debug.Print "Ratio = " & .GetDetailsOf(varFileName, 6)
                Debug.Print "CRC = " & .GetDetailsOf(varFileName, 8)
                Debug.Print vbCrLf & String(20, "-") & vbCrLf
            End With
            lngLoop = lngLoop + 1
        Next varFileName
    
        ListFilesInZip = strFiles

    End If

EndIt:
    Set objShell = Nothing
    DoCmd.Hourglass False
    Exit Function

EH:
    MsgBox Err.Number & ": " & Err.Description & vbCrLf & _
        "Occurred in ListFilesInZip"
    Resume EndIt

End Function
'-----------------------------------------------------------------------------------------------------------------------------------------
' Функция создает пустой Zip - файл
'-----------------------------------------------------------------------------------------------------------------------------------------
Private Sub InitializeZipFile(NameOfFile As String)

On Error GoTo EH

Dim intFile As Integer

    If Len(Dir(NameOfFile)) > 0 Then
        Kill NameOfFile
    End If

    intFile = FreeFile
    Open NameOfFile For Output As #intFile
    Print #intFile, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #intFile

EndIt:
    Exit Sub

EH:
    MsgBox Err.Number & ": " & Err.Description & vbCrLf & _
        "Occurred in InitalizeZipFile"
    Resume EndIt

End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------
' Вывод всех файлов из ZIP архива в заданный каталог
' Inputs:       ZipFile          String   Full path to the zip file.
'               DestFolder       String   Name of folder to use for unzipping.
'-----------------------------------------------------------------------------------------------------------------------------------------
Private Sub UnzipFilesToFolder(ZipFile As String, DestFolder As String)

On Error GoTo EH

Dim objShell As Object
Dim varDestFolder As Variant
Dim varZipFile As Variant

' Ensure that the folder and Zip file both exist.

    If Len(Dir(DestFolder, vbDirectory)) > 0 _
        And Len(Dir(ZipFile)) > 0 Then
  
' Instantiate an instance of the Shell Application object

        Set objShell = CreateObject("Shell.Application")

' Play along with the Shell's idiosyncracies, and use variants.
        varDestFolder = DestFolder
        varZipFile = ZipFile

' Note that if you only wanted to unzip a single file, you could change
' the line of code below to
'        objShell.Namespace(varDestFolder).CopyHere _
'            objShell.Namespace(varZipFile).items(FileName)
' where FileName would be the name of the file to unzip (no path)

        objShell.NameSpace(varDestFolder).CopyHere _
            objShell.NameSpace(varZipFile).Items
      
        Set objShell = Nothing

    End If
  
EndIt:
    Exit Sub

EH:
    MsgBox Err.Number & ": " & Err.Description & vbCrLf & _
        "Occurred in UnzipFilesToFolder"
    Resume EndIt

End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------
' Функция из заданной строки получает архив (разделители vbcrlf или ";"
'-----------------------------------------------------------------------------------------------------------------------------------------
Private Function CreateVar(sp As String) As Variant
Dim varFiles As Variant
'----------------------------------
    If Len(sp) > 0 Then
         If InStr(1, sp, ";") > 0 Then
            varFiles = Split(sp, ";")
         Else
            varFiles = Split(sp, vbCrLf)
         End If
    Else
         varFiles = ""
    End If
'----------------------------------------------------------
    CreateVar = varFiles  '!!!!
End Function
'-------------------------------------------------------------------------------------------------------------------------
' Функция транслирует тип FileInZip в KV - строку
'-------------------------------------------------------------------------------------------------------------------------
Private Function FileInZipToKV(fz As FileInZip, Optional sPairDelim As String = "=", _
                                                                               Optional sKVDelim As String = ";") As String
Dim sRes As String
'-----------------------------------------------------------------------
    sRes = "FileName" & sPairDelim & fz.fileName & sKVDelim & _
           "FilePath" & sPairDelim & fz.FilePath & sKVDelim & _
           "FileType" & sPairDelim & fz.FileTYPE & sKVDelim & _
           "LastModified" & sPairDelim & fz.LastModified & sKVDelim & _
           "UnpackedSize" & sPairDelim & fz.UnpackedSize & sKVDelim & _
           "PackedSize" & sPairDelim & fz.PackedSize & sKVDelim & _
           "Ratio" & sPairDelim & fz.Ratio & sKVDelim & _
           "CRC" & sPairDelim & fz.CRC & sKVDelim
'-------------------------------------------------------------------------
    FileInZipToKV = sRes  '!!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------
' Функция фильтрует строку от UNICODE
'------------------------------------------------------------------------------------------------------------------------
Private Function FilterUnicode(str As String) As String
Dim sRes As String
  Const IUNKCHAR As Integer = &H3F
  Dim I As Long, tChar As String
  
  If str = "" Then Exit Function
  '--------------------------
  For I = 1 To Len(str)
     tChar = Mid(str, I, 1)
    If Not ((Asc(tChar) = IUNKCHAR) And (AscW(tChar) <> IUNKCHAR)) Then ' Символ не UNICODE
      sRes = sRes & tChar
    End If
  Next
'---------------------------------------------------------
  FilterUnicode = sRes '!!!!!!!!!!!!!!!
End Function


