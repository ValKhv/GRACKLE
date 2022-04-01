Attribute VB_Name = "#_EXPORT"
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
'                 d$$$$$$$$$F                $P"                   ##  ##  ##  ##   #######   ####### ##      ### ######   ####  #####   ######
'                 $$$$$$$$$$F                                       ## ##  ##  ##  ##    ##     ##    ####   #### ##   ## ##  ## ##  ##    ##
'                  *$$$$$$$$"                                        ####  #####  ##     ##     ##    ## ## ## ## ######  ##  ## ##  ##    ##
'                    "***""               _____________                                         ##    ##  ###  ## ##      ##  ## ####      ##
' STANDARD MODULE WITH DEFAULT FUNCTIONS |v 2021/08/20 |                                        ##    ##       ## ##      ##  ## ##  ##    ##
' The module contains some functions for import and export from G-VBA library                 ######  ##       ## ##       ####  ##   ##   ##
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************Option Explicit
Option Explicit

Private Const MOD_NAME As String = "#_EXPORT"
'**********************************

'======================================================================================================================================================
' Get Value from another database
'======================================================================================================================================================
Public Function GetAValue(Optional id As Long, Optional sFLD As String, Optional TBL As String, Optional sPath As String) As Variant
Dim SQL As String, RS As DAO.Recordset, vRes As Variant, PathToDB As String

On Error GoTo ErrHandle
'---------------------------
    If sPath = "" Then
       PathToDB = OpenDialog(GC_FILE_PICKER, "Pick the database to retrieve some value", , False)
       If PathToDB = "" Then Exit Function
    Else
       PathToDB = sPath
    End If
    
    SQL = "SELECT  " & sFLD & " FROM " & TBL & " IN " & sCH(PathToDB) & " WHERE(ID = " & id & ")"
    Set RS = CurrentDb.OpenRecordset(SQL)
    With RS
        If Not .EOF Then
            .MoveLast: .MoveFirst
            vRes = Nz(.FIELDS(sFLD).value)
        End If
    End With
'---------------------------
ExitHere:
    GetAValue = vRes '!!!!!!!!!
    Exit Function
'---------
ErrHandle:
    ErrPrint2 "GetAValue", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'========================================================================================================================================================
' Extract and Save Object To File
'========================================================================================================================================================
Public Sub ExportDataToText(Optional sExportList As String, Optional sExceptList As String = "$$*", Optional ObjectType As AcObjectType = acTable, _
                                                                                                                            Optional DLM As String = ";")
Dim sOBJS() As String, nObjs As Integer, sWork As String, I As Integer
   Select Case ObjectType
   Case acTable:
        If sExportList <> "" Then
            sOBJS = Split(sExportList, DLM): nObjs = UBound(sOBJS)
        Else
            sWork = ListDBObjects(ACC_TBL_LOCAL, DLM): If sWork = "" Then Exit Sub
            sOBJS = Split(sWork, DLM): nObjs = UBound(sOBJS)
            For I = 0 To nObjs
                If Not IsWordInList(sOBJS(I), sExceptList, DLM) Then
                
                End If
            Next I
        End If
   Case acQuery:
   Case acForm:
   Case acReport:
   Case acModule:
   Case acMacro:
   Case Else:
   End Select


   sWork = SaveObjectToFile("ITEMS", acTable)
   Debug.Print sWork
End Sub

'========================================================================================================================================================
' Extract and Save Object To File
'========================================================================================================================================================
Public Function SaveObjectToFile(ObjectName As String, ObjectType As AcObjectType, Optional Folder As String, _
                   Optional bCreateSubFldAuth As Boolean = True, Optional UseTransferTextForTable As Boolean, Optional XMLExport As Boolean) As String
Dim sFolder As String, sFile As String


On Error GoTo ErrHandle
'------------------------------------------
    If Folder <> "" Then
        sFolder = Folder
        If InStr(1, sFolder, ":", vbBinaryCompare) = 0 Then sFolder = CurrentProject.Path & "\" & sFolder
    End If
'------------------------------------------
       Select Case ObjectType
       Case acForm:
            If (sFolder = "") And bCreateSubFldAuth Then sFolder = CurrentProject.Path & "\FORMS"
            If XMLExport Then
               sFile = ProceedXMLExport(ObjectName, ObjectType, sFolder)
            Else
               sFile = sFolder & "\" & ObjectName & ".frm": Call CheckFolderFile(sFolder, sFile)
               SaveAsText acForm, ObjectName, sFile & ".fr"
            End If
       Case acTable:
             If (sFolder = "") And bCreateSubFldAuth Then sFolder = CurrentProject.Path & "\DATA"
                          
             If XMLExport Then
                sFile = ProceedXMLExport(ObjectName, ObjectType, sFolder)
             ElseIf UseTransferTextForTable Then                        ' STANDARD EXPORT TO CVS
                   sFile = sFolder & "\" & ObjectName & ".tbl": Call CheckFolderFile(sFolder, sFile)
                   DoCmd.TransferText acExportDelim, , ObjectName, sFile, True, , acUTF16
             Else                                                       ' SAVE VIA CUSTOM PROCEDURE
                   sFile = TableToFile(ObjectName, "DATA", ";", ",", "csv")
             End If
       Case acQuery:
             If (sFolder = "") And bCreateSubFldAuth Then sFolder = CurrentProject.Path & "\QUERIES"
             If XMLExport Then
                sFile = ProceedXMLExport(ObjectName, ObjectType, sFolder)
             Else
                sFile = sFolder & "\" & ObjectName & ".sql": Call CheckFolderFile(sFolder, sFile)
                SaveAsText acQuery, ObjectName, sFile
             End If
       Case acReport:
            If (sFolder = "") And bCreateSubFldAuth Then sFolder = CurrentProject.Path & "\REPORTS"
            If XMLExport Then
               sFile = ProceedXMLExport(ObjectName, ObjectType, sFolder)
            Else
               sFile = sFolder & "\" & ObjectName & ".rep": Call CheckFolderFile(sFolder, sFile)
               SaveAsText acReport, ObjectName, sFile
            End If
       Case acModule:
            If (sFolder = "") And bCreateSubFldAuth Then sFolder = CurrentProject.Path & "\CODE"
            sFile = IIf(IsClassModule(ObjectName), ObjectName & ".cls", ObjectName & ".bas")
            Call CheckFolderFile(sFolder, sFile)
            SaveAsText acModule, ObjectName, sFile
       Case acMacro:
            If (sFolder = "") And bCreateSubFldAuth Then sFolder = CurrentProject.Path & "\CODE"
            sFile = sFolder & "\" & ObjectName & ".macro": Call CheckFolderFile(sFolder, sFile)
            SaveAsText acModule, ObjectName, sFile
       Case Else
            If (sFolder = "") And bCreateSubFldAuth Then sFolder = CurrentProject.Path & "\XYZ"
            sFile = sFolder & "\" & ObjectName & ".xyz": Call CheckFolderFile(sFolder, sFile)
            SaveAsText ObjectType, ObjectName, sFile
       End Select
'----------------------------------------------------
ExitHere:
    SaveObjectToFile = sFile '!!!!!!!!!!!
    Exit Function
'--------------
ErrHandle:
    ErrPrint2 "SaveObjectToFile", Err.Number, Err.Description, MOD_NAME
    Err.Clear: sFile = ""
End Function
'========================================================================================================================================================
' Load Object From File
'========================================================================================================================================================
Public Function LoadObjectFromFile(ObjectType As AcObjectType, ObjectFile As String) As Boolean
Dim ObjectName As String, iL As Integer
On Error GoTo ErrHandle
'------------------------------------------
  If Dir(ObjectFile) = "" Then Err.Raise 1000, , "Can't find file " & ObjectFile
  ObjectName = FileNameOnly(ObjectFile)
  iL = InStrRev(ObjectName, "."): ObjectName = Left(ObjectName, iL - 1)
  LoadFromText ObjectType, ObjectName, ObjectFile

'----------------------------------------------------
ExitHere:
    LoadObjectFromFile = True '!!!!!!!!!!!
    Exit Function
'--------------
ErrHandle:
    ErrPrint2 "LoadObjectFromFile", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'========================================================================================================================================================
' Extract and Save References To File
'========================================================================================================================================================
Public Function SaveRelationshipsToFile() As String
Dim sFile As String, sRes As String, sWork As String
Dim rel As DAO.Relation, FLD As DAO.Field
    
Const DLM As String = ";"
Const fLeft As String = "{", fRight As String = "}"

On Error GoTo ErrHandle
'-------------------------
    sFile = CurrentProject.Path & "\" & FilenameWithoutExtension(CurrentProject.Name) & ".rel"
    If Dir(sFile) <> "" Then
        If MsgBox("The file with relations is exist and will be removed. Shoud we proceed?", vbYesNoCancel + vbQuestion, "SaveRelations") <> vbYes Then
            Exit Function
        End If
    End If
'-------------------------
    For Each rel In CurrentDb.Relations
        With rel
            sWork = .Name & DLM & .Attributes & DLM & .Table & DLM & .foreignTable
            For Each FLD In .FIELDS
               If FLD.Name <> "" Then
                    sWork = sWork & DLM & "{" & FLD.Name & DLM & FLD.ForeignName & "}"
               End If
            Next
        End With
        sRes = sRes & sWork & vbCrLf
    Next
'-------------------------
If sRes <> "" Then Call WriteStringToFile(sFile, sRes)
'-------------------------
ExitHere:
    SaveRelationshipsToFile = sFile '!!!!!!!!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "SaveRelationshipsToFile", Err.Number, Err.Description, MOD_NAME
    Err.Clear: sFile = ""
End Function
'========================================================================================================================================================
' Restore REfs from text file (reuire delete all current refs before)
'========================================================================================================================================================
Public Sub RestoreRelationshipsFromFile(sFile As String)
Dim sRef As String, RFS() As String, nDim As Integer, I As Integer, iL As Long
Dim ATTRS() As String, nATTRS As Integer, FLDS() As String, nFlds As Integer, J As Integer

Const DLM As String = ";"
Const fLeft As String = "{", fRight As String = "}"

On Error GoTo ErrHandle
'-------------------------
sRef = ReadTextFile(sFile)
If sRef = "" Then Err.Raise "1000", , "Can't Read File " & sFile
RFS = Split(sRef, vbCrLf): nDim = UBound(RFS)
For I = 0 To nDim
   If RFS(I) <> "" Then
        iL = InStr(1, RFS(I), fLeft)
        If iL > 0 Then
           sRef = Left(RFS(I), iL - 1)
           ATTRS = Split(RFS(I), DLM): FLDS = TextBetweenTags(RFS(I), "([^{]*?)\w(?=\})")
           Call CreateRelationship(ATTRS(0), ATTRS(2), ATTRS(3), ATTRS(1), FLDS)
        End If
   End If
Next I
'-------------------------
ExitHere:
    Exit Sub
'----------
ErrHandle:
    ErrPrint2 "RestoreRelationshipsFromFile", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Sub

'======================================================================================================================================================
' Import Functions From File Or Directory To Some Module
'======================================================================================================================================================
Public Function ImportVBA(Optional sPath As String, Optional bFromVBALib As Boolean = True, Optional ProjectName As String, _
                                                                             Optional ModuleName As String, Optional DLM As String = "--") As String
Dim sFiles As String, sFolder As String, sRes As String, oMod As Module, sProjectName As String, sMod As String
Dim FLS() As String, nFLS As Integer, I As Integer, sCode As String, PARTS() As String
     
Const LOCAL_DLM As String = ";"
Const DEFAULT_MOD As String = "_TEST"
      On Error GoTo ErrHandle
'-------------------------
      If bFromVBALib Then
         sFolder = GetCodeFolder(GetGracklePath())
      Else
         sFolder = GetCodeFolder(CurrentDb.Name)
      End If

      If sPath <> "" Then
          sFiles = sPath
      Else
          sFiles = OpenDialog(GC_FILE_PICKER, "Pick the files", "VBA Files,*.vba", True, sFolder)
      End If
      
      If sFiles = "" Then Exit Function
      If InStr(1, sFiles, LOCAL_DLM) > 0 Then
          FLS = Split(sFiles, LOCAL_DLM): nFLS = UBound(FLS)
      Else
          nFLS = 0: ReDim FLS(nFLS): FLS(nFLS) = sFiles
      End If
      
      If ProjectName = "" Then
         sProjectName = Application.VBE.VBProjects(1).Name
      Else
         sProjectName = ProjectName
      End If
      
      For I = 0 To nFLS
           
           If ModuleName <> "" Then
               sMod = ModuleName
           Else
               sMod = GetModNameFromFile(FLS(I), sProjectName)
           End If
           If sMod = "" Then GoTo NextFile
           
           sCode = ReadTextFileUTF8(FLS(I))
           If sCode = "" Then GoTo NextFile
             
           sCode = Chr(39) & String(60, "<") & " IMPORT FROM " & FileNameOnly(FLS(I)) & vbCrLf & sCode
           Application.VBE.VBProjects(sProjectName).VBComponents(sMod).CodeModule.AddFromString sCode
           DoCmd.Save acModule, sMod
           
           sRes = sRes & LOCAL_DLM & sMod
           Call RenameFile(FLS(I), FolderNameOnly(FLS(I)) & "(installed)_" & FileNameOnly(FLS(I)))
NextFile:
      Next I
If sRes <> "" Then sRes = Right(sRes, Len(sRes) - Len(LOCAL_DLM))
'----------------------------
ExitHere:
      ImportVBA = sRes '!!!!!!!!!!!!!!!
      Exit Function
'--------------------
ErrHandle:
      ErrPrint2 "ImportVBA", Err.Number, Err.Description, MOD_NAME
      Err.Clear
End Function


'======================================================================================================================================================
' Save Function from memory to VBALib File
'======================================================================================================================================================
Public Function ExportVBA(Optional sModule As String, Optional FunctionCode As String, Optional ToVBALIB As Boolean = True, _
                                                                                                               Optional DLM As String = "--") As String
Dim sPath As String, sCode As String, sMod As String, sFuncName As String

Const VBA_FOLDER As String = "VBA"
    On Error GoTo ErrHandle
'---------------------------
    
    If FunctionCode = "" Then
        If MsgBox("We are going to export function from memory to file. Please be aware that your clipboard has a vba code only." & vbCrLf & _
              "Should we continue?", vbYesNoCancel + vbQuestion, "VBA Function Export") <> vbYes Then Exit Function
        sCode = FromClipboard()
    Else
        sCode = FunctionCode
    End If
    
    If sCode = "" Then Err.Raise 10005, , "No any code in memory"
    
    sCode = Chr(39) & String(60, ">") & "EXPORT " & Now() & " FROM " & CurrentProject.Name & vbCrLf & sCode
    If sModule <> "" Then
        sMod = sModule
    Else
        sMod = InputBox("Provide Module Name for code,please?", "Export VBA Function", "_TEST")
    End If
    
    sFuncName = GetTheVBAFuncName(sCode)
    If sFuncName = "" Then sFuncName = InputBox("Please, set function or sub name", "Export VBA Function")
    If sFuncName = "" Then Err.Raise 10007, , "Can't save code without File/Function Name"
    
    
    If ToVBALIB Then
       sPath = GetCodeFolder(GetGracklePath())
    Else
       sPath = GetCodeFolder(CurrentDb.Name)
    End If
    
    If Dir(sPath, vbDirectory) = "" Then Err.Raise 10006, , "Can't create the output folder"
    sPath = sPath & "\" & Format(Now(), "yyyymmdd") & DLM & sMod & DLM & sFuncName & ".vba"
    If Dir(sPath) <> "" Then Kill sPath
    
    If Not WriteStringToFileUTF8(sCode, sPath) Then sPath = ""
'---------------------------
ExitHere:
    ExportVBA = sPath '!!!!!!!!!!!!!!!!
    Debug.Print String(30, ">") & "  EXPORT VBA FUNCTION:" & sPath
    Exit Function
'-----------
ErrHandle:
    ErrPrint2 "ExportVBA", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'======================================================================================================================================================
' Export all code modules
'======================================================================================================================================================
Public Function ExportAllModules() As Integer
Dim sFolder As String, bRes As Boolean, nRes As Integer
Dim c As Object, Sfx As String, sFile As String

Const VBA_FOLDER = "CODE"

On Error GoTo ErrHandle
'------------------------------------------
sFolder = CurrentProject.Path & "\" & VBA_FOLDER
If Dir(sFolder, vbDirectory) = "" Then
        bRes = FolderCreate(sFolder)
Else
        bRes = True
End If
If Not bRes Then Err.Raise 10000, , "Can't create folder " & sFolder

sFolder = sFolder & "\" & Format(Now(), "yyyymmdd")
If Dir(sFolder, vbDirectory) = "" Then
         bRes = FolderCreate(sFolder)
Else
         bRes = True
End If
If Not bRes Then Err.Raise 10000, , "Can't create folder " & sFolder
    
'--------------------------------------
For Each c In Application.VBE.VBProjects(1).VBComponents
        Sfx = vbExtFromType(c.Type)
        If Sfx <> "" Then
            sFile = sFolder & "\" & c.Name & Sfx
            If Dir(sFile) <> "" Then Kill sFile
            c.Export _
                fileName:=sFile
            nRes = nRes + 1
        End If
Next c
'--------------------------------------
ExitHere:
    ExportAllModules = nRes '!!!!!!!!!!!!!!
    Debug.Print "The modules are saved to " & VBA_FOLDER & "; Count = " & nRes
    Exit Function
'------------
ErrHandle:
    ErrPrint2 "ExportAllModules", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function


'======================================================================================================================================================
' SAVE TABLE TO FILE AS CSV. If success - return path
'======================================================================================================================================================
Public Function TableToFile(sTableName As String, Optional sFolder As String = "DATA", Optional SDLM As String = ";", _
                                                             Optional ListDLM As String = ",", Optional sExt As String = "csv") As String
Dim sRes As String, I As Integer, sWRK As String, sRow As String, sData As String, iComplex As Integer, DLM As String
Dim sFile As String, sFlds As String, RS As DAO.Recordset, nFlds As Integer, EXTRAS As String, rst As DAO.Recordset, sFldrBase As String
   
Const STFILES As String = "EXTRAS"
Const iTextLim As Integer = 250                    ' Maximum of Text in Line
Const SSDLM As String = ";"
   
On Error GoTo ErrHandle
'----------------------------------------------
If SDLM = "" Then
    DLM = InputBox("You should setup a delimeter for csv File", "TableToFile", SSDLM)
    If DLM = "" Then Exit Function
Else
    DLM = SDLM
End If

DoCmd.Hourglass True

sFlds = GetFieldList(sTableName, DLM): If sFlds = "" Then Exit Function
nFlds = UBound(Split(sFlds, DLM))

sFldrBase = CurrentProject.Path & "\" & sFolder
If Dir(sFldrBase, vbDirectory) = "" Then FolderCreate (sFldrBase)
sFile = sFldrBase & "\" & sTableName & "." & sExt: If Dir(sFile) <> "" Then Kill sFile
'---------------------------------------------
   Set RS = CurrentDb.OpenRecordset(sTableName)
   With RS
       If Not .EOF Then
            .MoveLast: .MoveFirst
            '----------------------------------------------------------------------
            Do While Not .EOF
                     sRow = ""
                     For I = 0 To nFlds
                         sWRK = ""
                         
                         Select Case .FIELDS(I).Type
                              Case 101:                             ' Attachment: Save Files and Return [[SAVED_FILE_LIST]]
                                  Set rst = .FIELDS(I).value
                                  If rst.RecordCount > 0 Then
                                      If EXTRAS = "" Then EXTRAS = GetFOLDEREXTRACT(sFldrBase, sTableName)
                                      sWRK = _
                                      "[[" & AttachmentToDisk(sTableName, .FIELDS(I).Name, .FIELDS("ID").value, EXTRAS, "ID", ListDLM) & "]]"
                                   End If
                              Case 102, 103, 104, 105, 106, 107:    'MVF:  Create {{MVF_LIST}}
                                  Set rst = .FIELDS(I).value
                                  If rst.RecordCount > 0 Then
                                      sWRK = MVFToList(sTableName, .FIELDS(I).Name, .FIELDS("ID").value, "ID", ListDLM)
                                      If InStr(1, sWRK, ListDLM) > 0 Then sWRK = "{{" & sWRK & "}}"
                                      'Debug.Assert False
                                  End If
                              Case 9, 11:                           ' BINARY
                                  If Not IsNull(.FIELDS(I).value) Then
                                          If EXTRAS = "" Then EXTRAS = GetFOLDEREXTRACT(sFldrBase, sTableName)
                                          sWRK = _
                                          SHT(SaveBinaryFromDB(sTableName, .FIELDS(I).Name, .FIELDS("ID").value, EXTRAS, "ID"))
                                          End If
                              Case 12:                              ' MEMO/HYPERLINK
                                   sWRK = Nz(.FIELDS(I).value, "")
                                   If Len(sWRK) > iTextLim Or InStr(1, sWRK, vbCrLf) > 0 Then
                                          If EXTRAS = "" Then EXTRAS = GetFOLDEREXTRACT(sFldrBase, sTableName)
                                          sWRK = _
                                          SaveTextFromDB(sTableName, .FIELDS(I).Name, .FIELDS("ID").value, EXTRAS, "ID")
                                   End If
                              Case Else
                                   sWRK = Trim(Nz(.FIELDS(I).value, ""))
                          End Select
                          'If .Fields(i).Name = "Attachments" Then Debug.Assert False
                     sRow = sRow & Replace(sWRK, DLM, UFDELIM) & DLM
                     'Debug.Print sRow
                Next I
                
                If sRow <> "" Then
                       sRow = Left(sRow, Len(sRow) - Len(DLM)): sData = sData & sRow & vbCrLf
                End If
            
                .MoveNext
            Loop
       End If
   End With
'----------------------------------------------
sData = Trim(sData): If sData <> "" Then sData = Left(sData, Len(sData) - Len(vbCrLf))
sData = sFlds & vbCrLf & sData
sData = "sep=" & DLM & vbCrLf & sData
If WriteStringToFileUTF8(sData, sFile) Then sRes = sFile
'----------------------------------------------
ExitHere:
   TableToFile = sRes  '!!!!!!!
   DoCmd.Hourglass False
   Exit Function
'--------------------
ErrHandle:
   ErrPrint2 "TableToFile", Err.Number, Err.Description, MOD_NAME
   Err.Clear: Resume ExitHere
End Function
'======================================================================================================================================================
' The Function Import Form from another Database
'======================================================================================================================================================
Public Function ImportForm(sForm As String, sPath As String) As Boolean
Dim bRes As Boolean

On Error GoTo ErrHandle
'---------------------------
If sForm = "" Then Exit Function
    
    If Dir(sPath) = "" Then Err.Raise 10002, , "Can't read the external db: " & sPath
    If IsForm(sForm) Then ' The form is in this database already
        Debug.Print "the form " & sForm & " is already in the current database and does not require re-installation"
        bRes = True
    Else
        DoCmd.TransferDatabase acImport, "Microsoft Access", _
            sPath, acForm, sForm, sForm
        Wait 200
        bRes = IsForm(sForm)
    End If
    
'---------------------------
ExitHere:
    ImportForm = bRes '!!!!!!!!!
    Exit Function
'---------
ErrHandle:
    ErrPrint2 "ImportForm", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'======================================================================================================================================================
' Transfer Data to new DB
' PARAMETERS: sSQL    - SQL-statement with/without export params:
'                                     [sNewDB]  - full path to new database
'                                     [sNewTBL] - new table name in external db
'                                     [sWHERE]  - select params (without keyword WHERE)
'             sNewDB  - real path to new db
'             sNewTBL - real name to new table
' RETURNS   : True - if success
'======================================================================================================================================================
Public Function TransferSelectedData(sSQL As String, Optional sNewTBL As String = "", _
                                                Optional sNewDB As String = "", Optional sWhere As String = "") As Boolean
On Error GoTo ErrHandle
'---------------------------------------------------------------
' Set real value instead of sql params
 If InStr(1, sSQL, "[sNewTBL]") > 0 And sNewTBL <> "" Then sSQL = Replace(sSQL, "[sNewTBL]", sNewTBL)
 If InStr(1, sSQL, "[sNewDB]") > 0 And sNewDB <> "" Then sSQL = Replace(sSQL, "[sNewDB]", "'" & sNewDB & "'")
 If InStr(1, sSQL, "[sWHERE]") > 0 And sWhere <> "" Then sSQL = Replace(sSQL, "[sWHERE]", "WHERE (" & sWhere & ")")
'---------------------------------------------------------------
' Execute
 DoEvents
 CurrentDb.Execute sSQL
'---------------------------------------------------------------
ExitHere:
  TransferSelectedData = True '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  Exit Function
'---------------------------------------------------
ErrHandle:
  ErrPrint2 "TransferSelectedData", Err.Number, Err.Description, MOD_NAME
  Err.Clear
End Function

'===============================================================================================================================
' Import the specific module from external DB
'===============================================================================================================================
Public Function ImportVBAModuleFromLib(sModName As String, Optional sPath As String) As String
Dim sDBLibPath As String               ' Path to VBA Lib
Dim obj As AccessObject

On Error GoTo ErrHandle
'------------------------
    If sPath = "" Then
        sDBLibPath = OpenDialog(GC_FILE_PICKER, "Pick the new DB to export modules")
        If sDBLibPath = "" Then Exit Function
    Else
        sDBLibPath = sPath
    End If
    If Dir(sDBLibPath) = "" Then Err.Raise 10006, , "Can't find the DB: " & sDBLibPath
'---------------------------------------------------------
DoCmd.TransferDatabase acImport, "Microsoft Access", sDBLibPath, _
    acModule, sModName, sModName
'----------------------------------------------------------
ExitHere:
    ImportVBAModuleFromLib = sDBLibPath '!!!!!!!!!!!!!
    Exit Function
'--------------
ErrHandle:
    ErrPrint2 "ImportVBAModuleFromLib", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'======================================================================================================================================================
' Export all modules to external DB
'======================================================================================================================================================
Public Function ExportModulesToLib(Optional sModuleName As String = "", Optional sPath As String) As String
Dim sDBLibPath As String               ' Path to VBA Lib
Dim obj As AccessObject

On Error GoTo ErrHandle
'---------------------------
    If sPath = "" Then
        sDBLibPath = OpenDialog(GC_FILE_PICKER, "Pick the new DB to export modules")
        If sDBLibPath = "" Then Exit Function
    Else
        sDBLibPath = sPath
    End If
    If Dir(sDBLibPath) = "" Then Err.Raise 10006, , "Can't find the DB: " & sDBLibPath
'----------------------------------------------------------
If sModuleName = "" Then           ' Export all modules
    For Each obj In CurrentProject.AllModules
            DoCmd.TransferDatabase acExport, "Microsoft Access", sDBLibPath, _
                               acModule, obj.Name, obj.Name, False
    Next obj
Else                               ' Export specific module
    Set obj = CurrentProject.AllModules(sModuleName)
    DoCmd.TransferDatabase acExport, "Microsoft Access", sDBLibPath, _
                               acModule, obj.Name, obj.Name, False
End If
'------------------------------------
ExitHere:
    ExportModulesToLib = sDBLibPath '!!!!!!!!!!!!!
    Exit Function
'--------------
ErrHandle:
    ErrPrint2 "ExportModulesToLib", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'------------------------------------------------------------------------------------------------------------------------------------------------------
' The extension for different modules
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function vbExtFromType(ByVal ctype As Integer) As String
Const EXT_MODULE = ".bas"
Const EXT_CLASS = ".cls"
Const EXT_FORM = ".frm"
Const VB_MODULE = 1
Const VB_CLASS = 2
Const VB_FORM = 100

    Select Case ctype
        Case VB_MODULE
            vbExtFromType = EXT_MODULE
        Case VB_CLASS
            vbExtFromType = EXT_CLASS
        Case VB_FORM
            vbExtFromType = EXT_FORM
    End Select
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Get moduleName
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetModNameFromFile(sFileName As String, sProjectName As String, Optional DefaultModName As String = "_TEST", _
                                                                                                              Optional DLM As String = "--") As String
Dim PARTS() As String, iL As Integer, sModName As String, sRes As String

    On Error GoTo ErrHandle
'-----------------------------------
    iL = InStr(1, sFileName, DLM)
    If iL > 0 Then
        PARTS = Split(sFileName, DLM): sModName = Trim(PARTS(1))
    Else
        sModName = DefaultModName
    End If
        
        If IsVBAModule(sModName, sProjectName) Then
             sRes = sModName
        Else
             If CreateVBA(sModName, sProjectName) Then sRes = sModName
        End If
'-------------------------------------
ExitHere:
    GetModNameFromFile = sRes '!!!!!!!!!!!!!!
    Exit Function
'--------------
ErrHandle:
    ErrPrint2 "GetModNameFromFile", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Check if Folder exists and kill existing file
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub CheckFolderFile(Optional sFolder As String, Optional sFile As String)
On Error Resume Next
    If sFolder <> "" Then
       If Dir(sFolder, vbDirectory) = "" Then FolderCreate (sFolder)
    End If
    If sFile <> "" Then
       If Dir(sFile) <> "" Then Kill sFile
    End If
End Sub
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' XML Export
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function ProceedXMLExport(ObjectName As String, ObjectType As AcObjectType, sFolder As String) As String
Dim sXMLPath As String, sXSDPath As String, sXSVPath As String, sIMGPath As String

On Error GoTo ErrHandle
'-------------------------
    sXMLPath = sFolder & "\" & ObjectName & ".xml": Call CheckFolderFile(sFolder, sXMLPath)
    sXSDPath = sFolder & "\" & ObjectName & ".xsd": Call CheckFolderFile(, sXSDPath)

Select Case ObjectType
Case acTable:
    sIMGPath = sFolder & "\" & "FILES": Call CheckFolderFile(sIMGPath)
    Access.Application.ExportXML acExportTable, ObjectName, _
                     sXMLPath, sXSDPath, , , acUTF8, acExportAllTableAndFieldProperties
Case acQuery:
    Access.Application.ExportXML acExportQuery, ObjectName, sXMLPath
Case acForm:
    sXSVPath = sFolder & "\" & ObjectName & ".xvd": Call CheckFolderFile(, sXSVPath)
    Access.Application.ExportXML acExportForm, ObjectName, _
                  sXMLPath, sXSDPath, sXSVPath, , acUTF16
Case acReport:
    sXSVPath = sFolder & "\" & ObjectName & ".xvd": Call CheckFolderFile(, sXSVPath)
    Access.Application.ExportXML acExportReport, ObjectName, _
                  sXMLPath, sXSDPath, sXSVPath, , acUTF16
Case Else
    sXMLPath = ""
End Select
'-------------------------
ExitHere:
    ProceedXMLExport = sXMLPath '!!!!!!!!!!!!!!
    Exit Function
'--------------
ErrHandle:
    ErrPrint2 "ProceedXMLExport", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function


'-------------------------------------------------------------------------------------------------------------------------------------------
' Get Folders for Extras
'------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetFOLDEREXTRACT(sBaseFld As String, tblName As String) As String
Dim sRes As String
    
    sRes = sBaseFld & "\" & tblName & "_FILES"
    If Dir(sRes, vbDirectory) = "" Then
        If Not FolderCreate(sRes) Then sRes = ""
    End If
'---------------------------------------------
ExitHere:
   GetFOLDEREXTRACT = sRes '!!!!!!!
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------
' Function Form CODE Directory for Database
'-----------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetCodeFolder(sDBPath As String) As String
Dim sFolder As String, sRes As String
Const CODE_FOLDER As String = "CODE"
    If sDBPath = "" Then Exit Function
    
    sFolder = FolderNameOnly(sDBPath) & CODE_FOLDER
    If Dir(sFolder, vbDirectory) = "" Then
        If FolderCreate(sFolder) Then sRes = sFolder
    Else
        sRes = sFolder
    End If
'--------------------
ExitHere:
    GetCodeFolder = sRes '!!!!!!!!!!!!!!!!
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------
' Get Function Name
'-----------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetTheVBAFuncName(sCode As String) As String
Dim ROWS() As String, nRows As Integer, I As Integer
Dim sWork As String, iL As Integer, sName As String, sRes As String

    ROWS = Split(sCode, vbCrLf): nRows = UBound(ROWS)
    
    For I = 0 To nRows
        sWork = Trim(ROWS(I))
        
        If sWork <> "" Then
             If Left(sWork, 1) <> "'" Then
                    sRes = ExtractFuncName(sWork, "Function")
                    If sRes = "" Then sRes = ExtractFuncName(sWork, "Sub")
                    If sRes = "" Then sRes = ExtractFuncName(sWork, "Property")
                  
                    If sRes <> "" Then Exit For
             End If
        End If
    Next I
'-------------------------------
ExitHere:
   GetTheVBAFuncName = sRes '!!!!!!!!!!!!!!
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------------
' Function Extract the name from Row for coderow and keywords
'-----------------------------------------------------------------------------------------------------------------------------------------------------
Private Function ExtractFuncName(sRow As String, Optional sKeyWord As String = "Function") As String
Dim iL As Integer, sWork As String

    If sRow = "" Then Exit Function
    iL = InStr(1, sRow, sKeyWord, vbTextCompare): If iL = 0 Then Exit Function
    sWork = Trim(Right(sRow, Len(sRow) - iL - Len(sKeyWord))): If sWork = "" Then Exit Function
    iL = InStr(1, sWork, "("): If iL = 0 Then Exit Function
    sWork = Left(sWork, iL - 1)
'-----------------------------
ExitHere:
    ExtractFuncName = sWork '!!!!!!!!!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Check If VBA Module Is Class
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function IsClassModule(ModuleName As String) As Boolean

Dim iType As Integer

Const vbext_ct_ClassModule = 2
Const vbext_ct_Document = 100
Const vbext_ct_MSForm = 3
Const vbext_ct_StdModule = 1
   
   On Error Resume Next
'-----------------------------
  iType = VBE.ActiveVBProject.VBComponents(ModuleName).Type
  If (iType = vbext_ct_ClassModule) Or (iType = vbext_ct_Document) Then
          IsClassModule = True '!!!!!!!!!!!!!!!!
  End If
End Function


