VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAttachment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****************************************************************************************************************************************************
'****************************************************************************************************************************************************
' This class implements methods to work with attachment field
'****************************************************************************************************************************************************
'****************************************************************************************************************************************************
Option Explicit

Private m_Table As String
Private m_Field As String
Private m_Folder As String
Private m_FolderAuthomatic As Boolean

'****************************************************************************************************************************************************
'---------------------------------------------------------------------------------------------------------------------------------------------------
'  Get/Let table name to work with its attachments
'---------------------------------------------------------------------------------------------------------------------------------------------------
Public Property Get Table() As String
       Table = m_Table
End Property
Public Property Let Table(STABLE As String)
       m_Table = STABLE
End Property

'---------------------------------------------------------------------------------------------------------------------------------------------------
'  Get/Let field name for attachment one
'---------------------------------------------------------------------------------------------------------------------------------------------------
Public Property Get Field() As String
       Field = m_Field
End Property
Public Property Let Field(sField As String)
       m_Field = sField
End Property

'---------------------------------------------------------------------------------------------------------------------------------------------------
'  Get/Let folder path to save attchment
'---------------------------------------------------------------------------------------------------------------------------------------------------
Public Property Get Folder() As String
     If m_FolderAuthomatic Then BuildFolder
     Folder = m_Folder
End Property
Public Property Let Folder(sFolder As String)
     m_Folder = sFolder
     m_FolderAuthomatic = False
End Property

'---------------------------------------------------------------------------------------------------------------------------------------------------
'  Build Folder for Table
'---------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub BuildFolder(Optional IDD As Long = -1)
Dim sPath As String
   
   On Error GoTo ErrHandle
'------------------------
   sPath = CurrentProject.Path & "\DATA"
   If Dir(sPath, vbDirectory) = "" Then FolderCreate (sPath)
   sPath = sPath & "\" & m_Table & "_FILES"
   
   If Dir(sPath, vbDirectory) = "" Then FolderCreate (sPath)
   If IDD > 0 Then
       sPath = sPath & "\" & m_Field & "_" & IDD
       If Dir(sPath, vbDirectory) = "" Then FolderCreate (sPath)
   End If
'--------------------------
ExitHere:
   m_Folder = sPath  '!!!!!!!!!!!
   Exit Sub
'-------------
ErrHandle:
   ErrPrint "BuildFolder", Err.Number, Err.Description
   Err.Clear
End Sub

'================================================================================================================================================
' Check if provided field is attachment
'================================================================================================================================================
Public Function isAttachment() As Boolean
Dim sErr As String

On Error GoTo ErrHandle
'-----------------------------
ExitHere:
     isAttachment = (CurrentDb.TableDefs(m_Table).FIELDS(m_Field).Type = 101)
     Exit Function
'------------
ErrHandle:
     Select Case Err.Number
     Case 3265:
          sErr = "Wrong Table or Field Name"
     Case Else
          sErr = Err.Description
     End Select
     ErrPrint "IsAttachment", Err.Number, sErr
     Err.Clear
End Function

'================================================================================================================================================
' Add Attachment From Clipboard
'================================================================================================================================================
Public Function AttachFromClipboard(IDD As Long) As Boolean
Dim bRes As Boolean, RS As DAO.Recordset, rsAttach As DAO.Recordset
Dim TempPath As String, iCount As Integer

Const sExt As String = "PNG"
Const PRFX As String = "attch"

On Error GoTo ErrHandle
'-------------------------------------------
        If IDD <= 0 Then Exit Function
        TempPath = GetTempFile(PRFX, sExt): If TempPath = "" Then Exit Function
        TempPath = ImageFromClipboard(TempPath): If Dir(TempPath) = "" Then Exit Function
        
        Set RS = CurrentDb.OpenRecordset(m_Table)
        With RS
                .MoveLast: .MoveFirst
                .Index = "PrimaryKey"
                .Seek "=", IDD
                If .NoMatch Then
                   GoTo ExitHere
                Else
                   .Edit
                   Set rsAttach = .FIELDS(m_Field).value
                   If Not rsAttach.EOF Then
                             rsAttach.MoveLast:  rsAttach.MoveFirst
                   End If
                             iCount = rsAttach.RecordCount
                             rsAttach.AddNew
                             rsAttach.FIELDS("FileData").LoadFromFile (TempPath)
                             rsAttach.Update
                                   If rsAttach.RecordCount > iCount Then
                                       bRes = True
                                   End If
                             rsAttach.Close
                  
                End If
                .Update
                .Close
        End With
'-------------------------------------------
If Dir(TempPath) <> "" Then Kill TempPath
'-------------------------------------------
ExitHere:
        AttachFromClipboard = bRes '!!!!!!!!!!!!
        Set rsAttach = Nothing: Set RS = Nothing
        Exit Function
'----------------
ErrHandle:
        ErrPrint "AttachFromClipBoard", Err.Number, Err.Description
        Err.Clear: Resume ExitHere
End Function
'=================================================================================================================================================
' Save attachments to file, returns path to folder with files
'=================================================================================================================================================
Public Function SaveAttachment(Optional IDD As Long = -1, Optional sPath As String, Optional bClear As Boolean, _
                                                                                                           Optional DLM As String = ";") As String
Dim strFileName As String, strPath As String, SQL As String, sRes As String
Dim RS As DAO.Recordset2, rst As DAO.Recordset2


On Error GoTo ErrHandle
'------------------------------------------------------------
If IDD < 0 Then Exit Function
If m_FolderAuthomatic Then BuildFolder (IDD)
strPath = IIf(sPath = "", m_Folder, sPath)

SQL = "SELECT " & m_Field & " FROM " & m_Table & " WHERE(ID = " & IDD & ");"
Set RS = CurrentDb.OpenRecordset(SQL)
If Not RS.EOF Then
   Set rst = RS.FIELDS(m_Field).value
   With rst
    If .RecordCount > 0 Then .MoveFirst
     Do While Not .EOF
        strFileName = strPath & "\" & rst.FIELDS("FileName").value
        If Dir(strFileName) <> "" Then Kill strFileName
        
        .FIELDS("FileData").SaveToFile strFileName
        If Dir(strFileName) <> "" Then sRes = IIf(sRes <> "", sRes & DLM, "") & strFileName
        If bClear Then .Delete
        
        .MoveNext
     Loop
    End With
End If
'------------------------------
ExitHere:
    SaveAttachment = sRes '!!!!!!!!!!
    Set rst = Nothing: Set RS = Nothing
    Exit Function
'---------------------
ErrHandle:
    ErrPrint "SaveAttachment" & "(TBL = " & m_Table & "; AttachFLD =" & m_Field & "; ID =" & IDD, Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'================================================================================================================================================
' Get Attachment List
'================================================================================================================================================
Public Function AttachmentList(IDD As Long, Optional DLM As String = ";") As String
Dim RS As DAO.Recordset, rst As DAO.Recordset, SQL As String
Dim sRes As String

   On Error GoTo ErrHandle
'-----------------------
   If IDD < 0 Then Exit Function
   
   SQL = "SELECT " & m_Field & " FROM " & m_Table & " WHERE(ID = " & IDD & ");"
   Set RS = CurrentDb.OpenRecordset(SQL)
   Set rst = RS.FIELDS(m_Field).value
   With rst
       If Not .EOF Then
           .MoveLast: .MoveFirst
           Do While Not .EOF
            sRes = sRes & .FIELDS("FileName") & DLM
            .MoveNext
           Loop
       End If
   End With
 
If sRes <> "" Then sRes = Left(sRes, Len(sRes) - Len(DLM))
'------------------------
ExitHere:
   AttachmentList = sRes '!!!!!!!!!!!!!!!
   Set rst = Nothing: Set RS = Nothing
   Exit Function
'--------
ErrHandle:
   ErrPrint "AttachmentList", Err.Number, Err.Description
   Err.Clear
End Function
'================================================================================================================================================
' Attachment Count
'================================================================================================================================================
Public Function AttachmentCount(IDD As Long) As Integer
Dim RS As DAO.Recordset, rst As DAO.Recordset, SQL As String
Dim iRes As Integer

   On Error GoTo ErrHandle
'-----------------------
   If IDD < 0 Then Exit Function

   SQL = "SELECT " & m_Field & " FROM " & m_Table & " WHERE(ID = " & IDD & ");"
   Set RS = CurrentDb.OpenRecordset(SQL)
   Set rst = RS.FIELDS(m_Field).value: rst.MoveFirst: rst.MoveLast
   iRes = rst.RecordCount
'------------------------
ExitHere:
   AttachmentCount = iRes '!!!!!!!!!!!!!!!
   Set rst = Nothing: Set RS = Nothing
   Exit Function
'--------
ErrHandle:
   ErrPrint "AttachmentCount", Err.Number, Err.Description
   Err.Clear
End Function
'================================================================================================================================================
' Add Files to attchments. Return count of added files
'================================================================================================================================================
Public Function AddAttachment(IDD As Long, Optional Files As Variant, Optional DLM As String = ";") As Integer
Dim nFiles As Integer, iRes As Integer, FLS() As String, I As Integer, sWork As String
Dim RS As DAO.Recordset

On Error GoTo ErrHandle
'-----------------------------
     If Not isAttachment Then Exit Function
     If IDD <= 0 Then Exit Function
     
     If IsMissing(Files) Then
        sWork = OpenDialog(GC_FILE_PICKER, "Add Files", , True, CurrentProject.Path)
        If sWork = "" Then Exit Function
        FLS = BuildArray(sWork, DLM)
     Else
        FLS = BuildArray(Files, DLM)
     End If
     If FLS(0) = "" Then Exit Function
     nFiles = UBound(FLS)
'-----------------------------
     sWork = "SELECT " & m_Field & " FROM " & m_Table & " WHERE(ID = " & IDD & ");"
     Set RS = CurrentDb.OpenRecordset(sWork)
     If Not RS.EOF Then
            For I = 0 To nFiles
                  If Dir(FLS(I)) <> "" Then
                        If AddFile(RS, m_Field, FLS(I)) Then iRes = iRes + 1
                  End If
            Next I
     End If
'-----------------------------
ExitHere:
     Set RS = Nothing
     AddAttachment = iRes '!!!!!!!!!!!!!!
     Exit Function
'------------------
ErrHandle:
     ErrPrint "AddAttachment", Err.Number, Err.Description
     Err.Clear: Resume ExitHere
     
End Function

'================================================================================================================================================
' Function save all attachments to file and move it as a file link to another table + clear source table and reduce size of database
' (only for single record). Return the ID of new table records.
'================================================================================================================================================
Public Function ShiftAll(Optional sPath As String, Optional FileTable As String = "NOTES", Optional TitleField As String = "Title", _
               Optional PathField As String = "LINK", Optional ParentField As String = "ParentItem", Optional TypeField As String = "NoteType", _
                                                       Optional TypeValue As Integer = 1, Optional bInsertAsLink As Boolean = True) As Integer
Dim nRes As Integer
Dim RS As DAO.Recordset, IDD As Variant, I As Integer, nDim As Integer

On Error GoTo ErrHandle
'-----------------------------
    IDD = GETATTACHIDs(m_Table, "ID", m_Field, 1)
    If Not IsArray(IDD) Then GoTo ExitHere
    nDim = UBound(IDD, 2)
    
     For I = 0 To nDim
              nRes = nRes + _
              UBound(ShiftAttachment(CLng(IDD(0, I)), FileTable, TitleField, PathField, ParentField, TypeField, TypeValue, bInsertAsLink)) + 1
     Next I
'-----------------------------
ExitHere:
     ShiftAll = nRes '!!!!!!!!!
     Set RS = Nothing
     Exit Function
'-----------
ErrHandle:
    ErrPrint "ShiftAll", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------
' Get IDs for attachment field with records
'------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GETATTACHIDs(Optional STABLE As String = "ITEMS", Optional sIDFld As String = "ID", _
                                                Optional sAttachFld As String = "Attachments", Optional MinAttachCount As Integer = 1) As Variant
Dim RS As DAO.Recordset, SQL As String
Dim vARR As Variant

On Error GoTo ErrHandle
'------------------------------------
   SQL = "SELECT t." & sIDFld & _
         " FROM " & STABLE & " AS t GROUP BY t." & sIDFld & _
         " HAVING (((Count(t." & sAttachFld & ".FileName))>" & (MinAttachCount - 1) & "));"
         
   Set RS = CurrentDb.OpenRecordset(SQL)
   With RS
       .MoveLast: .MoveFirst
       vARR = .GetRows(.RecordCount)
   End With
'------------------------------------
ExitHere:
    GETATTACHIDs = vARR '!!!!!!!!!!!!!!!!!
    Set RS = Nothing
    Exit Function
'--------------------
ErrHandle:
    ErrPrint "GETATTACHIDs", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------
' Add File Link To Notes
'------------------------------------------------------------------------------------------------------------------------------------------------
Private Function AddFileLink(ParentIDD As Long, sPath As String, Optional FileTable As String = "NOTES", Optional TitleField As String = "Title", _
              Optional PathField As String = "LINK", Optional ParentField As String = "ParentItem", Optional TypeField As String = "NoteType", _
                                                       Optional TypeValue As Integer = 1, Optional bInsertAsLink As Boolean = True) As Long
Dim iRes As Long, RS As DAO.Recordset, sHASH As String

On Error GoTo ErrHandle
'------------------------------
      Set RS = CurrentDb.OpenRecordset(FileTable)
      With RS
            .AddNew
                If IsFieldInRS(RS, PathField) Then
                    .FIELDS(PathField).value = IIf(bInsertAsLink, "file:///" & sPath, sPath)
                Else
                    Err.Raise 1000, , "WRONG FILE PATH FIELD"
                End If
                '-------------------------------
                sHASH = GetHASH("FILE")
                .FIELDS("HASH") = sHASH
                                
                If IsFieldInRS(RS, TitleField) Then .FIELDS(TitleField).value = FileNameOnly(sPath)
                If IsFieldInRS(RS, TypeField) Then .FIELDS(TypeField).value = TypeValue
                If IsFieldInRS(RS, ParentField) Then .FIELDS(ParentField).value = ParentIDD
            .Update
            '--------------------------------------------------------
      End With
            
            iRes = Nz(DLookup("ID", FileTable, "HASH = " & sCH(sHASH)), 0)
'------------------------------
ExitHere:
        AddFileLink = iRes '!!!!!!!!
        Set RS = Nothing
        Exit Function
'-----------
ErrHandle:
        ErrPrint "AddFileLink", Err.Number, Err.Description
        Err.Clear: Resume ExitHere
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------
' Build FilePath as string
'------------------------------------------------------------------------------------------------------------------------------------------------
Public Function BuildFilePath(JustFileName As String, Optional sFolder As String, Optional IDD As Long) As String
Dim strPath As String
If m_FolderAuthomatic Then BuildFolder (IDD)
       strPath = IIf(sFolder = "", m_Folder, sFolder)
'------------------------------
ExitHere:
       BuildFilePath = strPath '!!!!!!!!!!!!!!!!!!
End Function

'================================================================================================================================================
' Function save attachment to file and move it as a file link to another table + clear source table and reduce size of database
' (only for single record). Return the ID of new table record.
'================================================================================================================================================
Public Function ShiftAttachment(IDD As Long, Optional FileTable As String = "NOTES", Optional TitleField As String = "Title", _
               Optional PathField As String = "LINK", Optional ParentField As String = "ParentItem", Optional TypeField As String = "NoteType", _
                                                       Optional TypeValue As Integer = 1, Optional bInsertAsLink As Boolean = True) As Long()
Dim iRes() As Long, sFiles As String, FLS() As String, I As Integer, nDim As Integer
Dim RS As DAO.Recordset, sHASH As String

Const DLM As String = ";"
On Error GoTo ErrHandle
'--------------------------------
If IDD <= 0 Then Exit Function
sFiles = SaveAttachment(IDD, , True, DLM)
If sFiles <> "" Then
      If InStr(1, sFiles, DLM) > 0 Then
            FLS = Split(sFiles, DLM): nDim = UBound(FLS)
      Else
            nDim = 0: ReDim FLS(0): FLS(0) = sFiles
      End If
      ReDim iRes(nDim)
   
      Set RS = CurrentDb.OpenRecordset(FileTable)
      With RS
          For I = 0 To nDim
            .AddNew
                If IsFieldInRS(RS, PathField) Then
                   .FIELDS(PathField).value = IIf(bInsertAsLink, "file:///" & FLS(I), FLS(I))
                Else
                   Err.Raise 1000, , "WRONG FILE PATH FIELD"
                End If
                '-------------------------------
                sHASH = GetHASH("FILE")
                .FIELDS("HASH") = sHASH
                                
                If IsFieldInRS(RS, TitleField) Then .FIELDS(TitleField).value = FileNameOnly(sFiles)
                If IsFieldInRS(RS, TypeField) Then .FIELDS(TypeField).value = TypeValue
                If IsFieldInRS(RS, ParentField) Then .FIELDS(ParentField).value = IDD
            .Update
            '--------------------------------------------------------
            iRes(I) = Nz(DLookup("ID", FileTable, "HASH = " & sCH(sHASH)), 0)
          Next I
      End With
End If
'--------------------------------
ExitHere:
      ShiftAttachment = iRes '!!!!!!!!!!!!!!!!!!!!!!!
      Set RS = Nothing
      Exit Function
'------------------
ErrHandle:
      ErrPrint "ShiftAttachments", Err.Number, Err.Description
      Err.Clear: Resume ExitHere
End Function


'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Check if recordset does have some field with specific name
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function IsFieldInRS(ByRef RS As DAO.Recordset, sFieldName As String) As Boolean

On Error GoTo ErrHandle
'-------------------------
ExitHere:
    IsFieldInRS = (RS.FIELDS(sFieldName).Name = sFieldName) '!!!!!!!!!
    Exit Function
'---------------
ErrHandle:
    Select Case Err.Number
    Case 3265:
           Err.Clear
    Case Else
           ErrPrint "IsFieldInRS", Err.Number, Err.Description
    End Select
End Function


'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Function add some file to attachment fiels, return True for success
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function AddFile(ByRef RS As DAO.Recordset, sFldName As String, sFile As String) As Boolean
Dim bRes As Boolean, rst As DAO.Recordset
On Error GoTo ErrHandle
'---------------------------
    RS.Edit
        Set rst = RS.FIELDS(sFldName).value
        rst.AddNew
           rst.FIELDS("FileData").LoadFromFile sFile
        rst.Update
    RS.Update
    bRes = True
'---------------------------
ExitHere:
    AddFile = bRes '!!!!!!!!!!!!!
    Set rst = Nothing
    Exit Function
ErrHandle:
    ErrPrint "AddFile", Err.Number, Err.Description
    Err.Clear
End Function


'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Return String Array from Dynamic array or string with possible delimeters
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function BuildArray(Arg As Variant, Optional DLM As String = ";") As String()
Dim sARR() As String, nArr As Integer

On Error GoTo ErrHandle
'-------------------------
    ReDim sARR(0): nArr = 0
    If IsArray(Arg) Then
          sARR = Arg
    ElseIf varType(Arg) = vbString Then
          If InStr(1, Arg, DLM) > 0 Then
               sARR = Split(CStr(Arg), DLM)
          Else
               sARR(0) = CStr(Arg)
          End If
    End If
'-------------------------
ExitHere:
    BuildArray = sARR '!!!!!!!!!!
    Exit Function
'-----------
ErrHandle:
   ErrPrint "BuildArray", Err.Number, Err.Description
   Err.Clear: Resume ExitHere
End Function

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
    m_Table = "ITEMS": m_Field = "Attachments"
    m_FolderAuthomatic = True
End Sub
Private Sub Class_Terminate()
   On Error Resume Next

End Sub
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Error Handler
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function ErrPrint(FuncName As String, ErrNumber As Long, ErrDescription As String, Optional bDebug As Boolean = True, _
                                                                                                  Optional sModName As String = "cAttachment") As String
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


