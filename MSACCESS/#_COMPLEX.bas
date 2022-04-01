Attribute VB_Name = "#_COMPLEX"
'********************************************************************************************************************************************************************************
'********************************************************************************************************************************************************************************
'********************************************************************************************************************************************************************************
' MVF AND ATTACHMENTS
'********************************************************************************************************************************************************************************
'********************************************************************************************************************************************************************************
'********************************************************************************************************************************************************************************
Option Compare Database
Option Explicit


Private Const MOD_NAME As String = "#_COMPLEX"
'*********************************

'=================================================================================================================================================
' Func save complex field to disk or return some list if MVF
'=================================================================================================================================================
Public Function SaveComplex(sTBL As String, sFLD As String, iFldType As Integer, IDD As Long, Optional sROOT As String, _
                                                                           Optional sIDFld As String = "ID", Optional DLM As String = ";") As String
Dim sRes As String, sWork As String

On Error GoTo ErrHandle
'-------------------------------
     Select Case iFldType
     Case 101:      ' Attachment
            sWork = AttachmentToDisk(sTBL, sFLD, IDD, sROOT, sIDFld)
            If sWork <> "" Then sRes = SHT(sWork)
     Case 102, 103, 104, 105, 106, 107:   'MVF
            sWork = MVFToList(sTBL, sFLD, IDD, DLM)
            sRes = "{" & sWork & "}"
     Case 9, 11:                          ' BINARY
            sWork = SaveBinaryFromDB(sTBL, sFLD, IDD, sROOT, sIDFld)
            If sWork <> "" Then sRes = SHT(sWork)
     Case 12:                             ' MEMO/HYPERLINK
           sWork = SaveTextFromDB(sTBL, sFLD, IDD, sROOT, sIDFld)
           If sWork <> "" Then sRes = "{" & sWork & "}"
     End Select
'-------------------------------
ExitHere:
    SaveComplex = sRes '!!!!!!!!!!!!
    Exit Function
'---------
ErrHandle:
    ErrPrint "SaveComplex", Err.Number, Err.Description
    Err.Clear
End Function


'=================================================================================================================================================
' Chek if fld require some special savements
'=================================================================================================================================================
Public Function IsComplexFld(Optional FldName As String, Optional TBL As String, Optional FldType As Integer = -1) As Integer
Dim iRes As Integer, iType  As Integer

On Error GoTo ErrHandle
'-------------------------------
      If FldType >= 0 Then
            iType = FldType
      Else
            iType = CurrentDb.TableDefs(TBL).FIELDS(FldName).Type
      End If
      '----------------------------------------------------
      Select Case iType
      Case 101:  ' Attachments               ' ATTACH
           iRes = iType
      Case 102, 103, 104, 105, 106, 107:   'MVF
           iRes = iType
      Case 9, 11:                          ' BINARY
           iRes = iType
      Case 12:                             ' MEMO/HYPERLINK
           iRes = iType
      Case Else
           iRes = 0
      End Select
'-------------------------------
ExitHere:
      IsComplexFld = iRes   '!!!!!!!!!!!!!!
      Exit Function
'----------
ErrHandle:
      Err.Clear
End Function
'=================================================================================================================================================
' Count of Attachments
'=================================================================================================================================================
Public Function AttachmentCount(TableName As String, Field As String, WhereClause As String)
    Dim rsRecords As DAO.Recordset, rsAttach As DAO.Recordset

    AttachmentCount = 0

    Set rsRecords = CurrentDb.OpenRecordset("SELECT * FROM [" & TableName & "] WHERE " & WhereClause, dbOpenDynaset)
    If rsRecords.EOF Then Exit Function

    Set rsAttach = rsRecords.FIELDS(Field).value
    If rsAttach.EOF Then Exit Function

    rsAttach.MoveLast
    rsAttach.MoveFirst
'------------------------------------------------
    AttachmentCount = rsAttach.RecordCount
End Function



'=================================================================================================================================================
' The function loads the attachment from the disk into the database according to some condition. If the attachment path is empty,
' then a file selection dialog is launched
' For correct operation, it is necessary that the share record is unique, otherwise the download goes to the very first record
' Returns a list of downloaded files
'=================================================================================================================================================
Public Function AttchmentLoad(TBL As String, AttachFld As String, Optional sWhere As String, Optional sFiles As String, _
                                                                                                          Optional DLM As String = ";") As String
Dim sFLS As String, FLS() As String, nFLS As Integer, sRes As String
Dim RS As DAO.Recordset, rst As DAO.Recordset, sFilePath As String, SQL As String, I As Integer

    On Error GoTo ErrHandle
'--------------------
    If IsBlank(TBL) Or IsBlank(AttachFld) Then Exit Function
    sFLS = sFiles: If IsBlank(sFLS) Then sFLS = OpenDialog(GC_FILE_PICKER, "Please select file(s)", "All Files,*.*", , GetLastFolder)
    If IsBlank(sFLS) Then Exit Function
    FLS = Split(sFLS, DLM): nFLS = UBound(FLS)
    
    If Not IsBlank(sWhere) Then
        SQL = "SELECT " & AttachFld & " FROM " & TBL & " WHERE (" & sWhere & ");"
    Else
        SQL = "SELECT " & AttachFld & " FROM " & TBL & ";"
    End If
    Set RS = CurrentDb.OpenRecordset(SQL)
    With RS
          If Not .EOF Then
               .MoveLast: .MoveFirst
               Set rst = .FIELDS(AttachFld).value
               
               For I = 0 To nFLS
                     If Not IsBlank(FLS(I)) Then
                            .Edit
                                rst.AddNew
                                    rst.FIELDS("FileData").LoadFromFile (FLS(I))
                                rst.Update
                                sRes = sRes & DLM & FLS(I)
                            .Update
                     End If
               Next I
               'rst.Close
          End If
    End With
    
    If Not IsBlank(sRes) Then sRes = Right(sRes, Len(sRes) - Len(DLM))
'--------------------
ExitHere:
    AttchmentLoad = sRes '!!!!!!!!!!!!!!!
    Set rst = Nothing: Set RS = Nothing
    Exit Function
'---------
ErrHandle:
    ErrPrint2 "AttchmentLoad", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Resume ExitHere
End Function
'=================================================================================================================================================
' function for saving the attached file to disk Similar to function 1, but with looser selection criteria - sWhere
' If the path is not specified, the download occurs directly to the directory of the current database.  Replace existing file
' If successful - returns a list of saved files
'=================================================================================================================================================
Public Function AttachmentSave(TBL As String, AttachFld As String, Optional sWhere As String, Optional ToFolder As String, _
                                                                                                          Optional DLM As String = ";") As String
Dim sARR() As String, nDim As Integer, sRes As String, sFolder As String
Dim RS As DAO.Recordset, rst As DAO.Recordset, sFilePath As String, SQL As String

    On Error GoTo ErrHandle
'--------------------
    If IsBlank(TBL) Or IsBlank(AttachFld) Then Exit Function
    
    nDim = -1: ReDim sARR(0)
    sFolder = ToFolder: If sFolder = "" Then sFolder = CurrentProject.Path
    
    If Not IsBlank(sWhere) Then
        SQL = "SELECT " & AttachFld & " FROM " & TBL & " WHERE (" & sWhere & ");"
    Else
        SQL = "SELECT " & AttachFld & " FROM " & TBL & ";"
    End If
    
    Set RS = CurrentDb.OpenRecordset(SQL)
    With RS
        If Not .EOF Then
            .MoveLast: .MoveFirst
            Do While Not .EOF
                Set rst = .FIELDS(AttachFld).value
                If Not rst.EOF Then
                       rst.MoveLast:  rst.MoveFirst
                       Do While Not rst.EOF
                           sFilePath = Nz(rst.FIELDS("FileName").value)
                           If sFilePath <> "" Then
                               sFilePath = sFolder & "\" & sFilePath
                               If Dir(sFilePath) <> "" Then Kill sFilePath
                               rst.FIELDS("FileData").SaveToFile sFilePath
                               If Not IsBlank(Dir(sFilePath)) Then
                                    nDim = nDim + 1: ReDim Preserve sARR(nDim)
                                    sARR(nDim) = sFilePath
                               End If
                           End If
                           rst.MoveNext
                       Loop
                End If
                .MoveNext
            Loop
        End If
    End With
    
    If nDim >= 0 Then sRes = Join(sARR, DLM)
'--------------------
ExitHere:
    AttachmentSave = sRes '!!!!!!!!!!!!!!!
    Set rst = Nothing: Set RS = Nothing
    Exit Function
'---------
ErrHandle:
    Debug.Print "AttachmentSaveToFiles", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Resume ExitHere
End Function
'=================================================================================================================================================
' Save attachments to file, returns path to folder with files
'=================================================================================================================================================
Public Function AttachmentToDisk(strTableName As String, strAttachmentField As String, Optional IDD As Long = -1, _
                             Optional sROOT As String, Optional strPrimaryKeyFieldName As String = "ID", Optional DLM As String = "¤") As String
Dim strFileName As String, strPath As String, SQL As String, sRes As String
Dim db As DAO.Database, rsParent As DAO.Recordset2, rsChild As DAO.Recordset2, FLD As DAO.Field2

Const DFLTPATH As String = "FILES"

On Error GoTo ErrHandle
'------------------------------------------------------------
strPath = AdjustFilePath(sROOT): If strPath = "" Then Exit Function
strPath = strPath & strTableName & "_" & strAttachmentField & IIf(IDD > 0, "-" & IDD, "") & "\"
If Dir(strPath, vbDirectory) = "" Then MkDir strPath
'------------------------------------------------------------
SQL = "SELECT " & strPrimaryKeyFieldName & ", " & strAttachmentField & " FROM " & strTableName
SQL = IIf(IDD > -1, SQL & " WHERE " & strPrimaryKeyFieldName & " = " & IDD, SQL)
    Set db = CurrentDb: Set rsParent = db.OpenRecordset(strTableName, dbOpenSnapshot)
    With rsParent
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            Set rsChild = rsParent(strAttachmentField).value
                 If rsChild.RecordCount > 0 Then rsChild.MoveFirst
                        
                        Do While Not rsChild.EOF
                                Set FLD = rsChild("FileData")
                                strFileName = strPath & rsChild("FileName")
                                If Len(Dir(strFileName)) <> 0 Then Kill strFileName
                                FLD.SaveToFile strFileName
                                
                                sRes = sRes & strFileName & DLM
                                
                            rsChild.MoveNext
                         Loop
            .MoveNext
        Loop
    End With
If sRes <> "" Then sRes = Left(sRes, Len(sRes) - Len(DLM))
'-------------------------------------------------------------
ExitHere:
    AttachmentToDisk = sRes '!!!!!!!!!!
    Set FLD = Nothing: Set rsChild = Nothing: Set rsParent = Nothing
    Set db = Nothing
    Exit Function
'---------------------
ErrHandle:
    ErrPrint "AttachmentToDisk" & "(TBL = " & strTableName & "; AttachFLD =" & strAttachmentField & "; ID =" & IDD, Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function

Public Sub TEST_MVF_List()
Dim iRes() As Long, sRes As String
    
    iRes = Get_MVF_VALUES(3, "TEST", "MVF")
    Debug.Assert False
    sRes = Join(iRes, ";")
End Sub

Public Sub TestAddValue()
Dim AddVals(1) As Long, bRes As Boolean

AddVals(0) = 2: AddVals(1) = 3
bRes = Add_To_MVF(AddVals, 5, "TEST", "MVF")
Debug.Assert False


End Sub
'=================================================================================================================================================
' Getting a text list for MVF Field
'=================================================================================================================================================
Public Function MVF_To_String(IDD As Long, TBL As String, MVFFLD As String, Optional iFormat As Integer, _
                                                                            Optional DLM As String = ";", Optional SEQV As String = "=") As String
Dim INDXS() As Long, nDim As Long, sRes As String, RS As DAO.Recordset, sVal() As String, J As Integer, sLoc As String

Dim sMVF_Source As String, sSQLParce As String, STABLE As String, sFld1 As String, sFld2 As String, sValues As String


On Error GoTo ErrHandle
'--------------------------------
    sMVF_Source = GetFieldLookUpInfo(TBL, MVFFLD, DLM, SEQV): If sMVF_Source = "" Then Exit Function
    sValues = MVFToList(TBL, MVFFLD, IDD, "ID", DLM):     If sValues = "" Then Exit Function
    sVal = Split(sValues, DLM)
    
    If GetValueForKey(sMVF_Source, "RowSourceType") = "Table/Query" Then         ' Table/Query
        
        sSQLParce = GetValueForKey(sMVF_Source, "RowSource", DLM, SEQV)
        
        If sSQLParce <> "" Then
             Set RS = CurrentDb.OpenRecordset(sSQLParce)
             With RS
                 If Not .EOF Then
                     .MoveLast: .MoveFirst
                     Do While Not .EOF
                         If IsValueInArr(sVal, Nz(.FIELDS(0).value)) Then
                                  
                                  Select Case iFormat
                                  Case 0:     '(ID=1,Field1=LBL1;ID=2,Field1=LBL2;ID=3,Field1=LBL3)
                                         For J = 0 To .FIELDS.Count - 1
                                                sLoc = IIf(sLoc <> "", sLoc & ",", "") & _
                                                      .FIELDS(J).Name & SEQV & .FIELDS(J).value
                                         Next J
                                  Case 1:     '(LBL1 = 1;LBL2 = 2;LBL3 = 3)
                                         If .FIELDS.Count >= 2 Then
                                                sLoc = IIf(sLoc <> "", sLoc & ",", "") & _
                                                      .FIELDS(1).value & SEQV & .FIELDS(0).value
                                         End If
                                  Case 2:     '(LBL1;LBL2;LBL3)
                                         If .FIELDS.Count >= 2 Then
                                                sLoc = IIf(sLoc <> "", sLoc & ",", "") & _
                                                      .FIELDS(1).value
                                         Else
                                                sLoc = IIf(sLoc <> "", sLoc & ",", "") & _
                                                      .FIELDS(0).value
                                         End If
                                  End Select
      
                         End If
                         If sLoc <> "" Then
                            sRes = IIf(sRes <> "", sRes & DLM, "") & sLoc: sLoc = ""
                         End If
                         
                         .MoveNext
                     Loop
                 End If
             End With
        End If
    Else                                                                        ' Value List
    End If
'--------------------------------
ExitHere:
     MVF_To_String = sRes '!!!!!!!!!!!!!!!
     Set RS = Nothing
     Exit Function
'-----------
ErrHandle:
     ErrPrint "MVF_To_String", Err.Number, Err.Description
     Err.Clear: Set RS = Nothing
End Function

'------------------------------------------------------------------------------------------------------------------------------------------
' Check If Value in Array
'------------------------------------------------------------------------------------------------------------------------------------------
Private Function IsValueInArr(vARR As Variant, vProbe As Variant) As Boolean
Dim nDim As Integer, I As Integer
    nDim = UBound(vARR)
    For I = 0 To nDim
       If CStr(vARR(I)) = CStr(vProbe) Then
            IsValueInArr = True
            Exit Function
       End If
    Next I
End Function

'=================================================================================================================================================
' The function Clear MVF
'=================================================================================================================================================
Public Function Clear_MVF(IDD As Long, TBL As String, MVFFLD As String, Optional iValue As Long) As Boolean
Dim SQL As String
On Error GoTo ErrHandle
'-----------------------------
If iValue > 0 Then
    SQL = "DELETE " & TBL & "." & MVFFLD & ".value FROM " & TBL & _
                    " WHERE((ID = " & IDD & ") AND (" & MVFFLD & ".value = " & iValue & "));"
Else
    SQL = "DELETE " & TBL & "." & MVFFLD & ".value FROM " & TBL & _
                      " WHERE(ID = " & IDD & ");"
End If
DoCmd.SetWarnings False
    CurrentDb.Execute SQL
'-----------------------------
ExitHere:
    Clear_MVF = True '!!!!!!!!!!!!!!
    DoCmd.SetWarnings True
    Exit Function
'-----------
ErrHandle:
    ErrPrint "Clear_MVF", Err.Number, Err.Description
    Err.Clear: DoCmd.SetWarnings False
End Function
'=================================================================================================================================================
' The function add some indexes (array) to MVF
'=================================================================================================================================================
Public Function Add_To_MVF(AddingValues() As Long, IDD As Long, TBL As String, MVFFLD As String) As Boolean
Dim RS As DAO.Recordset, rst As DAO.Recordset
Dim I As Long, nDim As Long, bUpdate As Boolean

On Error GoTo ErrHandle
'-----------------------------
nDim = UBound(AddingValues)

   Set RS = CurrentDb.OpenRecordset("SELECT " & MVFFLD & " FROM " & TBL & " WHERE(ID = " & IDD & ");")
   With RS
        If Not .EOF Then
            Set rst = .FIELDS(MVFFLD).value
            If Not rst.EOF Then
                rst.MoveLast: rst.MoveFirst
                For I = 0 To nDim
                    If AddingValues(I) <> 0 Then
                          If Not Check_IFinMVF(AddingValues(I), rst) Then
                            .Edit
                                bUpdate = True
                                rst.AddNew
                                    rst.FIELDS(0).value = AddingValues(I)
                                rst.Update
                          End If
                     End If
                Next I
            Else
                For I = 0 To nDim
                    If AddingValues(I) <> 0 Then
                            .Edit
                                bUpdate = True
                                rst.AddNew
                                    rst.FIELDS(0).value = AddingValues(I)
                                rst.Update
                    End If
                Next I
            End If
            If bUpdate Then .Update
        End If
   End With
'-----------------------------
ExitHere:
    Add_To_MVF = True '!!!!!!!!!!!!!!
    Set rst = Nothing: Set RS = Nothing
    Exit Function
'-----------
ErrHandle:
    ErrPrint "Add_To_MVF", Err.Number, Err.Description
    Err.Clear
End Function

'=================================================================================================================================================
' The function remove some index from MVF
'=================================================================================================================================================
Public Function Remove_From_MVF(RemoveValue As Long, IDD As Long, TBL As String, MVFFLD As String) As Boolean
Dim RS As DAO.Recordset, rst As DAO.Recordset, bRes As Boolean
Dim I As Long, nDim As Long

On Error GoTo ErrHandle
'-----------------------------
   Set RS = CurrentDb.OpenRecordset("SELECT " & MVFFLD & " FROM " & TBL & " WHERE(ID = " & IDD & ");")
   With RS
        If Not .EOF Then
            Set rst = .FIELDS(MVFFLD).value
            If Not rst.EOF Then
                    If Check_IFinMVF(RemoveValue, rst) Then
                          If rst.FIELDS(0).value = RemoveValue Then
                                    rst.Delete
                          End If
                    End If
                    
            End If
      End If
   End With
'-----------------------------
ExitHere:
    Remove_From_MVF = bRes '!!!!!!!!!!!!!!
    Set rst = Nothing: Set RS = Nothing
    Exit Function
'-----------
ErrHandle:
    ErrPrint "Remove_From_MVF", Err.Number, Err.Description
    Err.Clear
End Function

'=================================================================================================================================================
' The function check if some index placed in MVF
'=================================================================================================================================================
Public Function Check_IFinMVF(ByVal CheckedIndex As Long, ByRef MVFRS As Recordset) As Boolean
Dim bRes As Boolean

On Error GoTo ErrHandle
'-----------------------------
     MVFRS.FindFirst MVFRS.FIELDS(0).Name & " = " & CheckedIndex
     bRes = Not MVFRS.NoMatch
'-----------------------------
ExitHere:
    Check_IFinMVF = bRes  '!!!!!!!!!!!!!!
    Exit Function
'-----------
ErrHandle:
    ErrPrint "Check_IFinMVF", Err.Number, Err.Description
    Err.Clear
End Function
'=================================================================================================================================================
' The function retrieves all integer indexes stored in MVF to array
'=================================================================================================================================================
Public Function Get_MVF_VALUES(IDD As Long, TBL As String, MVFFLD As String) As Long()
Dim iRes() As Long, nDim As Long
Dim RS As DAO.Recordset, rst As DAO.Recordset

On Error GoTo ErrHandle
'-----------------------------
ReDim iRes(0): nDim = -1
   Set RS = CurrentDb.OpenRecordset("SELECT " & MVFFLD & " FROM " & TBL & " WHERE(ID = " & IDD & ");")
   With RS
        If Not .EOF Then
            Set rst = .FIELDS(MVFFLD).value
            If Not rst.EOF Then
                rst.MoveLast: rst.MoveFirst
                Do While Not rst.EOF
                    nDim = nDim + 1: ReDim Preserve iRes(nDim)
                    iRes(nDim) = rst.FIELDS(0).value
                    rst.MoveNext
                Loop
                rst.Close
            End If
        End If
   End With
'-----------------------------
ExitHere:
    Get_MVF_VALUES = iRes '!!!!!!!!!!!!!!
    Set rst = Nothing: Set RS = Nothing
    Exit Function
'-----------
ErrHandle:
    ErrPrint "Get_MVF_VALUES", Err.Number, Err.Description
    Err.Clear
End Function





'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'----------------------------------------------------------------------------------------------------------------------------------------------
' Error Handler
'----------------------------------------------------------------------------------------------------------------------------------------------
Private Function ErrPrint(FuncName As String, ErrNumber As Long, ErrDescription As String, Optional bDebug As Boolean = True) As String
Dim sRes As String
Const ErrChar As String = "#"
Const ErrRepeat As Integer = 60

'---------------------------------------------------------
sRes = String(ErrRepeat, ErrChar) & vbCrLf & "ERROR OF [" & "mod_MVF_" & FuncName & "]" & vbTab & "ERR#" & ErrNumber & vbTab & Now() & _
       vbCrLf & ErrDescription & vbCrLf & String(ErrRepeat, ErrChar)
If bDebug Then Debug.Print sRes
'----------------------------------------------------------
ExitHere:
       Beep
       ErrPrint = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'----------------------------------------------------------------------------------------------------------------------------------------------
' Function Return List of Values for MVF (no preliminary check)
'----------------------------------------------------------------------------------------------------------------------------------------------
Public Function MVFToList(strTableName As String, strMVFField As String, IDD As Long, _
                                                     Optional strPrimaryKeyFieldName As String = "ID", Optional DLM As String = ";") As String
Dim sRes As String, RS As DAO.Recordset, rst As DAO.Recordset
Dim SQL As String

On Error GoTo ErrHandle
'-------------------------------------------------------------
SQL = "SELECT " & strPrimaryKeyFieldName & ", " & strMVFField & " FROM " & _
                  strTableName & " WHERE (" & strPrimaryKeyFieldName & " = " & IDD & ")"
Set RS = CurrentDb.OpenRecordset(SQL)
With RS
    If Not .EOF Then
        .MoveLast: .MoveFirst
        Set rst = .FIELDS(strMVFField).value
        If Not rst.EOF Then
               rst.MoveLast: rst.MoveFirst
               Do While Not rst.EOF
                   sRes = sRes & rst.FIELDS(0).value & DLM
                   rst.MoveNext
               Loop
        End If
    End If
End With
If sRes <> "" Then sRes = Left(sRes, Len(sRes) - Len(DLM))
'-------------------------------------------------------------
ExitHere:
    MVFToList = sRes '!!!!!!!!!!
    Set RS = Nothing: Set rst = Nothing
    Exit Function
'---------------------
ErrHandle:
    ErrPrint "MVFToList", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------
' Get ID array for MVF
'--------------------------------------------------------------------------------------------------------------------------------------------------
Public Function GetMVFValues(IDD As Long, MVFFLD As String, TBL As String) As Long()
Dim iRes() As Long, nDim As Integer
Dim RS As DAO.Recordset, rst As DAO.Recordset, SQL As String

On Error GoTo ErrHandle
'----------------------
    ReDim iRes(0): nDim = -1
    If IDD <= 0 Then GoTo ExitHere
    
    SQL = "SELECT * FROM " & TBL & " WHERE (ID = " & IDD & ");"
    Set RS = CurrentDb.OpenRecordset(SQL)
    With RS
        If Not .EOF Then
            .MoveLast: .MoveFirst
            
            Set rst = RS.FIELDS(MVFFLD).value
            If Not rst.EOF Then
                rst.MoveLast: rst.MoveFirst
                Do While Not rst.EOF
                   nDim = nDim + 1: ReDim Preserve iRes(nDim)
                   iRes(nDim) = rst.FIELDS(0)
                   rst.MoveNext
                Loop
            End If
        End If
    End With
'----------------------
ExitHere:
    GetMVFValues = iRes '!!!!!!!!!!!!!!!
    Set rst = Nothing: Set RS = Nothing
    Exit Function
'---------
ErrHandle:
    ErrPrint "GetMVFValues", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function

Public Function SaveFileToDB(ByVal fileName As String, _
   RS As Object, FieldName As String) As Boolean
'**************************************************************
'PURPOSE: SAVES DATA FROM BINARY FILE (e.g., .EXE, WORD DOCUMENT
'CONTROL TO RECORDSET RS IN FIELD NAME FIELDNAME

'FIELD TYPE MUST BE BINARY (OLE OBJECT IN ACCESS)

'REQUIRES: REFERENCE TO MICROSOFT ACTIVE DATA OBJECTS 2.0 or ABOVE

'SAMPLE USAGE
'Dim sConn As String
'Dim oConn As New ADODB.Connection
'Dim oRs As New ADODB.Recordset
'
'
'sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MyDb.MDB;Persist Security Info=False"
'
'oConn.Open sConn
'oRs.Open "SELECT * FROM MYTABLE", oConn, adOpenKeyset, _
   adLockOptimistic
'oRs.AddNew

'SaveFileToDB "C:\MyDocuments\MyDoc.Doc", oRs, "MyFieldName"
'oRs.Update
'oRs.Close
'**************************************************************

Dim iFileNum As Integer
Dim lFileLength As Long

Dim abBytes() As Byte
Dim iCtr As Integer

On Error GoTo ErrorHandler
If Dir(fileName) = "" Then Exit Function
If Not TypeOf RS Is DAO.Recordset Then Exit Function

'read file contents to byte array
iFileNum = FreeFile
Open fileName For Binary Access Read As #iFileNum
lFileLength = LOF(iFileNum)
ReDim abBytes(lFileLength)
Get #iFileNum, , abBytes()

'put byte array contents into db field
RS.FIELDS(FieldName).AppendChunk abBytes()
Close #iFileNum

SaveFileToDB = True
ErrorHandler:
End Function
'-----------------------------------------------------------------------------------------------------------------
' Save Binary File From DB
'-----------------------------------------------------------------------------------------------------------------
Public Function SaveBinaryFromDB(TBL As String, FieldName As String, IDD As Long, _
                                                 Optional sROOT As String, Optional IDFld As String = "ID") As String
Dim iFileNum As Integer, lFileLength As Long, SQL As String, strPath As String, sRes As String
Dim abBytes() As Byte, RS As DAO.Recordset, iCtr As Integer, fileName As String

Const DFLTPATH As String = "FILES"
On Error GoTo ErrHandle
'-----------------------------------------------------------
fileName = TBL & "_" & FieldName & "-" & IDD & ".blob"
strPath = AdjustFilePath(sROOT, fileName): If strPath = "" Then Exit Function

iFileNum = FreeFile
Open strPath For Binary As #iFileNum
lFileLength = LenB(RS(FieldName))
'------------------------------------------------------------
    SQL = "SELECT " & IDFld & ", " & FieldName & " FROM " & TBL & " WHERE(" & IDFld & " = " & IDD & ")"
    Set RS = CurrentDb.OpenRecordset(SQL)
    With RS
        If Not .EOF Then
            .MoveLast: .MoveFirst
            abBytes = RS(FieldName).GetChunk(0, lFileLength)
            Put #iFileNum, , abBytes()
        End If
    End With
sRes = strPath
'------------------------------------------------------------
ExitHere:
    SaveBinaryFromDB = sRes '!!!!!!!!!
    Close #iFileNum
    Exit Function
'----------------
ErrHandle:
    ErrPrint "SaveBinaryFromDB" & "(TBL = " & TBL & "; FieldName = " & FieldName & ")", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Save Memo Or HyperLink to External File
'---------------------------------------------------------------------------------------------------------------------------------------------------
Public Function SaveTextFromDB(TBL As String, FieldName As String, IDD As Long, _
                                                 Optional sROOT As String, Optional IDFld As String = "ID") As String
Dim SQL As String, strPath As String, sRes As String, fileName As String, RS As DAO.Recordset
Dim sText As String

Const DFLTPATH As String = "FILES"

On Error GoTo ErrHandle
'-----------------------------------------------------------
fileName = TBL & "_" & FieldName & "-" & IDD & ".txt"
strPath = AdjustFilePath(sROOT, fileName): If strPath = "" Then Exit Function
'------------------------------------------------------------
    SQL = "SELECT " & IDFld & ", " & FieldName & " FROM " & SHT(TBL) & " WHERE(" & IDFld & " = " & IDD & ")"
    Set RS = CurrentDb.OpenRecordset(SQL)
    With RS
        If Not .EOF Then
            .MoveLast: .MoveFirst
            sText = .FIELDS(FieldName).value
            WriteStringToFile strPath, sText
        End If
    End With

sRes = strPath
'------------------------------------------------------------
ExitHere:
    SaveTextFromDB = sRes '!!!!!!!!!
    Exit Function
'----------------
ErrHandle:
    ErrPrint "SaveTextFromDB" & "(TBL =" & TBL & "; FLD = " & FieldName & "; ID = " & IDD & ")", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Function Check file path and ajust it
'---------------------------------------------------------------------------------------------------------------------------------------------------
Private Function AdjustFilePath(Optional sFolder As String, Optional sFileName As String, _
                                                        Optional GenerateAbsentFileName As Boolean, Optional bCreateDir As Boolean = True) As String
Dim sPath As String, sFile As String, sRes As String

Const DFLTPATH As String = "FILES"
Const DFLTEXT As String = ".dat"

On Error GoTo ErrHandle
'-----------------------------------------------------------
If sFileName = "" Then
   If sFolder = "" Then
      sPath = CurrentProject.Path & "\" & DFLTPATH
      If GenerateAbsentFileName Then sFile = Format(Now(), "yyyymmddhhnnss") & DFLTEXT
   Else
      If IsFilePath(sFolder) Then
         sPath = FolderNameOnlyAbstract(sFolder)
         sFile = FileNameOnly(sFolder)
      Else
         sPath = sFolder
         If GenerateAbsentFileName Then sFile = Format(Now(), "yyyymmddhhnnss") & DFLTEXT
      End If
   End If
Else
   If sFolder <> "" Then
        sFile = FileNameOnly(sFileName)
        sPath = FolderNameOnlyAbstract(sFolder)
        If sPath = "" Then sPath = sFolder
   Else
       sFile = FileNameOnly(sFileName)
       sPath = FolderNameOnlyAbstract(sFileName)
       If sPath = "" Then sPath = CurrentProject.Path & "\" & DFLTPATH
   End If
End If
'-----------------------------------------------------------
If (sPath = "") Then Exit Function
If bCreateDir Then
   If Dir(sPath, vbDirectory) = "" Then MkDir sPath
End If
If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
sRes = sPath & sFile
'-----------------------------------------------------------
ExitHere:
    AdjustFilePath = sRes '!!!!!!!!!
    Exit Function
'-----------------------------------------------------------
ErrHandle:
    ErrPrint "AdjustFilePath", Err.Number, Err.Description
    Err.Clear
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------
' check if the string is FilePath (require extention)
'------------------------------------------------------------------------------------------------------------------------------------------------
Private Function IsFilePath(sPath As String) As Boolean
Dim iL As Integer, bRes As Boolean
Const sDOt As String = "."
If sPath = "" Then Exit Function
     iL = InStrRev(sPath, sDOt)
     If iL > 0 Then
          iL = Len(sPath) - iL
          If ((iL > 1) And (iL < 7)) Then bRes = True
     End If
'----------------------------
   IsFilePath = bRes '!!!!!!!!!!!!!!!!!!!!!!!
End Function


