VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cVTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
' This Module was generated in 10. 01. 2021 5: 58: 10  This class allows you to store and add data like a virtual string table and in fact is
'  a two-dimensional array with a word addressing the names of the columns
'  @ Valery Khvatov (valery.khvatov@gmail.com), [01/20210110]
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
Option Compare Database
Option Explicit
'************************************************************
Private Type COLUMN
      Arr() As String
End Type

Private Const MOD_NAME As String = "cVTable"

Private FIELDS As cDictionary
Private INDEXES As cDictionary

Private m_MATRIX() As COLUMN

Private m_Fld() As TFLD
Private m_Records As Long
Private m_Fields As Long
Private m_TableName As String

Private m_IndexedField As String
Private m_IndexedFieldIndex As Long



Public Property Get TableName() As String
     TableName = m_TableName
End Property
Public Property Let TableName(sName As String)
     m_TableName = sName
End Property

Public Property Get Count() As Long
    Count = m_Records + 1
End Property

Public Property Get FieldCount() As Long
    FieldCount = m_Fields + 1
End Property

Public Property Get Item(Index As Long, Field As Variant) As String
Attribute Item.VB_UserMemId = 0
Dim iFld As Long

    If Index < 0 Or Index > m_Records Then Exit Property
    iFld = GetFLDIndex(Field): If iFld < 0 Then Exit Property
'-------------------------------------
    Item = GET_DATA(Index, iFld)
End Property
Public Property Let Item(Index As Long, Field As Variant, sVal As String)
Attribute Item.VB_UserMemId = 0
Dim iFld As Long, nFlds As Long
    
    On Error GoTo ErrHandle
'--------------------------------------
    iFld = GetFLDIndex(Field): If iFld < 0 Then Err.Raise 10001, , "Wrong Field"
    If Index < 0 Then Err.Raise 10001, , "Wrong Index"
    
     If iFld = m_IndexedFieldIndex Then
        If INDEXES.Exists(sVal) Then Err.Raise 10003, , "Can't add Value " & sVal & " for Indexed Field: this value is exists"
        INDEXES.Add sVal, Index
     End If

    If Index > m_Records Then
        m_Records = Index
        Call ReDim_By_Rows(m_Records)
    End If
'--------------------------------------
ExitHere:
     Call SET_DATA(Index, iFld, sVal)
     Exit Property
'----------------
ErrHandle:
     ErrPrint2 "Property Get Item", Err.Number, Err.Description, MOD_NAME
     Err.Clear
End Property

Public Property Get FieldType(Index As Long) As Integer
    If Index < 0 Or Index > m_Fields Then Exit Property
    FieldType = m_Fld(Index).Type
End Property
Public Property Let FieldType(Index As Long, iFieldType As Integer)
    If Index < 0 Or Index > m_Fields Then Exit Property
    m_Fld(Index).Type = iFieldType
End Property

Public Property Get FieldSize(Index As Long) As Integer
    If Index < 0 Or Index > m_Fields Then Exit Property
    FieldSize = m_Fld(Index).SIZE
End Property
Public Property Let FieldSize(Index As Long, iFieldSize As Integer)
    If Index < 0 Or Index > m_Fields Then Exit Property
    m_Fld(Index).SIZE = iFieldSize
End Property

Public Property Get FieldName(Index As Long) As String
    If Index < 0 Or Index > m_Fields Then Exit Property
    FieldName = m_Fld(Index).Name
End Property
Public Property Let FieldName(Index As Long, sFieldName As String)
Dim sOldFldName As String

    On Error GoTo ErrHandle
'-----------------------
    If sFieldName = "" Then Exit Property
    
    If Index < 0 Then Exit Property
    
    If FIELDS.Exists(sFieldName) Then
        Err.Raise 10009, , "Can't change field name to existing one: " & sFieldName
    Else
        If Index > m_Fields Then
            AddField sFieldName
        Else
            sOldFldName = m_Fld(Index).Name
            m_Fld(Index).Name = sFieldName
            FIELDS.Remove sOldFldName
            FIELDS.Add sFieldName, Index
        End If
    End If
'-----------------------
ExitHere:
    Exit Property
'-----------------
ErrHandle:
    ErrPrint2 "FieldName", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Property

'======================================================================================================================================================
' Add FieldList from string. Could add only not Indexed Text Fields
' Full FieldList comes as Fld_Name;Fld_Type;Fld_Size ^ Fld_Name;Fld_Type;Fld_Size ^ Fld_Name;Fld_Type;Fld_Size
' Short FieldList comes as on row: Fld_Name_1; Fld_Name_2; Fld_Name_3;...
'======================================================================================================================================================
Public Sub AddFieldList(sFieldList As String, Optional DLM As String = ";", Optional SEP As String = vbCrLf)
Dim FLDS() As TFLD, nFlds As Long, I As Long, bEasyList As Boolean
Dim sFLD() As String, fldSize As Integer, FldType As Integer

   On Error GoTo ErrHandle
'--------------------------
   If sFieldList = "" Then Exit Sub
   
   FLDS = FldListFromString(sFieldList, DLM, SEP): nFlds = UBound(FLDS)
   For I = 0 To nFlds
         Me.AddField FLDS(I).Name, FLDS(I).Type, FLDS(I).SIZE, FLDS(I).Required, FLDS(I).DefaultValue, FLDS(I).Description, FLDS(I).PK
   Next I
      
'--------------------------
ExitHere:
    Exit Sub
'------------
ErrHandle:
    ErrPrint2 "AddFieldList", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Sub

'======================================================================================================================================================
' Function returns the array of indexes for some vale of the field
'======================================================================================================================================================
Public Function Search(Field As String, sValue As String, Optional iCompare As VbCompareMethod = vbBinaryCompare) As Long()
Dim iFld As Long, iRes() As Long, nDim As Long, I As Long

    On Error GoTo ErrHandle
'---------------------------------------
    ReDim iRes(0): iRes(0) = -1: nDim = -1
    iFld = GetFLDIndex(Field)
    If iFld > 0 Then
         If iFld = m_IndexedFieldIndex Then
             If INDEXES.Exists(sValue) Then iRes(0) = iFld
         Else
             For I = 0 To m_Records
                 If StrComp(GET_DATA(I, iFld), sValue, iCompare) = 0 Then
                     nDim = nDim + 1
                     ReDim Preserve iRes(nDim): iRes(nDim) = I
                 End If
             Next I
         End If
    End If
'---------------------------------------
ExitHere:
    Search = iRes '!!!!!!!!!!!!
    Exit Function
'----------------
ErrHandle:
    ErrPrint2 "Search", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'======================================================================================================================================================
' Just Increment Array By Record Indexes with step 1
'======================================================================================================================================================
Public Sub AddRecord()
Dim nFlds As Long
    m_Records = m_Records + 1
    Call ReDim_By_Rows(m_Records)
End Sub
'======================================================================================================================================================
' Add Field/Column to Array and the class
'======================================================================================================================================================
Public Sub AddField(FieldName As String, Optional FldType As DataTypeEnum = dbText, Optional fldSize As Integer, _
             Optional fldRequied As Boolean, Optional fldDefaultValue As String, Optional fldDescripton As String, Optional Indexed As Boolean = False)

    On Error GoTo ErrHandle
'----------------------------------
    If FieldName = "" Then Exit Sub
    If FIELDS.Exists(FieldName) Then Err.Raise 10000, , "Can't add field " & FieldName & " This field is exist in this VTAble"
    
    m_Fields = m_Fields + 1
    
    Call Redim_By_Col(m_Fields)
    ReDim Preserve m_Fld(m_Fields)
    
    FIELDS.Add FieldName, m_Fields
    m_Fld(m_Fields).Name = FieldName: m_Fld(m_Fields).Type = FldType: m_Fld(m_Fields).SIZE = fldSize
    m_Fld(m_Fields).Required = fldRequied: m_Fld(m_Fields).DefaultValue = fldDefaultValue: m_Fld(m_Fields).Description = fldDescripton
    
    If Indexed Then Call SetIndex(FieldName, m_Fields)
'----------------------------------
ExitHere:
    Exit Sub
'-----------------
ErrHandle:
    ErrPrint2 "AddField", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Sub
'======================================================================================================================================================
' Edit Field Name and Field Type (no effect to stored data)
'======================================================================================================================================================
Public Sub EditField(OldFieldName As String, Optional NewFieldName As String, Optional NewFieldType As Integer, Optional NewFieldSize As Integer, _
                                Optional NewFieldRequired As Boolean, Optional NewFieldDefaultValue As String, Optional NewFieldDescription As String)

Dim iFld As Long

    On Error GoTo ErrHandle
'-----------------------------------------------
    If OldFieldName = "" Then Exit Sub
    iFld = GetFLDIndex(OldFieldName): If iFld < 0 Then Exit Sub
    
    If NewFieldName <> "" Then
       FIELDS.Remove OldFieldName
       FIELDS.Add NewFieldName, iFld
       m_Fld(iFld).Name = NewFieldName
    End If
    
    If NewFieldSize > 0 Then m_Fld(iFld).SIZE = NewFieldSize
    If NewFieldType > 0 Then m_Fld(iFld).Type = NewFieldType
    If NewFieldRequired = True Then m_Fld(iFld).Required = NewFieldRequired
    If NewFieldDefaultValue <> "" Then m_Fld(iFld).DefaultValue = NewFieldDefaultValue
    If NewFieldDescription <> "" Then m_Fld(iFld).Description = NewFieldDescription
'-----------------------------------------------
ExitHere:
    Exit Sub
'------------
ErrHandle:
    ErrPrint2 "EditField", Err.Number, Err.Description
    Err.Clear
End Sub
'======================================================================================================================================================
' Get Field String
'======================================================================================================================================================
Public Function GetFldString(Optional DLM As String = ";") As String
     If FIELDS.Count > 0 Then GetFldString = Join(FIELDS.Keys, DLM)
End Function
'======================================================================================================================================================
' Serialize to String
'======================================================================================================================================================
Public Function ToString(Optional WithHeaders As Boolean = True, Optional DLM As String = ";", Optional SEP As String = vbCrLf, _
                                                                                                           Optional FldLimit As Integer = 40) As String
Dim I As Long, J As Long, sHeader As String, sWork As String
Dim sRes As String, ROWS() As String, ARGS() As String, nARGS As Long

Const GS_ASC As Integer = 29

    On Error GoTo ErrHandle
'---------------------------
    nARGS = -1: ReDim ARGS(0): ReDim ROWS(m_Records)
    For I = 0 To m_Records
        For J = 0 To m_Fields
            sWork = StringFilter(m_MATRIX(J).Arr(I), True, DLM, SEP)
            If Len(sWork) > FldLimit Then
                nARGS = nARGS + 1: ReDim Preserve ARGS(nARGS)
                ARGS(nARGS) = SHT(m_Fld(J).Name & "_" & I) & SEP & sWork
                sWork = SHT(m_Fld(J).Name & "_" & I)
            End If
            ROWS(I) = ROWS(I) & DLM & sWork
        Next J
        ROWS(I) = Right(ROWS(I), Len(ROWS(I)) - Len(DLM))
    Next I
    sRes = Join(ROWS, SEP)
    If nARGS >= 0 Then sRes = sRes & SEP & Chr(GS_ASC) & SEP & Join(ARGS, SEP)
    If WithHeaders Then sRes = GetFldString(DLM) & SEP & sRes
'---------------------------
ExitHere:
    ToString = sRes '!!!!!!!!!!
    Exit Function
'----------------
ErrHandle:
    ErrPrint2 "ToString", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'======================================================================================================================================================
' Rebuild from String (only fo text fieldtype)
'======================================================================================================================================================
Public Sub FromString(str As String, Optional WithHeader As Boolean = True, Optional DLM As String = ";", Optional SEP As String = vbCrLf)
Dim ROWS() As String, sWork() As String, mFLDs As Long, nRows As Long, I As Long, J As Long, iRow As Long, sRows As String, sHeader As String
Dim nARGS As Long, dict As cDictionary
Dim sUnFiltered As String

Const GS_ASC As Integer = 29

   On Error GoTo ErrHandle
'----------------------------
   If str = "" Then Exit Sub
   sWork = Split(str, SEP & Chr(GS_ASC) & SEP)
   
   If UBound(sWork) > 0 Then
           sRows = sWork(0)
           Set dict = GetKVArg(sWork(1), SEP): nARGS = dict.Count
   Else
           sRows = sWork(0): nARGS = -1
   End If
   
   If sRows = "" Then Exit Sub
   
   If WithHeader Then
      I = InStr(1, sRows, SEP)
      If I > 0 Then
          sHeader = Left(sRows, I - 1)
          sRows = Right(sRows, Len(sRows) - I - Len(SEP) + 1)
      End If
   End If
   
   If sRows = "" Then Exit Sub
'----------------------------
   If m_Records > 0 Or m_Fields > 0 Then Call RemoveAll
   If sHeader <> "" Then Call AddFieldList(sHeader, DLM)
     
   ROWS = Split(sRows, SEP): nRows = UBound(ROWS)
   Call ReDim_By_Rows(nRows)
   
   If m_Fields < 0 Then
      m_Fields = UBound(Split(ROWS(0), DLM))
      If m_Fields < 0 Then Err.Raise 10005, , "Wrong string format"
      Call Redim_By_Col(m_Fields)
   End If
   
'----------------------------
   m_Records = nRows
   For I = 0 To nRows
       If ROWS(I) = "" Then GoTo NextRow
       sWork = Split(ROWS(I), DLM): mFLDs = UBound(sWork)
       
             For J = 0 To m_Fields
                    sUnFiltered = Trim(sWork(J))
                    
                    If sUnFiltered <> "" Then
                        If Left(sUnFiltered, 1) = "[" And Right(sUnFiltered, 1) = "]" And nARGS > 0 Then
                            If dict.Exists(sUnFiltered) Then sUnFiltered = dict(sUnFiltered)
                         End If
                         
                    End If
                    m_MATRIX(J).Arr(I) = StringFilter(sUnFiltered, False, DLM, SEP)
             Next J
NextRow:
   Next I
'----------------------------
ExitHere:
   Exit Sub
'-----------
ErrHandle:
   ErrPrint2 "FromString", Err.Number, Err.Description, MOD_NAME
   Err.Clear
End Sub


'------------------------------------------------------------------------------------------------------------------------------------------------------
' Create Argument Based KV-Pairs
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetKVArg(str As String, Optional SEP As String = vbCrLf) As Object
Dim art() As String, nART As Long, I As Long, sWork As String
Dim dict As New cDictionary
    
    On Error GoTo ErrHandle
'------------------------
    sWork = Trim(str)
    If sWork = "" Then GoTo ExitHere
    
    art = Split(sWork, SEP): nART = UBound(art)
    For I = 0 To nART
        sWork = Trim(art(I))
        If Left(sWork, 1) = "[" And Right(sWork, 1) = "]" Then
           If I < nART - 1 Then
              dict.Add sWork, Trim(art(I + 1))
           End If
        End If
    Next I
'------------------------
ExitHere:
    Set GetKVArg = dict '!!!!!!!!!!!!
    Exit Function
'----------------
ErrHandle:
    ErrPrint2 "GetKVArg", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------
' Filter string from DLM and SEP
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function StringFilter(str As String, Optional ModeForward As Boolean = True, Optional DLM As String = ";", _
                                     Optional SEP As String = vbCrLf, Optional REPL_DLM As Integer = 164, Optional REPL_SEP As Integer = 172) As String
Dim sRes As String
   If str = "" Then Exit Function
'--------------------------
   If ModeForward Then
        sRes = Replace(str, DLM, Chr(REPL_DLM))
        sRes = Replace(sRes, SEP, Chr(REPL_SEP))
   Else
        sRes = Replace(str, Chr(REPL_DLM), DLM)
        sRes = Replace(str, Chr(REPL_SEP), SEP)
   End If
'--------------------------
ExitHere:
   StringFilter = sRes '!!!!!!!!!!!!
End Function

'======================================================================================================================================================
' Get Row as Array
'======================================================================================================================================================
Public Function GetRow(iRow As Long) As String()
Dim MyArr() As String, nDim As Long, I As Long

     On Error GoTo ErrHandle
'--------------------------------
     ReDim MyArr(0)
     If iRow < 0 Or iRow > m_Records Then GoTo ExitHere
     If m_Fields < 0 Then GoTo ExitHere
     
     ReDim MyArr(m_Fields)
     For I = 0 To m_Fields
         MyArr(I) = m_MATRIX(I).Arr(iRow)
     Next I
'--------------------------------
ExitHere:
     GetRow = MyArr '!!!!!!!!!!!!
     Exit Function
'------------
ErrHandle:
     ErrPrint2 "GetRow", Err.Number, Err.Description, MOD_NAME
     Err.Clear
End Function

'======================================================================================================================================================
' Get Column as Array
'======================================================================================================================================================
Public Function GetColumn(iColumn As Long) As String()
Dim MyArr() As String, nDim As Long, I As Long

     On Error GoTo ErrHandle
'--------------------------------
     ReDim MyArr(0)
     If iColumn < 0 Or iColumn > m_Fields Then GoTo ExitHere
     If m_Records < 0 Then GoTo ExitHere
     
     ReDim MyArr(m_Records)
     For I = 0 To m_Records
        MyArr(I) = m_MATRIX(iColumn).Arr(I)
     Next I
'--------------------------------
ExitHere:
     GetColumn = MyArr '!!!!!!!!!!!!
     Exit Function
'------------
ErrHandle:
     ErrPrint2 "GetColumn", Err.Number, Err.Description, MOD_NAME
     Err.Clear
End Function
'======================================================================================================================================================
' Clear Data (and save the structure)
'======================================================================================================================================================
Public Sub Clear()
Dim J As Long
    
    On Error Resume Next
'----------------------
    Call ERASE_DATA
    For J = 0 To m_Fields
        ReDim m_MATRIX(J).Arr(0)
    Next J
    m_Records = -1
'----------------------
ExitHere:
    Exit Sub
End Sub
'======================================================================================================================================================
' Remove data and structure
'======================================================================================================================================================
Public Sub RemoveAll()

    On Error Resume Next
'----------------------
    Erase m_MATRIX
    ReDim m_MATRIX(0): ReDim m_MATRIX(0).Arr(0)
    
    m_Records = -1
    
    FIELDS.RemoveAll: INDEXES.RemoveAll
    Erase m_Fld: ReDim m_Fld(0): m_Fields = -1
'----------------------
ExitHere:
    Exit Sub
End Sub
'======================================================================================================================================================
' Save To File
'======================================================================================================================================================
Public Function SaveToFile(Optional sPath As String, Optional WithHeader As Boolean = True, Optional DLM As String = ";", _
                                                                                                             Optional SEP As String = vbCrLf) As String
Dim sRes As String, sToSave As String, sFile As String

    On Error GoTo ErrHandle
'--------------------------
    If sPath = "" Then
        sFile = CurrentProject.Path & "\" & GetTimestamp & "_" & m_TableName & ".dat"
    Else
        sFile = sPath
    End If
    If Dir(sFile) <> "" Then Kill sFile
    sToSave = ToString(WithHeader, DLM, SEP): If sToSave = "" Then Exit Function
    
    If WriteStringToFileUTF8(sToSave, sFile) Then sRes = sFile
'--------------------------
ExitHere:
    SaveToFile = sRes '!!!!!!!!!!!!!!
    Exit Function
'-------------
ErrHandle:
    ErrPrint2 "SaveToFile", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'======================================================================================================================================================
' Read From File
'======================================================================================================================================================
Public Sub ReadFromFile(Optional sPath As String, Optional WithHeader As Boolean = True, Optional DLM As String = ";", _
                                                                                                             Optional SEP As String = vbCrLf)
Dim sFile As String, sToBuild As String

    On Error GoTo ErrHandle
'--------------------------
    If sPath = "" Then
        sFile = OpenDialog(GC_OPEN_FILE, "Pick the file", , False, CurrentProject.Path)
    Else
        sFile = sPath
    End If
    If Dir(sFile) = "" Then Err.Raise 10005, , "Can't find the file " & sFile
'--------------------------
    sToBuild = ReadTextFileUTF8(sFile)
    If sToBuild <> "" Then
       Call FromString(sToBuild, WithHeader, DLM, SEP)
    End If
'--------------------------
ExitHere:
    Exit Sub
'---------
ErrHandle:
    ErrPrint2 "ReadFromFile", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Sub
'======================================================================================================================================================
' Read Structure and Data From Table/Query/Recordset
'======================================================================================================================================================
Public Function ReadFromTable(dbObject As Variant) As Boolean
Dim bRes As Boolean, sObject As String, RS As DAO.Recordset, I As Long, J As Long, iFieldIndex As Long

Const DLM As String = ";"
Const SEP As String = vbCrLf

    On Error GoTo ErrHandle
'-----------------------
    If m_Records > -1 Or m_Fields > -1 Then RemoveAll
    
    Call ReadFieldsFromTable(dbObject, DLM, SEP)
    
    If m_Fields = -1 Then Exit Function

    If varType(dbObject) = vbString Then ' Table Or Query
         sObject = CStr(dbObject)
         Set RS = CurrentDb.OpenRecordset(sObject)
    ElseIf varType(dbObject) = vbObject Or varType(dbObject) = vbDataObject Then
         Set RS = dbObject
    Else
         Err.Raise 10012, , "Wrong object type:" & CStr(dbObject)
    End If
    
    I = -1
    With RS
        If Not .EOF Then
             .MoveLast: .MoveFirst
             Do While Not .EOF
                I = I + 1: If I > m_Records Then Call AddRecord
                For J = 0 To .FIELDS.Count - 1
                   iFieldIndex = GetFLDIndex(.FIELDS(J).Name)
                   If iFieldIndex < 0 Or iFieldIndex > m_Fields Then _
                            Err.Raise 10013, , "Worong Field Index " & iFieldIndex & " for field " & SHT(.FIELDS(J).Name)
                   Call SET_DATA(I, iFieldIndex, CStr(Nz(.FIELDS(J).value)))
                Next J
                .MoveNext
             Loop
        End If
    End With
    bRes = True
'-----------------------
ExitHere:
    ReadFromTable = bRes '!!!!!!!!!!!!!!!
    Set RS = Nothing
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "ReadFromTable", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'======================================================================================================================================================
' Set Fields by Physical Table. dbObject could be Table Name, Query Name ot Recordset
'======================================================================================================================================================
Public Sub ReadFieldsFromTable(dbObject As Variant, Optional DLM As String = ";", Optional SEP As String = vbCrLf)
Dim I As Integer, sFields As String
    On Error GoTo ErrHandle
'--------------------------
    If m_Fields > -1 Or m_Records > 0 Then Call RemoveAll
    sFields = GetFieldList(dbObject, DLM, SEP)
    
    If sFields = "" Then Exit Sub
    Call AddFieldList(sFields, DLM, SEP)
    
'--------------------------
ExitHere:
    Exit Sub
'---------
ErrHandle:
    ErrPrint2 "ReadFieldsFromTable", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Sub

'======================================================================================================================================================
' Get/Set Indexed Field Name
'======================================================================================================================================================
Public Function GetIndexedField() As String
    GetIndexedField = m_IndexedField '!!!!!!!!!!!!!
End Function
Public Sub SetIndexedField(indx As Long)
    m_IndexedField = indx '!!!!!!!!!!!!!
End Sub

'======================================================================================================================================================
' Set Index for specific field (only first time is allowed)
'======================================================================================================================================================
Public Sub SetIndex(FieldName As String, IndexedFieldIndex As Long)
         
        On Error GoTo ErrHandle
'--------------------------------------
        If m_IndexedField <> "" Then Err.Raise 10002, , "The only one Index is allowed for VTable. It is " & m_IndexedField
        m_IndexedField = FieldName: m_IndexedFieldIndex = IndexedFieldIndex
'--------------------------------------
ExitHere:
        Exit Sub
'----------------
ErrHandle:
        ErrPrint2 "", Err.Number, Err.Description, MOD_NAME
        Err.Clear
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
    ReDim m_MATRIX(0): ReDim m_MATRIX(0).Arr(0)
    ReDim m_Fld(0)
    
    Set FIELDS = New cDictionary
    Set INDEXES = New cDictionary
    
    m_Records = -1: m_Fields = -1
    m_IndexedFieldIndex = -1
    
    m_TableName = GenRandomStr()
End Sub

Private Sub Class_Terminate()
    Erase m_MATRIX
    Set FIELDS = Nothing
    Set INDEXES = Nothing
End Sub


'------------------------------------------------------------------------------------------------------------------------------------------------------
'   Get FLD Index
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetFLDIndex(Field As Variant) As Long
Dim iRes As Long

   On Error GoTo ErrHandle
'----------------------------------
    iRes = -1
    If IsNumeric(Field) Then
       iRes = CLng(Field)
    Else
       If FIELDS.Exists(CStr(Field)) Then
          iRes = CLng(FIELDS(CStr(Field)))
       End If
    End If
    
    If iRes > m_Fields Then iRes = -1
'----------------------------------
ExitHere:
    GetFLDIndex = iRes '!!!!!!!!!!
    Exit Function
'----------------
ErrHandle:
    Err.Clear: GetFLDIndex = -1
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------
'   GET DATA
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GET_DATA(iRow As Long, iCol As Long) As String
      
   On Error Resume Next
'---------------------------------
   If iRow < 0 Or iRow > m_Records Then Exit Function
   If iCol < 0 Or iCol > m_Fields Then Exit Function
   
'-----------------------------------
ExitHere:
   GET_DATA = m_MATRIX(iCol).Arr(iRow)
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
'   SET DATA
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub SET_DATA(iRow As Long, iCol As Long, sValue As String)

   On Error GoTo ErrHandle
'---------------------------------
   If iRow < 0 Or iRow > m_Records Then Err.Raise 10020, , "Wrong Row Number: " & iRow & " with max records " & m_Records
   If iCol < 0 Or iCol > m_Fields Then Err.Raise 10021, , "Wrong Col Number: " & iCol & "with max columns " & m_Fields
   m_MATRIX(iCol).Arr(iRow) = sValue
'---------------------------------
ExitHere:
   Exit Sub
'----------
ErrHandle:
   ErrPrint2 "SET_DATA", Err.Number, Err.Description, MOD_NAME
   Err.Clear
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------------------
'   ReDim by Cols
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Redim_By_Col(nCols As Long)
    ReDim Preserve m_MATRIX(nCols)
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------------------
'   ReDim by Rows
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ReDim_By_Rows(nRows As Long)
Dim J As Long
    
    On Error GoTo ErrHandle
'--------------------------
    If m_Fields = -1 Then Exit Sub
    
    For J = 0 To m_Fields
        ReDim Preserve m_MATRIX(J).Arr(nRows)
    Next J
'--------------------------
ExitHere:
    Exit Sub
'----------
ErrHandle:
    ErrPrint2 "ReDim_By_Rows", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------------------
'   Erase Data
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ERASE_DATA()
Dim J As Long

    On Error GoTo ErrHandle
'--------------------------
    If m_Fields = -1 Then Exit Sub
    
    For J = 0 To m_Fields
        Erase m_MATRIX(J).Arr
    Next J
'--------------------------
ErrHandle:
    ErrPrint2 "ERASE_DATA", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Sub
