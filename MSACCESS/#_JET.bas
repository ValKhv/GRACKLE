Attribute VB_Name = "#_JET"
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
'                 $$$$$$$$$$F                                       ## ##  ##  ##  ##    ##      ## ###### ######
'                  *$$$$$$$$"                                        ####  #####  ##     ##      ## ##       ##
'                    "***""               _____________                                          ## ####     ##
' STANDARD MODULE WITH DEFAULT FUNCTIONS |v 2017/03/19 |                                         ## ##       ##
' The module contains frequently used functions and is part of the G-VBA library              ####  ######   ##
'****************************************************************************************************************************************************
'****************************************************************************************************************************************************
Option Explicit

Private Const MOD_NAME As String = "#_JET"
'**********************************************************************


Public Type TLookUpInfo
    RowSource As String
    RowSourceType As String
    DisplayControl As Integer
    ColumnCount As Integer
    ColumnWidths As String
End Type

Public Type TFLD
      Name As String
      Type As DataTypeEnum
      SIZE As Integer
      Required As Boolean
      DefaultValue As String
      Description As String
      PK As Boolean
      HyperLink As Boolean
      LookUpInfo As TLookUpInfo
End Type


Public Type TSQL                    ' Type for executing Insert and Update SQL
     key As String                       ' FIELD NAME
     val As String                       ' VALUE (STRING FORMAT)
     FldType As Integer                  ' VAL TYPE (0 - string, 1 - long, 2 - bool, 3 - date, 4 - GUID,5 - dbl)
     bMultiValued As Boolean
End Type
'**********************************************************************



Public Sub Test_SQL_SELECT_Parcer()
Dim sSQL As String, sRes As String

'sSQL = "SELECT FirstName , LastName FROM Employees;"
'sSQL = "SELECT * FROM Employees;"
'sSQL = "SELECT DISTINCT TOP 2  TEST.ID, TEST.Caption FROM TEST;"
'-----------
'sSQL = "SELECT Employees.Department, Supervisors.SupvName FROM Employees INNER JOIN Supervisors WHERE Employees.Department = Supervisors.Department;"

'sSQL = "SELECT MsysObjects.Name FROM MsysObjects " & vbCrLf & _
'       "WHERE (((MsysObjects.Name) Not Like '~*' And (MsysObjects.Name) Not Like 'MSys*') AND ((MsysObjects.Type)=1) AND ((MsysObjects.Flags)=0)) " & vbCrLf & _
'       "ORDER BY MsysObjects.Name;"
'-----------
'sSQL = "SELECT BirthDate " & vbCrLf & _
'       "AS Birth FROM Employees;"
'sSQL = "SELECT ALL *  FROM Employees ORDER BY EmployeeID;"
'sSQL = "SELECT DISTINCT LastName FROM Employees;"
'----------
'sSQL = "SELECT COUNT(EmployeeID) AS HeadCount FROM Employees;"
'sSQL = "SELECT DISTINCTROW CompanyName " & vbCrLf & _
'       "FROM Customers INNER JOIN Orders " & vbCrLf & _
'       "ON Customers.CustomerID = Orders.CustomerID " & vbCrLf & _
'       "ORDER BY CompanyName;"

sSQL = "SELECT [$$OBJECTS].ObjectGroup, [$$OBJECTS].ObjectName AS [Module], [$$FLDS].FldType AS PublicPrivate, [$$FLDS].FldName AS Function, [$$FLDS].FldCategory AS DataType, [$$FLDS].Description, [$$FLDS].BODY, [$$FLDS].ID, [$$FLDS].DateCreate, [$$FLDS].META, [$$FLDS].FldSize, Trim(Nz([Objectname],"")) & " & QR(".") & " & Nz([FldName]," & QR("") & ") AS FileAs " & vbCrLf & _
       "FROM [$$FLDS] RIGHT JOIN [$$OBJECTS] ON [$$FLDS].FldParent = [$$OBJECTS].ID " & vbCrLf & _
       "Where ((([$$OBJECTS].ObjectType) = -32761)) " & vbCrLf & _
       "ORDER BY [$$OBJECTS].ObjectGroup, [$$OBJECTS].ObjectName, [$$FLDS].FldType DESC , [$$FLDS].FldName;"

'sSQL = "SELECT TEST.ID, TEST.MVF, [_LBL].Field1 " & vbCrLf & _
'       "FROM TEST LEFT JOIN _LBL ON TEST.MVF.Value = [_LBL].ID;"

'sSQL = "SELECT TOP 25 FirstName , LastName " & vbCrLf & _
'       "FROM Students " & vbCrLf & _
'       "Where GraduationYear = 1994 " & vbCrLf & _
'       "ORDER BY GradePointAverage DESC;"
       
'sSQL = "SELECT TOP 10 PERCENT " & vbCrLf & _
'       "FirstName , LastName " & vbCrLf & _
'       "FROM Students " & vbCrLf & _
'       "Where GraduationYear = 1994 " & vbCrLf & _
'       "ORDER BY GradePointAverage ASC;"
'--------------
sRes = SQL_SELECT_Parcer(sSQL, ";", "=")
Debug.Print vbCrLf & sRes

End Sub

'======================================================================================================================================================
' DLookUp Replacement
' (c) Allen Browne, November 2003.  Updated April 2010.
'======================================================================================================================================================
Public Function ELookup(Expr As String, Domain As String, Optional Criteria As Variant, _
                                                                                                          Optional OrderClause As Variant) As Variant
Dim db As DAO.Database, RS As DAO.Recordset, rsMVF As DAO.Recordset
Dim varResult As Variant, strSQL As String, strOut As String
Dim lngLen As Long              'Length of string.
    Const strcSep = ","             'Separator between items in multi-value list.

On Error GoTo ErrHandle
'-----------------------------------------------------
    varResult = Null
    strSQL = "SELECT TOP 1 " & Expr & " FROM " & Domain
    
    If Not IsMissing(Criteria) Then
        strSQL = strSQL & " WHERE " & Criteria
    End If
    If Not IsMissing(OrderClause) Then
        strSQL = strSQL & " ORDER BY " & OrderClause
    End If
    strSQL = strSQL & ";"
'-------------------------------------------------------
'Lookup the value.
    Set db = DBEngine(0)(0):   Set RS = db.OpenRecordset(strSQL, dbOpenForwardOnly)
    If RS.RecordCount > 0 Then
        
        If varType(RS(0)) = vbObject Then   'Will be an object if multi-value field
            Set rsMVF = RS(0).value
            Do While Not rsMVF.EOF
                If RS(0).Type = 101 Then        'dbAttachment
                    strOut = strOut & rsMVF!fileName & strcSep
                Else
                    strOut = strOut & rsMVF![value].value & strcSep
                End If
                rsMVF.MoveNext
            Loop
            
            lngLen = Len(strOut) - Len(strcSep) ' 'Remove trailing separator
            If lngLen > 0& Then varResult = Left(strOut, lngLen)
            Set rsMVF = Nothing
        Else
            varResult = RS(0)  'Not a multi-value field: just return the value.
        End If
    End If
    RS.Close
'-----------------------------
ExitHere:
    ELookup = varResult  '!!!!!!!!!!!!!!!!!!!!!
    Set RS = Nothing: Set db = Nothing
    Exit Function
'-----------
ErrHandle:
    ErrPrint2 "ELookup", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Set RS = Nothing: Set db = Nothing
End Function
'==========================================================================================================================================
' SQL Parcer (retrive Table Name and list of fields)
'     SELECT [predicate] { * | table.* | [table.]field1 [AS alias1] [, [table.]field2 [AS alias2] [, …]]} FROM tableexpression [, …]
'              [IN externaldatabase] [WHERE… ] [GROUP BY… ]
'              [HAVING… ] [ORDER BY… ] [WITH OWNERACCESS OPTION]
'==========================================================================================================================================
Public Function SQL_SELECT_Parcer(SQL As String, Optional DLM As String = UFDELIM, Optional SEQV As String = "=") As String

Dim sRes As String, STABLE As String, sWork As String
Dim sLeft As String, sRight As String, bJoin As Boolean

Const SQL_KEYWORDS As String = "TABLE;FIELDS;FROM;LEFT JOIN;RIGHT JOIN;INNER JOIN;IN;WHERE;GROUP BY;HAVING;ORDER BY;WITH OWNERACCESS OPTION;TOP;ALL;DISTINCT;DISTINCTROW"

On Error GoTo ErrHandle
'------------------------------------------
If InStr(1, SQL, "SELECT", vbTextCompare) <= 0 Then Err.Raise 10000, , "The SQL = " & SQL & vbCrLf & " is not SELECT QUERY"
If InStr(1, SQL, "UNION", vbTextCompare) > 0 Then Err.Raise 1000, , "Can't proceed union query"

sWork = Replace(SQL, vbCrLf, " "): sWork = Replace(sWork, ";", "")
Call DIVIDEPARTS(sWork, "FROM", sLeft, sRight): If sLeft = "" Then Err.Raise 10000, , "Can't extract from clouse from SQL" & vbCrLf & SQL
    
    STABLE = FirstWord(sRight)
    
    sRight = SQLRight(sRight, DLM, SEQV): If GetWordPosition(sRight, "Join", 1, vbTextCompare) > 0 Then bJoin = True
    sLeft = SQLLeft(sLeft, STABLE, bJoin, "FIELDS", DLM, SEQV)
'---------------------------

sRes = "TABLE" & SEQV & STABLE & DLM & sLeft & DLM & sRight
'------------------------------------------
ExitHere:
    SQL_SELECT_Parcer = sRes '!!!!!!!!!!!!!!!!!!
    Exit Function
'-----------
ErrHandle:
   ErrPrint "SQL_SELECT_Parcer", Err.Number, Err.Description
   Err.Clear
End Function

Public Sub Test_FldListToString()
Dim FLDS() As TFLD
    ReDim FLDS(3)
    FLDS(0).Name = "ID": FLDS(1).Type = dbLong: FLDS(0).SIZE = 8: FLDS(0).Required = True: FLDS(0).DefaultValue = 1: FLDS(0).Description = "The Autonumber Field"
    FLDS(1).Name = "Title": FLDS(1).Type = dbText: FLDS(1).SIZE = 255: FLDS(1).Required = False: FLDS(1).DefaultValue = "": FLDS(1).Description = "The Caption"
    FLDS(2).Name = "Value": FLDS(1).Type = dbNumeric: FLDS(1).SIZE = 1: FLDS(0).Required = False: FLDS(0).DefaultValue = "": FLDS(0).Description = "The Autonumber Field"



End Sub
'======================================================================================================================================================
' Serialize Field Array to string
'======================================================================================================================================================
Public Function FldListToString(FLDS() As TFLD, Optional DLM As String = ";", Optional SEP As String = vbCrLf, _
                                                                                                         Optional bEasyList As Boolean = True) As String
Dim nFlds As Long, I As Long, sRes As String

    On Error GoTo ErrHandle
'---------------------
    nFlds = UBound(FLDS)
    For I = 0 To nFlds
        If bEasyList Then
             sRes = sRes & DLM & FLDS(I).Name
        Else
             sRes = sRes & SEP & FLDS(I).Name & DLM & FLDS(I).Type & DLM & FLDS(I).SIZE
             sRes = sRes & DLM & FLDS(I).DefaultValue & DLM & FLDS(I).Required & DLM & FLDS(I).Description
        End If
    Next I
    If bEasyList And sRes <> "" Then
        sRes = Right(sRes, Len(sRes) - Len(DLM))
    ElseIf (Not bEasyList) And sRes <> "" Then
        sRes = Right(sRes, Len(sRes) - Len(SEP))
    End If
'---------------------
ExitHere:
    FldListToString = sRes '!!!!!!!
    Exit Function
'-----------------
ErrHandle:
    ErrPrint2 "FldListToString", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'======================================================================================================================================================
' Build Field Array from String
'======================================================================================================================================================
Public Function FldListFromString(sFlds As String, Optional DLM As String = ";", Optional SEP As String = vbCrLf, _
                                                                                                       Optional bEasyList As Boolean = True) As TFLD()
Dim FLDS() As TFLD, nFlds As Long, I As Long
Dim sWorks() As String, nWorks As Integer, sFLD() As String

    On Error GoTo ErrHandle
'---------------------
ReDim FLDS(0): nFlds = -1
If sFlds = "" Then GoTo ExitHere
    If bEasyList Then
        sWorks = Split(sFlds, DLM): nWorks = UBound(sWorks)
        For I = 0 To nWorks
            If Trim(sWorks(I)) <> "" Then
                 nFlds = nFlds + 1: ReDim Preserve FLDS(nFlds)
                 FLDS(nFlds).Name = Trim(sWorks(I))
            End If
        Next I
    Else
        sWorks = Split(sFlds, SEP): nWorks = UBound(sWorks)
        For I = 0 To nWorks
            If Trim(sWorks(I)) <> "" Then
                nFlds = nFlds + 1: ReDim Preserve FLDS(nFlds)
                sFLD = Split(sWorks(I), DLM)
                If UBound(sFLD) >= 2 Then
                    FLDS(nFlds).Name = Trim(sFLD(0))
                    If IsNumeric(sFLD(1)) Then FLDS(nFlds).Type = CInt(sFLD(1))
                    If IsNumeric(sFLD(2)) Then FLDS(nFlds).SIZE = CInt(sFLD(2))
                End If
                If UBound(sFLD) >= 5 Then
                    FLDS(nFlds).Required = GetBool(sFLD(3))
                    FLDS(nFlds).DefaultValue = Trim(sFLD(4))
                    FLDS(nFlds).Description = Trim(sFLD(5))
                End If
            End If
        Next I
    End If
'---------------------
ExitHere:
    FldListFromString = FLDS '!!!!!!!
    Exit Function
'-----------------
ErrHandle:
    ErrPrint2 "FldListFromString", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'======================================================================================================================================================
' String to Fld Array
'======================================================================================================================================================
Public Function StringToFldList(sFlds As String, Optional DLM As String = ";", Optional SEP As String = vbCrLf, _
                                                                                                         Optional bEasyList As Boolean = True) As TFLD()
Dim FLDList() As TFLD, nFld As Long

    On Error GoTo ErrHandle
'-------------------------------
    ReDim FLDList(0): nFld = -1
    If sFlds <> "" Then
        
    End If
'-------------------------------
ExitHere:
    StringToFldList = FLDList '!!!!!!!!!!!!!
    Exit Function
'---------------
ErrHandle:
    ErrPrint2 "StringToFldList", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Parsing right parts of SQL
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function SQLRight(sSQLRight As String, Optional DLM As String = UFDELIM, Optional SEQV As String = "=") As String
Dim sKey As String, sKV As String, sWork As String, sRes As String

sWork = Trim(sSQLRight): If sWork = "" Then Exit Function

sKey = "WITH OWNERACCESS OPTION": Call ExtractSQLKV(sWork, sKV, sKey, SEQV)
If sKV <> "" Then sRes = sKV

sKey = "ORDER BY": Call ExtractSQLKV(sWork, sKV, sKey, SEQV)
If sKV <> "" Then sRes = IIf(sRes <> "", sRes & DLM, "") & sKV

sKey = "HAVING": Call ExtractSQLKV(sWork, sKV, sKey, SEQV)
If sKV <> "" Then sRes = IIf(sRes <> "", sRes & DLM, "") & sKV

sKey = "GROUP BY": Call ExtractSQLKV(sWork, sKV, sKey, SEQV)
If sKV <> "" Then sRes = IIf(sRes <> "", sRes & DLM, "") & sKV

sKey = "WHERE": Call ExtractSQLKV(sWork, sKV, sKey, SEQV)
If sKV <> "" Then sRes = IIf(sRes <> "", sRes & DLM, "") & sKV

sKey = "IN": Call ExtractSQLKV(sWork, sKV, sKey, SEQV)
If sKV <> "" Then sRes = IIf(sRes <> "", sRes & DLM, "") & sKV

sKey = "FROM": sKV = sKey & SEQV & sWork
If sKV <> "" Then sRes = IIf(sRes <> "", sRes & DLM, "") & sKV

sKey = "INNER JOIN": Call ExtractSQLKV(sWork, sKV, sKey, SEQV)
If sKV <> "" Then sRes = IIf(sRes <> "", sRes & DLM, "") & sKV

sKey = "RIGHT JOIN": Call ExtractSQLKV(sWork, sKV, sKey, SEQV)
If sKV <> "" Then sRes = IIf(sRes <> "", sRes & DLM, "") & sKV

sKey = "LEFT JOIN": Call ExtractSQLKV(sWork, sKV, sKey, SEQV)
If sKV <> "" Then sRes = IIf(sRes <> "", sRes & DLM, "") & sKV
'----------------------------
ExitHere:
    SQLRight = sRes '!!!!!!!!!!!!!
End Function
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' If keywords present - return this part of sentense and KV-pair
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ExtractSQLKV(ByRef sRow As String, ByRef sKV As String, sKey As String, Optional SEQV As String = "=")
Dim iL As Integer
    sKV = "": iL = GetWordPosition(sRow, sKey, 1, vbTextCompare)
    If iL > 0 Then
        sKV = Trim(sKey & SEQV & Right(sRow, Len(sRow) - iL - Len(sKey)))
        sRow = Left(sRow, iL - 1)
    End If
End Sub
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Proceed fields list and predicates
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function SQLLeft(sRow As String, Optional STABLE As String, Optional bJoin As Boolean, Optional FLDS_KEYWORD As String = "FIELDS", _
                                                                              Optional DLM As String = UFDELIM, Optional SEQV As String = "=") As String

Dim sWork As String, sPredicate As String, sFields() As String, nDim As Integer, I As Integer, sRes As String
Const FLDS_SEP As String = ","

On Error GoTo ErrHandle
'-------------------------------------------------------------
sWork = Trim(sRow): If sWork = "" Then Exit Function
If InStr(1, sWork, "Select", vbTextCompare) = 1 Then sWork = Trim(Right(sWork, Len(sWork) - 6))

Call ProcessPredicates(sWork, sPredicate, DLM, SEQV)
If sWork = "" Then Exit Function

sFields = Split(sWork, ","): nDim = UBound(sFields)
For I = 0 To nDim
     sRes = IIf(sRes <> "", sRes & FLDS_SEP, "") & ClearFld(sFields(I), STABLE, bJoin)
Next I

If sRes <> "" Then sRes = FLDS_KEYWORD & SEQV & sRes
If sPredicate <> "" Then sRes = sPredicate & DLM & sRes
'-------------------------------------------------------------
ExitHere:
    SQLLeft = sRes '!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'------------
ErrHandle:
    ErrPrint "SQLLeft", Err.Number, Err.Description
    Err.Clear
End Function
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Clear Table Name for simple query
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function ClearFld(sFLD As String, STABLE As String, bJoin As Boolean, Optional SEQV As String = "=") As String
Dim sRes As String, sVal As String, iL As Integer

Const ASDLM As String = ":"

        sRes = Trim(sFLD)
        iL = GetWordPosition(sRes, "AS", 1, vbTextCompare)
        If iL > 0 Then
           sVal = Trim(Left(sRes, iL - 1))
           sRes = Trim(Right(sRes, Len(sRes) - iL - 1))
        End If

        sRes = Replace(sRes, "[", ""): sRes = Replace(sRes, "]", "")
        If Not bJoin Then sRes = Replace(sRes, STABLE & ".", "", , , vbTextCompare)
        If sVal <> "" Then sRes = sRes & ASDLM & sVal
'---------------------
ExitHere:
        ClearFld = sRes '!!!!!!!!!!
End Function
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Process Predicates
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ProcessPredicates(ByRef sRow As String, ByRef sPredicate As String, Optional DLM As String = UFDELIM, Optional SEQV As String = "=")
Dim iL As Integer, iR As Integer, sWork As String

If sRow = "" Then Exit Sub

        iL = GetWordPosition(sRow, "TOP", 1, vbTextCompare)
        If iL > 0 Then    ' PROCESS "TOP" predicate
               sWork = NextWord(sRow, "TOP", iL)
               If IsNumeric(Left(sWork, 1)) Then
                   iR = InStr(1, sRow, "TOP " & sWork, vbTextCompare) + Len(sWork) + 3
               Else
                   iR = iL + 3
               End If
               
                  sPredicate = Trim("TOP" & SEQV & Mid(sRow, iL + 3, iR - iL))
                  sRow = Trim(CutSubString(sRow, iL, iR))
        End If
        
        iL = GetWordPosition(sRow, "ALL", 1, vbTextCompare)
        If iL > 0 Then    ' PROCESS "TOP" predicate
                 sPredicate = IIf(sPredicate <> "", sPredicate & DLM, "") & "ALL"
                 sRow = Trim(CutSubString(sRow, iL, iL + 3))
        End If
                
        iL = GetWordPosition(sRow, "DISTINCT", 1, vbTextCompare)
        If iL > 0 Then    ' PROCESS "TOP" predicate
                 sPredicate = IIf(sPredicate <> "", sPredicate & DLM, "") & "DISTINCT"
                 sRow = Trim(CutSubString(sRow, iL, iL + 8))
        End If
        
        iL = GetWordPosition(sRow, "DISTINCTROW", 1, vbTextCompare)
        If iL > 0 Then    ' PROCESS "TOP" predicate
                 sPredicate = IIf(sPredicate <> "", sPredicate & DLM, "") & "DISTINCTROW"
                 sRow = Trim(CutSubString(sRow, iL, iL + 11))
        End If
End Sub
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Function return Left and Right parts divided by single word
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub DIVIDEPARTS(sRow As String, sWord As String, ByRef sLeft As String, ByRef sRight As String)
Dim sU() As String, nDim As Integer, sWork As String

sLeft = "": sRight = "": If sRow = "" Then Exit Sub

    sWork = SplitByWord(sRow, sWord, UFDELIM, vbTextCompare): If sWork = "" Then Exit Sub
    sU = Split(sWork, UFDELIM)
    sLeft = sU(0): sRight = sU(1)
    
End Sub
'=======================================================================================================================
' Delete All Records in the Table
'=======================================================================================================================
Public Function ClearTable(STABLE As String) As Boolean
Dim sSQL As String, bRes As Boolean
On Error GoTo ErrHandle
           sSQL = "Delete * From " & SHT(STABLE) & ";"
           DoCmd.SetWarnings False
                 CurrentDb.Execute sSQL
                 bRes = True
'----------------------------
ExitHere:
           DoCmd.SetWarnings True
           ClearTable = bRes '!!!!!!!!!!
           Exit Function
'-------------------
ErrHandle:
           ErrPrint "ClearTable", Err.Number, Err.Description
           Err.Clear: Resume ExitHere
End Function

'=====================================================================================================================================================
' SETUP TABLE PROPERTIES
'=====================================================================================================================================================
Public Sub SetTableProperty(objTableObj As Object, strPropertyName As String, _
            intPropertyType As Integer, varPropertyValue As Variant)
Dim prpProperty As Variant
Const conErrPropertyNotFound = 3270
        
  On Error Resume Next                ' Don't trap errors.
'---------------------------------
 objTableObj.Properties(strPropertyName) = varPropertyValue
        If Err <> 0 Then                    ' Error occurred when value set.
            If Err <> conErrPropertyNotFound Then
                ' Error is unknown.
                ErrPrint "SetTableProperty", Err.Number, Err.Description
                Err.Clear
            Else
                ' Error is "Property not found", so add it to collection.
                Set prpProperty = objTableObj.CreateProperty(strPropertyName, _
                    intPropertyType, varPropertyValue)
                objTableObj.Properties.Append prpProperty
                Err.Clear
            End If
        End If
        objTableObj.Properties.Refresh
End Sub
'=====================================================================================================================================================
' RESET AUTO NUBER IN TaBLE
'=====================================================================================================================================================
Public Sub ResetID(TableName As String, Optional IDFld As String = "ID")

On Error GoTo ErrHandle
'----------------------------
DoCmd.SetWarnings False
        CurrentDb.Execute "ALTER TABLE " & SHT(TableName) & " ALTER COLUMN " & IDFld & " COUNTER(1,1)"
'----------------------------
ExitHere:
       DoCmd.SetWarnings True
       Exit Sub
'-------------
ErrHandle:
       ErrPrint "ResetID", Err.Number, Err.Description
       Err.Clear: Resume ExitHere
End Sub
'=====================================================================================================================================================
' SQL NORMALIZATION FOR DATE
'=====================================================================================================================================================
Public Function SQLDate(varDate As Variant) As String
    'Purpose:    Return a delimited string in the date format used natively by JET SQL.
    'Argument:   A date/time value.
    'Note:       Returns just the date format if the argument has no time component,
    '                or a date/time format if it does.
    'Author:     Allen Browne. allen@allenbrowne.com, June 2006.
    If IsDate(varDate) Then
        If DateValue(varDate) = varDate Then
            SQLDate = Format$(varDate, "\#mm\/dd\/yyyy\#")
        Else
            SQLDate = Format$(varDate, "\#mm\/dd\/yyyy hh\:nn\:ss\#")
        End If
    End If
End Function
'=====================================================================================================================================================
' Create String Array For Simple Select SQL Statement
'      String(0)  - FDLS IN SELECT PART (, is separator)
'      String (1) - Table Name
'      String (2)  - WHERE
'      String (3) - Order By
' DEPRECATED
'=====================================================================================================================================================
Public Function SelectSQLParcer(SelectSQL As String, Optional DLM As String = ";", Optional SEQV As String = "=") As String()
Dim s_Pars As String, sRes(3) As String

Const SQL_KEYWORDS As String = "TABLE;FIELDS;FROM;LEFT JOIN;RIGHT JOIN;INNER JOIN;IN;WHERE;GROUP BY;HAVING;ORDER BY;WITH OWNERACCESS OPTION;TOP;ALL;DISTINCT;DISTINCTROW"


     On Error GoTo ErrHandle
'-----------------------------------------------
s_Pars = SQL_SELECT_Parcer(SelectSQL, ";", "=")
If s_Pars = "" Then Err.Raise 1000, , "Wrong prosessing for SQL " & vbCrLf & SelectSQL
'-----------------------------------------------------------------------------
sRes(0) = GetValueForKey(s_Pars, "FIELDS", DLM, SEQV)
sRes(1) = GetValueForKey(s_Pars, "TABLE", DLM, SEQV)
sRes(2) = GetValueForKey(s_Pars, "WHERE", DLM, SEQV)
sRes(3) = GetValueForKey(s_Pars, "ORDER BY", DLM, SEQV)

'-------------------------------------------
ExitHere:
     SelectSQLParcer = sRes '!!!
     Exit Function
'-----------------
ErrHandle:
     ErrPrint "SelectSQLParcer", Err.Number, Err.Description
     Err.Clear
End Function

Private Function ExtractTextBetweenWords(sRow As String, sLeft As String, iLeft As Integer, iRight As Integer) As String
Dim sRes As String, iStart As Integer
    
    If iLeft <= 0 Then Exit Function
    
    If iRight > 0 Then
        iStart = iLeft + Len(sLeft)
        sRes = Mid(sRow, iStart, iRight - iStart)
    Else
        iStart = iLeft + Len(sLeft)
        sRes = Right(sRow, Len(sRow) - iStart)
    End If
    
'------------------------
ExitHere:
    ExtractTextBetweenWords = Trim(sRes) '!!!!!!!!!!!!!!!!!
End Function


'===================================================================================================================================================
' SQL NORMALIZATION FOR TEXT
'===================================================================================================================================================
Public Function SQLText2(str As String) As String
Dim sLeft As String
Dim sRight As String
     
sLeft = vbNullString
sRight = str
  Do Until Len(sRight) = 0 Or InStr(1, sRight, "|") = 0
        sLeft = sLeft & Left(sRight, InStr(1, sRight, "|") - 1) '& "|"
        sRight = Right(sRight, Len(sRight) - InStr(1, sRight, "|"))
  Loop
    str = sLeft & sRight
     
    sLeft = vbNullString
    sRight = str
     
    Do Until Len(sRight) = 0 Or InStr(1, sRight, """") = 0
        sLeft = sLeft & Left(sRight, InStr(1, sRight, """")) & """"
        sRight = Right(sRight, Len(sRight) - InStr(1, sRight, """"))
    Loop
    str = sLeft & sRight
     
    sLeft = vbNullString
    sRight = str
     
    Do Until Len(sRight) = 0 Or InStr(1, sRight, "'") = 0
        sLeft = sLeft & Left(sRight, InStr(1, sRight, "'")) & "'"
        sRight = Right(sRight, Len(sRight) - InStr(1, sRight, "'"))
    Loop
str = sLeft & sRight
'----------------------------------
ExitHere:
            SQLText2 = str '!!!!!!!!!!!!!!!
     
End Function

'====================================================================================================================
' Creation of a temporary request. Calls up the name of the created request
'====================================================================================================================
Public Function CreateTempQuery(QuerySQL As String, Optional QueryName As String = "", _
                                Optional TempPrefix As String = "%", Optional OpenQuery As Boolean = False) As String
Dim sQueryName As String
Dim qdf As QueryDef

On Error GoTo ErrHandle
'---------------------------------------------
    sQueryName = IIf(QueryName = "", GetTempTableName(TempPrefix), QueryName)
    Set qdf = CurrentDb.CreateQueryDef(sQueryName, QuerySQL)
    If OpenQuery Then DoCmd.OpenQuery qdf.Name
'-----------------------------------------------
ExitHere:
    CreateTempQuery = qdf.Name  '!!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'--------------------
ErrHandle:
        MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "CreateTempQuery ERROR"
        Err.Clear
End Function

'====================================================================================================================
' Add Field To Table
'====================================================================================================================
Public Function AddFieldToTable(strTable As String, strField As String, nFieldType As Integer, _
                                                                        Optional DefaultValue As Variant) As Boolean
Dim db As DAO.Database, tdf As DAO.TableDef, FLD As DAO.Field
Dim bRes As Boolean

On Error GoTo ErrHandle
'----------------------------------
    Set db = CurrentDb: Set tdf = db.TableDefs(strTable)
    Set FLD = tdf.CreateField(strField, nFieldType)
    
    If Not IsEmpty(DefaultValue) Then FLD.DefaultValue = DefaultValue
    
    tdf.FIELDS.Append FLD
    bRes = True
'----------------------------------
ExitHere:
    AddFieldToTable = bRes '!!!!!!!
    Set tdf = Nothing
    Set db = Nothing
    Exit Function
'-------------
ErrHandle:
        ErrPrint "AddFieldToTable", Err.Number, Err.Description
        Err.Clear: Resume ExitHere
End Function
'======================================================================================================================================================
' Create Table With Structure
'======================================================================================================================================================
Public Function CreateTable(ByVal TableName As String, FLDS() As TFLD, _
                                                                              Optional DLM As String = ";", Optional SEP As String = vbCrLf) As Boolean
Dim dbsCurrent As Object, TD As TableDef, f As Field, bRes As Boolean
Dim FieldsCount As Integer, I As Integer

On Error GoTo ErrHandle
'-------------------------------------------------------------------
Set dbsCurrent = CurrentDb()
If IsTable(TableName) Then Err.Raise 10001, , "The Table " & TableName & " is already exists"
Set TD = dbsCurrent.CreateTableDef(TableName)

    FieldsCount = UBound(FLDS)
        For I = 1 To FieldsCount
            If FLDS(I).Name = "" Then GoTo NextFLD
            If FLDS(I).Name <> "ID" Then
                Set f = TD.CreateField(FLDS(I).Name, FLDS(I).Type)
                    If FLDS(I).Type = dbText Then f.SIZE = FLDS(I).SIZE
                    If FLDS(I).Type = dbLong And FLDS(I).PK = True Then f.Attributes = dbAutoIncrField
                
                
                    TD.FIELDS.Append f    ' Append Field
                    If FLDS(I).DefaultValue <> "" Then Call SetDefaultValueForField(TableName, FLDS(I).Name, FLDS(I).DefaultValue)
                    If FLDS(I).HyperLink Then Call SetHyperLink(TableName, FLDS(I).Name)
             Else               '  Create Uniq Field ID (PK)
                Set f = TD.CreateField("ID", dbLong)
                    f.Attributes = dbAutoIncrField
                    TD.FIELDS.Append f
             
             End If
NextFLD:
        Next I
    dbsCurrent.TableDefs.Append TD
'-----------------------------------------------------
ExitHere:
    CreateTable = True  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'-----------------
ErrHandle:
    Set TD = Nothing
    Set f = Nothing
End Function
'======================================================================================================================================================
' Set DefaultValue To FLD
'======================================================================================================================================================
Public Sub SetDefaultValueForField(tblName As String, FldName As String, DefaultValue As Variant)

Dim FLD As DAO.Field, prp As DAO.Property

On Error GoTo ErrHandle
'----------------------------------
    CurrentDb.TableDefs.Refresh
    CurrentDb.TableDefs(tblName).FIELDS(FldName).DefaultValue = DefaultValue
'---------------------------
ExitHere:
     Set FLD = Nothing: Set prp = Nothing
     Exit Sub
'--------------
ErrHandle:
     ErrPrint2 "SetDefaultValueForField", Err.Number, Err.Description, MOD_NAME
     Err.Clear: Resume ExitHere
End Sub

'======================================================================================================================================================
' Set FLD As HyperLink
'======================================================================================================================================================
Public Sub SetHyperLink(tblName As String, FldName As String)

On Error GoTo ErrHandle
'----------------------------------
CurrentDb.TableDefs.Refresh
CurrentDb.TableDefs(tblName).FIELDS(FldName).Attributes = dbHyperlinkField Or dbVariableField
CurrentDb.TableDefs.Refresh

'---------------------------
ExitHere:
     Exit Sub
'--------------
ErrHandle:
     ErrPrint2 "SetHyperLink", Err.Number, Err.Description, MOD_NAME
     Err.Clear
End Sub

'======================================================================================================================================================
' Set FLD As LookUp
'======================================================================================================================================================
Public Sub SetLookUpFLD(tblName As String, FldName As String, pstrRowSource As String, Optional DisplayControl As Integer = acComboBox, _
                             Optional RowSourceType As String = "Table/Query", Optional ColumnCount As Integer = 2, _
                             Optional ColumnWidths As String = "0;1440")
Dim FLD As DAO.Field, prp As DAO.Property

On Error GoTo ErrHandle
'----------------------------------
CurrentDb.TableDefs.Refresh
'----------------------------------
Set FLD = CurrentDb.TableDefs(tblName).FIELDS(FldName)

Set prp = FLD.CreateProperty("DisplayControl", dbInteger, DisplayControl)
       CurrentDb.TableDefs(tblName).FIELDS(FldName).Properties.Append prp
       
Set prp = FLD.CreateProperty("RowSourceType", dbText, RowSourceType)
       CurrentDb.TableDefs(tblName).FIELDS(FldName).Properties.Append prp
       
Set prp = FLD.CreateProperty("ColumnCount", dbInteger, ColumnCount)
       CurrentDb.TableDefs(tblName).FIELDS(FldName).Properties.Append prp
       
Set prp = FLD.CreateProperty("ColumnWidths", dbText, ColumnWidths)
       CurrentDb.TableDefs(tblName).FIELDS(FldName).Properties.Append prp
       
Set prp = FLD.CreateProperty("RowSource", dbText, pstrRowSource)
        CurrentDb.TableDefs(tblName).FIELDS(FldName).Properties.Append prp
'---------------------------
ExitHere:
     Set FLD = Nothing: Set prp = Nothing
     Exit Sub
'--------------
ErrHandle:
     ErrPrint "setLookUpFld", Err.Number, Err.Description
     Err.Clear: Resume ExitHere
End Sub


'======================================================================================================================================================
' Function get RowSource for MVF Filed/LookUp Field and parce it: TableName;Field1;Field2
'======================================================================================================================================================
Public Function GetFieldLookUpInfo(tblName As String, MVFFLdName As String, Optional DLM As String = ";", Optional SEQV As String = "=", _
                                                                                                    Optional CheckMVF As Boolean) As String
Dim STABLE As String, MVFName As String, sRowSource As String, sRowSourceType As String, iColumnCount As Integer
Dim ColumnWidths As String, DisplayControl As Integer, BoundColumn As Integer, AllowMultipleValues As String
Dim db As DAO.Database, TBL As DAO.TableDef, FLD As Field, sRes As String, sWork As String

On Error GoTo ErrHandle
'---------------------------------
    Set db = CurrentDb(): Set TBL = db.TableDefs(tblName)
    Set FLD = TBL.FIELDS(MVFFLdName)
    
    If CheckMVF Then
        If Not FLD.Type = 104 Then Err.Raise 10000, , "The field " & MVFFLdName & " is not MVF"
    End If
    sRowSource = GetAccessProp(FLD, "RowSource"): sRowSourceType = GetAccessProp(FLD, "RowSourceType")
    If sRowSource = "" Then Exit Function
    
    sWork = GetAccessProp(FLD, "ColumnCount"): If sWork <> "" Then iColumnCount = CInt(sWork)
        
    Select Case sRowSourceType
    
    Case "Table/Query":
        sRowSourceType = "Table/Query"
    Case "Value List":
        sRowSourceType = "Value List"
    Case Else
    End Select
    
    
    sRes = "RowSourceType" & SEQV & sRowSourceType & DLM & _
           "RowSource" & SEQV & sRowSource & DLM & _
           "ColumnCount" & SEQV & iColumnCount
'------------------------------------------
ExitHere:
    GetFieldLookUpInfo = sRes '!!!!!!!!!!!!!!!!!!
    Set FLD = Nothing: Set TBL = Nothing
    Exit Function
'-----------
ErrHandle:
   ErrPrint "GetFieldLookUpInfo", Err.Number, Err.Description
   Err.Clear
End Function

'====================================================================================================================
' Create Temp Table
'====================================================================================================================
Public Function CreateTempTable(Optional TableName As String = "", Optional TempPrefix As String = "%", _
         Optional FieldList As String = "[ID] AUTOINCREMENT,[T1] TEXT(255),[I1] INTEGER,[DateCreate] DATETIME", _
                                 Optional FromSource As String = "", Optional bDelIfExist As Boolean = True) As String
Dim sTableName As String, bRes As Boolean, sRes As String
Dim sSQL As String, MyJet As New cJet
        On Error GoTo ErrHandle
'-----------------------------------------------------------------
' Check if table name is set. If no - generate new
    sTableName = SHT(IIf(TableName = "", GetTempTableName(TempPrefix), TableName))
'-------------------------------------------------------------------------------------------------------------------
If Not bDelIfExist Then ' If rewrite is impossible try to rrecreate new
       Do While Not MyJet.IsTableExists(sTableName)
                sTableName = sTableName & GenRandomStr(1, True, False, False)
       Loop
End If
'------------------------------------------------------------------
If FromSource <> "" Then
      bRes = MyJet.CreateTable2(sTableName, False, bDelIfExist, FromSource)
      GoTo ExitHere
End If
'------------------------------------------------------------------
'  Создаем таблицу с нуля
bRes = MyJet.CreateTable2(sTableName, False, bDelIfExist, "", FieldList)
'-----------------------------------------------------
ExitHere:
        sRes = IIf(bRes, sTableName, "")
        CreateTempTable = sRes '!!!!!!!!!!!!!!!!!
        Set MyJet = Nothing
        Exit Function
'-------------------------
ErrHandle:
        MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "CreateTempTable ERROR"
        Err.Clear: bRes = False
        Resume ExitHere
End Function
'=====================================================================================================================================================
' Create Table Link
'=====================================================================================================================================================
Public Function CreateTableLink(strTable As String, strPath As String, strBaseTable As String) As Boolean
Dim myDB As DAO.Database, tdf As TableDef
Dim strConnect As String, bRes As Boolean
    
    On Error GoTo ErrHandle
'------------------------------------
    DoCmd.SetWarnings False
    Set myDB = CurrentDb
    Set tdf = myDB.CreateTableDef(strTable)
    
    With tdf
        .Connect = ";DATABASE=" & strPath
        .SourceTableName = strBaseTable
    End With
    
    myDB.TableDefs.Append tdf
    bRes = True
'----------------------------------
ExitHere:
    CreateTableLink = bRes '!!!!!!!!!!!
    DoCmd.SetWarnings True
    Exit Function
'----------------
ErrHandle:
    If Err = 3110 Then
        Resume ExitHere
    ElseIf Err = 3011 Then
        Resume Next
    Else
        ErrPrint2 "CreateTableLink", Err.Number, Err.Description, MOD_NAME
    End If
End Function
'=====================================================================================================================================================
' Delete Link To Table
'=====================================================================================================================================================
Public Function DeLink(tblName As String) As Boolean
Dim bRes As Boolean
    On Error GoTo ErrHandle
'---------------------
    CurrentDb.TableDefs.Delete tblName
    bRes = True
'---------------------
ExitHere:
    DeLink = bRes '!!!!!!!!!!!!!!
    Exit Function
'---------
ErrHandle:
    ErrPrint2 "DeLink", Err.Number, Err.Description, MOD_NAME
End Function
'=====================================================================================================================================================
' RESET DATABASE
'=====================================================================================================================================================
Public Sub RESETDB(Optional ResetList As String, Optional SkipList As String = "_*; $$*", Optional DLM As String = ";")
Dim TBLS() As String, I As Integer, nDim As Integer, SKIPS() As String, mDim As Integer, J As Integer
Dim bRes As Boolean, sWork As String, sFile As String

On Error GoTo ErrHandle
'-----------------------------------------------
If ResetList <> "" Then
    TBLS = Split(ResetList, DLM)
Else
    sWork = ListDBObjects(ACC_TBL_LOCAL, DLM)
    If sWork = "" Then GoTo AFTERTBLS
    TBLS = Split(sWork, DLM)
End If
nDim = UBound(TBLS)
'--------------------------------
   If SkipList <> "" Then
         SKIPS = Split(SkipList, DLM): mDim = UBound(SKIPS)
   End If
'-----------------------------------------------
sFile = SaveRelationshipsToFile()
If sFile = "" Then Err.Raise 1000, , "Failed save dblinks to file"
Call DeleteAllRelationships
   
   For I = 0 To nDim
     For J = 0 To mDim
        If (TBLS(I) Like SKIPS(J)) Then GoTo NEXTTBL
     Next J
     
     bRes = ClearTable(TBLS(I))
     If bRes Then Call ResetID(sWork)
     
NEXTTBL:
   Next I
If sFile <> "" Then Call RestoreRelationshipsFromFile(sFile)
'-----------------------------------------------
AFTERTBLS:



'--------------------
ExitHere:
    Exit Sub
'----------
ErrHandle:
    ErrPrint "RESET" & IIf(sWork <> "", " (" & sWork & ") ", ""), Err.Number, Err.Description
    Err.Clear
End Sub


'======================================================================================================================================================
' Get Table list in current or External DataBase
'======================================================================================================================================================
Public Function GetTableList(Optional sDBPath As String, Optional DLM As String = ";", Optional SkipTablePrefix As String = "MSys;%") As String
Dim db As DAO.Database, tdf As DAO.TableDef
Dim sRes As String, PRFXS() As String, nPRFXS As Integer, I As Integer
On Error GoTo ErrHandle
'---------------------------
If sDBPath <> "" Then
   If Dir(sDBPath) = "" Then Err.Raise 10001, , "Can't find the database for " & sDBPath
   Set db = OpenDatabase(sDBPath)
Else
   Set db = CurrentDb()
End If
'---------------------------
    nPRFXS = -1
    If SkipTablePrefix <> "" Then
        PRFXS = Split(SkipTablePrefix, DLM): nPRFXS = UBound(PRFXS)
    End If
    
    For Each tdf In db.TableDefs
        If nPRFXS >= 0 Then
            For I = 0 To nPRFXS
               If Left(tdf.Name, Len(PRFXS(I))) = PRFXS(I) Then GoTo NextTable
            Next I
        End If
            sRes = sRes & DLM & tdf.Name
NextTable:
    Next tdf

If sRes <> "" Then sRes = Right(sRes, Len(sRes) - Len(DLM))
'---------------------------
ExitHere:
    GetTableList = sRes '!!!!!!!!!!!!
    Exit Function
'-----------
ErrHandle:
    ErrPrint2 "GetTableList", Err.Number, Err.Description, MOD_NAME
End Function




'==============================================================================================================================================
' DELETE ALL RELATIONS
'==============================================================================================================================================
Public Function DeleteAllRelationships() As Integer
Dim rex As Relations    ' Relations of currentDB.
Dim rel As Relation     ' Relationship being deleted.
Dim iKt As Integer      ' Count of relations deleted.
Dim sMsg As String      ' MsgBox string.

On Error GoTo ErrHandle
'------------------------------------------------
    sMsg = "About to delete ALL relationships between tables in the current database." & vbCrLf & "Continue?"
    If MsgBox(sMsg, vbQuestion + vbYesNo + vbDefaultButton2, "Are you sure?") = vbNo Then
        DeleteAllRelationships = "Operation cancelled"
        Exit Function
    End If
    '--------------------
    Set rex = CurrentDb.Relations
    iKt = rex.Count
    
    Do While rex.Count > 0
        rex.Delete rex(0).Name
    Loop
'--------------------------------------------------
ExitHere:
     DeleteAllRelationships = iKt '!!!!!!!!!!!!!!!
     Set rex = Nothing
     Exit Function
'-------------
ErrHandle:
     ErrPrint "DeleteAllRelationships", Err.Number, Err.Description
     Err.Clear: Resume ExitHere
End Function
'==============================================================================================================================================
' CREATE RELATION
'==============================================================================================================================================
Public Sub CreateRelationship(relName As String, relTable As String, forgnTable As String, sAttrib As String, FLDS() As String)
Dim nDim As Integer, I As Integer, rel As DAO.Relation
Dim sFldName As String, sFrgnName As String

Const DLM As String = ";"

On Error GoTo ErrHandle
'--------------------------
nDim = UBound(FLDS)
    With CurrentDb
        Set rel = .CreateRelation(Name:=relName, Table:=relTable, foreignTable:=forgnTable, Attributes:=sAttrib)
        For I = 0 To nDim
            sFldName = Split(FLDS(I), DLM)(0): sFrgnName = Split(FLDS(I), DLM)(1)
                 
            rel.FIELDS.Append rel.CreateField(sFldName)
            rel.FIELDS(sFldName).ForeignName = sFrgnName
        Next I
        
        .Relations.Append rel
    End With
'--------------------------
ExitHere:
    Exit Sub
'------------
ErrHandle:
    ErrPrint "CreateRelationship", Err.Number, Err.Description
    Err.Clear
End Sub

'=================================================================================================================================================
' Get Table Row List
'=================================================================================================================================================
Public Function PrintRows(SQL As String, Optional FldFilter As String, Optional bWithHeader As Boolean = True, Optional bForScreen As Boolean, _
                                     Optional nRowLimit As Integer = -1, Optional DLM As String = ";", Optional SEP As String = vbCrLf) As String
Dim sRes As String, RS As DAO.Recordset, sRow As String, sFLD As String
Dim HDRS() As String, nHDRS As Integer, ROWS() As String, nRows As Long, I As Integer, iRow As Long


    On Error GoTo ErrHandle
'--------------------
'[1]    PREPARE
    If IsBlank(SQL) Then Exit Function
    ReDim HDRS(0): nHDRS = -1: ReDim ROWS(0): nRows = -1
    
    Set RS = CurrentDb.OpenRecordset(SQL)
    
    With RS
'[2]    COLLECT HEADERS
        For I = 0 To .FIELDS.Count - 1
            sFLD = .FIELDS(I).Name
            
            If Not IsBlank(FldFilter) Then
                 If Not InList(FldFilter, sFLD, DLM) Then GoTo NextField
            End If
            
            nHDRS = nHDRS + 1: ReDim Preserve HDRS(nHDRS)
            HDRS(nHDRS) = sFLD
NextField:
        Next I
    
If nHDRS < 0 Then Exit Function               ' NO ANY FLDS TO PRINT OUT
    
'[3]    COLLECT ROWS
        If Not .EOF Then
            .MoveLast: .MoveFirst
            
            Do While Not .EOF
                iRow = iRow + 1: If nRowLimit > 0 And iRow > nRowLimit Then Exit Do
                sRow = vbNullString
                
                For I = 0 To nHDRS
                        sRow = sRow & DLM & PrintDBCell(RS, HDRS(I), .FIELDS(HDRS(I)).Type)
                Next I
            
                If Not IsBlank(sRow) Then
                        nRows = nRows + 1: ReDim Preserve ROWS(nRows)
                        ROWS(nRows) = Right(sRow, Len(sRow) - Len(DLM))
                End If
                
                .MoveNext
            Loop
        End If
    End With
    
'[4]    BUILD OUTPUT
    If nRows >= 0 Then sRes = Join(ROWS, SEP)
    If bWithHeader Then sRes = Join(HDRS, DLM) & IIf(Not IsBlank(sRes), SEP & sRes, vbNullString)
    If bForScreen Then sRes = ScreenOnlyRowsFormat(sRes, bWithHeader, DLM, SEP)
'--------------------
ExitHere:
    PrintRows = sRes '!!!!!!!!!!!!!!!
    Set RS = Nothing
    Exit Function
'---------
ErrHandle:
    ErrPrint2 "PrintRows", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Resume ExitHere
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------
' Format Table for Screen Only OutPut
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function ScreenOnlyRowsFormat(sRows As String, Optional bWithHeader As Boolean = True, Optional DLM As String = ";", _
                                                                                                            Optional SEP As String = vbCrLf) As String
Dim sRes As String, ROWS() As String, nRows As Long, COLS() As String, nCols As Integer, WDTHCOLS() As Integer
Dim I As Long, J As Integer, IWidth As Integer

Const SEP_FORMAT As String = vbCrLf
Const DLM_FORMAT As String = "   "
Const SEP_LINE As String = "-"

    On Error Resume Next
'---------------------------
'[1]    PREPARATION
    If IsBlank(sRows) Then Exit Function
    ROWS = Split(sRows, SEP): nRows = UBound(ROWS)
    nCols = UBound(Split(ROWS(0), DLM)): ReDim WDTHCOLS(nCols)
    
'[2]    FIRST ITERATION - SEARCH FOR COLUMN WIDTH
    For I = 0 To nRows
            If Not IsBlank(ROWS(I)) Then
                COLS = Split(ROWS(I), DLM)
                For J = 0 To nCols
                    IWidth = Len(COLS(J))
                    If IWidth > WDTHCOLS(J) Then WDTHCOLS(J) = IWidth
                Next J
            End If
    Next I
    
'[3]    SECOND ITERATION - FORMAT STRING (ALIGN COLUMNS)
    For I = 0 To nRows
            If Not IsBlank(ROWS(I)) Then
                COLS = Split(ROWS(I), DLM)
                For J = 0 To nCols
                    IWidth = Len(COLS(J))
                    If IWidth < WDTHCOLS(J) Then COLS(J) = COLS(J) & String(WDTHCOLS(J) - IWidth, " ")
                Next J
                ROWS(I) = Join(COLS, DLM_FORMAT)
                If I = 0 And bWithHeader Then ROWS(I) = ROWS(I) & SEP_FORMAT & String(Len(ROWS(I)), SEP_LINE)
            End If
    Next I
    
    sRes = Join(ROWS, SEP_FORMAT)
'---------------------------
ExitHere:
     ScreenOnlyRowsFormat = sRes '!!!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Get Fiekd Cell Value
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function PrintDBCell(ByRef RS As DAO.Recordset, FldName As String, Optional FieldType As Integer) As String
Dim sRes As String, FldType As Integer, rst As DAO.Recordset

Const DLM As String = ","
Const TEXT_LIMIT As Integer = 15

    On Error Resume Next
'-------------------------------
    Select Case FieldType
    Case dbAttachment:                                                                                      ' Attachment
           Set rst = RS(FldName).value
           If rst.RecordCount <> 0 Then
               rst.MoveLast
               sRes = SHT(rst.RecordCount)
           Else
               sRes = "[0]"
           End If
    Case dbComplexByte, dbComplexInteger, dbComplexLong, dbComplexSingle, dbComplexDouble, dbComplexGUID:   'MVF
           Set rst = RS(FldName).value
           With rst
                If Not .EOF Then
                    .MoveLast: .MoveFirst
                    Do While Not .EOF
                        sRes = sRes & DLM & Nz(.FIELDS(0).value, "")
                        .MoveNext
                    Loop
                End If
           End With
           If Not IsBlank(sRes) Then sRes = Right(sRes, Len(sRes) - Len(DLM))
    
     Case dbBinary, dbLongBinary:                                                                           ' BINARY
            sRes = "{" & RS(FldName).SIZE & "}"
     Case dbMemo:                                                                                           ' MEMO/HYPERLINK
            sRes = Nz(RS.FIELDS(FldName).value, "")
     Case Else                                                                                              ' REGULAR
            sRes = CStr(Nz(RS.FIELDS(FldName).value, ""))
     End Select
     
     If Len(sRes) > TEXT_LIMIT Then sRes = Left(sRes, TEXT_LIMIT) & ".."
'-------------------------------
ExitHere:
    PrintDBCell = sRes '!!!!!!!!!!!!!
    Set rst = Nothing
    Exit Function
'------------
ErrHandle:
    Err.Clear
    sRes = "###": Resume ExitHere
End Function

'============================================================================================================================
' By Tablename, ID and Attach Field add image from buffer
'===================================================================================================sTopic====================
Public Function AttachFromClipboard(TableName As String, id As Long, AttachFld As String) As Boolean
Dim bRes As Boolean, RS As DAO.Recordset, rsAttach As DAO.Recordset
Dim TempPath As String, iCount As Integer

Const sExt As String = "PNG"
Const PRFX As String = "attch"
On Error GoTo ErrHandle
'-------------------------------------------
        TempPath = GetTempFile(PRFX, sExt): If TempPath = "" Then Exit Function
        TempPath = ImageFromClipboard(TempPath): If Dir(TempPath) = "" Then Exit Function
        
        Set RS = CurrentDb.OpenRecordset(TableName)
        With RS
                .MoveLast: .MoveFirst
                .Index = "PrimaryKey"
                .Seek "=", id
                If .NoMatch Then
                   GoTo ExitHere
                Else
                   .Edit
                   Set rsAttach = .FIELDS(AttachFld).value
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

'============================================================================================================================
' Add To Table KVString KEY1=VAL1;KEY2=VAL2;
'============================================================================================================================
Public Function AddRecord(sTableName As String, sKV As String, _
                                          Optional KVDELIM As String = ";", Optional KDelim As String = "=") As Boolean
Dim bRes As Boolean
Dim sPairs() As String, sKeys As String, sValues As String
Dim mKV() As String, sWork As String
Dim nDim As Integer, I As Integer, sSQL As String
On Error GoTo ErrHandle
'----------------------------------------------------------------------
If sTableName = "" Then Exit Function
If sKV = "" Then Exit Function
sPairs = Split(sKV, KVDELIM): nDim = UBound(sPairs)
For I = 0 To nDim
    mKV = Split(sPairs(I), KDelim)
    sWork = SQLFormat(mKV(1), GetFieldType(sTableName, mKV(0)))
    If sWork <> "" Then
        sKeys = IIf(sKeys <> "", sKeys & "," & mKV(0), mKV(0))
        sValues = IIf(sValues <> "", sValues & "," & sWork, sWork)
    End If
Next I
'-------------------------------------------
If sKeys = "" Or sValues = "" Then GoTo ExitHere
sSQL = "INSERT INTO " & SHT(sTableName) & "(" & sKeys & ") VALUES (" & sValues & ");"
DoCmd.SetWarnings False
    CurrentDb.Execute sSQL
DoCmd.SetWarnings True
bRes = True
'--------------------------------------------
ExitHere:
    AddRecord = bRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'------------------
ErrHandle:
    Debug.Print String(40, "#") & vbCrLf & "ERR#" & Err.Number & vbCrLf & Err.Description & vbCrLf & String(40, "#")
    Err.Clear
    Resume ExitHere
End Function

'============================================================================================================================
' Функция обновляет данные, переданные в формате KVString KEY1=VAL1;KEY2=VAL2
'============================================================================================================================
Public Function UpdateRecord(sTableName As String, id As Long, sKV As String, _
                                          Optional KVDELIM As String = ";", Optional KDelim As String = "=") As Boolean
Dim bRes As Boolean
Dim sPairs() As String
Dim mKV() As String, sWork As String
Dim nDim As Integer, I As Integer, sSQL As String
On Error GoTo ErrHandle
'----------------------------------------------------------------------
If sTableName = "" Then Exit Function
If sKV = "" Then Exit Function
sPairs = Split(sKV, KVDELIM): nDim = UBound(sPairs)
For I = 0 To nDim
    mKV = Split(sPairs(I), KDelim)
    sWork = SQLFormat(mKV(1), GetFieldType(sTableName, mKV(0)))
    If sWork <> "" Then
        sSQL = IIf(sSQL <> "", sSQL & "," & mKV(0) & KDelim & sWork, mKV(0) & KDelim & sWork)
    End If
Next I
If sSQL = "" Then GoTo ExitHere
sSQL = "Update " & SHT(sTableName) & " SET " & sSQL & _
      " WHERE (ID=" & id & ");"
DoCmd.SetWarnings False
    CurrentDb.Execute sSQL
DoCmd.SetWarnings True
bRes = True
'--------------------------------------------
ExitHere:
    UpdateRecord = bRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'------------------
ErrHandle:
    Debug.Print String(40, "#") & vbCrLf & "ERR#" & Err.Number & vbCrLf & Err.Description & vbCrLf & String(40, "#")
    Err.Clear
    Resume ExitHere
End Function
'======================================================================================================================================================
' Create List of Tables/ Queries
'======================================================================================================================================================
Public Function AllTableListing(Optional PrefixExclude As String = "%;~;_;$$", Optional DLM As String = ";", Optional AddExtesion As Boolean) As String
Dim SQL As String, RS As DAO.Recordset
Dim sRes As String, sTableName As String, sExt As String

    On Error GoTo ErrHandle
'-----------------------
SQL = "SELECT MSysObjects.Database, MSysObjects.ForeignName, MSysObjects.Name, MSysObjects.Type " & _
      "FROM MSysObjects " & _
      "WHERE (((MSysObjects.Type)=1 Or (MSysObjects.Type)=5 Or (MSysObjects.Type)=6) AND ((Left([Name],4))<>" & sCH("MSys") & "));"
Set RS = CurrentDb.OpenRecordset(SQL)
With RS
    If Not .EOF Then
         .MoveLast: .MoveFirst
         Do While Not .EOF
            sTableName = Nz(RS.FIELDS("Name").value, "")
            
            If AddExtesion Then
                 'Debug.Print RS.Fields("Type").Value
                 If Nz(RS.FIELDS("Type").value, 0) = 5 Then
                      sExt = ".qry"
                 Else
                      sExt = ".tbl"
                 End If
            End If
            
            If sTableName <> "" Then
                If Left(sTableName, 2) <> "f_" And Not IsPrefixInList(PrefixExclude, sTableName, DLM) Then sRes = sRes & DLM & sTableName & sExt
            End If
            .MoveNext
         Loop
    End If
End With
If sRes <> "" Then sRes = Right(sRes, Len(sRes) - Len(DLM))
'-----------------------
ExitHere:
    AllTableListing = sRes '!!!!!!!!!!!!!!
    Set RS = Nothing
    Exit Function
'------------
ErrHandle:
    ErrPrint2 "AllTableListing", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Resume ExitHere
End Function

'========================================================================================================================
'Функция приводит текст к формату, совместимому с Access SQL
'PARAMS:  sVAR - входящий текст
'RETURN:  возвращает строковое значение, отформатированное в текст, заключенный в кавычки
'VERSION: 0.1           #2015.12.31
'========================================================================================================================
Public Function SQLText(sText As String) As String
Dim sRes As String

If Left(sText, 1) = "'" And Right(sText, 1) = "'" Then
    sRes = Mid(sText, 2, Len(sText) - 2)
Else
    sRes = sText
End If
'------------------------------------------------------------------------
sRes = Replace(sRes, "'", "''")
sRes = Replace(sRes, "\""", """")
sRes = Replace(sRes, """", "\""")
'-------------------------------------------------
ExitHere:
    SQLText = "'" & sRes & "'"   '!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'========================================================================================================================
'Функция приводит GUID к формату, совместимому с Access SQL
'PARAMS:  sVAR - входящая строка в виде GUID
'RETURN:  возвращает строковое значение, отформатированное в уникальный номер, принимаемый как репликатор
'VERSION: 0.1           #2015.12.31
'========================================================================================================================
Public Function SQLGUID(vVal As Variant) As String
Dim sRes As String                                       ' Возвращаемый результат
Dim sWork As String                                      ' Рабочая строка
'sRes = StringFromGUID(vVal)                             ' Осуществляем стандартное преобразование
 sRes = CStr(vVal)                                       ' Сначал для уверенности преобразовываем в строку
'------------------------------------------------------------------------------------------------
' Подразумеваем, что строка точно GUID
    sWork = sRes                                                  ' Удаляем все кавычки, если они имеются
    If Left(sWork, 7) = "{guid {" And Right(sWork, 2) = "}}" Then           ' Нормальная форма GUID'а
        sRes = sWork
    ElseIf Left(sWork, 1) = "{" And Right(sWork, 1) = "}" Then              ' GUID представлен VALUE- строкой
        sRes = "{guid " & sWork & "}"
    Else
        sRes = ""
    End If
'----------------------------------------------------------------------
        SQLGUID = sRes  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function

'======================================================================================================================================================
' Delete Table
'======================================================================================================================================================
Public Function DeleteTable(STABLE As String, Optional bConfirm As Boolean = True) As Boolean

On Error GoTo ErrHandle
            If bConfirm Then
                If MsgBox("Do you really want to delete table " & STABLE & "?", vbYesNoCancel, "Delete Table") <> vbYes Then
                     Exit Function
                End If
            End If
            DoCmd.DeleteObject acTable, STABLE
'----------------------------
ExitHere:
           DeleteTable = True '!!!!!!!!!!
           Exit Function
'-------------------
ErrHandle:
           ErrPrint "DeleteTable", Err.Number, Err.Description
           Err.Clear
End Function

'======================================================================================================================================================
' Create Field List of Tables/ Queries
'======================================================================================================================================================
Public Function TableQueryFieldList(sTableQuery As String, Optional DLM As String = ";") As String
Dim sRes As String
Dim db As DAO.Database, RS As DAO.Recordset, FLD As Field

    On Error GoTo ErrHandle
'---------------------
If sTableQuery = "" Then Exit Function
Set db = CurrentDb()
'---------------------
Set RS = CurrentDb.OpenRecordset(sTableQuery)

    For Each FLD In RS.FIELDS
        sRes = sRes & DLM & FLD.Name
    Next FLD

If sRes <> "" Then sRes = Right(sRes, Len(sRes) - Len(DLM))
'---------------------
ExitHere:
    TableQueryFieldList = sRes '!!!!!!!!!!!
    Set FLD = Nothing: Set RS = Nothing: Set db = Nothing
    Exit Function
'---------
ErrHandle:
    ErrPrint2 "TableQueryFieldList", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'========================================================================================================================
' Quick Insert up to 3 flds
'========================================================================================================================
Public Function QuickInsert(TableName As String, FLD1 As String, val1 As String, Optional FRMT1 As Integer = 0, _
                                   Optional FLD2 As String, Optional val2 As String, Optional FRMT2 As Integer = 0, _
                                 Optional FLD3 As String, Optional val3 As String, Optional FRMT3 As Integer = 0, _
                                 Optional sHASH As String) As Long
Dim KV(3) As TSQL, sSQL As String
Dim HSH As String, IDD As Long, bHASH As Boolean
On Error GoTo ErrHandle
'-----------------------------------------------
       bHASH = IsField(TableName, "HASH")
'-----------------------------------------------
       KV(0).key = FLD1: KV(0).val = val1: KV(0).FldType = FRMT1
       KV(1).key = FLD2: KV(1).val = val2: KV(1).FldType = FRMT2
       KV(2).key = FLD1: KV(2).val = val3: KV(2).FldType = FRMT3

       If bHASH Then
                 HSH = IIf(sHASH = "", GetHASH("un"), sHASH)
                 KV(3).key = "HASH": KV(3).val = HSH
       End If

       sSQL = BuildInsert(TableName, KV)
'-----------------------------------------------
       DoCmd.SetWarnings False
       CurrentDb.Execute sSQL
       '----------------------------------------
       If bHASH And HSH <> "" Then
           IDD = Nz(DLookup("ID", TableName, "HASH =" & SQLText(HSH)), -1)
       Else
           IDD = Nz(DLookup("ID", TableName, FLD1 & " = " & FormatFieldValue(val1, FRMT1) & _
                                        IIf(FLD2 = "", "", " AND (" & FLD2 & " = " & FormatFieldValue(val2, FRMT2) & ")")), -1)
       End If
'-----------------------------------------------
ExitHere:
       QuickInsert = IDD '!!!!!!!!!!!!!
       DoCmd.SetWarnings True
       Exit Function
'-------------------
ErrHandle:
       ErrPrint "QuickInsert", Err.Number, Err.Description
       Err.Clear: Resume ExitHere
End Function
'========================================================================================================================
' Build Insert SQL
'========================================================================================================================
Public Function BuildInsert(TableName As String, Pairs() As TSQL) As String
Dim sRes As String, Keys As String, VALS As String
Dim nDim As Integer, I As Integer

Const DLM As String = ","

On Error GoTo ErrHandle
'----------------------------------
nDim = UBound(Pairs)


For I = 0 To nDim
    If Pairs(I).key = "" Then GoTo NextRow
    If Pairs(I).val = "" Then GoTo NextRow
    '------------------------------------------------------------------------------
    Keys = Keys & Pairs(I).key & DLM
    Select Case Pairs(I).FldType
    Case 0: 'STRING
         VALS = VALS & SQLText(Pairs(I).val) & DLM
    Case 1: ' INT
         VALS = VALS & Pairs(I).val & DLM
    Case 2: ' bool
         VALS = VALS & Pairs(I).val & DLM
    Case 3: ' Date
         VALS = VALS & SQLDate(Pairs(I).val) & DLM
    Case 4: ' GUID
         VALS = VALS & SQLGUID(Pairs(I).val) & DLM
    Case Else
         VALS = VALS & SQLText(Pairs(I).val) & DLM
    End Select
    '------------------------------------------------------------------------------
NextRow:
Next I

VALS = Left(VALS, Len(VALS) - Len(DLM))
Keys = Left(Keys, Len(Keys) - Len(DLM))

sRes = "INSERT INTO " & SHT(TableName) & "(" & Keys & ") VALUES (" & VALS & ");"
'----------------------------------
ExitHere:
        BuildInsert = sRes '!!!!!!!!!!!!!!
        Exit Function
'-------------------
ErrHandle:
        ErrPrint "BuildInsert", Err.Number, Err.Description
        Err.Clear: Resume ExitHere
End Function
'========================================================================================================================
' Close all DB Connection (Just for Current DB
'========================================================================================================================
Public Sub StopConnection()
Dim SQL As String

On Error Resume Next
SQL = "ALTER DATABASE " & CurrentDb.Name & " SET OFFLINE WITH ROLLBACK IMMEDIATE;"
DoCmd.SetWarnings False
    CurrentDb.Execute SQL
DoCmd.SetWarnings True
   
End Sub
'========================================================================================================================
' Build Update SQL
'========================================================================================================================
Public Function BuildUpdate(TableName As String, Pairs() As TSQL, sWhere As String) As String
Dim sRes As String
Dim nDim As Integer, I As Integer

Const DLM As String = ","

On Error GoTo ErrHandle
'----------------------------------
nDim = UBound(Pairs)


For I = 0 To nDim
    If Pairs(I).key = "" Then GoTo NextRow
    '------------------------------------------------------------------------------
    Select Case Pairs(I).FldType
    Case 0: 'STRING
         sRes = sRes & Pairs(I).key & " = " & SQLText(Pairs(I).val) & DLM
    Case 1: ' INT
         sRes = sRes & Pairs(I).key & " = " & Pairs(I).val & DLM
    Case 2: ' bool
         sRes = sRes & Pairs(I).key & " = " & Pairs(I).val & DLM
    Case 3: ' Date
         sRes = sRes & Pairs(I).key & " = " & SQLDate(Pairs(I).val) & DLM
    Case 4: ' GUID
         sRes = sRes & Pairs(I).key & " = " & SQLGUID(Pairs(I).val) & DLM
    Case Else
         sRes = sRes & Pairs(I).key & " = " & SQLText(Pairs(I).val) & DLM
    End Select
    '------------------------------------------------------------------------------
NextRow:
Next I

sRes = Left(sRes, Len(sRes) - Len(DLM))
sRes = "UPDATE " & SHT(TableName) & " SET " & sRes & " WHERE (" & sWhere & ") ;"
'----------------------------------
ExitHere:
        BuildUpdate = sRes '!!!!!!!!!!!!!!
        Exit Function
'-------------------
ErrHandle:
        ErrPrint "BuildUpdate", Err.Number, Err.Description
        Err.Clear: Resume ExitHere
End Function

'========================================================================================================================
' Add Data to Table with MVF (NO UPDATE)
'========================================================================================================================
Public Function AddToTable(TableName As String, KVS() As TSQL, KeyFLDName As String) As Long
Dim nDim As Integer, I As Integer, iRes As Long
Dim RS As DAO.Recordset, rsm As DAO.Recordset, KeyValue As String, sWhere As String

On Error GoTo ErrHandle
'------------------------------------------------------
iRes = -1: nDim = UBound(KVS)
Set RS = CurrentDb.OpenRecordset(TableName)
      With RS
          If Not .EOF Then        ' REFRESH RS
              .MoveLast: .MoveFirst
          End If
              .AddNew             ' Start Editing
              For I = 0 To nDim
                    If KVS(I).key = "" Then GoTo NextFLD
                    If KVS(I).val = "" Then GoTo NextFLD
                    If KeyFLDName = KVS(I).key Then
                        Select Case KVS(I).FldType
                        Case 0:
                            sWhere = KVS(I).key & " = " & SQLText(KVS(I).val)
                        Case 1:
                            sWhere = KVS(I).key & " = " & KVS(I).val
                        Case 2:
                            sWhere = KVS(I).key & " = " & KVS(I).val
                        Case 3:
                            sWhere = KVS(I).key & " = " & SQLDate(KVS(I).val)
                        Case 4:
                            sWhere = KVS(I).key & " = " & SQLGUID(KVS(I).val)
                        Case 5:
                            sWhere = KVS(I).key & " = " & KVS(I).val
                        End Select
                    End If
                    If KVS(I).bMultiValued Then
                        Set rsm = RS.FIELDS(KVS(I).key).value
                        With rsm
                            If .RecordCount > 0 Then
                               .Edit
                            Else
                               .AddNew
                            End If
                            Select Case KVS(I).FldType
                            Case 0:
                                rsm.FIELDS(0).value = KVS(I).val
                            Case 1:
                                rsm.FIELDS(0).value = CLng(KVS(I).val)
                            Case 2:
                                rsm.FIELDS(0).value = CBool(KVS(I).val)
                            Case 3:
                                rsm.FIELDS(0).value = CDate(KVS(I).val)
                            Case 4:
                                rsm.FIELDS(0).value = KVS(I).val
                            Case 5:
                                rsm.FIELDS(0).value = CDbl(KVS(I).val)
                            End Select
                            .Update: .Close
                        End With
                    Else
                        Select Case KVS(I).FldType
                        Case 0:
                           RS.FIELDS(KVS(I).key) = KVS(I).val
                        Case 1:
                           RS.FIELDS(KVS(I).key) = CLng(KVS(I).val)
                        Case 2:
                           RS.FIELDS(KVS(I).key) = CBool(KVS(I).val)
                        Case 3:
                           RS.FIELDS(KVS(I).key) = CDate(KVS(I).val)
                        Case 4:
                           RS.FIELDS(KVS(I).key) = KVS(I).val
                        Case 5:
                           RS.FIELDS(KVS(I).key) = CDbl(KVS(I).val)
                        End Select
                    End If
                    
NextFLD:
              Next I
              .Update: .Close
      End With
      
      If sWhere <> "" Then iRes = DLookup("ID", TableName, sWhere)
'------------------------------------------------------
ExitHere:
      AddToTable = iRes '!!!!!!!!!!!!!!!
      Set RS = Nothing: Set rsm = Nothing
      Exit Function
'---------------------------------------
ErrHandle:
      ErrPrint "AddToTable", Err.Number, Err.Description
      Err.Clear: Resume ExitHere
End Function

Public Sub TEST_RecordsetToExternalDatabase()
Dim STABLE As String, sPath As String, RS As DAO.Recordset
Dim SQL As String, bRes As Boolean

    STABLE = "ITEMS"
    sPath = OpenDialog(GC_FILE_PICKER, "Pick he DB", , False, CurrentProject.Path)
    If sPath = "" Then Exit Sub
    SQL = "SELECT * FROM " & STABLE & " WHERE ID = 15"
    Set RS = CurrentDb.OpenRecordset(SQL)
    bRes = RecordsetToExternalDatabase(RS, sPath, STABLE)
    Debug.Print bRes
End Sub
'======================================================================================================================================================
' Copy Current Recordset to external database with similar structure
'======================================================================================================================================================
Public Function RecordsetToExternalDatabase(ByRef RS As DAO.Recordset, sExternalPath As String, sExternalTbl As String) As Boolean
Dim rs2 As DAO.Recordset, bRes As Boolean

    On Error GoTo ErrHandle
'---------------------------
    If RS.RecordCount <= 0 Then Exit Function
    If Dir(sExternalPath) = "" Then Err.Raise 10005, , "Wrong path to DB " & vbCrLf & sExternalPath
    
    Set rs2 = GetExternalRecordset(sExternalPath, sExternalTbl)
    bRes = RecordsetCopy(RS, rs2)
'---------------------------
ExitHere:
    RecordsetToExternalDatabase = bRes '!!!!!!!!!!!!
    Exit Function
'---------
ErrHandle:
    ErrPrint2 "RecordsetToExternalDatabase", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'======================================================================================================================================================
' Get Recordset Table for external database
'======================================================================================================================================================
Public Function GetExternalRecordset(sPath As String, STABLE As String, Optional bReadOnly As Boolean = False) As DAO.Recordset
Dim RS As DAO.Recordset, db As DAO.Database
Dim wsAccess As Workspace

    On Error GoTo ErrHandle
'---------------------------
    Set wsAccess = DBEngine(0)
    Set db = wsAccess.OpenDatabase(sPath, False, bReadOnly)
    Set RS = db.OpenRecordset(STABLE)
'---------------------------
ExitHere:
    Set GetExternalRecordset = RS '!!!!!!!!!!!!
    Set db = Nothing: Set wsAccess = Nothing
    Exit Function
'---------
ErrHandle:
    ErrPrint2 "GetExternalRecordset", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Recordset Copy RS1 (not Empty) To RS2 - add records
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function RecordsetCopy(ByRef RS1 As DAO.Recordset, ByRef rs2 As DAO.Recordset) As Boolean
Dim J As Integer, sField As String, childRS As DAO.Recordset
Dim RS2Fields As String, dict As New cDictionary, nFlds As Integer
    
Const DLM As String = ";"

    On Error GoTo ErrHandle
'-----------------------------
    nFlds = RS1.FIELDS.Count - 1: If nFlds < 0 Then Exit Function
    RS2Fields = GetRSFields(rs2, DLM): If RS2Fields = "" Then Exit Function
    
    Call dict.SetKVString(RS2Fields, , DLM)
    
    With RS1
        If Not .EOF Then
            .MoveLast: .MoveFirst
            Do While Not .EOF
               rs2.AddNew
                 For J = 0 To nFlds
                    sField = .FIELDS(J).Name:  If sField = "ID" Then GoTo NextFLD   ' SKIP ID FIELD
                    Debug.Print sField
                    'If sField = "Attachments" Then Debug.Assert False
                    If dict.Exists(sField) Then
                         If .FIELDS(J).Type = 104 Then                              ' MVF - create sub recordset
                                Call CopyMVF(RS1, rs2, sField)
                         ElseIf .FIELDS(J).Type = 101 Then                          ' Attachment Field
                                Call CopyATTACHMENT(RS1, rs2, sField)
                         Else                                                       ' Regular Field
                           If CStr(Nz(.FIELDS(J).value, "")) <> "" Then
                                      rs2.FIELDS(sField).value = .FIELDS(J).value
                           End If
                         End If
                    End If
NextFLD:
                 Next J
               rs2.Update
               .MoveNext
            Loop
        End If
    End With
'-----------------------------
ExitHere:
    RecordsetCopy = True '!!!!!!!!!!!!!!!!
    Set childRS = Nothing: Set dict = Nothing
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "RecordsetCopy", Err.Number, "For copiing " & sField & ": " & Err.Description, MOD_NAME
    Err.Clear
End Function

'======================================================================================================================================================
' Get all fields in Recordset in String Array
'======================================================================================================================================================
Public Function GetRSFields(ByRef RS As DAO.Recordset, Optional DLM As String = ";") As String
Dim sARR() As String, I As Integer, nDim As Integer
    
    On Error GoTo ErrHandle
'------------------------
    nDim = RS.FIELDS.Count - 1: ReDim sARR(nDim)
    For I = 0 To nDim
        sARR(I) = RS.FIELDS(I).Name
    Next I
'------------------------
ExitHere:
    GetRSFields = Join(sARR, DLM) '!!!!!!!!!
    Exit Function
'-------------
ErrHandle:
    ErrPrint2 "GetRSFields", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'======================================================================================================================================================
' SQL  to delimited string
'======================================================================================================================================================
Public Function RecordsToString(TBL As String, FLD As String, Optional sCriteria As String, Optional DLM As String = ";") As String

Dim RS As Recordset, sRes As String, SQL As String


On Error GoTo ErrHandle
'---------------------------------
    
    If sCriteria = "" Then
       SQL = "SELECT " & FLD & " FROM " & SHT(TBL) & ";"
    Else
       SQL = "SELECT " & FLD & " FROM " & SHT(TBL) & " WHERE(" & sCriteria & ");"
    End If

    Set RS = CurrentDb.OpenRecordset(SQL)
    With RS
        If Not .EOF Then
              .MoveLast: .MoveFirst
              Do While Not .EOF
                  sRes = sRes & Trim(CStr(RS.FIELDS(FLD).value)) & DLM
                  .MoveNext
              Loop
        End If
    End With
'---------------------------------
ExitHere:
    Set RS = Nothing
    If sRes <> "" Then sRes = Left(sRes, Len(sRes) - Len(DLM))
    RecordsToString = sRes '!!!!!!!!!!!!!!
    Exit Function
'----------------
ErrHandle:
    ErrPrint "RecordsToString", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
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
                                                                                                  Optional sModName As String = "#_JET") As String
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

'----------------------------------------------------------------------------------------------------------------------------
' Temp Table Name creation
'----------------------------------------------------------------------------------------------------------------------------
Private Function GetTempTableName(Optional TempPrefix As String = "%", Optional bTimeStamp As Boolean = True) As String
Const CountLimit As Integer = 10
Dim sRes As String, iCount As Integer
Dim TimeStamp As String
        TimeStamp = IIf(bTimeStamp, Format(Now(), "ssnnhhddmmyyyy"), "")
        Do While iCount < CountLimit
                iCount = iCount + 1
                sRes = TempPrefix & TimeStamp & GenRandomStr(4, False, True, True) & GenRandomStr(4, True, False, False)
                If IsTable(sRes) Then  ' Такой объект существует, пробуем иное имя (до 10 попыток)
                   sRes = ""
                Else
                   Exit Do
                End If
        Loop
'-------------------------------------------------------------
        GetTempTableName = sRes      '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'======================================================================================================================================================
' Return Field Type
'======================================================================================================================================================
Public Function GetFieldType(STABLE As String, sField As String) As Integer
Dim nRes As Integer

On Error GoTo ErrHandle
    nRes = -1
    nRes = CurrentDb.TableDefs(SHT(STABLE)).FIELDS(sField).Type
'------------------------------------
ExitHere:
    GetFieldType = nRes '!!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'------------------
ErrHandle:
    Err.Clear
    Resume ExitHere
End Function

'======================================================================================================================================================
' Get Table Field Description
'======================================================================================================================================================
Public Function GetFieldDescription(FLD As Variant, Optional TBL As String) As String
Dim sRes As String, sFldName As String, oFLD As Field
   
   On Error GoTo ErrHandle
'---------------------------------
If varType(FLD) = vbString Then
    If TBL = "" Then Err.Raise 10001, , "Can't recognize the table name: " & TBL
    sRes = ""
    sFldName = CStr(FLD)
    sRes = GetAccessProp(CurrentDb.TableDefs(TBL).FIELDS(sFldName), "Description")
ElseIf varType(FLD) = vbObject Or varType(FLD) = vbDataObject Then
    Set oFLD = FLD
    sRes = GetAccessProp(oFLD, "Description")
Else
    Err.Raise 10001, , "Wrong function params"
End If
'---------------------------------
ExitHere:
    GetFieldDescription = sRes '!!!!!!!
    Set oFLD = Nothing
    Exit Function
'-----------
ErrHandle:
    ErrPrint2 "GetFieldDescription", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'======================================================================================================================================================
' Set Table Field Description
'======================================================================================================================================================
Public Sub SetFieldDescription2(sDescription As String, FLD As Variant, Optional TBL As String)
Dim sFldName As String, oFLD As Field

   On Error GoTo ErrHandle
'---------------------------------
    If sDescription = "" Then Exit Sub
    
If varType(FLD) = vbString Then
    If TBL = "" Then Err.Raise 10001, , "Can't recognize the table name: " & TBL
    sFldName = CStr(FLD)
    
    SetAccessProp CurrentDb.TableDefs(TBL).FIELDS(sFldName), "Description", sDescription, dbText

ElseIf varType(FLD) = vbObject Or varType(FLD) = vbDataObject Then
    Set oFLD = FLD
    SetAccessProp oFLD, "Description", sDescription, dbText
Else
    Err.Raise 10001, , "Wrong function params"
End If
'---------------------------------
ExitHere:
    Set oFLD = Nothing
    Exit Sub
'-----------
ErrHandle:
    ErrPrint2 "GetFieldDescription", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Sub

'======================================================================================================================================================
' Copy QUERY TO TABLE
'======================================================================================================================================================
Public Function CopyQueryToNewTable(sQueryName As String, sNewTableName As String, Optional ObjectType As AcObjectType = acQuery) As String
Dim sRes As String

    On Error GoTo ErrHandle
'----------------------
    If IsTable(sNewTableName) Then
       If Not CloseOpenTableQuery(sQueryName, ObjectType) Then Err.Raise 10001, , "Can't close the object " & sQueryName
       
       If MsgBox("The table " & sNewTableName & " is exists. To proceed it will be removed and rewrite." & _
               vbCrLf & "Should we continue?", vbQuestion + vbYesNoCancel, "CopyQuerytoNewTable") = vbYes Then
                                DoCmd.DeleteObject acTable, sNewTableName
       Else
               Beep
               Exit Function
       End If
    End If
    
    DoCmd.CopyObject , sNewTableName, acTable, sQueryName
    sRes = sNewTableName
'----------------------
ExitHere:
    CopyQueryToNewTable = sRes '!!!!!!!!!!
    Exit Function
'---------------
ErrHandle:
    ErrPrint2 "CopyQueryToNewTable", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'======================================================================================================================================================
' Close QUERY or TABLE
'======================================================================================================================================================
Public Function CloseOpenTableQuery(sTableQuery As String, Optional ObjectType As AcObjectType = acTable) As Boolean
Dim ao As AccessObject

    On Error GoTo ErrHandle
'-----------------------
Select Case ObjectType
Case acTable:
    Set ao = CurrentData.AllTables(sTableQuery)
Case acQuery:
    Set ao = CurrentData.AllQueries(sTableQuery)
Case Else:
    ErrPrint2 "CloseOpenTableQuery", 0, "Wrong Type. This Function Close only Table/Query", MOD_NAME
    Exit Function
End Select

With ao
    If .IsLoaded Then DoCmd.Close ObjectType, sTableQuery, acSaveYes
End With
'-----------------------
ExitHere:
    CloseOpenTableQuery = True '!!!!!!!!!!!!!!!
    Set ao = Nothing
    Exit Function
'-----------
ErrHandle:
    ErrPrint2 "CloseOpenTableQuery", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Set ao = Nothing
End Function

'======================================================================================================================================================
' Check If Table contains some Field
'======================================================================================================================================================
Public Function IsField(sTableName As String, sFieldName As String) As Boolean
Dim cn As Object, bRes As Boolean
Dim RS As Object

Const AD_SCHEMACOLUMNS As Integer = 4

    On Error GoTo ErrHandle
'------------------------
    Set cn = CurrentProject.Connection

    Set RS = cn.OpenSchema(AD_SCHEMACOLUMNS, _
    Array(Empty, Empty, sTableName, sFieldName))

    If Not RS.EOF Then bRes = True
'------------------------
ExitHere:
    IsField = bRes '!!!!!!!!!!
    Set RS = Nothing: Set cn = Nothing
    Exit Function
'------------
ErrHandle:
    ErrPrint2 "IsField", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Set RS = Nothing: Set cn = Nothing
End Function


'======================================================================================================================================================
'  Save SQL to New SQL
'======================================================================================================================================================
Public Function SaveQuery(sQueryName As String, sSQL As String, Optional bAskToRewrite As Boolean) As Boolean
Dim qdf As QueryDef

Dim bRes As Boolean

    On Error GoTo ErrHandle
'-------------------------
   If sQueryName = "" Then Exit Function
   If sSQL = "" Then Exit Function
        
   If IsQuery(sQueryName) Then
        If bAskToRewrite Then
            If MsgBox("The query " & sQueryName & " exists. Shoul we overwrite it?", vbYesNoCancel + vbQuestion, "SaveQuery") <> vbYes Then Exit Function
        End If
           
        Call DoCmd.DeleteObject(acQuery, sQueryName)
   End If
   
   Set qdf = CurrentDb.CreateQueryDef(sQueryName, sSQL)
   bRes = True
'-------------------------
ExitHere:
    SaveQuery = bRes '!!!!!!!!!!!
    Exit Function
'--------
ErrHandle:
    ErrPrint2 "SaveQuery", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'======================================================================================================================================================
' FIELD LIST FOR RECORDSET, TABLE OR QUERY
' IsSimpleMode= True --> Just Row: Fld1;Fld2;Fld3;... (SEP is Ignored)
' IsSimpleMode= False --> Each FLD is separate row like: FLD_NAME;FLD_TYPE;FLD_SIZE
'======================================================================================================================================================
Public Function GetFieldList(ObjectOrTblName As Variant, Optional DLM As String = ";", Optional SEP As String = vbCrLf, _
                                                                                                     Optional bEasyList As Boolean = True) As String
Dim db As DAO.Database, FLD As DAO.Field, dbObject As Object, sObject As String
Dim sRes As String, FLDList() As TFLD, nFlds As Long, bTable As Boolean

On Error GoTo ErrHandle
'-----------------------
Set db = CurrentDb

If varType(ObjectOrTblName) = vbString Then  ' This is a Table or Query Name or SQL
    sObject = CStr(ObjectOrTblName)
    If IsTable(sObject) Then
           Set dbObject = db.TableDefs(sObject).FIELDS
           bTable = True
    ElseIf IsQuery(sObject) Then
           Set dbObject = db.QueryDefs(sObject)
    Else
           If InStr(1, sObject, "SELECT", vbTextCompare) > 0 Then
               Set dbObject = db.OpenRecordset(sObject)
           Else
               Err.Raise 10007, , "Can't recognize the source " & sObject
           End If
    End If

ElseIf varType(ObjectOrTblName) = vbDataObject Or varType(ObjectOrTblName) = vbObject Then
    Set dbObject = ObjectOrTblName.FIELDS
Else
    Err.Raise 10007, , "Can't recognize the source " & CStr(ObjectOrTblName)
End If
'---------------------
        ReDim FLDList(0): nFlds = -1
        For Each FLD In dbObject
                nFlds = nFlds + 1: ReDim Preserve FLDList(nFlds)
                If bEasyList Then
                    FLDList(nFlds).Name = FLD.Name
                Else
                    FLDList(nFlds).Name = FLD.Name
                    FLDList(nFlds).Type = FLD.Type
                    FLDList(nFlds).SIZE = FLD.SIZE
                    FLDList(nFlds).DefaultValue = CStr(FLD.DefaultValue)
                    FLDList(nFlds).Required = FLD.Required
                        FLDList(nFlds).Description = GetFieldDescription(FLD)
                End If
        Next FLD
                        
                    sRes = FldListToString(FLDList, DLM, SEP, bEasyList)
'-----------------------
ExitHere:
   GetFieldList = sRes '!!!!!!!!!!!!!!!!
   Set FLD = Nothing: Set dbObject = Nothing
   Exit Function
'-----------------
ErrHandle:
   ErrPrint2 "GetFieldList", Err.Number, Err.Description, MOD_NAME
   Err.Clear: Resume ExitHere
End Function
'-------------------------------------------------------------------------------------------------------------------------
' Format String by Recofnition
'-------------------------------------------------------------------------------------------------------------------------
Private Function SQLFormat(sVal As String, Optional nFormat As Integer = 10) As String
Dim sRes As String
    Select Case nFormat
    Case dbText ' 10 Chat,String
        sRes = SQLText(sVal)
    Case dbMemo ' 12 - Long Text
        sRes = SQLText(sVal)
    Case dbDate ' 8 - Date
        sRes = SQLDate(sVal)
    Case dbGUID ' 15
        sRes = SQLGUID(sVal)
    Case dbByte, dbInteger, dbLong, dbDouble, dbDecimal, dbCurrency ' Numbers
        sRes = sVal
    Case Else
        sRes = ""
    End Select
'-------------------------------------
    SQLFormat = sRes '!!!!!!!!!!
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------
' Format FLD according to type
'----------------------------------------------------------------------------------------------------------------------------------------------
Private Function FormatFieldValue(FLDValue As Variant, FLDFormat As Integer) As String
Dim sRes As String

On Error GoTo ErrHandle
'----------------------------------
Select Case FLDFormat
Case 0: ' TEXT
       sRes = SQLText(CStr(FLDValue))
Case 1: ' Long or other Numeric
       sRes = CStr(FLDValue)
Case 2: ' Boolean
       sRes = CStr(CBool(FLDValue))
Case 3: ' Date
       sRes = SQLDate(CStr(FLDValue))
Case 4: ' GUID
       sRes = SQLGUID(CStr(FLDValue))
Case Else
       sRes = CStr(FLDValue)
End Select
'-----------------------------
ExitHere:
    FormatFieldValue = sRes '!!!!!!!!
    Exit Function
'-----------------------------
ErrHandle:
    ErrPrint "FormatFieldValue", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'----------------------------------------------------------------------------------------------------------------------------------------------------
' Copy ATTACHMENT Field (RS2 should be in the edit mode)
'----------------------------------------------------------------------------------------------------------------------------------------------------
Private Function CopyATTACHMENT(ByRef RS1 As DAO.Recordset, ByRef rs2 As DAO.Recordset, AttchFLdName As String) As Boolean
Dim rst As DAO.Recordset

    On Error GoTo ErrHandle
'-----------------------------------
                                
                           Set rst = rs2.FIELDS(AttchFLdName).value
                           With RS1.FIELDS(AttchFLdName).value
                           If .RecordCount > 0 Then
                                If Not .EOF Then
                                   .MoveLast: .MoveFirst
                                   Do While Not .EOF
                                        rst.AddNew
                                            rst.FIELDS("FileData").value = .FIELDS("FileData").value
                                            rst.FIELDS("FileName").value = .FIELDS("FileName").value
                                        rst.Update
                                       
                                       .MoveNext
                                    Loop
                                End If
                           End If
                           End With
'-----------------------------------
ExitHere:
    CopyATTACHMENT = True '!!!!!!!!!!!!!!
    Set rst = Nothing
    Exit Function
'-----------
ErrHandle:
    ErrPrint2 "CopyMVF", Err.Number, Err.Description
    Err.Clear
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------
' Copy MVF Field (RS2 should be in the edit mode)
'----------------------------------------------------------------------------------------------------------------------------------------------------
Private Function CopyMVF(ByRef RS1 As DAO.Recordset, ByRef rs2 As DAO.Recordset, MVFFLdName As String) As Boolean
Dim rst As DAO.Recordset

    On Error GoTo ErrHandle
'-----------------------------------
                           If RS1.FIELDS(MVFFLdName).value.RecordCount > 0 Then
                                
                                Set rst = rs2.FIELDS(MVFFLdName).value
                                rst.AddNew
                                   rst.FIELDS(0).value = RS1.FIELDS(MVFFLdName).value.FIELDS(0).value
                                rst.Update
                           End If
'-----------------------------------
ExitHere:
    CopyMVF = True '!!!!!!!!!!!!!!
    Set rst = Nothing
    Exit Function
'-----------
ErrHandle:
    ErrPrint2 "CopyMVF", Err.Number, Err.Description
    Err.Clear
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Check If Prefix In List
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function IsPrefixInList(sPrefixList As String, sWord As String, Optional DLM As String = ";") As Boolean
Dim Arr() As String, nDim As Integer, I As Integer
Dim bRes As Boolean
    
    On Error Resume Next
'------------------
    If sPrefixList = "" Then Exit Function
    
    Arr = Split(sPrefixList, DLM): nDim = UBound(Arr)
    
    For I = 0 To nDim
         If Arr(I) = "" Then GoTo NextPrefix
         If Left(sWord, Len(Arr(I))) = Arr(I) Then
              bRes = True: Exit For
         End If
NextPrefix:
    Next I
'------------------
ExitHere:
    IsPrefixInList = bRes '!!!!!!!!
    Exit Function
End Function

