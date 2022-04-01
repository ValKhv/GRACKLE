VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSQLParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************************************************************************
'*********************************************************************************************************************************************************
' DESCRIPTION: This class is simple SQL Parser to provide its general structure
' DEPENDENCY: #_STRING.TextBetweenTwoWords (string parser)
' VERSION    : v 0.2 13.01.2019
' COPYRIGHT  : V.Khvatov@gmail.com DigitalXpert
'*********************************************************************************************************************************************************
'*********************************************************************************************************************************************************
Option Compare Database
Option Explicit

Private Type pFLD
    FldName As String
    FldExpression As String
End Type

Private Type pSelectSQL
    SQLBody As String
    tblName() As String
    FLDS() As String
    IsComplex As Boolean
    sWhere As String
    sOrderBy As String
    sGROUPBY As String
    sHaving As String
End Type

Private m_Body As String
Private m_WHERE As String
Private m_ORDERBY As String
Private m_GROUPBY As String
Private m_HAVING As String
'****************************************************************************************************
' SELECT CLAUSE (see https://support.office.com/en-us/article/access-sql-basic-concepts-vocabulary-and-syntax-444d0303-cde1-424e-9a74-e8dc3e460671)
'-----------------------------------------------------------------------------------------------------
'  KEY WORD |   DESCRIPTION                                                             |   REQUIRED
'-----------------------------------------------------------------------------------------------------
'   SELECT  |   Lists the fields that contain data of interest.                         |     Yes
'   FROM    |   Lists the tables that contain the fields listed in the SELECT clause.   |     Yes
'   WHERE   |   Specifies field criteria that must be met by each record to be included |      No
'           |     in the results.                                                       |
'   ORDER BY|   Specifies how to sort the results.                                      |      No
'   Group BY|   In a SQL statement that contains aggregate functions, lists fields that |      No
'           |   are not summarized in the SELECT clause. Only if there are such fields  |
'   HAVING  |   In a SQL statement that contains aggregate functions, specifies         |      No
'           |   conditions that apply to fields that are summarized in the              |
'           |   SELECT statement.                                                       |
'------------------------------------------------------------------------------------------------------

Private m_SELECT As pSelectSQL     ' The main object for Select SQL

Public Property Get BODY() As String
     BODY = m_Body
End Property
Public Property Let BODY(sBody As String)
     m_Body = sBody
End Property

Public Property Get Where() As String
     Where = m_WHERE
End Property
Public Property Let Where(sWhere As String)
    m_WHERE = sWhere
End Property

Public Property Get OrderBy() As String
    OrderBy = m_ORDERBY
End Property
Public Property Let OrderBy(sOrderBy As String)
    m_ORDERBY = sOrderBy
End Property

Public Property Get GroupBy() As String
    GroupBy = m_GROUPBY
End Property
Public Property Let GroupBy(sGROUPBY As String)
    m_GROUPBY = sGROUPBY
End Property

Public Property Get Having() As String
    Having = m_HAVING
End Property
Public Property Let Having(sHaving As String)
    m_HAVING = sHaving
End Property
'=========================================================================================================================================================
' Get Fields as string (included expressions) in format: FIELDNAME1;FIELDNAME2=EXPRESSION;FIELDNAME3
'=========================================================================================================================================================
Public Property Get FIELDS() As String
    
End Property


'==========================================================================================================================================================
' The main function to build and parse
'==========================================================================================================================================================
Public Function SELECTParse(sSQL As String) As Boolean
Dim bRes As Boolean, sFileds As String
Dim iWhere As Integer, iGroupBy As Integer, iOrderBy As Integer, iHaving As Integer
On Error GoTo ErrHandle
'------------------------------------
    m_SELECT.SQLBody = Trim(sSQL)
    If m_SELECT.SQLBody = "" Then Exit Function
    If InStr(1, m_SELECT.SQLBody, "SELECT", vbTextCompare) = 0 Then Err.Raise "The sql sentense is not Select-type query"
    If InStr(1, m_SELECT.SQLBody, "FROM", vbTextCompare) = 0 Then Err.Raise "Can't found the FROM Clause"

iWhere = InStr(1, m_SELECT.SQLBody, "WHERE", vbTextCompare)
iGroupBy = InStr(1, m_SELECT.SQLBody, "GROUP BY", vbTextCompare)
iOrderBy = InStr(1, m_SELECT.SQLBody, "ORDER BY", vbTextCompare)
iHaving = InStr(1, m_SELECT.SQLBody, "HAVING", vbTextCompare)

'SELECT TEST.ID, TEST.MVF, [_LBL].Field1
'FROM TEST LEFT JOIN _LBL ON TEST.MVF.Value = [_LBL].ID;


sFileds = TextBetweenTwoWords(m_SELECT.SQLBody, "SELECT", "FROM", True, vbTextCompare)

    
'------------------------------------
ExitHere:
    SELECTParse = bRes '!!!!!!!!!!!!
    Exit Function
'-----------
ErrHandle:
    ErrPrint "SELECTParse", Err.Number, Err.Description
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
                                                                                                  Optional sModName As String = "cSQLParser") As String
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

