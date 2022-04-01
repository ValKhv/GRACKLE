VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cJet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************************************
'*********************************************************************************************************************
'*********************************************************************************************************************
' DESCRIPTION: Класс представляет собой обертку Jet (Access) DB для автоматизации типовых операций и управления базой
' VERSION    : v 0.4 07.08.2014
' COPYRIGHT  : V.Khvatov DigitalXpert
'*********************************************************************************************************************
'*********************************************************************************************************************
'*********************************************************************************************************************

Option Compare Database
Option Explicit

' *******************************************************************
' Локальные объектные переменные
Private m_dbsCurrent As DAO.Database             ' Текущая база данных
Private m_fExclusive As Boolean                  ' База данных открыта эксклюзивно
Private m_fReadonly As Boolean                   ' База открыта только для чтения
Private m_varDatabasePassword As Variant         ' Пароль на Базу
Private FFIELDS() As FFIELDS                     ' Массив полей (для описания базы в случае необходимости)

' Специфические переменные класса
Private mfClassOpened As Boolean                 ' если класс открыл базу данных - в true
Private mfExclusive As Boolean
Private mfReadOnly As Boolean
Private mvarDatabasePassword As Variant

'*********************************************************************
' Тип объектов в базе данных
Public Enum EnumCJetObjectType
  cjoTypeTable = 0
  cjoTypeQuery = 1
  cjoTypeRelation = 6
  cjoTypeAccessForm = 2
  cjoTypeAccessReport = 3
  cjoTypeAccessMacro = 4
  cjoTypeAccessModule = 5
  cjoTypeALL = 6
End Enum

' *******************************************************************
' Перечисление типа ссылок Jet
Public Enum EnumCJetLinkType
  cjlTypeJet = 1
  cjlTypedBASE3 = 100
  cjlTypedBASE4 = 110
  cjlTypedBASE5 = 120
  cjlTypePdx3 = 200
  cjlTypePdx4 = 210
  cjlTypePdx5 = 220
  cjlTypeFoxPro2 = 300
  cjlTypeFoxPro25 = 310
  cjlTypeFoxPro26 = 320
  cjlTypeFoxPro30 = 330
  cjlTypeFoxProDBC = 340
  cjlTypeExcel3 = 400
  cjlTypeExcel4 = 410
  cjlTypeExcel5 = 420
  cjlTypeExcel97 = 430
  cjlTypeHTMLImport = 500
  cjlTypeHTMLExport = 510
  cjlTypeText = 600
  cjlTypeODBC = 700
  cjlTypeExchange = 800
End Enum
'***********************************************************
' Константы для определения ссылок Jet
Private Const mcstrLinkJet = ";"
Private Const mcstrLinkdBase3 = "dBASE III;"
Private Const mcstrLinkdBase4 = "dBASE IV;"
Private Const mcstrLinkdBase5 = "dBASE 5.0;"
Private Const mcstrLinkParadox3 = "Paradox 3.x;"
Private Const mcstrLinkParadox4 = "Paradox 4.x;"
Private Const mcstrLinkParadox5 = "Paradox 5.x;"
Private Const mcstrLinkFoxPro20 = "FoxPro 2.0;"
Private Const mcstrLinkFoxPro25 = "FoxPro 2.5;"
Private Const mcstrLinkFoxPro26 = "FoxPro 2.6;"
Private Const mcstrLinkFoxPro30 = "FoxPro 3.0;"
Private Const mcstrLinkFoxProDBC = "FoxPro DBC;"
Private Const mcstrLinkExcel3 = "Excel 3.0;"
Private Const mcstrLinkExcel4 = "Excel 4.0;"
Private Const mcstrLinkExcel5 = "Excel 5.0;"
Private Const mcstrLinkExcel97 = "Excel 97;"
Private Const mcstrHTMLImport = "HTML Import;"
Private Const mcstrHTMLExport = "HTML Export;"
Private Const mcstrLinkText = "Text;"
Private Const mcstrLinkODBC = "ODBC;"
Private Const mcstrLinkExchange = "Exchange 4.0;"

'**************************************************************************
Private Type FFIELDS       ' ТИП ПОЛЕ, простейший внутренний описатель
    Fname As String        ' Наименование поля
    fType As Integer       ' Тип поля
    fSize As Integer       ' Размер поля (только для Text)
    FisPK As Boolean       ' Флаг того, что данное поле есть ключ (только одно поле в таблице)
    FisReleased As Boolean ' Флаг того, что поле отображено на физическую таблицу
End Type
'*************************************************************************
Public Enum FFieldType     ' ПЕРЕЧИСЛЕНИЕ - ТИП ПОЛЯ
    fBoolean = dbBoolean
    fByte = dbByte
    fInteger = dbInteger
    fLong = dbLong
    fCurrency = dbCurrency
    fSingle = dbSingle
    fDouble = dbDouble
    fDate = dbDate
    fText = dbText
    fLongBinary = dbLongBinary
    fMemo = dbMemo
End Enum
'*************************************************************************************************************************
'*************************************************************************************************************************
'==========================================================================================================================
' Получение версии Microsoft Access для текущей базы данных. Свойство может не существовать, наприемр, если база была
' создана автоматически и никогда не открывалась Access - в таком случае управление передается обработчику ошибок
'==========================================================================================================================
Public Property Get AccessVersion() As String
  On Error GoTo PROC_ERR
  
  AccessVersion = m_dbsCurrent.Properties("AccessVersion")
'---------------------------------------
PROC_EXIT:
  Exit Property
'---------------------------------------
PROC_ERR:
  ' Свойство не существует, возможно база никогда не открывалась в Access
  AccessVersion = ""
  Exit Property
End Property
'==========================================================================================================================
' Возвращает указатель на текущую открытую базу данных
'==========================================================================================================================
Public Property Get Database() As DAO.Database
  On Error GoTo PROC_ERR
  
  Set Database = m_dbsCurrent
'----------------------------------------------------
PROC_EXIT:
  Exit Property
'----------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "Database"
  Resume PROC_EXIT
End Property
'==========================================================================================================================
' Установка указателя на УЖЕ ОТКРЫТУЮ базу данных
'==========================================================================================================================
Public Property Set Database(dbsValue As DAO.Database)
  On Error GoTo PROC_ERR
  
  Set m_dbsCurrent = dbsValue
'--------------------------------------------
PROC_EXIT:
  Exit Property
'--------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "Database"
  Resume PROC_EXIT
End Property
'==========================================================================================================================
' Получение имени базы (только чтение)
'==========================================================================================================================
Public Property Get DatabaseName() As String
  On Error GoTo PROC_ERR
     DatabaseName = Me.Database.Name
'----------------------------------------------------
PROC_EXIT:
  Exit Property
'----------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "Database"
  Resume PROC_EXIT
End Property

'==========================================================================================================================
' Получение пароля. который был использован при открытии базы данных. Работает только при открытии базы данных
' с помощью метода OpenDB. если база данных открыта вне класса, то свойство возвращает пустую строку
'==========================================================================================================================
Public Property Get DatabasePassword() As String
  DatabasePassword = m_varDatabasePassword
End Property
'==========================================================================================================================
' Возвращает True, если база данных была создана в Access или хотя бы раз там открывалась
'==========================================================================================================================
Public Property Get IsAccessDatabase() As Boolean
  Dim conTmp As DAO.Container
  
  On Error GoTo PROC_ERR
'------------------------------------------------------------------------------------------------------------
' Если база данных была создана или хоть раз открывалась в Access , то должна содержать объект Container для форм.
' Пробуем получить доступ к контейнеру, если получаем ошибку - значит НЕ БАЗА ACCESS.
  Set conTmp = m_dbsCurrent.Containers!Forms
  
  IsAccessDatabase = True
'--------------------------------------------
PROC_EXIT:
  Exit Property
'--------------------------------------------
PROC_ERR:
  ' Контейнер не существует, база никогда не создавалась и не открывалась в Access
  IsAccessDatabase = False
  Exit Property
   
End Property
'==========================================================================================================================
' Возвращает описание последней случившейся ошибки
'==========================================================================================================================
Public Property Get LastErrorDescription() As String
  On Error GoTo PROC_ERR
  
  If DAO.DBEngine.Errors.Count > 0 Then
    LastErrorDescription = DAO.DBEngine.Errors(0).Description
  End If
'---------------------------------------------------
PROC_EXIT:
  Exit Property
'---------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "LastErrorDescription"
  Resume PROC_EXIT
  
End Property
'==========================================================================================================================
' Возвращает номер последней случившейся ошибки
'==========================================================================================================================
Public Property Get LastErrorNumber() As Long
  On Error GoTo PROC_ERR
  
  If DAO.DBEngine.Errors.Count > 0 Then
    LastErrorNumber = DAO.DBEngine.Errors(0).Number
  End If
'-------------------------------------------------
PROC_EXIT:
  Exit Property
'------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "LastErrorNumber"
  Resume PROC_EXIT
  
End Property
'==========================================================================================================================
' Возвращает имя текущей базы данных (без пути)
'==========================================================================================================================
Public Property Get Name() As Variant
Dim MyStr As New cString
  On Error GoTo PROC_ERR
  
  Name = MyStr.GetFileNamePart(m_dbsCurrent.Name)
'-----------------------------------------------------
PROC_EXIT:
  Exit Property
'-----------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "Name"
  Resume PROC_EXIT
  
End Property
'==========================================================================================================================
' Получает свойство открытия: если база была открыта классом эксклюзивно - возвращается true
'==========================================================================================================================
Public Property Get OpenExclusive() As Boolean
  OpenExclusive = m_fExclusive
End Property
'==========================================================================================================================
' Получает свойство открытия: если база была открыта классом только для чтения - возвращается true
'==========================================================================================================================
Public Property Get OpenReadOnly() As Boolean
  OpenReadOnly = m_fReadonly
End Property
'==========================================================================================================================
' Current DB Path
'==========================================================================================================================
Public Property Get DBPath() As String
  
  On Error Resume Next
'--------------------------
  DBPath = CurrentProject.Path   '!!!!!!!!!!!!!!
End Property
'==========================================================================================================================
' Получает версию текущей базы данных (версию Jet Engine)
'==========================================================================================================================
Public Property Get Version() As Variant
  
  On Error Resume Next
'-------------------------------------
  Version = CurrentDb.Properties("Version")
End Property
'==========================================================================================================================
' Закрывает текущую открытую базу данных
'==========================================================================================================================
Public Sub CloseDB()
  On Error GoTo PROC_ERR
' -----------------------------------------------------------------------------------
' База данных закрывается только в том случае, если была открыта с помощью класса
  If mfClassOpened Then
    m_dbsCurrent.Close          ' закрываем базу и очищаем память
    Set m_dbsCurrent = Nothing
    ' Сбрасываем флаг
    mfClassOpened = False
  End If
'---------------------------------------------------
PROC_EXIT:
  Exit Sub
'---------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "CloseDB"
  Resume PROC_EXIT

End Sub


'============================================================================================================
' Проверка существования таблицы в текущей базе данных
'============================================================================================================
Public Function IsTableExists(TableName As String) As Boolean
Dim TD As TableDef                       ' табличное пространство

On Error Resume Next
'----------------------------------------------------
' Проверяем табличное пространство на предмет наличия таблицы. В случае отсутствия генерится ошибка
Set TD = m_dbsCurrent.TableDefs(TableName)
    IsTableExists = Err.Number = 0     '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Exit Function
'--------------------------------------------------------
' Обработчик ошибок
IsTableExists_ERR:
    Err.Raise Err.Number
    Exit Function
End Function
'============================================================================================================
' Проверка существования запроса в текущей базе данных
'============================================================================================================
Public Function IsQueryExists(QueryName As String) As Boolean
Dim QD As QueryDef                        ' запрос

On Error Resume Next
'----------------------------------------------------
' Проверяем пространство на предмет наличия запроса. В случае отсутствия генерится ошибка
Set QD = m_dbsCurrent.QueryDefs(QueryName)
    IsQueryExists = Err.Number = 0     '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Exit Function
'--------------------------------------------------------
' Обработчик ошибок
IsQueryExists_ERR:
    Err.Raise Err.Number
    Exit Function
End Function

'===========================================================================================================================
' Удаление записи из таблицы на основании критерия Criteria (в частном случае это KEY1=VALUE1;KEY2=VALUE2;...)
' Если bConfirm = True - запрос на подтверждение. Если Criteria ="", удаляются все записи
'==========================================================================================================================
Public Function DeleteRecordFromTable(ByVal TableName As String, Optional Criteria As String = "", _
                                                                          Optional bConfirm As Boolean = False) As Boolean
Dim strSQL As String                       ' Исполнимая строка запроса

On Error GoTo DeleteRecordFromTable_ERR

'----------------------------------------------------
' Проверяем необходимость проверки существования записи
 If Criteria = "" Then             ' Удаляем все записи в таблице
    strSQL = "DELETE * " & "FROM " & SHT(TableName) & ";"
 Else                              ' Удаляем записи, соотв. критерию
    strSQL = "DELETE * " & "FROM " & SHT(TableName) & "WHERE (" & NormKV(Criteria) & ");"
 End If
  
  DoCmd.SetWarnings False          ' Снимаем предупреждение
'----------------------------------------------------
' Выполняем запрос
        m_dbsCurrent.Execute strSQL
'-----------------------------------------------------
' Выходим из функции
        DeleteRecordFromTable = True   '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        Exit Function
'-----------------------------------------------------
' Обрабатываем ошибки
DeleteRecordFromTable_ERR:
    If Err <> 0 Then
        MsgBox Err.Description, vbExclamation, "DeleteRecordFromTable Error " & Err.Number
    End If
End Function
'=========================================================================================================================
' Очистка таблицы. Возвращает True в случае удачи
'=========================================================================================================================
Public Function EmptyTable(ByVal TableName As String) As Boolean
Dim bRes As Boolean
Dim KeyFieldName As String                 ' Ключевой столбец (автономер)
Dim strSQL As String                       '
On Error GoTo ClearTable_ERR

'-------------------------------------------------------------------
' Проверяем наличие таблицы
If Not IsTableExists(TableName) Then GoTo ClearTable_ERR
'---------------------------------------------------------------------
' Удаляем все записи из таблицы
  bRes = DeleteRecordFromTable(TableName)
'----------------------------------------------------
' Выход из функции
  EmptyTable = True  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  Exit Function
'-----------------------------------------------------
' Обрабатываем ошибки
ClearTable_ERR:
    If Err <> 0 Then
        MsgBox Err.Description, vbExclamation, "ClearTable Error " & Err.Number
    End If
End Function
'==========================================================================================================================
'  Функция создает налету объект Query с помощью DAO
'==========================================================================================================================
Public Function CreateQuery(QueryName As String, sSQL As String) As Boolean
Dim qdf As DAO.QueryDef                    ' Рабочий запрос

On Error GoTo CreateQuery_ERR
'----------------------------------------------------
' Создаем новый запрос
    Set qdf = m_dbsCurrent.CreateQueryDef(QueryName)
    'Set the SQL property to a string representing a SQL statement.
    qdf.SQL = sSQL
'----------------------------------------------------
' Выход из функции
  CreateQuery = True
  Set qdf = Nothing
  Exit Function
'-----------------------------------------------------
' Обрабатываем ошибки
CreateQuery_ERR:
    If Err <> 0 Then
        MsgBox Err.Description, vbExclamation, "CreateQuery Error " & Err.Number
    End If
End Function
'==========================================================================================================================
'  Функция получает актуальный размер поля для заданной ячейки(таблицы, поля и заданной строки)
'==========================================================================================================================
Public Function GetCellSize(TableName As String, FieldName As String, sCriteria As String) As Long
Dim RS As DAO.Recordset                         ' Открываем рекордсет
Dim sSQL As String                              ' Строка запроса
Dim lRes As Long                                ' Возвращаемый результат

On Error GoTo GetCellSize_ERR
'---------------------------------------------
lRes = -1
sSQL = "Select " & SHT(TableName) & "." & SHT(FieldName) & " FROM " & SHT(TableName) & " WHERE (" & NormKV(sCriteria) & ");"

Set RS = m_dbsCurrent.OpenRecordset(sSQL)
With RS
   If Not .EOF Then
     .MoveLast:       .MoveFirst
     lRes = RS.FIELDS(FieldName).FieldSize
   End If
End With
'--------------------------------------------
GetCellSize_EXIT:
    GetCellSize = lRes   '!!!!!!!!!!!!!!!!!!!!!
    Set RS = Nothing
    Exit Function
'---------------------
GetCellSize_ERR:
    MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "GetCellSize ERROR"
    Err.Clear
    Resume GetCellSize_EXIT
End Function
'==========================================================================================================================
'  Функция получает свойство DAO объекта
' Table (Native: non-linked) = 1; Table(ODBC - linked) = 4; Table (MS Jet/Access linked) = 6
' Query = 5; Form = -32768; Report = -32764; Macro = -32766; Module = -32761 ;
'==========================================================================================================================
Public Function GetObjProp(objName As String, ObjType As Integer, PropName As String) As String
Dim sRes As String
    
    On Error GoTo ErrHandle
    
    Select Case ObjType
    Case 1, 4, 6:
         sRes = CStr(CurrentDb.TableDefs(objName).Properties(PropName).value)
    Case 5:
         sRes = CStr(CurrentDb.QueryDefs(objName).Properties(PropName).value)
    Case -32768:
         sRes = CStr(CurrentDb.Containers!Forms.Documents(objName).Properties(PropName).value)
    Case -32764:
         sRes = CStr(CurrentDb.Containers!Reports.Documents(objName).Properties(PropName).value)
    Case -32766:
         sRes = CStr(CurrentDb.Containers!Scripts.Documents(objName).Properties(PropName).value)
    Case -32761:
         sRes = CStr(CurrentDb.Containers!Modules.Documents(objName).Properties(PropName).value)
    Case Else
         sRes = ""
    End Select
'------------------------------------------------------------------
ExitHere:
    GetObjProp = sRes '!!!!!!!!!!!!
    Exit Function
'-----------------------
ErrHandle:
    Debug.Print String(20, "#") & vbCrLf & "ERR#" & Err.Number & vbCrLf & Err.Description
    Err.Clear
    sRes = ""
    Resume ExitHere
End Function
'==========================================================================================================================
'  Функция получае с помощью DAO наименование стобца - счетчика, чаще всего являющеееся также PK
'==========================================================================================================================
Public Function GetAutoNumField(TableName As String) As String
Dim oDB As DAO.Database                    ' Рабочая база данных
Dim tdf As DAO.TableDef                    ' Рабочее табличное пространство
Dim FLD As DAO.Field                       ' Рабочее поле

On Error GoTo GetAutoNumField_ERR
'-------------------------------------------------------------------
' Инициализируемся
GetAutoNumField = ""
'-------------------------------------------------------------------
' Проверяем наличие таблицы
If Not IsTableExists(TableName) Then GoTo GetAutoNumField_ERR
'---------------------------------------------------------------------
' Определяем табличное пространство и проходим по столбцам
    Set tdf = m_dbsCurrent.TableDefs(TableName)
    
    For Each FLD In tdf.FIELDS
        If (FLD.Attributes And dbAutoIncrField) <> 0 Then
            GetAutoNumField = FLD.Name
            Exit For
        End If
    Next
'----------------------------------------------------
' Выход из функции
    Set FLD = Nothing
    Set tdf = Nothing
  Exit Function
'-----------------------------------------------------
' Обрабатываем ошибки
GetAutoNumField_ERR:
    If Err <> 0 Then
        MsgBox Err.Description, vbExclamation, "GetAutoNumField Error " & Err.Number
    End If
End Function
'===========================================================================================================================
' Получаем массив данных из запроса:
' sSQL - строка запроса. Если sSQL="*", то считываем из TableName всю информацию в массив. Иначе -
'                                                                                       игнорируем Tablename
' TableName - таблица, работает только, если sSQL = "*"
'===========================================================================================================================
Public Function GetArrayFromQuery(sSQL As String, Optional TableName As String = "", _
                                                                              Optional bWithColumns As Boolean) As Variant
Dim RS As DAO.Recordset                            ' Рабочий Recordset
Dim myArray() As String                            ' Массив переменных

On Error GoTo GetArrayFromQuery_ERR

'--------------------------------------------------------------------------------
' Инициализируем функцию
    GetArrayFromQuery = Empty
    If sSQL = "" Then GoTo GetArrayFromQuery_ERR
'-------------------------------------------------------------------------------
If sSQL = "*" Then                         ' Считываем таблицу целиком
         ' Проверяем наличие таблицы
         If Not IsTableExists(TableName) Then GoTo GetArrayFromQuery_ERR  '  Если таблицы нет, выходим
         Set RS = m_dbsCurrent.OpenRecordset("Select * From " & TableName & ";")
Else                                       ' Выполняем произвольный запрос и игнорируем таблицу
         Set RS = m_dbsCurrent.OpenRecordset(sSQL)
End If
'------------------------------------------------------------
' Получаем массив для заданного RS
   If Not IsNull(RS) Then
       GetArrayFromQuery = GetArrrayFromRS(RS, bWithColumns)
   End If
'----------------------------------------------------------------
' Выходим
GetArrayFromQuery_Exit:
   If Not IsNull(RS) Then RS.Close
   Set RS = Nothing
   Exit Function
'----------------------------------------------------------------
' Обработчик ошибок
GetArrayFromQuery_ERR:
   If Err > 0 Then MsgBox "GetArrayFromQuery_ERR: " & Err.Number & vbCrLf & Err.Description
   GoTo GetArrayFromQuery_Exit
End Function
'==========================================================================================================================
' Функция расширяет возможности стандартной DCount с точки зрения поиска не только по текущей базе,
' но по любой базе, учитывает порядок возвращаемого множества значений, число неповторяющихся записей
' Возвращает массив найденных значений (или Null, если ничего не найдено). Работает с многозначными полями,
' выполняется быстрее DCount
' Примеры использования:
'      1) Number of customers who have a region: ECount("Region", "Customers")
'      2) Number of customers who have no region: ECount("*", "Customers", "Region Is Null")
'      3) Number of distinct regions: ECount("Region", "Customers", ,True)
'========================================================================================================================
Public Function ECount(Expr As String, TableName As String, Optional Criteria As String, _
                                                                            Optional bCountDistinct As Boolean) As Long
    Dim RS As DAO.Recordset              ' Рабочий набор значений
    Dim strSQL As String                 ' Запрос SQL

On Error GoTo ECount_ERR
'-----------------------------------------------------------------
' Инициализируем функцию
    ECount = 0
'----------------------------------------------------------------
' Проверяем наличие таблицы
If Not IsTableExists(TableName) Then GoTo ECount_ERR  '  Если таблицы нет, выходим
'---------------------------------------------------------------------
' Начинаем подсчет
    If bCountDistinct Then              ' Считаем неповторяющиеся значения
        '---------------------------------------------------------------------------------
        If Expr <> "*" Then             'Cannot count distinct with the wildcard.
            strSQL = "SELECT " & Expr & " FROM " & TableName & " WHERE (" & Expr & " Is Not Null)"
            If Criteria <> vbNullString Then
                strSQL = strSQL & " AND (" & NormKV(Criteria) & ")"
            End If
            strSQL = strSQL & " GROUP BY " & Expr & ";"
            Set RS = m_dbsCurrent.OpenRecordset(strSQL)
            ' rs.MoveLast: rs.MoveFirst   ' Обновляем набор
            If RS.RecordCount > 0& Then
                RS.MoveLast
            End If
            ECount = RS.RecordCount     'Return the number of distinct records.
            RS.Close
        End If
    Else                                ' Обычный подсчет
        strSQL = "SELECT Count(" & Expr & ") AS TheCount FROM " & TableName
        If Criteria <> vbNullString Then
            strSQL = strSQL & " WHERE " & NormKV(Criteria)
        End If
        Set RS = m_dbsCurrent.OpenRecordset(strSQL)
        If RS.RecordCount > 0& Then
            ECount = RS!TheCount        'Return the count.
        End If
        RS.Close
    End If
'----------------------------------------------------------------
ECount_EXIT:
    Set RS = Nothing
    Exit Function
'---------------------------------------------------------------
ECount_ERR:
    MsgBox Err.Description, vbExclamation, "ECount Error " & Err.Number
    Resume ECount_EXIT
End Function
'===========================================================================================================================
' Функция НАХОДИТ СУММУ ПО УКАЗАННОМУ ПОЛЮ
'===========================================================================================================================
Public Function DSumX(Expr As String, TableName As String, Criteria As String) As Variant
Dim RS As DAO.Recordset
Dim sSQL As String
Dim vRes As Variant
On Error GoTo DSumX_ERR
'---------------------------------
vRes = 0
sSQL = "SELECT " & SHT(Expr) & " FROM " & SHT(TableName) & " WHERE " & NormKV(Criteria) & ";"
'-----------------------------------------
Set RS = m_dbsCurrent.OpenRecordset(sSQL)
    If RS.EOF Then GoTo DSumX_EXIT
    RS.MoveLast: RS.MoveFirst
    Do While Not RS.EOF
       vRes = vRes + RS.FIELDS(0).value
       RS.MoveNext
    Loop
'--------------------------------------------------------------------------
DSumX_EXIT:
     DSumX = vRes   '!!!!!!!!!!!!!!!!!
     Set RS = Nothing
     Exit Function
'--------------------------------------------------------------------------
DSumX_ERR:
    MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "DSumX ERROR"
    Set RS = Nothing
End Function

'===========================================================================================================================
' Функция возвращает значение для запроса, работает с Replication Code
'===========================================================================================================================
Public Function DLookUPX(Expr As String, TableName As String, Criteria As String) As Variant
Dim RS As DAO.Recordset
Dim sSQL As String
Dim vRes As Variant
On Error GoTo DLookUPX_ERR

sSQL = "SELECT " & SHT(Expr) & " FROM " & SHT(TableName) & " WHERE " & NormKV(Criteria) & ";"
'Debug.Print "sSQL=" & sSQL
Set RS = m_dbsCurrent.OpenRecordset(sSQL)
    If RS.EOF Then GoTo DLookUPX_EXIT
    RS.MoveLast: RS.MoveFirst
    vRes = RS.FIELDS(0).value
'--------------------------------------------------------------------------
DLookUPX_EXIT:
     DLookUPX = vRes   '!!!!!!!!!!!!!!!!!
     Set RS = Nothing
     Exit Function
'--------------------------------------------------------------------------
DLookUPX_ERR:
    MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description & vbCrLf & "SQL: " & sSQL, vbCritical, "DLookUPX ERROR"
    Debug.Print "----------------------------------------------------------------------------------------------"
    Debug.Print "Error SQL: " & sSQL
    Debug.Print "----------------------------------------------------------------------------------------------"
    Set RS = Nothing
End Function

'===========================================================================================================================
' Функция расширяет возможности стандартной DLookup с точки зрения поиска не только по текущей базе,
' но по любой базе, учитывает порядок возвращаемого множества значений
' Возвращает массив найденных значений (или Null, если ничего не найдено)
' bWithColumns = True - в нулевую строку записываются названия столбцов
' bWithColumns = False - обычный массив
' Работает с многозначными полями, выполняется быстрее DLookup
' Примеры использования:
'      1) ELookup("[Surname] & [FirstName]", "tblClient", , "ClientID DESC")
'      2) ELookup("ClientID", "tblClient", "Surname Is Not Null" , "Surname")
'===========================================================================================================================
Public Function ELookup(Expr As String, TableName As String, Optional Criteria As String, Optional OrderClause As Variant, _
                                                                         Optional bWithColumns As Boolean = True) As Variant
    Dim strSQL As String            ' SQL выражение
    Dim sResNull() As String        ' Нулевой массив
On Error GoTo ELookup_ERR
'--------------------------------------------------------------------------------
' Инициализируем функцию
    ELookup = Empty: ReDim sResNull(0, 0)
'-------------------------------------------------------------------------------
' Проверяем наличие таблицы
If Not IsTableExists(TableName) Then GoTo ELookup_ERR  '  Если таблицы нет, выходим
'----------------------------------------------------
' Строим строку запроса на основании входных данных
    strSQL = "SELECT " & SHT(Expr) & " FROM " & SHT(TableName)
    If Not IsMissing(Criteria) Then               ' Проверяем критерии
        strSQL = strSQL & " WHERE " & NormKV(Criteria)
    End If
    If Not IsMissing(OrderClause) Then
        strSQL = strSQL & " ORDER BY " & OrderClause
    End If
    strSQL = strSQL & ";"
'----------------------------------------------------
' Для заданного запроса получаем массив
   ELookup = GetArrayFromQuery(strSQL, , bWithColumns)
'----------------------------------------------------------------------------------
' Завершаем работу, возвращаем значения
ELookup_EXIT:
   If IsEmpty(ELookup) Then
      ELookup = sResNull             ' Создаем пустой масси
   End If
   Exit Function
'----------------------------------------------------------------------------------
' Обработчик значений
ELookup_ERR:
    If Err > 0 Then MsgBox Err.Description, vbExclamation, "ELookup Error " & Err.Number
    Resume ELookup_EXIT
End Function
'===========================================================================================================================
' Функция добавляет к существующей таблице поле (не PK)
'===========================================================================================================================
Public Function AddFieldToTable(FieldName As String, FieldType As FFieldType, FieldSize As Integer, _
                                                                                      ByVal TableName As String) As Boolean
On Error GoTo AddFieldToTable_ERR
Dim TD As TableDef                         ' Рабочая таблица
Dim f As Field
Dim bRes As Boolean

On Error GoTo AddFieldToTable_ERR
'-------------------------------------------------------------------
' Проверяем наличие таблицы
If Not IsTableExists(TableName) Then GoTo AddFieldToTable_ERR
'----------------------------------------------------
' Создаем объект табличного пространства
Set TD = m_dbsCurrent.TableDefs(TableName)
'----------------------------------------------------
' Создаем новое поле
Set f = TD.CreateField(FieldName, FieldType, FieldSize)
'----------------------------------------------------
' Присоединяем поле к таблице
TD.FIELDS.Append f

'-----------------------------------------------------
' Выходим
    Set f = Nothing
    Set TD = Nothing
    AddFieldToTable = True  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Exit Function
'-----------------------------------------------------
' Обрабатываем ошибки
AddFieldToTable_ERR:
    Set f = Nothing
    Set TD = Nothing
End Function
'=========================================================================================================================
' Функция  добавляет значение в таблице / поле такое-то (sKEYVALUE), на основании заданного id (Criteria)
'=========================================================================================================================
Public Function SetValue(ByVal TableName As String, sKEYVALUE As String, Criteria As String) As Variant
   SetValue = AddRecordToTable(TableName, sKEYVALUE, Criteria, True)
End Function
'=========================================================================================================================
' Функция добавляет запись к существующей таблице:
' PARAMETERS:
'     TableName = таблица, к которой добавляется запись
'     sKEYVALUE = строка вида KEY1=VALUE1;KEY2=VALUE2;... (определяет добавляемые значения)
'     Criteria = при наличии этой строчки проверяем записи на предмет обновления в соответствии с критерием
'     bUpdate =True - обновлять найденные данные; False - не обновлять (если bExist=true  - отказ)
' RETURNS: true, в случае успеха
' Если запись запись существует (в ответ на Criteria), то в зависимости от флага bUpdate она обновляется или
' осуществляется отказ - выход из функции
' Запись идентифицируется с помощью критерия (sCriteria), вид которой KEY=VALUE (например, ID = 100) или другой
' SQL совместимый
' Значения размещаются в строчке sVALUE вида KEY1=VALUE1;KEY2=VALUE2;....
'========================================================================================================================
Public Function AddRecordToTable(ByVal TableName As String, sKEYVALUE As String, _
                Optional Criteria As String = "", Optional bUpdate As Boolean = True, Optional ObjType As Integer = -1) As Variant
On Error GoTo AddRecordToTable_ERR
Dim strSQL As String                       ' Исполнимая строка запроса
Dim IDD As Variant                         ' Добавленный IDD
Dim TimeStamp As String                    ' Штамп даты времени
Dim iL As Long, iR As Long                 ' Для выделения позиции
Dim sHASH As String                        ' Переменная для HASH
TimeStamp = ""
'-------------------------------------------------------------------
' Проверяем наличие таблицы
If Not IsTableExists(TableName) Then GoTo AddRecordToTable_ERR
'--------------------------------------------------------------------
If IsFieldExists("TimeStampCode", TableName) Then    ' Присутсвует метка возрастающего времени
    TimeStamp = GetTimestamp()
    sKEYVALUE = sKEYVALUE & ";TimeStampCode='" & TimeStamp & "'"
End If
'---------------------------------------------------------------------
' Проверяем необходимость проверки существования записи
 If Criteria <> "" Then            ' Сначала проверяем, есть ли запись, потом решаем обновлять или добавлять
    If ECount("*", SHT(TableName), Criteria) > 0 Then      ' Нашли какое-то кол-во записей для обновления
       If Not bUpdate Then GoTo AddRecordToTable_ERR  ' Нельзя обновлять, выходим из функции
       '---------------------------------------------------------------------------------------------
       ' Обновляем записи в таблице
         strSQL = GetUpdateStatement(TableName, sKEYVALUE, Criteria)
         If strSQL = "" Then GoTo AddRecordToTable_ERR
       '---------------------------------------------------------------------------------------------
    End If
 Else                                 ' Сразу пытаемся записать, без проверки
    If IsFieldExists("HASH", TableName) Then                      ' Добавляем HASH по которому потом найдем запись
       If InStr(1, sKEYVALUE, "HASH", vbBinaryCompare) = 0 Then   ' А кроме того, такое поле HASH не обнаружено
           sHASH = "HASH=" & GetHASH(TableName)
           sKEYVALUE = sKEYVALUE & ";" & sHASH
       End If
    End If
    strSQL = GetInsertIntoStatement(TableName, sKEYVALUE)
    If strSQL = "" Then GoTo AddRecordToTable_ERR
 End If
'----------------------------------------------------
' Выполняем запрос
    m_dbsCurrent.Execute strSQL
'-----------------------------------------------------
' Выходим из функции, вычисляя добавленное ID
        If Criteria = "" Then        ' Критерии не были заданы, например, при UPDATE - пытаемся веделить HASH
           If sHASH <> "" Then
              strSQL = sHASH
           ElseIf InStr(1, sKEYVALUE, "HASH") > 0 Then
              strSQL = "HASH=" & GetValueForKey(sKEYVALUE, "HASH")
           End If
        Else
           strSQL = Criteria
        End If
        AddRecordToTable = GetLastIdentity(TableName, strSQL, TimeStamp) '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        Exit Function
'-----------------------------------------------------
' Обрабатываем ошибки
AddRecordToTable_ERR:
    If Err <> 0 Then
        MsgBox Err.Description, vbExclamation, "AddRecordToTable Error " & Err.Number
    End If
'-----------------------------------------------------------------------------------------
Debug.Print "------------------"
Debug.Print strSQL
Debug.Print "------------------"
'-----------------------------------------------------------------------------------------
End Function
'==========================================================================================================================
' Удаление только одного значения поля при сохранении записи
'==========================================================================================================================
Public Function DeleteValueFromField(TableName As String, FieldName As String, sCriteria As String) As Boolean
Dim sSQL As String                                 ' Выражение запроса
Dim bRes As Boolean                                ' Возвращаемый результат

On Error GoTo DeleteValueFromField_ERR
sSQL = "UPDATE " & TableName & " SET " & TableName & "." & FieldName & " = Null " & _
       "WHERE ((" & sCriteria & "));"
       
DoCmd.SetWarnings False
    m_dbsCurrent.Execute sSQL
DoCmd.SetWarnings True

bRes = True
'-----------------------------------
DeleteValueFromField_Exit:
    DeleteValueFromField = bRes '!!!!!!!!!!!!!!!
    Exit Function
'-----------
DeleteValueFromField_ERR:
    MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "DeleteValueFromField ERROR"
    Err.Clear
    Resume DeleteValueFromField_Exit
End Function
'=============================================================================================================================================
' Функция создает запрос (если не указано наименование - запрос временный - просто запрос)
'=============================================================================================================================================P
Public Function CreateQueryDefX(SQL As String, Optional QueryName As String = "", Optional bRecreate As Boolean = True) As DAO.Recordset
Dim qdf As QueryDef                                  ' Создаваемый запрос

    With CurrentDb
        '-------------------------------------------------------------------------------------------------------
        If QueryName <> "" Then                      ' Имеется наименование запроса
           If bRecreate Then                         ' Пересоздаем запрос - удаляем старый
              DelQuery (QueryName)
              Set qdf = .CreateQueryDef(QueryName, SQL)
           Else                                      ' Открываем запрос
              Set qdf = .QueryDefs(QueryName)
           End If
        Else                                         ' Имя запроса не задано, запрос - временный
              Set qdf = .CreateQueryDef("", SQL)
        End If
        '--------------------------------------------------------------------------------------------------------
        ' Открывает объект Recordset и возвращаем результат
        Set CreateQueryDefX = GetRstFromQDF(qdf)  '!!!!!!!!!!!!!!!!!!!!
    End With
    Exit Function
'----------------------------------------------------------------------------------------------------------------
CreateQueryDefX_ERR:
   MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "CreateQueryDefX ERROR"
   Err.Clear
End Function
'=============================================================================================================================================
' Функция возвращает набор данных для заданного запроса
'=============================================================================================================================================P
Public Function GetRstFromQDF(qdf As QueryDef)
Dim rst As Recordset
    With qdf
        Set rst = .OpenRecordset(dbOpenSnapshot)
        With rst
            ' Заполняет объект Recordset и печатает число записей.
            .MoveLast
            .Close
        End With
    End With
'--------------------------------------
    Set GetRstFromQDF = rst '!!!!!!!!!!!!!!!!!
End Function

'===========================================================================================================================
' Функция удаляет  запрос, если он существует для текущей базы данных
'===========================================================================================================================
Public Function DelQuery(QueryName As String) As Boolean
Dim QD As QueryDef                          ' Рабочий запрос
Dim bRes As Boolean                         ' Возвращаемый элемент

On Error GoTo DelQuery_ERR

'-------------------------------------------------------------------
' Проверяем наличие запроса
If Not IsQueryExists(QueryName) Then GoTo DelQuery_EXIT  '  Если запроса нет, выходим
'---------------------------------------------------------------------
' удаляем объект ТАБЛИЦА
  m_dbsCurrent.QueryDefs.Delete QueryName
  bRes = True
'-----------------------------------------------------
' Выходим из функции
DelQuery_EXIT:
    DelQuery = bRes  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!1
    Set QD = Nothing
    Exit Function
'-----------------------------------------------------
' Обрабатываем ошибки
DelQuery_ERR:
    MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "DelQuery ERROR"
    Err.Clear
    Resume DelQuery_EXIT
End Function
'===========================================================================================================================
' Функция удаляет  таблицу, если она существует для текущей базы данных
'===========================================================================================================================
Public Function DelTable(TableName As String) As Boolean
On Error GoTo DelTable_ERR
Dim TD As TableDef                         ' Рабочая таблица
On Error GoTo DelTable_ERR

'-------------------------------------------------------------------
' Проверяем наличие таблицы
If Not IsTableExists(TableName) Then GoTo DelTable_ERR  '  Если таблицы нет, выходим
'---------------------------------------------------------------------
' удаляем объект ТАБЛИЦА
  m_dbsCurrent.TableDefs.Delete TableName
'-----------------------------------------------------
' Выходим из функции
    DelTable = True  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!1
    Exit Function
'-----------------------------------------------------
' Обрабатываем ошибки
DelTable_ERR:
    Set TD = Nothing
    DelTable = False  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

End Function
'============================================================================================================
' Функция создает базу данных программно с помощью механизма DAO
' PARAMETERS: sDatabaseName - имя создаваемой базы, включая полный путь к ней
'             bReplace      - если True - заменять существующий файл
'             sLocale       - данные по кодировке (по умолчанию Cyrrilic)
'             sVersion      - данные по версии базы (по умолчанию - текущая)
'             sPWD          - пароль на базу (по умолчанию пустой = нет пароля)
'             isEncrypt     - шифровать (по умолчанию False = нет)
'             bOpen         - сразу открыть базу (по умолчанию False = нет)
' RETURNS:    True - в случае успеха
'============================================================================================================
Public Function CreateDatabase(ByVal sDatabaseName As String, Optional bReplace As Boolean = False, _
                           Optional sLocale As String = DAO.dbLangCyrillic, Optional sVersion As String = "", _
                           Optional sPWD As String = "", Optional IsEncrypt As Boolean = False, _
                                                                Optional bOPen As Boolean = False) As Boolean
Dim iRes As Long                                  ' Возвращаемый результат
Dim NewDB As DAO.Database                         ' Создаваемая база данных
Dim sConnect As String, sOption As String         ' Вспомогательные параметры
On Error GoTo CreateDatabase_ERR
'---------------------------------------------------------------------------
' Проверяем, что база данных с указанным именем отсутствует
   If Dir(sDatabaseName) <> "" Then
      If bReplace Then
            Kill sDatabaseName
      Else
            iRes = MsgBox("Файл " & sDatabaseName & " существует. Заменить?", vbYesNo, "Замена файла")
            If iRes = vbYes Then      ' Удаляем файл
               Kill sDatabaseName
            Else                      ' Выходим из функции
               MsgBox "Невозможно создать файл " & sDatabaseName & ". Повторите операцию позже или " & _
                      "замените исходный файл.", vbCritical, "Отмена операции"
               Exit Function
            End If
      End If
   End If
'--------------------------------------------------------------------------
' Формируем вспомогательные параметры
  If sLocale = "" Then sLocale = DAO.dbLangGeneral
     sConnect = sLocale
     If sPWD <> "" Then sConnect = sConnect & ";pwd=" & sPWD
     '-----
  If sVersion <> "" Then sOption = sVersion
  If IsEncrypt Then sOption = sOption & DAO.dbEncrypt
'----------------------------------------------------------------------
' Создаем базу данных
  If sOption <> "" Then        ' создаем с опциями
    Set NewDB = DAO.CreateDatabase(sDatabaseName, sConnect, sOption)
  Else                         ' создаем без опций
    Set NewDB = DAO.CreateDatabase(sDatabaseName, sConnect)
  End If
'---------------------------------------------------------------------
CreateDatabase_EXIT:
    If bOPen Then              ' Открываем базу (если bOpen = True)
       Set m_dbsCurrent = NewDB
    Else
       Set NewDB = Nothing
    End If
    CreateDatabase = True   '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'-------------------------------------------------
CreateDatabase_ERR:
   MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "CreateDatabase ERROR"
   Err.Clear
End Function
'============================================================================================================
' Альтернативное создание таблицы
'        TableName - наименование таблицы
'        IsID_PK - если True, то первым идет поле, которое ставится в AutoIncrement и PK
'        bDelExistTBL - если True, проверяется существование заданной таблицы и при совпадении она уничтожается
'        TableNameCopy - если не пустая строка, то таблица создается как копия с заданной
'        FieldList - список полей, разделенных сепаратором FLDSEP. Каждое поле задается в виде SQL инструкции
'                    вида:  FIELDNAME TYPE [(size)] [NOT NULL]  [index1]
'                    Допускаются следующие типы: TEXT (size), CHAR (size), MEMO, BYTE, SHORT, LONG,SINGLE,
'                                             DOUBLE, GUID,DATETIME,CURRENCY,COUNTER (seed, increment), YESNO
'============================================================================================================
Public Function CreateTable2(ByVal TableName As String, Optional IsID_PK As Boolean = True, _
                               Optional bDelExistTBL As Boolean = True, Optional TableNameCopy As String = "", _
                               Optional FieldList As String = "", Optional FLDSEP As String = ";") As Boolean
Dim bRes As Boolean                                  ' Результат исполнения функции
Dim sSQL As String                                   ' SQL инструкция для исполнения

On Error GoTo ErrHandle
'-----------------------------------------------------------------------
       If bDelExistTBL Then                                             ' ПРОВЕРЯЕМ НАЛИЧИЕ ТАБЛИЦЫ
             If IsTableExists(TableName) Then
                   If Not DelTable(TableName) Then Exit Function        ' УДАЛЯЕМ ТАБЛИЦУ
             End If
       End If
       '----------------------------------------------------------------
       If TableNameCopy <> "" Then                                      ' КОПИРУЕМ ТАБЛИЦУ ИЛИ SQL
          sSQL = "SELECT * INTO " & TableName & "FROM " & TableNameCopy & ";"
       Else                                                             ' СОСТАВЛЯЕМ НОВУЮ ИНСТРУКЦИЮ
          sSQL = "CREATE TABLE " & TableName & "(" & FieldList & ");"
       End If
'-----------------------------------------------------------------------
        m_dbsCurrent.Execute sSQL
        bRes = True
'-----------------------------------------------------------------------
ExitHere:
        CreateTable2 = bRes '!!!!!!!!!!!!!!!!!!!
        Exit Function
'-----------------------------
ErrHandle:
        Debug.Print String(20, "#") & vbCrLf & "ERR#" & Err.Number & vbCrLf & Err.Description
        Err.Clear
End Function
'------------------------------------------------------------------------------------------------------------
' Функция осуществляет процессинг списка полей, осуществляя их проверку
'------------------------------------------------------------------------------------------------------------
Private Function CheckFLDList(Optional FieldList As String = "", _
                                 Optional FLDSEP As String = ";", Optional IsID_PK As Boolean = True) As String
Dim myFLD() As String, nDim As Integer, I As Integer
Dim mWords() As String, mDim As Integer, J As Integer
Dim sRes As String, sFirst As String

sFirst = IIf(IsID_PK, "AUTOINCREMENT PRIMARY KEY", "")
    If FieldList = "" Then
       sRes = "[ID] " & sFirst & ",[Text1] TEXT(255)"
    Else
       myFLD() = Split(FieldList, FLDSEP): nDim = UBound(myFLD)
       For I = 0 To nDim
          mWords() = Split(myFLD(I), " "): mDim = UBound(mWords)                 ' Парсим поле на слова
          If Left(mWords(0), 1) <> "[" Then mWords(0) = "[" & mWords(0) & "]"    ' Одеваем наименование поля
          If I = 0 And IsID_PK Then                                              ' Проверяем инструкцию на индекс
             If mWords(1) <> "AUTOINCREMENT" Then mWords(0) = mWords(0) & " " & sFirst
          End If
          
       Next I
    End If
End Function
'============================================================================================================
' Функция создает таблицу и заносит в нее необходимые поля. Возможны следующие варианты:
' A. Список полей создается за счет внутренней структуры класса с использованием функции AddFieldsToList
'                                                                          или CreateFieldList из строки
' B. Список полей создается непосредственно в функции - передается строка в формате:
'                            Fld1Name,Fld1Type,Fld1Size,IsFld1PK;Fld1Name,Fld2Type,Fld2Size,IsFld2PK,...
' C. Список полей отсутствует, в этом случае по умолчанию создается только одна строка
' При внешней базе данных требуется указать имя - полный путь DatabaseName
'============================================================================================================
Public Function CreateTable(ByVal TableName As String, Optional FieldsList As String = "") As Boolean
On Error GoTo CreateTable_ERR
Dim TD As TableDef                         ' Рабочая таблица
Dim f As Field
Dim bRes As Boolean
Dim FieldsCount As Integer, I As Integer

On Error GoTo CreateTable_ERR
'-------------------------------------------------------------------
' Проверяем наличие таблицы
If IsTableExists(TableName) Then GoTo CreateTable_ERR
'---------------------------------------------------------------------
' Eсли инициализируемся через строку, проверяем успех создания объекта
If FieldsList <> "" Then
    bRes = CreateFieldsList(FieldsList)
    If Not bRes Then GoTo CreateTable_ERR     ' Неудачное создание объекта, выходим
End If
'----------------------------------------------------
' Создаем объект ТАБЛИЦА и заполняем ее полями
Set TD = m_dbsCurrent.CreateTableDef(TableName)
    FieldsCount = UBound(FFIELDS)
    If FieldsCount > 0 Then         ' Берем поля из списка (он создан заранее или получен из строки)
        '------------------------------
        For I = 1 To FieldsCount
            Set f = TD.CreateField(FFIELDS(I).Fname, FFIELDS(I).fType)
                If FFIELDS(I).fType = dbText Then      ' Для текстового поля определяем размер
                   f.SIZE = FFIELDS(I).fSize
                End If
                ' Проверяем, не является ли поле ключем
                If FFIELDS(I).fType = dbLong And FFIELDS(I).FisPK = True Then
                        f.Attributes = dbAutoIncrField
                End If
                ' Присоединяем поле к таблице
                TD.FIELDS.Append f
                FFIELDS(I).FisReleased = True ' Отмечаем поле реализованным
        Next I
        '------------------------------
    Else                            ' Создаем единственное инкрементальное поле ID (PK)
      Set f = TD.CreateField("ID", dbLong)
          f.Attributes = dbAutoIncrField
          TD.FIELDS.Append f         'Присоединяем поле к таблице

    End If
'-----------------------------------------------------
' Присоединяем таблицу к базе
        m_dbsCurrent.TableDefs.Append TD
'-----------------------------------------------------
' Выходим из функции
    CreateTable = True  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'-----------------------------------------------------
' Обрабатываем ошибки
CreateTable_ERR:
    Set TD = Nothing
    Set f = Nothing
End Function
'===========================================================================================================================
' Проверка существования поля в заданной таблице в базе данных (по умолчанию - в текущей)
'===========================================================================================================================
Public Function IsFieldExists(FieldName As String, TableName As String) As Boolean
Dim f As Field                                            ' рабочее поле

On Error GoTo IsFieldExists_ERR

'----------------------------------------------------
' Проверяем существование таблицы в базе
If Not IsTableExists(TableName) Then GoTo IsFieldExists_ERR
'----------------------------------------------------
' Проверяем табличное пространство
On Error Resume Next
        IsFieldExists = (m_dbsCurrent.TableDefs(TableName).FIELDS(FieldName).Name = FieldName)

Exit Function
'--------------------------------------------------------
' ловушка
IsFieldExists_ERR:
        Err.Raise Err.Number
        Exit Function
End Function
'===========================================================================================================================
' Функция создает список полей из строчки формата
'                                Fld1Name,Fld1Type,Fld1Size,IsFld1PK;Fld1Name,Fld2Type,Fld2Size,IsFld2PK,...
'===========================================================================================================================
 Public Function CreateFieldsList(FieldsList As String) As Boolean
 Dim MyFLDS() As String                          ' Массив полей с параметрами через запятую
 Dim nFlds As Integer                            ' Размерность массива полей
 Dim myPARAMS() As String                        ' Массив параметров
 Dim I As Integer                                ' Циклическая переменная
 Dim bRes As Boolean                             ' результат
 On Error GoTo CreateFieldsList_ERR
 '--------------------------------------------------------------------------------
 ' Получаем массив полей
 MyFLDS = Split(FieldsList, ";")
 nFlds = UBound(MyFLDS)
 If nFlds = 0 Then               ' Не удалось создать массив
    CreateFieldsList = False
    Exit Function
 End If
'--------------------------------------------------------------------------------
 ' Получаем массив параметров, проходя в цикле массив полей
        For I = 0 To nFlds
           myPARAMS = Split(MyFLDS(I), ",")           ' Должны получить массив размерности 4 (0-3)
           ' Вызываем функцию добавления значений
             'bRes = AddFieldsToList(MyParams(0), MyParams(1), MyParams(2), CBool(MyParams(3)))
             
        Next I
        CreateFieldsList = True  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'---------------------------------------------------------------------------------
CreateFieldsList_ERR:
        CreateFieldsList = False  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        Exit Function
 End Function
'============================================================================================================================
' Принудительная очистка массива полей для текущей таблицы
'============================================================================================================================
 Public Sub ClearFieldsList()
    ReDim FFIELDS(0)
 End Sub
'===========================================================================================================================
' Проверка того, что массив полей не содержит информации
'===========================================================================================================================
Public Function IsFieldsListEmpty() As Boolean
  If UBound(FFIELDS) = 0 Then IsFieldsListEmpty = True
End Function
'============================================================================================================================
' Добавление к массиву полей единичного поля. Функция возвращает True в случае удачи.
' Если массив уже содержит PK, то указатель на PK для следующего поля игнорируется
'============================================================================================================================
Public Function AddFieldsToList(FieldName As String, FieldType As FFieldType, _
                           Optional FieldSize As Integer = 50, Optional isPK As Boolean = False) As Boolean
Dim FieldCount As Integer                                                      ' Текущая размерность массива
'------------------------------------------------
' Проверяем наличие имени в списке, если имя есть - выходим
If FFIELDSIndexOf(FieldName) > 0 Then
       AddFieldsToList = False      '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
       Exit Function
End If
'------------------------------------------------
' Увеличиваем размерность массива
        FieldCount = UBound(FFIELDS)
        FieldCount = FieldCount + 1
        ReDim Preserve FFIELDS(FieldCount)
'------------------------------------------------
' Вносим поля в массив
        FFIELDS(FieldCount).Fname = FieldName
        FFIELDS(FieldCount).fType = FieldType
        FFIELDS(FieldCount).fSize = FieldSize
        If isPK Then                              ' Поле претендует на PK
                FFIELDS(FieldCount).FisPK = Not IsExistPKinFFIELDS
        Else
                FFIELDS(FieldCount).FisPK = False ' Поле не претендует на PK
        End If
        FFIELDS(FieldCount).FisReleased = False
End Function
'============================================================================================================================
' Функция ищет данные в конкретном поле конкретной таблицы и заменяет их на соответсвующие
' PARAMETERS: strTable - имя таблицы, которая содержит данные для изменения
'             strSearchField - наименование поля, в котором осуществляется поиск
'             varSearchFor - значение, которое ищется
'             strChangeField = наименование поля, в котором изменяется значение
'             varChangeTo - новое значение, которое подставляется
'             fFirstOnly - True, если только первая замена, False - все замены
' RETURNS   : Число измененных записей
'============================================================================================================================
Public Function ReplaceDataSequential(strTable As String, strSearchField As String, varSearchFor As Variant, _
                           strChangeField As String, varChangeTo As Variant, Optional fFirstOnly As Boolean = True) As Long
  Dim rstTmp As DAO.Recordset
  Dim lngChanged As Long
  Dim strCriteria As String
  
  On Error GoTo PROC_ERR
  
'-----------------------------------------------------------------------------------------
' Формируем динамический рекордсет для таблицы, в которой производим поиск
  Set rstTmp = m_dbsCurrent.OpenRecordset(strTable, DAO.dbOpenDynaset)
  ' Выстраиваем строку с критериями отбора (поиска)
  strCriteria = strSearchField & "=" & varSearchFor

  With rstTmp
    ' Ищем первое совпадение
    .FindFirst strCriteria
    Do Until .NoMatch
      If Not .NoMatch Then
        ' Обновляем поле strChangeField
        .Edit
        .FIELDS(strChangeField).value = varChangeTo
        .Update
        ' Увеличиваем счетчик
        lngChanged = lngChanged + 1
  
        ' Если поддерживается только первая замена - сразу выходим из цикла, позволив сделать изменения один раз
        If fFirstOnly Then
          Exit Do
        End If
      End If
  
      ' Ищем следующие изменения
      .FindNext strCriteria
    Loop
  
    ' Закрываем рекордсет
    .Close
    
  End With
  
  ReplaceDataSequential = lngChanged  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'---------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'--------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ReplaceDataSequential"
  Resume PROC_EXIT
End Function
'============================================================================================================================
' Функция изменяет данные в таблицы по Primary Key
' PARAMETERS: strTable - наименование таблицы, в которой осуществляем поиск
'             varKeySearch - Primary Key для которого ищется запись
'             strChangeField - Поле, в котором содержатся данные для изменения
'             varChangeValue - Новое значение
' RETURNS   : True если запись была изменена, False - запись не изменена (например, не найдена)
'============================================================================================================================
Public Function ReplaceValueInKeyedTable(strTable As String, varKeySearch As Variant, strChangeField As String, _
                                                                                    varChangeValue As Variant) As Boolean
  Dim rstTmp As DAO.Recordset
  
  On Error GoTo PROC_ERR
  
'----------------------------------------------------------------------------
' Открываем рекордсет, основанный на таблице - используем свойство Index и метод Seek
  Set rstTmp = m_dbsCurrent.OpenRecordset(strTable, dbOpenTable)
  
  With rstTmp
    .Index = "PrimaryKey"         ' Специфицируем, какой из индексов используем. Не все таблицы имеют PK.
    .Seek "=", varKeySearch       ' Находим первое совпадение
    '----------------------------------------
    If Not .NoMatch Then          ' Если данные найдены
      .Edit                       ' обновляем запись
      .FIELDS(strChangeField).value = varChangeValue
      .Update
    End If
    '---------------------------------------
    ' Закрываем рекордсет
    .Close
    
  End With
'---------------------------------------------------------------------------
  ReplaceValueInKeyedTable = True  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'---------------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'--------------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ChangeValueInKeyedTable"
  Resume PROC_EXIT
End Function
'============================================================================================================================
' Функция скрывает или показывает объекты текущей базы
'============================================================================================================================
Public Sub HideAllObjects(Optional booHide As Boolean, Optional ExludeFormName As String = "_HOME")

    Dim TBL    As TableDef
    Dim qry    As QueryDef
    Dim str    As String
    Dim I      As Integer

    On Error Resume Next
    
    'Set db = CurrentDb()


    For Each TBL In m_dbsCurrent.TableDefs
            Call SetHiddenAttribute(acTable, TBL.Name, booHide)
    Next TBL


    For Each qry In m_dbsCurrent.QueryDefs
        Call SetHiddenAttribute(acQuery, qry.Name, booHide)
    Next qry


    For I = 0 To m_dbsCurrent.Containers("Forms").Documents.Count - 1
        str = m_dbsCurrent.Containers("Forms").Documents(I).Name
        If str <> ExludeFormName Then
            Call SetHiddenAttribute(acForm, str, booHide)
        End If
    Next



    For I = 0 To m_dbsCurrent.Containers("Reports").Documents.Count - 1
        str = m_dbsCurrent.Containers("Reports").Documents(I).Name
        Call SetHiddenAttribute(acReport, str, booHide)
    Next I


    For I = 0 To m_dbsCurrent.Containers("Modules").Documents.Count - 1
        str = m_dbsCurrent.Containers("Modules").Documents(I).Name
        Call SetHiddenAttribute(acModule, str, booHide)
    Next I

    For I = 0 To m_dbsCurrent.Containers("Scripts").Documents.Count - 1
        str = m_dbsCurrent.Containers("Scripts").Documents(I).Name
        Call SetHiddenAttribute(acMacro, str, booHide)
    Next I

    'Set db = Nothing
End Sub
'============================================================================================================================
' Функция возвращает список таблиц для открытой базы
'============================================================================================================================
Public Function TableList(Optional bNoSystem As Boolean = True) As String
Dim sRes As String                       ' Возвращаемый результат
Dim tdfs As DAO.TableDefs                ' Коллекция таблиц в базе
Dim TD As TableDef, tName As String      ' Таблица и ее имя

sRes = ""
'-------------------------------------------
' Проверяем, что открыта база
If Me.Database Is Nothing Then
   MsgBox "Нет назначеннной базы. Сбой", vbCritical, "tableList ERROR"
   GoTo tableList_EXIT
End If
'--------------------------------------------
Set tdfs = Me.Database.TableDefs         ' Табличное пространство данной базы
For Each TD In tdfs
    tName = TD.Name
    If bNoSystem And Left(tName, 4) = "MSys" Then GoTo NEXTTBL
    sRes = sRes & TD.Name & ";"
NEXTTBL:
Next TD
'-------------------------------------------
tableList_EXIT:
    TableList = sRes '!!!!!!!!!!!!!!!!!!!!!
    Set tdfs = Nothing
    Exit Function
'-------------------------
tableList_ERR:
   MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "tableList ERROR"
   GoTo tableList_EXIT
End Function
'============================================================================================================================
' Функция определяет, существует ли запись в базе
'============================================================================================================================
Public Function IsRecordExist(STABLE As String, sCriteria As String) As Boolean
Dim rst As DAO.Recordset                        ' Пробный запрос
Dim sSQL As String                              ' Выражение запроса
Dim bRes As Boolean                             ' Возвращаемый результат

On Error GoTo IsRecordExist_ERR
'------------------------------------------------
If m_dbsCurrent Is Nothing Then
   MsgBox "Не задана база данных для запроса", vbCritical, "IsRecordExist ERROR"
   GoTo IsRecordExist_EXIT
End If
  If STABLE = "" Or sCriteria = "" Then
   MsgBox "Не заданы параметры", vbCritical, "IsRecordExist ERROR"
   GoTo IsRecordExist_EXIT
  End If
'--------------------------------------------------
sSQL = "SELECT * FROM " & STABLE & " WHERE " & sCriteria & ";"

Set rst = m_dbsCurrent.OpenRecordset(sSQL)
   rst.MoveLast: rst.MoveFirst
   If rst.RecordCount <> 0 Then
      bRes = True
   End If
'--------------------------------------------------
'---------------------------------------------------------------------------
IsRecordExist_EXIT:
  IsRecordExist = bRes
  Exit Function
'--------------------------------------------------------------------------
IsRecordExist_ERR:
  MsgBox "ERR#: " & Err.Number & vbCrLf & Err.Description, vbCritical, "IsRecordExist ERROR"
  Resume IsRecordExist_EXIT
End Function
'============================================================================================================================
' Функция экспортирует таблицу в текст, столбцы которого разделены ";"
' Parameters: strTable - Name of the table containing the data to export
'             strFile - Name of the file to export to (if it exists,
'             it will first be deleted)
'             strFDelimit - Character(s) to use to delimit fields. If you
'             specify a blank string, the default used is a semi-colon.
' Returns   : Number of records processed
'============================================================================================================================
Public Function ExportTableToText(strTable As String, strFile As String, strFDelimit As String) As Long
  Dim rstTmp As DAO.Recordset
  Dim intFile As Integer
  Dim intCounter As Integer
  Dim StrTmp As String
  Dim lngCount As Long
  
  On Error GoTo PROC_ERR
'-------------------------------------------------------------------
' Инициализируем установки по умолчанию
  If strFDelimit = "" Then
    strFDelimit = ";"
  End If
'-----------------------------------------------
' если файл существует - предварительно его удаляем
  On Error Resume Next
  Kill strFile
  On Error GoTo PROC_ERR

  ' Открываем файл для записи
  intFile = FreeFile
  Open strFile For Output As intFile

  ' Создаем рекордсет и записываем его
  Set rstTmp = m_dbsCurrent.OpenRecordset(strTable, DAO.dbOpenDynaset)
  With rstTmp
    '--------------------------------------------------------------------------------
    ' Идем вдоль записей
    Do Until .EOF
      StrTmp = ""  ' Подготавливаем строку
      ' Прокручиваем каждое поле
      '---------------------------------------------
      For intCounter = 0 To .FIELDS.Count - 1
        ' Добавляем значение соотв. поля к строке
        StrTmp = StrTmp & .FIELDS(intCounter).value
        ' Если поле не является последним добавляем разделитель
        If intCounter < .FIELDS.Count Then
          StrTmp = StrTmp & strFDelimit
        End If
      Next intCounter
      '--------------------------------------------
      ' Записываем строку в файл
      Print #intFile, StrTmp
      ' Переходим к следующей записи
      .MoveNext
      ' Прирост счетчика
      lngCount = lngCount + 1
    Loop
    '---------------------------------------------------------------------------------
    ' Закрываем рекордсет
    .Close
  End With
  ' Закрываем файл
  Close #intFile
  
  ExportTableToText = lngCount  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'---------------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'--------------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ExportTableToText"
  Resume PROC_EXIT
End Function
'============================================================================================================================
' Функция ищет дубликаты для заданного поля заданной таблицы и переносит записи в отдельную специф. таблицу
' Parameters: strTable - Table to look for duplicates in
'             strNewTable - Table to create holding duplicates
'             strField - Field to match duplicates on
' Returns   : Name of the created table containing the duplicates
'============================================================================================================================
Public Function FindDuplicates(strTable As String, strNewTable As String, strField As String) As String
  Dim strSQL As String
  
  On Error GoTo PROC_ERR
  
  ' Build the name of the new table
  ' Build the SQL string with the SELECT INTO syntax. The HAVING clause
  ' is what causes the query to find duplicate values.
  strSQL = "SELECT * INTO [" & strNewTable & "] FROM [" & strTable & "] " & _
           "WHERE [" & strField & "] IN " & "(SELECT [" & strField & "] " & _
           "FROM [" & strTable & "] " & "GROUP BY [" & strField & "] " & _
           "HAVING Count(*) > 1) " & "ORDER BY [" & strField & "];"
  ' Execute the SQL string as a temporary query
  m_dbsCurrent.Execute strSQL
  ' Return the new table name
  FindDuplicates = strNewTable
'--------------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'--------------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "FindDuplicates"
  Resume PROC_EXIT
End Function
'============================================================================================================================
' Функция возвращает число найденных совпадений специфицированного текста по всем полям заданного источника
' Parameters: strFind - Text to search for
'             strSource - Name of a table, query, or a
'             SQL string defining the data to search in
' Returns   : Number of occurrences
'============================================================================================================================
Public Function SearchAllFields(strFind As String, strSource As String) As Long
  Dim rstTmp As DAO.Recordset
  Dim fldTmp As DAO.Field
  Dim lngHits As Long
  
  On Error GoTo PROC_ERR
  
  ' Open a recordset on the source
  Set rstTmp = m_dbsCurrent.OpenRecordset(strSource, dbOpenDynaset)

  ' Loop through fields and records looking for value
  rstTmp.MoveFirst
    
  With rstTmp
  
    ' Loop through each record
    Do Until .EOF
    
      ' Loop through each field
      For Each fldTmp In .FIELDS
        ' Does it match?
        If fldTmp.value = strFind Then
          ' It matches, so increment the hit counter
          lngHits = lngHits + 1
        End If
      Next fldTmp
  
      ' Go to the next record
      .MoveNext
    Loop
  
    ' Close the recordset
    .Close
  End With

  SearchAllFields = lngHits
'--------------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "SearchAllFields"
  Resume PROC_EXIT
  
End Function
'==========================================================================================================================
' Очищаем таблицы от информации. Возможны варианты:
'           sTableList="" и sExcludeTable = "" - удаляем информацию из всех таблиц
'           sTableList <> "" - удаляем информацию только из перечисленных таблиц
'           sExcludeTable <> "" - удаляем информацию из всех таблиц, кроме перечисленных
' Параметры sTableList и sExcludeTable могут содержать как список полных наименований таблиц, разделенных сепаратором ";",
' так и шаблоны вида xxx* (то есть определять таблицы с определенным суффиксом).
' Если sTableList пуст, то подразумевается, что удаляем все таблицы, кроме
' Сепаратор может быть изменен с помощью параметра sSeparator
' Если bConfirm установлен в true (по умолчанию), то выдается предупреждение
'==========================================================================================================================
Public Sub EmptyTables(Optional sTableList As String = "", Optional sExcludeTable As String = "", _
                                                 Optional sSeparator As String = ";", Optional bConfirm As Boolean = True)
  Dim tdfTmp As DAO.TableDef                ' Табличное пространство
  Dim strSQL As String                      ' Строка запроса
  
  On Error GoTo PROC_ERR
  '---------------------------------------------------------------------------------------------
  ' если установлено - выдаем предупреждение
  If bConfirm Then               ' флаг установлен в True
      If MsgBox("Вы собираетесь очистить данные из нескоьких таблиц! " & _
             "Продолжить? ", _
             vbQuestion + vbYesNo) <> vbYes Then GoTo PROC_EXIT   ' выходим из функции
  End If
  '----------------------------------------------------------------------------------------------
  ' Начинаем цикл по таблицам, для каждой получаем имя и сравниваем его с имеющимся шаблоном
    For Each tdfTmp In m_dbsCurrent.TableDefs
      If Left(tdfTmp.Name, 4) <> "MSys" Then         ' Пропускаем системные таблицы
        '----------------------------------------------------------------------------------------
        ' Проверяем необходимость удаления таблицы
        If ShouldDeliteThisTable(tdfTmp.Name, sTableList, sExcludeTable) Then
           ' Строим оператор SQL
           strSQL = "DELETE * FROM [" & tdfTmp.Name & "]"
           ' Выполняем инструкцию очистки
           m_dbsCurrent.Execute strSQL
        End If
        '----------------------------------------------------------------------------------------
      End If
    Next tdfTmp
'------------------------------------------------------------------------------------------------
PROC_EXIT:
  Exit Sub
'------------------------------------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "EmptyAllTables"
  Resume PROC_EXIT
  
End Sub
'==========================================================================================================================
' Проверяет, пустая ли таблица
' Parameters: strTable - name of table to check
' Returns   : True - table is empty, False otherwise
'==========================================================================================================================
Public Function IsTableEmpty(strTable As String) As Boolean
  Dim rstTmp As DAO.Recordset
  
  On Error GoTo PROC_ERR
  
  ' Open a recordset on the table
  Set rstTmp = m_dbsCurrent.OpenRecordset(strTable, dbOpenDynaset)

  ' If the EOF (end of file) and BOF (beginning of file) properties are
  ' both true, the recordset has no records.
  IsTableEmpty = (rstTmp.EOF And rstTmp.BOF)

  ' Close the objects
  rstTmp.Close
'--------------------------------------------------------------
PROC_EXIT:
  Exit Function
'--------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "IsTableEmpty"
  Resume PROC_EXIT
End Function

'==========================================================================================================================
' Возвращает число форм MS Access в базе данных
'==========================================================================================================================
Public Function GetAccessFormCount() As Long
  Dim conTmp As DAO.Container
  
  On Error GoTo PROC_ERR

  ' Получаем коллекцию форм
  Set conTmp = m_dbsCurrent.Containers!Forms
  
  With conTmp
    ' Обновляем (refresh)
    .Documents.Refresh
    ' Получаем число
    GetAccessFormCount = .Documents.Count   '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  End With
'-----------------------------------------------
PROC_EXIT:
  Exit Function
'-----------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetAccessFormCount"
  Resume PROC_EXIT

End Function
'==========================================================================================================================
' Возвращает число макросов
'==========================================================================================================================
Public Function GetAccessMacroCount() As Long
  Dim conTmp As DAO.Container
  
  On Error GoTo PROC_ERR

  ' For obscure reasons known only to the Microsoft Jet 1.0 authors,
  ' macros were called 'Scripts'.
  Set conTmp = m_dbsCurrent.Containers!Scripts
  
  With conTmp
    ' Refresh the collection to get the latest info
    .Documents.Refresh
    ' Return the count
    GetAccessMacroCount = .Documents.Count
  End With
  
PROC_EXIT:
  Exit Function
    
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetAccessMacroCount"
  Resume PROC_EXIT

End Function
'==========================================================================================================================
' Возвращает число модулей
'==========================================================================================================================
Public Function GetAccessModuleCount() As Long
  ' Comments  : Returns the number of Microsoft Access modules in the database
  '             Note that pure-VB Jet databases do not contain
  '             Access objects.
  ' Parameters: None
  ' Returns   : Count of modules
  ' Source    : Total VB SourceBook 6
  '
  Dim conTmp As DAO.Container
  
  On Error GoTo PROC_ERR
  
  ' Get a handle to the collection of module objects
  Set conTmp = m_dbsCurrent.Containers!Modules
  
  With conTmp
    ' Refresh teh collection to get the latest info
    .Documents.Refresh
    ' Return the count
    GetAccessModuleCount = .Documents.Count
  End With
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetAccessModuleCount"
  Resume PROC_EXIT

End Function
'==========================================================================================================================
' Возвращает число отчетов
'==========================================================================================================================
Public Function GetAccessReportCount() As Long
  Dim conTmp As DAO.Container
  
  On Error GoTo PROC_ERR

  ' Get a handle to the collection of report objects
  Set conTmp = m_dbsCurrent.Containers!Reports
  
  With conTmp
    ' Refresh the collection to get the latest info
    .Documents.Refresh
    ' Return the count
    GetAccessReportCount = .Documents.Count
  End With
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetAccessReportCount"
  Resume PROC_EXIT

End Function
'==========================================================================================================================
' Возвращает число индексов
'==========================================================================================================================
Public Function GetDatabaseIndexCount() As Long
  Dim tdfTmp As DAO.TableDef
  Dim lngCounter As Long
  Dim lngCount As Long
  
  On Error GoTo PROC_ERR

  ' Loop through each table
  For Each tdfTmp In m_dbsCurrent.TableDefs
  
    ' Don't consider system tables. In most cases, the consumer
    ' won't have permissions to read them anyway.
    If Left(tdfTmp.Name, 4) <> "MSys" Then
    
      ' Add the table's count of indexes to the total
      lngCounter = lngCounter + tdfTmp.INDEXES.Count
      lngCount = lngCount + 1
      
    End If
    
  Next tdfTmp
  
  GetDatabaseIndexCount = lngCount
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetDatabaseIndexCount"
  Resume PROC_EXIT
  
End Function

'==========================================================================================================================
' Возвращает число объектов (tables, queries, Access forms, reports, macros and modules)
'==========================================================================================================================
Public Function GetObjectCount() As Long
  Dim lngCount As Long
  Dim conTmp As DAO.Container
  
  On Error GoTo PROC_ERR
  
  ' Jet represents objects as Container objects
  For Each conTmp In m_dbsCurrent.Containers
    
    Select Case conTmp.Name
      
      Case "Tables", "Forms", "Reports", "Scripts", "Modules", "Relationships"
        lngCount = lngCount + conTmp.Documents.Count
      
      Case Else
        ' Do not count databases, sysrel containers

    End Select

  Next conTmp
  
  GetObjectCount = lngCount
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetObjectCount"
  Resume PROC_EXIT

End Function
'==========================================================================================================================
' Возвращает число запросов (хранимых процедур)
'==========================================================================================================================
Public Function GetQueryCount(fRefresh As Boolean) As Long
  ' Comments  : Returns the count of queries in the currently open database
  ' Parameters: fRefresh - True to refresh the database's
  '             relations collection, False to get the current state
  ' Returns   : Count of queries
  ' Source    : Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR
    
  If fRefresh Then
    ' Refresh the collection to get the latest info
    m_dbsCurrent.QueryDefs.Refresh
  End If
  
  ' Return the count
  GetQueryCount = m_dbsCurrent.QueryDefs.Count
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetQueryCount"
  Resume PROC_EXIT
  
End Function

'==========================================================================================================================
' Возращает тип специфического запроса
'==========================================================================================================================
Public Function GetQueryType(strQuery As String) As String
  ' Comments  : Returns the type of the specified query
  ' Parameters: strQuery - Name of the query
  ' Returns   : Query type as as string
  ' Source    : Total VB SourceBook 6
  '
  Dim qdfTmp As DAO.QueryDef

  On Error GoTo PROC_ERR
  
  ' Get a handle to the query's definition so we can check the Type property
  Set qdfTmp = m_dbsCurrent.QueryDefs(strQuery)

  Select Case qdfTmp.Type
    Case dbQSelect
      GetQueryType = "Select"
      
    Case dbQAction
      GetQueryType = "Action"
      
    Case dbQCrosstab
      GetQueryType = "Crosstab"
      
    Case dbQDelete
      GetQueryType = "Delete"
      
    Case dbQUpdate
      GetQueryType = "Update"
      
    Case dbQAppend
      GetQueryType = "Append"
      
    Case dbQMakeTable
      GetQueryType = "Make-table"
      
    Case dbQDDL
      GetQueryType = "Data-definition"
      
    Case dbQSQLPassThrough
      GetQueryType = "Pass-through"
      
    Case dbQSetOperation
      GetQueryType = "Union"
      
    Case Else
      GetQueryType = "Undefined"
    
  End Select
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetQueryType"
  Resume PROC_EXIT
  
End Function
'==========================================================================================================================
' Возвращает число отношений
'==========================================================================================================================
Public Function GetRelationCount(fRefresh As Boolean) As Long
  ' Comments  : Returns the count of relations in the currently open database
  ' Parameters: fRefresh - True to refresh the database's
  '             relations collection, False to get the current state
  ' Returns   : Count of relations
  ' Source    : Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR
    
  If fRefresh Then
    ' Refresh the collection to get the latest info
    m_dbsCurrent.Relations.Refresh
  End If
  
  ' Return the count
  GetRelationCount = m_dbsCurrent.Relations.Count
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetRelationCount"
  Resume PROC_EXIT
  
End Function
'==========================================================================================================================
' Возвращает число таблиц
'==========================================================================================================================
Public Function GetTableCount(fRefresh As Boolean) As Long
  ' Comments  : Returns the count of tables in the currently open database
  ' Parameters: fRefresh - True to refresh the database's
  '             tables collection, False to get the current state
  ' Returns   : Count of tables
  ' Source    : Total VB SourceBook 6
  '
  On Error GoTo PROC_ERR
    
  If fRefresh Then
    ' Refresh the collection to get the latest info
    m_dbsCurrent.TableDefs.Refresh
  End If
  
  ' Return the count
  GetTableCount = m_dbsCurrent.TableDefs.Count
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetTableCount"
  Resume PROC_EXIT
  
End Function
'==========================================================================================================================
' Возвращает число индексов для специфицированной таблицы
'==========================================================================================================================
Public Function GetTableIndexCount(strTableName As String) As Long
  On Error GoTo PROC_ERR

  ' Return the count
  GetTableIndexCount = m_dbsCurrent.TableDefs(strTableName).INDEXES.Count
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetTableIndexCount"
  Resume PROC_EXIT
  
End Function
'==========================================================================================================================
' Заполняет массив списком полей специфицированной таблицы для специфицированного индекса
'==========================================================================================================================
Public Function IndexFieldsToArray(strTable As String, strIndex As String, astrIn() As String) As Integer
  Dim intCounter As Integer
  Dim intCount As Integer
  Dim idxTmp As DAO.Index

  On Error GoTo PROC_ERR

  ' Get a handle to the index's defintion
  Set idxTmp = m_dbsCurrent.TableDefs(strTable).INDEXES(strIndex)
  
  ' Count the index's fields to resize the array
  intCount = idxTmp.FIELDS.Count
  
  If intCount > 0 Then
    ' Resize the array to hold all field names
    ReDim astrIn(0 To intCount - 1)
    For intCounter = 0 To intCount - 1
      ' Add the field name to the array
      astrIn(intCounter) = idxTmp.FIELDS(intCounter).Name
    Next intCounter
  End If

  IndexFieldsToArray = intCount
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "IndexFieldsToArray"
  Resume PROC_EXIT

End Function
'==========================================================================================================================
' Функция возвращает число индексов в заданной таблице, а также строку со списком всех индексов
' PARAMETERS: strTable - Name of table to look in
'             strIndex - Name of the index
'             strIn - String to fill
'             chrDelimit - Single character delimiter
' RETURNS   : Integer number of indexes
'==========================================================================================================================
Public Function IndexFieldsToString(strTable As String, strIndex As String, strIn As String, chrDelimit As String) As Integer
  Dim intCounter As Integer
  Dim intCount As Integer
  Dim idxTmp As DAO.Index
  
  On Error GoTo PROC_ERR

  ' Get a handle to the index's definition
  Set idxTmp = m_dbsCurrent.TableDefs(strTable).INDEXES(strIndex)

  ' Count the fields so we know when to stop
  intCount = idxTmp.FIELDS.Count
  
  If intCount > 0 Then
    For intCounter = 0 To intCount - 1
      ' Add the name of the field to the string
      strIn = strIn & idxTmp.FIELDS(intCounter).Name
      ' If we aren't on the last field, append the delimiter
      If intCounter < intCount - 1 Then
        strIn = strIn & chrDelimit
      End If
    Next intCounter
  End If

  IndexFieldsToString = intCount
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "IndexFieldsToString"
  Resume PROC_EXIT

End Function
'==========================================================================================================================
' Проверяет, является ли передаваемое имя валидным с точки зрения правил Jet. Если да - возвращается true
'==========================================================================================================================
Public Function IsValidJetName(strName As String) As Boolean
  Dim intCounter As Integer
  Dim chrTmp As String * 1
  Dim fValid As Boolean

  On Error GoTo PROC_ERR
  
  ' Has to be at least one character
  If Len(strName) = 0 Then
    fValid = False

  ' Can't exceed 64 characters
  ElseIf Len(strName) > 64 Then
    fValid = False

  ' Can't start with a space or =
  ElseIf Left$(strName, 1) = " " Or Left$(strName, 1) = "=" Then
    fValid = False

  Else
    ' Search for illegal characters
    fValid = True
    For intCounter = 1 To Len(strName)
      chrTmp = Mid$(strName, intCounter, 1)

      If Asc(chrTmp) < 32 Then
        fValid = False
      ElseIf chrTmp = "!" Or chrTmp = "." Or _
        chrTmp = "[" Or chrTmp = "]" Or chrTmp = "`" Then
        fValid = False
      End If
    Next intCounter
  End If

  IsValidJetName = fValid
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "IsValidJetName"
  Resume PROC_EXIT
  
End Function
'==========================================================================================================================
' Определяет, существует ли именованный объект в текущей базе данных. Если да - возвращается true
' PARAMETERS: eType - Type of the object as defined by the
'             EnumCJetObjectType enumerated type
'             strName - name of the object to check
' RETURNS   : True if the object exists, False otherwise
'==========================================================================================================================
Public Function ObjectExists(eType As EnumCJetObjectType, strName As String) As Boolean
  Dim StrTmp As String
  
  On Error GoTo PROC_ERR
  
  Select Case eType
  
    Case cjoTypeTable
    
      ' The method we use to determine an object's existence is to
      ' disable error handling and try to refer to the object. If
      ' the object doesn't exist, an error will occur.
      
      On Error Resume Next
      StrTmp = m_dbsCurrent.TableDefs(strName).Name
      ObjectExists = (Err.Number = 0)
      On Error GoTo PROC_ERR
      
    Case cjoTypeQuery
      On Error Resume Next
      StrTmp = m_dbsCurrent.QueryDefs(strName).Name
      ObjectExists = (Err.Number = 0)
      On Error GoTo PROC_ERR
    
    Case cjoTypeRelation
    On Error Resume Next
      StrTmp = m_dbsCurrent.Relations(strName).Name
      ObjectExists = (Err.Number = 0)
      On Error GoTo PROC_ERR
    
    Case cjoTypeAccessForm
      On Error Resume Next
      StrTmp = m_dbsCurrent.Containers("Forms").Documents(strName).Name
      ObjectExists = (Err.Number = 0)
      On Error GoTo PROC_ERR
      
    Case cjoTypeAccessReport
      On Error Resume Next
      StrTmp = m_dbsCurrent.Containers("Reports").Documents(strName).Name
      ObjectExists = (Err.Number = 0)
      On Error GoTo PROC_ERR
    
    Case cjoTypeAccessMacro
      On Error Resume Next
      StrTmp = m_dbsCurrent.Containers("Scripts").Documents(strName).Name
      ObjectExists = (Err.Number = 0)
      On Error GoTo PROC_ERR
    
    Case cjoTypeAccessModule
      On Error Resume Next
      StrTmp = m_dbsCurrent.Containers("Modules").Documents(strName).Name
      ObjectExists = (Err.Number = 0)
      On Error GoTo PROC_ERR
    
    Case Else
      ' Function was passed an invalid object type
      Err.Raise vbObjectError + 1, , "Invalid object type."
    
  End Select
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ObjectExists"
  Resume PROC_EXIT
  
End Function
'============================================================================================================================
' Функция преобразовывает запись номер N в Recordset к виду KVString (Key=Value)
' PARAMETERS:  RS       - набор записей для преобразования
'              NRECORD  - запись, для которой строим KVString. Если отсутствует, то для первой
'              sDelim   - разделитель  между KV - парами в строке (по умолчанию (;))
'              lDelim   - разделитель между Key и Value (по умолчанию (=))
'============================================================================================================================
Public Function GetKVStringFromRS(RS As DAO.Recordset, Optional nRecord As Long, Optional sDelim As String = ";", _
                                                                                  Optional lDelim As String = "=") As String
Dim FieldName As String, intFieldCount As Integer, I As Integer       ' Наименование поля, кол-во полей, цикл. переменная
Dim varField As Variant                                               ' Текущее значение
Dim sWork As String                                                   ' Рабочая строка

On Error GoTo PROC_ERR

sWork = "": intFieldCount = RS.FIELDS.Count
'-------------------------------------------------------------------
' ИДЕМ ПО ПОЛЯМ RS записи номер NRECORD, формируя соотв. строку
  With RS
      If .EOF Then GoTo PROC_EXIT
      '-----------------------------
      If nRecord = 0 Then
        .MoveFirst
      Else
        .MoveFirst: .Move nRecord - 1    ' Переходим на запись NRECORD, считая от нуля
      End If
      '-----------------------------
      
      For I = 0 To intFieldCount - 1      ' Цикл по всем полям
        FieldName = RS.FIELDS(I).Name           ' Текущее имя поля
        varField = .FIELDS(I).value             ' Текущее значение
        If Not IsNull(varField) Then            ' Исключаем нулевые значения
          ' ДОБАВЛЯЕМ ПАРУ
          If sWork <> "" Then sWork = sWork & sDelim
          '------------------------------------------------
          sWork = sWork & FieldName & lDelim & varField
          '------------------------------------------------
        End If
      Next I
  End With
'-----------------------------------------------------------------------
PROC_EXIT:
    GetKVStringFromRS = sWork   '!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'---------------------------------------
PROC_ERR:
   MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "GetKVStringFromRS ERROR"
End Function
'==========================================================================================================================
' Заполняет массив наименованиями объектов определенного типа в текущей базе данных
' PARAMETERS: eType - Type of the object as defined by the
'             EnumCJetObjectType enumerated type
'             astrIn - Array of string variables (0-based)
' RETURNS   : Long integer indicating number of array elements
'==========================================================================================================================
Public Function ObjectsToArray(eType As EnumCJetObjectType, astrIn() As String) As Long
  Dim lngCount As Long
  Dim lngCounter As Long
  Dim conTmp As DAO.Container

  On Error GoTo PROC_ERR

  ' Figure out which type of object is specified
  Select Case eType

    Case cjoTypeTable
      ' For tables, use the DAO TableDefs collection
      lngCount = m_dbsCurrent.TableDefs.Count
      If lngCount > 0 Then
        ReDim astrIn(0 To lngCount - 1)
        For lngCounter = 0 To lngCount - 1
          astrIn(lngCounter) = m_dbsCurrent.TableDefs(lngCounter).Name
        Next lngCounter
      End If
 
    Case cjoTypeQuery
      ' For queries, use the DAO QueryDefs collection
      lngCount = m_dbsCurrent.QueryDefs.Count
      If lngCount > 0 Then
        ReDim astrIn(0 To lngCount - 1)
        For lngCounter = 0 To lngCount - 1
          astrIn(lngCounter) = m_dbsCurrent.QueryDefs(lngCounter).Name
        Next lngCounter
      End If

    Case cjoTypeRelation
      ' Fo relations, use the DAO Relations collection
      lngCount = m_dbsCurrent.Relations.Count
      If lngCount > 0 Then
        ReDim astrIn(0 To lngCount - 1)
        For lngCounter = 0 To lngCount - 1
          astrIn(lngCounter) = m_dbsCurrent.Relations(lngCounter).Name
        Next lngCounter
      End If
      
    Case Else
    
      ' The following handles Access objects, such as forms, reports,
      ' macros, and modules. These objects are not native to the DAO and
      ' Jet engine hierarchies, so they are contained in generic
      ' containers that Access defines.
      Select Case eType
        
        Case cjoTypeAccessForm
          Set conTmp = m_dbsCurrent.Containers("Forms")
        
        Case cjoTypeAccessReport
          Set conTmp = m_dbsCurrent.Containers("Reports")
        
        Case cjoTypeAccessMacro
          Set conTmp = m_dbsCurrent.Containers("Scripts")
        
        Case cjoTypeAccessModule
          Set conTmp = m_dbsCurrent.Containers("Modules")
        
        Case Else
          ' Function was passed an invalid object type
          Err.Raise vbObjectError + 2, CurrentProject.Name, "Invalid object type " & _
            Err.Number
      
      End Select

      ' Get the count of objects to resize the array
      lngCount = conTmp.Documents.Count
        
      If lngCount > 0 Then
        ' Resize the array
        ReDim astrIn(0 To lngCount - 1)
        ' Add the items
        For lngCounter = 0 To lngCount - 1
          astrIn(lngCounter) = conTmp.Documents(lngCounter).Name
        Next lngCounter
      End If
      
  End Select

  ObjectsToArray = lngCount
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ObjectsToArray"
  Resume PROC_EXIT

End Function
'==========================================================================================================================
' Функция возвращает число объектов и строку-список с именами объектов определенного типа
' PARAMETERS: eType - Type of the object as defined by the
'             EnumCJetObjectType enumerated type
'             strIn - string to fill with object names
'             chrDelimit - Single character delimiter to seperate names
' RETURNS   : Long integer indicating number of names added to string
'==========================================================================================================================
Public Function ObjectsToString(eType As EnumCJetObjectType, ByRef strIn As String, chrDelimit As String) As Long
  Dim conTmp As DAO.Container
  Dim lngCount As Long
  Dim lngCounter As Long

  On Error GoTo PROC_ERR

  ' Determine which object type was specified
  Select Case eType

    Case cjoTypeTable
      ' For tables, use the DAO TableDefs collection
      lngCount = m_dbsCurrent.TableDefs.Count
      If lngCount > 0 Then
        For lngCounter = 0 To lngCount - 1
          strIn = strIn & m_dbsCurrent.TableDefs(lngCounter).Name
          If lngCounter < lngCount - 1 Then
            strIn = strIn & chrDelimit
          End If
        Next lngCounter
      End If
 
    Case cjoTypeQuery
      ' For queries, use the DAO QueryDefs collection
      lngCount = m_dbsCurrent.QueryDefs.Count
      If lngCount > 0 Then
        strIn = strIn & m_dbsCurrent.QueryDefs(lngCounter).Name
        If lngCounter < lngCount - 1 Then
          strIn = strIn & chrDelimit
        End If
      End If

    Case cjoTypeRelation
      ' For relations, use the DAO Relations collection
      lngCount = m_dbsCurrent.Relations.Count
      If lngCount > 0 Then
        strIn = strIn & m_dbsCurrent.Relations(lngCounter).Name
        If lngCounter < lngCount - 1 Then
          strIn = strIn & chrDelimit
        End If
      End If

    Case Else
     ' The following handles Access objects, such as forms, reports,
      ' macros, and modules. These objects are not native to the DAO and
      ' Jet engine hierarchies, so they are contained in generic
      ' containers that Access defines.
      Select Case eType
        
        Case cjoTypeAccessForm
          Set conTmp = m_dbsCurrent.Containers("Forms")
        
        Case cjoTypeAccessReport
          Set conTmp = m_dbsCurrent.Containers("Reports")
        
        Case cjoTypeAccessMacro
          Set conTmp = m_dbsCurrent.Containers("Scripts")
        
        Case cjoTypeAccessModule
          Set conTmp = m_dbsCurrent.Containers("Modules")
        
        Case Else
          ' Function was passed an invalid object type
          Err.Raise vbObjectError + 2, CurrentProject.Name, "Error: " & _
            Err.Number & ", "
      
      End Select

      ' Get the count of objects to resize the array with
      lngCount = conTmp.Documents.Count
      If lngCount > 0 Then
        ' Add the name to the string
        strIn = strIn & conTmp.Documents(lngCounter).Name
        ' If we aren't on the last one, append the delimiter
        If lngCounter < lngCount - 1 Then
          strIn = strIn & chrDelimit
        End If
      End If
      
  End Select

  ObjectsToString = lngCount
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
   "ObjectsToString"
  Resume PROC_EXIT

End Function
'==========================================================================================================================
' Функция открывает базу данных
' PARAMETERS: strDatabaseName - Fully qualified path and name of the
'             database to open
'             fExclusive - True to open the database in exclusive mode,
'             False to open in shared mode
'             fReadOnly - True to open the database in read-only mode,
'             False to open for read/write access
'             varDatabasePassword - optional password if the database has a
'             password
' RETURNS   : True if successful, False otherwise
'==========================================================================================================================
Public Function OpenDB(strDatabaseName As String, fExclusive As Boolean, fReadOnly As Boolean, _
                                                                         Optional varDatabasePassword As Variant) As Boolean
  Dim strConnect As String
  
  On Error GoTo PROC_ERR
  
  ' Set this variable to True so we know later on that the class opened
  ' the database.
  mfClassOpened = True
  
  m_fExclusive = fExclusive
  m_fReadonly = fReadOnly
  
  ' If a password is specified, pass it to the OpenDatabase method through the
  ' Connect parameter.
  If Not IsMissing(varDatabasePassword) Then
    strConnect = ";pwd=" & varDatabasePassword
    m_varDatabasePassword = varDatabasePassword
  End If
  
  ' Open the database with the options specified
  Set m_dbsCurrent = DAO.DBEngine.OpenDatabase(strDatabaseName, _
    fExclusive, _
    fReadOnly, _
    strConnect)
  
  ' Save the opendatabase options for the compact/repair/encrypt methods
  mfExclusive = fExclusive
  mfReadOnly = fReadOnly
  If Not IsMissing(varDatabasePassword) Then
    mvarDatabasePassword = varDatabasePassword
  End If
  
  OpenDB = True
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "OpenDB"
  OpenDB = False
  Resume PROC_EXIT
  
End Function
'==========================================================================================================================
' Заполняет массив полями запроса
' PARAMETERS: strQuery - Name of query to look in
'             astrIn()  - Array to populate (0-based)
'             fTypeInfo - True to include field type info,
'             False to include only the field name
' RETURNS   : Integer number of elements in array
'==========================================================================================================================
Public Function QueryFieldsToArray(strQuery As String, astrIn() As String, fTypeInfo As Boolean) As Integer
  Dim intCounter As Integer
  Dim intCount As Integer
  Dim qdfTmp As DAO.QueryDef

  On Error GoTo PROC_ERR

  ' Get a handle to the query's definition
  Set qdfTmp = m_dbsCurrent.QueryDefs(strQuery)

  ' Count the fields for the array size
  intCount = qdfTmp.FIELDS.Count
  
  If intCount > 0 Then
    ' Resize the array to hold all fields
    ReDim astrIn(0 To intCount - 1)
    For intCounter = 0 To intCount - 1
      ' Add the field name to the array
      astrIn(intCounter) = qdfTmp.FIELDS(intCounter).Name
      ' If type information is specified, add it to the array
      If fTypeInfo Then
        astrIn(intCounter) = astrIn(intCounter) & " (" & _
          qdfTmp.FIELDS(intCounter).Type & ")"
      End If
    Next intCounter
  End If

  QueryFieldsToArray = intCount
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "QueryFieldsToArray"
  Resume PROC_EXIT

End Function

'==========================================================================================================================
' Возвращает число полей запроса, одновременно формируя строку со списком полей запроса
' PARAMETERS: strQuery - Name of table to look in
'             strIn - String to fill
'             chrDelimit - Single character delimiter
'             fTypeInfo - True to include field type info, False to include
'             only the name
' RETURNS   : Integer number of fields
'==========================================================================================================================
Public Function QueryFieldsToString(strQuery As String, strIn As String, chrDelimit As String, fTypeInfo As Boolean) As Integer
  Dim intCounter As Integer
  Dim intCount As Integer
  Dim qdfTmp As DAO.QueryDef

  On Error GoTo PROC_ERR

  ' Get a handle to the query's definition
  Set qdfTmp = m_dbsCurrent.QueryDefs(strQuery)

  ' Count the fields so we know when to stop
  intCount = qdfTmp.FIELDS.Count
  
  If intCount > 0 Then
    For intCounter = 0 To intCount - 1
      ' Add the field's name to the string
      strIn = strIn & qdfTmp.FIELDS(intCounter).Name
      ' If type information is specified, add it to the string
      If fTypeInfo Then
        strIn = strIn & " (" & qdfTmp.FIELDS(intCounter).Type & ")"
      End If
      
      ' If we aren't on the last field, append the delimiter
      If intCounter < intCount - 1 Then
        strIn = strIn & chrDelimit
      End If
    Next intCounter
  
  End If

  QueryFieldsToString = intCount
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "QueryFieldsToString"
  Resume PROC_EXIT

End Function
'==========================================================================================================================
' Функция возвращает число параметров запроса, а также заполняет массив этими параметрами
' Parameters: strQuery - Name of query to look in
'             astrIn()  - Array to populate (0-based)
' Returns   : Integer number of elements in array
'==========================================================================================================================
Public Function QueryParametersToArray(strQuery As String, astrIn() As String) As Integer
  Dim intCounter As Integer
  Dim intCount As Integer
  Dim qdfTmp As DAO.QueryDef

  On Error GoTo PROC_ERR
  
  ' Get a handle the query's defintion
  Set qdfTmp = m_dbsCurrent.QueryDefs(strQuery)

  ' Count the query's parameters to resize the array
  intCount = qdfTmp.Parameters.Count
  
  If intCount > 0 Then
    ' Resize the array to hold all parameters
    ReDim astrIn(0 To intCount - 1)
    For intCounter = 0 To intCount - 1
      ' Get the parameter name
      astrIn(intCounter) = qdfTmp.Parameters(intCounter).Name
    Next intCounter
  End If

  QueryParametersToArray = intCount
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "QueryParametersToArray"
  Resume PROC_EXIT

End Function
'==========================================================================================================================
' Функция возвращает число параметров запроса и формирует строку - список из параметров запроса
' Parameters: strQuery - Name of table to look in
'             strIn - String to fill
'             chrDelimit - Single character delimiter
' Returns   : Integer number of parameters
'==========================================================================================================================
Public Function QueryParametersToString(strQuery As String, strIn As String, chrDelimit As String) As Integer
  Dim qdfTmp As DAO.QueryDef
  Dim intCounter As Integer
  Dim intCount As Integer

  On Error GoTo PROC_ERR

  ' Get a handle to the query's definition
  Set qdfTmp = m_dbsCurrent.QueryDefs(strQuery)

  ' Count the parameters so we know when to stop
  intCount = qdfTmp.Parameters.Count
  
  If intCount > 0 Then
    For intCounter = 0 To intCount - 1
      ' Add the parameter name to the string
      strIn = strIn & qdfTmp.Parameters(intCounter).Name
      ' If we aren't on the last parameter, append the delimiter
      If intCounter < intCount - 1 Then
        strIn = strIn & chrDelimit
      End If
    Next intCounter
  End If

  QueryParametersToString = intCount
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "QueryParametersToString"
  Resume PROC_EXIT

End Function
'==========================================================================================================================
' Связи между полями в массив. Возвращает число связей и формирует указанный массив
' Parameters: strRelation - Name of Relation to look in
'             astrIn()  - Array to populate (0-based)
' Returns   : Integer number of elements in array
'==========================================================================================================================
Public Function RelationFieldsToArray(strRELATION As String, astrIn() As String) As Integer
  Dim intCounter As Integer
  Dim intCount As Integer
  Dim relTmp As DAO.Relation

  On Error GoTo PROC_ERR

  ' Get a handle to the relation's definition
  Set relTmp = m_dbsCurrent.Relations(strRELATION)

  ' Count the fields for the array size
  intCount = relTmp.FIELDS.Count
  
  If intCount > 0 Then
    ' Resize the array
    ReDim astrIn(0 To intCount - 1)
    For intCounter = 0 To intCount - 1
      ' Get the field name
      astrIn(intCounter) = relTmp.FIELDS(intCounter).Name
    Next intCounter
  End If
  
  RelationFieldsToArray = intCount
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "RelationFieldsToArray"
  Resume PROC_EXIT

End Function
'==========================================================================================================================
' Связи между полями в массив. Возвращает число связей и формирует строку - список со связами
' Parameters: strRelation - Name of Relation to look in
'             strIn - String to fill
'             chrDelimit - Single character delimiter
' Returns   : Integer number of fields
'==========================================================================================================================
Public Function RelationFieldsToString(strRELATION As String, strIn As String, chrDelimit As String) As Integer

  Dim intCounter As Integer
  Dim intCount As Integer
  Dim relTmp As DAO.Relation
    
  On Error GoTo PROC_ERR
  
  ' Get a handle to the relation's definition
  Set relTmp = m_dbsCurrent.Relations(strRELATION)

  ' Get the field count so we know when to stop
  intCount = relTmp.FIELDS.Count
  
  If intCount > 0 Then
    For intCounter = 0 To intCount - 1
      ' Add the fields name to the string
      strIn = strIn & relTmp.FIELDS(intCounter).Name
      ' If we're not on the last field, append the delimiter
      If intCounter < intCount - 1 Then
        strIn = strIn & chrDelimit
      End If
    Next intCounter
  End If

  RelationFieldsToString = intCount
'--------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "RelationFieldsToString"
  Resume PROC_EXIT

End Function
'==========================================================================================================================
' Функция закрывает текущую базу данных, восстанавливает это и открывает заново. Возвращает  размер восстановленной базы
' в байтах
'==========================================================================================================================
Public Function RepairDatabase() As Long
  Dim strFullPath As String
  
  On Error GoTo PROC_ERR

  ' Get the path so we know what to re-open
  strFullPath = m_dbsCurrent.Name
  
  ' Close and deinstantiate the database handle
  m_dbsCurrent.Close
  Set m_dbsCurrent = Nothing
  
  ' Repair the database
  DAO.DBEngine.RepairDatabase strFullPath
  
  ' Get the size of the resulting file. Note that we do this before we
  ' open the database. The act of opening the database changes its
  ' file size, so we want to do this first.
  RepairDatabase = FileLen(strFullPath)
  
  ' Reopen the database with the same options
  OpenDB strFullPath, mfExclusive, mfReadOnly, mvarDatabasePassword
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "RepairDatabase"
  Resume PROC_EXIT

End Function
'==========================================================================================================================
' Заполняет массив списком полей заданной таблицы и возвращает число таких полей
' Parameters: strTable - Name of table to look in
'             astrIn()  - Array to populate (0-based)
'             fTypeInfo - True to include field type info,
'                         False to include only the field name
' Returns   : Integer number of elements in array
'==========================================================================================================================
Public Function TableFieldsToArray(strTable As String, astrIn() As String, fTypeInfo As Boolean) As Integer
  Dim intCounter As Integer
  Dim intCount As Integer
  Dim tdfTmp As DAO.TableDef

  On Error GoTo PROC_ERR

  ' Get a pointer to the table's definition
  Set tdfTmp = m_dbsCurrent.TableDefs(strTable)

  ' Get a count of the fields to resize the array
  intCount = tdfTmp.FIELDS.Count
  
  If intCount > 0 Then
    ' Resize the array
    ReDim astrIn(0 To intCount - 1)
    ' Loop through the objects
    For intCounter = 0 To intCount - 1
      ' Get the field name
      astrIn(intCounter) = tdfTmp.FIELDS(intCounter).Name
      ' If type information is specified, get the field's type
      If fTypeInfo Then
        astrIn(intCounter) = astrIn(intCounter) & " (" & _
          tdfTmp.FIELDS(intCounter).Type & ")"
      End If
    Next intCounter
  End If

  TableFieldsToArray = intCount
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "TableFieldsToArray"
  Resume PROC_EXIT

End Function
'==========================================================================================================================
' Формирует строку - список из полей заданной таблицы и возвращает число полей.
' Parameters: strTable - Name of table to look in
'             strIn - String to fill
'             chrDelimit - Single character delimiter
'             fTypeInfo - True to include field type info, False to include
'             only the name
' Returns   : Number of fields
'==========================================================================================================================
Public Function TableFieldsToString(strTable As String, strIn As String, chrDelimit As String, fTypeInfo As Boolean) As Integer
  Dim intCounter As Integer
  Dim intCount As Integer
  Dim tdfTmp As DAO.TableDef

  On Error GoTo PROC_ERR

  ' Get a handle to the table's structure
  Set tdfTmp = m_dbsCurrent.TableDefs(strTable)

  ' Get a count of fields so we know when to stop
  intCount = tdfTmp.FIELDS.Count
  
  If intCount > 0 Then
    For intCounter = 0 To intCount - 1
      ' Add the name
      strIn = strIn & tdfTmp.FIELDS(intCounter).Name
      ' Add the type if specified
      If fTypeInfo Then
        strIn = strIn & " (" & tdfTmp.FIELDS(intCounter).Type & ")"
      End If
      If intCounter < intCount - 1 Then
        strIn = strIn & chrDelimit
      End If
    Next intCounter
  End If

  TableFieldsToString = intCount
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "TableFieldsToString"
  Resume PROC_EXIT

End Function
'==========================================================================================================================
' Заполняет массив списком индексов заданной таблицы и возвращает число таких индексов.
' Parameters: strTable - Name of table to look in
'             astrIn()  - Array to populate (0-based)
' Returns   : Integer number of elements in array
'==========================================================================================================================
Public Function TableIndexesToArray(strTable As String, astrIn() As String) As Integer
  Dim intCounter As Integer
  Dim intCount As Integer
  Dim tdfTmp As DAO.TableDef

  On Error GoTo PROC_ERR

  ' Get a handle to the table's definition
  Set tdfTmp = m_dbsCurrent.TableDefs(strTable)

  ' Get the count of indexes to resize the array
  intCount = tdfTmp.INDEXES.Count
  
  If intCount > 0 Then
    ' Resize the array to hold all index names
    ReDim astrIn(0 To intCount - 1)
    For intCounter = 0 To intCount - 1
      ' Add the index name
      astrIn(intCounter) = tdfTmp.INDEXES(intCounter).Name
    Next intCounter
  End If

  TableIndexesToArray = intCount
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "TableIndexesToArray"
  Resume PROC_EXIT

End Function
'==========================================================================================================================
' Формирует строку- список с индексами заданной таблицы и возвращает число таких индексов.
' Parameters: strTable - Name of table to look in
'             strIn - String to fill
'             chrDelimit - Single character delimiter
' Returns   : Integer number of indexes
'==========================================================================================================================
Public Function TableIndexesToString(strTable As String, strIn As String, chrDelimit As String) As Integer
  Dim tdfTmp As DAO.TableDef
  Dim intCounter As Integer
  Dim intCount As Integer

  On Error GoTo PROC_ERR

  ' Get a handle to the table's defintion
  Set tdfTmp = m_dbsCurrent.TableDefs(strTable)

  ' Count the indexes so we know when to stop
  intCount = tdfTmp.INDEXES.Count
  
  If intCount > 0 Then
    For intCounter = 0 To intCount - 1
      ' Add the index's name
      strIn = strIn & tdfTmp.INDEXES(intCounter).Name
      ' If we aren't on the last index, append the delimiter
      If intCounter < intCount - 1 Then
        strIn = strIn & chrDelimit
      End If
    Next intCounter
  End If

  TableIndexesToString = intCount
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "TableIndexesToString"
  Resume PROC_EXIT

End Function
'==========================================================================================================================
' Создает ссылку на внешнюю таблицу в заданной Access базе данных.
' Parameters: strDatabase - полный путь, включающий имя к базе данных,  к таблице которой строится ссылка
'             strSourceTable - имя таблице к которой строится ссылка
'             strTableName - имя линкованной таблицы (если опущена - совпадает с таблицей - источником)
' Returns   : True - в случае успешного создания ссылки
'==========================================================================================================================
Public Function CreateJetLink(strDatabase As String, strSourceTable As String, _
                                                                        Optional strTableName As String = "") As Boolean
Dim tdfTmp As DAO.TableDef                      ' Объект таблицы, для которого строятся ссылки
  
  On Error GoTo PROC_ERR
  
'----------------------------------------------------------------------
' Определяем имя прилинкованной таблицы
  If strTableName = "" Then strTableName = strSourceTable
'----------------------------------------------------------------------
' Создаем новый объект таблицы
  Set tdfTmp = m_dbsCurrent.CreateTableDef(strTableName)
  ' Устанавливаем свойства, необходимые для подсоединения таблицы
  tdfTmp.Connect = mcstrLinkJet & "DATABASE=" & strDatabase
  tdfTmp.SourceTableName = strSourceTable
'--------------------------------------------------------------------
' Сохраняем таблицу в базе данных
  m_dbsCurrent.TableDefs.Append tdfTmp
  CreateJetLink = True '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'-------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'-------------------------------------------------------------------
PROC_ERR:
  CreateJetLink = False '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "CreateJetLink"
  Resume PROC_EXIT
End Function
'==========================================================================================================================
' Функция создает ссылку на таблицу в источнике DSN (ODBC).
' Parameters: strDSN - имя зарегистрированного источника  ODBC
'             strSourceTable  - имя таблицы на сервере. которая должна быть прилинкована
'             strLinkName     - имя новой ссылки (прилинкованной таблицы)
'             strDatabase     - опционально база, которую требуется открыть. Если это название задано, то оно
'                               переопределяет базу специфицированную в DSN или необходима в том случае, когда
'                               DSN не содержит имя базы
'             strUID          - опциональное - имя пользователя, необходимо, если DSN не содержит имени (user id)
'             varPWD          - опционально определяет соотв. пользователю пароль, необходимо, если DSN не содержит аккаунта
'                             - если пароль отсутствует - необходима пустая строка
'             varLoginTimeout - опционально время ожидания (timeout) в миллисекундах (по умолчанию используется
'                               значение DAO/Jet)
' Returns   : true в случае удачи
'==========================================================================================================================
Public Function CreateODBCLink(strDSN As String, strSourceTable As String, Optional strLinkName As String = "", _
                                   Optional strDatabase As String, Optional strUID As String, Optional varPWD As Variant, _
                                                                            Optional varLoginTimeout As Variant) As Boolean
  Dim tdfTmp As DAO.TableDef                         ' Прилинкованная таблица
  Dim strConnect As String                           ' Строка доступа к DSN
  
  On Error GoTo PROC_ERR
  
'-------------------------------------------------------------------------
If strLinkName = "" Then strLinkName = strSourceTable
'--------------------------------------------------------------------------
' Строим строку доступа (connect string)
  strConnect = mcstrLinkODBC & "DSN=" & strDSN
  If strDatabase <> "" Then                          ' Если необходимо, добавляем параметр DB
    strConnect = strConnect & ";DATABASE=" & strDatabase
  End If
  If strUID <> "" Then                               ' Если необходимо, добавляем пользователя
    strConnect = strConnect & ";UID=" & strUID
  End If
  If Not IsMissing(varPWD) Then                      ' Если необходимо, добавляем пароль
    strConnect = strConnect & ";PWD=" & varPWD
  End If
  If Not IsMissing(varLoginTimeout) Then             ' Если необходимо, добавляем LoginTimeout
    strConnect = strConnect & ";LOGINTIMEOUT=" & varLoginTimeout
  End If
'-------------------------------------------------------------------------
' Создаем новый объект - таблица и устанавливаем специфическую информацию
  Set tdfTmp = m_dbsCurrent.CreateTableDef(strLinkName)
  tdfTmp.Connect = strConnect
  tdfTmp.SourceTableName = strSourceTable
'-------------------------------------------------------------------------
' Сохраняем прилинкованную таблицу к базе данных
  m_dbsCurrent.TableDefs.Append tdfTmp
  On Error GoTo PROC_ERR
  CreateODBCLink = True  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'-----------------------------------------------------
PROC_EXIT:
  Exit Function
'-----------------------------------------------------
PROC_ERR:
  CreateODBCLink = False '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "CreateODBCLink"
  Resume PROC_EXIT
End Function
'==========================================================================================================================
' Функция удаляет все линкованные таблицы в текущей базе данных. Возвращает True в случае успеха
'==========================================================================================================================
Public Function DeleteAllLinks() As Boolean
  Dim intCounter As Integer
  
  On Error GoTo PROC_ERR
  
  DeleteAllLinks = False
'----------------------------------------------------------------------------------
' Проходим по всем таблицам, если обнаруживаем связанную таблицу - удаляем ее
  For intCounter = m_dbsCurrent.TableDefs.Count - 1 To 0 Step -1
    If m_dbsCurrent.TableDefs(intCounter).Connect <> "" Then
       m_dbsCurrent.TableDefs.Delete m_dbsCurrent.TableDefs(intCounter).Name
    End If
  Next intCounter
  
  DeleteAllLinks = True
'-----------------------------------------------
PROC_EXIT:
  Exit Function
'-----------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "DeleteAllLinks"
  Resume PROC_EXIT
End Function
'==========================================================================================================================
' Функция возвращает путь линкования для заданной таблицы
' PARAMETERS: strTable - имя таблицы, для которой получаем путь
'==========================================================================================================================
Public Function GetLinkedPath(strTable As String) As String
  Dim tdfTmp As DAO.TableDef
  Dim strConnect As String
  Dim intPos As Integer
  
  On Error GoTo PROC_ERR
  
'-------------------------------------------------------------
' Получаем указатель на заданную таблицу
  Set tdfTmp = m_dbsCurrent.TableDefs(strTable)
' Считываем строку коннекта (например для Jet она будет  ";DATABASE=c:\mydata\nwind.mdb")
  strConnect = tdfTmp.Connect

  If strConnect <> "" Then          ' Найденная таблица - линкованная
    intPos = InStr(tdfTmp.Connect, "=")
    GetLinkedPath = Mid$(strConnect, intPos + 1, Len(strConnect)) '!!!!!!!!!!!!!!!!!!!!!!
  End If
'-----------------------------------------------------------------
PROC_EXIT:
  Exit Function
'----------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetLinkedPath"
  Resume PROC_EXIT
End Function
'==========================================================================================================================
' Возвращает тип связи для заданной таблицы
' Parameters: strTable - Name of table to check
' Returns   : Link type
'==========================================================================================================================
Public Function GetLinkedType(strTable As String) As String
  Dim tdfTmp As DAO.TableDef
  Dim strConnect As String
  Dim intPos As Integer
  Dim StrTmp As String

  On Error GoTo PROC_ERR
  
'---------------------------------------------------
' Get a handle to the table
  Set tdfTmp = m_dbsCurrent.TableDefs(strTable)
  ' Получаем заданную связь, например, "  ;DATABASE=c:\mydata\nwind.mdb"
  strConnect = tdfTmp.Connect
  

  If strConnect <> "" Then   ' Если строка не пустая - прилинкованная таблица
    ' Заполняем информацией
    intPos = InStr(tdfTmp.Connect, ";")
    StrTmp = Mid$(strConnect, 1, intPos - 1)
    ' If 'type' is blank, its an Access table
    If StrTmp = "" Then
      StrTmp = "Access"
    End If
  End If
  GetLinkedType = StrTmp
'---------------------------------------------------
PROC_EXIT:
  Exit Function
'---------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetLinkedType"
  Resume PROC_EXIT
End Function
'==========================================================================================================================
' Функция проверяет, является ли таблица линкованной
' Parameters: strTable - Name of table to check
' Returns   : True - Table is linked, False - table is not linked or
'==========================================================================================================================
Public Function IsTableLinked(strTable As String) As Boolean
  Dim tdfTmp As DAO.TableDef
  Dim lngSaveErr As Long
  Dim strSaveErr As String
  Dim fNoObject As Boolean
  Const clngErrObjectNotFound As Long = 3265
  
  On Error GoTo PROC_ERR
'-----------------------------------------------------------------
' Assume the object exists
  fNoObject = False
  ' Get a handle to the table's defintion. Test to ensure the table exists
  On Error Resume Next
  Set tdfTmp = m_dbsCurrent.TableDefs(strTable)
  lngSaveErr = Err.Number
  strSaveErr = Err.Description
  On Error GoTo PROC_ERR
  
  If lngSaveErr <> 0 Then
    If lngSaveErr = clngErrObjectNotFound Then
      IsTableLinked = False
      fNoObject = True
    Else
      Err.Raise 2000, "Error: " & strSaveErr
    End If
  End If

  If Not fNoObject Then
    ' If the connect property isn't blank, its a linked table
    If tdfTmp.Connect <> "" Then
      IsTableLinked = True
    Else
      IsTableLinked = False
    End If
  End If
'---------------------------------------------------------
PROC_EXIT:
  Exit Function
'---------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "IsTableLinked"
  Resume PROC_EXIT
End Function
'==========================================================================================================================
' Функция перелинковывает заданную таблицу по соответсвующему пути (база данных Jet)
' Parameters: strTable - Table name of the linked table
'             strPath - Path name of the database containing the real table
' Returns   : True if successful, False otherwise
'==========================================================================================================================
Public Function ReLinkTableJet(strTable As String, strPath As String) As Boolean
  Dim tdfTmp As DAO.TableDef
  Dim strPrefix As String
  Dim strNewConnect As String
  
  On Error GoTo PROC_ERR
  
  ' Get a handle to the table's defintion
  Set tdfTmp = m_dbsCurrent.TableDefs(strTable)
  
  ' Only try to relink if it is already a linked table
  If tdfTmp.Connect <> "" Then
    ' Build the new connect string
    strPrefix = Left$(tdfTmp.Connect, InStr(tdfTmp.Connect, "="))
    strNewConnect = strPrefix & strPath
    
    ' Relink the linked table
    tdfTmp.Connect = strNewConnect
    tdfTmp.RefreshLink
    
    ReLinkTableJet = True  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  Else
    ReLinkTableJet = False '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  End If
'----------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'----------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "ReLinkTable"
  Resume PROC_EXIT
End Function

'==========================================================================================================================
' Перелинковывает все таблицы Jet к новой базе данных (новый путь)
' Parameters: strNewPath - Full path of the database to link to
' Returns   : True if successful, False otherwise
'==========================================================================================================================
Public Function RelinkTablesJet(strNewPath As String) As Boolean
  Dim tdfTmp As DAO.TableDef
  Dim fOK As Boolean
  
  On Error GoTo PROC_ERR
  
'---------------------------------------------------------------------
  fOK = True
  ' Loop through each table
  For Each tdfTmp In m_dbsCurrent.TableDefs
    ' If it is a linked table
    If tdfTmp.Connect <> "" Then
      ' Relink the table
      fOK = ReLinkTableJet(tdfTmp.Name, strNewPath)
      ' If it fails, fall out of the loop
      If Not fOK Then
        Exit For
      End If
    End If
  Next tdfTmp
'--------------------------------------------------------------------
  RelinkTablesJet = (fOK = True)
'---------------------------------------------
PROC_EXIT:
  Exit Function
'--------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "RelinkTables"
  Resume PROC_EXIT
End Function

'==========================================================================================================================
' Функция тестирует прилинкованную именованную таблицу на предмет валидности связи
' Parameters: strTable - Name of the table to check
' Returns   : True if the table is correctly linked, False otherwise
'==========================================================================================================================
Public Function TestLinkedTable(strTable As String) As Boolean
  Dim tdfTmp As DAO.TableDef
  Dim StrTmp As String
  
  On Error GoTo PROC_ERR
  
  ' Get a handle to the table's definition
  Set tdfTmp = m_dbsCurrent.TableDefs(strTable)
  
  ' Is it a linked table?
  If tdfTmp.Connect <> "" Then
    ' To test, we disable error handling and try to get the name of the linked
    ' table's first field. If the table link is broken, this will generate an
    ' error. We then test for an error to determine if the link is broken.
    On Error Resume Next
    StrTmp = tdfTmp.FIELDS(0).Name
    
    TestLinkedTable = (Err.Number = 0)
    
    ' Re-enable error handling
    On Error GoTo PROC_ERR
  End If
'------------------------------------------
PROC_EXIT:
  Exit Function
'------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "TestLinkedTable"
  Resume PROC_EXIT
End Function
'==========================================================================================================================
' Тестирование всех прилинкованных таблиц на предмет валидности ссылок
' Parameters: None
' Returns   : True - Tables are correctly linked,
'==========================================================================================================================
Public Function TestLinkedTables() As Boolean
  Dim tdfTmp As DAO.TableDef
  Dim varTmp As Variant
  Dim fError As Boolean
  Dim lngErrorNumber As Long
  
  On Error GoTo PROC_ERR
'-------------------------------------------------------------------
  fError = False
  ' Look through each table in the database
  For Each tdfTmp In m_dbsCurrent.TableDefs
    ' See if its a linked table
    If tdfTmp.Connect <> "" Then
      ' Try and get the first field's name. If the table's link is broken,
      ' this will generate a runtime error.
      On Error Resume Next
      varTmp = tdfTmp.FIELDS(0).Name
      lngErrorNumber = Err.Number
      ' Resume Error handling
      On Error GoTo PROC_ERR
      ' Did an error occur?
      fError = (lngErrorNumber <> 0)
      ' Error, so drop out of the loop
      If fError Then
        Exit For
      End If
    End If
  Next tdfTmp
'-------------------------------------------
  TestLinkedTables = Not fError
'------------------------------------------------------------------------
PROC_EXIT:
  Exit Function
'------------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "TestLinkedTables"
  Resume PROC_EXIT
End Function
'==========================================================================================================================
' Генерация HASH для объекта.  Функция генерит HASH следующего формата:
'==========================================================================================================================
Public Function BuildHASH(Optional ObjectType As Integer = 0, Optional Prefix As String = "", _
                                                                                   Optional sName As String = "") As String
     BuildHASH = GetHASH(Prefix & sName)
End Function





















'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
' PRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATE
' PRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATE
' PRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATE
' PRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATE
' PRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATE
' PRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATE
'------------------------------------------------------------------------------------------------------------------------
' Инициализация класса
'------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Initialize()
   Set Me.Database = CurrentDb                             ' При инициализации присваиваем текущую базу данных
End Sub
'------------------------------------------------------------------------------------------------------------------------
' Завершение работы класса
'------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
   On Error GoTo PROC_ERR
  
  ' Если текущая база была открыта классом, то перед выхзодом закрываем эту базу
  If mfClassOpened Then
    If Not m_dbsCurrent Is Nothing Then
      m_dbsCurrent.Close
      Set m_dbsCurrent = Nothing
    End If
  End If
'--------------------------------------------------------------------------
PROC_EXIT:
  Exit Sub
'--------------------------------------------------------------------------
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "Class_Terminate"
  Resume PROC_EXIT
  
End Sub
'-----------------------------------------------------------------------------------------------------------------------
' Функция проверяет вхождение имени таблицы в список таблиц для удаления или наоборот, в список  для исключения.
' если таблица подлежит удалению - возвращается true, иначе false.
' Списки могут содержать как полные перечисления, так и шаблоны вида xxx* (то есть удаление по префиксу)
'    - Если sTableList и sExcludeTable - пустые, сразу возвращается true, так как в этом случае речь идет об удалении
'      всех таблиц
'    - Если sTableList не пустой и наименование таблицы в нем найдено - возвращаем True. Если не пустой и наименование
'      в нем не найдено - возвращаем False
'    - Если таблица найдена в непустом sExcludeTable - сразу возвращаем False
'-----------------------------------------------------------------------------------------------------------------------
Private Function ShouldDeliteThisTable(STABLE As String, sTableList As String, sExcludeTable As String) As Boolean
Dim MyStr As New cString                            ' Манипулятор строк
Dim ArrList() As String                              ' Массив - список найденных значений
Dim iCount As Integer
Dim I As Integer                                     ' Циклическая переменная
'---------------------------------------------------------------------------------------------------------
If sTableList = "" And sExcludeTable = "" Then       ' Случай удаления всех таблиц
   ShouldDeliteThisTable = True                      ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
ElseIf sExcludeTable <> "" Then                      ' Удаляем все таблицы, кроме входящих в sExcludeTable
   ' Проверяем, не входит ли sTable в sExcludeTable
   If Not MyStr.IsWordInList(sExcludeTable, STABLE) Then ShouldDeliteThisTable = True  ' !!!!!!!!!!!! (удаляем)
Else                                                 ' Удаляем все таблицы. содержащиеся в списке sTableList
   If MyStr.IsWordInList(sTableList, STABLE) Then ShouldDeliteThisTable = True ' !!!!!!!!!!!! (удаляем)
End If
End Function

'-------------------------------------------------------------------------------------------------------------------------
' Функция возвращает строку в квадратных скобках (для поля)
'-------------------------------------------------------------------------------------------------------------------------
Private Function SHT(sSTR As String) As String
Dim sWork As String
sWork = Trim(sSTR)
    If Left(sWork, 1) <> "[" Then
            SHT = "[" + sWork + "]"    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Else
            SHT = sWork                '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    End If
End Function
'------------------------------------------------------------------------------------------------------------------------------------
' Функция возвращает закавыченную для SQL строку для строковых значений,
' обычную строку для числовых значений (например ID = 3, а не ID ='3')
' и для дат - осуществляется форматирование дат в соответствии с форматом Jet
' Функция также производит распознавание GUID'ов
'----------------------------------------------------------------------------------------------------------------------------------
Private Function SH(sSTR As String) As String
Dim sRes As String, sWork As String                       ' Возвращаемый результат и рабочаЯ строка
      sWork = RemoveSH(sSTR)            ' Проверяем на уже включенные кавычки
      '------------------------------------------------------------------------------------------------
      If UCase(sWork) = "TRUE" Or UCase(sWork) = "FALSE" Then ' Выражение является булевским значением
         sRes = sWork
         GoTo sH_EXIT
      End If
      '-----------------------------------------------------------------------------------------------
      sRes = SQLDate(sWork)             ' Проверяем на дату
      '-----------------------------------------------------------------------------------------------
      If sRes = "" Then                 ' Это была не дата, проверяем дальше
      '-----------------------------------------------------------------------------------------------
        If IsNumeric(sWork) And Left(sSTR, 1) <> "'" Then              ' Выражени является числом
         sRes = sWork
      '----------------------------------------------------------------------------------------------
        ElseIf IsGuid(sWork) Then                ' Выражение является GUID'ом
         sRes = SQLGUID(sWork)
       '---------------------------------------------------------------------------------------------
        Else                                    ' Подразумеваем, что иначе - строка, закавычиваем
         sRes = "'" + sWork + "'"
        End If
      End If
'------------------------------------------------------------------------------
sH_EXIT:
     SH = sRes  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'-------------------------------------------------------------------------------------------------------------------------
' Функция нормализует строку, для еe использования в SQL -выражениях
' Если строка имеет вид KEY = VALUE => MYPARAM = VALUE         -->   [MYPARAM]='VALUE'
' Если строка имеет вид KEY1=VALUE1;KEY2=VALUE2;...=>
'       [MYPARAM1]='VALUE1' AND [MYPARAM2]='VALUE2' AND ...
' Если строка имеет другой синтаксис SQL, например, "Surname Is Not Null" то преобразование возвращает
' исходную строку
'-------------------------------------------------------------------------------------------------------------------------
Private Function NormKV(KVPar As String) As String
Dim inEQPos As Integer                         ' Позиция знака равенства
Dim sLeft As String, sRight As String
Dim myArray() As String, nMyArray As Integer   ' Строковый массив и его размерность
Dim I As Integer                               ' Циклическая переменная
Dim sWork As String                            ' Рабочая строка
'-------------------------------------------
' Инициализируем функцию
    NormKV = KVPar
'--------------------------------------------------------------------------------------------
' Определяем вид преобразования: если вид выражения A=B;C=D;...., пробуем разбить на пары (A=B) и (C=D)
Select Case GetKVPattern(KVPar)
'-------------------------------------------------------------------------------------
  Case "=":                            ' Одиночная пара KEY=VALUE
        ' Приводим выражение
                sLeft = GetKeyFromKV(KVPar): sRight = GetValueFromKV(KVPar)
                If sLeft <> "" And sRight <> "" Then
                    NormKV = sLeft + " = " + sRight
                End If
'-------------------------------------------------------------------------------------
  Case ";":                            ' Список пар KEY1=VALUE1;KEY2=VALUE2
        ' Разбиваем на пары, для каждой осуществл. преобразование
          myArray = Split(KVPar, ";"): nMyArray = UBound(myArray)
          sWork = ""
          For I = 1 To nMyArray
                sLeft = GetKeyFromKV(myArray(I)): sRight = GetValueFromKV(myArray(I))
                If sLeft <> "" And sRight <> "" Then
                    If sWork <> "" Then sWork = sWork & " And "
                    sWork = sWork + sLeft + " = " + sRight
                End If
          Next I
          NormKV = sWork
'-------------------------------------------------------------------------------------
  Case Else                            ' Другое выражение
End Select
'-------------------------------------------
End Function
'--------------------------------------------------------------------------------------------------------------------------
' Функция удаляет окаймляющие кавычки, если они имеются
'--------------------------------------------------------------------------------------------------------------------------
Private Function RemoveSH(sSTR As String) As String
Dim sRes As String
Dim sLeft As String, sRight As String

If sSTR = "" Then Exit Function
sRes = sSTR: sLeft = "'"
    Do While sLeft = "'"
        sLeft = Left(sRes, 1): sRight = Right(sRes, 1)
        If (sLeft = "'" And sRight = "'") Or (sLeft = Chr(34) And sRight = Chr(34)) Then
            sRes = Mid(sRes, 2, Len(sRes) - 2)
        End If
    Loop
'--------------------------------------------------------------------------------
    RemoveSH = sRes '!!!!!!!!!!!!!!!
End Function
'---------------------------------------------------------------------------------------------------------------------------
' Функция возвращает последний ID. Для этих целей используется несколько методов, включая заданный критерий
'---------------------------------------------------------------------------------------------------------------------------
Private Function GetLastIdentity(ByVal TableName As String, Optional sCriteria As String = "", _
                                                                            Optional TimeStampCode As String = "") As Variant
Dim iRes As Variant                          ' Возвращаемый результат
Dim RS As DAO.Recordset                      ' Спец рекордсет
Dim sSQL As String                           ' Строка запроса
Dim sTableName As String                     ' Наименование таблицы
Dim iLeft As Long, iRight As Long            ' Указатели выделения
sSQL = ""
'------------------------------------------------------
' Проверяем, не задано ли условие "[ID]=..."
   If InStr(1, sCriteria, "[ID]") > 0 Then
     iRes = GetValueForKey(sCriteria, "[ID]")
  ElseIf InStr(1, sCriteria, "ID") > 0 Then
     iRes = GetValueForKey(sCriteria, "ID")
  End If
     If iRes <> "" Then GoTo GetLastIdentity_EXIT
'-------------------------------------------------------------------------------------------------------------------------
' ДАЛЬНЕЙШЕЕ ЕСТЬ ВЕРОЯТНОСТНЫЕ МАНИПУЛЯЦИИ
If Not IsFieldExists("ID", TableName) Then Exit Function                      ' Функция работает только для поля с наименованием ID
If IsFieldExists("TimeStampCode", TableName) And TimeStampCode <> "" Then     ' Существует метка возрастающего времени,
     sSQL = "SELECT [ID] FROM " & SHT(TableName) & _
     " Where [TimeStampCode]='" & TimeStampCode & "';"
ElseIf IsFieldExists("DateCreate", TableName) Then                            ' Задано поле DateCreate (меньшая точность)
     ' sSQL = IIf(sCriteria <> "", " WHERE " & NormKV(sCriteria), "")
     sSQL = "SELECT [ID] FROM " & SHT(TableName) & sSQL & " ORDER BY " & SHT(TableName) & ".DateCreate" & ";"
ElseIf IsFieldExists("CreateDate", TableName) Then                            ' Задано поле DateCreate (меньшая точность)
     ' sSQL = IIf(sCriteria <> "", " WHERE " & NormKV(sCriteria), "")
     sSQL = "SELECT [ID] FROM " & SHT(TableName) & sSQL & " ORDER BY " & SHT(TableName) & ".CreateDate" & ";"
Else                                                                              ' нет надежного способа :(
     Exit Function
End If
'-------------------------------------------------------------------------------------------------------------------------
' Выполняем запрос
    Set RS = m_dbsCurrent.OpenRecordset(sSQL)
    RS.MoveFirst: RS.MoveLast
    iRes = RS.FIELDS(0).value               ' конечное значение
    RS.Close
If iRes <> 0 Then GoTo GetLastIdentity_EXIT
'------------------------------------------------------
' ОБЩИЙ СПОСОБ
sSQL = "SELECT @@IDENTITY AS LastID;"
Set RS = m_dbsCurrent.OpenRecordset(sSQL)
    iRes = RS!LastID
RS.Close
If iRes <> 0 Then GoTo GetLastIdentity_EXIT     ' Получили ID самым простым способом. Но эта функция Jet не очень надежна
'-------------------------------------------------------
GetLastIdentity_EXIT:
    GetLastIdentity = iRes   '!!!!!!!!!!!!!!!!!!
    Set RS = Nothing
    Exit Function
'------------------------------
GetLastIdentity_ERR:
    MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "GetLastIdentity ERROR"
    Err.Clear
    Resume GetLastIdentity_EXIT
End Function
'---------------------------------------------------------------------------------------------------------------------------
' Функция возвращает массив для DAO.RECORDSET, такой что массив ARRAY (I,J) cодержит значения
' При I = 1, Recordset.RECORDSCOUNT, I = J, Recordset.FIELDSSCOUNT
' а в первой строке I=0 содержатся названия столбцов.
' Функция поддерживает многозначные поля для MS Access 2007, сериализуя их в строки
'---------------------------------------------------------------------------------------------------------------------------
Private Function GetArrrayFromRS(RS As DAO.Recordset, _
                                                                    Optional bWithColumnNames As Boolean = True) As Variant
Dim myArray() As String                                     ' Исходный массив
Dim nRecords As Long                                        ' Кол-во записей
Dim nFields As Integer                                      ' Кол-во  полей
Dim J As Integer, I As Long                                 ' Циклические переменные
Dim rsMVF As DAO.Recordset                                  ' Подчиненный запрос для многозначных пролей

Dim strOut As String            ' Выходная строка  (multi-value field)
Dim lngLen As Long              ' длина строки
Const strcSep = ";"             ' Сепаратор для списка(multi-value list)

On Error GoTo GetArrrayFromRS_ERR

'-----------------------------------------------------------------------------------------------
' Инициализируем функцию
GetArrrayFromRS = Empty
'----------------------------------------------------------------------------------------------
' Проверяем, что Recordset не нулевой
            If RS.EOF = True And RS.BOF = True Then           ' Пустой массив
                GoTo GetArrrayFromRS_ERR
            End If
'-------------------------------------------------------------------------------------------------
' Выбираем метод получения массива
If bWithColumnNames Then                ' Переводим массив в ручную, включая имена столбцов
    '-----------------------------------------------------
    ' Обновляем Recordset
    RS.MoveLast: RS.MoveFirst
            nFields = RS.FIELDS.Count                ' посчитали кол-во столбцов
            nRecords = RS.RecordCount                ' посчитали кол-во строк
            ReDim myArray(nRecords, nFields)         ' инициализировали массив
     '-----------------------------------------------------
     ' Считываем имена строк и значения, проходя в цикле
     I = 0
     Do While Not RS.EOF
        For J = 1 To nFields
            If I = 0 Then                    ' Считываем заголовки
                    myArray(I, J) = Nz(RS.FIELDS(J - 1).Name, "")
            Else                             ' Считываем значения
                   '-------------------------------------------------------------------------------------
                   If varType(RS.FIELDS(J - 1).value) = vbObject Then     ' Обнаружили многозначное поле
                              Set rsMVF = RS.FIELDS(J - 1).value          ' Подчиненный запрос
                                  rsMVF.MoveLast: rsMVF.MoveFirst
                              Do While Not rsMVF.EOF
                                      If RS.FIELDS(J - 1).Type = 101 Then      'dbAttachment
                                          strOut = strOut & rsMVF!fileName & strcSep
                                       Else
                                          strOut = strOut & rsMVF![value].value & strcSep
                                       End If
                                   rsMVF.MoveNext
                               Loop
                                   lngLen = Len(strOut) - Len(strcSep) ' Удаляем лишний сепаратор
                               If lngLen > 0& Then
                                   myArray(I, J) = Left(strOut, lngLen)
                               End If
                                   Set rsMVF = Nothing
                    '--------------------------------------------------------------------------
                     Else                                 ' Обычное  поле, не multi-value
                          myArray(I, J) = Nz(RS.FIELDS(J - 1).value, "")
                    '--------------------------------------------------------------------------
                     End If
            RS.MoveNext         ' Продвигаемся вперед
            End If
            I = I + 1           ' ПРИРАЩЕНИЕ по i
            
        Next J
      Loop
'----------------------------------------------------------------------------------------
Else                                    ' Переводим стандартно и не заботимся о именах и типах полей
    myArray = RS.GetRows                '(CountOfReords = UBound(myArray, 2) + 1)
End If
'--------------------------------
' Завершаем функцию
GetArrrayFromRS = myArray
Exit Function
'-----------------------------------------------------------------------------
' Обрабатываем ошибки
GetArrrayFromRS_ERR:
  If Err <> 0 Then MsgBox "GetArrrayFromRS_ERR: " & Err.Number & vbCrLf & Err.Description
End Function

























'============================================================================================================
' Функция возвращает значение индекс поля в массиве, если он там присутсвует и 0 в противном случае
'============================================================================================================
Private Function FFIELDSIndexOf(sName) As Integer
Dim I As Integer                                      ' циклическая переменная
Dim FieldCount As Integer                             ' текущая размерность массива
    FieldCount = UBound(FFIELDS)
    If FieldCount > 0 Then                            ' есть что просматривать
        For I = 1 To FieldCount
           If FFIELDS(I).Fname = sName Then
                FFIELDSIndexOf = I
                Exit Function
           End If
        Next I
    Else                                              ' просматривать нечего, выходим
        FFIELDSIndexOf = 0
    End If
End Function
'============================================================================================================
' Функция проверяет, существует ли PK в массиве
'============================================================================================================
Private Function IsExistPKinFFIELDS() As Boolean
Dim I As Integer                                      ' циклическая переменная
Dim FieldCount As Integer                             ' текущая размерность массива

'------------------------------------------------------------------
' Значение по умолчанию
    IsExistPKinFFIELDS = False '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        FieldCount = UBound(FFIELDS)       ' ОПРЕДЕЛЯЕМ границы
        If FieldCount > 0 Then                        ' есть что просматривать
            For I = 1 To FieldCount
                If FFIELDS(I).FisPK Then   ' ОБНАРУЖИЛИ PK
                    IsExistPKinFFIELDS = True   '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                    Exit Function
                End If
            Next I
        End If
End Function

'============================================================================================================
' Функция формирует из строки KEY=VALUE и имени таблицы SQL запрос на вставку значения INSERT INTO
'============================================================================================================
Private Function GetInsertIntoStatement(TableName As String, KVString As String) As String
Dim KVPairs() As String                             ' Массив пар KEY=VALUE
Dim I As Integer, KVBound As Integer                ' Циклическая переменная и верхняя граница
Dim strCOLS As String, strVALS As String            ' Строка названий столбцов и названий строк
On Error GoTo GetInsertIntoStatement_ERR
'-------------------------------------------------------------------------------------
' Парсим и проверяем значимость входной строки
If KVString <> "" Then
    If InStr(1, KVString, ";") = 0 Then ' Присутсвует только одна пара
        ReDim KVPairs(0)
        KVPairs(0) = KVString
    Else
        KVPairs = Split(KVString, ";")
        KVBound = UBound(KVPairs)
        If KVBound < 1 Then GoTo GetInsertIntoStatement_ERR
    End If
Else
    GoTo GetInsertIntoStatement_ERR
End If
'-------------------------------------------------------------------------------------
' Из массива строк вида KEY=VALUE формируем  SQL подобную строку
      ' Нулевая строка
      strCOLS = GetKeyFromKV(KVPairs(0))
      strVALS = GetValueFromKV(KVPairs(0))
      ' Строка, начиная с первой (подразумевеатся, что хотя бы два столбца обновляется)
      If UBound(KVPairs) > 0 Then
            For I = 1 To KVBound
                strCOLS = strCOLS + "," + GetKeyFromKV(KVPairs(I))
                strVALS = strVALS + "," + GetValueFromKV(KVPairs(I))
            Next I
      End If
'-------------------------------------------------------------------------------------
' Формируем выходную строчку SQL
    GetInsertIntoStatement = "INSERT INTO " + SHT(TableName) + " (" + strCOLS + ") VALUES(" + strVALS + ");"
    Exit Function
'--------------------------------------------------------------
' Обрабатываем ошибку
GetInsertIntoStatement_ERR:
    GetInsertIntoStatement = ""
End Function

'============================================================================================================
' Функция формирует из строки KEY=VALUE и строки-критерия SUPKEY=SUPVALUE и имени таблицы SQL запрос
' на обновление Update
'============================================================================================================
Private Function GetUpdateStatement(TableName As String, KVString As String, Criteria As String) As String
Dim KVPairs() As String                             ' Массив пар KEY=VALUE
Dim I As Integer, KVBound As Integer                ' Циклическая переменная и верхняя граница
Dim strUpdate As String                             ' основная строка обновления
Dim strKriteria As String                           ' Нормализованная строка критериев
On Error GoTo GetUpdateStatement_ERR
'-------------------------------------------------------------------------------------
' Парсим и проверяем значимость входной строки
If KVString <> "" Then
    KVPairs = Split(KVString, ";")
    KVBound = UBound(KVPairs)
    ' If KVBound < 1 Then GoTo GetUpdateStatement_ERR
Else
    GoTo GetUpdateStatement_ERR
End If
'-------------------------------------------------------------------------------------
' Из массива строк вида KEY=VALUE формируем  SQL подобную строку
      ' Нулевая строка
      strUpdate = NormKV(KVPairs(0)) ' Для первой строки не нужны запятые
      ' Строка, начиная с первой (подразумевеатся, что хотя бы два столбца обновляется)
    For I = 1 To KVBound
      strUpdate = strUpdate + "," + NormKV(KVPairs(I))
    Next I
'-------------------------------------------------------------------------------------
' Нормализуем строку критериев
      strKriteria = NormKV(Criteria)
'-------------------------------------------------------------------------------------
' Формируем выходную строчку SQL
    GetUpdateStatement = "UPDATE " + SHT(TableName) + " SET " + strUpdate + " WHERE " + strKriteria + ";"
    Exit Function
'--------------------------------------------------------------
' Обрабатываем ошибку
GetUpdateStatement_ERR:
    GetUpdateStatement = ""
End Function
'=======================================================================================================
' Функция выделяет из строки KEY=VALUE ее ключ -- KEY и нормализует его для использования в SQL
' MYPARAM = VALUE         -->   [MYPARAM]
'=======================================================================================================
Private Function GetKeyFromKV(KVPar As String) As String
Dim inEQPos As Integer                         ' Позиция знака равенства
'-------------------------------------------
' Инициализируем функцию
    GetKeyFromKV = ""
'-------------------------------------------
' Приводим выражение
inEQPos = InStr(1, KVPar, "=") ' позиция знака равенства
    If inEQPos > 0 Then
      GetKeyFromKV = SHT(Left$(KVPar, inEQPos - 1))
    End If
End Function

'=======================================================================================================
' Функция выделяет из строки KEY=VALUE ее значение -- VALUE и нормализует его для использования в SQL
' MYPARAM = VALUE         -->   'VALUE'
'=======================================================================================================
Private Function GetValueFromKV(KVPar As String) As String
Dim inEQPos As Integer                         ' Позиция знака равенства
Dim sRes As String                             ' Возвращаемый результат
'-------------------------------------------
' Инициализируем функцию
    GetValueFromKV = ""
    sRes = ""
'-------------------------------------------
' Приводим выражение
inEQPos = InStr(1, KVPar, "=") ' позиция знака равенства
    If inEQPos > 0 Then
      sRes = Trim(Right$(KVPar, Len(KVPar) - inEQPos)) ' ПОЛУЧИЛИ САМО ВЫРАЖЕНИЕ
      '----------------------------------------------------------------------------------------------
      sRes = SH(sRes)       '///////////////////////////////////////////////////////////////////////
      '-----------------------------------------------------------------------------------------------
    End If
'--------------------------------------------------------------
    GetValueFromKV = sRes  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function



'=======================================================================================================
' Функция  выявляет паттерн в строке KV-String:
'  (1) KEY=VALUE     - возвращается "="
'  (2) KEY1=VALUE1;KEY2=VALUE2;....   - возвращается ";" или другой разделитель, если присутствует общий
'                                      вид паттерна, например "KEY1=VALUE1,KEY2=VALUE2" --> ","
'                                      разделителем не может быть "=" или ":"
'  (3) паттерн не выявлен (другой вид строки) - возвращается ""
'      Обязательным является использование в качестве разделителя KEY/VALUE "="
'=======================================================================================================
Private Function GetKVPattern(KVS As String) As String
Dim EqvCount As Long                  ' Кол-во вхождений знака "="
Dim SemColCount As Long               ' Кол-во вхождений знака ";"
'-------------------------------------------------------
' Инициализируем функцию
  GetKVPattern = ""
'-------------------------------------------------------
' Проверяем не входит ли в критерий ключевое слово SQL, в этом случае = > паттерн (3) сразу
   If IsStopWordInStr(KVS) Then Exit Function
'-------------------------------------------------------
' Подсчитываем вхождения ключевых разделителей
   EqvCount = InStrCount(1, KVS, "=")
   Select Case EqvCount
   Case 0:                       ' Разделитель "=" не найден, неизвестный паттен
       Exit Function
   Case 1:                       ' Присутствует всего одна пара
       GetKVPattern = "="
   Case Else                     ' Скорее всего присутствует несколько пар
       SemColCount = InStrCount(1, KVS, ";")
       If SemColCount >= 1 Then  ' Определили список
         GetKVPattern = ";"
       Else                      ' Сбой
         GetKVPattern = ""
       End If
   End Select
   
End Function

'=======================================================================================================
' Функция подсчитывает кол-во вхождений одной подстроки в другую (без создания массива с помощью Split)
' Подсчет начинается с указателя lStart
'=======================================================================================================
Private Function InStrCount(lStart As Long, sString1 As String, sString2 As String) As Long
Dim lPos As Long
lPos = InStr(lStart + lPos, sString1, sString2)
While lPos <> 0
   lPos = InStr(lStart + lPos + Len(sString2) - 1, sString1, sString2)
   InStrCount = InStrCount + 1
Wend
End Function

'=======================================================================================================
' Функция проверяет, что во входящей строке нет стоп-слова из списка. если Stop-Word обнаружено,
' возвращается True, иначе False. Слова должны присутствовать отдельно
'=======================================================================================================
Private Function IsStopWordInStr(str1 As String) As Boolean
Dim myList() As String                          ' Массив стоп-слов
Dim nMyList As Integer                          ' Размерность массива
Dim I As Integer                                ' Циклическая переменная
Const strStopWords = "SELECT;WHERE;LIKE;BY;ORDER;OR;AND;UPDATE"
'-----------------------------------------
' Получаем массив из строки
   myList = Split(strStopWords, ";"): nMyList = UBound(myList)
'-----------------------------------------
' предполагаем, что ключевое слово имеет пробел сначала и пробел или открывающую скобку в конце
' Проверяем по списку
   For I = 0 To nMyList
       
       If (InStr(1, str1, " " & myList(I) & " ") > 0) Or _
                        (InStr(1, str1, " " & myList(I) & "(") > 0) Then
            IsStopWordInStr = True
            Exit Function
       End If
   Next I
End Function
'----------------------------------------------------------------------------------------------------------------------------------
' Функция форматирует дату для использование в SQL запросах
' iFormat = 0 - полная дата (до секунд), 1 - средняя дата  (до минут), 2 - короткая дата (только день)
'----------------------------------------------------------------------------------------------------------------------------------
Private Function SQLDate(sValue As Variant, Optional iFormat As Integer = 0) As String
Dim sDate As String                       ' Возвращаемый формат

'---------------------------------------------------------------------
' Анализируем, передана ли дата
If Left(sValue, 1) = "#" And Right(sValue, 1) = "#" Then        ' Если уже заданы окаймляющие # - убираем их
       sDate = Mid(sValue, 2, Len(sValue) - 2)
Else
       sDate = sValue
End If
If Not IsDate(sDate) Then
    sDate = "": GoTo SQLDate_EXIT
End If
'-------------------------------------------------
Select Case iFormat
    Case 0: sDate = Format(Now, "MM/dd/yyyy HH:mm:ss")     ' полная дата
    Case 1: sDate = Format(Now, "MM/dd/yyyy HH:mm")        ' дата до минут
    Case 2: sDate = Format(Now, "MM/dd/yyyy")              ' дата до дней
    Case Else: sDate = Format(Now, "MM/dd/yyyy HH:mm:ss")  ' полная дата
End Select
'------------------------------------------------
    Mid$(sDate, 3) = "/" ' меняем разделители
    Mid$(sDate, 6) = "/"
    sDate = "#" & sDate & "#"
'-----------------------------------------------------------
SQLDate_EXIT:
    SQLDate = sDate  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'==========================================================================================================================
' Обработка GUID для представления в SQL
'==========================================================================================================================
Public Function SQLGUID(vVal As Variant) As String
Dim sRes As String                                       ' Возвращаемый результат
Dim sWork As String                                      ' Рабочая строка
'sRes = StringFromGUID(vVal)                             ' Осуществляем стандартное преобразование
 sRes = CStr(vVal)                                       ' Сначал для уверенности преобразовываем в строку
'------------------------------------------------------------------------------------------------
' Подразумеваем, что строка точно GUID
    sWork = RemoveSH(sRes)                                                  ' Удаляем все кавычки, если они имеются
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
'--------------------------------------------------------------------------------------------------------------------------
' Проверяем, не является ли строка GUID'ом
' Признаком является или приставка {guid, или окаймление {} , в которое вписано выражение вида
'                                                       00000000-0000-0000-0000-000000000000
' Не используется вариант с регулярными выражениями, чттобы не привлекать отдельную библиотеку
'--------------------------------------------------------------------------------------------------------------------------
Public Function IsGuid(vVal As Variant) As Boolean
Dim sWork As String                                    ' Рабочая строка
Dim bRes As Boolean                                    ' Возвращаемый результат
Dim iPos As Long                                       ' Позиция в строке

    sWork = StringFromGUID(vVal)                       ' Сразу преобразуем в строку. Усли конфликт типов - сразу выходим
    iPos = InStr(1, sWork, "{guid ")
    If iPos > 0 And iPos < 3 Then             ' Если нашли необходимую приставку - распознали сразу
       bRes = True
    ElseIf Left(sWork, 1) = "{" And Right(sWork, 1) = "}" And Len(sWork) = 38 Then  ' Всего со скобками 34 символа
       bRes = True
    End If
'---------------------------------------------
IsGUID_EXIT:
    IsGuid = bRes '!!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'---------------------------------------------
IsGUID_ERR:
    Err.Clear
    Resume IsGUID_EXIT
End Function
'---------------------------------------------------------------------------------------------------------------------------------
' Генерация случайной строки заданной длины
'---------------------------------------------------------------------------------------------------------------------------------
Private Function RandomString(LenOfString As Integer) As String

    Randomize
    Dim rgch As String
    rgch = "abcdefghijklmnopqrstuvwxyz"
    rgch = rgch & UCase(rgch) & "0123456789"

    Dim I As Long
    For I = 1 To LenOfString
        RandomString = RandomString & Mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
    Next

End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------
' Функция вычисляет значение ключа из KVString
'--------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetValueForKey(sKV As String, sKey As String) As String
Dim KVS As New cKVString
           KVS.KVString = sKV
           '----------------------------------
           GetValueForKey = KVS.GetValueOf(sKey)  '!!!!!!!!!!!!!!!!!!!!!!
'---------------------------------------------
Set KVS = Nothing
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
' Функция пробует найти последний ID по дате (поля CreateDate)
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetLastIDFromCreateDate(sTableName As String) As Variant
Dim vRes As Variant
Dim sSQL As String, RS As DAO.Recordset

On Error GoTo GetLastIDFromCreateDate_ERR
'-----------------------------------------------------------------------
sSQL = "SELECT top 1 " & SHT(sTableName) & ".[ID], " & SHT(sTableName) & _
        ".CreateDate FROM " & SHT(sTableName) & " ORDER BY " & SHT(sTableName) & ".CreateDate DESC;"
'-----------------------------------------------------------------------
If IsFieldExists("CreateDate", sTableName) Then
    Set RS = m_dbsCurrent.OpenRecordset(sSQL)
    If Not RS.EOF Then
        RS.MoveLast: RS.MoveFirst
        vRes = RS![id]
    Else
        vRes = -1
    End If
End If
'------------------------------------------------------------------------
GetLastIDFromCreateDate_EXIT:
        GetLastIDFromCreateDate = vRes '!!!!!!!!!!!!!!
        Exit Function
'-----------------------------------------------
GetLastIDFromCreateDate_ERR:
        Err.Clear
        Resume GetLastIDFromCreateDate_EXIT
End Function



