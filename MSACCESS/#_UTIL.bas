Attribute VB_Name = "#_UTIL"
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
'                 $$$$$$$$$$F                                       ## ##  ##  ##  ##    ##   ##   ##  ######  ######  ##
'                  *$$$$$$$$"                                        ####  #####  ##     ##   ##   ##    ##      ##    ##
'                    "***""               _____________                                       ##   ##    ##      ##    ##
' STANDARD MODULE WITH DEFAULT FUNCTIONS |v 2017/03/19 |                                      ##   ##    ##      ##    ##
' The module contains frequently used functions and is part of the G-VBA library               #####     ##    ######  #####
'****************************************************************************************************************************************************
'****************************************************************************************************************************************************
Option Explicit

Private Const MOD_NAME = "#_UTIL"

#If Win64 Then
       Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" _
            (ByVal vKey As Long) As Integer
#Else
       Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If


Public Enum LogCommand
        LOG_AddMsgToLog = 0
        LOG_SaveMsgToFile = 1
        LOG_SaveMsgToLogTable = 2
        LOG_OpenLastLogFile = 3
        LOG_ShowLogTable = 4
        LOG_ClearLogTable = 5
End Enum

' VIRTUAL KEY CODES
Private Const VK_TAB = 9
Private Const VK_SHIFT = 16
Private Const VK_CONTROL = 17
Private Const VK_Alt = 18
Private Const VK_ESCAPE = 27
Private Const VK_LEFT = 37
Private Const VK_UP = 38
Private Const VK_RIGHT = 39
Private Const VK_DOWN = 40
Private Const VK_F1 = 112


'****************************************************************************************************************************************************


'====================================================================================================================================================
' GET ERR MUMBER
'====================================================================================================================================================
Public Function GetErrID(ModName As String, FuncName As String, Optional MsgID = 0, Optional LangID As Integer = 0, _
                                                                                                       Optional RndNum As Boolean = True) As Long
Dim ErrID As Long, sWhere As String
Const MSGTBL As String = "$$MSGS"

On Error GoTo ErrHandle
'-----------------------
If Not IsEntityExist(MSGTBL) Then
      Call CreateMsgTable
      Exit Function
Else
      sWhere = "(ModName = " & sCH(ModName) & ") AND (FuncName = " & sCH(FuncName) & ")"
      ErrID = Nz(DLookup("ErrID", MSGTBL, sWhere), 0)
      If ErrID = 0 And RndNum Then
         ErrID = Int((65534 - 29999) * Rnd + 29999)
      End If
End If
'-----------------------
ExitHere:
    GetErrID = ErrID '!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint "GetErrID", Err.Number, Err.Description
    Err.Clear
End Function
'====================================================================================================================================================
' GET ERR DESCRIPTION
'====================================================================================================================================================
Public Function GetErrDescription(ModName As String, FuncName As String, Optional ErrID As Long = 0, Optional LangID As Integer = 0) As String
Dim sMsg As String
Const MSGTBL As String = "$$MSGS"

On Error GoTo ErrHandle
'-----------------------
    If ErrID > 0 Then
        sMsg = Nz(DLookup("Description", MSGTBL, "(ErrID =" & ErrID & ") AND (LangID = " & LangID & ")"), "")
    ElseIf (ModName <> "") And (FuncName <> "") Then
        sMsg = GetMsgs(ModName, FuncName, , LangID)
    End If
    
    If sMsg = "" Then sMsg = "Unlnow Error in function " & FuncName & " in " & ModName & " module"
'-----------------------
ExitHere:
    GetErrDescription = sMsg '!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint "GetErrDescription", Err.Number, Err.Description
    Err.Clear
End Function
'====================================================================================================================================================
' GET MESSAGES
'====================================================================================================================================================
Public Function GetMsgs(ModName As String, FuncName As String, Optional MsgID As Integer = 0, _
                                                                                  Optional LangID As Integer = 0, Optional IDD As Long = 0) As String
Dim sMsg As String, sWhere As String

Const MSGTBL As String = "$$MSGS"

On Error GoTo ErrHandle
'------------------------------
If Not IsEntityExist(MSGTBL) Then
      Call CreateMsgTable
      Exit Function
Else
      If IDD > 0 Then ' Looking by ID
          sMsg = Nz(DLookup("Description", MSGTBL, "ID = " & IDD), "")
      ElseIf (ModName <> "") And (FuncName <> "") Then
          sWhere = "(LangID = " & LangID & ") AND (ModName = " & sCH(ModName) & ") AND (FuncName = " & sCH(FuncName) & ")"
          sMsg = Nz(DLookup("Description", MSGTBL, sWhere), "")
      End If
End If
'------------------------------
ExitHere:
   GetMsgs = sMsg '!!!!!!!!!!!!!!
   Exit Function
'----------
ErrHandle:
    ErrPrint "GetMsgs", Err.Number, Err.Description
    Err.Clear
End Function
'======================================================================================================================================================
' Check if Debug.Mode
'======================================================================================================================================================
Public Function IsDebug() As Boolean
#If DEBUGMODE Then
    IsDebug = True '!!!!!!!!!!!!!
#Else
    IsDebug = False '!!!!!!!!!!!!
#End If
End Function

'======================================================================================================================================================
' Write Log File or Read It
'======================================================================================================================================================
Public Function LogLog(Optional iLogCommand As LogCommand = LOG_OpenLastLogFile, Optional sLog As String, Optional sMsg As String, _
                                                                      Optional sFuncName As String, Optional sContext As String, Optional id As Long, _
                                                                              Optional DLM As String = ";", Optional SEP As String = vbCrLf) As String
Dim sRes As String, sExt As String, sFolder As String, bRes As Boolean

Const LOG_FOLDER As String = "LOG"
Const LOG_TBL As String = "$$LOGS"
Const LOG_LOCAL_PARAM As String = "LASTLOGFILE"

    On Error GoTo ErrHandle
'-------------------------------------
Select Case iLogCommand
Case LOG_AddMsgToLog:           ' = 0
    If sMsg = "" Then
         sRes = sLog
    Else
         If sLog <> "" Then sLog = sLog & SEP
         sExt = sMsg & DLM & sFuncName & DLM & sContext & DLM & id
         If IsDebug Then Debug.Print sExt
         sRes = sLog & sExt
    End If
Case LOG_SaveMsgToFile:         ' = 1
    If sLog = "" Then Exit Function
    sExt = Format(Now(), "yyyymmdd") & Timer * 100 & "_" & FilenameWithoutExtension(CurrentProject.Name) & ".log"
    sFolder = CurrentProject.Path & "\" & LOG_FOLDER & "\"
    If Dir(sFolder, vbDirectory) = "" Then MkDir sFolder
    sRes = sFolder & sExt
    If Not WriteStringToFileUTF8(sLog, sRes) Then sRes = ""
    If Dir(sRes) <> "" Then Call SetLocal(LOG_LOCAL_PARAM, sRes, "Last Log File)")
    
Case LOG_SaveMsgToLogTable:     ' = 2
     If Not WriteLogToTBL(sLog, LOG_TBL, DLM, SEP) Then sRes = "The Log is saved to " & LOG_TBL
Case LOG_OpenLastLogFile:       ' = 3
     sExt = GetLocal(LOG_LOCAL_PARAM)
     If sExt <> "" Then
        If Dir(sExt) <> "" Then
            Call Shell("notepad.exe " & sExt, vbNormalFocus)
            sRes = sExt
        End If
     End If
Case LOG_ShowLogTable:          ' = 4
     If IsTable(LOG_TBL) Then
        sRes = LOG_TBL
        DoCmd.OpenTable (LOG_TBL)
     End If
Case LOG_ClearLogTable:         ' = 5
     If ClearTable(LOG_TBL) Then sRes = LOG_TBL
Case Else:
     sRes = ""
End Select
'-------------------------------------
ExitHere:
    LogLog = sRes '!!!!!!!!
    Exit Function
'----------------
ErrHandle:
    ErrPrint2 "LogLog", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------
' Write Log To Table
'----------------------------------------------------------------------------------------------------------------------------------------------------
Private Function WriteLogToTBL(sLog As String, Optional LOGTBL As String = "$$LOGS", Optional DLM As String = ";", _
                                                                                                         Optional SEP As String = vbCrLf) As Boolean
Dim bRes As Boolean, RS As DAO.Recordset
Dim LOGS() As String, nDim As Integer, I As Integer, WORKS() As String

Const LOG_DIM As Integer = 3

    On Error GoTo ErrHandle
'-----------------------
If sLog = "" Then Exit Function
LOGS = Split(sLog, SEP): nDim = UBound(LOGS)

If Not IsTable(LOGTBL) Then
   If Not CreateLogTable(LOGTBL) Then Err.Raise 10006, , "Can't create Log Table"
End If
     
    Set RS = CurrentDb.OpenRecordset(LOGTBL)
    With RS
        For I = 0 To nDim
            If LOGS(I) <> "" Then
                  WORKS = Split(LOGS(I), DLM)
                  If UBound(WORKS) <> LOG_DIM Then GoTo NextLog ' SKIP WRONG logMsg Format
                  .AddNew
                      If WORKS(0) <> "" Then .FIELDS("LogMsg").value = WORKS(0)
                      If WORKS(1) <> "" Then .FIELDS("FuncName").value = WORKS(1)
                      If WORKS(2) <> "" Then .FIELDS("Context").value = WORKS(2)
                      If WORKS(3) <> "" Then .FIELDS("LogID").value = CLng(WORKS(3))
                  .Update
            End If
NextLog:
        Next I
    End With
'-----------------------
ExitHere:
    WriteLogToTBL = True '!!!!!!
    Set RS = Nothing
    Exit Function
'---------
ErrHandle:
    ErrPrint2 "WriteLogToTBL", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Set RS = Nothing
End Function
'---------------------------------------------------------------------------------------------------------------------------------------------------
' Create the Log Table
'----------------------------------------------------------------------------------------------------------------------------------------------------
Private Function CreateLogTable(Optional LOGTBL As String = "$$LOGS") As Boolean
Dim SQL As String

On Error GoTo ErrHandle
'----------------------------
    SQL = "CREATE TABLE " & SHT(LOGTBL) _
        & "([ID] AUTOINCREMENT PRIMARY KEY, [DateCreate] DATETIME, [IsArchive] YESNO,[HASH] CHAR(250), " _
        & "[Context] TEXT(250),[FuncName] TEXT(250), [LogID] INTEGER, " _
        & "[LogMsg] MEMO);"
    
    DoCmd.SetWarnings False
        CurrentDb.Execute SQL
        '----------------------------------------
        CurrentDb.TableDefs.Refresh
        
        SetDefaultValueForField LOGTBL, "DateCreate", "= Now()"
'----------------------------
ExitHere:
    CreateLogTable = True '!!!!!!!!!!!!
    DoCmd.SetWarnings True
    Exit Function
'------------
ErrHandle:
    ErrPrint "CreateLogTable", Err.Number, Err.Description
    Err.Clear: DoCmd.SetWarnings True
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------
' Create the Messaging Table
'----------------------------------------------------------------------------------------------------------------------------------------------------
Private Function CreateMsgTable(Optional MSGTBL As String = "$$MSGS") As Boolean
Dim SQL As String

On Error GoTo ErrHandle
'----------------------------
    SQL = "CREATE TABLE " & SHT(MSGTBL) _
        & "([ID] AUTOINCREMENT PRIMARY KEY, [DateCreate] DATETIME, [IsArchive] YESNO,[HASH] CHAR(250), " _
        & "[MsgID] INTEGER, [ModName] TEXT(250),[FuncName] TEXT(250), [ErrID] INTEGER,[LangID] INTEGER, " _
        & "[Description] MEMO);"
    
    DoCmd.SetWarnings False
        CurrentDb.Execute SQL
        '----------------------------------------
        CurrentDb.TableDefs.Refresh
        
        SetDefaultValueForField MSGTBL, "DateCreate", "= Now()"
        Call SetLookUpFLD(MSGTBL, "LangID", "0;EN;1;RU", , "Value List")
        SetDefaultValueForField MSGTBL, "LangID", 0

'----------------------------
ExitHere:
    CreateMsgTable = True '!!!!!!!!!!!!
    DoCmd.SetWarnings True
    Exit Function
'------------
ErrHandle:
    ErrPrint "CreateMsgTable", Err.Number, Err.Description
    Err.Clear: DoCmd.SetWarnings True
End Function
'====================================================================================================================================================
' Check if specific Key os pressed
'====================================================================================================================================================
Public Function IsKeyPress(Optional iKey As Long = VK_ESCAPE) As Boolean
Dim I As Integer
On Error Resume Next
        I = DoEvents
        IsKeyPress = GetAsyncKeyState(iKey) '!!!!!!!!!!!!!!!!!!
End Function



'============================================================================================================================================
' Show Progress
'      ShowCommand = 0 - Off, 1 - On, 2 - Update
'============================================================================================================================================
Public Sub ProgressMeter(ShowCommand As Integer, iCount As Long, Optional TextShow As String)

    Select Case ShowCommand
        Case 0:
            SysCmd acSysCmdRemoveMeter
        Case 1:
            SysCmd acSysCmdInitMeter, TextShow, iCount
        Case 2:
            SysCmd acSysCmdUpdateMeter, iCount
    End Select
End Sub
'====================================================================================================================================================
' get Boolean from string or something like this
'====================================================================================================================================================
Public Function GetBool(v As Variant) As Boolean
Dim bRes As Boolean, sWork As String

On Error Resume Next
'--------------------------
    If IsEmpty(v) Then Exit Function
    
    If varType(v) = vbBoolean Then
    ElseIf varType(v) = vbString Then
           sWork = UCase(Trim(CStr(v)))
           If sWork = "TRUE" Or sWork = "FALSE" Then bRes = CBool(sWork)
    ElseIf IsNumeric(v) Then
           bRes = CBool(v)
    End If
'--------------------------
ExitHere:
    GetBool = bRes '!!!!!!!!!!
End Function

'====================================================================================================================================================
' Check if  VARIANT is Null or Empty
'====================================================================================================================================================
Public Function IsZero(v As Variant) As Boolean
Dim bRes As Boolean
Dim tenpStr As String, nDim As Long
    
On Error GoTo ErrHandle
'--------------------------------------
    If IsDate(v) Then
        tenpStr = Trim(Replace(Replace(UCase(CStr(v)), "AM", ""), "PM", ""))
        bRes = (tenpStr = "00:00:00") Or (tenpStr = "0:00:00") Or (tenpStr = "12:00:00")
        GoTo ExitHere
    ElseIf IsNumeric(v) Then
        bRes = v = 0
        GoTo ExitHere
    ElseIf IsArray(v) Then
        bRes = Not ((Not Not v) <> 0)
        GoTo ExitHere
    End If
'--------------------------------------
' IF WE ARE HERE THEN UNEXPECTED VARIANT
    bRes = IsMissing(v)
    If Not bRes Then
        bRes = IsEmpty(v)
    End If
    If Not bRes Then
       bRes = IsNull(v)
    End If
    '--------------------------------------
    If Not bRes Then
        tenpStr = Trim(CStr(v))
        If tenpStr = "" Or tenpStr = "-1" Then
            bRes = True
        End If
    End If
    '--------------------------------------
ExitHere:
        IsZero = bRes '!!!!!!!!!!!!!!!!!!!!
        Exit Function
    '--------------------------------------
ErrHandle:
        Err.Clear: bRes = True
        Resume ExitHere
End Function
'====================================================================================================================================================
' This finction create Boolean from string
'====================================================================================================================================================
Public Function DBool(vBL As Variant) As Boolean
  On Error Resume Next
        If IsEmpty(vBL) Then Exit Function
        If CStr(vBL) = "" Then Exit Function
'---------------------------------------------
    DBool = CBool(vBL) '!!!!!!!!!!!!
End Function

'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
' LOCAL SETTING MANIPULATION
'--------------------------------------------------------------------------------------------------------------------------
' Check if service Table $$Local Exists
' This function works only in Access VBA and supports functions GetLocal and SetLocal
'--------------------------------------------------------------------------------------------------------------------------
Public Function RecreateLocalTable(Optional LocalTable As String = "$$LOCAL") As String
Dim SQL As String, DLM As String, sRes As String

On Error GoTo ErrHandle
'-------------------------------------------
If IsNull(DLookup("Name", "MSysObjects", "Name=" & _
                  Chr(39) & LocalTable & Chr(39) & " And Type In (1,4,6)")) Then
    
    DLM = Chr(34) & "#" & Chr(34)
    SQL = "CREATE TABLE [" & LocalTable & "]" _
        & "([ID] AUTOINCREMENT PRIMARY KEY, [DateCreate] DATETIME, [DateUpdate] DATETIME, [IsArchive] YESNO,[HASH] CHAR(250), " _
        & "[DOMAIN] CHAR, " _
        & "[ParamName] CHAR, [ParamValue] TEXT(250),[DefaultValue] TEXT(250), [Description] TEXT(250), " _
        & "[xTension] MEMO,[IsDefault] YESNO);"
    CurrentDb.Execute SQL, dbFailOnError
    '---------------------------
    DoCmd.SetWarnings False
    CurrentDb.TableDefs.Refresh
   
    CurrentDb.TableDefs(LocalTable).FIELDS("DateCreate").DefaultValue = "= Now()"
    CurrentDb.TableDefs(LocalTable).FIELDS("DateUpdate").DefaultValue = "= Now()"
    CurrentDb.TableDefs(LocalTable).FIELDS("HASH").DefaultValue = "= " & DLM & _
                                                     " & (Int((1000000-10+1)*Rnd()+10)) & " & _
                                                     DLM & " & Date$() & " & DLM & " & Time$()"
    sRes = LocalTable
Else
    sRes = LocalTable
End If
'-------------------------------------------
ExitHere:
    RecreateLocalTable = sRes '!!!!!!!!!!
    DoCmd.SetWarnings True
    Exit Function
'----------------
ErrHandle:
    ErrPrint "RecreateLocalTable", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'==========================================================================================================================
' Read Param Value from LOCALPARAM. If absent then return "" + Environ("ComputerName")
'   @ Valery Khvatov (valery.khvatov@gmail.com), [01/20180501]
'==========================================================================================================================
Public Function GetLocal(ParamName As String, Optional bExtention As Boolean, _
                                                                        Optional LocalTbl As String = "$$LOCAL") As String
Dim sRes As String, sWhere As String
Dim LookFld As String

On Error GoTo ErrHandle
'---------------------------------------------------
If ParamName = "" Then Exit Function
If IsNull(DLookup("Name", "MSysObjects", "Name=" & _
                  Chr(39) & LocalTbl & Chr(39) & " And Type In (1,4,6)")) Then Exit Function
                  
LookFld = IIf(bExtention, "xTension", "ParamValue")

sWhere = "(ParamName = " & Chr(39) & ParamName & Chr(39) & ") AND (DOMAIN = " & Chr(39) & Environ("ComputerName") & Chr(39) & ")"
      sRes = Nz(DLookup(LookFld, SHT(LocalTbl), sWhere), "")
      If sRes = "" Then
           sWhere = "ParamName = " & Chr(39) & ParamName & Chr(39)
           sRes = Nz(DLookup(LookFld, SHT(LocalTbl), sWhere), "")
      End If
'----------------------------------------------------
ExitHere:
      GetLocal = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
      Exit Function
'--------------------
ErrHandle:
      ErrPrint "GetLocal", Err.Number, Err.Description
      Err.Clear: Resume ExitHere
End Function
'==========================================================================================================================
' Set Param Value in Local Param. If Param is absent thent create new ParamName
'==========================================================================================================================
Public Function SetLocal(ParamName As String, ParamValue As String, Optional sDescription As String = "", _
                                                          Optional sDefaultValue As String, Optional xTension As String, _
                                     Optional bDefault As Boolean = False, Optional LocalTbl As String = "$$LOCAL") As Long
Dim sSQL As String, IDD As Long, sWhere As String
Dim sParam As String, sValue As String, sHASH As String
Dim RS As DAO.Recordset, sDescr As String

On Error GoTo ErrHandle
'---------------------------------------
    If ParamName = "" Then Exit Function
    If RecreateLocalTable(LocalTbl) = "" Then Exit Function
'---------------------------------------
    sWhere = "(ParamName = " & Chr(39) & ParamName & Chr(39) & _
             ") AND (DOMAIN = " & Chr(39) & Environ("ComputerName") & Chr(39) & ")"
    
    IDD = Nz(DLookup("ID", "[" & LocalTbl & "]", sWhere), -1)
    '-------------------------------------------
    If IDD = -1 Then
        sSQL = "SELECT * FROM " & "[" & LocalTbl & "]"
        sHASH = GetHASH("#"): sDescr = IIf(sDescription <> "", sDescription, _
                                      "Added by " & Environ("UserName") & " at " & Now())
        Set RS = CurrentDb.OpenRecordset(sSQL)
        With RS
             If Not .EOF Then
                 .MoveLast: .MoveFirst
             End If
             .AddNew
                   !HASH = sHASH
                   !Domain = Environ("ComputerName")
                   !ParamName = ParamName
                   If ParamValue <> "" Then !ParamValue = ParamValue
                   If sDefaultValue <> "" Then !DefaultValue = sDefaultValue
                   If sDescr <> "" Then !Description = sDescr
                   If xTension <> "" Then !xTension = xTension
                   If bDefault Then !IsDefault = True
             .Update
        End With
        IDD = Nz(DLookup("ID", "[" & LocalTbl & "]", "HASH = " & sCH(sHASH)), -1)
    Else
        sSQL = "SELECT * FROM " & "[" & LocalTbl & "]" & " WHERE ID = " & IDD
        Set RS = CurrentDb.OpenRecordset(sSQL)
        With RS
            If Not .EOF Then
                 .MoveLast: .MoveFirst
                 .Edit
                   !Domain = Environ("ComputerName")
                   If ParamValue <> "" Then !ParamValue = ParamValue
                   If sDefaultValue <> "" Then !DefaultValue = sDefaultValue
                   If sDescription <> "" Then !Description = sDescription
                   If xTension <> "" Then !xTension = xTension
                   If bDefault Then !IsDefault = True
                 .Update
            End If
        End With
    End If
'------------------------------------
ExitHere:
    Set RS = Nothing
    DoCmd.SetWarnings True
    SetLocal = IDD '!!!!!!!!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint "SetLocal", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'=====================================================================================================================================================
' Change Root (Looking the table with name _Roots and Set As Root the default one
'    iOP = 0  - Save Provided Root To RootTable. If no any root table - create it
'    iOP = 1  - Get Root from Root Table and make it default root
'    iOP = 2  -
'=====================================================================================================================================================
Public Function ChangeRoot(Optional iOp As Integer = 0, Optional RootTable As String = "_Root", Optional sNewRoot As String, _
                                                                                                                   Optional sPath As String) As String
Dim sRes As String, SQL As String

Const Root_Table As String = "_ROOTS"

    On Error GoTo ErrHandle
'------------------
Select Case iOp
Case 0:    ' Save The Root  to Root Table

      If sNewRoot = "" Then Exit Function
      If Not IsTable(Root_Table) Then
         sRes = CreateRootTable(Root_Table)
         If sRes = "" Then Err.Raise 1007, , "Can't create Root Table"
      End If
      
      sRes = FileNameOnly(sNewRoot)
      SQL = "INSERT INTO " & Root_Table & "(ROOT, ROOTNAME)" & IIf(sPath <> "", " IN " & sCH(sPath), "") & _
              " VALUES(" & sCH(sNewRoot) & ", " & sCH(sRes) & ")"
      DoCmd.SetWarnings False
      CurrentDb.Execute SQL

Case 1:    '  Get Root from RootTable and make it default
      If Not IsTable(RootTable) Then Err.Raise 1009, , "Can't find Root Table " & Root_Table
      If sNewRoot <> "" Then
           sRes = Nz(DLookup("ROOT", Root_Table, "ROOTNAME = " & sCH(sNewRoot)), "")
           If sRes = "" Then Err.Raise 10009, , "Can't Get Root with name " & sNewRoot
      Else
           sRes = Nz(DLookup("ROOT", Root_Table, "ID > 0"), "")
      End If
           If sRes <> "" Then
                If Dir(sRes) <> "" Then SetRoot (sRes)
           End If
Case 2:   ' Save Current Root to Root Table
      If Not IsTable(Root_Table) Then
         sRes = CreateRootTable(Root_Table)
         If sRes = "" Then Err.Raise 1007, , "Can't create Root Table"
      End If
      
      sRes = GetRoot(): If sRes = "" Then Exit Function
      SQL = "INSERT INTO " & Root_Table & "(ROOT, ROOTNAME)" & IIf(sPath <> "", " IN " & sCH(sPath), "") & _
              " VALUES(" & sCH(sRes) & ", " & sCH(FileNameOnly(sRes)) & ")"
      DoCmd.SetWarnings False
      CurrentDb.Execute SQL
End Select
'------------------
ExitHere:
       ChangeRoot = sRes '!!!!!!!!!
       DoCmd.SetWarnings False
       Exit Function
'------------
ErrHandle:
        ErrPrint2 "ChangeRoot", Err.Number, Err.Description, MOD_NAME
        Err.Clear
End Function
'=====================================================================================================================================================
' Get Saved Root specified for this database
'=====================================================================================================================================================
Public Function GetRoot() As String
Dim sRes As String
     
    sRes = Nz(TempVars.Item("ROOT").value, "")
    If sRes = "" Then
        sRes = GetLocal("ROOT")
        If sRes <> "" Then TempVars.Item("ROOT").value = sRes
    End If
'-------------------
    GetRoot = sRes '!!!!!!!!!!!!!!!!
End Function


'--------------------------------------------------------------------------------------------------------------------------
' Create _Root Table on the fly
'--------------------------------------------------------------------------------------------------------------------------
Private Function CreateRootTable(Optional RootTable As String = "_ROOTS") As String
Dim SQL As String, sRes As String

On Error GoTo ErrHandle
'-------------------------------------------
    DoCmd.SetWarnings False
    
    SQL = "CREATE TABLE [" & RootTable & "]" _
        & "([ID] AUTOINCREMENT PRIMARY KEY, [IsArchive] YESNO, " _
        & "[ROOT] CHAR, [ROOTNAME] TEXT(250),[IsDefault] YESNO);"
    
    CurrentDb.Execute SQL, dbFailOnError
    '---------------------------
    CurrentDb.TableDefs.Refresh
   
    sRes = RootTable
'-------------------------------------------
ExitHere:
    CreateRootTable = sRes '!!!!!!!!!!
    DoCmd.SetWarnings True
    Exit Function
'----------------
ErrHandle:
    ErrPrint2 "CreateRootTable", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Resume ExitHere
End Function


'=====================================================================================================================================================
' Set Root for file storing folders
'=====================================================================================================================================================
Public Sub SetRoot(Optional sROOT As String)
Dim RootPath As String

On Error GoTo ErrHandle
'------------------------------
    If sROOT <> "" Then
       RootPath = sROOT
    Else
       RootPath = OpenDialog(GC_FOLDER_PICKER, "Set Root Path", , , CurrentProject.Path)
    End If
    If RootPath = "" Then Exit Sub
    
    If Not IsFolderExists(RootPath) Then Err.Raise 1000, , "Non existing path: " & sROOT
    SetLocal "ROOT", RootPath
'----------------------------
ExitHere:
    Exit Sub
'------------
ErrHandle:
    ErrPrint "SetRoot", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Sub
'=====================================================================================================================================================
' Get Last Folder (Only Access)
'=====================================================================================================================================================
Public Function GetLastFolder() As String
Dim sRes As String

    On Error Resume Next
'--------------------
    sRes = Nz(TempVars.Item("LastFolder").value, "")
    If sRes = "" Then sRes = GetLocal("LastFolder")
'--------------------
ExitHere:
    GetLastFolder = sRes '!!!!!!!!!!!!!!!
End Function
'=====================================================================================================================================================
' Get Last Folder (Only Access)
'=====================================================================================================================================================
Public Sub SetLastFolder(sFolder As String)
     TempVars.Item("LastFolder").value = sFolder
     SetLocal "LastFolder", sFolder, "This is recent folder"
End Sub


'=======================================================================================================================================================
' Get Number Of Array Dimensions
'=======================================================================================================================================================
Public Function NumberOfArrayDimensions(Arr As Variant) As Integer
Dim I As Integer, iRes As Integer

On Error Resume Next
Do
    I = I + 1
    iRes = UBound(Arr, I)
Loop Until Err.Number <> 0
'-----------------------
ExitHere:
    NumberOfArrayDimensions = I - 1
End Function
'======================================================================================================================================================
' Format Date in Knowledge Style (with Today and ago prefixes
'======================================================================================================================================================
Public Function FormatDate(ddd As Date, Optional sPatternDate As String = "dd.mm.yyyy", Optional sPatternTime As String = "hh:nn", _
                         Optional sToday As String = "Today", Optional sYesterday As String = "Yesterday", Optional sTomorrow As String = "Tomorrow", _
                                           Optional sAgo As String = "ago", Optional sHour As String = "hr", Optional sMin As String = "min") As String

Dim sDate As String, sTime As String
Dim dHours As Double

On Error GoTo ErrHandle
'--------------------------------------
sDate = TODAYYESTERDAY(ddd, sPatternDate, sToday, sYesterday, sTomorrow)
If sDate = sToday Then
   dHours = GetHoursAgo(ddd, Now())
   If dHours > 0 And dHours < 4 Then
      sTime = HourMins(dHours, sHour, sMin) & " " & sAgo
   Else
      sTime = Format(ddd, sPatternTime)
   End If
Else
   sTime = Format(ddd, sPatternTime)
End If
'----------------------------------------
ExitHere:
    FormatDate = sDate & " " & sTime  '!!!!!!!!!!!!!
    Exit Function
'-------------
ErrHandle:
    ErrPrint "FormatDate", Err.Number, Err.Description
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
                                                                                                  Optional sModName As String = "#_UTIL") As String
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
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Hours ago
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetHoursAgo(Date1 As Date, Date2 As Date) As Double
On Error Resume Next
    GetHoursAgo = DateDiff("n", Date1, Date2) / 60
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Get Hours and mins from dHours
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function HourMins(dHours As Double, Optional sHour As String = "hr", Optional sMin As String = "min") As String
Dim iHour As Integer, iMin As Integer
    iHour = Fix(dHours)
    iMin = Fix((dHours - iHour) * 60)
'---------------------------------------
    HourMins = iHour & Space(1) & sHour & Space(1) & iMin & Space(1) & sMin '!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Get a name for closest date
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function TODAYYESTERDAY(ddd As Date, Optional sPattern As String = "dd.mm.yyyy", Optional sToday As String = "Today", _
                                                      Optional sYesterday As String = "Yesterday", Optional sTomorrow As String = "Tomorrow") As String
Dim nDays As Long, sRes As String

On Error GoTo ErrHandle
'---------------------------------
    nDays = DateDiff("d", ddd, Now())
    Select Case nDays
    Case -1:
        sRes = sTomorrow
    Case 0:
        sRes = sToday
    Case 1:
        sRes = sYesterday
    Case Else
        sRes = Format(ddd, sPattern)
    End Select
'---------------------------------
ExitHere:
    TODAYYESTERDAY = sRes '!!!!!!!!!
    Exit Function
'------------------
ErrHandle:
    ErrPrint "TODAYYESTERDAY", Err.Number, Err.Description
    Err.Clear
End Function

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'=======================================================================================================================================================
' Public Error Handler
'=======================================================================================================================================================
Public Function ErrPrint2(FuncName As String, ErrNumber As Long, ErrDescription As String, Optional sModName As String = "NONAME MDULE", _
                                                                                                          Optional bDebug As Boolean = True) As String
Dim sRes As String
Const ERR_CHAR As String = "#"
Const ERR_REPEAT As Integer = 60

sRes = String(ERR_REPEAT, ERR_CHAR) & vbCrLf & "ERROR OF [" & sModName & ": " & FuncName & "]" & vbTab & "ERR#" & ErrNumber & vbTab & Now() & _
       vbCrLf & ErrDescription & vbCrLf & String(ERR_REPEAT, ERR_CHAR)
If bDebug Then Debug.Print sRes
'----------------------------------------------------------
ExitHere:
       Beep
       ErrPrint2 = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
