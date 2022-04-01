VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cVBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
' Класс реализует функции работы с кодом
' @ Valery Khvatov, (c)DigitalXpert 2011, v 0.1 10/03/2013
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
Option Explicit

Private Type Pairs
     FIRST As String
     Second As String
End Type

Private Const MOD_NAME As String = "cVBA"
'************************************************************************
Private Type VBAMODULE
    Name As String                     ' Module/Function Name
    FullName As String                 ' Full Name (the same for module, but not fo function/sub)
    fileName As String                 ' This is for exported functions
    Header As String                   ' Header for module/fumction - declaration
    Version As String                  ' version of module
    Author As String                   ' Module author
    CreateDate As Date                 '
    UpdateDate As Date                 '
    Type As String                     ' Module Type (STANDALONE = 5.0 - can't recognize, 5.1 - standard module, 5.2 - class,  Function/Sub/Property Get и )
    ParentObjName As String            ' Link to parent object ( 5.1 = Modules;  5.2 = Class Modules, for Function  - its type, i.e. String, Integer)
    code As String                     ' All code
    ObfuscatedCode As String           ' Cleared and obfuscated code
    HTMLCode As String                 ' HTML-formated code
    ROWS As Long                       ' Lines
    GlobalVars As String               ' Global Vars
    Functions As String                ' all functions in module (short names)
    Description As String              ' The Module description
    CONSTANTS As String                ' Constants
    APIs As String                     ' all presented API
    VARS As String                     ' all vars
    AllCodes As String                 ' Code with comments and others
    Metrics As String                  ' Metrics
End Type

Private Type TFUNCTION
    Name As String
    iPrivacy As Integer                ' 0 - Public, 1 - Private, 2 - Friend
    IsStatic As Boolean                ' Only For Static
    TheType As Integer                 ' 0 - Sub, 1 - Function, 2 - Property Get, 3 - Property Let, 4 - Property Set, 5 - Declare Sub, 6 Declare Function
    ArgumentList As String
    ReturnType As String
    Description As String
End Type

Private sMEMO As String                ' Concatention string
Private mVBAConst() As String          ' Constants

'******************************************************************************************************************************
Const PathToLib As String = "C:\Users\Valery\DBs\VBALIB\2015_VBALIB.accdb"
Const VBA1 As String = "AddHandler;AddressOf;Aggregate;Alias;And;AndAlso;Ansi;AsAssembly;Auto;AppActivate;Asc;AscW;"
Const VBA2 As String = "Binary;Boolean;ByRef;Byte;ByVal;Beep;"
Const VBA3 As String = "Call;Case;Catch;CBool;CByte;CChar;CDate;CDec;CDbl;Char;CInt;Class;CLng;CObj;Const;Continue;Compare;CSByte;CShort;CSng;"
Const VBA4 As String = "CStr;CType;CUInt;CULng;CUShort;Custom;"
Const VBA5 As String = "CallByName;CBool;CByte;CChar;CDate;CDbl;CDec;ChDir;ChDrive;Choose;Chr;ChrW;CInt;CLng;CObj;Command;CreateObject;CShort;CSng;CStr;CType;CurDir;"
Const VBA6 As String = "Date;Decimal;Declare;Default;Delegate;Dim;DirectCast;Distinct;Do;Double;DateAdd;DateDiff;DatePart;DateSerial;DateValue;Day;"
Const VBA7 As String = "DDB;DeleteSetting;Derived Math;Dir;"
Const VBA8 As String = "Each;Else;ElseIf;End;EndIf;Enum;Enumeration;Equals;Erase;Error;Event;Exit;Explicit;Environ;EOF;ErrorToString;"
Const VBA9 As String = "False;Finally;For;Friend;From;Function;FileAttr;FileClose;FileCopy;FileDateTime;FileGet;FileGetObject;FileLen;FileOpen;FilePut;"
Const VBA10 As String = "FilePutObject;FileWidth;Filter;Fix;Format;FormatCurrency;FormatDateTime;FormatNumber;FormatPercent;FreeFile;FV;"
Const VBA11 As String = "Get;GetType;GetXMLNamespace;Global;GoSub;GoTo;Group By;Group Join;GetAllSettings;GetAttr;GetChar;GetException;GetObject;GetSetting;Handles;Hex;Hour;"
Const VBA12 As String = "If;Into;Implements;Imports;In;Inherits;Integer;Interface;Is;IsFalse;IsNot;IsTrue;IIf;Input;InputBox;InputString;InStr;InStrRev;"
Const VBA13 As String = "Int;Fix;IPmt;IRR;IsArray;IsDate;IsDBNull;IsError;IsNothing;IsNumeric;IsReference;Join;Key;Kill;"
Const VBA14 As String = "Left;Let;Lib;Like;Long;Loop;LBound;LCase;Left;Len;LineInput;Loc;Lock;Unlock;LOF;LSet;"
Const VBA15 As String = "Me;Mid;Mod;Module;MustInherit;MustOverride;MyBase;MyClass;Mid;Minute;MIRR;MkDir;Month;MonthName;MsgBox;"
Const VBA16 As String = "Namespace;Narrowing;New;Next;Not;Nothing;NotInheritable;NotOverridable;NPer;NPV;"
Const VBA17 As String = "Object;Of;Off;On;Operator;Option;Optional;Or;Order By;OrElse;Overloads;Overridable;Overrides;Oct;"
Const VBA18 As String = "ParamArray;Partial;Preserve;Private;Property;Protected;Public;Partition;Pmt;PPmt;Print;PrintLine;PV;QBColor;"
Const VBA19 As String = "RaiseEvent;ReadOnly;ReDim;REM;RemoveHandler;Resume;Return;Randomize;Rate;Rename;Replace;Reset;RGB;Right;RmDir;Rnd;RSet;RTrim;"
Const VBA20 As String = "SByte;Select;Set;Skip;Skip While;Shadows;Shared;Short;Single;Static;Step;Stop;Strict;String;Structure;Sub;SyncLock;"
Const VBA21 As String = "SaveSetting;Second;Seek;SetAttr;Shell;SLN;Space;SPC;Split;Str;StrComp;StrConv;StrDup;StrReverse;Switch;SYD;SystemTypeName;"
Const VBA22 As String = "Take;Take While;Text;Then;Throw;To;True;Try;TryCast;TypeOf;TAB;TimeSerial;TimeValue;Trim;TypeName;"
Const VBA23 As String = "Variant;Variable;Val;VarType;VbTypeName;UInteger;Unicode;Until;ULong;UShort;Using;UBound;UCase;Unlock;"
Const VBA24 As String = "Wend;When;Where;While;Widening;With;WithEvents;WriteOnly;Weekday;WeekdayName;Write;WriteLine;Year;Xor;"
Const VBA25 As String = "#Const;#Else;#ElseIf;#End;#ExternalSource;#If;"
Const VBA26 As String = "=;&;&=;*;*=;/;/=;\;\=;^;^=;+;+=;-;-=;>>;>>=;<<;<<=;(;);"
'*****************************************************************************************************************************
Const vbext_ct_ClassModule = 2
Const vbext_ct_Document = 100
Const vbext_ct_MSForm = 3
Const vbext_ct_StdModule = 1



'******************************************************************************************************************************

'======================================================================================================================================================
' The function in the current database creates / supplements the $ VBACODE table and refills the data as needed
' Description of the module's functions
'======================================================================================================================================================
Public Sub WriteModInfo(ModuleName As String, Optional sTableName As String = "$VBACODE")
Dim bRes As Boolean                     ' Result
Dim ProcNames() As String, nDim As Integer, I As Integer
Dim ProcInfo() As String                ' Information about the function
Dim sHeader() As String
Dim IDD As Integer

On Error GoTo ErrHandle
' 1) We will check the value of the tables and if they are not created
   If Not TableExists(sTableName) Then
      bRes = CreateTableVBA(sTableName)
      If Not bRes Then Exit Sub
   End If
' 2) Getting the Name of the Modules Procedure
      ProcNames = Split(ListProcs(ModuleName), ";"): nDim = UBound(ProcNames)
' 3) Let's take each of the functions, display the information about her and write it to the table
      For I = 0 To nDim
          ProcInfo = ProcessProcedure(ModuleName, ProcNames(I))
          sHeader = Split(ProcInfo(1), ";")
          '------------------------------------------------------
          Debug.Print "Procedure: " & ProcNames(I)
          IDD = WriteToVBATAble(sTableName, ModuleName, ProcNames(I), sHeader(2), sHeader(3), sHeader(5), ProcInfo(0), CLng(ProcInfo(4)), ProcInfo(2), Replace(ProcInfo(3), Chr(39), ""))
          Erase ProcInfo
      Next I
'--------------------------------------------------------
ExitHere:
     Exit Sub
'------------------
ErrHandle:
     ErrPrint2 "WriteModInfo", Err.Number, Err.Description, MOD_NAME
     Err.Clear
End Sub
'======================================================================================================================================================
' For a preset procedure of a preset module, the function will run
' Calls up the data masses according to the selected procedure
'======================================================================================================================================================
Public Function ProcessProcedure(ModName As String, ProcName As String, Optional ProcType As Integer = 0) As String()
Dim sRes As String                      ' Result
Dim MyVBA As VBAMODULE                  ' Processed module
Dim myProc() As String                  ' Procedure's Params Array
Dim sCode As String                     ' The code of procedure

    On Error Resume Next
'--------------------------------------------------------------------------
    ReDim myProc(4)
        
    sCode = Trim(GetProcCODE(ModName, ProcName, ProcType))
    If sCode = "" Then GoTo ExitHere
    MyVBA = ParceProcedures(ModName & "." & ProcName, sCode)
    
    
    myProc(0) = MyVBA.AllCodes          ' All code with comments
    myProc(1) = MyVBA.Header            ' The header of procedure
    myProc(2) = MyVBA.Metrics           ' The metrics
    myProc(3) = MyVBA.Description       ' Function Description
    myProc(4) = MyVBA.ROWS              ' SLOC
'----------------------------------------------------------------------------------
ExitHere:
    ProcessProcedure = myProc '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'======================================================================================================================================================
' Get Detailed Report
'======================================================================================================================================================
Public Function GetFullReport() As String
Dim sRes As String, SMODS() As String, nDim As Integer, I As Integer, sSB() As String
Dim sWork As String, U() As String, W() As String, SBS() As String, nSBS As Integer, J As Integer
Dim nAllLines As Integer, nAllFuncs As Integer, sMod As String, sModName As String, sProcName As String

Const nLNG As Integer = 80

    On Error GoTo ErrHandle
'------------------------------------------
sWork = ListAllModules(";", "|"): If sWork = "" Then GoTo ExitHere
SMODS = Split(ListAllModules(";", "|"), ";"): nDim = UBound(SMODS)
For I = 0 To nDim
    U = Split(SMODS(I), "|")
    sModName = Trim(U(1))
        
    sRes = sRes & vbCrLf & sModName & vbTab & vbTab & "Lines: " & U(2) & vbTab & " Created: " & U(3)
    nAllLines = nAllLines + CInt(U(2))
           sWork = ListProcs(sModName)
           If sWork <> "" Then
                SBS = Split(sWork, ";"): nSBS = UBound(SBS)
                sMod = "": nAllFuncs = nAllFuncs + nSBS
                
                For J = 0 To nSBS
                   sSB = Split(SBS(J), "|"): sProcName = sSB(0)
                   '-------------------------------------
                   W = ProcessProcedure(sModName, sProcName, CInt(sSB(1)))
                   
                   If Trim(W(1)) <> "" Then
                       sWork = String(12, " ") & Split(W(1), ";")(2) & " " & _
                               Split(W(1), ";")(3) & " " & sProcName & "()"
                   Else
                       sWork = String(12, " ") & "---- ----- ----- ----- " & sProcName & "()"
                   End If
                   
                   sWork = sWork & String(70 - Len(sWork), " ") & "Lines: " & W(4)
                                    
                   sMod = sMod & sWork & ";"
                   '-------------------------------------
                Next J
           End If
           If sMod <> "" Then sMod = Replace(sMod, ";", vbCrLf)
           
    sRes = sRes & vbTab & " Functions: " & nSBS & vbCrLf & String(nLNG, "-") & vbCrLf & sMod
Next I
'-------------------
ExitHere:
      sRes = String(3, vbTab) & "VBA REPORT PREPARED " & Now() & vbCrLf & _
             "TotalLines: " & nAllLines & vbTab & vbTab & "Total Functions: " & nAllFuncs & _
             vbCrLf & String(nLNG, "*") & vbCrLf & String(nLNG, "*")
      
      GetFullReport = sRes '!!!!!!!!!!!!!!!!!
      Exit Function
'------------
ErrHandle:
      ErrPrint2 "GetFullReport", Err.Number, Err.Description, MOD_NAME
      Err.Clear
End Function

'======================================================================================================================================================
' Return vba module listing as String with format ProjectName|modName|modLines|modCreateData;ProjectName|modName|modLines|modCreateData
'======================================================================================================================================================
Public Function ListAllModules(Optional DLM As String = ";", Optional RDLM As String = "|") As String
Dim VBComp As Object, vbMod As Object, VBProj As Object
Dim sWork As String, sRes As String, sProjectName As String

    On Error GoTo ErrHandle
'-----------------------------------------------------------------------------
For Each VBProj In Application.VBE.VBProjects
           sProjectName = FileNameOnly(VBProj.fileName) & "." & VBProj.Name
           For Each VBComp In VBProj.VBComponents
           
                  Set vbMod = VBComp.CodeModule
                  sWork = sProjectName & RDLM & VBComp.Name & RDLM & _
                                vbMod.CountOfLines & RDLM & GetCreateDate(VBComp.Name)
                  
                  If sWork <> "" Then
                    sRes = sRes & sWork & DLM
                    sWork = ""
                  End If
             Next VBComp
Next VBProj
If sRes <> "" Then sRes = Left(sRes, Len(sRes) - Len(DLM))
'----------------------------------
ExitFunction:
       ListAllModules = sRes '!!!!!!!!!!!!!
       Exit Function
'-----------------
ErrHandle:
       ErrPrint2 "ListAllModules", Err.Number, Err.Description, MOD_NAME
       Err.Clear
End Function
'======================================================================================================================================================
' The function receives a list of modules for the current data base in the form of a line with a separator
' Separate lines are available for other modules and format modules and reports
'======================================================================================================================================================
Public Function ListModules(Optional SCENARIO As String = "STANDALONE", Optional sDelim As String = ";") As String
Dim sRes As String                      ' Result
Dim accObj As AccessObject              ' Access's object: module/form/report.
Dim bWasOpen As Boolean                 ' Is form open right now

    On Error GoTo ErrHandle
'-----------------------------------------
Select Case SCENARIO
Case "STANDALONE":
    For Each accObj In CurrentProject.AllModules
        sRes = IIf(sRes <> "", sRes & sDelim, sRes)
        sRes = sRes & accObj.Name
    Next accObj
Case "FORMS":
    For Each accObj In CurrentProject.AllForms
            bWasOpen = accObj.IsLoaded
            If Not bWasOpen Then
                DoCmd.OpenForm accObj.Name, acDesign, WindowMode:=acHidden
            End If
            If Forms(accObj.Name).HasModule Then
                        sRes = IIf(sRes <> "", sRes & sDelim, sRes)
                        sRes = sRes & "Form_" & accObj.Name
            End If
            If Not bWasOpen Then
                DoCmd.Close acForm, accObj.Name, acSaveNo
            End If
    Next
Case "REPORTS":
        For Each accObj In CurrentProject.AllReports
            bWasOpen = accObj.IsLoaded
            If Not bWasOpen Then
               'In Access 2000, remove the ", WindowMode:=acHidden" from the next line.
               DoCmd.OpenReport accObj.Name, acDesign, WindowMode:=acHidden
            End If
            If Reports(accObj.Name).HasModule Then
                        sRes = IIf(sRes <> "", sRes & sDelim, sRes)
                        sRes = sRes & "Report_" & accObj.Name
            End If
            If Not bWasOpen Then
                DoCmd.Close acReport, accObj.Name, acSaveNo
            End If
        Next
Case Else
     sRes = ""
End Select
'-----------------------------------------
ExitHere:
    ListModules = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'--------------------
ErrHandle:
    ErrPrint2 "ListModules", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Resume ExitHere
End Function
'======================================================================================================================================================
' The function receives a list of the preset module
'======================================================================================================================================================
Public Function GetVBACode(ModName As String, Optional nStart As Long = 1, Optional nEnd As Long = -1) As String
Dim iRes As Integer                                     ' Returning Result
Dim nLines As Long                                      ' Count of sLOC in Module
Dim sWork As String                                     ' Working var

On Error GoTo ErrHandle
'-----------------------------------------------------------------------
   sWork = ""
   iRes = CheckLoaded(ModName)
'-----------------------------------------------------------------------
   If iRes = 1 Or iRes = 2 Then                                        ' Check if is module in collection
        nLines = Application.Modules(ModName).CountOfLines             ' General SLOC
        If (nEnd > -1) And nEnd <= nLines Then nLines = nEnd           ' If nEnd is correct then place it to List
        sWork = Application.Modules(ModName).LINES(nStart, nLines)
   End If
'-----------------------------------------------------------------------
ExitHere:
        If iRes = 2 Then iRes = CheckLoaded(ModName, False)
        GetVBACode = sWork '!!!!!!!!!!!!!!!!
        Exit Function
'-------------------------
ErrHandle:
        ErrPrint2 "GetVBACode", Err.Number, Err.Description
        Resume ExitHere
End Function


'===============================================================================================================================
' The function is received for the preset function in the preset module
'               ProcType:
'                   vbext_pk_Get    lngR=3   A Property Get proc
'                   vbext_pk_Let    lngR=1   A Property Let proc
'                   vbext_pk_Proc   lngR=0   A Sub or Function proc
'                   vbext_pk_Set    lngR=2   A Property Set proc
'===============================================================================================================================
Public Function GetProcCODE(ModName As String, sProcName As String, Optional ProcType As Integer = 0) As String
Dim iRes As Integer                                     ' Returning result
Dim nLines As Long                                      ' count of sloc
Dim sWork As String                                     ' working var
Dim mdl As Module
Dim lngStartLine As Long, lngBodyLine As Long
Dim lngCount As Long, lngEndProc As Long

On Error GoTo ErrHandler
'-----------------------------------------------------------------------
   sWork = ""
   iRes = CheckLoaded(ModName)
'-----------------------------------------------------------------------
   If iRes = 1 Or iRes = 2 Then                                                         ' Collection has the module
        Set mdl = Modules(ModName)
        lngCount = mdl.ProcCountLines(sProcName, ProcType)                              ' calculate module sloc
        lngStartLine = mdl.ProcStartLine(sProcName, ProcType)                           ' find the first line
        lngBodyLine = mdl.ProcBodyLine(sProcName, ProcType)                             ' find first line to procedure body
        If lngBodyLine > lngStartLine Then
             sWork = mdl.LINES(lngStartLine, lngBodyLine - lngStartLine)                ' gather all lines before function declaration
        Else
             sWork = ""
        End If
        lngEndProc = (lngBodyLine + lngCount - 1) - Abs(lngBodyLine - lngStartLine)     ' last procedure line
        sWork = sWork & vbCrLf & mdl.LINES(lngBodyLine, (lngEndProc - lngBodyLine) + 1) ' all procedure body line
   End If
'-----------------------------------------------------------------------
ExitHere:
        If iRes = 2 Then iRes = CheckLoaded(ModName, False)
        GetProcCODE = sWork '!!!!!!!!!!!!!!!!
        Exit Function
'-------------------------
ErrHandler:
        Err.Clear
        Resume ExitHere
End Function
'------------------------------------------------------------------------------------------------------------------------------
' Function Return ProcType
'------------------------------------------------------------------------------------------------------------------------------
Private Function GetProcType(ModName As String, ProcName As String) As Integer
Dim iRes As Integer, I As Integer

On Error Resume Next
'-----------------------------------------------
Const vbext_pk_Get As Integer = 3 'A Property Get procedure.
Const vbext_pk_Let As Integer = 1 'A Property Let procedure.
Const vbext_pk_Set As Integer = 2 ' A Property Set procedure.
Const vbext_pk_Proc As Integer = 0 ' A Sub or Function procedure

iRes = -1
'----------------------------
For I = 0 To 3
     If ErrProcType(ModName, ProcName, I) Then
          iRes = I
          Exit For
     End If
Next I
'-------------------------------
     GetProcType = iRes '!!!!!!!!!!!!!!!!
End Function
Private Function ErrProcType(ModName As String, sProcName As String, ProcTypeTrial As Integer) As Boolean
Dim iRes As Integer

On Error GoTo ErrHandle
'------------------------------
     iRes = Modules(ModName).ProcCountLines(sProcName, ProcTypeTrial)
'------------------------------
ExitHere:
     If iRes > 0 Then ErrProcType = True '!!!!!!!!!!!!!
     Exit Function
'---------
ErrHandle:
     Err.Clear
End Function

'===============================================================================================================================
' The function receives a list of all functions in the specified module
'===============================================================================================================================
Public Function ListProcs(ModName As String, Optional sDelim As String = ";", Optional ProcTypeDLM As String = "|") As String
Dim iRes As Integer                                     ' Result
Dim sWork As String                                     ' work var
Dim mdl As Module
Dim lngCount As Long, lngCountDecl As Long, lngI As Long
Dim strProcName As String, astrProcNames() As String
Dim intI As Integer, lngR As Long

On Error GoTo ErrHandler
'-----------------------------------------------------------------------
   sWork = ""
   iRes = CheckLoaded(ModName)
'------------------------------------------------------------------------
   If (iRes = 1) Or (iRes = 2) Then
        Set mdl = Modules(ModName): lngCount = mdl.CountOfLines    ' Open the module and read the total number of lines
        lngCountDecl = mdl.CountOfDeclarationLines                 ' Read the number of lines on the declaration
        
        strProcName = mdl.ProcOfLine(lngCountDecl + 1, lngR)       ' Name of the first procedure
        
        intI = 0                                                   ' Counter initialization
        ReDim Preserve astrProcNames(intI)
        astrProcNames(intI) = strProcName & ProcTypeDLM & lngR
        '--------------------------------------------------------------------------
        For lngI = lngCountDecl + 1 To lngCount
                If strProcName <> mdl.ProcOfLine(lngI, lngR) Then
                    intI = intI + 1
                    strProcName = mdl.ProcOfLine(lngI, lngR)
                    ReDim Preserve astrProcNames(intI)
                    astrProcNames(intI) = strProcName & ProcTypeDLM & lngR
                End If
        Next lngI
        '--------------------------------------------------------------------------
        sWork = Join(astrProcNames, sDelim)
   End If
'----------------------------------------------------------------------------------------------
ExitHere:
     If iRes = 2 Then iRes = CheckLoaded(ModName, False)
     ListProcs = sWork '!!!!!!!!!!!!!!!!!
     Exit Function
'-----------------------------
ErrHandler:
     Err.Clear
     Resume ExitHere
End Function
'=========================================================================================================================
' This function gets the list of VBA modules in the given data base. If the data base is not set, the dialog box is displayed
'     If OneDB = True - display the dialog or display it on demand only for one database.
'        OneDB = False - scan all subscribed directories and create a separate table for all accdb databases.
'=========================================================================================================================
Public Sub GetAllModuleList(Optional sPath As String = "", Optional OneDB As Boolean = True)
Dim sPathToDB As String, sFileName As String                 ' FileName
Dim sList() As String, nDim As Long, I As Long               ' List of files
Dim sSQL As String                                           ' Query to db
Dim QryName As String                                        ' Query Name
Dim sPrefix As String                                        ' Rotaining Line
Dim mDBFile As String                                        ' The db file
Dim sText As String, IWidth As Long, iHeight As Long         ' The message size

   On Error GoTo ErrHandle
'-------------------------------------------------
If OneDB Then                                                ' Open database
    sPathToDB = IIf(sPath <> "", sPath, OpenDialog(1, "Выбор базы данных", "Файлы базы данных,*.accdb,*.mdb", False))
    If sPathToDB = "" Then Exit Sub
    sFileName = FileNameOnly(sPathToDB)
'--------------------------------------------------------------
    sSQL = "SELECT " & "'" & sFileName & "'" & "As DBName, MSysObjects.Id, MSysObjects.Name, MSysObjects.DateCreate, MSysObjects.DateUpdate " & _
       "FROM MSysObjects IN " & "'" & sPathToDB & "'" & _
       "WHERE (((MSysObjects.Type) = -32761)) ORDER BY MSysObjects.Name;"
'--------------------------------------------------------------
    QryName = CreateTempQuery(sSQL, , , True)
      MsgBox "Данные по модулям представлены временным запросом " & QryName
'------------------------------------------------------
Else                                                         ' List of database
   '-------------------------------------------------------------------------------------------------------------------
   ' As soon as the process is overridden - adjust the waiting method
   sPrefix = RotateLine("|"):   IWidth = 1000:   iHeight = 60
   '-------------------------------------------------------------------------------------------------------------------
   If sPath = "" Then
          sPathToDB = OpenDialog(4, "Выбор каталога", "", False)
   Else
          sPathToDB = sPath
   End If
   If sPathToDB = "" Then Exit Sub
   
   sText = sPrefix & "PROCESS IN " & Right(sPathToDB, Len(sPathToDB) - InStrRev(sPathToDB, "\"))
   
   'ShowMsg sText, , , , , , , iWidth, iHeight: Wait 3
      
      sList = LISTALLFILES(sPathToDB, "accdb;mdb")
      nDim = UBound(sList): If nDim = 0 And sList(0) = "" Then Exit Sub
      sPrefix = RotateLine(sPrefix): sText = sPrefix & "Found " & nDim & " dbs"
      'ShowMsg sText, , , , , , , iWidth, iHeight: Wait 1
      '-----------------------------------------------------------------
      ' Проверяем таблицу
        QryName = CreateTempTable(, , "ID LONG,DBName CHAR,DBPath TEXT(255),ModName CHAR, " & _
                                     "DateCreate DATETIME, DateUpdate DATETIME")
       If QryName = "" Then Exit Sub
       QryName = SHT(QryName)
      '-----------------------------------------------------------------
      'DoCmd.Hourglass True              ' START PROCESSING
      For I = 0 To nDim
             mDBFile = FileNameOnly(sList(I))
             sPrefix = RotateLine(sPrefix): sText = sPrefix & "Process " & I
             'ShowMsg sText, , , , , , , iWidth, iHeight
              InsertModInfo sList(I), QryName
      Next I
             sPrefix = RotateLine(sPrefix): sText = sPrefix & "Result In  " & QryName
            ' ShowMsg sText, , , , , , , iWidth, iHeight
             'Wait 5
      DoCmd.OpenTable Replace(Replace(QryName, "[", ""), "]", "")
End If
'-------------------------------------------------------------
ExitHere:
     'DoCmd.Hourglass False
     Exit Sub
'--------------------------------
ErrHandle:
    ErrPrint2 "", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Err.Clear
End Sub

'======================================================================================================================================================
' The function will be used by all modules in the file
'======================================================================================================================================================
Public Sub ExportAllCodeToFiles()
Dim c As Object, sPath As String, sFile As String
Dim Sfx As String
Const Folder As String = "CODE"

On Error GoTo ErrHandle
'---------------------------------------------------
sPath = CurrentProject.Path & "\" & Folder
If Dir(sPath, vbDirectory) = "" Then MkDir sPath
'---------------------------------------------------
For Each c In Application.VBE.VBProjects(1).VBComponents
    'Debug.Print c.Name
    Select Case c.Type
       
        Case vbext_ct_ClassModule
            Sfx = ".cls"
        Case vbext_ct_MSForm, vbext_ct_Document
            Sfx = ".frm"
        Case vbext_ct_StdModule
            Sfx = ".bas"
        Case Else
            Sfx = ""
    End Select
    If Sfx <> "" Then
        sFile = sPath & "\" & c.Name & Sfx
        If Dir(sFile) <> "" Then Kill sFile
        c.Export fileName:=sFile
    End If
Next c
'--------------------------------------------------
ExitHere:
      Exit Sub
'----------------
ErrHandle:
      ErrPrint2 "ExportAllCodeToFiles", Err.Number, Err.Description, MOD_NAME
      Err.Clear
End Sub
'===============================================================================================================================
' The function will operate the assigned module in the file, which will be opened in case of success
'===============================================================================================================================
Public Function ExportModule(ModName As String, Optional FolderName As String, Optional fileName As String, _
                              Optional OverwriteExisting As Boolean = True, Optional bAddVersion As Boolean = False) As String
Dim Extension As String
Dim Fname As String, fldr As String
Dim sRes As String, iRes As Integer, mdl As Module
On Error GoTo ErrHandler
'-----------------------------------------------------------------------
    iRes = CheckLoaded(ModName)
    Set mdl = Modules(ModName)
    Extension = FileExtension(mdl)
'-----------------------------------------------------------------------
   If iRes = 1 Or iRes = 2 Then                                                         ' Collection has this module
            '----------------------------------------------------------------------------------
            ' Assign file name
            If Trim(fileName) = vbNullString Then
                Fname = ModName & "." & Extension
            Else
                Fname = fileName
                If InStr(1, Fname, ".", vbBinaryCompare) = 0 Then
                    Fname = Fname & "." & Extension
                End If
            End If
            If bAddVersion Then Fname = Format(Now(), "yyyymmddhhnnss") & "_" & Fname
            '----------------------------------------------------------------------------------
            ' Letting folder
            If FolderName = "" Then
                fldr = CurrentProject.Path & "\"
            ElseIf StrComp(Right(FolderName, 1), "\", vbBinaryCompare) = 0 Then
                fldr = FolderName
            Else
                fldr = FolderName & "\"
            End If
            '----------------------------------------------------------------------------------
            ' Full file name
            Fname = fldr & Fname
            If Dir(Fname, vbNormal + vbHidden + vbSystem) <> vbNullString Then
                If OverwriteExisting = True Then
                    Kill Fname
                Else
                    MsgBox "Файл " & Fname & vbCrLf & "существует, отказ в операции"
                    Exit Function
                End If
            End If
            '----------------------------------------------------------------------------------
            ' export itself
              VBE.ActiveVBProject.VBComponents(ModName).Export Fname
              sRes = Fname
   End If
'-----------------------------------------------------------------------
ExitHere:
        If iRes = 2 Then iRes = CheckLoaded(ModName, False)
        ExportModule = sRes '!!!!!!!!!!!!!!!!
        Exit Function
'-------------------------
ErrHandler:
        Err.Clear
        Resume ExitHere
End Function

'======================================================================================================================================================
' Import module from file
'       modType = 1 - vbext_ct_StdModule; 2- vbext_ct_ClassModule;3- vbext_ct_MSForm
'======================================================================================================================================================
Public Function ImportModule(ModName As String, Optional sFile As String = "", Optional sSource As String, _
                                      Optional modType As Integer = 1, Optional OverwriteExisting As Boolean = True) As Boolean
Dim bRes As Boolean
On Error GoTo ErrHandler
'-----------------------------------------------------------------------
    If sFile <> "" Then '  add Via File
        Call AddModFromFile(sFile, ModName, OverwriteExisting)
        bRes = True
    Else                '  add via string
        Call CreateNewModuleFromString(ModName, sSource)
        bRes = True
    End If
'-----------------------------------------------------------------------
ExitHere:
        If bRes Then bRes = CheckLoaded(ModName, False)
        ImportModule = bRes '!!!!!!!!!!!!!!!!
        Exit Function
'-------------------------
ErrHandler:
        ErrPrint2 "ImportModule", Err.Number, Err.Description, MOD_NAME
        Err.Clear: Resume ExitHere
End Function
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
Dim sWork As String
      sWork = VBA1 & VBA2 & VBA3 & VBA4 & VBA5 & VBA6 & VBA7 & VBA8 & VBA9 & VBA10
      sWork = sWork & VBA11 & VBA12 & VBA13 & VBA14 & VBA15 & VBA16 & VBA17 & VBA18 & VBA19
      sWork = sWork & VBA20 & VBA21 & VBA22 & VBA23 & VBA24 & VBA25 & VBA26 & ";" & Chr(34)
      mVBAConst = Split(Replace(sWork, vbCrLf, ""), ";")
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------
' The function will check the correct module
'------------------------------------------------------------------------------------------------------------------------------------------
Private Function IsModuleExist(ModName As String) As Boolean
Dim myComponent As Object
Dim bRes As Boolean

On Error GoTo ErrHandler
'--------------------------
        Set myComponent = VBE.ActiveVBProject.VBComponents(ModName)
        bRes = True
'----------------------------------
ExitHere:
   IsModuleExist = bRes '!!!!!!!!!!!!!!!
   Exit Function
'-------------------
ErrHandler:
   Err.Clear
   Resume ExitHere
End Function

'------------------------------------------------------------------------------------------------------------------------------------------
' The function adds a standard module with a given name
'      modType = 1 - vbext_ct_StdModule; 2- vbext_ct_ClassModule;3- vbext_ct_MSForm
'------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AddModFromString(ModName As String, Optional modType As Integer = 1, Optional OverwriteExisting As Boolean = True, _
                                                                                                        Optional sText As String = "")
Dim myComponent As Object 'VBComponent

On Error GoTo ErrHandler
'--------------------------
If IsModuleExist(ModName) Then
   If OverwriteExisting Then
      RemoveModule ModName
      DoCmd.Save
   Else
      MsgBox "Модуль " & ModName & " существует. Операция добавления прервана"
      Exit Sub
   End If
End If
'--------------------------
Set myComponent = VBE.ActiveVBProject.VBComponents.Add(modType)
    myComponent.Name = ModName
    If sText <> "" Then    ' Insert text
        Call myComponent.CodeModule.InsertLines(1, "Option Explicit")
    End If
'----------------------------------
ExitHere:
   DoCmd.Save
   Exit Sub
'-------------------
ErrHandler:
   Err.Clear
   Resume ExitHere
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------
' The function will remove the module with the given name
'------------------------------------------------------------------------------------------------------------------------------------------
Private Sub RemoveModule(ModName As String)
Dim myComponent As Object 'VBComponent

On Error GoTo ErrHandler
'--------------------------
Set myComponent = VBE.ActiveVBProject.VBComponents(ModName)
    Call VBE.ActiveVBProject.VBComponents.Remove(myComponent)
'----------------------------------
ExitHere:
   DoCmd.Save
   Exit Sub
'-------------------
ErrHandler:
   Err.Clear
   Resume ExitHere
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------
' The function defines the name of the module by the name of the file
'------------------------------------------------------------------------------------------------------------------------------------------
Private Function ExtractModName(sFile As String) As String
Dim sRes As String
Dim iPos As Integer
If sFile = "" Then Exit Function
'------------------------------
sRes = sFile
    iPos = InStrRev(sFile, "\"):    If iPos > 0 Then sRes = Right(sRes, Len(sRes) - iPos)
    iPos = InStrRev(sFile, "."):    If iPos > 0 Then sRes = Left(sFile, iPos - 1)
    iPos = InStr(1, sRes, "_"):     If iPos > 0 Then sRes = Right(sRes, Len(sRes) - iPos)
'------------------------------
    ExtractModName = sRes '!!!!!!!!!!!!!!!!!!!!!1
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
' The function receives the expansion depending on the type
'------------------------------------------------------------------------------------------------------------------------------------------
Private Function FileExtension(vbaMod As Module) As String
Dim sName As String
Dim sRes As String
sName = vbaMod.Name
If Left(sName, 5) = "Form_" Then       ' FORMS
    sRes = "frm"
ElseIf Left(sName, 7) = "Report_" Then ' REPORTS
    sRes = "rep"
Else                                   ' STANDARD MODULE
        Select Case vbaMod.Type
            Case 0:
                sRes = "bas"
            Case 1:
                 sRes = "cls"
            Case Else:
                  sRes = "bas"
        End Select
End If
'----------------------------------------------------------------
        FileExtension = sRes '!!!!!!!!!!!!!!!!!
End Function



'------------------------------------------------------------------------------------------------------------------------------
'Function Serializes VBACode array
'------------------------------------------------------------------------------------------------------------------------------
Private Function VBACodeArrToStr(arrVBA() As VBAMODULE) As String
Dim sRes As String, nDim As Integer, I As Integer, arrSTR() As String
  nDim = UBound(arrVBA): ReDim arrSTR(nDim)
  For I = 0 To nDim
    arrSTR(I) = VBACodeToStr(arrVBA(I))
  Next I
  sRes = Join(arrSTR, Chr(29))
'------------------------------------
VBACodeArrToStr = sRes  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------
'The function restores the VBACode array - from the line
'------------------------------------------------------------------------------------------------------------------------------
Private Function VBACodeArrFromStr(VBAArrStr As String) As Variant 'VBAMODULE()
Dim arrVBA() As VBAMODULE, nDim As Integer, sARR() As String, I As Integer
    sARR = Split(VBAArrStr, Chr(29)): nDim = UBound(sARR): ReDim arrVBA(nDim)
    For I = 0 To nDim
        arrVBA(I) = VBACodeFromStr(sARR(I))
    Next I
'------------------------------------------------------
    VBACodeArrFromStr = sARR  '!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------
'Function Serializes VBACode in line
'------------------------------------------------------------------------------------------------------------------------------
Private Function VBACodeToStr(mVBA As VBAMODULE) As String
Dim PDelim As String
Dim sRes As String
    PDelim = Chr(31)  ' Unit Separator
    sRes = mVBA.Name & PDelim & mVBA.Description & PDelim & mVBA.FullName & PDelim & mVBA.CreateDate & PDelim & mVBA.UpdateDate & PDelim & mVBA.Type & _
                      PDelim & mVBA.fileName & PDelim & mVBA.ParentObjName
    sRes = sRes & PDelim & mVBA.ROWS & PDelim & mVBA.Header & PDelim & mVBA.Author & PDelim & mVBA.Version & _
                      PDelim & mVBA.Functions & PDelim & mVBA.GlobalVars
    sRes = sRes & PDelim & mVBA.code & PDelim & mVBA.HTMLCode & PDelim & mVBA.ObfuscatedCode & PDelim & _
           mVBA.CONSTANTS & PDelim & mVBA.APIs & PDelim & mVBA.VARS
'-----------------------------------------------------------
   VBACodeToStr = sRes '!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------
' The function restores VBACode from the line
'------------------------------------------------------------------------------------------------------------------------------
Private Function VBACodeFromStr(sVBA As String) As VBAMODULE
Dim PDelim As String
Dim mVBA As VBAMODULE, tempArr() As String

tempArr = Split(sVBA, PDelim)
'-----------------------------------
    mVBA.Name = tempArr(0): mVBA.Description = tempArr(1): mVBA.FullName = tempArr(2)
    mVBA.CreateDate = tempArr(3): mVBA.UpdateDate = tempArr(4)
    mVBA.Type = tempArr(5): mVBA.fileName = tempArr(6): mVBA.ParentObjName = tempArr(7)
    mVBA.ROWS = tempArr(8): mVBA.Header = tempArr(9): mVBA.Author = tempArr(10): mVBA.Version = tempArr(11)
    mVBA.Functions = tempArr(12): mVBA.GlobalVars = tempArr(13)
    mVBA.code = tempArr(14): mVBA.HTMLCode = tempArr(15): mVBA.ObfuscatedCode = tempArr(16)
    mVBA.CONSTANTS = tempArr(17): mVBA.APIs = tempArr(18): mVBA.VARS = tempArr(19)
'-------------------------------------------------------------------------------
VBACodeFromStr = mVBA  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------
' Function of parameters of the procedure, retrieved from unnecessary information
'  sCode - all the code of the procedure
'  bToHTML = True - so we parse HTML (this is color coded)
'------------------------------------------------------------------------------------------------------------------------------
Private Function ParceProcedures(ProcName As String, sCode As String, Optional bToHTML As Boolean = False) As VBAMODULE
Dim MyVBA  As VBAMODULE                                         ' Result
Dim CONSTANTS() As String, nConst As Integer, I As Integer      ' Array of constants
Dim ConstList() As Pairs, nPairs As Integer, sWork As String    ' List of constatnts
Dim sWorkCode As String                                         ' Working code
Dim sClearCode As String                                        ' Clear code

On Error GoTo ErrHandle
'---------------------------------
Dim nLines As Integer, nSplitLines As Integer, nComments As Integer, nCodeLines As Integer, sDescript As String
Dim sHeader As String, sVARS As String
'---------------------------------
Const constPrefix As String = "const" '
'*********************************************************************************************************
If sCode = "" Then Exit Function
    MyVBA.AllCodes = sCode                                     ' whole procedure listing
'-------------------------------------------------------------------------
' 1) START UP ON AUXILIARY VALUES -> Receive a code without a warning
    nPairs = -1: ReDim ConstList(0)
    CONSTANTS = GetConstArrayFromCode(sCode)
    
    If CONSTANTS(0) = "" Then                                       ' List of constants not updated
        sWorkCode = sCode
    ' 2) CLEAR CODE FROM COMMENTS
        sClearCode = ProcWorkOut(ProcName, sWorkCode, nLines, nSplitLines, nComments, nCodeLines, sDescript, sHeader, sVARS)
    Else                                                            ' Standards updated, we can add them
        sWorkCode = sCode: nConst = UBound(CONSTANTS)
        For I = 0 To nConst
            nPairs = nPairs + 1: ReDim Preserve ConstList(nPairs)
            ConstList(nPairs).FIRST = CONSTANTS(I)
            ConstList(nPairs).Second = SHT(GetUniqueConstantName(constPrefix, I))
            sWorkCode = Replace(sWorkCode, ConstList(nPairs).FIRST, ConstList(nPairs).Second)     ' Play Submenu
        Next I
        ' 2) CLEAR CODE FROM COMMENTS and GET DESCRIPTION
        sClearCode = ProcWorkOut(ProcName, sWorkCode, nLines, nSplitLines, nComments, nCodeLines, sDescript, sHeader, sVARS)
        For I = 0 To nPairs
            sClearCode = Replace(sClearCode, ConstList(I).Second, ConstList(I).FIRST)
        Next I
    End If
'-------------------------------------------------------------------------
' 3) RESULTS:
    MyVBA.Description = sDescript                             ' Introduced description
    MyVBA.ROWS = nLines                                       ' All line (important)
    MyVBA.Metrics = "Всего строк: " & nLines & ";" & vbCrLf & _
                    "Кол-во комментариев: " & nComments & ";" & vbCrLf & _
                    "Количество разделителей: " & nSplitLines
    MyVBA.Header = sHeader                                    ' Heading (normalized heading)
    MyVBA.ObfuscatedCode = sClearCode                         ' Set of code with elements of operation
'----------------------------------------------------------------------------------
ExitHere:
        ParceProcedures = MyVBA '!!!!!!!!!!!!!!!!!!!!
        Exit Function
'----------------------
ErrHandle:
        ErrPrint2 "ParceProcedure (Proc: " & ProcName & ")", Err.Number, Err.Description, MOD_NAME
        Err.Clear: Resume ExitHere
End Function
'----------------------------------------------------------------------------------------------------------------------------
' Generate Unique ConstantName
'----------------------------------------------------------------------------------------------------------------------------
Private Function GetUniqueConstantName(sPrefix As String, nDim As Integer) As String
         GetUniqueConstantName = GetRandomAlphaString(4, 16, sPrefix) & nDim & Timer '!!!!!!!!!!!!
End Function
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'----------------------------------------------------------------------------------------------------------------------------
' The function replaces all the examples of the constant according to the list with the new setting or if it stops.
'     bOP = True - Menu Values on Keys
'     bOP = False - Menu Keys to Values
'----------------------------------------------------------------------------------------------------------------------------
Private Function ReplaceVars(sCode As String, VarList As cDictionary, Optional bOP As Boolean = True) As String
Dim sRes As String, sKey As String, sValue As String
Dim I As Integer, nDim As Integer
     nDim = VarList.Count - 1
     If nDim < 0 Then Exit Function
     sRes = sCode
For I = 0 To nDim
        sKey = VarList.GetKey(I): sValue = VarList.GetValue(I)
    If bOP Then          ' Set the value on the key
        sRes = Replace(sRes, sValue, sKey)
    Else                 ' Меняем ключи на значения
        sRes = Replace(sRes, sKey, sValue)
    End If
Next I
'----------------------------------------------------------------
    ReplaceVars = sRes '!!!!!!!!!!!!!!!!!!!
End Function
'----------------------------------------------------------------------------------------------------------------------------
' The function will work with the commenters and empty lines;
'   nLines - total number of lines (full text)
'   nSplitLines - number of lines, dividers
'   nComments - number of commented lines (included separately)
'   nCodeLines - string-only code
'   sDescript - description of the function (if it is preceded by the declaration or follows directly after)
'   sHeader - Common Header
'   sVARS - list of enumerated variables
'----------------------------------------------------------------------------------------------------------------------------
Private Function ProcWorkOut(ProcName As String, ByVal sCode As String, ByRef nLines As Integer, ByRef nSplitLines As Integer, _
                               ByRef nComments As Integer, ByRef nCodeLines As Integer, ByRef sDescript As String, _
                               ByRef sHeader As String, ByRef sVARS As String) As String
Dim sRes() As String, nRes As Integer                         ' Возвращаемый результат (очищенный текст)
Dim WorkCode() As String, I As Integer, J As Integer          ' Рабочий код (если многострочный текст)
Dim sWork As String, sComments() As String                    ' Рабочая строка
Dim sMEMO As String, RealCode() As String, nReal As Integer   ' Строка для запоминания, массив для разделения
Dim nFuncDeclaration As Integer, sWork2 As String             ' Флаг того, что найден заголовок функции (номер строки)

Const ComentMark1 As String = "Rem:"                          ' Символ комментария 1
Const ComentMark2 As String = "'"                             ' Символ комментария 2
Const JoinMark = " _"                                         ' Символ объединения строки
Const SplitMark = ": "                                        ' Символ разделения строки

On Error GoTo ErrHandle
'--------------------------------------------------------------------

nLines = 0: nSplitLines = 0: nComments = 0: nCodeLines = -1: sDescript = ""
sMEMO = ""
'---------------------------------------------------------------------
If sCode = "" Then Exit Function
WorkCode = Split(sCode, vbCrLf): nLines = UBound(WorkCode)
For I = 0 To nLines                                                ' ИДЕМ ПО СТРОЧКАМ КОДА
          
          sWork = Trim(WorkCode(I))
          'If ProcName = "cVBA.ParceProcedures" And I = 32 Then
          '   Debug.Print "I = " & I & "  sLine = " & sWork
          '   Debug.Assert False
          'End If
          
          If sWork = "" Then GoTo NextLine                         ' Пропускаем пустые строчки
          If Left(sWork, Len(ComentMark1)) = ComentMark1 Then      ' Первой идет REM - комментарий
              nComments = nComments + 1                            ' Увеличиваем счетчик комментариев (считаем не может быть разделителя)
          ElseIf Left(sWork, Len(ComentMark2)) = ComentMark2 Then  ' Обычная строка-комментарий
          '---------------------------------------------------
              If Not IsStringLine(sWork) Then                      ' Строка не представляет собой разделитель, которые мы отбрасываем
                            nComments = nComments + 1
                            If nFuncDeclaration = 0 Or nCodeLines = 0 Then  ' Декларация еще не найдена, считаем весь текст перед функцией ее описанием
                               If sDescript <> "" Then sDescript = sDescript & vbCrLf
                               sDescript = sDescript & sWork
                            End If
              Else
                            nSplitLines = nSplitLines + 1
              End If
          Else                                                     ' Имеем код
          '---------------------------------------------------
                sComments = VBASPLITCOMMENT(sWork)
                      '-------------------------------------------------------------------------
                      If sComments(0) <> "" Then                   ' Выделена кодовая часть
                                '-------------------------------------------------------------------------
                                ' ПРОВЕРЯЕМ, ЧТО НЕТ ПРИЗНАКА КОНКАТЕНЦИИ. ЕСЛИ ЕСТЬ, ТО ЗАПОМИНАЕМ ДЛЯ ПОСЛЕДУЮЩЕЙ СВЯЗКИ
                                  If Right(sComments(0), Len(JoinMark)) = JoinMark Then        ' (помещаем в sMEMO)
                                         sMEMO = sMEMO & Left(sComments(0), Len(sComments(0)) - Len(JoinMark))
                                  Else    ' ДОБАВЛЕНИЕ КОДА
                                         sWork = sMEMO & sComments(0): sMEMO = ""               ' (извлекаем из накопителя)
                                         '-----------------------------------------------------------------
                                         ' ОБРАБАТЫВАЕМ ЗАГОЛОВОК
                                         If nFuncDeclaration = 0 Then                          ' Еще не обрабатывали заголовок
                                            If IsLineFuncHeader(sWork) Then                    ' Обнаружили заголовок
                                                nCodeLines = nCodeLines + 1
                                                nFuncDeclaration = 1                           ' Пометили строку заголовка
                                                nRes = nRes + 1: ReDim Preserve sRes(nRes)
                                                sRes(nRes) = sWork
                                                sHeader = ProcessFuncHeader(sWork)             ' РАЗОБРАЛИ ЗАГОЛОВОК
                                            End If
                                         '-----------------------------------------------------------------
                                         Else                                                   ' ОБЫЧНЫЙ КОД
                                            ' ПРОВЕРЯЕМ  - НЕТ ЛИ СИМВОЛА РАЗДЕЛЕНИЯ СТРОК (:)
                                            RealCode = ProcessLineSplitters(sWork, SplitMark): nReal = UBound(RealCode)
                                            For J = 0 To nReal
                                                nCodeLines = nCodeLines + 1
                                                nRes = nRes + 1: ReDim Preserve sRes(nRes)
                                                sRes(nRes) = RealCode(J)                         ' Добавляем код
                                                '---------------------------------------------------------------
                                                ' ОБРАБАТЫВАЕМ ВОЗМОЖНЫЕ ПЕРЕМЕННЫЕ
                                                sWork2 = GetVARS(RealCode(J))
                                                If sWork2 <> "" Then
                                                    If sVARS <> "" Then sVARS = sVARS & ";"
                                                       sVARS = sVARS & sWork2
                                                End If
                                            Next J
                                         End If
                                  End If
                      End If
                      If sComments(1) <> "" Then                   ' Обнаружили комментарий
                            nComments = nComments + 1
                      End If
          '---------------------------------------------------
          End If
NextLine:
Next I
'---------------------------------------------------------------------
ExitHere:
    nCodeLines = nCodeLines + 1         ' Обрабатывали код с -1
    ProcWorkOut = Trim(Join(sRes, vbCrLf)) '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
ErrHandle:
    ErrPrint2 "ProcWorkOut (" & ProcName & ")", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'----------------------------------------------------------------------------------------------------------------------------
' The function reads the line and in the case of a divider (:) changes the line masses
'----------------------------------------------------------------------------------------------------------------------------
Private Function ProcessLineSplitters(sCode As String, Optional SplitMark As String = ": ") As String()
    ProcessLineSplitters = Split(sCode, SplitMark) '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'----------------------------------------------------------------------------------------------------------------------------
' Check is Codeline  a header
'----------------------------------------------------------------------------------------------------------------------------
Private Function IsLineFuncHeader(sCodeLine As String) As Boolean
Dim bRes As Boolean
If IsWordInString("End", sCodeLine) > 0 Or IsWordInString("Exit", sCodeLine) > 0 Then Exit Function ' Can't be declaration,
                                                                              
If IsWordInString("Function", sCodeLine) Then
            bRes = True
ElseIf IsWordInString("Sub", sCodeLine) Then
            bRes = True
'-----------------------------------------------------------------------------------------
ElseIf IsWordInString("Property", sCodeLine) Then
       If InStr(1, sCodeLine, "Property Let ") Then
            bRes = True
       ElseIf InStr(1, sCodeLine, "Property Get") Then
            bRes = True
       ElseIf InStr(1, sCodeLine, "Property Set") Then
            bRes = True
       End If
End If
'--------------------------------------------------------------------------------
ExitHere:
      IsLineFuncHeader = bRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'----------------------------------------------------------------------------------------------------------------------------
' The function recognizes the common information from the separator line
'----------------------------------------------------------------------------------------------------------------------------
Private Function IsStringLine(sInput As String) As Boolean
Dim bRes As Boolean, sWork As String
    sWork = Trim(sInput)
    If Left(sWork, 1) <> "'" Then
        Exit Function  ' Строка не закомментирована
    Else
        sWork = Trim(Right(sWork, Len(sWork) - 1))   ' Повторно удаляем пробелы
    End If
'---------------------------------------------------------
' Если добрались сюда, значит строка точно комментарий. Осталось проверить наличие большого числа одинаковых симоволов
        If InStrCount(1, sWork, Left(sWork, 1)) > 10 Then            ' Пробуем первый символ
            bRes = True
        ElseIf InStrCount(2, sWork, Mid(sWork, 2, 1)) > 10 Then      ' На всякий случай пробуем 2-й символ
            bRes = True
        Else
            Exit Function
        End If
'--------------------------------------------------
        IsStringLine = bRes '!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'----------------------------------------------------------------------------------------------------------------------------
' The function calculates the number of inputs from sString2 to sString1
'----------------------------------------------------------------------------------------------------------------------------
Private Function InStrCount(lStart As Long, sString1 As String, sString2 As String) As Long
Dim lPos As Long
lPos = InStr(lStart + lPos, sString1, sString2)
While lPos <> 0
   lPos = InStr(lStart + lPos + Len(sString2) - 1, sString1, sString2)
   InStrCount = InStrCount + 1
Wend
End Function
'----------------------------------------------------------------------------------------------------------------------------
' The function sets the list of constants contained in the code
' Standards all over the place with a measure (by default)
'----------------------------------------------------------------------------------------------------------------------------
Public Function GetConstArrayFromCode(sCode As String, Optional sMarker As String = """") As String()
Dim WorkCode() As String, nDim As Integer, I As Integer    ' Work string with code
Dim sRes() As String, nRes As Integer                      ' Constant array
Dim iL As Long, iR As Long, iLen As Long                   ' Pointers to substring
Dim sWork As String                                        ' Temp string
Dim specChar As String                                     ' replacmenter

ReDim sRes(0): nRes = -1
If sCode = "" Then Exit Function
specChar = Chr(34) & Chr(34) & Chr(34) & Chr(34)           ' Combination """" = "

WorkCode = Split(sCode, vbCrLf): nDim = UBound(WorkCode)   ' Get array of lines
For I = 0 To nDim
    WorkCode(I) = Trim(ClearCode(WorkCode(I)))             ' Clean from comment
    If WorkCode(I) = "" Then GoTo NextLine
    If InStr(1, WorkCode(I), specChar) > 0 Then            ' Looking for combination
       WorkCode(I) = Replace(WorkCode(I), specChar, Chr(2))
       nRes = nRes + 1: ReDim Preserve sRes(nRes)
       sRes(nRes) = specChar
    End If
    
    iL = 1                                                 ' Start search from number 1
    Do While iL > 0
          iL = InStr(iL, WorkCode(I), sMarker)             ' Look for left quate
       If iL > 0 Then
          iR = InStr(iL + 1, WorkCode(I), sMarker)         ' Look for right quate
          If iR > 0 Then
             '-------------------------------------------------------------------------------
             ' SELECT CONSTANTES
             sWork = Mid(WorkCode(I), iL, iR - iL + 1)
             If InStr(1, Join(sRes, Chr(1)), sWork) = 0 Then
                nRes = nRes + 1: ReDim Preserve sRes(nRes)
                sRes(nRes) = Mid(WorkCode(I), iL, iR - iL + 1)
             End If
             iL = iR + 1
             If iL > Len(WorkCode(I)) Then Exit Do
             '-------------------------------------------------------------------------------
          Else                                             ' If can't find right quate
             Exit Do
          End If
       Else                                                ' Can't find
           Exit Do
       End If
    Loop
NextLine:
Next I
'-------------------------------------------------------
ExitHere:
    GetConstArrayFromCode = sRes '!!!!!!!!!!!!!!!!!!!!
End Function
'----------------------------------------------------------------------------------------------------------------------------
' Function on subSttring basis Separates the line s on the side of the dividers.
' The separators are so evolving in the form of Resultat - a mass written in an inert line
' For example, "sdfgAeyuBdsgjhAihdvB", "A;B" --> "sdfg;A;eyu;B;dsgjh;A;ihdv;B"
' The limitation of the function makes the detector unavailable for repetition of the support outside the context -
' select all substring
'----------------------------------------------------------------------------------------------------------------------------
Private Function GetSubStringArray(s As String, subString As String, Optional sDelim As String = ";") As String
Dim sDLMS() As String, nDim As Integer, I As Integer                  ' Array of delimeters
Dim sWork As String                                                   ' Worinf string

If s = "" Or subString = "" Then Exit Function
'------------------------------------------------
sDLMS = Split(subString, sDelim): nDim = UBound(sDLMS)                ' Array of DLMS
sWork = s
For I = 0 To nDim
    sWork = Replace(sWork, sDLMS(I), sDelim & sDLMS(I) & sDelim)
Next I
If Left(sWork, Len(sDelim)) = sDelim Then sWork = Right(sWork, Len(sDelim))
If Right(sWork, Len(sDelim)) = sDelim Then sWork = Left(sWork, Len(sWork) - Len(sDelim))
'-------------------------------------------------
ExitHere:
       GetSubStringArray = sWork '!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'---------------------------------------------------------------------------------------------------------------------------
' The function receives the input line with the code and changes the mass of two parameters: separate 0
' Delivered comment, if present - item 1
'---------------------------------------------------------------------------------------------------------------------------
Private Function VBASPLITCOMMENT(sCode As String) As String()
Dim sRes(1) As String, sWork As String                       ' Results
Dim iL As Long                                               ' String pointer
sWork = Trim(sCode): If sWork = "" Then GoTo ExitHere
            iL = InStr(1, sWork, "'")                        ' Looking comments
            '--------------------------------------------------------------------------
            If iL = 0 Then                                   ' whole line is code
               sRes(0) = sWork
            ElseIf iL = 1 Then                               ' whole line is comment
               sRes(1) = Right(sWork, Len(sWork) - 1)
            Else                                             ' probably there is a comment
               sRes(0) = ClearCode(sWork)
               sRes(1) = GetComnts(sWork, sRes(0)):
               If Len(sRes(1)) > 0 Then sRes(1) = Trim(Right(sRes(1), Len(sRes(1)) - 1))
            End If
'--------------------------------------------
ExitHere:
            VBASPLITCOMMENT = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'----------------------------------------------------------------------------------------------------------------------------
' The function sounds full line and full line
'----------------------------------------------------------------------------------------------------------------------------
Private Function GetComnts(sFull As String, sClearCode As String) As String
Dim sRes As String
Dim iL As Integer
    iL = Len(sClearCode)
    If iL < Len(sFull) Then   ' There is a comment
        sRes = Trim(Right(sFull, Len(sFull) - iL))
    Else
        sRes = ""
    End If
'----------------------------------------------------
ExitHere:
      GetComnts = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'----------------------------------------------------------------------------------------------------------------------------
' Cleaning line from any comments
'----------------------------------------------------------------------------------------------------------------------------
Private Function ClearCode(sCode As String) As String
Dim iL As Long, sRes As String                                        ' Result and commet varkm
Dim LeftSubstring As String, RightSubstring As String                 ' Substrings

Const CommentMarker As String = "'"
    
  
    iL = 1
    Do While iL > 0                                                   ' Go through markers
                iL = InStr(iL, sCode, CommentMarker)
                If iL > 0 Then                                        ' If found check if it is not in substring
                     LeftSubstring = Left(sCode, iL - 1)              ' select substring
                     RightSubstring = Right(sCode, Len(sCode) - iL)
                     '-------------------------------
                     If InStr(1, RightSubstring, Chr(34)) > 0 Then
                        
                        ' Check if marker is not inside substring
                        's = a + " 'hjgj'" 'mnmn "
                        If GetSinglBalance(LeftSubstring, Chr(34), Chr(34)) > 0 Then ' Can't find ballance from left
                                                                      ' But right hold quate
                                iL = iL + 1                           ' Find new marker
                                GoTo NextMarker
                        End If
                        ' This is correct comment
                        sRes = Trim(LeftSubstring)
                        GoTo ExitHere
                     Else                                             ' If not any code - whole right part is comment
                        sRes = Trim(LeftSubstring)
                        GoTo ExitHere
                     End If
                     '--------------------------------
                End If
NextMarker:
    Loop
    If sRes = "" Then sRes = sCode                                    ' No found comment marker, return the whole line
'---------------------------------------------------------------------------
ExitHere:
    ClearCode = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'-----------------------------------------------------------------------------------------------------------------
' Check the brackets ballance
'-----------------------------------------------------------------------------------------------------------------
Private Function GetSinglBalance(s As String, _
                                   Optional LeftTag As String = "(", Optional RightTag As String = ")") As Integer
Dim BL As Integer                                ' Current ballance
Dim iL As Long, iLen As Long                     ' Left pointer and lenght
Dim iR As Long, lRPL As String, rRPL As String   ' Right pointer and sime replacements
Dim sWork As String                              ' Working var
Const lReplcmnt As String = "X"
Const rReplcmnt As String = "Y"

'---------------------------------------------------------
' Генерим подмены
lRPL = String(Len(LeftTag), lReplcmnt)
rRPL = String(Len(RightTag), rReplcmnt)
'--------------------------------------------------------
sWork = s: iL = 1: BL = 0: iLen = Len(sWork)
    Do While iL > 0
        iL = InStr(iL, sWork, LeftTag)
        If iL = 0 Then Exit Do
        BL = BL + 1: sWork = ReplaceSymbol(sWork, iL, lReplcmnt)
        iL = iL + 1: If iL > iLen Then Exit Do
            iR = InStr(iL, sWork, RightTag)
            If iR > 0 Then
                BL = BL - 1: sWork = ReplaceSymbol(sWork, iR, rReplcmnt)
            End If
        iL = iL + 1
     Loop
'------------------------------------------------------------------------
             GetSinglBalance = BL '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'---------------------------------------------------------------------------------------------------------------------------
' The function replaces the symbol 1, which is used in the given position in the line with the symbol 2
'---------------------------------------------------------------------------------------------------------------------------
Private Function ReplaceSymbol(s As String, POS As Long, NewSymb As String) As String
Dim sRes As String
     If POS < 1 Or POS > Len(s) Then
         sRes = s
         GoTo ExitHere
     End If
'------------------------------------------
     sRes = Left(s, POS - 1) & NewSymb & Right(s, Len(s) - POS)
'------------------------------------------
ExitHere:
     ReplaceSymbol = sRes '!!!!!!!!!!!
End Function
'-----------------------------------------------------------------------------------------------------------------------
' Function generates a random line according to the given rules
'-----------------------------------------------------------------------------------------------------------------------
Private Function GetRandomAlphaString(Optional minLen As Integer = 4, Optional maxLen = 16, _
                                                                            Optional sPrefix As String = "") As String
Dim FileNameLen As Integer, I As Integer
Dim sRes As String, iLen As Integer
'--------------------------------------------------------------
Const MASAR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
 Call Randomize: iLen = Len(MASAR)
 FileNameLen = CInt(Int((maxLen * Rnd()) + minLen))    ' lenght 4 to 16
     sRes = Mid(MASAR, CInt(Int(((iLen - 10) * Rnd()) + 1)), 1)   ' 1-position is letter
 For I = 1 To FileNameLen                                         ' go throgh iteration
     sRes = sRes & Mid(MASAR, CInt(Int((iLen * Rnd()) + 1)), 1)
 Next I
'--------------------------------------------------------------
     GetRandomAlphaString = sPrefix & LCase(sRes)
End Function







'======================================================================================================================================================
' List Public Function and Procedures in Module
'======================================================================================================================================================
Public Function ListProcedures(Optional sModName As String = "#_ENVIRONMENT", Optional sProjectName As String = "_GRACKLE", _
                                                                       Optional bOnlyPublic As Boolean = True, Optional DLM As String = ";") As String
Dim VBPrj As Object
Dim VBComp As Object, CodeMod As Object, FRES As TFUNCTION
Dim ProcName As String, LineNum As Long
Dim LINES() As String, nDim As Integer, I As Integer, sRes() As String, iRes As Integer, sWork As String

    On Error GoTo ErrHandle
'------------------------------------
iRes = -1: ReDim Preserve sRes(0)

Set VBPrj = Application.VBE.VBProjects(sProjectName)
Set VBComp = VBPrj.VBComponents(sModName)
Set CodeMod = VBComp.CodeModule


With CodeMod
        sWork = .LINES(1, .CountOfLines): If sWork = "" Then Exit Function
        sWork = Replace(sWork, "_" & vbCrLf, " ")           ' Remove Any Row breaks
        LINES = Split(sWork, vbCrLf): nDim = UBound(LINES)
        For I = 0 To nDim
            sWork = Trim(LINES(I))
            If IsFunctionHeader(sWork) Then
                FRES = ParseFunctionHeader(sWork)
                If FRES.Name = "" Then GoTo NextLine
                If bOnlyPublic And FRES.iPrivacy > 0 Then GoTo NextLine
                iRes = iRes + 1: ReDim Preserve sRes(iRes)
                sRes(iRes) = FRES.Name
            End If
NextLine:
        Next I
End With
If iRes > 1 Then BubbleSort sRes
'------------------------------------
ExitHere:
    ListProcedures = Join(sRes, DLM)
    Set CodeMod = Nothing: Set VBComp = Nothing: Set VBPrj = Nothing
    Exit Function
'-----------
ErrHandle:
    ErrPrint2 "ListProcedures", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'======================================================================================================================================================
' The function is received for the preset module of its declaration (if it is measured before the first time)
'======================================================================================================================================================
Public Function GetModuleDeclaration(ModName As String) As String
Dim iRes As Integer                                     ' Result
Dim mdl As Module
Dim lngCount As Long, lngCountDecl As Long, lngI As Long
Dim strDeclaration As String

On Error Resume Next
'-----------------------------------------------------------------------
   iRes = CheckLoaded(ModName)
'------------------------------------------------------------------------
   If (iRes = 1) Or (iRes = 2) Then
        Set mdl = Modules(ModName): lngCount = mdl.CountOfLines    ' Open the module and read the total number of lines
        lngCountDecl = mdl.CountOfDeclarationLines                 ' Read the number of declaration lines
        strDeclaration = mdl.LINES(1, lngCountDecl)                ' Declaration's lines
   End If
'------------------------------------------------------------------------
ExitHere:
        GetModuleDeclaration = strDeclaration '!!!!!!!!!!!
End Function

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Parcing finction header. This is an alternative implementation of parsing the function header according to the VB7 spec
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function ParseFunctionHeader(sFuncHeader As String) As TFUNCTION
Dim WORDS() As String, nWords As Integer, I As Integer, IsDeclare As Boolean
Dim FRES As TFUNCTION, IsProperty As Boolean, IsNameReady As Boolean, iL As Integer

Const DLM As String = ";"


   On Error Resume Next
'---------------------
    If sFuncHeader = "" Then Exit Function
    WORDS = Split(sFuncHeader, " "): nWords = UBound(WORDS)

        For I = 0 To nWords
            Select Case WORDS(I)
            Case "Public":
                  FRES.iPrivacy = 0
            Case "Private":
                  FRES.iPrivacy = 1
            Case "Fiend":
                  FRES.iPrivacy = 2
            Case "Declare":
                  IsDeclare = True
            Case "Static":
                  FRES.IsStatic = True
            Case "Function", "Sub":
                  If WORDS(I) = "Function" Then
                       FRES.ReturnType = GetReturnType(sFuncHeader)
                       If IsDeclare Then
                            FRES.TheType = 6
                       Else
                            FRES.TheType = 1
                       End If
                  Else
                       If IsDeclare Then
                            FRES.TheType = 5
                       Else
                            FRES.TheType = 0
                       End If
                  End If
                  IsNameReady = True
            Case "Property":
                  IsProperty = True
            Case "Let":
                  If IsProperty Then FRES.TheType = 3
                  IsNameReady = True
            Case "Get":
                  If IsProperty Then FRES.TheType = 2
                  FRES.ReturnType = GetReturnType(sFuncHeader)
                  IsNameReady = True
            Case "Set":
                  If IsProperty Then FRES.TheType = 4
                  IsNameReady = True
            Case Else:
                 If IsNameReady Then
                       iL = InStr(1, WORDS(I), "(")
                       If iL > 0 Then
                            FRES.Name = Left(WORDS(I), iL - 1)
                       Else
                            FRES.Name = WORDS(I)
                       End If
                       IsNameReady = False
                 End If
            End Select
         Next I
    FRES.ArgumentList = GetFuncArgList(sFuncHeader)
'---------------------
ExitHere:
    ParseFunctionHeader = FRES '!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Get ReturnType
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetReturnType(sFuncHeader As String) As String
Dim sRes As String, iL As Integer, sWork As String, iR As Integer
Dim sArrXT As String

    On Error GoTo ErrHandle
'----------------------
If sFuncHeader = "" Then Exit Function
sWork = Trim(sFuncHeader)
If Right(sWork, 2) = "()" Then
    sWork = Left(sWork, Len(sWork) - 2): sArrXT = "()"
End If
iL = InStrRev(sWork, ")"): If iL = 0 Then Exit Function
iL = InStr(iL + 1, sWork, "As")
If iL > 0 Then sRes = Trim(Right(sWork, Len(sWork) - iL - 1))
'----------------------
ExitHere:
    GetReturnType = sRes & sArrXT '!!!!!!!!!!!!!!!!
    Exit Function
'--------
ErrHandle:
    ErrPrint2 "GetReturnType", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Get Argument List
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetFuncArgList(sFuncHeader As String) As String
Dim sRes As String, iL As Integer, sWork As String, iR As Integer

    On Error GoTo ErrHandle
'---------------------
If sFuncHeader = "" Then Exit Function
sWork = Trim(sFuncHeader)
If Right(sWork, 2) = "()" Then sWork = Left(sWork, Len(sWork) - 2)

iL = InStr(1, sWork, "("): If iL = 0 Then Exit Function
iR = InStrRev(sWork, ")"): If iR = 0 Then Exit Function
sRes = Mid(sWork, iL + 1, iR - iL - 1)
'---------------------
ExitHere:
    GetFuncArgList = sRes '!!!!!!!!!!!
    Exit Function
'-----------
ErrHandle:
    ErrPrint2 "GetFuncArgList", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Check if this line declare the function
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function IsFunctionHeader(sLine As String) As Boolean
Dim WORDS() As String, nWords As Integer, sWork As String

Const DLM As String = ";"
Const KEY_WORDS As String = "Function;Sub;Property"
Const IGNOR_LIST As String = "Friend;Static;Declare;PtrSafe"

    On Error Resume Next
'---------------------
If sLine = "" Then Exit Function
If Left(sLine, 1) = Chr(39) Then Exit Function

sWork = Trim(FiltreString(sLine, Split(IGNOR_LIST, DLM)))

WORDS = Split(sWork, " "): nWords = UBound(WORDS)
If nWords < 1 Then Exit Function
    Select Case WORDS(0)
    Case "Sub", "Function", "Property":   ' The Public Function/Sub without modificators

    Case "Public", "Private":             ' Could be function
           If nWords < 2 Then Exit Function
           If Not InList(KEY_WORDS, WORDS(1), DLM) Then Exit Function
    Case Else:
        Exit Function
    End Select
'---------------------
ExitHere:
    IsFunctionHeader = True '!!!!!!!!!!!
End Function

'--------------------------------------------------------------------------------------------------------------------------------
' The function replaces all separate words
'--------------------------------------------------------------------------------------------------------------------------------
Private Function ReplaceWord(ByVal expression As String, ByVal Find As String, ByVal Replcmnt As String, _
                            Optional Compare As VbCompareMethod = vbBinaryCompare) As String
Dim ExpressionLen As Long, FindLen As Long, iL As Long, iPrev As Integer
Dim PartOfExpression As String, sRes As String

iL = 1: PartOfExpression = ""
ExpressionLen = Len(expression): FindLen = Len(Find)
'--------------------------------------------------------------------
If FindLen <= ExpressionLen And FindLen > 0 Then
    Do While (iL <= ExpressionLen)
        iL = InStrWord(iL, expression, Find, Compare)
        '------------------------------------------------------------
        If iL > 0 Then                    ' Found sothing
           If iL = 1 Then
               sRes = sRes & Replcmnt
               iPrev = iL
           ElseIf iPrev > 0 Then
               PartOfExpression = Mid(expression, iPrev + FindLen, iL - iPrev - FindLen)
               sRes = sRes & PartOfExpression & Replcmnt
               iPrev = iL
           End If
        Else                             ' Nothing else
               PartOfExpression = Right(expression, ExpressionLen - iPrev - FindLen + 1)
               sRes = sRes & PartOfExpression
               Exit Do
        End If
            iL = iL + FindLen
    Loop
'---------------------------------------------------------------------
End If 'if string less than ..
'----------------------------------------------------------------------
ExitHere:
        sRes = IIf(FindLen = 0 Or expression = "", expression, sRes)
        ReplaceWord = sRes   '!!!!!!!
End Function
'--------------------------------------------------------------------------------------------------------------------------------
' The function is determined by either letter or number. In this case, it is called True. First False
'--------------------------------------------------------------------------------------------------------------------------------
Private Function IsAlphaNumeric(sChar As String) As Boolean
    If sChar = "" Then Exit Function
    If Asc(sChar) < 65 And Not IsNumeric(sChar) Then Exit Function
    IsAlphaNumeric = True '!!!!!!!!!!!!!!
End Function
'----------------------------------------------------------------------------------------------------------------------------
' The function will form html presetting the code, depending on the specific type
'----------------------------------------------------------------------------------------------------------------------------
Private Function GetHTMLCode(sCode As String, Optional sComment As String = "", Optional nType As Integer = 0) As String
Dim sRes As String, sWprk As String                     ' retutning string with html code
Dim nDim As Integer, I As Integer                       ' Keywords dimention
Const bgColor As String = "#ffffff"
Const frColor As String = "#000000"
Const cmntColor As String = "#696969"
Const KeyWordColor As String = "#800000"
Const constColor As String = "#0000e6"

           Select Case nType
              Case 0:               ' HTML initialization
                  sRes = "<pre style=" + Chr(39) + "color:" + frColor & ";background:" + frColor + ";" + Chr(39) + _
                         "><html><body style=" & Chr(39) + "color:" + frColor + "; background:" & frColor + "; " + Chr(39) + "><pre>"
              Case 1:               ' Add empty row
                  sRes = sCode
              Case 2:               ' Divider
                  sRes = "<span style=" + Chr(39) + "color:" + cmntColor + "; " + Chr(39) + ">" + sCode + _
                         "</span>"
              Case 3:               ' regular code
                nDim = UBound(mVBAConst): sRes = sCode
                For I = 0 To nDim
                      If InStr(1, sCode, mVBAConst(I)) Then
                          sWprk = "<span style=" + Chr(39) + "color:" + KeyWordColor + "; font-weight:bold; " + Chr(39) + ">" + _
                                 mVBAConst(I) + "</span>"
                          sRes = ReplaceWord(sRes, mVBAConst(I), sWprk)
                      End If
                Next I
                If sComment <> "" Then
                    sRes = sRes + "<span style=" + Chr(39) + "color:" + cmntColor + "; " + Chr(39) + ">" + _
                               sComment + "</span>"
                End If
              Case 4:             ' constant highlight
                    sRes = "<span style=" + Chr(39) + "color:" + constColor + "; " + Chr(39) + ">" + _
                           sCode + "</span>"
              Case 5:             ' stop
                    sRes = "</pre></body></html>"
           End Select
'-----------------------------------------------
ExitHere:
       GetHTMLCode = sRes '!!!!!!!!!!!!!!!!!!!!
End Function
'---------------------------------------------------------------------------------------------------------------------------
' The function is set to the default table with a defined structure information about the modules
'---------------------------------------------------------------------------------------------------------------------------
Private Sub InsertModInfo(pthToDB As String, TableName As String)
Dim sSQL As String

On Error GoTo ErrHandle

      If pthToDB = "" Then Exit Sub
      DoCmd.SetWarnings False
'-------------------------------------------------------------------------------------------------
             sSQL = " INSERT INTO " & TableName & " ( Id, DBName, DBPath, ModName, DateCreate, DateUpdate ) " & _
                   "SELECT MSysObjects.Id,  " & sCH(FileNameOnly(pthToDB)) & " As DBName, " & sCH(pthToDB) & " As DBPath, " & _
                   "MSysObjects.Name, MSysObjects.DateCreate, MSysObjects.DateUpdate " & _
                   "FROM MSysObjects IN " & sCH(pthToDB) & " " & _
                   "WHERE (((MSysObjects.Type) = -32761)) ORDER BY MSysObjects.Name;"
             CurrentDb.Execute sSQL     ' Try to get result
'------------------------------------------
ExitHere:
      DoCmd.SetWarnings True
      Exit Sub
'--------------------
ErrHandle:
      Debug.Print "ERR#" & Err.Number & vbCrLf & Err.Description
End Sub
'---------------------------------------------------------------------------------------------------------------------------
' Effect to rotate line
'---------------------------------------------------------------------------------------------------------------------------
Private Function RotateLine(sChar As String) As String
Dim sWork As String, sRes As String
    sWork = Trim(sChar):
    Select Case sWork
    Case "/":
         sRes = "-"
    Case "-":
         sRes = "\"
    Case "|":
         sRes = "/"
    Case "\"
         sRes = "|"
    End Select
'----------------------------------
ExitHere:
    RotateLine = sRes & String(10, " ")  '!!!!!!!!!!!!!!
End Function
'-------------------------------------------------------------------------------------------------------------------------------
' Create new table
'-------------------------------------------------------------------------------------------------------------------------------
Private Function CreateTableVBA(sTableName As String) As Boolean
On Error GoTo ErrHandle
  
DoCmd.SetWarnings False
  CurrentDb.Execute "CREATE TABLE " & SHT(sTableName) & "(ID COUNTER(1, 1) PRIMARY KEY, " & _
                    "CreateDate DATETIME, HASH VARCHAR(255), ModName VARCHAR(100)," & _
                    "ProcName CHAR, Publ CHAR, FuncType CHAR,Header CHAR, Source MEMO, " & _
                    "LineCount LONG, Metrics CHAR,Description MEMO, HTML MEMO, Version CHAR,Author CHAR, Flag YESNO);", dbFailOnError
DoCmd.SetWarnings True
  CurrentDb.TableDefs.Refresh
'------------------------------------
ExitHere:
     CreateTableVBA = True '!!!!!!!!!!!!!
     Exit Function
'------------------
ErrHandle:
     Debug.Print String(40, "#") & vbCrLf & "ERR#" & Err.Number & vbCrLf & Err.Description & vbCrLf & String(40, "#")
     Err.Clear
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
' Check if table exists
'-----------------------------------------------------------------------------------------------------------------------------------
Private Function TableExists(strTableName As String) As Boolean
On Error Resume Next
    TableExists = IsObject(CurrentDb.TableDefs(strTableName)) '!!!!!!!!!
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
' The function for the preset module and the function in the VBA table returns a line, if they input - ID and 0, output
'-----------------------------------------------------------------------------------------------------------------------------------
Private Function GetRecordID(sTableName As String, ModName As String, ProcName As String) As Long
    GetRecordID = Nz(DLookup("ID", SHT(sTableName), "(ModName=" & Chr(39) & ModName & _
                              Chr(39) & ") AND (ProcName=" & Chr(39) & ProcName & Chr(39) & ")"), 0)
End Function
'-----------------------------------------------------------------------------------------------------------------------------------
' The function will add the record or update the record in the VBA tab, change the ID
'-----------------------------------------------------------------------------------------------------------------------------------
Private Function WriteToVBATAble(sTableName As String, ModName As String, ProcName As String, _
                     Optional PUBL As String, Optional FuncType As String, Optional sHeader As String, Optional sSource As String, _
                     Optional LineCount As Long, Optional sMetrics As String, Optional sDescription As String, Optional sHTML As String, _
                     Optional sVersion As String, Optional sAuthor As String)
Dim id As Integer, sHASH As String
Dim sSQL1 As String, sSQL2 As String

On Error GoTo ErrHandle
If Not TableExists(sTableName) Then Exit Function
'-----------------------------------------------------------------------------------
id = GetRecordID(sTableName, ModName, ProcName)
If id = 0 Then       ' First record insert
    sHASH = GetHASH("VBA" & ModName & ProcName)
    sSQL1 = "INSERT INTO " & SHT(sTableName) & " (HASH,CreateDate,ModName,ProcName"
    sSQL2 = " VALUES (" & sCH(sHASH) & "," & SQLDate(Now()) & "," & sCH(ModName) & "," & sCH(ProcName)
    If PUBL <> "" Then
       sSQL1 = sSQL1 & ",Publ": sSQL2 = sSQL2 & "," & sCH(PUBL)
    End If
    If FuncType <> "" Then
       sSQL1 = sSQL1 & ",FuncType": sSQL2 = sSQL2 & "," & sCH(FuncType)
    End If
    If sHeader <> "" Then
       sSQL1 = sSQL1 & ",Header": sSQL2 = sSQL2 & "," & SQLText2(sHeader)
    End If
    If sSource <> "" Then
       sSQL1 = sSQL1 & ",Source": sSQL2 = sSQL2 & "," & SQLText2(sSource)
    End If
    If LineCount > 0 Then
       sSQL1 = sSQL1 & ",LineCount": sSQL2 = sSQL2 & "," & LineCount
    End If
    If sMetrics <> "" Then
       sSQL1 = sSQL1 & ",Metrics": sSQL2 = sSQL2 & "," & SQLText2(sMetrics)
    End If
    If sDescription <> "" Then
        sSQL1 = sSQL1 & ",Description": sSQL2 = sSQL2 & "," & SQLText2(sDescription)
    End If
    If sVersion <> "" Then
       sSQL1 = sSQL1 & ",Version": sSQL2 = sSQL2 & "," & SQLText2(sVersion)
    End If
    If sAuthor <> "" Then
        sSQL1 = sSQL1 & ",Author": sSQL2 = sSQL2 & "," & SQLText2(sAuthor)
    End If
    sSQL1 = sSQL1 & ") " & sSQL2 & ");"
Else                 ' Update
    sSQL1 = "UPDATE " & SHT(sTableName) & " SET ProcName =" & SQLText2(ProcName): sSQL2 = " WHERE (ID=" & id & ");"
    If PUBL <> "" Then sSQL1 = sSQL1 & ", Publ =" & sCH(PUBL)
    
    If FuncType <> "" Then sSQL1 = sSQL1 & ",FuncType=" & sCH(FuncType)
    If sHeader <> "" Then sSQL1 = sSQL1 & ",Header=" & SQLText2(sHeader)
    If sSource <> "" Then sSQL1 = sSQL1 & ",Source=" & SQLText2(sSource)
    If LineCount > 0 Then sSQL1 = sSQL1 & ",LineCount=" & LineCount
    If sMetrics <> "" Then sSQL1 = sSQL1 & ",Metrics=" & SQLText2(sMetrics)
    If sDescription <> "" Then sSQL1 = sSQL1 & ",Description=" & SQLText2(sDescription)
    If sVersion <> "" Then sSQL1 = sSQL1 & ",Version=" & SQLText2(sVersion)
    If sAuthor <> "" Then sSQL1 = sSQL1 & ",Author=" & SQLText2(sAuthor)
    
    sSQL1 = sSQL1 & sSQL2
End If
'-------------------------------------------------------------------------------
DoCmd.SetWarnings False
    CurrentDb.Execute sSQL1
DoCmd.SetWarnings True
'---------------------------------------------------------------
    If (id = 0) And sHASH <> "" Then id = GetRecordID(sTableName, ModName, ProcName)
'----------------------------
ExitHere:
     WriteToVBATAble = id '!!!!!!!!!!!!!!
     Exit Function
'------------
ErrHandle:
     Debug.Print String(40, "#") & vbCrLf & "ERR#" & Err.Number & vbCrLf & Err.Description & vbCrLf & String(40, "#")
     Err.Clear
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------------
' Function Extract Value for Key.
'-----------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetStringPart(sSource As String, sKey As String, Optional sSeparatorList As String = ",;") As String
Dim iL As Integer, sLD As String
Dim I As Integer
     iL = InStr(1, sSource, sKey, vbTextCompare)
     If iL > 0 Then
          sLD = Right(sSource, Len(sSource) - (iL + Len(sKey)) + 1)
          '--------------------------------------------------------
          ' Check if some separators
          For I = 1 To Len(sSeparatorList)
                iL = InStr(1, sLD, Mid(sSeparatorList, I, 1))
                If iL > 0 Then
                     sLD = Left(sLD, iL - 1)
                End If
          Next I
          '--------------------------------------------------------
          ' GETTING VALUE
          sLD = Trim(sLD)
     End If
'------------------------------------------------
ExitHere:
     GetStringPart = sLD '!!!!!!!!!!!!!
End Function

'-------------------------------------------------------------------------------------------------------------------------------------
' Function shuts down variable declarations if they are
' The function will immediately check the acquisition, if this is the case, but
' in the line [VAR_NAME],[VAR_TYPE],[CONST/PUBLIC/PRIVATE/DIM] - as declarated,[VALUE] - only for constants
'          i.e.      ;[VAR_NAME],[VAR_TYPE],...
'-------------------------------------------------------------------------------------------------------------------------------------
Private Function GetVARS(sCode As String, Optional sDLM1 As String = ",", Optional sDLM2 As String = ";") As String
Dim sRes As String, iL As Integer                       ' Results and string pointer
Dim OBYAVA As String, sWork As String
Dim MyArr() As String, nArr As Integer, I As Integer    ' Array of parameters
Dim sName As String, sType As String, sValue As String  ' Working string
'Const i As Integer = 1, n As Integer = 5
'--------------------------------------------------------------------------------------------------------------
' Проверяем, что есть признак
       If InStr(1, sCode, "Function") Or InStr(1, sCode, "Sub") Or InStr(1, sCode, "Property") Then ' Fail
            sRes = "": GoTo ExitHere
       ElseIf InStr(1, sCode, "Type") Or InStr(1, sCode, "Enum") Then                               ' Fail
            sRes = "": GoTo ExitHere
       ElseIf InStr(1, sCode, "Const") Then                                                         ' Fail
            iL = InStr(1, sCode, "Const")
            If iL = 1 Then
                OBYAVA = "CONST"
                sRes = Mid(sCode, iL + Len("Const") + 1, Len(sCode) - Len("Public") - iL + 1)
            End If
       ElseIf InStr(1, sCode, "Private") Then                                                      ' May be module
            iL = InStr(1, sCode, "Private"): OBYAVA = "PRIVATE"
            sRes = Mid(sCode, iL + Len("Private") + 1, Len(sCode) - Len("Private") - iL + 1)
       ElseIf InStr(1, sCode, "Public") Then
            iL = InStr(1, sCode, "Public"): OBYAVA = "PUBLIC"
            sRes = Mid(sCode, iL + Len("Public") + 1, Len(sCode) - Len("Public") - iL + 1)
       ElseIf InStr(1, sCode, "Dim") Then
            iL = InStr(1, sCode, "Dim"): OBYAVA = "DIM"
            sRes = Mid(sCode, iL + Len("Dim") + 1, Len(sCode) - Len("Dim") - iL + 1)
       Else
            sRes = ""
       End If
'------------------------------------------------------------------------------------------------
If sRes <> "" Then            ' param array
       MyArr = Split(sRes, sDLM1): nArr = UBound(MyArr)
       sRes = ""
       For I = 0 To nArr
            If OBYAVA = "CONST" Then
               If InStr(1, MyArr(I), "=") Then
                    sValue = Trim(Split(MyArr(I), "=")(1))
                    MyArr(I) = Trim(Split(MyArr(I), "=")(0))
               End If
            End If
            '--------------------------------------------------
            If InStr(1, MyArr(I), "As") Then
                sName = Trim(Split(MyArr(I), "As")(0))
                sType = Trim(Split(MyArr(I), "As")(1))
            Else
                sName = Trim(MyArr(I))
                sType = "Variant"
            End If
            sRes = sRes & sDLM2 & sName & sDLM1 & sType & sDLM1 & OBYAVA & sDLM1 & sValue
            sName = "": sType = "": sValue = ""
       Next I
End If
If Left(sRes, Len(sDLM2)) = sDLM2 Then sRes = Right(sRes, Len(sRes) - Len(sDLM2))
'-------------------------------------------------------------------------------------------------
ExitHere:
       GetVARS = sRes '!!!!!!!!!!!!!!!!!
End Function
'--------------------------------------------------------------------------------------------------------------------------------
' The function allows you to search for a specific word (in some other words). In case of success - evokes - position
'--------------------------------------------------------------------------------------------------------------------------------
Private Function InStrWord(Start As Long, expression As String, Find As String, _
                                                        Optional Compare As VbCompareMethod = vbBinaryCompare) As Long
Dim iL As Long, iRes As Long                                  ' String Pointer
Dim chtLeft As String, chtRight As String                     ' Chars around

If expression = "" Or Find = "" Then Exit Function
If Start < 1 Or Start > Len(expression) Then Exit Function
'----------------------------------------------------------------------------------------
        iL = Start
    Do While iL > 0
        iL = InStr(iL, expression, Find, Compare)
        If iL = 0 Then Exit Do
        If iL <= 1 Then
           chtLeft = ""
        Else
           chtLeft = Mid(expression, iL - 1, 1)
        End If
        If iL + Len(Find) > Len(expression) Then
           chtRight = ""
        Else
           chtRight = Mid(expression, iL + Len(Find), 1)
        End If
'-----------------------------------------------------------------
        If IsAlphaNumeric(chtLeft) Or IsAlphaNumeric(chtRight) Then
            iL = iL + 1: If iL > Len(expression) Then Exit Do
        Else
            iRes = iL
            Exit Do
        End If
     Loop
'---------------------------------------------------
    InStrWord = iRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function

'---------------------------------------------------------------------------------------------------------------------------
' Function Selects the first line of the function. If the liquid is burnt, just heading (full)
' Calls up the line of view:
' Name;Function Type (function case);Private/Public;
'         Sub/Function/Property Get/Property Let/Property Set;Var1 as String, Var2 as Long;Full Declaraion Line
'---------------------------------------------------------------------------------------------------------------------------
Private Function ProcessFuncHeader(sHeader As String, Optional sDelim As String = ";") As String
Dim sRes As String

If InStr(1, sHeader, "Function") Then
    sRes = StripHeader(sHeader, "Function", sDelim)
ElseIf InStr(1, sHeader, "Sub") Then
    sRes = StripHeader(sHeader, "Sub", sDelim)
'-----------------------------------------------------------------------------------------
ElseIf InStr(1, sHeader, "Property") Then
       If InStr(1, sHeader, "Property Let") Then
            sRes = StripHeader(sHeader, "Property Let", sDelim)
       ElseIf InStr(1, sHeader, "Property Get") Then
            sRes = StripHeader(sHeader, "Property Get", sDelim)
       ElseIf InStr(1, sHeader, "Property Set") Then
            sRes = StripHeader(sHeader, "Property Set", sDelim)
       Else
            sRes = "": GoTo ExitHere
       End If
'-----------------------------------------------------------------------------------------
Else
    sRes = "": GoTo ExitHere
End If
'--------------------------------------------------------------------------------
ExitHere:
      ProcessFuncHeader = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'--------------------------------------------------------------------------------------------------------------------------
' Assisted Header Discharge Function. Returns the Prepared Result
'--------------------------------------------------------------------------------------------------------------------------
Private Function StripHeader(sHeader As String, sSubType As String, Optional sDelim As String = ";") As String
Dim sFuncName As String, LSubType As Integer        ' Func Name and Class (Sub/ Function / Property)
Dim sFuncType As String, sPublicType As String      ' Returning Type (String, Integer, Boolean), тип открытости (Public/Private)
Dim sVARS As String                                 ' Its parameters
Dim sFullDeclaration As String                      ' Full declaration line
Dim sRes As String, iL As Integer, iR As Integer    ' Results
'-----------------------------------------------------------
    sFullDeclaration = sHeader: LSubType = Len(sSubType)
    iL = InStr(1, sHeader, sSubType)
    sPublicType = IIf(iL = 1, "Public", IIf(InStr(1, sHeader, "Private"), "Private", "Public"))
    iR = InStr(1, sHeader, "(")
    If iR > 0 Then
           sFuncName = Trim(Mid(sHeader, iL + LSubType, iR - iL - LSubType))
           iL = InStr(1, sHeader, ")")
           sVARS = Mid(sHeader, iR + 1, iL - iR - 1)
           iL = InStr(iL, sHeader, "As")
              If sSubType = "Function" Or sSubType = "Property Get" Then
                       sFuncType = IIf(iL > 0, Trim(Right(sHeader, Len(sHeader) - iL - 1)), "Variant")
              Else
                       sFuncType = ""
              End If
    Else
       sRes = "": GoTo ExitHere    ' Fail
    End If
'---------------------------------
       sRes = sFuncName & sDelim & sFuncType & sDelim & sPublicType & sDelim & _
                      sSubType & sDelim & sVARS & sDelim & sFullDeclaration
'--------------------------------------------------------------------------------
ExitHere:
      StripHeader = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' The function will form the title of the module or the function
' For this function, it receives the text of the first (!) Lines,
' The function of sam does not read the code, but if it analyzes it, it changes the type:
' MODULE_NAME;DESCRIPTION;AUTHOR;COPYRIGHT;VERSION;DATE;DEPENDENCY
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetHeader(Optional ModName As String, Optional FuncName As String, Optional sSource As String, _
                                                                                       Optional DLM As String = ";") As String
Dim sRes As String
Dim sHeader() As String, nDim As Integer, I As Integer, iL As Integer, iR As Integer
Dim sName As String, sDescription As String, sAuthor As String, sVersion As String, sCopyright As String, sDependency As String

'---------------------------------------------------------------
If sSource = "" Then Exit Function
    sHeader = Split(sSource, vbCrLf): nDim = UBound(sHeader)
'-----------------------------------------------------------------
' JUST HEADER OF MODULE
If FuncName = "" Then ' ANALYSE MODULE HEADER
    For I = 0 To nDim
        If Trim(sHeader(I)) = "" Then GoTo skipit
        If IsStringLine(sHeader(I)) Then GoTo skipit
        If Left(sHeader(I), 1) = "'" Then          ' READ HEADER
           ' Check the name
           sName = GetStringPart(sHeader(I), "NAME: "): If sName <> "" Then GoTo skipit
           sAuthor = GetStringPart(sHeader(I), "@")
           sCopyright = GetStringPart(sHeader(I), "(c)")
           If sAuthor <> "" Or sCopyright <> "" Then GoTo skipit
             
           
           iL = InStr(1, sHeader(I), "Name: ", vbTextCompare)
           If iL > 0 Then   ' Check The Name
                   sName = Right(sHeader(I), Len(sHeader(I)) - (iL + Len("Name: ")))
                   GoTo skipit
           End If
           '---------------------------------------------------------------------------
           'Check the Author and copyright
           iL = InStr(1, sHeader(I), "@")
           If iL > 0 Then
              iR = InStr(iL + 1, sHeader(I), ",")
              If iR > 0 Then
                 sAuthor = Trim(Mid(sHeader(I), iL + 1, iR - 1))
                 '--------------------------------------------------------------------
                 'Check Copyright in Author Linne
                 'iL = InStr(
              Else
                 sAuthor = Right(sHeader(I), Len(sHeader(I)) - iL - 1)
              End If
           End If
           If sAuthor = "" Then
               iL = InStr(1, sHeader(I), "AUTHOR: ", vbTextCompare)
               If iL > 0 Then
                  sAuthor = Trim(Right(sHeader(I), Len(sHeader(I)) - iL + 1))
                  GoTo skipit
               End If
           Else
               GoTo skipit
           End If
           '----------------------------------------------------------------------------
           
           
           
           
           
        End If
skipit:
    Next I
End If
'--------------------------------------------------------
ExitHere:
      GetHeader = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Get Date Creation
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetCreateDate(sModName As String) As String
Dim sSysObj As String, sRes As String
On Error Resume Next

If Left(sModName, 1) = "F" And Len(sModName) > 5 Then
   If Left(sModName, 5) = "Form_" Then
       sSysObj = Right(sModName, Len(sModName) - 5)
   End If
End If
If sSysObj = "" Then sSysObj = sModName
'----------------------------------------------
sRes = Nz(DLookup("DateCreate", "MSysObjects", "Name = " & sCH(sSysObj)), "")
'---------------------------------
ExitHere:
    GetCreateDate = sRes '!!!!!!!!!!!
End Function

'------------------------------------------------------------------------------------------------------------------------------------------
' The function will add the module from the file (needlessly)
'------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AddModFromFile(FilePath As String, Optional ModName As String = "", Optional OverwriteExisting As Boolean = True)
Dim bExist As Boolean
Dim sName As String, myComponent As Object

    On Error GoTo ErrHandle
'--------------------------
If Dir(FilePath, vbNormal + vbHidden + vbSystem) <> vbNullString Then
            If ModName <> "" Then
                sName = ModName
            Else
                sName = ExtractModName(FilePath)
            End If
                '----------------------------------------------------------------------------
                bExist = IsModuleExist(sName)
                If bExist Then
                   If OverwriteExisting Then
                      RemoveModule sName
                   Else
                      MsgBox "Модуль " & sName & " существует. Отказ операции импорта"
                      Exit Sub
                   End If
                End If
            '----------------------------------------------------------------------------------
            ' IMPORT Itself
              Set myComponent = VBE.ActiveVBProject.VBComponents.Import(FilePath)
              myComponent.Name = sName
   End If
'----------------------------------
ExitHere:
   DoCmd.Save
   Exit Sub
'-------------------
ErrHandle:
   ErrPrint2 "", Err.Number, Err.Description, MOD_NAME
   Err.Clear:  Resume ExitHere
End Sub
