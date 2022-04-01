Attribute VB_Name = "#_ACCESS"
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
'                 d$$$$$$$$$F                $P"                   ##  ##  ##  ##   #######        ####  ####  #### ####    ####  ####
'                 $$$$$$$$$$F                                       ## ##  ##  ##  ##    ##       ## ## ##    ##    ##     ##    ##
'                  *$$$$$$$$"                                        ####  #####  ##     ##      ##  ## ##    ##    ###     ##     ##
'                    "***""               _____________                                         ####### ##    ##    ###       #     #
' STANDARD MODULE WITH DEFAULT FUNCTIONS |v 2021/08/20 |                                       ##    ## ##    ##    ##      ##    ##
' The module contains some functions to work with MS Access and is part of the G-VBA library  ##     ##  ####  #### ####  ###   ###
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
Option Explicit

Private Const MOD_NAME As String = "#_ACCESS"

Public Enum ACC_OBJTYPE
    ACC_ALL_OBJECTS = 0
    ACC_TBL_LOCAL = 1
    ACC_TBL_LINKED_ODBC = 4
    ACC_TBL_LINKED = 6
    ACC_QUERIS = 5
    ACC_FORMS = -32768
    ACC_REPORTS = -32764
    ACC_MACROS = -32766
    ACC_MODULES = -32761
End Enum

Private colForms As New Collection
Private mintForm As Integer

Private Const acbcOffsetHoriz = 75
Private Const acbcOffsetVert = 375


Private Const HKEY_CLASSES_ROOT As Long = &H80000000
Private Const READ_CONTROL As Long = &H20000
Private Const STANDARD_RIGHTS_READ As Long = (READ_CONTROL)
Private Const KEY_QUERY_VALUE As Long = &H1
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Private Const KEY_NOTIFY As Long = &H10
Private Const SYNCHRONIZE As Long = &H100000
Private Const KEY_READ As Long = (( _
                  STANDARD_RIGHTS_READ _
               Or KEY_QUERY_VALUE _
               Or KEY_ENUMERATE_SUB_KEYS _
               Or KEY_NOTIFY) _
               And (Not SYNCHRONIZE))
Private Const ERROR_SUCCESS As Long = 0&
Private Const ERROR_NO_MORE_ITEMS As Long = 259&

'**************************
#If Win64 Then
    Private Declare PtrSafe Function RegOpenKeyEx _
        Lib "advapi32.dll" Alias "RegOpenKeyExA" ( _
        ByVal hKey As LongPtr, _
        ByVal lpSubKey As String, _
        ByVal ulOptions As Long, _
        ByVal samDesired As Long, _
        phkResult As LongPtr) As Long

    Private Declare PtrSafe Function RegEnumKey _
        Lib "advapi32.dll" Alias "RegEnumKeyA" ( _
        ByVal hKey As LongPtr, _
        ByVal dwIndex As Long, _
        ByVal lpName As String, _
        ByVal cbName As Long) As Long
        
    Private Declare PtrSafe Function RegQueryValue _
        Lib "advapi32.dll" Alias "RegQueryValueA" ( _
        ByVal hKey As LongPtr, _
        ByVal lpSubKey As String, _
        ByVal lpValue As String, _
        lpcbValue As Long) As Long
    Private Declare PtrSafe Function RegCloseKey _
        Lib "advapi32.dll" ( _
        ByVal hKey As LongPtr) As Long

#Else
    Private Declare Function RegOpenKeyEx _
        Lib "advapi32.dll" Alias "RegOpenKeyExA" ( _
            ByVal hKey As Long, _
            ByVal lpSubKey As String, _
            ByVal ulOptions As Long, _
            ByVal samDesired As Long, _
            ByRef phkResult As Long) As Long
 
    Private Declare Function RegEnumKey _
        Lib "advapi32.dll" Alias "RegEnumKeyA" ( _
            ByVal hKey As Long, _
            ByVal dwIndex As Long, _
            ByVal lpName As String, _
            ByVal cbName As Long) As Long
 
    Private Declare Function RegQueryValue _
        Lib "advapi32.dll" Alias "RegQueryValueA" ( _
            ByVal hKey As Long, _
            ByVal lpSubKey As String, _
            ByVal lpValue As String, _
            ByRef lpcbValue As Long) As Long
 
    Private Declare Function RegCloseKey _
        Lib "advapi32.dll" ( _
            ByVal hKey As Long) As Long
#End If

'************************************

'======================================================================================================================================================
'  Show Navigation Pane
'======================================================================================================================================================
Public Sub ShowNavPane()
    Call DoCmd.SelectObject(acTable, , True)
End Sub

'======================================================================================================================================================
'  Hide Navigation Pane
'======================================================================================================================================================
Public Sub HideNavPane()
    Call DoCmd.NavigateTo("acNavigationCategoryObjectType") 'select the navigation pange
    Call DoCmd.RunCommand(acCmdWindowHide) 'hide the selected object
End Sub

'======================================================================================================================================================
'  Hide Ribbon
'======================================================================================================================================================
Public Sub HideRibbon()
    Call DoCmd.ShowToolbar("Ribbon", acToolbarNo)
End Sub
'======================================================================================================================================================
'  Show Ribbon
'======================================================================================================================================================
Public Sub ShowRibbon()
    Call DoCmd.ShowToolbar("Ribbon", acToolbarYes)
End Sub

'======================================================================================================================================================
'  Change AppDB Icon and Title
'======================================================================================================================================================
Public Function AppIconTitle(Optional sIconPath As String, Optional sTitle As String) As Boolean
Dim bRes As Boolean

 Const DB_Text As Long = 10
 
    On Error Resume Next
'------------------------------
 If sIconPath = "" And sTitle = "" Then Exit Function
 
 If sTitle <> "" Then
    bRes = AddAppProperty("APP_TITLE", DB_Text, sTitle)
 End If
 
 If sIconPath <> "" Then
     If Dir(sIconPath) <> "" Then
        bRes = AddAppProperty("AppIcon", DB_Text, sIconPath) Or bRes
        If Not bRes Then GoTo ExitHere
        
        CurrentDb.Properties("UseAppIconForFrmRpt") = 1
     End If
 End If
 
 Application.RefreshTitleBar
'------------------------------
ExitHere:
    AppIconTitle = bRes '!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'=====================================================================================================================================================
' Listing objects in DB
'=====================================================================================================================================================
Public Function ListDBObjects(Optional lObjectType As ACC_OBJTYPE = ACC_ALL_OBJECTS, Optional DLM As String = ";") As String
Dim db  As DAO.Database, RS   As DAO.Recordset
Dim sSQL As String, sRes As String
 
On Error GoTo ErrHandle
'----------------------------------

sSQL = "SELECT MsysObjects.Name FROM MsysObjects " & _
       "WHERE ((((MsysObjects.Name) Not Like '~*') And ((MsysObjects.Name) Not Like 'MSys*')) " & _
       IIf(lObjectType <> ACC_ALL_OBJECTS, " AND ((MsysObjects.Type)=" & lObjectType & ")", "") & _
       " AND ((MsysObjects.Flags)=0)) ORDER BY MsysObjects.Name;"

 
    Set db = CurrentDb
    Set RS = db.OpenRecordset(sSQL, dbOpenSnapshot)
    With RS
        If .RecordCount <> 0 Then
            Do While Not .EOF
                sRes = sRes & Trim(CStr(![Name])) & DLM
                .MoveNext
            Loop
        End If
    End With
'-------------------------------------
ExitHere:
    If sRes <> "" Then sRes = Left(sRes, Len(sRes) - Len(DLM))
    ListDBObjects = sRes '!!!!!!!!!!!!
    If Not RS Is Nothing Then
        RS.Close
        Set RS = Nothing
    End If
    If Not db Is Nothing Then Set db = Nothing
    Exit Function
'--------------
ErrHandle:
    ErrPrint2 "ListDBObjects", Err.Number, Err.Description, MOD_NAME
    Err.Clear:  Resume ExitHere
End Function

'======================================================================================================================================================
' Check If Some Object (Table/Query/Module/Form/Report) Exists
'======================================================================================================================================================
Public Function IsAObject(sObjectName As String, Optional ObjType As ACC_OBJTYPE = ACC_ALL_OBJECTS) As Boolean
Dim sName As String
    
    On Error Resume Next
'-----------------------
If ObjType = ACC_ALL_OBJECTS Then
    sName = Nz(DLookup("Name", "MSysObjects", "Name = " & sCH(sObjectName)), "")
Else
    sName = Nz(DLookup("Name", "MSysObjects", "Name = " & sCH(sObjectName)) & " AND Type = " & ObjType, "")
End If
'-----------------------
ExitHere:
    IsAObject = sName <> "" '!!!!!!!!!!!!!!
End Function
'======================================================================================================================================================
' Check if the table with name sTable is exists
'======================================================================================================================================================
Public Function IsTable(STABLE As String, Optional bCheckSystem As Boolean = True) As Boolean
Dim bRes As Boolean

On Error Resume Next
'-----------------------
   If bCheckSystem Then
    If Nz(DLookup("MSysObjects.ID", "MSysObjects", _
            "(Name = " & sCH(STABLE) & ") AND (Type In (1, 4, 6))"), _
            -1) <> -1 Then bRes = True
   Else
       If CurrentDb.TableDefs(STABLE).Name = STABLE Then bRes = True
   End If
'----------------------
ExitHere:
    IsTable = bRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'======================================================================================================================================================
' Check if  query is exists
'======================================================================================================================================================
Public Function IsQuery(sQuery As String, Optional bCheckSystem As Boolean = True) As Boolean
Dim bRes As Boolean

On Error Resume Next
'-----------------------
   If bCheckSystem Then
    If Nz(DLookup("MSysObjects.ID", "MSysObjects", _
            "(Name = " & sCH(sQuery) & ") AND (Type = 5)"), _
            -1) <> -1 Then bRes = True
   Else
       If CurrentDb.QueryDefs(sQuery).Name = sQuery Then bRes = True
   End If
'----------------------
ExitHere:
    IsQuery = bRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'======================================================================================================================================================
' Check if  form is exists
'======================================================================================================================================================
Public Function IsForm(sForm As String) As Boolean
Dim bRes As Boolean

On Error Resume Next
'-----------------------
    If Nz(DLookup("MSysObjects.ID", "MSysObjects", _
            "(Name = " & sCH(sForm) & ") AND (Type = -32768)"), _
            -1) <> -1 Then bRes = True
'----------------------
ExitHere:
    IsForm = bRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'======================================================================================================================================================
' Show System ProgressBar
' The Command: 0 - Init; 1 - Show Progress; 2 - Show only Text;  -1 - Clear
' If Esc Pressed, then return False
'======================================================================================================================================================
Public Function ProgressMeter(Optional dCurrent As Double, Optional sText As String, Optional iCommand As Integer = 1) As Boolean

Const MAX_LEN As Long = 100
Const VB_ESC As Long = 27

    On Error Resume Next
'----------------------------
Select Case iCommand
Case 0:   ' INIT METER
    SysCmd acSysCmdInitMeter, sText, MAX_LEN
Case 1:   ' Show Update (Progress)
    SysCmd acSysCmdUpdateMeter, MAX_LEN * dCurrent
Case 2:   ' Show only Text
    SysCmd acSysCmdSetStatus, sText
Case -1:  ' Close and clean
    SysCmd acSysCmdClearStatus
End Select

DoEvents

If IsKeyPress(VB_ESC) Then Exit Function

'-----------------------------
ExitHere:
     ProgressMeter = True '!!!!!!!!!!!!!!!!!
     Exit Function
End Function

'======================================================================================================================================================
' Backup Current DB
'======================================================================================================================================================
Public Function BackupDB(Optional sSourcePath As String, Optional sBackupFolder As String) As String
Dim sNewFile As String, sCurrentFile As String
Dim sFolder As String, bRes As Boolean

Const BKP_FOLDER As String = "BACKUP"

On Error GoTo ErrHandle
'-----------------------------
    sFolder = IIf(sBackupFolder <> "", sBackupFolder, CurrentProject.Path & "\" & BKP_FOLDER)
    If Right(sFolder, 1) = "\" Then sFolder = Left(sFolder, Len(sFolder) - 1)
    If Dir(sFolder, vbDirectory) = "" Then
           bRes = FolderCreate(sFolder)
    Else
           bRes = True
    End If
    If Not bRes Then Err.Raise 1000, , "Can't create the backup folder " & sFolder

sCurrentFile = IIf(sSourcePath <> "", sSourcePath, CurrentProject.Path & "\" & CurrentProject.Name)
If Dir(sCurrentFile) = "" Then Err.Raise 1000, , "Absent Source File " & sCurrentFile

sNewFile = sFolder & "\" & Format(Now(), "yyyymmdd") & "_" & FileNameOnly(sCurrentFile)
If Dir(sNewFile) <> "" Then Kill sNewFile

    bRes = CopyFile(sCurrentFile, sNewFile)
    If Not bRes Then Exit Function
'-----------------------------
ExitHere:
      BackupDB = sNewFile '!!!!!!!!!!!!!!!
      Debug.Print "The backup file is created"
      Exit Function
'--------------
ErrHandle:
      ErrPrint2 "BackupDB", Err.Number, Err.Description, MOD_NAME
      Err.Clear
End Function

'======================================================================================================================================================
' Open Form in Gracle
'======================================================================================================================================================
Public Function OpenExternalForm(Optional sForm As String = "f_VBA", _
                                                       Optional sPath As String = "C:\Users\valer\Google Drive\_ZWORKS\_DATABASES\VBALIB\GRACKLE.accdb")
Dim Acc As Object
    
On Error GoTo ErrHandle
'--------------------------------
    If Dir(sPath) = "" Then Exit Function
    
    Set Acc = CreateObject("Access.Application")
    Call Acc.OpenCurrentDatabase(sPath)
    Call Acc.DoCmd.OpenForm(sForm)
'----------------------------------
ExitHere:
    OpenExternalForm = True '!!!!!!!!!!
    Set Acc = Nothing
    Exit Function
'----------------
ErrHandle:
    ErrPrint2 "OpenExternalForm", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'======================================================================================================================================================
' Show (bHidden = False) /Hide (bHidden = False) the object in DB
'======================================================================================================================================================
Public Sub SetHiddenAttribute(ObjectType As AcObjectType, ObjectName As String, Optional bHidden As Boolean)
    
    On Error Resume Next
'-----------------
    Application.SetHiddenAttribute ObjectType, ObjectName, bHidden
End Sub

'======================================================================================================================================
' Get some property for db object (MS Access Only)
'======================================================================================================================================
Public Function GetAccessProp(obj As Object, PropName As String) As String
    On Error Resume Next
    GetAccessProp = obj.Properties(PropName)  '!!!!!!!!!!!!!!!!!
End Function
'======================================================================================================================================
' Assign New Property (MS Access Only)
'======================================================================================================================================
Public Sub SetAccessProp(obj As Object, PropName As String, PropValue As Variant, Optional PROPTYPE As Integer = dbText)

On Error GoTo ErrHandle
    obj.Properties(PropName) = PropValue '!!!!!!!!!!!!!!!!!
    Exit Sub
'----------------------------------------------------------
ErrHandle:

    If (Err.Number = 3270) Or (Err.Number = 2455) Then      ' Свойства не существует
       Call CreateAccessProp(obj, PropName, PropValue, PROPTYPE)
    Else
       ErrPrint2 "SetAccessProp", Err.Number, Err.Description, MOD_NAME
    End If
End Sub
'======================================================================================================================================
' Add new Property to DB Objects
'======================================================================================================================================
Public Sub CreateAccessProp(obj As Object, PropName As String, PropValue As Variant, Optional PROPTYPE As Integer = dbText)
On Error Resume Next
    obj.Properties.Append obj.CreateProperty(PropName, PROPTYPE, PropValue)
End Sub

'=================================================================================================================================================
' Check if Form is loaded
'=================================================================================================================================================
Public Function IsFormLoaded(formName As String) As Boolean
On Error Resume Next
   IsFormLoaded = CurrentProject.AllForms(formName).IsLoaded
   Err.Clear
End Function
'=================================================================================================================================================
' OPEN DETAIL FORM FOR SPECIFIC ID
'  To create another instance of existing form use
'                                           Set frm = New Form_FORMNAME   and pass it as FormNewInstance
'==================================================================================================================================================
Public Function OpenDetailForm(formName As String, Optional CurrentID As Variant, Optional WinMode As AcWindowMode = acDialog, _
                                                                     Optional FormNewInstance As Variant, Optional sARG As String = "") As Boolean
Dim IDD As String                        ' String ID representation

On Error GoTo ErrHandle
'-----------------------------------------------------------------------------------------------
If Not IsMissing(CurrentID) Then
    IDD = IsGuid(CStr(CurrentID))
    If IDD = "" Then IDD = CurrentID
End If
'--------------------------------------------------------------------------------------------
If IsFormLoaded(formName) Then                                 '  Then Open another Instance
        If IsMissing(FormNewInstance) Then Err.Raise 10000, , "No FormNewInstance param"
        If Not IsObject(FormNewInstance) Then Err.Raise 10000, , "Wrong FormNewInstance Type: should be a Form"
        Call OpenNewInstance(FormNewInstance, IDD, sARG)
Else                                                           '  Open a New Instance
        If IDD = "" Then       ' CurrentID - Empty
                DoCmd.OpenForm formName, , , , , acHidden
                DoCmd.GoToRecord acDataForm, formName, acNewRec
                If sARG <> "" Then
                        DoCmd.OpenForm formName, , , , , WinMode, sARG
                Else
                        DoCmd.OpenForm formName, , , , , WinMode
                End If
        Else
                If sARG <> "" Then
                        DoCmd.OpenForm formName, , , "ID=" & IDD, , WinMode, sARG
                Else
                        DoCmd.OpenForm formName, , , "ID=" & IDD, , WinMode
                End If
        End If
End If
'----------------------------------------------------------------------------------------------
ExitHere:
       OpenDetailForm = True '!!!!!!!!!!!!
       Exit Function
'----------------------------------------------
ErrHandle:
       ErrPrint2 "OpenDetailForm", Err.Number, Err.Description, MOD_NAME
       Err.Clear
End Function

'====================================================================================================================================================
' The function writes the positions to the local storage so that when the next time it is opened, the specified geometry is restored.
' Requires a special class cAccessWindows to manipulate windows
' If bOnlyFirstTime = True, write it only for first call
'=====================================================================================================================================================
Public Sub WriteWindowState(f As Form, Optional bOnlyFirstTime As Boolean = False)
Dim AWin As New cAccessWindows                           ' Манипулятор окон
Dim sState As String                                     ' Статус - строка
Dim parName As String                                    ' Параметр для хранения статуса
    
On Error GoTo ErrHandle
'-------------------------------------------------------------------------------
        parName = "WinState_" & f.Name
        If bOnlyFirstTime Then
                 If GetLocal(parName) <> "" Then GoTo ExitHere 'Выходим, если запись уже есть
        End If
        
        If Trim(GetLocal("AllowFormResize")) = "0" Then Exit Sub
        sState = AWin.GetWindowState(f)
        If sState = "" Then Exit Sub
        '---------------------------------------------------------------------------
        ' Записываем статус
        SetLocal parName, sState, "Координаты формы " & f.Name
'-------------------------------------------------------------------------------
ExitHere:
        Set AWin = Nothing
        Exit Sub
'------------------------
ErrHandle:
        ErrPrint2 "WriteWindowState", Err.Number, Err.Description, MOD_NAME
        Err.Clear: Resume ExitHere
End Sub
'===============================================================================================================================================
' The function restores the form window based on the entries made earlier in the local store.
' If there are no such records, the form takes the values ??Access
' If bCentralize = true, then setup only size but not the position and then form will centralized
'===============================================================================================================================================
Public Sub RestoreWindowState(f As Form, Optional bCentralize As Boolean = False)
Dim AWin As New cAccessWindows                           ' Манипулятор окон
Dim sState As String                                     ' Статус - строка
Dim parName As String                                    ' Параметр для хранения статуса
        
On Error GoTo ErrHandle
'-----------------------------------------------------------------------
    parName = "WinState_" & f.Name
    sState = GetLocal(parName)
    If sState = "" Then Exit Sub
'-----------------------------------------------------------------------
        AWin.SetWindowState f, sState
            If bCentralize Then                               ' Если задана централизация
                AWin.CenterForm f
            End If
    '-----------------------------------------------------------------------
ExitHere:
        Set AWin = Nothing
        Exit Sub
    '------------------------
ErrHandle:
        ErrPrint2 "RestoreWindowState", Err.Number, Err.Description, MOD_NAME
        Err.Clear: Resume ExitHere
End Sub

'=============================================================================================================
'  COMPILE VBA MODULES
'=============================================================================================================
Public Sub CompileNow()

On Error Resume Next
'-----------------------------------------
    Const COMPILE_SAVE As Integer = 16483
    Const COMPILE_NOSAVE As Integer = 16484

    SysCmd 504, 16483
'Debug.Print "The VBA is Compiled"
End Sub
'=====================================================================================================================================================
' Compact and Backup
'=====================================================================================================================================================
Public Sub CompactDB(Optional sPathToDB As String, Optional sCopyPathForNextVersion As String)
Dim sPath As String

Const OP_NUMBER As Integer = 602

On Error Resume Next
'-------------------------------
If MsgBox("Compact DB right now? It could take some time and close then restore this DB", _
                          vbYesNoCancel + vbQuestion, "Compact & Repair") <> vbYes Then Exit Sub
               
    sPath = IIf(sPathToDB <> "", sPathToDB, CurrentProject.FullName)
    If sCopyPathForNextVersion <> "" Then
    Else
        SysCmd OP_NUMBER, sPath, sCopyPathForNextVersion
    End If
     
End Sub
'======================================================================================================================================================
' Get VBA REFERENCESS (adapted)
' Authors: Dirk Goldgar,  Contributor: Tom van Stiphout (https://www.devhut.net/2017/03/03/vba-list-references/)
' Return str array (with sep delimeter), each row has format: NAME;FULLPATH;GUID; KIND; BUILIN; IsBroken
'======================================================================================================================================================
Public Function ListReferences(Optional DLM As String, Optional SEP As String = vbCrLf) As String
Dim ref As Object, sRes As String, lngCount As Long, lngBrokenCount As Long
Dim bBroken As Boolean, sOUT() As String, nOut As Integer

    On Error Resume Next
'------------------------------
    nOut = -1: ReDim sOUT(0)
    
    For Each ref In Application.References
        sRes = vbNullString: sRes = GetRefDetail(ref)
        If sRes <> "" Then
           nOut = nOut + 1: ReDim Preserve sOUT(nOut)
           sOUT(nOut) = sRes
        End If
    Next ref
    
sRes = Join(sOUT, SEP)
'------------------------------
ExitHere:
    ListReferences = sRes '!!!!!!!!!!!!!
End Function
'======================================================================================================================================================
' App available referenses as list
'======================================================================================================================================================
Public Function AllAvailableRefList(Optional DLM As String = ";", Optional SEP As String = vbCrLf) As String
Dim R1 As Long, R2 As Long

#If Win64 Then
    Dim hHK1 As LongPtr, hHK2 As LongPtr, hHK3 As LongPtr, hHK4 As LongPtr
#Else
    Dim hHK1 As Long, hHK2 As Long, hHK3 As Long, hHK4 As Long
#End If
Dim I As Long, J As Long, sRes As String, sRow As String
Dim lpPath As String, lpGUID As String, lpName As String, lpValue As String

    On Error Resume Next
'---------------------------------
   lpPath = String$(128, vbNullChar): lpValue = String$(128, vbNullChar)
   lpName = String$(128, vbNullChar): lpGUID = String$(128, vbNullChar)
   
   R1 = RegOpenKeyEx(HKEY_CLASSES_ROOT, "TypeLib", ByVal 0&, KEY_READ, hHK1)
   If R1 = ERROR_SUCCESS Then
       I = 0
       Do While Not R1 = ERROR_NO_MORE_ITEMS
           R1 = RegEnumKey(hHK1, I, lpGUID, Len(lpGUID))
           If R1 = ERROR_SUCCESS Then
               R2 = RegOpenKeyEx(hHK1, lpGUID, ByVal 0&, KEY_READ, hHK2)
                                                      sRow = TrimToNull(lpGUID) ' Left(lpGUID, InStr(lpGUID, "}"))
               If R2 = ERROR_SUCCESS Then
                   J = 0
                   Do While Not R2 = ERROR_NO_MORE_ITEMS
                       R2 = RegEnumKey(hHK2, J, lpName, Len(lpName)) '1.0
                       If R2 = ERROR_SUCCESS Then
                           RegQueryValue hHK2, lpName, lpValue, Len(lpValue)
                           
                                                       sRow = sRow & DLM & TrimToNull(lpValue)
                           
                           RegOpenKeyEx hHK2, lpName, ByVal 0&, KEY_READ, hHK3
                           RegOpenKeyEx hHK3, "0", ByVal 0&, KEY_READ, hHK4
                           RegQueryValue hHK4, "win32", lpPath, Len(lpPath)
                           
                                                      sRow = sRow & DLM & TrimToNull(lpPath)
                                                      sRes = sRes & SEP & sRow
                           J = J + 1: sRow = ""
                           

                       End If
                   Loop
               End If
           End If
           I = I + 1
       Loop
       
       RegCloseKey hHK1: RegCloseKey hHK2: RegCloseKey hHK3: RegCloseKey hHK4
   End If
   If sRes <> "" Then sRes = Right(sRes, Len(sRes) - Len(SEP))
'-------------------------------
ExitHere:
    AllAvailableRefList = sRes '!!!!!!!!!!!
    Exit Function
'------------
ErrHandle:
    ErrPrint2 "AllAvailableRefList", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'======================================================================================================================================================
' Check if App has Nesessary Refs
'======================================================================================================================================================
Public Function CheckRef(Optional RefNeeded As String = "_GRACKLE") As Boolean
Dim sList As String, REFS() As String, nRefs As Integer
Dim bRes As Boolean, I As Integer, sWork As String
    
Const DLM As String = ";"
Const SEP As String = vbCrLf

    On Error GoTo ErrHandle
'---------------------
    sList = ListReferences(DLM, SEP)
    If sList = "" Then Exit Function
    REFS = Split(sList, SEP): nRefs = UBound(REFS)
    For I = 0 To nRefs
        If REFS(I) <> "" Then
            sWork = Split(REFS(I), DLM)(0)
            If RefNeeded = sWork Then
                 bRes = True: Exit For
            End If
        End If
    Next I
'---------------------
ExitHere:
    CheckRef = bRes '!!!!!!!!!!
    Exit Function
'----------------
ErrHandle:
    ErrPrint2 "CheckRef", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'======================================================================================================================================================
'RemoveLib: Removes a library reference to this script programmatically
'======================================================================================================================================================
Public Function RemoveLib(LibName As String, Optional FilePath As String, Optional GUID As String) As String
Dim oApp As Object, oRef As Object, bRes As Boolean, sName As String

Const APPLICATION_OBJ_NAME As String = "Access.Application"

    On Error GoTo ErrHandle
'--------------------------------------------------
Set oApp = GetObject(, APPLICATION_OBJ_NAME)
    
    bRes = IsLib(LibName, FilePath, GUID) ' Check if the library has already been added
    If Not bRes Then
       bRes = True
       GoTo ExitHere
    End If
    Call oApp.References.Remove(LibName)
'---------------------------
ExitHere:
    RemoveLib = bRes '!!!!!!!!!!!!!!
    Exit Function
'----------------
ErrHandle:
    ErrPrint2 "RemoveLib", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'======================================================================================================================================================
'AddLib: Adds a library reference to this script programmatically, so that
'        libraries do not need to be added manually.
'======================================================================================================================================================
Public Function AddLib(LibName As String, Optional FilePath As String, Optional GUID As String, Optional major As Long, _
                                                                                                                    Optional minor As Long) As Boolean
Dim oApp As Object, oRef As Object, bRes As Boolean, sName As String

Const APPLICATION_OBJ_NAME As String = "Access.Application"

    On Error GoTo ErrHandle
'--------------------------------------------------
If FilePath = "" And GUID = "" Then Err.Raise 10000, , "Can't add library for empty params: set FilePath or Guid"

Set oApp = GetObject(, APPLICATION_OBJ_NAME)
    
    bRes = IsLib(LibName, FilePath, GUID) ' Check if the library has already been added
    If bRes Then GoTo ExitHere
        
If FilePath <> "" Then
    Set oRef = oApp.References.AddFromFile(FilePath)
Else
    Set oRef = oApp.References.AddFromGuid(GUID, major, minor)
End If
    sName = oRef.Name: bRes = True
'---------------------------
ExitHere:
    AddLib = bRes '!!!!!!!!!!!!!!
    Exit Function
'----------------
ErrHandle:
    ErrPrint2 "AddLib", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'======================================================================================================================================================
' Find the Library (check if lib is install)
'======================================================================================================================================================
Public Function IsLib(Optional LibName As String, Optional FilePath As String, Optional sGUID As String) As Boolean
Dim oRef As Object, oApp As Object, bRes As Boolean

Const APPLICATION_OBJ_NAME As String = "Access.Application"

    On Error Resume Next
'---------------------
    Set oApp = GetObject(, APPLICATION_OBJ_NAME)
    
    For Each oRef In oApp.References
        If oRef.Name = LibName Then
           bRes = True
        ElseIf oRef.FullPath = FilePath Then
           bRes = True
        ElseIf oRef.GUID = sGUID Then
           bRes = True
        End If
        If bRes Then Exit For
    Next
'---------------------
ExitHere:
    IsLib = bRes '!!!!!!!!!!!!
End Function


'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'------------------------------------------------------------------------------------------------------------------------------------------------------
'   Create or change project properties
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function AddAppProperty(strName As String, varType As Variant, varValue As Variant) As Boolean

Dim dbs As Object, prp As Variant, bRes As Boolean

Const conPropNotFoundError = 3270
 
    On Error GoTo ErrHandle
'-------------------------------------
 Set dbs = CurrentDb
 
 dbs.Properties(strName) = varValue
 bRes = True

'------------------------
ExitHere:
    AddAppProperty = bRes '!!!!!
    Exit Function
'--------
ErrHandle:
 If Err = conPropNotFoundError Then
    Set prp = dbs.CreateProperty(strName, varType, varValue)
    dbs.Properties.Append prp
    Resume
 Else
    AddAppProperty = False
    Resume ExitHere
 End If
End Function
'------------------------------------------------------------------------------------------------------------------------
' Open another instance
'------------------------------------------------------------------------------------------------------------------------
Private Sub OpenNewInstance(FormInstance As Variant, Optional id As String, Optional sARG As String)
Dim frm As Form

On Error GoTo ErrHandle
'----------------------------------------------------------------------------------------
    If Not IsObject(FormInstance) Then Err.Raise 10000, , "Wrong format of FormInstance: should be a form"
    
    Set frm = FormInstance
    
    colForms.Add Item:=frm, key:=FormInstance.Name & " " & frm.hWnd & ""
    mintForm = mintForm + 1
    frm.Caption = IIf(frm.Caption <> "", frm.Caption, frm.Name) & "/" & mintForm
    
    frm.SetFocus:   DoCmd.MoveSize mintForm * acbcOffsetHoriz, mintForm * acbcOffsetVert
    If sARG <> "" Then frm.OpenArgs = sARG

'---------------------------------------------------------------------------------------
    If id <> "" Then frm.RecordSource = "SELECT * FROM " & _
                                          frm.RecordSource & " WHERE ([ID]=" & id & ");"

    frm.Visible = True
'------------------
ExitHere:
       Set frm = Nothing
       Exit Sub
'----------------------------------------------
ErrHandle:
       ErrPrint2 "OpenNewInstance", Err.Number, Err.Description, MOD_NAME
       Err.Clear: Resume ExitHere
End Sub

'------------------------------------------------------------------------------------------------------------------------------------------------------
' Function process reference with safety
' NAME;FULLPATH;GUID; KIND; BUILIN; IsBroken
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetRefDetail(ByRef ref As Object, Optional DLM As String = ";") As String
Dim sRes As String

    On Error Resume Next
'-----------------------
    With ref
        sRes = .Name & DLM & .FullPath & DLM & .GUID & DLM & .Kind & DLM & .IsBroken
    End With
'-----------------------
    GetRefDetail = sRes '!!!!!!!!!!!!
End Function



