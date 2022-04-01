Attribute VB_Name = "#_DOCUMENTER"
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
' DOCUMENTATION MANAGER OF ACCESS DATABASE
' @ Valery Khvatov (valery.khvatov@gmail.com),  (c) DigitalXpert Inc.,  01/20140808
'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
Option Compare Database
Option Explicit

Private Const MOD_NAME As String = "#_DOCUMENTER"
'**********************************
Private Type dbField
           id As Integer
           DateCreated As Date
           HASH As String
           Name As String
           FldCategory As String
           FldType As String
           Header As String
           Description As String
           BODY As String
           fldSize As Long
           META As String
End Type
'**********************************
Private Type dbObject             ' The Data Base Object
            id As Integer
            HASH As String
            Parent As String      ' Name of Parent Object
            Title As String       ' The Object Title (descriptively)
            Name As String        ' Object Name
            Type As Integer       ' Object Type (1;Table;5;Query;-32768;Form;-32764;Report;-32766;Macro;-32761;Module;9000;Command)
            DateCreated As Date   ' The created date
            BODY As String        ' Object content
            DataType As Integer   ' Object Type
            Description As String ' Object Description
            Group As Integer      ' The group of object (-3;SYS0;-2;SYS1;-1;SYS2;1;Entity;2;Ref;3;CrossLink;4;QueryExtension;5;WindowForm;
                                  '                  6;ListingForm,7;DetailedForm;8;SubForm;9;Standalone;10;Class;11;VBA;12;Other)
            META As String        ' Information Extention (KEY-VALUES STRING)
            Path As String        ' Path to object (including import and export)
            Version As String     ' Object Version
End Type
'**********************************
' TYPICAL CONSTANTS
Private Const SepLineCounter As Integer = 120 ' Lenght of divider string
Public VBWork As String

'**********************************
'==========================================================================================================================================
' generate Full report and save it to external File
'==========================================================================================================================================
Public Sub VBAREPORT()
Dim MyVBA As New cVBA, sRes As String
Dim FileNumber As Integer, sPath As String

Const fileName As String = "VBA_LISTING.txt"

On Error GoTo ErrHandle
'--------------------------------------------
   sRes = MyVBA.GetFullReport

If sRes <> "" Then
    sPath = CurrentProject.Path & "\" & fileName
    If Len(Dir(sPath)) > 0 Then Kill sPath
    FileNumber = FreeFile: Open sPath For Append As #FileNumber
    
    Print #FileNumber, sRes
End If

'-----------------------------
ExitHere:
        Set MyVBA = Nothing
        Close #FileNumber
        Exit Sub
'--------------
ErrHandle:
        ErrPrint "VBAREPORT", Err.Number, Err.Description
        Err.Clear: Resume ExitHere
End Sub


'==========================================================================================================================================
'  Write Function from Clipboard/Text To Disk for Next Processing
'==========================================================================================================================================
Public Function FuncToDisk(Optional sFunc As String, Optional sFileName As String, Optional bUseClipboard As Boolean = True) As String
Dim sCode As String, sPath As String, bFolder As Boolean, sFuncName As String

Const CODEBASE As String = "CODE"
Const FUNC_REPO As String = "VBA"

On Error GoTo ErrHandle
'--------------------------------------------
sCode = sFunc: If sCode = "" Then sCode = FromClipboard()
If sCode = "" Then Exit Function

sPath = sFileName
If sPath = "" Then
    sPath = CurrentProject.Path & "\" & CODEBASE: bFolder = True
    If Dir(sPath, vbDirectory) = "" Then Call FolderCreate(sPath)
    sPath = sPath & "\" & FUNC_REPO: If Dir(sPath, vbDirectory) = "" Then Call FolderCreate(sPath)
Else
    If FileExt(sPath) = "" Then bFolder = True ' This is a folder
End If
If Dir(sPath, vbDirectory) = "" Then Err.Raise 10000, , "Wrong path " & sPath
'-----------------------------
If bFolder Then ' EXTRACT FUNCNAME
   sFuncName = ExtractFuncName(sCode)
   If sFuncName <> "" Then
       sPath = BuildPath(sPath, sFuncName & ".vba")
   Else
       sPath = BuildPath(sPath, FilenameWithoutExtension(TempFileName) & ".vba")
   End If
End If
'-----------------------------
   bFolder = WriteStringToFile(sPath, sCode)
'-----------------------------
ExitHere:
        If bFolder Then
            FuncToDisk = sPath
            Debug.Print "Code is saved to " & sPath
        End If
        
        Exit Function
'--------------
ErrHandle:
        ErrPrint "FuncToDisk", Err.Number, Err.Description
        Err.Clear: Resume ExitHere
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
' Function to find Function/ Sub Name in Text
'------------------------------------------------------------------------------------------------------------------------------------------
Private Function ExtractFuncName(sCode As String) As String
Dim iL As Integer, sRes As String, sWork As String

If sCode = "" Then Exit Function
sWork = ClearComments(sCode)

iL = InStr(1, sWork, "Function", vbTextCompare)
If iL > 0 Then
        sRes = NextWord(sWork, "Function", iL): GoTo ExitHere
End If

iL = InStr(1, sWork, "Sub", vbTextCompare)
If iL > 0 Then
        sRes = NextWord(sWork, "Sub", iL): GoTo ExitHere
End If

iL = InStr(1, sWork, "Property", vbTextCompare)
If iL > 0 Then
        sRes = NextWord(sWork, "Property", iL): GoTo ExitHere
End If
'------------------------------------
ExitHere:
     ExtractFuncName = sRes '!!!!!!!!!!!!!!
End Function
'======================================================================================================================================================
'  Clear all comments
'======================================================================================================================================================
Public Function ClearComments(sCode As String) As String
Dim sRes As String, sWork As String
Dim CODES() As String, I As Integer, nDim As Integer

On Error GoTo ErrHandle
'------------------------
       If sCode = "" Then Exit Function
       CODES = Split(sCode, vbCrLf): nDim = UBound(CODES)
       '-----------------------
       For I = 0 To nDim
           If Trim(CODES(I)) <> "" Then
              sWork = RemoveRowComment(CODES(I))
              If sWork <> "" Then sRes = IIf(sRes <> "", sRes & vbCrLf, "") & sWork
           Else
              sRes = IIf(sRes <> "", sRes & vbCrLf, "") & CODES(I)
           End If
       Next I
'------------------------
ExitHere:
       ClearComments = sRes '!!!!!!!!!!!
       Exit Function
'--------------------
ErrHandle:
       ErrPrint2 "ClearComments", Err.Number, Err.Description
       Err.Clear
End Function

'-----------------------------------------------------------------------------------------------------------------------------------------
'  Remove Comment for Row
'-----------------------------------------------------------------------------------------------------------------------------------------
Private Function RemoveRowComment(sRow As String) As String
Dim iL As Integer, sRes As String

     If Left(Trim(sRow), 1) = Chr(39) Then Exit Function  ' whole row is comment
     iL = CalcCommentMarkPosition(sRow)
     If iL > 0 Then
           sRes = Left(sRow, iL - 1)
     Else
           sRes = sRow
     End If
'------------------------
ExitHere:
     RemoveRowComment = sRes '!!!!!!!!!!!!!!
End Function
'----------------------------------------------------------------------------------------------------------------------------------------------------------
' Search comment-mark position outside of double-quate string
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CalcCommentMarkPosition(sRow As String) As Integer
Dim sParen() As String, iRes As Integer, iL As Integer, iR As Integer
Dim sWork As String

       iL = InStr(1, sRow, Chr(39)): If iL = 0 Then Exit Function
       iR = InStr(1, sRow, Chr(34)): If iR = 0 Then GoTo ExitHere  '-------------------------------------- No any "-mark
'------------------------------  So IL > 0 and IR > 0, we should check the situation of "'"

        Do While iL > 0
            sParen = SplitParen(sRow, Chr(34), Chr(34), iR - 1)
                 ' Debug.Print vbTab & "--> " & Join(sParen, ";")
            If sParen(0) = "" Then                                   '------------------------------------- No " "-situation
               Exit Do ' We Found IL
            Else    '-------------------------------------------------------------------------------------- There is " "-situation
               iR = Len(sParen(0)) + Len(sParen(1)) + 4 ' start of row-tail (after " ")
               
               If iL > iR Then                                       '-------------------------------------- "'" lays right of " "
                       '  continue search
               ElseIf iL < Len(sParen(0)) Then                       ' -------------------------------------- "'" lays left of " "
                     Exit Do   ' IL is start of comment
               Else                                                  ' --------------------------------------"'" lays inside of " "
                     iL = InStr(iL + 1, sRow, Chr(39)) ' Look for new position of iL
               End If
            End If
        Loop
       
'------------------------------
ExitHere:
       CalcCommentMarkPosition = iL '!!!!!!!!!!!!!!!!
End Function


'==========================================================================================================================================
'  Get Proper Func Name
'==========================================================================================================================================
Public Function GetFuncName(FuncID As Long, Optional FuncTBL As String = "$$FLDS", Optional ObjTBL As String = "$$OBJECTS") As String
Dim sFunc As String, sModule As String, ModID As Long

On Error Resume Next
'-----------------------------------------------------
    sFunc = Nz(DLookup("FldName", SHT(FuncTBL), "ID = " & FuncID), "")
    ModID = Nz(DLookup("FldParent", SHT(FuncTBL), "ID = " & FuncID), 0)
    If ModID > 0 Then
         sModule = Trim(Nz(DLookup("ObjectName", SHT(ObjTBL), "ID = " & ModID), ""))
         If sModule <> "" Then sModule = sModule & "."
    End If
'--------------------------------
ExitHere:
   GetFuncName = sModule & sFunc '!!!!!!!!!!!
End Function
'=============================================================================================================================================
' Save all code to subfolder CODE
'=============================================================================================================================================
Public Sub SaveAllCode()
Dim MyVBA As New cVBA

On Error GoTo ErrHandle
'-------------------------
    Call MyVBA.ExportAllCodeToFiles
'-------------------------
ExitHere:
    MsgBox "All VBA modules save to files in the folder CODE", vbInformation, "SaveAllCode"
    Set MyVBA = Nothing
    Exit Sub
'----------
ErrHandle:
    ErrPrint "SaveAllCode", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Sub

'=============================================================================================================================================
' Comment Text - Insert comment mark for text in clipboard
'=============================================================================================================================================
Public Sub CommentText()
Dim sText As String
Dim txt() As String, nTXT As Integer, I As Integer

Const sMark As String = "'"

   On Error GoTo ErrHandle
'--------------------------------
   sText = FromClipboard()
   If sText = "" Then
       Debug.Print "No Text in clipboard"
       Beep
       Exit Sub
   End If
'--------------------------------
   txt = Split(sText, vbCrLf): nTXT = UBound(txt)
   
   For I = 0 To nTXT
      txt(I) = sMark & txt(I)
   Next I
'--------------------------------
   sText = Join(txt, vbCrLf)
   ToClipBoard sText
   Debug.Print "The text in clipboard is commented now"
'--------------------------------
ExitHere:
   Debug.Print "Commented Text Is Placed in Clipboard"
   Exit Sub
'--------------
ErrHandle:
   ErrPrint "CommentText", Err.Number, Err.Description
   Err.Clear
End Sub
'=============================================================================================================================================
' Function is SHELL for VBA Suppurt
'=============================================================================================================================================
Public Function DOCDOC(Optional sCommand As String) As String
Dim sRes As String, sCMD As String

On Error GoTo ErrHandle
'----------------------------------------------
    sCMD = IIf(sCommand = "", "/?", sCommand)
Select Case sCMD
Case "/?":                 ' Display Option
     sRes = BuildOptionList("", "/build", "Build the CODEBASE", 2)
     sRes = BuildOptionList(sRes, "/backup", "Backup current db", 2)
     sRes = BuildOptionList(sRes, "/classreuse", "Set Class Attributes to multiuse")
     sRes = BuildOptionList(sRes, "/commenta", "Set Text in clipboard is commented")
     sRes = BuildOptionList(sRes, "/compile", "Compile VBA Project Now")
     sRes = BuildOptionList(sRes, "/divider", "Copy standard divider to ClipBoard")
     sRes = BuildOptionList(sRes, "/exportcode", "Export all vba to CODE")
     sRes = BuildOptionList(sRes, "/codetodisk", "Save vba from clopboard  to CODE")
     
     sRes = BuildOptionList(sRes, "/printmod", "Build Module Template")
     sRes = BuildOptionList(sRes, "/printfunc", "Build Function Template")
     sRes = BuildOptionList(sRes, "/reupdescr", "Place description from $$OBJECTS to real objects")
     sRes = BuildOptionList(sRes, "/charcode", "Show Char Code for Unicode from clipboard")
     
     
Case "/build":
     If MsgBox("Do you want to build/rebuild the CODEBASE?", vbQuestion + vbYesNoCancel, "DOCDOC") = vbYes Then
        Call BuildDocTables
     End If
Case "/compile":
     Call CompileNow
     sRes = "The GRACLE is Compile"
Case "/printfunc":
     Call PrintFunc
Case "/printmod":
     Call PrintMod
Case "/divider":
     Call Divider
Case "/classreuse":
     Call SetClassAttributeToMultiUse
Case "/reupdescr":
     Call ReUpDoc
Case "/commenta":
     Call CommentText
Case "/backup"
     Call BackupDB
Case "/exportcode"
     Call ExportAllModules
Case "/codetodisk"
     Call FuncToDisk
Case "/charcode"
     Call CharCode(True, True)
Case Else
End Select
'----------------------------------------------
ExitHere:
    DOCDOC = sRes '!!!!!!!!!!!!!
    If sRes <> "" Then Debug.Print sRes
    Exit Function
'---------
ErrHandle:
    ErrPrint "DOCDOC", Err.Number, Err.Description
    Err.Clear
End Function
Private Function BuildOptionList(sList As String, sKey As String, sDescription As String, Optional nTabs As Integer = 1, Optional DLM As String = vbCrLf) As String
    BuildOptionList = IIf(sList <> "", sList & DLM, "") & sKey & String(nTabs, vbTab) & "- " & sDescription '!!!!!!!!!!!!
End Function

'=============================================================================================================================================
' Create Note to DB
'=============================================================================================================================================
Public Sub NoteNote(Optional iSet As Integer = 0)
Dim sMsg As String

Const Note_ParamName As String = "NOTENOTE"

    On Error GoTo ErrHandle
'-----------------------
If iSet = 0 Then
    sMsg = InputBox("Set Note to this DB", "Note to DB")
    If sMsg <> "" Then
        Call SetLocal(Note_ParamName, "Note is created " & Now(), "The note for this DB", , sMsg)
    End If
    MsgBox "The note is placeâ to local table", vbInformation, "NoteNote"
ElseIf iSet = 1 Then
    sMsg = GetLocal(Note_ParamName, True)
    MsgBox sMsg, vbInformation, "NoteNote"
ElseIf iSet = 2 Then
    sMsg = InputBox("Add Note to this DB", "Note to DB")
    If sMsg <> "" Then
        sMsg = sMsg & vbCrLf & GetLocal(Note_ParamName, True)
        Call SetLocal(Note_ParamName, "Note is created " & Now(), "The note for this DB", , sMsg)
     End If
End If
'---------------
ExitHere:
    Exit Sub
'------------
ErrHandle:
    ErrPrint2 "NoteNote", Err.Number, Err.Description, MOD_NAME
End Sub

'=============================================================================================================================================
' Create Document of all dbObjects
'=============================================================================================================================================
Public Sub BuildDocTables()
Dim OBJS() As dbObject, nObj As Integer, FLDS() As dbField, nFld As Integer
Dim I As Integer, IDD As Integer

On Error GoTo ErrHandle
'----------------------------------------------
If IsEntityExist("$$OBJECTS") Then
    If MsgBox("Some information about DB Objects is existed and stored here. " & _
        vbCrLf & "Do you want remove all object's data and reload it?", vbYesNoCancel, "DOCDOC") = vbYes Then
        If MsgBox("Do you want restore description from doctables to the physical objects?", vbYesNo, "DOCDOC") = vbYes Then
             Call ReUpDoc   ' DESCRIPTION RESTORE
        End If
    Else
        Exit Sub
    End If
End If

DoCmd.Hourglass True
'----------------------------------------------
' 0. RECREATE DOC TABLES
     If Not RecreateDocTables Then Exit Sub
'----------------------------------------------
' 1. PROCESS TABLES
OBJS = GetObjectListFromDB(1): nObj = UBound(OBJS)
Call WriteObjects(OBJS)
For I = 0 To nObj
     IDD = GetObjectID(OBJS(I).Name, OBJS(I).Type)
       If IDD > 0 Then
            FLDS = GetTableFields(OBJS(I).Name)
            Call WriteFlds(FLDS, IDD)
       End If
Next I
'----------------------------------------------
' 2. PROCESS QUERIES
OBJS = GetObjectListFromDB(5): nObj = UBound(OBJS)
Call WriteObjects(OBJS)
For I = 0 To nObj
     IDD = GetObjectID(OBJS(I).Name, OBJS(I).Type)
       If IDD > 0 Then
            FLDS = GetQueryFlds(OBJS(I).Name)
            Call WriteFlds(FLDS, IDD)
       End If
Next I
'----------------------------------------------
' 3. PROCESS FORMS
OBJS = GetObjectListFromDB(-32768): nObj = UBound(OBJS)
Call WriteObjects(OBJS)
'----------------------------------------------
' 4. PROCESS REPORTS
OBJS = GetObjectListFromDB(-32764): nObj = UBound(OBJS)
Call WriteObjects(OBJS)
'----------------------------------------------
' 5. PROCESS MACROS
OBJS = GetObjectListFromDB(-32766): nObj = UBound(OBJS)
Call WriteObjects(OBJS)
'----------------------------------------------
' 6. PROCESS VBA
OBJS = GetObjectListFromDB(-32761): nObj = UBound(OBJS)
Call WriteObjects(OBJS)
For I = 0 To nObj
     IDD = GetObjectID(OBJS(I).Name, OBJS(I).Type)
       If IDD > 0 Then
            FLDS = VBAListProc(OBJS(I).Name)
            Call WriteFlds(FLDS, IDD)
       End If
Next I
'----------------------------------------------
ExitHere:
     DoCmd.Hourglass False
     MsgBox "The Information has stored in doctables($$OBJECTS , $$FLDS)", vbInformation, "DOCDOC"
     Exit Sub
'-------------------
ErrHandle:
     ErrPrint "BuildDocTables", Err.Number, Err.Description
     Err.Clear: DoCmd.Hourglass False
End Sub
Public Function AllProcs(ByVal strDatabasePath As String, ByVal strModuleName As String)
    Dim appAccess As Access.Application
    Dim mdl As Module
    Dim lngCount As Long
    Dim lngCountDecl As Long
    Dim lngI As Long
    Dim strProcName As String
    Dim astrProcNames() As String
    Dim intI As Integer
    Dim strMsg As String
    Dim lngR As Long

    Set appAccess = New Access.Application

    appAccess.OpenCurrentDatabase strDatabasePath
    ' Open specified Module object.
    appAccess.DoCmd.OpenModule strModuleName
    ' Return reference to Module object.
    Set mdl = appAccess.Modules(strModuleName)
    ' Count lines in module.
    lngCount = mdl.CountOfLines
    ' Count lines in Declaration section in module.
    lngCountDecl = mdl.CountOfDeclarationLines
    ' Determine name of first procedure.
    strProcName = mdl.ProcOfLine(lngCountDecl + 1, lngR)
    ' Initialize counter variable.
    intI = 0        ' Redimension array.
    ReDim Preserve astrProcNames(intI)
    ' Store name of first procedure in array.
    astrProcNames(intI) = strProcName
    ' Determine procedure name for each line after declarations.
    For lngI = lngCountDecl + 1 To lngCount
        ' Compare procedure name with ProcOfLine property value.
        If strProcName <> mdl.ProcOfLine(lngI, lngR) Then
            ' Increment counter.
            intI = intI + 1
            strProcName = mdl.ProcOfLine(lngI, lngR)
            ReDim Preserve astrProcNames(intI)
            ' Assign unique procedure names to array.
            astrProcNames(intI) = strProcName
        End If
    Next lngI
    strMsg = "Procedures in module '" & strModuleName & "': " & vbCrLf & vbCrLf
    For intI = 0 To UBound(astrProcNames)
        strMsg = strMsg & astrProcNames(intI) & vbCrLf
    Next intI
    ' Message box listing all procedures in module.
    Debug.Print strMsg
    appAccess.CloseCurrentDatabase
    appAccess.Quit
    Set appAccess = Nothing
End Function
'======================================================================================================================================================
' Set DB NOTE
'======================================================================================================================================================
'==============================================================================================================================================
' Rewrite descriptions from doctables to physical objects
'==============================================================================================================================================
Public Sub ReUpDoc()
     Call ReUpDescription("$$OBJECTS", "$$FLDS")
     Debug.Print "All description set from $$OBJECTS to objects"
End Sub
'======================================================================================================================================================
'  Generate Function Template
'======================================================================================================================================================
Public Sub GenerateFunc()
Dim sFuncName As String, sDescription As String
Dim sRes As String
    sFuncName = "SomeFunction"
    sDescription = "This Function was generated " & Now() & " for proporses"
    sRes = PrintFunc(False, "Function", "I As Integer", "Public", sFuncName, "String", sDescription)
    If sRes <> "" Then
         ToClipBoard sRes
         MsgBox "The function " & sFuncName & " was generated and put in the clipboard", , "Generate Func"
    End If
End Sub
'=====================================================================================================================================================
' GENERATE HEADER OF MODULE
'=====================================================================================================================================================
Public Function GenerateMod() As String
Dim ModName As String, sRes As String, sWorkOut As String, sWRK() As String, nSWRK As Integer
Dim DESCR As String, Version As String, Author As String, Copyright As String

Const DLM As String = "¤"
Const SepChar As String = "*"

On Error GoTo ErrHandle
'----------------------------------------------------------
ModName = InputBox("Provide Module Name", "GenerateMod", GetCurrentMod())
If ModName = "" Then Exit Function
If IsModule(ModName) Then
     sWorkOut = ExtractHeader(ModName, DLM)
     If sWorkOut <> "" Then
         If Left(sWorkOut, 1) <> DLM Then
             sWRK = Split(sWorkOut, DLM): nSWRK = UBound(sWRK)
             If nSWRK < 4 Then Err.Raise 1000, , "Wrong Format of Header"
             DESCR = sWRK(1)
             If sWRK(1) = "" Then DESCR = sWRK(0)
             Version = sWRK(2)
             Author = sWRK(3)
             Copyright = sWRK(4)
         End If
     End If
End If
'--------------------------------------------------------
NewHeader:
     If DESCR = "" Then DESCR = "THIS MODILE ..."
     If Version = "" Then Version = "v. " & "01/" & Format(Now(), "yyyymmdd")
     If Author = "" Then Author = "@ " & "Valery Khvatov" & " " & "(valery.khvatov@gmail.com)"
     If Copyright = "" Then Copyright = "(c) DigitalXpert Inc."
'---------------------------------------------------------
CompileHere:
     sWorkOut = Chr(39) & String(SepLineCounter, SepChar)
     DESCR = InputBox("Type Description of Module " & ModName, "GenerateMod", DESCR)
     Version = InputBox("Type Version of Module " & ModName, "GenerateMod", Version)
     Author = InputBox("Type Author Info of Module " & ModName, "GenerateMod", Author)
     Copyright = InputBox("Type Copyright of Module " & ModName, "GenerateMod", Copyright)
     
     sRes = sWorkOut & vbCrLf & sWorkOut & vbCrLf & sWorkOut
     sRes = sRes & vbCrLf & MultiLineCommenting(DESCR) & vbCrLf & _
            Chr(39) & " " & Author & ",  " & IIf(Copyright <> "", Copyright & ",  ", "") & _
            Version & vbCrLf & sRes
            
     Call ToClipBoard(sRes)
     MsgBox "The Header of module " & ModName & " is generated and copy to ClipBoard", vbInformation, ""
'----------------------------------------------------------
ExitHere:
     GenerateMod = sRes '!!!!!!!!!!!!!
     Exit Function
'-------------
ErrHandle:
     If Err.Number = 1000 Then
        Err.Clear
                If MsgBox("Can't recognize Header of module " & ModName & _
                ". But you can generate new one. Should we Process it?", _
                vbYesNo + vbExclamation, "GenerateMod") = vbYes Then GoTo NewHeader
     Else
        ErrPrint "GenerateMod", Err.Number, Err.Description
        Err.Clear: Resume ExitHere
     End If
End Function
'------------------------------------------------------------------------------------------------------------------
' Multiline Commenting
'------------------------------------------------------------------------------------------------------------------
Private Function MultiLineCommenting(str As String) As String
    MultiLineCommenting = Chr(39) & " " & Replace(str, vbCrLf, vbCrLf & Chr(39) & " ") '!!!!!!!!!!!!
End Function
'===================================================================================================================
' Check If Module Exists
'=====================================================================================================================
Public Function IsModule(ModName As String) As Boolean
On Error GoTo ErrHandle
    IsModule = ModName = CurrentDb.Containers("Modules").Documents(ModName).Name
    Exit Function
'-----------------
ErrHandle:
    Err.Clear
End Function
'======================================================================================================================================================
' Create New Module From String
'======================================================================================================================================================
Public Function CreateNewModuleFromString(pModName As String, sBody As String) As String
Dim MyModule As Module, sRes As String

    On Error GoTo ErrHandle
'-----------------------------------------
If pModName = "" Then Exit Function
If sBody = "" Then Exit Function

    DoCmd.RunCommand acCmdNewObjectModule ' Create the module.
    Set MyModule = Application.Modules.Item(Application.Modules.Count - 1)

    DoCmd.Save acModule, MyModule
    DoCmd.Close acModule, MyModule, acSaveYes
    DoCmd.Rename pModName, acModule, MyModule
'-----------------------------------------
    MyModule.AddFromString sBody
    DoCmd.Save acModule, MyModule
    
    sRes = pModName
'-----------------------------------------
ExitHere:
    CreateNewModuleFromString = sRes '!!!!!!!!!!!!!
    Set MyModule = Nothing
    Exit Function
'----------------
ErrHandle:
    ErrPrint2 "CreateNewModuleFromString", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Resume ExitHere
End Function

'=====================================================================================================================
' Function Return Structure that contains some parts:
'        (0) - FULL HEADER (as Is)
'        (1) - ONLY DESCRIPTION (without version and copyright)
'        (2) - ONLY VERSION INFO
'        (3) - ONLY COPYRIGHT INFO
'        (4) - ONLY AUTHORITY INFO
'=====================================================================================================================
Public Function ExtractHeader(Optional sModuleName As String, Optional DLM As String = "¤") As String
Dim ModName As String, mdl As Module, lngCount As Long, lngCountDecl As Long
Dim sRes As String, SRS() As String, nDim As Integer, I As Integer
Dim sWork As String, sChar As String, sProb(4) As String, sPRB As String, bSpec As Boolean

Const DLMRowOccur As Integer = 10                    ' Repeated Symbols
Const VERSINFO As String = "v.;v;V.;V.;Vers;VERS;VERSION;version"
Const COPYRIGHTINFO As String = "(c);©;Copr.;Copyright;®;Cprght."
Const AUTHORITYINFO As String = "@;developed by;DEVELOPED BY"

On Error GoTo ErrHandle
'-------------------------------------------------
ModName = IIf(sModuleName = "", GetCurrentMod(), sModuleName)
DoCmd.OpenModule ModName: Set mdl = Modules(ModName)  ' Open specified Module object
    lngCount = mdl.CountOfLines                       ' Count lines in module.
    If lngCount = 0 Then Exit Function                ' No source code in the module
    lngCountDecl = mdl.CountOfDeclarationLines        ' Count lines in Declaration section in module.
    If lngCountDecl = 0 Then Exit Function            ' No Declaration Lines
    sWork = mdl.LINES(1, lngCountDecl)                ' Get All Declaration Lines
    If sWork = "" Then Exit Function                  ' Something wrong with declaration
'-------------------------------------------------
SRS = Split(sWork, vbCrLf): nDim = UBound(SRS)
For I = 0 To nDim
    SRS(I) = Trim(SRS(I))
    If SRS(I) = "" Then GoTo NextLine                 ' Skip all empty lines
    If Left(SRS(I), 1) <> "'" Then Exit For           ' Assume that header is placed before any code
    
    sWork = Trim(Right(SRS(I), Len(SRS(I)) - 1))      ' Get valuable text
    If sWork = "" Then GoTo NextLine                  ' Commented empty line
    sChar = Left(sWork, 1)                            ' Check first symbol of valuable text to skip full-row comment
    If StringCountOccurrences(sWork, sChar) > DLMRowOccur Then GoTo NextLine  ' Ignor Full-Row Comment
    
    If sProb(0) <> "" Then sProb(0) = sProb(0) & vbCrLf
    sProb(0) = sProb(0) & sWork                       ' Add text to our reading stack
    '--------------------------------------------------------------------------------------
    ' CHECK ANY SPECIAL COMMENTS
    sPRB = GetValueForWRDS(sWork, VERSINFO)
    If sPRB <> "" Then
       sProb(2) = Trim(sProb(2) & " " & sPRB): bSpec = True
    End If
    sPRB = GetValueForWRDS(sWork, COPYRIGHTINFO)
    If sPRB <> "" Then
       sProb(3) = Trim(sProb(3) & " " & sPRB): bSpec = True
    End If
    sPRB = GetValueForWRDS(sWork, AUTHORITYINFO)
    If sPRB <> "" Then
       sProb(4) = Trim(sProb(4) & " " & sPRB): bSpec = True
    End If
    '--------------------
    If Not bSpec Then
        If sProb(1) <> "" Then sProb(1) = sProb(1) & vbCrLf
        sProb(1) = sProb(1) & UCase(sWork)             ' Add text to our reading stack
    End If
    bSpec = False
    '--------------------------------------------------------------------------------------
NextLine:
Next I
    sRes = Join(sProb, DLM)
'-----------------------------
ExitHere:
    ExtractHeader = sRes '!!!!!!!!!!!!!!!!!
    Exit Function
'-----------------
ErrHandle:
     ErrPrint "ExtractHeader", Err.Number, Err.Description
     Err.Clear: Resume ExitHere
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Get Current Module Name
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetCurrentMod() As String
On Error Resume Next
    GetCurrentMod = VBE.ActiveCodePane.CodeModule '!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' SEARCH WORDS FROM LIST AND RETURNS VALUE FOR ITS
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function GetValueForWRDS(sLine As String, sWords As String, Optional DLM As String = ";") As String
Dim sRes As String, WRDS() As String, I As Integer, nDim As Integer
Dim iL As Long

On Error GoTo ErrHandle
'------------------------
WRDS = Split(sWords, DLM): nDim = UBound(WRDS)
For I = 0 To nDim
    iL = IsWordInStr(sLine, WRDS(I))
    If iL > 0 Then        ' FOUND KEYWORD
         sRes = WRDExtract(sLine, WRDS(I), iL)
         Exit For
    End If
Next I
If Left(sRes, 1) = "." Then sRes = Trim(Right(sRes, Len(sRes) - 1))
'------------------------
ExitHere:
         GetValueForWRDS = sRes '!!!!!!!!!!!!
         Exit Function
'---------
ErrHandle:
         ErrPrint "GetValueForWRDS", Err.Number, Err.Description
End Function
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Extract Value from iPos to Next Divider
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function WRDExtract(str As String, Word As String, iPos As Long, Optional DelimChars As String = ",;") As String
Dim sRes As String, sWork As String, VALS() As String, nDim As Integer, I As Integer
Dim SeLST() As String, nSep As Integer, iL As Long
Const UDLM As String = "¤"

On Error GoTo ErrHandle
'------------------------
sWork = Right(str, Len(str) - iPos): If sWork = "" Then Exit Function
sWork = Replace(sWork, "  ", UDLM): sWork = Replace(sWork, vbTab, UDLM)

SeLST = Split(StrConv(DelimChars, 64), Chr(0)): nSep = UBound(SeLST)
For I = 0 To nSep
     sWork = Replace(sWork, SeLST(I), UDLM)
Next I

VALS = Split(sWork, UDLM): nDim = UBound(VALS)
sRes = Trim(Replace(VALS(0), Word, ""))
sRes = Trim(Replace(sRes, "=", "")): sRes = Trim(Replace(sRes, ":", "")): sRes = Trim(Replace(sRes, "-", ""))
'------------------------------
ExitHere:
         WRDExtract = sRes '!!!!!!!!!!!!
         Exit Function
'---------
ErrHandle:
         ErrPrint "WRDExtract", Err.Number, Err.Description
End Function
'======================================================================================================================================================
' Check if VBA Module Exists
'======================================================================================================================================================
Public Function IsVBAModule(sModuleName As String, Optional VBProjectName As String) As Boolean
Dim sProjectName As String, bRes As Boolean
   
   On Error GoTo ErrHandle
'--------------------------
   If VBProjectName <> "" Then
      sProjectName = VBProjectName
   Else
      sProjectName = Application.VBE.VBProjects(1).Name
   End If
   If Application.VBE.VBProjects(sProjectName).VBComponents(sModuleName).Name = sModuleName Then
       bRes = True
   End If
'--------------------------
ExitHere:
   IsVBAModule = bRes '!!!!!!!!!!
   Exit Function
'-------------------
ErrHandle:
   Err.Clear
End Function
'======================================================================================================================================================
' Create VB Module With Name
'======================================================================================================================================================
Public Function CreateVBA(sModName As String, Optional VBProjectName As String) ', Optional ComponentType As vb) As Boolean
Dim bRes As Boolean, vbc As Object, sProjectName As String


Const VBA_StdModule As Long = 1

   On Error GoTo ErrHandle
'-------------------------
  If sModName = "" Then Exit Function
  If VBProjectName <> "" Then
      sProjectName = VBProjectName
   Else
      sProjectName = Application.VBE.VBProjects(1).Name
   End If
   
   Set vbc = Application.VBE.VBProjects(sProjectName).VBComponents.Add(VBA_StdModule)
   
   vbc.Name = sModName
   
   VBWork = ModHeader(sModName)
   
   
   Call CompileNow
   
       With vbc.CodeModule
          .InsertLines 1, VBWork
       End With
        
   DoCmd.Save acModule, sModName
   bRes = True
'-------------------------
ExitHere:
   CreateVBA = bRes '!!!!!!!!!!
   Set vbc = Nothing
   Exit Function
'------------
ErrHandle:
    ErrPrint2 "CreateVBA", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'======================================================================================================================================================
' The function will check the opening of a module (available in Modules) with bLoad = True
'                                                           with bLoad = False the module is closed by command
' Returns: 1 = The module has been verified and successfully opened;
'          2 = The module is checked, not open, but open the clock
'          3 = module closed on demand
'======================================================================================================================================================
Public Function CheckLoaded(ModName As String, Optional bLoad As Boolean = True) As Integer
Dim accObj As AccessObject              ' Renewed variable, assigned to each object module/form/report.
Dim sWork As String
Dim iRes As Integer

On Error GoTo ErrHandler
'-----------------------------------------------------
If bLoad = True Then                    ' CHECK THE OPENNESS OF MODULE
        sWork = ModName: iRes = 1
      If Left(sWork, 5) = "Form_" Then                                                  ' FORMS
        sWork = Replace(sWork, "Form_", "")
        Set accObj = CurrentProject.AllForms(sWork)
        If Not accObj.IsLoaded Then
           DoCmd.OpenForm accObj.Name, acDesign, WindowMode:=acHidden
                    iRes = 2
        End If
      ElseIf Left(sWork, 7) = "Report_" Then                                            ' REPORTSÛ
        sWork = Replace(sWork, "Report_", "")
        Set accObj = CurrentProject.AllReports(sWork)
        If Not accObj.IsLoaded Then
           DoCmd.OpenReport accObj.Name, acDesign, WindowMode:=acHidden
                    iRes = 2
        End If
      Else                                                                              ' MODULES
        Set accObj = CurrentProject.AllModules(sWork)
        If Not accObj.IsLoaded Then
           DoCmd.OpenModule accObj.Name
                    iRes = 2
        End If
      End If
      '------------------------------------------------------------------
Else                                  ' CLOSE MODULE
            DoCmd.Close acModule, accObj.Name, acSaveNo
            iRes = 3
End If
'-----------------------------------------------------
ExitHere:
    CheckLoaded = iRes '!!!!!!!!!!!!!!!!!
    Exit Function
'----------------------------------
ErrHandler:
    Err.Clear
    iRes = 0
    Resume ExitHere
End Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'----------------------------------------------------------------------------------------------------------------------------------------------
' Error Handler
'----------------------------------------------------------------------------------------------------------------------------------------------
Private Function ErrPrint(FuncName As String, ErrNumber As Long, ErrDescription As String, Optional bDebug As Boolean = True, _
                                                                                               Optional sModName As String = "mod_DOCUMENTER") As String
Dim sRes As String
Const ErrChar As String = "#"
Const ErrRepeat As Integer = 60

sRes = String(ErrRepeat, ErrChar) & vbCrLf & "ERROR OF [" & sModName & ": " & FuncName & "]" & vbTab & "ERR#" & ErrNumber & vbTab & Now() & _
       vbCrLf & ErrDescription & vbCrLf & String(ErrRepeat, ErrChar)
If bDebug Then Debug.Print sRes
'----------------------------------------------------------
ExitHere:
       Beep
       ErrPrint = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------------------
' Function Split some description for rows limited by lenghth
'------------------------------------------------------------------------------------------------------------------------------------------
Private Function FormatComments(sComments As String, Optional nLimit As Integer = 90, Optional bTranslit As Boolean = True) As String
Dim ROWS() As String, nDim As Integer, I As Integer
Dim sWork As String, sRes As String, W() As String, NW As Integer

On Error GoTo ErrHandle
'---------------------------------------------
If sComments = "" Then Exit Function
sWork = IIf(bTranslit, Translit(sComments), sComments)
ReDim ROWS(0)
'---------------------------------------------
W = SplitByWords(sWork, nLimit): NW = UBound(W): sWork = ""

For I = 0 To NW
   If Len(ROWS(nDim) & " " & W(I)) < nLimit - 2 Then
       If ROWS(nDim) <> "" Then ROWS(nDim) = ROWS(nDim) & " "
       ROWS(nDim) = ROWS(nDim) & W(I)
   Else
       nDim = nDim + 1: ReDim Preserve ROWS(nDim)
       ROWS(nDim) = W(I)
   End If
Next I
'---------------------------
ExitHere:
    FormatComments = Chr(39) & " " & Join(ROWS, vbCrLf & Chr(39) & " ") '!!!!!!!!!!!!!!!!
    Exit Function
'--------------
ErrHandle:
    ErrPrint "FormatComments", Err.Number, Err.Description
    Err.Clear
End Function
'-------------------------------------------------------------------------------------------------------------------------------
' Split By Words With Limit
'-------------------------------------------------------------------------------------------------------------------------------
Private Function SplitByWords(ByVal str As String, Optional nLimit As Integer = 10) As String()
Dim sRes() As String, nRes As Integer
Dim W() As String, nDim As Integer, I As Integer, U() As String
Dim CommonSeparator As String, nSep As Integer, DLM As String, sWork As String
Const WordDLMS As String = " ,.:;=-+/\|!?" & vbCr

'------------------------------------------------------
nRes = -1: ReDim sRes(0)
If str = "" Then GoTo ExitHere
'------------------------------------------------------
' FIRST - SPLIT BY WORDS
CommonSeparator = Chr(30): sWork = str
nSep = Len(WordDLMS)
For I = 1 To nSep - 1
    DLM = Mid(WordDLMS, I, 1)
    sWork = Replace(sWork, DLM, DLM & CommonSeparator)
Next I
W = Split(sWork, CommonSeparator): nDim = UBound(W)
'--------------------------------------------------------
' SECOND - SKIP EMPTY WORDS AND SPLIT LONG WORDS
For I = 0 To nDim
    If Trim(W(I)) <> "" Then
       If Len(W(I)) <= nLimit Then
          nRes = nRes + 1: ReDim Preserve sRes(nRes)
          sRes(nRes) = Trim(W(I))
       Else
          U = SplitString(W(I), nLimit)
          nRes = ImplantArr(sRes, U)
       End If
    End If
Next I
'--------------------------------------------------------
ExitHere:
    SplitByWords = sRes '!!!!!!!!!!!!
End Function
'-------------------------------------------------------------------------------------------------------------------------------
' Function Impalnt some array aftter end another array
'-------------------------------------------------------------------------------------------------------------------------------
Private Function ImplantArr(ByRef MainArr() As String, AddArr() As String) As Integer
Dim nDim As Integer, mDim As Integer, I As Integer
    nDim = UBound(MainArr): mDim = UBound(AddArr)
    ReDim Preserve MainArr(nDim + mDim + 1)
    For I = 0 To mDim
        MainArr(nDim + I + 1) = AddArr(I)
    Next I
'----------------------------------
    ImplantArr = nDim + mDim + 1 '!!!!!!!!!!!!
End Function
'-------------------------------------------------------------------------------------------------------------------------------
' Set Comments to Row
'-------------------------------------------------------------------------------------------------------------------------------
Private Function RowComment(sRow As String) As String
    RowComment = IIf(Left(sRow, 1) = Chr(39), sRow, Chr(39) & " " & sRow) '!!!!!!!!!!!!!!
End Function
'---------------------------------------------------------------------------------------------------------------------------------------
' Lit all object in this databases with description as array
'---------------------------------------------------------------------------------------------------------------------------------------
Private Function GetObjectListFromDB(Optional ObjType As Integer = 1, Optional bExcludeSysteTable As Boolean = True) As dbObject()
   Dim MyObjects() As dbObject, nDim As Long                       ' Âîçâðàùàåìûé ìàññèâ îáúåêòîâ è åãî ðàçìåðíîñòü
   Dim db As DAO.Database, mO As Object, mOBJS As Object           ' Ññûëêè íà îáúåêòû âíóòðè áàçû äàííûõ

On Error GoTo ErrHandle
'------------------------------------------------------
   Set db = CurrentDb()
       Select Case ObjType
       Case 1:           ' ÒÀÁËÈÖÀ
            Set mOBJS = db.TableDefs
       Case 5:           ' ÇÀÏÐÎÑ
            Set mOBJS = db.QueryDefs
       Case -32768:      ' ÔÎÐÌÛ
            Set mOBJS = CurrentProject.AllForms
       Case -32764:      ' ÎÒ×ÅÒÛ
            Set mOBJS = CurrentProject.AllReports
       Case -32766:      ' ÌÀÊÐÎÑÛ
            Set mOBJS = CurrentProject.AllMacros
       Case -32761:      ' ÌÎÄÓËÈ
            Set mOBJS = CurrentProject.AllModules
       Case Else
            Exit Function
       End Select
   nDim = -1: ReDim MyObjects(0)
   For Each mO In mOBJS
        '------------------------------------------------------------------------------
        If bExcludeSysteTable Then
           If Left(mO.Name, 1) = "~" Or Left(mO.Name, 1) = "%" Or Left(mO.Name, 4) = "MSys" Then GoTo LoopNext
        End If
        '------------------------------------------------------------------------------
        'Debug.Print "ÎÁÚÅÊÒ: " & mO.Name
        nDim = nDim + 1: ReDim Preserve MyObjects(nDim)
        MyObjects(nDim).Type = ObjType   ' Òàáëèöà, çàïðîñ è ò.ä.
        MyObjects(nDim).Name = mO.Name
        MyObjects(nDim).Description = GetAccessProp(mO, "Description")
            If ObjType = -32761 Then
                 Call GetVBADetails(MyObjects(nDim))
            End If
            
        MyObjects(nDim).DateCreated = mO.DateCreated
        MyObjects(nDim).Group = GetObjGroup(MyObjects(nDim).Name, MyObjects(nDim).Type)
        MyObjects(nDim).Parent = GetObjParent(MyObjects(nDim).Name, MyObjects(nDim).Type)
        MyObjects(nDim).BODY = GetObjBody(MyObjects(nDim).Name, MyObjects(nDim).Type)
LoopNext:
    Next mO
'----------------------------------------------------------------
ExitHere:
        GetObjectListFromDB = MyObjects '!!!!!!!!!!!!!
        Exit Function
'----------------------------------------------------------------
ErrHandle:
        ErrPrint "GetObjectListFromDB", Err.Number, Err.Description
        Err.Clear
End Function
'--------------------------------------------------------------------------------------------------------------------------------------
' Get Object Body
'--------------------------------------------------------------------------------------------------------------------------------------
Private Function GetObjBody(objName As String, ObjType As Integer) As String
Dim sRes As String, iWork As Integer
'Dim myVBA As cVBA
        If ObjType = 1 Then
            sRes = ""
        ElseIf ObjType = 5 Then
            sRes = CurrentDb.QueryDefs(objName).SQL
        'ElseIf ObjType = -32761 Then
            'Set myVBA = New cVBA
            'sRes = myVBA.GetVBACode(ObjName)
            'Set myVBA = Nothing
        Else
            sRes = ""
        End If
'--------------------------------------------------------------
    GetObjBody = sRes '!!!!!!!!!!!!!!!!!
End Function

'--------------------------------------------------------------------------------------------------------------------------------------
' get parent object for db entity
'--------------------------------------------------------------------------------------------------------------------------------------
Private Function GetObjParent(objName As String, ObjType As Integer) As String
Dim sRes As String, iWork As Integer
        If ObjType = 1 Then
            sRes = CurrentProject.Name
        ElseIf ObjType = 5 Then
            sRes = GetParentForQuery(objName)
        ElseIf ObjType = -32768 Then
            iWork = GetObjGroup(objName, ObjType)
            If iWork = 5 Then                                          'window
                sRes = CurrentProject.Name
            ElseIf iWork = 6 Or iWork = 8 Then                         'lst / subf
                sRes = Right(objName, Len(objName) - 4)
                If Not IsEntityExist(sRes, 1) Then sRes = "q_" & sRes
                If Not IsEntityExist(sRes, 5) Then sRes = ""
            ElseIf iWork = 7 Then                                      'detailed
                sRes = "lst_" & Right(objName, Len(objName) - 2)
                If Not IsEntityExist(sRes, -32768) Then sRes = "_" & Right(sRes, Len(sRes) - 4)
                If Not IsEntityExist(sRes, -32768) Then sRes = ""
            End If
        ElseIf ObjType = -32761 Then
            sRes = CurrentProject.Name
        ElseIf ObjType = -32764 Then
            sRes = CurrentProject.Name
        ElseIf ObjType = -32766 Then
            sRes = CurrentProject.Name
        Else
             sRes = ""
        End If
'--------------------------------------------------------------
    GetObjParent = sRes '!!!!!!!!!!!!!!!!!
End Function
'--------------------------------------------------------------------------------------------------------------------------------------
' Create group by name mnemonic
'--------------------------------------------------------------------------------------------------------------------------------------
Private Function GetObjGroup(objName As String, ObjType As Integer) As Integer
Dim nGroup As Integer

Select Case ObjType
    Case 1:                                 ' ÒÀÁËÈÖÛ
        If Left(objName, 3) = "$$$" Then
          nGroup = -3  ' SYS0
        ElseIf Left(objName, 2) = "$$" Then
          nGroup = -2  ' SYS1
        ElseIf Left(objName, 1) = "$" Then
          nGroup = -1  ' SYS2
        ElseIf Left(objName, 1) = "_" Then
          nGroup = 2   ' Ref
        ElseIf Left(objName, 8) = "_" Then
          nGroup = 3   ' CrossLink
        ElseIf objName = UCase(objName) Then
          nGroup = 1   ' Entity
        Else
          nGroup = 12  ' Other
        End If
    Case 5:                                 ' ÇÀÏÐÎÑÛ
        If IsEntityExist(Right(objName, Len(objName) - 2)) Then
          nGroup = 4    ' QueryExtension
        Else
          nGroup = 12   ' Other
        End If
    Case -32768:                            ' ÔÎÐÌÛ
        If Left(objName, 1) = "_" Then
            nGroup = 5  ' WindowForm
        ElseIf Left(objName, 3) = "lst" Then
            nGroup = 6  ' ListingForm
        ElseIf Left(objName, 2) = "f_" Then
            nGroup = 7  ' DetailedForm
        ElseIf Left(objName, 4) = "subf" Then
            nGroup = 8  ' SubForm
        Else
            nGroup = 12 ' Other
        End If
    Case -32764:                            ' ÎÒ×ÅÒÛ
          nGroup = 12   ' Other
    Case -32766:                            ' ÌÀÊÐÎÑÛ
          nGroup = 12   ' Other
    Case -32761:                            ' ÌÎÄÓËÈ
           If Left(objName, 1) = "c" Then
              nGroup = 10
           Else
              nGroup = 9
           End If
    Case Else:
           nGroup = 12   ' Other
End Select
'-------------------------------------------
    GetObjGroup = nGroup '!!!!!!!!!!!
End Function
'--------------------------------------------------------------------------------------------------------------------------------------
' Check what table for this query
'--------------------------------------------------------------------------------------------------------------------------------------
Private Function GetParentForQuery(QryName As String) As String
Dim sRes As String, sWork As String
Dim iPoint As Long
On Error GoTo ErrHandle
'-------------------------------------------------------------------
    sWork = UCase(CurrentDb.QueryDefs(QryName).SQL)
    iPoint = InStr(1, sWork, "FROM") + 4
    sWork = Right(sWork, Len(sWork) - iPoint)
    iPoint = InStr(1, sWork, vbCrLf)
    If iPoint > 0 Then sWork = Left(sWork, iPoint)
    iPoint = InStr(1, sWork, " ")
    If iPoint > 0 Then sWork = Left(sWork, iPoint)
    iPoint = InStr(1, sWork, ";")
    If iPoint > 0 Then sWork = Left(sWork, iPoint - 1)
    If Left(sWork, 1) = "[" Then sWork = Mid(sWork, 2, Len(sWork) - 2)
    If sWork <> "" Then sRes = sWork
'-------------------------------------------------------------------
ExitHere:
    GetParentForQuery = sRes '!!!!!!!!!!!!!!
    Exit Function
'----------------------
ErrHandle:
    Err.Clear
    Resume ExitHere
End Function
'======================================================================================================================================
' Check Is Table Exist
'======================================================================================================================================
Public Function IsEntityExist(EntityName As String, Optional iType As Integer = 1) As Boolean
Dim sWhere As String, bRes As Boolean
'-----------------------------------------
    sWhere = "(MSysObjects.Type=" & iType & ") AND (MSysObjects.Name)='" & EntityName & "'"
    If Nz(DLookup("MSysObjects.ID", "MSysObjects", sWhere), -1) <> -1 Then bRes = True
'--------------------------------------------
    IsEntityExist = bRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'---------------------------------------------------------------------------------------------------------------------------------------
' Function Process VB Module and Extract Function List
'---------------------------------------------------------------------------------------------------------------------------------------
Private Function VBAListProc(ModName As String) As dbField()
Dim MyVB As New cVBA, FLDS() As dbField
Dim sWork As String, PROCS() As String, nDim As Integer, I As Integer
Dim W As Variant, U() As String, iType As Integer

Const DLM As String = ";"
On Error GoTo ErrHandle

nDim = -1: ReDim FLDS(0)
'-------------------------------------------
    sWork = MyVB.ListProcs(ModName, DLM)
    If sWork = "" Then GoTo ExitHere
    
       PROCS = Split(sWork, DLM): nDim = UBound(PROCS): ReDim FLDS(nDim)
       For I = 0 To nDim
          sWork = Trim(PROCS(I))
          FLDS(I).Name = Split(sWork, "|")(0)
          iType = CInt(Split(sWork, "|")(1))
          
          W = MyVB.ProcessProcedure(ModName, FLDS(I).Name, iType)
          If Not IsEmpty(W) Then
              FLDS(I).BODY = Trim(W(0))
              If W(4) <> "" Then FLDS(I).fldSize = CInt(W(4))
              FLDS(I).Description = ClearFirstApostrophe(CStr(W(3)))
              If Trim(CStr(W(1))) <> "" Then
                U = Split(W(1), DLM)
                FLDS(I).FldCategory = U(1)
                FLDS(I).FldType = UCase(U(2))
                If U(5) <> "" Then FLDS(I).META = U(5)
              End If
              
              If CStr(W(2)) <> "" Then FLDS(I).META = FLDS(I).META & vbCrLf & W(2)
          End If
          
       Next I
       
'-------------------------------------------
ExitHere:
    VBAListProc = FLDS '!!!!!!!!!!!!!!!!!
    Set MyVB = Nothing
    Exit Function
'--------------------
ErrHandle:
    ErrPrint "VBAListProc: " & ModName & "." & Left(sWork, 10), Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------
' Function Return String without First apostrofe
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Function ClearFirstApostrophe(str As String) As String
Dim W() As String, nDim As Integer, I As Integer
Dim sRes As String
    If str = "" Then Exit Function
    W = Split(str, vbCrLf): nDim = UBound(W)
    If nDim = -1 Then Exit Function
    '----------------------------------
    For I = 0 To nDim
        If sRes <> "" Then sRes = sRes & vbCrLf
        If Asc(Left(W(I), 1)) = 39 Then W(I) = Trim(Right(W(I), Len(W(I)) - 1))
        sRes = sRes & W(I)
    Next I
'-------------------------------
ExitHere:
   ClearFirstApostrophe = sRes '!!!!!!!!!!!
End Function
'---------------------------------------------------------------------------------------------------------------------------------------------
' Function Get VBA Module Header & Code
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GetVBADetails(DBOBJ As dbObject)
Dim VBAProcessor As New cVBA, sCode As String, sDECLARATION As String
On Error GoTo ErrHandle
'---------------------------
   sCode = VBAProcessor.GetVBACode(DBOBJ.Name)
   sDECLARATION = VBAProcessor.GetModuleDeclaration(DBOBJ.Name)
'---------------------------
  DBOBJ.BODY = sCode
  DBOBJ.Description = GetVBADescription(sDECLARATION)
ExitHere:
    Set VBAProcessor = Nothing
    Exit Sub
'-----------
ErrHandle:
    ErrPrint "GetVBADetails", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Sub
'---------------------------------------------------------------------------------------------------------------------------------------------
' Function Look For Description Through Module Declartion
' 1) It Is Commented Line
' 2) It Is Not a Separate Line
' 3) It is Placed on The top
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetVBADescription(sDECLARATION As String) As String
Dim sRes As String, SRS() As String, nDim As Integer, I As Integer
Dim sWork As String, sChar As String

Const DLMRowOccur As Integer = 10
'-----------------------------
If sDECLARATION = "" Then Exit Function
SRS = Split(sDECLARATION, vbCrLf): nDim = UBound(SRS)
For I = 0 To nDim
    SRS(I) = Trim(SRS(I))
    If SRS(I) = "" Then GoTo NextLine
    If Left(SRS(I), 1) <> "'" Then GoTo NextLine ' Process Only Full-Row Comments
    sWork = Trim(Right(SRS(I), Len(SRS(I)) - 1))
    If sWork = "" Then GoTo NextLine
    sChar = Left(sWork, 1)
    If StringCountOccurrences(sWork, sChar) > DLMRowOccur Then GoTo NextLine  ' Ignor Full-Row Comment
    If sRes <> "" Then sRes = sRes & vbCrLf
    sRes = sRes & sWork
NextLine:
Next I
'-----------------------------
    GetVBADescription = sRes '!!!!!!!!!!!!!!!!!
End Function
'---------------------------------------------------------------------------------------------------------------------------------------------
' Get Object ID
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetObjectID(ObjectName As String, ObjectType As Integer, Optional tblName As String = "$$OBJECTS") As Integer
    GetObjectID = Nz(DLookup("ID", SHT(tblName), "(ObjectName = " & _
                      sCH(ObjectName) & ") AND (ObjectType = " & ObjectType & ")"), -1) '!!!!!!!!!!!!!!!!!
End Function
'---------------------------------------------------------------------------------------------------------------------------------------------
' Get Fields for Query
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetQueryFlds(QryName As String) As dbField()
Dim FLDS() As dbField, nDim As Integer
Dim FLD As DAO.Field, qdf As DAO.QueryDef, db As DAO.Database


On Error GoTo ErrHandle
'----------------------------------------------
   nDim = -1: ReDim FLDS(0)
   Set db = CurrentDb(): Set qdf = db.QueryDefs(QryName)
   For Each FLD In qdf.FIELDS
       nDim = nDim + 1: ReDim Preserve FLDS(nDim)
       '---------------------------------------
       FLDS(nDim).Name = FLD.Name
       FLDS(nDim).fldSize = FLD.SIZE
       FLDS(nDim).FldType = FieldTypeName(FLD)
       FLDS(nDim).Description = GetAccessProp(FLD, "Description")
       'FLDS(nDim).Body = GetDataSample(TblName, FLDS(nDim).Name)
       FLDS(nDim).FldCategory = "Query Field"
   Next
'----------------------------------------------
ExitHere:
       GetQueryFlds = FLDS '!!!!!!!!!!
       Set db = Nothing
       Exit Function
'-------------------
ErrHandle:
       ErrPrint "GetTableFields", Err.Number, Err.Description
       Err.Clear: Resume ExitHere
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------
' Get Fields for Table
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetTableFields(tblName As String) As dbField()
Dim FLDS() As dbField, nDim As Integer
Dim FLD As DAO.Field, tdf As DAO.TableDef, db As DAO.Database


On Error GoTo ErrHandle
'----------------------------------------------
   nDim = -1: ReDim FLDS(0)
   Set db = CurrentDb(): Set tdf = db.TableDefs(tblName)
   For Each FLD In tdf.FIELDS
       nDim = nDim + 1: ReDim Preserve FLDS(nDim)
       '---------------------------------------
       FLDS(nDim).Name = FLD.Name
       FLDS(nDim).fldSize = FLD.SIZE
       FLDS(nDim).FldType = FieldTypeName(FLD)
       FLDS(nDim).Description = GetAccessProp(FLD, "Description")
       FLDS(nDim).BODY = GetDataSample(tblName, FLDS(nDim).Name)
       FLDS(nDim).FldCategory = "Table Field"
   Next
'----------------------------------------------
ExitHere:
       GetTableFields = FLDS '!!!!!!!!!!
       Set db = Nothing
       Exit Function
'-------------------
ErrHandle:
       ErrPrint "GetTableFields", Err.Number, Err.Description
       Err.Clear: Resume ExitHere
End Function
'-----------------------------------------------------------------------------------------------------------------------------
' Get Data Sample
'-----------------------------------------------------------------------------------------------------------------------------
Private Function GetDataSample(tblName As String, FldName As String, Optional sCriteria As String = "ID > 0") As String
On Error Resume Next
        GetDataSample = CStr(Nz(DLookup(FldName, SHT(tblName), sCriteria), "")) '!!!!!!!!!!!!!!!!!!!!
End Function
'-----------------------------------------------------------------------------------------------------------------------------
' Return field type names for access tables
'-----------------------------------------------------------------------------------------------------------------------------
Public Function FieldTypeName(FLD As DAO.Field) As String
    'Purpose: Converts the numeric results of DAO Field.Type to text.
    Dim strReturn As String    'Name to return

    Select Case CLng(FLD.Type) 'fld.Type is Integer, but constants are Long.
        Case dbBoolean: strReturn = "Yes/No"            ' 1
        Case dbByte: strReturn = "Byte"                 ' 2
        Case dbInteger: strReturn = "Integer"           ' 3
        Case dbLong                                     ' 4
            If (FLD.Attributes And dbAutoIncrField) = 0& Then
                strReturn = "Long Integer"
            Else
                strReturn = "AutoNumber"
            End If
        Case dbCurrency: strReturn = "Currency"         ' 5
        Case dbSingle: strReturn = "Single"             ' 6
        Case dbDouble: strReturn = "Double"             ' 7
        Case dbDate: strReturn = "Date/Time"            ' 8
        Case dbBinary: strReturn = "Binary"             ' 9 (no interface)
        Case dbText                                     '10
            If (FLD.Attributes And dbFixedField) = 0& Then
                strReturn = "Text"
            Else
                strReturn = "Text (fixed width)"        '(no interface)
            End If
        Case dbLongBinary: strReturn = "OLE Object"     '11
        Case dbMemo                                     '12
            If (FLD.Attributes And dbHyperlinkField) = 0& Then
                strReturn = "Memo"
            Else
                strReturn = "Hyperlink"
            End If
        Case dbGUID: strReturn = "GUID"                 '15

        'Attached tables only: cannot create these in JET.
        Case dbBigInt: strReturn = "Big Integer"        '16
        Case dbVarBinary: strReturn = "VarBinary"       '17
        Case dbChar: strReturn = "Char"                 '18
        Case dbNumeric: strReturn = "Numeric"           '19
        Case dbDecimal: strReturn = "Decimal"           '20
        Case dbFloat: strReturn = "Float"               '21
        Case dbTime: strReturn = "Time"                 '22
        Case dbTimeStamp: strReturn = "Time Stamp"      '23

        'Constants for complex types don't work prior to Access 2007 and later.
        Case 101&: strReturn = "Attachment"         'dbAttachment
        Case 102&: strReturn = "Complex Byte"       'dbComplexByte
        Case 103&: strReturn = "Complex Integer"    'dbComplexInteger
        Case 104&: strReturn = "Complex Long"       'dbComplexLong
        Case 105&: strReturn = "Complex Single"     'dbComplexSingle
        Case 106&: strReturn = "Complex Double"     'dbComplexDouble
        Case 107&: strReturn = "Complex GUID"       'dbComplexGUID
        Case 108&: strReturn = "Complex Decimal"    'dbComplexDecimal
        Case 109&: strReturn = "Complex Text"       'dbComplexText
        Case Else: strReturn = "Field type " & FLD.Type & " unknown"
    End Select

    FieldTypeName = strReturn
End Function
'---------------------------------------------------------------------------------------------------------------------------------------------
' Function put array of dbFLDS to Fld Tables (without their fields)
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub WriteFlds(FLDS() As dbField, iParent As Integer, Optional tblName As String = "$$FLDS")
Dim RS As DAO.Recordset, SQL As String, sHASH As String
Dim nDim As Integer, I As Integer, sWork As String

On Error GoTo ErrHandle
'------------------------------
SQL = "SELECT * FROM " & SHT(tblName) & ";": Set RS = CurrentDb.OpenRecordset(SQL)
nDim = UBound(FLDS)
With RS
    For I = 0 To nDim
           sHASH = GetHASH("FLD" & I)
           .AddNew
                !HASH = sHASH
                If FLDS(I).DateCreated > CDate(0) Then !DateCreate = FLDS(I).DateCreated
                sWork = FLDS(I).Name
                If sWork <> "" Then !FldName = sWork
                !FldParent = iParent
                If FLDS(I).FldType <> "" Then !FldType = FLDS(I).FldType
                If FLDS(I).FldCategory <> "" Then !FldCategory = FLDS(I).FldCategory
                '--------------------------------------------
                If FLDS(I).fldSize > 0 Then !fldSize = FLDS(I).fldSize
                If FLDS(I).Description <> "" Then !Description = FLDS(I).Description
                '--------------------------------------------
                If FLDS(I).BODY <> "" Then !BODY = FLDS(I).BODY
                If FLDS(I).META <> "" Then !META = FLDS(I).META
           .Update
    Next I
End With

'------------------------------
ExitHere:
    Set RS = Nothing
    Exit Sub
'------------
ErrHandle:
    ErrPrint "WriteFlds (" & sWork & ")", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Sub
'---------------------------------------------------------------------------------------------------------------------------------------------
' Function put array of dbObject to Object Tables (without their fields)
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Sub WriteObjects(OBJS() As dbObject, Optional tblName As String = "$$Objects")
Dim RS As DAO.Recordset, SQL As String, sHASH As String, sWork As String
Dim nDim As Integer, I As Integer

On Error GoTo ErrHandle
'------------------------------
SQL = "SELECT * FROM " & SHT(tblName) & ";": Set RS = CurrentDb.OpenRecordset(SQL)
nDim = UBound(OBJS)
With RS
    For I = 0 To nDim
           sHASH = GetHASH("OBJ" & I)
           .AddNew
                !HASH = sHASH
                If OBJS(I).DateCreated > CDate(0) Then !DateCreate = OBJS(I).DateCreated
                sWork = OBJS(I).Name
                If sWork <> "" Then !ObjectName = sWork
                If OBJS(I).Type <> 0 Then !ObjectType = OBJS(I).Type
                '--------------------------------------------
                !ObjectGroup = OBJS(I).Group
                If OBJS(I).DataType <> 0 Then !DataType = OBJS(I).DataType
                If OBJS(I).Parent <> "" Then !ObjectParent = OBJS(I).Parent
                If OBJS(I).Description <> "" Then !Description = OBJS(I).Description
                '--------------------------------------------
                If OBJS(I).BODY <> "" Then !BODY = OBJS(I).BODY
                If OBJS(I).META <> "" Then !META = OBJS(I).META
           .Update
    Next I
End With

'------------------------------
ExitHere:
    Set RS = Nothing
    Exit Sub
'------------
ErrHandle:
    ErrPrint "WriteObjects (" & sWork & ")", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Sub
'---------------------------------------------------------------------------------------------------------------------------------------------
' Function Recreate SPECIAL TABLES ($$OBJECTS and $$FLDS)
'---------------------------------------------------------------------------------------------------------------------------------------------
Private Function RecreateDocTables(Optional ObjTBL As String = "$$OBJECTS", Optional FldTBL As String = "$$FLDS", _
                                                                                           Optional bDelIfExist As Boolean = True) As Boolean
Dim bRes As Boolean, SQL As String
On Error GoTo ErrHandle
'-------------------------------------------
DoCmd.SetWarnings False
If bDelIfExist Then
    If IsEntityExist(ObjTBL) Then DeleteTable (ObjTBL)
    If IsEntityExist(FldTBL) Then DeleteTable (FldTBL)
End If
'-------------------------------------------
If Not IsEntityExist(ObjTBL) Then
     
    SQL = "CREATE TABLE " & SHT(ObjTBL) _
        & "([ID] AUTOINCREMENT PRIMARY KEY, [DateCreate] DATETIME, [DateUpdate] DATETIME, [IsArchive] YESNO,[HASH] CHAR(250), " _
        & "[ObjectName] CHAR, [ObjectType] INTEGER, [ObjectGroup] INTEGER,[ObjectParent] TEXT(250), [DataType] INTEGER, " _
        & "[Description] MEMO,[BODY] MEMO,[META] CHAR);"
    CurrentDb.Execute SQL
    '----------------------------------------
    CurrentDb.TableDefs.Refresh
    SetDefaultValueForField ObjTBL, "DateCreate", "= Now()"
    SetDefaultValueForField ObjTBL, "DateUpdate", "= Now()"
    
    Call SetLookUpFLD(ObjTBL, "ObjectType", "1;Table;5;Query;-32768;Form;-32764;Report;-32766;Macro;-32761;VBA", , "Value List")
    SetDefaultValueForField ObjTBL, "ObjectType", 1
    Call SetLookUpFLD(ObjTBL, "ObjectGroup", "-3;SYS0;-2;SYS1;-1;SYS2;0;Undefinite;1;Entity;2;Reference;" & _
                     "3;Crosslinks;4;Query Extension;5;Window Form;6;Listing Form;7;Detailed Form;8;SubForm;" & _
                     "9;VBA Module;10;VBA Class;11;AccessObj;12;Other", , "Value List")
    SetDefaultValueForField ObjTBL, "ObjectGroup", 0
End If
If Not IsEntityExist(FldTBL) Then
     
    SQL = "CREATE TABLE " & SHT(FldTBL) _
        & "([ID] AUTOINCREMENT PRIMARY KEY, [DateCreate] DATETIME, [DateUpdate] DATETIME, [IsArchive] YESNO,[HASH] CHAR(250), " _
        & "[FldName] CHAR, [FldType] CHAR, [FldSize] INTEGER,[FldParent] LONG, [FldCategory] CHAR, " _
        & "[Description] MEMO,[BODY] MEMO,[META] MEMO);"
    CurrentDb.Execute SQL
    '----------------------------------------
    CurrentDb.TableDefs.Refresh
    SetDefaultValueForField FldTBL, "DateCreate", "= Now()"
    SetDefaultValueForField FldTBL, "DateUpdate", "= Now()"
    
    Call SetLookUpFLD(FldTBL, "FldParent", "SELECT ID, ObjectName FROM " & SHT(ObjTBL))
    
End If

bRes = True
'--------------------
ExitHere:
     RecreateDocTables = bRes '!!!!!!!!!!!!!!!
     DoCmd.SetWarnings True
     Exit Function
'-----------
ErrHandle:
     ErrPrint "RecreateDocTables", Err.Number, Err.Description
     Err.Clear: Resume ExitHere
End Function

'--------------------------------------------------------------------------------------------------------------------------------
' Delete Table
'--------------------------------------------------------------------------------------------------------------------------------
Private Sub DeleteTable(TableName As String, Optional bWarnings As Boolean = True)

On Error GoTo ErrHandle
'---------------------------------
If bWarnings Then DoCmd.SetWarnings False

      CurrentDb.Execute "DROP TABLE " & SHT(TableName) & ";"
'---------------------------------
ExitHere:
      If bWarnings Then DoCmd.SetWarnings True
      Exit Sub
'--------
ErrHandle:
      ErrPrint "DeleteTable", Err.Number, Err.Description
      Err.Clear: Resume ExitHere
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------------
' Function Get Field Description from $$FLDS and reup it for database fld only
'------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ReUPFldDescript(Optional tblName As String = "$$FLDS", Optional ObjTBL As String = "$$OBJECTS")
Dim STABLE As String, sFLD As String, sDescript As String, nType As Integer, iParent As Integer
Dim RS As DAO.Recordset, SQL As String

On Error GoTo ErrHandle
'----------------------------
SQL = "SELECT FldName, FldParent, Description FROM " & SHT(tblName) & ";"
Set RS = CurrentDb.OpenRecordset(SQL)

With RS
    If Not .EOF Then
        .MoveLast: .MoveFirst
        Do While Not .EOF
            sFLD = Trim(Nz(!FldName, "")): iParent = Nz(!FldParent, -1): sDescript = Trim(Nz(!Description, ""))
            If iParent = -1 Then GoTo NextStep
            If Nz(DLookup("ObjectType", SHT(ObjTBL), "ID = " & iParent), -1) <> 1 Then GoTo NextStep
            STABLE = Trim(Nz(DLookup("ObjectName", SHT(ObjTBL), "ID = " & iParent), ""))
            Call SetFieldDescription(STABLE, sFLD, sDescript)
NextStep:
            .MoveNext
        Loop
    End If
End With
'------------------------------
ExitHere:
    Set RS = Nothing
    Exit Sub
'-----
ErrHandle:
    ErrPrint "ReUPFldDescript", Err.Number, Err.Description
    Err.Clear
End Sub
'======================================================================================================================================================
' Set Description for field
'======================================================================================================================================================
Public Sub SetFieldDescription(tblName As String, FldName As String, Optional sDescription As String)
Dim FLD As DAO.Field, tdf As DAO.TableDef, db As DAO.Database

Dim sText As String

On Error GoTo ErrHandle
'----------------------------------------------
Select Case UCase(FldName)
Case "ID":
    sText = "Unique identifier and auto-generated primary key for table record. Can be Long autonumber or GUID"
Case "HASH":
    sText = "Generated unique string for replication proposes, store some information about environment and date of the record  creation"
Case "ISARCHIVE":
    sText = "Archive flag, when is true it means that this record is not actual now. By default is false"
Case "CREATEDATE":
    sText = "Date of the record creation. By default is current date"
Case "UPDATEDATE":
    sText = "Date of the record update. By default is current date"
Case "UPDATEBY":
    sText = "Username who update this record."
Case "TITLE":
    sText = "Entity title or name"
Case "DESCRIPTION":
    sText = "Some long text descriptes this record's entity"
Case "ATTACHMENTS":
    sText = "The specific Access solution to store multiple binary object in database"
Case Else
    sText = sDescription
End Select
'----------------------------------------------
If sText = "" Then Exit Sub

   Set db = CurrentDb(): Set tdf = db.TableDefs(tblName)
   Set FLD = tdf.FIELDS(FldName)
   Call SetAccessProp(FLD, "Description", sText)
'---------------------------------------------
ExitHere:
    Set db = Nothing
    Exit Sub
'------------
ErrHandle:
    ErrPrint "SetFieldDescription", Err.Number, Err.Description
    Err.Clear
End Sub
'----------------------------------------------------------------------------------------------------------------------------------------------
' Function ReUP Description from Database Objects to objects
'----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ReUpDescription(Optional ObjTBL As String = "$$OBJECTS", Optional FldTBL As String = "$$FLDS")
Dim RS As DAO.Recordset, sDescript As String, sName As String, iType As Integer
Dim sSQL As String, TBL As String

TBL = SHT(ObjTBL)
    sSQL = "SELECT " & TBL & ".ObjectName," & TBL & ".ObjectType," & TBL & ".Description FROM " & TBL & _
           " WHERE (((" & TBL & ".IsArchive)=False));"
'---------------------------------------------------------------------------------------------------------
Set RS = CurrentDb.OpenRecordset(sSQL)
    With RS
         If Not .EOF Then
            .MoveLast: .MoveFirst
            '---------------------------------------------------------------------------
            Do While Not .EOF
               sName = Trim(Nz(RS![ObjectName], "")): iType = Nz(RS![ObjectType], 1)
               sDescript = Trim(Nz(RS![Description], ""))
               If sDescript <> "" Then
                  If IsEntityExist(sName, iType) Then
                      '-----------------------------------------------------------
                      Select Case iType
                            Case 1:           ' TABLE
                                    Call SetDescrTable(sName, sDescript)
                            Case 5:           ' QUERY
                                    Call SetAccessProp(CurrentDb.QueryDefs(sName), "Description", sDescript)
                            Case -32768:      ' FORMS
                                    Call SetAccessProp(CurrentProject.AllForms(sName), "Description", sDescript)
                            Case -32764:      ' REPORTS
                                    Call SetAccessProp(CurrentProject.AllReports(sName), "Description", sDescript)
                            Case -32766:      ' MACRO
                                    Call SetAccessProp(CurrentProject.AllMacros(sName), "Description", sDescript)
                            Case -32761:      ' MODULE
                                    Call SetAccessProp(CurrentProject.AllModules(sName), "Description", sDescript)
                            Case Else
                                    GoTo LoopNext
                      End Select
                      '-----------------------------------------------------------
                  End If
               End If
LoopNext:
               .MoveNext
            Loop
            '--------------------------------------------------------------------------
         End If
    End With
'---------------------------------------------------------------------------------------------------------
    Call ReUPFldDescript(FldTBL, ObjTBL)  ' SAVE FIELDS
'---------------------------------------------------------------------------------------------------------
ExitHere:
    Set RS = Nothing
    Exit Sub
'------------------------
ErrHandle:
    ErrPrint "ReUpDescription", Err.Number, Err.Description
    Err.Clear:  Resume ExitHere
End Sub

Public Sub TetDescr()
    Call SetDescrTable("$$CACHE", "Store operational information")
End Sub
'-----------------------------------------------------------------------------------------------------------------------------------------------------
' Set Description for Table
'-----------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub SetDescrTable(tblName As String, sDescr As String)
On Error Resume Next
    Call SetAccessProp(CurrentDb.TableDefs(SHT(tblName)), "Description", sDescr)
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------
' Get Copyright Label
'-----------------------------------------------------------------------------------------------------------------------------------------------------
Private Function Copyright()
Dim sCopyright

On Error Resume Next
'--------------------------
      sCopyright = Chr(39) & vbTab & "@ Valery Khvatov (valery.khvatov@gmail.com), [01/" & Format(Now(), "yyyymmdd") & "]"
      ToClipBoard sCopyright
'--------------------------
ExitHere:
      Copyright = sCopyright '!!!!!!!!!!!!!!!!!!!!
End Function
'======================================================================================================================================================
' Get Divider
'======================================================================================================================================================
Public Function Divider(Optional sMark As String = "=", Optional LenDiv As Integer = 150) As String
Dim DLM As String

'--------------------------------------------
    If sMark = "" Then
        DLM = InputBox("Please set divider mark: ", "Divider", "=")
    Else
        DLM = sMark
    End If
'---------------------------------------------
    Divider = Chr(39) & String(LenDiv, DLM)  '!!!!!!!!!!!!!!
    ToClipBoard Divider
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------
' Get ErrHandle and ExitHere
'-----------------------------------------------------------------------------------------------------------------------------------------------------
Private Function FuncTail(Optional FuncName As String, Optional bFunc As Boolean = True, Optional bErrPrint2 As Boolean) As String
Dim sFunc As String, sRes As String

If FuncName = "" Then
    sFunc = InputBox("Please set Func Name:", "FuncTail", "NewFunction")
Else
    sFunc = FuncName
End If
'-----------------------------------
sRes = Chr(39) & String(30, "-")
sRes = sRes & vbCrLf & "ExitHere:"
sRes = sRes & vbCrLf & vbTab & IIf(bFunc, sFunc & " = sRes " & vbTab & Chr(39) & String(20, "!") & vbCrLf & vbTab & "Exit Function", "Exit Sub")
sRes = sRes & vbCrLf & Chr(39) & String(15, "-")
sRes = sRes & vbCrLf & "ErrHandle:"
sRes = sRes & vbCrLf & vbTab & IIf(bErrPrint2, "ErrPrint2 ", "ErrPrint ") & Chr(34) & sFunc & Chr(34) & ", Err.Number, Err.Description" & IIf(bErrPrint2, ", MOD_NAME", "")
sRes = sRes & vbCrLf & vbTab & "Err.Clear:    Resume ExitHere"
'-----------------------------------
ExitHere:
      FuncTail = sRes '!!!!!!!!!!!!!!!!!
      ToClipBoard sRes
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------
' Func Header
'-----------------------------------------------------------------------------------------------------------------------------------------------------
Public Function FuncHeader(Optional PublicPrivate As String = "Public", Optional sDescription As String, Optional LenDiv As Integer = 150) As String
Dim sRes As String, sSep As String, sDescript As String
On Error Resume Next
'------------------------------------
    sSep = IIf(PublicPrivate = "Public", "=", "-")
    sDescript = FormatComments(IIf(sDescription <> "", sDescription, " Description: " & " This Function was generated in " & Now() & " ..."))
    sRes = Divider(sSep, LenDiv) & vbCrLf & Chr(39) & vbTab & sDescript
    sRes = sRes & vbCrLf & Copyright & vbCrLf & Divider(sSep, LenDiv)
'--------------------------------------
ExitHere:
    FuncHeader = sRes '!!!!!!!!!!!!!!!!!!
    ToClipBoard sRes
End Function

'=====================================================================================================================================================
' VBA Function Template Generator
'=====================================================================================================================================================
Public Function PrintFunc(Optional bHeaderOnly As Boolean, Optional FuncSub As String = "Function", Optional sARG As String, _
                Optional PublicPrivate As String = "Public", Optional FuncName As String, Optional FuncType As String = "String", _
                        Optional FuncDescription As String, Optional AddErrHandler As Boolean = True, Optional bErrPrint2 As Boolean = True, _
                                                                                                           Optional LenDiv As Integer = 150) As String
Dim HDR(6) As String, sText As String, sFunct As String

On Error GoTo ErrHandle
'---------------------------------------
HDR(0) = FuncHeader(PublicPrivate, FuncDescription, LenDiv)

If FuncName = "" Then sFunct = InputBox("Enter Function Name", "PrintFunc", "SomeFunction")
sFunct = IIf(sFunct <> "", sFunct, "FUNC" & Format(Now(), "yyyymmddhhnnss"))

HDR(1) = IIf(PublicPrivate = "Public", "Public", "Private") & " " & IIf(FuncSub = "Function", "Function", "Sub") & " " & sFunct & _
         IIf(sARG <> "", "(" & sARG & ")", "() ") & IIf(FuncSub = "Function", " As " & FuncType, "")

HDR(2) = "Dim sRes as String"
HDR(3) = vbCrLf & "On Error GoTo ErrHandle" & vbCrLf & Chr(39) & String(30, "-")
HDR(4) = vbCrLf & vbCrLf & "sRes = " & Chr(34) & "SOMETHING" & Chr(34) & vbCrLf

HDR(5) = FuncTail(sFunct, FuncSub = "Function", bErrPrint2)
HDR(6) = IIf(FuncSub = "Function", "End Function", "End Sub")

sText = Join(HDR, vbCrLf)
'-----------------------------------
ExitHere:
    PrintFunc = sText '!!!!!!!!!!!!!!!!!!!!
    ToClipBoard sText
    Debug.Print "The Function " & sFunct; " is placed to clipboard"
    Exit Function
'---------------------------------------
ErrHandle:
    ErrPrint "PrintFunc", Err.Number, Err.Description
    Err.Clear
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------
' Module Header
'-----------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ModHeader(Optional ModName As String, Optional sDescription As String, Optional LenDiv As Integer = 150) As String
Dim sRes As String, sSep As String, sDescript As String
On Error Resume Next
'------------------------------------
    sSep = "*"
    sDescript = FormatComments(IIf(sDescription <> "", sDescription, " Description: " & " This Module was generated in " & Now() & " ..."))
    
    sRes = Divider(sSep, LenDiv) & vbCrLf & Divider(sSep, LenDiv) & vbCrLf & sDescript
    
    sRes = sRes & vbCrLf & Copyright & vbCrLf & Divider(sSep) & vbCrLf & Divider(sSep)
'--------------------------------------
ExitHere:
    ModHeader = sRes '!!!!!!!!!!!!!!!!!!
    ToClipBoard sRes
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------
' Generate ERRPRINT FUNCTION
'-----------------------------------------------------------------------------------------------------------------------------------------------------
Public Function PrintErrPrint(Optional ModName As String) As String

Dim sRes As String, sSep As String
Dim HDR(6) As String, sModName As String

On Error Resume Next
'------------------------------------
    sModName = IIf(ModName <> "", ModName, "Module" & Format(Now(), "yyyymmddhhnnss"))
    sSep = "/"
    
    HDR(0) = Divider(sSep) & vbCrLf & Divider(sSep) & vbCrLf & Divider(sSep) & vbCrLf & Divider(sSep) & vbCrLf & Divider(sSep)
    sSep = "-"
    HDR(1) = Divider(sSep) & vbCrLf & Chr(39) & vbTab & "ERROR HANDLER FOR " & sModName & vbCrLf & Divider(sSep)
    
    HDR(2) = "Private Function ErrPrint(FuncName As String, ErrNumber As Long, ErrDescription As String, Optional bDebug As Boolean = True, _" & vbCrLf & _
             "                                                                                   Optional sModName As String = " & _
             Chr(34) & sModName & Chr(34) & " As String"
    
    HDR(3) = "Dim sREs As String" & vbCrLf & "Const ErrChar As String = " & Chr(34) & "#" & Chr(34) & vbCrLf & _
             "Const ErrRepeat As Integer = 60" & vbCrLf & vbCrLf
    
    HDR(4) = "sREs = String(ErrRepeat, ErrChar) & vbCrLf & " & QR("ERROR OF [") & " & sModName & " & QR(": ") & " & FuncName & " & QR("]") & _
             " & vbTab & " & QR("ERR#") & " & ErrNumber & vbTab & Now() & _" & vbCrLf & _
             "                                                          vbCrLf & ErrDescription & vbCrLf & String(ErrRepeat, ErrChar)"
    
    HDR(5) = vbCrLf & vbCrLf & "If bDebug Then Debug.Print sREs" & vbCrLf & String(50, "-")
    HDR(6) = "ExitHere:" & vbCrLf & vbTab & "Beep" & vbCrLf & vbTab & "ErrPrint = sREs" & "   " & Chr(39) & String(30, "!") & vbCrLf & "End Function"
    
    sRes = Join(HDR, vbCrLf)
'--------------------------------------
ExitHere:
    PrintErrPrint = sRes '!!!!!!!!!!!!!!!!!!
    ToClipBoard sRes
End Function
'======================================================================================================================================================
' Print New Module
'======================================================================================================================================================
Public Function PrintMod(Optional ModName As String, Optional sDescription As String, Optional LenDiv As Integer = 150) As String
Dim sMod As String
Dim HDR(6) As String, sRes As String
On Error Resume Next
'---------------------------------------
     sMod = IIf(ModName <> "", ModName, "Module" & Format(Now(), "yyyymmddhhnnss"))
     
     HDR(0) = ModHeader(sMod, sDescription)
     HDR(1) = "Option Compare Database" & vbCrLf & "Option Explicit" & vbCrLf & Chr(39) & String(60, "*")
     HDR(2) = vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
     HDR(3) = PrintErrPrint(sMod)
'----------------------------------------
ExitHere:
     sRes = Join(HDR, vbCrLf)
     ToClipBoard sRes
     Debug.Print "The Module " & sMod & " template is placed in clipboard"
     PrintMod = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function

'======================================================================================================================================================
' Function  : SetClassAttributeToMultiUse
' DateTime  : 16-10-2005 13:49
' Author    : d
' Purpose   : Ensures the instancing property of all class modules is set to 5
'             (multi-use) so that they can be instantiated from a linked db.
' Notes     : VBE.ActiveVBProject.VBComponents.Item(x).Type
'             (Microsoft Visual Basic For Applications Extensibilty 5.3 -  VBIDE)
'               1   - Standard Module   (vbext_ct_StdModule)
'               2   - Class Module      (vbext_ct_ClassModule)
'               100 - Access Form       (vbext_ct_Document)
'             VBE.ActiveVBProject.VBComponents.Item(x).Properties("Instancing").Value
'               1   - Private
'               2   - PublicNotCreatable
'               5   - GlobalMultiUse
'======================================================================================================================================================
Public Function SetClassAttributeToMultiUse() As Boolean
    Dim blRet As Boolean, I As Integer

On Error GoTo ErrHandle
    
    With VBE.ActiveVBProject.VBComponents
        For I = 1 To .Count
            If .Item(I).Type = 2 Then               ' Class module (vbext_ct_ClassModule)
                With .Item(I).Properties.Item(2)    ' "Instancing"
                    If .value <> 5 Then
                           .value = 5
                    End If
                End With
            End If
        Next I
    End With
    blRet = True
'------------------------------------
ExitHere:
    SetClassAttributeToMultiUse = blRet
    Debug.Print "The attributes for all classes set to multiuse"
    Exit Function
'------------
ErrHandle:
    ErrPrint "SetClassAttributeToMultiUse", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function

'======================================================================================================================================================
' Copy Procedure from clipboard to module in Gracle
'======================================================================================================================================================
Public Function ImportCodeFromClipBoard(Optional strCode As String, Optional ProjectName As String = "_GRACKLE", _
                                                       Optional ModuleName As String = "Module1", Optional bFromClipBoard As Boolean = True) As Boolean
Dim iRow As Integer, sCode As String

On Error GoTo ErrHandle
'--------------------------------
     sCode = strCode: If sCode = "" And bFromClipBoard Then sCode = FromClipboard()
     If sCode = "" Then Exit Function
     
    With Application.VBE.VBProjects(ProjectName).VBComponents.Item(ModuleName).CodeModule
            iRow = .CountOfLines + 1
            .InsertLines iRow, sCode
    End With
'----------------------------------
ExitHere:
    ImportCodeFromClipBoard = True '!!!!!!!!!!
    Exit Function
'----------------
ErrHandle:
    ErrPrint2 "ImportCodeFromClipBoard", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function


'======================================================================================================================================================
' Calculate Total  Code Lines in the project
'======================================================================================================================================================
Public Function TotalCodeLines(Optional sProjectName As String) As Long
Dim VBProj As Object, VBComp As Object, LineCount As Long

    On Error GoTo ErrHandle
'---------------------------
LineCount = -1
If sProjectName = "" Then
    Set VBProj = Application.VBE.ActiveVBProject
Else
    Set VBProj = Application.VBE.VBProjects(sProjectName)
End If

        
        For Each VBComp In VBProj.VBComponents
            LineCount = LineCount + VBComp.CodeModule.CountOfLines
        Next VBComp
'---------------------------
ExitHere:
    TotalCodeLines = LineCount  '!!!!!!!!!!!!
    Set VBComp = Nothing: Set VBProj = Nothing
    Exit Function
'--------------
ErrHandle:
    ErrPrint2 "TotalCodeLines", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Resume ExitHere
End Function



