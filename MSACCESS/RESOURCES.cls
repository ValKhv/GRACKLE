VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RESOURCES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
' Static class for managing library resources
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
Option Explicit


Private Const MOD_NAME As String = "RESOURCES"
Private Const MOD_VERSION As String = "20220331"
'****************************************

'======================================================================================================================================================
' Standard About Method
'======================================================================================================================================================
Public Function About() As String
    About = MOD_VERSION  '!!!!!!!!!!!!!!!!!!!
End Function

'======================================================================================================================================================
' A resource loader for icon that can take a path to a file on disk or a reference to a resource stored in a private resource
' type of resources (loading using the RESOURCES component). Return Resource Pointer (Long or LongPtr)
'======================================================================================================================================================
Public Function LoadIcon(ByVal IconLocator As Variant, Optional ByVal iconHeight As Long, Optional ByVal iconWidth As Long, _
                                                                                                           Optional ByVal Addin As Boolean) As Variant
Dim vRes As Variant, sRL As String

Const ERR_NO_RESOURCE_LOCATOR As Long = 10067

    On Error GoTo ErrHandle
'--------------------------
    vRes = -1
    If varType(IconLocator) = vbString Then
         sRL = CStr(IconLocator): If IsBlank(sRL) Then GoTo ExitHere
         If InStr(1, sRL, "\") > 0 Then                         ' This is a path to external file
             If InStr(1, sRL, ":") = 0 Then   ' This Is a Realative Path
                 sRL = CurrentProject.Path & sRL
             End If
             If IsBlank(Dir(sRL)) Then Err.Raise ERR_NO_RESOURCE_LOCATOR, , "Can't find the file " & sRL
             
             vRes = LoadPicturehIcon(sRL, iconHeight, iconWidth)
             
         Else                                                   ' This is a Name in Resource Table
             If RESOURCES.GetResorceIDByName(sRL) < 0 Then Err.Raise ERR_NO_RESOURCE_LOCATOR, , "Can't find the file " & sRL
             vRes = AttachmentTohIcon("MSysResources", "Data", "Name", dbText, sRL, iconHeight, iconWidth)
         End If
    End If
'--------------------------
ExitHere:
    LoadIcon = vRes  '!!!!!!!!!!!!!!
    Exit Function
'-----------
ErrHandle:
    ErrPrint2 "LoadIcon", Err.Number, Err.Description, "MOD_NAME"
    Err.Clear: vRes = -1: Resume ExitHere
End Function
'=================================================================================================================================================
' Get Resource Table List
'=================================================================================================================================================
Public Function GetResourceTableList(Optional ResName As String, Optional TBL As String = "MSysResources", Optional bWithHeader As Boolean = True, _
                                   Optional bForScreen As Boolean = True, Optional DLM As String = ";", Optional SEP As String = vbCrLf) As String

Dim sRes As String, SQL As String
       
    On Error GoTo ErrHandle
'--------------------
    SQL = "SELECT * FROM " & TBL & IIf(IsBlank(ResName), ";", " WHERE (Name = " & sCH(ResName) & ");")
    sRes = PrintRows(SQL, , bWithHeader, bForScreen, , DLM, SEP)
'--------------------
ExitHere:
    GetResourceTableList = sRes '!!!!!!!!!!!!!!!
    Exit Function
'---------
ErrHandle:
    ErrPrint2 "GetResourceTableList", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Resume ExitHere
End Function

'======================================================================================================================================================
' The function loads resources inside the system resource table. Return ID of loaded resource
'======================================================================================================================================================
Public Function ResourceTableLoad(Optional sFiles As String, Optional EntryName As String, Optional TBL As String = "MSysResources", _
                                                                                                             Optional DLM As String = ";") As Long
Dim nRes As Long, sFLS As String, FLS() As String, nFLS As Integer, sName As String, sExtention As String, sType As String
Dim IDD As Long, RS As DAO.Recordset

Const ATTACH_FLD As String = "Data"

    On Error GoTo ErrHandle
'--------------------
    nRes = -1: sFLS = sFiles: If IsBlank(sFLS) Then sFLS = OpenDialog(GC_FILE_PICKER, "Please select file(s)", "All Files,*.*", , GetLastFolder)
    If IsBlank(sFLS) Then Exit Function
    FLS = Split(sFLS, DLM): nFLS = UBound(FLS)
    sName = EntryName: If IsBlank(sName) Then sName = FileNameOnly(FLS(0))
    If sName = "" Then Exit Function
    If InStr(1, sName, ".") > 0 Then sName = Split(sName, ".")(0)
    sExtention = FileExt(FLS(0)): sType = GetFileType(sExtention)
    
    IDD = GetResorceIDByName(sName)
    If IDD < 0 Then
         Set RS = CurrentDb.OpenRecordset(TBL)
         With RS
             .AddNew
                   .FIELDS("Name").value = sName
                   If Not IsBlank(sType) Then .FIELDS("Type").value = sType
                   If Not IsBlank(sExtention) Then .FIELDS("Extension").value = sExtention
             .Update
         End With
         IDD = GetResorceIDByName(sName)
    End If
    
    If IDD > 0 Then
       sFLS = AttchmentLoad(TBL, ATTACH_FLD, "ID = " & IDD, sFLS, DLM)
       If Not IsBlank(sFLS) Then nRes = IDD
    End If
    
'--------------------
ExitHere:
    ResourceTableLoad = nRes '!!!!!!!!!!!!!!!
    Set RS = Nothing
    Exit Function
'---------
ErrHandle:
    Debug.Print "ResourceTableLoad", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Resume ExitHere
End Function


'=================================================================================================================================================
' Get Resource ID By Name
'=================================================================================================================================================
Public Function GetResorceIDByName(sName As String, Optional FLD As String = "Name", Optional TBL As String = "MSysResources") As Long
        GetResorceIDByName = Nz(DLookup("Id", TBL, FLD & " = " & sCH(sName)), -1)        '!!!!!!!!!!!!!!
End Function


'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Get File Type by Extension
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetFileType(sExt As String) As String
Dim sRes As String
    
    On Error Resume Next
'------------------------
    Select Case sExt
    Case "jpeg", "jpg", "png", "bmp", "tiff", "gif":
          sRes = "img"
    Case "mp3", "wav":
          sRes = "sound"
    Case "doc", "docx", "xls", "xlsx", "ppt", "pptx":
          sRes = "document"
    Case "txt":
          sRes = "text"
    Case "pdf":
          sRes = "pdf"
    Case Else:
          sRes = "unknow"
    End Select
'------------------------
ExitHere:
    GetFileType = sRes '!!!!!!!!!
End Function

