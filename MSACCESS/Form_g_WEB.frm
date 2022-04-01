VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_g_WEB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
' This the web form, part of GFORMS Library
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
Option Explicit

Private Const DIALOGWEB As String = "DIALOGWEB"
'**********************

Private Sub btOK_Click()
     TempVars(DIALOGWEB).value = "True"
     
     DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_Load()
Dim sHome As String, sRef As String
Dim Arr() As String, iZoom As Long

Const DLM As String = ";"

   On Error Resume Next
'--------------------------------
    sRef = Nz(Me.OpenArgs, "")
    If sRef <> "" Then
        Arr = Split(sRef, DLM)
        sHome = Arr(0)
        Me.Caption = Arr(1)
        Me.lblPrompt.Caption = Arr(2)
        iZoom = CLng(Arr(3))
    Else
        sHome = "https://www.google.com/"
    End If

    sHome = "C:\Users\valer\Google Drive\_ZWORKS\_DATABASES\VBALIB\1.html"
    HomePage sHome
    If iZoom > 0 Then Call Zoom(iZoom)
'--------------------------------
ExitHere:
    Exit Sub
'--------
ErrHandle:
    ErrPrint2 "Form_Load", Err.Number, Err.Description, Me.Name
    Err.Clear
End Sub

'------------------------------------------------------------------------------------------------------------------------------------------------------
' ZOOM CONTENT
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Zoom(Optional iZoom As Long = 200)

Const OLECMDID_OPTICAL_ZOOM As Integer = 63
Const OLECMDEXECOPT_DONTPROMPTUSER As Integer = 2

    On Error GoTo ErrHandle
'------------------------------'
DoEvents

Me.BROWSER.Object.ExecWB OLECMDID_OPTICAL_ZOOM, _
OLECMDEXECOPT_DONTPROMPTUSER, _
CLng(iZoom), vbNull


'-----------------------
ExitHere:
    Exit Sub
'---------
ErrHandle:
    ErrPrint2 "Zoom", Err.Number, Err.Description
    Err.Clear
End Sub

'------------------------------------------------------------------------------------------------------------------------------------------------------
' RELOAD CONTENT
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function WriteContent(Optional sContent As String = "<html><head></head><body><p>Some content.</p></body></html>") As Boolean

    On Error GoTo ErrHandle
'--------------------
With Me.BROWSER.Object.Document
     .Open
        .Write sContent
     .Close
End With
'--------------------
ExitHere:
    WriteContent = True '!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'------
ErrHandle:
    ErrPrint2 "WriteContent", Err.Number, Err.Description, Me.Name
    Err.Clear
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------
' OPEN HOME PAGE
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function HomePage(sPage As String) As Boolean
Dim sURL As String
    
    On Error GoTo ErrHandle
'----------------------
    If Mid(sPage, 2, 1) = ":" Then ' This a local file
         sURL = FileRecode(sPage)
    Else
         sURL = sPage
    End If
    
    'Me.BROWSER.ControlSource = sURL
    Me.BROWSER.ControlSource = "=" & Chr(34) & sURL & Chr(34)
'----------------------
ExitHere:
    HomePage = True '!!!!!!!!!!!!
    Exit Function
'-----
ErrHandle:
    ErrPrint2 "HomePage", Err.Number, Err.Description, Me.Name
    Err.Clear
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------
' RE-CODE FILE PATH
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function FileRecode(sFile As String) As String
Dim sRes As String

    On Error Resume Next
'---------------------
sRes = Replace(sFile, ":", "$"): sRes = Replace(sRes, "\", "/")
sRes = "file://127.0.0.1/" & sRes
'---------------------
ExitHere:
    FileRecode = sRes '!!!!!!!!!!!!!
End Function

Private Sub lblPrompt_DblClick(Cancel As Integer)
Dim sFile As String
    sFile = OpenDialog(GC_FILE_PICKER, "Pick the File", "All Files,*.*", False, CurrentProject.Path)
    If sFile <> "" Then
        sFile = FileRecode(sFile)
        Me.BROWSER.ControlSource = "=" & Chr(34) & sFile & Chr(34)
    End If
End Sub
