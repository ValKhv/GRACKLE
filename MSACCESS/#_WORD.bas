Attribute VB_Name = "#_WORD"
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
'                 $$$$$$$$$$F                                       ## ##  ##  ##  ##    ##   ###    ##    ###  ##    ######  ######
'                  *$$$$$$$$"                                        ####  #####  ##     ##    ##    ##   ##   ## ##  ##  ##  ##  ##
'                    "***""               _____________                                         ###  ##  ##    ## ##  ####    ## ##
' STANDARD MODULE WITH DEFAULT FUNCTIONS |v 2020/12/01 |                                          ## ## ##     ## ##  ## ##   ## ##
' The module contains some functions to work with Word and is part of the G-VBA library            ######       ##    ##  ##  ####
'****************************************************************************************************************************************************
'****************************************************************************************************************************************************
Option Explicit

Private Const MOD_NAME As String = "#_WORD"


Public Sub EXTRACTCOMMENTS_TEST()
Dim sPath As String, sRes As String

sPath = OpenDialog(GC_FILE_PICKER, "Pick a file", , False, CurrentProject.Path)

sRes = EXTRACTCOMMENTS(sPath)

End Sub

'======================================================================================================================================================
' Change Commentator Name
'======================================================================================================================================================
Public Sub ChangeAllAuthorNamesInComments(Optional sFile As String, Optional sAuthor1 As String, Optional sAuthor2 As String, _
                                                                                                                       Optional bFromComment As Boolean)
Dim wdApp As Object, wdDoc As Object, sWordDocName As String
Dim sPath As String, sAuth1 As String, sAuth2 As String, sInitial As String, sLastName As String
Dim objComment As Object
    On Error GoTo ErrHandle
'----------------------------------------------
    If sFile = "" Then
       sPath = OpenDialog(GC_FILE_PICKER, "Open Word Document", "Word Documents,*.docx", False, GetLastFolder())
    Else
       sPath = sFile
    End If
    
    If sPath = "" Then Exit Sub
    If Dir(sPath) = "" Then Err.Raise 10003, , "Can't find the file: " & sPath
    
    If sAuthor1 <> "" Then
       sAuth1 = sAuthor1
    Else
       sAuth1 = InputBox("Enter the original author you want to find", "Setup Author 1")
       If sAuth1 = "" And Not bFromComment Then Exit Sub
    End If
        
    If sAuthor2 <> "" Then
       sAuth2 = sAuthor2
    Else
       sAuth2 = InputBox("Please add the author to replace the previous one", "Setup Author 2")
       If sAuth2 = "" Then Exit Sub
    End If
'----------------------------------------------
    sInitial = GetInitials(sAuth2): If bFromComment Then sLastName = GetLastName(sAuth2)
    
    Set wdApp = GetWordApp()
    If wdApp Is Nothing Then Err.Raise 1000, , "Can't start Word App Object"
    
    sWordDocName = FileNameOnly(sPath):
    Set wdDoc = IsWordDoc(wdApp, sWordDocName)
    
    If wdDoc Is Nothing Then Set wdDoc = wdApp.Documents.Open(sPath)

    If wdDoc Is Nothing Then Err.Raise 1000, , "Can't open word Document: " & sPath
    wdDoc.Activate

    
  ' Change all author names in comments
  For Each objComment In wdDoc.Comments
    If Not bFromComment Then
        If objComment.Author = sAuth1 Then
            objComment.Author = sAuth2
            objComment.Initial = sInitial
        End If
    Else
        If Left(objComment.Range, Len(sLastName)) = sLastName Then
            objComment.Author = sAuth2
            objComment.Initial = sInitial
        End If
    End If
    
  Next objComment
'---------------------------------
ExitHere:
    Exit Sub
'----------------
ErrHandle:
    ErrPrint2 "ChangeAllAuthorNamesInComments", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Get LastName
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetLastName(sName As String) As String
Dim sRes As String
    If InStr(1, sName, " ") Then
        sRes = Split(sName, " ")(0)
    Else
        sRes = sName
    End If
'-----------------
ExitHere:
    GetLastName = sRes '!!!!!!!!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Get INITIALS
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetInitials(sName As String) As String
Dim sRes As String, sASS() As String, sWork As String
    sWork = Trim(sName): sWork = Replace(sWork, "  ", " ")
    
    If InStr(1, sName, " ") > 0 Then
         sASS = Split(sName): sWork = Left(sASS(0), 1) & Left(sASS(1), 1)
    Else
         sWork = Left(sWork, 1)
    End If
'-------------------------
ExitHere:
    GetInitials = sWork '!!!!!!!!!!!!!!!!
End Function

'======================================================================================================================================================
' Get Comment List
'======================================================================================================================================================
Public Function EXTRACTCOMMENTS(sPath As String, Optional bClose As Boolean, Optional DLM As String = "|", Optional SEP As String = vbCrLf) As String
Dim wdApp As Object, wdDoc As Object, sWordDocName As String, sRes As String, wdComment As Object
Dim I As Integer, sWork As String

Const wdGoToLine As Integer = 3
Const wdGoToAbsolute As Integer = 1
Const wdFirstCharacterLineNumber As Integer = 10
Const wdActiveEndAdjustedPageNumber = 1

   On Error GoTo ErrHandle
'------------------------------

Set wdApp = GetWordApp()
If wdApp Is Nothing Then Err.Raise 1000, , "Can't start Word App Object"

sWordDocName = FileNameOnly(sPath):
Set wdDoc = IsWordDoc(wdApp, sWordDocName)
If wdDoc Is Nothing Then Set wdDoc = wdApp.Documents.Open(sPath)

If wdDoc Is Nothing Then Err.Raise 1000, , "Can't open word Document: " & sPath
wdDoc.Activate
wdApp.Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=4
'--------------------------------------------------------------------------
DoCmd.Hourglass True

With wdDoc
  ' Process the Comments
  For I = 1 To .Comments.Count
    If sRes <> "" Then sRes = sRes & SEP
    Set wdComment = .Comments(I)
      sWork = wdComment.Index & DLM                                                         ' Index
      sWork = sWork & wdComment.Reference.Information(wdActiveEndAdjustedPageNumber) & DLM  ' Page Number
      sWork = sWork & wdComment.Reference.Information(wdFirstCharacterLineNumber) & DLM     ' Line Number
      sWork = sWork & wdComment.Author & DLM & wdComment.Initial & DLM                      ' Comment Author and his/her Initial
      sWork = sWork & wdComment.Date & DLM                                                  ' Comment Date
      sWork = sWork & wdComment.Range.Text & DLM                                            ' Comment Text itself
      sWork = sWork & wdComment.Reference.Text & DLM                                        ' Comment Reference Text ???
      sWork = sWork & wdComment.Scope.Text & DLM                                            ' Comment highlight words
      sWork = sWork & wdComment.Replies.Count & DLM                                         ' Comment How Many Replies
      If Not wdComment.Ancestor Is Nothing Then
         sWork = sWork & wdComment.Ancestor.Index
      End If
      Debug.Print sWork
      sRes = sRes & sWork
  Next I
End With


'---------------------------------
ExitHere:
    Debug.Print "The comment list for " & sWordDocName & " has built"
    DoCmd.Hourglass False: Set wdComment = Nothing
    EXTRACTCOMMENTS = sRes '!!!!!!!!!!!!!!!
    If bClose Then
       wdDoc.Close
       wdApp.Quit
       Set wdDoc = Nothing: Set wdApp = Nothing
    End If
    Exit Function
'--------------------
ErrHandle:
    ErrPrint2 "EXTRACTCOMMENTS", Err.Number, Err.Description, MOD_NAME
    Err.Clear: DoCmd.Hourglass False
End Function

Public Sub EXTRACTOUTLINE_TEST()
Dim sPath As String, sRes As String

sPath = OpenDialog(GC_FILE_PICKER, "Pick a file", , False, CurrentProject.Path)

sRes = EXTRACTOUTLINE(sPath)

End Sub

'======================================================================================================================================================
' Get WORD DOC STRUCTURE (OUTLINE)
'======================================================================================================================================================
Public Function EXTRACTOUTLINE(sPath As String, Optional bClose As Boolean, Optional DLM As String = "|", Optional SEP As String = vbCrLf) As String
Dim wdApp As Object, wdDoc As Object, sWordDocName As String

Dim iHead As Integer, sHead As String, Heading_txt As String, Heading_lvl As Variant, Heading_lne As Long, Heading_pge As Integer, Heading_outline As String
Dim sWork As String

Const TO_HEADER_LEVEL As Integer = 5
Const wdStory As Integer = 5
Const wdMove As Integer = 0
Const wdCollapseEnd As Integer = 0
Const wdLine = 5
Const wdActiveEndPageNumber As Integer = 3
Const wdCollapseStart As Integer = 1
Const wdGoToAbsolute As Integer = 1
Const wdGoToLine As Integer = 3

   On Error GoTo ErrHandle
'------------------------------

Set wdApp = GetWordApp()
If wdApp Is Nothing Then Err.Raise 1000, , "Can't start Word App Object"

sWordDocName = FileNameOnly(sPath):
Set wdDoc = IsWordDoc(wdApp, sWordDocName)
If wdDoc Is Nothing Then Set wdDoc = wdApp.Documents.Open(sPath)

If wdDoc Is Nothing Then Err.Raise 1000, , "Can't open word Document: " & sPath
wdDoc.Activate
wdApp.Selection.GoTo What:=wdGoToLine, Which:=wdGoToAbsolute, Count:=4
'--------------------------------------------------------------------------
    DoCmd.Hourglass True
For iHead = 1 To TO_HEADER_LEVEL
    sHead = ("Heading " & iHead)
    wdApp.Selection.HomeKey wdStory, wdMove

    Do
       If sWork <> "" Then sWork = sWork & SEP
       With wdApp.Selection
          .MoveStart Unit:=wdLine, Count:=1
          .Collapse Direction:=wdCollapseEnd
       End With
        
       With wdApp.Selection.Find
          .ClearFormatting:          .Text = "":
          .MatchWildcards = False:   .Forward = True
          .style = wdDoc.Styles(sHead)
         If .Execute = False Then GoTo Level_exit
            .ClearFormatting
       End With

       Heading_txt = wdApp.CleanString(wdApp.Selection.Range.Text)
       Heading_lvl = wdApp.Selection.Range.ListFormat.ListLevelNumber
       Heading_outline = wdApp.Selection.Range.ListFormat.ListString
       Heading_lne = wdDoc.Range(0, wdApp.Selection.Range.End).Paragraphs.Count
       Heading_pge = wdApp.Selection.Information(wdActiveEndPageNumber)
       
       sWork = sWork & Heading_outline & DLM & Heading_lne & DLM & Heading_pge & DLM & Heading_txt
       
       If wdApp.Selection.style = "Heading 1" Then GoTo Level_exit
       wdApp.Selection.Collapse Direction:=wdCollapseStart
   Loop
Level_exit:
Next iHead
'---------------------------------
ExitHere:
    Debug.Print "The structure for " & sWordDocName & " has built"
    DoCmd.Hourglass False
    EXTRACTOUTLINE = sWork '!!!!!!!!!!!!!!!
    If bClose Then
       wdDoc.Close
       wdApp.Quit
       Set wdDoc = Nothing: Set wdApp = Nothing
    End If
    Exit Function
'--------------------
ErrHandle:
    ErrPrint2 "EXTRACTOUTLINE", Err.Number, Err.Description, MOD_NAME
    Err.Clear: DoCmd.Hourglass False
End Function

Private Function GetLevel(strItem As String) As Integer
    ' Return the heading level of a header from the
    ' array returned by Word.

    ' The number of leading spaces indicates the
    ' outline level (2 spaces per level: H1 has
    ' 0 spaces, H2 has 2 spaces, H3 has 4 spaces.

    Dim strTemp As String
    Dim strOriginal As String
    Dim intDiff As Integer

    ' Get rid of all trailing spaces.
    strOriginal = RTrim$(strItem)

    ' Trim leading spaces, and then compare with
    ' the original.
    strTemp = LTrim$(strOriginal)

    ' Subtract to find the number of
    ' leading spaces in the original string.
    intDiff = Len(strOriginal) - Len(strTemp)
    GetLevel = (intDiff / 2) + 1
End Function

'======================================================================================================================================================
' Check If word doc is openned
'======================================================================================================================================================
Public Function IsWordDoc(ByRef wdApp As Object, sDocName As String) As Object
Dim wdDoc As Object, I As Integer, sWork As String

On Error GoTo ErrHandle
'-----------------------
    For I = 1 To wdApp.Documents.Count
       sWork = wdApp.Documents(I)
       If sDocName = sWork Then
          Set wdDoc = wdApp.Documents(I)
          Exit For
       End If
    Next I
    If wdDoc Is Nothing Then Exit Function
'-----------------------
ExitHere:
    Set IsWordDoc = wdDoc '!!!!!!!!!!!!!
    Exit Function
'------------
ErrHandle:
    ErrPrint2 "IsWordDoc", Err.Number, Err.Description
    Err.Clear
End Function
'-------------------------------------------------------------------------------------------------------------------------------------------------
' Get Word Objects
'-------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetWordApp() As Object
Dim objApp As Object, strMsg As String

Const WORD_CLASS As String = "Word.Application"

On Error GoTo ErrHandle
'-----------------------------
    Set objApp = GetObject(, WORD_CLASS)

'-----------------------------
ExitHere:
    On Error GoTo 0
    Set GetWordApp = objApp
    Exit Function
'-----------
ErrHandle:
    Select Case Err.Number
    Case 429 ' ActiveX component can't create object
        Set objApp = CreateObject(WORD_CLASS)
        Resume ExitHere
    Case Else
        strMsg = "Error " & Err.Number & " (" & Err.Description _
            & ") in procedure GiveMeAnApp"
        ErrPrint2 "GetWordApp", Err.Number, Err.Description, MOD_NAME
        GoTo ExitHere
    End Select
End Function

