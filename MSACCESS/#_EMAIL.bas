Attribute VB_Name = "#_EMAIL"
'***************************************************************************************************************************************************************************
'***************************************************************************************************************************************************************************
' E-MAILand RELATED FUNCTIONS
'***************************************************************************************************************************************************************************
'***************************************************************************************************************************************************************************

Option Compare Database
Option Explicit


Private Type EmailTemplate
     TemplateName As String
     TemplateFile As String          ' html or otf file
     TMPLTEXT As String
     HTMLBody As String
     bKV As Boolean
     Keys() As String
     VALS() As String
     FromAccount As String
     BCC As String
     Subject As String
     SIGNATURE As String
     sTO As String
     IsEdit As Boolean
End Type

'=========================================================================================================================================================================
' SENT E-MAIL TO INVESTOR
'=========================================================================================================================================================================
Public Function SendSerialEmails() As Integer

End Function

'=========================================================================================================================================================================
' Функция отсылает быстрый e-Mail
'=========================================================================================================================================================================
Public Function SendMail(Optional sTO As String, Optional ContactName As String, Optional Subject As String, Optional MAILBODY As String, Optional bEDIT As Boolean = True)
Dim bRes As Boolean, Send_To As String, Send_Subject As String

On Error GoTo ErrHandle
'------------------------------------------
Send_To = ContactName & IIf(sTO <> "", " [" & sTO & "]", "someeamil@mail.com")
Send_Subject = IIf(Subject <> "", Subject, "E-MAIL SUBJECT " & Now())

DoCmd.SendObject , "", "", Send_To, "", "", Send_Subject, MAILBODY, bEDIT, ""
bRes = True
'------------------------------------------
ExitHere:
     SendMail = bRes '!!!!!!!!!!!!!!!!
     Exit Function
'----------------
ErrHandle:
     Select Case Err.Number
        Case 2501:  ' User Cancel E-Mail
                Err.Clear:  bRes = False
                Resume ExitHere
        Case Else:
                ErrPrint "SendMail", Err.Number, Err.Description
                Err.Clear: Resume ExitHere
     End Select
End Function
'=========================================================================================================================================================================
' SENT E-MAIL TO INVESTOR
'=========================================================================================================================================================================
Public Sub SendEMAILToInvestor(InvestorEMAIL As String, sTitle As String, PATHTOTMPL As String, FromName As String, sFrom As String, _
                                                                             sBCC As String, bEDIT As Boolean, sBody As String, sDEAR As String, Optional sKVS As String)
Dim HTMLBody As String, KVS() As String, nDim As Integer, I As Integer
Dim eTMPL As EmailTemplate, sTemplate As String

Const KVSDELIM As String = ";"
Const KVSSEQ As String = "="

On Error GoTo ErrHandle
'---------------------------------------------------
eTMPL.TemplateFile = PATHTOTMPL
eTMPL.BCC = sBCC
eTMPL.SIGNATURE = FromName
eTMPL.TemplateName = sTitle
eTMPL.FromAccount = sFrom
eTMPL.TMPLTEXT = sBody
eTMPL.Subject = sTitle
eTMPL.sTO = InvestorEMAIL
eTMPL.IsEdit = bEDIT
'---------------------------------------------------
If sKVS <> "" Then sKVS = sKVS & KVSDELIM
   sKVS = sKVS & "[DEAR INVESTOR]" & KVSSEQ & sDEAR & KVSDELIM & "[FROMNAMEFROMNAME]" & KVSSEQ & FromName
    
   KVS = Split(sKVS, KVSDELIM): nDim = UBound(KVS): eTMPL.bKV = True
   ReDim eTMPL.Keys(nDim): ReDim eTMPL.VALS(nDim)
    
    For I = 0 To nDim
        eTMPL.Keys(I) = Trim(Split(KVS(I), "=")(0))
        eTMPL.VALS(I) = Trim(Split(KVS(I), "=")(1))
    Next I
'---------------------------------------------------
If eTMPL.TemplateFile <> "" Then
    If Left(FileExt(eTMPL.TemplateFile), 3) = "htm" Then
          eTMPL.HTMLBody = ReadTextFile(PATHTOTMPL)
    End If
End If
'-----------------------------------------
'HTMLBody = ReadTextFile(PATHTOTMPL)
'HTMLBody = Replace(HTMLBody, "[DEAR INVESTOR]", sDEAR)
'HTMLBody = Replace(HTMLBody, "[BODYBODYBODY]", sBody)
'HTMLBody = Replace(HTMLBody, "[FROMNAMEFROMNAME]", FromName)
'------------------------------------------
sTemplate = EmlStr(eTMPL):


    Call SendOutlook(InvestorEMAIL, , sBCC, sFrom, FromName, sTitle, HTMLBody, 2, , , bEDIT, , sTemplate)
'------------------------------------------
ExitHere:
    Exit Sub
'---------------
ErrHandle:
    ErrPrint "SendMilToInvestor", Err.Number, Err.Description
    Err.Clear
End Sub



'Sub testEmail()
'   SendOutlook "valery.khvatov@gmail.com", , "alla.khvatov@gmail.com", "valery.khvatov@gmail.com", "Valery Khvatov+", "This is the the mail", "<html><head></head><body><p><b>TEST<b></p></body></html>", 2, , , True
'"C:\Users\valer\AppData\Roaming\Microsoft\Templates\EMAIL4.oft"
'End Sub


'=========================================================================================================================================================================
' Функция отсылает  e-Mail с помощью Outlook
'=========================================================================================================================================================================
Public Function SendOutlook(Optional sTO As String, Optional sCC As String, Optional sBCC As String, Optional sFrom As String, Optional FromName As String, _
                                                                                                                                            Optional Subject As String, _
                            Optional sText As String, Optional iFormat As Integer = 1, Optional AttachmentFiles As String, Optional iImportance As Integer = 1, _
                                                            Optional bEDIT As Boolean = True, Optional DLM As String = ";", Optional sTemplate As String) As Boolean
Dim oApp As Object, oMSG As Object, oAPPRecip As Object, oAPPAttach As Object
Dim accNo As Integer, foundAccount As Boolean, bFromTemplate As Boolean, eTPL As EmailTemplate
Dim sWork() As String, nDim As Integer, I As Integer, bRes As Boolean

Const o_olFormatHTML As Integer = 2
Const o_olFormatPlain  As Integer = 1
Const o_olFormatRichText  As Integer = 3
Const o_olFormatUnspecified  As Integer = 0

Const o_olMailItem As Integer = 0

Const o_olBCC As Integer = 3
Const o_olCC As Integer = 2
Const o_olTo As Integer = 1

Const o_olImportanceHigh As Integer = 2
Const o_olImportanceLow As Integer = 0
Const o_olImportanceNormal As Integer = 1
Const o_olFolderDrafts As Integer = 16

On Error GoTo ErrHandle
'-------------------------------------------------------------------------------------------------------------------------------------------
If sTemplate <> "" Then
        eTPL = StrEml(sTemplate)
        bFromTemplate = True
Else
        eTPL.sTO = sTO: eTPL.BCC = sBCC: eTPL.SIGNATURE = FromName: eTPL.FromAccount = sFrom
        eTPL.TemplateName = Subject: eTPL.Subject = Subject: eTPL.TMPLTEXT = sText
        eTPL.IsEdit = bEDIT: eTPL.TMPLTEXT = sText
        eTPL.BCC = sBCC
End If

If eTPL.sTO = "" Then Err.Raise 1000, , "Missing T0:"
If Not IsEmail(eTPL.sTO) Then Err.Raise 1000, , "Wrong format T0:" & eTPL.sTO
'------------------------------------------
Set oApp = CreateObject("Outlook.Application") ' Create the Outlook session

If bFromTemplate Then
  Set oMSG = oApp.CreateItemFromTemplate(eTPL.TemplateFile, _
            oApp.Session.GetDefaultFolder(o_olFolderDrafts))
Else
  Set oMSG = oApp.CreateItem(o_olMailItem) ' Create the message
End If
'----------------------------------------------------------------------------------
' SEARCH ACCOUNT
If sFrom <> "" Then
    For I = 1 To oApp.Session.Accounts.Count
       If oApp.Session.Accounts.Item(I).smtpAddress = eTPL.FromAccount Then
          accNo = I: foundAccount = True: Exit For
       End If
    Next I
End If
'----------------------------------------------------------------------------------
With oMSG
              ' Add the To recipient(s) to the message.
              If eTPL.sTO <> "" Then
                sWork = Split(eTPL.sTO, "DLM"): nDim = UBound(sWork)
                For I = 0 To nDim
                   Set oAPPRecip = .Recipients.Add(sWork(I))
                    oAPPRecip.Type = o_olTo
                Next I
              End If
              ' Add the CC recipient(s) to the message.
              If sCC <> "" Then
                sWork = Split(sCC, "DLM"): nDim = UBound(sWork)
                For I = 0 To nDim
                   Set oAPPRecip = .Recipients.Add(sWork(I))
                  oAPPRecip.Type = o_olCC
                Next I
              End If
             ' Add the BCC recipient(s) to the message.
              If sBCC <> "" Then
                sWork = Split(sBCC, "DLM"): nDim = UBound(sWork)
                For I = 0 To nDim
                   Set oAPPRecip = .Recipients.Add(sWork(I))
                  oAPPRecip.Type = o_olBCC
                Next I
              End If
             '-----------------------------------------------------------------------------------------------------
             If foundAccount Then
                 Set .SendUsingAccount = oApp.Session.Accounts.Item(accNo)
             End If
             '------------------------------------------------------------------------------------------------------
             ' Set the Subject, Body, and Importance of the message.
             If eTPL.Subject = "" Then
                    .Subject = "This is an Automation test with Microsoft Outlook"
             Else
                    .Subject = eTPL.Subject
             End If
             
             If bFromTemplate Then
                 eTPL.HTMLBody = .HTMLBody
                 Call ProcessEmailTemplate(eTPL)
             End If
             
             If eTPL.TMPLTEXT <> "" Then
                  Select Case iFormat
                  Case 1:         ' PLAIN
                       .BodyFormat = o_olFormatPlain
                       .BODY = eTPL.TMPLTEXT
                  Case 2:         ' HTML
                       .BodyFormat = o_olFormatHTML
                       .HTMLBody = eTPL.HTMLBody
                  End Select
             '.HTMLBody = Forms![frmDMPClientEmail]![ClientMiscLtr]    'Note:[ClientMiscLtr] will contain HTML text similar to this: "<B><u><i> This is HTML Text </b></u></i>"
             End If
             .SentOnBehalfOfName = "BGX Project <team@bgx.ai>"
             .Importance = iImportance ' olImportanceHigh  'High importance
             
             If AttachmentFiles <> "" Then
                 sWork = Split(AttachmentFiles, DLM): nDim = UBound(sWork)
                 For I = 0 To nDim
                     Set oAPPAttach = .Attachments.Add(sWork(I))
                 Next I
             End If
             '------------------------------------------------------------------------------
             ' Resolve each Recipient's name.
             For Each oAPPRecip In .Recipients
                 oAPPRecip.Resolve
             Next

             ' Should we display the message before sending?
             If bEDIT Then
                 .Display
             Else
                 .Save
                 .Send
             End If
          End With
'-----------------------------------------
bRes = True
'------------------------------------------
ExitHere:
     SendOutlook = bRes '!!!!!!!!!!!!!!!!
     Set oAPPAttach = Nothing: Set oAPPRecip = Nothing:
     Set oMSG = Nothing: Set oApp = Nothing
     Exit Function
'----------------
ErrHandle:
                ErrPrint "SendOutlook", Err.Number, Err.Description
                Err.Clear: Resume ExitHere
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
                                                                                                    Optional sModName As String = "mod_EMAIL") As String
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

Public Sub TestEmailTemplSer()
Dim TMPL As EmailTemplate, sWork As String
Dim NewTMPL As EmailTemplate

ReDim TMPL.Keys(4): ReDim TMPL.VALS(4)

    TMPL.TemplateName = "Test": TMPL.bKV = True
    TMPL.TemplateFile = "C:/Users/valer/Documents/DBs/BGXUSERS/TEMPLATES/EMAIL4_files/image001.gif"
    TMPL.BCC = "alla@cogeco.ca"
    TMPL.FromAccount = "valery.khvatov@gmail.com"
    TMPL.Subject = "THIS IS A TITLE"
    TMPL.SIGNATURE = "Elena Gibas"
    TMPL.HTMLBody = "<img width=280 height=82 src=KEY1" & Chr(34) & "cid:image002.png@01D3D0BA.553CAA90" & Chr(34) & "v:shapes=" & Chr(34) & "_x0000_i1025" & Chr(34) & "><![endif]></span><o:p></o:p></p><" & vbCrLf & _
                    "p align=center style=KEY2'text-align:center'><span lang=EN-US style='font-size:10.0pt;font-family:" & Chr(34) & "Verdana KEY3" & Chr(34) & ",sans-serif;mso-ansi-language:EN-US'>" & vbCrLf & _
                    "Dear Alla,</span><o:p></o:p></p><p align=center style='text-align:center'>KEY4<span lang=EN-US style='font-size:10.0pt;font-family:" & vbCrLf & _
                    "Verdana" & Chr(34) & ",sans-serif;mso-ansi-language:EN-US'>"
    TMPL.Keys(0) = "KEY0": TMPL.VALS(0) = "VAL0"
    TMPL.Keys(1) = "KEY1": TMPL.VALS(1) = "VAL1"
    TMPL.Keys(2) = "KEY2": TMPL.VALS(2) = "VAL2"
    TMPL.Keys(3) = "KEY3": TMPL.VALS(3) = "VAL3"
    TMPL.Keys(4) = "KEY4": TMPL.VALS(4) = "VAL4"
    
    sWork = EmlStr(TMPL)
    
    NewTMPL = StrEml(sWork)
    
    Call ProcessEmailTemplate(NewTMPL)
    
    Debug.Assert False
End Sub

Public Sub TestReadTMPL()
Dim TMPL As EmailTemplate
      TMPL = ReadTemplate(2)
      Debug.Assert False
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------------------
' READ EMAIL TEMPLATE FROM TABLE
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function ReadTemplate(TemplateID As Long, Optional TBL As String = "_TEMPLATES", Optional DLM As String = ";", _
                                                                                                       Optional SEQV As String = "=") As EmailTemplate
Dim eTMPL As EmailTemplate, RS As DAO.Recordset, SQL As String
Dim sKeyWords As String, KVS() As String, nDim As Integer, I As Integer

On Error GoTo ErrHandle
'-----------------------------------------------------------------------
SQL = "SELECT * FROM " & SHT(TBL) & " WHERE ID = " & TemplateID & ";"
Set RS = CurrentDb.OpenRecordset(SQL)
With RS
    If Not .EOF Then
         .MoveLast: .MoveFirst
             eTMPL.TemplateName = Nz(!TMPLName, "")
             sKeyWords = Nz(!TMPLSQL, "")
             eTMPL.HTMLBody = Nz(!TMPLBody, "")
             eTMPL.BCC = Nz(!BCC, "")
             eTMPL.SIGNATURE = Nz(!SIGNTEXT, "")
             eTMPL.FromAccount = Nz(!FROMEACCAOUNT, "")
             eTMPL.TemplateFile = Nz(!TMPLFile, "")
    End If
End With
'--------------------------------------------------
    eTMPL.Subject = eTMPL.TemplateName
If sKeyWords <> "" Then
    KVS = Split(sKeyWords, DLM): nDim = UBound(KVS)
    ReDim eTMPL.Keys(nDim): ReDim eTMPL.VALS(nDim)
    For I = 0 To nDim
        eTMPL.Keys(I) = Trim(Split(KVS(I), "=")(0))
        eTMPL.VALS(I) = Trim(Split(KVS(I), "=")(1))
    Next I
End If
'--------------------------------------------------
ExitHere:
    ReadTemplate = eTMPL  '!!!!!!!!!!!!!!!!!!!!!!
    Set RS = Nothing
    Exit Function
'-------------
ErrHandle:
     ErrPrint "ReadTemplate", Err.Number, Err.Description
     Err.Clear
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' FUNC SERIALIZE TO STRING EMAIL TEMPLATE
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function EmlStr(eTMPL As EmailTemplate, Optional DLM = "¤", Optional SEQV As String = "") As String
Dim nDim As Integer, sRes As String, sKV As String, I As Integer

On Error GoTo ErrHandle
'----------------------------------------------------------
If eTMPL.bKV Then
   nDim = UBound(eTMPL.Keys)
   For I = 0 To nDim
        sKV = sKV & eTMPL.Keys(I) & SEQV & eTMPL.VALS(I) & DLM
   Next I
End If
If sKV <> "" Then sKV = Left(sKV, Len(sKV) - Len(DLM))


sRes = "[EMAILTEMPLATE0000]" & " " & eTMPL.TemplateName & vbCrLf & "[EMAILTEMPLATEFILE]" & SEQV & eTMPL.TemplateFile & vbCrLf
sRes = sRes & "[EMTMPLFROMACCOUNT]" & SEQV & eTMPL.FromAccount & vbCrLf & "[EMLTMPLBCC0000000]" & SEQV & eTMPL.BCC & vbCrLf & _
              "[EMTMPLSIGNATURE00]" & SEQV & eTMPL.SIGNATURE & vbCrLf & "[EMTMPLATETO000000]" & SEQV & eTMPL.sTO & vbCrLf & _
              "[EMLTMPLATESUBJECT]" & SEQV & eTMPL.Subject & vbCrLf & "[EMLTMPLTISEDIT000]" & SEQV & eTMPL.IsEdit & vbCrLf
sRes = sRes & "[EMAILTEMPLATESKV0]" & vbCrLf & sKV & vbCrLf & "[EMAILTEMPLATEBODY]" & vbCrLf & eTMPL.TMPLTEXT
'----------------------------------------------------------
ExitHere:
      EmlStr = sRes '!!!!!!!!!!!!!!
      Exit Function
'---------------------
ErrHandle:
      ErrPrint "EmlStr", Err.Number, Err.Description
      Err.Clear
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' FUNC CONVERT SERIALIZED STRING TO EMAIL TEMPLATE
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function StrEml(sTMPL As String, Optional DLM = "¤", Optional SEQV As String = "") As EmailTemplate
Dim eTMPL As EmailTemplate, sWork As String, nDim As Integer, I As Integer, sKV As String, ROWS() As String
Dim bReadKV As Boolean, bBody As Boolean, sBody As String, KV() As String

Const WLIm As Integer = 19

On Error GoTo ErrHandle
'----------------------------------------------------------
If sTMPL = "" Then Exit Function

ROWS = Split(sTMPL, vbCrLf): nDim = UBound(ROWS)
If Left(ROWS(0), WLIm) <> "[EMAILTEMPLATE0000]" Then Err.Raise 1000, , "Error with Template Serialization"
'---------------------------------------
For I = 0 To nDim
      sWork = Left(ROWS(I), WLIm)
      Select Case sWork
      Case "[EMAILTEMPLATE0000]":
           eTMPL.TemplateName = Trim(Replace(ROWS(I), "[EMAILTEMPLATE0000]", ""))
      Case "[EMAILTEMPLATEFILE]":
           eTMPL.TemplateFile = Trim(Split(ROWS(I), SEQV)(1))
      Case "[EMTMPLFROMACCOUNT]":
           eTMPL.FromAccount = Trim(Split(ROWS(I), SEQV)(1))
      Case "[EMLTMPLBCC0000000]":
           eTMPL.BCC = Trim(Split(ROWS(I), SEQV)(1))
      Case "[EMTMPLATETO000000]":
           eTMPL.sTO = Trim(Split(ROWS(I), SEQV)(1))
      Case "[EMTMPLSIGNATURE00]":
           eTMPL.SIGNATURE = Trim(Split(ROWS(I), SEQV)(1))
      Case "[EMLTMPLATESUBJECT]":
           eTMPL.Subject = Trim(Split(ROWS(I), SEQV)(1))
      Case "[EMLTMPLTISEDIT000]":
           eTMPL.IsEdit = CBool(Trim(Split(ROWS(I), SEQV)(1)))
      Case "[EMAILTEMPLATESKV0]":
           eTMPL.bKV = True: bReadKV = True
      Case "[EMAILTEMPLATEBODY]":
           bReadKV = False: bBody = True
      Case Else
           If bReadKV Then
              sKV = sKV & ROWS(I) & DLM
           ElseIf bBody Then
              sBody = sBody & ROWS(I) & vbCrLf
           End If
      End Select
Next I

eTMPL.TMPLTEXT = Trim(sBody):
nDim = -1: ReDim eTMPL.Keys(0): ReDim eTMPL.VALS(0)

If sKV <> "" Then
   sKV = Left(sKV, Len(sKV) - Len(DLM))
    KV = Split(sKV, DLM): nDim = UBound(KV)
    ReDim eTMPL.Keys(nDim): ReDim eTMPL.VALS(nDim)
    For I = 0 To nDim
        eTMPL.Keys(I) = Trim(Split(KV(I), SEQV)(0))
        eTMPL.VALS(I) = Trim(Split(KV(I), SEQV)(1))
    Next I
End If
'----------------------------------------------------------
ExitHere:
      StrEml = eTMPL '!!!!!!!!!!!!!!
      Exit Function
'---------------------
ErrHandle:
      ErrPrint "StrEml", Err.Number, Err.Description
      Err.Clear
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' FUNC PROCESS TEMPLATE
'------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ProcessEmailTemplate(ByRef eTMPL As EmailTemplate)
Dim nDim As Integer, I As Integer, sRes As String, sText As String

On Error GoTo ErrHandle

If eTMPL.TMPLTEXT <> "" Then
   If eTMPL.HTMLBody <> "" Then
      eTMPL.HTMLBody = Replace(eTMPL.HTMLBody, "[BODYBODYBODY]", eTMPL.TMPLTEXT)
   Else
      eTMPL.HTMLBody = eTMPL.TMPLTEXT
   End If
End If
    
'HTMLBody = Replace(HTMLBody, "[DEAR INVESTOR]", sDEAR)
'HTMLBody = Replace(HTMLBody, "[BODYBODYBODY]", sBody)
'HTMLBody = Replace(HTMLBody, "[FROMNAMEFROMNAME]", FromName)
'----------------------------------------------------------
If eTMPL.bKV Then
    nDim = UBound(eTMPL.Keys)
    For I = 0 To nDim
        eTMPL.HTMLBody = Replace(eTMPL.HTMLBody, eTMPL.Keys(I), eTMPL.VALS(I))
    Next I
End If
'----------------------------------------------------------
ExitHere:
      Exit Sub
'---------------------
ErrHandle:
      ErrPrint "ProcessEmailTemplate", Err.Number, Err.Description
      Err.Clear
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------
' Last Email Date
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function LastEmail(UserId As Long, Optional TBL As String = "NOTES") As String
    LastEmail = CStr(Nz(DMax("DateUpdate", TBL, "ParentUser = " & UserId & " AND NoteType = 6"), "")) '!!!!!!!!!!
End Function
'----------------------------------------------------------------------------------------------------------------------------------------------------------
' Email Sum for User
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function EmailCount(UserId As Long, Optional TBL As String = "NOTES") As Integer
    EmailCount = Nz(DCount("DateUpdate", TBL, "ParentUser = " & UserId & " AND NoteType = 6"), 0) '!!!!!!!!!!
End Function
