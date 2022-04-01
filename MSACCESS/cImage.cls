VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
' DESCRIPTION: Image Processor (Use IrfanView for some image operation)
' VERSION    : v 0.2 30.01.2015
' COPYRIGHT  : V.Khvatov DigitalXpert
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
Option Compare Database
Option Explicit


'######################################################################################################################################################
'======================================================================================================================================================
' Save Clipboard Image To File
'======================================================================================================================================================
Public Function SaveClipboardToImage(Optional sFile As String, Optional sFolder As String, Optional sAddOption As String) As String
Dim sIMAGEPROCESSOR As String, sEXECUTE As String, SaveFolder As String, ImageFile As String


Const FOR_IMAGE As String = "IMAGES"

On Error GoTo ErrHandle
'-------------------------------------------------------------------
    sIMAGEPROCESSOR = GetExecutor("IRFANVIEW")
    If sIMAGEPROCESSOR = "" Then Err.Raise 1000, , "Image Processor (IrfanView) is absent on this PC"
    
    If sFolder <> "" Then
         SaveFolder = sFolder
    Else
         SaveFolder = CurrentProject.Path & "\" & FOR_IMAGE
    End If
    If Not IsFolderExists(SaveFolder) Then FolderCreate SaveFolder
    If Right(SaveFolder, 1) <> "\" Then SaveFolder = SaveFolder & "\"
    
    If sFile <> "" Then
       ImageFile = sFile
    Else
       ImageFile = GetRandomAlphaString & ".jpg"
    End If
    
    
    sEXECUTE = sIMAGEPROCESSOR & IIf(sAddOption = "", "", " " & Trim(sAddOption)) & _
                          " /silent /clippaste /convert=" & SaveFolder & ImageFile
    Shell sEXECUTE
'-----------------------------------------------------------------------------
ExitHere:
     SaveClipboardToImage = SaveFolder & ImageFile  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
     Exit Function
'---------------------------------
ErrHandle:
    Beep
    ErrPrint "SaveClipboardToImage", Err.Number, Err.Description
    Err.Clear
End Function
'======================================================================================================================================================
' Get thumbnail for media file
'======================================================================================================================================================
Public Function ThumbnailFile(sFile As String, Optional ThumbFolder As String, Optional Prefix As String = "th_", Optional IWidth As Long = 200, _
                                                                              Optional iHeight As Long = 100, Optional WaitTime As Long = 0) As String
Dim sFolder As String, sIMAGEPROCESSOR As String, sCmdExec As String, sThumbnail As String

Const FOR_THUMBNAILS As String = "THUMB"

On Error GoTo ErrHandle
'------------------------------
    sIMAGEPROCESSOR = GetExecutor("IRFANVIEW")
    If sIMAGEPROCESSOR = "" Then Err.Raise 1000, , "Image Processor (IrfanView) is absent on this PC"
    If ThumbFolder <> "" Then
       sFolder = ThumbFolder
    Else
       sFolder = CurrentProject.Path & "\" & FOR_THUMBNAILS
    End If
    If Not IsFolderExists(sFolder) Then FolderCreate sFolder
    If Right(sFolder, 1) <> "\" Then sFolder = sFolder & "\"
    
    sThumbnail = ChangeXtension(sFolder & Prefix & FileNameOnly(sFile), "jpg")
    If IsFileExists(sThumbnail) Then Kill sThumbnail
'------------------------------------------------------------------------
     sCmdExec = sIMAGEPROCESSOR & " " & sFile & " /silent /thumbs /resize=(" & IWidth & "," & iHeight & ") /aspectratio /resample /convert=" & _
                               sThumbnail
     Shell sCmdExec
     If WaitTime > 0 Then Wait WaitTime
'---------------------------------------------------------------------------------------------------
ExitHere:
     If Dir(sThumbnail) <> "" Then ThumbnailFile = sThumbnail '!!!!!!!
     Exit Function
'----------
ErrHandle:
    ErrPrint "ThumbnailFile", Err.Number, Err.Description
    Err.Clear:  Resume ExitHere
End Function

'######################################################################################################################################################
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Get BinPath To IrfanView
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetIrfanViewPath() As String
Dim oWScript As Object, sPath As String

Const REG_PATH As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\IrfanView\shell\open\command\"
On Error GoTo ErrHandle
'-----------------------------------------------
  sPath = GetLocal("IrfanView")                 ' TRY TO GET IRFANVIEW FROM LOCAL
  If sPath <> "" Then GoTo ExitHere             ' SUCCESS

  Set oWScript = CreateObject("WScript.Shell")  ' TRY TO GET IRFANVIEW FROM REGISTRY
  sPath = oWScript.RegRead(REG_PATH)
  If sPath <> "" Then
     sPath = Replace(sPath, "%1", ""): sPath = Trim(Replace(sPath, Chr(34), ""))
     SetLocal "IrfanView", sPath
     GoTo ExitHere
  End If
                                                ' TRY TO GET IRFANVIEW PATH MNUALLY
  If MsgBox("Can't extract Binary Path for IrfanView automatically. To process some image operation you should have Installed IrfanView on your computer." & vbCrLf & _
             "Do you want navigate to IrfanView manually?", vbYesNoCancel + vbQuestion, "IrfanViewPath") = vbYes Then
             sPath = OpenDialog(GC_FILE_PICKER, "IrfanView Binary Path", "Binary Files,*.exe", False, CurrentProject.Path)
             If sPath <> "" Then SetLocal "IrfanView", sPath
  End If
'------------------------------------------------
ExitHere:
  GetIrfanViewPath = sPath '!!!!!!!!!!!!!!!
  Set oWScript = Nothing
  Exit Function
'---------------
ErrHandle:
  ErrPrint "GetIrfanViewPath", Err.Number, Err.Description
  Err.Clear: Resume ExitHere
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
                                                                                                  Optional sModName As String = "cImage") As String
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

