VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cClipBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************************************************************************************************
' The class is manipulator for clipboard object
'******************************************************************************************************************************************************************
Option Explicit

Private Const DATAOBJECT_BINDING As String = "new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}"
Private Const MOD_NAME As String = "cClipBoard"

' Required data structures
Private Type POINTAPI
    x As Long
    y As Long
End Type

#If Mac Then
    ' do nothing
#Else
    #If Win64 Then
' Clipboard Manager Functions
        Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
        Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
        Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
        Private Declare PtrSafe Function SetClipboardData Lib "user32" Alias "SetClipboardDataA" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
        Private Declare PtrSafe Function GetClipboardData Lib "user32" Alias "GetClipboardDataA" (ByVal wFormat As Long) As LongPtr
        Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long

' Other required Win32 APIs
        Private Declare PtrSafe Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal HDROP As LongPtr, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
        Private Declare PtrSafe Function DragQueryPoint Lib "shell32.dll" (ByVal HDROP As LongPtr, lpPoint As POINTAPI) As Long
        
        Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
        Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
        Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
        Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
        Private Declare PtrSafe Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
        
        Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
                                                         ByVal lpString2 As Any) As LongPtr
    #Else
        
' Clipboard Manager Functions
        Private Declare Function EmptyClipboard Lib "user32" () As Long
        Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
        Private Declare Function CloseClipboard Lib "user32" () As Long
        Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
        Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
        Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long

' Other required Win32 APIs
        Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal HDROP As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
        Private Declare Function DragQueryPoint Lib "shell32.dll" (ByVal HDROP As Long, lpPoint As POINTAPI) As Long
        Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
        Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
        Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
        Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
        Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
        Private Declare Function lstrcpyn Lib "kernel32" Alias "lstrcpynA" (ByVal lpString1 As String, ByVal lpString2 As String, ByVal iMaxLength As Long) As Long
    #End If
#End If

'***********************************************************
' Predefined Clipboard Formats
Private Const CF_TEXT As Long = 1
Private Const CF_BITMAP As Long = 2
Private Const CF_METAFILEPICT As Long = 3
Private Const CF_SYLK As Long = 4
Private Const CF_DIF As Long = 5
Private Const CF_TIFF As Long = 6
Private Const CF_OEMTEXT As Long = 7
Private Const CF_DIB As Long = 8
Private Const CF_PALETTE As Long = 9
Private Const CF_PENDATA As Long = 10
Private Const CF_RIFF As Long = 11
Private Const CF_WAVE As Long = 12
Private Const CF_UNICODETEXT As Long = 13
Private Const CF_ENHMETAFILE As Long = 14
Private Const CF_HDROP As Long = 15
Private Const CF_LOCALE As Long = 16
Private Const CF_MAX As Long = 17

' New shell-oriented clipboard formats
Private Const CFSTR_SHELLIDLIST As String = "Shell IDList Array"
Private Const CFSTR_SHELLIDLISTOFFSET As String = "Shell Object Offsets"
Private Const CFSTR_NETRESOURCES As String = "Net Resource"
Private Const CFSTR_FILEDESCRIPTOR As String = "FileGroupDescriptor"
Private Const CFSTR_FILECONTENTS As String = "FileContents"
Private Const CFSTR_FILENAME As String = "FileName"
Private Const CFSTR_PRINTERGROUP As String = "PrinterFriendlyName"
Private Const CFSTR_FILENAMEMAP As String = "FileNameMap"

' Global Memory Flags
Private Const GMEM_FIXED = &H0
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_NOCOMPACT = &H10
Private Const GMEM_NODISCARD = &H20
Private Const GMEM_ZEROINIT = &H40
Private Const GMEM_MODIFY = &H80
Private Const GMEM_DISCARDABLE = &H100
Private Const GMEM_NOT_BANKED = &H1000
Private Const GMEM_SHARE = &H2000
Private Const GMEM_DDESHARE = &H2000
Private Const GMEM_NOTIFY = &H4000
Private Const GMEM_LOWER = GMEM_NOT_BANKED
Private Const GMEM_VALID_FLAGS = &H7F72
Private Const GMEM_INVALID_HANDLE = &H8000
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Private Type DROPFILES
    pFiles As Long
    pt As POINTAPI
    fNC As Long
    fWide As Long
End Type
'***************************************************************
Private Const MAXSIZE = 4096
Private DataObj As Object

'================================================================================================================================================================
' Save file from clipboard to Disk
'================================================================================================================================================================
Public Function SaveFilesFromClipboard(DestinationFolder As String) As Boolean
Dim sRes As String
   
   On Error GoTo ErrHandle
'----------------------------------
            SaveFilesFromClipboard = False
            If Not CBool(IsClipboardFormatAvailable(CF_HDROP)) Then Exit Function
            CreateObject("Shell.Application").NameSpace(CVar(DestinationFolder)).Self.InvokeVerb "Paste"
            SaveFilesFromClipboard = True
'----------------------------------
ExitHere:
       SaveFilesFromClipboard = sRes '1111111111111111
       Exit Function
'-------------------
ErrHandle:
       ErrPrint2 "SaveFilesFromClipboard", Err.Number, Err.Description
       Err.Clear
End Function
        
'================================================================================================================================================================
' Check if Text in Clipboard
'================================================================================================================================================================
Public Function Is_Txt_in_Clipboard() As Boolean
Dim bRes As Boolean

   On Error GoTo ErrHandle
'----------------------------------
If IsClipboardFormatAvailable(CF_TEXT) Then bRes = True
If IsClipboardFormatAvailable(CF_OEMTEXT) Then bRes = True
If IsClipboardFormatAvailable(CF_UNICODETEXT) Then bRes = True

'----------------------------------
ExitHere:
       Is_Txt_in_Clipboard = bRes '1111111111111111
       Exit Function
'-------------------
ErrHandle:
       ErrPrint2 "Is_Txt_in_Clipboard", Err.Number, Err.Description
       Err.Clear
       
End Function
'================================================================================================================================================================
' Chtck if Image in Clipboard
'================================================================================================================================================================
Public Function Is_Pic_in_Clipboard() As Boolean
Dim bRes As Boolean

   On Error GoTo ErrHandle
'----------------------------------
If IsClipboardFormatAvailable(CF_BITMAP) Then bRes = True
If IsClipboardFormatAvailable(CF_DIB) Then bRes = True
'If IsClipboardFormatAvailable(CF_ENHMETAFILE) Then bRes = True

'----------------------------------
ExitHere:
       Is_Pic_in_Clipboard = bRes '1111111111111111
       Exit Function
'-------------------
ErrHandle:
       ErrPrint2 "Is_Pic_in_Clipboard", Err.Number, Err.Description
       Err.Clear
       
End Function

'================================================================================================================================================================
' Copy Single File To ClipBoard
'================================================================================================================================================================
Public Function ClipboardCopySingleFile(sFile As String) As Boolean

Dim Files(0) As String
Files(0) = sFile
ClipboardCopyFiles Files()

End Function

'================================================================================================================================================================
' Copy File Bundles To ClipBoard
'================================================================================================================================================================
Public Function ClipboardCopyFiles(Files() As String) As Boolean

Dim Data As String, df As DROPFILES
Dim I As Long

#If Win64 Then
    Dim hGlobal As LongPtr, lpGlobal As LongPtr
#Else
    Dim hGlobal As Long, lpGlobal As LongPtr
#End If

On Error GoTo ErrHandle
'-------------------------------------
If OpenClipboard(0&) Then                             ' Open and clear existing crud off clipboard.
    Call EmptyClipboard
    
    For I = LBound(Files) To UBound(Files)            ' Build double-null terminated list of files.
        Data = Data & Files(I) & vbNullChar
    Next
    Data = Data & vbNullChar

    hGlobal = GlobalAlloc(GHND, Len(df) + Len(Data))  ' Allocate and get pointer to global memory,
    If hGlobal Then                                   ' then copy file list to it.
        lpGlobal = GlobalLock(hGlobal)

        
        df.pFiles = Len(df)                           ' Build DROPFILES structure in global memory.
        Call CopyMem(ByVal lpGlobal, df, Len(df))
        Call CopyMem(ByVal (lpGlobal + Len(df)), ByVal Data, Len(Data))
        Call GlobalUnlock(hGlobal)

        
        If SetClipboardData(CF_HDROP, hGlobal) Then   ' Copy data to clipboard, and return success.
            ClipboardCopyFiles = True
        End If
    End If
    
    Call CloseClipboard                               ' Clean up
End If
'-------------------------------------
ExitHere:
     Exit Function
'-----------------
ErrHandle:
     ErrPrint2 "ClipboardCopyFiles", Err.Number, Err.Description, MOD_NAME
     Err.Clear
End Function

'=============================================================================================================================================================
' Save Clipboard File to new path
'============================================================================================================================================================
Public Function ClipboardPasteSingleFile(sNewFilePath As String) As Long
Dim Files(0) As String
     Files(0) = sNewFilePath
'-----------------------------
ExitHere:
    ClipboardPasteSingleFile = ClipboardPasteFiles(Files) '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'=============================================================================================================================================================
' Save Clipboard Files to new path
'============================================================================================================================================================
Public Function ClipboardPasteFiles(Files() As String) As Long

Dim nFiles As Long, I As Long
Dim desc As String, fileName As String
Dim pt As POINTAPI

Const MAX_PATH As Long = 260

On Error GoTo ErrHandle
'-----------------------------
#If Win64 Then
   Dim HDROP As LongPtr
#Else
   Dim HDROP As Long
#End If

If IsClipboardFormatAvailable(CF_HDROP) Then    ' Insure desired format is there, and open clipboard.
    If OpenClipboard(0&) Then
        HDROP = GetClipboardData(CF_HDROP)      ' Get handle to Dropped Filelist data, and number of files.
        nFiles = DragQueryFile(HDROP, -1&, "", 0)

        'ReDim Files(0 To nFiles - 1) As String  ' Allocate space for return and working variables.
        fileName = Space(MAX_PATH)

        For I = 0 To nFiles - 1                 ' Retrieve each filename in Dropped Filelist.
            Call DragQueryFile(HDROP, I, fileName, Len(fileName))
            Files(I) = TrimNull(fileName)
        Next

        Call CloseClipboard                     ' Clean up
    End If

    ClipboardPasteFiles = nFiles                ' Assign return value equal to number of files dropped
End If
'--------------------------------
ExitHere:
     Exit Function
'------------------
ErrHandle:
     ErrPrint2 "ClipboardPasteFiles", Err.Number, Err.Description, MOD_NAME
     Err.Clear
End Function

'=============================================================================================================================================================
' Empty Clipboard
'============================================================================================================================================================
Public Sub ClearClipboard()
    OpenClipboard (0&)
    EmptyClipboard
    CloseClipboard
End Sub
'=============================================================================================================================================================
' Paste Text
'============================================================================================================================================================
Public Function Paste() As String

On Error GoTo ErrHandle
'----------------------------------
     With DataObj
        .GetFromClipboard
         Paste = .GetText
     End With
     Exit Function
'----------------------------------
ErrHandle:
     ErrPrint2 "Paste", Err.Number, Err.Description, MOD_NAME
     Err.Clear
End Function
'=============================================================================================================================================================
' Copy Text
'============================================================================================================================================================
Public Sub Copy(ByVal sText As String)
On Error GoTo ErrHandle
'----------------------------------
With DataObj
     .SetText sText
     .PutInClipboard
End With
     Exit Sub
'----------------------------------
ErrHandle:
     ErrPrint2 "Copy", Err.Number, Err.Description, MOD_NAME
     Err.Clear
End Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
    On Error GoTo ErrHandle
'-----------------------------------------------
    Set DataObj = CreateObject(DATAOBJECT_BINDING)
    Exit Sub
'-------------
ErrHandle:
    ErrPrint2 "Class_Initialize", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Sub
Private Sub Class_Terminate()
     Set DataObj = Nothing
End Sub

Private Function TrimNull(ByVal sTmp As String) As String

Dim nNul As Long

'
' Truncate input sTmpg at first Null.
' If no Nulls, perform ordinary Trim.
'
nNul = InStr(sTmp, vbNullChar)
Select Case nNul
    Case Is > 1
    TrimNull = Left(sTmp, nNul - 1)
    Case 1
    TrimNull = ""
    Case 0
    TrimNull = Trim(sTmp)
End Select
End Function

