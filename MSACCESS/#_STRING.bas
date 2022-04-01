Attribute VB_Name = "#_STRING"
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
'                 $$$$$$$$$$F                                       ## ##  ##  ##  ##    ##    ##### ###### #####   ####  ##   ## #####
'                  *$$$$$$$$"                                        ####  #####  ##     ##    ##      ##   ##  ##   ##   ###  ## ##
'                    "***""               _____________                                          ##    ##   ####     ##   ## # ## ##  ###
' STANDARD MODULE WITH DEFAULT FUNCTIONS |v 2017/03/19 |                                          ##   ##   ##  ##   ##   ##  ### ##   ##
' The module contains frequently used functions and is part of the G-VBA library               #####   ##   ##   ## ####  ##   ##  ######
'****************************************************************************************************************************************************
'****************************************************************************************************************************************************
Option Explicit

Public Const UFDELIM As String = "¤"
Public Const EQ As String = "="
Public Const KVDELIM As String = ";"

Private Const CP_UTF8 = 65001
Private Const MOD_NAME As String = "#_STRING"

#If VBA7 Then
  Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long _
    ) As Long
    
    Public Declare PtrSafe Sub Mem_Read2 Lib "msvbvm60" Alias "GetMem2" (ByRef Source As Any, ByRef Destination As Any)
    Public Declare PtrSafe Sub Mem_Copy Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef Source As Any, _
    ByVal Length As Long)
#Else
  Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long _
    ) As Long
#End If


'***************************************

'======================================================================================================================================================
' Glue Text (clear from input text)
' The function reads the text copied earlier to the clipboard, then clears it of line breaks, forming a single paragraph
'======================================================================================================================================================
Public Sub GlueText()

Dim sText As String

    On Error Resume Next
'------------------------
    sText = FromClipboard()
    If sText <> "" Then
        sText = Replace(sText, vbCr, Chr(29))
        sText = Replace(sText, vbLf, Chr(29))
        sText = Replace(sText, Chr(29), " ") & vbCrLf
        If InStr(1, sText, "[") > 0 Then
           sText = Replace(sText, "[", vbCrLf & "[")
        End If
    End If
'------------------------
ExitHere:
    If sText <> "" Then Call ToClipBoard(sText)
End Sub

'======================================================================================================================================================
' Function check if string is null  or ncontains no characters or is only whitespace
'======================================================================================================================================================
Public Function IsBlank(str As String) As Boolean
Dim sRes As String
    
    On Error Resume Next
'---------------------
    sRes = Replace(Replace(Replace(Trim$(str), vbTab, ""), vbCr, ""), vbLf, "")
'---------------------
ExitHere:
    IsBlank = sRes = vbNullString '!!!!!!!!!!!!!!!
End Function

'======================================================================================================================================================
' Get all subset for string-list
'======================================================================================================================================================
Public Function ListSubsets(sList As String, Optional DLM As String = ";", Optional SEP As String = vbCrLf) As String
Dim Arr() As String, nDim As Integer, I As Integer, ArrIndx() As Integer
Dim sRes As String, sWork As String
Dim done As Boolean
Dim OddStep As Boolean

    On Error GoTo ErrHandle
'-----------------------------
If sList = "" Then Exit Function
sWork = SquaredFilter(sList): If sWork = "" Then Exit Function

Arr = Split(sWork, DLM): nDim = UBound(Arr)
ReDim ArrIndx(nDim) 'it starts all 0


    OddStep = True
'---------------------------
    Do Until done        'Add a new subset according to current contents of ArrIndx
        
        sWork = ""
        
        For I = 0 To nDim
            If ArrIndx(I) = 1 Then
                If sWork = "" Then
                    sWork = Arr(I)
                Else
                    sWork = sWork & DLM & Arr(I)
                End If
            End If
        Next I
        
        If sWork <> "" Then sRes = sRes & vbCrLf & Trim(sWork)
        
        If OddStep Then                         'update ArrIndx
            ArrIndx(0) = 1 - ArrIndx(0)         'just flip first bit
        Else
            I = 0                               'first locate first 1
            Do While ArrIndx(I) <> 1
                I = I + 1
            Loop
            
            If I = nDim Then                    'done if i = nDim
                done = True
            Else
                I = I + 1                       'if not done then flip the *next* bit
                ArrIndx(I) = 1 - ArrIndx(I)
            End If
        End If
        OddStep = Not OddStep                   'toggles between even and odd steps
    Loop
    
    If sRes <> "" Then sRes = Right(sRes, Len(sRes) - Len(SEP))
'-----------------------------
ExitHere:
    ListSubsets = sRes     '!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'----------------
ErrHandle:
    ErrPrint2 "ListSubsets", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'======================================================================================================================================================
' Squared Filter
'======================================================================================================================================================
Public Function SquaredFilter(sList As String, Optional DLM As String = ";") As String
Dim Arr() As String, I As Integer, nDim As Integer
Dim sRes As String
     If sList = "" Then Exit Function
     Arr = Split(sList, DLM): nDim = UBound(Arr)
     For I = 0 To nDim
        If Left(Arr(I), 1) <> "[" And Right(Arr(I), 1) <> "]" Then
            sRes = sRes & DLM & Arr(I)
        End If
     Next I
If sRes <> "" Then sRes = Right(sRes, Len(sRes) - Len(DLM))
'------------------------
ExitHere:
    SquaredFilter = sRes '!!!!!!!!!!!!!!
End Function

'======================================================================================================================================================
' Create Generalization Hierarchy Prefix (bPrefix = True) /Suffix Tree for string array
' If bSuppressNotGenralized  Then Remove words without general ancestors
'======================================================================================================================================================
Public Function GeneralizationList(sList As String, Optional bPrefix As Boolean = True, Optional bSuppressNotGenralized As Boolean = True, _
                                                                              Optional DLM As String = ";", Optional SEP As String = vbCrLf) As String
Dim Arr() As String, nDim As Long, I As Long, dict As Object, sRes As String
Dim nWork As Integer, J As Integer, sWork() As String, s As String, key As Variant, sTail As String

    On Error GoTo ErrHandle
'-----------------------------
Set dict = CreateObject("Scripting.Dictionary")

Arr = Split(sList, DLM): nDim = UBound(Arr)
BubbleSort Arr

For I = 0 To nDim                                                  ' CREATE DICTIONARY WITH REDUSED SRINGS
      s = Trim(Arr(I)): If s = "" Then GoTo NextLine
      
      s = StrReduceArr(s, bPrefix, DLM)
      sWork = Split(s, DLM): nWork = UBound(sWork)
      For J = 0 To nWork
          key = sWork(J)
          If Not dict.Exists(key) Then
                 dict.Add key, Arr(I)
          Else
                 dict(key) = dict(key) & DLM & Arr(I)
          End If
      Next J
NextLine:
Next I

For Each key In dict.Keys                                          ' PREPARE RESULT
    If bSuppressNotGenralized Then
          s = dict(key): sWork = Split(s, DLM): nWork = UBound(sWork)
          If nWork > 0 Then sRes = sRes & SEP & key & ": " & s
    Else
          sRes = sRes & SEP & key & ": " & dict(key)
    End If
Next key

If sRes <> "" Then sRes = Right(sRes, Len(sRes) - Len(SEP))
'-----------------------------
ExitHere:
    GeneralizationList = sRes '!!!!!!!!!!!!
    Set dict = Nothing
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "GeneralizationList", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------
' This function returns the portion of the string Text that is to the left of
' TrimChar. If SearchFromRight is omitted or False, the returned string
' is that string to the left of the FIRST occurrence of TrimChar.
'------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function TrimToChar(sText As String, TrimChar As String, Optional bSearchFromRight As Boolean = False) As String
Dim iPos As Integer, sRes As String
    
    If TrimChar = vbNullString Then
        sRes = sText
        GoTo ExitHere
    End If
'-----------------------
    If bSearchFromRight = True Then
        iPos = InStrRev(sText, TrimChar, -1, vbTextCompare)
    Else
        iPos = InStr(1, sText, TrimChar, vbTextCompare)
    End If
    
    If iPos > 0 Then
        sRes = Left(sText, iPos - 1)
    Else
        sRes = sText
    End If
'-----------------------
ExitHere:
    TrimToChar = sRes '!!!!!!!!!!!!!
End Function

'======================================================================================================================================================
' The function remove from string any substring that contains filter array
'======================================================================================================================================================
Public Function FiltreString(str As String, FLTR As Variant) As String
Dim sRes As String, nFLTR As Integer, I As Integer

    On Error GoTo ErrHandle
'---------------------------
sRes = Trim(str)
If sRes = "" Then Exit Function
    nFLTR = UBound(FLTR)
    For I = 0 To nFLTR
            sRes = Trim(Replace(sRes, FLTR(I), ""))
    Next I
'---------------------------
ExitHere:
    FiltreString = sRes '!!!!!!!!!
    Exit Function
'--------------
ErrHandle:
    ErrPrint2 "FiltreString", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'======================================================================================================================================================
' Get String to left from Null Symbol
'======================================================================================================================================================
Public Function TrimToNull(sText As String) As String
Dim iPos As Integer, sRes As String

    iPos = InStr(1, sText, vbNullChar)
    If iPos > 0 Then
        sRes = Left(sText, iPos - 1)
    Else
        sRes = sText
    End If
'-----------------
ExitHere:
    TrimToNull = sRes '!!!!!!!!!!!
End Function

'======================================================================================================================================================
' Parse the hyperlink data and extract the full address
'======================================================================================================================================================
Public Function GetHyperlinkFullAddress(ByVal hyperlinkData As Variant, Optional ByVal removeMailto As Boolean) As Variant

    Const SEPARATOR As String = "#"

    Dim RetVal As Variant
    Dim tmpArr As Variant
    
    If IsNull(hyperlinkData) Then
        RetVal = hyperlinkData
    Else
        
        If InStr(hyperlinkData, SEPARATOR) > 0 Then
            ' I append 4 separators at the end, so I don't have to worry about the
            ' lenght of the array returned by Split()
            hyperlinkData = hyperlinkData & String(4, SEPARATOR)
            tmpArr = Split(hyperlinkData, SEPARATOR)
            
            If Len(tmpArr(1)) > 0 Then
                RetVal = tmpArr(1)
                If Len(tmpArr(2)) > 0 Then
                    RetVal = RetVal & "#" & tmpArr(2)
                End If
            End If
        Else
            RetVal = hyperlinkData
        End If
    
        If Left(RetVal, 7) = "mailto:" Then
            RetVal = Mid(RetVal, 8)
        End If
    
    End If

    GetHyperlinkFullAddress = RetVal

End Function
'======================================================================================================================================================
' Getting Query Body (SQL)
'======================================================================================================================================================
Public Function GetQuerySQL(sQueryName As String) As String
Dim qdf As QueryDef, sRes As String

    On Error GoTo ErrHandle
'------------------------
    If Not IsQuery(sQueryName) Then Exit Function
    
    Set qdf = CurrentDb.QueryDefs(sQueryName)
    sRes = qdf.SQL
    qdf.Close
'------------------------
ExitHere:
    GetQuerySQL = sRes '!!!!!!!!!!!
    Set qdf = Nothing
    Exit Function
'------------
ErrHandle:
    ErrPrint2 "GetQuerySQL", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'======================================================================================================================================================
' String concate with DLM
'======================================================================================================================================================
Public Function ConcateString(str1 As String, str2 As String, Optional DLM As String = ";") As String
Dim sRes As String
    If str1 = "" Then
         sRes = str2
    ElseIf str2 = "" Then
         sRes = str1
    Else
         sRes = str1 & DLM & str2
    End If
'----------------------------
    ConcateString = sRes '!!!!!!!!!!
End Function
'======================================================================================================================================================
' Filtre list string with like (i.e. sFiltre = "*ar*;Go?d;[!a]")
'======================================================================================================================================================
Public Function FiltreList(sList As String, sFiltre As String, Optional DLM As String = ";") As String
Dim LST() As String, nLST As Integer, FLTR() As String, nFLTR As Integer
Dim I As Integer, J As Integer, sRes As String

    On Error GoTo ErrHandle
'--------------------------
    If sList = "" Then Exit Function
    If sFiltre = "" Then GoTo ExitHere
    
    LST = Split(sList, DLM): nLST = UBound(LST)
    FLTR = Split(sFiltre, DLM): nFLTR = UBound(FLTR)
    
    For I = 0 To nLST
        For J = 0 To nFLTR
            If Not LST(I) Like FLTR(J) Then
                 sRes = sRes & DLM & LST(I)
            End If
        Next J
    Next I
If sRes <> "" Then sRes = Right(sRes, Len(sRes) - Len(DLM))
'--------------------------
ExitHere:
    FiltreList = sRes
    Exit Function
'---------
ErrHandle:
    ErrPrint2 "FiltreList", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'======================================================================================================================================================
' This function is of an auxiliary nature and helps to enter a string (character) from the Unicode composition interactively.
' This feature allows you to bypass the limitations of the VBA IDE, which does not support Unicode.
'======================================================================================================================================================
Public Function CharCode(Optional str As String, Optional bShow As Boolean, Optional bFromClipBoard As Boolean) As String
Dim sTest As String, sRes As String
    
    On Error Resume Next
'------------------------------
If str <> "" Then
        sTest = str
Else
    If bFromClipBoard Then
        sTest = FromClipboard()
        If sTest <> "" Then sTest = Left(sTest, 1)
    Else
        sTest = InputBox("Add Char or CharCode", "Cahr Code")
    End If
End If

    If sTest = "" Then Exit Function
    
    If IsNumeric(sTest) Then
        sRes = ChrW(CInt(sTest))
    Else
        sRes = CStr(AscW(sTest))
    End If
'----------------------
ExitHere:
    CharCode = sRes  '!!!!!!!!!!!!!
    If bShow Then Call MsgBoxW("The input is " & sTest & " = " & sRes, vbOKOnly, "Char Code")
End Function
'======================================================================================================================================================
' Append String
'======================================================================================================================================================
Public Function AppendString(sSource As String, sRow As String, Optional DLM As String = vbCrLf, Optional nShift As Long = 0) As String
Dim sRes As String, sShift As String

    On Error Resume Next
'--------------------
   If sRow = "" And sSource = "" Then Exit Function
   
   sShift = String(nShift, " ")
   If sRow = "" Then
      sRes = sShift & sSource
   Else
      sRes = sShift & IIf(sSource <> vbNullString, sSource & DLM & sRow, sRow)
   End If
'--------------------
ExitHere:
    AppendString = sRes '!!!!!!!!!!!!!!
End Function

'=======================================================================================================================================================
' SQL TEXT QUATATION
'=======================================================================================================================================================
Public Function q(sT As String) As String
    If sT <> "" Then
         q = Chr(39) & sT & Chr(39) ' !!!!!!!!!!!!!
    Else
         q = "" '!!!!!!!!!!!!!!!!!!!!!!!!
    End If
End Function
'=======================================================================================================================================================
' GENERAL TEXT QUATATION
'=======================================================================================================================================================
Public Function QR(sT As String) As String
       QR = Chr(34) & sT & Chr(34) ' !!!!!!!!!!!!!
End Function
'=======================================================================================================================================================
' SINGLE QUATATION (EQUIVALENT OF Q)
'=======================================================================================================================================================
Public Function sCH(sSTR As String) As String
   sCH = Chr(39) & sSTR & Chr(39)
End Function
'=======================================================================================================================================================
'GENERAL QUATATION
'=======================================================================================================================================================
Public Function SH(sSTR As String) As String
  SH = Chr(34) & sSTR & Chr(34)   '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'=======================================================================================================================================================
' PLACE ARGUMENT TO SQUARED BARCKETS
'=======================================================================================================================================================
Public Function SHT(sSTR As String) As String
Dim sRes As String
If Left(sSTR, 1) <> "[" Then
      sRes = "[" & sSTR & "]"
Else
      sRes = sSTR
End If
'---------------------
      SHT = sRes   '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'=======================================================================================================================================================
' Count the occurrnces of a substring in a string
'=======================================================================================================================================================
Public Function CountOfSubstring(str As String, substr As String, Optional ByVal Start& = 1, _
                                                                                     Optional Compare As VbCompareMethod = vbBinaryCompare) As Integer
    Dim s2L&
    
    If Compare = vbBinaryCompare Then
        s2L = LenB(substr)
        If s2L Then
            Start = InStrB(Start, str, substr)
            Do While Start
                CountOfSubstring = CountOfSubstring + 1
                Start = InStrB(Start + s2L, str, substr)
            Loop
        End If
    Else
        CountOfSubstring = CountOfSubstring(LCase$(str), LCase$(substr), Start)
    End If
End Function
'=======================================================================================================================================================
' PLACE ARGUMENT TO ROUND BARCKETS
'=======================================================================================================================================================
Public Function SHS(sSTR As String) As String
Dim sRes As String
If Left(sSTR, 1) <> "(" Then
      sRes = "(" & sSTR & ")"
Else
      sRes = sSTR
End If
'---------------------
      SHS = sRes   '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'========================================================================================================================================================
' Truncate the string
'========================================================================================================================================================
Public Function Truncate(src As String, Optional iLimit As Integer = 40, Optional sEndLess As String = "..")
    If src = "" Then Exit Function
    Truncate = Left(Trim(src), iLimit - Len(sEndLess)) & sEndLess
End Function
'========================================================================================================================================================
' Creating a Random String
'========================================================================================================================================================
Public Function GetRandomAlphaString(Optional minLen As Integer = 4, Optional maxLen = 16, _
                                                                                Optional sPrefix As String = "") As String
 Dim FileNameLen As Integer, I As Integer
 Dim sRes As String, iLen As Integer
    '--------------------------------------------------------------
    Const MASAR As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
     Call Randomize: iLen = Len(MASAR)
     FileNameLen = CInt(Int((maxLen * Rnd()) + minLen))    ' длина строки от 4 до 16
         sRes = Mid(MASAR, CInt(Int(((iLen - 10) * Rnd()) + 1)), 1)   ' 1-ой должна быть буква
     For I = 1 To FileNameLen                                         ' Начинаем итерации
         sRes = sRes & Mid(MASAR, CInt(Int((iLen * Rnd()) + 1)), 1)
     Next I
 '--------------------------------------------------------------
         GetRandomAlphaString = sPrefix & LCase(sRes)
 End Function
'=====================================================================================================================================================
' The function Tag/Edging the string. If DLM <> "" (vbCRLF for example) tags write separately
'=====================================================================================================================================================
Public Function StrTag(str As String, Optional sLeftTag As String = "<", Optional sRightTag As String = ">", Optional DLM As String = "") As String
    If str = "" Then Exit Function
    If Left(str, Len(sLeftTag)) = sLeftTag Then Exit Function ' Prevent Duplicate Tagging
    StrTag = sLeftTag & DLM & str & DLM & sRightTag '!!!!!!!!!!!!!!!!!
End Function
'=====================================================================================================================================================
' Getting Random Alpha-Numeric String
' PARAMS:   iNoChars      - No of characters the random string should be in length
'           bNumeric      - Should the random string include Numeric characters
'           bUpperAlpha   - Should the random string include Uppercase Alphabet characters
'           bLowerAlpha   - Should the random string include Lowercase Alphabet characters
'=====================================================================================================================================================
Public Function GenRandomStr(Optional iNoChars As Integer = 12, Optional bNumeric As Boolean = True, _
                                                      Optional bUpperAlpha As Boolean = True, Optional bLowerAlpha As Boolean = True)
       
Dim AllowedChars() As Variant, iEleCounter As Integer
Dim I  As Integer, iRndChar  As Integer, iNoAllowedChars As Integer
     
On Error GoTo ErrHandle
'----------------------------------------------
        ReDim Preserve AllowedChars(0)         'Initialize our array otherwise it throws an error
        AllowedChars(0) = ""
     
        Randomize
     
'-----------------------------------------------
        If bNumeric = True Then                 'Numeric
            For I = 48 To 57                    '48-57
                iEleCounter = UBound(AllowedChars)
                ReDim Preserve AllowedChars(iEleCounter + 1)
                AllowedChars(iEleCounter + 1) = I
            Next I
        End If
'------------------------------------------------
        If bUpperAlpha = True Then              'Uppercase alphabet
            For I = 65 To 90                    '65-90
                ReDim Preserve AllowedChars(UBound(AllowedChars) + 1)
                iEleCounter = UBound(AllowedChars)
                AllowedChars(iEleCounter) = I
            Next I
        End If
'------------------------------------------------
        If bLowerAlpha = True Then              'Lowercase alphabet
            For I = 97 To 122                   '97-122
                ReDim Preserve AllowedChars(UBound(AllowedChars) + 1)
                iEleCounter = UBound(AllowedChars)
                AllowedChars(iEleCounter) = I
            Next I
        End If
'------------------------------------------------
        iNoAllowedChars = UBound(AllowedChars)
        For I = 1 To iNoChars
            iRndChar = Int((iNoAllowedChars * Rnd) + 1)
            GenRandomStr = GenRandomStr & Chr(AllowedChars(iRndChar))
        Next I
'--------------------------------------------------------
ExitHere:
        Exit Function
'--------------------------------------------------------
ErrHandle:
        ErrPrint "GenRandomStr", Err.Number, Err.Description
End Function

'=====================================================================================================================================================
' Функция находит массив всех вхождений заданной подстроки в другую строку
'=====================================================================================================================================================
Public Function GetInStrArray(sSource As String, sTag As String, _
                                                         Optional iCompare As VbCompareMethod = vbTextCompare) As Long()
Dim muRes() As Long, nDim As Integer
Dim iL As Long

iL = 1: nDim = -1: ReDim muRes(0)
'----------------------------------------------------
    Do While iL > 0
       iL = InStr(iL, sSource, sTag, iCompare)
       If iL > 0 Then
           nDim = nDim + 1: ReDim Preserve muRes(nDim)
           muRes(nDim) = iL: iL = iL + Len(sTag)
           If iL > Len(sSource) Then GoTo ExitHere
       End If
    Loop
'-----------------------------------------------
ExitHere:
    GetInStrArray = muRes '!!!!!!!!!!!!!!!
End Function
'======================================================================================================================================================
' Function Get Count of character in string
'======================================================================================================================================================
Public Function StringCountOccurrences(strText As String, strFind As String, _
                                Optional lngCompare As VbCompareMethod) As Long
Dim lngPos As Long
Dim lngTemp As Long
Dim lngCount As Long
    If Len(strText) = 0 Then Exit Function
    If Len(strFind) = 0 Then Exit Function
    lngPos = 1
    Do
        lngPos = InStr(lngPos, strText, strFind, lngCompare)
        lngTemp = lngPos
        If lngPos > 0 Then
            lngCount = lngCount + 1
            lngPos = lngPos + Len(strFind)
        End If
    Loop Until lngPos = 0
    StringCountOccurrences = lngCount
End Function

'======================================================================================================================================================
' Get Levenshtein distance between two strings (Iterative with full matrix/ Wagner–Fischer algorithm)
'======================================================================================================================================================
Public Function Levenshtein(s1 As String, s2 As String) As Integer
Dim n As Integer, m As Integer
Dim D() As Integer, I As Integer, J As Integer
    
    On Error Resume Next
'-----------------------------------
n = Len(s1) + 1: m = Len(s2) + 1
 ReDim D(n, m)
 
If n = 1 Then
        Levenshtein = m - 1
        Exit Function
Else
        If m = 1 Then
            Levenshtein = n - 1
            Exit Function
        End If
End If
 
For I = 1 To n
        D(I, 1) = I - 1
Next I
 
For J = 1 To m
        D(1, J) = J - 1
Next J
 
For I = 2 To n
        For J = 2 To m
            D(I, J) = min(D(I - 1, J) + 1, _
                           D(I, J - 1) + 1, _
                           (D(I - 1, J - 1) - (Mid(s1, I - 1, 1) <> Mid(s2, J - 1, 1))))
        Next J
Next I
'------------------------------
ExitHere:
    Levenshtein = D(n, m)   '!!!!!!!!!!!!!!!!!!!!!!!
End Function
'=====================================================================================================================================================
' This finction check if some word is in list. Returns index of word if success
'=====================================================================================================================================================
Public Function IsInList(sWord As String, sList As String, Optional sSep As String = ",") As Integer
Dim LISTT() As String, nList As Integer, I As Integer, iRes As Integer
    If (sWord = "") Or (sList = "") Then Exit Function
            LISTT = SplitToWords(sList, sSep): nList = UBound(LISTT): iRes = 0
            For I = 0 To nList
                If UCase(sWord) = UCase(LISTT(I)) Then
                    iRes = I + 1: Exit For
                End If
            Next I
    '---------------------------------------------------
ExitHere:
            IsInList = iRes '!!!!!!!!!!!!!!!!!!!!!!
    End Function
'===========================================================================================================================================
' Check if words is presented in this sentence as separate unit
'===========================================================================================================================================
Public Function IsWordInString(sWord As String, sRow As String, Optional SEPS As String = ":;,.=+-()[]{}/\|!?<>") As Boolean
Dim iL As Integer, dSeps As String, sLeft As String, sRight As String, iR As Integer
Dim zWord As String, zRow As String

On Error Resume Next
'------------------------------------------------
zWord = Trim(sWord): zRow = Trim(sRow)
If (zWord = "") Or (zRow = "") Then Exit Function

dSeps = SEPS & " " & vbTab & vbCrLf
iL = InStr(1, zRow, zWord, vbTextCompare): If iL = 0 Then Exit Function

If iL > 1 Then
   sLeft = Mid(zRow, iL - 1, 1)
   If InStr(1, dSeps, sLeft) = 0 Then Exit Function
End If
   iR = iL + Len(zWord) - 1
If iR < Len(sRow) Then
   sRight = Mid(zRow, iR + 1, 1)
   If InStr(1, dSeps, sRight) = 0 Then Exit Function
End If
'------------------------------------------------
ExitHere:
   IsWordInString = True '!!!!!!!!!!!!!!
End Function
'===========================================================================================================================================
' Function split the sentence by separat words an return array with word and its position
'===========================================================================================================================================
Public Function SeparateWords(sSentence As String, Optional DLM As String = ";", Optional SEPS As String = ":;,.=+-()[]{}/\|!?<>") As String()
Dim sRes() As String, nDim As Integer, iSentenceLen As Integer, I As Integer
Dim sWork As String, sTry As String, bWord As Boolean, iWord As Integer
    sTry = SEPS & vbTab & " " & vbCrLf: nDim = -1

iSentenceLen = Len(sSentence)
    For I = 1 To iSentenceLen
         sWork = Mid$(sSentence, I, 1)
         
         If InStr(1, sTry, sWork) > 0 Then       '    SEP
                 bWord = False
         Else                                    '    NOT SEP
                If Not bWord Then
                     nDim = nDim + 1: ReDim Preserve sRes(nDim)
                     bWord = True: sRes(nDim) = I & DLM
                End If
                sRes(nDim) = sRes(nDim) & sWork
         End If
    Next I
'------------------------------
ExitHere:
   SeparateWords = sRes  '!!!!!!!!!!!!!!!!!!
End Function
'=====================================================================================================================================================
' This finction split any string on separaete words
'=====================================================================================================================================================
Public Function SplitToWords(ByRef Text As String, Optional DelimChars As String = ",;:><=+-()[]'?!~/\|") As String()
Dim DelimLen As Long, iDelim As Long
Dim strTemp As String, ThisDelim As String
    
    
    If Text = "" Then Exit Function
    '---------------------------------------
    ' GENERAL SEPARATORS
    strTemp = Trim(Text): strTemp = Replace(strTemp, "  ", " "): strTemp = Replace(strTemp, " ", UFDELIM)
    strTemp = Replace(strTemp, vbTab, UFDELIM): strTemp = Replace(strTemp, vbCrLf, UFDELIM)
    strTemp = Replace(strTemp, vbNewLine, UFDELIM): strTemp = Replace(strTemp, Chr(10), UFDELIM)
    '------------------------------------------------------------
        DelimLen = Len(DelimChars)
        For iDelim = 1 To DelimLen
            ThisDelim = Mid$(DelimChars, iDelim, 1)
            If InStr(strTemp, ThisDelim) <> 0 Then _
                strTemp = Replace(strTemp, ThisDelim, UFDELIM)
        Next iDelim
    '-------------------------------------------------------------
        SplitToWords = Split(strTemp, UFDELIM)  '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'=====================================================================================================================================================
' Check if obkect in List with Patterns
'=====================================================================================================================================================
Public Function IsWordInList(sWord As String, sList As String, Optional DLM As String = ";") As Boolean
Dim LST() As String, nLST As Integer, I As Integer
Dim bRes As Boolean
'---------------------------------------
    If sWord = "" Then Exit Function
    If sList = "" Then Exit Function
    LST = Split(sList, DLM): nLST = UBound(LST)
    For I = 0 To nLST
      If UCase(sWord) Like UCase(LST(I)) Then
        bRes = True: GoTo ExitHere
      End If
    Next I
'---------------------------------------
ExitHere:
   IsWordInList = bRes '!!!!!!!!!!
End Function
'======================================================================================================================================================
' Join To Lists without repeated words
'======================================================================================================================================================
Public Function JoinLists(sList1 As String, sList2 As String, Optional DLM As String = ";") As String
Dim A1() As String, A2() As String, n1 As Integer, n2 As Integer, I As Integer, J As Integer
Dim sRes As String, sWork As String

'------------------------
If sList1 = "" Then
   sRes = sList2: GoTo ExitHere
End If

If sList2 = "" Then
   sRes = sList1: GoTo ExitHere
End If

    A1 = Split(sList1, DLM): n1 = UBound(A1)
    A2 = Split(sList2, DLM): n2 = UBound(A2)
    
    For J = 0 To n2
        For I = 0 To n1
            If A1(I) = A2(J) Then
               sWork = ""
               Exit For
            Else
                sWork = A2(J)
            End If
        Next I
        
       If sWork <> "" Then
          sRes = sRes & DLM & sWork
          sWork = ""
       End If
    
    Next J
    
    sRes = sList1 & sRes
'------------------------
ExitHere:
    JoinLists = sRes '!!!!!!!!!!!
End Function

'=====================================================================================================================================================
' Build String List
'=====================================================================================================================================================
Public Function BuildList(sArg1 As String, Optional sArg2 As String, Optional sArg3 As String, Optional sArg4 As String, Optional sArg5 As String, _
                                                        Optional sArg6 As String, Optional sArg7 As String, Optional DLM As String = ";") As String
Dim sRes As String

On Error Resume Next
'-----------------------------
If sArg1 <> "" Then sRes = sArg1
If sArg2 <> "" Then sRes = IIf(sRes <> "", sRes & DLM & sArg2, sArg2)
If sArg3 <> "" Then sRes = IIf(sRes <> "", sRes & DLM & sArg3, sArg3)
If sArg4 <> "" Then sRes = IIf(sRes <> "", sRes & DLM & sArg4, sArg4)
If sArg5 <> "" Then sRes = IIf(sRes <> "", sRes & DLM & sArg5, sArg5)
If sArg6 <> "" Then sRes = IIf(sRes <> "", sRes & DLM & sArg6, sArg6)
If sArg7 <> "" Then sRes = IIf(sRes <> "", sRes & DLM & sArg7, sArg7)
'----------------------------
ExitHere:
    BuildList = sRes '!!!!!!!!!!!!!!!!!
End Function
'=====================================================================================================================================================
' Translit from Russian to English (Require Option Compare Binary in the module level )
'http://www.utf8-chartable.de/unicode-utf8-table.pl?start=1024&utf8=-&unicodeinhtml=dec
'use numerical HTML column
'=====================================================================================================================================================
Public Function Translit(Russian As String) As String
Dim dict As Object, simb As String, sRes As String, val As String
Dim letters As Variant, Letter As Variant, I As Integer

On Error GoTo ErrHandle
'---------------------------------------
Set dict = CreateObject("Scripting.Dictionary"): dict.CompareMode = vbBinaryCompare
letters = Array("A", "B", "V", "G", "D", "E", "YO", "ZH", "Z", "I", "Y", "K", "L", _
        "M", "N", "O", "P", "R", "S", "T", "U", "F", "KH", "TZ", "CH", "SH", "SCH", "", "Y", _
        "", "E", "YU", "YA", "a", "b", "v", "g", "d", "e", "yo", "zh", "z", "i", "y", "k", "l", "m", _
        "n", "o", "p", "r", "s", "t", "u", "f", "h", "tz", "ch", "sh", "sch", "", "y", "", "e", "yu", "ya", "#")
    I = 1040
For Each Letter In letters
        Select Case Letter
            Case "YO"
                If StrComp(Letter, "YO", vbBinaryCompare) = 0 Then
                     val = ChrW(1025)
                Else
                     val = ChrW(1105)
                End If
            Case "#"
                val = ChrW(8470)
            Case Else
                val = ChrW(I)
                I = I + 1
        End Select
        dict.Add val, Letter
Next Letter
'-----------------------------------------------------------------
For I = 1 To Len(Russian)
        simb = Mid(Russian, I, 1)
        If dict.Exists(simb) Then simb = dict(simb)
        sRes = sRes & simb
Next I
'----------------------------
ExitHere:
    Translit = sRes '!!!!!!!!!!
    Set dict = Nothing
    Exit Function
'---------------
ErrHandle:
    ErrPrint "Translit", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'=====================================================================================================================================================
' Split string by number of characters
'=====================================================================================================================================================
Public Function SplitString(ByVal TheString As String, ByVal StringLen As Integer) As String()
Dim ArrCount As Integer
Dim I As Long
Dim TempArray() As String
  ReDim TempArray((Len(TheString) - 1) \ StringLen)
  For I = 1 To Len(TheString) Step StringLen
    TempArray(ArrCount) = Mid$(TheString, I, StringLen)
    ArrCount = ArrCount + 1
  Next
'----------------------------------------------------
  SplitString = TempArray   '!!!!!!!!!!!!!!!
End Function
'=====================================================================================================================================================
' Get arg in List by index
'=====================================================================================================================================================
Public Function GetListArgByIndex(sList As String, Optional indx As Integer = 0, Optional DLM As String = ";") As String
Dim sRes As String
Dim ar() As String

On Error Resume Next
'-----------------------------
If sList = "" Then Exit Function
If InStr(1, sList, DLM) <= 0 Then
    If indx = 0 Then               ' Check if Indx = 0
       sRes = sList
    Else
       Exit Function
    End If
Else
    '---------------------
       ar = Split(sList, DLM)
       If indx > UBound(ar) Then Exit Function
       sRes = ar(indx)
End If
'-----------------------------
ExitHere:
     GetListArgByIndex = sRes '!!!!!!!!!!!!!!!
End Function
'=====================================================================================================================================================
' Search word in string
'=====================================================================================================================================================
Public Function IsWordInStr(str As String, Word As String, Optional DelimChars As String = ",;:><=+-()[]'?!~/\\|") As Long
Dim sWork As String, SeLST() As String, nSep As Integer, I As Integer
Const UDLM As String = "¤"

On Error GoTo ErrHandle
'---------------------------------------------
If str = "" Then Exit Function
If Word = "" Then Exit Function

    sWork = Replace(str, vbCrLf, UDLM & UDLM)           ' If it is multiline string replace new line as 2--symbol diveder
    sWork = UDLM & Replace(sWork, " ", UDLM) & UDLM     ' Replacr  " " To UDLM
    SeLST = Split(StrConv(DelimChars, 64), Chr(0)): nSep = UBound(SeLST)
    For I = 0 To nSep
         sWork = Replace(sWork, SeLST(I), UDLM)
    Next I
'---------------------------------------------
ExitHere:
    IsWordInStr = InStr(1, sWork, UDLM & Word & UDLM)  '!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'-------------------------
ErrHandle:
    ErrPrint "IsWordInStr", Err.Number, Err.Description
    Err.Clear
End Function

'==========================================================================================================================================================
' Return first quated text
' Example:  QuatedText("dfg = " & CHR(34) & "bvbnvb" & CHR(34))
'==========================================================================================================================================================
Public Function QuatedText(str As String) As String
Dim SS() As String

On Error Resume Next
'------------------------------------
    If InStr(1, str, Chr(34)) = 0 Then Exit Function
    SS = TextBetweenTags(str, "(['" & Chr(34) & "])(?:(?!\1|\\).|\\.)*\1")
    QuatedText = SS(0)
End Function
'==========================================================================================================================================================
' Split Row by position
'==========================================================================================================================================================
Public Function SplitByPosition(sRow As String, iPos As Integer) As String()
Dim sRes(1) As String, iL As Integer
'------------------------
    If sRow = "" Then GoTo ExitHere
    
    iL = Len(sRow)
    
    If iPos > iL Then
        sRes(0) = sRow
    ElseIf iPos < 1 Then
        sRes(1) = sRow
    Else
        sRes(0) = Left(sRow, iPos): sRes(1) = Right(sRow, iL - iPos)
    End If
'------------------------
ExitHere:
    SplitByPosition = sRes '!!!!!!!!!!!!!!!
End Function
'==========================================================================================================================================================
' Get Word From List by Position
'==========================================================================================================================================================
Public Function GetWordFromList(sList As String, Optional indx As Long = 0, Optional DLM As String = ";") As String
Dim Arr() As String, nDim As Long
Dim sRes As String

    On Error Resume Next
'--------------------------
    If sList = "" Then Exit Function
    Arr = Split(sList, DLM): nDim = UBound(Arr): If indx > nDim Then Exit Function
    sRes = Trim(Arr(indx))
'--------------------------
ExitHere:
    GetWordFromList = sRes '!!!!!!!!!!!!!!
End Function
'==========================================================================================================================================================
' The Function Split  the sentence by given word
'==========================================================================================================================================================
Public Function SplitByWord(sRow As String, sWord As String, Optional DLM As String = UFDELIM, _
                                                                                     Optional CompareMode As VbCompareMethod = vbTextCompare) As String
Dim iL As Integer, sRes As String, sWork As String

On Error GoTo ErrHandle
'-----------------------------------
iL = InStr(1, sRow, sWord, CompareMode): If iL = 0 Then Exit Function
sWork = sRow
'------------------------
Do While iL > 0
      If CheckSeparateWord(iL, sWord, sWork) Then
            sRes = IIf(sRes <> "", sRes & DLM, "") & Left(sWork, iL - 1)
            sWork = Trim(Right(sWork, Len(sWork) - iL - Len(sWord) + 1))
      End If
      
      iL = InStr(1, sWork, sWord, CompareMode)
Loop
    If sWork <> "" Then sRes = sRes & DLM & sWork
'------------------------
ExitHere:
    SplitByWord = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'----------------
ErrHandle:
    ErrPrint "SplitByWord", Err.Number, Err.Description
    Err.Clear
End Function

'==========================================================================================================================================================
' The Function look for separate word position in the source
'==========================================================================================================================================================
Public Function GetWordPosition(sRow As String, sWord As String, Optional iStartPosition As Integer = 1, _
                                                                                      Optional CompareMode As VbCompareMethod = vbBinaryCompare) As Integer
Dim iL As Integer, iRes As Integer

On Error GoTo ErrHandle
'--------------------------
      iL = InStr(iStartPosition, sRow, sWord, CompareMode)
      Do While iL > 0
           If CheckSeparateWord(iL, sWord, sRow) Then
               iRes = iL: Exit Do
           End If
           iL = InStr(iL + 1, sRow, sWord, CompareMode)
      Loop
'--------------------------
ExitHere:
     GetWordPosition = iRes '!!!!!!!!!!!!
     Exit Function
'-----------------
ErrHandle:
     ErrPrint "GetWordPosition", Err.Number, Err.Description
     Err.Clear
End Function
'==========================================================================================================================================================
' Function cut some substring fom iStart > 1 and  iEnd < Len(sRow)
'==========================================================================================================================================================
Public Function CutSubString(sRow As String, iStart As Integer, iEnd As Integer) As String
Dim sRes As String, iLen As Integer

'-----------------------------------
    If sRow = "" Then Exit Function:   If iStart < 1 Then Exit Function
    iLen = Len(sRow): If iEnd > iLen Then iEnd = iLen: If iEnd <= iStart Then Exit Function
    
    If iStart > 1 Then sRes = Left(sRow, iStart - 1)
    If iEnd < iLen Then sRes = sRes & Right(sRow, iLen - iEnd)
'-----------------------------------
    CutSubString = sRes '!!!!!!!!!!!!
End Function
'==========================================================================================================================================================
' Function get first word for sentence
'==========================================================================================================================================================
Public Function FirstWord(sRow As String) As String
Dim sChar As String, iFirstWord As Integer, iEndFirstWord As Integer
'----------------------
    sChar = Left(sRow, 1)
    If IsSeparator(sChar) Then
        iFirstWord = GetNextWordPeriod(sRow, 1)
        If iFirstWord >= Len(sRow) Then Exit Function
    Else
        iFirstWord = 1
    End If
    iEndFirstWord = GetNextWordPeriod(sRow, iFirstWord)
    If iEndFirstWord = 0 Then Exit Function
'----------------------
    FirstWord = Mid(sRow, iFirstWord, iEndFirstWord - iFirstWord) '!!!!!!!!!!!!!!!!!
End Function

'==========================================================================================================================================================
' Function get next word after given
'==========================================================================================================================================================
Public Function NextWord(sRow As String, sWord As String, iWord As Integer) As String
Dim sRes As String
Dim iSepBeginning As Integer, iNextWord As Integer, iEndNextWord As Integer

On Error GoTo ErrHandle
'-------------------------
    iSepBeginning = iWord + Len(sWord)
    iNextWord = GetNextWordPeriod(sRow, iSepBeginning): If iNextWord = Len(sRow) Then Exit Function
    iEndNextWord = GetNextWordPeriod(sRow, iNextWord): If iEndNextWord = 0 Then Exit Function
    
    sRes = Mid(sRow, iNextWord, iEndNextWord - iNextWord)
'-------------------------
ExitHere:
    NextWord = sRes '!!!!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint "NextWord", Err.Number, Err.Description
    Exit Function
End Function
'----------------------------------------------------------------------------------------------------------------------------------------------------------
' Function extract period (separators vs words ) in sentence from iPosition till next period
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetNextWordPeriod(sRow As String, iPosition) As Integer
Dim I As Integer, sChar As String, nL As Integer, iResult As Integer
Dim bSep As Boolean

    If sRow = "" Then Exit Function
    nL = Len(sRow): If nL - iPosition <= 0 Then Exit Function
    sChar = Mid(sRow, iPosition, 1): bSep = IsSeparator(sChar)
    
    For I = iPosition + 1 To nL
        sChar = Mid(sRow, I, 1):
        If IsSeparator(sChar) <> bSep Then
              iResult = I: Exit For
        End If
    Next I
'-----------------------------
If iResult = 0 Then iResult = nL + 1 '(The end of sentence)
'-----------------------------
    GetNextWordPeriod = iResult '!!!!!!!!!!!!!!!
End Function
'==========================================================================================================================================================
' Split Text For Parens: ROW --> Before_LeftTag, BetweenTags, After_RightTag
'==========================================================================================================================================================
Public Function SplitParen(sRow As String, LeftTag As String, RightTag As String, Optional iStart As Integer = 1, _
                                                                                     Optional CompareMode As VbCompareMethod = vbBinaryCompare) As String()
Dim sOUT(2) As String
Dim iL As Integer, iR As Integer
'-----------------------------
     If sRow = "" Then GoTo ExitHere:
     iL = GetTagPosition(sRow, LeftTag, iStart, CompareMode): iR = GetTagPosition(sRow, RightTag, iL + 1, CompareMode)
     If iL > 0 And iR > 0 Then
            sOUT(0) = Left(sRow, iL - 1)
            sOUT(1) = Mid(sRow, iL + Len(LeftTag), iR - iL - Len(LeftTag))
            sOUT(2) = Right(sRow, Len(sRow) - iR - Len(RightTag) + 1)
     End If
'-----------------------------
ExitHere:
     SplitParen = sOUT '!!!!!!!!!!!!!!!!!!!!!!
End Function
'----------------------------------------------------------------------------------------------------------------------------------------------------------
' Get Tag Position, correct work with Empty Tag
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetTagPosition(sRow As String, sTag As String, Optional iStart As Integer = 1, _
                                                                                      Optional CompareMode As VbCompareMethod = vbBinaryCompare) As Integer
      If sTag = "" Then Exit Function
      GetTagPosition = InStr(iStart, sRow, sTag, CompareMode) '!!!!!!!!!!!!!!!!!!!!!!!
End Function
'==========================================================================================================================================================
' Get Text between 2 words. If first word is not included, then it extracts text from beginning to second word. If the second word is not included
' then it extracts text from first word to the end of string. If both words are not included into sentence then it returns empty string
'==========================================================================================================================================================
Public Function TextBetweenTwoWords(sRow As String, FirstWord As String, SecondWord As String, Optional iStartPosition As Integer = 1, _
                                             Optional bSeparateWord As Boolean = True, Optional CompareMode As VbCompareMethod = vbBinaryCompare) As String
Dim sRes As String, sWork As String
Dim iL As Integer, iR As Integer

On Error GoTo ErrHandle
'-------------------------
sWork = Trim(sRow): If sWork = "" Then Exit Function
If bSeparateWord Then
   iL = GetWordPosition(sWork, FirstWord, iStartPosition, CompareMode)
Else
   iL = InStr(iStartPosition, sWork, FirstWord, CompareMode)
End If
   If iL <= 0 Then Exit Function
   iL = iL + Len(FirstWord)

If bSeparateWord Then
   iR = GetWordPosition(sWork, SecondWord, iL + 1, CompareMode)
Else
   iR = InStr(iL + 1, sWork, SecondWord, CompareMode)
End If
   If iR <= 0 Then Exit Function
   
   sRes = Mid(sWork, iL, iR - iL)

'-------------------------
ExitHere:
    TextBetweenTwoWords = Trim(sRes) '!!!!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint "TextBetweenTwoWords", Err.Number, Err.Description
    Exit Function
End Function
'==========================================================================================================================================================
' GET TEXT BETWEEN TAGS
' '"<FONT xxx>Value</FONT> <FONT yyy>zzz</FONT>" << - "<FONT([^>]*)>([^(</FONT>)]*)</FONT>"
' "b" -> b  ===> '(?<=")[^"]+(?=")'
' http://www.xyz.com, www.xyz.com, http://www.abc.com/category -> xyz.com   << - "http://|[:/].*"
' <b>TEXT</b> -> TEXT ==> "<(?'tag'\w+?).*>" + "(?'text'.*?)" + "</\k'tag'>", TФП = b
'"38c6v5hrk[x]537fhvvb" -> x ===>  "(\[x\])|(\[\d*\])"
' dddj{ddfht},m,x{fdgg}d{wwmn} --> ([^{]*?)\w(?=\})
' https://www.experts-exchange.com/articles/1336/Using-Regular-Expressions-in-Visual-Basic-for-Applications-and-Visual-Basic-6.html
'==========================================================================================================================================================
Public Function TextBetweenTags(sSource As String, Optional sPattern As String = "([^{]*?)\w(?=\})") As String()
Dim Res() As String, nDim As Integer, I As Integer
Dim RegX As Object, mC As Object
                                                     
On Error GoTo ErrHandle
'-------------------------------------------------------
Set RegX = CreateObject("VBScript.RegExp")
    With RegX
        .Global = True
        .IgnoreCase = True
        .MultiLine = False
        .Pattern = sPattern
        
        Set mC = .Execute(sSource)
   End With

nDim = mC.Count - 1: ReDim Res(nDim)
For I = 0 To nDim
     Res(I) = mC(I).value
Next I
'--------------------------------
ExitHere:
     TextBetweenTags = Res '!!!!!!!!!!!!!!!!!!!
     Set mC = Nothing: Set RegX = Nothing
     Exit Function
'-------------
ErrHandle:
     ErrPrint "TextBetweenTags", Err.Number, Err.Description
     Err.Clear: Resume ExitHere
End Function

'=====================================================================================================================================================
' Get Sunctring Beetween Tags
'=====================================================================================================================================================
Public Function GetSubstringLR(sST As String, iStart As Integer, LeftTag As String, RightTag As String) As String
Dim iLeft As String, iRight As String
    iLeft = InStr(iStart, sST, LeftTag): If iLeft = 0 Then Exit Function
    iRight = InStr(iLeft + 1, sST, RightTag): If iRight = 0 Then Exit Function
   
    GetSubstringLR = Mid$(sST, iLeft + Len(LeftTag), iRight - iLeft - Len(LeftTag) - 1) '!!!!!!!!!!!!!!!!!
End Function

'=========================================================================================================================================================
' Function return Value for KV String
'=========================================================================================================================================================
Public Function GetValueForKey(sKV As String, sKey As String, Optional DLM As String = ";", Optional SEQV As String = "=") As String

Dim KV() As String, nDim As Integer, I As Integer
Dim Pair() As String, sRes As String

On Error GoTo ErrHandle
'------------------------
    If sKV = "" Then Exit Function
    KV = Split(sKV, DLM): nDim = UBound(KV)
        For I = 0 To nDim
            If InStr(1, KV(I), SEQV) > 0 Then
                  Pair = Split(KV(I), SEQV)
                  If UCase(Pair(0)) = UCase(sKey) Then
                        sRes = Trim(Pair(1))
                        Exit For
                  End If
            ElseIf InStr(1, KV(I), sKey) > 0 Then
                        sRes = "TRUE"
                        Exit For
            End If
        Next I
'---------------------------------------
ExitHere:
     GetValueForKey = sRes '!!!!!!!!!!!
     Exit Function
'------------
ErrHandle:
     ErrPrint "GetValueForKey", Err.Number, Err.Description
     Err.Clear: Resume ExitHere
    End Function
    
'====================================================================================================================================================================
' Check is String Match Patter
'====================================================================================================================================================================
Public Function IsRegexMatch(ByRef sText As String, ByVal sPattern As String) As Boolean
Dim RegEx As Object, Matches As Object, bRes As Boolean

On Error GoTo ErrHandle
'--------------------------------
    IsRegexMatch = False

    Set RegEx = CreateObject("vbscript.regexp")

    RegEx.IgnoreCase = True
    RegEx.Global = True
    RegEx.Pattern = sPattern

    Set Matches = RegEx.Execute(sText)
    If Matches.Count = 1 Then bRes = True
'---------------------------------
ExitHere:
    IsRegexMatch = bRes '!!!!!!!!!!!!
    Set Matches = Nothing: Set RegEx = Nothing
    Exit Function
'----------------
ErrHandle:
    ErrPrint "IsRegMatch", Err.Number, Err.Description
    Err.Clear
End Function
'====================================================================================================================================================================
' Check if URL Correct
'====================================================================================================================================================================
Public Function IsURL(ByRef sURL As String) As Boolean
Dim sPattern As String
    IsURL = False

    sPattern = "^(?:http(s)?:\/\/)?[\w.-]+(?:\.[\w\.-]+)+[\w\-\._~:/?#[\]@!\$&'\(\)\*\+,;=.]+$"
    
    IsURL = IsRegexMatch(sURL, sPattern)
'(\.[\w\d\-])* => (\.[\w\d\-]+)*
End Function
'====================================================================================================================================================================
' Check if email is correct
'====================================================================================================================================================================
Public Function IsEmail(sEMAIL As String) As Boolean
Dim sPattern As String
    sPattern = "[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}"
    IsEmail = IsRegexMatch(sEMAIL, sPattern)
End Function

'=====================================================================================================================================================
' Is FileName Correct
'=====================================================================================================================================================
Public Function IsValidFileName(ByVal TheFileName As String) As Boolean
Dim RegEx As Object
    Set RegEx = CreateObject("vbscript.regexp")
    RegEx.Pattern = "\A(?!(?:COM[0-9]|CON|LPT[0-9]|NUL|PRN|AUX|com[0-9]|con|lpt[0-9]|nul|prn|aux)|[\s\.])[^\\\/:*" & Chr(34) & "?<>|]{1,254}\z"
    
    IsValidFileName = Not RegEx.Test(TheFileName)

Set RegEx = Nothing

End Function
'====================================================================================================================================================================
' Check Is String - in Unicode
'====================================================================================================================================================================
Public Function IsUnicode(Text As String) As Boolean
    Dim s As String
    
    On Error Resume Next
'----------------------
    s = RemoveUnicode(Text)
'----------------------
ExitHere:
    IsUnicode = Len(s) <> Len(Text)
End Function


'====================================================================================================================================================================
' Check if URL Correct
'====================================================================================================================================================================
Public Function IsValidURL(ByRef sURL As String) As Boolean
Dim sPattern As String
    IsValidURL = False
    sPattern = "^(?:http(s)?:\/\/)?[\w.-]+(?:\.[\w\.-]+)+[\w\-\._~:/?#[\]@!\$&'\(\)\*\+,;=.]+$"
    
    IsValidURL = IsRegexMatch(sURL, sPattern)
'(\.[\w\d\-])* => (\.[\w\d\-]+)*
End Function



'====================================================================================================================================================================
' Extract Domain From URL
'====================================================================================================================================================================
Public Function ExtractDomain(ByVal url As String) As String
  
On Error Resume Next
'-------------------------------------------------------------
  If InStr(url, "//") Then url = Mid(url, InStr(url, "//") + 2)
  If Left(url, 4) Like "[Ww][Ww][Ww0-9]." Then url = Mid(url, 5)
  
  ExtractDomain = Split(url, "/")(0)
End Function
'======================================================================================================================================================
' Calculate If Word In List
'======================================================================================================================================================
Public Function InList(sList As String, sWord As String, Optional sDelim As String = ";", _
                                                                     Optional iCompare As VbCompareMethod = vbTextCompare) As Boolean
Dim bRes As Boolean
Dim MyArr() As String, nDim As Integer, I As Integer
Dim sWork As String

sWork = Trim(sWord)
'-----------------------------------
If sList = "" Or sWork = "" Then Exit Function
If InStr(1, sList, sWork, iCompare) = 0 Then Exit Function
If InStr(1, sList, sDelim, iCompare) = 0 Then
    If StrComp(sList, sWork, iCompare) = 0 Then
        bRes = True: GoTo ExitHere
    Else
        Exit Function
    End If
End If
'-----------------------------------
MyArr = Split(sList, sDelim): nDim = UBound(MyArr)
    For I = 0 To nDim
        If StrComp(Trim(MyArr(I)), sWord, iCompare) = 0 Then
           bRes = True: GoTo ExitHere
        End If
    Next I
'-------------------------------------------------------------
ExitHere:
    InList = bRes '!!!!!!!!!!!!!!!!
End Function

'=====================================================================================================================================================
' Remove all non-printable symbols:
'      "[^\u0000-\u007F]" - all non -ASC-II
'      "[^\u0000-\u04FF]" - all ASC-II + Russian
'=====================================================================================================================================================
Public Function RemoveUnicode(str As String, Optional CharReplace As String = "", Optional RegTemplate = "[^\u0000-\u04FF]", _
                                                                Optional bIgnoreCase As Boolean = True, Optional bGlobal As Boolean = True) As String
Dim RegStr As Object, sRes As String

On Error GoTo ErrHandle
'-----------------------
    Set RegStr = CreateObject("VBScript.RegExp")
    With RegStr
        .Global = bGlobal
        .Pattern = RegTemplate
        .IgnoreCase = bIgnoreCase
    End With
    
    sRes = RegStr.Replace(str, CharReplace)
'-----------------------
ExitHere:
    RemoveUnicode = sRes '!!!!!!!!!!!!
    Set RegStr = Nothing
    Exit Function
'-------------
ErrHandle:
    ErrPrint "RemoveUnicode", Err.Number, Err.Description
    Err.Clear
End Function


'=====================================================================================================================================================
' Build KV-String
'=====================================================================================================================================================
Public Function BuildKV(src As String, sKey As String, sValue As String, Optional DLM As String = ";", Optional SEQV As String = "=", _
                                                                                                      Optional Replacement As String = "¤") As String
Dim sRes As String, Pair() As String
Dim dict As New cDictionary, ssKey As String, ssVal As String

On Error GoTo ErrHandle
'-----------------------------
  ssKey = CheckKV(sKey, Replacement, DLM, SEQV): ssVal = CheckKV(sValue, Replacement, DLM, SEQV)

  If src = "" Then
       sRes = ssKey & SEQV & ssVal
  Else
       dict.SetKVString src, SEQV, DLM
       If dict.Exists(ssKey) Then
          dict(ssKey) = ssVal
       Else
          dict.Add ssKey, ssVal
       End If
'-----------------------------
       sRes = dict.GetKVString(SEQV, DLM)
  End If

'-----------------------------
ExitHere:
    BuildKV = sRes '!!!!!!!!!!!!
    Set dict = Nothing
    Exit Function
'------------------------
ErrHandle:
    ErrPrint "BuildKV", Err.Number, Err.Description
    Err.Clear: Resume ExitHere
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------
' Check is KV cotains delimeters
'-----------------------------------------------------------------------------------------------------------------------------------------------------
Private Function CheckKV(sCheck As String, Optional Replacement As String = "¤", Optional DLM As String = ";", Optional SEQV As String = "=") As String
      If sCheck = "" Then Exit Function
      CheckKV = Replace(Replace(sCheck, DLM, Replacement), SEQV, Replacement)
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------
' Check is word separate
'------------------------------------------------------------------------------------------------------------------------------------------------
Public Function CheckSeparateWord(iWord As Integer, sWord As String, sRow As String) As Boolean
Dim bLeft As Boolean, bRight As Boolean, TryChar As String, iL As Integer

If iWord = 1 Then
    bLeft = True
Else
    TryChar = Mid(sRow, iWord - 1, 1)
    If IsSeparator(TryChar) Then bLeft = True
End If

iL = iWord + Len(sWord) - 1
If iL >= Len(sRow) Then
    bRight = True
Else
    TryChar = Mid(sRow, iL + 1, 1)
    If IsSeparator(TryChar) Then bRight = True
End If
'-------------------------------------
ExitHere:
    CheckSeparateWord = bLeft And bRight '!!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------
' Check is symbol is separator
'------------------------------------------------------------------------------------------------------------------------------------------------
Private Function IsSeparator(sChar As String) As Boolean
Dim bRes As Boolean, sSeps As String

Const SEPARATOTS As String = "=-+:;',.!?/\[]{}()|\/*^`~"
sSeps = SEPARATOTS & vbTab & vbCrLf
      
      If sChar = " " Then
          bRes = True
      ElseIf Asc(sChar) <= 47 Then
          bRes = True
      ElseIf InStr(1, sSeps, sChar, vbBinaryCompare) > 0 Then
          bRes = True
      End If
'--------------------
      IsSeparator = bRes '!!!!!!!!!!!!!
End Function

'=====================================================================================================================================================
' Convert string to char array
'=====================================================================================================================================================
Public Function CharacterArray(value As String) As String()
    value = StrConv(value, vbUnicode)
    CharacterArray = Split(Left(value, Len(value) - 1), vbNullChar)
End Function

'=====================================================================================================================================================
' Format Phone Number (+# (###) ###-####, +# ### ###-##-##, \(@@@\)\ @@@\-@@@@ )
'=====================================================================================================================================================
Public Function PhoneFormat(ByVal strPhoneNumber As String, Optional phnFormat As String = "+# (###) ###-####", Optional CountryCode As String, _
                                                                                                             Optional RegionCode As String) As String
Dim sPhone As String, sExt As String, sPlus As String

  On Error GoTo ErrHandle
  If strPhoneNumber = "" Then Exit Function
'-----------------------------------------------------------
' CLEAR PHONE --> NUMERIC
  sPhone = PhoneClear(strPhoneNumber, CountryCode, RegionCode) 'Clear phone to numeric grade
  
  If Left(sPhone, 1) = "+" Then
      sPlus = "+": sPhone = Right(sPhone, Len(sPhone) - 1)
  End If
  If Len(sPhone) > 11 Then
         sExt = Right(sPhone, Len(sPhone) - 11): sPhone = Left(sPhone, 11)
  End If
  If Not IsNumeric(sPhone) Then
         sPhone = strPhoneNumber: GoTo ExitHere
  End If
'-----------------------------------------------------------
' NORMILIZE PHONE --> 11 DIGITS
  Select Case Len(sPhone)
    Case 11:                                                  ' WITH COUNTRY CODE  +7 967 030 5810 --> 79670305810 (11 digits);
      If Left(sPhone, 1) = 8 Then                             '                     8 967 030 5810 --> 89670305810 (11 digits)
          sPhone = "7" & Right(sPhone, Len(sPhone) - 1)
      End If
   Case 10:                                                   ' ONLY AREA CODE AND PHONE. 905 827 5942 --> 9058275942 (10 digits)
          If sPlus <> "" Then
            sPhone = strPhoneNumber: GoTo ExitHere
          Else
            sPhone = CountryCode & sPhone
          End If
   Case 7:                                                    ' ONLY PHONE WITHOUT COUNTRY & REGION CODE
          If sPlus <> "" Then
            sPhone = strPhoneNumber: GoTo ExitHere
          Else
            sPhone = CountryCode & RegionCode & sPhone
          End If
   Case Else                                                  ' Can't Recognize Phone Format
          sPhone = strPhoneNumber
          GoTo ExitHere
  End Select
'-----------------------------------------------------------
' FORMAT PHONE
    sPhone = Trim(Format(sPhone, phnFormat) & Space(1) & sExt)
'---------------------------
ExitHere:
    PhoneFormat = sPhone '!!!!!!!!!!!!!!!!!!!!
    Exit Function
'----------------
ErrHandle:
    ErrPrint "", Err.Number, Err.Description
    Err.Clear
End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Clear Phone from extra symbols
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function PhoneClear(strPhoneNumber As String, Optional CountryCode As String = "+7", Optional RegionCode As String = "495") As String
Dim sPhone As String, iLength As Integer, sPlus As String
Dim I As Integer
  If strPhoneNumber = "" Then Exit Function
'--------------------------------------------------
' Remove any style characters from the user input
  sPhone = Replace(strPhoneNumber, ")", "")
  sPhone = Replace(sPhone, "(", "")
  sPhone = Replace(sPhone, "-", "")
  sPhone = Replace(sPhone, ".", "")
  sPhone = Replace(sPhone, Space(1), "")
        
  iLength = Len(sPhone)
  If Not IsNumeric(sPhone) Then     'convert any letters to numbers
    For I = 1 To iLength
        Mid$(sPhone, I, I) = _
            PhoneLetterToDigit(Mid$(sPhone, I, I))
    Next I
  End If
'----------------
ExitHere:
  PhoneClear = sPhone '!!!!!!!!!!!!!!!
End Function
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Phone Letter to Digit
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function PhoneLetterToDigit(ByVal strPhoneLetter As String) As String
Dim intDigit As Integer
  
  intDigit = Asc(UCase$(strPhoneLetter))
    
  If intDigit >= 65 And intDigit <= 90 Then

    If intDigit = 81 Or 90 Then ' Q or Z
      intDigit = intDigit - 1
    End If

    intDigit = (((intDigit - 65) \ 3) + 2)
    PhoneLetterToDigit = intDigit
  Else
    PhoneLetterToDigit = strPhoneLetter
  End If

End Function



'=====================================================================================================================================================
' Convert STRING UNICODE (UTF-16) To UTF-8
'=====================================================================================================================================================
Public Function UTF16To8(ByVal UTF16 As String) As String
Dim sBuffer As String
Dim lLength As Long
If UTF16 <> "" Then
    #If VBA7 Then
        lLength = WideCharToMultiByte(CP_UTF8, 0, CLngPtr(StrPtr(UTF16)), -1, 0, 0, 0, 0)
    #Else
        lLength = WideCharToMultiByte(CP_UTF8, 0, StrPtr(UTF16), -1, 0, 0, 0, 0)
    #End If
    sBuffer = Space$(lLength)
    #If VBA7 Then
        lLength = WideCharToMultiByte(CP_UTF8, 0, CLngPtr(StrPtr(UTF16)), -1, CLngPtr(StrPtr(sBuffer)), LenB(sBuffer), 0, 0)
    #Else
        lLength = WideCharToMultiByte(CP_UTF8, 0, StrPtr(UTF16), -1, StrPtr(sBuffer), LenB(sBuffer), 0, 0)
    #End If
    sBuffer = StrConv(sBuffer, vbUnicode)
    UTF16To8 = Left$(sBuffer, lLength - 1)
Else
    UTF16To8 = ""
End If
End Function

'=====================================================================================================================================================
'  This function depends on the correct code page. For some versions of Windows, Unicode characters will not display correctly in VBA.
'  A short patch for Russian fonts - applying a simple file to the Windows registry
'       1251.reg
'   -------------------------------------------------------------------------
'  |     Windows Registry Editor Version 5.00                                |
'  |                                                                         |
'  |     [HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Nls\CodePage]  |
'  |     "1250"="c_1251.nls"                                                 |
'  |     "1251"="c_1251.nls"                                                 |
'  |     "1252"="c_1251.nls"                                                 |
'  | _________________________________________________________________________|
'=====================================================================================================================================================
Public Function URL_Encode(sText As String, Optional UseADODBStream As Boolean = False) As String
    Dim sRes As String, buffer As String, I As Long, c As Long, n As Long
    
#If Mac Then
    buffer = String$(Len(sText) * 12, "%")
 
    For I = 1 To Len(txt)
        c = AscW(Mid$(txt, I, 1)) And 65535
 
        Select Case c
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95  ' Unescaped 0-9A-Za-z-._ '
                n = n + 1
                Mid$(buffer, n) = ChrW(c)
            Case Is <= 127            ' Escaped UTF-8 1 bytes U+0000 to U+007F '
                n = n + 3
                Mid$(buffer, n - 1) = Right$(Hex$(256 + c), 2)
            Case Is <= 2047           ' Escaped UTF-8 2 bytes U+0080 to U+07FF '
                n = n + 6
                Mid$(buffer, n - 4) = Hex$(192 + (c \ 64))
                Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
            Case 55296 To 57343       ' Escaped UTF-8 4 bytes U+010000 to U+10FFFF '
                I = I + 1
                c = 65536 + (c Mod 1024) * 1024 + (AscW(Mid$(txt, I, 1)) And 1023)
                n = n + 12
                Mid$(buffer, n - 10) = Hex$(240 + (c \ 262144))
                Mid$(buffer, n - 7) = Hex$(128 + ((c \ 4096) Mod 64))
                Mid$(buffer, n - 4) = Hex$(128 + ((c \ 64) Mod 64))
                Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
            Case Else                 ' Escaped UTF-8 3 bytes U+0800 to U+FFFF '
                n = n + 9
                Mid$(buffer, n - 7) = Hex$(224 + (c \ 4096))
                Mid$(buffer, n - 4) = Hex$(128 + ((c \ 64) Mod 64))
                Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
        End Select
    Next
    sRes = Left$(buffer, n)
#Else
     If UseADODBStream Then
          URL_Encode = ADODB_URL_Encode(sText)
     Else
          URL_Encode = URLEncode_API(sText)
     End If
#End If
End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Encode String with ADODBStream
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function ADODB_URL_Encode(StringVal As Variant, Optional SpaceAsPlus As Boolean = False) As String

Dim bytes() As Byte, b As Byte, I As Integer, Space As String
Dim oStream As Object
  
Const adModeReadWrite = 3
Const adTypeText = 2
Const adTypeBinary = 1

    On Error GoTo ErrHandle
'---------------------------
If SpaceAsPlus Then Space = "+" Else Space = "%20"

If Len(StringVal) > 0 Then

    Set oStream = CreateObject("ADODB.Stream")
     
    With oStream
      .mode = adModeReadWrite
      .Type = adTypeText
      .CHARSET = "UTF-8"
      .Open
      .WriteText StringVal
      .Position = 0
      .Type = adTypeBinary
      .Position = 3 ' skip BOM
      bytes = .Read
    End With

    ReDim result(UBound(bytes)) As String

    For I = UBound(bytes) To 0 Step -1
      b = bytes(I)
      Select Case b
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
          result(I) = Chr(b)
        Case 32
          result(I) = Space
        Case 0 To 15
          result(I) = "%0" & Hex(b)
        Case Else
          result(I) = "%" & Hex(b)
      End Select
    Next I

    ADODB_URL_Encode = Join(result, "")
  End If
'------------------------
ExitHere:
    Exit Function
'----------------
ErrHandle:
    ErrPrint2 "ADODB_URL_Encode", Err.Number, Err.Description
    Err.Clear
End Function

'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Encode String with API
'-------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function URLEncode_API(StringVal As String, Optional SpaceAsPlus As Boolean = False, Optional UTF8Encode As Boolean = True) As String

Dim StringValCopy As String: StringValCopy = IIf(UTF8Encode, UTF16To8(StringVal), StringVal)
Dim StringLen As Long: StringLen = Len(StringValCopy)

If StringLen > 0 Then
    ReDim result(StringLen) As String
    Dim I As Long, CharCode As Integer
    Dim Char As String, Space As String

  If SpaceAsPlus Then Space = "+" Else Space = "%20"

  For I = 1 To StringLen
    Char = Mid$(StringValCopy, I, 1)
    CharCode = Asc(Char)
    Select Case CharCode
      Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
        result(I) = Char
      Case 32
        result(I) = Space
      Case 0 To 15
        result(I) = "%0" & Hex(CharCode)
      Case Else
        result(I) = "%" & Hex(CharCode)
    End Select
  Next I
  URLEncode_API = Join(result, "")

End If

End Function
'=====================================================================================================================================================
' Replace Separators and Delimetr
'=====================================================================================================================================================
Public Function ReplaceDLM(str As String, Optional LookFor As String = ";", Optional ReplaceWith As String = "¤") As String
Dim C1 As String, C2 As String
   
   If str = "" Then Exit Function
'----------------------------------
   If IsNumeric(LookFor) Then
       C1 = Chr(LookFor)
   Else
       C1 = LookFor
   End If
   If IsNumeric(ReplaceWith) Then
       C2 = Chr(ReplaceWith)
   Else
       C2 = ReplaceWith
   End If
'----------------------------------
   ReplaceDLM = Replace(str, LookFor, ReplaceWith) '!!!!!!!!!!!!!!!
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
                                                                                                  Optional sModName As String = "#_STRING") As String
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
' Create Reduce List: string -> strin -> stri -> str -> st -> s
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function StrReduceArr(s As String, Optional bPrefix As Boolean = True, Optional DLM As String = ";") As String
Dim I As Integer, iLen As Integer
Dim buff() As String

    On Error Resume Next
'--------------------
If s = "" Then Exit Function
iLen = Len(s)

ReDim buff(iLen - 1): buff(0) = s

For I = 1 To iLen - 1
    If bPrefix Then
        buff(I) = Left(buff(I - 1), Len(buff(I - 1)) - 1)
    Else
        buff(I) = Right(buff(I - 1), Len(buff(I - 1)) - 1)
    End If
Next
'--------------------
ExitHere:
    StrReduceArr = Join(buff, DLM) '!!!!!!!!!!!!!!
End Function

