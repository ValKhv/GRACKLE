VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************
' This class provides two functional mode:
'    (1) CSV (Comma-Separated Values) parser to convert CSV text with some separator to Collection or Array of the CSV table contents
'    (2) CSV-writer as VBA functions. The CSV writer converts 2-dimensional array to CSV text
'  The parsers and writer are compliant with the CSV format defined in RFC4180 (https://www.ietf.org/rfc/rfc4180.txt),
'  which allows commas, line breaks, and double-quotes included in the fields
'************************************************************************************************************************************************************************
'************************************************************************************************************************************************************************
Option Compare Database
Option Explicit


'----- Enum -------------------------------------------------------------------------

' Field Quoting
'   Used for the argument 'quoting' of ConvertArrayToCSV()
'   This argument controls what kind of fields to be quoted
Public Enum CSVUtilsQuote
    MINIMAL = 0     ' quote the fields that requires quotation (i.e., that includes comma, return code, quotation mark)
    all = 1         ' quote all the fields
    NONNUMERIC = 2  ' quote non-numeric (Not IsNumeric()) fields
End Enum

'----- Global variables -------------------------------------------------------------

Private CSVUtilsAnyErrorIsFatal As Boolean  'default False





'
Public Sub SetCSVUtilsAnyErrorIsFatal(ByRef value As Boolean)
  CSVUtilsAnyErrorIsFatal = value
End Sub

'====================================================================================================================================================================
' Parce CSV String To Collection
'   Return a Collection of records; record is a Collection of fields
'   When error, return Nothing
'====================================================================================================================================================================
Public Function ParseCSVToCollection(ByRef csvText As String, Optional ByRef allowVariableNumOfFields As Boolean = False) As Collection
    Err.Clear
    If CSVUtilsAnyErrorIsFatal Then GoTo Head
    On Error Resume Next
Head:
    Dim csvPos As Long
    Dim recordPos As Long
    Dim recordText As String, recordTextComma As String
    Dim fieldText As String
    Dim FIELDS As Collection
    Dim csvCollection As Collection
    
    Set csvCollection = New Collection 'empty collection
    
    Set ParseCSVToCollection = csvCollection
    
    'for empty text
    If csvText = "" Then Exit Function 'return empty collection
    
    'extract records and fields
    csvPos = 1
    Do While GetOneRecord(csvText, csvPos, recordText)
        Set FIELDS = New Collection
        recordPos = 1
        recordTextComma = recordText & ","
        Do While FindNextSeparator(recordTextComma, recordPos, fieldText, ",", "")
            If InStr(fieldText, """") > 0 Then
                fieldText = TrimQuotes(fieldText) 'get internal of double-quotes
                fieldText = Replace(fieldText, """""", """") 'un-escape double quote
                If Left(fieldText, 2) = "=""" And Right(fieldText, 1) = """" Then fieldText = Mid(fieldText, 3, Len(fieldText) - 3) 'remove MS quote (="...")
                'add to collection
                FIELDS.Add typingField(fieldText, True)
            Else
                'add to collection
                FIELDS.Add typingField(fieldText, False)
            End If
        Loop
        csvCollection.Add FIELDS
        
        If Not allowVariableNumOfFields And csvCollection.Item(1).Count <> FIELDS.Count Then
            Err.Raise 10001, "ParseCSVToCollection", "Syntax Error in CSV: numbers of fields are different among records"
            GoTo ErrHandle
        End If
    Loop
    If Err.Number <> 0 Then GoTo ErrHandle
'------------
ExitHere:
    Set ParseCSVToCollection = csvCollection
    Exit Function
'------------
ErrHandle:
    Set ParseCSVToCollection = Nothing
    
End Function

'====================================================================================================================================================================
' Parse CSV text and return 2-dim array
'
'  Return 2-dim array --- String(1 TO recordCount, 1 TO fieldCount)
'  When CSV text is "", return empty array --- String(0 TO -1)
'  When error, return Null
'====================================================================================================================================================================
Public Function ParseCSVToArray(ByRef csvText As String, Optional ByRef allowVariableNumOfFields As Boolean = False) As Variant
    ' "On Error Resume Next" only if CSVUtilsAnyErrorIsFatal is True
    Err.Clear
    If CSVUtilsAnyErrorIsFatal Then GoTo Head
    On Error Resume Next
Head:
    Dim csv As Collection
    Dim recCnt As Long, fldCnt As Long
    Dim csvArray() As String
    Dim ri As Long, fi As Long
    Dim rc As Variant, cc As Variant
    
    ParseCSVToArray = Null 'for error
  
    ' convert CSV text to Collection
    Set csv = ParseCSVToCollection(csvText, allowVariableNumOfFields)
    If csv Is Nothing Then  'error occur
        Exit Function
    End If
    
    ' get size of collections
    recCnt = csv.Count
    If recCnt = 0 Then
        ParseCSVToArray = Split("", "/") 'return empty(zero length) String array of bound 0 TO -1
                                         '(https://msdn.microsoft.com/ja-jp/library/office/gg278528.aspx)
        Exit Function
    End If
    fldCnt = 0
    For ri = 1 To csv.Count
      If fldCnt < csv.Item(ri).Count Then fldCnt = csv.Item(ri).Count
    Next
    
    ' copy collection to array
    ReDim csvArray(1 To recCnt, 1 To fldCnt) As String
    ri = 1
    For Each rc In csv 'for each is faster for Collection
      fi = 1
      For Each cc In rc
        csvArray(ri, fi) = cc
        fi = fi + 1
      Next
      ri = ri + 1
    Next
    
    ParseCSVToArray = csvArray
End Function


'====================================================================================================================================================================
' Convert 2-dim array to CSV text string
'
'  inArray : 2-dim array of arbitary size/range and type.
'  fmtDate : format used for conversion from type Date to type String
'  When error, return ""
'====================================================================================================================================================================
Public Function ConvertArrayToCSV(inArray As Variant, Optional fmtDate As String = "yyyy/m/d", _
                          Optional ByVal quoting As CSVUtilsQuote = CSVUtilsQuote.MINIMAL, _
                          Optional ByVal recordSeparator As String = vbCrLf) As String
    ' "On Error Resume Next" only if CSVUtilsAnyErrorIsFatal is True
    Err.Clear
    If CSVUtilsAnyErrorIsFatal Then GoTo Head
    On Error Resume Next
Head:
    Dim csv As String
    Dim R As Long, c As Long, ub2 As Long
    Dim v As Variant
    Dim cell As String
    Dim arrRecord As Variant, arrField As Variant
    
    'error check
    If Not IsArray(inArray) Then
        Err.Raise 10004, "ConvertArrayToCSV", "Input argument inArray is not array"
        GoTo ErrHandle
    End If
    ub2 = UBound(inArray, 2)
    If Err.Number <> 0 Then 'expecting Err.Number = 9, Err.Description = "Subscript out of range", for inArray is 1-dim
        GoTo ErrHandle
    End If

    Dim rc As Long, cc As Long
    ReDim arrRecord(LBound(inArray, 1) To UBound(inArray, 1)) As String 'temporary array
    ReDim arrField(LBound(inArray, 2) To UBound(inArray, 2)) As String 'temporary array
    
    For R = LBound(inArray, 1) To UBound(inArray, 1)
      For c = LBound(inArray, 2) To UBound(inArray, 2)
        v = inArray(R, c)
        'formatting
        cell = IIf(IsNull(v), "", v)
        If TypeName(v) = "Date" Then cell = Format(v, fmtDate)
        'quote and escape
        If quoting = CSVUtilsQuote.all Or _
           (quoting = CSVUtilsQuote.NONNUMERIC And Not IsNumeric(v)) Or _
           InStr(cell, ",") > 0 Or InStr(cell, """") > 0 Or InStr(cell, vbCr) > 0 Or InStr(cell, vbLf) > 0 Then
          cell = Replace(cell, """", """""")
          cell = """" & cell & """"
        End If
        'add to array
        arrField(c) = cell
      Next
      arrRecord(R) = Join(arrField, ",") & recordSeparator
    Next
    If Err.Number <> 0 Then GoTo ErrHandle 'unexpected error
    
    ConvertArrayToCSV = Join(arrRecord, "")
    Exit Function
ErrHandle:
    ConvertArrayToCSV = ""
End Function


'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function typingField(fieldText As String, quoted As Boolean) As Variant
    typingField = fieldText
End Function
 

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Get the next one record from csvText, and put it into recordText
'     updating csvPos
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetOneRecord(ByRef csvText As String, ByRef csvPos As Long, ByRef recordText As String) As Boolean
  GetOneRecord = FindNextSeparator(csvText, csvPos, recordText, "" & vbCr, "" & vbLf)
  If Not GetOneRecord Then Exit Function
  If Mid(csvText, csvPos - 1, 2) = vbCr & vbLf Then csvPos = csvPos + 1 'for CR+LF
End Function
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Find next separator in inText starting with the position "start"
'   foundText = substring [start, found_separator-1] of inText
'   start = found_separator + 1
'   assume that a virtual separator exists at the end of string if there is no separator there.
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function FindNextSeparator(ByRef inText As String, ByRef Start As Long, ByRef foundText As String, ByRef sep1 As String, Optional ByRef sep2 As String = "") As Boolean
    Dim dQuateCnt As Long
    Dim init_start As Long, lenText As Long, p2 As Long, found As Long
    
    FindNextSeparator = False
    lenText = Len(inText)
    init_start = Start
        
    If Start > lenText Then Exit Function 'over-run
    
    dQuateCnt = 0
    Do While Start <= lenText
        'find next separator
        found = InStr(Start, inText, sep1)
        If sep2 <> "" Then
          p2 = InStr(Start, inText, sep2)
          If p2 <> 0 And (found = 0 Or p2 < found) Then found = p2
        End If
        If found = 0 Then found = lenText + 1 'EOF
                
        dQuateCnt = dQuateCnt + StrCount(inText, """", Start, found - 1) 'number of double quates in inText
        Start = found + 1
        If dQuateCnt Mod 2 = 0 Then  'if the number of double-quates is even, then the separator is not fake
            FindNextSeparator = True
            foundText = Mid(inText, init_start, found - init_start)
            Exit Function
        End If
    Loop
    
    Err.Raise 10002, "ParseCSVToCollection", "Syntax Error in CSV: illegal double-quote code"
End Function



'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
' count the string Target in [p0, p1] of Source
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function StrCount(Source As String, Target As String, p0 As Long, p1 As Long) As Long
    Dim n As Long, cnt As Long
    n = p0 - 1
    cnt = 0
    Do
        n = InStr(n + 1, Source, Target)
        If n = 0 Or n > p1 Then Exit Do
        cnt = cnt + 1
    Loop
    StrCount = cnt
End Function

'---------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Trim all before and after doube-quote
'   * text MUST include two or more double-quotes (")
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function TrimQuotes(ByRef Text As String) As String
    'If InStr(text, """") = 0 Then Err.Raise 9999, "", "program error"
    Dim p0 As Long, p1 As Long
    Dim s As String
    
    'trim tail
    For p1 = Len(Text) To 1 Step -1
      s = Mid(Text, p1, 1)
      If (s = """") Then Exit For
    Next
    'trim head
    For p0 = 1 To p1
      s = Mid(Text, p0, 1)
      If (s = """") Then Exit For
    Next
    'return
    TrimQuotes = Mid(Text, p0 + 1, p1 - p0 - 1)
End Function


'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
    
End Sub
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'-------------------------------------------------------------------------------------------------------------------------------------------------------
' Error Handler
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function ErrPrint(FuncName As String, ErrNumber As Long, ErrDescription As String, Optional bDebug As Boolean = True, _
                                                                                                  Optional sModName As String = "cLink") As String
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


