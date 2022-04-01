Attribute VB_Name = "#_EXCEL"
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
'                 $$$$$$$$$$F                                       ## ##  ##  ##  ##    ##   #####  ##  ##  ####  ###### ##
'                  *$$$$$$$$"                                        ####  #####  ##     ##   ##       ###   ##    ##     ##
'                    "***""               _____________                                       ###      ##    ##    ###    ##
' STANDARD MODULE WITH DEFAULT FUNCTIONS |v 2020/12/01 |                                      ##      ## ##  ##    ##     ##
' The module contains some functions to work with Excel and is part of the G-VBA library      #####  ##   ## ####  #####  ######  #####
'****************************************************************************************************************************************************
'****************************************************************************************************************************************************
Option Explicit

Private Const MOD_NAME As String = "#_EXCEL"
'**************************

Public Sub TestImportExcel()
Dim sPath As String, sSheet As String, iCols As Integer
Dim sRes As String, iRowStart As Integer
    sPath = OpenDialog(GC_OPEN_FILE, "Pick the Excel Jpurnal", , False)
    sSheet = "Sheet1": iCols = 10: iRowStart = 2
    If sPath = "" Then Exit Sub
    'sRes = ImportExcel(sPath, sSheet, iRowStart, iCols)
    sRes = ImportExcel2(sPath, sSheet, iRowStart)
    Debug.Assert False
End Sub
'===================================================================================================================================================
' Import Data from Excel Sheet to String
'===================================================================================================================================================
Public Function ImportExcel(ExcelPath As String, Optional SheetName As String = "1", Optional iRowStart As Integer = 1, _
                             Optional iColumns As Integer = 2, Optional sPassword As String = vbNullString, _
                             Optional bVisible As Boolean = False, Optional DLM As String = "|", Optional SEP As String = vbCrLf) As String
Dim sRes As String, nRows As Long, I As Long, J As Long, iBlankCounter As Integer
Dim xlApp As Object, xlWB As Object, xlSheet As Object, sRow As String, sWork As String


Const iBlankRowLimit As Integer = 2
Const bReadOnly = True

On Error GoTo ErrHandle
'----------------------
    DoCmd.Hourglass True
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = bVisible
    Set xlWB = xlApp.Workbooks.Open(ExcelPath, , bReadOnly, , _
      sPassword)
      
    If IsNumeric(SheetName) Then
        Set xlSheet = xlWB.Worksheets(CInt(SheetName))
    Else
        Set xlSheet = xlWB.Worksheets(SheetName)
    End If
    '------------------
    nRows = xlSheet.ROWS.Count
    
    For I = iRowStart To nRows - 1
        sRow = ""
        For J = 1 To iColumns
             sWork = sWork & xlSheet.ROWS.Cells(I, J) & DLM
        Next J
             sWork = Left(sWork, Len(sWork) - Len(DLM))
           
        If Trim(Replace(sWork, DLM, "")) = "" Then
           iBlankCounter = iBlankCounter + 1
           If iBlankCounter > iBlankRowLimit Then Exit For
        Else
            sRes = sRes & IIf(sRes <> "", SEP, "") & sWork
            sWork = ""
        End If
    Next I
'----------------------
ExitHere:
    ImportExcel = sRes
    DoCmd.Hourglass False
    Set xlSheet = Nothing: Set xlWB = Nothing: Set xlApp = Nothing
    Exit Function
'--------
ErrHandle:
    ErrPrint2 "ImportExcel", Err.Number, Err.Description, MOD_NAME
    Err.Clear: DoCmd.Hourglass False
End Function
'===================================================================================================================================================
' Import Data from Excel Sheet to String-2
'===================================================================================================================================================
Public Function ImportExcel2(ExcelPath As String, Optional SheetName As String = "1", Optional iRowStart As Integer = 1, _
                             Optional DLM As String = "|", Optional SEP As String = vbCrLf) As String

Dim sRes As String, nRows As Long, I As Long, J As Long, iBlankCounter As Integer
Dim SQL As String, RS As DAO.Recordset, sRow As String, nCols As Long

Const ISAM_STRING As String = "Excel 12.0;HDR=YES;IMEX=1"

Const iBlankRowLimit As Integer = 2
'----------------------
SQL = "SELECT T1.* FROM [" & ISAM_STRING & ";Database=" & ExcelPath & "].[" & SheetName & "$A" & iRowStart & ":U65536] AS T1;"

Set RS = CurrentDb.OpenRecordset(SQL)
With RS
    If Not .EOF Then
        nCols = .FIELDS.Count - 1
        For I = 0 To nCols
           sRow = sRow & IIf(sRow <> "", DLM, "") & .FIELDS(I).Name
        Next I
        If sRow <> "" Then sRes = sRow
                        
        .MoveLast: .MoveFirst
        Do While Not .EOF
            sRow = ""
            For I = 0 To nCols
                sRow = sRow & IIf(sRow <> "", DLM, "") & .FIELDS(I).value
            Next I
    
                If Trim(Replace(sRow, DLM, "")) = "" Then
                    iBlankCounter = iBlankCounter + 1
                    If iBlankCounter > iBlankRowLimit Then GoTo ExitHere
                Else
                    sRes = sRes & IIf(sRes <> "", SEP, "") & sRow
                End If
    
            
            .MoveNext
        Loop
    End If
End With
'----------------------
ExitHere:
    ImportExcel2 = sRes '!!!!!!!!!
    Set RS = Nothing
    Exit Function
'--------
ErrHandle:
    ErrPrint2 "ImportExcel2", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function


Public Sub Test_PrintToExcel()
Dim sHeader As String, sData As String, DLM As String, sTitle As String, sColumnFormat As String

DLM = "|"
sHeader = "Header 1 - Number" & DLM & "Header 2 - Title " & DLM & "Header 3 - Long Text" & DLM & "Header 4 - Date" & DLM & "Header 5 - Currency"
sColumnFormat = "0|0|1|2|3"
sData = "1" & DLM & "Title-1" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "2" & DLM & "Title-2" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "3" & DLM & "Title-3" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "4" & DLM & "Title-4" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "5" & DLM & "Title-5" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "6" & DLM & "Title-6" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "7" & DLM & "Title-7" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "8" & DLM & "Title-8" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "9" & DLM & "Title-9" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "10" & DLM & "Title-10" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "11" & DLM & "Title-11" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "12" & DLM & "Title-12" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "13" & DLM & "Title-13" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"

sTitle = "THIS IS AN EXAMPLE."
Debug.Print PrintToExcel(sData, sHeader, , , , "SHHET45", sTitle, 5, sColumnFormat)

End Sub

Public Sub Test_PrintToExcelTemplate()
Dim sHeader As String, sData As String, DLM As String, sTitle As String, sColumnFormat As String, xLPath As String


xLPath = "C:\Users\valer\Template_2.xlsx"
DLM = "|"

sData = "1" & DLM & "Title-1" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "2" & DLM & "Title-2" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "3" & DLM & "Title-3" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "4" & DLM & "Title-4" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "5" & DLM & "Title-5" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "6" & DLM & "Title-6" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "7" & DLM & "Title-7" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "8" & DLM & "Title-8" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "9" & DLM & "Title-9" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "10" & DLM & "Title-10" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "11" & DLM & "Title-11" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "12" & DLM & "Title-12" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"
sData = sData & vbCrLf & "13" & DLM & "Title-13" & _
              DLM & "This is a long text to test Excel output blah blah blah. Is It correct Now? If YES it is perfect" & _
              DLM & "12/03/1967" & DLM & "1500000"

Debug.Print PrintToExcel(sData, , xLPath, False, 3, "Sheet1")

End Sub

'===================================================================================================================================================
' Create Excel and Print some 2-D Array to excel with or without header
' Return path to Excel. Column format is string with the same DLM and numbers∆
'                       0 =  As Is + AutoFit (for all if ColumnFormat is Empty)
'                       1 =  Long Text - Wrap
'                       2 =  Date
'                       3 = Currency, for example "$#,##0.00"
'===================================================================================================================================================
Public Function PrintToExcel(sExport As String, Optional sHeader As String, Optional xLPath As String, Optional bNewFile As Boolean = True, _
                                                                                Optional iStartRow As Long = 1, Optional SheetName As String = "", _
                     Optional sTitle As String, Optional iTitleLen As Integer = 4, Optional ColumnFormat As String, Optional DLM As String = "|", _
           Optional SEP As String = vbCrLf, Optional DateFormat As String = "dd/MM/yyyy", Optional CurrencyFormat As String = "$#,##0.00") As String

Dim xlApp As Object, xlWkBk As Object, bOpenExcel As Boolean, xlSheet As Object
Dim Arr() As String, WORKS() As String, FORMATS() As String, nDim As Long, nRows As Long, I As Long, J As Long, iStart As Long

Const HEADER_SHADOW As Long = 14474460
Const xlCenter As Integer = -4108
Const LEN_TO_WRAP As Integer = 25
Const StandardWidth As Integer = 15            ' Width for General Columns
Const XTWidth As Integer = 40                  ' Width for Long Text Column

  On Error GoTo ErrHandle
'-------------------------------
If sExport = "" Then Exit Function
If Not bNewFile Then
   If Dir(xLPath, vbNormal) = "" Then Err.Raise 10000, , "Can't Find The Existing Excel File: " & xLPath
End If
Arr = Split(sExport, SEP): nRows = UBound(Arr)

If ColumnFormat <> "" Then
   FORMATS = Split(ColumnFormat, DLM)
End If
'-------------------------------
  On Error Resume Next
        Set xlApp = GetObject(, "Excel.Application") 'Start Excel if it isn't running
        If xlApp Is Nothing Then
                Set xlApp = CreateObject("Excel.Application")
                bOpenExcel = True
                If xlApp Is Nothing Then
                      MsgBox "Can't start Excel", vbExclamation, "PrintToExcel"
                      Exit Function
                End If
        End If
  On Error GoTo ErrHandle
'-------------------------------
With xlApp
        DoCmd.Hourglass True
        .Visible = False
        
        If bNewFile Then                 ' Create new Worksheet
            Set xlWkBk = .Workbooks.Add
            Set xlSheet = xlWkBk.Worksheets(1)
        Else                             ' Open Existing Template
            Set xlWkBk = .Workbooks.Open(xLPath)
           If SheetName <> "" Then
                Set xlSheet = xlWkBk.Worksheets(SheetName)
           Else
                Set xlSheet = xlWkBk.Worksheets(1)
           End If
        End If
        
        '---------------------------------------------------------------------------------------
        With xlSheet
            If ColumnFormat <> "" Then      ' Set Columns Heading
               For J = 0 To UBound(FORMATS)
                   If FORMATS(J) = "1" Then
                      .COLUMNS(ConvertToLetter(J + 1)).ColumnWidth = XTWidth
                   Else
                      .COLUMNS(ConvertToLetter(J + 1)).ColumnWidth = StandardWidth
                   End If
               Next J
            End If
            
            If sTitle <> "" Then            ' PRINT TITLE
               .Cells(iStartRow, 1).value = sTitle
               .Range("A" & iStartRow & ":" & ConvertToLetter(iTitleLen) & iStartRow).Merge
               .Range("A" & iStartRow & ":" & ConvertToLetter(iTitleLen) & iStartRow).HorizontalAlignment = xlCenter
               .Range("A" & iStartRow & ":" & ConvertToLetter(iTitleLen) & iStartRow).Interior.Color = HEADER_SHADOW
               .Range("A" & iStartRow & ":" & ConvertToLetter(iTitleLen) & iStartRow).Font.Bold = True
               
               iStart = iStartRow + 1
            Else
               iStart = iStartRow
            End If
            
            If sHeader <> "" Then           ' PRINT HEADER
               WORKS = Split(sHeader, DLM)
               For J = 0 To UBound(WORKS)
                    .Cells(iStart, J + 1).value = WORKS(J)
                    If Len(WORKS(J)) > LEN_TO_WRAP Then .Cells(iStartRow, J + 1).WrapText = True
                    
                    .Cells(iStart, J + 1).HorizontalAlignment = xlCenter
                    .Cells(iStart, J + 1).Interior.Color = HEADER_SHADOW
                    .Cells(iStart, J + 1).Font.Bold = True
               Next J
               iStart = iStart + 1
            End If
            
            For I = 0 To nRows              ' PRINT CELLS
                WORKS = Split(Arr(I), DLM)
                For J = 0 To UBound(WORKS)
                    If ColumnFormat <> "" Then
                        If FORMATS(J) = 0 Or FORMATS(J) = "" Then
                            .Cells(I + iStart, J + 1).value = WORKS(J)
                        ElseIf FORMATS(J) = 1 Then
                            .Cells(I + iStart, J + 1).value = WORKS(J)
                            .Cells(I + iStart, J + 1).WrapText = True
                        ElseIf FORMATS(J) = 2 Then
                            .Cells(I + iStart, J + 1).value = Format(GetProperDate(WORKS(J)), DateFormat)
                            .Cells(I + iStart, J + 1).NumberFormat = DateFormat
                        ElseIf FORMATS(J) = 3 Then
                             .Cells(I + iStart, J + 1).value = CDbl(WORKS(J))
                             .Cells(I + iStart, J + 1).NumberFormat = CurrencyFormat
                        Else
                             .Cells(I + iStart, J + 1).value = WORKS(J)
                             .Cells(I + iStart, J + 1).AutoFit
                        End If
                    Else
                        .Cells(I + iStart, J + 1).value = WORKS(J)
                    End If
                Next J
            Next I
            '------------------------------------
            ' FORMAT COLUMNS WITHOUT FORMATTING
            If ColumnFormat = "" Then
                .COLUMNS("A:" & ConvertToLetter(UBound(WORKS) + 1)).AutoFit
            End If
            
            If bNewFile And SheetName <> "" Then
                 .Name = SheetName
            End If
            
        End With
        '---------------------------------------------------------------------------------------
        
        DoCmd.Hourglass False
        If xLPath <> "" And bNewFile Then
            If IsFileExists(xLPath) Then Kill xLPath
            xlWkBk.SaveAs xLPath
            xlWkBk.Close
            If bOpenExcel Then xlApp.Quit
            MsgBox "Excel Workbook updates finished. Your data in the file " & vbCrLf & xLPath, vbOKOnly, "PrintToExcel"
        ElseIf xLPath <> "" And Not bNewFile Then
            xlWkBk.Save
            MsgBox "Excel Workbook updates finished. Your data in the file " & vbCrLf & xLPath, vbOKOnly, "PrintToExcel"
            .Visible = True
        Else
            MsgBox "Excel Workbook updates finished. You Can save It Now", vbOKOnly, "PrintToExcel"
            .Visible = True
        End If
        
End With

'-------------------------------
ExitHere:
    Set xlSheet = Nothing: Set xlWkBk = Nothing: Set xlApp = Nothing
    PrintToExcel = xLPath '!!!!!!!!!!!!!!!
    DoCmd.Hourglass False
    Exit Function
'---------------------
ErrHandle:
    ErrPrint2 "PrintToExcel", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Resume ExitHere
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------
' Function Convert num of column to alphabetic name
'--------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ConvertToLetter(iCol As Integer) As String
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      ConvertToLetter = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
   End If
End Function

'===================================================================================================================================================
' Create Excel and Print some 2-D Array to excel with or without header
'===================================================================================================================================================
Public Function ExcelExport(sQuery As String, Optional xLPath As String, Optional sSheetName As String, Optional colsCurrency As String, _
                            Optional colsDate As String, Optional FRMT_CUR As String = "$#,##0.00", _
                            Optional FRMT_DATE As String = "yyyy-mm-dd", Optional DLM As String = "|") As String
Dim xlApp As Object, xlWkBk As Object, bOpenExcel As Boolean, xlSheet As Object
Dim sPath As String

Const BGRD_COLOR As Long = 13158500
Const xlCenter As Integer = -4108
'----------------------------------------------------
On Error GoTo ErrHandle

DoCmd.Hourglass True

If xLPath <> "" Then
   sPath = xLPath
Else
   sPath = CurrentProject.Path & "\" & FilenameWithoutExtension(CurrentProject.Name) & "_" & ProveFileName(sQuery) & ".xlsx"
End If

If Dir(sPath) <> "" Then Kill sPath

DoCmd.TransferSpreadsheet TransferType:=acExport, _
        SpreadsheetType:=acSpreadsheetTypeExcel12Xml, _
        TableName:=sQuery, fileName:=sPath, HasFieldNames:=True
'----------------------------------------------
  On Error Resume Next
  
        If Trim(sSheetName & colsCurrency & colsDate) = "" Then GoTo ExitHere
        
        Set xlApp = GetObject(, "Excel.Application") 'Start Excel if it isn't running
        If xlApp Is Nothing Then
                Set xlApp = CreateObject("Excel.Application")
                bOpenExcel = True
                If xlApp Is Nothing Then
                      MsgBox "Can't start Excel", vbExclamation, "PrintToExcel"
                      Exit Function
                End If
        End If
  On Error GoTo ErrHandle
  
  If Dir(sPath) = "" Then Err.Raise 10000, , "Can't Find the Excel File " & sPath
'-------------------------------
With xlApp
  .Visible = False
  Set xlWkBk = .Workbooks.Open(sPath)
  Set xlSheet = xlWkBk.Worksheets(1)
  
  If sSheetName <> "" Then xlSheet.Name = sSheetName
  If colsCurrency <> "" Then Call NumberFormatColumns(xlSheet, colsCurrency, FRMT_CUR, DLM)
  If colsDate <> "" Then Call NumberFormatColumns(xlSheet, colsCurrency, FRMT_DATE, DLM)
  
  With xlSheet.Cells
        .Select
        .EntireColumn.AutoFit
        .HorizontalAlignment = xlCenter
  End With

  xlWkBk.Save
  
  .Visible
  If bOpenExcel Then .Quit
End With
'-------------------------------
ExitHere:
    MsgBox "The Data are saved to Excel File " & xLPath, vbInformation, "ExcelExport"
    
    Set xlSheet = Nothing: Set xlWkBk = Nothing: Set xlApp = Nothing
    DoCmd.Hourglass False
    ExcelExport = xLPath '!!!!!!!!!!!!!!!
    Exit Function
'----------------
ErrHandle:
    ErrPrint2 "ExcelExport", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Resume ExitHere
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------------
' NumberFormat Field
'--------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub NumberFormatColumns(ByRef xlSheet As Object, sColumns As String, sFormat As String, Optional DLM As String = "|")
Dim COLUMNS() As Integer, I As Integer, nDim As Integer

On Error Resume Next
'-----------------------
If sColumns = "" Then Exit Sub
If sFormat = "" Then Exit Sub
'-----------------------
COLUMNS = Split(sColumns, DLM): nDim = UBound(COLUMNS)
For I = 0 To nDim
    With xlSheet.COLUMNS(COLUMNS(I))
        .NumberFormat = sFormat
    End With
Next I

End Sub
