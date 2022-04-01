Attribute VB_Name = "#_DATETIME"
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
'                 $$$$$$$$$$F                                       ## ##  ##  ##  ##    ## ####         ##  ###### ###### ###### #### ##   ## #####
'                  *$$$$$$$$"                                        ####  #####  ##     ## ##  ##     ## ##   ##   ##       ##    ##  ### ### ##
'                    "***""               _____________                                     ##   ##   ##  ##   ##   ####     ##    ##  ## # ## ####
' STANDARD MODULE WITH DEFAULT FUNCTIONS |v 2017/03/19 |                                    ##   ##  #######   ##   ##       ##    ##  ##   ## ##
' The module contains frequently used functions and is part of the G-VBA library             #####  ##    ##   ##   ######   ##   #### ##   ## #####
'****************************************************************************************************************************************************
'****************************************************************************************************************************************************
Option Explicit

#If VBA7 And Win64 Then
    Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public Declare PtrSafe Function GetTickCount Lib "kernel32.dll" () As Long
#Else
    Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
#End If



Private Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Private Const MOD_NAME As String = "#_DATETIME"
'****************************************************************************************************************************************************

'====================================================================================================================================================
'   Get timestamp
'====================================================================================================================================================
Public Function GetTimestamp(Optional InDate As Date) As Double
Dim STT As SYSTEMTIME, dt As Date
Dim iOffSet As Long   ' Days between 1970/01/01 and 1900/01/01 (

On Error Resume Next
   
#If Mac Then
    iOffSet = 24107
#Else
    iOffSet = 25569
#End If
    
    If IsZero(InDate) Then
           dt = Now()
    Else
           dt = Now()
    End If
    
    GetSystemTime STT
    GetTimestamp = (((dt - iOffSet) * 86400) - (3600 * 9)) * 1000 + STT.wMilliseconds
End Function

Public Function FormatDateEx(dt As Double) As String
    FormatDateEx = Format(dt / 86400, "yyyy-mm-dd HH:mm:ss") & "." & ((dt - Fix(dt)) * 1000)
End Function

'====================================================================================================================================================
' Stop Program for Some time
'====================================================================================================================================================
Public Sub Wait(Optional mTimeMSec As Long = 1000)

On Error Resume Next
'-------------------
    Sleep mTimeMSec    '1000 = wait 1 second
End Sub
'====================================================================================================================================================
' Soft Breakpoint
'====================================================================================================================================================
Public Sub Pause()
       Debug.Assert False
End Sub

'======================================================================================================================================================
' Return Zero Date - i.e. 0:00:00
'======================================================================================================================================================
Public Function ZeroDate() As Date

End Function
'======================================================================================================================================================
' Get Proper Date
'======================================================================================================================================================
Public Function GetProperDate(sDate As String, Optional DateFormat As String = "dd/MM/yyyy") As Date
Dim dRes As Date, sWork As String, sDateParts() As String

On Error GoTo ErrHandle
'-------------------------
    If sDate = "" Then Exit Function
    If IsDate(sDate) Then
        dRes = CDate(sDate)
    Else
        sWork = TryDate(sDate)
        If IsDate(sWork) Then
            dRes = CDate(sWork)
        Else
                sWork = DateInvert(sDate)
                If IsDate(sWork) Then dRes = CDate(sWork)
        End If
    End If
'-------------------------
ExitHere:
    GetProperDate = dRes  '!!!!!!!!!!!!!!!!!!!!
    Exit Function
'-----------
ErrHandle:
    ErrPrint2 "GetProperDate", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------------
' Function try to get date indirect
'--------------------------------------------------------------------------------------------------------------------------------------------------
Private Function TryDate(sDate As String) As String
On Error Resume Next
    TryDate = CStr(DateValue(sDate)) '!!!!!!!!!!!!
End Function
'====================================================================================================================================================
' Correct Date according current format
'       RUSSIAN DATE INVERT TO USA AND VICE VERSE
'====================================================================================================================================================
Public Function DateInvert(sDate As Variant) As String
Dim sCurrentPattern As String, WW() As String, sTime As String, nDim As Integer, I As Integer
Dim sRes As String, iL As Integer, YYYY As String, MMM As String, DD As String

Const InitYear As Integer = 1600
Const FinitYear As Integer = 2030
Const DayLimit As Integer = 31

On Error Resume Next
'----------------------------------
sRes = UCase(Trim(CStr(sDate)))
sTime = TimeFromDate(sRes)
If sTime <> "" Then sRes = Trim(Replace(sRes, sTime, ""))
sTime = CStr(TimeValue(sTime))
'--------------------------------
sRes = Replace(sRes, "/", "."): sRes = Replace(sRes, "\", "."): sRes = Replace(sRes, "-", ".")
sRes = Replace(sRes, ",", "."): sRes = Replace(sRes, "  ", " "): sRes = Replace(sRes, " ", ".")

If IsNumeric(sRes) Then ' No any separaters - means ISO-like format
      If CInt(Right(sRes, 4)) > InitYear And CInt(Right(sRes, 4)) < FinitYear Then
        sRes = Left(sRes, 2) & "." & Mid(sRes, 3, 2) & "." & Right(sRes, 4)
      ElseIf CInt(Left(sRes, 4)) > InitYear And CInt(Left(sRes, 4)) < FinitYear Then
        sRes = Left(sRes, 4) & "." & Mid(sRes, 5, 2) & "." & Right(sRes, 2)
      End If
End If

iL = InStr(1, sRes, "."): If iL = 0 Then GoTo ExitHere
WW = Split(sRes, "."): nDim = UBound(WW): If nDim > 5 Then GoTo ExitHere
iL = -1
'---------------------------------
' RECOGNIZE PARTS
  MMM = FindMonthName(WW(0), WW(1), WW(2)) ' First try to find not numerical part
  
For I = 0 To nDim              ' Processing numerical parts
     If IsNumeric(WW(I)) Then
           If CInt(WW(I)) > DayLimit Then
                iL = I
                If CInt(WW(I)) > InitYear Then
                    YYYY = WW(I)
                ElseIf CInt(WW(I)) < 35 Then
                    YYYY = "20" & WW(I)
                Else
                    YYYY = "19" & WW(I)
                End If
           ElseIf CInt(WW(I)) > 12 Then
                DD = WW(I)
           Else
                If MMM = "" Then  ' No Month Yet
                   MMM = WW(I)
                Else
                   DD = WW(I)
                End If
           End If
     End If
Next I

If DD = "" Then
   DD = IIf(iL = 0, WW(2), WW(0))
   MMM = IIf(iL = 0, WW(1), WW(2))
End If
'---------------------------------------------------------------------------------------
sRes = Trim(CStr(DateSerial(CInt(YYYY), CInt(MMM), CInt(DD))) & " " & sTime)
'---------------------------------------------
ExitHere:
     DateInvert = sRes '!!!!!!!!!!!!!
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------
' Find month name for DateParts array
'----------------------------------------------------------------------------------------------------------------------------------------------------
Private Function FindMonthName(part1 As String, part2 As String, part3 As String) As String
Dim iRes As Integer

On Error Resume Next
   If Not IsNumeric(part1) Then
          iRes = ConvertMonthName(part1)
   ElseIf Not IsNumeric(part2) Then
          iRes = ConvertMonthName(part2)
   ElseIf Not IsNumeric(part3) Then
          iRes = ConvertMonthName(part3)
   End If
'--------------------------
    If iRes > 0 Then FindMonthName = CStr(iRes) '!!!!!!!!
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------
' Convert month name to number
'----------------------------------------------------------------------------------------------------------------------------------------------------
Private Function ConvertMonthName(MonthName As String) As Integer
Dim iRes As Integer
         Select Case MonthName
         Case "JA", "JAN", "JANUARY", "ßÍ", "ßÍÂ", "ßÍÂÀÐÜ", "ßÍÂÀÐß", "JANVIER":
                iRes = 1
         Case "FE", "FEB", "FEBRUARY", "ÔÅÂ", "ÔÅÂÐÀËÜ", "ÔÅÂÐÀËß", "FEVRIER":
                iRes = 2
         Case "MR", "MAR", "MARCH", "ÌÀÐ", "ÌÀÐÒ", "ÌÀÐÒÀ", "MARS":
                iRes = 3
         Case "AL", "APR", "APRIL", "ÀÏÐ", "ÀÏÐÅËÜ", "ÀÏÐÅËß", "AVRIL":
                iRes = 4
         Case "MA", "MAY", "ÌÀ", "ÌÀÉ", "ÌÀß", "MAI":
                iRes = 5
         Case "JN", "JUN", "JUNE", "ÈÞÍ", "ÈÞÍÜ", "ÈÞÍß", "JUIN":
                iRes = 6
         Case "JL", "JUL", "JULY", "ÈÞË", "ÈÞËÜ", "ÈÞËß", "JUILLET":
                iRes = 7
         Case "AU", "AUG", "AUGUST", "ÀÂÃ", "ÀÂÃÓÑÒÀ", "AOUT":
                iRes = 8
         Case "SE", "SEPT", "SEPTEMBER", "ÑÅÍ", "ÑÅÍÒßÁÐÜ", "ÑÅÍÒßÁÐß", "SEPTEMBRE":
                iRes = 9
         Case "OC", "OCT", "OCTOBER", "ÎÊÒ", "ÎÊÒßÁÐÜ", "ÎÊÒßÁÐß", "OCTOBRE":
                iRes = 10
         Case "NO", "NOV", "NOVEMBER", "ÍÎß", "ÍÎßÁÐÜ", "ÍÎßÁÐß", "NOVEMBRE":
                iRes = 11
         Case "DE", "DEC", "DECEMBER", "ÄÅÊ", "ÄÅÊÀÁÐÜ", "ÄÅÊÀÁÐß", "DECEMBRE":
                iRes = 12
         End Select
'----------------------------------------
ExitHere:
         ConvertMonthName = iRes '!!!!!!!!!!!!!!
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------
' Extract Time from General Date
'-----------------------------------------------------------------------------------------------------------------------------------------------------
Public Function TimeFromDate(sDate As String) As String
Dim iL As Integer, sRes As String
    sRes = Trim(sDate)
    iL = InStr(1, sRes, ":")
    If iL > 0 Then
         iL = InStrRev(sRes, " ", iL)
         If iL > 0 Then
             sRes = Right(sRes, Len(sRes) - iL)
         Else
             sRes = ""
         End If
    Else
         sRes = ""
    End If
'----------------------------------------
ExitHere:
    TimeFromDate = sRes '!!!!!!!!!!!!!
End Function
'-----------------------------------------------------------------------------------------------------------------------------------------------------
' Function check date format
'------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function DateFormat() As String
  DateFormat = CStr(DateSerial(1999, 1, 2))
  DateFormat = Replace(DateFormat, "1999", "yyyy")
  DateFormat = Replace(DateFormat, "99", "yy")
  DateFormat = Replace(DateFormat, "01", "mm")
  DateFormat = Replace(DateFormat, "1", "m")
  DateFormat = Replace(DateFormat, "02", "dd")
  DateFormat = Replace(DateFormat, "2", "d")
  DateFormat = Replace(DateFormat, MonthName(1), "mmmm")
  DateFormat = Replace(DateFormat, MonthName(1, True), "mmm")
End Function



Private Function RetDate(ByVal sText As String, Optional DLM As String = ";") As String
Dim iL As Integer, LDate As Integer, LYYYY As Integer
Dim I As Integer, YYYY As String
Dim sRes As String, sWork As String

Const StartYear As Integer = 2015
Const EndYear As Integer = 2016


For I = StartYear To EndYear
    YYYY = CStr(I)
    iL = InStr(1, sText, CStr(I))
    If iL > 0 Then
            Do While iL > 0
             ' TEMPLATE DD/MM/YYYY (DD.MM.YYYY;DD-MM-YYYY)
                 sWork = Mid(sText, iL - 6, 10)
                 If IsDate(sWork) Then
                    sRes = sRes & sWork & DLM
                    GoTo NextLoop
                 End If
             ' TEMPLATE D/MM/YYYY (D.MM.YYYY;D-MM-YYYY)
                 sWork = Mid(sText, iL - 5, 9)
                 If IsDate(sWork) Then
                    sRes = sRes & sWork & DLM
                    GoTo NextLoop
                 End If
             ' TEMPLATE YYYY/MM/DD (YYYY.MM.DD,YYYY-MM-DD)
                 sWork = Mid(sText, iL, 10)
                 If IsDate(sWork) Then
                    sRes = sRes & sWork & DLM
                    GoTo NextLoop
                  End If
             ' TEMPLATE YYYY/MM/D (YYYY.MM.D,YYYY-MM-D)
                 sWork = Mid(sText, iL, 9)
                 If IsDate(sWork) Then
                     sRes = sRes & sWork & DLM
                     GoTo NextLoop
                 End If
             ' TEMPLATE
            '----------------------------------------------------------
NextLoop:
              iL = iL + 4
              iL = InStr(iL, sText, CStr(I))
            Loop
   End If
Next I
'--------------------------------------------------------------------------
If Right(sRes, 1) = DLM Then sRes = Left(sRes, Len(sRes) - 1)
ExitHere:
   RetDate = sRes '!!!!!!!!!!!!!!!!!!!!!
End Function

Private Function MonthLong(sTRR As String, YYYY As String, YYYYPos As Integer, Optional DLM As String) As String
Dim MonthAray() As String, nDim As Integer, sRes As String, sRepl  As String
Dim I As Integer, J As Integer, iL As Integer, sWork As String, m As Integer, n As Integer, mP As Integer


Const MonthList As String = "January;February;March;April;May;June;July;August;September;October;November;December"
Const MonthListShort As String = "Jan;Feb;Mar;Apr;May;June;July;Aug;Sept;Oct;Nov;Dec"
Const Radius As Integer = 8

MonthAray = Split(MonthListShort & ";" & MonthList, ";"): nDim = UBound(MonthAray)



m = YYYYPos - Radius: n = YYYYPos + Radius
If m < 1 Then m = 1: If n > Len(sTRR) Then n = Len(sTRR)
sWork = Mid(sTRR, m, n - m)
'------------------------------------------------------------------------
For I = 0 To nDim
      iL = InStr(1, sWork, MonthAray(I), vbTextCompare)
      If iL > 0 Then
           If iL > 2 Then                   ' Format 14 Jul 1790/2016-Jul-07/01.Jul.2005/20 Jul,2016
                If Mid(sTRR, iL - 1, 1) = " " And IsNumeric(Mid(sTRR, iL - 2, 1)) Then
                     sRes = "."
                     For J = iL - 1 To 0 Step -1 ' Going Back
                         'If Mid()
                     Next J
                End If
       End If
                                            
                                            ' Format
      End If
      
           sWork = IIf(I < 9, "0" & (I + 1), CStr(I + 1))
           
            For mP = 1 To 5         ' Going Back
                sRepl = Mid(sTRR, iL - mP, 1)
                If sRepl = "" Then
                    sWork = sRes
                End If
            Next mP
            mP = iL - 1
            
           
           
            sRepl = IIf(I < 9, "0" & (I + 1), CStr(I + 1))
            
                
            
                 sWork = " " & sWork: sRepl = "." & sRepl
            
            
            If Mid(sTRR, iL + Len(sWork) + 1, 1) = " " And IsNumeric(Mid(sTRR, iL - 2, 1)) Then ' Format 14 Jul 1790
                 sWork = " " & sWork: sRepl = "." & sRepl
            End If
            
      
Next I

End Function
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
