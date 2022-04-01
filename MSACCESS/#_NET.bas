Attribute VB_Name = "#_NET"
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
'                 $$$$$$$$$$F                                       ## ##  ##  ##  ##    ##   ##   ## ###### ######
'                  *$$$$$$$$"                                        ####  #####  ##     ##   ###  ## ##       ##
'                    "***""               _____________                                       ## # ## ####     ##
' STANDARD MODULE WITH DEFAULT FUNCTIONS |v 2017/03/19 |                                      ##  ### ##       ##
' The module contains frequently used functions and is part of the G-VBA library              ##   ## #####    ##
'****************************************************************************************************************************************************
'****************************************************************************************************************************************************
Option Explicit

Const MOD_NAME As String = "#_NET"
'*******************************


'====================================================================================================================================================
' Function Return Currency Conversion (use exchangeratesapi)
'====================================================================================================================================================
Public Function CurrencyConversion(Optional DD As Date, Optional CurrencyIn As String = "EUR", Optional CurrencyOut As String = "CAD") As Double
Dim MyRequest As Object, JSON As Object, sWork As String
Dim sURL As String, sDate As String, a As Object, dRes As Double

Const sURLBase As String = "https://api.exchangeratesapi.io/history?"

On Error GoTo ErrHandle
'----------------------------
If IsZero(DD) Then
    sDate = Format(Now(), "YYYY-MM-DD")
Else
    sDate = Format(DD, "YYYY-MM-DD")
End If

sURL = sURLBase & "start_at=" & sDate & "&end_at=" & sDate & _
       "&symbols=" & CurrencyOut & "&base=" & CurrencyIn

Set MyRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
MyRequest.Open "GET", sURL
MyRequest.Send

'----------------------------
    If MyRequest.ResponseText <> "" Then
        Set JSON = [#_JSON].ParseJSON(MyRequest.ResponseText)
        Set a = JSON("rates")(sDate): sWork = a(CurrencyOut)
        If IsNumeric(sWork) Then dRes = CDbl(sWork)
    End If
'----------------------------
ExitHere:
    CurrencyConversion = dRes '!!!!!!!!!!!!!!!!!!!
    Set JSON = Nothing: Set MyRequest = Nothing
    Exit Function
'---------------
ErrHandle:
    ErrPrint2 "CurrencyConversion", Err.Number, Err.Description, MOD_NAME
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
                                                                                                  Optional sModName As String = "#_NET") As String
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



