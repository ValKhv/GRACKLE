Attribute VB_Name = "#_MATH"
'*********************************.ze$$e. ************************************************************************************************************
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
'                 d$$$$$$$$$F                $P"                   ##  ##  ##  ##   #######    ##       ##      ####  ######  ##   ##
'                 $$$$$$$$$$F                                       ## ##  ##  ##  ##    ##    ###     ###     ## ##    ##    ##   ##
'                  *$$$$$$$$"                                        ####  #####  ##     ##    ## ## ## ##    ##  ##    ##    ##   ##
'                    "***""               _____________                                        ##   #   ##   #######    ##    #######
' STANDARD MODULE WITH DEFAULT FUNCTIONS |v 2021/08/20 |                                       ##       ##  ##    ##    ##    ##   ##
' The module contains some functions to work with MS Access and is part of the G-VBA library   ##       ## ##     ##    ##    ##   ##
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
Option Explicit



Private Const MOD_NAME As String = "#_MATH"
'***************************

'====================================================================================================================================================
' Max of two or more numbers
'====================================================================================================================================================
Public Function max(V1 As Variant, Optional V2 As Variant, Optional V3 As Variant) As Variant
Dim vMax As Variant, tempArr() As String

On Error GoTo ErrHandle
'---------------------------------------
If IsArray(V1) Then
            vMax = MaxArray(V1)
Else
            If Not IsMissing(V2) Then
                    vMax = IIf(V1 > V2, V1, V2)
            
            Else
                    vMax = V1
            End If
            If Not IsMissing(V3) Then
                    vMax = IIf(vMax > V3, vMax, V3)
            End If
End If
'---------------------------------------
ExitHere:
    max = vMax '!!!!!!!!!
    Exit Function
'------------------
ErrHandle:
    ErrPrint2 "max", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'====================================================================================================================================================
' Min of two or more numbers
'====================================================================================================================================================
Public Function min(V1 As Variant, Optional V2 As Variant, Optional V3 As Variant) As Variant
Dim vMin As Variant

On Error GoTo ErrHandle
'---------------------------------------
If IsArray(V1) Then
            vMin = MinArray(V1)
Else
            If Not IsMissing(V2) Then
                    vMin = IIf(V1 < V2, V1, V2)
            Else
                    vMin = V1
            End If
            If Not IsMissing(V3) Then
                    vMin = IIf(vMin < V3, vMin, V3)
            End If
End If
'---------------------------------------
ExitHere:
    min = vMin '!!!!!!!!!
    Exit Function
'------------------
ErrHandle:
    ErrPrint2 "min", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'====================================================================================================================================================
' Get Random Number between two numbers
'====================================================================================================================================================
Public Function GetRandom(Optional lowerBound As Variant = 0, Optional upperBound As Variant = 100, Optional bInt As Boolean = True) As Variant
Dim vRes As Variant

On Error Resume Next
'--------------------------------------
    Randomize
    vRes = Int((upperBound - lowerBound + 1) * Rnd + lowerBound)
    
    If Not bInt Then vRes = vRes + Rnd(1)
'--------------------------------------
ExitHere:
    GetRandom = vRes '!!!!!!!!!!!!
End Function

