Attribute VB_Name = "#_STRUCTURES"
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
'                 $$$$$$$$$$F                                       ## ##  ##  ##  ##    ##          #### ###### ######  ## ##  #### ######
'                  *$$$$$$$$"                                        ####  #####  ##     ##          ##     ##   ##  ##  ## ## ##      ##
'                    "***""               _____________                                               ##    ##   ####    ## ## ##      ##
' STANDARD MODULE WITH DEFAULT FUNCTIONS |v 2020/01/12 |                                               ##   ##   ## ##   ## ## ##      ##
' The module contains functions to process some logic structurea as part of the G-VBA library        ####   ##   ##  ###  ###   ####   ##
'****************************************************************************************************************************************************
'****************************************************************************************************************************************************
Option Explicit
Option Compare Database

Private Const MOD_NAME As String = "modSort"
'**********************

'====================================================================================================================================================
' The Function Return one Column from 2D-Array
'====================================================================================================================================================
Private Function GetColumn(TwoDArr As Variant, ColumIndex As Long) As Variant
Dim nRows As Long, vR() As Variant, I As Long

    On Error GoTo ErrHandle
'-------------------------
nRows = UBound(TwoDArr, 1)
ReDim vR(nRows)

For I = 0 To nRows
    vR(I) = TwoDArr(I, ColumIndex)
Next I
'-------------------------
ExitHere:
    GetColumn = vR '!!!!!!!!!!!!
    Exit Function
'-----------------
ErrHandle:
    ErrPrint2 "GetColumn", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'====================================================================================================================================================
' Get Max for Array
'====================================================================================================================================================
Public Function MaxArray(Arr As Variant) As Variant
Dim nDim As Long, mDim As Long, I As Long
Dim iMax As Variant

On Error Resume Next
'--------------------------------
   If Not IsArray(Arr) Then Exit Function
   nDim = UBound(Arr): mDim = LBound(Arr): iMax = Arr(mDim)
   For I = mDim To nDim
        If Arr(I) > iMax Then iMax = Arr(I)
   Next I
'--------------------------------
ExitHere:
  MaxArray = iMax '!!!!!!!!!!!!!!!!!!!
End Function
'====================================================================================================================================================
' Get Min for Array
'====================================================================================================================================================
Public Function MinArray(Arr As Variant) As Variant
Dim nDim As Long, mDim As Long, I As Long
Dim iMin As Variant

On Error Resume Next
'--------------------------------
   If Not IsArray(Arr) Then Exit Function
   nDim = UBound(Arr): mDim = LBound(Arr): iMin = Arr(mDim)
   For I = mDim To nDim
        If Arr(I) < iMin Then iMin = Arr(I)
   Next I
'--------------------------------
ExitHere:
  MinArray = iMin '!!!!!!!!!!!!!!!!!!!
End Function

'======================================================================================================================================================
' Two Arrays Intersection
'   The function compares the elements of the second array (arr2) with the first one (arr1) and returns an array
'   with those elements of the second array that do not match the first
'======================================================================================================================================================
Public Function ArrayIntersection(ARR1 As Variant, ARR2 As Variant, _
                                                        Optional CompareMode As VbCompareMethod = vbTextCompare) As Variant
Dim Arr As Variant, nArr As Long
Dim LBound_arr1 As Long, UBound_arr1 As Long, LBound_arr2 As Long, UBound_arr2 As Long
Dim I As Long, J As Long, iCalc As Long

On Error GoTo ErrHandle
'---------------------------
If Not IsArray(ARR1) Then Exit Function
If Not IsArray(ARR2) Then Exit Function

    If IsObject(ARR1(LBound(ARR1))) Then Err.Raise 10000, , "Can't intersect for object arrays"
    If varType(ARR1(LBound(ARR1))) >= vbArray Then Err.Raise 10000, , "Can't itersect for unlnow type arrays"
    If varType(ARR1(LBound(ARR1))) = vbUserDefinedType Then Err.Raise 10000, , "Can't itersect for arrays of User Defined Types"

    ReDim Arr(0): nArr = -1                                ' PREPARE THE RESULT ARRAY
    For I = LBound(ARR2) To UBound(ARR2)                   ' Each element of the second array
        iCalc = 0                                          ' is compared with the first one
        For J = LBound(ARR1) To UBound(ARR1)
            If IsNumeric(ARR1(I)) Then
               If CDbl(ARR2(I)) = CDbl(ARR1(J)) Then iCalc = iCalc + 1
            Else
               If StrComp(CStr(ARR2(I)), CStr(ARR1(J)), CompareMode) = 0 Then iCalc = iCalc + 1
            End If
        Next J
        
        If iCalc = 0 Then                                  ' If iCalc = 0 means no matches were found and
                nArr = nArr + 1: ReDim Preserve Arr(nArr)  ' the element is placed in the resulting array
                Arr(nArr) = ARR2(I)
        End If
    Next I
'---------------------------
ExitHere:
     ArrayIntersection = Arr   '!!!!!!!!!!!!!!!!!
     Exit Function
'-----------------
ErrHandle:
     ErrPrint2 "ArrayIntersection", Err.Number, Err.Description, MOD_NAME
     Err.Clear
End Function

'======================================================================================================================================================
' Merge two arrays
'======================================================================================================================================================
Public Function MergeArrays(ByVal ARR1 As Variant, ByVal ARR2 As Variant) As Variant
        Dim tmpArr As Variant, upper1 As Long, upper2 As Long
        Dim higherUpper As Long, I As Long, newIndex As Long
        upper1 = UBound(ARR1) + 1: upper2 = UBound(ARR2) + 1
        higherUpper = IIf(upper1 >= upper2, upper1, upper2)
        ReDim tmpArr(upper1 + upper2 - 1)
 
        For I = 0 To higherUpper
            If I < upper1 Then
                tmpArr(newIndex) = ARR1(I)
                newIndex = newIndex + 1
            End If
 
            If I < upper2 Then
                tmpArr(newIndex) = ARR2(I)
                newIndex = newIndex + 1
            End If
        Next I
        MergeArrays = tmpArr
End Function

'======================================================================================================================================================
' QUICK-SORT FOR 1-D ARRAY
'======================================================================================================================================================
Public Sub Quicksort(Arr As Variant, lB As Long, ub As Long)
Dim tmpArr As Variant, tmpArr2    As Variant
Dim iLow   As Long, iHi    As Long
 
    On Error GoTo ErrHandle
'------------------------------------
iLow = lB
iHi = ub
tmpArr = Arr((lB + ub) \ 2)
 
While (iLow <= iHi)                                                             ' partionning array first
   While (Arr(iLow) < tmpArr And iLow < ub)
      iLow = iLow + 1
   Wend
  
   While (tmpArr < Arr(iHi) And iHi > lB)
      iHi = iHi - 1
   Wend
 
   If (iLow <= iHi) Then
      tmpArr2 = Arr(iLow)
      Arr(iLow) = Arr(iHi)
      Arr(iHi) = tmpArr2
      iLow = iLow + 1: iHi = iHi - 1
   End If
Wend
 
  If (lB < iHi) Then Quicksort Arr, lB, iHi                    ' loop through recursively
  If (iLow < ub) Then Quicksort Arr, iLow, ub
'--------------------------
ExitHere:
    Exit Sub
'--------
ErrHandle:
    ErrPrint2 "Quicksort", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Sub

'======================================================================================================================================================
' BUBBLE SORT FOR 1-D ARRAY
'======================================================================================================================================================
Public Sub BubbleSort(Arr As Variant)

Dim I As Long, J As Long
Dim temp As Variant
 
    On Error GoTo ErrHandle
'------------------------------------
For I = LBound(Arr) To UBound(Arr) - 1
    For J = I + 1 To UBound(Arr)
        If Arr(I) > Arr(J) Then
            temp = Arr(J)
            Arr(J) = Arr(I)
            Arr(I) = temp
        End If
    Next J
Next I
'--------------------------
ExitHere:
    Exit Sub
'--------
ErrHandle:
    ErrPrint2 "BubbleSort", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Sub


'=========================================================================================================================================== ===========
' SORT VIA ARRAY LIST STRUCTURE - using Microsoft's built-in ArrayList class (System.Collections)
'======================================================================================================================================================
Public Sub ArrayListSort(Arr As Variant)
Dim list As Object, lB As Long, ub As Long, I As Long


    On Error GoTo ErrHandle
'------------------------------------
Set list = CreateObject("System.Collections.ArrayList")
lB = LBound(Arr): ub = UBound(Arr)                             '  Overhead for filling the structure
For I = lB To ub
    list.Add Arr(I)
Next I

list.Sort                                                       ' Using system library sort method
Arr = list.ToArray

'--------------------------
ExitHere:
    Set list = Nothing
    Exit Sub
'--------
ErrHandle:
    ErrPrint2 "BubbleSort", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Resume ExitHere
End Sub


'=========================================================================================================================================== ===========
' HEAP SORT FOR 1-D Array
'======================================================================================================================================================
Public Sub HeapSort(Arr As Variant)
Dim lB As Long, Count As Long
Dim iStart As Long, iEnd As Long
 
    On Error GoTo ErrHandle
'------------------------------------
lB = LBound(Arr)
Count = UBound(Arr) - lB + 1
iStart = (Count - 2) \ 2
iEnd = Count - 1

    While iStart >= 0
        heapTopDown Arr, iStart, iEnd
        iStart = iStart - 1
    Wend
 
    Dim temp As Variant
 
    While iEnd > 0
        temp = Arr(lB + iEnd)
        Arr(lB + iEnd) = Arr(lB)
        Arr(lB) = temp
 
        iEnd = iEnd - 1
 
        heapTopDown Arr, 0, iEnd
    Wend
'--------------------------
ExitHere:
    Exit Sub
'--------
ErrHandle:
    ErrPrint2 "HeapSort", Err.Number, Err.Description, MOD_NAME
    Err.Clear: Resume ExitHere
End Sub
'------------------------------------------------------------------------------------------------------------------------------------------------------
'  (Heap) pyramid creation helper
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub heapTopDown(Arr As Variant, iStart As Long, iEnd As Long)
Dim root As Long
    Dim lB As Long
    Dim temp As Variant
 
    On Error Resume Next
'------------------------------------
root = iStart: lB = LBound(Arr)

    While root * 2 + 1 <= iEnd
        Dim child As Long: child = root * 2 + 1
        If child + 1 <= iEnd Then
            If Arr(lB + child) < Arr(lB + child + 1) Then
                child = child + 1
            End If
        End If
        If Arr(lB + root) < Arr(lB + child) Then
            temp = Arr(lB + root)
            Arr(lB + root) = Arr(lB + child)
            Arr(lB + child) = temp

            root = child
        Else
            Exit Sub
        End If
    Wend
End Sub


'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'  TEST SEVERAL SORT PROCEDURES
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Sub TEST_Sort()
Dim Arr() As Variant, sARR() As Variant
Dim I As Integer, nDim As Integer
Dim dTime As Double

    sARR = Array("Bella", "Anna", "Eugene", "Fedora", "Alevtiva", "Marina", "Elena", "Julia", "Tatyana", "Inna")
    nDim = 1000
    
    Debug.Print "The Input:" & Join(sARR, ";")
 
    
    dTime = Timer
    For I = 0 To nDim
       Arr = sARR
       Quicksort Arr, LBound(Arr), UBound(Arr)
    Next I
    
    Debug.Print String(40, "-")
    
    dTime = Timer - dTime
    Debug.Print "Quicsort time: " & dTime
    Debug.Print "Quicksort result: " & Join(Arr, ";")
    
    Debug.Print String(40, "-")
    
    dTime = Timer
    For I = 0 To nDim
       Arr = sARR
       BubbleSort Arr
    Next I
    dTime = Timer - dTime
    Debug.Print "BubbleSort time: " & dTime
    Debug.Print "BubbleSort result: " & Join(Arr, ";")
    
    
    Debug.Print String(40, "-")
    
    
    dTime = Timer
    For I = 0 To nDim
       Arr = sARR
       ArrayListSort Arr
    Next I
    dTime = Timer - dTime
    Debug.Print "ArrayListSort time: " & dTime
    Debug.Print "ArrayListSort result: " & Join(Arr, ";")
    
    dTime = Timer
    For I = 0 To nDim
       Arr = sARR
       HeapSort Arr
    Next I
    dTime = Timer - dTime
    Debug.Print "HeapSort time: " & dTime
    Debug.Print "HeapSort result: " & Join(Arr, ";")
    
    
End Sub














