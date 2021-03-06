VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cKVString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************************************************
'********************************************************************************************************************************
'********************************************************************************************************************************
' ����� ������������� KV-string, ������������� ������ ������ ��� Key=Value ����������� SDELIM, ����.
' (c) 2012 V.Khvatov DigitalXpert Inc
' �� ��������� SDELIM = ";", �� ��������� ����������� ����� Key � Value - "="
' ������������ ����� ����� �������������� SDELIM
'*********************************************************************************************************************************
'*********************************************************************************************************************************
'*********************************************************************************************************************************

Option Compare Database
Option Explicit

Private Const SDELIM_DEFAULT As String = ";"                   ' ����������� ����� ������ �� ���������
Private Const KVDELIM_DEFAULT As String = "="                  ' ����������� ����� ������ � ��������� �� ���������
'**************************************************************
Private Type sKV             ' ��� ����-��������
     sKey As String
     sValue As String
End Type
'**************************************************************
Private Type QuickStack      ' ��� ��� ����������
    High As Long
    Low As Long
End Type

Private sDelim As String                       ' ����������� ������
Private KVDELIM As String                      ' ����������� Key � Value
Private mKVString As String                    ' KV ������
Private nDim As Long                           ' ����������� ������� ����- ��������
Private KV() As sKV                            ' ������ ��� ����-�������� (���������� �������� ������)
Private bRightString As Boolean                ' ���� ������������ ������


'=================================================================================================================================
' ��������: ����������� KVString (�� ��������� ";")
'=================================================================================================================================
Public Property Get DLM() As String
    DLM = sDelim
End Property
Public Property Let DLM(sDelimiter As String)
    sDelim = sDelimiter
End Property
'=================================================================================================================================
' ��������: ����������� ���� Key � Value (�� ��������� "=")
'=================================================================================================================================
Public Property Get SEP() As String
    SEP = KVDELIM
End Property
Public Property Let SEP(sDelimiter As String)
    KVDELIM = sDelimiter
End Property
'==================================================================================================================================
' ��������: KV - ������
'=================================================================================================================================
Public Property Get KVString() As String
    If Not IsKV Then Exit Sub                ' ��� �������������� ������� ������� ������ ������
    Call BuildString
    KVString = mKVString  '!!!!!!!!!!!!!
End Property
Public Property Let KVString(sKVString As String)
Dim sWork As String                                  ' ������� ������
Dim sC() As String                                   ' ��������� ����� ��� ���
Dim iLeft1 As Long, iLeft2 As Long, iRight As Long
'---------------------------------------------------------------------
' �������� ����������� ����� � ����� (����� ����� �����������)
If Left(sKVString, 1) = sDelim Then
            sWork = Right(sKVString, Len(sKVString) - 1)
ElseIf Right(sKVString, 1) = sDelim Then
            sWork = Left(sKVString, Len(sKVString) - 1)
Else
            sWork = sKVString
End If
'----------------------------------------------------------------------
' ���������, ��� ����������� ������� ���� ������� � ��� ����������, ���� ��� - ����������� ������������ ����
       mKVString = CheckStringKV(sWork)
       If mKVString <> "" Then
            bRightString = True        ' ������ ���������
            Call BuildArray            ' ������������� ������
       End If
End Property
'==================================================================================================================================
' ������� ��� KV- ������ �� ��������� ����� ��������� ��������
'==================================================================================================================================
Public Function GetValueFromKV(KVString As String, key As String, _
             Optional PairDelim As String = ";", Optional KVDELIM As String = "=", Optional iCompare As VbCompareMethod = vbTextCompare) As String
Dim sRes As String, sKey As String
Dim MyArr() As String, nDim As Integer, I As Integer

If KVString = "" Or key = "" Then Exit Function
If InStr(1, KVString, key, iCompare) = 0 Then Exit Function
If InStr(1, KVString, PairDelim, iCompare) = 0 Then
     If InStr(1, KVString, KVDELIM, iCompare) = 0 Then Exit Function
     sRes = Split(KVString, KVDELIM)(1)
     GoTo ExitHere
End If
'-------------------------------------------------------------------------------
MyArr = Split(KVString, PairDelim): nDim = UBound(MyArr)
For I = 0 To nDim
     sKey = Trim(Split(MyArr(I), KVDELIM)(0))
     If sKey = Trim(key) Then
        sRes = Trim(Split(MyArr(I), KVDELIM)(1))
        GoTo ExitHere
     End If
Next I
'----------------------------------------------------------
ExitHere:
    GetValueFromKV = sRes '!!!!!!!!!!!!!!
End Function
'==================================================================================================================================
' ����� ���������� ������ ��� ��������� �����
'=================================================================================================================================
Public Function IndexOf(sKey As String) As Long
Dim I As Long                                          ' ����������� ���������
Dim nRes As Long                                       ' ������������ ���������

'--------------------------------------
nRes = -1
If Not IsKV Then GoTo IndexOf_EXIT
'--------------------------------------
    For I = 0 To nDim
        If UCase(Trim(sKey)) = UCase(Trim(KV(I).sKey)) Then      ' ����� ����
            nRes = I
            Exit For
        End If
    Next I
'----------------------------------------------------------
IndexOf_EXIT:
            IndexOf = nRes
End Function
'==================================================================================================================================
' ����� ����������, ������� �� ��� ����� ����
'=================================================================================================================================
Public Function IsKeyExist(sKey As String) As Boolean
    If IndexOf(sKey) > -1 Then IsKeyExist = True  '!!!!!!!!!!!!!!!
End Function
'==================================================================================================================================
' ����� ��� ��������� ����� ���������� ��� ��������
' ���� ���� ����������, ������������ ������ ������
'=================================================================================================================================
Public Function GetValueOf(sKey As String) As String
Dim iINDX As Long                              ' ������ ��������������� �����
Dim sRes As String                             ' ������������ ���������

If Not IsKV Then Exit Function                 ' ��� �������������� ������� �������� �������� ������
sRes = ""
'----------------------------------------------------------------------------
    iINDX = IndexOf(sKey)
    If iINDX > -1 Then sRes = KV(iINDX).sValue
'-----------------------------------------------------------------------------
    GetValueOf = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'==================================================================================================================================
' Implant New KV
'==================================================================================================================================
Public Sub ImplantNewKV(subKV As String)
Dim Pairs() As String, VALS() As String, nDim As Integer, I As Integer
    Pairs = Split(subKV, sDelim): nDim = UBound(Pairs)
    
    For I = 0 To nDim
        VALS = Split(Pairs(I), KVDELIM)
        Call AddValueFor(Trim(VALS(0)), Trim(VALS(1)))
    Next I
    
End Sub
'==================================================================================================================================
' ����� ��������� ���� � ��������. ���� ���� ���������� - �������� �����������, ���� ����� �� ���������� - �� �� �����������
' ������� ���������� ������
'=================================================================================================================================
Public Function AddValueFor(sKey As String, sValue As String) As Long
Dim iINDX As Long
     If Not IsKV Then Exit Function                ' ��� �������������� ������� �������
     iINDX = IndexOf(sKey)
     If iINDX = -1 Then     ' ���� �� ����������, ��������� � ��������, � ����
       nDim = nDim + 1: ReDim Preserve KV(nDim)
       KV(nDim).sKey = sKey
       KV(nDim).sValue = sValue
       iINDX = nDim
     Else                   ' ���� ����������, ��� ��������� ����� ��������� ��������
       KV(iINDX).sValue = sValue
    End If
'------------------------------------------------------
    AddValueFor = iINDX     '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'==================================================================================================================================
' ������� ������� ���� � ��������������� ��������. ���� ���� �� ������, ������ �� ����������
'==================================================================================================================================
Public Sub RemoveKey(sKey As String)
Dim iINDX As Long                                                        ' ������ ���������� �����
Dim sTempArray() As sKV, I As Long, J As Long                            ' ��������� ������ � ����������� ����������

     iINDX = IndexOf(sKey)
     If iINDX = -1 Then Exit Sub          ' ���� �� ����������, �������
     J = -1: ReDim sTempArray(0)          ' �������������� ��������� ������
     '--------------------------------------------------------------------
     ' ������ �������� �������
     For I = 0 To nDim
         If iINDX <> I Then
            J = J + 1: ReDim Preserve sTempArray(J)
            sTempArray(J).sKey = KV(I).sKey
            sTempArray(J).sValue = KV(I).sValue
          End If
     Next I
     '--------------------------------------------------------------------
     ' ������������ ��������� ������ � ����������
     nDim = nDim - 1: ReDim KV(nDim)
     For I = 0 To nDim
        KV(I).sKey = sTempArray(I).sKey
        KV(I).sValue = sTempArray(I).sValue
     Next I
'---------------------------------------------------------------------------------------------
    Erase sTempArray
End Sub
'==================================================================================================================================
' ������� ���������� ������ ������ � ���� ������, ����������� sDelim
'==================================================================================================================================
Public Function GetKeyList() As String
Dim I As Long, sRes As String
    If Not IsKV Then Exit Function                ' ��� �������������� ������� �������� �������� ������
    For I = 0 To nDim
        sRes = IIf(I < nDim, sRes & KV(I).sKey & sDelim, sRes & KV(I).sKey)
    Next I
'---------------------------------------------
    GetKeyList = sRes '!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'==================================================================================================================================
' ������� ���������� ������ �������� � ���� ������, ����������� sDelim
'==================================================================================================================================
Public Function GetValueList() As String
Dim I As Long, sRes As String
    If Not IsKV Then Exit Function                ' ��� �������������� ������� �������� �������� ������
    For I = 0 To nDim
        sRes = IIf(I < nDim, sRes & KV(I).sValue & sDelim, sRes & KV(I).sValue)
    Next I
'---------------------------------------------
    GetValueList = sRes '!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'==================================================================================================================================
' ������� ���������� ������ ��������� ��� Key=Value
'==================================================================================================================================
Public Function GetKVArray() As String()
Dim I As Long, sRes() As String

If Not IsKV Then Exit Function                ' ��� �������������� ������� �������� �������� ������
ReDim sRes(nDim)
    For I = 0 To nDim
        sRes(I) = KV(I).sKey & KVDELIM & KV(I).sValue
    Next I
'---------------------------------------------
    GetKVArray = sRes '!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'==================================================================================================================================
' ������� ��������� ��������� ������ �� ������ �� ������ �������� � ��������
'==================================================================================================================================
Public Sub Sort(Optional descending As Boolean)
Dim SKVTemp As Variant                                      ' ������ ��������� ��� ���� Key=Value (��� ����������)
Dim I As Long, sPAIR() As String                            ' ����������� ���������� � ������ ���� ��� ��������

If Not IsKV Then Exit Sub                ' ��� �������������� ������� ����������� ������
SKVTemp = GetKVArray()                                                      ' �������� ������ ���
QuickSortNonRecursive SKVTemp, descending                                   ' ��������� ������ ���
'-----------------------------------------------------------------------------------------------------
' �������������� �������� ������, ��� �� �������������� �������
For I = 0 To nDim
    sPAIR = Split(SKVTemp(I), KVDELIM)
    KV(I).sKey = sPAIR(0)
    KV(I).sValue = sPAIR(1)
Next I
'------------------------------------------------------------------------------------------------------
Erase SKVTemp
End Sub
'###############################################################################################################################
'###############################################################################################################################
'######/////////###//////////###/////////####//###########//#######///#########//////////###/////////###########################
'######//#####//###//######//######///########//#########//#######//#//############//#######//##################################
'######//#####//###//######//######///#########//#######//#######//###//###########//#######//##################################
'######/////////###//////////######///##########//#####//#######//#####//##########//#######/////###############################
'######//##########//###//#########///###########//###//#######///////////#########//#######//##################################
'######//##########//####//########///############//#//#######//#########//########//#######//##################################
'######//##########//#####///###/////////##########///#######//###########//#######//#######/////////###########################
'###############################################################################################################################
'###############################################################################################################################
'-------------------------------------------------------------------------------------------------------------------------------
' ������������� ������
'-------------------------------------------------------------------------------------------------------------------------------
 Private Sub Class_Initialize()
    sDelim = SDELIM_DEFAULT
    KVDELIM = KVDELIM_DEFAULT
    mKVString = ""
    bRightString = False           ' ���������� ������ �� ���������
 End Sub
'-------------------------------------------------------------------------------------------------------------------------------
' ���������� ������
'-------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
   Erase KV
End Sub
'-------------------------------------------------------------------------------------------------------------------------------
' ������� ��������� ������ �� ������������. ���� ������ ���������� - ��� ������������.
' ���� ������ �� ���������� - �������� ������� ��������� ������.
' ���� � ���������� ����������� (�������������� ������������ ��� ��� ��� ���������� ������������) �������� ������ ������
' ��������� ������
'-------------------------------------------------------------------------------------------------------------------------------
Private Function CheckStringKV(sKV As String) As String
Dim sRes As String                                                               ' ������� ����������
Dim sTemp1() As String, sTemp2() As String, k As Long, n As Long                 ' ��������� ������� � ����������� ����������

On Error GoTo IsStringKV_EXIT

sRes = ""
'----------------------------------------------------------------------
' ����� ������� �������� ���������, ��� ����������� ������� ���� ������� � ��� ����������
    If CountSubString(sKV, KVDELIM) - CountSubString(sKV, sDelim) = 1 Then       ' ������ ����������
       sRes = sKV
       GoTo IsStringKV_EXIT
    Else                                                                         ' ������ �� �����, �������� ���������� �� �����
       sTemp1 = Split(sKV, sDelim): n = UBound(sTemp1)
       For k = 0 To n                                                            ' ��������� �� �����
        '--------------------------------------------------------------------
        If InStr(1, sTemp1(k), KVDELIM) > 0 Then    ' ����������� 2 ���������
            sTemp2 = Split(sTemp1(k), KVDELIM)
            If sTemp2(0) <> "" Then
                sRes = sRes & sTemp2(0) & KVDELIM & sTemp2(1) & sDelim
            End If
        End If
        '---------------------------------------------------------------------
       Next k
    End If
If Right(sRes, Len(sDelim)) = sDelim Then sRes = Left(sRes, Len(sRes) - Len(sDelim)) ' ������� ����� ������ �����������
If sRes = "" Then MsgBox "������. ������ " & vbCrLf & IIf(Len(sKV) > 60, Left(sKV, 59) & "...", sKV) & "�� �������� ����������� ���"
'----------------------------------------------------------------------------
IsStringKV_EXIT:
        CheckStringKV = sRes '!!!!!!!!!!!!!!!!!!!!!!!
End Function
'-------------------------------------------------------------------------------------------------------------------------------
' ������� ������ �� ��������� ������ ������, � ������� ����� �������� ����������
' ���� ������ ���� �������� ���������, �� ������������ True
'-------------------------------------------------------------------------------------------------------------------------------
Private Sub BuildArray()
Dim sTemp() As String, I As Long, J As Long             ' ��������� ������ � ����������� ����������
Dim TestNDim As Long                                    ' �������� ������ �������
Dim sTempKV() As String                                 ' ������ �� ����� � ��������
Dim sWork As String                                     ' ������� �������
Const SBlank As String = " "                            ' ����������� ������

If Not bRightString Then Exit Sub
    sWork = ""
    sTemp = Split(mKVString, sDelim): TestNDim = UBound(sTemp)
    nDim = -1: ReDim KV(0)                                         ' �������������� �������� ������
    '------------------------------------------------------------------------------------
    For I = 0 To TestNDim                                          ' ��������� ��������� � �����
           If sTemp(I) <> "" Then                                  ' ���������� ������ ����
            sTempKV = Split(sTemp(I), KVDELIM)                     ' ��������� ����
            If sTempKV(0) <> "" Then                               ' ���� �� ����� ���� ������. ������ ���� - ����������
                If InStr(1, sWork, SBlank & sTempKV(0) & SBlank) = 0 Then       ' ��������������, ��� ��� �� ������� ����� ����
                    sWork = sWork & SBlank & sTempKV(0) & SBlank                ' ���������� ������ ��� ����������� ��������
                    nDim = nDim + 1: ReDim Preserve KV(nDim)                    ' ������������ ������
                    KV(nDim).sKey = Trim(sTempKV(0))                                  ' ����������� ����
                    KV(nDim).sValue = Trim(sTempKV(1))                                ' ����������� ��������
                 Else                                                           ' ����� ���� ��� ����������, ���� ����� ������������
                     MsgBox "������� ����� ���������� ����� " & sTempKV(0) & vbCrLf & _
                            "����� �������� ����� ����� ��������������� ", vbCritical, "���������� ������"
                 End If
            End If
           End If
    Next I
    '------------------------------------------------------------------------------------
End Sub
'-------------------------------------------------------------------------------------------------------------------------------
' ������� ������ ������ KV �� �������
'-------------------------------------------------------------------------------------------------------------------------------
Private Sub BuildString()
Dim I As Long


mKVString = ""
    For I = 0 To nDim
        mKVString = mKVString & KV(I).sKey & KVDELIM & KV(I).sValue & IIf(I < nDim, sDelim, "")
    Next I
End Sub
'-------------------------------------------------------------------------------------------------------------------------------
' ������� ������ ������ KV �� �������
'-------------------------------------------------------------------------------------------------------------------------------
Private Function CountSubString(sString As String, sSubString As String) As Long
Dim lRes As String                   ' ������������ ���������
lRes = 0
If sString = "" Or sSubString = "" Then Exit Function
        lRes = (Len(sString) - Len(Replace(sString, sSubString, ""))) / Len(sSubString)
'-----------------------------------------------
        CountSubString = lRes    '!!!!!!!!!!!!!!!!!!!!
End Function
'-------------------------------------------------------------------------------------------------------------------------------
' �������, �� ����������� ���������� �������
'-------------------------------------------------------------------------------------------------------------------------------
Private Sub QuickSortNonRecursive(ByRef SortArray As Variant, Optional descending As Boolean)
    Dim I As Long, J As Long, lB As Long, ub As Long
    Dim stack() As QuickStack, stackpos As Long, maxstackpos As Long, stposArrMax As Long, ppos As Long, pivot As Variant, swp
    
    lB = LBound(SortArray)
    ub = UBound(SortArray)
    stposArrMax = 16
    ReDim stack(stposArrMax)
    
    stackpos = 1
    maxstackpos = 1
    stack(1).Low = lB
    stack(1).High = ub
    Do
        lB = stack(stackpos).Low
        ub = stack(stackpos).High
        stackpos = stackpos - 1
        Do
            ppos = (lB + ub) \ 2
            I = lB: J = ub: pivot = SortArray(ppos)
            Do
                While IIf(descending, SortArray(I) > pivot, SortArray(I) < pivot): I = I + 1: Wend
                While IIf(descending, pivot > SortArray(J), pivot < SortArray(J)): J = J - 1: Wend
                If I > J Then Exit Do
                swp = SortArray(I): SortArray(I) = SortArray(J): SortArray(J) = swp
                I = I + 1
                J = J - 1
           Loop While I <= J

            If I < ppos Then
                stackpos = stackpos + 1
                If stackpos > maxstackpos Then maxstackpos = stackpos
                If stackpos > stposArrMax Then stposArrMax = stposArrMax * 2: ReDim Preserve stack(stposArrMax)
                stack(stackpos).Low = I
                stack(stackpos).High = ub
                ub = J
            Else
                If J > lB Then
                    stackpos = stackpos + 1
                    If stackpos > maxstackpos Then maxstackpos = stackpos
                    If stackpos > stposArrMax Then stposArrMax = stposArrMax * 2: ReDim Preserve stack(stposArrMax)
                    stack(stackpos).Low = lB
                    stack(stackpos).High = J
                End If
                lB = I
            End If
        Loop While lB < ub
    Loop While stackpos
End Sub
'---------------------------------------------------------------------------------------------------------------------------------
' ������� ���������, ��������������� �� �������� ������
'---------------------------------------------------------------------------------------------------------------------------------
Private Function IsKV() As Boolean
Dim bRes As Boolean

On Error GoTo IsKV_ERR
   If KV(0).sKey <> "" Then
      bRes = True
   End If
'-----------------------------------------------------------
   IsKV = bRes '!!!!!!!!!!!!!!!!!
   Exit Function
'------------------------------------------------------------
IsKV_ERR:
     Err.Clear
End Function


