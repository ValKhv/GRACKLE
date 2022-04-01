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
' Класс инкапсулирует KV-string, сериализуемую строку набора пар Key=Value разделенных SDELIM, спец.
' (c) 2012 V.Khvatov DigitalXpert Inc
' По умолчанию SDELIM = ";", по умолчанию разделитель между Key и Value - "="
' Пользователь может также переопределить SDELIM
'*********************************************************************************************************************************
'*********************************************************************************************************************************
'*********************************************************************************************************************************

Option Compare Database
Option Explicit

Private Const SDELIM_DEFAULT As String = ";"                   ' Разделитель между парами по умолчанию
Private Const KVDELIM_DEFAULT As String = "="                  ' Разделитель между ключом и значением по умолчанию
'**************************************************************
Private Type sKV             ' Тип КЛЮЧ-ЗНАЧЕНИЕ
     sKey As String
     sValue As String
End Type
'**************************************************************
Private Type QuickStack      ' Тип для сортировки
    High As Long
    Low As Long
End Type

Private sDelim As String                       ' Разделитель строки
Private KVDELIM As String                      ' Разделитель Key и Value
Private mKVString As String                    ' KV строка
Private nDim As Long                           ' размерность массива ключ- значение
Private KV() As sKV                            ' Массив пар ключ-значение (внутреннее хранение строки)
Private bRightString As Boolean                ' Флаг правильности строки


'=================================================================================================================================
' Свойство: разделитель KVString (по умолчанию ";")
'=================================================================================================================================
Public Property Get DLM() As String
    DLM = sDelim
End Property
Public Property Let DLM(sDelimiter As String)
    sDelim = sDelimiter
End Property
'=================================================================================================================================
' Свойство: разделитель пары Key и Value (по умолчанию "=")
'=================================================================================================================================
Public Property Get SEP() As String
    SEP = KVDELIM
End Property
Public Property Let SEP(sDelimiter As String)
    KVDELIM = sDelimiter
End Property
'==================================================================================================================================
' Свойство: KV - строка
'=================================================================================================================================
Public Property Get KVString() As String
    If Not IsKV Then Exit Sub                ' ПРи несуществующем массиве строить строку нельзя
    Call BuildString
    KVString = mKVString  '!!!!!!!!!!!!!
End Property
Public Property Let KVString(sKVString As String)
Dim sWork As String                                  ' Рабочая строка
Dim sC() As String                                   ' Временный архив для пар
Dim iLeft1 As Long, iLeft2 As Long, iRight As Long
'---------------------------------------------------------------------
' Отсекаем разделители слева и права (такое часто встречается)
If Left(sKVString, 1) = sDelim Then
            sWork = Right(sKVString, Len(sKVString) - 1)
ElseIf Right(sKVString, 1) = sDelim Then
            sWork = Left(sKVString, Len(sKVString) - 1)
Else
            sWork = sKVString
End If
'----------------------------------------------------------------------
' Проверяем, что разделители нужного типа имеются и они чередуются, если нет - отбрасываем некорректные пары
       mKVString = CheckStringKV(sWork)
       If mKVString <> "" Then
            bRightString = True        ' Строка проверена
            Call BuildArray            ' Перестраиваем массив
       End If
End Property
'==================================================================================================================================
' Функция для KV- строки по заданному ключу вычисляет значение
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
' Метод возвращает индекс для заданного ключа
'=================================================================================================================================
Public Function IndexOf(sKey As String) As Long
Dim I As Long                                          ' Циклическая переменна
Dim nRes As Long                                       ' Возвращаемый результат

'--------------------------------------
nRes = -1
If Not IsKV Then GoTo IndexOf_EXIT
'--------------------------------------
    For I = 0 To nDim
        If UCase(Trim(sKey)) = UCase(Trim(KV(I).sKey)) Then      ' Нашли ключ
            nRes = I
            Exit For
        End If
    Next I
'----------------------------------------------------------
IndexOf_EXIT:
            IndexOf = nRes
End Function
'==================================================================================================================================
' Метод определяет, имеется ли уже такой ключ
'=================================================================================================================================
Public Function IsKeyExist(sKey As String) As Boolean
    If IndexOf(sKey) > -1 Then IsKeyExist = True  '!!!!!!!!!!!!!!!
End Function
'==================================================================================================================================
' Метод для заданного ключа возвращает его значение
' Если ключ отсутсвует, возвращается пустая строка
'=================================================================================================================================
Public Function GetValueOf(sKey As String) As String
Dim iINDX As Long                              ' Индекс соответсвующего ключа
Dim sRes As String                             ' Возвращаемый результат

If Not IsKV Then Exit Function                 ' ПРи несуществующем массиве получать значения нельзя
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
' Метод добавляет ключ и значение. Если ключ существует - значение обновляется, если ключа не существует - то он добавляется
' Функция возвращает индекс
'=================================================================================================================================
Public Function AddValueFor(sKey As String, sValue As String) As Long
Dim iINDX As Long
     If Not IsKV Then Exit Function                ' При несуществующем массиве выходим
     iINDX = IndexOf(sKey)
     If iINDX = -1 Then     ' Ключ не существует, добавляем и значение, и ключ
       nDim = nDim + 1: ReDim Preserve KV(nDim)
       KV(nDim).sKey = sKey
       KV(nDim).sValue = sValue
       iINDX = nDim
     Else                   ' Ключ существует, для заданного ключа обновляем значение
       KV(iINDX).sValue = sValue
    End If
'------------------------------------------------------
    AddValueFor = iINDX     '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'==================================================================================================================================
' Функция удаляет ключ и соответствующее значение. Если ключ не найден, ничего не происходит
'==================================================================================================================================
Public Sub RemoveKey(sKey As String)
Dim iINDX As Long                                                        ' Индекс удаляемого ключа
Dim sTempArray() As sKV, I As Long, J As Long                            ' Временный массив и циклические переменные

     iINDX = IndexOf(sKey)
     If iINDX = -1 Then Exit Sub          ' Ключ не существует, выходим
     J = -1: ReDim sTempArray(0)          ' Подготавливаем временный массив
     '--------------------------------------------------------------------
     ' Прямая прогонка массива
     For I = 0 To nDim
         If iINDX <> I Then
            J = J + 1: ReDim Preserve sTempArray(J)
            sTempArray(J).sKey = KV(I).sKey
            sTempArray(J).sValue = KV(I).sValue
          End If
     Next I
     '--------------------------------------------------------------------
     ' Перезаливаем временный массив в постоянный
     nDim = nDim - 1: ReDim KV(nDim)
     For I = 0 To nDim
        KV(I).sKey = sTempArray(I).sKey
        KV(I).sValue = sTempArray(I).sValue
     Next I
'---------------------------------------------------------------------------------------------
    Erase sTempArray
End Sub
'==================================================================================================================================
' Функция возвращает список ключей в виде строки, разделенной sDelim
'==================================================================================================================================
Public Function GetKeyList() As String
Dim I As Long, sRes As String
    If Not IsKV Then Exit Function                ' ПРи несуществующем массиве получать значения нельзя
    For I = 0 To nDim
        sRes = IIf(I < nDim, sRes & KV(I).sKey & sDelim, sRes & KV(I).sKey)
    Next I
'---------------------------------------------
    GetKeyList = sRes '!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'==================================================================================================================================
' Функция возвращает список значений в виде строки, разделенной sDelim
'==================================================================================================================================
Public Function GetValueList() As String
Dim I As Long, sRes As String
    If Not IsKV Then Exit Function                ' ПРи несуществующем массиве получать значения нельзя
    For I = 0 To nDim
        sRes = IIf(I < nDim, sRes & KV(I).sValue & sDelim, sRes & KV(I).sValue)
    Next I
'---------------------------------------------
    GetValueList = sRes '!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'==================================================================================================================================
' Функция возвращает массив строковых пар Key=Value
'==================================================================================================================================
Public Function GetKVArray() As String()
Dim I As Long, sRes() As String

If Not IsKV Then Exit Function                ' ПРи несуществующем массиве получать значения нельзя
ReDim sRes(nDim)
    For I = 0 To nDim
        sRes(I) = KV(I).sKey & KVDELIM & KV(I).sValue
    Next I
'---------------------------------------------
    GetKVArray = sRes '!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'==================================================================================================================================
' Функция сортирует имеющийся массив по ключам от самого меньшего к большему
'==================================================================================================================================
Public Sub Sort(Optional descending As Boolean)
Dim SKVTemp As Variant                                      ' Массив строковых пар вида Key=Value (для сортировки)
Dim I As Long, sPAIR() As String                            ' Циклическая переменная и массив пары для парсинга

If Not IsKV Then Exit Sub                ' ПРи несуществующем массиве сортировать нельзя
SKVTemp = GetKVArray()                                                      ' Получаем массив пар
QuickSortNonRecursive SKVTemp, descending                                   ' Сортируем массив пар
'-----------------------------------------------------------------------------------------------------
' Перезаписываем основной массив, идя по сортированному массиву
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
' Инициализация класса
'-------------------------------------------------------------------------------------------------------------------------------
 Private Sub Class_Initialize()
    sDelim = SDELIM_DEFAULT
    KVDELIM = KVDELIM_DEFAULT
    mKVString = ""
    bRightString = False           ' Правильная строка не загружена
 End Sub
'-------------------------------------------------------------------------------------------------------------------------------
' Разрушение класса
'-------------------------------------------------------------------------------------------------------------------------------
Private Sub Class_Terminate()
   Erase KV
End Sub
'-------------------------------------------------------------------------------------------------------------------------------
' Функция проверяет строку на правильность. Если строка правильная - она возвращается.
' Если строка не правильная - делается попытка исправить строку.
' Если в результате исправления (отбраковывания неправильных пар или при отсутствии разделителей) остается пустая строка
' Генерится ошибка
'-------------------------------------------------------------------------------------------------------------------------------
Private Function CheckStringKV(sKV As String) As String
Dim sRes As String                                                               ' Рабочие переменные
Dim sTemp1() As String, sTemp2() As String, k As Long, n As Long                 ' Временные массивы и циклические переменные

On Error GoTo IsStringKV_EXIT

sRes = ""
'----------------------------------------------------------------------
' Общая быстрая проверка Проверяем, что разделители нужного типа имеются и они чередуются
    If CountSubString(sKV, KVDELIM) - CountSubString(sKV, sDelim) = 1 Then       ' Строка правильная
       sRes = sKV
       GoTo IsStringKV_EXIT
    Else                                                                         ' Строка не верна, начинаем процессинг по парам
       sTemp1 = Split(sKV, sDelim): n = UBound(sTemp1)
       For k = 0 To n                                                            ' Итерируем по парам
        '--------------------------------------------------------------------
        If InStr(1, sTemp1(k), KVDELIM) > 0 Then    ' Разделитель 2 обнаружен
            sTemp2 = Split(sTemp1(k), KVDELIM)
            If sTemp2(0) <> "" Then
                sRes = sRes & sTemp2(0) & KVDELIM & sTemp2(1) & sDelim
            End If
        End If
        '---------------------------------------------------------------------
       Next k
    End If
If Right(sRes, Len(sDelim)) = sDelim Then sRes = Left(sRes, Len(sRes) - Len(sDelim)) ' Убираем самый правый разделитель
If sRes = "" Then MsgBox "Ошибка. Строка " & vbCrLf & IIf(Len(sKV) > 60, Left(sKV, 59) & "...", sKV) & "не содержит необходимых пар"
'----------------------------------------------------------------------------
IsStringKV_EXIT:
        CheckStringKV = sRes '!!!!!!!!!!!!!!!!!!!!!!!
End Function
'-------------------------------------------------------------------------------------------------------------------------------
' Функция строит из имеющейся строки массив, в которым потом хранится информация
' Если каждая пара проходит валидацию, то возвращается True
'-------------------------------------------------------------------------------------------------------------------------------
Private Sub BuildArray()
Dim sTemp() As String, I As Long, J As Long             ' Временный массив и циклическая переменная
Dim TestNDim As Long                                    ' Реальный размер массива
Dim sTempKV() As String                                 ' Массив из ключа и значения
Dim sWork As String                                     ' Рабочая строчка
Const SBlank As String = " "                            ' Окаймляющий пробел

If Not bRightString Then Exit Sub
    sWork = ""
    sTemp = Split(mKVString, sDelim): TestNDim = UBound(sTemp)
    nDim = -1: ReDim KV(0)                                         ' Подготавливаем основной массив
    '------------------------------------------------------------------------------------
    For I = 0 To TestNDim                                          ' Заполняем структуру в цикле
           If sTemp(I) <> "" Then                                  ' Пропускаем пустые пары
            sTempKV = Split(sTemp(I), KVDELIM)                     ' Разбиваем пару
            If sTempKV(0) <> "" Then                               ' Ключ не может быть пустым. Пустой ключ - игнорируем
                If InStr(1, sWork, SBlank & sTempKV(0) & SBlank) = 0 Then       ' Удостоверяемся, что еще не вводили такой ключ
                    sWork = sWork & SBlank & sTempKV(0) & SBlank                ' Наращиваем строку для последующей проверки
                    nDim = nDim + 1: ReDim Preserve KV(nDim)                    ' Масштабируем массив
                    KV(nDim).sKey = Trim(sTempKV(0))                                  ' Присваиваем ключ
                    KV(nDim).sValue = Trim(sTempKV(1))                                ' Присваиваем значение
                 Else                                                           ' Такой ключ уже существует, даем знать пользователю
                     MsgBox "Попытка ввода повторного ключа " & sTempKV(0) & vbCrLf & _
                            "Новое значение ключа будет проигнорировано ", vbCritical, "Добавление строки"
                 End If
            End If
           End If
    Next I
    '------------------------------------------------------------------------------------
End Sub
'-------------------------------------------------------------------------------------------------------------------------------
' Функция строит строку KV из массива
'-------------------------------------------------------------------------------------------------------------------------------
Private Sub BuildString()
Dim I As Long


mKVString = ""
    For I = 0 To nDim
        mKVString = mKVString & KV(I).sKey & KVDELIM & KV(I).sValue & IIf(I < nDim, sDelim, "")
    Next I
End Sub
'-------------------------------------------------------------------------------------------------------------------------------
' Функция строит строку KV из массива
'-------------------------------------------------------------------------------------------------------------------------------
Private Function CountSubString(sString As String, sSubString As String) As Long
Dim lRes As String                   ' Возвращаемый результат
lRes = 0
If sString = "" Or sSubString = "" Then Exit Function
        lRes = (Len(sString) - Len(Replace(sString, sSubString, ""))) / Len(sSubString)
'-----------------------------------------------
        CountSubString = lRes    '!!!!!!!!!!!!!!!!!!!!
End Function
'-------------------------------------------------------------------------------------------------------------------------------
' Быстрая, не рекурсивная сортировка массива
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
' Функция проверяет, инициализирован ли основной массив
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


