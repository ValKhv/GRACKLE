VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************************************
'*********************************************************************************************************************
'*********************************************************************************************************************
' DESCRIPTION: Класс представляет собой манипулятор строками и реализует всяческий парсинг
' VERSION    : v 0.1 19.05.2010
' DEPENDENCY : RegExp (MS vbscript.dll)
' COPYRIGHT  :
'*********************************************************************************************************************
'*********************************************************************************************************************
Option Explicit

'*********************************************************************
'======================================================================================================================
' Функция принимает на вход два списка, возвращает общие значения (пересечения двух списков)
'======================================================================================================================
Public Function MultiplicateList(FirstList As String, SecondList As String, Optional sDelim As String = ";") As String
Dim sRes As String                                                                       ' Возвращаемый результат
Dim f() As String, NF As Integer, s() As String, NS As Integer, I As Integer, J As Integer

On Error GoTo ErrHandle
'---------------------------------------------------------------------------------------------
If FirstList = "" Or SecondList = "" Then Exit Function
f = Split(FirstList, sDelim): NF = UBound(f): s = Split(SecondList, sDelim): NS = UBound(s)
    For I = 0 To NF
        For J = 0 To NS
            If f(I) = s(J) Then
                sRes = sRes & f(I) & sDelim
            End If
        Next J
    Next I
    If Len(sRes) > 0 Then sRes = Left(sRes, Len(sRes) - 1)
'-----------------------------------------------------------------
ExitHere:
        MultiplicateList = sRes '!!!!!!!!!!!!!!!
        Exit Function
'-----------------------------
ErrHandle:
        Debug.Print Err.Description
        Err.Clear
        MultiplicateList = "" '!!!!!!!!!!!!!!!
End Function
'======================================================================================================================
' Функция вычисляет, входит ли в список заданное выражение - слово
'======================================================================================================================
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
    If StrComp(sList, sWork, iCompare) Then
        bRes = True
        GoTo ExitHere
    Else
        Exit Function
    End If
End If
'-----------------------------------
MyArr = Split(sList, sDelim): nDim = UBound(MyArr)
    For I = 0 To nDim
        If StrComp(Trim(MyArr(I)), sWord, iCompare) Then
           bRes = True
           GoTo ExitHere
        End If
    Next I
'-------------------------------------------------------------
ExitHere:
    InList = bRes '!!!!!!!!!!!!!!!!
End Function
'======================================================================================================================
' Функция проверяет, входит ли заданное слово в строку, которая может содержать шаблоны.
' При этом решается задача обратная FindWord - в списке могут быть заданы шаблоны
' PARAMETERS:  strList     - входящая строка / список, разделенный ";"
'              strWord     - слово, которое проверяется на нахождение в списке
'              bIgnoreCase - игнорирование регистра
' RETURNS:     True, если слово обнаружено.
'======================================================================================================================
Public Function IsWordInList(strList As String, strWord As String, Optional bIgnoreCase As Boolean = True) As Boolean
Dim strWORDS() As String, nDim As Long                      ' Массив слов в списке и его размерность
Dim I As Long                                               ' Циклическая переменная и рабочая строка
Dim bRes As Boolean                                         ' Возвращаемый результат
Dim RWork As String, LWork As String                        ' Рабочие переменные

'-------------------------------------------------------
    If bIgnoreCase Then             ' преобразуем регистр
           LWork = UCase(strWord)
    Else
           LWork = strWord
    End If
'----------------------------------------------------------
' Получаем массив и его размерность
strWORDS = SplitX(strList)                                  ' Получили список слов
        For I = 0 To UBound(strWORDS)
                '---------------------------------------------
                If bIgnoreCase Then
                   RWork = UCase(strWORDS(I))
                Else
                   RWork = strWORDS(I)
                End If
                '---------------------------------------------
                ' СОБСТВЕННО СРАВНЕНИЕ
                If strWord Like RWork Then   ' Совпадение найдено
                    bRes = True
                    Exit For
                End If
        Next I
'-----------------------------------------------------------
' Возвращаем результат
       IsWordInList = bRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function

'======================================================================================================================
' Функция производит поиск слова целиком внутри строки, в том числе с использованием паттерна, но без использования
' регулярных выражений. разбиение на слова дается стандартной функцией с использованием разграничителей в виде знаков
' препинания и пробелов. Если входящую строку не удается разбить на слова, поиск переходит к сравнению со всей строкой
' (как одно слово). Возвращает все найденные вхожденяи
' PARAMETERS:   strSource   - исходная строка в которой производится поиск
'               strMatch    - подстрока, которая ищется внутристроки strSource. Может быть задана паттерном:
'                              ?           - любой отдельный знак
'                              *           - ноль или более знаков
'                              #           - любая цифра (0-9)
'                              [charlist]  - любой отдельный знак. входящий в список charlist
'                              [!charlist] - любой знак, не входящий в charlist
'               strFounded  - строковый массив, содержащий найленные данные в формате: NPos;SValue;NLen
'                             NPos   - вхождение найденной подстроки (слова) в strSource
'                             SValue - найденная подстрока/ слово (может отличаться, если используется шаблон)
'                             NLen   - длина найденной подстроки/ слова
'               bIgnoreCase - игнорировать разницу в написании
' RETURN:       кол-во найденных вхождений
'               через массив strFounded() - найденные вхождения
'======================================================================================================================
Public Function FindWord(strSource As String, strMatch As String, strFounded() As String, _
                                                                       Optional bIgnoreCase As Boolean = True) As Long
   Dim astrText()    As String              ' Поисковый массив слов (при невозможности разбить - целиком строка)
   Dim lngCount As Long, I As Long          ' Кол-во найденных совпадений и циклическая переменная
   Dim strWork As String, iWork As Long     ' Рабочее слово и рабочая позиция

'----------------------------------------------------------------------------------------
' Инициализируемся
   lngCount = 0                                   ' инициализируем счетчик
   ReDim strFounded(lngCount)                     ' инициализируем массив
   If bIgnoreCase Then strMatch = UCase(strMatch) ' в случае поиска буз учета в регистр - переводим в верхний регистр
'---------------------------------------------------------------------------------------
' Идем в цикле по словам с проверкой маски
   astrText = SplitX(strSource)             ' Получаем собственно массив слов
   iWork = 1                                ' Cчитаем, что первое слово начинается с 1- ой позиции
   For I = 0 To UBound(astrText)
      strWork = astrText(I)                    ' Очередное рабочее слово
      iWork = InStr(iWork, strSource, strWork) ' Ищем позицию нового слова, начиная с предыдущего слова
      ' если игнорируем регистр, сравниваемое слово переводим в верхний регистр
      If bIgnoreCase Then strWork = UCase(strWork)
      '-----------------------------------------------------------------
      ' производим сравнение
         If strWork Like strMatch Then     ' СОВПАДЕНИЕ !!!
           lngCount = lngCount + 1                               ' Наращиваем счетчик
           ReDim Preserve strFounded(lngCount)                   ' Сдвигаем массив
           ' формируем выходной массив
           strFounded(lngCount) = iWork & ";" & astrText(I) & ";" & Len(strWork)
         End If
      '-----------------------------------------------------------------
      iWork = iWork + Len(strWork)                                 ' сдвиагем позицию для нового поиска
   Next I
'--------------------------------------------------------------------------------------
' Завершаем функцию
   FindWord = lngCount   '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'====================================================================================================================
' Функция ищет вхождение подстроки strMatch в строку strSource. Отличие от стандартной функции InStr состоит в том,
' что возвращаются все вхождения, а не только первое. Кроме того в качестве strMatch может использоваться шаблон c
' использованием знаков * и ?. Для обработки применяются регулярные выражения.
' PARAMETERS:   strSource   - исходная строка в которой производится поиск
'               strMatch    - подстрока, которая ищется внутристроки strSource
'               strFounded  - строковый массив, содержащий найленные данные в формате: NPos;SValue;NLen
'                             NPos   - вхождение найденной подстроки в strSource
'                             SValue - найденная подстрока (может отличаться, если используется шаблон)
'                             NLen   - длина найденной подстроки
'               bIgnoreCase - игнорировать разницу в написании
' RETURN:       кол-во найденных вхождений
' EXAMPLE:                        Dim MyStr As New cString
'                                 Dim i As Long, sFounded() As String
'                                 i = MyStr.FindSubstr("tblXXX;tblYY;_erv;sysTBL,__ERT;asdf;tblCV", "tbl", sFounded())
'                                 (возвращается i=4)
'====================================================================================================================
Public Function FindSubstr(strSource As String, strMatch As String, strFounded() As String, _
                                                                      Optional bIgnoreCase As Boolean = True) As Long
Dim re As Object                    '!!! движок регулярных выражений (требуется подключенная библиотека vbscript.dll)
Dim matchs As Object                ' Управление шаблонами
Dim m As Object                     ' Вхождение совпадения в коллекцию совпадений
Dim I As Long, iCount As Long       ' Рабочая циклическая переменная и найденные вхождения

Set re = CreateObject("VBScript.RegExp")
'--------------------------------------------------------------------------------------------
' Проверяем вхождения
re.Pattern = strMatch                  ' Установили паттерн поиска
re.Global = True                       ' для поиска по всему тексту
re.IgnoreCase = bIgnoreCase            ' Игнорировтаь или нет регистр
'-------------------------------------------------------------------------------------------
' Осуществляем поиск
Set matchs = re.Execute(strSource)
    iCount = matchs.Count              ' Считываем кол-во найденных соответствий
    If iCount > 0 Then
        ReDim strFounded(iCount)           ' Подстраиваем массив соотвествий
        I = 0                              ' инициализируем циклическую переменную
        ' Идем по коллекции соответствий
        For Each m In matchs
           I = I + 1
           strFounded(I) = m.FirstIndex + 1 & ";" & m.value & ";" & m.Length
        Next
    End If
'-----------------------------------------------------------------------------
' Завершение функции
        FindSubstr = iCount
        Set re = Nothing
End Function
'=================================================================================================================================================
' Функция удаляет из передаваемого текста массив подстрок (очищает текст)
'=================================================================================================================================================
Public Function ClearText(InputText As String, sClear As String, Optional sDelim As String = ";") As String
Dim Clear() As String, nClear As Integer, I As Integer                ' Размерность массива подстрок для очистки и циклическая переменная
Dim sRes As String                                                    ' Возвращаемый результат
Clear = Split(sClear, sDelim): nClear = UBound(Clear)
'----------------------------------------------------------------------------------
sRes = InputText
For I = 0 To nClear
    sRes = Replace(sRes, Clear(I), "")
Next I
'-------------------------------------------------------
    ClearText = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'====================================================================================================================
' Функция позволяет вычислять из передаваемой строки номер заданного слова по порядку (начиная с 1).
' Если слова нет, возвращается 0
' PARAMETERS:  sSource   - входящая строка, в которой ищется слово
'              sWord     - искомое слово
'              iCompare  - метод сравнения
' RETURNS:     номер заданного слова sWORD в строке sSource
'====================================================================================================================
Public Function GetNumOfWord(ByVal sSource As String, sWord As String, Optional iCompare As VbCompareMethod) As Long
Dim iRes As Long                                 ' Возвращаемый результат
Dim sW() As String, nDim As Long, I As Long      ' массив слов, его размерность и циклическая переменная
iRes = 0
'------------------------------------------------------------------------------------
If InStr(1, sSource, sWord, iCompare) Then       ' если слово вообще находится
    sW = SplitX(sSource)   ' Разбиваем на слова стандартным образом
    nDim = UBound(sW)      ' Размерность массива
    For I = 0 To nDim
       If StrComp(sW(I), sWord, iCompare) Then  ' НАШЛИ !!!
            iRes = I + 1
            Exit For       ' Выходим из цикла
       End If
    Next I
End If
'-----------------------------------------------------------------------
GetNumOfWord_EXIT:
    GetNumOfWord = iRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'====================================================================================================================
' Функция позволяет вычленять слово, стоящее перед заданной подстрокй (в частных случаях может быть символом или
' словом или частью слова). Возвращается именно отдельно стоящее слово. Поиск ведется, начиная с заданной позиции
' до первого совпадения, которое и возвращается
' PARAMETERS:  sSource      - входящая строка, в которой ищется слово
'              sMark        - маркер, перед которым ищется слово
'              iPos         - позиция, начиная с которой ищется sMark
'              iCompare     - метод сравнения (учет регистра)
'              sFIND        - выделенное слово (выходящий параметр) (слово в строке sSource стоящее после
'                            заданного маркером)
'              iRes         - позиция в которой стоит найденное слово (выходящий параметр)
'              sStop        -  подстрока, ограничивающая поиск, если она встречается, то движение по строке в
'                             заданном направлении (назад) прекращается и возвращается текущий результат
'                             то есть поиск - не далее, чем расположение sStop
'              bNoSepRepeat - запрет на повторяющиеся символы списка сепараторов
' RETURNS:                  - позицию, в которой стоит найденный маркер
'====================================================================================================================
Public Function GetPrevWord(ByVal iStart As Long, ByVal sSource As String, ByVal sMark As String, sFind As String, _
                            iRes As Long, Optional iCompare As VbCompareMethod, Optional sStop As String = "", _
                                                                            Optional bNoSepRepeat As Boolean) As Long
Dim sWork As String           ' Текущая строка
Dim bStart As Boolean         ' Флаг начала считывания слова
Dim bFINISH As Boolean        ' Флаг окончания считывания слов
Dim iL As Long, iC As Long    ' Текущий указатель найденного sMARK и сдвиг относительно него
Dim iStop As Long             ' Ограничитель поиска
Dim bInSEPLIST As Boolean     ' Знак, присутствует в списке разделителей SEPS
Dim iRepeat As Integer        ' Повторяющиеся символы
Const NoRepeatCount = 2       ' Ограничение по повторяющимся символам из списка
Const SEPS = ".,:-+/\|^&*!?~ ()[]{}<>'""" & vbTab & vbCrLf

iRes = 0: iC = 0: sFind = "": iRepeat = 0
'------------------------------------
        iL = InStr(iStart, sSource, sMark, iCompare)
'-----------------------
If iL > 0 Then                 ' НАШЛИ sMARK
        If sStop <> "" Then                                         ' Расчитываем ограничитель
              iStop = InStrRev(sSource, sStop, iL, iCompare)        ' Ищем в обратном направлении
              If iStop <> 0 Then iStop = iStop + Len(sStop) + 1     ' По передней грани ограничителя
        Else
              iStop = 0
        End If
    '--------------------------------------------------
    ' Начинаем движение в обратную сторону, начиная с нулевого сдвига относительно sMark,
    ' пока не встретим границу слова
    Do
      sWork = Mid(sSource, iL - iC, 1)              ' Проверяемый символ
      bInSEPLIST = InStr(1, SEPS, sWork, vbTextCompare)
        
      If Not bStart And bInSEPLIST Then bStart = True               ' Включаем запись, как только встретили разделитель
      If bStart And bInSEPLIST And sFind <> "" Then bFINISH = True  ' Включаем, если после начала записи встретили разделитель
              
      If Not bInSEPLIST And bStart And Not bFINISH Then
                    sFind = sWork & sFind
                    iRes = iL - iC
      End If
            ' Проверяем повторение символов из списка (подряд более одного разделителя слов)
                If bNoSepRepeat And bInSEPLIST Then iRepeat = iRepeat + 1
                If iRepeat >= NoRepeatCount Then bFINISH = True
      If bFINISH Then Exit Do                       ' При окончании - выходим из цикла
      iC = iC + 1                                   ' Наращиваем шаги назад
    Loop While iC < iL - iStop
    '-------------------------------------------------
End If
'--------------------------------------
GetPrevWord_EXIT:
    GetPrevWord = iL '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'====================================================================================================================
' Функция позволяет вычленять слово, стоящее после заданной подстрокй (в частных случаях может быть символом или
' словом или частью слова). Возвращается именно отдельно стоящее слово. Поиск ведется, начиная с заданного номера
' до первого совпадения, которое и возвращается
' PARAMETERS:  sSource     - входящая строка, в которой ищется слово
'              sMark       - маркер, перед которым ищется слово
'              iStart        - позиция, начиная с которой ищется sMark (сам поиск ТОЛЬКО ВПЕРЕД)
'              iCompare    - метод сравнения (учет регистра)
'              sFIND       - выделенное слово (выходящий параметр) (слово в строке sSource стоящее перед
'                            заданным маркером)
'              iRes        - позиция в которой стоит найденное слово (выходящий параметр)
'              sStop       -  подстрока, ограничивающая поиск, если она встречается, то движение по строке в
'                            заданном направлении (вперед) прекращается и возвращается текущий результат
'                            то есть поиск - не далее, чем расположение sStop
' RETURNS:                 - позицию, в которой стоит найденный маркер
'====================================================================================================================
Public Function GetNextWord(ByVal iStart As Long, ByVal sSource As String, ByVal sMark As String, sFind As String, _
                        iRes As Long, Optional iCompare As VbCompareMethod, Optional sStop As String = "") As Long
Dim sWork As String           ' Текущая строка
Dim bStart As Boolean         ' Флаг начала считывания слова
Dim bFINISH As Boolean        ' Флаг окончания считывания слов
Dim iL As Long, iC As Long    ' Текущий указатель
Dim iADD As Long              ' Поправка на длину sMark
Dim iStop As Long             ' Ограничитель поиска
Dim bInSEPLIST As Boolean     ' Знак, присутствует в списке разделителей SEPS

Const SEPS = ".,:-+/\|^&*!?~ ()[]{}<>'""" & vbTab & vbCrLf

iRes = 0: iC = 0: sFind = ""
'------------------------------------
        iL = InStr(iStart, sSource, sMark, iCompare)   ' Текущее положение sMark в строке
'-----------------------
If iL > 0 Then                 ' НАШЛИ sMARK
        iADD = Len(sMark) - 1
        iStop = Len(sSource)   ' первоначальное приближение
        If sStop <> "" Then    ' расчитываем актуальное значение ограничителя
              iStop = InStr(iL, sSource, sStop, iCompare)
        End If
    '--------------------------------------------------
    ' Начинаем движение вперед, начиная с нулевого сдвига относительно окончания sMark,
    ' пока не встретим границу слова
    bStart = False: bFINISH = False
    Do
      sWork = Mid(sSource, iL + iC + iADD, 1)            ' Проверяемый символ
      bInSEPLIST = InStr(1, SEPS, sWork, vbTextCompare)
        
      If Not bStart And bInSEPLIST Then bStart = True               ' Включаем запись, как только встретили разделитель
      If bStart And bInSEPLIST And sFind <> "" Then bFINISH = True  ' Включаем, если после начала записи встретили разделитель
        
      If Not bInSEPLIST And bStart And Not bFINISH Then  ' (намеренно так, чтобы пропустить двойные пробелы и т.п.)
                    sFind = sFind & sWork
                    If iRes = 0 Then iRes = iL + iC + iADD
      End If
      If bFINISH Then Exit Do                       ' При окончании - выходим из цикла
      iC = iC + 1                                   ' Наращиваем шаги вперед
    Loop While iC <= iStop - iL - iADD
    '-------------------------------------------------
End If
'--------------------------------------
GetNextWord_EXIT:
    GetNextWord = iL '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'====================================================================================================================
' Функция позволяет вычислять из передаваемой строки слово по заданному номеру. Если номер находится за границами
' диапазона, то возвращается пустая строка
' PARAMETERS:  sSource   - входящая строка, в которой ищется слово
'              Num       - заданный номер, считаемый от 1
' RETURNS:     слово в строке sSource с заданным номером
'====================================================================================================================
Public Function GetWordOfNum(ByVal sSource As String, NUM As Long) As String
Dim sRes As String                               ' Возвращаемый результат
Dim sW() As String, nDim As Long                 ' массив слов, его размерность, рабочая переменная

sRes = ""
'------------------------------------------------------------------------------------
' Получаем массив слов
    sW = SplitX(sSource)   ' Разбиваем на слова стандартным образом
    nDim = UBound(sW)      ' Размерность массива
    If NUM > 1 And NUM - 1 <= nDim Then
        sRes = sW(NUM - 1)
    End If
'-----------------------------------------------------------------------
GetWordOfNum_EXIT:
    GetWordOfNum = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'====================================================================================================================
' Расширенная функция Split, разбивающая входящую строку на отдельные слова.
' PARAMETERS: InputText - входящая строка
'             Delimiter - опциональный разделитель, если обычные (знаки препинания и пробелы) не подходят
' RETURNS:    строковый массив отдельных слов
'====================================================================================================================
Public Function SplitX(ByVal InputText As String, Optional ByVal Delimiter As String) As Variant
    
    Const CHARS = ".!?,;:""'()[]{}"    ' Массив символов, которые разбивают строку на слова
    Dim strReplacedText As String      ' рабочая строковая переменная с учетом подмены символов
    Dim intIndex As Integer            ' рабочий указатель внутри строки

    ' Заменяем tab на space
    strReplacedText = Trim(Replace(InputText, vbTab, " "))
    ' Заменяем все спец. символы, определенные в CHARS на пробелы
    For intIndex = 1 To Len(CHARS)
        strReplacedText = Trim(Replace(strReplacedText, _
            Mid(CHARS, intIndex, 1), " "))
    Next intIndex
    '----------------------------------------------------------------------------------
    ' Заменяем все двойные пробелы на одинарные
    Do While InStr(strReplacedText, "  ")
        strReplacedText = Replace(strReplacedText, "  ", " ")
    Loop
    '----------------------------------------------------------------------------------
    ' Разделяем строку на слова с учетом имеющегося разделителя или без учета такового
    If Len(Delimiter) = 0 Then
        SplitX = Split(strReplacedText)               '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Else
        SplitX = Split(strReplacedText, Delimiter)    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    End If
End Function

'====================================================================================================================
' Функция подсчитывает кол-во слов в строке
'====================================================================================================================
Public Function CountWords(InputText As String) As Long
    Dim astrWords() As String

    ' Split the input string into an array of words.
    astrWords = SplitX(InputText)

    ' Return the number of elements in the array.
    CountWords = UBound(astrWords) - _
        LBound(astrWords) + 1
End Function
'====================================================================================================================
' Вспомогательная функция. выделяет имя из полного пути strIn. Возвращает файловое имя с расширением
'====================================================================================================================
Public Function GetFileNamePart(strIn As String) As String
  Dim intCounter As Integer
  Dim StrTmp As String
  
  On Error GoTo PROC_ERR
  
  For intCounter = Len(strIn) To 1 Step -1
    If Mid$(strIn, intCounter, 1) <> "\" Then
      StrTmp = Mid$(strIn, intCounter, 1) & StrTmp
    Else
      Exit For
    End If
  Next intCounter

  GetFileNamePart = StrTmp
  
PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "GetFileNamePart"
  Resume PROC_EXIT
  
End Function

'=====================================================================================================
' Функция парсит KV-string, то есть строку вида KEY1=VALUE1;KEY2=VALUE2;KEY3=VALUE3 и для заданного
' KEY возвращает VALUE. Если не удалось найти вхождение KEY или возник иной сбой, возвращается
' DEFAULTVALUE
'=====================================================================================================
Public Function GetValueFromKVString(sSource As String, forKEY As String, _
                                          Optional KVSsep As String = ";", Optional Pairsep = "=", _
                                                     Optional DefaultValue As String = "") As String
Dim iKey As Long                                    ' Начало пары
Dim iValue As Long                                  ' Позиция значения
Dim iValueF As Long                                 ' Окончание позиции значения
Dim LenValue As Long                                ' Длина выражения значения
Dim sWork As String                                 ' Рабочая строка

sWork = forKEY & Pairsep                            ' Ключ вместе с разделителем пары
iKey = InStr(1, sSource, sWork, vbTextCompare)
If iKey > 0 Then                       ' Нашли вхождение ключа
  iValue = iKey + Len(sWork)           ' Стартовая позиция подстроки VALUE
  iValueF = InStr(iValue, sSource, KVSsep, vbTextCompare) ' ищем вхождение разделителя KVString
  If iValueF > 0 Then                               ' нашли вхождение следующего разделителя KV- строки
     LenValue = iValueF - iValue
     sWork = Mid(sSource, iValue, iValueF - iValue)
  Else                                              ' не нашли вхождения (возможно - конец строки, возвращ хвост)
     sWork = Right(sSource, Len(sSource) - iValue + 1) ' конец строки
  End If
Else                                   ' Вхождение ключа не найдено
  sWork = DefaultValue
End If
GetValueFromKVString = sWork '!!!!!!!!!!!!!!!!!!!!!!! возвращаем знач по умолчанию
End Function
'================================================================================================================================
' Функция выделяет из строки подстроку по двум тегам LEFTAG и RIGHTTAG. Подразумевается, что закрывающий
' тэг идет после открывающего
'================================================================================================================================
Public Function GetSubstringTAGS(sSource As String, LeftTag As String, RightTag As String) As String
Dim iLeft As Long, iRight As Long                    ' Ограничивающие позиции
Dim sWork As String                                  ' Рабочая строка
sWork = ""
'-----------------------------------------------------------
    iLeft = InStr(1, sSource, LeftTag)
    If iLeft > 0 Then           ' Нашли открывающий тег, пересчитываем позиции
       iLeft = iLeft + Len(LeftTag)
       iRight = InStr(iLeft, sSource, RightTag)
       If iRight > 0 Then       ' Нашли и правый, закрывающий тэг
         sWork = Mid(sSource, iLeft, iRight - iLeft)   ' Выделяем строку
       End If
    End If
'-----------------------------------------------------------
GetSubstringTAGS = sWork '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'=============================================================================================================================
' Функция возвращает для двух тегов (открывающего и закрывающего) подстроку, но с широким правилом выделения:
' Левый тэг ищется первым слева, а правый - первым справа. То есть берется максимально широкая строка. При этом внутри
' выделенной подстроки могут находится другие тэги. Если соответсвующего правого тэга, расположенного справа, но после
' левого тэга не найдено - возвращается пустая строка
'============================================================================================================================
Public Function GetSubstringTAGSMax(sSource As String, LeftTag As String, RightTag As String) As String
Dim iLeft As Long, iRight As Long                    ' Ограничивающие позиции
Dim sWork As String                                  ' Рабочая строка
sWork = ""
'-----------------------------------------------------------
    iLeft = InStr(1, sSource, LeftTag)
    If iLeft > 0 Then                         ' Нашли открывающий тег слева
       iRight = InStrRev(sSource, RightTag)   ' Нашли закрывающий тэг справа
       If iRight > iLeft Then                 ' Пара тэгов верифицирована
         sWork = Mid(sSource, iLeft + Len(LeftTag), iRight - iLeft - Len(LeftTag)) ' Выделяем строку
       End If
    End If
'-----------------------------------------------------------
GetSubstringTAGSMax = sWork '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'============================================================================================================================
' Функция парсит входящую строку, извлекая из нее вложенные тэги (nested tags) и возвращает строковый массив, начинающийся
' с исходной строки, затем следующие по порядку теги и заканчивающийся самой вложеной, самой маленькой строкой.
' Каждая строка массива представляет собой структурированную строку.
' PARAMETERS:   sSource - входящая строка
'               LEFTTAG - левый (открывающий)  тэг
'               RIGHTTAG - правый (закрывающий)  тэг
' RETURNS:      массив структурированных строк вида (*)
' (*) nPosInSourceStr||nParentSubStr||sSubstrWithoutTAGS, где
'   nPosInSourceStr - позиция строки в исходной,nParentSubStr - ссылка на номер родительской строки, куда вложена данная
'   sSubstrWithoutTAGS - вложенная строка без соотв. тегов
'============================================================================================================================
Public Function GetSubstringTAGSArray(sSource As String, LeftTag As String, RightTag As String, _
                                          Optional iCompare As VbCompareMethod, Optional sDelim As String = "||") As String()
Dim NESTEDS() As String, UDim As Long, I As Long        ' Массив вложенных строк, его размерность и циклическая переменная
Dim LeftLen As Integer, iL As Long                      ' Длины левого тега и текущий указатель
Dim iPosLefts() As Long, iPosRights() As Long           ' Все позиции левых и правых тэгов
Dim nLefts As Long, nRights As Long                     ' Соответсвующие размерности массивов позиций левого и правого тегов
Dim NOLeft As Long, NORoot As Long                      ' Номер текущего левого тега и номер ссылки на родителя
Dim bDirect As Boolean                                  ' Флаг, указывающий направление сортировки узлов
Dim curRight As Long, curLeft As Long, nextLeft As Long ' Текущие указатели левого. след левого и правого тегов
On Error GoTo GetSubstringTAGSArray_ERR
'---------------------------------------------------------------------
    ReDim NESTEDS(0): UDim = 0                       ' Формируем выходный массив
    NESTEDS(0) = 1 & sDelim & 0 & sDelim & sSource
    If LeftTag = RightTag Then GoTo GetSubstringTAGSArray_EXIT
'----------------------------------------------------------------------
' Собираем позиции все тегов. Сначала левых, затем правых.
    iL = InStrArray(1, sSource, LeftTag, iPosLefts, iCompare)
    nLefts = iL - 1
    If iL > 0 Then   ' Нашли левые теги, переходим к выделению правых тегов
       iL = InStrArray(1, sSource, RightTag, iPosRights, iCompare)
       nRights = iL - 1
    End If
    If iL = 0 Then GoTo GetSubstringTAGSArray_EXIT
'---------------------------------------------------------------------
' Готовимся к вычислениям
    LeftLen = Len(LeftTag)                            ' Длины левог тэга, необходимая для вычленения строк
    NORoot = 0                                        ' Номер уровня текущего тэга
    bDirect = True                                    ' Прямая сортировка вначале
'----------------------------------------------------------------------------------------------------------
' Осуществляем проход по правым (закрывающим тэгам), для каждого из них восстанавливаем левы тэг (при
' возможности) за два прохода
    NOLeft = 0
    For I = 0 To nRights ' Идем по всем правым
                    curRight = iPosRights(I) ' текущий правый
      '-----------------------------------------------------&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
      If bDirect Then                                              ' &&&&&&&&&&&&&&&&&&&&     ПРЯМАЯ СОРТИРОВКА
        Do While NOLeft <= nLefts
            curLeft = iPosLefts(NOLeft)
            '------------------------------------------
            If NOLeft < nLefts Then   ' До предпоследнего узла смотрим его и следующий левый
                nextLeft = iPosLefts(NOLeft + 1)
                '--------------------------
                If curRight < curLeft Then                                             '(ПРАВЫЙ ЛЕВЕЕ ЛЕВОГО)
                    Exit Do           ' Выходим, чтобы перейти к следующему правому
                ElseIf curRight > curLeft And curRight < nextLeft Then                 '(ПРАВЫЙ МЕЖДУ) ! пара
                    UDim = UDim + 1: ReDim Preserve NESTEDS(UDim)
                    iL = iPosLefts(NOLeft) + LeftLen
                    NESTEDS(UDim) = iL & sDelim & NORoot & sDelim & _
                                    Mid(sSource, iL, iPosRights(I) - iL)
                    iPosLefts(NOLeft) = 0  ' сбрасываем текущий левый, чтобы дальше его не рассматривать
                    NOLeft = NOLeft + 1
                    Exit Do                ' переходим к следующем правому
                Else                                                                   '(ПРАВЫЙ СЛИШКОМ ПРАВЕЕ)
                    NOLeft = NOLeft + 1    ' пропускаем текущий левый
                    NORoot = NORoot + 1    ' повышаем уровень вложенности
                End If
            Else                                                                        ' ПОСЛЕДНИЙ ЛЕВЫЙ
               bDirect = False             ' Закрываем прямую сортировку
               If curRight > curLeft Then
                    UDim = UDim + 1: ReDim Preserve NESTEDS(UDim)
                    iL = iPosLefts(NOLeft) + LeftLen
                    NESTEDS(UDim) = iL & sDelim & NORoot & sDelim & _
                                    Mid(sSource, iL, iPosRights(I) - iL)
                    iPosLefts(NOLeft) = 0  ' сбрасываем текущий левый, чтобы дальше его не рассматривать
                    Exit Do
               Else                        ' пропускаем последний левый
                    NORoot = NORoot + 1    ' повышаем уровень вложенности
               End If
            End If
        Loop
      '-----------------------------------------------------&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
      Else                                                         ' &&&&&&&&&&&&&&&&&&&&     ОБРАТНАЯ СОРТИРОВКА
            NOLeft = nLefts
            Do While NOLeft >= 0
               curLeft = iPosLefts(NOLeft)
               If curLeft = 0 Then               ' игнорируем левый
                  NOLeft = NOLeft - 1            ' сдвигаемся вниз
               Else                              ' обрабатываем текущий левый
                    UDim = UDim + 1: ReDim Preserve NESTEDS(UDim)
                    iL = iPosLefts(NOLeft) + LeftLen
                    NORoot = NORoot - 1    ' повышаем уровень
                    NESTEDS(UDim) = iL & sDelim & NORoot & sDelim & _
                                    Mid(sSource, iL, iPosRights(I) - iL)
                    iPosLefts(NOLeft) = 0  ' сбрасываем текущий левый, чтобы дальше его не рассматривать
                    Exit Do                ' переходим к следующем правому
               End If
            Loop
      End If
      '----------------------------------------------------
     Next I
'----------------------------------------------------------------------
GetSubstringTAGSArray_EXIT:
      GetSubstringTAGSArray = NESTEDS     '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
      Exit Function
'----------------------------------------------------------------------
GetSubstringTAGSArray_ERR:
      If Err.Number > 0 Then
        MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "GetSubstringTAGSArray ERROR"
      End If
End Function
'===============================================================================================================================
' Функция возвращает позицию заданной подстроки слева от второй подстроки внутри данной строки.
' Возвращается первое вхождение. если базовая подстрока, относительно которой ищется вхождение
' Не присутсвует в строке или не найдена искомая строка - возвращается 0
' PARAMETERS: iStart    - позиция, начиная с которой ищется sBase
'             sSource   - исходная строка, в которой осуществляется весь поиск
'             sBase     - строка- якорь, относительно которой ищутся остальные строки
'             sFindLeft - подстрока, которая ищется левее sBase
'             iCompare  - сравнение (c extnjv htubcnhf bkb ,tp)
' RETURNS:    позицию первого вхождения sFindLeft в sSource левее sFindLeft/ Если sBase или  sFindLeft не найдены - 0
'==============================================================================================================================
Public Function InStrLeft(iStart As Long, sSource As String, sBase As String, sFindLeft As String, _
                                                                                  Optional iCompare As VbCompareMethod) As Long
Dim iRes As Long                ' Возвращаемый результат

'-------------------------------------------------------------------
' Ищем базовую подстроку
iRes = InStr(iStart, sSource, sBase, iCompare)
If iRes = 0 Then GoTo InStrLeft_EXIT
'------------------------------------------------------------------
' Ищем левую подстроку (слева)
iRes = InStrRev(sSource, sFindLeft, iRes, iCompare)
'-----------------------------------------
InStrLeft_EXIT:
            InStrLeft = iRes    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'===============================================================================================================================
' Функция возвращает позицию заданной подстроки слева от второй подстроки внутри данной строки.
' Возвращается первое вхождение. если базовая подстрока, относительно которой ищется вхождение
' Не присутсвует в строке или не найдена искомая строка - возвращается 0
' PARAMETERS: iStart    - позиция, начиная с которой ищется sBase
'             sSource   - исходная строка, в которой осуществляется весь поиск
'             sBase     - строка- якорь, относительно которой ищутся остальные строки
'             sFindLeft - подстрока, которая ищется левее sBase
'             iCompare  - сравнение (c extnjv htubcnhf bkb ,tp)
' RETURNS:    позицию первого вхождения sFindLeft в sSource левее sFindLeft/ Если sBase или  sFindLeft не найдены - 0
'==============================================================================================================================
Public Function InStrRight(iStart As Long, sSource As String, sBase As String, sFindRight As String, _
                                                                                  Optional iCompare As VbCompareMethod) As Long
Dim iRes As Long                ' Возвращаемый результат

'-------------------------------------------------------------------
' Ищем базовую подстроку
iRes = InStr(iStart, sSource, sBase, iCompare)
If iRes = 0 Then GoTo InStrRight_EXIT
'------------------------------------------------------------------
' Ищем левую подстроку (слева)
iRes = InStr(iRes, sSource, sFindRight, iCompare)
'-----------------------------------------
InStrRight_EXIT:
            InStrRight = iRes    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'=====================================================================================================
' Функция парсит входящую строку состоящую из подстрок с сепараторами, извлекая из нее подстроку с
' индексом (просто облочка к SPLIT) - 1- based index
'=====================================================================================================
Public Function GetSubstring(sSource As String, SEP As String, indx As Long) As String
On Error GoTo zGetSubstring_ERR
    GetSubstring = Split(sSource, SEP)(indx - 1)
    Exit Function
zGetSubstring_ERR:
    Err.Clear           ' очищаем список ошибок
    GetSubstring = ""   ' В случае ошибки (например вышли за пределы массива) возвращаем нулевую строку
End Function
'===============================================================================================================================
' Функция парсит входящую строку состоящую из подстрок с сепараторами и находит индекс для подстроки
' в которую входит некоторый заданный текст. если текст не найден - возврат "-1"
' Предполагается, что идентифицирующий шаблон входит в общую строку единожды - 1-based index
'===============================================================================================================================
Public Function GetSubstrINDX(sSource As String, SEP As String, sMatch As String) As Long
Dim iPos As Long            ' Позиция поискового шаблона внутри искомой лексемы
Dim iSepLeft As Long        ' Ближайшая слева от найденного шаблона позиция сепаратора
Dim iSepRight As Long       ' Ближайшая слева от найденного шаблона позиция сепаратора
Dim indx As Long            ' искомый индекс

'-----------------------------------------------------------------------------------------
' Инициализируем значение
indx = -1
'-----------------------------------------------------------------------------------------
' Начинаем поиск
     iPos = InStr(1, sSource, sMatch, vbTextCompare)   ' Вычисляем позицию вхождения
     If iPos > 0 Then             ' Вхождение найдено, переходим к вычислению индекса
        iSepLeft = InStrRev(sSource, SEP, iPos, vbTextCompare)
        iSepRight = InStr(iPos, sSource, SEP, vbTextCompare)
        
        If iSepRight > 0 And iSepLeft = 0 Then               ' Искомая лексема раполагается первой
           indx = 1   ' Первая лексема
        ElseIf iSepLeft > 0 And iSepRight = 0 Then           ' Искомая лексема - последняя
           indx = CountSubstrAsWord(Left(sSource, iSepLeft), SEP) + 1
        Else                                                 ' Искомая лексема где-то посередине
           ' Вычисляем число вхождений сепаратора в подстроке слева и добавляем 1
           indx = CountSubstrAsWord(Left(sSource, iSepRight - 1), SEP) + 1
        End If
     Else                         ' Вхождение ненайдено, выходим с возвратом "-1"
       
     End If
     GetSubstrINDX = indx
End Function
'==================================================================================================================================
' Функция, проверяет что в состав строки, входит более NCount подстрок вида sSubstr
' PARAMETERS: sSource  - строка-источник, в которой ищется
'             sSubstr  - искомая подстрока
'             iCompare - флаг, определяющий сравнение
' RETURNS:    True     - если кол-во вхождений более NCount
'==================================================================================================================================
Public Function CountSubstrTillNCount(sSource As String, sSubstr As String, nCount As Long, _
                                                                                   Optional iCompare As VbCompareMethod) As Boolean
Dim lngPos As Long
Dim lngTemp As Long
Dim lngCount As Long
    
    If Len(sSource) = 0 Then Exit Function
    If Len(sSubstr) = 0 Then Exit Function
    lngPos = 1
'-------------------------------------------------------------
    Do
        lngPos = InStr(lngPos, sSource, sSubstr, iCompare)
        lngTemp = lngPos
        If lngPos > 0 Then
            lngCount = lngCount + 1
            '-------------------------------------------------
            If lngCount >= nCount Then
                CountSubstrTillNCount = True '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                Exit Function
            End If
            '-------------------------------------------------
            lngPos = lngPos + Len(sSubstr)
        End If
    Loop Until lngPos = 0
'-----------------------------------------------------------------------
End Function
'===============================================================================================================================
' Функция обрезает входящую строку вплоть до некоторого слова, которое ищется с заданной позиции включительно
' То есть отсекается левая часть. используется первое вхождение. Если подстрока не находится - возвращается таже самая строка
' Если подстрока стоит в самом конце, обрезание возвращает полную строку
'===============================================================================================================================
Public Function TruncateStrTillWord(iStart As Long, sSource As String, sTruncateWord As String, _
                                                                                         Optional iCompare As VbCompareMethod)
Dim iL As Long, LenSource As Long                                 ' Рабочая позиция и новая длина строки
Dim sRes As String                                                ' Возвращаемый результат
'---------------------------------------------------------------
sRes = ""
iL = InStr(iStart, sSource, sTruncateWord, iCompare)
    If iL > 0 Then
        LenSource = iL + Len(sTruncateWord)
        sRes = Left(sSource, LenSource)
    Else
        sRes = sSource
    End If
'--------------------------------------------------------------
        TruncateStrTillWord = sRes      '!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'===============================================================================================================================
' Функция обрезает входящую строку, начиная с некоторого слова, которое ищется с заданной позиции включительно
' То есть отсекается левая часть. используется первое вхождение. Если подстрока не находится - возвращается таже самая строка
' Если подстрока стоит в самом конце, обрезание возвращает пустую строку
'===============================================================================================================================
Public Function TruncateStrFromWord(iStart As Long, sSource As String, sTruncateWord As String, _
                                                                                         Optional iCompare As VbCompareMethod)
Dim iL As Long, LenSource As Long                                 ' Рабочая позиция и новая длина строки
Dim sRes As String                                                ' Возвращаемый результат
'---------------------------------------------------------------
sRes = ""
iL = InStr(iStart, sSource, sTruncateWord, iCompare)
    If iL > 0 Then
        LenSource = iL + Len(sTruncateWord)
        If LenSource < Len(sSource) Then sRes = Right(sSource, Len(sSource) - LenSource)
    Else
        sRes = sSource
    End If
'--------------------------------------------------------------
        TruncateStrFromWord = sRes      '!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'===============================================================================================================================
' Функция проверяет, что данная подстрока, находящаяся в строке лежит внутри пары не пересекающихся тегов LeftTag и RightTag
' Иными словами, находится внутри скобок. При этом скобки могут быть одинаковыми, например, это может быть закавыченная
' строка. Функция проверяет, что такие тэги образуют пару вокруг заданной подстроки. Подстрока задается парой значений -
' начало и длина. Функция работает с любой, а не только с правильной скобочной последовательностью.
' PARAMETERS: sSource  - исходная строка, в которой ищется вхождение
'             iStart   - начало вхождения подстроки
'             iLen     - длина вхождения подстроки
'             LeftTag  - левый тэг
'             RightTag - правый тэг
'             iCompare - метод сравнения, применяемый к поиску тэгов
' RETURNS:    Если тэги найдены  - срока "iLeft;iRight". Если не найдены (свободная подстрока) - пустая строка
'===============================================================================================================================
    Public Function GetTagsContainedSubString(sSource As String, iStart As Long, iLen As Long, LeftTag As String, RightTag As String, _
                                                                                     Optional iCompare As VbCompareMethod) As String
    Dim iCount As Integer, iL As Long, iR As Long             ' Счетчик, левый текущий указатель, правый текущий указатель
    Dim iLeft As Long, iRight As Long, sRes As String         ' Найденные результаты
    
    '-----------------------------------------------------
    ' Инициализируемся
    iLeft = 0: iRight = 0: sRes = ""
    '-----------------------------------------------------
    ' Поиск скобок вперед
    iCount = 0: iL = iStart + iLen - 1: iR = iL     ' Поиск левого начиная с самого конца подстроки и
                                                    ' правого, начиная сразу после подстроки
    Do While iR > 0
       iR = InStr(iR + 1, sSource, RightTag, iCompare)
       If iR > 0 Then iCount = iCount - 1
       Do While iL < iR And iL > 0
         iL = InStr(iL + 1, sSource, LeftTag, iCompare)
         If iL < iR And iL > 0 Then iCount = iCount + 1
       Loop
         If iCount = -1 Then Exit Do
    Loop
        iRight = iR
    '-----------------------------------------------------
    ' Поиск скобок назад
    iCount = 0: iL = iStart: iR = iL
    Do While iL > 0
        iL = InStrRev(sSource, LeftTag, iL - 1, iCompare)
        If iL > 0 Then iCount = iCount + 1
        Do While iR > iL And iR > 0
            iR = InStrRev(sSource, RightTag, iR - 1, iCompare)
            If iR > 0 And iR > iL Then iCount = iCount - 1
        Loop
          If iCount = 1 Then Exit Do
    Loop
       iLeft = iL
    '-----------------------------------------------------
    ' Формируем выход
    If iLeft > 0 And iRight > 0 Then sRes = iLeft & ";" & iRight
    '---------------------------------------------------
    ' Выход функции
    GetTagsContainedSubString = sRes    '!!!!!!!!!!!!!!!!!!!
End Function
'-----------------------------------------------------------------------------------------------------------------------
' Функция создает скобочные последовательности для заданной строки так, чтобы она была в общем сбалансирована
'-----------------------------------------------------------------------------------------------------------------------
Public Function GetBrackets(str As String, _
                                                  Optional LeftTag As String = "(", Optional RightTag As String = ")", _
                                            Optional iCompare As VbCompareMethod, Optional DLM As String = ";") As String
Dim lefts() As Long, nDim As Integer, I As Integer                   ' массив положений левых тегов
Dim rights() As Long, mDim As Integer, J As Integer                  ' массив положений правых тегов
Dim sWork As String, sRes() As String, kDim As Integer               ' рабочая переменная, выходной массив и его разм
Dim sDelim As String


sDelim = IIf(DLM = "", Chr(2), DLM)
lefts = GetInStrArray(str, LeftTag, iCompare):   nDim = UBound(lefts)
rights = GetInStrArray(str, RightTag, iCompare): mDim = UBound(rights)

If nDim <> mDim Then Exit Function              ' Первая элементарная проверка, позволяющая предположить, что строка
                                                ' в целом сбалансирована (лучше проверять полный баланс)
ReDim sRes(0): kDim = -1
'---------------------------------------------------------------------------------------------------------
For I = 0 To nDim                               ' Идем по левому тегу, проверяя для него каждый правый
     For J = 0 To mDim                          ' - пока не находим баланс
         If lefts(I) < rights(J) Then
            sWork = Mid(str, lefts(I) + 1, rights(J) - lefts(I) - 1)
            If StringBalance(sWork, LeftTag, RightTag) = 0 Then
               kDim = kDim + 1: ReDim Preserve sRes(kDim)
               sRes(kDim) = sWork
            End If
         End If
     Next J
Next I
'---------------------------------------------------------------------------------------------------------
    GetBrackets = Join(sRes, sDelim)  '!!!!!!!!!!!!!!!!!!!!!!!
End Function
' Функция проверяет баланс скобок (односимвольных), требует доработки
' Возвращает целое число. При нуле - сбалансировано
'-----------------------------------------------------------------------------------------------------------------
Private Function StringBalance(s As String, _
                                   Optional LeftTag As String = "(", Optional RightTag As String = ")") As Integer
Dim Depth As Integer                             ' Глубина входа
Dim ch As String * 1                             ' Пробный символ
Dim I As Integer                                 ' Циклическая переменная


Depth = 0                                        ' Задаемся глубиной 0
'-----------------------------------------------------------------------
For I = 1 To Len(s)
  ch = Mid$(s, I, 1)
  If ch = LeftTag Then Depth = Depth + 1
  If ch = RightTag Then
    If Depth = 0 Then 'not balanced
      StringBalance = -1
      Exit Function
    Else
      Depth = Depth - 1
    End If
  End If
Next
'------------------------------------------------------------------------
             StringBalance = Depth '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function

'-----------------------------------------------------------------------------------------------------------------------
' Функция извлекает скобки для заданной последовательности
'-----------------------------------------------------------------------------------------------------------------------
Public Function GetParen(strIn As String, sPattern As String) As String
    Dim objRegex As Object
    Dim objRegMC As Object
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
        '.Pattern = "\((.+?)\)"
        .Pattern = sPattern
        
        '.Pattern = "/(\{(?:\{.*\}|[^\{])*\})/m"
        If .Test(strIn) Then
            Set objRegMC = .Execute(strIn)
            GetParen = objRegMC(0).submatches(0)
        Else
            GetParen = "No match"
        End If
    End With
    Set objRegex = Nothing
End Function
'===============================================================================================================================
' Функция проверяет закавыченность подстроки. Иными словами, если в строке присутсвуют одинаковые тэги (левый и правый)
' и открытость исчисляется слева, функция подсчитывает, находится ли подстрока, заданная своей позицией и длиной внутри пары
' одинаковых тегов, то есть между нечетным и четным тегом, считая слева
'===============================================================================================================================
Public Function IsSubStringInQuotes(sSource As String, iStart As Long, iLen As Long, Optional sTag As String = """", _
                                                                                 Optional iCompare As VbCompareMethod) As Boolean
Dim iCount As Integer, iL As Integer
Dim bRes As Boolean

iL = 1: iCount = 1         ' Ищем с самого начала входящей строки. обнуляем счетчик
'---------------------------------------------------------------------------------
' Рассматриваем отрезок меньший подстроки
Do While (iL > 0 And iL < iStart)
    iL = InStr(iL, sSource, sTag, iCompare)
    If iL > 0 And iL < iStart Then
       iCount = iCount * (-1)
       iL = iL + 1
    End If
Loop
'-------------------------------------------------------------------------------
If iCount > 0 Then Exit Function  ' Все кавычки закрыты, возвращаем False
'-------------------------------------------------------------------------------
' Рассматриваем второй отрезок, чтобы убедиться, в закрытии
                  iL = InStr(iStart + iLen, sSource, sTag, iCompare)
                  If iL >= iStart + iLen Then bRes = True
'-----------------------------------------------------
            IsSubStringInQuotes = bRes    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'===============================================================================================================================
' Функция расширяет функцию Split за счет исключения из знаков разделения тех, которые содержатся в ""
'===============================================================================================================================
Public Function SplitXCludeQuates(sSource As String, sDelim As String, Optional sXClude As String = """", _
                                                                              Optional iCompare As VbCompareMethod) As String()
Dim iL As Long, iLen As Long                                    ' текущие указатель и длина
Dim sWork As String, lDelim As String                           ' рабочая строка, временный разделитель

'------------------------------------------------
' ИНИЦИАЛИЗАЦИЯ И ПРОВЕРКИ
       sWork = sSource: lDelim = sDelim
    If Trim(sWork) = "" Or sDelim = "" Then GoTo SplitXCludeQuates_EXIT           ' Вхождение пустой строки
    If InStr(1, sWork, sDelim, iCompare) = 0 Then GoTo SplitXCludeQuates_EXIT     ' В строку не входит разделитель
    If InStr(1, sWork, sXClude, iCompare) = 0 Then GoTo SplitXCludeQuates_EXIT  ' Исключать нечего
    '-------------------------------------------
        iLen = Len(sDelim): iL = 1 - iLen
        lDelim = "<!&>6s"                                         ' Временный разделитель
'-----------------------------------------------------------------------------------------------------------
' Идем по всем разделителям, в случае правильного, наращиваем массив за счет предыдущей строки
Do
      iL = InStr(iL + iLen, sWork, sDelim, iCompare)
      '---------------------------------------------------------------------------------------
      If iL > 0 Then
         '------------------------------------------------------------
         If Not IsSubStringInQuotes(sWork, iL, iLen, sXClude, iCompare) Then
                     sWork = ReplaceString(sWork, iL, iL + iLen - 1, lDelim)
         End If
      End If
Loop Until iL = 0
'---------------------------------------------------------------------------
SplitXCludeQuates_EXIT:
    SplitXCludeQuates = Split(sWork, lDelim) '!!!!!!!!!!!!!!!!!!!!!!!
End Function
'===============================================================================================================================
' Функция проверяет, что ни одно из стоп-слов не входит в заданную строку. Если обнаруживается слово (именно слово)
' из стоп-листа,
' возвращается False, если строка свободна от входящих стоп- слов - возвращается True
' Используется функция ListInWord с такими же параметрами
' PARAMETERS: sStr     - входная строка, в которой ищется стоп-слово
'             sSTOPs   - список стоп-слов, разделенных sDelim (по умолчанию ";")
'             iCompare - метод сравнения (по умолчанию с учетом регистра)
'             sDelim   - разделитель в списке стоп-слов (по умолчанию ";")
'             IsAsWord - рассматривать вхождение как отдельное слово
' RETURNS:    True - если в строке не содержится стоп-слова, False - если содержится
'===============================================================================================================================
Public Function IsStrFreeOfSTOPS(sSTR As String, sSTOPs As String, Optional iCompare As VbCompareMethod, _
                                              Optional sDelim As String = ";", Optional IsAsWord As Boolean = True) As Boolean
Dim MyWords() As String, nDim As Integer      ' Список слов и его размерность
Dim iL As Long, I As Integer                  ' Циклическая переменная и указатель
'---------------------------------------------------------------------------
' Инициализируем функцию и проверяем входящие
If sSTR = "" Or sSTOPs = "" Then GoTo IsStrFreeOfSTOPS_EXIT
'--------------------------------------------------------
' Выделяем слова из списка
MyWords = Split(sSTOPs, sDelim): nDim = UBound(MyWords)
'---------------------------------------------------------
For I = 0 To nDim
    iL = InStr(1, sSTR, MyWords(I), iCompare)
    If iL = 0 Then GoTo NextWord       ' Обнаружили вхождение слова.
         '----------------------------------------------------------------------
            If IsAsWord Then           ' Проверяем, что слово отдельное
               If IsInThisPosWord(sSTR, iL, Len(MyWords(I))) Then Exit Function
            Else                       ' Сразу выходим с отрицательным результатом
               Exit Function
            End If
         '----------------------------------------------------------------------
NextWord:
Next I
'----------------------------------------------------------
' Выход
IsStrFreeOfSTOPS_EXIT:
  IsStrFreeOfSTOPS = True   '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function

'==============================================================================================================================
' Функция находит слово из списка sWORDS в строке sStr, как отдельное слово. Если слово найдено, то оно и возвращается.
' Если не найдено ни одного слова, то возвращается пустая строка.
' PARAMETERS: sStr     - входная строка, в которой ищется стоп-слово
'             sLIST    - список слов, разделенных sDelim (по умолчанию ";"), каждое из которых ищется в строке
'             iCompare - метод сравнения (по умолчанию с учетом регистра)
'             sDelim   - разделитель в списке стоп-слов (по умолчанию ";")
' RETURNS:    True - если в строке не содержится стоп-слова, False - если содержится
'==============================================================================================================================
Public Function ListInWord(sSTR As String, sList As String, Optional iCompare As VbCompareMethod, _
                                                                                       Optional sDelim As String = ";") As String
Dim MyWords() As String, nDim As Integer      ' Список слов и его размерность
Dim I As Integer, sRes As String             ' Циклическая переменная и ожидаемый результат

On Error GoTo ListInWord_ERR

sRes = ""
'---------------------------------------------------------------------------
' Получаем рабочий массив и входим в цикл
    MyWords = Split(sList, sDelim): nDim = UBound(MyWords)
    For I = 0 To nDim                         ' Идем по циклу  заданного списка слов
          If InStrAsWord(1, sSTR, MyWords(I), iCompare) Then
                      sRes = MyWords(I)            ' НАШЛИ СЛОВО СПИСКА
                      Exit For                     ' выходим из цикла
          End If
    Next I
'----------------------------------------------------------
' Выход
ListInWord_EXIT:
  ListInWord = sRes   '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  Exit Function
'---------------------------------------------------------
ListInWord_ERR:
  MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "ListInWord ERROR"
  GoTo ListInWord_EXIT
End Function

'===============================================================================================================================
' Функция InStrAsWord является заменой стандартной функции InStr и позволяет искать вхождение sSubstr в строку sSource
' начиная с позиции iStart как отдельное слово.
' PARAMETERS:  iStart   - позиция, с которой производится поиск
'              sSource  - исходная строка, в которой ищется
'              sSubstr  - строка, которая ищется
'              iCompare - флаг, который определяет, стоит ли сравнивать с учетом регистра или без
'===============================================================================================================================
Public Function InStrAsWord(iStart As Long, sSource As String, sSubstr As String, Optional iCompare As VbCompareMethod) As Long
Dim iPos As Long                         ' Рабочая позиция / возвращаемый результат
Dim lLen As Long                         ' Длина подстроки
'-----------------------------------------------
lLen = Len(sSubstr)
iPos = InStr(iStart, sSource, sSubstr, iCompare)      ' Первичное нахождение

Do While iPos > 0
    If IsInThisPosWord(sSource, iPos, lLen) Then Exit Do ' Нашли слово
    iPos = InStr(iPos + lLen, sSource, sSubstr, iCompare)
Loop
'------------------------------------------------------------------------------------------------------
     InStrAsWord = iPos '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'==================================================================================================================================
' Функция подсчитывает вхождение строки sFind в строку sSource без учета выделенных слов (быстрый вариант)
'==================================================================================================================================
Public Function CountSubstr(sSource As String, sFind As String, Optional iCompare As VbCompareMethod) As Long
' Counts occurrences of a particular character or characters.
' If lngCompare argument is omitted, procedure performs binary comparison.
' Пример:
' ?StringCountOccurrences("","") = 0
' ?StringCountOccurrences("","a") = 0
' ?StringCountOccurrences("aaa","a") = 3
' ?StringCountOccurrences("aaa","b") = 0
' ?StringCountOccurrences("aaa","aa") = 1

Dim lngPos As Long
Dim lngTemp As Long
Dim lngCount As Long
    
    If Len(sSource) = 0 Then Exit Function
    If Len(sFind) = 0 Then Exit Function
    lngPos = 1
'-------------------------------------------------------------
    Do
        lngPos = InStr(lngPos, sSource, sFind, iCompare)
        lngTemp = lngPos
        If lngPos > 0 Then
            lngCount = lngCount + 1
            lngPos = lngPos + Len(sSource)
        End If
    Loop Until lngPos = 0
'-----------------------------------------------------------------------
    CountSubstr = lngCount
End Function
'=====================================================================================================
' Функция считает кол-во вхождений подстроки substr в строку str
' Parameters:
'           sSource - исходная строка, в которой считаются подстроки
'           sSubstr - считаемые подстроки
'           bCountAsWord - считать ли вхождение как отдельного слова (слева и справа пробелы,
'                           непечатные символы или знаки препинания)
' RETURNS: кол-во вхождений с учетом bCountAsWord
'=====================================================================================================
Public Function CountSubstrAsWord(sSource As String, sSubstr As String, _
                                                   Optional bCountAsWord As Boolean = False) As Long
Dim lCount As Long                                ' Рабочий счетчик
Dim iStart As Long                                ' Найденная позиция подстроки
Dim iLen As Long, iLenSub As Long                 ' Длина строки поиска
Dim iCharLeft As Long                             ' Ограничивающий слово(подстроку) символ слева
Dim iCharRight As Long                            ' Ограничивающий слово (подстроку) символ справа

lCount = 0                  ' Инициализируем счетчик
iLen = Len(sSource)         ' Вычисляем длину строки, в которой осуществляем поиск
iLenSub = Len(sSubstr)      ' Вычисляем длину подстроки
iStart = 0                  ' Начинаем искать с первого символа (0 с учетом последующего сдвига)

Do
  iStart = iStart + 1                                      ' Сдвигаемся
  iStart = InStr(iStart, sSource, sSubstr, vbTextCompare)  ' Находим очередное вхождение подстроки
  '-------------------------------------------------------------
  If iStart > 0 Then        ' Что-то нашли
     '---------------------------------------------------------
     If bCountAsWord Then   ' В этом случае считаем как слово - то есть проверяем символы слева и справа
         If GetCharASCCategory(Mid(sSource, iStart - 1, 1)) < 3 Then  ' Слева содержится непечатаемый знак или знак препинания
              If iStart + iLenSub > iLen Then
                  lCount = lCount + 1     ' Подтвердили признак слова (достигли конца строки)
              ElseIf GetCharASCCategory(Mid(sSource, iStart + iLenSub, 1)) < 3 Then ' Справа также непечатаемый символ или знак препинания
                  lCount = lCount + 1     ' Подтвердили признак слова
              End If
         End If
     Else                   ' Не считаем как слово, поэтому сразу наращиваем счетчик
                 lCount = lCount + 1
     End If
     '---------------------------------------------------------
  End If
Loop While iStart > 0 And iStart < iLen                    ' Двигаемся пока находим результаты и не достигли границы

'--------------------------------------------------------------------------------
CountSubstrAsWord = lCount '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'===============================================================================================================================
' Функция расширяет имеющуюся функцию InStr  и возвращает не пощицию вхождения, а число вхождений и массив  sArrayвсех вхождений
' строки sFIND в строку sSource, начиная с позиции iStart.
' PARAMETERS:
'===============================================================================================================================
Public Function InStrArray(iStart As Long, sSource As String, sFind As String, sArray() As Long, _
                                                                                  Optional iCompare As VbCompareMethod) As Long
Dim iL As Long, nCount As Long                          ' Найденное кол-во вхождений и текущий указатель

'-----------------------------------------------------------------------
nCount = -1: ReDim sArray(0)                    ' Инициализируемся в 0 с удалением старой информации
iL = iStart                                     ' Начальное положение указателя в строке
'------------------------------------------------------------------------
' Начинаем поиск
    Do While iL > 0
       iL = InStr(iL, sSource, sFind, iCompare)
       If iL > 0 Then
            nCount = nCount + 1
            If nCount > UBound(sArray) Then ReDim Preserve sArray(nCount)
            sArray(nCount) = iL
            iL = iL + 1
       End If
    Loop
'------------------------------------------
InStrArray_EXIT:
    InStrArray = nCount + 1           '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'==============================================================================================================================
' Функция склеивает две подстроки в одну с использованием разделителя. Если одна из строк "", то возвращается только
' вторая строка. Если обе строки пустые, то возвращается пустая строк
'==============================================================================================================================
Public Function GlueString(str1 As String, str2 As String, Optional sDelim As String = ";", _
                                                Optional bNoRepeat As Boolean, Optional iCompare As VbCompareMethod) As String
Dim sRes As String
'-----------------------------------------------------
    If str1 = "" Then
                sRes = str2: GoTo GlueString_EXIT
    ElseIf str2 = "" Then
                sRes = str1: GoTo GlueString_EXIT
    End If
'-----------------------------------------------------
' Проверяем, что присоединяемая строка не является списком
If InStr(1, str2, sDelim) Then  ' Требуется дополнительная обработка
    sRes = ListMerge(str1, str2, sDelim, iCompare)
Else                            ' Работаем, как с обычной лексемой
        If bNoRepeat And InStr(1, str1, str2, iCompare) > 0 Then ' Проверяем, чтобы при присоединении не было повторов
                            sRes = str1
        Else
                            sRes = str1 & sDelim & str2
        End If
End If
'-------------------------------------------------
GlueString_EXIT:
        GlueString = sRes   '!!!!!!!!!!!!!!!!!!!!
End Function
'==============================================================================================================================
' Функция совмещает два списка с одинаковыми разделителями. При этом, если подстроки из второго уже встречаются в
' первом списке, они не добавляются. Возвращается обобщенный список
'==============================================================================================================================
Public Function ListMerge(List1 As String, List2 As String, Optional sDelim As String = ";", _
                                                                              Optional iCompare As VbCompareMethod) As String
Dim myList2() As String, nDim As Long, I As Long             ' Рабочий массив, его размерность и цикл переменная
Dim sRes As String                                           ' Возвращаемый результат

sRes = List1
'---------------------------------------------------------
myList2 = Split(List2, sDelim): nDim = UBound(myList2)
    For I = 0 To nDim
      If InStr(1, List1, myList2(I), iCompare) = 0 Then sRes = sRes & sDelim & myList2(I)
    Next I
'---------------------------------------
    ListMerge = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'==============================================================================================================================
' Функция исключает из списка List1 посдстроки списка List2. Оба списка представляют собой строки
' с одинаковыми разделителями
'==============================================================================================================================
Public Function ListDiff(List1 As String, List2 As String, Optional sDelim As String = ";", _
                                                                              Optional iCompare As VbCompareMethod) As String
Dim myList1() As String, nDim As Long, I As Long             ' Рабочий массив 1, его размерность и цикл переменная
Dim myList2() As String, mDim As Long, J As Long             ' Рабочий массив 2, его размерность и цикл переменная
Dim bRes As Boolean, sRes As String                          ' Флаг совпадения и возвращаемый результат

sRes = ""
'---------------------------------------------------------
myList1 = Split(List1, sDelim): nDim = UBound(myList1)
myList2 = Split(List2, sDelim): mDim = UBound(myList2)
'------------------------
    For I = 0 To nDim
        bRes = False
        For J = 0 To mDim
            If myList1(I) = myList2(J) Then
                bRes = True
                Exit For
            End If
        Next J
      If Not bRes Then sRes = sRes & myList1(I) & sDelim
    Next I
'---------------------------------------
    ListDiff = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'==============================================================================================================================
' Функция исключает из строки sSource любые подстроки из списка ListStr
'==============================================================================================================================
Public Function ListReplace(sSource As String, ListRpl As String, sREPLACE As String, Optional sDelim As String = ";", _
                                                                              Optional iCompare As VbCompareMethod) As String
Dim myList() As String, nDim As Integer, I As Integer
Dim sRes As String, iL As Long

sRes = sSource
myList = Split(ListRpl, sDelim): nDim = UBound(myList)
'--------------------------------------------------------
For I = 0 To nDim
   If InStr(1, sRes, myList(I), iCompare) Then  ' Нашли вхождение, обрабатываем его более корректно
       Do
        iL = InStrAsWord(iL + 1, sRes, myList(I), iCompare)
        If iL > 0 Then sRes = ReplaceString(sRes, iL, iL + Len(myList(I)) - 1, sREPLACE)
       Loop While iL > 0
   End If
Next I
'--------------------------------------------------------
ListReplace = sRes  '!!!!!!!!!!!!!!!!!!!!!!!
End Function
'==============================================================================================================================
' Функция заменяет в строке sSource часть подстроки, ограниченную iStart и iFinish на новую построку sReplace.
'==============================================================================================================================
Public Function ReplaceString(sSource As String, iStart As Long, iFinish As Long, sREPLACE As String) As String
    ReplaceString = Left(sSource, iStart - 1) + sREPLACE + Right(sSource, Len(sSource) - iFinish)
End Function

'************************************************************************************************************************
'************************************************************************************************************************
'************************************************************************************************************************
' PRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATE
' PRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATE
' PRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATE
' PRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATE
' PRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATE
' PRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATEPRIVATE
'=====================================================================================================
' Функция проверяет входящий символ (первый символ входящей строки) и возвращает его категорию
' Данная функция предоставляет упрощенную версию, не работающую с Юникодом и использует ASC - коды
' вместо Win API функций (GetStringType или IsCharAlpha,IsCharAlphaNumeric,IsTextUnicode,FoldString)
' RETURNS:
'         0 - непечатаемый символ (перевод каретки, нулевой символ, табуляция и т.п.)
'         1 - знаки препинания, скобки, математические знаки, офиц разделители (. , ; : ? ! ' - = + > [ ))
'         2 - псевдографика
'         3 - цифры
'         4 - большие латинские буквы
'         5 - малые латинские буквы
'         6 - большие киррилические буквы
'         7 - малые кирриличекие буквы
'         8 - другие знаки (корень, параграф, номер и т.п.)
'=====================================================================================================
Private Function GetCharASCCategory(sChar As String) As Integer
Dim lASCII As Integer           ' Вычисляемый ASC _ символ
Dim iReturn As Integer          ' Возвращаемое значение
lASCII = Asc(Nz(sChar, " "))            ' Вычисляем код входящего символа
'---------------------------------------------------------------------------
' Выбор диапазона ASC - символов
      Select Case lASCII
          Case 0 To 31, 127                             ' Непечатаемые (управляющие знаки)
             iReturn = 0
          Case 32 To 47, 58 To 64, 91 To 96, 123 To 126 ' Пробел, знаки препинания, скобки, мат. знаки, спец. знаки
             iReturn = 1
          Case 48 To 57                                 ' Цифры
             iReturn = 3
          Case 65 To 90                                 ' Большие латинские буквы
             iReturn = 4
          Case 97 To 122                                ' Малые латинские буквы
             iReturn = 5
          Case 128 To 159, 240                          ' Большие русские буквы (cp866)
             iReturn = 6
          Case 160 To 175, 224 To 239, 241              ' Малые киррилические буквы (cp866)
             iReturn = 7
          Case 176 To 223                               ' Псевдографика
             iReturn = 2
          Case Else                                     ' Иные знаки (типа № ¤ °)
             iReturn = 8
      End Select
GetCharASCCategory = iReturn                             ' Возвращаем категорию !!!!!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------
' Функция проверяет, что на указанной позиции в слове находится непечатный символ или знак препинания,
' пробле и прочее. Корректно обрабатывается ситуация, указатель за пределами слова
'------------------------------------------------------------------------------------------------------------------
Private Function IsCharBlank(sSTR As String, iR As Long) As Boolean
Dim bRes As Boolean
Dim iCategory As Integer
  If iR > Len(sSTR) Or iR < 1 Then
                bRes = True
  Else
    iCategory = GetCharASCCategory(Mid(sSTR, iR, 1))
    If iCategory < 3 Or iCategory = 8 Then bRes = True       '! Не печатаемый символ или знак препинания, разделитель
  End If
'----------------------------------------
IsCharBlank = bRes        '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------
' Функция проверяет, что в строке str в позиции iPos стоит отдельное слово с длиной iLen
'------------------------------------------------------------------------------------------------------------------
Private Function IsInThisPosWord(str As String, iPos As Long, iLen As Long) As Boolean
Dim sChar As String, iR As Long, bRes As Boolean
Const SEPS = " .,'!?~^&*()[]{}<>-+=/\|" & vbTab & vbCrLf    ' Сепараторы, отделяющие отдельное слово

If iPos > 1 Then
    sChar = Mid(str, iPos - 1, 1)
    If InStr(1, SEPS, sChar) = 0 Then Exit Function
End If
'--------------------
iR = iPos + iLen
If iR <= Len(str) Then
    sChar = Mid(str, iR, 1)
    If InStr(1, SEPS, sChar) = 0 Then Exit Function
End If
'-------------------------------------------------
        IsInThisPosWord = True '!!!!!!!!!!!!!!!!!!!!!!
End Function




