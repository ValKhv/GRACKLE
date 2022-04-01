VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****************************************************************************************************************************************************************
'****************************************************************************************************************************************************************
'****************************************************************************************************************************************************************
'****************************************************************************************************************************************************************
' DICTIONARY CLASS INCAPSULATION
'****************************************************************************************************************************************************************
'****************************************************************************************************************************************************************
'****************************************************************************************************************************************************************
'****************************************************************************************************************************************************************
Option Explicit

#Const UseScriptingDictionaryIfAvailable = True
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    ' KeyValue 0: FormattedKey, 1: OriginalKey, 2: Value
    Private pKeyValues As Collection
    Private pKeys() As Variant
    Private pItems() As Variant
    Private pCompareMode As CompareMethod
#Else
    Private dict As Object
#End If



Public Enum CompareMethod
    BinaryCompare = vbBinaryCompare
    TextCompare = vbTextCompare
    DatabaseCompare = vbDatabaseCompare
End Enum

Const MOD_NAME As String = "cDictionary"
'****************************************************************************************************************************************************************
Public Property Get CompareMode() As CompareMethod
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
        CompareMode = pCompareMode
#Else
        CompareMode = dict.CompareMode
#End If
End Property
Public Property Let CompareMode(value As CompareMethod)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
        If Me.Count > 0 Then
            ' Can't change CompareMode for Dictionary that contains data
            ' http://msdn.microsoft.com/en-us/library/office/gg278481(v=office.15).aspx
            Err.Raise 5 ' Invalid procedure call or argument
        End If
        pCompareMode = value
#Else
        dict.CompareMode = value
#End If
End Property

Public Property Get Count() As Long
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Count = pKeyValues.Count
#Else
    Count = dict.Count
#End If
End Property

Public Property Get Item(key As Variant) As Variant
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Dim KeyValue As Variant
    KeyValue = GetKeyValue(key)
    
    If Not IsEmpty(KeyValue) Then
        If IsObject(KeyValue(2)) Then
            Set Item = KeyValue(2)
        Else
            Item = KeyValue(2)
        End If
    Else
        ' Not found -> Returns Empty
    End If
#Else
    If IsObject(dict.Item(key)) Then
        Set Item = dict.Item(key)
    Else
        Item = dict.Item(key)
    End If
#End If
End Property
Public Property Let Item(key As Variant, value As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.Exists(key) Then
        ReplaceKeyValue GetKeyValue(key), key, value
    Else
        DirectAddKeyValue key, value
    End If
#Else
    dict.Item(key) = value
#End If
End Property
Public Property Set Item(key As Variant, value As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.Exists(key) Then
        ReplaceKeyValue GetKeyValue(key), key, value
    Else
        DirectAddKeyValue key, value
    End If
#Else
    Set dict.Item(key) = value
#End If
End Property

Public Property Let key(Previous As Variant, Updated As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Dim KeyValue As Variant
    KeyValue = GetKeyValue(Previous)
    
    If Not IsEmpty(KeyValue) Then
        ReplaceKeyValue KeyValue, Updated, KeyValue(2)
    End If
#Else
    dict.key(Previous) = Updated
#End If
End Property

Public Sub Add(key As Variant, Item As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Not Me.Exists(key) Then
        DirectAddKeyValue key, Item
    Else
        ' This key is already associated with an element of this collection
        Err.Raise 457
    End If
#Else
    dict.Add key, Item
#End If
End Sub

Public Function Exists(key As Variant) As Boolean
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Exists = Not IsEmpty(GetKeyValue(key))
#Else
    Exists = dict.Exists(key)
#End If
End Function

Public Function Items() As Variant
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.Count > 0 Then
        Items = pItems
    Else
        ' Split("") creates initialized empty array that matches Dictionary Keys and Items
        Items = Split("")
    End If
#Else
    Items = dict.Items
#End If
End Function

Public Function Keys() As Variant
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.Count > 0 Then
        Keys = pKeys
    Else
        ' Split("") creates initialized empty array that matches Dictionary Keys and Items
        Keys = Split("")
    End If
#Else
    Keys = dict.Keys
#End If
End Function

Public Sub Remove(key As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Dim KeyValue As Variant
    KeyValue = GetKeyValue(key)
    
    If Not IsEmpty(KeyValue) Then
        RemoveKeyValue KeyValue
    Else
        ' Application-defined or object-defined error
        Err.Raise 32811
    End If
#Else
    dict.Remove key
#End If
End Sub

Public Sub RemoveAll()
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Set pKeyValues = New Collection
    
    Erase pKeys
    Erase pItems
#Else
    dict.RemoveAll
#End If
End Sub








'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Set pKeyValues = New Collection
    
    Erase pKeys
    Erase pItems
#Else
    Set dict = CreateObject("Scripting.Dictionary")
#End If
End Sub
Private Sub Class_Terminate()
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Set pKeyValues = Nothing
#Else
    Set dict = Nothing
#End If
End Sub


#If Mac Or Not UseScriptingDictionaryIfAvailable Then

Private Function GetKeyValue(key As Variant) As Variant
    On Error Resume Next
    GetKeyValue = pKeyValues(GetFormattedKey(key))
    Err.Clear
End Function

Private Sub DirectAddKeyValue(key As Variant, value As Variant, Optional Index As Long = -1)
    If Me.Count = 0 Then '89
        ReDim pKeys(0 To 0)
        ReDim pItems(0 To 0)
    Else
        ReDim Preserve pKeys(0 To UBound(pKeys) + 1)
        ReDim Preserve pItems(0 To UBound(pItems) + 1)
    End If
    
    Dim FormattedKey As String
    FormattedKey = GetFormattedKey(key)
    
    If Index > 0 And Index <= pKeyValues.Count Then
        Dim I As Long
        For I = UBound(pKeys) To Index Step -1
            pKeys(I) = pKeys(I - 1)
            If IsObject(pItems(I - 1)) Then
                Set pItems(I) = pItems(I - 1)
            Else
                pItems(I) = pItems(I - 1)
            End If
        Next I
        
        pKeys(Index - 1) = key
        If IsObject(value) Then
            Set pItems(Index - 1) = value
        Else
            pItems(Index - 1) = value
        End If
        
        pKeyValues.Add Array(FormattedKey, key, value), FormattedKey, Before:=Index
    Else
        pKeys(UBound(pKeys)) = key
        If IsObject(value) Then
            Set pItems(UBound(pItems)) = value
        Else
            pItems(UBound(pItems)) = value
        End If
        
        pKeyValues.Add Array(FormattedKey, key, value), FormattedKey
    End If
End Sub

Private Sub ReplaceKeyValue(KeyValue As Variant, key As Variant, value As Variant)
    Dim Index As Long
    Dim I As Integer
    
    For I = 0 To UBound(pKeys)
        If pKeys(I) = KeyValue(1) Then
            Index = I + 1
            Exit For
        End If
    Next I
    
    ' Remove existing value
    RemoveKeyValue KeyValue, Index
    
    ' Add new key value back
    DirectAddKeyValue key, value, Index
End Sub

Private Sub RemoveKeyValue(KeyValue As Variant, Optional ByVal Index As Long = -1)
    Dim I As Long
    If Index = -1 Then
        For I = 0 To UBound(pKeys)
            If pKeys(I) = KeyValue(1) Then
                Index = I
            End If
        Next I
    Else
        Index = Index - 1
    End If
    
    If Index >= 0 And Index <= UBound(pKeys) Then
        For I = Index To UBound(pKeys) - 1
            pKeys(I) = pKeys(I + 1)
            
            If IsObject(pItems(I + 1)) Then
                Set pItems(I) = pItems(I + 1)
            Else
                pItems(I) = pItems(I + 1)
            End If
        Next I
        
        If UBound(pKeys) = 0 Then
            Erase pKeys
            Erase pItems
        Else
            ReDim Preserve pKeys(0 To UBound(pKeys) - 1)
            ReDim Preserve pItems(0 To UBound(pItems) - 1)
        End If
    End If
    
    pKeyValues.Remove KeyValue(0)
End Sub

Private Function GetFormattedKey(key As Variant) As String
    GetFormattedKey = CStr(key)
    If Me.CompareMode = CompareMethod.BinaryCompare Then
        ' Collection does not have method of setting key comparison
        ' So case-sensitive keys aren't supported by default
        ' -> Approach: Append lowercase characters to original key
        '    AbC -> AbC__b, abc -> abc__abc, ABC -> ABC
        '    Won't work in very strange cases, but should work for now
        '    AbBb -> AbBb__bb matches AbbB -> AbbB__bb
        Dim Lowercase As String
        Lowercase = ""
        
        Dim I As Integer
        Dim Ascii As Integer
        Dim Char As String
        For I = 1 To Len(GetFormattedKey)
            Char = VBA.Mid$(GetFormattedKey, I, 1)
            Ascii = Asc(Char)
            If Ascii >= 97 And Ascii <= 122 Then
                Lowercase = Lowercase & Char
            End If
        Next I
        
        If Lowercase <> "" Then
            GetFormattedKey = GetFormattedKey & "__" & Lowercase
        End If
    End If
End Function

#End If
'========================================================================================================================
' Add Element
'========================================================================================================================
Public Property Let value(key As String, iValue As Variant)
Attribute value.VB_UserMemId = 0
On Error GoTo ErrHandle
'------------------------------------------
    If dict.Exists(key) Then
           If varType(iValue) = vbObject Then
              Set dict(key) = iValue
           Else
              dict(key) = iValue
           End If
    Else
            dict.Add key, iValue
    End If
'----------------------------
ExitHere:
    Exit Property
'-----------------
ErrHandle:
    ErrPrint2 "Let Value", Err.Number, Err.Description
    Err.Clear
End Property
Public Property Get value(key As String) As Variant

On Error GoTo ErrHandle
'------------------------------------------
    If dict.Exists(key) Then
                  If varType(dict(key)) = vbObject Then
                      Set value = dict(key)
                  Else
                      value = dict(key)
                  End If
                  
    Else
                  value = ""
    End If
'----------------------------
ExitHere:
    Exit Property
'-----------------
ErrHandle:
    ErrPrint "Get Value", Err.Number, Err.Description
    Err.Clear
End Property
'========================================================================================================================
' Extract Dictionary Itself
'========================================================================================================================
Public Function GetDictObject() As Object
     Set GetDictObject = dict
End Function
'========================================================================================================================
' Return Value by Index
'========================================================================================================================
Public Function GetValue(ItemNumber As Integer) As Variant
       GetValue = dict.Items()(ItemNumber)
End Function
'========================================================================================================================
' Return Key By Index
'========================================================================================================================
Public Function GetKey(ItemNumber As Integer) As String
    GetKey = dict.Keys()(ItemNumber)
End Function
'========================================================================================================================
' Return Key Arrays
'========================================================================================================================
Public Function GetKeys() As Variant
        GetKeys = dict.Keys
End Function
'========================================================================================================================
' Return Values Array
'========================================================================================================================
Public Function GetValues() As Variant
    GetValues = dict.Items
End Function
'========================================================================================================================
' Sort By Keys
'========================================================================================================================
Public Sub SortByKey()
    Call SortDictionary(1)
End Sub
'========================================================================================================================
' Sort By Values
'========================================================================================================================
Public Sub SortByValue()
    Call SortDictionary(2)
End Sub

'========================================================================================================================
' Return KeyValue By Key
'========================================================================================================================
Public Function GetKeyValue(sKey As String) As Object
Dim KV As New cKeyValue

On Error GoTo ErrHandle
'--------------------------
    If Not dict.Exists(sKey) Then Err.Raise 1000, , "No Such Key"
    KV.key = sKey
    If varType(dict(sKey)) = vbObject Then
       Set KV.value = dict(sKey)
    Else
       KV.value = dict(sKey)
    End If
'------------------------
ExitHere:
    Set GetKeyValue = KV  '!!!!!!!!!!!!!!
    Exit Function
'---------------
ErrHandle:
    ErrPrint "GetKeyValue", Err.Number, Err.Description
    Err.Clear: Set KV = Nothing
End Function
'========================================================================================================================
' Add KeyValue To Dictionary
'========================================================================================================================
Public Sub SetKeyValue(KV As Object)
Dim sKey As String, sValue As Variant
Dim iType As Integer

On Error GoTo ErrHandle
'--------------------------
    sKey = KV.key: iType = KV.GetType()
    If iType = vbObject Then
       Set sValue = KV.value
    Else
       sValue = KV.value
    End If
    
    If Not dict.Exists(sKey) Then
       dict.Add sKey, sValue
    Else
       If iType = vbObject Then
          Set dict(sKey) = sValue
       Else
          dict(sKey) = sValue
       End If
    End If
'------------------------
ExitHere:
    Exit Sub
'---------------
ErrHandle:
    ErrPrint "SetKeyValue", Err.Number, Err.Description
    Err.Clear: Set KV = Nothing
End Sub
'========================================================================================================================
' Add KeyValue Pair To Dictionary
'========================================================================================================================
Public Sub AddKeyValue(sKV As String, Optional SEQV As String = "=")
Dim sKey As String, sValue As String, KV() As String
    If sKV = "" Then Exit Sub
    
    KV = Split(sKV, SEQV)
    If UBound(KV) < 1 Then
       sKey = KV(0): sValue = ""
    Else
       sKey = KV(0): sValue = KV(1)
    End If
    'sKey = KV(0): sValue = KV(1)
    Add sKey, sValue
End Sub
'========================================================================================================================
' Get KVs Array
'========================================================================================================================
Public Function GetKVArray() As Variant
Dim KVS() As cKeyValue, nDim As Integer, sKey As Variant
On Error GoTo ErrHandle
'-------------------------
    If Me.Count = 0 Then Exit Function
    nDim = -1: ReDim KVS(0)
   
    For Each sKey In dict.Keys
        nDim = nDim + 1: ReDim Preserve KVS(nDim)
        Set KVS(nDim) = New cKeyValue
        KVS(nDim).key = CStr(sKey)
        If varType(dict(sKey)) = vbObject Then
             Set KVS(nDim).value = dict(sKey)
        Else
             KVS(nDim).value = dict(sKey)
        End If
        
    Next sKey
   
'-------------------------
ExitHere:
   GetKVArray = KVS  '!!!!!!!!!!!!!!
   Exit Function
'----------
ErrHandle:
   ErrPrint "GetKVArray", Err.Number, Err.Description
   Err.Clear
End Function
'========================================================================================================================
' Add From KV-String
'========================================================================================================================
Public Sub SetKVString(KVString As String, Optional SEQV As String = "=", Optional DLM As String = ";")
Dim KVS() As String, nDim As Integer, I As Integer

On Error GoTo ErrHandle
'-------------------------
   If KVString = "" Then Exit Sub
   KVS = Split(KVString, DLM): nDim = UBound(KVS)
   
   For I = 0 To nDim
        AddKeyValue KVS(I), SEQV
   Next I
'-------------------------
ExitHere:
   Exit Sub
'----------
ErrHandle:
   ErrPrint "SetKVString", Err.Number, Err.Description
   Err.Clear
End Sub
'========================================================================================================================
' Extract KVString
'========================================================================================================================
Public Function GetKVString(Optional SEQV As String = "=", Optional DLM As String = ";") As String
Dim nDim As Integer, I As Integer
Dim sRes As String

On Error GoTo ErrHandle
'-------------------------
    If Me.Count = 0 Then Exit Function
    nDim = UBound(dict.Keys)
    
    For I = 0 To nDim
        sRes = sRes & GetKey(I) & SEQV & GetValue(I) & DLM
    Next I
'--------------------------------------------------------------
    If sRes <> "" Then sRes = Left(sRes, Len(sRes) - Len(DLM))
'--------------------------------------------------------------
ExitHere:
    GetKVString = sRes '!!!!!!!!!!!!!
    Exit Function
'---------------
ErrHandle:
   ErrPrint "GetKVString", Err.Number, Err.Description
   Err.Clear
End Function

'========================================================================================================================
' Load From 2D Array
'========================================================================================================================
Public Function LoadFromArray(Arr As Variant) As Boolean

Dim I As Long, ItemVar As Variant
Dim KeyVal As String

On Error GoTo ErrHandle
'---------------------------------------------
If IsArray(Arr) = False Then Err.Raise 1000, , "Unqualified Array"
If NumberOfArrayDimensions(Arr:=Arr) <> 2 Then Err.Raise 1001, , "Wrong Array Dimension"
    
For I = LBound(Arr, 1) To UBound(Arr, 1)
    dict.Add key:=Arr(I, LBound(Arr, 2) + 1), Item:=Arr(I, LBound(Arr, 2))
Next I
'-----------------------------
ExitHere:
   LoadFromArray = True  '!!!!!!!!!
   Exit Function
ErrHandle:
   ErrPrint "LoadFromArray", Err.Number, Err.Description
   Err.Clear
End Function

'========================================================================================================================
' Convert to Collection
'========================================================================================================================
Public Function ToCollection() As Collection
Dim PS As Collection, I As Integer, p As cKeyValue
Dim sItem As Variant, sKey As Variant

On Error GoTo ErrHandle
'---------------------------------
   If Me.Count = 0 Then Exit Function
   Set PS = New Collection

   For Each sKey In dict.Keys
        Set p = New cKeyValue
        p.key = CStr(sKey)
        If IsObject(dict(sKey)) Then
           Set p.value = dict(sKey)
        Else
           p.value = dict(sKey)
        End If
        
        PS.Add p, CStr(sKey)
        
   Next sKey
'----------------------------------
ExitHere:
   Set ToCollection = PS
   Exit Function
'---------------
ErrHandle:
   ErrPrint "ToCollection", Err.Number, Err.Description
   Err.Clear
End Function
'-----------------------------------------------------------------------------------------------------------------------
' Sorting Dictionary
'-----------------------------------------------------------------------------------------------------------------------
Private Sub SortDictionary(intSort)
Dim strDict() As String
Dim ObjKey
Dim strKey As String, strItem As String
Dim x, y, z As Integer
        
Const dictKey = 1
Const dictItem = 2
'------------------------------------------
  z = dict.Count
  If z > 1 Then
    ReDim strDict(z, 2)
    x = 0
    ' ----------------------
    For Each ObjKey In dict
        strDict(x, dictKey) = CStr(ObjKey)
        strDict(x, dictItem) = CStr(dict(ObjKey))
        x = x + 1
    Next

    ' ----------------------
    For x = 0 To (z - 2)
      For y = x To (z - 1)
        If StrComp(strDict(x, intSort), strDict(y, intSort), vbTextCompare) > 0 Then
            strKey = strDict(x, dictKey)
            strItem = strDict(x, dictItem)
            strDict(x, dictKey) = strDict(y, dictKey)
            strDict(x, dictItem) = strDict(y, dictItem)
            strDict(y, dictKey) = strKey
            strDict(y, dictItem) = strItem
        End If
      Next
    Next
    ' --------------------------------------------------------------------
    dict.RemoveAll
    For x = 0 To (z - 1)
      dict.Add strDict(x, dictKey), strDict(x, dictItem)
    Next
    ' --------------------------------------------------------------------
  End If
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
                                                                                                     Optional sModName As String = "cDictionary") As String
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
