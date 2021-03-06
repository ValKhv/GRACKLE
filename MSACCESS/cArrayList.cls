VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
' DESCRIPTION: ArrayList (i.e. Vector)
' VERSION    : v 0.1 30.01.2017
' COPYRIGHT  : V.Khvatov DigitalXpert
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
Option Compare Database
Option Explicit


Private m_List As Object
'######################################################################################################################################################

'======================================================================================================================================================
' Add new Element
'=====================================================================================================================================================
Public Sub Add(vItem As Variant)
     m_List.Add vItem
End Sub
'======================================================================================================================================================
' Add new Element
'=====================================================================================================================================================
Public Function Count() As Long
    Count = m_List.Count
End Function
'======================================================================================================================================================
' Exttract Item
'=====================================================================================================================================================
Public Property Get Item(cnt As Integer) As Variant
  Item = m_List.Item(cnt)
End Property
'======================================================================================================================================================
' Directed Add Item
'=====================================================================================================================================================
Public Property Let Item(cnt As Integer, value As Variant)
Attribute Item.VB_UserMemId = 0
    m_List.Item(cnt) = value
End Property


'======================================================================================================================================================
' Get/set Capacity
'=====================================================================================================================================================
Public Property Get Capacity() As Long
    Capacity = m_List.Capacity
End Property
Public Property Let Capacity(ByVal lNewValue As Long)
    m_List.Capacity = lNewValue
End Property

'======================================================================================================================================================
' Clear all elements
'=====================================================================================================================================================
Public Sub Clear()
    m_List.Clear
End Sub

'======================================================================================================================================================
' Convert To String
'=====================================================================================================================================================
Public Function ToString(Optional DLM As String = vbCrLf) As String
   If Me.Count = 0 Then Exit Function
   ToString = Join(m_List.ToArray(), DLM)
End Function
'======================================================================================================================================================
'Determines whether an element is in the ArrayList.
'======================================================================================================================================================
Public Function Contains(ObjectA) As Boolean
    Contains = m_List.Contains(ObjectA)
End Function
'======================================================================================================================================================
' Clone elements
'=====================================================================================================================================================
Public Function Clone() As Object
    Set Clone = m_List.Clone
End Function

'======================================================================================================================================================
' ToArray
'=====================================================================================================================================================
Public Function ToArray() As Variant
    ToArray = m_List.ToArray
End Function
'=====================================================================================================================================================
'Sets the capacity to the actual number of elements in the ArrayList.
'=====================================================================================================================================================
Public Sub TrimToSize()
    m_List.TrimToSize
End Sub
'=====================================================================================================================================================
'Inserts an element into the ArrayList at the specified index.
'=====================================================================================================================================================
Public Sub Insert(Start As Long, value As Variant)
    m_List.Insert Start, value
End Sub

'=====================================================================================================================================================
'Search Unique Element
'=====================================================================================================================================================
Public Function IndexOf(value As Variant, Optional Start As Long = -1, Optional CountOf As Long = -1) As Long
    If Start = -1 Then
        IndexOf = m_List.IndexOf_3(value)
    Else
        IndexOf = m_List.IndexOf_2(value, Start, CountOf)
    End If
End Function

'=====================================================================================================================================================
'Sorts the elements in the entire ArrayList
'=====================================================================================================================================================
Public Sub Sort()
    m_List.Sort
End Sub
'=====================================================================================================================================================
'Removes the first occurrence of a specific object from the ArrayList.
'=====================================================================================================================================================
Public Sub Remove(ObjectA As Object)
    m_List.Remove
End Sub
'=====================================================================================================================================================
'Removes the element at the specified index of the ArrayList.
'=====================================================================================================================================================
Public Sub RemoveAt(Start As Long)
    m_List.RemoveAt Start
End Sub
'=====================================================================================================================================================
'Returns an ArrayList which represents a subset of the elements in the source ArrayList.
'=====================================================================================================================================================
Public Function GetRange(Start As Long, CountOf As Long) As Object
    Set GetRange = m_List.GetRange(Start, CountOf)
End Function

'=====================================================================================================================================================
'Reverse(): Reverses the order of the elements in the entire ArrayList.
'=====================================================================================================================================================
Public Sub Reverse(Optional Start As Long = -1, Optional CountOf As Long = -1)
    If Start = -1 Then
        m_List.Reverse
    Else
        m_List.Reverse_2 Start, CountOf
    End If
End Sub

'=====================================================================================================================================================
'Determines whether the specified object is equal to the current object.(Inherited from Object.)
'=====================================================================================================================================================
Public Function Equals(Control As Object) As Boolean
    Equals = m_List.Equals(Control)
End Function

'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
    Set m_List = CreateObject("System.Collections.ArrayList") 'Create the ArrayList
End Sub
Private Sub Class_Terminate()
    Set m_List = Nothing
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
                                                                                                  Optional sModName As String = "cArrayList") As String
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
