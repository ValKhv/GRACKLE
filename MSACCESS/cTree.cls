VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*********************************************************************************************************************
'**********************************************************************************************************************************************
'**********************************************************************************************************************************************
' DESCRIPTION: ����� ������������ ����� ���������� ������
' VERSION    : v 0.2 29.05.2011
' COPYRIGHT  :
'***********************************************************************************************************************************************
'***********************************************************************************************************************************************
Option Compare Database
Option Explicit

Private Type Attribut                     ' ������� (��������)
   NodeID As Long                         ' ����� ����, � �������� ������������� �������
   ANAME As String                        ' ������������ ��������
   aValue As String                       ' �������� ��������
   ATYPE As Integer                       ' ��� ��������: 0  - ������, 1 - Long, 2 - Double, 3 - Date, 4 - Currency (41 - USD, 42 - EUR, 43 - CAD, 44 - RUB)
   ATAG As String                         ' ��� (����������)
End Type

Private Type Node
    Index As Long                         ' ����� � �������
    NodeID As Long                        ' ��������� ����, ��������� � ���������� ���� � �������
                                          ' � ����� ������, ������������ ��� ����������
    ParentID As Long                      ' ������ �� ��������
    sKey As String                        ' ����
    sValue As String                      ' ��������
    sTitle As String                      ' ���������
    sTag As String                        ' ��� - ��� ���������� ��������
    ATTRS() As Attribut                   ' ������ ��������� (�������)
    CHILDS() As Long                      ' ������ ��� ���������������� �����
End Type



Private UARRAY() As Node                  ' ������ - ������
Private UARRAYDim As Long                 ' ����������� ������� - ������

'============================================================================================================================================
' �������� ��� ������ - �������� ����������� ������ (���-�� ���������)
'============================================================================================================================================
Public Property Get Count()
    Count = UARRAYDim
End Property
'============================================================================================================================================
' �������� ID (�������������� ID)
'============================================================================================================================================
Public Property Get id(Index As Long) As Long
   id = UARRAY(Index).NodeID
End Property
Public Property Let id(Index As Long, iValue As Long)
  UARRAY(Index).NodeID = iValue
End Property
'============================================================================================================================================
' �������� ���� (Name)
'============================================================================================================================================
Public Property Get key(Index As Long) As String
   If Index < 0 Then Exit Property
   key = UARRAY(Index).sKey
End Property
Public Property Let key(Index As Long, sValue As String)
  If Index < 0 Then Exit Property
  UARRAY(Index).sKey = sValue
End Property
'============================================================================================================================================
' �������� ���
'============================================================================================================================================
Public Property Get TAG(Index As Long) As String
   TAG = UARRAY(Index).sTag
End Property
Public Property Let TAG(Index As Long, sValue As String)
  UARRAY(Index).sTag = sValue
End Property
'============================================================================================================================================
' �������� ��������
'============================================================================================================================================
Public Property Get value(Index As Long) As String
   value = UARRAY(Index).sValue
End Property
Public Property Let value(Index As Long, sValue As String)
  UARRAY(Index).sValue = sValue
End Property
'============================================================================================================================================
' �������� ������������
'============================================================================================================================================
Public Property Get Title(Index As Long) As String
   Title = UARRAY(Index).sTitle
End Property
Public Property Let Title(Index As Long, sValue As String)
  UARRAY(Index).sTitle = sValue
End Property
'============================================================================================================================================
' �������� ��������
'============================================================================================================================================
Public Property Get Parent(Index As Long) As Long
   Parent = UARRAY(Index).ParentID
End Property
Public Property Let Parent(Index As Long, iValue As Long)
   UARRAY(Index).ParentID = iValue
End Property
'============================================================================================================================================
' ����� �������� �������� ����
'============================================================================================================================================
Public Function GetProperty(NodeID As Long, PropertyName As String) As String
Dim iProp As Long
    iProp = GetPropertyID(NodeID, PropertyName)
    If iProp < 0 Then Exit Function
    '------------------------------------------------
    GetProperty = UARRAY(NodeID).ATTRS(iProp).aValue '!!!!!!!!!!!!!!!!!!
End Function
'============================================================================================================================================
' ����� �������� ������� (��������) � ���� ��� �������� ���. � ������ ������ ���������� ����� �������� - ��� ������
'============================================================================================================================================
Public Function SetProperty(NodeID As Long, PropertyName As String, PropertyValue As String, Optional PropertyType As Integer = -1) As Long
Dim iRes As Long                         ' ������������ ���������
Dim nDim As Long                         ' ����������� ������� �������

On Error GoTo ErrHandle
'-------------------------------------------------------------------------------
   If NodeID < 0 Or NodeID > UARRAYDim Then        ' ��������� ������� ����
       iRes = -1: GoTo ExitHere
   End If
'-------------------------------------------------------------------------------
   iRes = GetPropertyID(NodeID, PropertyName)
   If iRes = -1 Then        ' ������ �������� ���, ������� ��� �������
      nDim = GetArrDim(NodeID)
      If nDim < 0 Then      ' ������� ���, ������� ��
          ReDim UARRAY(NodeID).ATTRS(0)
          nDim = 0
      Else                  ' �������� ����, ��������� ������
          nDim = nDim + 1
          ReDim Preserve UARRAY(NodeID).ATTRS(nDim)
      End If
      ' ������� ��������
      iRes = nDim
      UARRAY(NodeID).ATTRS(iRes).ANAME = PropertyName
      UARRAY(NodeID).ATTRS(iRes).aValue = PropertyValue
      If PropertyType > -1 Then UARRAY(NodeID).ATTRS(iRes).ATYPE = PropertyType
   Else                     ' �������� �������, ������ ����������� ��� ��������
      UARRAY(NodeID).ATTRS(iRes).aValue = PropertyValue
      If PropertyType > -1 Then UARRAY(NodeID).ATTRS(iRes).ATYPE = PropertyType
   End If
'-----------------------------------
ExitHere:
     SetProperty = iRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
     Exit Function
'----------------------
ErrHandle:
     Debug.Print String(40, "#") & vbCrLf & "ERR#" & Err.Number & vbCrLf & Err.Description & vbCrLf & String(40, "#")
     Err.Clear
     Resume ExitHere
End Function
'--------------------------------------------------------------------------------------------------------------------------------------------
' ������� ���������� ��� ��������� ���� ID �������� �� ��� �����. ���� ������ �������� �� ����������, ������������ -1
'--------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetPropertyID(NodeID As Long, PropertyName As String) As Long
Dim PropID As Long                           ' ������������ ���������
Dim nDim As Long, I As Long                  ' ����������� ������� ������� � ����������� ����������

On Error GoTo ErrHandle
PropID = -1: If PropertyName = "" Then GoTo ExitHere
If NodeID > UARRAYDim Then GoTo ExitHere
'---------------------
nDim = GetArrDim(NodeID)
If nDim = -1 Then GoTo ExitHere
    For I = 0 To nDim
        If UCase(UARRAY(NodeID).ATTRS(I).ANAME) = UCase(PropertyName) Then
            PropID = I
            Exit For
        End If
    Next I
'---------------------
ExitHere:
    GetPropertyID = PropID '!!!!!!!!!!!!!!!!!!
    Exit Function
'-------------
ErrHandle:
    Err.Clear
    Resume ExitHere
End Function
'============================================================================================================================================
' ����� ��������. ���������� ID (����� �������� �������) - ������� ��� ��������
'============================================================================================================================================
Public Function Add(Optional sKey As String = "", Optional sValue As String = "", Optional sTitle As String = "", _
                 Optional iParent As Long = -1, Optional NodeID As Long = -1, Optional sTag As String = "") As Long
Dim nChilds As Long
On Error GoTo ErrHandle
'------------------------------------------------------------------------------------------------
    UARRAYDim = UARRAYDim + 1: ReDim Preserve UARRAY(UARRAYDim)
    UARRAY(UARRAYDim).Index = UARRAYDim
    If sKey <> "" Then UARRAY(UARRAYDim).sKey = sKey
    If sValue <> "" Then UARRAY(UARRAYDim).sValue = sValue
    If sTitle <> "" Then UARRAY(UARRAYDim).sTitle = sTitle
    If NodeID <> -1 Then UARRAY(UARRAYDim).NodeID = NodeID
    If sTag <> "" Then UARRAY(UARRAYDim).sTag = sTag
    '------------------------------------------------------------------------------
    UARRAY(UARRAYDim).ParentID = iParent                   ' ������������ ��������, �������� ������ � ����� ������,
    If iParent > -1 Then                                   ' ������� �� ������������ ID ��� ������������ ����������
          If iParent < UARRAYDim Then                      ' ������ �� ������������� ��������
                 nChilds = GetChildDim(iParent)
                 If nChilds > -1 Then                      ' ������ �������� ����� ��� �����
                    nChilds = nChilds + 1: ReDim Preserve UARRAY(iParent).CHILDS(nChilds)
                    UARRAY(iParent).CHILDS(nChilds) = UARRAYDim
                 Else                                      ' ������������� ������
                    ReDim UARRAY(iParent).CHILDS(0)
                    UARRAY(iParent).CHILDS(0) = UARRAYDim
                 End If
          End If
    End If
'------------------------------------------------------------------------------
ExitHere:
    Add = UARRAYDim  '!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'------------------------------------
ErrHandle:
    Debug.Print String(40, "#") & vbCrLf & "ERR#" & Err.Number & vbCrLf & Err.Description & vbCrLf & String(40, "#")
    Err.Clear
End Function
'============================================================================================================================================
' ����� ���������� � ������ ������ ������� ������
' �����: index - ������������ ���� � �������� ����������� ���������, SubTree - ����������� ������
'============================================================================================================================================
Public Function InsertSubTree(ParentIndex As Long, SubTree As cTree) As Long
Dim I As Long, nDD As Long                                      ' ����������� ��������� � ���������� ����������� ����������
Dim wNode As Node, wParent As Long                              ' ������� ���� � ������� ������ �� ������� ����
'-------------------------------------------------------------------------------------------
' ��������� �������, ������� ������ ��������� � ������� �������
   wNode = StringToNode(SubTree.GetNodeAsString(0)): wNode.ParentID = ParentIndex
   wParent = AddNODE(wNode)                     ' ����� ������������ ������ (������ 0)
   nDD = SubTree.Count                          ' ���������� ��������� �� ����������� �����
'-------------------------------------------------------------------------------------------
'  ��������� ��������� ����, ���������� �� ������������ ���� - wParent
   If nDD > 0 Then
        For I = 1 To nDD
           wNode = StringToNode(SubTree.GetNodeAsString(I))   ' �������� ��������� ����
           wNode.ParentID = wParent + wNode.ParentID          ' ������������� ��������
           AddNODE wNode                                      ' ��������� � �������� ������
        Next I
   End If
'--------------------------------------------------------
   InsertSubTree = nDD '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'============================================================================================================================================
' ����� ��� ��������� ���� �������� ������ ��� �������� ����� (� ���� �������)
' ���� ������ bForce, �� ������ ���������������, ����� - ������� �� ��� ������������
'============================================================================================================================================
Public Function GetChilds(Index As Long, Optional bForce As Boolean = False) As Long()
Dim iRes() As Long, nDim As Long                          ' ������������ ���������
Dim I As Long                                             ' ������������� ������

On Error GoTo ErrHandle
'----------------------------------------------------------------------------------
nDim = GetChildDim(Index)                                 ' �������� ����������� �������� ������� �������� �����
If bForce Or nDim = -1 Then                               ' �������� ���� ����������
nDim = -1: ReDim iRes(0): iRes(0) = 1
            For I = 0 To UARRAYDim
                If UARRAY(I).ParentID = Index Then
                    nDim = nDim + 1: ReDim Preserve iRes(nDim)
                    iRes(nDim) = I
                End If
            Next I
            If (Index > -1) And (nDim > -1) Then UARRAY(Index).CHILDS = iRes
Else                                                    ' ����� ��� ��������� ������
                     iRes = UARRAY(Index).CHILDS
End If
'----------------------------------------------------
ExitHere:
                 GetChilds = iRes '!!!!!!!!!!!!!!!!!
                 Exit Function
'-----------------------------------------
ErrHandle:
    Debug.Print String(40, "#") & vbCrLf & "ERR#" & Err.Number & vbCrLf & Err.Description & vbCrLf & String(40, "#")
    Err.Clear
End Function
'============================================================================================================================================
' ����� ��������� ��� ��������� ���� �������� ������ ��� �������� �����, ���� ������� �������� �������� �����
'============================================================================================================================================
Public Function GetChildsWithKey(Index As Long, sKey As String) As Long()
Dim iRes() As Long, I As Long, nDim As Long
Dim iChols() As Long, mDim As Long

On Error GoTo ErrHandle
'----------------------------------------------------
nDim = -1: ReDim iRes(0): iRes(0) = -1
        iChols = GetChilds(Index): mDim = UBound(iChols)      ' �������� ������ �������� �����
        If iChols(0) = -1 Then                                ' �������� ����� ���
           iRes(0) = -1
           GoTo ExitHere
        End If
        For I = 0 To mDim
            If InStr(1, UARRAY(iChols(I)).sKey, sKey) > 0 Then
               nDim = nDim + 1
               ReDim Preserve iRes(nDim): iRes(nDim) = iChols(I)
            End If
        Next I
'----------------------------------------------------
ExitHere:
                 GetChildsWithKey = iRes '!!!!!!!!!!!!!!!!!
                 Exit Function
'-----------------------------------------
ErrHandle:
    Debug.Print String(40, "#") & vbCrLf & "ERR#" & Err.Number & vbCrLf & Err.Description & vbCrLf & String(40, "#")
    Err.Clear
End Function
'============================================================================================================================================
' ������� ����������� ������ � ������
'============================================================================================================================================
Public Function TreeToString(Optional sNodeDelim As String = "", Optional sTreeDelim As String = "") As String
Dim I As Long
Dim sRes As String, sNode As String

    For I = 0 To UARRAYDim
        sNode = NodeToString(UARRAY(I), sNodeDelim)
        sRes = IIf(sRes = "", sNode, sRes & sTreeDelim & sNode)
    Next I
'--------------------------------------------------------------------------------------------------------------------------------------------
        TreeToString = sRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'============================================================================================================================================
' ������� ��������������� ������ �� ������
'============================================================================================================================================
Public Function StringToTree(StrTree As String, Optional sNodeDelim As String = "", Optional sTreeDelim As String = "") As Long
Dim Nodes() As String, nNODES As Long                                   ' ������ �����
Dim I As Long, wNode As Node                                            ' ����������� ���������� � ������� ����
Dim TreeCount As Long
Call Initialize                                                                   ' ����������������
Nodes = Split(StrTree, sTreeDelim): nNODES = UBound(Nodes)                        ' ������� ������ �����
'-----------------------------------------------------------------------------------------------------------------
For I = 0 To nNODES              ' ���� �� ������� ����, �������� ��������
   If Nodes(I) <> "" Then
        wNode = StringToNode(Nodes(I), sNodeDelim)
        TreeCount = AddNODE(wNode)
   End If
Next I
'--------------------------------------------------------
                StringToTree = Me.Count '!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'============================================================================================================================================
'  ����� �������� ���� � ���� ��������������� ������
'============================================================================================================================================
Public Function GetNodeAsString(Index As Long, Optional sNodeDelim As String = "") As String
Dim sRes As String
    If Index <= UBound(UARRAY) Then
        sRes = NodeToString(UARRAY(Index), sNodeDelim)
    Else
        sRes = ""
    End If
'------------------------------------------------
    GetNodeAsString = sRes '!!!!!!!!!!!!!!!!!!!!!!!!
End Function
'============================================================================================================================================
' ����� ��������� ������ �� ������� (������ ������������)
' PARAMS:   TableName - ������������ �������, �� ������� ��������� ������
'           ID_FIELDNAME - ������������ ����, ������� ������ �������������  ���� (�� ��������� ID)
'           RecordID     - �������� ID, ��� �������� ���������� ����� ������. ���� "-1" - ����� �������. ������������ �������������
'           KeyFieldName - ������������ ����, ������ ������ ���� (���� "" - �� ���������)
'           ValueFieldName - �������� ����
'           TitleFieldName - ������������ ��������� ����
'           ParentFieldName - ������������ ����, ������� ������ ������ �� ������������ ����
'           ParentFieldValue - ������������ ������������� ����, �� �������� ���������� ����������
'============================================================================================================================================
Public Function LoadFromTable(TableName As String, Optional ID_FLDNAME As String = "ID", Optional RecordID As Variant = -1, _
                                       Optional Key_FLDNAME As String = "", Optional Title_FLDNAME As String = "", Optional Value_FLDNAME As String = "", _
                                       Optional Parent_FLDNAME As String = "", Optional RANG_FLDNAME As String = "", _
                                       Optional sCriteria As String = "") As Boolean
Dim bRes As Boolean                                                                         ' ������������ ���������
Dim sTableName As String, sID_FLDNAME As String, sKey_FLDNAME As String                     ' ��������������� ������������ �����
Dim sTitle_FLDNAME As String, sValue_FLDNAME As String, sParent_FLDNAME As String           ' ��������������� ������������ �����
On Error GoTo LoadFromTable_ERR

'-----------------------------------------------------
If RecordID = -1 Then Exit Function
    sTableName = SHT(TableName): sID_FLDNAME = SHT(ID_FLDNAME): sKey_FLDNAME = SHT(Key_FLDNAME)
    sTitle_FLDNAME = SHT(Title_FLDNAME): sValue_FLDNAME = SHT(Value_FLDNAME): sParent_FLDNAME = SHT(Parent_FLDNAME)
    Erase UARRAY:     ReDim UARRAY(0): UARRAYDim = 0  ' ��������� ��������� ������ ������
'-----------------------------------------------------
' ��������� �������� ����-��������, �� �������� ������ ��������� ������
    UARRAY(0) = GetNodeFromTable(sTableName, sID_FLDNAME, RecordID, sKey_FLDNAME, sTitle_FLDNAME, sValue_FLDNAME, sParent_FLDNAME)
' �������� �������� ������� ����, ���� ��� ����
    bRes = GetChildsNodesFromTable(0, sParent_FLDNAME, RecordID, sTableName, sID_FLDNAME, sKey_FLDNAME, _
                                   sTitle_FLDNAME, sValue_FLDNAME, RANG_FLDNAME, sCriteria)
'------------------------------------------------------------------------------------
LoadFromTable_EXIT:
    LoadFromTable = bRes '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'--------------------------------
LoadFromTable_ERR:
    MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "LoadFromTable ERROR"
    Resume LoadFromTable_EXIT
End Function

'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()
    Call Initialize
End Sub
Private Sub Class_Terminate()
    Erase UARRAY
End Sub
'----------------------------------------------------------------------------------------------------------------------------------------------
' ������� �������������� �����
'----------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Initialize()
    UARRAYDim = -1
    ReDim UARRAY(0)
End Sub
'----------------------------------------------------------------------------------------------------------------------------------------------
' ������� �� ����� ��������� ��������� ���� NODE, ��������� ID
'----------------------------------------------------------------------------------------------------------------------------------------------
Private Function AddNODE(MyNode As Node) As Long
 AddNODE = Me.Add(MyNode.sKey, MyNode.sValue, MyNode.sTitle, MyNode.ParentID, MyNode.NodeID, MyNode.sTag)
End Function
'----------------------------------------------------------------------------------------------------------------------------------------------
' ������� �� ��������� ParentID �������� �� �������� ������� �������� ���� Node, ������� ������ �� �������� ���� � ��������� �� � ������
'----------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetChildsNodesFromTable(ByVal ArrayID As Long, ByVal Parent_FLDNAME As String, ByVal ParentID As Long, _
                                         ByVal TableName As String, ByVal ID_FLDNAME As String, _
                                         Optional Key_FLDNAME As String = "", Optional ByVal Title_FLDNAME As String = "", _
                                         Optional ByVal Value_FLDNAME As String = "", Optional ByVal RANG_FLDNAME As String = "", _
                                         Optional ByVal sCriteria As String = "") As Boolean
Dim CurrentID As Long, CurArrayID As Long          ' ������� ID � �������, ������� ID � �������
Dim cNode As Node                                  ' ������� ����
Dim ParentIDD As Long                              ' ������������ ����, ��� �������� ���� ��������
Dim sSORT As String                                ' ���� ��� ���������� ���������
Dim sWhere As String                               ' �������������� �������
Dim sSQL As String, rst As DAO.Recordset           ' ��������� ������� � ����� ������
Dim bRes As Boolean
    On Error GoTo GetChildsNodesFromTable_ERR

bRes = False
ParentIDD = ParentID         ' ����� ����, ��� ������� ������� �������
'----------------------------------------------------------------------------
' �������� �������� ������� ����, ���� ��� ����
    sWhere = " WHERE (((" & TableName & "." & Parent_FLDNAME & ")=" & ParentIDD & ")"
    sWhere = IIf(sCriteria = "", sWhere & ")", sWhere & " AND (" & sCriteria & "))")
    sSORT = IIf(RANG_FLDNAME = "", "", " ORDER BY " & TableName & "." & RANG_FLDNAME)
    sSQL = "SELECT " & TableName & "." & ID_FLDNAME & " FROM " & TableName & _
           sWhere & sSORT & ";"
   '--------------------------------------------------------------------------------------
                            Set rst = CurrentDb.OpenRecordset(sSQL, dbOpenDynaset, dbSeeChanges)
'*****************
If Not rst.EOF Then
                            rst.MoveLast: rst.MoveFirst
   '--------------------------------------------------------------------------------------
   With rst
      Do While Not .EOF
        CurrentID = rst(ID_FLDNAME)
            cNode = GetNodeFromTable(TableName, ID_FLDNAME, CurrentID, Key_FLDNAME, Title_FLDNAME, Value_FLDNAME, Parent_FLDNAME)
            cNode.ParentID = ArrayID                    ' ������ ������ �� ������������ ���� � ���������� �������� �� �������� � �������
            CurArrayID = AddNODE(cNode)                 ' ��������� ���� � ������������ �������
            '----------------------------------------------------------------------------
            ' ���������� �������� ���� ��� ���������� �������� ����� �� ��������� � ��������
            
            bRes = GetChildsNodesFromTable(CurArrayID, Parent_FLDNAME, CurrentID, TableName, ID_FLDNAME, Key_FLDNAME, _
                                           Title_FLDNAME, Value_FLDNAME, RANG_FLDNAME, sCriteria)
            '----------------------------------------------------------------------------
        .MoveNext
      Loop
   End With
End If
'*****************
   bRes = True
'-----------------------------------------------------------------------------------------
GetChildsNodesFromTable_EXIT:
    GetChildsNodesFromTable = bRes '!!!!!!!!!!!!!!!!!!!!!!!!!
    Set rst = Nothing
    Exit Function
'-------------------------------------------
GetChildsNodesFromTable_ERR:
    MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "GetChildsNodesFromTable ERROR"
    Err.Clear
    Resume GetChildsNodesFromTable_EXIT
End Function
'----------------------------------------------------------------------------------------------------------------------------------------------
' ������� �� ��������� ID �������� �� �������� ������� �������� � ��������� ��� ��������� ���� Node
'----------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetNodeFromTable(ByVal TableName As String, ByVal ID_FIELDNAME As String, ByVal RecordID As Variant, _
                                  Optional Key_FLDNAME As String = "", Optional Title_FLDNAME As String = "", Optional Value_FLDNAME As String = "", _
                                  Optional Parent_FLDNAME As String = "") As Node
Dim cNode As Node                                                 ' ������� ����
Dim STABLE As String                                              ' ������� ������, � ������� ����������� ������
Dim RS As DAO.Recordset                                           ' ����� ������
Dim sSQL As String                                                ' ������ � �������

On Error GoTo GetNodeFromTable_ERR
'-----------------------------------------------------
'  ��������� �����
   sSQL = " FROM " & TableName & " WHERE (((" & TableName & "." & ID_FIELDNAME & ")=" & RecordID & "));"
   If Value_FLDNAME <> "" Then sSQL = "," & TableName & "." & Value_FLDNAME & sSQL
   If Key_FLDNAME <> "" Then sSQL = "," & TableName & "." & Key_FLDNAME & sSQL
   If Title_FLDNAME <> "" Then sSQL = "," & TableName & "." & Title_FLDNAME & sSQL
   If Parent_FLDNAME <> "" Then sSQL = "," & TableName & "." & Parent_FLDNAME & sSQL
   sSQL = "SELECT " & TableName & "." & ID_FIELDNAME & sSQL
   '--------------------------------------------------------------------------------------
                            Set RS = CurrentDb.OpenRecordset(sSQL, dbOpenDynaset, dbSeeChanges): RS.MoveLast: RS.MoveFirst
   '--------------------------------------------------------------------------------------
    If Not RS.EOF Then        ' ��������� ������ ������ ������
        cNode.NodeID = RecordID
        If Key_FLDNAME <> "" Then cNode.sKey = Nz(RS(Key_FLDNAME), "")
        If Value_FLDNAME <> "" Then cNode.sValue = Nz(RS(Value_FLDNAME), "")
        If Title_FLDNAME <> "" Then cNode.sTitle = Nz(RS(Title_FLDNAME), "")
        If Parent_FLDNAME <> "" Then cNode.ParentID = Nz(RS(Parent_FLDNAME), -1)
    End If
'------------------------------------------------------------------------------------
GetNodeFromTable_EXIT:
    Set RS = Nothing
    GetNodeFromTable = cNode '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Exit Function
'--------------------------------
GetNodeFromTable_ERR:
    MsgBox "ERR#" & Err.Number & vbCrLf & Err.Description, vbCritical, "GetNodeFromTable ERROR"
    Resume GetNodeFromTable_EXIT


End Function
'------------------------------------------------------------------------------------------------------------------------------------------------
' ������� ��������� ������ � ���������� ������
'------------------------------------------------------------------------------------------------------------------------------------------------
Private Function SHT(sI As String) As String
  SHT = IIf(Left(sI, 1) = "[", sI, "[" & sI & "]") '!!!!!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------
' ������� ����������� ���� � ������
'------------------------------------------------------------------------------------------------------------------------------------------------
Private Function NodeToString(nNode As Node, Optional sNodeDelim As String = "") As String
    NodeToString = nNode.Index & sNodeDelim & nNode.NodeID & sNodeDelim & nNode.sKey & sNodeDelim & _
                     nNode.sValue & sNodeDelim & nNode.sTitle & sNodeDelim & nNode.ParentID & sNodeDelim & nNode.sTag    '!!!!!!!!!!!!!!!!!!
'-------------------------------------------------------------------------------------------------------------
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------
' ������� ����������� ������ � ����
'------------------------------------------------------------------------------------------------------------------------------------------------
Private Function StringToNode(sNode As String, Optional sNodeDelim As String = "") As Node
Dim nyNode As Node, sKV() As String                            ' ���� � ������ ��������
If sNode <> "" Then
    sKV = Split(sNode, sNodeDelim)
    nyNode.Index = sKV(0): nyNode.NodeID = sKV(1)
    nyNode.sKey = sKV(2):    nyNode.sValue = sKV(3)
    nyNode.sTitle = sKV(4):    nyNode.ParentID = sKV(5): nyNode.sTag = sKV(6)
End If
'---------------------------------------------------------------------
   StringToNode = nyNode '!!!!!!!!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------
' ������� �������� ����������� ������� �������, � ������ ���� �������� ������
'------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetArrDim(NodeID As Long) As Long
Dim iRes As Long

On Error GoTo ErrHandle
        iRes = UBound(UARRAY(NodeID).ATTRS)
'---------------------------------------------------------
ExitHere:
        GetArrDim = iRes '!!!!!!!!!!!!!!!!
        Exit Function
'--------------
ErrHandle:
        Err.Clear
        iRes = -1
        Resume ExitHere
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------
' ������� �������� ����������� ������� �������� �����, � ������ ���� �������� ������
'------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetChildDim(NodeID As Long) As Long
Dim iRes As Long

On Error GoTo ErrHandle
        iRes = UBound(UARRAY(NodeID).CHILDS)
'---------------------------------------------------------
ExitHere:
        GetChildDim = iRes '!!!!!!!!!!!!!!!!
        Exit Function
'--------------
ErrHandle:
        Err.Clear
        iRes = -1
        Resume ExitHere
End Function

