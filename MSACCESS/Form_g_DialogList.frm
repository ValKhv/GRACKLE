VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_g_DialogList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
' LIST DIALOG (MsgBox with options)
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
Option Compare Database
Option Explicit

    Private Const DIALOGLIST As String = "DIALOGLIST"
'**************************

Private Sub btOK_Click()
Dim sRes As String
    
    On Error Resume Next
  '-------------------------------
  sRes = Nz(LISTLIST.value, "")
  If sRes <> "" Then
      TempVars(DIALOGLIST).value = sRes
  End If
  '-------------------------------
  DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_Load()
Dim sRef As String, Arr() As String
Dim I As Integer, sList As String, nDim As Integer
    
Const DLM As String = ";"

    On Error Resume Next
'--------------------------------

    sRef = Nz(Me.OpenArgs, "")
    If sRef <> "" Then
        Arr = Split(sRef, DLM): nDim = UBound(Arr)
        Me.Caption = Arr(0)
        For I = 1 To nDim
           sList = sList & DLM & Arr(I)
        Next I
        If sList <> "" Then sList = Right(sList, Len(sList) - Len(DLM))
        If sList <> "" Then LISTLIST.RowSource = sList
    Else
        Me.Caption = "DEMO of Dialog List"
        LISTLIST.RowSource = "Demo Option 1;Demo Option 2;Demo Option 3"
    End If


End Sub



