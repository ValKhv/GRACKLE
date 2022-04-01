VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_g_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Const DIALOGOPT As String = "DIALOGOPT"
'**************************************
Private Sub btOK_Click()
Dim iRes As Long, sRes As String
  iRes = THEOPTS.value
  
  Select Case iRes
  Case 1:   sRes = Me.lblOpt1.Caption
  Case 2:   sRes = Me.lblOpt2.Caption
  Case 3:   sRes = Me.lblOpt3.Caption
  End Select
  
  TempVars(DIALOGOPT).value = sRes
  
  DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_Load()
    '----------------------------------
        RestoreWindowState Me
    '----------------------------------
  Call ProcessArgs(Nz(Me.OpenArgs, ""))
End Sub


Private Sub ProcessArgs(sARG As String, Optional DLM As String = ";")
Dim VARG() As String, nDim As Integer
  
  On Error GoTo ErrHandle
'---------------------
If sARG = "" Then Exit Sub

    VARG = Split(sARG, DLM)
 
    Me.Caption = VARG(0)
    Me.lblOptGroup.Caption = VARG(1)
 
    Me.lblOpt1.Caption = VARG(2)
    Me.lblOpt2.Caption = VARG(3)
 
 If VARG(4) = "" Then
     Me.lblOpt3.Visible = False
     Me.Opt3.Visible = False
 Else
     Me.lblOpt3.Visible = True
     Me.Opt3.Visible = True
     Me.lblOpt3.Caption = VARG(4)
 End If
'---------------------
ExitHere:
    Exit Sub
'-----------
ErrHandle:
    ErrPrint2 "ProcessArgs", Err.Number, Err.Description, Me.Name
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '----------------------------------
        WriteWindowState Me
    '----------------------------------
End Sub
