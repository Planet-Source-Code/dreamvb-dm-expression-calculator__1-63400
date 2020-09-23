VERSION 5.00
Begin VB.Form frmVariables 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Variables"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox LstVals 
      Height          =   450
      Left            =   3885
      TabIndex        =   10
      Top             =   15
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Exit"
      Height          =   345
      Left            =   4875
      TabIndex        =   9
      Top             =   2370
      Width           =   1035
   End
   Begin VB.CommandButton cmddel 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4860
      TabIndex        =   8
      Top             =   1935
      Width           =   1035
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Enabled         =   0   'False
      Height          =   345
      Left            =   3615
      TabIndex        =   7
      Top             =   2370
      Width           =   1035
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Enabled         =   0   'False
      Height          =   345
      Left            =   3615
      TabIndex        =   6
      Top             =   1935
      Width           =   1035
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Left            =   3090
      TabIndex        =   5
      Top             =   1485
      Width           =   2760
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   3090
      MaxLength       =   10
      TabIndex        =   3
      Top             =   825
      Width           =   2760
   End
   Begin VB.ListBox lstVars 
      Height          =   2205
      Left            =   105
      TabIndex        =   0
      Top             =   510
      Width           =   2880
   End
   Begin VB.Label lblvardata 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Variable Value:"
      Height          =   195
      Left            =   3105
      TabIndex        =   4
      Top             =   1215
      Width           =   1065
   End
   Begin VB.Label lblvarname 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Variable Name:"
      Height          =   195
      Left            =   3090
      TabIndex        =   2
      Top             =   555
      Width           =   1080
   End
   Begin VB.Label lblVariables 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Variables"
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   270
      Width           =   1200
   End
End
Attribute VB_Name = "frmVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VarMemFile As String
Private aValues() As Variant

Sub ClearBoxes()
    txtName.Text = ""
    txtData.Text = ""
    cmddel.Enabled = False
    cmdUpdate.Enabled = False
End Sub

Function inList(lb As ListBox, Item As String) As Boolean
Dim x As Integer
    inList = False
    
    For x = 0 To lb.ListCount - 1
        If LCase(lb.List(x)) = LCase(Item) Then
            inList = True
            Exit For
        End If
    Next
    
End Function

Private Sub cmdAdd_Click()
    If inList(lstVars, txtName.Text) Then
        MsgBox "Variable ' " & txtName.Text & "' is already in the list.", vbInformation, "Item Already Found"
        Exit Sub
    Else
        lstVars.AddItem txtName.Text
        LstVals.AddItem txtData.Text
        txtName.Text = "": txtData.Text = ""
    End If
    
End Sub

Private Sub cmdclose_Click()
Dim StrA As String, x As Integer
On Error Resume Next
    'save variable list
    For x = 0 To lstVars.ListCount - 1
        StrA = StrA & lstVars.List(x) & "=" & LstVals.List(x) & vbCrLf
    Next x
        
    SetAttr VarMemFile, vbNormal
    Kill VarMemFile

    Open VarMemFile For Output As #1
        Print #1, StrA
    Close #1
        
    'Clean up
    lstVars.Clear: LstVals.Clear
    ClearBoxes
    'Reaload the variables
    frmExpression.ExprFx.LoadMem VarMemFile
    Unload frmVariables
    
End Sub

Sub OpenVarFile(lpFile As String)
Dim fp As Long, sIn As String, sVal As String, e_pos As Integer, iCnt As Integer
On Error Resume Next
    iCnt = 0
    
    fp = FreeFile
    Open lpFile For Input As #fp
        Do While Not EOF(fp)
            Input #fp, sIn
            If Len(Trim(sIn)) <> 0 Then
                e_pos = InStr(1, sIn, "=", vbBinaryCompare)
                If e_pos <> 0 Then
                    ReDim Preserve aValues(iCnt)
                    sVal = Trim$(Mid(sIn, e_pos + 1, Len(sIn))) 'Variable Value
                    sIn = Trim$(Left(sIn, e_pos - 1)) 'Variable Name
                    lstVars.AddItem sIn
                    LstVals.AddItem sVal
                    
                    'lstVars.ItemData(CVar(lstVars.ListCount - 1)) = sVal
                    aValues(iCnt) = sVal
                    iCnt = iCnt + 1
                    sVal = ""
                    sIn = ""
                End If
            End If
        Loop
    Close #fp
    
End Sub

Private Sub cmddel_Click()
Dim iSize As Integer
On Error Resume Next
Top:
    If lstVars.ListCount = 0 Then
        ClearBoxes
        Exit Sub
    End If
    
    If lstVars.ListIndex = -1 Then Exit Sub
    iSize = (lstVars.ListCount - 1)
    
    LstVals.RemoveItem lstVars.ListIndex
    lstVars.RemoveItem lstVars.ListIndex
   'LstVals.RemoveItem lstVars.ListIndex
    
    If lstVars.ListCount = 0 Then GoTo Top:
    
End Sub

Private Sub cmdUpdate_Click()
Dim idx As Integer
    idx = lstVars.ListIndex
    lstVars.List(idx) = txtName.Text
    LstVals.List(idx) = txtData.Text
End Sub

Private Sub Form_Load()
Dim x As Integer
    lstVars.Clear: LstVals.Clear
    VarMemFile = frmExpression.CalMem
    
    If Not frmExpression.FileExsits(VarMemFile) Then
        Exit Sub
    Else
        OpenVarFile VarMemFile
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmVariables = Nothing
End Sub

Private Sub lstVars_Click()
    cmddel.Enabled = True
    cmdUpdate.Enabled = True
    
    txtName.Text = lstVars.Text
    txtData.Text = LstVals.List(lstVars.ListIndex)
    
End Sub

Private Sub txtData_Change()
    txtName_Change
End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 8) Or (KeyAscii = 46) Then Exit Sub
    If Not IsNumeric(Chr(KeyAscii)) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub txtName_Change()
    If (Len(txtName.Text) = 0) Or (Len(txtData.Text) = 0) Then
        cmdAdd.Enabled = False
    Else
        cmdAdd.Enabled = True
    End If
   
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
Dim sCheck As String
    sCheck = UCase(Chr(KeyAscii))
    
    If KeyAscii = 8 Then Exit Sub
    
    If (sCheck < "A") Or (sCheck > "Z") Then
        Beep
        KeyAscii = 0
    End If
    
    
End Sub
