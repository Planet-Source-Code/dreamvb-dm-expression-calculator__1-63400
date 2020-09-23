VERSION 5.00
Begin VB.Form frmExpression 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Expression IT {Calculator}"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdVar 
      Caption         =   "Var"
      Height          =   360
      Left            =   4470
      TabIndex        =   34
      Top             =   2445
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   7170
      TabIndex        =   30
      Top             =   3105
      Width           =   7170
      Begin VB.CommandButton cmdexit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   6000
         TabIndex        =   33
         Top             =   75
         Width           =   1095
      End
      Begin VB.CommandButton cmdabout 
         Caption         =   "&About"
         Height          =   375
         Left            =   4800
         TabIndex        =   32
         Top             =   75
         Width           =   1095
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Show Answer"
         Height          =   375
         Left            =   3135
         TabIndex        =   31
         Top             =   75
         Width           =   1530
      End
   End
   Begin VB.CommandButton cmdOps 
      Caption         =   ")"
      Height          =   360
      Index           =   11
      Left            =   6555
      TabIndex        =   26
      Top             =   2445
      Width           =   510
   End
   Begin VB.CommandButton cmdOps 
      Caption         =   "("
      Height          =   360
      Index           =   10
      Left            =   6555
      TabIndex        =   25
      Top             =   1995
      Width           =   510
   End
   Begin VB.CommandButton cmdOps 
      Caption         =   ">"
      Height          =   360
      Index           =   9
      Left            =   6555
      TabIndex        =   24
      Top             =   1545
      Width           =   510
   End
   Begin VB.CommandButton cmdOps 
      Caption         =   "<"
      Height          =   360
      Index           =   8
      Left            =   6555
      TabIndex        =   23
      Top             =   1080
      Width           =   510
   End
   Begin VB.CommandButton cmdOps 
      Caption         =   "Or"
      Height          =   360
      Index           =   7
      Left            =   5910
      TabIndex        =   22
      Top             =   2445
      Width           =   510
   End
   Begin VB.CommandButton cmdOps 
      Caption         =   "mod"
      Height          =   360
      Index           =   6
      Left            =   5910
      TabIndex        =   21
      Top             =   1995
      Width           =   510
   End
   Begin VB.CommandButton cmdOps 
      Caption         =   "And"
      Height          =   360
      Index           =   5
      Left            =   5910
      TabIndex        =   20
      Top             =   1545
      Width           =   510
   End
   Begin VB.CommandButton cmdOps 
      Caption         =   "Pow"
      Height          =   360
      Index           =   4
      Left            =   5910
      TabIndex        =   19
      Top             =   1080
      Width           =   510
   End
   Begin VB.CommandButton cmdNums 
      Caption         =   "0"
      Height          =   360
      Index           =   9
      Left            =   4935
      TabIndex        =   18
      Top             =   2445
      Width           =   375
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C"
      Height          =   360
      Left            =   4005
      TabIndex        =   17
      Top             =   2445
      Width           =   375
   End
   Begin VB.CommandButton cmdOps 
      Caption         =   "+"
      Height          =   360
      Index           =   3
      Left            =   5415
      TabIndex        =   16
      Top             =   1995
      Width           =   375
   End
   Begin VB.CommandButton cmdOps 
      Caption         =   "-"
      Height          =   360
      Index           =   2
      Left            =   5415
      TabIndex        =   15
      Top             =   2445
      Width           =   375
   End
   Begin VB.CommandButton cmdOps 
      Caption         =   "*"
      Height          =   360
      Index           =   1
      Left            =   5415
      TabIndex        =   14
      Top             =   1545
      Width           =   375
   End
   Begin VB.CommandButton cmdNums 
      Caption         =   "9"
      Height          =   360
      Index           =   8
      Left            =   4935
      TabIndex        =   13
      Top             =   1995
      Width           =   375
   End
   Begin VB.CommandButton cmdNums 
      Caption         =   "8"
      Height          =   360
      Index           =   7
      Left            =   4470
      TabIndex        =   12
      Top             =   1995
      Width           =   375
   End
   Begin VB.CommandButton cmdNums 
      Caption         =   "7"
      Height          =   360
      Index           =   6
      Left            =   3990
      TabIndex        =   11
      Top             =   1995
      Width           =   375
   End
   Begin VB.CommandButton cmdNums 
      Caption         =   "6"
      Height          =   360
      Index           =   5
      Left            =   4935
      TabIndex        =   10
      Top             =   1545
      Width           =   375
   End
   Begin VB.CommandButton cmdNums 
      Caption         =   "5"
      Height          =   360
      Index           =   4
      Left            =   4470
      TabIndex        =   9
      Top             =   1545
      Width           =   375
   End
   Begin VB.CommandButton cmdNums 
      Caption         =   "4"
      Height          =   360
      Index           =   3
      Left            =   3990
      TabIndex        =   8
      Top             =   1545
      Width           =   375
   End
   Begin VB.CommandButton cmdOps 
      Caption         =   "/"
      Height          =   360
      Index           =   0
      Left            =   5415
      TabIndex        =   7
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmdNums 
      Caption         =   "3"
      Height          =   360
      Index           =   2
      Left            =   4935
      TabIndex        =   6
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmdNums 
      Caption         =   "2"
      Height          =   360
      Index           =   1
      Left            =   4470
      TabIndex        =   5
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmdNums 
      Caption         =   "1"
      Height          =   360
      Index           =   0
      Left            =   3990
      TabIndex        =   4
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtExpr 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Text            =   "2 * (3 + 4) * A / NumTen ^ (4 + pi )"
      Top             =   540
      Width           =   6315
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   -15
      X2              =   765
      Y1              =   3090
      Y2              =   3090
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   -15
      X2              =   780
      Y1              =   3075
      Y2              =   3075
   End
   Begin VB.Label lblBin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   240
      TabIndex        =   29
      Top             =   1860
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Binary Result:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   28
      Top             =   1650
      Width           =   1200
   End
   Begin VB.Label lblErr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   240
      TabIndex        =   27
      Top             =   2115
      Width           =   45
   End
   Begin VB.Label lblOut 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   240
      TabIndex        =   3
      Top             =   1365
      Width           =   90
   End
   Begin VB.Label lblResult 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Result:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1140
      Width           =   615
   End
   Begin VB.Label lblexpression 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expression:"
      Height          =   195
      Left            =   225
      TabIndex        =   0
      Top             =   240
      Width           =   810
   End
End
Attribute VB_Name = "frmExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ExprFx As clsExpr
Attribute ExprFx.VB_VarHelpID = -1
Dim CboTmp As String
Public CalMem As String 'Were all the Variables are kept
Dim xStart As Integer

Function Dec2Bin(IntDec As Long) As String
Dim sBinCheck As Boolean
Dim sBin As String
  ' Returns the binary string of an integer
    Do While IntDec <> 0
        sBinCheck = IntDec Mod 2
        If sBinCheck Then
            sBin = "1" & sBin
        Else
            sBin = "0" & sBin
        End If
        IntDec = IntDec \ 2
    Loop
    If sBin = "" Then
        Dec2Bin = "0"
    Else
        Dec2Bin = sBin
    End If
    sBin = ""
    
End Function

Private Sub cmdabout_Click()
    MsgBox frmExpression.Caption & vbCrLf & vbTab & "By Dreamvb", vbInformation, "About.."
End Sub

Private Sub cmdClear_Click()
    txtExpr.Text = "": txtExpr.SetFocus
End Sub

Private Sub cmdVar_Click()
    frmVariables.Show vbModal, frmExpression
End Sub

Private Sub ExprFx_ExprError(sError As String)
    Beep
    lblErr.Caption = "Error: " & sError
End Sub

Sub AddText(mText As String)
On Error Resume Next
Dim e As String

    txtExpr.SelStart = xStart
    
    If xStart <> 0 Then
        If Mid(txtExpr.Text, xStart, 1) <> " " Then
            e = " "
        Else
            e = ""
        End If
    End If
    
    txtExpr.SelText = e + mText
    txtExpr.SetFocus
    txtExpr_Click
End Sub

Public Function FileExsits(Filename As String) As Boolean
    'Find a file
    FileExsits = LenB(Dir(Filename)) <> 0
End Function

Private Sub cmdExit_Click()
    Unload frmExpression
End Sub

Private Sub cmdNums_Click(Index As Integer)
    AddText cmdNums(Index).Caption
End Sub

Private Sub cmdOps_Click(Index As Integer)
Dim addTx As String

    If Index = 4 Then
        addTx = "^"
    ElseIf Index = 5 Then
        addTx = "&"
    ElseIf Index = 6 Then
        addTx = "%"
    ElseIf Index = 7 Then
        addTx = "!"
    Else
        addTx = cmdOps(Index).Caption
    End If
        
    AddText addTx
    addTx = ""
    
End Sub

Private Sub cmdTest_Click()
Dim Result As Variant
    lblErr.Caption = ""
    ExprFx.Expr = txtExpr.Text
    ExprFx.InitExp
    Result = ExprFx.Expression
    lblOut = Result
    lblBin.Caption = Dec2Bin(CLng(Result))
End Sub

Function FixPath(lpPath As String) As String
    If Right(lpPath, 1) = "\" Then FixPath = lpPath Else FixPath = lpPath & "\"
End Function

Private Sub Form_Load()
    Set ExprFx = New clsExpr
    CalMem = FixPath(App.Path) & "variables.txt"
    Me.KeyPreview = True
    
    'Load in the calculators variables
    If FileExsits(CalMem) Then ExprFx.LoadMem CalMem



End Sub

Private Sub Form_Resize()
    Line1(0).X2 = frmExpression.ScaleWidth
    Line1(1).X2 = Line1(0).X2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DoEvents
    'Write the calulators memory
    CalMem = ""
    CboTmp = ""
    Set ExprFx = Nothing
    Set frmExpression = Nothing
End Sub

Private Sub txtExpr_Click()
    xStart = txtExpr.SelStart
End Sub

