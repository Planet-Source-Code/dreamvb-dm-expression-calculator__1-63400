VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'VER 1
'Hi all this is a simple expression Phaser, That I got fed up of working on in VB
'This project came about when I was reading though a compiler book in C and decided to convert it over to VB
'It can just works things like + - *\ ^ and can also process things in brackets like 2 * (5 + 5)
'Note returns Integer result. nor does it support floating points
'Tho you can add this with out to much problums
'Anyway hope you find it useful.

'VER 2
' Little small update I made to support varisbles
' Your variables are also saved to the a file in the same folder when you exit. so you can;t lose them
' You also have a built in Variable editor to edit or delete add, etc your variables

' Add some more operators & % ! Thay are AND MOD and OR also added / < > =
' Updated the GUI so it look a little more Friendly to use.
' Added error event to show any errors
' Last but not least Added a binary output of your result from the expression
' If I decide to-update this agian I work on functions

' Any new updates to all my programs at my forum
' http://www.eraystudios.com/forum

Dim e_pos As Integer
Dim look As String
Dim iExpression As String
Dim iGetNum  As String

Dim bErrFlag As Boolean
Public Event ExprError(sError As String)

'Added to support variables
Private Type VarStack
    mVarName As String
    mVarValue As Variant
End Type

Private VarCounter As Integer
Private Variables() As VarStack
Private Const NO_VARIABLE = -1

Sub ResetVarStack()
    'Erase variables
    VarCounter = 0
    Erase Variables
End Sub

Function AddVar(lpVarName As String, Optional lpVarData As Variant = 0)
    'Adds a new variable to the variable stack
    ReDim Preserve Variables(VarCounter)
    Variables(VarCounter).mVarName = lpVarName
    Variables(VarCounter).mVarValue = lpVarData
    VarCounter = VarCounter + 1
End Function

Function VarIndex(varname As String) As Integer
Dim idx As Integer, x As Integer
On Error GoTo nError:
    idx = NO_VARIABLE
    'Return a variables index in the variables stack
    For x = 0 To VarCounter
        If LCase$(varname) = LCase(Variables(x).mVarName) Then
            idx = x
            Exit For
        End If
    Next x
    
    x = 0
    
    VarIndex = idx
    Exit Function
nError:
    
    VarIndex = NO_VARIABLE
    
End Function

Sub GetChar()
    'Sb used to get the next char from the input string
    e_pos = e_pos + 1 'Add one to our char pos counter
    look = Mid(iExpression, e_pos, 1) 'Store char x into Look
End Sub

Private Sub SkipWhite()
    'Jump over white spaces
    Do While (isWhite(look))
        GetChar 'Call GetChar to get the next char
    Loop
End Sub

Private Function isWhite(c As String) As Integer
    'Check if c is a white space or a tab
    If ((c = " ") Or (c = vbTab)) Then
        isWhite = 1
    Else
        isWhite = 0
    End If
End Function

Private Function isAlpha(c As String) As Integer
    'Checks to see if c is alpha ex a-z or A-Z
    Select Case c
        Case "A" To "Z": isAlpha = 1
        Case "a" To "z": isAlpha = 1
        Case Else: isAlpha = 0
    End Select
End Function

Private Function isDigit(c As String) As Integer
    'Used to check if c is a digit 0-9
    Select Case c
        Case 0 To 9: isDigit = 1
        Case Else: isDigit = 0
    End Select
End Function

Private Sub Match(x As String)
    'check if look is euale to x if it is Get the next char
    If (look = x) Then
        GetChar 'Get next char
        SkipWhite ' Jump over white spaces
    Else
        RaiseEvent ExprError("Math error (" & look & ")")
    End If
End Sub

Private Function GetName() As Variant
Dim lName As Integer, Token As String
    Token = ""
    
    If isAlpha(look) = 0 Then RaiseEvent ExprError("Expected Name")
    
    Do While isAlpha(look)
        'While we only have alpha build a token
        Token = Token & look
        GetChar
    Loop
    
    GetName = Token
    SkipWhite
End Function

Private Function GetNum() As Variant
Dim Value As Variant
    'Return a number found in look
    If isDigit(look) = 0 Then
        RaiseEvent ExprError("Invaild Digit found " & "'" & look & "'")
        bErrFlag = True
        Exit Function
    End If
    
    'Look while we have a digit
    Do While (isDigit(look)) = 1
        Value = 10 * Value + CDbl(look) - 0 'return value
        GetChar 'Get the next look char
    Loop
    SkipWhite 'Jump over white space
    GetNum = Value 'Return value
    
End Function

Function Expression() As Variant
Dim Value As Variant
On Error GoTo nErr:
    If bErrFlag Then Exit Function
    
    If (IsAddOp(look)) = 1 Then
        Value = 0
    Else
        Value = Term
        'Loop while we have a add-op
        Do While (IsAddOp(look)) = 1
            Select Case look
                Case "+"
                    Match ("+")
                    Value = Value + Term
                Case "-"
                    Match ("-")
                    Value = Value - Term
                
            End Select
        Loop
    End If
    
    Expression = Value
    Exit Function

nErr:
    If Not bErrFlag Then RaiseEvent ExprError(Err.Description)
End Function

Private Function IsAddOp(c As String) As Integer
    'Check to see if we have a add op + or -
    If ((c = "+") Or (c = "-")) Then
        IsAddOp = 1
    Else
        IsAddOp = 0
    End If
End Function

Private Sub init()
    bErrFlag = False
    If Len(Trim(iExpression)) = 0 Then
        bErrFlag = True
        RaiseEvent ExprError("Expression is empty")
        Exit Sub
    Else
        GetChar
        SkipWhite
    End If
End Sub

Private Function Term() As Variant
Dim Value As Variant

    Value = Factor() 'Get Factor
    
    Do While ((look = "*") Or (look = "\") Or (look = "^") Or _
    (look = ">") Or (look = "<") Or (look = "=") Or (look = "%") _
    Or (look = "!") Or (look = "&") Or (look = "/") Or (look = "."))
        
        Select Case look
            Case "*"
                Match ("*")
                Value = Value * Factor()
            Case "\"
                Match ("\")
                Value = Value \ Factor()
            Case "^"
                Match ("^")
                Value = Value ^ Factor()
            Case ">"
                Match (">")
                Value = Abs(Value > Factor())
            Case "<"
                Match ("<")
                Value = Abs(Value < Factor())
            Case "="
                Match ("=")
                Value = Abs(Value = Factor())
            Case "%"
                Match ("%")
                Value = Abs(Value Mod Factor())
            Case "!"
                Match ("!")
                Value = Abs(Value Or Factor())
            Case "&"
                Match ("&")
                Value = Abs(Value And Factor())
            Case "/"
                Match ("/")
                Value = Value / Factor()
            Case "."
                Match (".")
                Value = Value & Factor()
        End Select
    Loop
    
    Term = Value
    
End Function

Private Function Factor() As Variant
Dim Value As Variant
Dim idx As Integer
    'Used to work with bracets ()
    If (look = "(") Then
        Match ("(")
        Factor = Expression()
        Match (")")
    ElseIf isAlpha(look) Then
        idx = VarIndex(GetName) 'Get Variable Index
        If (idx = NO_VARIABLE) Then
            'Opps no variable found so we must show a error
            RaiseEvent ExprError("Expected Variable")
        Else
            'Return the variables value
            Factor = Variables(idx).mVarValue
        End If
        SkipWhite 'Skip white spaces
    Else
        Factor = GetNum 'Return a Number
    End If
    
End Function

Public Sub InitExp()
    e_pos = 0
    iGetNum = ""
    look = ""
    cGetName = ""
    init
End Sub

Public Property Get Expr() As String
    Expr = iExpression
End Property

Public Property Let Expr(ByVal vNewValue As String)
    iExpression = vNewValue
End Property

Public Property Get Variable(Index As Integer) As Variant
    'Return variable
    Variable = VarStack(Index)
End Property

Public Property Let Variable(Index As Integer, ByVal vNewVal As Variant)
    'Add variable
    VarStack(Index) = vNewVal
End Property

Public Sub LoadMem(lpFile As String)
Dim fp As Long, sIn As String, sVal As String, e_pos As Integer
On Error Resume Next
    Call ResetVarStack
    'Load the variables onto the variable stack
    fp = FreeFile
    Open lpFile For Input As #fp
        Do While Not EOF(fp)
            Input #fp, sIn
            If Len(Trim(sIn)) <> 0 Then
                e_pos = InStr(1, sIn, "=", vbBinaryCompare)
                If e_pos <> 0 Then
                    sVal = Trim$(Mid(sIn, e_pos + 1, Len(sIn))) 'Variable Value
                    sIn = Trim$(Left(sIn, e_pos - 1)) 'Variable Name
                    AddVar sIn, sVal
                    sVal = ""
                    sIn = ""
                End If
            End If
        Loop
    Close #fp

End Sub


