VERSION 5.00
Begin VB.Form calculator 
   BackColor       =   &H00000000&
   Caption         =   "Standard Calculator"
   ClientHeight    =   8055
   ClientLeft      =   6330
   ClientTop       =   1620
   ClientWidth     =   4335
   FillColor       =   &H000080FF&
   ForeColor       =   &H00C000C0&
   HasDC           =   0   'False
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   537
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   289
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOneDividedByX 
      BackColor       =   &H0000FF00&
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   23
      Left            =   3120
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "One divided by x"
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdSquare 
      BackColor       =   &H0000FF00&
      Caption         =   "x²"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   22
      Left            =   2160
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Square"
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdSquareRoot 
      BackColor       =   &H0000FF00&
      Caption         =   "sqrt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   21
      Left            =   1200
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Squareroot"
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdModulus 
      BackColor       =   &H0000FF00&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   20
      Left            =   240
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Modulus"
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H0000FF00&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   18
      Left            =   1200
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Clear"
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmd2 
      BackColor       =   &H0000FF00&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   1200
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H0000FF00&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   240
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton cmd5 
      BackColor       =   &H0000FF00&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   5
      Left            =   1200
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton cmd8 
      BackColor       =   &H0000FF00&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   8
      Left            =   1200
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdNeg 
      BackColor       =   &H0000FF00&
      Caption         =   "±"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   10
      Left            =   240
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Negative or positive"
      Top             =   6840
      Width           =   855
   End
   Begin VB.CommandButton cmdEquals 
      BackColor       =   &H0000FF00&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   12
      Left            =   3120
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Equals"
      Top             =   6840
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton cmdDivide 
      BackColor       =   &H0000FF00&
      Caption         =   "÷"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   16
      Left            =   3120
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Division"
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdMultiply 
      BackColor       =   &H0000FF00&
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   15
      Left            =   3120
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Multiplication"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdSubtract 
      BackColor       =   &H0000FF00&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   14
      Left            =   3120
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Subtraction"
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H0000FF00&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   13
      Left            =   3120
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Addition"
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton cmdDot 
      BackColor       =   &H0000FF00&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   11
      Left            =   2160
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6840
      Width           =   855
   End
   Begin VB.CommandButton cmd0 
      BackColor       =   &H0000FF00&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   1200
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6840
      Width           =   855
   End
   Begin VB.CommandButton cmd3 
      BackColor       =   &H0000FF00&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   2160
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton cmd6 
      BackColor       =   &H0000FF00&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   6
      Left            =   2160
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton cmd4 
      BackColor       =   &H0000FF00&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   240
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton cmd9 
      BackColor       =   &H0000FF00&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   9
      Left            =   2160
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmd7 
      BackColor       =   &H0000FF00&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   7
      Left            =   240
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdClearEquation 
      BackColor       =   &H0000FF00&
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   17
      Left            =   240
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Clear equation"
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdBackSpace 
      BackColor       =   &H0000FF00&
      Caption         =   "‹-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   19
      Left            =   2160
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Back space"
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox txtAnswer 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FF00&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   720
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "0"
      Top             =   840
      Width           =   3735
   End
   Begin VB.Line Line13 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Index           =   3
      X1              =   272
      X2              =   280
      Y1              =   128
      Y2              =   136
   End
   Begin VB.Line Line13 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Index           =   2
      X1              =   272
      X2              =   280
      Y1              =   120
      Y2              =   128
   End
   Begin VB.Line Line13 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Index           =   1
      X1              =   272
      X2              =   280
      Y1              =   112
      Y2              =   120
   End
   Begin VB.Line Line13 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Index           =   0
      X1              =   272
      X2              =   280
      Y1              =   48
      Y2              =   56
   End
   Begin VB.Line Line12 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   272
      X2              =   280
      Y1              =   520
      Y2              =   528
   End
   Begin VB.Line Line11 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   16
      X2              =   280
      Y1              =   528
      Y2              =   528
   End
   Begin VB.Line Line10 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   8
      X2              =   16
      Y1              =   520
      Y2              =   528
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   280
      X2              =   280
      Y1              =   16
      Y2              =   528
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   272
      X2              =   280
      Y1              =   8
      Y2              =   16
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Index           =   1
      X1              =   8
      X2              =   272
      Y1              =   8
      Y2              =   8
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Standard Calculator"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   25
      ToolTipText     =   "Click me to show the about page."
      Top             =   240
      Width           =   3735
   End
   Begin VB.Line Line9 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   8
      X2              =   8
      Y1              =   8
      Y2              =   520
   End
   Begin VB.Line Line8 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   272
      X2              =   272
      Y1              =   8
      Y2              =   520
   End
   Begin VB.Line Line7 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   8
      X2              =   272
      Y1              =   48
      Y2              =   48
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   8
      X2              =   272
      Y1              =   112
      Y2              =   112
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   8
      X2              =   272
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Index           =   0
      X1              =   8
      X2              =   272
      Y1              =   520
      Y2              =   520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   8
      X2              =   272
      Y1              =   128
      Y2              =   128
   End
End
Attribute VB_Name = "calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================
'= PROJECT IN PROGRAMMING    =
'= VISUAL BASIC CALCULATOR   =
'= COPYRIGHT not COPY PASTE  =
'= MARCH 9, 2019             =
'=============================

' Used built-in functions:
'  * Len(string) -> returns string's length
'  * Mid(string, start, end) -> returns a substring
'  * InStr(start, string, string_to_find) -> returns index of found string
'
'<Sequences>
'
' Number button is pressed:
'     * Functions:
'       - handle_pressed_num_button
'       - check_if_zero
'     * Human:
'       - If `txtAnswer` is 0, then set `txtAnswer` to nothing
'       - Append the number at the end of `txtAnswer`
' Operator button is pressed:
'     * Functions:
'       - handle_pressed_operator_button
'       - do_calculation
'       - zero_txtAsnwer_and_False_pressedDot
'     * Human:
'       - If operation is to add, then add
'       - Else if operation is to minus, then minus
'       - Else if operation is to multiply, then multiply
'       - Else if operation is to divide, then divide
'       - Else if operation is to get the modulus, then get that
'       - Else if there is no operation yet, assign the `txtAnswer` to `tempNum`
'       - Assign the `op` to `operator`
'       - Set `txtAnswer` to 0
'       - Set `pressedDot` to False
' Equals button is pressed:
'     * Function
'       - do_calculation
'     * Human:
'       - Do the calculation
'       - Assign `tempNum` to `txtAnswer`
'       - Set `tempNum` to 0
'       - Set `operator` to nothing
'       - If there is NO dot in `txtAnswer`, then set `pressedDot` to False
'       - If there is, then set `pressedDot` to True
'
'</Sequences>


'Global variable definitions
Dim tempNum As Double
Dim operator As String
Dim pressedDot As Boolean


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 97 Then
        'Hide the `calculator` form
        calculator.Hide
        'Show the `About` form
        About.Show
    End If
End Sub

'When form loads,
Private Sub Form_Load()
    'These will be set to these
    txtAnswer = 0
    tempNum = 0
    pressedDot = False
    operator = ""
End Sub


'Removes the right most number
Private Sub cmdBackSpace_Click(Index As Integer)
    'If `txtAnswer` is not 0
    If Not txtAnswer = "0" Then
        'Remove the number at the very right end
        txtAnswer = Mid(txtAnswer, 1, Len(txtAnswer) - 1)
    End If
    'If answer box has nothing or it only has a dash
    If txtAnswer = "" Or txtAnswer = "-" Then
        'Set `txtAnswer` to 0
        txtAnswer = "0"
    End If
End Sub


'Clears the `txtAnswer`
Private Sub cmdClearEquation_Click(Index As Integer)
    txtAnswer = "0"
End Sub


'Clears the answer box and defaults all global vars
Private Sub cmdClear_Click(Index As Integer)
    zero_txtAsnwer_and_False_pressedDot
    operator = ""
    tempNum = 0
End Sub


'Checks if `txtAnswer` is 0
Private Sub check_if_zero()
    'If `txtAnswer` is 0
    If txtAnswer = "0" Then
        'Set `txtAnswer` to nothing
        txtAnswer = ""
    End If
End Sub


'Sets `txtAnswer` to 0 and `pressedDot` to False
Private Sub zero_txtAsnwer_and_False_pressedDot()
    txtAnswer = "0"
    pressedDot = False
End Sub


'Does the calculation
Private Sub do_calculation()
'Turns on error catching
On Error GoTo area_that_says_error
    If operator = "+" Then
        'For addition
        tempNum = tempNum + Val(txtAnswer)
    ElseIf operator = "-" Then
        'For subtraction
        tempNum = tempNum - Val(txtAnswer)
    ElseIf operator = "*" Then
        'For multiplication
        tempNum = tempNum * Val(txtAnswer)
    ElseIf operator = "/" Then
        'For division
        'If `txtAnswer` is 0
        If Not Val(txtAnswer) = 0 Then
            'Divide `tempNum` to `txtAnswer`
            tempNum = tempNum / Val(txtAnswer)
        Else
            'If `txtAnswer` is zero, output the error
            show_error_for "division"
        End If
    ElseIf operator = "Mod" Then
        'For modulus.
        'This is where the overflow "might" happen
        tempNum = tempNum Mod Val(txtAnswer)
    Else
        'If there's no operator yet, `txtAnswer` value will
        'be assigned to `tempNum`
        tempNum = Val(txtAnswer)
    End If
Exit Sub
'Execution jumps here if there is an overflow error
area_that_says_error:
    show_error_for "overflow"
End Sub


'For limiting numbers in `txtAnswer`
Private Sub txtAnswer_Change()
    'If `txtAnswer` length is greater than 10
    If Len(txtAnswer) > 10 Then
        'Limit `txtAnswer` to 10 numbers only
        txtAnswer = Mid(txtAnswer, 1, 10)
    End If
End Sub


'Auxiliary function for reducing code redundancy
Private Sub handle_pressed_num_button(Num As String)
    'Check first if `txtAnswer` is 0
    check_if_zero
    'Append the number at the end of `txtAnswer`
    txtAnswer = txtAnswer & Num
End Sub


'Auxiliary function for reducing code redundancy
Private Sub handle_pressed_operator_button(op As String)
    'Do the calculation first
    do_calculation
    'Assign the `op` to `operator`
    operator = op
    'Empty `txtAnswer` and set `pressedDot` to False
    zero_txtAsnwer_and_False_pressedDot
End Sub


'Function for showing different kinds of errors
Private Sub show_error_for(Err As String)
    Select Case Err
        Case "overflow"
            MsgBox "That is an overflow.", vbOKOnly + vbExclamation, "Error"
        Case "division"
            MsgBox "You cannot divide any number to zero.", vbOKOnly + vbExclamation, "Error"
        Case "negative"
            MsgBox "You cannot get the square root of negative numbers.", vbOKOnly + vbExclamation, "Error"
    End Select
    'In every error, the program defaults everything
    cmdClear_Click (18)
End Sub


'Toggle function for making `txtAnswer` value negative or positive
Private Sub cmdNeg_Click(Index As Integer)
    'If `txtAnswer` is already negative
    If Mid(txtAnswer, 1, 1) = "-" Then
        'Remove the dash
        txtAnswer = Mid(txtAnswer, 2, Len(txtAnswer))
    'If `txtAnswer` is not 0
    ElseIf Not txtAnswer = "0" Then
        'Append a dash at the left most end
        txtAnswer = "-" & txtAnswer
    End If
End Sub


'Adding a dot to `txtAnswer`
Private Sub cmdDot_Click(Index As Integer)
   'If `pressedDot` is False
    If Not pressedDot Then
        'Append the dot
        txtAnswer = txtAnswer & "."
        'Set `pressedDot` to True
        pressedDot = True
    End If
End Sub


'<Numbers>
Private Sub cmd0_Click(Index As Integer)
    handle_pressed_num_button "0"
End Sub
Private Sub cmd1_Click(Index As Integer)
    handle_pressed_num_button "1"
End Sub
Private Sub cmd2_Click(Index As Integer)
    handle_pressed_num_button "2"
End Sub
Private Sub cmd3_Click(Index As Integer)
    handle_pressed_num_button "3"
End Sub
Private Sub cmd4_Click(Index As Integer)
    handle_pressed_num_button "4"
End Sub
Private Sub cmd5_Click(Index As Integer)
    handle_pressed_num_button "5"
End Sub
Private Sub cmd6_Click(Index As Integer)
    handle_pressed_num_button "6"
End Sub
Private Sub cmd7_Click(Index As Integer)
    handle_pressed_num_button "7"
End Sub
Private Sub cmd8_Click(Index As Integer)
    handle_pressed_num_button "8"
End Sub
Private Sub cmd9_Click(Index As Integer)
    handle_pressed_num_button "9"
End Sub
'<Numbers/>



'<Operators>
Private Sub cmdAdd_Click(Index As Integer)
    handle_pressed_operator_button "+"
End Sub
Private Sub cmdSubtract_Click(Index As Integer)
    handle_pressed_operator_button "-"
End Sub
Private Sub cmdMultiply_Click(Index As Integer)
    handle_pressed_operator_button "*"
End Sub
Private Sub cmdDivide_Click(Index As Integer)
    handle_pressed_operator_button "/"
End Sub
Private Sub cmdModulus_Click(Index As Integer)
    handle_pressed_operator_button "Mod"
End Sub
'<Operators/>


'Equals
Private Sub cmdEquals_Click(Index As Integer)
    'Do the calculation first
    do_calculation
    'Assign `tempNum` to `txtAnswer`
    txtAnswer = tempNum
    'Set `tempNum` to 0
    tempNum = 0
    'Set `operator` to nothing
    operator = ""
    'If there is NO dot in `txtAnswer`
    If InStr(1, txtAnswer, ".") = 0 Then
        'Set `pressedDot` to False
        pressedDot = False
    Else
        'Set `pressedDot` to True
        pressedDot = True
    End If
End Sub


'Gets the square root
Private Sub cmdSquareRoot_Click(Index As Integer)
    'If `txtAnswer` is negative
    If Val(txtAnswer) < 0 Then
        'Show the error
        show_error_for "negative"
    Else
        'Assign the square root to `txtAnswer` itself
        txtAnswer = Val(txtAnswer) ^ 0.5
    End If
End Sub


'Gets the square
Private Sub cmdSquare_Click(Index As Integer)
'Turns on error handling catching
On Error GoTo area_that_says_error
    'If it overflows here,
    txtAnswer = Val(txtAnswer) ^ 2
Exit Sub
'Then execution will jump here
area_that_says_error:
    show_error_for "overflow"
End Sub


'Divides by one
Private Sub cmdOneDividedByX_Click(Index As Integer)
'Turns on error handling catching
On Error GoTo area_that_says_error
    'If divided by zero,
    txtAnswer = 1 / Val(txtAnswer)
Exit Sub
'Then execution will jump here
area_that_says_error:
    show_error_for "division"
End Sub


'About
Private Sub Label1_Click()
    'Hide the `calculator` form
    calculator.Hide
    'Show the `About` form
    About.Show
End Sub


'When application terminates
Private Sub Form_Terminate()
    MsgBox "Bye! Have a nice day!", vbInformation + vbOKOnly, "Information"
End Sub
