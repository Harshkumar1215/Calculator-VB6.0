VERSION 5.00
Begin VB.Form calculatorx 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  Calculator"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "calculatorx"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   10320
      Top             =   3240
   End
   Begin VB.TextBox Hidden 
      Height          =   375
      Left            =   8640
      TabIndex        =   24
      Top             =   5640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox equal_dec 
      Height          =   285
      Left            =   9480
      TabIndex        =   23
      Text            =   "n"
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox dec_choice 
      Height          =   285
      Left            =   10680
      TabIndex        =   22
      Text            =   "n"
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox decider 
      Height          =   285
      Left            =   9000
      TabIndex        =   21
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox num2 
      Height          =   375
      Left            =   9000
      TabIndex        =   20
      Text            =   "0"
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox num1 
      Height          =   375
      Left            =   9000
      TabIndex        =   19
      Text            =   "0"
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton equals_old 
      Caption         =   "="
      Height          =   375
      Left            =   7440
      TabIndex        =   18
      ToolTipText     =   "Equals"
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Clear_old 
      Caption         =   "C"
      Height          =   375
      Left            =   8160
      TabIndex        =   17
      ToolTipText     =   "Clear Screen"
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton add_old 
      Caption         =   "+"
      Height          =   375
      Left            =   7200
      TabIndex        =   16
      ToolTipText     =   "Addition"
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton subtract_old 
      Caption         =   "-"
      Height          =   375
      Left            =   8160
      TabIndex        =   15
      ToolTipText     =   "Subtract"
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton multiply_old 
      Caption         =   "*"
      Height          =   375
      Left            =   10440
      TabIndex        =   14
      ToolTipText     =   "Multiply"
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton divide_old 
      Caption         =   "/"
      Height          =   375
      Left            =   9960
      TabIndex        =   13
      ToolTipText     =   "Divide"
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton decimal_old 
      Caption         =   "."
      Height          =   375
      Left            =   8040
      TabIndex        =   12
      ToolTipText     =   "Decimal Point"
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton plus_minus_old 
      Caption         =   "+/-"
      Height          =   375
      Left            =   7560
      TabIndex        =   11
      ToolTipText     =   "Turn number in to a minus number"
      Top             =   4440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton zero_old 
      Caption         =   "0"
      Height          =   375
      Left            =   7800
      TabIndex        =   10
      ToolTipText     =   "Zero"
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton nine_old 
      Caption         =   "9"
      Height          =   375
      Left            =   9840
      TabIndex        =   9
      ToolTipText     =   "Nine"
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton eight_old 
      Caption         =   "8"
      Height          =   375
      Left            =   9960
      TabIndex        =   8
      ToolTipText     =   "Eight"
      Top             =   4680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton seven_old 
      Caption         =   "7"
      Height          =   375
      Left            =   9000
      TabIndex        =   7
      ToolTipText     =   "Seven"
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton six_old 
      Caption         =   "6"
      Height          =   375
      Left            =   9000
      TabIndex        =   6
      ToolTipText     =   "Six"
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton five_old 
      Caption         =   "5"
      Height          =   375
      Left            =   9000
      TabIndex        =   5
      ToolTipText     =   "Five"
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton four_old 
      Caption         =   "4"
      Height          =   375
      Left            =   9000
      TabIndex        =   4
      ToolTipText     =   "Four"
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Three_old 
      Caption         =   "3"
      Height          =   375
      Left            =   9000
      TabIndex        =   3
      ToolTipText     =   "Three"
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton two_old 
      Caption         =   "2"
      Height          =   375
      Left            =   9000
      TabIndex        =   2
      ToolTipText     =   "Two"
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton one_old 
      Caption         =   "1"
      Height          =   375
      Left            =   9000
      TabIndex        =   1
      ToolTipText     =   "One"
      Top             =   4320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox display 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      ToolTipText     =   "Calculator Screen"
      Top             =   600
      Width           =   2415
   End
   Begin VB.Shape Shape19 
      Height          =   615
      Left            =   2400
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label Ans 
      Alignment       =   2  'Center
      Caption         =   "ANS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2475
      TabIndex        =   44
      Top             =   4800
      Width           =   495
   End
   Begin VB.Shape Shape18 
      Height          =   615
      Left            =   3240
      Top             =   600
      Width           =   615
   End
   Begin VB.Shape Shape17 
      Height          =   615
      Left            =   3240
      Top             =   1440
      Width           =   615
   End
   Begin VB.Shape Shape16 
      Height          =   615
      Left            =   3240
      Top             =   2280
      Width           =   615
   End
   Begin VB.Shape Shape15 
      Height          =   615
      Left            =   3240
      Top             =   3120
      Width           =   615
   End
   Begin VB.Shape Shape14 
      Height          =   615
      Left            =   3240
      Top             =   3960
      Width           =   615
   End
   Begin VB.Shape Shape13 
      Height          =   615
      Left            =   3240
      Top             =   4680
      Width           =   615
   End
   Begin VB.Shape Shape12 
      Height          =   615
      Left            =   2400
      Top             =   1440
      Width           =   615
   End
   Begin VB.Shape Shape11 
      Height          =   615
      Left            =   1560
      Top             =   1440
      Width           =   615
   End
   Begin VB.Shape Shape10 
      Height          =   615
      Left            =   720
      Top             =   1440
      Width           =   615
   End
   Begin VB.Shape Shape9 
      Height          =   615
      Left            =   2400
      Top             =   2280
      Width           =   615
   End
   Begin VB.Shape Shape8 
      Height          =   615
      Left            =   1560
      Top             =   2280
      Width           =   615
   End
   Begin VB.Shape Shape7 
      Height          =   615
      Left            =   720
      Top             =   2280
      Width           =   615
   End
   Begin VB.Shape Shape6 
      Height          =   615
      Left            =   2400
      Top             =   3120
      Width           =   615
   End
   Begin VB.Shape Shape5 
      Height          =   615
      Left            =   1560
      Top             =   3120
      Width           =   615
   End
   Begin VB.Shape Shape4 
      Height          =   615
      Left            =   720
      Top             =   3120
      Width           =   615
   End
   Begin VB.Shape Shape3 
      Height          =   615
      Left            =   2400
      Top             =   3960
      Width           =   615
   End
   Begin VB.Shape Shape2 
      Height          =   615
      Left            =   1560
      Top             =   3960
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   720
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label decimalx 
      Alignment       =   2  'Center
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   43
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label plus_minus 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   42
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label zero 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   41
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label equals 
      Alignment       =   2  'Center
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   40
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label add 
      Alignment       =   2  'Center
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   39
      Top             =   4080
      Width           =   375
   End
   Begin VB.Label subtract 
      Alignment       =   2  'Center
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   38
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label multiply 
      Alignment       =   2  'Center
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   37
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label divide 
      Alignment       =   2  'Center
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   36
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label clear 
      Alignment       =   2  'Center
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   35
      Top             =   720
      Width           =   375
   End
   Begin VB.Label nine 
      Alignment       =   2  'Center
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   34
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label eight 
      Alignment       =   2  'Center
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   33
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label seven 
      Alignment       =   2  'Center
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   32
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label six 
      Alignment       =   2  'Center
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   31
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label five 
      Alignment       =   2  'Center
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   30
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label four 
      Alignment       =   2  'Center
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   29
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label three 
      Alignment       =   2  'Center
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   28
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label two 
      Alignment       =   2  'Center
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   27
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label one 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   26
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label closex 
      Alignment       =   2  'Center
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   25
      Top             =   4920
      Width           =   615
   End
   Begin VB.Shape Shape20 
      Height          =   495
      Left            =   720
      Top             =   4800
      Width           =   855
   End
End
Attribute VB_Name = "calculatorx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add_Click()

num1 = display
decider.Text = "+"
display = ""
display.SetFocus

dec_choice = "n"

End Sub

Private Sub add_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

add.ForeColor = &HFF0000
Shape14.BorderColor = &HFF0000

End Sub

Private Sub Ans_Click()

display.Text = Hidden.Text

End Sub

Private Sub Ans_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Ans.ForeColor = &HFF0000
Shape19.BorderColor = &HFF0000

End Sub

Private Sub Clear_Click()

display = ""
num1 = ""
num2 = ""
decider = ""
dec_choice = "n"
equal_dec = "n"

display.SetFocus

End Sub

Private Sub clear_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

clear.ForeColor = &HFF0000
Shape18.BorderColor = &HFF0000

End Sub

Private Sub closex_Click()

Unload Me

End Sub

Private Sub closex_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

closex.ForeColor = &HFF0000
Shape20.BorderColor = &HFF0000

End Sub

Private Sub decimalx_Click()

On Error GoTo number_problem
    
    If dec_choice = "Y" Then
        Exit Sub
    End If
    
    dec_choice.Text = "Y"
    display = display.Text & "."

Exit Sub
number_problem:
    MsgBox Error & " (" & Err & ")"
Exit Sub

End Sub

Private Sub decimal_Click()

End Sub

Private Sub decimalx_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
decimalx.ForeColor = &HFF0000
Shape3.BorderColor = &HFF0000
End Sub

Private Sub display_KeyPress(KeyAscii As Integer)

On Error GoTo key_error

'MsgBox KeyAscii

If KeyAscii = 27 Then
    closex.ForeColor = &HFF0000
    Shape20.BorderColor = &HFF0000
    Unload Me
End If

If KeyAscii = 46 Then
    decimalx.ForeColor = &HFF0000
    Shape3.BorderColor = &HFF0000
End If

If KeyAscii = 48 Then
    zero.ForeColor = &HFF0000
    Shape1.BorderColor = &HFF0000
End If

If KeyAscii = 49 Then
    one.ForeColor = &HFF0000
    Shape10.BorderColor = &HFF0000
End If

If KeyAscii = 50 Then
    two.ForeColor = &HFF0000
    Shape11.BorderColor = &HFF0000
End If

If KeyAscii = 51 Then
    three.ForeColor = &HFF0000
    Shape12.BorderColor = &HFF0000
End If

If KeyAscii = 52 Then
    four.ForeColor = &HFF0000
    Shape7.BorderColor = &HFF0000
End If

If KeyAscii = 53 Then
    five.ForeColor = &HFF0000
    Shape8.BorderColor = &HFF0000
End If

If KeyAscii = 54 Then
    six.ForeColor = &HFF0000
    Shape9.BorderColor = &HFF0000
End If

If KeyAscii = 55 Then
    seven.ForeColor = &HFF0000
    Shape4.BorderColor = &HFF0000
End If

If KeyAscii = 56 Then
    eight.ForeColor = &HFF0000
    Shape5.BorderColor = &HFF0000
End If

If KeyAscii = 57 Then
    nine.ForeColor = &HFF0000
    Shape6.BorderColor = &HFF0000
End If

'#######################################
If KeyAscii = 42 Then
    multiply.ForeColor = &HFF0000
    Shape16.BorderColor = &HFF0000
    num1 = display
    decider.Text = "*"
    display = ""
    display.SetFocus
    dec_choice = "n"
End If

If KeyAscii = 43 Then
    add.ForeColor = &HFF0000
    Shape14.BorderColor = &HFF0000
    num1 = display
    decider.Text = "+"
    display = ""
    display.SetFocus
    dec_choice = "n"
End If

If KeyAscii = 45 Then
    subtract.ForeColor = &HFF0000
    Shape15.BorderColor = &HFF0000
    num1 = display
    decider.Text = "-"
    display = ""
    display.SetFocus
    dec_choice = "n"
End If

If KeyAscii = 47 Then
    divide.ForeColor = &HFF0000
    Shape17.BorderColor = &HFF0000
    num1 = display
    decider.Text = "/"
    display = ""
    display.SetFocus
    dec_choice = "n"
End If

If KeyAscii = 13 Then
    
    equals.ForeColor = &HFF0000
    Shape13.BorderColor = &HFF0000
    
    If decider <> "" Then
        num2.Text = display.Text
    End If

    equal_dec = "Y"
    
    If num1 And num2 <> "" And decider = "+" Then
        display.Text = Val(num1.Text) + Val(num2.Text)
    End If
    
    If num1 And num2 <> "" And decider = "-" Then
        display.Text = Val(num1.Text) - Val(num2.Text)
    End If
    
    If num1 And num2 <> "" And decider = "/" Then
        display.Text = Val(num1.Text) / Val(num2.Text)
    End If
    
    If num1 And num2 <> "" And decider = "*" Then
        display.Text = Val(num1.Text) * Val(num2.Text)
    End If
    
    Hidden.Text = display.Text
End If

If KeyAscii = 8 Then
    clear.ForeColor = &HFF0000
    Shape18.BorderColor = &HFF0000
    display = ""
    num1 = ""
    num2 = ""
    decider = ""
    dec_choice = "n"
    equal_dec = "n"
    display.SetFocus
End If


Timer1.Enabled = True
Timer1.Interval = 500

Exit Sub

key_error:
    MsgBox Error & " (" & Err & ")"
Exit Sub

End Sub

Private Sub display_KeyUp(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo problem
    'MsgBox KeyCode
    
    If KeyCode = 107 Then 'Or 106 Or 109 Or 111 Or 13
        display = ""
    End If
    
    If KeyCode = 109 Then 'Or 106 Or 109 Or 111 Or 13
        display = ""
    End If
    
    If KeyCode = 106 Then 'Or 106 Or 109 Or 111 Or 13
        display = ""
    End If
    
    If KeyCode = 111 Then 'Or 106 Or 109 Or 111 Or 13
        display = ""
    End If
    
    Exit Sub
    
problem:
    MsgBox Error & " (" & Err & ")"
Exit Sub

    
End Sub

Private Sub divide_Click()

num1 = display
decider.Text = "/"
display = ""
display.SetFocus

dec_choice = "n"

End Sub

Private Sub divide_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

divide.ForeColor = &HFF0000
Shape17.BorderColor = &HFF0000

End Sub

Private Sub eight_Click()

On Error GoTo number_problem

If equal_dec = "Y" Then
    display.Text = ""
    equal_dec = "n"
End If

    display = display.Text & 8

Exit Sub
number_problem:
    MsgBox Error & " (" & Err & ")"
Exit Sub

End Sub

Private Sub eight_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

eight.ForeColor = &HFF0000
Shape5.BorderColor = &HFF0000

End Sub

Private Sub equals_Click()

On Error GoTo problem

If decider <> "" Then
    num2.Text = display.Text
End If

equal_dec = "Y"

If num1 And num2 <> "" And decider = "+" Then
    display.Text = Val(num1.Text) + Val(num2.Text)
End If

If num1 And num2 <> "" And decider = "-" Then
    display.Text = Val(num1.Text) - Val(num2.Text)
End If

If num1 And num2 <> "" And decider = "/" Then
    display.Text = Val(num1.Text) / Val(num2.Text)
End If

If num1 And num2 <> "" And decider = "*" Then
    display.Text = Val(num1.Text) * Val(num2.Text)
End If

Hidden.Text = display.Text


Exit Sub

problem:
    If Err = 13 Then
        'MsgBox "You can only use numbers"
        num1.Text = ""
        num2.Text = ""
        display.Text = ""
        decider.Text = ""
        dec_choice = "n"
        Exit Sub
    End If
    
    MsgBox Error & " (" & Err & ")"
Exit Sub




End Sub

Private Sub equals_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

equals.ForeColor = &HFF0000
Shape13.BorderColor = &HFF0000

End Sub

Private Sub five_Click()

On Error GoTo number_problem

If equal_dec = "Y" Then
    display.Text = ""
    equal_dec = "n"
End If

    display = display.Text & 5

Exit Sub
number_problem:
    MsgBox Error & " (" & Err & ")"
Exit Sub

End Sub

Private Sub five_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

five.ForeColor = &HFF0000
Shape8.BorderColor = &HFF0000

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

one.ForeColor = &H80000012
two.ForeColor = &H80000012
three.ForeColor = &H80000012
four.ForeColor = &H80000012
five.ForeColor = &H80000012
six.ForeColor = &H80000012
seven.ForeColor = &H80000012
eight.ForeColor = &H80000012
nine.ForeColor = &H80000012
zero.ForeColor = &H80000012
plus_minus.ForeColor = &H80000012
add.ForeColor = &H80000012
subtract.ForeColor = &H80000012
divide.ForeColor = &H80000012
equals.ForeColor = &H80000012
decimalx.ForeColor = &H80000012
multiply.ForeColor = &H80000012
clear.ForeColor = &H80000012
Ans.ForeColor = &H80000012
closex.ForeColor = &H80000012

Shape1.BorderColor = &H80000012
Shape2.BorderColor = &H80000012
Shape3.BorderColor = &H80000012
Shape4.BorderColor = &H80000012
Shape5.BorderColor = &H80000012
Shape6.BorderColor = &H80000012
Shape7.BorderColor = &H80000012
Shape8.BorderColor = &H80000012
Shape9.BorderColor = &H80000012
Shape10.BorderColor = &H80000012
Shape11.BorderColor = &H80000012
Shape12.BorderColor = &H80000012
Shape13.BorderColor = &H80000012
Shape14.BorderColor = &H80000012
Shape15.BorderColor = &H80000012
Shape16.BorderColor = &H80000012
Shape17.BorderColor = &H80000012
Shape18.BorderColor = &H80000012
Shape19.BorderColor = &H80000012
Shape20.BorderColor = &H80000012

End Sub

Private Sub four_Click()

On Error GoTo number_problem

If equal_dec = "Y" Then
    display.Text = ""
    equal_dec = "n"
End If

    display = display.Text & 4

Exit Sub
number_problem:
    MsgBox Error & " (" & Err & ")"
Exit Sub

End Sub

Private Sub four_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

four.ForeColor = &HFF0000
Shape7.BorderColor = &HFF0000

End Sub

Private Sub multiply_Click()

num1 = display
decider.Text = "*"
display = ""
display.SetFocus

dec_choice = "n"

End Sub

Private Sub multiply_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

multiply.ForeColor = &HFF0000
Shape16.BorderColor = &HFF0000

End Sub

Private Sub nine_Click()

On Error GoTo number_problem

If equal_dec = "Y" Then
    display.Text = ""
    equal_dec = "n"
End If

    display = display.Text & 9

Exit Sub
number_problem:
    MsgBox Error & " (" & Err & ")"
Exit Sub

End Sub

Private Sub nine_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

nine.ForeColor = &HFF0000
Shape6.BorderColor = &HFF0000

End Sub

Private Sub one_Click()

On Error GoTo number_problem

If equal_dec = "Y" Then
    display.Text = ""
    equal_dec = "n"
End If

    display = display.Text & 1

Exit Sub
number_problem:
    MsgBox Error & " (" & Err & ")"
Exit Sub

End Sub

Private Sub one_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

one.ForeColor = &HFF0000
Shape10.BorderColor = &HFF0000

End Sub

Private Sub plus_minus_Click()

On Error GoTo prob

display.Text = display.Text - display.Text - display.Text

Exit Sub

prob:
    If Err = 13 Then
        Exit Sub
    Else
        MsgBox Error & " (" & Err & ")"
    Exit Sub
    
    End If
    

End Sub

Private Sub plus_minus_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

plus_minus.ForeColor = &HFF0000
Shape2.BorderColor = &HFF0000

End Sub

Private Sub seven_Click()

On Error GoTo number_problem

If equal_dec = "Y" Then
    display.Text = ""
    equal_dec = "n"
End If

    display = display.Text & 7

Exit Sub
number_problem:
    MsgBox Error & " (" & Err & ")"
Exit Sub

End Sub

Private Sub seven_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

seven.ForeColor = &HFF0000
Shape4.BorderColor = &HFF0000

End Sub

Private Sub six_Click()

On Error GoTo number_problem

If equal_dec = "Y" Then
    display.Text = ""
    equal_dec = "n"
End If

    display = display.Text & 6

Exit Sub
number_problem:
    MsgBox Error & " (" & Err & ")"
Exit Sub

End Sub

Private Sub six_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

six.ForeColor = &HFF0000
Shape9.BorderColor = &HFF0000

End Sub

Private Sub subtract_Click()

num1 = display
decider.Text = "-"
display = ""
display.SetFocus

dec_choice = "n"

End Sub

Private Sub subtract_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

subtract.ForeColor = &HFF0000
Shape15.BorderColor = &HFF0000

End Sub

Private Sub Three_Click()

On Error GoTo number_problem

If equal_dec = "Y" Then
    display.Text = ""
    equal_dec = "n"
End If
    display = display.Text & 3

Exit Sub

number_problem:
    MsgBox Error & " (" & Err & ")"
Exit Sub

End Sub

Private Sub three_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

three.ForeColor = &HFF0000
Shape12.BorderColor = &HFF0000

End Sub

Private Sub Timer1_Timer()

one.ForeColor = &H80000012
two.ForeColor = &H80000012
three.ForeColor = &H80000012
four.ForeColor = &H80000012
five.ForeColor = &H80000012
six.ForeColor = &H80000012
seven.ForeColor = &H80000012
eight.ForeColor = &H80000012
nine.ForeColor = &H80000012
zero.ForeColor = &H80000012
plus_minus.ForeColor = &H80000012
add.ForeColor = &H80000012
subtract.ForeColor = &H80000012
divide.ForeColor = &H80000012
equals.ForeColor = &H80000012
decimalx.ForeColor = &H80000012
multiply.ForeColor = &H80000012
clear.ForeColor = &H80000012
Ans.ForeColor = &H80000012

Shape1.BorderColor = &H80000012
Shape2.BorderColor = &H80000012
Shape3.BorderColor = &H80000012
Shape4.BorderColor = &H80000012
Shape5.BorderColor = &H80000012
Shape6.BorderColor = &H80000012
Shape7.BorderColor = &H80000012
Shape8.BorderColor = &H80000012
Shape9.BorderColor = &H80000012
Shape10.BorderColor = &H80000012
Shape11.BorderColor = &H80000012
Shape12.BorderColor = &H80000012
Shape13.BorderColor = &H80000012
Shape14.BorderColor = &H80000012
Shape15.BorderColor = &H80000012
Shape16.BorderColor = &H80000012
Shape17.BorderColor = &H80000012
Shape18.BorderColor = &H80000012
Shape19.BorderColor = &H80000012

End Sub

Private Sub two_Click()

 On Error GoTo number_problem

If equal_dec = "Y" Then
    display.Text = ""
    equal_dec = "n"
End If
    display = display.Text & 2

Exit Sub

number_problem:
    MsgBox Error & " (" & Err & ")"
Exit Sub

End Sub

Private Sub two_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

two.ForeColor = &HFF0000
Shape11.BorderColor = &HFF0000

End Sub

Private Sub zero_Click()

On Error GoTo number_problem

If equal_dec = "Y" Then
    display.Text = ""
    equal_dec = "n"
End If

    display = display.Text & 0

Exit Sub
number_problem:
    MsgBox Error & " (" & Err & ")"
Exit Sub

End Sub

Private Sub zero_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

zero.ForeColor = &HFF0000
Shape1.BorderColor = &HFF0000

End Sub


