VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Technology Calculator"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7035
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "calculator.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   38
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Function Data"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4920
      TabIndex        =   33
      Top             =   0
      Width           =   1815
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1605
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CALCULATOR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   32
      Top             =   120
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   2280
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TIME"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   31
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSin 
      Caption         =   "Sin"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4800
      TabIndex        =   30
      Top             =   3195
      Width           =   540
   End
   Begin VB.CommandButton cmdCos 
      Caption         =   "Cos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4800
      TabIndex        =   29
      Top             =   2565
      Width           =   540
   End
   Begin VB.CommandButton cmdSquareRoot 
      Caption         =   "sqrt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4800
      TabIndex        =   28
      Top             =   1935
      Width           =   540
   End
   Begin VB.Frame Frame2 
      Caption         =   "Memory Contents"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3840
      TabIndex        =   26
      Top             =   4680
      Width           =   2535
      Begin VB.Label lblMemory1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Left            =   105
         TabIndex        =   27
         Top             =   210
         Width           =   2325
      End
   End
   Begin VB.CommandButton cmdMemory3 
      Caption         =   "MC"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5670
      TabIndex        =   25
      Top             =   1920
      Width           =   540
   End
   Begin VB.CommandButton cmdMemory2 
      Caption         =   "MR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5670
      TabIndex        =   24
      Top             =   2445
      Width           =   540
   End
   Begin VB.CommandButton cmdMemory1 
      Caption         =   "M+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5670
      TabIndex        =   23
      Top             =   2970
      Width           =   540
   End
   Begin VB.Frame Frame1 
      Caption         =   "Answer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1080
      TabIndex        =   21
      Top             =   4680
      Width           =   2535
      Begin VB.Label lblRunning 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   330
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   2325
      End
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "AC"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   20
      Top             =   1080
      Width           =   900
   End
   Begin VB.CommandButton cmdEquals 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3495
      TabIndex        =   19
      Top             =   3825
      Width           =   1170
   End
   Begin VB.CommandButton cmdDivide 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4125
      TabIndex        =   18
      Top             =   3195
      Width           =   540
   End
   Begin VB.CommandButton cmdTimes 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3495
      TabIndex        =   17
      Top             =   3195
      Width           =   540
   End
   Begin VB.CommandButton cmdMinus 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4125
      TabIndex        =   16
      Top             =   2565
      Width           =   540
   End
   Begin VB.CommandButton cmdPlusMinus 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3495
      TabIndex        =   15
      Top             =   1935
      Width           =   540
   End
   Begin VB.CommandButton cmdOver 
      Caption         =   "1/X"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4125
      TabIndex        =   14
      Top             =   1935
      Width           =   540
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3495
      TabIndex        =   13
      Top             =   2565
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4800
      TabIndex        =   12
      Top             =   3825
      Width           =   540
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "CE / C"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   900
   End
   Begin VB.CommandButton cmdDigits 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   9
      Left            =   2670
      TabIndex        =   10
      Top             =   3195
      Width           =   540
   End
   Begin VB.CommandButton cmdDigits 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   8
      Left            =   2040
      TabIndex        =   9
      Top             =   3195
      Width           =   540
   End
   Begin VB.CommandButton cmdDigits 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   7
      Left            =   1410
      TabIndex        =   8
      Top             =   3195
      Width           =   540
   End
   Begin VB.CommandButton cmdDigits 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   6
      Left            =   2670
      TabIndex        =   7
      Top             =   2565
      Width           =   540
   End
   Begin VB.CommandButton cmdDigits 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   5
      Left            =   2040
      TabIndex        =   6
      Top             =   2565
      Width           =   540
   End
   Begin VB.CommandButton cmdDigits 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   4
      Left            =   1410
      TabIndex        =   5
      Top             =   2565
      Width           =   540
   End
   Begin VB.CommandButton cmdDigits 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   3
      Left            =   2670
      TabIndex        =   4
      Top             =   1935
      Width           =   540
   End
   Begin VB.CommandButton cmdDigits 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   2
      Left            =   2040
      TabIndex        =   3
      Top             =   1935
      Width           =   540
   End
   Begin VB.CommandButton cmdDigits 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   1
      Left            =   1410
      TabIndex        =   2
      Top             =   1935
      Width           =   540
   End
   Begin VB.CommandButton cmdDigits 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   3840
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Time"
      Height          =   240
      Left            =   1560
      TabIndex        =   39
      Top             =   1560
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Month"
      Height          =   195
      Left            =   3240
      TabIndex        =   37
      Top             =   1560
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      Height          =   195
      Left            =   4200
      TabIndex        =   36
      Top             =   1560
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Year"
      Height          =   195
      Left            =   1560
      TabIndex        =   35
      Top             =   1560
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblDisplay 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1080
      Width           =   4590
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim operand1 As Double, operand2 As Double
Dim operator As String
Dim cleardisplay As Boolean
Dim memory1 As Double
Dim percent As Double



Private Sub cmdClear_Click()
If Label2.Caption <> "Calculator" Then
MsgBox "Click On The Calculator Button First!!"
End If
lblDisplay.Caption = ""

End Sub

Private Sub cmdClearAll_Click()
If Label2.Caption <> "Calculator" Then
MsgBox "Click On The Calculator Button First!!"
End If
operand1 = 0
operand2 = 0
lblDisplay.Caption = ""
lblRunning.Caption = ""

End Sub

Private Sub cmdCos_Click()
lblDisplay.Caption = Cos(Val(lblDisplay.Caption))

End Sub

Private Sub cmdDigits_Click(Index As Integer)

If cleardisplay Then
    lblDisplay.Caption = ""
    cleardisplay = False
End If
If Len(lblDisplay.Caption) < 10 Then
   lblDisplay.Caption = lblDisplay.Caption + cmdDigits(Index).Caption
    Else
    End If


End Sub

Private Sub cmdDivide_Click()

operand1 = Val(lblDisplay.Caption)
operator = "/"
lblDisplay.Caption = ""

End Sub

Private Sub cmdEquals_Click()

On Error GoTo errorhandler

Dim result As Double

operand2 = Val(lblDisplay.Caption)
If operator = "+" Then result = operand1 + operand2
If operator = "-" Then result = operand1 - operand2
If operator = "*" Then result = operand1 * operand2
If operator = "/" And operand2 <> "0" Then _
                result = operand1 / operand2
lblDisplay.Caption = result
operand1 = result
lblRunning.Caption = result
Exit Sub

errorhandler:
MsgBox "The operation resulted in the following error" & _
    vbCrLf & Err.Description
lblDisplay.Caption = "ERROR"
cleardisplay = True

End Sub

Private Sub cmdMemory1_Click()
memory1 = lblDisplay.Caption
lblMemory1 = memory1
End Sub

Private Sub cmdMemory2_Click()
lblDisplay.Caption = memory1
End Sub

Private Sub cmdMemory3_Click()
memory1 = 0
lblMemory1.Caption = ""

End Sub

Private Sub cmdMinus_Click()

operand1 = Val(lblDisplay.Caption)
operator = "-"
lblDisplay.Caption = ""

End Sub

Private Sub cmdOver_Click()

If Val(lblDisplay.Caption) <> 0 Then lblDisplay.Caption = _
                        1 / Val(lblDisplay.Caption)
                        
End Sub


Private Sub cmdPlus_Click()

operand1 = Val(lblDisplay.Caption)
operator = "+"
lblDisplay.Caption = ""
lblRunning.Caption = operand1

End Sub

Private Sub cmdPlusMinus_Click()

lblDisplay.Caption = -Val(lblDisplay.Caption)

End Sub

Private Sub cmdSin_Click()
lblDisplay.Caption = Sin(Val(lblDisplay.Caption))
End Sub

Private Sub cmdSquareRoot_Click()
If lblDisplay.Caption < 0 Then
MsgBox "Can't calculate the square root of a negative number"
Else
lblDisplay.Caption = Sqr(Val(lblDisplay.Caption))
End If
End Sub

Private Sub cmdTimes_Click()

operand1 = Val(lblDisplay.Caption)
operator = "*"
lblDisplay.Caption = ""

End Sub

Private Sub Command1_Click()

If InStr(lblDisplay.Caption, ".") Then
    Exit Sub
Else
    lblDisplay.Caption = lblDisplay.Caption + "."
End If

End Sub

Private Sub Command2_Click()
Label1.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = True
Label2.Caption = "Clock"
Timer1.Enabled = True
End Sub

Private Sub Command3_Click()
Label2.Caption = "Calculator"
Timer1.Enabled = False
Label1.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
cmdClearAll_Click
End Sub

Private Sub Command4_Click()
Label2.Caption = "Date"
Timer1.Enabled = False
Label1.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = False
lblDisplay.Caption = "   " & Year(Now) & "            " & Month(Now) & "       " & Day(Now)
End Sub

Private Sub Form_Load()
Label2.Caption = "Calculator"
End Sub

Private Sub Timer1_Timer()
lblDisplay.Caption = Format(Now, "hh:mm:ss AM/PM")
End Sub
