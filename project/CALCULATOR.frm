VERSION 5.00
Begin VB.Form CALCULATOR 
   Caption         =   "Math Calculator"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3060
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   2910
   ScaleWidth      =   3060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Over 
      Caption         =   "1/X"
      Height          =   450
      Left            =   2475
      TabIndex        =   19
      Top             =   705
      Width           =   465
   End
   Begin VB.CommandButton PlusMinus 
      Caption         =   "+/-"
      Height          =   450
      Left            =   1875
      TabIndex        =   18
      Top             =   705
      Width           =   465
   End
   Begin VB.CommandButton Digits 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   1
      Left            =   90
      TabIndex        =   17
      Top             =   705
      Width           =   465
   End
   Begin VB.CommandButton Digits 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   2
      Left            =   645
      TabIndex        =   16
      Top             =   705
      Width           =   465
   End
   Begin VB.CommandButton Digits 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   3
      Left            =   1215
      TabIndex        =   15
      Top             =   705
      Width           =   465
   End
   Begin VB.CommandButton Equals 
      Caption         =   "="
      Height          =   450
      Left            =   1875
      TabIndex        =   14
      Top             =   2340
      Width           =   1095
   End
   Begin VB.CommandButton Div 
      Caption         =   "/"
      Height          =   450
      Left            =   2475
      TabIndex        =   13
      Top             =   1800
      Width           =   465
   End
   Begin VB.CommandButton Times 
      Caption         =   "*"
      Height          =   450
      Left            =   1875
      TabIndex        =   12
      Top             =   1800
      Width           =   465
   End
   Begin VB.CommandButton Minus 
      Caption         =   "-"
      Height          =   450
      Left            =   2475
      TabIndex        =   11
      Top             =   1245
      Width           =   465
   End
   Begin VB.CommandButton Plus 
      Caption         =   "+"
      Height          =   450
      Left            =   1875
      TabIndex        =   10
      Top             =   1245
      Width           =   465
   End
   Begin VB.CommandButton ClearBttn 
      Caption         =   "C"
      Height          =   450
      Left            =   90
      TabIndex        =   9
      Top             =   2340
      Width           =   465
   End
   Begin VB.CommandButton DotBttn 
      Caption         =   "."
      Height          =   450
      Left            =   1215
      TabIndex        =   8
      Top             =   2340
      Width           =   465
   End
   Begin VB.CommandButton Digits 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   9
      Left            =   1215
      TabIndex        =   7
      Top             =   1800
      Width           =   465
   End
   Begin VB.CommandButton Digits 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   8
      Left            =   645
      TabIndex        =   6
      Top             =   1800
      Width           =   465
   End
   Begin VB.CommandButton Digits 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   7
      Left            =   90
      TabIndex        =   5
      Top             =   1800
      Width           =   465
   End
   Begin VB.CommandButton Digits 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   6
      Left            =   1215
      TabIndex        =   4
      Top             =   1245
      Width           =   465
   End
   Begin VB.CommandButton Digits 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   5
      Left            =   645
      TabIndex        =   3
      Top             =   1245
      Width           =   465
   End
   Begin VB.CommandButton Digits 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   4
      Left            =   90
      TabIndex        =   2
      Top             =   1245
      Width           =   465
   End
   Begin VB.CommandButton Digits 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   645
      TabIndex        =   1
      Top             =   2340
      Width           =   465
   End
   Begin VB.Label Display 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   390
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   2820
   End
End
Attribute VB_Name = "CALCULATOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ******************************'
'  ******************************'
'  ** CALCULATOR FORM          **'
'  ******************************'
'  ******************************'
Option Explicit
Dim Operand1 As Double, Operand2 As Double
Dim Operator As String
Dim ClearDisplay As Boolean

Private Sub ClearBttn_Click()
    Display.Caption = ""
End Sub

Private Sub Digits_Click(Index As Integer)
    If ClearDisplay Then
        Display.Caption = ""
        ClearDisplay = False
    End If
    Display.Caption = Display.Caption + Digits(Index).Caption
End Sub

Private Sub Display_Click()

End Sub

Private Sub Div_Click()
    Operand1 = Val(Display.Caption)
    Operator = "/"
    Display.Caption = ""
End Sub

Private Sub DotBttn_Click()
    If ClearDisplay Then
        Display.Caption = ""
        ClearDisplay = False
    End If
    If InStr(Display.Caption, ".") Then
        Exit Sub
    Else
        Display.Caption = Display.Caption + "."
    End If
End Sub

Private Sub Equals_Click()
Dim result As Double

On Error GoTo ErrorHandler
    Operand2 = Val(Display.Caption)
    If Operator = "+" Then result = Operand1 + Operand2
    If Operator = "-" Then result = Operand1 - Operand2
    If Operator = "*" Then result = Operand1 * Operand2
    If Operator = "/" And Operand2 <> "0" Then result = Operand1 / Operand2
    Display.Caption = result
    ClearDisplay = True
    Exit Sub
ErrorHandler:
    MsgBox "The operation resulted in the following error" & vbCrLf & Err.Description
    Display.Caption = "ERROR"
    ClearDisplay = True
End Sub

Private Sub Minus_Click()
    Operand1 = Val(Display.Caption)
    Operator = "-"
    Display.Caption = ""
End Sub

Private Sub Over_Click()
    If Val(Display.Caption) <> 0 Then Display.Caption = 1 / Val(Display.Caption)
End Sub

Private Sub Plus_Click()
    Operand1 = Val(Display.Caption)
    Operator = "+"
    Display.Caption = ""
End Sub

Private Sub PlusMinus_Click()
    Display.Caption = -Val(Display.Caption)
End Sub

Private Sub Times_Click()
    Operand1 = Val(Display.Caption)
    Operator = "*"
    Display.Caption = ""
End Sub
