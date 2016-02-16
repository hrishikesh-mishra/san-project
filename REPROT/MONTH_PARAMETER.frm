VERSION 5.00
Begin VB.Form MONTH_PARAMETER 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MONTH PARAMETER"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4395
   Begin VB.ComboBox monthCombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "ESP"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton okCmd 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "ESP"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Month :"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Select the month  of payment"
      BeginProperty Font 
         Name            =   "Dobkin"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "MONTH_PARAMETER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************'
'**                      **'
'** MONTH NAME PARAMETER **'
'**                      **'
'**************************'

'VARIABLE DECLARATION

 Option Explicit
 
 Dim j As Integer

Private Sub Command2_Click()
 Unload Me
End Sub

Private Sub Form_Load()
  
  For j = 1 To 12
    monthCombo.AddItem MonthName(j)
  Next
  monthCombo.ListIndex = 0

End Sub

Private Sub OkCmd_Click()

  If monthCombo.Text <> "" Then
     ModuleVarious.monthN = monthCombo.Text
     Unload Me
     SANPROJECT.formShow
  End If

End Sub


