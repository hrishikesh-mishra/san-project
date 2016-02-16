VERSION 5.00
Begin VB.Form DATE_PARAMETER 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DATE PARAMETER"
   ClientHeight    =   3480
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4815
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton betweenRdBtn 
      Caption         =   "Between"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3240
      TabIndex        =   18
      Top             =   720
      Width           =   1215
   End
   Begin VB.OptionButton specificRdBtn 
      Caption         =   "Specific"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1920
      TabIndex        =   17
      Top             =   720
      Width           =   1215
   End
   Begin VB.OptionButton todayRdBtn 
      Caption         =   "Today"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   600
      TabIndex        =   16
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox EdayCombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   690
      TabIndex        =   11
      Text            =   " "
      Top             =   2520
      Width           =   855
   End
   Begin VB.ComboBox EyearCombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3150
      TabIndex        =   10
      Text            =   " "
      Top             =   2520
      Width           =   975
   End
   Begin VB.ComboBox EmonthCombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1575
      TabIndex        =   9
      Text            =   " "
      Top             =   2520
      Width           =   1575
   End
   Begin VB.ComboBox BdayCombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   671
      TabIndex        =   5
      Text            =   " "
      Top             =   1560
      Width           =   855
   End
   Begin VB.ComboBox ByearCombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3101
      TabIndex        =   4
      Text            =   " "
      Top             =   1560
      Width           =   975
   End
   Begin VB.ComboBox BmonthCombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1526
      TabIndex        =   3
      Text            =   " "
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "ESP"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "ESP"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000040C0&
      BorderStyle     =   3  'Dot
      X1              =   4560
      X2              =   2400
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000040C0&
      BorderStyle     =   3  'Dot
      X1              =   4560
      X2              =   4560
      Y1              =   1200
      Y2              =   600
   End
   Begin VB.Label Label6 
      Caption         =   "Date Option"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   1080
      TabIndex        =   19
      Top             =   480
      Width           =   1335
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000040C0&
      BorderStyle     =   3  'Dot
      X1              =   360
      X2              =   960
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000040C0&
      BorderStyle     =   3  'Dot
      X1              =   360
      X2              =   360
      Y1              =   600
      Y2              =   1200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000040C0&
      BorderStyle     =   3  'Dot
      X1              =   360
      X2              =   4560
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label5 
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Day"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3045
      TabIndex        =   8
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1755
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Day"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   675
      TabIndex        =   6
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Select The Date Between"
      BeginProperty Font 
         Name            =   "CloisterBlack BT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   0
      Width           =   3960
   End
End
Attribute VB_Name = "DATE_PARAMETER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************'
'**                     **'
'** DATA PARAMETER FORM **'
'**                     **'
'*************************'

'VARIABLE DECLARATION

Option Explicit

Dim chkDat1 As Boolean
Dim chkDat2 As Boolean
 

Private Sub betweenRdBtn_Click()
  
  EdayCombo.Enabled = True   'DOING ENDING DATE COMBOS ENABLE
  EmonthCombo.Enabled = True
  EyearCombo.Enabled = True

  BdayCombo.Enabled = True   'DOING BEGINING DATE COMBOS ENABLE
  BmonthCombo.Enabled = True
  ByearCombo.Enabled = True

  'CALLING SUBROUTINE FOR ADDING DATES TO THE COMBO BOX
  Call FILLCOMBODATE(EdayCombo, EmonthCombo, EyearCombo)
  Call FILLCOMBODATE(BdayCombo, BmonthCombo, ByearCombo)

  BdayCombo.Text = "5"        'SPECIFYING THE BEGINING DATES
  BmonthCombo.Text = "June"
  ByearCombo.Text = "2005"
  
End Sub

Private Sub CancelButton_Click()

  Unload Me

End Sub

Private Sub Form_Load()

  'CALLING SUBROUTINE FOR ADDING DATES TO THE COMBO BOX
  Call FILLCOMBODATE(BdayCombo, BmonthCombo, ByearCombo)
  
  BdayCombo.Text = "5"       'SPECIFYING THE BEGINING DATES
  BmonthCombo.Text = "June"
  ByearCombo.Text = "2005"
  
  'CALLING SUBROUTINE FOR ADDING DATES TO THE COMBO BOX
  Call FILLCOMBODATE(EdayCombo, EmonthCombo, EyearCombo)
  betweenRdBtn.Value = True

End Sub


Private Sub OKButton_Click()
  
  'CALLING FUNCTION FOR CHECKING DATES
  chkDat1 = VERIFY_DATE(BdayCombo.Text, BmonthCombo.Text, ByearCombo.Text)
  chkDat2 = VERIFY_DATE(EdayCombo.Text, EmonthCombo.Text, EyearCombo.Text)
  
  If chkDat1 = False Or chkDat2 = False Then  'IF DATE IS INVALID DATE
      MsgBox "Error ! Invalid date .", vbCritical, "Date Error:"
      Exit Sub
  End If
 
  If todayRdBtn.Value = True Or betweenRdBtn.Value = True Then
      ModuleVarious.sDate = BdayCombo.Text & "-" & BmonthCombo.Text & "-" & ByearCombo.Text
      ModuleVarious.eDate = EdayCombo.Text + "-" + EmonthCombo.Text + "-" + EyearCombo.Text
      SANPROJECT.formShow
      Unload Me
  Else
      ModuleVarious.sDate = BdayCombo.Text & "-" & BmonthCombo.Text & "-" & ByearCombo.Text
      ModuleVarious.eDate = BdayCombo.Text & "-" & BmonthCombo.Text & "-" & ByearCombo.Text
      SANPROJECT.formShow
      Unload Me
 End If

End Sub

Private Sub specificRdBtn_Click()
  
  EdayCombo.Enabled = False
  EmonthCombo.Enabled = False
  EyearCombo.Enabled = False

  BdayCombo.Enabled = True
  BmonthCombo.Enabled = True
  ByearCombo.Enabled = True
  Call FILLCOMBODATE(BdayCombo, BmonthCombo, ByearCombo)

End Sub

Private Sub todayRdBtn_Click()
   
   EdayCombo.Enabled = False
   EmonthCombo.Enabled = False
   EyearCombo.Enabled = False
 
   BdayCombo.Enabled = False
   BmonthCombo.Enabled = False
   ByearCombo.Enabled = False

   Call FILLCOMBODATE(EdayCombo, EmonthCombo, EyearCombo)
   Call FILLCOMBODATE(BdayCombo, BmonthCombo, ByearCombo)

End Sub
