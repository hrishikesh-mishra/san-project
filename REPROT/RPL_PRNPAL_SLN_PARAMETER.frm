VERSION 5.00
Begin VB.Form RPL_PRNPAL_SLN_PARAMETER 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REPLACE TO PRINCIPAL"
   ClientHeight    =   2655
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox rplSlnCombo 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "ESP"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "ESP"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Repl  Sln . :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Replace To Principal"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Select The Replacement Sln. "
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "RPL_PRNPAL_SLN_PARAMETER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************'
'**                                    **'
'** REPLACE TO PRINCIPAL SLN PARAMETER **'
'**                                    **'
'****************************************'

'VARIABLE DECLARATION

Option Explicit

Dim rSPSession As OraSession
Dim rSPDatabase As OraDatabase
Dim rSPDyn     As OraDynaset

Private Sub CancelButton_Click()
  
  Unload Me

End Sub

Private Sub Form_Load()

  Set rSPSession = CreateObject("oracleinprocserver.xorasession")
  Set rSPDatabase = rSPSession.OpenDatabase("jms", "hrishi/jms", &H0&)
  Set rSPDyn = rSPDatabase.CreateDynaset("select rlp_sln from replace_to_principal_detail", &H0&)
  
  If rSPDyn.EOF Then
     MsgBox "Error ! Database empty .", vbExclamation, "Empty:"
     Unload Me
  End If

  While Not rSPDyn.EOF
     rplSlnCombo.AddItem rSPDyn.Fields(0)
     rSPDyn.MoveNext
  Wend

  rplSlnCombo.ListIndex = 0

End Sub

Private Sub OKButton_Click()

  If rplSlnCombo.Text <> "" Then
     ModuleVarious.repSlnPrnpal = rplSlnCombo.Text
     SANPROJECT.formShow
     Unload Me
  End If

End Sub

