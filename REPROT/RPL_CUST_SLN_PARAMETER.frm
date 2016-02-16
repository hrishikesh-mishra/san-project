VERSION 5.00
Begin VB.Form RPL_CUST_SLN_PARAMETER 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "REPLACEMENT FROM CUSTOMER PARAMETER"
   ClientHeight    =   2745
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
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
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
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
      Left            =   2280
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ComboBox rplSlnCombo 
      Height          =   315
      Left            =   2040
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
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
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   593
      TabIndex        =   5
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Repaclement  from Customer"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   593
      TabIndex        =   4
      Top             =   360
      Width           =   3015
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
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
End
Attribute VB_Name = "RPL_CUST_SLN_PARAMETER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************'
'**                                **'
'** CUSTOMER REPLACE SLN PARAMETER **'
'**                                **'
'************************************'

'VARIABLE DECLARATION

Option Explicit

Dim rSCSession As OraSession
Dim rSCDatabase As OraDatabase
Dim rSCDyn     As OraDynaset

Private Sub CancelButton_Click()

  Unload Me

End Sub

Private Sub Form_Load()
  
  Set rSCSession = CreateObject("oracleinprocserver.xorasession")
  Set rSCDatabase = rSCSession.OpenDatabase("jms", "hrishi/jms", &H0&)
  Set rSCDyn = rSCDatabase.CreateDynaset("select rpl_sln from replacement_detail", &H0&)
  
  If rSCDyn.EOF Then
     MsgBox "Error ! Database empty .", vbExclamation, "Empty:"
     Unload Me
  End If
  
  While Not rSCDyn.EOF
    rplSlnCombo.AddItem rSCDyn.Fields(0)
    rSCDyn.MoveNext
  Wend

  rplSlnCombo.ListIndex = 0

End Sub

Private Sub OKButton_Click()

  If rplSlnCombo.Text <> "" Then
     ModuleVarious.repSlnCust = rplSlnCombo.Text
     SANPROJECT.formShow
     Unload Me
  End If
  
End Sub
