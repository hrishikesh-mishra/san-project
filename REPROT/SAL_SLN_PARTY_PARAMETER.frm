VERSION 5.00
Begin VB.Form SAL_SLN_PARTY_PARAMETER 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SALE SLN PARAMETER"
   ClientHeight    =   2475
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ComboBox saleSlnCombo 
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Sale Sln . :-"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Select The Sale Sln. "
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   990
      TabIndex        =   4
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "By Party"
      BeginProperty Font 
         Name            =   "Novelty Script"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   990
      TabIndex        =   3
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "SAL_SLN_PARTY_PARAMETER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************'
'**                          **'
'** PARTY SALE SLN PARAMETER **'
'**                          **'
'******************************'

'VARIABLE DECLARATION

Option Explicit

Dim salSlnSession As OraSession
Dim SalSlnDatabase As OraDatabase
Dim salSlnDyn      As OraDynaset

Private Sub CancelButton_Click()
  
  Unload Me

End Sub

Private Sub Form_Load()
  
  Set salSlnSession = CreateObject("oracleinprocserver.xorasession")
  Set SalSlnDatabase = salSlnSession.OpenDatabase("jms", "hrishi/jms", &H0&)
  Set salSlnDyn = SalSlnDatabase.CreateDynaset("select sale_sln from hrishi.sale_detail where type_of_cust='party'", &H4&)
  
  If salSlnDyn.EOF Then
     MsgBox "The database is empty . ", vbExclamation, "Empty:"
     Unload Me
  End If

  While Not salSlnDyn.EOF
     saleSlnCombo.AddItem salSlnDyn.Fields(0)
     saleSlnCombo.ListIndex = 0
     salSlnDyn.MoveNext
  Wend

End Sub

Private Sub OKButton_Click()
   
   If saleSlnCombo.Text <> "" Then
      ModuleVarious.salSlnParty = saleSlnCombo.Text
      SANPROJECT.formShow
      Unload Me
 End If

End Sub
