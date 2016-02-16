VERSION 5.00
Begin VB.Form PUR_SLN_PARAMETER 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PURCHASE SLN. PARAMETER"
   ClientHeight    =   1995
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox purSlnCombo 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   600
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
      Left            =   1920
      TabIndex        =   1
      Top             =   1440
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
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Pur Sln . :-"
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
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Select The Purchase Sln. "
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "PUR_SLN_PARAMETER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 '**********************************'
 '**                              **'
 '**  PURCHASE Sln No. PARAMETER  **'
 '**                              **'
 '**********************************'
 
 ' VARIABLE DECLEARATION


Option Explicit
Dim purSlnSession As OraSession
Dim purSlnDatabase As OraDatabase
Dim purSlnDyn  As OraDynaset

Private Sub CancelButton_Click()

   Unload Me

End Sub

Private Sub Form_Load()

   Set purSlnSession = CreateObject("oracleinprocserver.xorasession")
   Set purSlnDatabase = purSlnSession.OpenDatabase("jms", "hrishi/jms", &H0&)
   Set purSlnDyn = purSlnDatabase.CreateDynaset("select pru_sln from hrishi.purchase_detail", &H0&)
   
   If purSlnDyn.EOF Then
      MsgBox "Error ! The Purchase database is Empty .", vbCritical, "Empty Database :"
      Exit Sub
   End If

   While Not purSlnDyn.EOF
       purSlnCombo.AddItem purSlnDyn.Fields(0)
       purSlnCombo.ListIndex = 0
       purSlnDyn.MoveNext
   Wend

End Sub

Private Sub OKButton_Click()

   If purSlnCombo.Text <> "" Then
      ModuleVarious.purSln = purSlnCombo.Text
      SANPROJECT.formShow
      Unload Me
   End If

End Sub
