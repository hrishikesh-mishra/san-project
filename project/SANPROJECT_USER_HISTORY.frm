VERSION 5.00
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "ORADC.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form SANPROJECT_HISTORY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SANPROJECT USER HISTORY"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   8385
   Begin VB.CommandButton OkCmd 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "ESP"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton clearHisCmd 
      Caption         =   "Clear History"
      BeginProperty Font 
         Name            =   "ESP"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   4800
      Width           =   2415
   End
   Begin ORADCLibCtl.ORADC ORADC1 
      Height          =   375
      Left            =   2880
      Top             =   4200
      Width           =   2895
      _Version        =   65536
      _ExtentX        =   5106
      _ExtentY        =   661
      _StockProps     =   207
      Caption         =   "HISTORY"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DatabaseName    =   "JMS"
      Connect         =   "HRISHI/JMS"
      RecordSource    =   "SELECT * FROM SANPROJECT_HISTORY"
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "SANPROJECT_USER_HISTORY.frx":0000
      Height          =   3255
      Left            =   0
      OleObjectBlob   =   "SANPROJECT_USER_HISTORY.frx":0015
      TabIndex        =   0
      Top             =   960
      Width           =   8295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "SANPROJECT USER HISTORY"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "SANPROJECT_HISTORY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************'
'**                                    **'
'** SANPROJECT LOGIN USER HISTORY FORM **'
'**                                    **'
'****************************************'

'VARIABLE DECLARATION

Option Explicit

Dim hisSession As OraSession
Dim hisDatabase As OraDatabase
Dim clearSql   As String

Private Sub clearHisCmd_Click()
   
   clearSql = " delete from hrishi.sanproject_history "
 
   If MsgBox("Do you really clear the history.", vbExclamation + vbYesNo, "Conformation:") = vbYes Then
       hisDatabase.ExecuteSQL (clearSql)
       ORADC1.Refresh
   End If
  
End Sub

Private Sub Form_Load()
    
    Set hisSession = CreateObject("oracleinprocserver.xorasession")
    Set hisDatabase = hisSession.OpenDatabase("jms", "hrishi/jms", &H0&)

End Sub

Private Sub OkCmd_Click()
    
    Unload Me

End Sub
