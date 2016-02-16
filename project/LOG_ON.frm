VERSION 5.00
Begin VB.Form LOG_ON 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LOG ON"
   ClientHeight    =   3990
   ClientLeft      =   5445
   ClientTop       =   4050
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "LOG_ON.frx":0000
   ScaleHeight     =   3990
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cancelCmd 
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
      Left            =   2880
      TabIndex        =   5
      ToolTipText     =   "Exit from this."
      Top             =   3000
      Width           =   1215
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
      Left            =   480
      TabIndex        =   4
      ToolTipText     =   "Log On"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox passwordText 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   3
      ToolTipText     =   "Enter the Password."
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox userIDText 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      MaxLength       =   15
      TabIndex        =   2
      ToolTipText     =   "Enter the user."
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   0
      Picture         =   "LOG_ON.frx":6077A
      Top             =   0
      Width           =   3000
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Password :"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      Caption         =   "User ID :"
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
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "LOG_ON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************'
'**             **'
'** LOG ON FORM **'
'**             **'
'*****************'

'VARIABLE DECLERATION

Option Explicit

Dim lgNSession As OraSession
Dim lgNdatabase As OraDatabase
Dim lgNDyn      As OraDynaset
Dim countAtmp       As Integer
Dim dateStr         As String

Private Sub cancelCmd_Click()

  Unload SANPROJECT
  Unload Me

End Sub

Private Sub Form_Load()
   
  On Error GoTo ERRORHANDLER
  'CREATING SESSION, OPENING DATABAE AND CREATING DYNASET
  
  Set lgNSession = CreateObject("oracleinprocserver.xorasession")
  Set lgNdatabase = lgNSession.OpenDatabase("jms", "hrishi/jms", &H0&)
  Set SANPROJECT.hisSession = CreateObject("oracleinprocserver.xorasession")
  Set SANPROJECT.hisDatabase = SANPROJECT.hisSession.OpenDatabase("jms", "hrishi/jms", &H4&)
  countAtmp = 0
  
ERRORHANDLER:  'CODDING FOR ERROR DEDECTION
    If lgNSession.LastServerErr = 0 Then
      If lgNdatabase.LastServerErr = 0 Then
         If Err.Number = 0 Then
         Else
           MsgBox "VB ERROR:" & Err.Number & Err.Description, vbCritical, "VB ERROR:"
           Exit Sub
         End If
      Else
        MsgBox "DATABASE ERROR:" & lgNdatabase.LastServerErr & lgNdatabase.LastServerErrText _
                , vbCritical, "DATABASE ERROR:"
        lgNdatabase.LastServerErrReset
         Exit Sub
      End If
  Else
    MsgBox "SESSION ERROR:" & lgNSession.LastServerErr & lgNSession.LastServerErrText _
          , vbCritical, "SESSION ERROR:"
    lgNSession.LastServerErrReset
    Exit Sub
  End If
  
End Sub

Private Sub OkCmd_Click()
  
  If userIDText.Text = "" Then 'CHECKING EMPTY USER ID
     MsgBox "Empty ! UserId is Blank . ", vbCritical, "Empty:"
     userIDText.SetFocus
     Exit Sub
  ElseIf passwordText.Text = "" Then ' CHECKING FOR EMPTY PASSWORD
     MsgBox "Empty ! Password is Blank.", vbCritical, "Empty:"
     passwordText.SetFocus
     Exit Sub
  End If
 
  Set lgNDyn = lgNdatabase.CreateDynaset("select * from hrishi.sanproject_user where user_id ='" & userIDText.Text & "' and password = '" & passwordText.Text & "'", &H4&)
  countAtmp = countAtmp + 1
  If lgNDyn.EOF Then  'CHECKING FOR AUTHORIZED USER
     MsgBox "Invalid userID and password .", vbCritical, "Databased Error:"
     If countAtmp = 3 Then
        MsgBox "After 3 Unsuccess Atemption You Loged Out. ", vbExclamation, "Log Out:"
        Unload Me
        Unload SANPROJECT
        Exit Sub
     End If
    userIDText.Text = ""
    passwordText.Text = ""
    userIDText.SetFocus
    Exit Sub
 Else
   'ASSIGNING THEIR JOBS
   SANPROJECT.administration.Enabled = lgNDyn.Fields("ADM")
   SANPROJECT.admProEntry.Enabled = lgNDyn.Fields("ADMPE")
   SANPROJECT.admProUpd.Enabled = lgNDyn.Fields("ADMPU")
   SANPROJECT.adm_createUser.Enabled = lgNDyn.Fields("ADMCU")
   SANPROJECT.adm_modifyUser.Enabled = lgNDyn.Fields("ADMMU")
   SANPROJECT.adm_ChangePassword.Enabled = lgNDyn.Fields("ADMCP")
   SANPROJECT.adm_backUp.Enabled = lgNDyn.Fields("ADMBU")
   SANPROJECT.adm_recovery.Enabled = lgNDyn.Fields("ADMRV")
    
   SANPROJECT.purchase.Enabled = lgNDyn.Fields("PUR")
   SANPROJECT.pur_ven.Enabled = lgNDyn.Fields("PURVEN")
   SANPROJECT.pur_purEntry.Enabled = lgNDyn.Fields("PURPUR")
   SANPROJECT.pur_purReturn.Enabled = lgNDyn.Fields("PURPR")
  
   SANPROJECT.sale.Enabled = lgNDyn.Fields("SAL")
   SANPROJECT.sale_cust.Enabled = lgNDyn.Fields("SALCUST")
   SANPROJECT.sale_party.Enabled = lgNDyn.Fields("SALPRT")
   SANPROJECT.sale_saleEntry.Enabled = lgNDyn.Fields("SALSAL")
   SANPROJECT.sale_saleReturn.Enabled = lgNDyn.Fields("SALSR")
  
   SANPROJECT.repla.Enabled = lgNDyn.Fields("RPL")
   SANPROJECT.repl_on_spt.Enabled = lgNDyn.Fields("RPLOSR")
   SANPROJECT.repl_repl_from.Enabled = lgNDyn.Fields("RPLRF")
   SANPROJECT.repl_repl_to_prin.Enabled = lgNDyn.Fields("RPLRTP")
   SANPROJECT.repl_defc_pro.Enabled = lgNDyn.Fields("RPLDPE")
  
   SANPROJECT.empSuport.Enabled = lgNDyn.Fields("EMPS")
   SANPROJECT.emp_join_entry.Enabled = lgNDyn.Fields("EMPSJE")
   SANPROJECT.emp_relEmp.Enabled = lgNDyn.Fields("EMPSRE")
   SANPROJECT.emp_sal.Enabled = lgNDyn.Fields("EMPSES")
  
   SANPROJECT.masterDetail.Enabled = lgNDyn.Fields("MATD")
   SANPROJECT.mas_stk_detail.Enabled = lgNDyn.Fields("MATDSD")
   SANPROJECT.mas_pur_detail.Enabled = lgNDyn.Fields("MATDPD")
   SANPROJECT.mas_sale_detail.Enabled = lgNDyn.Fields("MATDSALD")
   SANPROJECT.mas_onSpotReplacementDetail.Enabled = lgNDyn.Fields("MATDOSRD")
   SANPROJECT.mas_replaceToPrnplDetail.Enabled = lgNDyn.Fields("MATDRTP")
   SANPROJECT.mas_replaceFromCust.Enabled = lgNDyn.Fields("MATDRFC")
   SANPROJECT.mas_currEmpDetail.Enabled = lgNDyn.Fields("MATDED")
   SANPROJECT.mas_relieveEmpDetail.Enabled = lgNDyn.Fields("MATDRED")
  
   SANPROJECT.report.Enabled = lgNDyn.Fields("RPT")
   SANPROJECT.rep_pro.Enabled = lgNDyn.Fields("RPTPR")
   SANPROJECT.rep_ven.Enabled = lgNDyn.Fields("RPTVR")
   SANPROJECT.rep_party.Enabled = lgNDyn.Fields("RPTPRTR")
   SANPROJECT.rep_cust.Enabled = lgNDyn.Fields("RPTCR")
   SANPROJECT.rep_pur.Enabled = lgNDyn.Fields("RPTPURR")
   SANPROJECT.rep_sale.Enabled = lgNDyn.Fields("RPTSALR")
   SANPROJECT.rep_stk.Enabled = lgNDyn.Fields("RPTSTKR")
   SANPROJECT.rep_rpl.Enabled = lgNDyn.Fields("RPTRPLR")
   SANPROJECT.rep_defPro.Enabled = lgNDyn.Fields("RPTDPR")
   SANPROJECT.rep_emp.Enabled = lgNDyn.Fields("RPTER")
   
   SANPROJECT.log_logOn.Enabled = False
   SANPROJECT.log_logOff.Enabled = True
   SANPROJECT.fav.Enabled = True
   SANPROJECT.window.Enabled = True
   SANPROJECT.help.Enabled = True
  
   dateStr = DAY(Date) & "-" & MonthName(MONTH(Date)) & "-" & YEAR(Date)
   ModuleVarious.LogOnUser = userIDText.Text
   ModuleVarious.workingDate = dateStr
   ModuleVarious.Stime = Time
   
   SANPROJECT.Show
   Unload Me
   
  End If
 
End Sub

Private Sub passwordText_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     okCmd.SetFocus
  End If

End Sub

Private Sub userIDText_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 Then
    passwordText.SetFocus
 End If

End Sub
