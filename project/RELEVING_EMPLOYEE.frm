VERSION 5.00
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "ORADC.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form RELIEVING_EMPLOYEE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RELEVING EMPLOYEE FORM"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10860
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   10860
   Begin MSDBGrid.DBGrid RELGrid 
      Bindings        =   "RELEVING_EMPLOYEE.frx":0000
      Height          =   2055
      Left            =   120
      OleObjectBlob   =   "RELEVING_EMPLOYEE.frx":0017
      TabIndex        =   19
      ToolTipText     =   "Relieved Employee Detail."
      Top             =   5760
      Width           =   10575
   End
   Begin ORADCLibCtl.ORADC RELORADC 
      Height          =   375
      Left            =   3960
      Top             =   5040
      Width           =   4095
      _Version        =   65536
      _ExtentX        =   7223
      _ExtentY        =   661
      _StockProps     =   207
      Caption         =   "RELIEVED  EMPLOYEE"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DatabaseName    =   "JMS"
      Connect         =   "HRISHI/JMS"
      RecordSource    =   "SELECT * FROM HRISHI.RELIEVED_EMP_DETAIL"
   End
   Begin VB.CommandButton extCmd 
      BackColor       =   &H80000004&
      Caption         =   "&Exit"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Exit "
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton rlvCmd 
      BackColor       =   &H80000004&
      Caption         =   "&Relieve"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Relieving Employee "
      Top             =   4680
      Width           =   1455
   End
   Begin VB.ComboBox yearcombo 
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   4320
      TabIndex        =   12
      Text            =   " "
      ToolTipText     =   "Select the Year."
      Top             =   3840
      Width           =   975
   End
   Begin VB.ComboBox monthCombo 
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   2880
      TabIndex        =   11
      Text            =   " "
      ToolTipText     =   "Select the Month."
      Top             =   3840
      Width           =   1455
   End
   Begin VB.ComboBox dayCombo 
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   2160
      TabIndex        =   10
      Text            =   " "
      ToolTipText     =   "Select the Day."
      Top             =   3840
      Width           =   735
   End
   Begin VB.ComboBox eidCombo 
      Height          =   360
      Left            =   2880
      TabIndex        =   4
      Text            =   "Select Id"
      ToolTipText     =   "Select the employee ID."
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   10800
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   0
      Picture         =   "RELEVING_EMPLOYEE.frx":10A6
      ToolTipText     =   "San's Relieving Employee Form"
      Top             =   0
      Width           =   3000
   End
   Begin VB.Label yearLabel 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   9960
      TabIndex        =   22
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label monthlabel 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   8400
      TabIndex        =   21
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label dayLabel 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   7560
      TabIndex        =   20
      Top             =   3120
      Width           =   615
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   5400
      X2              =   5400
      Y1              =   1920
      Y2              =   4560
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   10920
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   10920
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   11040
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label refDate 
      Caption         =   "Relieve Date :-"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   " Year"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4320
      TabIndex        =   15
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   " Month"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   " Day"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2160
      TabIndex        =   13
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label doj 
      Caption         =   "Date of joining :-"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   5520
      TabIndex        =   9
      Top             =   3090
      Width           =   1935
   End
   Begin VB.Label destiLabel 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   2280
      TabIndex        =   8
      Top             =   3120
      Width           =   2535
   End
   Begin VB.Label destin 
      Caption         =   "Destination :-"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label empNameLabel 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   7920
      TabIndex        =   6
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label empName 
      Caption         =   " Employee Name :-"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   5520
      TabIndex        =   5
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label selLabel 
      Caption         =   "Please select Emp Id :-"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label rlvSlnLabel 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label relLabel2 
      Caption         =   "Releive No. :-"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "RELIEVING EMPLOYEE FORM"
      BeginProperty Font 
         Name            =   "Ivy League Outline"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   600
      Width           =   5415
   End
End
Attribute VB_Name = "RELIEVING_EMPLOYEE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 '**************************'
 '**                      **'
 '** RELIEVING EMP DETAIL **'
 '**                      **'
 '**************************'
 
 'VARIABLE DECLEARATION
 
 Option Explicit
 
 Dim relSession          As OraSession
 Dim relDatabase         As OraDatabase
 Dim relSlnDyn           As OraDynaset
 Dim eidDyn              As OraDynaset
 Dim relDyn              As OraDynaset
 Dim insertSql           As String
 Dim delSql              As String
 Dim dateStr             As String
 Dim dojStr              As String
 Dim chkDate             As Boolean
 
Private Sub eidCombo_Click()
 If eidCombo.Text = "" Then
    Exit Sub
 End If
 
 Set relDyn = relDatabase.CreateDynaset("select * from HRISHI.CURRENT_EMP_DETAIL where EID =" & Val(eidCombo.Text) & "", &H0&)
 
 empNameLabel.Caption = relDyn.Fields("ENAME")
 destiLabel.Caption = relDyn.Fields("DESTINATION")
 dayLabel.Caption = DAY(relDyn.Fields("JOIN_DATE"))
 monthlabel.Caption = MonthName(MONTH(relDyn.Fields("JOIN_DATE")))
 yearLabel.Caption = YEAR(relDyn.Fields("JOIN_DATE"))
 
End Sub

Private Sub extCmd_Click()

 If MsgBox("Do you want to exit ?", vbExclamation + vbYesNo, "Exit:") = vbYes Then
    Unload Me
 End If

End Sub

Private Sub Form_Load()

  Call FILLCOMBODATE(dayCombo, monthCombo, yearCombo)
  
  Set relSession = CreateObject("oracleinprocserver.xorasession")
  Set relDatabase = relSession.OpenDatabase("jms", "hrishi/jms", &H0&)
  Set relSlnDyn = relDatabase.CreateDynaset("select RELIEVE_SLN.NEXTVAL from DUAL", &H4&)
  Set eidDyn = relDatabase.CreateDynaset("select EID from HRISHI.CURRENT_EMP_DETAIL", &H0&)
  rlvSlnLabel.Caption = relSlnDyn.Fields(0)

  If eidDyn.EOF Then
     MsgBox "Nothing is Relieve to.", vbCritical, "Nothing :"
     rlvCmd.Enabled = False
  End If

  While Not eidDyn.EOF
     eidCombo.AddItem eidDyn.Fields(0)
     eidDyn.MoveNext
  Wend

End Sub

Private Sub rlvCmd_Click()

  If eidCombo.Text = "" Then
     Exit Sub
  End If

  chkDate = VERIFY_DATE(Val(dayCombo.Text), monthCombo.Text, Val(yearCombo.Text))
  
  If chkDate = False Then
     MsgBox "Date Error ! Invalid date.", vbCritical, "Date Error:"
     Exit Sub
  End If

  On Error GoTo ERRORHANDLER
  dateStr = dayCombo.Text + "-" + monthCombo.Text + "-" + yearCombo.Text
  dojStr = dayLabel.Caption + "-" + monthlabel.Caption + "-" + yearLabel.Caption
  insertSql = "insert into HRISHI.RELIEVED_EMP_DETAIL values (" & Val(rlvSlnLabel.Caption) & "," & Val(eidCombo.Text) & ",'" & _
             UCase(empNameLabel.Caption) & "','" & UCase(destiLabel.Caption) & "','" & dojStr & "','" & dateStr & "')"
  delSql = "delete from HRISHI.CURRENT_EMP_DETAIL where EID=" & Val(eidCombo.Text) & ""
    
  If MsgBox("Do you want to continue ", vbInformation + vbYesNo, "Conformation:") = vbYes Then
     relDatabase.ExecuteSQL (insertSql)
     relDatabase.ExecuteSQL (delSql)
     
     If MsgBox("Sucess ! data is savaed . " & vbCrLf & "Do you want continue .", vbInformation + vbYesNo, "Sucess:") = vbYes Then
        RELORADC.Refresh
        eidCombo.CLEAR
        eidCombo.Text = "Select id"
        empNameLabel.Caption = ""
        destiLabel.Caption = ""
        dayLabel.Caption = ""
        monthlabel.Caption = ""
        yearLabel.Caption = ""
        dayCombo.Text = DAY(Date)
        monthCombo.Text = MonthName(MONTH(Date))
        yearCombo.Text = YEAR(Date)
        Set relSlnDyn = relDatabase.CreateDynaset("select RELIEVE_SLN.NEXTVAL from DUAL", &H4&)
        rlvSlnLabel.Caption = relSlnDyn.Fields(0)
        Set eidDyn = relDatabase.CreateDynaset("select EID from HRISHI.CURRENT_EMP_DETAIL", &H0&)
        
        While Not eidDyn.EOF
           eidCombo.AddItem eidDyn.Fields(0)
           eidDyn.MoveNext
        Wend
        Exit Sub
     Else
        Unload Me
   End If
 End If

ERRORHANDLER:
 If relSession.LastServerErr = 0 Then
    If relDatabase.LastServerErr = 0 Then
       If Err.Number = 0 Then
       Else
         MsgBox "VB ERROR:" & vbCrLf & Err.Number & Err.Description, vbCritical, "VB Error:"
         Exit Sub
       End If
    Else
     MsgBox "DATABASE ERROR:" & vbCrLf & relDatabase.LastServerErr & relDatabase.LastServerErrText, vbCritical, "DATABASE Error:"
     relDatabase.LastServerErrReset
     Exit Sub
   End If
 Else
    MsgBox "SESSION ERROR :" & vbCrLf & relSession.LastServerErr & relSession.LastServerErrText, vbCritical, "SESSION Error:"
    relSession.LastServerErrReset
    Exit Sub
 End If

End Sub
