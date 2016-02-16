VERSION 5.00
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "ORADC.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form CURRENT_EMP_UPDATE 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CURRENT EMPLOYEE FORM [UPDATE]"
   ClientHeight    =   8205
   ClientLeft      =   2100
   ClientTop       =   1290
   ClientWidth     =   9570
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   9570
   Begin MSDBGrid.DBGrid empDetailGrid 
      Bindings        =   "CURRENT_EMP_UPDATE.frx":0000
      Height          =   1815
      Left            =   0
      OleObjectBlob   =   "CURRENT_EMP_UPDATE.frx":001B
      TabIndex        =   26
      Top             =   6240
      Width           =   9495
   End
   Begin ORADCLibCtl.ORADC empDetOracle 
      Height          =   375
      Left            =   2520
      Top             =   5760
      Width           =   4575
      _Version        =   65536
      _ExtentX        =   8070
      _ExtentY        =   661
      _StockProps     =   207
      Caption         =   "EMP DETAIL"
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
      RecordSource    =   "SELECT * FROM HRISHI.CURRENT_EMP_DETAIL "
   End
   Begin VB.CommandButton extCmd 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   5678
      TabIndex        =   13
      ToolTipText     =   "Exit from this "
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton delCmd 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   3998
      TabIndex        =   12
      ToolTipText     =   "Delete the information"
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton updCmd 
      Caption         =   "&Update"
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      ToolTipText     =   "Update the information."
      Top             =   5160
      Width           =   1335
   End
   Begin VB.ComboBox eidCombo 
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   1440
      TabIndex        =   0
      ToolTipText     =   "Select the Employee ID."
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox empNameText 
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   6360
      MaxLength       =   25
      TabIndex        =   1
      ToolTipText     =   "Enter the name of Employee"
      Top             =   1560
      Width           =   3015
   End
   Begin VB.OptionButton msexRBtn 
      Caption         =   "Male"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      ToolTipText     =   "Select the sex."
      Top             =   2280
      Width           =   975
   End
   Begin VB.OptionButton fSexRBtn 
      Caption         =   "Female"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      ToolTipText     =   "Select the sex."
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox empAddText 
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   1440
      MaxLength       =   50
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      ToolTipText     =   "Enter the Address."
      Top             =   2880
      Width           =   2895
   End
   Begin VB.TextBox empPhnText 
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   6360
      MaxLength       =   15
      TabIndex        =   6
      ToolTipText     =   "Enter the Phone No."
      Top             =   3240
      Width           =   2535
   End
   Begin VB.ComboBox dayCombo 
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   6360
      TabIndex        =   8
      Text            =   " "
      ToolTipText     =   "Select the Day."
      Top             =   4200
      Width           =   735
   End
   Begin VB.ComboBox monthCombo 
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   7080
      TabIndex        =   9
      Text            =   " "
      ToolTipText     =   "Select the Month."
      Top             =   4200
      Width           =   1455
   End
   Begin VB.ComboBox yearcombo 
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   8520
      TabIndex        =   10
      Text            =   " "
      ToolTipText     =   "Select the year."
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox empAgeText 
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   6360
      MaxLength       =   3
      TabIndex        =   4
      ToolTipText     =   "Enter the age of employee"
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox destiText 
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1560
      MaxLength       =   15
      TabIndex        =   7
      ToolTipText     =   "Enter the Destination"
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   0
      Picture         =   "CURRENT_EMP_UPDATE.frx":13F8
      ToolTipText     =   "San's Current Employee Update Form."
      Top             =   0
      Width           =   3000
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   2
      Index           =   3
      X1              =   3960
      X2              =   3960
      Y1              =   5040
      Y2              =   5640
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   2
      Index           =   2
      X1              =   5520
      X2              =   5520
      Y1              =   5040
      Y2              =   5640
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   2
      Index           =   1
      X1              =   7320
      X2              =   9600
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   2400
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line9 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      X1              =   0
      X2              =   9480
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   2
      X1              =   7320
      X2              =   7320
      Y1              =   5640
      Y2              =   6120
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   2
      X1              =   2400
      X2              =   2400
      Y1              =   5640
      Y2              =   6120
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   2
      Index           =   0
      X1              =   2400
      X2              =   2400
      Y1              =   5040
      Y2              =   5640
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   2
      X1              =   7320
      X2              =   7320
      Y1              =   5040
      Y2              =   5640
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   2
      X1              =   2400
      X2              =   7320
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   2
      X1              =   2400
      X2              =   7320
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      X1              =   0
      X2              =   9480
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      X1              =   4560
      X2              =   4560
      Y1              =   1320
      Y2              =   4920
   End
   Begin VB.Label Label5 
      Caption         =   "CURRENT EMPLOYEE UPDATE"
      BeginProperty Font 
         Name            =   "TACOMA"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   1680
      TabIndex        =   25
      Top             =   720
      Width           =   7215
   End
   Begin VB.Label empNameLabel 
      Caption         =   "Name :-"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4800
      TabIndex        =   24
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label ageLabel 
      Caption         =   "Age :-"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4800
      TabIndex        =   23
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label empPhnLabel 
      Caption         =   "Phone No. :-"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4800
      TabIndex        =   22
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   " Day"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6360
      TabIndex        =   21
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   " Month"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   7200
      TabIndex        =   20
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   " Year"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   8520
      TabIndex        =   19
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Joining Date :-"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4800
      TabIndex        =   18
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label empAddLabel 
      Caption         =   "Address :-"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label empSexLabel 
      Caption         =   "Sex :-"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label empIdLabel1 
      Caption         =   "Emp ID :-"
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label destiLabel 
      Caption         =   "Destination :-"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Width           =   1455
   End
End
Attribute VB_Name = "CURRENT_EMP_UPDATE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************'
'**                         **'
'** CURRENT EMPLOYEE UPDATE **'
'**                         **'
'*****************************'

'VARIABLE DECLARATION
 
 Option Explicit
  
 Dim ceuSession         As OraSession
 Dim ceuDatabase        As OraDatabase
 Dim ceuDyn             As OraDynaset
 Dim eidDyn             As OraDynaset
 Dim sex                As String
 Dim dateStr            As String
 Dim updateSql          As String
 Dim delSql             As String
 Dim chkDate            As Boolean
 Dim FlagEmpty          As Boolean
Private Sub chkEmpty()
    
    'A SUBROUTINE FOR CHECKING THE EMPTY
    If empNameText.Text = "" Then
       MsgBox "Empty Error: Name isn't present.", vbCritical, "Empty:"
       FlagEmpty = True
    ElseIf empAgeText.Text = "" Then
       MsgBox "Empty Error : Age isn't present.", vbCritical, "Empty:"
       FlagEmpty = True
    ElseIf IsNumeric(empAgeText.Text) = False Then
       MsgBox "Missmatch datatype of Age !", vbCritical, "Error:"
       FlagEmpty = True
       empAgeText.SetFocus
    ElseIf empAddText.Text = "" Then
       MsgBox "Empty Error: Address isn't present! ", vbCritical, "Empty:"
       FlagEmpty = True
    ElseIf empPhnText.Text = "" Then
       MsgBox "Empty Error : Phone number isn't present !", vbCritical, "Empty:"
       FlagEmpty = True
    ElseIf destiText.Text = "" Then
       MsgBox "Empty Error : Destination isn't present ! ", vbCritical, "Empty:"
       FlagEmpty = True
    Else
      FlagEmpty = False
    End If
  
End Sub

Private Sub dayCombo_KeyPress(KeyAscii As Integer)
  
    If KeyAscii = 13 Then 'IF ENTER KEY PRESSED
      monthCombo.SetFocus
    End If
    
End Sub

Private Sub delCmd_Click()
    
    'THE SQL COMMAND FOR DELETING THE EMPLOYEE
    delSql = "delete from HRISHI.CURRENT_EMP_DETAIL where EID=" & Val(eidCombo.Text) & ""
    
    'ASKING FOR CONFIRMATION
    If MsgBox("The data will be parmanentally removed ." & vbCrLf & " Continue ?", vbYesNo, "Warning:") = vbYes Then
      ceuDatabase.ExecuteSQL (delSql)
      MsgBox "Sucess ! The deletion is made .", vbInformation, "Sucess:"
      empDetOracle.Refresh
      eidCombo.SetFocus
      Exit Sub
    End If

End Sub

Private Sub destiText_KeyPress(KeyAscii As Integer)
 
   If KeyAscii = 13 Then 'IF ENTER KEY PRESSED
      dayCombo.SetFocus
   End If
   
End Sub

Private Sub eidCombo_Click()
  
  If eidCombo.Text = "" Then
     Exit Sub
  End If
  
  Set ceuDyn = ceuDatabase.CreateDynaset("select * from HRISHI.CURRENT_EMP_DETAIL Where EID = " _
               & Val(eidCombo.Text) & "", &H0&)
  
  empNameText.Text = ceuDyn.Fields("ENAME")
  sex = ceuDyn.Fields("SEX")
  
  If sex = "M" Then
    msexRBtn.Value = True
  Else
    fSexRBtn.Value = True
  End If

  empAgeText.Text = ceuDyn.Fields("AGE")
  empAddText.Text = ceuDyn.Fields("ADDRESS")
  empPhnText.Text = ceuDyn.Fields("PHNO")
  destiText.Text = ceuDyn.Fields("DESTINATION")
  dayCombo.Text = DAY(ceuDyn.Fields("JOIN_DATE"))
  monthCombo.Text = MonthName(MONTH(ceuDyn.Fields("JOIN_DATE")))
  yearCombo.Text = YEAR(ceuDyn.Fields("JOIN_DATE"))

End Sub

Private Sub eidCombo_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then 'IF KEY PRESS IS ENTER
     empNameText.SetFocus
  End If
  
End Sub

Private Sub empAddText_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then 'IF ENTER KEY IS PRESSED
      empPhnText.SetFocus
   End If
   
End Sub

Private Sub empAgeText_KeyPress(KeyAscii As Integer)
   
    If KeyAscii = 13 Then 'IF KEY PRESS IS ENTER
       empAddText.SetFocus
    End If
    
End Sub

Private Sub empNameText_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     msexRBtn.SetFocus
  End If
  
End Sub

Private Sub empPhnText_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then 'IF ENTER KEY IS PRESSED
      destiText.SetFocus
   End If
   
End Sub

Private Sub extCmd_Click()

  If MsgBox("Do you want to exit ?", vbExclamation + vbYesNo, "Exit:") = vbYes Then
     Unload Me
  End If

End Sub

Private Sub Form_Load()

  'CALL A SUBROUTINE FOR ADDING THE DATE IN THEIR COMBO'S
  Call FILLCOMBODATE(dayCombo, monthCombo, yearCombo)
  
  On Error GoTo ERRORHANDLER
  'CREATING SESSION, OPENING DATABASE AND CREATING DATABASE
  Set ceuSession = CreateObject("oracleinprocserver.xorasession")
  Set ceuDatabase = ceuSession.OpenDatabase("jms", "hrishi/jms", &H4&)
  Set eidDyn = ceuDatabase.CreateDynaset("select EID from HRISHI.CURRENT_EMP_DETAIL", &H4&)
  
  While Not eidDyn.EOF      'ADDING EMPLOYEE ID IN ID COMBO
      eidCombo.AddItem eidDyn.Fields(0)
      eidDyn.MoveNext
      eidCombo.ListIndex = 0
  Wend

ERRORHANDLER:
  If ceuSession.LastServerErr = 0 Then
     If ceuDatabase.LastServerErr = 0 Then
        If Err.Number = 0 Then
        Else
           MsgBox "VB ERROR:" & Err.Number & Err.Description, vbCritical, "VB Error:"
           Unload Me
        End If
     Else
        MsgBox "DATABASE ERROR :" & ceuDatabase.LastServerErr & ceuSession.LastServerErrText _
                , vbCritical, "DATABASE ERROR:"
        ceuDatabase.LastServerErrReset
        Unload Me
     End If
  Else
    MsgBox "SESSION ERROR:" & ceuSession.LastServerErr & ceuSession.LastServerErrText _
            , vbCritical, "SESSION ERROR:"
    ceuSession.LastServerErrReset
    Unload Me
  End If
  
End Sub

Private Sub monthCombo_KeyPress(KeyAscii As Integer)
  
   If KeyAscii = 13 Then 'IF ENTER KEY IS PRESSED
      yearCombo.SetFocus
   End If
   
 
End Sub

Private Sub msexRBtn_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then 'IF KEY PRESS IS ENTER
     empAgeText.SetFocus
  End If
  
End Sub

Private Sub updCmd_Click()
 
 'CALLING A FUNCTION FOR VERIFICATION OF DATE
 chkDate = VERIFY_DATE(Val(dayCombo.Text), monthCombo.Text, Val(yearCombo.Text))
 'CALLING A SUBROUTINE FOR CHECKING EMPTY
 Call chkEmpty
 
 On Error GoTo eh
 
 If chkDate = False Then         'IF ENTER DATE IS INVALID DATE THEN
   MsgBox "Date Error : Invalid date .", vbCritical, "DATE Error:"
   Exit Sub
 End If
 
 If FlagEmpty = True Then        'IF ANY EMPTY
   Exit Sub
 End If

 dateStr = dayCombo.Text + "-" + monthCombo.Text + "-" + yearCombo.Text
 If msexRBtn.Value = True Then  'FINDIG THE SEX OF EMPLOYEE
    sex = "M"
 Else
    sex = "F"
 End If

 'THE SQL COMMAND FOR UPDATING THE DATABASE OF EMPLOYEE
 updateSql = "update HRISHI.CURRENT_EMP_DETAIL set ENAME ='" & UCase(empNameText.Text) & "',SEX = '" & _
              sex & "', AGE =" & Val(empAgeText.Text) & ",ADDRESS = '" & UCase(empAddText.Text) & "',PHNO ='" & _
                empPhnText.Text & "', DESTINATION = '" & UCase(destiText.Text) & "',JOIN_DATE ='" & dateStr & "' where EID= " & Val(eidCombo.Text) & ""
 
 'ASKING THE CONFIRMATION FOR UPDATION
 If MsgBox("Are you ready to update ?", vbInformation + vbYesNo, "CONFORMATION :") = vbYes Then
    ceuDatabase.ExecuteSQL (updateSql)
    'IF NO ERROR IS PRODUCED THEN
    MsgBox "Sucess ! Database is updated .", vbInformation, "Sucess:"
    empDetOracle.Refresh
    eidCombo.SetFocus
    Exit Sub
 End If

eh: 'CODING FOR ERROR HANDLER
 If ceuSession.LastServerErr = 0 Then
    If ceuDatabase.LastServerErr = 0 Then
       If Err.Number = 0 Then
       Else
         MsgBox "VB ERROR :" & vbCrLf & Err.Number & Err.Description, vbCritical, "VB Error:"
         Exit Sub
       End If
    Else
      MsgBox "DATABASE ERROR :" & vbCrLf & ceuDatabase.LastServerErr & ceuDatabase.LastServerErrText, vbCritical, "DATABASE Error:"
      ceuDatabase.LastServerErrReset
      Exit Sub
   End If
 Else
   MsgBox "SESSION ERROR :" & vbCrLf & ceuSession.LastServerErr & ceuSession.LastServerErrText, vbCritical, "SESSION Error:"
   ceuSession.LastServerErrReset
   Exit Sub
 End If

End Sub

Private Sub yearcombo_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then 'IF ENTER KEY IS PRESSED
      updCmd.SetFocus
  End If
  
End Sub
