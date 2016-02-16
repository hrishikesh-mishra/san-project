VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form EMP_JOINING_ENTRY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EMPLOYEE JOINING FORM"
   ClientHeight    =   6630
   ClientLeft      =   1920
   ClientTop       =   2055
   ClientWidth     =   9450
   FillColor       =   &H8000000F&
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
   ScaleHeight     =   6630
   ScaleWidth      =   9450
   Begin VB.TextBox destiText 
      ForeColor       =   &H00404080&
      Height          =   375
      Left            =   1560
      MaxLength       =   15
      TabIndex        =   7
      ToolTipText     =   "Enter the destination of employee"
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton saveCmd 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2048
      TabIndex        =   12
      ToolTipText     =   "Save the information. "
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cancelCmd 
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   3728
      TabIndex        =   13
      ToolTipText     =   "Cancel the process."
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton extCmd 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      ToolTipText     =   "Exit from this."
      Top             =   5520
      Width           =   1455
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   25
      Top             =   6255
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1940
            MinWidth        =   1940
            TextSave        =   "9:07 AM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2116
            MinWidth        =   2116
            TextSave        =   "7/23/2005"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9949
            MinWidth        =   9949
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1481
            MinWidth        =   1481
            TextSave        =   "NUM"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox empAgeText 
      ForeColor       =   &H00404080&
      Height          =   360
      Left            =   6240
      MaxLength       =   3
      TabIndex        =   4
      ToolTipText     =   "Enter the age of employee"
      Top             =   2280
      Width           =   615
   End
   Begin VB.ComboBox yearcombo 
      ForeColor       =   &H00404080&
      Height          =   360
      Left            =   8400
      TabIndex        =   11
      Text            =   " "
      ToolTipText     =   "Select the year"
      Top             =   4080
      Width           =   975
   End
   Begin VB.ComboBox monthCombo 
      ForeColor       =   &H00404080&
      Height          =   360
      Left            =   6960
      TabIndex        =   10
      Text            =   " "
      ToolTipText     =   "Select the month"
      Top             =   4080
      Width           =   1455
   End
   Begin VB.ComboBox dayCombo 
      ForeColor       =   &H00404080&
      Height          =   360
      Left            =   6240
      TabIndex        =   8
      Text            =   " "
      ToolTipText     =   "Select the day "
      Top             =   4080
      Width           =   735
   End
   Begin VB.TextBox empPhnText 
      ForeColor       =   &H00404080&
      Height          =   360
      Left            =   6240
      MaxLength       =   15
      TabIndex        =   6
      ToolTipText     =   "Enter the Phone no of employee"
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox empAddText 
      ForeColor       =   &H00404080&
      Height          =   975
      Left            =   1440
      MaxLength       =   50
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      ToolTipText     =   "Enter the address"
      Top             =   2760
      Width           =   2895
   End
   Begin VB.OptionButton fSexRBtn 
      Caption         =   "Female"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      ToolTipText     =   "Select the sex of employee"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.OptionButton msexRBtn 
      Caption         =   "Male"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      ToolTipText     =   "Select the sex of employee"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox empNameText 
      ForeColor       =   &H00404080&
      Height          =   360
      Left            =   6240
      MaxLength       =   25
      TabIndex        =   1
      ToolTipText     =   "Enter the name of employee."
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   2
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   600
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   1800
      X2              =   7320
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   0
      X2              =   9480
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   0
      Picture         =   "EMPLOYEE.frx":0000
      ToolTipText     =   "San's Employee Joining Entry Form"
      Top             =   0
      Width           =   3000
   End
   Begin VB.Label destiLabel 
      Caption         =   "Destination :-"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   1815
      X2              =   7335
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   1815
      X2              =   7335
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   7335
      X2              =   7335
      Y1              =   5400
      Y2              =   6000
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   1815
      X2              =   1815
      Y1              =   5400
      Y2              =   6000
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   5415
      X2              =   5415
      Y1              =   5400
      Y2              =   6000
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   3495
      X2              =   3495
      Y1              =   5400
      Y2              =   6000
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   9480
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   4560
      X2              =   4560
      Y1              =   1200
      Y2              =   4680
   End
   Begin VB.Label empIdLabel1 
      Caption         =   "Emp ID :-"
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label empSexLabel 
      Caption         =   "Sex :-"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label empAddLabel 
      Caption         =   "Address :-"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Joining Date :-"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4680
      TabIndex        =   21
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   " Year"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   8400
      TabIndex        =   20
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   " Month"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   7080
      TabIndex        =   19
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   " Day"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   6240
      TabIndex        =   18
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label empPhnLabel 
      Caption         =   "Phone No. :-"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4680
      TabIndex        =   17
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label ageLabel 
      Caption         =   "Age :-"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label empNameLabel 
      Caption         =   "Name :-"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4680
      TabIndex        =   15
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label empIdLabel 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label EMP 
      Alignment       =   2  'Center
      Caption         =   "EMPLOYEE JOINING  FORM"
      BeginProperty Font 
         Name            =   "HILLARY"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   720
      Width           =   6135
   End
End
Attribute VB_Name = "EMP_JOINING_ENTRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************'
'**                       **'
'** EMPLOYEE JOINING FORM **'
'**                       **'
'***************************'

'VARIABLE DECLARATION
Option Explicit

Dim empJSession        As OraSession
Dim empJDatabase       As OraDatabase
Dim empIDDyn           As OraDynaset
Dim sex                As String
Dim dateStr            As String
Dim insertSql          As String
Dim chkDate            As Boolean
Dim FlagEmpty          As Boolean

Private Sub chkEmpty()
  
  'SUBROUNTINE FOR CHECKING EMPTY
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
  
  Private Sub CLEAR()
   
   empNameText.Text = ""  'CLEARING ALL TEXT BOX AND COMBO BOX
   msexRBtn.Value = True
   empAgeText.Text = ""
   empAddText.Text = ""
   empPhnText.Text = ""
   destiText.Text = ""
   dayCombo.Text = DAY(Date)
   monthCombo.Text = MonthName(MONTH(Date))
   yearCombo.Text = YEAR(Date)
   empNameText.SetFocus
   
     
  End Sub


Private Sub cancelCmd_Click()
 
  Call CLEAR

End Sub

Private Sub cancelCmd_GotFocus()

  StatusBar1.Panels(3) = "CLEAR THE INFORMATION ..."

End Sub

Private Sub dayCombo_GotFocus()

  StatusBar1.Panels(3) = "SELECT THE JOING DATE..."

End Sub

Private Sub dayCombo_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then  'THE ENTER KEY PRESSED
    monthCombo.SetFocus
 End If

End Sub

Private Sub destiText_GotFocus()

 StatusBar1.Panels(3) = "ENTER THE DESTINATION .."

End Sub

Private Sub destiText_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 Then  'THE ENTER KEY PRESSED
    dayCombo.SetFocus
 End If
 
End Sub

Private Sub empAddText_GotFocus()
 
 StatusBar1.Panels(3) = "ENTER THE ADDRESS .."
 
End Sub

Private Sub empAddText_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then  'THE ENTER KEY PRESSED
   empPhnText.SetFocus
  End If

End Sub

Private Sub empAgeText_GotFocus()

  StatusBar1.Panels(3) = "ENTER EMPLOYEE AGE .."

End Sub

Private Sub empAgeText_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then  'THE ENTER KEY PRESSED
     empAddText.SetFocus
  End If

End Sub

Private Sub empNameText_GotFocus()
 
   StatusBar1.Panels(3) = "ENTER EMPLOYEE NAME .."
 
End Sub

Private Sub empNameText_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 Then  'THE ENTER KEY PRESSED
    msexRBtn.SetFocus
 End If
 
End Sub

Private Sub empPhnText_GotFocus()

  StatusBar1.Panels(3) = "ENTER THE PHONE OF EMPLOYEE.."

End Sub

Private Sub empPhnText_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then 'THE ENTER KEY PRESSED
     destiText.SetFocus
  End If

End Sub

Private Sub extCmd_Click()
 
 If MsgBox("Do you want to exit ?", vbInformation + vbYesNo, "Exit:") = vbYes Then
   Unload Me
 End If
 
End Sub

Private Sub extCmd_GotFocus()
  
  StatusBar1.Panels(3) = "EXIT"

End Sub

Private Sub Form_GotFocus()

   StatusBar1.Panels(3) = "EMPLOYEE JOINING FORM.."

End Sub

Private Sub Form_Load()
   
   'CALLING THE SUBROUTINE FOR ADDING DATES IN THEIR COMBOS
   Call FILLCOMBODATE(dayCombo, monthCombo, yearCombo)

  On Error GoTo ERRORHANDLER
  'CREATING SESSION, OPEN DATABASE AND CREATING DYNASET
  Set empJSession = CreateObject("oracleinprocserver.xorasession")
  Set empJDatabase = empJSession.OpenDatabase("jms", "hrishi/jms", &H0&)
  Set empIDDyn = empJDatabase.CreateDynaset("select EMP_ID.NEXTVAL from dual", &H4&)

  empIdLabel.Caption = empIDDyn.Fields(0)
  msexRBtn.Value = True

ERRORHANDLER: 'CODDING FOR ERROR DETECTION
   If empJSession.LastServerErr = 0 Then
      If empJDatabase.LastServerErr = 0 Then
          If Err.Number = 0 Then
          Else
            MsgBox "VB ERROR:" & vbCrLf & Err.Number & Err.Description, vbCritical, "VB Error:"
            Unload Me
          End If
     Else
          MsgBox "DATABASE ERROR :" & vbCrLf & empJDatabase.LastServerErr & empJDatabase.LastServerErrText, vbCritical, "DATABASE Error:"
          empJDatabase.LastServerErrReset
          Unload Me
     End If
  Else
     MsgBox "SESSION ERROR :" & vbCrLf & empJSession.LastServerErr & empJSession.LastServerErrText, vbCritical, "SESSION Error:"
     empJSession.LastServerErrReset
     Exit Sub
  End If

End Sub

Private Sub fSexRBtn_GotFocus()

StatusBar1.Panels(3) = "SELECT THE SEX OF EMPLOYEE.."

End Sub

Private Sub fSexRBtn_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then  'THE ENTER KEY PRESSED
      empAgeText.SetFocus
   End If

End Sub

Private Sub monthCombo_GotFocus()

  StatusBar1.Panels(3) = "SELECT THE JOINING DATE.."

End Sub

Private Sub monthCombo_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then 'THE ENTER KEY PRESSED
      yearCombo.SetFocus
   End If

End Sub

Private Sub msexRBtn_GotFocus()

  StatusBar1.Panels(3) = "SELECT THE SEX OF EMPLOYEE.."

End Sub

Private Sub msexRBtn_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then  'THE ENTER KEY PRESSED
      empAgeText.SetFocus
  End If

End Sub

Private Sub saveCmd_Click()
 
 'CALLING FUNCTION FOR CHECK DATE
 chkDate = VERIFY_DATE(Val(dayCombo.Text), monthCombo.Text, Val(yearCombo.Text))
 Call chkEmpty
 
 On Error GoTo ERRORHANDLER
 
  If chkDate = False Then 'IF INVALID DATE THEN EXIT
     MsgBox "Date Error : Invalid date .", vbCritical, "DATE Error:"
     Exit Sub
  End If
 
   If FlagEmpty = True Then  'IF EMPTY ANY ONE THEN EXIT
      Exit Sub
   End If

   dateStr = dayCombo.Text + "-" + monthCombo.Text + "-" + yearCombo.Text
   
   If msexRBtn.Value = True Then  'FINDING SEX OF EMPLOYEE
      sex = "M"
   Else
      sex = "F"
   End If
 
  'THE SQL FOR INSERTION OF RECORD
  insertSql = "insert into HRISHI.CURRENT_EMP_DETAIL values( " & Val(empIdLabel.Caption) & ",'" & _
           UCase(empNameText.Text) & "','" & sex & "'," & Val(empAgeText.Text) & ",'" & UCase(empAddText.Text) & "','" & _
           empPhnText.Text & "','" & UCase(destiText.Text) & "','" & dateStr & "')"
           
  If MsgBox("Do you really save this data .", vbInformation + vbYesNo, "Conformation:") = vbYes Then
     'INSERTING RECORD
     empJDatabase.ExecuteSQL (insertSql)
  Else
     empNameText.SetFocus
     Exit Sub
  End If
 
  If MsgBox("Sucess ! Data is saved ." & vbCrLf & "Do you Continue ?", vbInformation + vbYesNo, "Sucess:") = vbYes Then
     Set empIDDyn = empJDatabase.CreateDynaset("select EMP_ID.NEXTVAL from dual", &H4&)
     empIdLabel.Caption = empIDDyn.Fields(0)
     Call CLEAR
     empNameText.SetFocus
     Exit Sub
  Else
     Unload Me
 End If

 
 
ERRORHANDLER: 'CODDING FOR ERROR DETECTION
    If empJSession.LastServerErr = 0 Then
       If empJDatabase.LastServerErr = 0 Then
         If Err.Number = 0 Then
         Else
           MsgBox "VB ERROR :" & vbCrLf & Err.Number & Err.Description, vbCritical, "VB ERROR:"
           Exit Sub
         End If
       Else
         MsgBox "DATABASE ERROR:" & vbCrLf & empJDatabase.LastServerErr & empJDatabase.LastServerErr, vbCritical, "DATABASE Error:"
         empJDatabase.LastServerErrReset
         Exit Sub
       End If
    Else
      MsgBox "SESSION ERROR :" & vbCrLf & empJSession.LastServerErr & empJSession.LastServerErrText, vbCritical, "SESSION Error:"
      empJSession.LastServerErrReset
      Exit Sub
    End If
   
     
End Sub

Private Sub saveCmd_GotFocus()
  
  StatusBar1.Panels(3) = "SAVE THE DATA.."

End Sub

Private Sub yearcombo_GotFocus()

 StatusBar1.Panels(3) = "SELECT THE JOINING DATE.."

End Sub

Private Sub yearcombo_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then  'THE ENTER KEY PRESSED
    saveCmd.SetFocus
  End If

End Sub
