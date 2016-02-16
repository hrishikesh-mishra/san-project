VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form EMP_SALARY_DETAIL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EMPLOYEE SALARY DETAIL "
   ClientHeight    =   8295
   ClientLeft      =   2505
   ClientTop       =   2220
   ClientWidth     =   10245
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
   ScaleHeight     =   8295
   ScaleWidth      =   10245
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   44
      Top             =   7800
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "5:53 PM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "9/25/2005"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "NUM"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton extCmd 
      Caption         =   "&Exit "
      Height          =   375
      Left            =   5640
      TabIndex        =   15
      ToolTipText     =   "Exit form this."
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cancelCmd 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   16
      ToolTipText     =   "Cancel the process."
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton okCmd 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      ToolTipText     =   "Save the data."
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox netSalText 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8520
      TabIndex        =   43
      Text            =   "0.00"
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox ttlDesText 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5424
      TabIndex        =   41
      Text            =   "0.00"
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox gSalText 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1848
      TabIndex        =   39
      Text            =   "0.00"
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox fPaytext 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6360
      TabIndex        =   13
      Text            =   "0.00"
      ToolTipText     =   "Enter Festival Pay"
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox sPaytext 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6360
      TabIndex        =   12
      Text            =   "0.00"
      ToolTipText     =   "Enter Special pay"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox taxText 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6360
      MaxLength       =   5
      TabIndex        =   11
      Text            =   "0.00"
      ToolTipText     =   "Enter Tax"
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox descText 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6360
      MaxLength       =   5
      TabIndex        =   9
      Text            =   "0.00"
      ToolTipText     =   "Enter Deduction"
      Top             =   3120
      Width           =   975
   End
   Begin VB.ComboBox monthCombo 
      Height          =   360
      Left            =   7215
      TabIndex        =   3
      Text            =   " "
      ToolTipText     =   "Select the month "
      Top             =   2280
      Width           =   1575
   End
   Begin VB.ComboBox yearCombo 
      Height          =   360
      Left            =   8790
      TabIndex        =   4
      Text            =   " "
      ToolTipText     =   "Select the year"
      Top             =   2280
      Width           =   975
   End
   Begin VB.ComboBox dayCombo 
      Height          =   360
      Left            =   6360
      TabIndex        =   2
      Text            =   " "
      ToolTipText     =   "Select the day "
      Top             =   2280
      Width           =   855
   End
   Begin VB.ComboBox salMonthCombo 
      Height          =   360
      Left            =   2160
      TabIndex        =   1
      Text            =   "Select month"
      ToolTipText     =   "Select the month of salary."
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox taText 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   8
      Text            =   "0.00"
      ToolTipText     =   "Enter TA"
      Top             =   4920
      Width           =   975
   End
   Begin VB.TextBox daText 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   7
      Text            =   "0.00"
      ToolTipText     =   "Enter DA"
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox hraText 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   6
      Text            =   "0.00"
      ToolTipText     =   "Enter the HRA."
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox bsText 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      ToolTipText     =   "Enter the basic salary"
      Top             =   3120
      Width           =   1575
   End
   Begin VB.ComboBox empIDcombo 
      Height          =   360
      Left            =   2160
      TabIndex        =   0
      Text            =   "Select Id"
      ToolTipText     =   "Select the employee no"
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   0
      Picture         =   "EMP_SALARY_DETAIL.frx":0000
      ToolTipText     =   "San's Employee Salary Entry Form"
      Top             =   0
      Width           =   3000
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   0
      X2              =   10200
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   4080
      X2              =   4080
      Y1              =   2880
      Y2              =   5400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   0
      X2              =   10200
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Label netSal 
      Caption         =   "Net Salary :-"
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   7200
      TabIndex        =   42
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label14 
      Caption         =   "Total Deduction :-"
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   3480
      TabIndex        =   40
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label13 
      Caption         =   "Gross Salary :-"
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   120
      TabIndex        =   38
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   0
      X2              =   10080
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   0
      X2              =   10200
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   0
      X2              =   10200
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label6 
      Caption         =   "%"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   7320
      TabIndex        =   37
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "%"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   7320
      TabIndex        =   36
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label12 
      Caption         =   "Festival Pay:-"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4320
      TabIndex        =   35
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "Special Pay :-"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4320
      TabIndex        =   34
      Top             =   4365
      Width           =   1695
   End
   Begin VB.Label Label10 
      Caption         =   "Day"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   6360
      TabIndex        =   33
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Month"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   7440
      TabIndex        =   32
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Year"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   8760
      TabIndex        =   31
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label eDaeLabel 
      Caption         =   "Entry Date :-"
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   4320
      TabIndex        =   30
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Salary for Month :-"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label taxLabel 
      Caption         =   "Tax :-"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4320
      TabIndex        =   28
      Top             =   3795
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "%"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   3120
      TabIndex        =   27
      Top             =   4920
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "%"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   3120
      TabIndex        =   26
      Top             =   4380
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "%"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   3120
      TabIndex        =   25
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Deduction :-"
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   4320
      TabIndex        =   24
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label taLabel 
      Caption         =   "T A :-"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label daLabel 
      Caption         =   "D A :-"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label hraLabel 
      Caption         =   "H R A :-"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label bsLabel 
      Caption         =   "Basic Salary :-"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label empNameLabel 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   6480
      TabIndex        =   19
      Top             =   1680
      Width           =   75
   End
   Begin VB.Label enamelabel 
      Caption         =   "Employee Name :-"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4320
      TabIndex        =   18
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label eidlabel 
      Caption         =   "Employee ID :-"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label EMPH 
      Alignment       =   2  'Center
      Caption         =   "EMPLOYEE SALARY  DETAIL ENTRY FORM"
      BeginProperty Font 
         Name            =   "PAQUITO"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   960
      TabIndex        =   10
      Top             =   720
      Width           =   9015
   End
End
Attribute VB_Name = "EMP_SALARY_DETAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************'
'**                        **'
'** EMPLOYEE SALARY DETAIL **'
'**                        **'
'****************************'
'VARIABLE DECLARAION

Option Explicit

Dim esdSession    As OraSession
Dim esdDatabase   As OraDatabase
Dim eidDyna       As OraDynaset
Dim enameDyna     As OraDynaset
Dim basic         As Double
Dim hra           As Double
Dim da            As Double
Dim ta            As Double
Dim deduc         As Double
Dim tax           As Double
Dim m             As Integer
Dim spay          As Double
Dim fpay          As Double
Dim grossSalary      As Double
Dim ttlDeduc      As Double
Dim netSalary        As Double
Dim checkDate     As Boolean
Dim flag          As Boolean
Dim dateStr       As String
Dim insertSql     As String

Private Sub checkValues()
 
 If bsText.Text = "" Or Val(bsText.Text) = 0 Then
    MsgBox "Error ! Basic salary is not present .", vbInformation, "Empty Error:"
    bsText.SetFocus
    flag = True
 ElseIf IsNumeric(bsText.Text) = False Or Val(bsText.Text) < 0 Then
   MsgBox "Error ! Invalid value in basic Salay.", vbInformation, "Invalid Value:"
   bsText.SetFocus
   flag = True
 ElseIf hraText.Text = "" Then
       hraText.Text = "0.00"
 ElseIf IsNumeric(hraText.Text) = False Or Val(hraText.Text) < 0 Or Val(hraText.Text) > 99 Then
    MsgBox "Error ! Invalid Value in H R A .", vbInformation, "Invalid Value:"
    flag = True
    bsText.SetFocus
 ElseIf daText.Text = "" Then
        daText.Text = "0.00"
 ElseIf IsNumeric(daText.Text) = False Or Val(daText.Text) < 0 Or Val(daText.Text) > 99 Then
     MsgBox "Error ! Invalid value in  D A.", vbInformation, "Invalid Value:"
     flag = True
     daText.SetFocus
 ElseIf taText.Text = "" Then
     taText.Text = "0.00"
 ElseIf IsNumeric(taText.Text) = False Or Val(taText.Text) < 0 Or Val(taText.Text) > 99 Then
     MsgBox "Error ! Invalid value in T A .", vbInformation, "Invalid Value:"
     flag = True
     taText.SetFocus
 ElseIf descText.Text = "" Then
     descText.Text = "0.00"
 ElseIf IsNumeric(descText.Text) = False Or Val(descText.Text) < 0 Or Val(descText.Text) > 99 Then
     MsgBox "Error ! Invalid value in Deduction .", vbInformation, "Invalid Value:"
     flag = True
     descText.SetFocus
 ElseIf taxText.Text = "" Then
     taxText.Text = "0.00"
 ElseIf IsNumeric(taxText.Text) = False Or Val(taxText.Text) < 0 Or Val(taxText.Text) > 99 Then
     MsgBox "Error ! Invalid value in Tax .", vbInformation, "Invalid Value:"
     flag = True
     taxText.SetFocus
 ElseIf fPaytext.Text = "" Then
     fPaytext.Text = "0.00"
 ElseIf IsNumeric(fPaytext.Text) = False Or Val(fPaytext.Text) < 0 Then
     MsgBox "Error ! Invalid value in Festival Pay .", vbInformation, "Invalid Value:"
     flag = True
     fPaytext.SetFocus
 ElseIf sPaytext.Text = "" Then
     sPaytext.Text = "0.00"
 ElseIf IsNumeric(sPaytext.Text) = False Or Val(sPaytext.Text) < 0 Then
     MsgBox "Error ! Invalid value in Special Pay .", vbInformation, "Invalid Value:"
     flag = True
    sPaytext.SetFocus
 Else
   flag = False
 End If
 
End Sub
Private Sub calculateValue()
  
  basic = Val(bsText.Text)
   
  hra = basic * Val(hraText.Text) / 100
  da = basic * Val(daText.Text) / 100
  ta = basic * Val(taText.Text) / 100
  deduc = basic * Val(descText.Text) / 100
  tax = basic * Val(taxText.Text) / 100
  spay = Val(sPaytext.Text)
  fpay = Val(fPaytext.Text)
  grossSalary = basic + hra + da + ta + spay + fpay
  ttlDeduc = tax + deduc
  netSalary = grossSalary - ttlDeduc
  
   gSalText.Text = grossSalary
   ttlDesText.Text = ttlDeduc
   netSalText.Text = netSalary
  
End Sub

Private Sub CLEAR()
 
  Call FILLCOMBODATE(dayCombo, monthCombo, yearCombo)
 
  bsText.Text = ""
  hraText.Text = "0.00"
  daText.Text = "0.00"
  taText.Text = "0.00"
  descText.Text = "0.00"
  taxText.Text = "0.00"
  sPaytext.Text = "0.00"
  fPaytext.Text = "0.00"
  gSalText.Text = "0.00"
  ttlDesText.Text = "0.00"
  netSalText.Text = "0.00"
End Sub

Private Sub bsText_GotFocus()

  StatusBar1.Panels(3) = "Enter the basic salary."

End Sub

Private Sub bsText_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     hraText.SetFocus
  End If

End Sub

Private Sub bsText_LostFocus()
  
  StatusBar1.Panels(3) = ""

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command1_GotFocus()
  
  StatusBar1.Panels(3) = "Clear the information."

End Sub

Private Sub Command1_LostFocus()
  
  StatusBar1.Panels(3) = ""

End Sub

Private Sub cancelCmd_Click()
  
  Call CLEAR

End Sub

Private Sub daText_GotFocus()
   
   StatusBar1.Panels(3) = "Ente the da."

End Sub

Private Sub daText_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
     taText.SetFocus
  End If

End Sub

Private Sub daText_LostFocus()

  StatusBar1.Panels(3) = ""

End Sub

Private Sub dayCombo_GotFocus()
 
   StatusBar1.Panels(3) = " Select the day."
 
End Sub

Private Sub dayCombo_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
      monthCombo.SetFocus
  End If

End Sub

Private Sub dayCombo_LostFocus()
  
  StatusBar1.Panels(3) = ""

End Sub

Private Sub descText_GotFocus()
   
   StatusBar1.Panels(3) = "Enter the Deduction."

End Sub

Private Sub descText_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     taxText.SetFocus
  End If

End Sub

Private Sub descText_LostFocus()
  
  StatusBar1.Panels(3) = ""

End Sub

Private Sub empIDcombo_Click()
  
  If empIDcombo.Text = "" Then
     Exit Sub
  End If

 Set enameDyna = esdDatabase.CreateDynaset("select ENAME from HRISHI.CURRENT_EMP_DETAIL where EID =" & Val(empIDcombo.Text) & "", &H0&)
 empNameLabel.Caption = enameDyna.Fields(0)

End Sub

Private Sub empIDcombo_GotFocus()
 
  StatusBar1.Panels(3) = "Select the Employee id."
 
End Sub

Private Sub empIDcombo_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     salMonthCombo.SetFocus
  End If

End Sub

Private Sub empIDcombo_LostFocus()
   
   StatusBar1.Panels(3) = ""

End Sub

Private Sub extCmd_Click()
 
   If MsgBox("Do you want exit ?", vbExclamation + vbYesNo, "Exit:") = vbYes Then
      Unload Me
   End If
 
End Sub

Private Sub extCmd_GotFocus()
   
   StatusBar1.Panels(3) = "Exit"

End Sub

Private Sub extCmd_LostFocus()
   
   StatusBar1.Panels(3) = ""

End Sub

Private Sub Form_Load()

  Call FILLCOMBODATE(dayCombo, monthCombo, yearCombo)
  For m = 1 To 12
    salMonthCombo.AddItem MonthName(m)
  Next
  
  salMonthCombo.Text = MonthName(MONTH(Date))
 
 On Error GoTo ERRORHANDLER

  Set esdSession = CreateObject("oracleinprocserver.xorasession")
  Set esdDatabase = esdSession.OpenDatabase("jms", "hrishi/jms", &H0&)
  Set eidDyna = esdDatabase.CreateDynaset("Select EID from HRISHI.CURRENT_EMP_DETAIL", &H0&)

  If eidDyna.EOF Then
    MsgBox "Current employee is zero.", vbInformation
    okCmd.Enabled = False
  End If

  While Not eidDyna.EOF
      empIDcombo.AddItem eidDyna.Fields(0)
      eidDyna.MoveNext
      empIDcombo.ListIndex = 0
  Wend

ERRORHANDLER:
 If esdSession.LastServerErr = 0 Then
     If esdDatabase.LastServerErr = 0 Then
        If Err.Number = 0 Then
        Else
           MsgBox "VB ERROR :" & vbCrLf & Err.Number & Err.Description, vbCritical, "VB ERROR"
        End If
    Else
        MsgBox "DATABASE ERROR:" & vbCrLf & esdDatabase.LastServerErr & esdDatabase.LastServerErrText, vbCritical, "DATABASE ERROR:"
        esdDatabase.LastServerErrReset
        Unload Me
    End If
  Else
     MsgBox "SESSION ERROR:" & vbCrLf & esdSession.LastServerErr & esdSession.LastServerErrText, vbCritical, "SESSION ERROR:"
     esdSession.LastServerErrReset
  End If
  
End Sub

Private Sub fPaytext_GotFocus()

  StatusBar1.Panels(3) = "Enter the Festival pay."

End Sub

Private Sub fPaytext_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       okCmd.SetFocus
    End If

End Sub

Private Sub fPaytext_LostFocus()
    
    StatusBar1.Panels(3) = ""

End Sub


Private Sub hraText_GotFocus()
    
    StatusBar1.Panels(3) = "Enter the  hra ."

End Sub

Private Sub hraText_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        daText.SetFocus
    End If

End Sub

Private Sub hraText_LostFocus()
    
    StatusBar1.Panels(3) = ""

End Sub

Private Sub monthCombo_GotFocus()
   
   StatusBar1.Panels(3) = "Select the month ."

End Sub

Private Sub monthCombo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        yearCombo.SetFocus
    End If

End Sub

Private Sub monthCombo_LostFocus()
    
    StatusBar1.Panels(3) = ""

End Sub

Private Sub OkCmd_Click()
  
  checkDate = VERIFY_DATE(Val(dayCombo.Text), monthCombo, Val(yearCombo.Text))
  Call checkValues
  
  If checkDate = False Then
     MsgBox "Error ! Invalid date .", vbCritical, "Date Error."
     Exit Sub
  ElseIf flag Then
     Exit Sub
  End If

  Call calculateValue
  On Error GoTo ERRORHANDLER


   dateStr = dayCombo.Text + "-" + monthCombo.Text + "-" + yearCombo.Text
   insertSql = "insert into HRISHI.EMP_SALARY_DETAIL values(" & Val(empIDcombo.Text) & ",'" & _
            UCase(empNameLabel.Caption) & "','" & salMonthCombo.Text & "'," & Val(bsText.Text) & "," & _
            Val(hraText.Text) & "," & Val(daText.Text) & "," & Val(taText.Text) & "," & Val(descText.Text) & "," & Val(taxText.Text) & "," & spay & "," & fpay & "," & _
            grossSalary & "," & ttlDeduc & "," & netSalary & ",'" & dateStr & "')"
        
  If MsgBox("Are sure to save the data ? " & vbCrLf & "Continue.", vbInformation + vbOKCancel, "Conformation:") = vbOK Then
      esdDatabase.ExecuteSQL (insertSql)
      If MsgBox("Sucess ! data is saved ." & vbCrLf & "Do you want continue ? ", vbInformation + vbYesNo, "Sucess:") = vbYes Then
         Call CLEAR
         Exit Sub
     Else
         Unload Me
     End If
 Else
    Exit Sub
 End If

ERRORHANDLER:
If esdSession.LastServerErr = 0 Then
    If esdDatabase.LastServerErr = 0 Then
      If Err.Number = 0 Then
      Else
         MsgBox "VB ERROR :" & vbCrLf & Err.Number & Err.Description, vbCritical, "VB ERROR"
         Exit Sub
      End If
    Else
      MsgBox "DATABASE ERROR:" & vbCrLf & esdDatabase.LastServerErr & esdDatabase.LastServerErrText, vbCritical, "DATABASE ERROR:"
      esdDatabase.LastServerErrReset
      Exit Sub
    End If
Else
    MsgBox "SESSION ERROR:" & vbCrLf & esdSession.LastServerErr & esdSession.LastServerErrText, vbCritical, "SESSION ERROR:"
    esdSession.LastServerErrReset
    Exit Sub
End If
  
End Sub

Private Sub okCmd_GotFocus()
  
  StatusBar1.Panels(3) = "Save the Data ."

End Sub

Private Sub okCmd_LostFocus()
  
  StatusBar1.Panels(3) = ""

End Sub

Private Sub salMonthCombo_GotFocus()
  
  StatusBar1.Panels(3) = " Select the month of salary."

End Sub

Private Sub salMonthCombo_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     dayCombo.SetFocus
  End If

End Sub

Private Sub salMonthCombo_LostFocus()

  StatusBar1.Panels(3) = ""

End Sub

Private Sub sPaytext_GotFocus()
   
   StatusBar1.Panels(3) = "Enter the  Special pay. "

End Sub

Private Sub sPaytext_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     fPaytext.SetFocus
  End If

End Sub

Private Sub sPaytext_LostFocus()
   
   StatusBar1.Panels(3) = ""

End Sub

Private Sub taText_GotFocus()
   
   StatusBar1.Panels(3) = "Enter the TA."

End Sub

Private Sub taText_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     descText.SetFocus
  End If

End Sub

Private Sub taText_LostFocus()
  
  StatusBar1.Panels(3) = ""

End Sub

Private Sub taxText_GotFocus()
  
  StatusBar1.Panels(3) = "Enter the Tax."

End Sub

Private Sub taxText_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     sPaytext.SetFocus
  End If

End Sub

Private Sub taxText_LostFocus()
  
  StatusBar1.Panels(3) = ""

End Sub

Private Sub yearcombo_GotFocus()
   
   StatusBar1.Panels(3) = "Select the year ."

End Sub

Private Sub yearcombo_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 Then
   bsText.SetFocus
 End If

End Sub

Private Sub yearCombo_LostFocus()

  StatusBar1.Panels(3) = ""

End Sub
