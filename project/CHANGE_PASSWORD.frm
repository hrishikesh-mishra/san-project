VERSION 5.00
Begin VB.Form CHANGE_PASSWORD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CHANGE PASSWORD"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   6435
   Begin VB.CommandButton exitCmd 
      Caption         =   "Exit"
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
      Left            =   4680
      TabIndex        =   6
      ToolTipText     =   "Exit From This."
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cancelCmd 
      Caption         =   "Cancel"
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
      TabIndex        =   5
      ToolTipText     =   "Cancel the process "
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton okCmd 
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
      Left            =   480
      TabIndex        =   4
      ToolTipText     =   "Save the setting "
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox conformPasswordText 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   3
      ToolTipText     =   "Enter Confirm password"
      Top             =   3840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox newPasswordText 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Enter new password"
      Top             =   3120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox oldPasswordText 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Enter old password"
      Top             =   2400
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox userIDText 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      ToolTipText     =   "Enter user ID"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   585
      Left            =   0
      Picture         =   "CHANGE_PASSWORD.frx":0000
      ToolTipText     =   "San's Change Password"
      Top             =   0
      Width           =   3000
   End
   Begin VB.Label Label5 
      Caption         =   "CHANGE PASSWORD"
      BeginProperty Font 
         Name            =   "SEYMOUR"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1680
      TabIndex        =   11
      Top             =   720
      Width           =   3855
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000001&
      X1              =   2040
      X2              =   2040
      Y1              =   4440
      Y2              =   5400
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000001&
      X1              =   4080
      X2              =   4080
      Y1              =   4440
      Y2              =   5400
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   6360
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   6360
      X2              =   6360
      Y1              =   1320
      Y2              =   5400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   1320
      Y2              =   5400
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C00000&
      X1              =   2520
      X2              =   2520
      Y1              =   1320
      Y2              =   4440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   120
      X2              =   6360
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   6360
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000A&
      Caption         =   "Conform Password"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000A&
      Caption         =   "New Password"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      Caption         =   "Old Password"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   2175
   End
End
Attribute VB_Name = "CHANGE_PASSWORD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************'
'**                      **'
'** CHANGE PASSWORD FORM **'
'**                      **'
'**************************'

'VARIABLE DECLERATION
 
 Option Explicit
 
 Dim chpSession As OraSession
 Dim chpDatabase As OraDatabase
 Dim chpDyn      As OraDynaset
 Dim changeSql   As String
 Dim EmptyFlag   As Boolean

Private Sub cancelCmd_Click()
  
  userIDText.Text = ""           'IF CANCEL CMD IS CLICK THEN
  oldPasswordText.Text = ""      'CLEAR ALL TEXT BOX
  newPasswordText.Text = ""
  conformPasswordText.Text = ""

End Sub
Private Sub CheckEmpty()
                                
  'A SUBROUTINE FOR CHECKING EMPTY TEXTBOX AND GENERATING ERROR MESSAGE
 If userIDText.Text = "" Then
     MsgBox "Empty ! UserId  is Blank.", vbExclamation, "Empty:"
     EmptyFlag = True
     userIDText.SetFocus
 ElseIf oldPasswordText.Text = "" And ModuleVarious.LogOnUser <> "administrator" Then
     MsgBox "Empty ! Old Password is Blank.", vbExclamation, "Empty:"
     EmptyFlag = True
     oldPasswordText.SetFocus
 ElseIf newPasswordText.Text = "" Then
     MsgBox "Empty ! New Password is Blank.", vbExclamation, "Empty:"
     EmptyFlag = True
     newPasswordText.SetFocus
 ElseIf conformPasswordText.Text = "" Then
     MsgBox "Empty ! Conform Password is Blank .", vbExclamation, "Empty:"
     EmptyFlag = True
     conformPasswordText.SetFocus
 Else
    EmptyFlag = False
 End If

End Sub

Private Sub conformPasswordText_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then 'IF ENTER KEY IS PRESSED
   OkCmd.SetFocus
End If
 
End Sub

Private Sub exitCmd_Click()

 'CONFIRMING FOR EXITING
 If MsgBox("Do you want to exit ?", vbExclamation + vbYesNo, "Exit:") = vbYes Then
    Unload Me
 End If

End Sub

Private Sub Form_Load()

 On Error GoTo ERRORHANDLER
 Set chpSession = CreateObject("oracleinprocserver.xorasession")
 Set chpDatabase = chpSession.OpenDatabase("jms", "hrishi/jms", &H0&)
 
ERRORHANDLER:          'CODING FOR ERROR HANDLER
  If chpSession.LastServerErr = 0 Then
     If chpDatabase.LastServerErr = 0 Then
        If Err.Number = 0 Then
        Else
           MsgBox "VB ERROR: " & Err.Number & Err.Description, vbCritical, "VB Error:"
           End
        End If
     Else
        MsgBox "DATABASE ERROR:" & chpDatabase.LastServerErr & chpDatabase.LastServerErrText, vbCritical, "SESSION Error:"
        chpDatabase.LastServerErrReset
        End If
 Else
  MsgBox "SESSION ERROR:" & chpSession.LastServerErr & chpSession.LastServerErrText, vbCritical, "SESSION Error:"
  chpSession.LastServerErrReset
  End
End If

End Sub

Private Sub newPasswordText_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then           'IF ENTER KEY IS PRESSED
   conformPasswordText.Visible = True
   conformPasswordText.SetFocus
 End If
 
End Sub

Private Sub OkCmd_Click()

  On Error GoTo ERRORHANDLER
  Call CheckEmpty   'CALLING SUBROUTINE FOR CHECKING EMPTY
  If EmptyFlag Then  'IF ANY ONE EMPTY THEN EXIT SUB
     Exit Sub
  End If
                   'CHECKING FOR NEW PASSWORD AND CONFIRM PASSWORD IS SAME
  If newPasswordText.Text <> conformPasswordText.Text Then
     MsgBox "Both New Password aren't match.", vbExclamation, "San's Project:"
     newPasswordText.SetFocus
     Exit Sub
  End If
                    'IF LOGIN USER IS ADMINISTRATOR THEN OLD PASSWORD ISN'T REQURID
  If ModuleVarious.LogOnUser = "administrator" Then
     Set chpDyn = chpDatabase.CreateDynaset("select * from hrishi.sanproject_user where user_id='" & userIDText.Text & "' ", &H4&)
  Else
     Set chpDyn = chpDatabase.CreateDynaset("select * from hrishi.sanproject_user where user_id='" & userIDText.Text & "' and password='" & oldPasswordText.Text & "'", &H4&)
  End If
  'CHECKING FOR AUTHORIZED USER WHO WISH TO CHANGE THE PASSWORD
  'IF NOT THEN EXIT SUB
  If chpDyn.EOF Then
     MsgBox "Error ! Not Authorized User .", vbCritical, "San's Project:"
     Exit Sub
  End If
   
  'CREATING SQL FOR CHANGING PASSWORD
  changeSql = "update hrishi.sanproject_user set password='" & newPasswordText.Text & _
              "'where user_id ='" & userIDText.Text & "'"
  chpDatabase.ExecuteSQL (changeSql)

ERRORHANDLER:
  If chpSession.LastServerErr = 0 Then
     If chpDatabase.LastServerErr = 0 Then
        If Err.Number = 0 Then
           MsgBox "Success ! " & vbCrLf & userIDText.Text & " ,Your Password is Changed.", _
                   vbInformation, "Success:"
           Unload Me
           Exit Sub
        Else
           MsgBox "Vb Error : " & Err.Number & Err.Description, vbCritical, "VB Error:"
           Exit Sub
        End If
     Else
        MsgBox "Database Error :" & chpDatabase.LastServerErr & chpDatabase.LastServerErrText, vbCritical, "Database Error:"
        chpDatabase.LastServerErrReset
        Exit Sub
     End If
  Else
      MsgBox "Session Error :" & chpSession.LastServerErr & chpSession.LastServerErrText, vbCritical, "Session Error:"
      chpSession.LastServerErrReset
      Exit Sub
  End If

End Sub

Private Sub oldPasswordText_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 Then          'IF ENTER KEY IS PRESSED
    newPasswordText.Visible = True
    newPasswordText.SetFocus
  End If
 
End Sub

Private Sub userIDText_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 Then 'IF ENTER KEY IS PRESSED
    oldPasswordText.Visible = True
    oldPasswordText.SetFocus
    
 End If

End Sub
