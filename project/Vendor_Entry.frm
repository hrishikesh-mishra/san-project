VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form VEN_ENTRY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VENDOR FORM [ENTRY]"
   ClientHeight    =   8295
   ClientLeft      =   3960
   ClientTop       =   2775
   ClientWidth     =   9165
   DrawMode        =   9  'Not Mask Pen
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
   ScaleWidth      =   9165
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   20
      Top             =   7800
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2011
            MinWidth        =   2011
            TextSave        =   "9/23/2005"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   2011
            MinWidth        =   2011
            TextSave        =   "2:03 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   11360
            MinWidth        =   11360
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton extCmd 
      Caption         =   "&Exit"
      Height          =   615
      Left            =   6000
      TabIndex        =   18
      ToolTipText     =   "Exit."
      Top             =   6960
      Width           =   1335
   End
   Begin VB.ComboBox dayCombo 
      Height          =   360
      ItemData        =   "Vendor_Entry.frx":0000
      Left            =   3600
      List            =   "Vendor_Entry.frx":0002
      TabIndex        =   17
      ToolTipText     =   "Select Day."
      Top             =   6120
      Width           =   855
   End
   Begin VB.TextBox delOfText 
      Height          =   360
      Left            =   3600
      TabIndex        =   16
      ToolTipText     =   "Enter Detail"
      Top             =   5160
      Width           =   3855
   End
   Begin VB.TextBox venEIDText 
      Height          =   360
      Left            =   3600
      TabIndex        =   15
      ToolTipText     =   "Enter Email ID."
      Top             =   4440
      Width           =   5055
   End
   Begin VB.TextBox venPhNText 
      Height          =   360
      Left            =   3600
      TabIndex        =   14
      ToolTipText     =   "Enter Phone NO."
      Top             =   3600
      Width           =   3855
   End
   Begin VB.TextBox venAddText 
      Height          =   360
      Left            =   3600
      TabIndex        =   13
      ToolTipText     =   "Enter vendor Address."
      Top             =   2760
      Width           =   5055
   End
   Begin VB.TextBox venNameText 
      Height          =   360
      Left            =   3600
      TabIndex        =   12
      ToolTipText     =   "Enter vendor Name."
      Top             =   2040
      Width           =   3975
   End
   Begin VB.ComboBox monthCombo 
      Height          =   360
      Left            =   4440
      TabIndex        =   4
      ToolTipText     =   "Select Month."
      Top             =   6120
      Width           =   1815
   End
   Begin VB.ComboBox yearCombo 
      Height          =   360
      Left            =   6240
      TabIndex        =   3
      ToolTipText     =   "Select Year."
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cancelCmd 
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      ToolTipText     =   "Clear the Information."
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton saveCmd 
      Caption         =   "&Save"
      Height          =   615
      Left            =   4200
      TabIndex        =   2
      ToolTipText     =   "Save the information."
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FF0000&
      X1              =   2160
      X2              =   7560
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   0
      Picture         =   "Vendor_Entry.frx":0004
      ToolTipText     =   "San's Vendor Entry."
      Top             =   0
      Width           =   3000
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   9240
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3600
      TabIndex        =   19
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label venIdlabel 
      Caption         =   "Vendor  ID :-"
      ForeColor       =   &H00008000&
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   11
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label venNameLabel 
      Caption         =   "Vendor Name:-"
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Vendor Address :-"
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label venPhNLabel 
      Caption         =   "Vendor Phone No :-"
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label venEIDLabel 
      Caption         =   "Vendor  Email ID :-"
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label delOfLabel 
      Caption         =   "Deler Of :-"
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label EDateLabel 
      Caption         =   "Entry Date:- "
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FF0000&
      X1              =   5760
      X2              =   5760
      Y1              =   6840
      Y2              =   7680
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF0000&
      X1              =   3960
      X2              =   3960
      Y1              =   6840
      Y2              =   7680
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF0000&
      X1              =   7560
      X2              =   7560
      Y1              =   6840
      Y2              =   7680
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      X1              =   2160
      X2              =   2160
      Y1              =   6840
      Y2              =   7680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      X1              =   2160
      X2              =   7560
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line1 
      X1              =   2160
      X2              =   7560
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label venLabel 
      Alignment       =   2  'Center
      Caption         =   "VENDOR ENTRY"
      BeginProperty Font 
         Name            =   "ZINNIA"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Index           =   0
      Left            =   2520
      TabIndex        =   0
      Top             =   600
      Width           =   4455
   End
End
Attribute VB_Name = "VEN_ENTRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************'
'**                   **'
'** VENDOR ENTRY FORM **'
'**                   **'
'***********************'

' VARIABLE DECLARATION

Option Explicit

Dim venSession  As OraSession
Dim venDatabase As OraDatabase
Dim venIdSln    As OraDynaset
Dim insertStr   As String
Dim dateStr     As String
Dim checkDate   As Boolean
Dim FlagEmpty   As Boolean

Private Sub cancelCmd_GotFocus()
    
    StatusBar1.Panels(3) = "Clear the Information ..."

End Sub

Private Sub dayCombo_GotFocus()

    StatusBar1.Panels(3) = "Select the day.. "

End Sub

Private Sub dayCombo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        monthCombo.SetFocus
    End If

End Sub

Private Sub delOfText_GotFocus()
    
    StatusBar1.Panels(3) = "Enter the Deler of  ... "

End Sub

Private Sub delOfText_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        dayCombo.SetFocus
    End If
 
End Sub

Private Sub extCmd_Click()

    If MsgBox("DO YOU REALLY WANT TO EXIT", vbYesNo + vbExclamation, "EXIT") = vbYes Then
        Unload Me
    End If

End Sub

Private Sub check_empty()
    
    If venNameText.Text = "" Then
      MsgBox "Error ! Vendor Name isn't present ."
      FlagEmpty = True
    ElseIf venAddText.Text = "" Then
      MsgBox "Error ! Vendor address isn't present ."
      FlagEmpty = True
    ElseIf venPhNText.Text = "" Then
      MsgBox "Error ! Vendor Phone No isn't present. "
      FlagEmpty = True
    ElseIf venEIDText.Text = "" Then
      MsgBox "Error ! vendor Email Id isn't present. "
      FlagEmpty = True
    ElseIf delOfText.Text = "" Then
      MsgBox "Error ! Deler of column is empty."
      FlagEmpty = True
    Else
       FlagEmpty = False
    End If
    
    If InStr(venEIDText.Text, ".") And InStr(venEIDText.Text, "@") Then
    Else
       MsgBox "Error ! Email ID isn't properly defined "
       FlagEmpty = True
       venEIDText.SetFocus
    End If
    
End Sub

Private Sub extCmd_GotFocus()
    
    StatusBar1.Panels(3) = "Exit ...."

End Sub

Private Sub Form_Load()
    
        Call FILLCOMBODATE(dayCombo, monthCombo, yearCombo)
        
        Set venSession = CreateObject("oracleinprocserver.xorasession")
        Set venDatabase = venSession.OpenDatabase("jms", "hrishi/jms", &H0&)
        Set venIdSln = venDatabase.CreateDynaset("select  vendor_no.nextval from dual ", &H4&)
        Label2.Caption = venIdSln.Fields(0)

End Sub

Private Sub monthCombo_GotFocus()
    
    StatusBar1.Panels(3) = "Select the month ..."

End Sub

Private Sub monthCombo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        yearCombo.SetFocus
    End If

End Sub

Private Sub saveCmd_Click()
    Call check_empty       'Calling a subroutine for checking for the _
                            blank column
    checkDate = VERIFY_DATE(Val(dayCombo.Text), monthCombo.Text, Val(yearCombo.Text))
    If checkDate = False Then
        MsgBox "Error ! Invaild date .", vbCritical, "Error"
        Exit Sub
    ElseIf FlagEmpty Then
        Exit Sub
    End If

    On Error GoTo ERRORHANDLER
        
    dateStr = dayCombo.Text + "-" + monthCombo.Text + "-" + yearCombo.Text
    insertStr = "insert into HRISHI.VENDOR_DETAIL values( " & (Label2.Caption) _
             & ",'" & UCase(venNameText.Text) & "','" & UCase(venAddText.Text) _
             & "','" & UCase(venPhNText.Text) & "','" & LCase(venEIDText.Text) _
             & "','" & UCase(delOfText.Text) & "','" & dateStr & "')"

    venDatabase.ExecuteSQL (insertStr)
  
ERRORHANDLER:                        'CODING FOR ERRORHANDLER
   If venSession.LastServerErr = 0 Then
      If venDatabase.LastServerErr = 0 Then
         If Err.Number = 0 Then
         If MsgBox("Sucess ! Do you want more ...", vbYesNo + _
            vbExclamation, "EXIT") = vbYes Then
            venNameText.SetFocus
            Set venIdSln = venDatabase.CreateDynaset("select  vendor_no.nextval from dual ", &H4&)
             Label2.Caption = venIdSln.Fields(0)
         Else
           Unload Me
         End If
         Else
           MsgBox "Vb Error : " & Err.Number & Err.Description
           Exit Sub
         End If
      Else
        MsgBox " Database  Error : " & venDatabase.LastServerErr & venDatabase.LastServerErrText
        venDatabase.LastServerErrReset
        Exit Sub
      End If
   Else
     MsgBox "Session Error : " & venSession.LastServerErr & venSession.LastServerErrText
     venSession.LastServerErrReset
     Exit Sub
   End If
  
   
End Sub


Private Sub saveCmd_GotFocus()

    StatusBar1.Panels(3) = "Save the Information.... "

End Sub

Private Sub venAddText_GotFocus()
    
    StatusBar1.Panels(3) = "Enter the Address..."

End Sub

Private Sub venAddText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        venPhNText.SetFocus
    End If

End Sub

Private Sub venEIDText_GotFocus()
    
    StatusBar1.Panels(3) = "Enter the Email ID ... "

End Sub

Private Sub venEIDText_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        delOfText.SetFocus
    End If

End Sub

Private Sub venNameText_GotFocus()

    StatusBar1.Panels(3) = "Enter the Name..."

End Sub

Private Sub venNameText_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       venAddText.SetFocus
    End If

End Sub

Private Sub venPhNText_GotFocus()
    
    StatusBar1.Panels(3) = "Enter the Phone No ... "

End Sub

Private Sub venPhNText_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        venEIDText.SetFocus
    End If

End Sub

Private Sub yearcombo_GotFocus()
    
    StatusBar1.Panels(3) = "Select the year ..."

End Sub

Private Sub yearcombo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       saveCmd.SetFocus
    End If

End Sub




