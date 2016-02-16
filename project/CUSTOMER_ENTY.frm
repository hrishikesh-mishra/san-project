VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form CUSTOMER_ENTRY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CUSTOMER [ENTRY]"
   ClientHeight    =   7515
   ClientLeft      =   3360
   ClientTop       =   2925
   ClientWidth     =   8790
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   8790
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   18
      Top             =   7020
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1960
            MinWidth        =   1960
            TextSave        =   "12/16/2005"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1940
            MinWidth        =   1940
            TextSave        =   "9:31 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   11359
            MinWidth        =   11359
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox custNameText 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      ToolTipText     =   "Enter the Custome's name "
      Top             =   2115
      Width           =   4215
   End
   Begin VB.TextBox custAddText 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      ToolTipText     =   "Enter the Customer's Address."
      Top             =   2790
      Width           =   5055
   End
   Begin VB.TextBox custPhNText 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      ToolTipText     =   "Enter the Customer's Phone No."
      Top             =   3450
      Width           =   4575
   End
   Begin VB.TextBox totalSaleText 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   16
      Text            =   "0.00"
      Top             =   4125
      Width           =   2535
   End
   Begin VB.ComboBox dayCombo 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   3360
      TabIndex        =   3
      ToolTipText     =   "Select The Day."
      Top             =   4800
      Width           =   1095
   End
   Begin VB.ComboBox monthCombo 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   4440
      TabIndex        =   4
      ToolTipText     =   "Select the Month."
      Top             =   4800
      Width           =   1935
   End
   Begin VB.ComboBox yearCombo 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   6360
      TabIndex        =   5
      ToolTipText     =   "Select The year."
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton ExitCmd 
      Caption         =   "&EXIT"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5153
      TabIndex        =   9
      ToolTipText     =   "Exit from this."
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton clearCmd 
      Caption         =   "&CLEAR"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3953
      TabIndex        =   8
      ToolTipText     =   "Clear the Information"
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton oKCmd 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2633
      TabIndex        =   7
      ToolTipText     =   "Save the Information"
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   0
      Picture         =   "CUSTOMER_ENTY.frx":0000
      ToolTipText     =   "San's Cutomer Detail Enter Form"
      Top             =   0
      Width           =   3000
   End
   Begin VB.Line Line18 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   8760
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line17 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   3240
      X2              =   3240
      Y1              =   1920
      Y2              =   5400
   End
   Begin VB.Line Line16 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   3240
      X2              =   8640
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line15 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   8640
      X2              =   8640
      Y1              =   1200
      Y2              =   5400
   End
   Begin VB.Line Line14 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   1080
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line13 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   1080
      X2              =   1080
      Y1              =   1920
      Y2              =   5400
   End
   Begin VB.Line Line12 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   8760
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line11 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   8760
      X2              =   0
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   6600
      X2              =   6600
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   2280
      X2              =   2280
      Y1              =   5640
      Y2              =   6720
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   2280
      X2              =   6600
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   2280
      X2              =   6600
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   5040
      X2              =   5040
      Y1              =   5760
      Y2              =   6600
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   3840
      X2              =   3840
      Y1              =   5760
      Y2              =   6600
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   6480
      X2              =   6480
      Y1              =   5760
      Y2              =   6600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   2400
      X2              =   2400
      Y1              =   5760
      Y2              =   6600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   2400
      X2              =   6480
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   2400
      X2              =   6480
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label custIdLabel 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3360
      TabIndex        =   17
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label CUST_ID 
      Caption         =   "Customer Id :-"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label custNameLabel 
      Caption         =   "Name :-"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   2115
      Width           =   1815
   End
   Begin VB.Label custAddLabel 
      Caption         =   "Address :-"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   2790
      Width           =   1815
   End
   Begin VB.Label custPhNLabel 
      Caption         =   "Phone No. :-"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   3450
      Width           =   1815
   End
   Begin VB.Label edateLabel 
      Caption         =   "Entry Date  :-"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label TotalSaleLabel 
      Caption         =   "Total Sale"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   1200
      TabIndex        =   10
      Top             =   4125
      Width           =   1815
   End
   Begin VB.Label CUST_FORM 
      Alignment       =   2  'Center
      Caption         =   "CUSTOMER DETAIL FORM"
      BeginProperty Font 
         Name            =   "HAYWARD"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   720
      Width           =   6375
   End
End
Attribute VB_Name = "CUSTOMER_ENTRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+++++++++++++++++++++++++'
'++                     ++'
'++ CUSTOMER ENTRY FORM ++'
'++                     ++'
'+++++++++++++++++++++++++'

'VARIABLE DECLARATION
Option Explicit

Dim CustSession    As OraSession
Dim CustDatabase   As OraDatabase
Dim CustIdDyn      As OraDynaset
Dim FlagEmpty      As Boolean
Dim checkDate      As Boolean
Dim dateStr        As String
Dim insertSql       As String

Private Sub SaveCustData()

  On Error GoTo ERRORHANDLER
  'SQL COMMAND FOR INSERT THE INFORAMTION OF CUSTOMER
  insertSql = "INSERT INTO HRISHI.CUSTOMER_DETAIL VALUES ( " & Val(custIdLabel) & _
            ",'" & UCase(custNameText.Text) & "' ,'" & UCase(custAddText.Text) & "','" _
            & custPhNText.Text & "',0,'" & dateStr & "')"
  
  'INSERTING THE CUSTOMER INFORMATION
  CustDatabase.ExecuteSQL (insertSql)

ERRORHANDLER:
  If CustSession.LastServerErr = 0 Then
      If CustDatabase.LastServerErr = 0 Then
         If Err.Number = 0 Then
            SALE_ENTRY.IdLabel.Caption = custIdLabel.Caption
            SALE_ENTRY.NameLabel.Caption = UCase(custNameText.Text)
            SALE_ENTRY.typOfCust.Caption = "customer"
            Unload Me
        Else
            MsgBox "VB ERROR : " & Err.Number & " :: " & Err.Description, _
            vbCritical, "VB Error :"
            Exit Sub
        End If
      Else
         MsgBox "DATABASE ERROR :" & CustDatabase.LastServerErr & _
         CustDatabase.LastServerErrText, vbCritical, "DATABASE Error :"
         CustDatabase.LastServerErrReset
         Exit Sub
      End If
  Else
     MsgBox "SESSION ERROR :" & CustSession.LastServerErr & CustSession.LastServerErrText, _
     vbCritical, "SESSION Error"
     CustSession.LastServerErrReset
     Exit Sub
  End If

End Sub
Private Sub CheckEmpty()
  
  'A SUBROUTINE FOR CHECKING THE EMPTY TEXT BOX
  If custNameText.Text = "" Then
     MsgBox "ERROR  ! CUSTOMER NAME IS BLANK .", vbInformation, "SAN PROJECT: "
     FlagEmpty = True
  ElseIf custAddText.Text = "" Then
     MsgBox "ERROR ! CUSTOMER ADDRESS IS BLANK .", vbInformation, "SAN PROJECT :"
     FlagEmpty = True
  ElseIf custPhNText.Text = "" Then
     MsgBox "ERROR ! CUSTOMER PHONE NO IS BLANK.", vbInformation, "SAN PROJECT :"
     FlagEmpty = True
  Else
     FlagEmpty = False
  End If
 
End Sub

Private Sub custAddText_GotFocus()

StatusBar1.Panels(3) = "Entry the customer address.."

End Sub

Private Sub custAddText_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 Then 'IF ENTER KEY IS PRESSED
   custPhNText.SetFocus
 End If
 
End Sub

Private Sub custAddText_LostFocus()
  
  StatusBar1.Panels(3) = ""

End Sub

Private Sub custNameText_GotFocus()
   
   StatusBar1.Panels(3) = "Entry the customer name .."

End Sub

Private Sub custNameText_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then  'IF ENTER KEY IS PRESSED
     custAddText.SetFocus
  End If
 
End Sub

Private Sub custNameText_LostFocus()
  
  StatusBar1.Panels(3) = " "

End Sub

Private Sub custPhNText_GotFocus()
 
 StatusBar1.Panels(3) = "Enter the customer Phone number.."

End Sub

Private Sub custPhNText_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then 'IF ENTER KEY IS PRESSED
     dayCombo.SetFocus
  End If

End Sub

Private Sub custPhNText_LostFocus()
  
  StatusBar1.Panels(3) = ""

End Sub

Private Sub dayCombo_GotFocus()
 
  StatusBar1.Panels(3) = "Select the day .."

End Sub

Private Sub dayCombo_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then 'IF ENTER KEY IS PRESSED
     monthCombo.SetFocus
  End If

End Sub

Private Sub dayCombo_LostFocus()

  StatusBar1.Panels(3) = ""

End Sub

Private Sub exitCmd_Click()
   
   If MsgBox("DO YOU WANT TO EXIT ", vbExclamation + vbYesNo, "EXIT:") = vbYes Then
      Unload Me
   End If
 
End Sub

Private Sub Form_Load()

  On Error GoTo ERRORHANDLER
 
  'CREATING SESSION, OPENING DATABASE AND CREATING DYNASET
  Set CustSession = CreateObject("oracleinprocserver.xorasession")
  Set CustDatabase = CustSession.OpenDatabase("jms", "hrishi/jms", &H4&)
  Set CustIdDyn = CustDatabase.CreateDynaset("select CUSTOMER_ID.NEXTVAL from dual", &H4&)

  custIdLabel.Caption = CustIdDyn.Fields(0).Value
  'CALLING A SUBROTINE FOR ADDING THE DATE IN THEIR COMBO'S
  Call FILLCOMBODATE(dayCombo, monthCombo, yearCombo)

ERRORHANDLER:  'CODING FOR ERROR HANDLER
   If CustSession.LastServerErr = 0 Then
      If CustDatabase.LastServerErr = 0 Then
         If Err.Number = 0 Then
       Else
          MsgBox "VB ERROR : " & Err.Number & " :: " & Err.Description, _
                 vbCritical, "VB Error :"
          Exit Sub
       End If
     Else
        MsgBox "DATABASE ERROR :" & CustDatabase.LastServerErr & _
               CustDatabase.LastServerErrText, vbCritical, "DATABASE Error :"
       CustDatabase.LastServerErrReset
       Exit Sub
     End If
   Else
      MsgBox "SESSION ERROR :" & CustSession.LastServerErr & _
              CustSession.LastServerErrText, vbCritical, "SESSION Error"
      CustSession.LastServerErrReset
      Exit Sub
   End If

End Sub

Private Sub Label1_Click()

End Sub

Private Sub monthCombo_GotFocus()

   StatusBar1.Panels(3) = "Select the month .."

End Sub

Private Sub monthCombo_KeyPress(KeyAscii As Integer)
   
   yearCombo.SetFocus

End Sub

Private Sub monthCombo_LostFocus()
   
   StatusBar1.Panels(3) = ""

End Sub

Private Sub OkCmd_Click()
 
 'CALL A SUBROUTINE TO FILLING THE DATE TO THEIR COMBO'S
 checkDate = VERIFY_DATE(dayCombo.Text, monthCombo.Text, yearCombo.Text)
 'CALL A SUBROUTINE FOR CHECKING EMPTY
 Call CheckEmpty
 
 If checkDate = False Then
    MsgBox "Error ! Invalid Date .", vbCritical, "DATE Errro ."
    Exit Sub
 ElseIf FlagEmpty = True Then
    Exit Sub
 End If

   dateStr = dayCombo.Text + "-" + monthCombo.Text + "-" + yearCombo.Text
    
   'CALLING A SUBROUTINE FOR SAVING THE DATA OF CUTOMER
   Call SaveCustData

End Sub

Private Sub yearcombo_GotFocus()
 
 StatusBar1.Panels(3) = "Select the year .."

End Sub

Private Sub yearcombo_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 Then
    OkCmd.SetFocus
 End If
 
End Sub

Private Sub yearCombo_LostFocus()
  
  StatusBar1.Panels(3) = ""

End Sub
