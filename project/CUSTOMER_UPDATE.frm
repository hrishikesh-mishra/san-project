VERSION 5.00
Begin VB.Form CUSTOMER_UPDATE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CUSTOMER FORM [UPDATE]"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9720
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
   ScaleHeight     =   7560
   ScaleWidth      =   9720
   Begin VB.ComboBox Daycombo 
      ForeColor       =   &H000040C0&
      Height          =   360
      Left            =   2520
      TabIndex        =   5
      ToolTipText     =   "Select the day."
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox ttlSaleText 
      Enabled         =   0   'False
      ForeColor       =   &H000040C0&
      Height          =   360
      Left            =   2520
      TabIndex        =   4
      Text            =   " "
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox custPhNText 
      ForeColor       =   &H000040C0&
      Height          =   360
      Left            =   2520
      TabIndex        =   3
      Text            =   " "
      ToolTipText     =   "Enter the Cutomer's Phone No."
      Top             =   3960
      Width           =   3975
   End
   Begin VB.TextBox custAddText 
      ForeColor       =   &H000040C0&
      Height          =   360
      Left            =   2520
      TabIndex        =   2
      ToolTipText     =   "Enter the Customer's Address."
      Top             =   3285
      Width           =   3255
   End
   Begin VB.TextBox custNameText 
      ForeColor       =   &H000040C0&
      Height          =   360
      Left            =   2520
      TabIndex        =   1
      ToolTipText     =   "Enter the Customer's name."
      Top             =   2655
      Width           =   3615
   End
   Begin VB.ComboBox custIdCombo 
      ForeColor       =   &H000040C0&
      Height          =   360
      Left            =   2520
      TabIndex        =   0
      ToolTipText     =   "Select customer ID."
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton ExtCmd 
      Height          =   615
      Left            =   5400
      Picture         =   "CUSTOMER_UPDATE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Exit from this."
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton okCmd 
      Height          =   615
      Left            =   3720
      Picture         =   "CUSTOMER_UPDATE.frx":006B
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6600
      Width           =   855
   End
   Begin VB.ComboBox yearCombo 
      ForeColor       =   &H000040C0&
      Height          =   360
      Left            =   5310
      TabIndex        =   7
      Text            =   " "
      ToolTipText     =   "Select the year."
      Top             =   5760
      Width           =   1455
   End
   Begin VB.ComboBox monthCombo 
      ForeColor       =   &H000040C0&
      Height          =   360
      Left            =   3615
      TabIndex        =   6
      Text            =   " "
      ToolTipText     =   "Select the Month."
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton mvFirstCmd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Picture         =   "CUSTOMER_UPDATE.frx":03F2
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Move First"
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton mvNextCmd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1110
      Picture         =   "CUSTOMER_UPDATE.frx":082F
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Move Next "
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton mvLastCmd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1605
      Picture         =   "CUSTOMER_UPDATE.frx":0C6A
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Move Last"
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton viewCmd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&View"
      Height          =   495
      Left            =   6675
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "View the store data."
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton editCmd 
      BackColor       =   &H00FFFF80&
      Caption         =   "&Edit"
      Height          =   495
      Left            =   7770
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Edit the store data."
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton delCmd 
      BackColor       =   &H008080FF&
      Caption         =   "&Delete"
      Height          =   495
      Left            =   8625
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Delete the store data."
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton mvPreCmd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   615
      Picture         =   "CUSTOMER_UPDATE.frx":10A2
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Move Previous."
      Top             =   1200
      Width           =   495
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   9720
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0080FF80&
      BorderWidth     =   2
      X1              =   6960
      X2              =   9720
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0080FF80&
      BorderWidth     =   2
      X1              =   6960
      X2              =   6960
      Y1              =   1920
      Y2              =   6360
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0080FF80&
      BorderWidth     =   2
      X1              =   0
      X2              =   9720
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0080FF80&
      BorderWidth     =   2
      X1              =   0
      X2              =   360
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080FF80&
      BorderWidth     =   2
      X1              =   360
      X2              =   360
      Y1              =   1920
      Y2              =   6360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   9720
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   0
      Picture         =   "CUSTOMER_UPDATE.frx":14BB
      ToolTipText     =   "San's Customer Update Form"
      Top             =   0
      Width           =   3000
   End
   Begin VB.Label Label3 
      Caption         =   "Year"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   5520
      TabIndex        =   27
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Month"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3840
      TabIndex        =   26
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Day"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2640
      TabIndex        =   25
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Entry Date"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   570
      TabIndex        =   24
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label ttlSale 
      Caption         =   "Total Sale"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   570
      TabIndex        =   23
      Top             =   4770
      Width           =   1575
   End
   Begin VB.Label custPhNLabel 
      Caption         =   "Customer Phone"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   570
      TabIndex        =   22
      Top             =   4035
      Width           =   1935
   End
   Begin VB.Label custAddLabel 
      Caption         =   "Customer Addres"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   570
      TabIndex        =   21
      Top             =   3285
      Width           =   1935
   End
   Begin VB.Label custNameLabel 
      Caption         =   "Customer Name"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   570
      TabIndex        =   20
      Top             =   2670
      Width           =   1815
   End
   Begin VB.Label custIdLabel 
      Caption         =   "Customer ID"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   570
      TabIndex        =   19
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label curWorkLabel 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "SPENCER"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2100
      TabIndex        =   18
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Label custUpLabel 
      Alignment       =   2  'Center
      Caption         =   "CUSTOMER UPDATE Form"
      BeginProperty Font 
         Name            =   "GENNIFER"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2880
      TabIndex        =   17
      Top             =   600
      Width           =   3735
   End
End
Attribute VB_Name = "CUSTOMER_UPDATE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**************************'
'**                      **'
'** CUSTOMER UPDATE FORM **'
'**                      **'
'**************************'

'VARIABLE DECLARATION

Option Explicit

Dim custUpDSession          As OraSession
Dim custUpDDatabase         As OraDatabase
Dim custSlnDyn              As OraDynaset
Dim custDyn                 As OraDynaset
Dim delSql                  As String
Dim dateStr                 As String
Dim custUpdSql              As String
Dim saleUpdSql              As String
Dim FlagEmpty               As Boolean
Dim checkDate               As Boolean
Private Sub displayValue()
     
     'A SUBROUTINE FOR DISPLAYING VALUE
      custIdCombo.Text = custDyn.Fields("CUST_ID").Value
      custNameText.Text = custDyn.Fields("CUST_NAME").Value
      custAddText.Text = custDyn.Fields("CUST_ADD").Value
      custPhNText.Text = custDyn.Fields("CUST_PHN").Value
      ttlSaleText.Text = custDyn.Fields("TOTAL_SALE").Value
      dayCombo.Text = DAY(custDyn.Fields("ENTRY_DATE").Value)
      monthCombo.Text = MonthName(MONTH(custDyn.Fields("ENTRY_DATE").Value))
      yearCombo.Text = YEAR(custDyn.Fields("ENTRY_DATE").Value)
        
 End Sub
 Private Sub EnableFalse()
   
   'A SUBROUNTINE FOR DOING THE ENABLE FALSE OF TEXT BOX
   custNameText.Enabled = False
   custAddText.Enabled = False
   custPhNText.Enabled = False
   dayCombo.Enabled = False
   monthCombo.Enabled = False
   yearCombo.Enabled = False
     
 End Sub
Private Sub check_empty()

      'CHECKING FOR EMPTY
     If custNameText.Text = "" Then
         MsgBox "Error ! Customer name isn't present.", vbInformation, "Empty:"
         FlagEmpty = True
     ElseIf custAddText.Text = "" Then
         MsgBox "Error ! Customer address isn't present.", vbInformation, "Empty:"
         FlagEmpty = True
     ElseIf custPhNText.Text = "" Then
         MsgBox "Error ! Customer Phone No. isn't present.", vbInformation, "Empty:"
         FlagEmpty = True
     Else
         FlagEmpty = False
     End If
     
End Sub

Private Sub EnableTrue()
   
   'A SUBROUNTINE FOR DOING THEN ENABLE TRUE OF TEXT BOX
   custNameText.Enabled = True
   custAddText.Enabled = True
   custPhNText.Enabled = True
   dayCombo.Enabled = True
   monthCombo.Enabled = True
   yearCombo.Enabled = True
   
End Sub


Private Sub custAddText_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 Then  'IF ENTER KEY IS PRESSED
    custPhNText.SetFocus
 End If

End Sub

Private Sub custIdCombo_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 Then  'IF ENTER KEY IS PRESSED
     custNameText.SetFocus
 End If

End Sub

Private Sub custNameText_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 Then  'IF ENTER KEY IS PRESSED
    custAddText.SetFocus
 End If

End Sub

Private Sub custPhNText_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 Then  'IF ENTER KEY IS PRESSED
   dayCombo.SetFocus
 End If

End Sub

Private Sub dayCombo_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 Then  'IF ENTER KEY IS PRESSED
    monthCombo.SetFocus
 End If

End Sub

Private Sub delCmd_Click()

  If custIdCombo.Text = "" Then
     Exit Sub
  End If
  
  curWorkLabel.Caption = "DELETE"
  curWorkLabel.ForeColor = vbRed
  OkCmd.Enabled = False
  custIdCombo.Enabled = True
  Call EnableFalse
   
  On Error GoTo ERRORHANDLER
  'SQL FOR DELETING CUSTOMER RECORD
  delSql = "delete from HRISHI.CUSTOMER_DETAIL where CUST_ID = " & Val(custIdCombo.Text) & ""
  'ASKING FOR CONFIRMATION
  If MsgBox("Warning ! The Data will lost ,CONTINUE", vbCritical + vbYesNo, "Warning:") = vbYes Then
    custUpDDatabase.ExecuteSQL (delSql)
  Else
    Call ViewCmd_Click
    Exit Sub
  End If
  
ERRORHANDLER: 'CODING FOR ERROR HANDLER
   If custUpDSession.LastServerErr = 0 Then
      If custUpDDatabase.LastServerErr = 0 Then
         If Err.Number = 0 Then
            MsgBox "Sucess ! The deletion is made. ", vbInformation, "Sucess"
            custDyn.Refresh
            Call displayValue
            Call ViewCmd_Click
         Else
            MsgBox "VB ERROR:" & Err.Number & Err.Description, vbCritical, "VB Error:"
         End If
      Else
        MsgBox "DATABASE ERROR:" & custUpDDatabase.LastServerErr & custUpDDatabase.LastServerErrText, _
                vbCritical, "DATABASE Error."
        custUpDDatabase.LastServerErrReset
      End If
   Else
     MsgBox "SESSION ERRRO:" & custUpDSession.LastServerErr & custUpDSession.LastServerErrText, _
     vbCritical, "SESSION ERROR:"
     custUpDSession.LastServerErrReset
  End If
  
End Sub

Private Sub editCmd_Click()
   
   If custIdCombo.Text = "" Then
      Exit Sub
   End If
 
   curWorkLabel.Caption = "EDIT"
   curWorkLabel.ForeColor = vbGreen
   custIdCombo.Enabled = False
   OkCmd.Enabled = True
   
   Call EnableTrue
  
End Sub

Private Sub extCmd_Click()
   
   If MsgBox("Do you want to exit ?", vbExclamation + vbYesNo, "Exit:") = vbYes Then
      Unload Me
   End If

End Sub

Private Sub Form_Load()
 
  'CALLING A SUBROUTINE FOR ADDING DATES TO THEIR COMBO'S
   Call FILLCOMBODATE(dayCombo, monthCombo, yearCombo)
   OkCmd.Enabled = False
   curWorkLabel.Caption = "VIEW"
   Call EnableFalse

   'CREATING SESSION, OPENING SESSION AND CREATING DYANASET
   Set custUpDSession = CreateObject("oracleinprocserver.xorasession")
   Set custUpDDatabase = custUpDSession.OpenDatabase("jms", "hrishi/jms", &H0&)
   Set custSlnDyn = custUpDDatabase.CreateDynaset("select CUST_ID FROM HRISHI.CUSTOMER_DETAIL ", &H4&)
   
   While Not custSlnDyn.EOF  ' ADDING CUSTOMER ID TO THE COMBO BOX
      custIdCombo.AddItem custSlnDyn.Fields(0)
      custSlnDyn.MoveNext
   Wend

   Set custDyn = custUpDDatabase.CreateDynaset("select * from HRISHI.CUSTOMER_DETAIL", &H4&)
   'CHECKING FOR ANY RECORDS ARE PRESENT IN DATABASE OR NOT
   If Not custDyn.EOF Then
      Call displayValue
   Else
         MsgBox "Nothing to display ..", vbInformation, "DATABASE Info."
         mvFirstCmd.Enabled = False
         mvPreCmd.Enabled = False
         mvNextCmd.Enabled = False
         mvLastCmd.Enabled = False
   End If

End Sub

Private Sub monthCombo_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then  'IF ENTER KEY IS PRESSED
     yearCombo.SetFocus
  End If

End Sub

Private Sub mvFirstCmd_Click()
    
    custDyn.MoveFirst
    Call displayValue

End Sub

Private Sub mvLastCmd_Click()
  
  custDyn.MoveLast
  Call displayValue

End Sub

Private Sub mvNextCmd_Click()
 
 custDyn.MoveNext
  If custDyn.EOF Then
    MsgBox "Nothing is in After ...", vbInformation, "DATABASE Info."
    custDyn.MoveLast
  Else
     Call displayValue
  End If
  
End Sub

Private Sub mvPreCmd_Click()
    
    custDyn.MovePrevious
    If custDyn.BOF Then
      MsgBox "Nothing is in Before...", vbInformation, "DATABASE Info."
      custDyn.MoveFirst
    Else
      Call displayValue
    End If
   
End Sub

Private Sub OkCmd_Click()
   
   'CALLING A SUBROUTINE FOR CHECKING EMPTY
   Call check_empty
   'CALLING A FUNCTION FOR FIND VALID DATE
   checkDate = VERIFY_DATE(Val(dayCombo.Text), monthCombo.Text, Val(yearCombo.Text))
   
   If checkDate = False Then
     MsgBox " Error ! Invalid date. ", vbCritical, "DATE Error."
     Exit Sub
   ElseIf FlagEmpty Then
     Exit Sub
   End If
   
   On Error GoTo ERRORHANDLER
   
    dateStr = dayCombo.Text + "-" + monthCombo.Text + "-" + yearCombo.Text
   'SQL FOR CUSTOMER DATA UPDATION
    custUpdSql = " update HRISHI.CUSTOMER_DETAIL set CUST_NAME= '" & _
                   UCase(custNameText.Text) & "', CUST_ADD='" & _
                   UCase(custAddText.Text) & "', CUST_PHN='" & _
                   UCase(custPhNText.Text) & "',ENTRY_DATE ='" & _
                   dateStr & "' where CUST_ID = " & Val(custIdCombo.Text) & ""
   saleUpdSql = " update HRISHI.SALE_DETAIL set CUST_NAME='" & _
                  UCase(custNameText.Text) & "' where TYPE_OF_CUST= 'customer' and  CUST_ID = " & Val(custIdCombo.Text) & ""
                  
   custUpDDatabase.ExecuteSQL (custUpdSql)
   custUpDDatabase.ExecuteSQL (saleUpdSql)
                  
ERRORHANDLER: 'CODING FOR ERROR HANDLER
    If custUpDSession.LastServerErr = 0 Then
       If custUpDDatabase.LastServerErr = 0 Then
          If Err.Number = 0 Then
              MsgBox "Sucess ! Updation is made.", vbInformation, "Sucess."
              custDyn.Refresh
              Call ViewCmd_Click
           Else
             MsgBox "VB ERROR ! " & Err.Number & Err.Description, vbCritical, "VB Error:"
           End If
       Else
         MsgBox "DATABASE ERROR: " & custUpDDatabase.LastServerErr & custUpDDatabase.LastServerErrText, vbCritical, "DATABASE Error:"
         custUpDDatabase.LastServerErrReset
       End If
    Else
       MsgBox "SESSION ERROR :" & custUpDSession.LastServerErr & custUpDSession.LastServerErrText, vbCritical, " SESSION Error:"
       custUpDSession.LastServerErrReset
    End If
    
       
End Sub

Private Sub okCmd_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then 'IF ENTER KEY IS PRESSED
     ViewCmd.SetFocus
  End If

End Sub

Private Sub ViewCmd_Click()

 If custIdCombo.Text = "" Then  'IF ENTER KEY IS PRESSED
     Exit Sub
 End If

  curWorkLabel.Caption = "VIEW"
  curWorkLabel.ForeColor = vbBlue
  OkCmd.Enabled = False
  custIdCombo.Enabled = True
  Call EnableFalse

End Sub

Private Sub yearcombo_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then  'IF ENTER KEY IS PRESSED
     OkCmd.SetFocus
  End If

End Sub
