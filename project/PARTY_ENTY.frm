VERSION 5.00
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "ORADC.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form PARTY_ENTRY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PARTY ENTRY"
   ClientHeight    =   9960
   ClientLeft      =   2670
   ClientTop       =   1365
   ClientWidth     =   9600
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
   ScaleHeight     =   9960
   ScaleWidth      =   9600
   Begin VB.CommandButton exitCmd 
      Caption         =   "&EXIT"
      Height          =   495
      Left            =   5640
      TabIndex        =   16
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton clearCmd 
      Caption         =   "&CLEAR"
      Height          =   495
      Left            =   4080
      TabIndex        =   15
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton okCmd 
      Caption         =   "&OK"
      Height          =   495
      Left            =   2520
      TabIndex        =   14
      Top             =   7560
      Width           =   1215
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "PARTY_ENTY.frx":0000
      Height          =   1455
      Left            =   -120
      OleObjectBlob   =   "PARTY_ENTY.frx":0019
      TabIndex        =   29
      ToolTipText     =   "Party Infromation."
      Top             =   8400
      Width           =   9495
   End
   Begin ORADCLibCtl.ORADC ORADCPARTY 
      Height          =   375
      Left            =   6120
      Top             =   4080
      Visible         =   0   'False
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   661
      _StockProps     =   207
      Caption         =   "PARTY ORDACE"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DatabaseName    =   "jms"
      Connect         =   "hrishi/jms"
      RecordSource    =   "select * from party_detail"
   End
   Begin VB.TextBox nameText 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      ToolTipText     =   "Enter the Name."
      Top             =   2655
      Width           =   4095
   End
   Begin VB.ComboBox prtyIdCombo 
      Height          =   360
      Left            =   2160
      TabIndex        =   3
      ToolTipText     =   "Select the ID."
      Top             =   2130
      Width           =   1575
   End
   Begin VB.TextBox streetText 
      Height          =   360
      Left            =   2160
      TabIndex        =   5
      Text            =   " "
      ToolTipText     =   "Enter Street."
      Top             =   3195
      Width           =   3855
   End
   Begin VB.TextBox cityText 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Text            =   " "
      ToolTipText     =   "Enter City."
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox stateText 
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      ToolTipText     =   "Enter state."
      Top             =   4260
      Width           =   2415
   End
   Begin VB.TextBox pinText 
      Height          =   375
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   8
      ToolTipText     =   "Enter PIN Code."
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox eidText 
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Text            =   " "
      ToolTipText     =   "Enter the Email Id"
      Top             =   5340
      Width           =   4575
   End
   Begin VB.TextBox phnText 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Text            =   " "
      ToolTipText     =   "Enter the Phone No."
      Top             =   5880
      Width           =   2895
   End
   Begin VB.TextBox ttlSaleText 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   27
      Text            =   "0.00 "
      Top             =   6420
      Width           =   2055
   End
   Begin VB.ComboBox dayCombo 
      Height          =   360
      Left            =   2160
      TabIndex        =   11
      Text            =   " "
      ToolTipText     =   "Select The Day."
      Top             =   6960
      Width           =   1095
   End
   Begin VB.ComboBox yearCombo 
      Height          =   360
      Left            =   4560
      TabIndex        =   13
      Text            =   " "
      ToolTipText     =   "Select The Year."
      Top             =   6960
      Width           =   1215
   End
   Begin VB.ComboBox monthCombo 
      Height          =   360
      Left            =   3240
      TabIndex        =   12
      Text            =   " "
      ToolTipText     =   "Select the Month."
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Frame partyFrame 
      Caption         =   "Select for the party"
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Select the Party type"
      Top             =   960
      Width           =   8415
      Begin VB.OptionButton OldPartyRadBtn 
         Caption         =   "Old Party"
         ForeColor       =   &H00004080&
         Height          =   375
         Left            =   5040
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton NewPartyRadBtn 
         Caption         =   "New Party"
         ForeColor       =   &H00004080&
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   0
      Picture         =   "PARTY_ENTY.frx":1764
      ToolTipText     =   "San's Party Entry"
      Top             =   0
      Width           =   3000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      X1              =   -120
      X2              =   9360
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      X1              =   0
      X2              =   9360
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      X1              =   -120
      X2              =   9480
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label partyenty 
      Alignment       =   2  'Center
      Caption         =   "PARTY ENTRY "
      BeginProperty Font 
         Name            =   "BrushScriptSWCondensed"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3000
      TabIndex        =   28
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label partyId 
      Caption         =   "Party ID"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   -120
      TabIndex        =   26
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label parytName 
      Caption         =   "Name :-"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   2805
      Width           =   1575
   End
   Begin VB.Label streetLabel 
      Caption         =   "Street :-"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3345
      Width           =   1575
   End
   Begin VB.Label citLabel 
      Caption         =   "City :-"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3870
      Width           =   1575
   End
   Begin VB.Label stateLabel 
      Caption         =   "State :-"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   4410
      Width           =   1575
   End
   Begin VB.Label pinLabel 
      Caption         =   "PIN Code. :-"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4950
      Width           =   1575
   End
   Begin VB.Label eidLabel 
      Caption         =   "Email-ID :-"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   5490
      Width           =   1575
   End
   Begin VB.Label phnLabel 
      Caption         =   "Phone No :-"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   6030
      Width           =   1575
   End
   Begin VB.Label ttlSaleLabel 
      Caption         =   "Total Sale :-"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   6570
      Width           =   1575
   End
   Begin VB.Label dateLabel 
      Caption         =   "Date :-"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   7110
      Width           =   1575
   End
End
Attribute VB_Name = "PARTY_ENTRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************'
'**                    **'
'**  PARTY ENTRY FORM  **'
'**                    **'
'************************'

'VARIABLE DECLERATION

Option Explicit

Dim prtySession     As OraSession
Dim prtyDatabase    As OraDatabase
Dim prtySlnDyn      As OraDynaset
Dim prtyIdDyn       As OraDynaset
Dim prtyDyn         As OraDynaset
Dim dateStr         As String
Dim insertStr       As String
Dim checkDate       As Boolean
Dim FlagEmpty       As Boolean
Dim prtySln         As Integer

Private Sub EmptyCheck()  'CHECKING FOR NULL TEXTBOX AND COMBOMOX
 
 If nameText.Text = "" Then
    MsgBox "ERROR ! Name is Empty. ", vbInformation, "EMPTY Error:"
    FlagEmpty = True
 ElseIf streetText.Text = "" Then
    MsgBox "ERROR ! Street is Empty.", vbInformation, "EMPTY Error:"
    FlagEmpty = True
 ElseIf cityText.Text = "" Then
    MsgBox "ERROR ! City is Empty.", vbInformation, "EMPTY Error:"
    FlagEmpty = True
 ElseIf stateText.Text = "" Then
    MsgBox "ERROR ! State is Empty.", vbInformation, "EMPTY Error:"
    FlagEmpty = True
 ElseIf pinText.Text = "" Then
    MsgBox "ERROR ! Pin is Empty .", vbInformation, "EMPTY Error:"
    FlagEmpty = True
 ElseIf phnText.Text = "" Then
    MsgBox "ERROR ! PhoneNo. is Empty .", vbInformation, "EMPTY Error:"
    FlagEmpty = True
 Else
   FlagEmpty = True  'IN CASE WHERE NO ONE IS BLANK
 End If
 
 End Sub
 
 Private Sub CLEAR()  'FOR CLEARING THE TEXT BOX FOR USER INPUT NEW INFORMATION
  
  nameText.Text = ""
  streetText.Text = ""
  cityText.Text = ""
  stateText.Text = ""
  pinText.Text = ""
  eidText.Text = ""
  phnText.Text = ""
  dayCombo.Text = DAY(Date) 'INSERTING THE CURRENT DAY IN DAYCOMBO BOX
  monthCombo.Text = MonthName(MONTH(Date)) 'INSERTING THE CURRENT MONTH NAME IN MONTHCOMBO
  yearCombo.Text = YEAR(Date)  'INSERTING THE CURRENT YEAR IN YEAR COMBO BOX
  
 End Sub
 
Private Sub EnabelFalse()
  
 nameText.Enabled = False   'DOING ENABLE FALSE OF ALL TEXT BOX  IN CASE WHEN OLD DATA IS BEING
 streetText.Enabled = False  ' DISPLAY
 cityText.Enabled = False
 stateText.Enabled = False
 pinText.Enabled = False
 eidText.Enabled = False
 phnText.Enabled = False
 dayCombo.Enabled = False
 monthCombo.Enabled = False
 yearCombo.Enabled = False
 
End Sub

Private Sub EnableTrue()
 
 nameText.Enabled = True   'DOING ENABLE TRUE OF ALL TEXT BOX IN CASE WHERE USER WANT TO
 streetText.Enabled = True 'INSERT NEW ONE DATA
 cityText.Enabled = True
 stateText.Enabled = True
 pinText.Enabled = True
 eidText.Enabled = True
 phnText.Enabled = True
 dayCombo.Enabled = True
 monthCombo.Enabled = True
 yearCombo.Enabled = True
End Sub
Private Sub displayValue()
   
   If prtyDyn.EOF = True Then  'CHECKING FOR END OF FILE IF IT BEACOME TRUE
   Exit Sub                    ' THEN NOTHING IS TO DISPLAY AND EXIT
   End If
   
    
  prtyIdCombo.Text = prtyDyn.Fields("PARTY_ID").Value
  nameText.Text = prtyDyn.Fields("PARTY_NAME").Value
  streetText.Text = prtyDyn.Fields("PARTY_ADD_STREET").Value
  cityText.Text = prtyDyn.Fields("PARTY_ADD_CITY").Value
  stateText.Text = prtyDyn.Fields("PARTY_ADD_STATE").Value
  pinText.Text = prtyDyn.Fields("PARTY_ADD_PIN").Value
  eidText.Text = prtyDyn.Fields("PARTY_EID").Value
  phnText.Text = prtyDyn.Fields("PARTY_PHN").Value
  ttlSaleText.Text = prtyDyn.Fields("PARTY_TOTAL_SALE").Value
  dayCombo.Text = DAY(prtyDyn.Fields("ENTRY_DATE").Value)
  monthCombo.Text = MonthName(MONTH(prtyDyn.Fields("ENTRY_DATE").Value))
  yearCombo.Text = YEAR(prtyDyn.Fields("ENTRY_DATE").Value)
  
  
End Sub


Private Sub cityText_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then             'IF KEYPRESS IS ENTER THEN MOVE CURSOR IN
    stateText.SetFocus              'STATE TEXT BOX
  End If

End Sub

Private Sub clearCmd_Click()

 If NewPartyRadBtn.Value = True Then
   Call CLEAR
 End If

End Sub

Private Sub dayCombo_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then          'IF KEYPRESS IS ENTER THEN MOVE CURSOR IN
     monthCombo.SetFocus           'MONTH COMBO BOX
  End If


End Sub

Private Sub eidText_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then          'IF KEYPRESS IS ENTER THEN MOVE CURSOR IN
     phnText.SetFocus              'PHONE TEXT BOX
  End If

End Sub

Private Sub exitCmd_Click()

  'ASKING QUESTION FOR EXIT IN THIS FORM IF YES THE EXIT ELSE NOT THAN STAY
   If MsgBox("DO YOU WANT TO EXIT ? ", vbExclamation + vbYesNo, "EXIT :") = vbYes Then
      Unload Me
   End If

End Sub

Private Sub Form_Load()

   On Error GoTo ERRORHANDLER
  
  Set prtySession = CreateObject("oracleinprocserver.xorasession")
  Set prtyDatabase = prtySession.OpenDatabase("jms", "hrishi/jms", &H4&)
  Set prtySlnDyn = prtyDatabase.CreateDynaset("select PARTY_SLN.NEXTVAL from dual ", &H4&)
  Set prtyIdDyn = prtyDatabase.CreateDynaset("select PARTY_ID from HRISHI.PARTY_DETAIL", &H4&)
  Set prtyDyn = prtyDatabase.CreateDynaset("select * from HRISHI.PARTY_DETAIL", &H4&)

  While Not prtyIdDyn.EOF      'ADDING PARTY ID IN PARTY ID COMBO BOX
      prtyIdCombo.AddItem prtyIdDyn.Fields("PARTY_ID").Value
      prtyIdDyn.MoveNext
  Wend

  prtyIdCombo.Enabled = False
  prtySln = prtySlnDyn.Fields(0) 'STORING THE NEW PARTY ID NUMBER IN A VARIABLE FOR LATER USE
  prtyIdCombo.Text = prtySlnDyn.Fields(0)

  Call FILLCOMBODATE(dayCombo, monthCombo, yearCombo) 'FILLING DATES IN COMBO BOX WITH THE HELP OF PUBLIC
                                                    'SUBROUTINE WHICH DEFINE IN THE MODULE FILLWITHDATE

ERRORHANDLER:     'CODING FOR ERROR HANDLER

If prtySession.LastServerErr = 0 Then
   If prtyDatabase.LastServerErr = 0 Then
     If Err.Number = 0 Then
     Else
        MsgBox "VB ERROR :" & Err.Number & " :: " & Err.Description, vbCritical, "VB Error :"
        End
     End If
   Else
     MsgBox "DATABASE ERROR " & prtyDatabase.LastServerErr & prtyDatabase.LastServerErrText, vbCritical, "DATABASE Error:"
     prtyDatabase.LastServerErrReset
     End
   End If
Else
   MsgBox "SESSION ERROR " & prtySession.LastServerErr & prtySession.LastServerErrText, vbCritical, "DATABASE Error:"
   prtySession.LastServerErrReset
   End
End If

   

End Sub

Private Sub monthCombo_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then        'IF KEYPRESS IS ENTER THEN MOVE THE CURSOR IN
    yearCombo.SetFocus         'YEAR COMBO BOX
  End If

End Sub

Private Sub nameText_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then      'IF KEYPRESS IS ENTER THEN MOVE THE CURSRO IN
     streetText.SetFocus      'THE STREET TEXT BOX
  End If

End Sub

Private Sub NewPartyRadBtn_Click()

  prtyIdCombo.Text = prtySln  'ADDING THE NEW PARTY ID NUMBER IN THE PARTY ID COMBO WHICH STORE IN VARIABLE
  prtyIdCombo.Enabled = False  'AND BECOME ITS ENABLE FALSE FOR ACCIDENTALLY CHANGE
  Call EnableTrue         'CALLING ENABLETRUE FOR USER INPUT
  Call CLEAR
  nameText.SetFocus
  
End Sub

Private Sub OkCmd_Click()
  
  If NewPartyRadBtn.Value = True Then
     
     checkDate = VERIFY_DATE(Val(dayCombo.Text), monthCombo.Text, YEAR(yearCombo.Text))
     If checkDate = False Then
        MsgBox "Error ! Invalid Date .", vbCritical, "DATE Error "
        Exit Sub
     ElseIf FlagEmpty Then
        Exit Sub
     End If


  On Error GoTo ERRORHANDLER
  
  dateStr = dayCombo.Text + "_" + monthCombo + "_" + yearCombo
  insertStr = "insert into HRISHI.PARTY_DETAIL VALUES ( " & Val(prtyIdCombo.Text) & _
            ",'" & UCase(nameText.Text) & "','" & UCase(streetText.Text) & "','" & UCase(cityText.Text) & _
            "','" & UCase(stateText.Text) & "','" & pinText.Text & "','" & LCase(eidText) & _
            "','" & phnText.TabIndex & "',0,'" & dateStr & "')"
            
  prtyDatabase.ExecuteSQL (insertStr)

ERRORHANDLER:

If prtySession.LastServerErr = 0 Then
  If prtyDatabase.LastServerErr = 0 Then
    If Err.Number = 0 Then
     SALE_ENTRY.IdLabel.Caption = prtyIdCombo.Text
     SALE_ENTRY.NameLabel.Caption = UCase(nameText.Text)
     SALE_ENTRY.typOfCust.Caption = "party"
      MsgBox "DATA IS SAVED ", vbInformation, "INFORMATION :"
      ORADCPARTY.Refresh
      Unload Me
     Else
      MsgBox "VB ERROR :" & Err.Number & " :: " & Err.Description, vbCritical, "VB Error:"
      Exit Sub
     End If
  Else
    MsgBox "DATABASE ERROR" & prtyDatabase.LastServerErr & prtyDatabase.LastServerErrText, vbCritical, "DATABASE Error :"
    prtyDatabase.LastServerErrReset
    Exit Sub
 End If
Else
  MsgBox " SESSION ERROR " & prtySession.LastServerErr & prtySession.LastServerErrText, vbCritical, "SESSION Error:"
  prtySession.LastServerErrReset
  Exit Sub
End If

Else
  SALE_ENTRY.IdLabel.Caption = prtyIdCombo.Text
  SALE_ENTRY.NameLabel.Caption = nameText.Text
  SALE_ENTRY.typOfCust.Caption = "party"
  Unload Me
End If

End Sub

Private Sub OldPartyRadBtn_Click()
  prtyIdCombo.Enabled = True
  Call displayValue
  Call EnabelFalse
End Sub

Private Sub phnText_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
     dayCombo.SetFocus
  End If

End Sub

Private Sub pinText_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     eidText.SetFocus
  End If

End Sub

Private Sub prtyIdCombo_Click()

   If OldPartyRadBtn.Value = True Then
      Set prtyDyn = prtyDatabase.CreateDynaset("select * from HRISHI.PARTY_DETAIL where party_id = " & prtyIdCombo.Text & "", &H4&)
      Call displayValue
   End If

End Sub

Private Sub prtyIdCombo_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     nameText.SetFocus
  End If

End Sub

Private Sub stateText_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
     pinText.SetFocus
  End If

End Sub

Private Sub streetText_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     cityText.SetFocus
  End If

End Sub

Private Sub yearcombo_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     okCmd.SetFocus
  End If

End Sub
