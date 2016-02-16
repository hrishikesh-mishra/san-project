VERSION 5.00
Begin VB.Form VEN_UPDATE 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VENDOR FORM [UPDATE]"
   ClientHeight    =   7455
   ClientLeft      =   5205
   ClientTop       =   2775
   ClientWidth     =   9405
   FillColor       =   &H00404040&
   ForeColor       =   &H00400040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "VEN_UPDATE.frx":0000
   ScaleHeight     =   7455
   ScaleWidth      =   9405
   Begin VB.CommandButton cancelCmd 
      Height          =   615
      Left            =   5160
      Picture         =   "VEN_UPDATE.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton okCmd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      Picture         =   "VEN_UPDATE.frx":03AD
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6600
      Width           =   855
   End
   Begin VB.ComboBox yearCombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   5160
      TabIndex        =   23
      Top             =   6000
      Width           =   1095
   End
   Begin VB.ComboBox monthCombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   3360
      TabIndex        =   22
      Top             =   6000
      Width           =   1815
   End
   Begin VB.ComboBox dayCombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      ItemData        =   "VEN_UPDATE.frx":0734
      Left            =   2520
      List            =   "VEN_UPDATE.frx":0736
      TabIndex        =   21
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox delOfText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   2520
      TabIndex        =   19
      Top             =   5280
      Width           =   3855
   End
   Begin VB.TextBox venEIDText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   2520
      TabIndex        =   17
      Top             =   4560
      Width           =   5055
   End
   Begin VB.TextBox venPhNText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   2520
      TabIndex        =   15
      Top             =   3840
      Width           =   3855
   End
   Begin VB.TextBox venAddText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   3120
      Width           =   4575
   End
   Begin VB.TextBox venNameText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   2520
      Width           =   3615
   End
   Begin VB.ComboBox venIdcombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   2520
      TabIndex        =   9
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton DelCmd 
      BackColor       =   &H008080FF&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton editCmd 
      BackColor       =   &H0080FF80&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton viewCmd 
      BackColor       =   &H00FF8080&
      Caption         =   "&View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton mvNextCmd 
      Height          =   495
      Left            =   1200
      Picture         =   "VEN_UPDATE.frx":0738
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton mvLastCmd 
      BackColor       =   &H8000000A&
      Height          =   495
      Left            =   1800
      Picture         =   "VEN_UPDATE.frx":0B73
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton mvPreCmd 
      Height          =   495
      Left            =   600
      Picture         =   "VEN_UPDATE.frx":0FAB
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton mvFirstCmd 
      Height          =   495
      Left            =   0
      Picture         =   "VEN_UPDATE.frx":13C4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   615
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   9360
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   9360
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   9360
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   9360
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   555
      Left            =   0
      Picture         =   "VEN_UPDATE.frx":1801
      Top             =   0
      Width           =   2970
   End
   Begin VB.Label EDateLabel 
      Caption         =   "Entry Date:- "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Label delOfLabel 
      Caption         =   "Deler Of :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   120
      TabIndex        =   18
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label venEIDLabel 
      Caption         =   "Vendor  Email ID :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label venPhNLabel 
      Caption         =   "Vendor Phone No :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label venAddLabel 
      Caption         =   "Vendor Address :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label venNameLabel 
      Caption         =   "Vendor Name :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label venIdLabel 
      Caption         =   "Vendor ID:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label curWorkLabel 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label venUpdLael 
      Alignment       =   2  'Center
      Caption         =   "VENDOR UPDATE "
      BeginProperty Font 
         Name            =   "CARLIN"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   2880
      TabIndex        =   26
      Top             =   600
      Width           =   3135
   End
End
Attribute VB_Name = "VEN_UPDATE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************'
'**                    **'
'** VENDOR UPDATE FORM **'
'**                    **'
'************************'

'VARIBLE DECLEARATION
 
 Option Explicit
 Dim venUpdSession  As OraSession
 Dim venUpdDatabase As OraDatabase
 Dim venIdDyn       As OraDynaset
 Dim venDyn         As OraDynaset
 Dim i              As Integer
 Dim j              As Integer
 Dim K              As Integer
 Dim FlagEmpty      As Boolean
 Dim checkDate      As Boolean
 Dim updateStr      As String
 Dim delStr         As String
 Dim dateStr        As String
 Private Sub EnableFalse()
   
   venNameText.Enabled = False
   venAddText.Enabled = False
   venPhNText.Enabled = False
   venEIDText.Enabled = False
   delOfText.Enabled = False
   dayCombo.Enabled = False
   monthCombo.Enabled = False
   yearCombo.Enabled = False
 
 End Sub
 Private Sub EnableTrue()
   
   venNameText.Enabled = True
   venAddText.Enabled = True
   venPhNText.Enabled = True
   venEIDText.Enabled = True
   delOfText.Enabled = True
   dayCombo.Enabled = True
   monthCombo.Enabled = True
   yearCombo.Enabled = True
 
 End Sub
 Private Sub fillWithValues()
    
    Call FILLCOMBODATE(dayCombo, monthCombo, yearCombo) 'calling subroutine for filling the combo of date
       
    venIdCombo.Text = venDyn.Fields("vendor_id").Value
    venNameText.Text = venDyn.Fields("vendor_name").Value
    venAddText.Text = venDyn.Fields("vendor_add").Value
    venPhNText.Text = venDyn.Fields("vendor_pho").Value
    venEIDText.Text = venDyn.Fields("vendor_eid").Value
    delOfText.Text = venDyn.Fields("deler_of").Value
    dayCombo.Text = DAY(venDyn.Fields("entry_date").Value)
    monthCombo.Text = MonthName(MONTH(venDyn.Fields("entry_date").Value))
    yearCombo.Text = YEAR(venDyn.Fields("entry_date").Value)

 End Sub
 

Private Sub check_empty()
    
    If venNameText.Text = "" Then
      MsgBox "Error ! Vendor Name isn't present .", vbInformation, "Empty:"
      FlagEmpty = True
    ElseIf venAddText.Text = "" Then
      MsgBox "Error ! Vendor address isn't present .", vbInformation, "Empty:"
      FlagEmpty = True
    ElseIf venPhNText.Text = "" Then
      MsgBox "Error ! Vendor Phone No isn't present. ", vbInformation, "Empty:"
      FlagEmpty = True
    ElseIf venEIDText.Text = "" Then
      MsgBox "Error ! vendor Email Id isn't present. ", vbInformation, "Empty:"
      FlagEmpty = True
    ElseIf delOfText.Text = "" Then
      MsgBox "Error ! Deler of column is empty.", vbInformation, "Empty:"
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

Private Sub cancelCmd_Click()
    
    If MsgBox("DO YOU REALLY WANT TO EXIT", vbYesNo + vbExclamation, "EXIT") = vbYes Then
        Unload Me
    End If

End Sub

Private Sub delCmd_Click()

    If venIdCombo.Text = "" Then
        Exit Sub
    End If

    curWorkLabel.Caption = "DELETE"
    OkCmd.Enabled = False
    venIdCombo.Enabled = True
    Call EnableFalse  'Calling subroutine for enable false of Some Component

    On Error GoTo ERRORHANDLER

    delStr = "delete from VENDOR_DETAIL where vendor_id = " & Val(venIdCombo.Text) & " "
    If MsgBox("Warning ! The data will lost , CONTINUE ", vbYesNo + _
                                vbCritical, "Warning") = vbYes Then
        venUpdDatabase.ExecuteSQL (delStr)
    Else
        Call ViewCmd_Click
        Exit Sub
    End If
ERRORHANDLER:                            'CODING FOR ERRORHANDLER
   If venUpdSession.LastServerErr = 0 Then
      If venUpdDatabase.LastServerErr = 0 Then
         If Err.Number = 0 Then
          MsgBox "Sucess ! The deletion is made.", vbInformation, "Sucess"
          venDyn.Refresh
          Call fillWithValues
          Call ViewCmd_Click
       
         Else
           MsgBox "Vb Error : " & Err.Number & Err.Description
           Exit Sub
         End If
      Else
        MsgBox " Database  Error : " & venUpdDatabase.LastServerErr & venUpdDatabase.LastServerErrText
        venUpdDatabase.LastServerErrReset
        Exit Sub
      End If
   Else
     MsgBox "Session Error : " & venUpdSession.LastServerErr & venUpdSession.LastServerErrText
     venUpdSession.LastServerErrReset
     Exit Sub
   End If

End Sub

Private Sub editCmd_Click()

    If venIdCombo.Text = "" Then
        Exit Sub
    End If

    curWorkLabel.Caption = "EDIT"
    venIdCombo.Enabled = False
    OkCmd.Enabled = True

    Call EnableTrue       'Calling subroutine for enable True of Some Component
 
End Sub

Private Sub Form_Load()

    Call EnableFalse
    curWorkLabel.Caption = "VIEW"
    OkCmd.Enabled = False

    Set venUpdSession = CreateObject("oracleinprocserver.xorasession")
    Set venUpdDatabase = venUpdSession.OpenDatabase("jms", "hrishi/jms", &H0&)
    Set venIdDyn = venUpdDatabase.CreateDynaset("select vendor_id  from vendor_detail ", &H0&)
    
    While Not venIdDyn.EOF
        venIdCombo.AddItem venIdDyn.Fields(0)
        venIdDyn.MoveNext
    Wend
    Set venDyn = venUpdDatabase.CreateDynaset("Select *  from  vendor_detail", &H0&)
    If Not venDyn.EOF Then
        Call fillWithValues  'calling this subroutine to display data
    Else
          MsgBox "Nothing to display .. ", vbInformation, "DATABASE Info."
          mvFirstCmd.Enabled = False
          mvPreCmd.Enabled = False
          mvNextCmd.Enabled = False
          mvLastCmd.Enabled = False
   End If
 
End Sub

Private Sub mvFirstCmd_Click()
     
     venDyn.MoveFirst
     Call fillWithValues  'calling this subroutine to display data
 
 End Sub

Private Sub mvLastCmd_Click()
       
       venDyn.MoveLast
      Call fillWithValues   'calling this subroutine to display data

End Sub

Private Sub mvNextCmd_Click()

    venDyn.MoveNext
    If venDyn.EOF Then
        MsgBox "Nothing is in After...", vbInformation
        venDyn.MoveLast
    Else
        Call fillWithValues   'calling this subrountine to display data
    End If

End Sub

Private Sub mvPreCmd_Click()
 
    venDyn.MovePrevious
    If venDyn.BOF Then
        MsgBox "Nothing is in Before...", vbInformation
        venDyn.MoveFirst
    Else
        Call fillWithValues
    End If
      
End Sub

Private Sub OkCmd_Click()

    Call check_empty      'calling subroutine for checking blank fields
    checkDate = VERIFY_DATE(Val(dayCombo.Text), monthCombo.Text, Val(yearCombo.Text))
    If checkDate = False Then
        MsgBox "Error ! Invalid date ", vbCritical, "Error"
        Exit Sub
    ElseIf FlagEmpty Then
        Exit Sub
    End If
 
 On Error GoTo ERRORHANDLER
     dateStr = dayCombo.Text + "-" + monthCombo.Text + "-" + yearCombo.Text
     updateStr = " update VENDOR_DETAIL set VENDOR_NAME = '" & _
            UCase(venNameText.Text) & "',VENDOR_ADD= '" & _
            UCase(venAddText.Text) & "',VENDOR_PHO='" & _
            venPhNText.Text & "',VENDOR_EID='" & _
            LCase(venEIDText.Text) & "',DELER_OF='" & _
            UCase(delOfText.Text) & "',ENTRY_DATE= '" & _
            dateStr & "' WHERE VENDOR_ID= " & Val(venIdCombo.Text) & ""
    venUpdDatabase.ExecuteSQL (updateStr)


ERRORHANDLER:                        'CODING FOR ERRORHANDLER
   If venUpdSession.LastServerErr = 0 Then
      If venUpdDatabase.LastServerErr = 0 Then
         If Err.Number = 0 Then
          MsgBox "Sucess ! The updation is made.", vbInformation, "Sucess"
          venDyn.Refresh
          Call ViewCmd_Click
       
         Else
           MsgBox "Vb Error : " & Err.Number & Err.Description
           Exit Sub
         End If
      Else
        MsgBox " Database  Error : " & venUpdDatabase.LastServerErr & venUpdDatabase.LastServerErrText
        venUpdDatabase.LastServerErrReset
        Exit Sub
      End If
   Else
     MsgBox "Session Error : " & venUpdSession.LastServerErr & venUpdSession.LastServerErrText
     venUpdSession.LastServerErrReset
     Exit Sub
   End If

End Sub

Private Sub ViewCmd_Click()
    If venIdCombo.Text = "" Then
        Exit Sub
    End If

    curWorkLabel.Caption = "VIEW"
    OkCmd.Enabled = False
    venIdCombo.Enabled = True
    Call EnableFalse       'Calling subroutine for enable false of _
                            Some Component
                     
End Sub

