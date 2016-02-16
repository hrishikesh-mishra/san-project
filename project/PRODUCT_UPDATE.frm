VERSION 5.00
Begin VB.Form PRODUCT_UPDATE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PRODUCT UPDATE"
   ClientHeight    =   5940
   ClientLeft      =   3360
   ClientTop       =   675
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "PRODUCT_UPDATE.frx":0000
   ScaleHeight     =   5940
   ScaleWidth      =   8235
   Begin VB.CommandButton searchCmd 
      Height          =   375
      Left            =   6600
      Picture         =   "PRODUCT_UPDATE.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton extCmd 
      Height          =   495
      Left            =   4560
      Picture         =   "PRODUCT_UPDATE.frx":096A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Exit."
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton okCmd 
      Height          =   495
      Left            =   3000
      Picture         =   "PRODUCT_UPDATE.frx":09D5
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5160
      Width           =   735
   End
   Begin VB.ComboBox proNameCombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      TabIndex        =   0
      ToolTipText     =   "Enter the product Name."
      Top             =   2400
      Width           =   3495
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
      Height          =   360
      Left            =   2760
      TabIndex        =   3
      ToolTipText     =   "Select the day."
      Top             =   4440
      Width           =   975
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
      Height          =   360
      Left            =   3720
      TabIndex        =   4
      ToolTipText     =   "Select the Month."
      Top             =   4440
      Width           =   1695
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
      Height          =   360
      Left            =   5400
      TabIndex        =   5
      ToolTipText     =   "Select the year."
      Top             =   4440
      Width           =   1215
   End
   Begin VB.ComboBox cateCombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      TabIndex        =   1
      ToolTipText     =   "Select the Category."
      Top             =   3000
      Width           =   2895
   End
   Begin VB.TextBox manufacNameText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      ToolTipText     =   "Enter the Manufacturer."
      Top             =   3720
      Width           =   2895
   End
   Begin VB.CommandButton delCmd 
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Delete the records."
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton editCmd 
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Edit the records."
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton viewCmd 
      BackColor       =   &H00E0E0E0&
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
      Height          =   375
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "View the records."
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton mvLastCmd 
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   1080
      Picture         =   "PRODUCT_UPDATE.frx":0D5C
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Move Last"
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton mvNextCmd 
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   720
      Picture         =   "PRODUCT_UPDATE.frx":0F98
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Move Next."
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton mPreCmd 
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      Picture         =   "PRODUCT_UPDATE.frx":10E5
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Move Previous."
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton mvFirstCmd 
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      Picture         =   "PRODUCT_UPDATE.frx":1234
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Move First"
      Top             =   1200
      Width           =   375
   End
   Begin VB.Line Line10 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   7800
      X2              =   7800
      Y1              =   1680
      Y2              =   5040
   End
   Begin VB.Line Line9 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   240
      X2              =   240
      Y1              =   1680
      Y2              =   5040
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   6480
      X2              =   6480
      Y1              =   2280
      Y2              =   2880
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   7560
      X2              =   7560
      Y1              =   2280
      Y2              =   2880
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   6480
      X2              =   7560
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   6480
      X2              =   7560
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   8280
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label3 
      Caption         =   "PRODUCT UPDATE FORM"
      BeginProperty Font 
         Name            =   "CURTIS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2280
      TabIndex        =   21
      Top             =   600
      Width           =   4095
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   0
      Picture         =   "PRODUCT_UPDATE.frx":1477
      ToolTipText     =   "San's Product Update Form"
      Top             =   0
      Width           =   3000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   8280
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   8280
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   8280
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label dateLabel 
      Caption         =   "Date :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Manufacture:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Category :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label proNameLabel 
      Caption         =   "Product Name :- "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Left            =   360
      TabIndex        =   16
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label curWorkLabel 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "NICOLE"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   1440
      TabIndex        =   14
      Top             =   1200
      Width           =   4215
   End
End
Attribute VB_Name = "PRODUCT_UPDATE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************'
'**                **'
'** PRODUCT UPDATE **'
'**                **'
'********************'

'VARIABLE DECLEARATION

Option Explicit
Dim proUpdSession  As OraSession
Dim proUpdDatabase As OraDatabase
Dim proUpdDyn      As OraDynaset
Dim proNameDyn     As OraDynaset
Dim proCatDyn      As OraDynaset
Dim stkUpdSession  As OraSession
Dim stkUpdDatabase As OraDatabase
Dim i              As Integer
Dim j              As Integer
Dim K              As Integer
Dim checkDate      As Boolean
Dim FlagEmpty      As Boolean
Dim dateStr        As String
Dim proUpdStr      As String
Dim proDelStr      As String
Dim proName        As String
Dim DelProStr      As String
Dim DelStkStr      As String
Dim stkUpdStr      As String

Private Sub EnableTrue()
   
   cateCombo.Enabled = True
   manufacNameText.Enabled = True
   dayCombo.Enabled = True
   monthCombo.Enabled = True
   yearCombo.Enabled = True

End Sub
Private Sub EnableFalse()
  
  cateCombo.Enabled = False
  manufacNameText.Enabled = False
  dayCombo.Enabled = False
  monthCombo.Enabled = False
  yearCombo.Enabled = False

End Sub
  
Private Sub FillWithValue()
  
  If proUpdDyn.EOF And proUpdDyn.BOF Then
     proNameCombo.Text = ""
     cateCombo.Text = ""
     manufacNameText.Text = ""
     dayCombo.Text = ""
     monthCombo.Text = ""
     yearCombo.Text = ""
     Exit Sub
  End If
   
  Call comboDate
  proNameCombo.Text = proUpdDyn.Fields("product_name").Value
  cateCombo.Text = proUpdDyn.Fields("Category").Value
  manufacNameText.Text = proUpdDyn.Fields("manuf").Value
  dayCombo.Text = DAY(proUpdDyn.Fields("e_date").Value)
  monthCombo.Text = MonthName(MONTH(proUpdDyn.Fields("e_date").Value))
  yearCombo.Text = YEAR(proUpdDyn.Fields("e_date"))

End Sub

 Private Sub comboDate()

  For i = 1 To 31         'ADDING DAYS INTO DDCOMBOBOX
    dayCombo.AddItem i
  Next

  For j = 1 To 12         'ADDING MONTHS INTO MMCOMBO BOX
   monthCombo.AddItem MonthName(j)
  Next

 For K = 1900 To 2100    'ADDING YEAR INTO YYCOMBO BOX
    yearCombo.AddItem K
 Next

  dayCombo.Text = DAY(Date)  'ADDING CURRENT DATE IN THE ABOVE COMBO BOX
  monthCombo.Text = MonthName(MONTH(Date))
  yearCombo.Text = YEAR(Date)

End Sub

Private Sub CheckForEmpty() 'This subroutine check for the empty of the _
                             textbox and the combobox
  
  If proNameCombo.Text = "" Then
      MsgBox ("Error ! Product Name is not preset"), vbCritical, "Error"
      FlagEmpty = True
  ElseIf cateCombo.Text = "" Then
      MsgBox ("Error ! Category is not present "), vbCritical, "Error"
      FlagEmpty = True
  ElseIf manufacNameText.Text = "" Then
      MsgBox ("Error ! Manufacture is not defined "), vbCritical, "Error"
      FlagEmpty = True
  Else
      FlagEmpty = False
  End If
  
End Sub


Private Sub delCmd_Click()

  If proNameCombo.Text = "" Then
     Exit Sub
  End If
  On Error GoTo ED

  OkCmd.Enabled = False
  curWorkLabel.Caption = "DELETE"
  Call EnableFalse

  DelProStr = "delete from HRISHI.PRODUCT_DETAIL where PRODUCT_NAME = '" & proNameCombo.Text & "'"
  DelStkStr = "delete from HRISHI.STOCK_DETAIL where PRODUCT_NAME = '" & proNameCombo.Text & "'"

  If MsgBox("The data will loss Do you want continue", vbExclamation + vbYesNo) = vbYes Then
     proUpdDatabase.ExecuteSQL (DelProStr)
     stkUpdDatabase.ExecuteSQL (DelStkStr)
  Else
    Exit Sub
  End If

ED:
If proUpdSession.LastServerErr = 0 Then
   If proUpdDatabase.LastServerErr = 0 Then
      If Err.Number = 0 Then
         MsgBox "Success ! Deletion is made ", vbInformation, "Success:"
         proUpdDyn.Refresh
         Call ViewCmd_Click
         proUpdDyn.MoveFirst
       Else
         MsgBox "VB ERROR : " & Err.Number & " :: " & Err.Description, vbCritical, "VB Error :"
         Call ViewCmd_Click
         Exit Sub
       End If
   Else
       MsgBox "DATABASE ERROR :" & proUpdDatabase.LastServerErr & proUpdDatabase.LastServerErrText, vbCritical, "DATABASE Error"
       proUpdDatabase.LastServerErrReset
       Call ViewCmd_Click
       Exit Sub
   End If
Else
   MsgBox "SESSION ERROR :" & proUpdSession.LastServerErr & proUpdSession.LastServerErrText, vbCritical, "DATABASE Error:"
    proUpdSession.LastServerErrReset
    Call ViewCmd_Click
    Exit Sub
 End If
 

End Sub

Private Sub editCmd_Click()

  If proNameCombo.Text = "" Then
     Exit Sub
  End If

  proName = proNameCombo.Text
  curWorkLabel.Caption = "EDIT"
  OkCmd.Enabled = True
  Call EnableTrue

End Sub

Private Sub extCmd_Click()

  If MsgBox("DO YOU REALLY WANT TO EXIT", vbYesNo + vbExclamation, "EXIT") = vbYes Then
     Unload Me
  End If

End Sub

Private Sub Form_Load()

  curWorkLabel.Caption = "VIEW"
  OkCmd.Enabled = False

  Call EnableFalse
  On Error GoTo ERRORHANDLER
  
  Set proUpdSession = CreateObject("oracleinprocserver.xorasession")
  Set proUpdDatabase = proUpdSession.OpenDatabase("jms", "hrishi/jms", &H4&)
  Set proUpdDyn = proUpdDatabase.CreateDynaset("Select * from hrishi.product_detail", &H0&)
  Set proNameDyn = proUpdDatabase.CreateDynaset("select product_name from hrishi.product_detail", &H4&)
  Set proCatDyn = proUpdDatabase.CreateDynaset("Select distinct category from hrishi.product_detail", &H4&)
  Set stkUpdSession = CreateObject("oracleinprocserver.xorasession")
  Set stkUpdDatabase = stkUpdSession.OpenDatabase("jms", "hrishi/jms", &H4&)


  If proUpdDyn.EOF Then
     MsgBox "There is nothing to Update.", vbInformation, "Error"
     Exit Sub
  End If
 
 Call FillWithValue
 
 While Not proNameDyn.EOF
    proNameCombo.AddItem proNameDyn.Fields("product_name").Value
    proNameDyn.MoveNext
    proNameCombo.ListIndex = 0
 Wend
 
 While Not proCatDyn.EOF
   cateCombo.AddItem proCatDyn.Fields("Category").Value
   proCatDyn.MoveNext
   cateCombo.ListIndex = 0
 Wend
  
 
ERRORHANDLER:
  If proUpdSession.LastServerErr = 0 Then
     If proUpdDatabase.LastServerErr = 0 Then
        If Err.Number = 0 Then
        Else
           MsgBox "VB Error : " & Err.Number & " :: " & Err.Description, vbCritical, "Error: "
           End
        End If
     Else
        MsgBox "Database Error : " & proUpdDatabase.LastServerErr & proUpdDatabase.LastServerErrText, vbCritical, "Error"
        proUpdDatabase.LastServerErrReset
        End
     End If
 Else
     MsgBox "Session Error :" & proUpdSession.LastServerErr & proUpdSession.LastServerErrText, vbCritical, "Error"
     proUpdSession.LastServerErrReset
     End
 End If
 
End Sub

Private Sub mPreCmd_Click()

 If proUpdDyn.EOF Then
    Exit Sub
 End If

  proUpdDyn.MovePrevious
   If proUpdDyn.BOF Then
      MsgBox "Nothing is in Before...", vbInformation, "Information :"
      proUpdDyn.MoveFirst
   Else
     Call FillWithValue
   End If
 
End Sub

Private Sub mvFirstCmd_Click()

  proUpdDyn.MoveFirst
  Call FillWithValue
  
End Sub

Private Sub mvLastCmd_Click()

   proUpdDyn.MoveLast
   Call FillWithValue
   
End Sub

Private Sub mvNextCmd_Click()
  
  If proUpdDyn.EOF Then
     Exit Sub
  End If

  proUpdDyn.MoveNext
  If proUpdDyn.EOF Then
     MsgBox "Nothing is in After...", vbInformation, "Information :"
     proUpdDyn.MoveLast
  Else
     Call FillWithValue
  End If

End Sub

Private Sub OkCmd_Click()
 
 Call CheckForEmpty
 checkDate = VERIFY_DATE(Val(dayCombo.Text), monthCombo.Text, Val(yearCombo.Text))
 
 If checkDate = False Then
    MsgBox "Error ! Invalid date. ", vbCritical, "Error"
    Exit Sub
 ElseIf FlagEmpty Then
    Exit Sub
 End If

  On Error GoTo ERRORHANDLER
  dateStr = dayCombo.Text + "-" + monthCombo.Text + "-" + yearCombo.Text
  proUpdStr = "update HRISHI.PRODUCT_DETAIL set PRODUCT_NAME ='" _
             & UCase(proNameCombo.Text) & "',CATEGORY='" _
             & UCase(cateCombo.Text) & "',MANUF='" _
             & UCase(manufacNameText.Text) & "',E_DATE= '" _
             & dateStr & "' where PRODUCT_NAME = '" _
             & proName & "'"
 
 
 
  stkUpdStr = "update  HRISHI.STOCK_DETAIL SET  PRODUCT_NAME ='" _
             & UCase(proNameCombo.Text) & "', CATEGORY ='" _
             & UCase(cateCombo.Text) & "' where PRODUCT_NAME ='" _
             & proName & "'"
            
             
             
  proUpdDatabase.ExecuteSQL (proUpdStr)
  stkUpdDatabase.ExecuteSQL (stkUpdStr)

ERRORHANDLER:
 proUpdDyn.Refresh
 If proUpdSession.LastServerErr = 0 Then
    If proUpdDatabase.LastServerErr = 0 Then
      If Err.Number = 0 Then
        MsgBox "Success ! Updatation is made ", vbInformation, "Success :"
     
        Call ViewCmd_Click
      Else
        MsgBox "VB Error  :" & Err.Number & " :: " & Err.Description, vbCritical, "VB Error:"
        Exit Sub
      End If
   Else
      MsgBox "DATABASE ERROR : " & proUpdDatabase.LastServerErr & proUpdDatabase.LastServerErrText, vbCritical, " DATABASE Error: "
      proUpdDatabase.LastServerErrReset
      Exit Sub
      
   End If
Else
   MsgBox "SESSION ERROR : " & proUpdSession.LastServerErr & proUpdSession.LastServerErrText, vbCritical, "SESSION Error :"
   proUpdSession.LastServerErrReset
   Exit Sub
End If
   
End Sub

Private Sub searchCmd_Click()

  Set proUpdDyn = proUpdDatabase.CreateDynaset("select * from hrishi.product_detail where product_name='" & proNameCombo.Text & "'", &H4&)
  If proUpdDyn.EOF Then
     MsgBox "The data isn't exist in the database ", vbInformation, "Database : "
     Set proUpdDyn = proUpdDatabase.CreateDynaset("select * from hrishi.product_detail ", &H4&)
     Call FillWithValue
  Else
     Call FillWithValue
  End If

End Sub

Private Sub ViewCmd_Click()
  
  If proNameCombo.Text = "" Then
     Exit Sub
  End If

  curWorkLabel = "VIEW"
  OkCmd.Enabled = False
  proNameCombo.Enabled = True
  Call FillWithValue
  Call EnableFalse

End Sub

