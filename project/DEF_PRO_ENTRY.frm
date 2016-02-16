VERSION 5.00
Begin VB.Form DEF_PRO_ENTRY 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DEFECTIVE PRODUCT [ENTRY]"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8655
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
   ScaleHeight     =   7320
   ScaleWidth      =   8655
   Begin VB.TextBox qtyText 
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      ToolTipText     =   "Enter the number of defected quantity."
      Top             =   4800
      Width           =   1215
   End
   Begin VB.ComboBox pronameCombo 
      ForeColor       =   &H00004000&
      Height          =   360
      Left            =   2160
      TabIndex        =   1
      ToolTipText     =   "Select the product name."
      Top             =   3960
      Width           =   4335
   End
   Begin VB.ComboBox catCombo 
      ForeColor       =   &H00004000&
      Height          =   360
      Left            =   2160
      TabIndex        =   0
      ToolTipText     =   "Select the category of product."
      Top             =   3240
      Width           =   2895
   End
   Begin VB.CommandButton okCmd 
      Height          =   495
      Left            =   2760
      Picture         =   "DEF_PRO_ENTRY.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton extCmd 
      Height          =   495
      Left            =   4320
      Picture         =   "DEF_PRO_ENTRY.frx":0387
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Exit from this."
      Top             =   6480
      Width           =   735
   End
   Begin VB.ComboBox dayCombo 
      ForeColor       =   &H00004000&
      Height          =   360
      Left            =   2160
      TabIndex        =   3
      Text            =   " "
      ToolTipText     =   "Select the Day."
      Top             =   5640
      Width           =   855
   End
   Begin VB.ComboBox yearCombo 
      ForeColor       =   &H00004000&
      Height          =   360
      Left            =   4590
      TabIndex        =   5
      Text            =   " "
      ToolTipText     =   "Select the year."
      Top             =   5640
      Width           =   975
   End
   Begin VB.ComboBox monthCombo 
      ForeColor       =   &H00004000&
      Height          =   360
      Left            =   3015
      TabIndex        =   4
      Text            =   " "
      ToolTipText     =   "Select the Month"
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton mvFirstCmd 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Picture         =   "DEF_PRO_ENTRY.frx":03F2
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Move First"
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton mPreCmd 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      Picture         =   "DEF_PRO_ENTRY.frx":0635
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Move Previous"
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton mvNextCmd 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   750
      Picture         =   "DEF_PRO_ENTRY.frx":0784
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Move Next"
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton mvLastCmd 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1125
      Picture         =   "DEF_PRO_ENTRY.frx":08D1
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Move Last"
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton delCmd 
      BackColor       =   &H80000004&
      Caption         =   " &Delete"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Delete the data."
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton editCmd 
      BackColor       =   &H80000004&
      Caption         =   " &Edit"
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Enter the store data."
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton addCmd 
      BackColor       =   &H80000004&
      Caption         =   " &Add"
      Height          =   375
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Add New one."
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton ViewCmd 
      BackColor       =   &H80000004&
      Caption         =   " &View"
      Height          =   375
      Left            =   5235
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "View the store data."
      Top             =   1560
      Width           =   855
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   2
      DrawMode        =   12  'Nop
      X1              =   6960
      X2              =   6960
      Y1              =   2160
      Y2              =   6240
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   2
      DrawMode        =   12  'Nop
      X1              =   120
      X2              =   120
      Y1              =   2160
      Y2              =   6240
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      DrawMode        =   12  'Nop
      X1              =   0
      X2              =   8640
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      DrawMode        =   12  'Nop
      X1              =   0
      X2              =   8640
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      DrawMode        =   12  'Nop
      X1              =   0
      X2              =   8640
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   2
      DrawMode        =   12  'Nop
      X1              =   0
      X2              =   8640
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   0
      Picture         =   "DEF_PRO_ENTRY.frx":0B0D
      ToolTipText     =   "San's Defective Product Entry Form"
      Top             =   0
      Width           =   3000
   End
   Begin VB.Label defSln 
      AutoSize        =   -1  'True
      Caption         =   "Sl.No. "
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   240
      TabIndex        =   26
      Top             =   2400
      Width           =   705
   End
   Begin VB.Label proNameLabel 
      Caption         =   "Product Name :-"
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   240
      TabIndex        =   24
      Top             =   4020
      Width           =   1815
   End
   Begin VB.Label qtyLabel 
      Caption         =   "Quantity :-"
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   240
      TabIndex        =   23
      Top             =   4830
      Width           =   1455
   End
   Begin VB.Label eDaeLabel 
      Caption         =   "Entry Date :-"
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   240
      TabIndex        =   22
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label slnLabel 
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
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   2160
      TabIndex        =   21
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Year"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   4560
      TabIndex        =   20
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Month"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3240
      TabIndex        =   19
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Day"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2160
      TabIndex        =   18
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label curWorkLabel 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "CURTIS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1500
      TabIndex        =   13
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Label defLabel 
      Alignment       =   2  'Center
      Caption         =   " DEFECTIVE PRODUCT ENTRY"
      BeginProperty Font 
         Name            =   "CLINTON"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   495
      Left            =   2040
      TabIndex        =   12
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label CatLabel 
      Caption         =   " Category :-"
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   240
      TabIndex        =   25
      Top             =   3210
      Width           =   1215
   End
End
Attribute VB_Name = "DEF_PRO_ENTRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************'
'**                                    **'
'** DEFECTIVE PRODUCT ENTRY AND UPDATE **'
'**                                    **'
'****************************************'

'VARIABLE DECLARATION
Option Explicit

Dim dfcSession       As OraSession
Dim dfcDatabase      As OraDatabase
Dim dfcSlnDyn        As OraDynaset
Dim cateGoDyn        As OraDynaset
Dim proNameDyn       As OraDynaset
Dim dfcDyn           As OraDynaset
Dim FlagEmpty        As Boolean
Dim checkDate        As Boolean
Dim insertSql        As String
Dim dateStr          As String
Dim updSql           As String
Dim delSql           As String

Private Sub displayValue()
   
   'SUBROUTINE FOR DISPLAYING VALUE OF PARTICULAR RECORD
   
   If dfcDyn.EOF Or dfcDyn.BOF Then
      Exit Sub
   End If
  
  slnLabel.Caption = dfcDyn.Fields(0)
  proNameCombo.Text = dfcDyn.Fields(1)
  catCombo.Text = dfcDyn.Fields(2)
  qtyText.Text = dfcDyn.Fields(3)
  dayCombo.Text = DAY(dfcDyn.Fields(4))
  monthCombo.Text = MonthName(MONTH(dfcDyn.Fields(4)))
  yearCombo.Text = YEAR(dfcDyn.Fields(4))
    
End Sub

Private Sub addCmd_Click()
 
 Call EnableTrue
 dayCombo.Text = DAY(Date)
 monthCombo.Text = MonthName(MONTH(Date))
 yearCombo.Text = YEAR(Date)

 curWorkLabel.Caption = "ADD"
 curWorkLabel.ForeColor = vbGreen
 OkCmd.Enabled = True
 Set dfcSlnDyn = dfcDatabase.CreateDynaset("select DEFCT_SLN.NEXTVAL FROM DUAL", &H4&)
 slnLabel.Caption = dfcSlnDyn.Fields(0)
 qtyText.Text = ""
 
End Sub

Private Sub catCombo_Click()
 
   If catCombo.Text = "" Then
      Exit Sub
   End If
 
 proNameCombo.CLEAR
 proNameCombo.Text = ""
 Set proNameDyn = dfcDatabase.CreateDynaset("select  PRODUCT_NAME from HRISHI.PRODUCT_DETAIL where category='" & catCombo.Text & "'", &H0&)
 
 While Not proNameDyn.EOF 'ADDING ALL PRODUCT NAME OF PARTICULAR CATEGORY TO THE PRODUCT COMBO
     proNameCombo.AddItem proNameDyn.Fields(0)
     proNameDyn.MoveNext
     proNameCombo.ListIndex = 0
 Wend
 
End Sub

Private Sub catCombo_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then
      proNameCombo.SetFocus
   End If

End Sub

Private Sub dayCombo_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then
      monthCombo.SetFocus
   End If

End Sub

Private Sub delCmd_Click()

  curWorkLabel.Caption = "DELETE"
  curWorkLabel.ForeColor = vbRed
  Call EnableFalse
  
  On Error GoTo ERRORHANDLER
  'SQL COMMAND FOR DELETING
  delSql = "delete from HRISHI.DEFECTIVE_PRO_DETAIL where DEFSLN=" & Val(slnLabel.Caption) & ""

  If MsgBox("The data will parmanentally lost." & vbCrLf & "Continue ?", vbCritical + vbYesNo, "Conformation.") = vbYes Then
      dfcDatabase.ExecuteSQL (delSql)
      MsgBox "Sucess ! The deletion is made.", vbInformation, "Sucess:"
      dfcDyn.Refresh
      Exit Sub
  Else
      Call ViewCmd_Click
      Exit Sub
  End If


ERRORHANDLER:
 If dfcSession.LastServerErr = 0 Then
     If dfcDatabase.LastServerErr = 0 Then
        If Err.Number = 0 Then
        Else
          MsgBox "VB ERROR :" & vbCrLf & Err.Number & Err.Description, vbCritical, "VB Error:"
          Exit Sub
        End If
     Else
        MsgBox "DATABASE ERROR:" & vbCrLf & dfcDatabase.LastServerErr & dfcDatabase.LastServerErrText, vbCritical, "DATABASE Error:"
        dfcDatabase.LastServerErrReset
        Exit Sub
     End If
  Else
     MsgBox "SESSION ERROR:" & vbCrLf & dfcSession.LastServerErr & dfcSession.LastServerErrText, vbCritical, "SESSION Error:"
     dfcSession.LastServerErrReset
     Exit Sub
  End If
       
End Sub

Private Sub editCmd_Click()

   curWorkLabel.Caption = "EDIT"
   curWorkLabel.ForeColor = vbYellow
   Call EnableTrue
   OkCmd.Enabled = True

End Sub

Private Sub extCmd_Click()
   
   If MsgBox("Do you want to exit ? ", vbExclamation + vbYesNo, "Exit:") = vbYes Then
       Unload Me
   End If
   
End Sub
Private Sub EnableFalse()
  
  catCombo.Enabled = False
  proNameCombo.Enabled = False
  qtyText.Enabled = False
  dayCombo.Enabled = False
  monthCombo.Enabled = False
  yearCombo.Enabled = False
  
End Sub

Private Sub EnableTrue()
 
 catCombo.Enabled = True
 proNameCombo.Enabled = True
 qtyText.Enabled = True
 dayCombo.Enabled = True
 monthCombo.Enabled = True
 yearCombo.Enabled = True

 End Sub
 
 Private Sub checkForBlank()
   
   'SUBROUTINE FOR CHECKING BLANK
   If catCombo.Text = "" Then
      MsgBox "Category isn't present. ", vbInformation, "Blank Error:"
      FlagEmpty = True
   ElseIf proNameCombo.Text = "" Then
     MsgBox "Product name isn't preset.", vbInformation, "Blank Error:"
     FlagEmpty = True
   ElseIf qtyText.Text = "" Then
     MsgBox "Quantity isn't present.", vbInformation, "Blank Error:"
     FlagEmpty = True
   Else
     FlagEmpty = False
   End If
 
 End Sub
Private Sub Form_Load()
 
 Call FILLCOMBODATE(dayCombo, monthCombo, yearCombo)
 Call EnableFalse
 
 On Error GoTo ERRORHANDLER
 
  'CREATING SESSION, OPENING DATA AND CREATING DYNASETS
  Set dfcSession = CreateObject("oracleinprocserveR.xorasession")
  Set dfcDatabase = dfcSession.OpenDatabase("jms", "hrishi/jms", &H0&)
  Set cateGoDyn = dfcDatabase.CreateDynaset("select distinct (category) from HRISHI.PRODUCT_DETAIL", &H4&)
  Set dfcDyn = dfcDatabase.CreateDynaset("select * from HRISHI.DEFECTIVE_PRO_DETAIL", &H0&)
  
  
  While Not cateGoDyn.EOF  'ADDING CATEGORY TO THE COMBOBOX
     catCombo.AddItem cateGoDyn.Fields(0)
     cateGoDyn.MoveNext
     catCombo.ListIndex = 0
  Wend
  
  Call displayValue
  OkCmd.Enabled = False
  curWorkLabel.Caption = "VIEW"
  curWorkLabel.ForeColor = vbBlue
  


ERRORHANDLER: 'CODDING FOR ERROR DETECTION
   If dfcSession.LastServerErr = 0 Then
     If dfcDatabase.LastServerErr = 0 Then
        If Err.Number = 0 Then
        Else
          MsgBox "VB ERROR :" & vbCrLf & Err.Number & Err.Description, vbCritical, "VB Error:"
          Unload Me
        End If
     Else
        MsgBox "DATABASE ERROR:" & vbCrLf & dfcDatabase.LastServerErr & dfcDatabase.LastServerErrText, vbCritical, "DATABASE Error:"
        dfcDatabase.LastServerErrReset
        Unload Me
     End If
  Else
     MsgBox "SESSION ERROR:" & vbCrLf & dfcSession.LastServerErr & dfcSession.LastServerErrText, vbCritical, "SESSION Error:"
     dfcSession.LastServerErrReset
     Unload Me
  End If
   
End Sub

Private Sub monthCombo_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then  'THE ENTER KEY PRESSED
     yearCombo.SetFocus
  End If

End Sub

Private Sub mPreCmd_Click()
  
  If dfcDyn.EOF Then
     Exit Sub
  End If
 
  dfcDyn.MovePrevious 'DISPLAYING THE PREVIOUS RECORD
   If dfcDyn.BOF Then
      MsgBox "Nothing is in Before...", vbInformation, "Information:"
      dfcDyn.MoveFirst
  Else
     Call displayValue
  End If

End Sub

Private Sub mvFirstCmd_Click()

  dfcDyn.MoveFirst 'DISPLAYING THE FIRST RECORD
  Call displayValue

End Sub

Private Sub mvLastCmd_Click()

  dfcDyn.MoveLast 'DISPLAYING THE LAST RECORD
  Call displayValue

End Sub

Private Sub mvNextCmd_Click()

  If dfcDyn.EOF Then
     Exit Sub
  End If

  dfcDyn.MoveNext 'DISPLAYING THE NEXT RECORD
  If dfcDyn.EOF Then
      MsgBox "Nothing is int After..", vbInformation, "Infromation:"
      dfcDyn.MoveLast
  Else
      Call displayValue
  End If

End Sub

Private Sub OkCmd_Click()
  
  On Error GoTo ERRORHANDLER
    'CALLING SUBROUTINE FOR VERIFY DATE
    checkDate = VERIFY_DATE(Val(dayCombo.Text), monthCombo.Text, Val(yearCombo.Text))
    'CALLING SUBROUTINE FOR CHEKING BLANK
    Call checkForBlank
    
    If checkDate = False Then
       MsgBox "Error" & vbCrLf & "Invalid date .", vbCritical, "Date Error:"
       Exit Sub
    ElseIf FlagEmpty Then
       Exit Sub
    End If
    
    dateStr = dayCombo.Text + monthCombo.Text + "-" + yearCombo.Text
     
     If curWorkLabel.Caption = "ADD" Then
      'SQL FOR INSERTION
      insertSql = "insert into HRISHI.DEFECTIVE_PRO_DETAIL values (" & Val(slnLabel.Caption) & ",'" & _
               UCase(proNameCombo.Text) & "','" & UCase(catCombo.Text) & "', " & Val(qtyText.Text) & ",'" & dateStr & "')"
      'INSERTING THE RECORD
      dfcDatabase.ExecuteSQL (insertSql)
    
    MsgBox "Sucess ! your data is save", vbInformation, "Sucess:"
    slnLabel.Caption = ""
    OkCmd.Enabled = False
    ViewCmd.SetFocus
    dfcDyn.Refresh
    Call ViewCmd_Click
    
   End If
   
   
 If curWorkLabel.Caption = "EDIT" Then
   'SQL COMMAND FOR UPDATION
   updSql = " update HRISHI.DEFECTIVE_PRO_DETAIL set PRODUCT_NAME='" & UCase(proNameCombo.Text) & "',CATEGORY='" & _
   UCase(catCombo.Text) & "',QTY= " & Val(qtyText.Text) & ",E_DATE='" & dateStr & "' where DEFSLN = " & Val(slnLabel.Caption) & ""
                 
   'UPDATING RECORD
   dfcDatabase.ExecuteSQL (updSql)
   
   MsgBox "Sucess ! database is updated.", vbInformation, "Sucess:"
   slnLabel.Caption = ""
   OkCmd.Enabled = False
   ViewCmd.SetFocus
   dfcDyn.Refresh
   Call ViewCmd_Click
   
 End If
 
ERRORHANDLER:  'CODDING FOR ERROR HANDLER
  If dfcSession.LastServerErr = 0 Then
     If dfcDatabase.LastServerErr = 0 Then
        If Err.Number = 0 Then
        Else
          MsgBox "VB ERROR :" & vbCrLf & Err.Number & Err.Description, vbCritical, "VB Error:"
          Exit Sub
        End If
     Else
        MsgBox "DATABASE ERROR:" & vbCrLf & dfcDatabase.LastServerErr & dfcDatabase.LastServerErrText, vbCritical, "DATABASE Error:"
        dfcDatabase.LastServerErrReset
        Exit Sub
     End If
  Else
     MsgBox "SESSION ERROR:" & vbCrLf & dfcSession.LastServerErr & dfcSession.LastServerErrText, vbCritical, "SESSION Error:"
     dfcSession.LastServerErrReset
     Exit Sub
  End If
   
End Sub

Private Sub proNameCombo_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then   'THE ENTER KEY PRESSED
     qtyText.SetFocus
  End If

End Sub

Private Sub qtyText_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 Then   'THE ENTER KEY PRESSED
    dayCombo.SetFocus
 End If

End Sub

Private Sub ViewCmd_Click()

  curWorkLabel.Caption = "VIEW"
  curWorkLabel.ForeColor = vbBlue
  Call EnableFalse

End Sub

Private Sub yearcombo_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then  'THE ENTER KEY PRESSED
     OkCmd.SetFocus
  End If

End Sub
