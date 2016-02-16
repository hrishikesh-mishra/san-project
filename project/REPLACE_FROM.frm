VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form REPLACE_FROM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REPLACE FROM"
   ClientHeight    =   8835
   ClientLeft      =   3180
   ClientTop       =   2055
   ClientWidth     =   9795
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
   ScaleHeight     =   8835
   ScaleWidth      =   9795
   Begin VB.ComboBox proNameCombo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7080
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox itemText 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox itemCombo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7320
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox rplrNameText 
      Height          =   375
      Left            =   6360
      TabIndex        =   20
      Text            =   " "
      Top             =   3000
      Width           =   3255
   End
   Begin VB.ComboBox rplrIdCombo 
      Height          =   360
      Left            =   2040
      TabIndex        =   17
      Text            =   " "
      ToolTipText     =   "Select the Replacer ID."
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton extCmd 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   5760
      TabIndex        =   16
      ToolTipText     =   "Exit."
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton cancelCmd 
      Caption         =   "&Canel"
      Height          =   495
      Left            =   4290
      TabIndex        =   15
      ToolTipText     =   "Clear."
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton okCmd 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   2850
      TabIndex        =   14
      ToolTipText     =   "Save the Data."
      Top             =   7560
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid rplGrid 
      Height          =   2775
      Left            =   240
      TabIndex        =   13
      ToolTipText     =   "Enter the replace product entry."
      Top             =   4080
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4895
      _Version        =   393216
      Rows            =   100
      Cols            =   4
   End
   Begin VB.OptionButton custRadBtn 
      Caption         =   "Customer"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      ToolTipText     =   "Select the Replacer Type."
      Top             =   2400
      Width           =   1575
   End
   Begin VB.OptionButton partyRadBtn 
      Caption         =   "Party"
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      ToolTipText     =   "Select the replacer Type."
      Top             =   2400
      Width           =   1335
   End
   Begin VB.ComboBox yearcombo 
      Height          =   360
      Left            =   8640
      TabIndex        =   6
      Text            =   " "
      ToolTipText     =   "Select the Year."
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ComboBox monthCombo 
      Height          =   360
      Left            =   6960
      TabIndex        =   5
      Text            =   " "
      ToolTipText     =   "Select the Month."
      Top             =   1560
      Width           =   1695
   End
   Begin VB.ComboBox dayCombo 
      Height          =   360
      Left            =   6000
      TabIndex        =   4
      Text            =   " "
      ToolTipText     =   "Select the Day."
      Top             =   1560
      Width           =   975
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   555
      Left            =   0
      Picture         =   "REPLACE_FROM.frx":0000
      ToolTipText     =   "San's Replace From Entry"
      Top             =   0
      Width           =   2970
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   0
      X2              =   9720
      Y1              =   8400
      Y2              =   8400
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   4200
      X2              =   4200
      Y1              =   1080
      Y2              =   2280
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   120
      X2              =   720
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label4 
      Caption         =   "Enter the information of replace product "
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   840
      TabIndex        =   24
      Top             =   3840
      Width           =   4215
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   9720
      X2              =   9720
      Y1              =   3960
      Y2              =   7200
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   3960
      Y2              =   7200
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   120
      X2              =   9720
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   120
      X2              =   9720
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   4920
      X2              =   9720
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   7200
      X2              =   7200
      Y1              =   7440
      Y2              =   8160
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   2760
      X2              =   2760
      Y1              =   7440
      Y2              =   8160
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   2760
      X2              =   7200
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   2760
      X2              =   7200
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   9720
      X2              =   9720
      Y1              =   2280
      Y2              =   3480
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   2280
      Y2              =   3480
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   4080
      X2              =   4080
      Y1              =   2880
      Y2              =   3480
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   120
      X2              =   9720
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   120
      X2              =   9720
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   120
      X2              =   9720
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label rplrIDlabel 
      Caption         =   "Replacer ID :-"
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label rplrNameLabel 
      Caption         =   "Replacer Name :-"
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   4320
      TabIndex        =   18
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label rplrTypLabel 
      Caption         =   "Replacer Type :-"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label SLN 
      Caption         =   "Replace No. :-"
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label rplSlnLabel 
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label dateLabel 
      Caption         =   "Date :-"
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   " Year"
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   8640
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   " Month"
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   " Day"
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label RPLLABEL 
      Alignment       =   2  'Center
      Caption         =   "REPLACE FROM ENTRY"
      BeginProperty Font 
         Name            =   "LINCOLN"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2393
      TabIndex        =   0
      Top             =   600
      Width           =   4935
   End
End
Attribute VB_Name = "REPLACE_FROM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************'
'**                                  **'
'** REPLACE FROM CUSTOMER ENTRY FORM **'
'**                                  **'
'**************************************'

'VARIABLE DECLARATION

Option Explicit

Dim rplSession       As OraSession
Dim rplDatabase      As OraDatabase
Dim rplSlnDyn        As OraDynaset
Dim CustIdDyn        As OraDynaset
Dim custNameDyn      As OraDynaset
Dim partyIdDyn       As OraDynaset
Dim partyNameDyn     As OraDynaset
Dim itemDyn          As OraDynaset
Dim proDyn           As OraDynaset
Dim rowCount         As Integer
Dim r                As Integer
Dim c                As Integer
Dim d                As Integer
Dim numOfRow         As Integer
Dim cateGo           As String
Dim dateStr          As String
Dim insertSql        As String
Dim typOfCust        As String
Dim S                As String
Dim checkDate        As Boolean
Dim chkBlk           As Boolean

Private Sub Customer_Id()

  rplrIdCombo.CLEAR
  Set CustIdDyn = rplDatabase.CreateDynaset("select CUST_ID,CUST_NAME from HRISHI.CUSTOMER_DETAIL", &H4&)
  While Not CustIdDyn.EOF
     rplrIdCombo.AddItem CustIdDyn.Fields("CUST_ID").Value
     rplrIdCombo.ListIndex = 0
     CustIdDyn.MoveNext
  Wend
 
End Sub
Private Sub Party_Id()
 
 rplrIdCombo.CLEAR
 Set partyIdDyn = rplDatabase.CreateDynaset("select PARTY_ID from HRISHI.PARTY_DETAIL", &H4&)
 While Not partyIdDyn.EOF
       rplrIdCombo.AddItem partyIdDyn.Fields("PARTY_ID")
       rplrIdCombo.ListIndex = 0
       partyIdDyn.MoveNext
 Wend
 
End Sub
Private Sub CLEAR()
  
  dayCombo.Text = DAY(Date)
  monthCombo.Text = MonthName(MONTH(Date))
  yearcombo.Text = YEAR(Date)
  
  For r = 1 To 99
     For c = 1 To 4
        rplGrid.TextMatrix(r, c) = ""
     Next
  Next
  rplGrid.Col = 1
  rplGrid.Row = 1
  rplGrid.RowSel = 1
  rplGrid.ColSel = 1
  rowCount = 1
  rplGrid.TextMatrix(1, 1) = rowCount
  rowCount = rowCount + 1
  itemText.Text = ""
  dayCombo.SetFocus
  
End Sub

Private Sub cancelCmd_Click()
 
 Call CLEAR
 
End Sub

Private Sub custRadBtn_Click()

  Call Customer_Id

End Sub

Private Sub custRadBtn_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
     rplrIdCombo.SetFocus
  End If

End Sub

Private Sub dayCombo_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
     monthCombo.SetFocus
  End If
  
End Sub

Private Sub extCmd_Click()
  
  If MsgBox("Do you want to exit ?", vbExclamation + vbYesNo, "Exit") = vbYes Then
     Unload Me
  End If
  
End Sub

Private Sub Form_Load()
  
  Call FILLCOMBODATE(dayCombo, monthCombo, yearcombo)
    
  On Error GoTo ERRORHANDLER
  S$ = "|Slno.|^            Item name              |^       Qty       |^                                     Description              "
  rplGrid.formatString = S$
  
  Set rplSession = CreateObject("oracleinprocserver.xorasession")
  Set rplDatabase = rplSession.OpenDatabase("jms", "hrishi/jms", &H0&)
  Set rplSlnDyn = rplDatabase.CreateDynaset("select REPLACE_Sln.nextval from dual", &H4&)
  rplSlnLabel.Caption = rplSlnDyn.Fields(0)
  custRadBtn.Value = True
  
  rowCount = 1
  rplGrid.Col = 1
  rplGrid.Row = 1
  rplGrid.TextMatrix(1, 1) = rowCount
  rowCount = rowCount + 1
  
ERRORHANDLER:
   If rplSession.LastServerErr = 0 Then
      If rplDatabase.LastServerErr = 0 Then
         If Err.Number = 0 Then
         Else
           MsgBox "VB ERROR:" & Err.Number & Err.Description, vbCritical, "VB Error:"
           Exit Sub
         End If
      Else
        MsgBox "DATABASE ERROR :" & rplDatabase.LastServerErr & rplDatabase.LastServerErrText, vbCritical, "DATABASE Error:"
        rplDatabase.LastServerErrReset
        Exit Sub
      End If
  Else
      MsgBox "SESSION ERROR :" & rplSession.LastServerErr & rplSession.LastServerErrText, vbCritical, "SESSION Error:"
      rplSession.LastServerErrReset
      Exit Sub
  End If
  
End Sub

Private Sub itemCombo_Click()
   
   cateGo = itemCombo.Text
   proNameCombo.CLEAR
   
   itemCombo.Visible = False
     
   Set proDyn = rplDatabase.CreateDynaset("select PRODUCT_NAME from HRISHI.PRODUCT_DETAIL where CATEGORY = '" & cateGo & "'", &H4&)
   
   While Not proDyn.EOF
      proNameCombo.AddItem proDyn.Fields("PRODUCT_NAME").Value
      proDyn.MoveNext
   Wend
   
   proNameCombo.Visible = True
   proNameCombo.Top = itemCombo.Top
   proNameCombo.Left = itemCombo.Left
   proNameCombo.Width = itemCombo.Width
   proNameCombo.ListIndex = 0
   
End Sub

Private Sub itemText_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
    If rplGrid.Col < 4 Then
       rplGrid.Col = rplGrid.Col + 1
       rplGrid.ColSel = rplGrid.Col
       If rplGrid.Col = 2 Then
           itemCombo.SetFocus
       Else
          itemText.SetFocus
       End If
       
   Else
     itemText.Visible = False
     rplGrid.Col = 1
     If rplGrid.Row = 99 Then
        rplGrid.Row = 1
     Else
        rplGrid.Row = rplGrid.Row + 1
     End If
     rplGrid.ColSel = rplGrid.Col
     rplGrid.RowSel = rplGrid.Row
     rplGrid.TextMatrix(rplGrid.Row, rplGrid.Col) = rowCount
     rowCount = rowCount + 1
     rplGrid.Col = 2
     rplGrid.ColSel = rplGrid.Col
     
   End If
 End If
  
End Sub

Private Sub monthCombo_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 Then
    yearcombo.SetFocus
 End If

End Sub

Private Sub OkCmd_Click()
  
  checkDate = VERIFY_DATE(Val(dayCombo.Text), monthCombo.Text, Val(yearcombo.Text))
  If checkDate = False Then
    MsgBox "Date Error ! Invalid date .", vbCritical, "DATE Error:"
    Exit Sub
  End If
  
  If rplrIdCombo.Text = "" Or rplrNameText.Text = "" Then
     MsgBox "Empty ! Replacer is Blank.", vbCritical, "Empty Error:"
     Exit Sub
  End If
  
 On Error GoTo ERRORHANDLER
    If okCmd.Caption = "&Ok" Then
       okCmd.Caption = "&Save"
       itemCombo.Enabled = False
       itemText.Enabled = False
       proNameCombo.Enabled = False
       itemCombo.CLEAR
       itemText.Text = ""
       proNameCombo.CLEAR
       chkBlk = False
       
       For r = 1 To 99
          For c = 1 To 4
             If rplGrid.TextMatrix(r, 1) = "" Then
                chkBlk = True
             End If
             If chkBlk Then
                rplGrid.TextMatrix(r, c) = ""
             End If
                   
          Next
          
      If rplGrid.TextMatrix(r, 2) = "" Or rplGrid.TextMatrix(r, 3) = "" Or rplGrid.TextMatrix(r, 4) = "" Then
         For d = 1 To 4
            rplGrid.TextMatrix(r, d) = ""
          Next
      End If
    Next
     cancelCmd.Enabled = False
     Exit Sub
  End If
   
 If okCmd.Caption = "&Save" Then
     itemCombo.Enabled = True
     itemText.Enabled = True
     proNameCombo.Enabled = True
     itemCombo.CLEAR
     numOfRow = 0
     chkBlk = True
     For r = 1 To rowCount
         If rplGrid.TextMatrix(r, 1) <> "" Then
           chkBlk = False
           numOfRow = numOfRow + 1
         End If
     Next
  
  If chkBlk = False Then
     If custRadBtn.Value Then
        typOfCust = "customer"
     Else
        typOfCust = "party"
     End If
     dateStr = dayCombo.Text + "-" + monthCombo.Text + "-" + yearcombo.Text
     insertSql = "insert into HRISHI.REPLACEMENT_DETAIL values (" & Val(rplSlnLabel.Caption) & ",'" & _
                  typOfCust & "'," & Val(rplrIdCombo.Text) & ",'" & UCase(rplrNameText.Text) & "','" & dateStr & "')"
                  
     rplDatabase.ExecuteSQL (insertSql)
     

     For r = 1 To numOfRow
       
       insertSql = "insert into HRISHI.REPLACED_ITEM_DETAIL values ( " & Val(rplSlnLabel.Caption) & ", '" & _
                    UCase(rplGrid.TextMatrix(r, 2)) & "'," & Val(rplGrid.TextMatrix(r, 3)) & ",'" & UCase(rplGrid.TextMatrix(r, 4)) & "')"
                   
        rplDatabase.ExecuteSQL (insertSql)
     Next
     
   End If
   
     If MsgBox("Data is Saved ! do you want Contine.", vbInformation + vbYesNo, "Continue :") = vbYes Then
        okCmd.Caption = "&Ok"
        cancelCmd.Enabled = True
        Set rplSlnDyn = rplDatabase.CreateDynaset("select REPLACE_Sln.nextval from dual", &H4&)
        rplSlnLabel.Caption = rplSlnDyn.Fields(0)
        Call CLEAR
     Else
        Unload Me
     End If
   Exit Sub
 End If

ERRORHANDLER:
  If rplSession.LastServerErr = 0 Then
     If rplDatabase.LastServerErr = 0 Then
        If Err.Number = 0 Then
        Else
          MsgBox " VB ERROR :" & Err.Number & Err.Description, vbCritical, "VB Error:"
          Exit Sub
        End If
      Else
        MsgBox "DATABASE ERROR :" & rplDatabase.LastServerErr & rplDatabase.LastServerErrText, vbCritical, "DATABASE Error:"
        rplDatabase.LastServerErrReset
        Exit Sub
      End If
  Else
    MsgBox " SESSION ERROR:" & rplSession.LastServerErr & rplSession.LastServerErrText, vbCritical, "SESSION Error:"
    rplSession.LastServerErrReset
    Exit Sub
 End If
 
End Sub

Private Sub partyRadBtn_Click()
 
 Call Party_Id

End Sub

Private Sub partyRadBtn_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
    rplrIdCombo.SetFocus
  End If
  
End Sub

Private Sub proNameCombo_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     rplGrid.Col = rplGrid.Col + 1
     rplGrid.ColSel = rplGrid.Col
  End If
  
End Sub

Private Sub rplGrid_EnterCell()

  If rplGrid.Col = 2 And rplGrid.Enabled Then
     itemCombo.Visible = True
     itemCombo.Top = rplGrid.Top + rplGrid.CellTop
     itemCombo.Left = rplGrid.Left + rplGrid.CellLeft
     itemCombo.Width = rplGrid.CellWidth
     itemCombo.CLEAR
     Set itemDyn = rplDatabase.CreateDynaset("select distinct (category) from HRISHI.PRODUCT_DETAIL", &H0&)
     
     While Not itemDyn.EOF
       itemCombo.AddItem itemDyn.Fields(0)
       itemDyn.MoveNext
     Wend
  Else
    If rplGrid.Col = 1 Then
      Exit Sub
    End If
      
    If itemText.Enabled Then
       proNameCombo.Visible = False
       itemText.Text = ""
       itemText.Top = rplGrid.Top + rplGrid.CellTop
       itemText.Left = rplGrid.Left + rplGrid.CellLeft
       itemText.Width = rplGrid.CellWidth
       itemText.Height = rplGrid.CellHeight
       itemText.Visible = True
       itemText.SetFocus
    End If
 End If
 
End Sub

Private Sub rplGrid_LeaveCell()

  If rplGrid.Col = 2 Then
     rplGrid.Text = proNameCombo.Text
     proNameCombo.Visible = False
  End If
 
  If rplGrid.Col = 3 And itemText.Enabled Then
    If itemText.Text = "" Then
        itemText.Text = "0.00"
    End If
    rplGrid.Text = itemText.Text
    itemText.Text = ""
    itemText.Visible = False
  End If
  If rplGrid.Col = 4 And itemText.Enabled Then
     If itemText.Text = "" Then
       itemText.Text = "---DEFECTIVE---"
  End If
      rplGrid.Text = itemText.Text
      itemText.Text = ""
      itemText.Visible = False
  End If
 
End Sub

Private Sub RPLLABEL_Click()

End Sub

Private Sub rplrIdCombo_Click()
 
 If rplrIdCombo.Text <> "" Then
      If custRadBtn.Value Then
         Set custNameDyn = rplDatabase.CreateDynaset("select CUST_NAME from HRISHI.CUSTOMER_DETAIL where CUST_ID= " & rplrIdCombo.Text & "", &H4&)
         rplrNameText.Text = custNameDyn.Fields("CUST_NAME").Value
      Else
        Set partyNameDyn = rplDatabase.CreateDynaset("select PARTY_NAME from HRISHI.PARTY_DETAIL where PARTY_ID= " & rplrIdCombo.Text & "", &H4&)
        rplrNameText.Text = partyNameDyn.Fields("PARTY_NAME").Value
      End If
   End If
   
End Sub

Private Sub rplrIdCombo_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
     rplGrid.SetFocus
  End If

End Sub

Private Sub yearcombo_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
     custRadBtn.SetFocus
  End If
 
End Sub
