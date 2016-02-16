VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ON_SPOT_RLMNT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ON SPOT REPLACEMENT [ENTRY]"
   ClientHeight    =   10395
   ClientLeft      =   1560
   ClientTop       =   900
   ClientWidth     =   9585
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10395
   ScaleWidth      =   9585
   Begin VB.ComboBox proNameCombo2 
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
      Left            =   7560
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.ComboBox itemCombo2 
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
      Left            =   7560
      TabIndex        =   28
      Top             =   -120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox itemText2 
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
      Left            =   7920
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox itemCombo1 
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
      Left            =   7680
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox itemText1 
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
      Left            =   7680
      TabIndex        =   23
      Top             =   -240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox proNameCombo1 
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
      Left            =   7560
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton extCmd 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   5625
      TabIndex        =   10
      ToolTipText     =   "Exit from this."
      Top             =   9720
      Width           =   1095
   End
   Begin VB.CommandButton cancelCmd 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4305
      TabIndex        =   9
      ToolTipText     =   "Canel the Process."
      Top             =   9720
      Width           =   1095
   End
   Begin VB.CommandButton okCmd 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   2865
      TabIndex        =   8
      ToolTipText     =   "Save the data."
      Top             =   9720
      Width           =   1215
   End
   Begin VB.ComboBox dayCombo 
      Height          =   360
      Left            =   5520
      TabIndex        =   1
      Text            =   " "
      ToolTipText     =   "Select the Day."
      Top             =   1680
      Width           =   975
   End
   Begin VB.ComboBox monthCombo 
      Height          =   360
      Left            =   6480
      TabIndex        =   2
      Text            =   " "
      ToolTipText     =   "Select the Month"
      Top             =   1680
      Width           =   1695
   End
   Begin VB.ComboBox yearcombo 
      Height          =   360
      Left            =   8160
      TabIndex        =   3
      Text            =   " "
      ToolTipText     =   "Select The Year."
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Frame rplframe 
      Caption         =   "Enter the replacer information :-"
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Select the replacer information."
      Top             =   2280
      Width           =   9015
      Begin VB.OptionButton custRadBtn 
         Caption         =   "Customer"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton partyRadBtn 
         Caption         =   "Party"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox rplrIdCombo 
         Height          =   360
         Left            =   5280
         TabIndex        =   5
         Text            =   " "
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox rplrNameText 
         Height          =   375
         Left            =   5280
         TabIndex        =   19
         Text            =   " "
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Line Line18 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   0
         X2              =   120
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line17 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   0
         X2              =   0
         Y1              =   120
         Y2              =   1680
      End
      Begin VB.Line Line16 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   3480
         X2              =   9000
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line15 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   9000
         X2              =   9000
         Y1              =   120
         Y2              =   1680
      End
      Begin VB.Line Line14 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   0
         X2              =   9000
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label rplrIDlabel 
         Caption         =   "Replacer ID :-"
         Height          =   375
         Left            =   3240
         TabIndex        =   21
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label rplrNameLabel 
         Caption         =   "Replacer Name :-"
         Height          =   375
         Left            =   3240
         TabIndex        =   20
         Top             =   1080
         Width           =   1935
      End
   End
   Begin MSFlexGridLib.MSFlexGrid aoGrid 
      Height          =   2055
      Left            =   360
      TabIndex        =   7
      ToolTipText     =   "Enter the detail of Defective Product."
      Top             =   7320
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3625
      _Version        =   393216
      Rows            =   100
      Cols            =   4
      BackColor       =   16777215
      ForeColor       =   8520959
   End
   Begin MSFlexGridLib.MSFlexGrid ridGrid 
      Height          =   2055
      Left            =   360
      TabIndex        =   6
      ToolTipText     =   "Enter the replaced product entry"
      Top             =   4560
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3625
      _Version        =   393216
      Rows            =   100
      Cols            =   3
      BackColor       =   16777215
      ForeColor       =   16711808
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   0
      Picture         =   "ON_SPOT_RLMNT.frx":0000
      ToolTipText     =   "San's On Spot Replacement Entry"
      Top             =   120
      Width           =   3000
   End
   Begin VB.Line Line13 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   120
      X2              =   9240
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line12 
      X1              =   2400
      X2              =   2520
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line11 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   120
      X2              =   120
      Y1              =   4440
      Y2              =   6840
   End
   Begin VB.Line Line10 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   120
      X2              =   1080
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line9 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   9240
      X2              =   9240
      Y1              =   4440
      Y2              =   6840
   End
   Begin VB.Line Line8 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   3600
      X2              =   9240
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line7 
      X1              =   9240
      X2              =   9255
      Y1              =   6240
      Y2              =   6255
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   9240
      X2              =   9240
      Y1              =   7200
      Y2              =   9600
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   120
      X2              =   9240
      Y1              =   9600
      Y2              =   9600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   120
      X2              =   120
      Y1              =   7200
      Y2              =   9600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   120
      X2              =   960
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   2280
      X2              =   9240
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Label Label5 
      Caption         =   "Against of :-"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   960
      TabIndex        =   26
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Replaced item detail :-"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   1080
      TabIndex        =   25
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label rpl 
      Alignment       =   2  'Center
      Caption         =   "ON SPOT REPLACEMENT ENTRY"
      BeginProperty Font 
         Name            =   "KAROLYN"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   1545
      TabIndex        =   18
      Top             =   720
      Width           =   6495
   End
   Begin VB.Label Label1 
      Caption         =   " Day"
      Height          =   375
      Left            =   5520
      TabIndex        =   17
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   " Month"
      Height          =   375
      Left            =   6600
      TabIndex        =   16
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   " Year"
      Height          =   375
      Left            =   8160
      TabIndex        =   15
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label dateLabel 
      Caption         =   "Date :-"
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label slnLabel 
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
      Left            =   1680
      TabIndex        =   12
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   9600
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label slnnlabel 
      Caption         =   "Sln No. :-"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "ON_SPOT_RLMNT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************** *'
'**                           **'
'** ON SPOT REPLACEMENT ENTRY **'
'**                           **'
'*******************************'

'VARIABLE DECLARATION

Option Explicit

Dim osrSession   As OraSession
Dim osrDatabase  As OraDatabase
Dim slnDyn       As OraDynaset
Dim CustIdDyn    As OraDynaset
Dim partyIdDyn   As OraDynaset
Dim custNameDyn  As OraDynaset
Dim partyNameDyn As OraDynaset
Dim itemDyn      As OraDynaset
Dim proDyn       As OraDynaset
Dim s1           As String
Dim s2           As String
Dim cateGo       As String
Dim dateStr      As String
Dim insertSql    As String
Dim typOfCust    As String
Dim r1           As Integer
Dim r2           As Integer
Dim c1           As Integer
Dim c2           As Integer
Dim d1           As Integer
Dim d2           As Integer
Dim numOfRow1    As Integer
Dim numOfRow2    As Integer
Dim rowCount1    As Integer
Dim rowCount2    As Integer
Dim checkDate    As Boolean
Dim chkBlk1      As Boolean
Dim chkBlk2      As Boolean

Private Sub Customer_Id()
 
 rplrIdCombo.CLEAR
 Set CustIdDyn = osrDatabase.CreateDynaset("select CUST_ID,CUST_NAME from HRISHI.CUSTOMER_DETAIL", &H4&)
 
 While Not CustIdDyn.EOF 'ADDING CUSTOMER ID IN CUSTOMER ID COMBO
    rplrIdCombo.AddItem CustIdDyn.Fields("CUST_ID").Value
    rplrIdCombo.ListIndex = 0
    CustIdDyn.MoveNext
 Wend
  
End Sub
Private Sub Party_Id()
 
 rplrIdCombo.CLEAR
 Set partyIdDyn = osrDatabase.CreateDynaset("select PARTY_ID from HRISHI.PARTY_DETAIL ", &H4&)
 
 While Not partyIdDyn.EOF     'ADDING PARTY ID IN PARTY ID COMBO
       rplrIdCombo.AddItem partyIdDyn.Fields("PARTY_ID")
       rplrIdCombo.ListIndex = 0
       partyIdDyn.MoveNext
 Wend
 
End Sub

Private Sub CLEAR()
 
 'A SUBROUTINE WHICH CLEAR ALL THE ENTERED INFROMATION
 dayCombo.Text = DAY(Date)
 monthCombo.Text = MonthName(MONTH(Date))
 yearCombo.Text = YEAR(Date)
  
  For r1 = 1 To 99
     For c1 = 1 To 4
        aoGrid.TextMatrix(r1, c1) = ""
     Next
  Next

  For r2 = 1 To 99
     For c2 = 1 To 3
        ridGrid.TextMatrix(r2, c2) = ""
     Next
  Next

  ridGrid.Col = 1
  ridGrid.Row = 1
  ridGrid.ColSel = 1
  ridGrid.RowSel = 1
  rowCount2 = 1
  ridGrid.TextMatrix(1, 1) = rowCount2
  rowCount2 = rowCount2 + 1
  itemText1.Text = ""
  
  aoGrid.Col = 1
  aoGrid.Row = 1
  aoGrid.ColSel = 1
  aoGrid.RowSel = 1
  rowCount1 = 1
  aoGrid.TextMatrix(1, 1) = 1
  rowCount1 = rowCount1 + 1
  itemText2.Text = ""
  dayCombo.SetFocus
  
End Sub

Private Sub aoGrid_EnterCell()
 
   If aoGrid.Col = 2 And aoGrid.Enabled Then
       itemCombo1.Visible = True
       itemCombo1.Top = aoGrid.Top + aoGrid.CellTop
       itemCombo1.Left = aoGrid.Left + aoGrid.CellLeft
       itemCombo1.Width = aoGrid.CellWidth
       itemCombo1.CLEAR
   
     Set itemDyn = osrDatabase.CreateDynaset("select distinct (category) from HRISHI.PRODUCT_DETAIL", &H0&)
     
     While Not itemDyn.EOF
         itemCombo1.AddItem itemDyn.Fields(0)
         itemDyn.MoveNext
     Wend
   Else
     If aoGrid.Col = 1 Then
        Exit Sub
     End If
      
     If itemText1.Enabled Then
        proNameCombo1.Visible = False
        itemText1.Text = ""
        itemText1.Top = aoGrid.Top + aoGrid.CellTop
        itemText1.Left = aoGrid.Left + aoGrid.CellLeft
        itemText1.Width = aoGrid.CellWidth
        itemText1.Height = aoGrid.CellHeight
        itemText1.Visible = True
        itemText1.SetFocus
     End If
   End If
           
End Sub

Private Sub aoGrid_LeaveCell()
 
  If aoGrid.Col = 2 Then
     aoGrid.Text = proNameCombo1.Text
     proNameCombo1.Visible = False
  End If
 
  If aoGrid.Col = 3 And itemText1.Enabled Then
     If itemText1.Text = "" Then
        itemText1.Text = "0.00"
     End If
     aoGrid.Text = itemText1.Text
     itemText1.Text = ""
     itemText1.Visible = False
  End If
  
  If aoGrid.Col = 4 And itemText1.Enabled Then
     If itemText1.Text = "" Then
        itemText1.Text = "---DEFECTIVE---"
     End If
     aoGrid.Text = itemText1.Text
     itemText1.Text = ""
     itemText1.Visible = False
  End If
  
End Sub

Private Sub cancelCmd_Click()

  Call CLEAR

End Sub

Private Sub custRadBtn_Click()
 
  Call Customer_Id
 
End Sub

Private Sub dayCombo_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     monthCombo.SetFocus
  End If
 
End Sub

Private Sub extCmd_Click()
 
  If MsgBox("Do you want to exit.", vbExclamation + vbYesNo, "Exit:") = vbYes Then
     Unload Me
  End If
 
End Sub

Private Sub Form_Load()

  Call FILLCOMBODATE(dayCombo, monthCombo, yearCombo)
  s1$ = "| Slno. |^                                Item name                                         |^               Qty          "
  s2$ = "| Slno.  |^             Item name                                 |^     Qty           |^            Description  "
  ridGrid.formatString = s1$
  aoGrid.formatString = s2$

  Set osrSession = CreateObject("oracleinprocserver.xorasession")
  Set osrDatabase = osrSession.OpenDatabase("jms", "hrishi/jms", &H0&)
  Set slnDyn = osrDatabase.CreateDynaset("select OSPR_SLN.NEXTVAL from dual", &H4&)
  slnLabel.Caption = slnDyn.Fields(0)
  custRadBtn.Value = True

  rowCount1 = 1
  rowCount2 = 1
  ridGrid.Col = 1
  ridGrid.Row = 1
  aoGrid.Col = 1
  aoGrid.Row = 1
  ridGrid.TextMatrix(1, 1) = 1
  aoGrid.TextMatrix(1, 1) = 1
  rowCount1 = rowCount1 + 1
  rowCount2 = rowCount2 + 1

End Sub

Private Sub itemCombo1_Click()

  cateGo = itemCombo1.Text
  proNameCombo1.CLEAR
  itemCombo1.Visible = False
  
  Set proDyn = osrDatabase.CreateDynaset("select PRODUCT_NAME from HRISHI.PRODUCT_DETAIL where CATEGORY='" & cateGo & "'", &H4&)
  
  While Not proDyn.EOF
     proNameCombo1.AddItem proDyn.Fields(0)
     proDyn.MoveNext
  Wend

  proNameCombo1.Visible = True
  proNameCombo1.Top = itemCombo1.Top
  proNameCombo1.Left = itemCombo1.Left
  proNameCombo1.Width = itemCombo1.Width
  
  proNameCombo1.ListIndex = 0

End Sub

Private Sub itemCombo2_Click()
  
  cateGo = itemCombo2.Text
  proNameCombo2.CLEAR
  itemCombo2.Visible = False
  
  Set proDyn = osrDatabase.CreateDynaset("select PRODUCT_NAME from HRISHI.PRODUCT_DETAIL where CATEGORY = '" & cateGo & "'", &H4&)
  
  While Not proDyn.EOF
     proNameCombo2.AddItem proDyn.Fields("PRODUCT_NAME").Value
     proDyn.MoveNext
  Wend
   
  proNameCombo2.Visible = True
  proNameCombo2.Top = itemCombo2.Top
  proNameCombo2.Left = itemCombo2.Left
  proNameCombo2.Width = itemCombo2.Width
  proNameCombo2.ListIndex = 0

End Sub

Private Sub itemText1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If aoGrid.Col < 4 Then
        aoGrid.Col = aoGrid.Col + 1
        aoGrid.ColSel = aoGrid.Col
        If aoGrid.Col = 2 Then
           itemCombo1.SetFocus
        Else
           itemText1.SetFocus
        End If
     Else
        itemText1.Visible = False
        aoGrid.Col = 1
     If aoGrid.Col = 99 Then
        aoGrid.Row = 1
     Else
        aoGrid.Row = aoGrid.Row + 1
     End If
     aoGrid.ColSel = aoGrid.Col
     aoGrid.RowSel = aoGrid.Row
     aoGrid.TextMatrix(aoGrid.Row, aoGrid.Col) = rowCount1
     rowCount1 = rowCount1 + 1
     aoGrid.Col = 2
     aoGrid.ColSel = aoGrid.Col
    End If
  End If
   
End Sub

Private Sub itemText2_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
     If ridGrid.Col < 3 Then
        ridGrid.Col = ridGrid.Col + 1
        ridGrid.ColSel = ridGrid.Col
        If ridGrid.Col = 2 Then
           itemCombo2.SetFocus
        Else
           itemText2.SetFocus
        End If
     Else
        itemText2.Visible = False
        ridGrid.Col = 1
        If ridGrid.Row = 99 Then
           ridGrid.Row = 1
       Else
          ridGrid.Row = ridGrid.Row + 1
       End If
          ridGrid.RowSel = ridGrid.Row
          ridGrid.ColSel = ridGrid.Col
          ridGrid.TextMatrix(ridGrid.Row, ridGrid.Col) = rowCount2
          rowCount2 = rowCount2 + 1
          ridGrid.Col = 2
          ridGrid.ColSel = ridGrid.Col
          
        
      End If
   End If
         
End Sub

Private Sub monthCombo_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
     yearCombo.SetFocus
  End If

End Sub

Private Sub OkCmd_Click()
   
   checkDate = VERIFY_DATE(Val(dayCombo.Text), monthCombo.Text, Val(yearCombo.Text))

  If checkDate = False Then
      MsgBox "Error ! Invaid date .", vbCritical, "DATE Error:"
      Exit Sub
  End If
 
  If rplrIdCombo.Text = "" Or rplrNameText.Text = "" Then
     MsgBox "Empty ! Replacer is Blank.", vbCritical, "Empty Error:"
     Exit Sub
  End If


 On Error GoTo ERRORHANDLER
    
    If OkCmd.Caption = "&Ok" Then
       OkCmd.Caption = "&Save"
       itemCombo1.Visible = False
       itemCombo2.Visible = False
       proNameCombo1.Visible = False
       proNameCombo2.Visible = False
       itemCombo1.CLEAR
       itemCombo2.CLEAR
       proNameCombo1.CLEAR
       proNameCombo2.CLEAR
       itemText1.Text = ""
       itemText2.Text = ""
       itemText1.Visible = False
       itemText2.Visible = False
       ridGrid.Enabled = False
       aoGrid.Enabled = False
       For r1 = 1 To 99
           For c1 = 1 To 4
              If aoGrid.TextMatrix(r1, 1) = "" Then
                 chkBlk1 = True
              End If
              If chkBlk1 Then
                  aoGrid.TextMatrix(r1, c1) = ""
              End If
           Next
           
           If aoGrid.TextMatrix(r1, 2) = "" Or aoGrid.TextMatrix(r1, 3) = "" Or aoGrid.TextMatrix(r1, 4) = "" Then
              For d1 = 1 To 4
                aoGrid.TextMatrix(r1, d1) = ""
              Next
           End If
       Next
         
         For r2 = 1 To 99
           For c2 = 1 To 3
               If ridGrid.TextMatrix(r2, 1) = "" Then
                  chkBlk2 = True
               End If
               If chkBlk2 Then
                  ridGrid.TextMatrix(r2, c2) = ""
               End If
           Next
           
           If ridGrid.TextMatrix(r2, 2) = "" Or ridGrid.TextMatrix(r2, 3) = "" Then
              For d2 = 1 To 3
                  ridGrid.TextMatrix(r2, d2) = ""
              Next
           End If
         Next
         
         cancelCmd.Enabled = False
         Exit Sub
    End If
      
      
  If OkCmd.Caption = "&Save" Then
     numOfRow1 = 0
     numOfRow2 = 0
     chkBlk1 = True
     chkBlk2 = True
    
  For r1 = 1 To rowCount1
      If aoGrid.TextMatrix(r1, 1) <> "" Then
          chkBlk1 = False
          numOfRow1 = numOfRow1 + 1
      End If
  Next
   
  For r2 = 1 To rowCount2
      If ridGrid.TextMatrix(r2, 1) <> "" Then
         chkBlk2 = False
         numOfRow2 = numOfRow2 + 1
      End If
      
  Next
   
   
  If chkBlk1 = False And chkBlk2 = False Then
    If custRadBtn.Value Then
        typOfCust = "customer"
    Else
        typOfCust = "party"
    End If
     
     dateStr = dayCombo.Text + "-" + monthCombo.Text + "-" + yearCombo.Text
      
     insertSql = "insert into HRISHI.ON_SPOT_REP_DETAIL values (" & Val(slnLabel.Caption) & ",'" & _
                   typOfCust & "'," & Val(rplrIdCombo.Text) & ",'" & UCase(rplrNameText.Text) & "','" & dateStr & "')"
                   
     osrDatabase.ExecuteSQL (insertSql)
     
     If chkBlk1 = False Then
    
       For r1 = 1 To numOfRow1

        insertSql = "insert into HRISHI.ON_SPOT_DEFECT_ITEM_DETAIL values (" & Val(slnLabel.Caption) & ",'" & _
                     UCase(aoGrid.TextMatrix(r1, 2)) & "', " & Val(aoGrid.TextMatrix(r1, 3)) & ",'" & UCase(aoGrid.TextMatrix(r1, 4)) & "')"
                     
        osrDatabase.ExecuteSQL (insertSql)
        Next
     End If
     
     If chkBlk2 = False Then
        For r2 = 1 To numOfRow2
        
         insertSql = "insert into HRISHI.ON_SPOT_REPLACED_ITEM_DETAIL values(" & Val(slnLabel.Caption) & ",'" & _
                      UCase(ridGrid.TextMatrix(r2, 2)) & "'," & Val(ridGrid.TextMatrix(r2, 3)) & ")"
                      
         osrDatabase.ExecuteSQL (insertSql)
        Next
      End If
     Else
       MsgBox "YOUR DATABASE ISN'T SAVE ,DATA ISN'T PRESENT IN BOTH GRID. ", vbCritical, "ERROR :"
       
       Unload Me
       Exit Sub
       
      
    End If
    
    If MsgBox("Sucess ! data is saved ." & vbCrLf & "Do you want to Continue ?", vbInformation + vbYesNo, "Suscess:") = vbYes Then
      OkCmd.Caption = "&Ok"
      cancelCmd.Enabled = True
      Set slnDyn = osrDatabase.CreateDynaset("select OSPR_SLN.NEXTVAL from dual", &H4&)
      slnLabel.Caption = slnDyn.Fields(0)
      Call CLEAR
      ridGrid.Enabled = True
      aoGrid.Enabled = True
    Else
     Unload Me
    End If
    
    
    Exit Sub
  End If


ERRORHANDLER:
   If osrSession.LastServerErr = 0 Then
      If osrDatabase.LastServerErr = 0 Then
         If Err.Number = 0 Then
         Else
          MsgBox "VB ERROR :" & vbCrLf & Err.Number & Err.Description, vbCritical, "VB Error:"
          Exit Sub
         End If
      Else
        MsgBox "DATABASE ERROR :" & vbCrLf & osrDatabase.LastServerErr & osrDatabase.LastServerErrText, vbCritical, "DATABASE Error :"
        osrDatabase.LastServerErrReset
        Exit Sub
      End If
  Else
    MsgBox "SESSION ERROR :" & vbCrLf & osrSession.LastServerErr & osrSession.LastServerErrText, vbCritical, "SESSION Error:"
    osrSession.LastServerErrReset
    Exit Sub
  End If
    
End Sub

Private Sub partyRadBtn_Click()

  Call Party_Id

End Sub

Private Sub proNameCombo1_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 Then
    aoGrid.Col = aoGrid.Col + 1
    aoGrid.ColSel = aoGrid.Col
 End If
 
End Sub

Private Sub proNameCombo2_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
     ridGrid.Col = ridGrid.Col + 1
     ridGrid.ColSel = ridGrid.Col
  End If

End Sub

Private Sub ridGrid_EnterCell()
 
 If ridGrid.Col = 2 Then
    itemCombo2.Visible = True
    itemCombo2.Top = ridGrid.Top + ridGrid.CellTop
    itemCombo2.Left = ridGrid.Left + ridGrid.CellLeft
    itemCombo2.Width = ridGrid.CellWidth
    itemCombo2.CLEAR
    
    Set itemDyn = osrDatabase.CreateDynaset("select distinct (category) from HRISHI.PRODUCT_DETAIL", &H0&)
   
    While Not itemDyn.EOF
       itemCombo2.AddItem itemDyn.Fields(0)
       itemDyn.MoveNext
    Wend
    Else
     If ridGrid.Col = 1 Then
         Exit Sub
     End If
     If itemText2.Enabled Then
         proNameCombo2.Visible = False
         itemText2.Text = ""
         itemText2.Top = ridGrid.Top + ridGrid.CellTop
         itemText2.Left = ridGrid.Left + ridGrid.CellLeft
         itemText2.Width = ridGrid.CellWidth
         itemText2.Height = ridGrid.CellHeight
         itemText2.Visible = True
         itemText2.SetFocus
      End If
  End If
  
     
End Sub

Private Sub ridGrid_LeaveCell()

  If ridGrid.Col = 2 Then
     ridGrid.Text = proNameCombo2.Text
     proNameCombo2.Visible = False
  End If

  If ridGrid.Col = 3 And itemText2.Enabled Then
     If itemText2.Text = "" Then
        itemText2.Text = "0.00"
     End If
     ridGrid.Text = itemText2.Text
     itemText2.Text = ""
     itemText2.Visible = False
  End If
    
End Sub

Private Sub rplrIdCombo_Click()

  If rplrIdCombo.Text <> "" Then
      If custRadBtn.Value Then
         Set custNameDyn = osrDatabase.CreateDynaset("select CUST_NAME from HRISHI.CUSTOMER_DETAIL where CUST_ID= " & rplrIdCombo.Text & "", &H4&)
         rplrNameText.Text = custNameDyn.Fields("CUST_NAME").Value
      Else
        Set partyNameDyn = osrDatabase.CreateDynaset("select PARTY_NAME from HRISHI.PARTY_DETAIL where PARTY_ID= " & rplrIdCombo.Text & "", &H4&)
        rplrNameText.Text = partyNameDyn.Fields("PARTY_NAME").Value
      End If
  End If

End Sub

Private Sub rplrIdCombo_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
     ridGrid.SetFocus
  End If

End Sub

Private Sub yearcombo_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     custRadBtn.SetFocus
  End If
 
End Sub
