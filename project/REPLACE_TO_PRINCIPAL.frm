VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form REPLACE_TO_PRINCIPAL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REPLACE TO PRINCIPAL  FORM [ENTRY]"
   ClientHeight    =   8070
   ClientLeft      =   2850
   ClientTop       =   720
   ClientWidth     =   9300
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
   ScaleHeight     =   8070
   ScaleWidth      =   9300
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   7695
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "12:40 PM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "9/23/2005"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10584
            MinWidth        =   10584
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   6600
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
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
      Left            =   7800
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
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
      Left            =   6600
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton extcmd 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   5520
      TabIndex        =   7
      ToolTipText     =   "Exit from this."
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton cancelcmd 
      Caption         =   "&Canel"
      Height          =   495
      Left            =   4170
      TabIndex        =   6
      ToolTipText     =   "Clear the Information."
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton okcmd 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   2850
      TabIndex        =   5
      ToolTipText     =   "Save the data."
      Top             =   6960
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid rtpGrid 
      Height          =   3015
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Enter the Replace Item Detail"
      Top             =   3240
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5318
      _Version        =   393216
      Rows            =   100
      Cols            =   5
      BackColor       =   16777215
      ForeColor       =   16711935
      GridLinesFixed  =   1
      FormatString    =   $"REPLACE_TO_PRINCIPAL.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox venIdCombo 
      Height          =   360
      Left            =   1440
      TabIndex        =   3
      ToolTipText     =   "Select the Vendor ID"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.ComboBox dayCombo 
      Height          =   360
      Left            =   5760
      TabIndex        =   0
      Text            =   " "
      Top             =   1320
      Width           =   855
   End
   Begin VB.ComboBox monthCombo 
      Height          =   360
      Left            =   6600
      TabIndex        =   1
      Text            =   " "
      Top             =   1320
      Width           =   1575
   End
   Begin VB.ComboBox yearcombo 
      Height          =   360
      Left            =   8160
      TabIndex        =   2
      Text            =   " "
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox venNameText 
      Height          =   375
      Left            =   5640
      TabIndex        =   15
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   0
      Picture         =   "REPLACE_TO_PRINCIPAL.frx":0006
      ToolTipText     =   "San's Replace to Principal Entry"
      Top             =   0
      Width           =   3000
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   3720
      X2              =   3720
      Y1              =   1080
      Y2              =   2760
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   -120
      X2              =   9360
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   -120
      X2              =   -120
      Y1              =   2040
      Y2              =   5280
   End
   Begin VB.Label venNameLabel 
      Caption         =   "Vendor Name :-"
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00008000&
      X1              =   5280
      X2              =   5280
      Y1              =   6840
      Y2              =   7560
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00008000&
      X1              =   3960
      X2              =   3960
      Y1              =   6840
      Y2              =   7560
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   6600
      X2              =   6600
      Y1              =   6840
      Y2              =   7560
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   2640
      X2              =   6600
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   2640
      X2              =   2640
      Y1              =   6840
      Y2              =   7560
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   2640
      X2              =   6600
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   120
      X2              =   840
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   3120
      Y2              =   6360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   120
      X2              =   9120
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   9120
      X2              =   9120
      Y1              =   3120
      Y2              =   6360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   3360
      X2              =   9120
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label5 
      Caption         =   "Replaced item detail :-"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   840
      TabIndex        =   18
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label venIdlabel 
      Caption         =   "Vendor ID. :-"
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   " Day"
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   " Month"
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   " Year"
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   8280
      TabIndex        =   12
      Top             =   960
      Width           =   735
   End
   Begin VB.Label dateLabel 
      Caption         =   "Date :-"
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label rpllable 
      Alignment       =   2  'Center
      Caption         =   "REPLACE TO PRINCIPAL"
      BeginProperty Font 
         Name            =   "FABIAN"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label slnLabel 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Slno. :-"
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1320
      Width           =   975
   End
End
Attribute VB_Name = "REPLACE_TO_PRINCIPAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************'
'**                                 **'
'** REPLACE TO PRINCIPAL ENTRY FORM **'
'**                                 **'
'*************************************'

'VARIABLE DECLARATION

Option Explicit


Dim rtpSession    As OraSession
Dim rtpDatabase   As OraDatabase
Dim slnDyn        As OraDynaset
Dim vidDyn        As OraDynaset
Dim venNameDyn    As OraDynaset
Dim itemDyn       As OraDynaset
Dim proDyn        As OraDynaset
Dim r             As Integer
Dim c             As Integer
Dim d             As Integer
Dim rowCount      As Integer
Dim numOfRow      As Integer
Dim cateGo        As String
Dim dateStr       As String
Dim insertSql     As String
Dim formatString  As String
Dim checkDate     As Boolean
Dim chkBlk        As Boolean


Private Sub CLEAR()
  
  dayCombo.Text = DAY(Date)
  monthCombo.Text = MonthName(MONTH(Date))
  yearCombo.Text = YEAR(Date)
  
  For r = 1 To 99
      For c = 1 To 4
        rtpGrid.TextMatrix(r, c) = ""
      Next
 Next
 
 rtpGrid.Col = 1
 rtpGrid.Row = 1
 rtpGrid.RowSel = 1
 rtpGrid.ColSel = 1
 rowCount = 1
 rtpGrid.TextMatrix(1, 1) = rowCount
 rowCount = rowCount + 1
 itemText.Text = ""
 dayCombo.SetFocus

End Sub

Private Sub cancelCmd_Click()
 
 Call CLEAR
 
End Sub

Private Sub dayCombo_GotFocus()
  
  StatusBar1.Panels(3) = "SELECT THE DAY.."

End Sub

Private Sub extCmd_Click()
 
 If MsgBox("Do you want to Exit ?", vbExclamation + vbYesNo, "Exit:") = vbYes Then
    Unload Me
 End If

End Sub

Private Sub Form_Load()

  Call FILLCOMBODATE(dayCombo, monthCombo, yearCombo)
  formatString$ = "| Slno. |^              Item Name                       |^       Qty      |^                            Description"
  rtpGrid.formatString = formatString$

  On Error GoTo ERRORHANDLER

  Set rtpSession = CreateObject("oracleinprocserver.xorasession")
  Set rtpDatabase = rtpSession.OpenDatabase("jms", "hrishi/jms", &H0&)
  Set slnDyn = rtpDatabase.CreateDynaset("select RPL_TO_PRNPAL_SLN.NEXTVAL from DUAL", &H4&)
  Set vidDyn = rtpDatabase.CreateDynaset("select VENDOR_ID from VENDOR_DETAIL", &H4&)
  slnLabel.Caption = slnDyn.Fields(0).Value
  
  While Not vidDyn.EOF
     venIdCombo.AddItem vidDyn.Fields(0)
      vidDyn.MoveNext
  Wend
 
 venIdCombo.ListIndex = 0
 rowCount = 1
 rtpGrid.Col = 1
 rtpGrid.Row = 1
 rtpGrid.TextMatrix(1, 1) = rowCount
 rowCount = rowCount + 1

ERRORHANDLER:
  If rtpSession.LastServerErr = 0 Then
     If rtpDatabase.LastServerErr = 0 Then
        If Err.Number = 0 Then
        Else
          MsgBox "VB ERROR:" & vbCrLf & Err.Number & Err.Description, vbCritical, "VB Error:"
          Exit Sub
        End If
     Else
       MsgBox "DATABASE ERROR:" & vbCrLf & rtpDatabase.LastServerErr & rtpDatabase.LastServerErrText, vbCritical, "DATABASE Error:"
       rtpDatabase.LastServerErrReset
       Exit Sub
     End If
 Else
   MsgBox "SESSION ERROR:" & vbCrLf & rtpSession.LastServerErr & rtpSession.LastServerErrText, vbCritical, "SESSION Error:"
   rtpSession.LastServerErrReset
   Exit Sub
 End If

End Sub

Private Sub itemCombo_Click()
 
 cateGo = itemCombo.Text
 proNameCombo.CLEAR
 itemCombo.Visible = False
 
 Set proDyn = rtpDatabase.CreateDynaset("select PRODUCT_NAME FROM HRISHI.PRODUCT_DETAIL where CATEGORY='" & cateGo & "'", &H0&)
 
 While Not proDyn.EOF
     proNameCombo.AddItem proDyn.Fields(0)
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
     If rtpGrid.Col < 4 Then
        rtpGrid.Col = rtpGrid.Col + 1
        rtpGrid.ColSel = rtpGrid.Col
        If rtpGrid.Col = 2 Then
           itemCombo.SetFocus
        Else
           itemText.SetFocus
        End If
     Else
        itemText.Visible = False
        rtpGrid.Col = 1
        If rtpGrid.Row = 99 Then
           rtpGrid.Row = 1
        Else
           rtpGrid.Row = rtpGrid.Row + 1
        End If
        rtpGrid.ColSel = rtpGrid.Col
        rtpGrid.RowSel = rtpGrid.Row
        rtpGrid.TextMatrix(rtpGrid.Row, rtpGrid.Col) = rowCount
        rowCount = rowCount + 1
        rtpGrid.Col = 2
        rtpGrid.ColSel = rtpGrid.Col
    End If
  End If
   
End Sub

Private Sub monthCombo_GotFocus()

   StatusBar1.Panels(3) = "SELECT THE MONTH .."

End Sub

Private Sub OkCmd_Click()
  
  checkDate = VERIFY_DATE(Val(dayCombo.Text), monthCombo.Text, Val(yearCombo.Text))
  
  If checkDate = False Then
       MsgBox "Error ! Invalid date", vbCritical, "DATE Error:"
       Exit Sub
  End If
   
  If venIdCombo.Text = "" Or venNameText.Text = "" Then
     MsgBox "Empty !  Vendor information isn't present ", vbCritical, "EMPTY Error:"
     Exit Sub
  End If
   
  On Error GoTo ERRORHANDLER
   
   If OkCmd.Caption = "&Ok" Then
      OkCmd.Caption = "&Save"
      itemCombo.Visible = False
      proNameCombo.Visible = False
      itemText.Visible = False
      itemText.Text = ""
      itemCombo.CLEAR
      proNameCombo.CLEAR
      rtpGrid.Enabled = False
      chkBlk = False
      
      For r = 1 To 99
        For c = 1 To 4
           If rtpGrid.TextMatrix(r, 1) = "" Then
              chkBlk = True
           End If
           If chkBlk Then
              rtpGrid.TextMatrix(r, c) = ""
           End If
        Next
        
        If rtpGrid.TextMatrix(r, 2) = "" Or rtpGrid.TextMatrix(r, 3) = "" Or rtpGrid.TextMatrix(r, 4) = "" Then
           For d = 1 To 4
               rtpGrid.TextMatrix(r, d) = ""
           Next
        End If
    Next
    cancelCmd.Enabled = False
    Exit Sub
  End If
  
 If OkCmd.Caption = "&Save" Then
    numOfRow = 0
    chkBlk = True
    
    For r = 1 To rowCount
       If rtpGrid.TextMatrix(r, 1) <> "" Then
         chkBlk = False
         numOfRow = numOfRow + 1
       End If
     Next
     
     If chkBlk = False Then
     
    dateStr = dayCombo.Text + "-" + monthCombo.Text + "-" + yearCombo.Text
    
    insertSql = "insert into HRISHI.REPLACE_TO_PRINCIPAL_DETAIL values ( " & Val(slnLabel.Caption) & "," & Val(venIdCombo.Text) & ",'" & _
                 UCase(venNameText.Text) & "','" & dateStr & "')"
    rtpDatabase.ExecuteSQL (insertSql)
    
       
  For r = 1 To numOfRow
    insertSql = "insert into HRISHI.REPLACE_TO_PRNPAL_ITEM_DETAIL values (" & Val(slnLabel.Caption) & ",'" & UCase(rtpGrid.TextMatrix(r, 2)) & "'," & _
                Val(rtpGrid.TextMatrix(r, 3)) & ",'" & UCase(rtpGrid.TextMatrix(r, 4)) & "')"
    rtpDatabase.ExecuteSQL (insertSql)
  Next
  End If
  If MsgBox("Sucess ! data is saved." & vbCrLf & "Do you want to continue ?", vbInformation + vbYesNo, "Sucess:") = vbYes Then
    OkCmd.Caption = "&Ok"
    cancelCmd.Enabled = True
    Set slnDyn = rtpDatabase.CreateDynaset("select RPL_TO_PRNPAL_SLN.NEXTVAL from DUAL", &H4&)
    slnLabel.Caption = slnDyn.Fields(0)
    Call CLEAR
    rtpGrid.Enabled = True
  Else
    Unload Me
  End If
  
  Exit Sub
End If


ERRORHANDLER:
  If rtpSession.LastServerErr = 0 Then
     If rtpDatabase.LastServerErr = 0 Then
        If Err.Number = 0 Then
        Else
          MsgBox "VB ERROR :" & vbCrLf & Err.Number & Err.Description, vbCritical, "VB Error:"
          Exit Sub
        End If
     Else
      MsgBox "DATABASE ERROR :" & vbCrLf & rtpDatabase.LastServerErr & rtpDatabase.LastServerErrText, vbCritical, "DATABASE Error:"
      rtpDatabase.LastServerErrReset
      Exit Sub
     End If
  Else
   MsgBox "SESSION ERROR :" & vbCrLf & rtpSession.LastServerErr & rtpSession.LastServerErrText, vbCritical, "SESSION Error:"
   rtpSession.LastServerErrReset
   Exit Sub
  End If
  
End Sub

Private Sub proNameCombo_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 Then
    rtpGrid.Col = rtpGrid.Col + 1
    rtpGrid.ColSel = rtpGrid.Col
  End If
 
End Sub

Private Sub rtpGrid_EnterCell()
 
  If rtpGrid.Col = 2 And rtpGrid.Enabled Then
    itemCombo.Visible = True
    itemCombo.Top = rtpGrid.Top + rtpGrid.CellTop
    itemCombo.Left = rtpGrid.Left + rtpGrid.CellLeft
    itemCombo.Width = rtpGrid.CellWidth
    itemCombo.CLEAR
        
    Set itemDyn = rtpDatabase.CreateDynaset("select distinct (category) from HRISHI.PRODUCT_DETAIL ", &H0&)
    While Not itemDyn.EOF
         itemCombo.AddItem itemDyn.Fields(0)
         itemDyn.MoveNext
    Wend
  Else
     If rtpGrid.Col = 1 Then
        Exit Sub
     End If
     
     If itemText.Enabled Then
         proNameCombo.Visible = False
         itemText.Text = ""
         itemText.Top = rtpGrid.Top + rtpGrid.CellTop
         itemText.Left = rtpGrid.Left + rtpGrid.CellLeft
         itemText.Width = rtpGrid.CellWidth
         itemText.Height = rtpGrid.CellHeight
         itemText.Visible = True
         itemText.SetFocus
     End If
   End If
     
End Sub

Private Sub rtpGrid_GotFocus()

  StatusBar1.Panels(3) = "ENTER THE INFORMATION OF REPLACEMENT"

End Sub

Private Sub rtpGrid_LeaveCell()
  
  If rtpGrid.Col = 2 Then
     rtpGrid.Text = proNameCombo.Text
     proNameCombo.Visible = False
  End If
  
  If rtpGrid.Col = 3 And itemText.Enabled Then
     If itemText.Text = "" Then
         itemText.Text = "0.00"
     End If
     rtpGrid.Text = itemText.Text
     itemText.Text = ""
     itemText.Visible = False
   End If
   
   If rtpGrid.Col = 4 And itemText.Enabled Then
      If itemText.Text = "" Then
         itemText.Text = "---DEFECTIVE---"
      End If
      rtpGrid.Text = itemText.Text
      itemText.Text = ""
      itemText.Visible = False
   End If
   
End Sub

Private Sub venIdCombo_Click()

  Set venNameDyn = rtpDatabase.CreateDynaset("select VENDOR_NAME from VENDOR_DETAIL where VENDOR_ID =" & venIdCombo.Text & "", &H0&)
   venNameText.Text = venNameDyn.Fields(0)

End Sub

Private Sub venIdCombo_GotFocus()

   StatusBar1.Panels(3) = "SELECT THE VENDOR ID"

End Sub

Private Sub yearcombo_GotFocus()

  StatusBar1.Panels(3) = "SELECT THE YEAR"

End Sub
