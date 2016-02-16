VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form SALE_ENTRY 
   AutoRedraw      =   -1  'True
   Caption         =   "SALE FORM [ENTRY]"
   ClientHeight    =   10155
   ClientLeft      =   2505
   ClientTop       =   690
   ClientWidth     =   10770
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
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   10155
   ScaleWidth      =   10770
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
      Left            =   7440
      TabIndex        =   25
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
      Left            =   7680
      TabIndex        =   24
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
      Left            =   7560
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton OkCmd 
      BackColor       =   &H0000C000&
      Caption         =   "&Ok"
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
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Save the Information."
      Top             =   8760
      Width           =   1095
   End
   Begin VB.CommandButton CancelCmd 
      BackColor       =   &H00FF8080&
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Clear the information."
      Top             =   8760
      Width           =   1095
   End
   Begin VB.CommandButton extCmd 
      BackColor       =   &H008080FF&
      Caption         =   "&Exit"
      Height          =   615
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Exit from this."
      Top             =   8760
      Width           =   1095
   End
   Begin VB.TextBox taxText 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   7920
      TabIndex        =   19
      Text            =   "0.00"
      ToolTipText     =   "Enter the Tax "
      Top             =   8640
      Width           =   615
   End
   Begin VB.TextBox discText 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   7920
      TabIndex        =   18
      Text            =   "0.00"
      ToolTipText     =   "Enter the Discont."
      Top             =   8280
      Width           =   615
   End
   Begin VB.TextBox totalSaleAmountText 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   7920
      TabIndex        =   17
      Top             =   9720
      Width           =   2055
   End
   Begin VB.TextBox disAmountText 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   7920
      TabIndex        =   16
      Top             =   9000
      Width           =   2055
   End
   Begin VB.TextBox taxAmountText 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   7920
      TabIndex        =   15
      Top             =   9360
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid saleGrid 
      Height          =   3855
      Left            =   480
      TabIndex        =   10
      ToolTipText     =   "Enter Sale infromation."
      Top             =   3720
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   6800
      _Version        =   393216
      Rows            =   100
      Cols            =   5
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
   Begin VB.Frame selTypfrm 
      Caption         =   "Select the type of customer :-"
      ForeColor       =   &H00C0C000&
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Enter the customer Information"
      Top             =   1920
      Width           =   10575
      Begin VB.CommandButton selTypCmd 
         Caption         =   "Create Customer"
         Height          =   375
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   2535
      End
      Begin VB.OptionButton partyRadBtn 
         Caption         =   "Party"
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   4560
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton custRadBtn 
         Caption         =   "Customer"
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label typOfCust 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2760
         TabIndex        =   28
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label IdLabel 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4815
         TabIndex        =   27
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label NameLabel 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   6150
         TabIndex        =   26
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "To whome Sale :- "
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   1935
      End
   End
   Begin VB.ComboBox yearCombo 
      Height          =   360
      Left            =   9270
      TabIndex        =   5
      ToolTipText     =   "Select the Year."
      Top             =   1320
      Width           =   1335
   End
   Begin VB.ComboBox monthCombo 
      Height          =   360
      Left            =   7695
      TabIndex        =   4
      ToolTipText     =   "Select the  Month."
      Top             =   1320
      Width           =   1575
   End
   Begin VB.ComboBox dayCombo 
      Height          =   360
      Left            =   6720
      TabIndex        =   3
      ToolTipText     =   "Select the Day."
      Top             =   1320
      Width           =   975
   End
   Begin VB.Line Line11 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   10800
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   4920
      X2              =   4920
      Y1              =   8520
      Y2              =   9600
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   720
      X2              =   4920
      Y1              =   9600
      Y2              =   9600
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   720
      X2              =   720
      Y1              =   9600
      Y2              =   8520
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   720
      X2              =   4920
      Y1              =   8520
      Y2              =   8520
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00C000C0&
      BorderWidth     =   2
      X1              =   240
      X2              =   1200
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label3 
      Caption         =   "Sale Entry :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1320
      TabIndex        =   35
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   0
      Picture         =   "SALE_ENTRY.frx":0000
      ToolTipText     =   "San's Sale Entry."
      Top             =   0
      Width           =   3000
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00C000C0&
      BorderWidth     =   2
      X1              =   10560
      X2              =   10560
      Y1              =   3480
      Y2              =   8160
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C000C0&
      BorderWidth     =   2
      X1              =   240
      X2              =   240
      Y1              =   3480
      Y2              =   8160
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C000C0&
      BorderWidth     =   2
      X1              =   240
      X2              =   10560
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C000C0&
      BorderWidth     =   2
      X1              =   2400
      X2              =   10560
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   10320
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label SaleEntryLabel 
      Alignment       =   2  'Center
      Caption         =   "SALE ENTRY"
      BeginProperty Font 
         Name            =   "BAILEY"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3735
      TabIndex        =   34
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label6 
      Caption         =   "Total amount :-"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   5520
      TabIndex        =   33
      Top             =   9720
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Total tax :-"
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   5520
      TabIndex        =   32
      Top             =   9360
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Total discount:-"
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   5520
      TabIndex        =   31
      Top             =   9000
      Width           =   1935
   End
   Begin VB.Label taxLabel 
      Caption         =   "Tax :-"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   5520
      TabIndex        =   30
      Top             =   8640
      Width           =   1935
   End
   Begin VB.Label discLabel 
      Caption         =   "Discount :-"
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   5520
      TabIndex        =   29
      Top             =   8280
      Width           =   1935
   End
   Begin VB.Label qtyLabel 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5040
      TabIndex        =   14
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label totalLabel 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7680
      TabIndex        =   13
      Top             =   7680
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   4800
      TabIndex        =   12
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label dateLabel 
      Caption         =   "Date :- "
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label slnLabel 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label slnLabel1 
      Caption         =   "Sl No. :-"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "SALE_ENTRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*********************'
'**                 **'
'** SALE ENTRY FORM **'
'**                 **'
'*********************'

'VARIABLE DECLERATION

Option Explicit

Dim saleSession        As OraSession
Dim saleDatabase       As OraDatabase
Dim SaleSlnDyn         As OraDynaset
Dim itemDyn            As OraDynaset
Dim proDyn             As OraDynaset
Dim maxQtyDyn          As OraDynaset
Dim S                  As String
Dim cateStr            As String
Dim dateStr            As String
Dim insertSql          As String
Dim updStkSql          As String
Dim updCustSql         As String
Public rowCount        As Integer
Dim numOfRow           As Integer
Dim qty                As Integer
Dim maxQty             As Integer
Dim r                  As Integer
Dim c                  As Integer
Dim d                  As Integer
Dim price              As Double
Dim total              As Double
Dim checkDate          As Boolean
Dim chkBlk             As Boolean
 

Private Sub CLEAR()
  
  dayCombo.Text = DAY(Date)
  monthCombo.Text = MonthName(MONTH(Date))
  yearCombo.Text = YEAR(Date)
  typOfCust.Caption = ""
  IdLabel.Caption = ""
  NameLabel.Caption = ""
  
  
  For r = 1 To 99
    For c = 1 To 5
       saleGrid.TextMatrix(r, c) = ""
    Next
 Next

  saleGrid.Col = 1
  saleGrid.Row = 1
  saleGrid.ColSel = 1
  saleGrid.RowSel = 1
  rowCount = 1
  saleGrid.TextMatrix(1, 1) = rowCount
  rowCount = rowCount + 1
  dayCombo.SetFocus
  qtyLabel.Caption = "0.00"
  totalLabel.Caption = "0.00"
  discText.Text = "0.00"
  taxText.Text = "0.00"
  disAmountText.Text = ""
  taxAmountText.Text = ""
  totalSaleAmountText.Text = ""
  itemText.Text = ""
  
End Sub

Private Sub cancelCmd_Click()

  Call CLEAR

End Sub

Private Sub custRadBtn_Click()

  selTypCmd.Caption = "Create Customer"

End Sub

Private Sub dayCombo_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     monthCombo.SetFocus
  End If

End Sub

Private Sub extCmd_Click()
  
  If MsgBox("Do you want to exit ?", vbExclamation + vbYesNo, "Exit:") = vbYes Then
     Unload Me
  End If

End Sub

Private Sub Form_Load()

    On Error GoTo ERRORHANDLER
    custRadBtn.Value = True
    Call FILLCOMBODATE(dayCombo, monthCombo, yearCombo)
    S$ = "|Slno.|^                 Item name                                 |^       Qty    | Price per Unit |^        Total            "
    saleGrid.formatString = S$

    Set saleSession = CreateObject("oracleinprocserver.xorasession")
    Set saleDatabase = saleSession.OpenDatabase("jms", "hrishi/jms", &H4&)
    Set SaleSlnDyn = saleDatabase.CreateDynaset("select sale_sln.nextval from dual ", &H4&)
    slnLabel = SaleSlnDyn.Fields(0)

    rowCount = 1              'INITIALINGING THE VALUE OF ROWCOUNT
    saleGrid.Col = 1
    saleGrid.Row = 1
    saleGrid.TextMatrix(1, 1) = rowCount
    rowCount = rowCount + 1
    qtyLabel.Caption = "0.00"
    totalLabel.Caption = "0.00"
  
ERRORHANDLER:

If saleSession.LastServerErr = 0 Then
   If saleDatabase.LastServerErr = 0 Then
     If Err.Number = 0 Then
     Else
        MsgBox "VB ERROR :" & Err.Number & " :: " & Err.Description, vbCritical, "VB Error:"
        Unload Me
     End If
   Else
      MsgBox "DATABASE ERROR :" & saleDatabase.LastServerErr & saleDatabase.LastServerErrText, vbCritical, "DATABASE Error :"
      saleDatabase.LastServerErrReset
      Unload Me
   End If
Else
   MsgBox "SESSION ERROR :" & saleSession.LastServerErr & saleSession.LastServerErrText, vbCritical, "DATABASE Error :"
   saleSession.LastServerErrReset
   Unload Me
End If

End Sub

Private Sub itemCombo_Click()

    cateStr = itemCombo.Text
    proNameCombo.CLEAR
    itemCombo.Visible = False

    Set proDyn = saleDatabase.CreateDynaset("Select PRODUCT_NAME from HRISHI.PRODUCT_DETAIL  where CATEGORY = '" & cateStr & "'", &H4&)

    While Not proDyn.EOF
         proNameCombo.AddItem proDyn.Fields(0)
         proDyn.MoveNext
    Wend


    proNameCombo.Top = itemCombo.Top
    proNameCombo.Left = itemCombo.Left
    proNameCombo.Width = itemCombo.Width
    proNameCombo.ListIndex = 0
    proNameCombo.Visible = True

End Sub

Private Sub itemText_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       If saleGrid.Col < 4 Then
            saleGrid.Col = saleGrid.Col + 1
            saleGrid.ColSel = saleGrid.Col
            If saleGrid.Col = 2 Then
                itemCombo.SetFocus
            Else
            itemText.SetFocus
            End If
       Else
       itemText.Visible = False
       saleGrid.Col = 1
       If saleGrid.Row = 99 Then
          saleGrid.Row = 1
       Else
          saleGrid.Row = saleGrid.Row + 1
       End If
        saleGrid.ColSel = saleGrid.Col
        saleGrid.RowSel = saleGrid.Row
        saleGrid.TextMatrix(saleGrid.Row, saleGrid.Col) = rowCount
        rowCount = rowCount + 1
        saleGrid.Col = 2
        saleGrid.ColSel = saleGrid.Col
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
        MsgBox "Error ! Invalid Date .", vbCritical, "DATE Error "
        Exit Sub
    End If
    If IdLabel.Caption = " " Or NameLabel.Caption = "" Then
        Exit Sub
    End If


    On Error GoTo ERRORHANDLER
    If OkCmd.Caption = "&Ok" Then
         OkCmd.Caption = "&Save"
         itemCombo.Enabled = False
         itemText.Enabled = False
         proNameCombo.Enabled = False
         chkBlk = False
         itemCombo.CLEAR
         itemText.Text = ""
         proNameCombo.CLEAR
         disAmountText.Text = Val(totalLabel.Caption) * Val(discText.Text) / 100
         taxAmountText.Text = Val(totalLabel.Caption) * Val(taxText.Text) / 100
         totalSaleAmountText.Text = Val(totalLabel.Caption) - Val(disAmountText.Text) + Val(taxAmountText.Text)
     
         For r = 1 To 99
             For c = 1 To 5
                  If saleGrid.TextMatrix(r, 1) = "" Then
                        chkBlk = True
                  End If
                  If chkBlk Then
                        saleGrid.TextMatrix(r, c) = ""
                  End If
             Next
       
             If saleGrid.TextMatrix(r, 2) = "" Or saleGrid.TextMatrix(r, 3) = "" Or saleGrid.TextMatrix(r, 4) = "" Then
                For d = 1 To 5
                     saleGrid.TextMatrix(r, d) = ""
                Next
             End If
         Next
        cancelCmd.Enabled = False
        Exit Sub
    End If
 
 
    If OkCmd.Caption = "&Save" Then
         itemCombo.Enabled = True
         itemText.Enabled = True
         proNameCombo.Enabled = True
         itemCombo.CLEAR
         numOfRow = 0
         chkBlk = True
         For r = 1 To rowCount
            If saleGrid.TextMatrix(r, 1) <> "" Then
                chkBlk = False
                numOfRow = numOfRow + 1
            End If
         Next
    
    If chkBlk = False Then
        dateStr = dayCombo.Text + "-" + monthCombo.Text + "-" + yearCombo.Text
        insertSql = "insert into HRISHI.SALE_DETAIL values (" & Val(slnLabel.Caption) & ",'" & typOfCust.Caption & "'," & _
                    Val(IdLabel.Caption) & ",'" & NameLabel.Caption & "','" & dateStr & "'," & Val(totalLabel.Caption) & "," & _
                    Val(discText.Text) & "," & Val(taxText.Text) & "," & Val(totalSaleAmountText.Text) & ")"
                    
         saleDatabase.ExecuteSQL (insertSql)
         
         If typOfCust = "party" Then
             updCustSql = "update HRISHI.PARTY_DETAIL set PARTY_TOTAL_SALE=PARTY_TOTAL_SALE + " & Val(totalSaleAmountText.Text) & " where PARTY_ID= " & Val(IdLabel.Caption) & ""
         Else
             updCustSql = "update HRISHI.CUSTOMER_DETAIL set TOTAL_SALE= " & Val(totalSaleAmountText.Text) & " where CUST_ID= " & Val(IdLabel.Caption) & ""
         End If
         
         saleDatabase.ExecuteSQL (updCustSql)
        
        For r = 1 To numOfRow
              insertSql = " insert into HRISHI.SALE_ITEM_DETAIL values( " & Val(slnLabel.Caption) & ",'" & _
                     (saleGrid.TextMatrix(r, 2)) & "'," & Val(saleGrid.TextMatrix(r, 3)) & "," & _
                     Val(saleGrid.TextMatrix(r, 4)) & "," & Val(saleGrid.TextMatrix(r, 5)) & " )"
                updStkSql = "update HRISHI.STOCK_DETAIL set TOTAL_SALE_QTY = TOTAL_SALE_QTY + " & Val(saleGrid.TextMatrix(r, 3)) & ",STOCK_IN_HAND =STOCK_IN_HAND - " & Val(saleGrid.TextMatrix(r, 3)) & ",LAST_MODIFY_DATE = SYSDATE where PRODUCT_NAME='" & (saleGrid.TextMatrix(r, 2)) & "'"
                    
                saleDatabase.ExecuteSQL (insertSql)
               saleDatabase.ExecuteSQL (updStkSql)
        Next
       End If
       If MsgBox("Data is Saved ! do you want Continue.", vbInformation + vbYesNo, "Continue.") = vbYes Then
            OkCmd.Caption = "&Ok"
            cancelCmd.Enabled = True
            Set SaleSlnDyn = saleDatabase.CreateDynaset("select SALE_SLN.NEXTVAL from dual", &H4&)
            Call CLEAR
       Else
            Unload Me
       End If
       Exit Sub
    End If
      
ERRORHANDLER:
  If saleSession.LastServerErr = 0 Then
     If saleDatabase.LastServerErr = 0 Then
       If Err.Number = 0 Then
       Else
          MsgBox "VB ERROR: " & CStr(Err.Number) & Err.Description, vbCritical, "VB Error:"
          Exit Sub
       End If
     Else
       MsgBox "DATABASE ERROR :" & saleDatabase.LastServerErr & saleDatabase.LastServerErrText, vbCritical, "DATABASE Error:"
       saleDatabase.LastServerErrReset
       Exit Sub
     End If
 Else
   MsgBox " SESSION ERROR:" & saleSession.LastServerErr & saleSession.LastServerErrText, vbCritical, "SESSION Error:"
    saleSession.LastServerErrReset
    Exit Sub
 End If
 
End Sub

Private Sub partyRadBtn_Click()

    selTypCmd.Caption = "Create Party"

End Sub

Private Sub proNameCombo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        saleGrid.Col = saleGrid.Col + 1
        saleGrid.ColSel = saleGrid.Col
    End If

End Sub

Private Sub saleGrid_EnterCell()

    If saleGrid.Col = 2 And saleGrid.Enabled Then
        itemCombo.Visible = True
        itemCombo.Top = saleGrid.Top + saleGrid.CellTop
        itemCombo.Left = saleGrid.Left + saleGrid.CellLeft
        itemCombo.Width = saleGrid.CellWidth
        itemCombo.CLEAR
        Set itemDyn = saleDatabase.CreateDynaset("select distinct (category) from HRISHI.PRODUCT_DETAIL", &H4&)
  
        While Not itemDyn.EOF
            itemCombo.AddItem itemDyn.Fields(0)
            itemDyn.MoveNext
        Wend
    Else
        If saleGrid.Col = 1 Then
            Exit Sub
        End If
   
        If itemText.Enabled Then
            proNameCombo.Visible = False
            If saleGrid.Col > 3 Then
                itemText.Text = ""
            End If
            itemText.Top = saleGrid.Top + saleGrid.CellTop
            itemText.Left = saleGrid.Left + saleGrid.CellLeft
            itemText.Width = saleGrid.CellWidth
            itemText.Height = saleGrid.CellHeight
            itemText.Visible = True
            itemText.SetFocus
        End If
    End If
 
End Sub

Private Sub saleGrid_LeaveCell()
 
     If saleGrid.Col = 2 Then
        saleGrid.Text = proNameCombo.Text
        If proNameCombo.Text <> "" Then
            Set maxQtyDyn = saleDatabase.CreateDynaset("select STOCK_IN_HAND from HRISHI.STOCK_DETAIL where PRODUCT_NAME ='" & proNameCombo.Text & "'", &H4&)
            maxQty = maxQtyDyn.Fields(0)
        End If
        proNameCombo.Visible = False
        itemText.Text = maxQty
     End If
 
     If saleGrid.Col = 3 And Val(itemText.Text) > maxQty Then
        itemText.Text = maxQty
     End If
 
    If saleGrid.Col > 2 And itemText.Enabled Then
        If itemText.Text = "" Then
        itemText.Text = "0.00"
        End If
        saleGrid.Text = itemText.Text
        itemText.Visible = False
    End If
 
    If saleGrid.Col = 3 Then
        qtyLabel.Caption = Val(qtyLabel.Caption) + Val(saleGrid.TextMatrix(saleGrid.Row, 3))
    End If
 
    If saleGrid.Col = 4 Then
        qty = Val(saleGrid.TextMatrix(saleGrid.Row, saleGrid.Col - 1))
        price = Val(saleGrid.TextMatrix(saleGrid.Row, saleGrid.Col))
        total = qty * price
        saleGrid.TextMatrix(saleGrid.Row, saleGrid.Col + 1) = total
        totalLabel.Caption = Val(totalLabel.Caption) + total
    End If
 
End Sub

Private Sub selTypCmd_Click()
    
    If selTypCmd.Caption = "Create Customer" Then
        CUSTOMER_ENTRY.Show
    End If
    
    If selTypCmd.Caption = "Create Party" Then
        PARTY_ENTRY.Show
    End If

    If custRadBtn.Value = True Then
        IdLabel.Caption = custFind.custId
        NameLabel.Caption = custFind.custName
    Else
        IdLabel.Caption = custFind.partyId
        NameLabel.Caption = custFind.partyName
    End If

End Sub

