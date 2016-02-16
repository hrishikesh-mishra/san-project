VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PUR_ENTRY 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PURCHASE FORM [ENTRY]"
   ClientHeight    =   9780
   ClientLeft      =   690
   ClientTop       =   120
   ClientWidth     =   10770
   BeginProperty Font 
      Name            =   "Tahoma"
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
   ScaleHeight     =   9780
   ScaleWidth      =   10770
   Begin VB.TextBox taxAmountText 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8400
      TabIndex        =   32
      Top             =   8280
      Width           =   2055
   End
   Begin VB.TextBox disAmountText 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8400
      TabIndex        =   31
      Top             =   7800
      Width           =   2055
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   30
      Top             =   9405
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   661
      SimpleText      =   "\\"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Bevel           =   2
            Object.Width           =   2364
            MinWidth        =   2364
            TextSave        =   "9/23/2005"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Bevel           =   2
            Object.Width           =   2011
            MinWidth        =   2011
            TextSave        =   "2:03 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   11360
            MinWidth        =   11360
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   2364
            TextSave        =   "NUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox vNameTxt 
      Height          =   375
      Left            =   6840
      TabIndex        =   29
      Top             =   1920
      Width           =   3615
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
      TabIndex        =   28
      Top             =   -240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton exitcmd 
      BackColor       =   &H008080FF&
      Caption         =   "&Exit"
      Height          =   615
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Exit from this."
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton okCmd 
      BackColor       =   &H0080FF80&
      Caption         =   "&Ok"
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Save the data."
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cancelCmd 
      BackColor       =   &H00FF8080&
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Clear the information."
      Top             =   8040
      Width           =   1215
   End
   Begin VB.TextBox totalPurAmountText 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8400
      TabIndex        =   24
      Top             =   8760
      Width           =   2055
   End
   Begin VB.TextBox discText 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   8520
      TabIndex        =   20
      Text            =   "0.00"
      ToolTipText     =   "Enter the Discount in Percent."
      Top             =   6840
      Width           =   615
   End
   Begin VB.TextBox taxText 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   8520
      TabIndex        =   19
      Text            =   "0.00"
      ToolTipText     =   "Enter the Tax in  percent"
      Top             =   7320
      Width           =   615
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
      Left            =   7080
      TabIndex        =   12
      Top             =   -120
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
      Left            =   7200
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSFlexGridLib.MSFlexGrid purGrid 
      Height          =   3375
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Enter the Purchase information."
      Top             =   2760
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   5953
      _Version        =   393216
      Rows            =   100
      Cols            =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox ddCombo 
      Height          =   360
      Left            =   7080
      TabIndex        =   1
      ToolTipText     =   "Select The Day."
      Top             =   1200
      Width           =   735
   End
   Begin VB.ComboBox venIdCombo 
      Height          =   360
      Left            =   1560
      TabIndex        =   4
      ToolTipText     =   "Select the Vendor ID."
      Top             =   1800
      Width           =   1575
   End
   Begin VB.ComboBox yyCombo 
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
      Left            =   9480
      TabIndex        =   3
      ToolTipText     =   "Select The Year."
      Top             =   1200
      Width           =   975
   End
   Begin VB.ComboBox mmCombo 
      Height          =   360
      Left            =   7800
      TabIndex        =   2
      ToolTipText     =   "Select the Month."
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox slnTxt 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      TabIndex        =   7
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Line Line14 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      X1              =   3720
      X2              =   3720
      Y1              =   7920
      Y2              =   8760
   End
   Begin VB.Line Line13 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      X1              =   2040
      X2              =   2040
      Y1              =   7920
      Y2              =   8760
   End
   Begin VB.Line Line12 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      X1              =   5280
      X2              =   5280
      Y1              =   7920
      Y2              =   8760
   End
   Begin VB.Line Line11 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      X1              =   5280
      X2              =   480
      Y1              =   8760
      Y2              =   8760
   End
   Begin VB.Line Line10 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      X1              =   480
      X2              =   480
      Y1              =   8760
      Y2              =   7920
   End
   Begin VB.Line Line9 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      X1              =   480
      X2              =   5280
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   0
      Picture         =   "PUR_ENTRY.frx":0000
      ToolTipText     =   "San's Purchase Entry."
      Top             =   0
      Width           =   3000
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000C0&
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      X1              =   0
      X2              =   10680
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000C0&
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      X1              =   4560
      X2              =   4560
      Y1              =   1080
      Y2              =   2640
   End
   Begin VB.Line Line6 
      X1              =   10800
      X2              =   10800
      Y1              =   2040
      Y2              =   6120
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   0
      Y1              =   2640
      Y2              =   6600
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   360
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label7 
      Caption         =   "Enter information of Purchase "
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   33
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   3480
      X2              =   10800
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   5520
      X2              =   5520
      Y1              =   6600
      Y2              =   9480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   10800
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Label Label6 
      Caption         =   "Total amount :-"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   6240
      TabIndex        =   23
      Top             =   8760
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Total tax :-"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   6240
      TabIndex        =   22
      Top             =   8280
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Total discount:-"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   6240
      TabIndex        =   21
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Label percent1 
      Caption         =   "%"
      Height          =   375
      Left            =   9240
      TabIndex        =   18
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label taxLabel 
      Caption         =   "Tax :-"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   6240
      TabIndex        =   17
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label percent2 
      Caption         =   " %"
      Height          =   375
      Left            =   9240
      TabIndex        =   16
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label discLabel 
      Caption         =   "Discount :-"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   6240
      TabIndex        =   15
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label totalLabel 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8160
      TabIndex        =   14
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Label qtyLabel 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5400
      TabIndex        =   13
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Vendor Name :-"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Vendor Id. :-"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Date :-"
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label SLNO 
      Caption         =   "Sl No. :-"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label PUR_ENTRY 
      Alignment       =   2  'Center
      Caption         =   "PURCHASE ENTRY "
      BeginProperty Font 
         Name            =   "SEYMOUR"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3885
      TabIndex        =   0
      Top             =   600
      Width           =   3975
   End
End
Attribute VB_Name = "PUR_ENTRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 '***************************'
 '**                       **'
 '**  PURCHASE ENTRY FORM  **'
 '**                       **'
 '***************************'
 
 ' VARIABLE DECLEARATION
 
 Option Explicit
 
 Dim S              As String
 Dim sql            As String
 Dim cateGo         As String
 Dim insertSql      As String
 Dim chkBlk         As Boolean
 Public rowCount    As Integer
 Dim numOfRow       As Integer
 Dim purSession     As OraSession
 Dim purDatabase    As OraDatabase
 Dim itemDyna       As OraDynaset
 Dim proDyna        As OraDynaset
 Dim vIdDyna        As OraDynaset
 Dim vNameDyna      As OraDynaset
 Dim slnDyna        As OraDynaset
 Dim checkDate      As Boolean
 Dim dateStr        As String
 Dim updStkStr      As String
 Dim i As Integer, j As Integer, K As Integer
 Dim r As Integer, c As Integer, d As Integer

Private Sub cancelCmd_Click()
  
  Call CLEAR
      
End Sub

Private Sub cancelCmd_GotFocus()
  
  StatusBar1.Panels(3) = "Click for clear the above informatio..."

End Sub
Private Sub ddCombo_GotFocus()
 
 StatusBar1.Panels(3) = "Select the day... "

End Sub
Private Sub ddCombo_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then     'IF KEYPRESS IS ENTER
     mmCombo.SetFocus
  End If

End Sub
Private Sub discText_GotFocus()
  
  StatusBar1.Panels(3) = "Input the discount in percent ..."

End Sub

Private Sub discText_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 Then      ' IF KEYPRESS IS ENTER
   taxText.SetFocus
 End If

End Sub
Private Sub discText_LostFocus()
 
 If Val(discText.Text) > 100 Then 'CHECKING THE DISCOUNT IS GREATER THAN 100
    MsgBox "Discount is too much !", vbCritical
    discText.SetFocus
 End If

End Sub
Private Sub exitCmd_Click()
 
 If MsgBox("DO YOU REALLY WANT TO EXIT", vbYesNo + vbExclamation, "EXIT") = vbYes Then
    Unload Me
 End If

End Sub

Private Sub exitcmd_GotFocus()
  
  StatusBar1.Panels(3) = "Click to exit from this"

End Sub
Private Sub Form_Load()
  
  Call FILLCOMBODATE(ddCombo, mmCombo, yyCombo)   'CALLING AS SUBRUTION COMBODATE
  S$ = "|Slno.|^            Item name                                 |^       Qty    |Price per Unit|^        Total            "
  purGrid.formatString = S$
  rowCount = 1
  
  On Error GoTo ERRORHANDLER

  Set purSession = CreateObject("oracleinprocserver.xorasession")
  Set purDatabase = purSession.OpenDatabase("jms", "hrishi/jms", &H0&)
  Set vIdDyna = purDatabase.CreateDynaset("select vendor_id from vendor_detail", &H4&)
  Set slnDyna = purDatabase.CreateDynaset("select pur_sln.nextval from dual ", &H4&)
  slnTxt = slnDyna.Fields(0) 'INSERT THE VALUE OF PUR_SLN SEQUENCE INOT SLNTXT
         
  While Not vIdDyna.EOF      'ADDING VENDOR ID INTO VENIDCOMBO BOX
     venIdCombo.AddItem vIdDyna.Fields(0)
    vIdDyna.MoveNext
  Wend
  venIdCombo.ListIndex = 0
   
  rowCount = 1              'INITIALINGING THE VALUE OF ROWCOUNT
  purGrid.Col = 1
  purGrid.Row = 1
  purGrid.TextMatrix(1, 1) = rowCount
  rowCount = rowCount + 1
  qtyLabel.Caption = "0.00"
  totalLabel.Caption = "0.00"
 
    
ERRORHANDLER:                   'CODING FOR ERRORHANDLER
   If purSession.LastServerErr = 0 Then
      If purDatabase.LastServerErr = 0 Then
         If Err.Number = 0 Then
         Else
           MsgBox "Vb Error : " & Err.Number & Err.Description
           End
         End If
      Else
        MsgBox " Database  Error : " & purDatabase.LastServerErr & purDatabase.LastServerErrText
        purDatabase.LastServerErrReset
        End
      End If
   Else
     MsgBox "Session Error : " & purSession.LastServerErr & purSession.LastServerErrText
     purSession.LastServerErrReset
     End
   End If
   
End Sub

Private Sub itemCombo_Click()
  
  cateGo = itemCombo.Text
  proNameCombo.CLEAR
  itemCombo.Visible = False
                          'QURRY FOR PRRODUCT NAME
  sql = "select product_name from product_detail where category=  '" & cateGo & "' "
  Set proDyna = purDatabase.CreateDynaset(sql, ORADYN_DEFAULT)
  
  While Not proDyna.EOF
     proNameCombo.AddItem proDyna.Fields(0)
     proDyna.MoveNext
  Wend
                          'POSITIONING AND SIZING THE COMBO BOX CONTROL OVER  GRID CELL
  proNameCombo.Visible = True
  proNameCombo.Top = itemCombo.Top
  proNameCombo.Left = itemCombo.Left
  proNameCombo.Width = itemCombo.Width
  proNameCombo.ListIndex = 0
  
End Sub

Private Sub itemText_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then     'IF KEYPRESS IS ENTER
     If purGrid.Col < 4 Then
        purGrid.Col = purGrid.Col + 1
        purGrid.ColSel = purGrid.Col
        If purGrid.Col = 2 Then
           itemCombo.SetFocus
        Else
           itemText.SetFocus
        End If
     Else
        itemText.Visible = False
        purGrid.Col = 1
        If purGrid.Row = 99 Then
           purGrid.Row = 1
        Else
           purGrid.Row = purGrid.Row + 1
        End If
    purGrid.ColSel = purGrid.Col
    purGrid.RowSel = purGrid.Row
    purGrid.TextMatrix(purGrid.Row, purGrid.Col) = rowCount
    rowCount = rowCount + 1
    purGrid.Col = 2
    purGrid.ColSel = purGrid.Col
    End If
  End If
    
  If KeyAscii = 27 Then
     discText.SetFocus
  End If

End Sub
Private Sub mmCombo_GotFocus()
  
  StatusBar1.Panels(3) = "Select the month ... "

End Sub
Private Sub mmCombo_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     yyCombo.SetFocus
  End If

End Sub
Private Sub OkCmd_Click()

  checkDate = VERIFY_DATE(Val(ddCombo.Text), mmCombo.Text, Val(yyCombo.Text))

  If checkDate = False Then
     MsgBox "Error ! Invalid Date."
     Exit Sub
  End If
 

  On Error GoTo ERRORHANDLER
  
  If OkCmd.Caption = "&Ok" Then
      OkCmd.Caption = "&Save"
      itemCombo.Enabled = False
      itemText.Enabled = False
      proNameCombo.Enabled = False
      itemCombo.CLEAR
      itemText.Text = ""
      proNameCombo.CLEAR
      disAmountText.Text = Val(totalLabel.Caption) * Val(discText.Text) / 100
      taxAmountText.Text = Val(totalLabel.Caption) * Val(taxText.Text) / 100
      totalPurAmountText.Text = Val(totalLabel.Caption) - Val(disAmountText.Text) + Val(taxAmountText.Text)
      For r = 1 To 99
          For c = 1 To 5
              If purGrid.TextMatrix(r, 1) = "" Then
                  chkBlk = True
              End If
              If chkBlk Then
                 purGrid.TextMatrix(r, c) = ""
              End If
          Next
          If purGrid.TextMatrix(r, 2) = "" Or purGrid.TextMatrix(r, 3) = "" Or purGrid.TextMatrix(r, 4) = "" Then
             For d = 1 To 5
                 purGrid.TextMatrix(r, d) = ""
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
         If purGrid.TextMatrix(r, 1) <> "" Then
            chkBlk = False
            numOfRow = numOfRow + 1
         End If
      Next
      If chkBlk = False Then
         dateStr = ddCombo.Text + "-" + mmCombo.Text + "-" + yyCombo
     
         insertSql = "insert into hrishi.purchase_detail values(" & Val(slnTxt.Text) & "," & Val(venIdCombo.Text) & ", '" & (vNameTxt.Text) & "','" & dateStr & "', " & Val(totalLabel.Caption) & "," & Val(discText.Text) & ", " & Val(taxText.Text) & ", " & Val(totalPurAmountText.Text) & " )"
         purDatabase.ExecuteSQL (insertSql)
         
         For r = 1 To numOfRow
             insertSql = " insert into hrishi.purchase_item_detail values( " & Val(slnTxt.Text) & ", '" & (purGrid.TextMatrix(r, 2)) & "', " & Val(purGrid.TextMatrix(r, 3)) & "  , " & Val(purGrid.TextMatrix(r, 4)) & " ,  " & Val(purGrid.TextMatrix(r, 5)) & " )"
             updStkStr = " update HRISHI.STOCK_DETAIL  set TOTAL_PUR_QTY = TOTAL_PUR_QTY +  " & Val(purGrid.TextMatrix(r, 3)) & " , STOCK_IN_HAND  = STOCK_IN_HAND + " & Val(purGrid.TextMatrix(r, 3)) & " ,LAST_MODIFY_DATE =SYSDATE where PRODUCT_NAME ='" & (purGrid.TextMatrix(r, 2)) & "' "
             purDatabase.ExecuteSQL (insertSql)
             purDatabase.ExecuteSQL (updStkStr)
         Next
      
      End If
      If MsgBox("DO YOU WANT TO CONTINUE ...", vbYesNo + vbExclamation, "EXIT") = vbYes Then
         OkCmd.Caption = "&Ok"
         cancelCmd.Enabled = True
         Set slnDyna = purDatabase.CreateDynaset("select pur_sln.nextval from dual ", &H4&)
         slnTxt = slnDyna.Fields(0) 'INSERT THE VALUE OF PUR_SLN SEQUENCE INOT SLNTXT
         'Call CLEAR
      Else
         Unload Me
      End If
      
           
      Exit Sub
  End If
  
ERRORHANDLER:
  If purSession.LastServerErr = 0 Then
     If purDatabase.LastServerErr = 0 Then
       If Err.Number = 0 Then
       
        Else
           MsgBox "Vb Error : " & Err.Number & Err.Description
           Exit Sub
        End If
      Else
         MsgBox " Database  Error : " & purDatabase.LastServerErr & purDatabase.LastServerErrText
         purDatabase.LastServerErrReset
         Exit Sub
      End If
   Else
       MsgBox "Session Error : " & purSession.LastServerErr & purSession.LastServerErrText
       purSession.LastServerErrReset
       Exit Sub
   End If

End Sub
Private Sub CLEAR()
  
  ddCombo.Text = DAY(Date)  'ADDING CURRENT DATE IN THE ABOVE COMBO BOX
  mmCombo.Text = MonthName(MONTH(Date))
  yyCombo.Text = YEAR(Date)
   
  For r = 1 To 99
      For c = 1 To 5
        purGrid.TextMatrix(r, c) = ""
      Next
  Next
  
  purGrid.Col = 1
  purGrid.Row = 1
  purGrid.RowSel = 1
  purGrid.ColSel = 1
  rowCount = 1
  purGrid.TextMatrix(1, 1) = rowCount
  rowCount = rowCount + 1
  slnTxt.SetFocus
  qtyLabel.Caption = "0.00"
  totalLabel.Caption = "0.00"
  discText.Text = "0.00"
  taxText.Text = "0.00"
  disAmountText.Text = ""
  taxAmountText.Text = ""
  totalPurAmountText.Text = ""

End Sub

Private Sub okCmd_GotFocus()
 
 If OkCmd.Caption = "&Ok" Then
    StatusBar1.Panels(3) = "Click for verificaion of above informaion "
  Else
      StatusBar1.Panels(3) = "Click for save the above informaion "
  End If
  
End Sub

Private Sub proNameCombo_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
     purGrid.Col = purGrid.Col + 1
     purGrid.ColSel = purGrid.Col
  End If
  
  If KeyAscii = 27 Then
        discText.SetFocus
  End If

End Sub

Private Sub purGrid_EnterCell()

 If purGrid.Col = 2 And purGrid.Enabled Then
    itemCombo.Visible = True
    itemCombo.Top = purGrid.Top + purGrid.CellTop
    itemCombo.Left = purGrid.Left + purGrid.CellLeft
    itemCombo.Width = purGrid.CellWidth
   
    itemCombo.CLEAR
    Set itemDyna = purDatabase.CreateDynaset("select  distinct (category) from product_detail", &H0&)

    While Not itemDyna.EOF
       itemCombo.AddItem itemDyna.Fields(0)
       itemDyna.MoveNext
    Wend
     
 Else
    If purGrid.Col = 1 Then
       Exit Sub
       End If
       If itemText.Enabled Then
          proNameCombo.Visible = False
          itemText.Text = ""
          itemText.Top = purGrid.Top + purGrid.CellTop
          itemText.Left = purGrid.Left + purGrid.CellLeft
          itemText.Width = purGrid.CellWidth
          itemText.Height = purGrid.CellHeight
          itemText.Visible = True
          itemText.SetFocus
       End If
 End If

End Sub
    
Private Sub purGrid_GotFocus()

  StatusBar1.Panels(3) = "Input the purchase information..."

End Sub

Private Sub purGrid_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
     If purGrid.Col < 4 Then
        purGrid.Col = purGrid.Col + 1
        purGrid.ColSel = purGrid.Col
     Else
        purGrid.Col = 1
        If purGrid.Row = 99 Then
           purGrid.Row = 1
        Else
           purGrid.Row = purGrid.Row + 1
        End If
        
        purGrid.ColSel = purGrid.Col
        purGrid.RowSel = purGrid.Row
        purGrid.TextMatrix(purGrid.Row, purGrid.Col) = rowCount
        rowCount = rowCount + 1
     End If
  End If
    
  If purGrid.Col = 2 Then
     itemCombo.SetFocus
  Else
     itemText.SetFocus
  End If

End Sub

Private Sub purGrid_LeaveCell()

  Dim qty  As Double
  Dim price As Double
  Dim total As Double
  
  If purGrid.Col = 2 Then
     purGrid.Text = proNameCombo.Text
     proNameCombo.Visible = False
  End If
 
 If purGrid.Col > 2 And itemText.Enabled Then
    If itemText.Text = "" Then
       itemText.Text = "0.00"
    End If
    purGrid.Text = itemText.Text
    itemText.Visible = False
 End If
 
 If purGrid.Col = 3 Then
    qtyLabel.Caption = Val(qtyLabel.Caption) + Val(purGrid.TextMatrix(purGrid.Row, 3))
 End If
 
 If purGrid.Col = 4 Then
    qty = Val(purGrid.TextMatrix(purGrid.Row, purGrid.Col - 1))
    price = Val(purGrid.TextMatrix(purGrid.Row, purGrid.Col))
    total = qty * price
    purGrid.TextMatrix(purGrid.Row, purGrid.Col + 1) = total
    totalLabel.Caption = Val(totalLabel.Caption) + total
 End If
  
End Sub

Private Sub slnTxt_GotFocus()

  StatusBar1.Panels(3) = ""

End Sub

Private Sub slnTxt_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     ddCombo.SetFocus
  End If
 
End Sub

Private Sub taxText_GotFocus()
  
  StatusBar1.Panels(3) = "Input the tax in percent..."

End Sub

Private Sub taxText_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then
      cancelCmd.SetFocus
   End If
 
End Sub

Private Sub taxText_LostFocus()
 
 If Val(taxText.Text) > 100 Then
    MsgBox "The tax percentage too much ! ", vbCritical
    taxText.SetFocus
 End If
  
End Sub



Private Sub venIdCombo_Click()
 
 Set vNameDyna = purDatabase.CreateDynaset("select vendor_name from vendor_detail where vendor_id = " & venIdCombo.Text & " ", &H0&)
 vNameTxt.Text = vNameDyna.Fields(0)
 
End Sub

Private Sub venIdCombo_GotFocus()
  
  StatusBar1.Panels(3) = "Select the vendor id no... "

End Sub

Private Sub venIdCombo_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 Then
   purGrid.SetFocus
 End If

End Sub
 
Private Sub vNameTxt_GotFocus()
  
  StatusBar1.Panels(3) = ""

End Sub

Private Sub yyCombo_GotFocus()
  
  StatusBar1.Panels(3) = "Select the year.."

End Sub

Private Sub yyCombo_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 Then
    venIdCombo.SetFocus
 End If

End Sub
