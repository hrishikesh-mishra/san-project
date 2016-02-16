VERSION 5.00
Object = "{92096210-97DF-11CF-9F27-02608C4BF3B5}#1.0#0"; "ORADC.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form PRODUCT_ENTRY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PRODUCT FORM [ENTRY]"
   ClientHeight    =   7890
   ClientLeft      =   4590
   ClientTop       =   1710
   ClientWidth     =   8295
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   8295
   Begin MSDBGrid.DBGrid proDBGrid 
      Bindings        =   "PRODUCT_ENTRY.frx":0000
      Height          =   1575
      Left            =   0
      OleObjectBlob   =   "PRODUCT_ENTRY.frx":001C
      TabIndex        =   15
      ToolTipText     =   "Product Detail."
      Top             =   5760
      Width           =   8295
   End
   Begin ORADCLibCtl.ORADC ORADC_PRODUCT 
      Height          =   375
      Left            =   2520
      Top             =   5160
      Visible         =   0   'False
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
      _ExtentY        =   661
      _StockProps     =   207
      Caption         =   "PRODUCT DETAIL"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.74
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DatabaseName    =   "jms"
      Connect         =   "hrishi/jms"
      RecordSource    =   "select * from product_detail"
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   7395
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1835
            MinWidth        =   1835
            TextSave        =   "9/23/2005"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1834
            MinWidth        =   1834
            TextSave        =   "2:03 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   11536
            MinWidth        =   11536
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Left            =   3000
      TabIndex        =   3
      ToolTipText     =   "Enter the manufacture of Product."
      Top             =   2880
      Width           =   4575
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
      Left            =   3000
      TabIndex        =   2
      ToolTipText     =   "Select or Enter the Product Category."
      Top             =   2160
      Width           =   4455
   End
   Begin VB.CommandButton extCmd 
      Caption         =   "&Exit"
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
      Left            =   5400
      TabIndex        =   9
      ToolTipText     =   "Exit From This."
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton saveCmd 
      Caption         =   "&Save "
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
      Left            =   1440
      TabIndex        =   8
      ToolTipText     =   "Save the Information."
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cancelCmd 
      Caption         =   "&Cancel"
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Clear the Entered Data."
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox proNameText 
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
      Left            =   3000
      TabIndex        =   12
      ToolTipText     =   "Enter the Product Name."
      Top             =   1440
      Width           =   4455
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
      Left            =   5640
      TabIndex        =   6
      ToolTipText     =   "Select the Year."
      Top             =   3600
      Width           =   1215
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
      Left            =   3960
      TabIndex        =   5
      ToolTipText     =   "Select the Month."
      Top             =   3600
      Width           =   1695
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
      Left            =   3000
      TabIndex        =   4
      ToolTipText     =   "Select the Day."
      Top             =   3600
      Width           =   975
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   5880
      X2              =   5880
      Y1              =   5160
      Y2              =   5640
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   2400
      X2              =   5880
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   2400
      X2              =   2400
      Y1              =   5160
      Y2              =   5640
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   8280
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   0
      Picture         =   "PRODUCT_ENTRY.frx":0D71
      ToolTipText     =   "San's Product Entry Form."
      Top             =   0
      Width           =   3000
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   8280
      Y1              =   1080
      Y2              =   1080
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
      TabIndex        =   13
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   5040
      X2              =   5040
      Y1              =   4440
      Y2              =   5160
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   3000
      X2              =   3000
      Y1              =   4440
      Y2              =   5160
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   1200
      X2              =   6840
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   6840
      X2              =   6840
      Y1              =   4440
      Y2              =   5160
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      Index           =   0
      X1              =   1200
      X2              =   1200
      Y1              =   4440
      Y2              =   5160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   1200
      X2              =   6840
      Y1              =   4440
      Y2              =   4440
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
      TabIndex        =   11
      Top             =   3600
      Width           =   1335
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
      TabIndex        =   10
      Top             =   2160
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
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label proEntryLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "PRODUCT ENTRY FORM "
      BeginProperty Font 
         Name            =   "GLENNA"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   600
      Width           =   3375
   End
End
Attribute VB_Name = "PRODUCT_ENTRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************'
'**                       **'
'**  PRODUCT ENTRY FORM   **'
'**                       **'
'***************************'


'VARIBLE DECLEARATION
Option Explicit

Dim i As Integer, j As Integer, K As Integer
Dim proSession      As OraSession
Dim stockSession    As OraSession
Dim stockDatabase   As OraDatabase
Dim proDatabase     As OraDatabase
Dim catdyna         As OraDynaset
Dim insertSql       As String
Dim stockSql        As String
Dim dateStr         As String
Dim zeroFlag        As Boolean
Dim flag            As Boolean
Dim checkDate       As Boolean
Dim stockflag       As Boolean

Private Sub categoryList()
 
 zeroFlag = False
 Set catdyna = proDatabase.CreateDynaset("Select distinct category from product_detail", &H0&)
 
 While Not catdyna.EOF
     cateCombo.AddItem catdyna.Fields(0)
     catdyna.MoveNext
     zeroFlag = True
  Wend
 
 If zeroFlag Then
   cateCombo.ListIndex = 0
 End If
 
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
 
Private Sub checkAll()
  
  If proNameText.Text = "" Then
      MsgBox ("Error ! Product Name is not preset."), vbCritical, "Error"
      flag = True
  ElseIf cateCombo.Text = "" Then
      MsgBox ("Error ! Category is not present. "), vbCritical, "Error"
      flag = True
  ElseIf manufacNameText.Text = "" Then
      MsgBox ("Error ! Manufacture is not defined "), vbCritical, "Error"
      flag = True
  Else
      flag = False
      End If
  
End Sub

Private Sub Stock_Insert()

 Set stockSession = CreateObject("oracleinprocserver.xorasession")
 Set stockDatabase = stockSession.OpenDatabase("jms", "hrishi/jms", &H0&)
 
 stockSql = " insert into HRISHI.STOCK_DETAIL values  ('" & _
              UCase(proNameText.Text) & "','" & UCase(cateCombo.Text) _
              & "',0,0,0,SYSDATE)"
 stockDatabase.ExecuteSQL (stockSql)
 
 
 If stockSession.LastServerErr = 0 Then
   If stockDatabase.LastServerErr = 0 Then
     If Err.Number = 0 Then
        stockflag = False
     Else
       MsgBox " VB ERROR: " & Err.Number & "::" & Err.Description, vbCritical, "VB Error"
       Exit Sub
     End If
   Else
       MsgBox "DATABASE ERROR : " & stockDatabase.LastServerErr & stockDatabase.LastServerErrText, vbCritical, "DATABASE Error"
       stockSession.LastServerErrReset
       Exit Sub
    End If
 Else
   MsgBox "SESSION ERROR : " & stockSession.LastServerErr & stockSession.LastServerErrText, vbCritical, "SESSION Error"
   stockSession.LastServerErrReset
   Exit Sub
 End If
 
End Sub
Private Sub cancelCmd_Click()

  Call CLEAR
  Call categoryList

End Sub

Private Sub cancelCmd_GotFocus()

   StatusBar1.Panels(3) = "Clear the above information..."

End Sub

Private Sub cateCombo_GotFocus()

  StatusBar1.Panels(3) = "Select the Category..."

End Sub

Private Sub cateCombo_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
     manufacNameText.SetFocus
  End If

End Sub

Private Sub dayCombo_GotFocus()

  StatusBar1.Panels(3) = "Select the date ..."

End Sub

Private Sub dayCombo_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
     monthCombo.SetFocus
  End If

End Sub

Private Sub extCmd_Click()

  If MsgBox("DO YOU REALLY WANT TO EXIT", vbYesNo + vbExclamation, "EXIT") = vbYes Then
      Unload Me
  End If

End Sub

Private Sub CLEAR()
    
    proNameText.Text = ""
    cateCombo.Text = ""
    manufacNameText.Text = ""
    dayCombo.Text = DAY(Date)
    monthCombo.Text = MonthName(MONTH(Date))
    yearCombo.Text = YEAR(Date)

End Sub

Private Sub extCmd_GotFocus()
  
  StatusBar1.Panels(3) = "Exit...."

End Sub

Private Sub Form_Load()
 
 Call comboDate  'CALLING AS SUBRUTION COMBODATE

 On Error GoTo ERRORHANDLER

  Set proSession = CreateObject("oracleinprocserver.xorasession")
  Set proDatabase = proSession.OpenDatabase("jms", "hrishi/jms", &H0&)

  Call categoryList

   
ERRORHANDLER:                     'CODING FOR ERRORHANDLER
    If proSession.LastServerErr = 0 Then
      If proDatabase.LastServerErr = 0 Then
         If Err.Number = 0 Then
         Else
           MsgBox "Vb Error : " & Err.Number & Err.Description, vbCritical, "Error"
           End
           End If
      Else
        MsgBox " Database  Error : " & proDatabase.LastServerErr & proDatabase.LastServerErrText, vbCritical, "Error"
        proDatabase.LastServerErrReset
        End
      End If
   Else
     MsgBox "Session Error : " & proSession.LastServerErr & proSession.LastServerErrText, vbCritical, "Error"
     proSession.LastServerErrReset
     End
   End If
  
End Sub

Private Sub manufacNameText_GotFocus()
  
  StatusBar1.Panels(3) = "Enter the Manufacture..."

End Sub

Private Sub manufacNameText_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then
      dayCombo.SetFocus
   End If

End Sub

Private Sub monthCombo_GotFocus()
   
   StatusBar1.Panels(3) = "Select the month..."

End Sub

Private Sub monthCombo_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     yearCombo.SetFocus
  End If
 
End Sub

Private Sub proNameText_GotFocus()
  
  StatusBar1.Panels(3) = "Enter the Product name ..."

End Sub

Private Sub proNameText_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     cateCombo.SetFocus
  End If

End Sub

Private Sub saveCmd_Click()

 On Error GoTo ERRORHANDLER
 
 Call checkAll
 checkDate = VERIFY_DATE(Val(dayCombo.Text), monthCombo.Text, Val(yearCombo.Text))
  
 If checkDate = False Then
     MsgBox "Error ! Invalid date", vbCritical, "Error"
     Exit Sub
 ElseIf flag Then
     Exit Sub
 End If
 
 dateStr = dayCombo.Text + "-" + monthCombo.Text + "-" + yearCombo.Text
 insertSql = "insert into HRISHI.PRODUCT_DETAIL values( '" & _
     UCase(proNameText.Text) & "', '" & UCase(cateCombo.Text) & _
     "','" & UCase(manufacNameText.Text) & "','" & dateStr & "' )"
      
 proDatabase.ExecuteSQL (insertSql)
 

           

ERRORHANDLER:                         'CODING FOR ERRORHANDLER
   If proSession.LastServerErr = 0 Then
      If proDatabase.LastServerErr = 0 Then
         If Err.Number = 0 Then
            If MsgBox("Sucess ! Do you want more ...", vbInformation + vbYesNo, "EXIT") = vbYes Then
              ORADC_PRODUCT.Refresh
               Call Stock_Insert
               proNameText.SetFocus
               
               Call categoryList
            Else
              Unload Me
            End If
         Else
            MsgBox "Vb Error : " & Err.Number & " :: " & Err.Description, vbCritical, "Error"
            Exit Sub
         End If
      Else
         MsgBox " Database  Error : " & proDatabase.LastServerErr & proDatabase.LastServerErrText, vbCritical, "Error"
         proDatabase.LastServerErrReset
         Exit Sub
      End If
   Else
     MsgBox "Session Error : " & proSession.LastServerErr & proSession.LastServerErrText, vbCritical, "Error"
     proSession.LastServerErrReset
     Exit Sub
   End If
   
   
End Sub

Private Sub saveCmd_GotFocus()
  
  StatusBar1.Panels(3) = "Save the above information.."

End Sub

Private Sub yearcombo_GotFocus()
   
   StatusBar1.Panels(3) = "Select the year..."

End Sub

Private Sub yearcombo_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then
     saveCmd.SetFocus
  End If

End Sub
