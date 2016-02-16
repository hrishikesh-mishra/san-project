VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PRO_ENTRY 
   Caption         =   "Form1"
   ClientHeight    =   6900
   ClientLeft      =   6120
   ClientTop       =   4545
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   8070
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   6405
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1835
            MinWidth        =   1835
            TextSave        =   "7/2/2005"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1834
            MinWidth        =   1834
            TextSave        =   "10:39 PM"
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
      TabIndex        =   12
      Top             =   2760
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
      TabIndex        =   11
      Top             =   2040
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
      Height          =   615
      Left            =   5400
      TabIndex        =   10
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
      Height          =   615
      Left            =   3480
      TabIndex        =   9
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
      Height          =   615
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   1320
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
      Top             =   3480
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
      Top             =   3480
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
      Top             =   3480
      Width           =   975
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
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Line Line7 
      X1              =   5040
      X2              =   5040
      Y1              =   4440
      Y2              =   5280
   End
   Begin VB.Line Line6 
      X1              =   3000
      X2              =   3000
      Y1              =   4440
      Y2              =   5280
   End
   Begin VB.Line Line5 
      X1              =   1440
      X2              =   2760
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line4 
      X1              =   1080
      X2              =   6960
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line3 
      X1              =   6960
      X2              =   6960
      Y1              =   4440
      Y2              =   5280
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   1080
      X2              =   1080
      Y1              =   4440
      Y2              =   5280
   End
   Begin VB.Line Line1 
      X1              =   1080
      X2              =   6960
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
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3480
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
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2040
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
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label proEntryLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "PRODUCT ENTRY FORM "
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   2865
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Pro_Entry"
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
Dim i As Integer, j As Integer, k As Integer
Dim proSession As OraSession
Dim proDatabase As OraDatabase
Dim catdyna As OraDynaset
Dim flag As Boolean
Dim insertSql As String
Dim checkDate As Boolean
Dim dateStr As String
Dim zeroFlag As Boolean
Private Sub categoryList()
 zeroFlag = False
  Set catdyna = proDatabase.CreateDynaset("Select distinct category from Product_detail ", &H0&)
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

For k = 1900 To 2100    'ADDING YEAR INTO YYCOMBO BOX
 yearCombo.AddItem k
Next

dayCombo.Text = DAY(Date)  'ADDING CURRENT DATE IN THE ABOVE COMBO BOX
monthCombo.Text = MonthName(MONTH(Date))
yearCombo.Text = YEAR(Date)

End Sub
 
Private Sub checkAll()
  If proNameText.Text = "" Then
      MsgBox ("Error ! Product Name is not preset")
      flag = False
      
  ElseIf cateCombo.Text = "" Then
     MsgBox ("Error ! Category is not present ")
     flag = False
     ElseIf manufacNameText.Text = "" Then
    MsgBox ("Error ! Manufacture is not defined ")
    flag = False
  Else
    flag = True
  End If
  
  
  
  
  
End Sub


Private Sub cancelCmd_Click()
Call clear
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

Private Sub clear()
proNameText.Text = ""
cateCombo.Text = ""
manufacNameText.Text = ""
dayCombo.Text = DAY(Date)
monthCombo.Text = MonthName(MONTH(Date))
yearCombo.Text = YEAR(Date)
End Sub

Private Sub monthCombo1_Change()

End Sub

Private Sub extCmd_GotFocus()
StatusBar1.Panels(3) = "Exit...."
End Sub

Private Sub Form_Load()


 Call comboDate  'CALLING AS SUBRUTION COMBODATE
   On Error GoTo errorhandler
 
 Set proSession = CreateObject("oracleinprocserver.xorasession")
 Set proDatabase = proSession.OpenDatabase("jms", "hrishi/jms", &H0&)
 
 Call categoryList
 
 
errorhandler:                     'CODING FOR ERRORHANDLER
   If proSession.LastServerErr = 0 Then
      If proDatabase.LastServerErr = 0 Then
         If Err.Number = 0 Then
         Else
           MsgBox "Vb Error : " & Err.Number & Err.Description
           End
           End If
      Else
        MsgBox " Database  Error : " & proDatabase.LastServerErr & proDatabase.LastServerErrText
        proDatabase.LastServerErrReset
        End
      End If
   Else
     MsgBox "Session Error : " & proDatabase.LastServerErr & proDatabase.LastServerErrText
     proDatabase.LastServerErrReset
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
 Call checkAll

 checkDate = VERIFY_DATE(Val(dayCombo.Text), monthCombo.Text, _
                         Val(yearCombo.Text))
  If checkDate = False Then
    MsgBox "Error ! Not Valid date"
 End If
   
                         
 On Error GoTo errorhandler
 
 If flag And checkDate Then
     dateStr = dayCombo.Text + "-" + monthCombo.Text + "-" + _
     yearCombo
     
     insertSql = "insert into HRISHI.PRODUCT_DETAIL values( '" & _
     UCase(proNameText.Text) & "', '" & UCase(cateCombo.Text) & _
     "','" & UCase(manufacNameText.Text) & "','" & dateStr & "' )"
     
     proDatabase.ExecuteSQL (insertSql)
 
errorhandler:                       'CODING FOR ERRORHANDLER
   If proSession.LastServerErr = 0 Then
      If proDatabase.LastServerErr = 0 Then
         If Err.Number = 0 Then
         If MsgBox("Sucess ! Do you want more ...", vbYesNo + _
            vbExclamation, "EXIT") = vbYes Then
            proNameText.SetFocus
            Call categoryList
          Else
           End
         End If
         Else
           MsgBox "Vb Error : " & Err.Number & Err.Description
           
         End If
      Else
        MsgBox " Database  Error : " & proDatabase.LastServerErr & proDatabase.LastServerErrText
        proDatabase.LastServerErrReset
      
      End If
   Else
     MsgBox "Session Error : " & proDatabase.LastServerErr & proDatabase.LastServerErrText
     proDatabase.LastServerErrReset
     
   End If
   End If
   
End Sub

Private Sub saveCmd_GotFocus()
StatusBar1.Panels(3) = "Save the above information.."
End Sub

Private Sub yearCombo_GotFocus()
StatusBar1.Panels(3) = "Select the year..."
End Sub

Private Sub yearCombo_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
  saveCmd.SetFocus
End If

End Sub
