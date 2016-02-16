VERSION 5.00
Begin VB.Form STOCK_DETAIL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STOCK DETAIL:-"
   ClientHeight    =   7920
   ClientLeft      =   3705
   ClientTop       =   2580
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   9180
   Begin VB.CommandButton extCmd 
      BackColor       =   &H80000004&
      Caption         =   "&EXIT"
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
      Left            =   3750
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Exit "
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton mvLastCmd 
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   5160
      Picture         =   "STOCK_DETAIL.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Move Last."
      Top             =   6600
      Width           =   495
   End
   Begin VB.CommandButton mvNextCmd 
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   4680
      Picture         =   "STOCK_DETAIL.frx":023C
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Move Next."
      Top             =   6600
      Width           =   495
   End
   Begin VB.CommandButton mvPreCmd 
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   4200
      Picture         =   "STOCK_DETAIL.frx":0389
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Move Previous."
      Top             =   6600
      Width           =   495
   End
   Begin VB.CommandButton mvFirstCmd 
      BackColor       =   &H000000FF&
      Height          =   375
      Left            =   3720
      MaskColor       =   &H00C00000&
      Picture         =   "STOCK_DETAIL.frx":04D8
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Move First."
      Top             =   6600
      Width           =   495
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   120
      Picture         =   "STOCK_DETAIL.frx":071B
      ToolTipText     =   "San's stock Detail"
      Top             =   0
      Width           =   3000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00800080&
      BorderWidth     =   2
      X1              =   120
      X2              =   9360
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800080&
      BorderWidth     =   2
      X1              =   120
      X2              =   9360
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800080&
      BorderWidth     =   2
      X1              =   120
      X2              =   9240
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label proNameLabel2 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3240
      TabIndex        =   12
      Top             =   1560
      Width           =   4215
   End
   Begin VB.Label categoryLabel2 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3240
      TabIndex        =   11
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Label totalPurLabel2 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3240
      TabIndex        =   10
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label stkInHandLabel2 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3240
      TabIndex        =   9
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label totalSaleLabel2 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3240
      TabIndex        =   8
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label lastTranLabel2 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3240
      TabIndex        =   7
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Label proNameLabel1 
      Caption         =   "PRODUCT NAME :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label categoryLabel1 
      Caption         =   "CATEGORY :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2430
      Width           =   1935
   End
   Begin VB.Label toalPurLabel1 
      Caption         =   "TOTAL PURCHASE :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3165
      Width           =   2415
   End
   Begin VB.Label totalSaleLabel1 
      Caption         =   "TOTAL SALE :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4035
      Width           =   1695
   End
   Begin VB.Label stkInHandLabel1 
      Caption         =   "STOCK IN HAND :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   4890
      Width           =   2055
   End
   Begin VB.Label lastTranLabel1 
      Caption         =   "LAST TRANSACTION ON:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   5760
      Width           =   2895
   End
   Begin VB.Label stkLabel 
      Alignment       =   2  'Center
      Caption         =   "STOCK DETAIL "
      BeginProperty Font 
         Name            =   "LOUISETTA"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Left            =   3015
      TabIndex        =   0
      Top             =   600
      Width           =   3000
   End
End
Attribute VB_Name = "STOCK_DETAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************'
'**                   **'
'** STOCK DETAIL FORM **'
'**                   **'
'***********************'

'VARIABLE DECLARATION

Option Explicit

Dim stkSession  As OraSession
Dim stkDatabase As OraDatabase
Dim stkdyn      As OraDynaset

Private Sub FillWithValue()

 If stkdyn.EOF Then
    Exit Sub
 End If
 proNameLabel2.Caption = stkdyn.Fields("PRODUCT_NAME").Value
 categoryLabel2.Caption = stkdyn.Fields("CATEGORY").Value
 totalPurLabel2.Caption = stkdyn.Fields("TOTAL_PUR_QTY").Value
 totalSaleLabel2.Caption = stkdyn.Fields("TOTAL_SALE_QTY").Value
 stkInHandLabel2.Caption = stkdyn.Fields("STOCK_IN_HAND").Value
 lastTranLabel2.Caption = stkdyn.Fields("LAST_MODIFY_DATE").Value

End Sub

Private Sub extCmd_Click()

    If MsgBox("DO YOU WANT TO EXIT .. ", vbExclamation + vbYesNo, "EXIT :") = vbYes Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()

   On Error GoTo ERRH
    Set stkSession = CreateObject("oracleinprocserver.xorasession")
    Set stkDatabase = stkSession.OpenDatabase("jms", "hrishi/jms", &H0&)
    Set stkdyn = stkDatabase.CreateDynaset("Select * from HRISHI.STOCK_DETAIL", &H4&)

    If stkdyn.EOF And stkdyn.BOF Then
        MsgBox "ERROR ! THERE IS NOTHING IN THE DATABASE ", vbCritical, "ERROR:"
        Exit Sub
    Else
        stkdyn.MoveFirst
        Call FillWithValue
    End If
ERRH:
    If stkSession.LastServerErr = 0 Then
        If stkDatabase.LastServerErr = 0 Then
            If Err.Number = 0 Then
            Else
              MsgBox "VB ERROR : " & Err.Number & " :: " & Err.Description, vbCritical, "VB Error :"
              Exit Sub
            End If
        Else
            MsgBox "DATABASE ERROR : " & stkDatabase.LastServerErr & stkDatabase.LastServerErrText, vbCritical, "DATABASE Error:"
            stkDatabase.LastServerErrReset
            Exit Sub
        End If
    Else
        MsgBox "SESSION ERROR :" & stkSession.LastServerErr & stkSession.LastServerErrText, vbCritical, "SESSION Error :"
        stkSession.LastServerErrReset
        Exit Sub
    End If

End Sub



Private Sub mvFirstCmd_Click()
   
   mvFirstCmd.BackColor = vbGreen
   stkdyn.MoveFirst
   Call FillWithValue

End Sub

Private Sub mvLastCmd_Click()

    mvLastCmd.BackColor = vbGreen
    stkdyn.MoveLast
    Call FillWithValue

End Sub

Private Sub mvNextCmd_Click()
    
    mvNextCmd.BackColor = vbGreen
    If stkdyn.EOF And stkdyn.BOF Then
        Exit Sub
    End If

    stkdyn.MoveNext
    If stkdyn.EOF Then
        MsgBox "Nothing is in After..", vbInformation, "Database Information :"
        stkdyn.MoveLast
    Else
        Call FillWithValue
    End If

End Sub

Private Sub mvPreCmd_Click()
    
    mvPreCmd.BackColor = vbGreen
    If stkdyn.BOF And stkdyn.EOF Then
        Exit Sub
    End If
    
    stkdyn.MovePrevious
    If stkdyn.BOF Then
        MsgBox "Nothing is in Before ..", vbInformation, "Database Information"
        stkdyn.MoveFirst
    Else
        Call FillWithValue
    End If

End Sub

