VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form MASTER_RELIVED_EMP_DETAIL 
   Caption         =   "MASTER RELIEVED EMPLOYEE DETAIL "
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   8685
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "MASTER_RELIVED_EMP_DETAIL.frx":0000
      Height          =   3855
      Left            =   480
      TabIndex        =   2
      ToolTipText     =   "RELIEVED EMPLOYEE DETAIL "
      Top             =   1680
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6800
      _Version        =   393216
      ForeColor       =   16711935
      Cols            =   6
      FixedCols       =   0
      ForeColorFixed  =   16744576
      AllowUserResizing=   3
      DataMember      =   "materRelievedEmpDetail"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman CE"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0)._NumMapCols=   6
      _Band(0)._MapCol(0)._Name=   "RELSLN"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "EID"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "ENAME"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "DESTINATION"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(4)._Name=   "DOJ"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(5)._Name=   "DOR"
      _Band(0)._MapCol(5)._RSIndex=   5
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "ESP"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   240
      X2              =   8640
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   8640
      X2              =   8640
      Y1              =   5640
      Y2              =   1440
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   5520
      X2              =   8640
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   5520
      X2              =   5520
      Y1              =   5640
      Y2              =   6120
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   2760
      X2              =   5520
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   2760
      X2              =   2760
      Y1              =   5640
      Y2              =   6120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   2760
      X2              =   240
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   240
      X2              =   240
      Y1              =   1440
      Y2              =   5640
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   555
      Left            =   0
      Picture         =   "MASTER_RELIVED_EMP_DETAIL.frx":0020
      ToolTipText     =   "San's Master Detail of RELIVED EMPLOYEE."
      Top             =   0
      Width           =   2970
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "MASTER RELIVED EMPLOYEE DETAIL"
      BeginProperty Font 
         Name            =   "ANDREIAN"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   8175
   End
End
Attribute VB_Name = "MASTER_RELIVED_EMP_DETAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me

End Sub

