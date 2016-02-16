VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form MASTER_CURRENT_EMP_DETAIL 
   Caption         =   " MASTER CURRENT EMPLOYEE DETAIL"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13455
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   13455
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "MASTER_CURRENT_EMP_DETAIL.frx":0000
      Height          =   4575
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Detail of Current Employee."
      Top             =   1200
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   8070
      _Version        =   393216
      ForeColor       =   16711680
      Cols            =   21
      FixedCols       =   0
      ForeColorFixed  =   16711680
      AllowUserResizing=   3
      DataMember      =   "masterCurEmpDetail"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial CE"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   2
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   8
      _Band(0)._MapCol(0)._Name=   "EID"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "ENAME"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "SEX"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "AGE"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(4)._Name=   "ADDRESS"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(5)._Name=   "PHNO"
      _Band(0)._MapCol(5)._RSIndex=   5
      _Band(0)._MapCol(6)._Name=   "DESTINATION"
      _Band(0)._MapCol(6)._RSIndex=   6
      _Band(0)._MapCol(7)._Name=   "JOIN_DATE"
      _Band(0)._MapCol(7)._RSIndex=   7
      _Band(1).BandIndent=   1
      _Band(1).Cols   =   13
      _Band(1).GridLinesBand=   1
      _Band(1).TextStyleBand=   0
      _Band(1).TextStyleHeader=   0
      _Band(1)._ParentBand=   0
      _Band(1)._NumMapCols=   15
      _Band(1)._MapCol(0)._Name=   "EID"
      _Band(1)._MapCol(0)._RSIndex=   0
      _Band(1)._MapCol(0)._Hidden=   -1  'True
      _Band(1)._MapCol(1)._Name=   "ENAME"
      _Band(1)._MapCol(1)._RSIndex=   1
      _Band(1)._MapCol(1)._Hidden=   -1  'True
      _Band(1)._MapCol(2)._Name=   "SAL_FOR_MONTH"
      _Band(1)._MapCol(2)._RSIndex=   2
      _Band(1)._MapCol(3)._Name=   "BASIC"
      _Band(1)._MapCol(3)._RSIndex=   3
      _Band(1)._MapCol(4)._Name=   "HRA"
      _Band(1)._MapCol(4)._RSIndex=   4
      _Band(1)._MapCol(5)._Name=   "DA"
      _Band(1)._MapCol(5)._RSIndex=   5
      _Band(1)._MapCol(6)._Name=   "TA"
      _Band(1)._MapCol(6)._RSIndex=   6
      _Band(1)._MapCol(7)._Name=   "DEDUCTION"
      _Band(1)._MapCol(7)._RSIndex=   7
      _Band(1)._MapCol(8)._Name=   "TAX"
      _Band(1)._MapCol(8)._RSIndex=   8
      _Band(1)._MapCol(9)._Name=   "SPECIAL_PAY"
      _Band(1)._MapCol(9)._RSIndex=   9
      _Band(1)._MapCol(10)._Name=   "FESTIVAL_PAY"
      _Band(1)._MapCol(10)._RSIndex=   10
      _Band(1)._MapCol(11)._Name=   "GROSS_SAL"
      _Band(1)._MapCol(11)._RSIndex=   11
      _Band(1)._MapCol(12)._Name=   "TOTAL_DEDUCTION"
      _Band(1)._MapCol(12)._RSIndex=   12
      _Band(1)._MapCol(13)._Name=   "NET_SAL"
      _Band(1)._MapCol(13)._RSIndex=   13
      _Band(1)._MapCol(14)._Name=   "ENTRY_DATE"
      _Band(1)._MapCol(14)._RSIndex=   14
   End
   Begin VB.CommandButton Command1 
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
      Left            =   5460
      TabIndex        =   0
      Top             =   5760
      Width           =   2535
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   0
      Picture         =   "MASTER_CURRENT_EMP_DETAIL.frx":0020
      ToolTipText     =   "San's Master Detail of Current Employee."
      Top             =   0
      Width           =   3000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "MASTER CURRENT EMPLOYEE DETAIL"
      BeginProperty Font 
         Name            =   "Airfoil Script SSi"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   600
      Width           =   9015
   End
End
Attribute VB_Name = "MASTER_CURRENT_EMP_DETAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub
