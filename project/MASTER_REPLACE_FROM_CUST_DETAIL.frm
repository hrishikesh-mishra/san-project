VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form MASTER_REPLACE_FROM_CUST_DETAIL 
   Caption         =   "MASTER REPLACE FROM CUSTOMER DETAIL"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   10755
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "MASTER_REPLACE_FROM_CUST_DETAIL.frx":0000
      Height          =   4575
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Replace from customer detail."
      Top             =   1200
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8070
      _Version        =   393216
      ForeColor       =   4227327
      Cols            =   8
      FixedCols       =   0
      ForeColorFixed  =   128
      AllowUserResizing=   3
      DataMember      =   "masterCustRplDetail"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   2
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   5
      _Band(0)._MapCol(0)._Name=   "RPL_SLN"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "REPLACER_TYPE"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "REPLACER_ID"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "REPLACER_NAME"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(4)._Name=   "REPLACE_DATE"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(1).BandIndent=   1
      _Band(1).Cols   =   3
      _Band(1).GridLinesBand=   1
      _Band(1).TextStyleBand=   0
      _Band(1).TextStyleHeader=   0
      _Band(1)._ParentBand=   0
      _Band(1)._NumMapCols=   4
      _Band(1)._MapCol(0)._Name=   "RPL_SLN"
      _Band(1)._MapCol(0)._RSIndex=   0
      _Band(1)._MapCol(0)._Hidden=   -1  'True
      _Band(1)._MapCol(1)._Name=   "ITEM_NAME"
      _Band(1)._MapCol(1)._RSIndex=   1
      _Band(1)._MapCol(2)._Name=   "QTY"
      _Band(1)._MapCol(2)._RSIndex=   2
      _Band(1)._MapCol(3)._Name=   "DESCRIP"
      _Band(1)._MapCol(3)._RSIndex=   3
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
      Left            =   3960
      TabIndex        =   0
      Top             =   5880
      Width           =   2535
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   6600
      X2              =   6600
      Y1              =   5880
      Y2              =   6360
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   3840
      X2              =   6600
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   3840
      X2              =   3840
      Y1              =   5880
      Y2              =   6360
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   120
      X2              =   3840
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   1080
      Y2              =   5880
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   6600
      X2              =   10560
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   10560
      X2              =   10560
      Y1              =   1080
      Y2              =   5880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   120
      X2              =   10560
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   555
      Left            =   0
      Picture         =   "MASTER_REPLACE_FROM_CUST_DETAIL.frx":0020
      ToolTipText     =   "San's Master Replace From Customer Detail"
      Top             =   0
      Width           =   2970
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "MASTER REPLACE FROM CUSTOMER DETAIL"
      BeginProperty Font 
         Name            =   "PAULINE"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   7335
   End
End
Attribute VB_Name = "MASTER_REPLACE_FROM_CUST_DETAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

