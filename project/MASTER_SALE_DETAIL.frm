VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form MASTER_SALE_DETAIL 
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14610
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5280
   ScaleWidth      =   14610
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
      Height          =   495
      Left            =   6345
      TabIndex        =   2
      Top             =   4560
      Width           =   1815
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "MASTER_SALE_DETAIL.frx":0000
      Height          =   3375
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Sale Detail"
      Top             =   1080
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   5953
      _Version        =   393216
      ForeColor       =   2645968
      Cols            =   13
      FixedCols       =   0
      ForeColorFixed  =   65535
      AllowUserResizing=   3
      DataMember      =   "masterSaleDetail"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   2
      _Band(0).Cols   =   9
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   9
      _Band(0)._MapCol(0)._Name=   "SALE_SLN"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "TYPE_OF_CUST"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "CUST_ID"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "CUST_NAME"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(4)._Name=   "SALE_DATE"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(5)._Name=   "TOTAL_SALE_AMOUNT"
      _Band(0)._MapCol(5)._RSIndex=   5
      _Band(0)._MapCol(6)._Name=   "DISC"
      _Band(0)._MapCol(6)._RSIndex=   6
      _Band(0)._MapCol(7)._Name=   "TAX"
      _Band(0)._MapCol(7)._RSIndex=   7
      _Band(0)._MapCol(8)._Name=   "TOTAL_AMOUNT"
      _Band(0)._MapCol(8)._RSIndex=   8
      _Band(1).BandIndent=   1
      _Band(1).Cols   =   4
      _Band(1).GridLinesBand=   1
      _Band(1).TextStyleBand=   0
      _Band(1).TextStyleHeader=   0
      _Band(1)._ParentBand=   0
      _Band(1)._NumMapCols=   5
      _Band(1)._MapCol(0)._Name=   "ITEM_SLN"
      _Band(1)._MapCol(0)._RSIndex=   0
      _Band(1)._MapCol(0)._Hidden=   -1  'True
      _Band(1)._MapCol(1)._Name=   "ITEM_NAME"
      _Band(1)._MapCol(1)._RSIndex=   1
      _Band(1)._MapCol(2)._Name=   "QTY"
      _Band(1)._MapCol(2)._RSIndex=   2
      _Band(1)._MapCol(3)._Name=   "PPU"
      _Band(1)._MapCol(3)._RSIndex=   3
      _Band(1)._MapCol(4)._Name=   "TOTAL"
      _Band(1)._MapCol(4)._RSIndex=   4
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   14520
      X2              =   8400
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   14520
      X2              =   14520
      Y1              =   4560
      Y2              =   960
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   120
      X2              =   14520
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   8400
      X2              =   8400
      Y1              =   4560
      Y2              =   5160
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   8400
      X2              =   6120
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   6120
      X2              =   6120
      Y1              =   4560
      Y2              =   5160
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   120
      X2              =   6120
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   960
      Y2              =   4560
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   555
      Left            =   0
      Picture         =   "MASTER_SALE_DETAIL.frx":0020
      ToolTipText     =   "San's Master Sale Detail"
      Top             =   0
      Width           =   2970
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "MASTER SALE DETAIL "
      BeginProperty Font 
         Name            =   "TIMORA"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   5078
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "MASTER_SALE_DETAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub
