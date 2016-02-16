VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form MASTER_PURCHASE_DETAIL 
   Caption         =   "MASTER DETAIL [PURCHASE DETAIL]"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14460
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   14460
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "MASTER_PURCHASE_DETAIL.frx":0000
      Height          =   4575
      Left            =   480
      TabIndex        =   0
      ToolTipText     =   "Purchase Detail."
      Top             =   960
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   8070
      _Version        =   393216
      ForeColor       =   13454032
      Cols            =   12
      FixedCols       =   0
      ForeColorFixed  =   4194432
      AllowUserResizing=   3
      DataMember      =   "masterPurDetail"
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   2
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   8
      _Band(0)._MapCol(0)._Name=   "PRU_SLN"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "VENDOR_ID"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "VENDOR_NAME"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "PUR_DATE"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(4)._Name=   "TOTAL_PUR_AMOUNT"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(5)._Name=   "DISC"
      _Band(0)._MapCol(5)._RSIndex=   5
      _Band(0)._MapCol(6)._Name=   "TAX"
      _Band(0)._MapCol(6)._RSIndex=   6
      _Band(0)._MapCol(7)._Name=   "TOTAL_AMOUNT"
      _Band(0)._MapCol(7)._RSIndex=   7
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
      Left            =   5963
      TabIndex        =   1
      Top             =   5760
      Width           =   2535
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   0
      Picture         =   "MASTER_PURCHASE_DETAIL.frx":0020
      ToolTipText     =   "San's Master Detail of Purchase."
      Top             =   0
      Width           =   3000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "MASTER PURCHSE DETAIL"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   360
      Width           =   6015
   End
End
Attribute VB_Name = "MASTER_PURCHASE_DETAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

