VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form MASTER_ONSPOT_REPLACE_DETAIL 
   Caption         =   "MASTER ON SPOT REPLACEMENT DETAIL "
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6480
   ScaleWidth      =   11895
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
      Left            =   4680
      TabIndex        =   3
      Top             =   5880
      Width           =   2535
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "MASTER_ONSPOT_REPLACE_DETAIL.frx":0000
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7011
      _Version        =   393216
      ForeColor       =   8388863
      Rows            =   3
      Cols            =   10
      FixedCols       =   0
      ForeColorFixed  =   16711935
      AllowUserResizing=   3
      DataMember      =   "masterOnSpotDetail"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   3
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   5
      _Band(0)._MapCol(0)._Name=   "ONSPT_SLN"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "CUST_TYPE"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "CUST_ID"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "CUST_NAME"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(4)._Name=   "ONSRPL_DATE"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(1).BandIndent=   1
      _Band(1).Cols   =   3
      _Band(1).GridLinesBand=   1
      _Band(1).TextStyleBand=   0
      _Band(1).TextStyleHeader=   0
      _Band(1)._ParentBand=   0
      _Band(1)._NumMapCols=   4
      _Band(1)._MapCol(0)._Name=   "ONSPT_SLN"
      _Band(1)._MapCol(0)._RSIndex=   0
      _Band(1)._MapCol(0)._Hidden=   -1  'True
      _Band(1)._MapCol(1)._Name=   "ITEM_NAME"
      _Band(1)._MapCol(1)._RSIndex=   1
      _Band(1)._MapCol(2)._Name=   "QTY"
      _Band(1)._MapCol(2)._RSIndex=   2
      _Band(1)._MapCol(3)._Name=   "DESCRIBLE"
      _Band(1)._MapCol(3)._RSIndex=   3
      _Band(2).BandIndent=   2
      _Band(2).Cols   =   2
      _Band(2).GridLinesBand=   1
      _Band(2).TextStyleBand=   0
      _Band(2).TextStyleHeader=   0
      _Band(2)._ParentBand=   0
      _Band(2)._NumMapCols=   3
      _Band(2)._MapCol(0)._Name=   "ONSPT_SLN"
      _Band(2)._MapCol(0)._RSIndex=   0
      _Band(2)._MapCol(0)._Hidden=   -1  'True
      _Band(2)._MapCol(1)._Name=   "ITME_NAME"
      _Band(2)._MapCol(1)._RSIndex=   1
      _Band(2)._MapCol(2)._Name=   "QTY"
      _Band(2)._MapCol(2)._RSIndex=   2
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   0
      Picture         =   "MASTER_ONSPOT_REPLACE_DETAIL.frx":0020
      Top             =   0
      Width           =   3000
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "MASTER ON SPOT REPLACEMENT DETAIL "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   480
      Width           =   6975
   End
   Begin VB.Label Label2 
      Caption         =   "REPLACED ITEM DETAIL"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   8160
      TabIndex        =   2
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "DEFECTIVE ITEM DETAIL "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
   End
End
Attribute VB_Name = "MASTER_ONSPOT_REPLACE_DETAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

