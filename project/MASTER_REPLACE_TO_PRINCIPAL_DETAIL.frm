VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form MASTER_REPLACE_TO_PRINCIPAL_DETAIL 
   Caption         =   "2"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   9945
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   2535
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "MASTER_REPLACE_TO_PRINCIPAL_DETAIL.frx":0000
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Replace to Principal detail"
      Top             =   1200
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8493
      _Version        =   393216
      ForeColor       =   6842891
      Cols            =   7
      FixedCols       =   0
      ForeColorFixed  =   255
      AllowUserResizing=   3
      DataMember      =   "masterReplaceToPrnplDetail"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New Greek"
         Size            =   11.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   2
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   4
      _Band(0)._MapCol(0)._Name=   "RLP_SLN"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "VEN_ID"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "VEN_NAME"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "REP_DATE"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(1).BandIndent=   1
      _Band(1).Cols   =   3
      _Band(1).GridLinesBand=   1
      _Band(1).TextStyleBand=   0
      _Band(1).TextStyleHeader=   0
      _Band(1)._ParentBand=   0
      _Band(1)._NumMapCols=   4
      _Band(1)._MapCol(0)._Name=   "RLP_SLN"
      _Band(1)._MapCol(0)._RSIndex=   0
      _Band(1)._MapCol(0)._Hidden=   -1  'True
      _Band(1)._MapCol(1)._Name=   "ITEM_NAME"
      _Band(1)._MapCol(1)._RSIndex=   1
      _Band(1)._MapCol(2)._Name=   "QTY"
      _Band(1)._MapCol(2)._RSIndex=   2
      _Band(1)._MapCol(3)._Name=   "DESCRIP"
      _Band(1)._MapCol(3)._RSIndex=   3
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   9720
      X2              =   9720
      Y1              =   1080
      Y2              =   6120
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   6480
      X2              =   9720
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   6480
      X2              =   6480
      Y1              =   6120
      Y2              =   6600
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   6480
      X2              =   3480
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   3480
      X2              =   3480
      Y1              =   6120
      Y2              =   6600
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   120
      X2              =   3480
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   6120
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   9720
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   555
      Left            =   0
      Picture         =   "MASTER_REPLACE_TO_PRINCIPAL_DETAIL.frx":0020
      ToolTipText     =   "San's Master Replace To Principal Detail"
      Top             =   0
      Width           =   2970
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "MASTER REPLACE TO PRINCIPAL DETAIL"
      BeginProperty Font 
         Name            =   "OSGOOD"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Left            =   1530
      TabIndex        =   2
      Top             =   600
      Width           =   6735
   End
End
Attribute VB_Name = "MASTER_REPLACE_TO_PRINCIPAL_DETAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

