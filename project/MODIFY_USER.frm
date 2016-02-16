VERSION 5.00
Begin VB.Form MODIFY_USER 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MODIFY USER"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton exitCmd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "ESP"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7212
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Exit"
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton delCmd 
      BackColor       =   &H008080FF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "ESP"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5284
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Delete"
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton editCmd 
      BackColor       =   &H00FFFF80&
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "ESP"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3244
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Edit "
      Top             =   7440
      Width           =   1335
   End
   Begin VB.TextBox passwordText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   7560
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.ComboBox userIDCombo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   2280
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CheckBox admChk 
      Caption         =   "Administration"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   360
      TabIndex        =   50
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CheckBox purVenChk 
      Caption         =   "Vendor"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   3600
      TabIndex        =   49
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CheckBox PurChk 
      Caption         =   "Purcahse "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   3360
      TabIndex        =   48
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CheckBox admRvChk 
      Caption         =   "Recovery"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   600
      TabIndex        =   47
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CheckBox admBUChk 
      Caption         =   "Back Up"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   600
      TabIndex        =   46
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CheckBox admMUChk 
      Caption         =   "Modify User"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   600
      TabIndex        =   45
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CheckBox admCPChk 
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   600
      TabIndex        =   44
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CheckBox admCUChk 
      Caption         =   "Create User"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   600
      TabIndex        =   43
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CheckBox admPUChk 
      Caption         =   "Product Update"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   600
      TabIndex        =   42
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CheckBox admPEChk 
      Caption         =   "Product Entry"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   600
      TabIndex        =   41
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CheckBox purPurChk 
      Caption         =   "Purchase "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3600
      TabIndex        =   40
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CheckBox purPRChk 
      Caption         =   "Purchase Retrun"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3600
      TabIndex        =   39
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CheckBox salChk 
      Caption         =   "Sale"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   6000
      TabIndex        =   38
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CheckBox salCustChk 
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   6120
      TabIndex        =   37
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CheckBox salPrtChk 
      Caption         =   "Party"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6120
      TabIndex        =   36
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CheckBox salSalChk 
      Caption         =   "Sale"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   6120
      TabIndex        =   35
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CheckBox salSRChk 
      Caption         =   "Sale Retrun"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   6120
      TabIndex        =   34
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CheckBox rplChk 
      Caption         =   "Replacement"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   8280
      TabIndex        =   33
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CheckBox rplOSRChk 
      Caption         =   "On Spot Replacement"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   8400
      TabIndex        =   32
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CheckBox rplRTPChk 
      Caption         =   "Replace To Principal"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   8400
      TabIndex        =   31
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CheckBox rplRFChk 
      Caption         =   "Replace From  "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   8400
      TabIndex        =   30
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CheckBox rplDPEChk 
      Caption         =   "Defective Product Entry"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   8400
      TabIndex        =   29
      Top             =   2880
      Width           =   2895
   End
   Begin VB.CheckBox empSChk 
      Caption         =   "Employee Support"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   360
      TabIndex        =   28
      Top             =   4200
      Width           =   2775
   End
   Begin VB.CheckBox empSJEChk 
      Caption         =   "Join Employee"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   480
      TabIndex        =   27
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CheckBox empSREChk 
      Caption         =   "Relieving Employee"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   480
      TabIndex        =   26
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CheckBox empSESChk 
      Caption         =   "Employee Salary"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   480
      TabIndex        =   25
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CheckBox matDChk 
      Caption         =   "Master Detail "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   4320
      TabIndex        =   24
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CheckBox matDSDChk 
      Caption         =   "Stock Detail"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4560
      TabIndex        =   23
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CheckBox matDPDChk 
      Caption         =   "Purchase Detail"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4560
      TabIndex        =   22
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CheckBox matDSalDChk 
      Caption         =   "Sale Detail "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4560
      TabIndex        =   21
      Top             =   5160
      Width           =   2055
   End
   Begin VB.CheckBox matDOSRDChk 
      Caption         =   "On Spot Replacement Detail"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   4560
      TabIndex        =   20
      Top             =   5400
      Width           =   3255
   End
   Begin VB.CheckBox matDRTPChk 
      Caption         =   "Replace To Principal"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   4560
      TabIndex        =   19
      Top             =   5640
      Width           =   2535
   End
   Begin VB.CheckBox matDRFCChk 
      Caption         =   "Replace From Customer"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   4560
      TabIndex        =   18
      Top             =   5880
      Width           =   2895
   End
   Begin VB.CheckBox matDCEDChk 
      Caption         =   "Current Employee Detail "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   255
      Left            =   4560
      TabIndex        =   17
      Top             =   6120
      Width           =   3015
   End
   Begin VB.CheckBox matDREDChk 
      Caption         =   "Relieved Employee Detail"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   255
      Left            =   4560
      TabIndex        =   16
      Top             =   6360
      Width           =   3135
   End
   Begin VB.CheckBox rptChk 
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   8280
      TabIndex        =   15
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CheckBox rptPRChk 
      Caption         =   "Product Report"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   8640
      TabIndex        =   14
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CheckBox rptVRChk 
      Caption         =   "Vendor Report "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   8640
      TabIndex        =   13
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CheckBox rptPrtRChk 
      Caption         =   "Party Report "
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   8640
      TabIndex        =   12
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CheckBox rptCRChk 
      Caption         =   "Customer Report"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   8640
      TabIndex        =   11
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CheckBox rptPurRChk 
      Caption         =   "Purhcase Report"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   8640
      TabIndex        =   10
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CheckBox rptSalRChk 
      Caption         =   "Sale Report"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   8640
      TabIndex        =   9
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CheckBox rptStkRChk 
      Caption         =   "Stock Report"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   8640
      TabIndex        =   8
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CheckBox rptRplRChk 
      Caption         =   "Replacement Report"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   255
      Left            =   8640
      TabIndex        =   7
      Top             =   6360
      Width           =   2415
   End
   Begin VB.CheckBox rptDPRChk 
      Caption         =   "Defective Product Report"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   255
      Left            =   8640
      TabIndex        =   6
      Top             =   6600
      Width           =   2895
   End
   Begin VB.CheckBox rptERChk 
      Caption         =   "Employee Report"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   8640
      TabIndex        =   5
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Line Line26 
      X1              =   6960
      X2              =   6960
      Y1              =   7200
      Y2              =   8040
   End
   Begin VB.Line Line25 
      X1              =   4920
      X2              =   4920
      Y1              =   7200
      Y2              =   8040
   End
   Begin VB.Line Line24 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   8760
      X2              =   8760
      Y1              =   7200
      Y2              =   8040
   End
   Begin VB.Line Line23 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   3000
      X2              =   8760
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Line Line22 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   3000
      X2              =   3000
      Y1              =   7200
      Y2              =   8040
   End
   Begin VB.Line Line21 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   9600
      X2              =   9600
      Y1              =   960
      Y2              =   600
   End
   Begin VB.Line Line20 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   2640
      X2              =   9600
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line19 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   2640
      X2              =   2640
      Y1              =   600
      Y2              =   960
   End
   Begin VB.Line Line18 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   1560
      Y2              =   960
   End
   Begin VB.Line Line17 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   11640
      X2              =   11640
      Y1              =   1560
      Y2              =   960
   End
   Begin VB.Line Line16 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   11640
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   0
      Picture         =   "MODIFY_USER.frx":0000
      ToolTipText     =   "San's Modify the User Permission "
      Top             =   0
      Width           =   3000
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "MODIFY EXISTING USER AND THEIR PERMISSIONS FORM"
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   14.25
         Charset         =   255
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2640
      TabIndex        =   53
      Top             =   600
      Width           =   6960
   End
   Begin VB.Label Label2 
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   5280
      TabIndex        =   52
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "USER ID"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   360
      TabIndex        =   51
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   11640
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   11640
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   11640
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   11640
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   1560
      Y2              =   3960
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   11640
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   11640
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line8 
      X1              =   3240
      X2              =   3240
      Y1              =   1560
      Y2              =   3960
   End
   Begin VB.Line Line9 
      X1              =   5880
      X2              =   5880
      Y1              =   1560
      Y2              =   3960
   End
   Begin VB.Line Line10 
      X1              =   8160
      X2              =   8160
      Y1              =   1560
      Y2              =   3960
   End
   Begin VB.Line Line11 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   11640
      X2              =   11640
      Y1              =   1560
      Y2              =   3960
   End
   Begin VB.Line Line12 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   3960
      Y2              =   7200
   End
   Begin VB.Line Line13 
      X1              =   3720
      X2              =   3720
      Y1              =   4080
      Y2              =   7200
   End
   Begin VB.Line Line14 
      X1              =   8160
      X2              =   8160
      Y1              =   4080
      Y2              =   7200
   End
   Begin VB.Line Line15 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   11640
      X2              =   11640
      Y1              =   3960
      Y2              =   7200
   End
End
Attribute VB_Name = "MODIFY_USER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************'
'**                           **'
'** MODIFY EXISTING USER FORM **'
'**                           **'
'*******************************'

'VARIABLE DECLERATION

Option Explicit

Dim muSession As OraSession
Dim muDatabase As OraDatabase
Dim muDyn     As OraDynaset
Dim userIdDyn As OraDynaset
Dim updateSql As String
Dim delSql    As String

Private Sub UserId()
 
 userIDCombo.CLEAR
 Set userIdDyn = muDatabase.CreateDynaset("Select user_id from sanproject_user where user_id !='administrator' ", &H4&)

 While Not userIdDyn.EOF  'ADDING USER ID IN ID COMBO
    userIDCombo.AddItem userIdDyn.Fields(0)
    userIdDyn.MoveNext
 Wend
 userIDCombo.Text = "Select user ID."

End Sub


Private Sub delCmd_Click()
 
 On Error GoTo ErrorHandeler
 
 If passwordText.Text = "" And ModuleVarious.LogOnUser <> "administrator" Then
     MsgBox "Password is Blank .", vbExclamation, "San's Project:"
     passwordText.SetFocus
     Exit Sub
 End If

 If ModuleVarious.LogOnUser = "administrator" Then
    Set muDyn = muDatabase.CreateDynaset("Select * from hrishi.sanproject_user where user_id ='" & userIDCombo.Text & "'", &H4&)
 Else
   Set muDyn = muDatabase.CreateDynaset("Select * from hrishi.sanproject_user where user_id ='" & userIDCombo.Text & "' and password ='" & passwordText.Text & "'", &H4&)
 End If

   
 If muDyn.EOF Then
   MsgBox "Unauthorized user ." & vbCrLf & " Not Permisssion to change.", vbCritical, "San's Project."
   userIDCombo.SetFocus
   Exit Sub
 End If
 
 delSql = "delete from  hrishi.sanproject_user where user_id = '" & userIDCombo.Text & "' "
  
  
 If MsgBox("Do you really want to remove this user ?", vbExclamation + vbYesNo, "Conformation:") = vbYes Then
    muDatabase.ExecuteSQL (delSql)
 Else
     Call UserId
     passwordText.Text = ""
     Exit Sub
 End If
 
 
  
ErrorHandeler:
  If muSession.LastServerErr = 0 Then
     If muDatabase.LastServerErr = 0 Then
        If Err.Number = 0 Then
           MsgBox " Success ! " & vbCrLf & userIDCombo.Text & ":  is Deleted.", vbInformation, "San's Project :"
           Call UserId
           passwordText.Text = ""
        Else
         MsgBox "VB Error :" & Err.Number & Err.Description, vbCritical, "VB Error:"
         Exit Sub
        End If
    Else
      MsgBox "Database Error :" & muDatabase.LastServerErr & muDatabase.LastServerErrText, vbCritical, "Database Error:"
      muDatabase.LastServerErrReset
      Exit Sub
    End If
 Else
   MsgBox "Session Error:" & muSession.LastServerErr & muSession.LastServerErrText, vbCritical, "Session Error:"
   muSession.LastServerErrReset
   Exit Sub
 End If
 
End Sub

Private Sub editCmd_Click()

If editCmd.Caption = "Edit" Then
  
     If passwordText.Text = "" And ModuleVarious.LogOnUser <> "administrator" Then
       MsgBox "Password is Blank .", vbExclamation, "San's Project:"
       passwordText.SetFocus
       Exit Sub
    End If
 
  
  
 On Error GoTo ErrorHandeler
 
 If ModuleVarious.LogOnUser = "administrator" Then
    Set muDyn = muDatabase.CreateDynaset("Select * from hrishi.sanproject_user where user_id ='" & userIDCombo.Text & "'", &H4&)
 Else
    Set muDyn = muDatabase.CreateDynaset("Select * from hrishi.sanproject_user where user_id ='" & userIDCombo.Text & "' and password ='" & passwordText.Text & "'", &H4&)
 End If
 
 If muDyn.EOF Then
   MsgBox "Unauthorized user ." & vbCrLf & " Not Permisssion to change.", vbCritical, "San's Project."
   userIDCombo.SetFocus
   Exit Sub
 End If
 
 admChk.Value = muDyn.Fields("ADM")
 admPEChk.Value = muDyn.Fields("ADMPE")
 admPUChk.Value = muDyn.Fields("ADMPU")
 admCUChk.Value = muDyn.Fields("ADMCU")
 admMUChk.Value = muDyn.Fields("ADMMU")
 admCPChk.Value = muDyn.Fields("ADMCP")
 admBUChk.Value = muDyn.Fields("ADMBU")
 admRvChk.Value = muDyn.Fields("ADMRV")
 
 PurChk.Value = muDyn.Fields("PUR")
 purVenChk.Value = muDyn.Fields("PURVEN")
 purPurChk.Value = muDyn.Fields("PURPUR")
 purPRChk.Value = muDyn.Fields("PURPR")
 
 salChk.Value = muDyn.Fields("SAL")
 salCustChk.Value = muDyn.Fields("SALCUST")
 salPrtChk.Value = muDyn.Fields("SALPRT")
 salSalChk.Value = muDyn.Fields("SALSAL")
 salSRChk.Value = muDyn.Fields("SALSR")
 
 rplChk.Value = muDyn.Fields("RPL")
 rplOSRChk.Value = muDyn.Fields("RPLOSR")
 rplRFChk.Value = muDyn.Fields("RPLRF")
 rplRTPChk.Value = muDyn.Fields("RPLRTP")
 rplDPEChk.Value = muDyn.Fields("RPLDPE")
 
 empSChk.Value = muDyn.Fields("EMPS")
 empSJEChk.Value = muDyn.Fields("EMPSJE")
 empSREChk.Value = muDyn.Fields("EMPSRE")
 empSESChk.Value = muDyn.Fields("EMPSES")
 
 matDChk.Value = muDyn.Fields("MATD")
 matDSDChk.Value = muDyn.Fields("MATDSD")
 matDPDChk.Value = muDyn.Fields("MATDPD")
 matDSalDChk.Value = muDyn.Fields("MATDSALD")
 matDOSRDChk.Value = muDyn.Fields("MATDOSRD")
 matDRTPChk.Value = muDyn.Fields("MATDRTP")
 matDRFCChk.Value = muDyn.Fields("MATDRFC")
 matDCEDChk.Value = muDyn.Fields("MATDED")
 matDREDChk.Value = muDyn.Fields("MATDRED")
 
 rptChk.Value = muDyn.Fields("RPT")
 rptPRChk.Value = muDyn.Fields("RPTPR")
 rptVRChk.Value = muDyn.Fields("RPTVR")
 rptPrtRChk.Value = muDyn.Fields("RPTPRTR")
 rptPurRChk.Value = muDyn.Fields("RPTPURR")
 rptSalRChk.Value = muDyn.Fields("RPTSALR")
 rptStkRChk.Value = muDyn.Fields("RPTSTKR")
 rptRplRChk.Value = muDyn.Fields("RPTRPLR")
 rptDPRChk.Value = muDyn.Fields("RPTDPR")
 rptERChk.Value = muDyn.Fields("RPTER")
 
 editCmd.Caption = "Ok"
 userIDCombo.Enabled = False
 passwordText.Enabled = False
 Exit Sub
End If

If editCmd.Caption = "Ok" Then

 
 updateSql = " update    hrishi.sanproject_user  set ADM= " & admChk.Value & ",ADMPE= " & admPEChk.Value & ",ADMPU=" & admPUChk.Value & ",ADMCU=" & admCUChk.Value & ",ADMMU=" & admMUChk.Value _
             & ",ADMCP=" & admCPChk.Value & ",ADMBU=" & admBUChk.Value & ",ADMRV=" & admRvChk.Value & ",PUR=" & PurChk.Value & ",PURVEN= " & _
             purVenChk.Value & ",PURPUR=" & purPurChk.Value & ",PURPR=" & purPRChk.Value & ",SAL=" & salChk.Value & ",SALCUST=" & salCustChk.Value _
             & ",SALPRT=" & salPrtChk.Value & ",SALSAL=" & salSalChk.Value & ",SALSR=" & salSRChk.Value & ",RPL=" & rplChk.Value & ",RPLOSR=" & rplOSRChk.Value _
             & ",RPLRF=" & rplRFChk.Value & " ,RPLRTP=" & rplRTPChk.Value & ",RPLDPE=" & rplDPEChk.Value & " ,EMPS=" & empSChk.Value _
             & ",EMPSJE=" & empSJEChk.Value & ",EMPSRE=" & empSREChk.Value & ",EMPSES=" & empSESChk.Value & ",MATD=" & matDChk.Value & ",MATDSD=" & _
             matDSDChk.Value & ",MATDPD=" & matDPDChk.Value & ",MATDSALD=" & matDSalDChk.Value & ",MATDOSRD=" & matDOSRDChk.Value & " ,MATDRTP=" & matDRTPChk.Value _
             & ",MATDRFC=" & matDRFCChk.Value & ",MATDED=" & matDCEDChk.Value & ",MATDRED=" & matDREDChk.Value & ",RPT=" & rptChk.Value & ",RPTPR=" & rptPRChk.Value _
             & " ,RPTVR=" & rptVRChk.Value & ",RPTPRTR=" & rptPrtRChk.Value & ",RPTCR=" & rptCRChk.Value & ",RPTPURR=" & rptPurRChk.Value & ",RPTSALR=" & rptSalRChk.Value _
             & ",RPTSTKR=" & rptStkRChk.Value & ",RPTRPLR=" & rptRplRChk.Value & ",RPTDPR=" & rptDPRChk.Value & ",RPTER=" & rptERChk.Value & " where user_id ='" & userIDCombo.Text & "'"
  
  
  muDatabase.ExecuteSQL (updateSql)
  
ErrorHandeler:
  If muSession.LastServerErr = 0 Then
     If muDatabase.LastServerErr = 0 Then
        If Err.Number = 0 Then
           MsgBox " Success ! " & vbCrLf & userIDCombo.Text & ": your information is Modified.", vbInformation, "San's Project :"
           userIDCombo.Enabled = True
           passwordText.Enabled = True
           editCmd.Caption = "Edit"
           Call UserId
           passwordText.Text = ""
            admChk.Value = 0
            admPEChk.Value = 0
            admPUChk.Value = 0
            admCUChk.Value = 0
            admMUChk.Value = 0
            admCPChk.Value = 0
            admBUChk.Value = 0
            admRvChk.Value = 0
 
            PurChk.Value = 0
            purVenChk.Value = 0
            purPurChk.Value = 0
            purPRChk.Value = 0
 
            salChk.Value = 0
            salCustChk.Value = 0
            salPrtChk.Value = 0
            salSalChk.Value = 0
            salSRChk.Value = 0
 
            rplChk.Value = 0
            rplOSRChk.Value = 0
            rplRFChk.Value = 0
            rplRTPChk.Value = 0
            rplDPEChk.Value = 0
 
            empSChk.Value = 0
            empSJEChk.Value = 0
            empSREChk.Value = 0
            empSESChk.Value = 0
 
            matDChk.Value = 0
            matDSDChk.Value = 0
            matDPDChk.Value = 0
            matDSalDChk.Value = 0
            matDOSRDChk.Value = 0
            matDRTPChk.Value = 0
            matDRFCChk.Value = 0
            matDCEDChk.Value = 0
            matDREDChk.Value = 0
 
            rptChk.Value = 0
            rptPRChk.Value = 0
            rptVRChk.Value = 0
            rptPrtRChk.Value = 0
            rptPurRChk.Value = 0
            rptSalRChk.Value = 0
            rptStkRChk.Value = 0
            rptRplRChk.Value = 0
            rptDPRChk.Value = 0
            rptERChk.Value = 0
  
        Else
         MsgBox "VB Error :" & Err.Number & Err.Description, vbCritical, "VB Error:"
         Exit Sub
        End If
    Else
      MsgBox "Database Error :" & muDatabase.LastServerErr & muDatabase.LastServerErrText, vbCritical, "Database Error:"
      muDatabase.LastServerErrReset
      Exit Sub
    End If
 Else
   MsgBox "Session Error:" & muSession.LastServerErr & muSession.LastServerErrText, vbCritical, "Session Error:"
   muSession.LastServerErrReset
   Exit Sub
 End If
 
 End If
 
 End Sub

Private Sub exitCmd_Click()

  If MsgBox("Do you exit ?", vbExclamation + vbYesNo, "Exit:") = vbYes Then
     Unload Me
  End If
  
End Sub

Private Sub Form_Load()

  Set muSession = CreateObject("oracleinprocserver.xorasession")
  Set muDatabase = muSession.OpenDatabase("jms", "hrishi/jms", &H0&)
  Set userIdDyn = muDatabase.CreateDynaset("Select user_id from sanproject_user where user_id !='administrator' ", &H4&)

  While Not userIdDyn.EOF
      userIDCombo.AddItem userIdDyn.Fields(0)
      userIdDyn.MoveNext
  Wend

  userIDCombo.Text = "Select user ID."
  
End Sub


Private Sub userIDCombo_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     passwordText.SetFocus
  End If

End Sub
