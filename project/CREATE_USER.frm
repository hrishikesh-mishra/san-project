VERSION 5.00
Begin VB.Form CREATE_USER 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CREATE USER "
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11820
   FillColor       =   &H000000FF&
   ForeColor       =   &H00C000C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton exitCmd 
      BackColor       =   &H00808080&
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Exit from this."
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton cancelCmd 
      BackColor       =   &H00808080&
      Caption         =   "Cancel"
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Cancel the Process"
      Top             =   7560
      Width           =   1320
   End
   Begin VB.CommandButton okCmd 
      BackColor       =   &H00808080&
      Caption         =   "Ok"
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
      Left            =   3360
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Create the user"
      Top             =   7560
      Width           =   1095
   End
   Begin VB.TextBox rePasswordText 
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   9840
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Enter the confirm password"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox passwordText 
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   5760
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Enter the password"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox userIDText 
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1560
      MaxLength       =   15
      TabIndex        =   0
      ToolTipText     =   "Enter the user ID"
      Top             =   1080
      Width           =   1575
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
      TabIndex        =   48
      Top             =   6960
      Width           =   2175
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
      TabIndex        =   47
      Top             =   6720
      Width           =   2895
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
      TabIndex        =   46
      Top             =   6480
      Width           =   2415
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
      TabIndex        =   45
      Top             =   6240
      Width           =   1695
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
      TabIndex        =   44
      Top             =   6000
      Width           =   1575
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
      TabIndex        =   43
      Top             =   5760
      Width           =   2175
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
      TabIndex        =   42
      Top             =   5520
      Width           =   2175
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
      TabIndex        =   41
      Top             =   5280
      Width           =   1695
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
      TabIndex        =   40
      Top             =   5040
      Width           =   1935
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
      TabIndex        =   39
      Top             =   4800
      Width           =   1935
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
      TabIndex        =   38
      Top             =   4320
      Width           =   2415
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
      TabIndex        =   37
      Top             =   6480
      Width           =   3135
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
      TabIndex        =   36
      Top             =   6240
      Width           =   3015
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
      TabIndex        =   35
      Top             =   6000
      Width           =   2895
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
      TabIndex        =   34
      Top             =   5760
      Width           =   2535
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
      TabIndex        =   33
      Top             =   5520
      Width           =   3255
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
      TabIndex        =   32
      Top             =   5280
      Width           =   2055
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
      TabIndex        =   31
      Top             =   5040
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
      TabIndex        =   30
      Top             =   4800
      Width           =   1935
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
      TabIndex        =   29
      Top             =   4320
      Width           =   2175
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
      TabIndex        =   28
      Top             =   5280
      Width           =   2175
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
      TabIndex        =   27
      Top             =   5040
      Width           =   2415
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
      TabIndex        =   26
      Top             =   4800
      Width           =   1935
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
      TabIndex        =   25
      Top             =   4320
      Width           =   2775
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
      TabIndex        =   24
      Top             =   3000
      Width           =   2895
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
      TabIndex        =   22
      Top             =   2520
      Width           =   1815
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
      TabIndex        =   23
      Top             =   2760
      Width           =   2535
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
      TabIndex        =   21
      Top             =   2280
      Width           =   2655
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
      TabIndex        =   20
      Top             =   1800
      Width           =   1935
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
      TabIndex        =   19
      Top             =   3000
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
      TabIndex        =   18
      Top             =   2760
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
      TabIndex        =   17
      Top             =   2520
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
      TabIndex        =   16
      Top             =   2280
      Width           =   1455
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
      TabIndex        =   15
      Top             =   1800
      Width           =   1575
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
      TabIndex        =   14
      Top             =   2760
      Width           =   2055
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
      TabIndex        =   13
      Top             =   2520
      Width           =   1455
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
      TabIndex        =   4
      Top             =   2280
      Width           =   1815
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
      TabIndex        =   5
      Top             =   2520
      Width           =   1935
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
      TabIndex        =   6
      Top             =   2760
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
      TabIndex        =   8
      Top             =   3240
      Width           =   2175
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
      TabIndex        =   7
      Top             =   3000
      Width           =   1695
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
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
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
      TabIndex        =   10
      Top             =   3720
      Width           =   1455
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
      TabIndex        =   11
      Top             =   1800
      Width           =   1815
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
      TabIndex        =   12
      Top             =   2280
      Width           =   1095
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
      TabIndex        =   3
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Line Line23 
      BorderColor     =   &H000000FF&
      X1              =   7560
      X2              =   7560
      Y1              =   960
      Y2              =   1560
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   585
      Left            =   0
      Picture         =   "CREATE_USER.frx":0000
      ToolTipText     =   "San's Create User Form."
      Top             =   0
      Width           =   3000
   End
   Begin VB.Line Line26 
      BorderColor     =   &H000000FF&
      X1              =   120
      X2              =   11520
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line25 
      BorderColor     =   &H000000FF&
      X1              =   120
      X2              =   120
      Y1              =   960
      Y2              =   1560
   End
   Begin VB.Line Line24 
      BorderColor     =   &H000000FF&
      X1              =   11520
      X2              =   11520
      Y1              =   960
      Y2              =   1560
   End
   Begin VB.Line Line22 
      BorderColor     =   &H000000FF&
      X1              =   120
      X2              =   11520
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line21 
      BorderColor     =   &H000000FF&
      X1              =   3360
      X2              =   3360
      Y1              =   960
      Y2              =   1560
   End
   Begin VB.Line Line20 
      BorderColor     =   &H0000C0C0&
      X1              =   6960
      X2              =   6960
      Y1              =   7320
      Y2              =   8160
   End
   Begin VB.Line Line19 
      BorderColor     =   &H0000C0C0&
      X1              =   4920
      X2              =   4920
      Y1              =   7320
      Y2              =   8160
   End
   Begin VB.Line Line18 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      X1              =   8760
      X2              =   8760
      Y1              =   8160
      Y2              =   7320
   End
   Begin VB.Line Line17 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      X1              =   3000
      X2              =   8760
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line16 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      X1              =   3000
      X2              =   3000
      Y1              =   7320
      Y2              =   8160
   End
   Begin VB.Label Label4 
      Caption         =   "ReType PassWord:"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   7800
      TabIndex        =   55
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "PassWord:"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3720
      TabIndex        =   54
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "User ID: "
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   240
      TabIndex        =   53
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   11640
      X2              =   11640
      Y1              =   4200
      Y2              =   7320
   End
   Begin VB.Line Line14 
      BorderColor     =   &H0000FF00&
      X1              =   8160
      X2              =   8160
      Y1              =   4200
      Y2              =   7320
   End
   Begin VB.Line Line13 
      BorderColor     =   &H0000FF00&
      X1              =   3720
      X2              =   3720
      Y1              =   4200
      Y2              =   7320
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   4200
      Y2              =   7320
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   11640
      X2              =   11640
      Y1              =   1680
      Y2              =   4080
   End
   Begin VB.Line Line10 
      BorderColor     =   &H0000FF00&
      X1              =   8160
      X2              =   8160
      Y1              =   1680
      Y2              =   4080
   End
   Begin VB.Line Line9 
      BorderColor     =   &H0000FF00&
      X1              =   5880
      X2              =   5880
      Y1              =   1680
      Y2              =   4080
   End
   Begin VB.Line Line8 
      BorderColor     =   &H0000FF00&
      X1              =   3240
      X2              =   3240
      Y1              =   1680
      Y2              =   4080
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   120
      X2              =   11640
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   120
      X2              =   11640
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   1680
      Y2              =   4080
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   120
      X2              =   11640
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   120
      X2              =   11640
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   120
      X2              =   11640
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   120
      X2              =   11640
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "CREATE USER AND GRANT PREMISSIONS FORM"
      BeginProperty Font 
         Name            =   "RAYLENE"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1770
      TabIndex        =   52
      Top             =   600
      Width           =   8280
   End
End
Attribute VB_Name = "CREATE_USER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************'
'**                  **'
'** CREATE USER FORM **'
'**                  **'
'**********************'

'VARIABLE DECLERATION

Option Explicit

Dim CuSession As OraSession
Dim CuDatabase As OraDatabase
Dim userIdDyn As OraDynaset
Dim insertSql As String


Private Sub admChk_Click()
          
  If admChk.Value = 1 Then  'IF ADMINISTRATOR CHECK BOX IS CLICK AND CHECKED
     admPEChk.Value = 1     'THEN ALL CHECK BOX WHICH IS BELOW TO THIS IS ALSO
     admPUChk.Value = 1     'CHECKED
     admCUChk.Value = 1
     admMUChk.Value = 1
     admCPChk.Value = 1
     admBUChk.Value = 1
     admRvChk.Value = 1
  Else                      'IF ADMINSTRATOR CHECK BOX IS CLICK AND CHECK BOX IS CLEAR
     admPEChk.Value = 0     'THEN ALL CHECK BOX WHICH IS BELOW IS ALSO CLEARED
     admPUChk.Value = 0
     admCUChk.Value = 0
     admMUChk.Value = 0
     admCPChk.Value = 0
     admBUChk.Value = 0
     admRvChk.Value = 0
  End If
 
End Sub

Private Sub cancelCmd_Click()
            
            'IF CANCEL CMD IS CLICK THEN ALL CHECK BOX AND TEXT BOX BECOME CLEAR
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
            
            userIDText.Text = ""
            passwordText.Text = ""
            rePasswordText.Text = ""

End Sub

Private Sub empSChk_Click()
     
 If empSChk.Value = 1 Then  'IF EMPLOYEE SUPPORT CHECK BOX IS CLICKED AND IT
    empSJEChk.Value = 1     'BECOME CHECKED THEN ALL CHECK BOX BELLOW IT ALSO
    empSREChk.Value = 1     'CHECKED
    empSESChk.Value = 1
 Else                       'ELSE CLEAR THEN ALL CHECK BOX BELLOW IT BECOME
    empSJEChk.Value = 0     'CLEAR
    empSREChk.Value = 0
    empSESChk.Value = 0
 End If

End Sub

Private Sub exitCmd_Click()
  
  'CONFIRMING FOR EXIT
  If MsgBox("Do you want to exit ?", vbExclamation + vbYesNo, "Exit:") = vbYes Then
     Unload Me
  End If
  
End Sub

Private Sub Form_Load()

 'CREATING SESSION AND OPENING DATABASE
 On Error GoTo ERRORHANDLER
 Set CuSession = CreateObject("oracleinprocserver.xorasession")
 Set CuDatabase = CuSession.OpenDatabase("jms", "hrishi/jms", &H0&)
 
ERRORHANDLER:
   'CODING FOR ERROR HANDLER
 If CuSession.LastServerErr = 0 Then
    If CuDatabase.LastServerErr = 0 Then
       If Err.Number = 0 Then
       Else
          MsgBox "VB ERROR :" & Err.Number & Err.Description, vbCritical, "VB Error:"
          Unload Me
       End If
    Else
       MsgBox "DATABASE ERROR:" & CuDatabase.LastServerErr & CuDatabase.LastServerErrText, _
              vbCritical, "DATABASE Error:"
       CuDatabase.LastServerErrReset
       Unload Me
    End If
 Else
    MsgBox "SESSION ERROR:" & CuSession.LastServerErr & CuSession.LastServerErrText, _
            vbCritical, "SESSION Error:"
    CuSession.LastServerErrReset
    Unload Me
 End If
 
End Sub

Private Sub matDChk_Click()

  If matDChk.Value = 1 Then     'IF MASTER DETAIL CHECK BOX IS CLICK AND CHECKED
     matDSDChk.Value = 1        'THEN ALL CHECK BOX BELOW IT BECOME CHECKED
     matDPDChk.Value = 1
     matDSalDChk.Value = 1
     matDOSRDChk.Value = 1
     matDRTPChk.Value = 1
     matDRFCChk.Value = 1
     matDCEDChk.Value = 1
     matDREDChk.Value = 1
  Else                        'IF MASTER DETAIL CHECK BOX IS CLICK AND CLEARED
     matDSDChk.Value = 0      'THEN ALL CHECK BOX BELOW IT BECOME CLEARED
     matDPDChk.Value = 0
     matDSalDChk.Value = 0
     matDOSRDChk.Value = 0
     matDRTPChk.Value = 0
     matDRFCChk.Value = 0
     matDCEDChk.Value = 0
     matDREDChk.Value = 0
  End If

End Sub

Private Sub OkCmd_Click()

  If userIDText.Text = "" Then                'CHECK FOR EMPTY
     MsgBox "User ID is Empty .", vbCritical, "Empty:"
     userIDText.SetFocus
     Exit Sub
  ElseIf passwordText.Text = "" Then
     MsgBox "Password is Blank .", vbCritical, "Empty:"
     passwordText.SetFocus
     Exit Sub
  ElseIf rePasswordText.Text = "" Then
    MsgBox "ReType Password is Blank .", vbCritical, "Empty:"
    rePasswordText.SetFocus
    Exit Sub
  End If
   'VARIFING THE PASSWORD AND REENTER PASSWORD IS SAME
  If passwordText.Text <> rePasswordText.Text Then
     MsgBox "Both password don't match."
     passwordText.SetFocus
     Exit Sub
  End If

  On Error GoTo ERRORHANDLER
  Set userIdDyn = CuDatabase.CreateDynaset("select * from hrishi.sanproject_user where user_id='" & userIDText.Text & "'", &H0&)
   
  'CHECKING THE USER ID IS AN EXISTING USER
  If Not (userIdDyn.EOF) Then
     MsgBox "The UserID is Exist." & vbCrLf & " Change UserID .", vbCritical, "Duplicate UserID.:"
     userIDText.SetFocus
     Exit Sub
  End If

 'THE SQL FOR CREATING USER
  insertSql = "insert into hrishi.sanproject_user values ('" & userIDText.Text & "','" & passwordText.Text & "', sysdate," _
             & admChk.Value & ", " & admPEChk.Value & "," & admPUChk.Value & "," & admCUChk.Value & "," & admMUChk.Value _
             & "," & admCPChk.Value & "," & admBUChk.Value & "," & admRvChk.Value & "," & PurChk.Value & ", " & _
             purVenChk.Value & "," & purPurChk.Value & "," & purPRChk.Value & "," & salChk.Value & "," & salCustChk.Value _
             & "," & salPrtChk.Value & "," & salSalChk.Value & "," & salSRChk.Value & "," & rplChk.Value & "," & rplOSRChk.Value _
             & "," & rplRFChk.Value & " ," & rplRTPChk.Value & "," & rplDPEChk.Value & " ," & empSChk.Value _
             & "," & empSJEChk.Value & "," & empSREChk.Value & "," & empSESChk.Value & "," & matDChk.Value & "," & _
             matDSDChk.Value & "," & matDPDChk.Value & "," & matDSalDChk.Value & "," & matDOSRDChk.Value & " ," & matDRTPChk.Value _
             & "," & matDRFCChk.Value & "," & matDCEDChk.Value & "," & matDREDChk.Value & "," & rptChk.Value & "," & rptPRChk.Value _
             & " ," & rptVRChk.Value & "," & rptPrtRChk.Value & "," & rptCRChk.Value & "," & rptPurRChk.Value & "," & rptSalRChk.Value _
             & "," & rptStkRChk.Value & "," & rptRplRChk.Value & "," & rptDPRChk.Value & "," & rptERChk.Value & ")"
             
  'CREATING USER
  CuDatabase.ExecuteSQL (insertSql)

ERRORHANDLER:
  If CuSession.LastServerErr = 0 Then
     If CuDatabase.LastServerErr = 0 Then
        If Err.Number = 0 Then
           MsgBox " Success ! User is Created .", vbInformation, "Success :"
           Unload Me
        Else
           MsgBox "VB Error :" & Err.Number & Err.Description, vbCritical, " VB Error:"
           Exit Sub
        End If
     Else
        MsgBox "Database Error :" & CuDatabase.LastServerErr & CuDatabase.LastServerErrText, vbCritical, "Database Error:"
        CuDatabase.LastServerErrReset
        Exit Sub
     End If
  Else
     MsgBox " Session Error :" & CuSession.LastServerErr & CuSession.LastServerErrText, vbCritical, "Session Error:"
     CuSession.LastServerErrReset
     Exit Sub
  End If

End Sub

Private Sub passwordText_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then  'IF ENTER KEY IS PRESSED
     rePasswordText.SetFocus
  End If
  
End Sub

Private Sub PurChk_Click()

  If PurChk.Value = 1 Then   'IF PURCHASE CHECK BOX IS CLICKED AND BECOME CHECKED
     purVenChk.Value = 1    'THEN ALL CHECK BOX BELOW IT BECOME  CHECKED ALSO
     purPurChk.Value = 1
     purPRChk.Value = 1
  Else                       'IF PURCHASE CHECK BOX IS CLICKED AND BECOME CLEARED
     purVenChk.Value = 0    'THEN ALL CHECK BOX BELOW IT BECOME CLEARED ALSO
     purPurChk.Value = 0
     purPRChk.Value = 0
  End If
 
End Sub

Private Sub rplChk_Click()

  If rplChk.Value = 1 Then  'IF REPLACEMENT CHECK BOX IS CLICKED AND BECOME CHECKED
     rplOSRChk.Value = 1    'THEN ALL CHECK BOX BELOW IT BECOME CHECKED ALSO
     rplRFChk.Value = 1
     rplRTPChk.Value = 1
     rplDPEChk.Value = 1
  Else                     'IF REPLACEMENT CHECK BOX IS CLICKED AND BECOME CLEARED
     rplOSRChk.Value = 0   'THEN ALL CHECK BOX BELOW IT ALSO BECOME CLEARED
     rplRFChk.Value = 0
     rplRTPChk.Value = 0
     rplDPEChk.Value = 0
  End If

   
End Sub

Private Sub rptChk_Click()

  If rptChk.Value = 1 Then   'IF REPORT CHECK BOX IS CLICKED AND BECOME CHECKED
     rptPRChk.Value = 1      'THEN ALL CHECK BOX BELOW IT ALSO BECOME CHECKED
     rptVRChk.Value = 1
     rptPrtRChk.Value = 1
     rptCRChk.Value = 1
     rptPurRChk.Value = 1
     rptSalRChk.Value = 1
     rptStkRChk.Value = 1
     rptRplRChk.Value = 1
     rptDPRChk.Value = 1
     rptERChk.Value = 1
  Else                      'IF REPORT CHECK BOX IS CLICKED AND BECOME CLEARED
     rptPRChk.Value = 0     'THEN ALL CHECK BOX BELOW IT ALSO BECOME CLEARED
     rptVRChk.Value = 0
     rptPrtRChk.Value = 0
     rptCRChk.Value = 0
     rptPurRChk.Value = 0
     rptSalRChk.Value = 0
     rptStkRChk.Value = 0
     rptRplRChk.Value = 0
     rptDPRChk.Value = 0
     rptERChk.Value = 0
 End If
 
End Sub

Private Sub salChk_Click()

  If salChk.Value = 1 Then   'IF SALE CHECK BOX IS CLICKED AND BECOME CHECKED
     salCustChk.Value = 1    'THEN ALL CHECK BOX BELOW IT ALSO BECOME CHECKED
     salPrtChk.Value = 1
     salSalChk.Value = 1
     salSRChk.Value = 1
  Else                       'IF SALE CHECK BOX IS CLICKED AND BECOME CLEARED
     salCustChk.Value = 0    'THEN ALL CHECK BOX BELOW IT ALSO BECOME CLEARED
     salPrtChk.Value = 0
     salSalChk.Value = 0
     salSRChk.Value = 0
  End If
 
End Sub

Private Sub userIDText_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then     'IF ENTER KEY IS PRESSED
     passwordText.SetFocus
   End If
 
End Sub
