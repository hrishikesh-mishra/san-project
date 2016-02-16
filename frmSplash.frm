VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4530
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7425
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H000000C0&
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   960
         Top             =   3360
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1600
         Left            =   360
         Top             =   3360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         X1              =   2520
         X2              =   5520
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "SEVOKE ROAD SILIGURI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   2160
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "TIRUPATI APEX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "SAN'S COMP SYSTEM"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "IGNOU"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   20.25
            Charset         =   0
            Weight          =   600
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2040
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label lblCopyright 
         Caption         =   $"frmSplash.frx":0316
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   1215
         Left            =   3720
         TabIndex        =   3
         Top             =   2760
         Width           =   3135
      End
      Begin VB.Label lblCompany 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   2
         Top             =   3270
         Width           =   2415
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Caption         =   "Version      1.0.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4680
         TabIndex        =   4
         Top             =   2400
         Width           =   1770
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "BCA PROJECT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   660
         Left            =   2160
         TabIndex        =   6
         Top             =   840
         Width           =   4020
      End
      Begin VB.Label lblLicenseTo 
         Caption         =   "LicenseTo:   Sanjay Saha The EO of San's Comp System Siliguri (W/B)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2355
         TabIndex        =   5
         Top             =   705
         Width           =   105
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    
    Timer1.Enabled = True
    Timer2.Enabled = True
    
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Timer2.Enabled = False
Unload Me
LOG_ON.Show
End Sub

Private Sub Timer2_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
If ProgressBar1.Value = 100 Then
   Timer2.Enabled = False
End If



End Sub
