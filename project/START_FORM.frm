VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form START_FORM 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "START UP  FORM"
   ClientHeight    =   4860
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   7920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "START_FORM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   4560
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4650
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   7920
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   6960
         Top             =   120
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1600
         Left            =   6480
         Top             =   120
      End
      Begin VB.Label Label8 
         BackColor       =   &H000000FF&
         Height          =   855
         Left            =   2280
         TabIndex        =   15
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   "project"
         BeginProperty Font 
            Name            =   "DANIEL"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   750
         Index           =   1
         Left            =   4680
         TabIndex        =   14
         Top             =   960
         Width           =   3045
      End
      Begin VB.Image Image1 
         Height          =   525
         Left            =   -120
         Picture         =   "START_FORM.frx":000C
         Top             =   480
         Width           =   2940
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         Caption         =   "EXCLUSIVELY DEVELOPED ON DAY-TO-DAY TRANSACTION OF"
         BeginProperty Font 
            Name            =   "RONALD"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   2640
         TabIndex        =   13
         Top             =   1800
         Width           =   5055
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000009&
         Caption         =   "110 068"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
         Caption         =   "MAIDAN GARHI, NEW DELHI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   3480
         Width           =   3015
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000009&
         Caption         =   "INDIRA GANDHI NATIONAL OPEN UNIVERSITY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   4215
      End
      Begin VB.Image imgLogo 
         Height          =   705
         Index           =   1
         Left            =   1440
         Picture         =   "START_FORM.frx":3F2D
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         X1              =   3720
         X2              =   6720
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H80000009&
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
         Left            =   3840
         TabIndex        =   8
         Top             =   2640
         Width           =   2895
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
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
         Left            =   4320
         TabIndex        =   7
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
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
         Left            =   3720
         TabIndex        =   6
         Top             =   2040
         Width           =   3255
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H80000009&
         Caption         =   $"START_FORM.frx":432B
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
         Left            =   4560
         TabIndex        =   3
         Top             =   3360
         Width           =   3135
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H80000009&
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
         BackColor       =   &H80000009&
         Caption         =   "Version      1.1.1"
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
         Left            =   6000
         TabIndex        =   4
         Top             =   3000
         Width           =   1770
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "SAN'S"
         BeginProperty Font 
            Name            =   "DANIEL"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   750
         Index           =   0
         Left            =   2595
         TabIndex        =   5
         Top             =   960
         Width           =   2145
      End
      Begin VB.Label lblLicenseTo 
         BackColor       =   &H80000009&
         Caption         =   "LicenseTo:   Sanjay Saha The CEO of San's Comp System Siliguri (W/B)"
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
   End
End
Attribute VB_Name = "START_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************'
'**                       **'
'** SANPROJECT START FORM **'
'**                       **'
'***************************'

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
