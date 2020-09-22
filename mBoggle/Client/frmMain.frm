VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Boggle"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10335
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraPlayer 
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   240
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Frame fraPFound 
         Caption         =   "Already Found (0)"
         Height          =   5295
         Left            =   2400
         TabIndex        =   31
         Top             =   2520
         Width           =   2295
         Begin VB.ListBox lstPFound 
            Height          =   4740
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame fraPAccepted 
         Caption         =   "Accepted (0)"
         Height          =   5295
         Left            =   0
         TabIndex        =   33
         Top             =   2520
         Width           =   2295
         Begin VB.ListBox lstPAccepted 
            BackColor       =   &H00FFFFFF&
            Height          =   4935
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame fraAwards 
         Caption         =   "Awards won"
         Height          =   1935
         Left            =   2760
         TabIndex        =   29
         Top             =   480
         Width           =   2055
         Begin prjBClient.CustomListBox clbAwards 
            Height          =   1575
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   2778
            BackColor       =   16777215
            BeginProperty FontInfo {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
            Graphical       =   0   'False
            Picture         =   "frmMain.frx":0CCA
            ScrollBarBackColor=   12632256
            ScrollBarBorderColor=   8421504
            SelBoxColor     =   12632064
            Sorted          =   0   'False
         End
      End
      Begin VB.Label lblPMedals 
         Caption         =   "0"
         Height          =   255
         Left            =   1560
         TabIndex        =   57
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         Caption         =   "No. of medals won:"
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   56
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblPSpeed 
         Caption         =   "0 kpm"
         Height          =   255
         Left            =   1560
         TabIndex        =   38
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblPFound 
         Caption         =   "0"
         Height          =   255
         Left            =   1560
         TabIndex        =   37
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblPPos 
         Caption         =   "0"
         Height          =   255
         Left            =   1560
         TabIndex        =   36
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblPPoints 
         Caption         =   "0"
         Height          =   255
         Left            =   1560
         TabIndex        =   35
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         Caption         =   "Position:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblInfo 
         Caption         =   "Typing Speed:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblInfo 
         Caption         =   "Words Found:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   27
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblInfo 
         Caption         =   "Points:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblStats 
         Caption         =   "Statistics of Player:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   0
         Width           =   4575
      End
   End
   Begin VB.Frame fraHelp 
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   240
      TabIndex        =   46
      Top             =   600
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Frame fraCredits 
         Caption         =   "Credits"
         Height          =   3375
         Left            =   0
         TabIndex        =   68
         Top             =   4440
         Visible         =   0   'False
         Width           =   4815
         Begin VB.Label lblInfo 
            Caption         =   "WavSource for the sounds"
            Height          =   255
            Index           =   22
            Left            =   600
            TabIndex        =   94
            Top             =   2400
            Width           =   3855
         End
         Begin VB.Label lblInfo 
            Caption         =   "Enable2k word lexicon"
            Height          =   255
            Index           =   21
            Left            =   600
            TabIndex        =   93
            Top             =   2640
            Width           =   3855
         End
         Begin VB.Label lblInfo 
            Caption         =   "Hasbro for Boggle board graphics and dictionary"
            Height          =   255
            Index           =   9
            Left            =   600
            TabIndex        =   92
            Top             =   2160
            Width           =   3855
         End
         Begin VB.Label lblInfo 
            Caption         =   "Patrick Gillespie for his custom listbox"
            Height          =   255
            Index           =   20
            Left            =   600
            TabIndex        =   75
            Top             =   1680
            Width           =   3855
         End
         Begin VB.Label lblInfo 
            Caption         =   "Carles P.V. for his multicolumn listbox"
            Height          =   255
            Index           =   19
            Left            =   600
            TabIndex        =   74
            Top             =   1920
            Width           =   3855
         End
         Begin VB.Label lblInfo 
            Caption         =   "Steven Hoyt for his algorithm for the Boggle Solver"
            Height          =   255
            Index           =   18
            Left            =   600
            TabIndex        =   73
            Top             =   1440
            Width           =   3855
         End
         Begin VB.Label Label1 
            Caption         =   "Acknowledgements:"
            Height          =   375
            Left            =   240
            TabIndex        =   72
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label lblInfo 
            Caption         =   "Chia Yan Sheng Alexander"
            Height          =   375
            Index           =   17
            Left            =   1800
            TabIndex        =   71
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label lblInfo 
            Caption         =   "Programmed by:"
            Height          =   255
            Index           =   16
            Left            =   240
            TabIndex        =   70
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.Frame fraHAwards 
         Caption         =   "Awards that can be won"
         Height          =   3375
         Left            =   0
         TabIndex        =   51
         Top             =   4440
         Width           =   2295
         Begin prjBClient.CustomListBox clbInfo 
            Height          =   3015
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   5318
            BackColor       =   16777215
            BeginProperty FontInfo {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   -2147483640
            Graphical       =   0   'False
            Picture         =   "frmMain.frx":0CE6
            ScrollBarBackColor=   12632256
            ScrollBarBorderColor=   8421504
            SelBoxColor     =   12632064
            Sorted          =   0   'False
         End
      End
      Begin MSComctlLib.TabStrip tabCredits 
         Height          =   375
         Left            =   3240
         TabIndex        =   69
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Style           =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Awards"
               Key             =   "tAward"
               Object.ToolTipText     =   "View awards that can be won each round"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Credits"
               Key             =   "tCred"
               Object.ToolTipText     =   "View the credits for this program"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Frame fraHScore 
         Caption         =   "Scoring"
         Height          =   1935
         Left            =   0
         TabIndex        =   66
         Top             =   2160
         Width           =   2415
         Begin prjBClient.ucReportList ucrInfo 
            Height          =   1575
            Left            =   120
            TabIndex        =   67
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   2778
            BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame fraHDescribe 
         Caption         =   "Description"
         Height          =   3375
         Left            =   2400
         TabIndex        =   53
         Top             =   4440
         Width           =   2415
         Begin VB.Label lblMedalD 
            Height          =   2895
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmMain.frx":0D02
         Height          =   855
         Index           =   12
         Left            =   120
         TabIndex        =   49
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmMain.frx":0DCB
         Height          =   855
         Index           =   11
         Left            =   120
         TabIndex        =   48
         Top             =   1440
         Width           =   4575
      End
      Begin VB.Label lblInfo 
         Caption         =   "Help and Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   10
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Width           =   3135
      End
   End
   Begin VB.CheckBox chkSounds 
      Alignment       =   1  'Right Justify
      Caption         =   "Sounds:"
      Height          =   195
      Left            =   5280
      TabIndex        =   95
      Top             =   7680
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.Frame fraStats 
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   240
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   4815
      Begin VB.Frame fraTop 
         Caption         =   "Top 50 Highest Valued Words"
         Height          =   2775
         Left            =   120
         TabIndex        =   45
         Top             =   1320
         Width           =   4575
         Begin prjBClient.ucReportList lstTop 
            Height          =   2415
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   4260
            BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Label lblHM 
         Caption         =   "None."
         Height          =   2655
         Left            =   480
         TabIndex        =   64
         Top             =   4560
         Width           =   4095
      End
      Begin VB.Label lblInfo 
         Caption         =   "Honourable Mention:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   120
         TabIndex        =   63
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label lblIntervalT 
         Caption         =   "Interval Time remaining:   3:00"
         Height          =   255
         Left            =   2160
         TabIndex        =   44
         Top             =   7440
         Width           =   2655
      End
      Begin VB.Label lblWFound 
         Alignment       =   1  'Right Justify
         Caption         =   "-"
         Height          =   375
         Left            =   2520
         TabIndex        =   43
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblInfo 
         Caption         =   "Total number of words found:"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   42
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label lblWinner 
         Alignment       =   1  'Right Justify
         Caption         =   "-"
         Height          =   375
         Left            =   960
         TabIndex        =   41
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label lblInfo 
         Caption         =   "Winner:"
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblRNum 
         Caption         =   "Statistics for Round 1:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   0
         Width           =   3975
      End
   End
   Begin VB.Frame fraBoard 
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   240
      TabIndex        =   12
      Top             =   600
      Width           =   4815
      Begin VB.TextBox txtTime 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "3:00"
         Top             =   1940
         Width           =   2415
      End
      Begin VB.Frame fraFound 
         Caption         =   "Already Found (0)"
         Height          =   1935
         Left            =   1560
         TabIndex        =   21
         Top             =   0
         Width           =   1575
         Begin VB.ListBox lstFound 
            Height          =   1620
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame fraInvalid 
         Caption         =   "Invalid (0)"
         Height          =   1935
         Left            =   3120
         TabIndex        =   18
         Top             =   0
         Width           =   1575
         Begin VB.ListBox lstInv 
            Height          =   1620
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame fraAccepted 
         Caption         =   "Accepted (0)"
         Height          =   1935
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   1575
         Begin VB.ListBox lstAccepted 
            BackColor       =   &H00FFFFFF&
            Height          =   1620
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.TextBox txtWord 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   7320
         Width           =   4575
      End
      Begin VB.Label lblG 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Index           =   15
         Left            =   3540
         TabIndex        =   91
         Tag             =   "0"
         Top             =   5875
         Width           =   615
      End
      Begin VB.Label lblG 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Index           =   14
         Left            =   2660
         TabIndex        =   90
         Tag             =   "0"
         Top             =   5875
         Width           =   615
      End
      Begin VB.Label lblG 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Index           =   13
         Left            =   1750
         TabIndex        =   89
         Tag             =   "0"
         Top             =   5875
         Width           =   615
      End
      Begin VB.Label lblG 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Index           =   12
         Left            =   855
         TabIndex        =   88
         Tag             =   "0"
         Top             =   5875
         Width           =   615
      End
      Begin VB.Label lblG 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Index           =   11
         Left            =   3540
         TabIndex        =   87
         Tag             =   "0"
         Top             =   4980
         Width           =   615
      End
      Begin VB.Label lblG 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Index           =   10
         Left            =   2660
         TabIndex        =   86
         Tag             =   "0"
         Top             =   4980
         Width           =   615
      End
      Begin VB.Label lblG 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Index           =   9
         Left            =   1750
         TabIndex        =   85
         Tag             =   "0"
         Top             =   4980
         Width           =   615
      End
      Begin VB.Label lblG 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Index           =   8
         Left            =   855
         TabIndex        =   84
         Tag             =   "0"
         Top             =   4980
         Width           =   615
      End
      Begin VB.Label lblG 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Index           =   7
         Left            =   3540
         TabIndex        =   83
         Tag             =   "0"
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label lblG 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Index           =   6
         Left            =   2660
         TabIndex        =   82
         Tag             =   "0"
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label lblG 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Index           =   5
         Left            =   1750
         TabIndex        =   81
         Tag             =   "0"
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label lblG 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Index           =   4
         Left            =   855
         TabIndex        =   80
         Tag             =   "0"
         Top             =   4080
         Width           =   615
      End
      Begin VB.Label lblG 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Index           =   3
         Left            =   3540
         TabIndex        =   79
         Tag             =   "0"
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label lblG 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Index           =   2
         Left            =   2660
         TabIndex        =   78
         Tag             =   "0"
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label lblG 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Index           =   1
         Left            =   1750
         TabIndex        =   77
         Tag             =   "0"
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label lblG 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Index           =   0
         Left            =   855
         TabIndex        =   76
         Tag             =   "0"
         Top             =   3240
         Width           =   615
      End
      Begin VB.Image imgBog 
         Height          =   4380
         Left            =   300
         Picture         =   "frmMain.frx":0E66
         Top             =   2640
         Width           =   4275
      End
      Begin VB.Label lblInterval 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Interval Time:"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   2160
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblNFound 
         BackColor       =   &H008080FF&
         Caption         =   "Word Not Found"
         Height          =   255
         Left            =   3240
         TabIndex        =   50
         Top             =   7080
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Frame fraStatus 
      Caption         =   "Info"
      Height          =   2895
      Left            =   5280
      TabIndex        =   59
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      Begin VB.Label lblADesc 
         Height          =   2535
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblInfo 
         Caption         =   "Interval Time left:"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblTLeft 
         Alignment       =   1  'Right Justify
         Caption         =   "3:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   60
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Timer tmrF 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   400
      Left            =   9240
      Top             =   2880
   End
   Begin VB.Timer tmrF 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   500
      Left            =   8760
      Top             =   2880
   End
   Begin VB.Timer tmrF 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   500
      Left            =   8280
      Top             =   2880
   End
   Begin VB.Timer tmrF 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   500
      Left            =   7800
      Top             =   2880
   End
   Begin VB.Timer tmrR 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6960
      Top             =   2760
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   8415
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   14843
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Board"
            Key             =   "tBoard"
            Object.ToolTipText     =   "Game Board"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Key             =   "tHelp"
            Object.ToolTipText     =   "Help and Information"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   9480
      TabIndex        =   10
      Top             =   7560
      Width           =   735
   End
   Begin VB.TextBox txtC 
      Height          =   375
      Left            =   5280
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   8040
      Width           =   4935
   End
   Begin MSWinsockLib.Winsock wsck 
      Left            =   2760
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtChat 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   5280
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3000
      Width           =   4935
   End
   Begin VB.Frame fraCon 
      Caption         =   "Settings"
      Height          =   1695
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.TextBox txtServerPort 
         Height          =   375
         Left            =   3960
         TabIndex        =   7
         Text            =   "7200"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtPName 
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Text            =   "Player"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtServerIP 
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Text            =   "127.0.0.1"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "&Connect"
         Default         =   -1  'True
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblInfo 
         Caption         =   "Port:"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Player Name:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Server IP:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraScoreBoard 
      Caption         =   "Scoreboard"
      Height          =   2895
      Left            =   5280
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   4935
      Begin VB.ListBox lstP 
         Height          =   2595
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Image imgMedal 
      Height          =   480
      Index           =   6
      Left            =   7680
      Picture         =   "frmMain.frx":4E34
      Top             =   7440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMedal 
      Height          =   480
      Index           =   9
      Left            =   8280
      Picture         =   "frmMain.frx":56FE
      Top             =   7440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMedal 
      Height          =   480
      Index           =   8
      Left            =   7320
      Picture         =   "frmMain.frx":5FC8
      Top             =   7440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMedal 
      Height          =   480
      Index           =   7
      Left            =   6960
      Picture         =   "frmMain.frx":6892
      Top             =   7440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMedal 
      Height          =   480
      Index           =   5
      Left            =   5520
      Picture         =   "frmMain.frx":715C
      Top             =   7440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMedal 
      Height          =   480
      Index           =   0
      Left            =   6000
      Picture         =   "frmMain.frx":7A26
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMedal 
      Height          =   600
      Index           =   4
      Left            =   8640
      Picture         =   "frmMain.frx":8868
      Top             =   7440
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgMedal 
      Height          =   960
      Index           =   3
      Left            =   7680
      Picture         =   "frmMain.frx":9432
      Top             =   6960
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image imgMedal 
      Height          =   960
      Index           =   2
      Left            =   6960
      Picture         =   "frmMain.frx":AA7C
      Top             =   6960
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image imgMedal 
      Height          =   960
      Index           =   1
      Left            =   6240
      Picture         =   "frmMain.frx":C0C6
      Top             =   6960
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gameStr As String
Dim gameState As Integer
Dim rNum As Integer
Dim curTab As String
Dim numP As Integer
Dim SPlay As Boolean
Dim ply(0 To 1000) As PType
Dim bog(0 To 5, 0 To 5) As String
Dim lAccept(1 To 1000) As BogWord
Dim nAccept As Integer
Dim lAF(1 To 1000) As BogWord
Dim tnWords As Integer      'total no of words in board
Dim tnFound As Integer      'total no of words found
Dim nAF As Integer
Dim nInv As Integer
Dim lPlay(1 To 1000) As BogWord
Dim nPlay As Integer
Dim minLen As Integer
Dim wFound As Boolean
Dim startT As Long
Dim winPts As Integer
Dim TTS, TTR, TTI As Long
Dim curPlay As Integer
Dim topWords(1 To numTopWords) As TTopWord

Private Sub clbAwards_Click()
On Error Resume Next
    Dim aName As String
    Dim i, j, k As Integer
    aName = clbAwards.List(clbAwards.ListIndex)
    
    j = -1
    For i = 1 To numMedals
        If Medals(i).name = aName Then
            j = i
            Exit For
        End If
    Next i
    If j <> -1 Then
        lblADesc.Caption = Medals(j).desc
        lblADesc.Visible = True
    Else
        For i = 1 To 3
            If Prizes(i).name = aName Then
                j = i
                Exit For
            End If
        Next i
        If j <> -1 Then
            lblADesc.Caption = Prizes(j).desc
            lblADesc.Visible = True
        Else
            lblADesc.Visible = False
        End If
    End If
End Sub

Private Sub clbAwards_KeyDown(KeyCode As Integer, Shift As Integer)
    clbAwards_Click
End Sub

Private Sub clbAwards_LostFocus()
    lblADesc.Visible = False
End Sub

Private Sub clbInfo_Click()
    Dim aName As String
    Dim i, j, k As Integer
    aName = clbInfo.List(clbInfo.ListIndex)
    
    j = -1
    For i = 1 To numMedals
        If Medals(i).name = aName Then
            j = i
            Exit For
        End If
    Next i
    If j <> -1 Then
        lblMedalD.Caption = Medals(j).desc
    End If
End Sub

Private Sub clbInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    clbInfo_Click
End Sub

Private Sub cmdClear_Click()
    txtChat.Text = ""
End Sub

Private Sub cmdConnect_Click()
    cmdConnect.Enabled = False
    txtPName.Enabled = False
    txtServerIP.Enabled = False
    txtServerPort.Enabled = False
    
    txtPName.Text = Replace(txtPName.Text, "=", "")
    Me.Caption = "Boggle - " & txtPName.Text
    wsck.RemoteHost = txtServerIP.Text
    wsck.RemotePort = txtServerPort.Text
    wsck.Connect
End Sub

Private Sub lstTop_Click()
    Dim aName As String
    Dim dIndex As Long
    Dim i, j, k As Integer
    Dim a, b As Integer

    If DictLoad Then
    
        aName = lstTop.List(lstTop.ListIndex)
        
        dIndex = -1
        If aName <> "" Then
            a = InStr(1, aName, vbTab)
            b = InStr(a + 1, aName, vbTab)
            aName = Mid(aName, a + 1, b - a - 1)
            dIndex = SearchD(aName)
        End If
        If aName <> "" Then
            If dIndex = -1 Then
                lblADesc.Caption = "Word not found in dictionary."
            Else
                lblADesc.Caption = "'" & aName & "'" & vbCrLf & vbCrLf & DDef(dIndex)
            End If
            
            lblADesc.Visible = True
        Else
            lblADesc.Visible = False
        End If
    
    End If
End Sub

Private Sub lstTop_KeyDown(KeyCode As Integer, Shift As Integer)
    lstTop_Click
End Sub

Private Sub lstTop_LostFocus()
    lblADesc.Visible = False
End Sub

Private Function SearchD(wor As String) As Long
    Dim i As Long
    SearchD = -1
    
    'search in dict
    For i = 1 To numDWor
        If DWor(i).word = LCase(wor) Then
            SearchD = DWor(i).def
            Exit For
        End If
    Next i
    
End Function

Private Sub Form_Load()
    Dim i As Integer
    
    txtPName.Text = "Player" & (Rnd * 10002) Mod 10000
    curTab = "tBoard"
    curPlay = 0
    numP = 0
    SPlay = False
    ply(0).Con = 1
    lstAccepted.BackColor = RGB(245, 255, 245)
    lstFound.BackColor = RGB(255, 255, 245)
    lstInv.BackColor = RGB(255, 245, 245)
    lstPAccepted.BackColor = lstAccepted.BackColor
    lstPFound.BackColor = lstFound.BackColor
    
    lstTop.AddHeader 20, LeftJustify, ""
    lstTop.AddHeader 100, LeftJustify, "Word"
    lstTop.AddHeader 165, LeftJustify, "Found By:"
    
    ucrInfo.AddHeader 100, LeftJustify, "Number of letters"
    ucrInfo.AddHeader 40, LeftJustify, "Points"
    ucrInfo.AddItem "3" & vbTab & "1"
    ucrInfo.AddItem "4" & vbTab & "1"
    ucrInfo.AddItem "5" & vbTab & "2"
    ucrInfo.AddItem "6" & vbTab & "3"
    ucrInfo.AddItem "7" & vbTab & "5"
    ucrInfo.AddItem "8 or more" & vbTab & "11"
    
    For i = 0 To 9
        clbInfo.AddImage imgMedal(i).Picture
        clbAwards.AddImage imgMedal(i).Picture
    Next i
    
    For i = 1 To numMedals
        clbInfo.AddItem Medals(i).name, Medals(i).gNum
    Next i
End Sub

Private Sub tabCredits_Click()
    If tabCredits.SelectedItem.Key = "tCred" Then
        fraCredits.Visible = True
    Else
        fraCredits.Visible = False
    End If
End Sub

Private Sub tabMain_Click()
    If curTab <> tabMain.SelectedItem.Key Then
        ChangeTab tabMain.SelectedItem.Key
    End If
End Sub

Private Sub EndGame()
    txtWord.Enabled = False
End Sub

Private Sub ShowPTabs()
    Dim i As Integer
    
    If SPlay = False Then
        SPlay = True
        
        fraScoreBoard.Left = 7560
        fraScoreBoard.Width = 2655
        lstP.Width = 2415
        fraStatus.Visible = True
        lblADesc.Visible = False
        
        DoHM
        
        tabMain.Tabs.Add 3, "tStats", "Statistics"
        tabMain.Tabs.Item(3).ToolTipText = "End Game Statistics"
        
        i = 0
        tabMain.Tabs.Add , "tP" & i, ply(i).name
        tabMain.Tabs.Item(tabMain.Tabs.Count).Tag = i
        
        For i = 1 To numP
            If ply(i).Con = 1 Then  'connected
                tabMain.Tabs.Add , "tP" & i, ply(i).name
            End If
        Next i
        
        lblRNum.Caption = "Statistics for Round " & rNum & ":"
        
        tabMain.MultiRow = False
        'If tabMain.Tabs.Count > 4 Then tabMain.Style = tabButtons
        
        tabMain.Tabs.Item(3).Selected = True
        
    End If
End Sub

Private Sub DoHM() 'honourable mention
    Dim i, j, k As Integer
    Dim z As Integer
    Dim strWin As String
    
        lblHM.Caption = "None."
        
        addHM lblWinner.Caption & " won this round with " & winPts & " points."
        
        'solitaire
        z = 0
        strWin = ""
        For i = 0 To numP   'winner
            If ply(i).Con = 1 Then  'connected
                If ply(i).bAward(1) Then
                    z = z + 1
                    If z = 1 Then
                        strWin = ply(i).name
                    ElseIf z = 2 Then
                        strWin = strWin & " and " & ply(i).name
                    Else
                        strWin = ply(i).name & ", " & strWin
                    End If
                End If
            End If
        Next i
        If strWin <> "" Then
            addHM "Solitaire award conferred to " & strWin & " for performing outstandingly well. You are exceptional!"
        End If
        
        'prizes
        k = 0
        strWin = ""
        z = 0
        For i = 0 To numP   'winner
            If ply(i).Con = 1 Then  'connected
                If ply(i).Prize = 1 Then
                    z = z + 1
                    If z = 1 Then
                        strWin = ply(i).name
                    ElseIf z = 2 Then
                        strWin = strWin & " and " & ply(i).name
                    Else
                        strWin = ply(i).name & ", " & strWin
                    End If
                End If
            End If
        Next i
        If strWin <> "" Then
            addHM1 "Gold Star awardees: " & strWin
            k = 1
        End If
        
        strWin = ""
        z = 0
         For i = 0 To numP   'winner
            If ply(i).Con = 1 Then  'connected
                If ply(i).Prize = 2 Then
                    z = z + 1
                    If z = 1 Then
                        strWin = ply(i).name
                    ElseIf z = 2 Then
                        strWin = strWin & " and " & ply(i).name
                    Else
                        strWin = ply(i).name & ", " & strWin
                    End If
                End If
            End If
        Next i
        If strWin <> "" Then
            k = 1
            addHM1 "Silver Star awardees: " & strWin
        End If
        
        strWin = ""
        z = 0
        For i = 0 To numP   'winner
            If ply(i).Con = 1 Then  'connected
                If ply(i).Prize = 3 Then
                    z = z + 1
                    If z = 1 Then
                        strWin = ply(i).name
                    ElseIf z = 2 Then
                        strWin = strWin & " and " & ply(i).name
                    Else
                        strWin = ply(i).name & ", " & strWin
                    End If
                End If
            End If
        Next i
        If strWin <> "" Then
            k = 1
            addHM1 "Bronze Star awardees: " & strWin
        End If
        If k = 1 Then addHM1 " "
        
        'high flyer
        z = 0
        strWin = ""
        For i = 0 To numP   'winner
            If ply(i).Con = 1 Then  'connected
                If ply(i).bAward(3) Then
                    z = z + 1
                    If z = 1 Then
                        strWin = ply(i).name
                    ElseIf z = 2 Then
                        strWin = strWin & " and " & ply(i).name
                    Else
                        strWin = ply(i).name & ", " & strWin
                    End If
                End If
            End If
        Next i
        If strWin <> "" Then
            addHM "High Flyer award conferred to " & strWin & " for achieving the feat of obtaining a score of over 50 points or finding all words in the board. Good job!"
        End If
        
        'detective
        z = 0
        strWin = ""
        For i = 0 To numP   'winner
            If ply(i).Con = 1 Then  'connected
                If ply(i).bAward(4) Then
                    z = z + 1
                    If z = 1 Then
                        strWin = ply(i).name
                    ElseIf z = 2 Then
                        strWin = strWin & " and " & ply(i).name
                    Else
                        strWin = ply(i).name & ", " & strWin
                    End If
                End If
            End If
        Next i
        If strWin <> "" Then
            addHM "Detective award conferred to " & strWin & " for finding all 3 top valued words. Outstanding."
        End If
        
End Sub

Private Sub addHM(tos As String)
    If lblHM.Caption = "None." Then lblHM.Caption = ""
    lblHM.Caption = lblHM.Caption & tos & vbCrLf & vbCrLf
End Sub

Private Sub addHM1(tos As String)
    If lblHM.Caption = "None." Then lblHM.Caption = ""
    lblHM.Caption = lblHM.Caption & tos & vbCrLf
End Sub

Private Sub HidePTabs()
    Dim i, j As Integer
    If SPlay = True Then
        SPlay = False
        
        fraScoreBoard.Left = 5280
        fraScoreBoard.Width = 4935
        lstP.Width = 4695
        fraStatus.Visible = False
        For i = 0 To numP
            For j = 1 To tabMain.Tabs.Count
                If tabMain.Tabs.Item(j).Caption = ply(i).name Then
                    tabMain.Tabs.Remove j
                    Exit For
                End If
            Next j
        Next i
        tabMain.Tabs.Remove 3
        tabMain.Style = tabTabs
    End If
    tabMain_Click
End Sub

Private Sub ChangeTab(tabName As String)
    Dim i As Integer
    
    curTab = tabName

    Select Case tabName
        Case "tBoard"   'main board
            fraBoard.Visible = True
            fraStats.Visible = False
            fraPlayer.Visible = False
            fraHelp.Visible = False
        Case "tHelp"    'help
            fraBoard.Visible = False
            fraStats.Visible = False
            fraPlayer.Visible = False
            fraHelp.Visible = True
        Case "tStats"   'stats
            fraBoard.Visible = False
            fraStats.Visible = True
            fraPlayer.Visible = False
            fraHelp.Visible = False
        Case Else       'player stats
            fraBoard.Visible = False
            fraStats.Visible = False
            fraPlayer.Visible = True
            fraHelp.Visible = False
            ShowPFrame tabMain.SelectedItem.Caption
    End Select
End Sub

Private Sub ShowPFrame(pName As String)
    Dim i, j, k As Integer
    Dim pIndex As Integer
    Dim aStr As String
    
    pIndex = -1
    
    For i = 0 To numP
        If ply(i).Con = 1 And ply(i).name = pName Then
            pIndex = i
            Exit For
        End If
    Next i
    If pIndex <> -1 Then
        curPlay = pIndex
        lblStats.Caption = "Statistics of " & ply(curPlay).name & ":"
        lblPPoints.Caption = ply(curPlay).Score
        lblPPos.Caption = ply(curPlay).Pos
        aStr = " ("
        If tnWords = 0 Then
            aStr = ""
        Else
            j = Int((ply(curPlay).numFound * 100 / tnWords) + 0.5)
            aStr = aStr & j & "%)"
        End If
        lblPFound.Caption = ply(curPlay).numFound & aStr
        lblPSpeed.Caption = ply(curPlay).speed & " kpm"
        lblPMedals.Caption = ply(curPlay).numAwards
        
        SortList 4
        SortList 5
        LRefresh 4
        LRefresh 5
        LRefresh 6
    End If
    
    
End Sub

Private Sub FlashB(Index As Integer)
    '0 = accepted
    '1 = af
    '2 = invalid
    '3 = accepted -> af
    Select Case Index
        Case 0
            lstAccepted.BackColor = RGB(100, 255, 100)
            tmrF(0).Enabled = True
        Case 1
            lstFound.BackColor = RGB(255, 255, 100)
            tmrF(1).Enabled = True
        Case 2
            lstInv.BackColor = RGB(255, 100, 100)
            tmrF(2).Enabled = True
        Case 3
            lstAccepted.BackColor = RGB(255, 100, 100)
            tmrF(0).Enabled = True
    End Select
End Sub

Private Sub tmrF_Timer(Index As Integer)
    tmrF(Index).Enabled = False
    Select Case Index
        Case 0
            lstAccepted.BackColor = RGB(245, 255, 245)
        Case 1
            lstFound.BackColor = RGB(255, 255, 245)
        Case 2
            lstInv.BackColor = RGB(255, 245, 245)
        Case 3
            lblNFound.Visible = False
    End Select
End Sub

Private Sub tmrR_Timer()
    Dim curT As Long
    Dim timeLeft As Long
    Dim min As Integer
    Dim sec As Integer
    Dim strS As String
    Dim bCol, sCol As Long
    
    curT = GetTickCount
    
    bCol = vbWhite
    sCol = Me.BackColor
    
    Select Case gameState
        Case 1
            timeLeft = TTS - (curT - startT)
        Case 2
            timeLeft = TTR - (curT - startT)
            If timeLeft < 30000 Then
                bCol = RGB(255, 180, 100)
            End If
            If timeLeft < 10000 Then
                bCol = RGB(255, 100, 100)
            End If
            If timeLeft < 3000 Then
                bCol = vbRed
            End If
        Case 3 'interval time
            timeLeft = TTI - (curT - startT)
            bCol = vbCyan
            If timeLeft < 10000 Then
                sCol = RGB(255, 100, 100)
            End If
            If timeLeft < 5000 Then
                bCol = RGB(255, 100, 100)
            End If
    End Select
    
    If timeLeft < 0 Then timeLeft = 0
    timeLeft = timeLeft / 1000
    sec = timeLeft Mod 60
    min = (timeLeft - sec) / 60
    If sec = 0 Then
        strS = "00"
    ElseIf sec < 10 Then
        strS = "0" & sec
    Else
        strS = Str(sec)
    End If
    
    txtTime.Text = min & ":" & strS
    If bCol <> txtTime.BackColor Then txtTime.BackColor = bCol
    
    If gameState = 3 Then
        lblInterval.Visible = True
        lblIntervalT.Caption = "Interval Time remaining:   " & min & ":" & strS
        If sCol <> lblTLeft.BackColor Then lblTLeft.BackColor = sCol
        lblTLeft.Caption = Str(min) & ":" & strS
    End If
End Sub

Private Sub txtC_KeyPress(KeyAscii As Integer)
    Dim strC As String
    txtC.Text = Replace(txtC.Text, vbCrLf, "")
    If KeyAscii = vbKeyReturn Then
        strC = Replace(txtC.Text, "=", "!%")
        WSend "chat=" & strC
        'ProcessCmd txtC.Text
        txtC.Text = ""
    End If
End Sub

Private Sub txtChat_Change()
    txtChat.SelStart = Len(txtChat.Text)
    txtChat.SelLength = 0
End Sub

Private Sub txtWord_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtWord.Text = Replace(txtWord.Text, vbCrLf, "")
        If gameState = 2 Then
            If Len(txtWord.Text) >= minLen Then
                SubmitWord txtWord.Text
            End If
        End If
        txtWord.Text = ""
    End If
End Sub

Private Sub SubmitWord(strWord As String)
    Dim tmpStr As String
    
    tmpStr = UCase(strWord)
    tmpStr = Replace(tmpStr, "QU", "!")
    
    CheckValid tmpStr
    If wFound Then
        WSend "word=" & UCase(strWord)
    Else
        tmrF(3).Enabled = True
        lblNFound.Visible = True
    End If
End Sub

Private Sub CheckValid(strWord As String)
    wFound = False
    DFS strWord
End Sub

Private Sub DFS(wor As String)  'Depth First Search
    Dim i, j, k, X, Y, z As Integer
    Dim wHist(1 To 17) As sp_coord 'dfs history
    Dim wLen As Integer
    Dim nextC As sp_coord
    Dim skip As Boolean
    Dim nPass As Boolean

    For i = 1 To 17
        wHist(i).stopNum = -1
    Next i

    For i = 1 To 4
        For j = 1 To 4
        
            If bog(i, j) = Mid(wor, 1, 1) Then
                
                wLen = 1
                wHist(wLen).X = i
                wHist(wLen).Y = j
                
                Do
                    skip = False
                    Select Case wHist(wLen).stopNum
                        Case 1
                            GoTo go2
                        Case 2
                            GoTo go3
                        Case 3
                            GoTo go4
                        Case 4
                            GoTo go5
                        Case 5
                            GoTo go6
                        Case 6
                            GoTo go7
                        Case 7
                            GoTo go8
                        Case 8
                            GoTo gofin
                    End Select
                    
go1:
                    nextC.X = wHist(wLen).X - 1
                    nextC.Y = wHist(wLen).Y - 1
                    If bog(nextC.X, nextC.Y) = Mid(wor, wLen + 1, 1) And skip = False Then
                        nPass = True
                        For k = 1 To wLen - 1
                            If wHist(k).X = nextC.X And wHist(k).Y = nextC.Y Then
                                nPass = False
                                Exit For
                            End If
                        Next k
                        If nPass Then
                            skip = True
                            wHist(wLen).stopNum = 1
                            wLen = wLen + 1
                            wHist(wLen).X = nextC.X
                            wHist(wLen).Y = nextC.Y
                            If wLen = Len(wor) Then
                                wFound = True
                                Exit Do
                            End If
                        End If
                    End If
                    
go2:
                    nextC.X = wHist(wLen).X
                    nextC.Y = wHist(wLen).Y - 1
                    If bog(nextC.X, nextC.Y) = Mid(wor, wLen + 1, 1) And skip = False Then
                        nPass = True
                        For k = 1 To wLen - 1
                            If wHist(k).X = nextC.X And wHist(k).Y = nextC.Y Then
                                nPass = False
                                Exit For
                            End If
                        Next k
                        If nPass Then
                            skip = True
                            wHist(wLen).stopNum = 2
                            wLen = wLen + 1
                            wHist(wLen).X = nextC.X
                            wHist(wLen).Y = nextC.Y
                            If wLen = Len(wor) Then
                                wFound = True
                                Exit Do
                            End If
                        End If
                    End If
                    
go3:
                    nextC.X = wHist(wLen).X + 1
                    nextC.Y = wHist(wLen).Y - 1
                    If bog(nextC.X, nextC.Y) = Mid(wor, wLen + 1, 1) And skip = False Then
                        nPass = True
                        For k = 1 To wLen - 1
                            If wHist(k).X = nextC.X And wHist(k).Y = nextC.Y Then
                                nPass = False
                                Exit For
                            End If
                        Next k
                        If nPass Then
                            skip = True
                            wHist(wLen).stopNum = 3
                            wLen = wLen + 1
                            wHist(wLen).X = nextC.X
                            wHist(wLen).Y = nextC.Y
                            If wLen = Len(wor) Then
                                wFound = True
                                Exit Do
                            End If
                        End If
                    End If
                    
go4:
                    nextC.X = wHist(wLen).X - 1
                    nextC.Y = wHist(wLen).Y
                    If bog(nextC.X, nextC.Y) = Mid(wor, wLen + 1, 1) And skip = False Then
                        nPass = True
                        For k = 1 To wLen - 1
                            If wHist(k).X = nextC.X And wHist(k).Y = nextC.Y Then
                                nPass = False
                                Exit For
                            End If
                        Next k
                        If nPass Then
                            skip = True
                            wHist(wLen).stopNum = 4
                            wLen = wLen + 1
                            wHist(wLen).X = nextC.X
                            wHist(wLen).Y = nextC.Y
                            If wLen = Len(wor) Then
                                wFound = True
                                Exit Do
                            End If
                        End If
                    End If
                    
go5:
                    nextC.X = wHist(wLen).X + 1
                    nextC.Y = wHist(wLen).Y
                    If bog(nextC.X, nextC.Y) = Mid(wor, wLen + 1, 1) And skip = False Then
                        nPass = True
                        For k = 1 To wLen - 1
                            If wHist(k).X = nextC.X And wHist(k).Y = nextC.Y Then
                                nPass = False
                                Exit For
                            End If
                        Next k
                        If nPass Then
                            skip = True
                            wHist(wLen).stopNum = 5
                            wLen = wLen + 1
                            wHist(wLen).X = nextC.X
                            wHist(wLen).Y = nextC.Y
                            If wLen = Len(wor) Then
                                wFound = True
                                Exit Do
                            End If
                        End If
                    End If
                    
go6:
                    nextC.X = wHist(wLen).X - 1
                    nextC.Y = wHist(wLen).Y + 1
                    If bog(nextC.X, nextC.Y) = Mid(wor, wLen + 1, 1) And skip = False Then
                        nPass = True
                        For k = 1 To wLen - 1
                            If wHist(k).X = nextC.X And wHist(k).Y = nextC.Y Then
                                nPass = False
                                Exit For
                            End If
                        Next k
                        If nPass Then
                            skip = True
                            wHist(wLen).stopNum = 6
                            wLen = wLen + 1
                            wHist(wLen).X = nextC.X
                            wHist(wLen).Y = nextC.Y
                            If wLen = Len(wor) Then
                                wFound = True
                                Exit Do
                            End If
                        End If
                    End If
                    
go7:
                    nextC.X = wHist(wLen).X
                    nextC.Y = wHist(wLen).Y + 1
                    If bog(nextC.X, nextC.Y) = Mid(wor, wLen + 1, 1) And skip = False Then
                        nPass = True
                        For k = 1 To wLen - 1
                            If wHist(k).X = nextC.X And wHist(k).Y = nextC.Y Then
                                nPass = False
                                Exit For
                            End If
                        Next k
                        If nPass Then
                            skip = True
                            wHist(wLen).stopNum = 7
                            wLen = wLen + 1
                            wHist(wLen).X = nextC.X
                            wHist(wLen).Y = nextC.Y
                            If wLen = Len(wor) Then
                                wFound = True
                                Exit Do
                            End If
                        End If
                    End If
                    
go8:
                    nextC.X = wHist(wLen).X + 1
                    nextC.Y = wHist(wLen).Y + 1
                    If bog(nextC.X, nextC.Y) = Mid(wor, wLen + 1, 1) And skip = False Then
                        nPass = True
                        For k = 1 To wLen - 1
                            If wHist(k).X = nextC.X And wHist(k).Y = nextC.Y Then
                                nPass = False
                                Exit For
                            End If
                        Next k
                        If nPass Then
                            skip = True
                            wHist(wLen).stopNum = 8
                            wLen = wLen + 1
                            wHist(wLen).X = nextC.X
                            wHist(wLen).Y = nextC.Y
                            If wLen = Len(wor) Then
                                wFound = True
                                Exit Do
                            End If
                        End If
                    End If
                    
gofin:
                    
                    If Not skip Then
                        wHist(wLen).stopNum = -1
                        wLen = wLen - 1
                    End If
                    If wLen = 0 Then Exit Do
                Loop
            End If
            If wFound Then Exit For
        Next j
        If wFound Then Exit For
    Next i
End Sub

Private Sub wsck_Close()
    MsgBox "Connection with server lost.", vbCritical, "Error"
    End
End Sub

Private Sub wsck_Connect()
    InitCon
End Sub

Private Sub wsck_DataArrival(ByVal bytesTotal As Long)
    Dim gStr As String
    
    wsck.GetData gStr
    buf = buf & gStr
    Do
        a = InStr(1, buf, vbNullChar)
        If a = 0 Then Exit Do
        gStr = Left$(buf, a - 1)
        buf = Right$(buf, Len(buf) - a)
        ProcessCmd gStr
        DoEvents
    Loop
    
End Sub

Private Sub ProcessCmd(strC As String)
    Dim comD() As String
    Dim i, j, k As Integer
    Dim tmpStr As String
    Dim aStr As String
    Dim strZZ As String
    
    If strC = "" Then Exit Sub
    comD = Split(strC, "=")
    'AddChat "c=" & strC
    
    Select Case comD(0)
        Case "name?"
            WSend "name=" & txtPName.Text
        Case "con"
            AddChat "Connected to " & txtServerIP.Text & ":" & txtServerPort.Text & "."
        Case "echo"
            aStr = comD(1)
            aStr = Replace(aStr, "!%", "=")
            AddChat aStr
        Case "addP" 'add new player
            numP = numP + 1
            ply(numP).name = comD(1)
            ply(numP).Con = 1
            ply(numP).Index = numP
            RefreshP
        Case "leaveP" 'player left
            For i = 1 To numP
                If ply(i).name = comD(1) Then
                    ply(i).Con = 0
                End If
            Next i
            RefreshP
        Case "interval" 'interval
            lblInterval.Visible = True
            tmrR.Enabled = True
        Case "game" 'New game
            gameStr = comD(1)
            minLen = comD(2)
            rNum = comD(3)
            
            For i = 0 To 5
                For j = 0 To 5
                    bog(i, j) = vbNullChar
                Next j
            Next i
            
            For i = 1 To 16
                j = Int((i - 1) / 4) + 1
                k = i Mod 4
                If k = 0 Then k = 4
                aStr = Mid(gameStr, i, 1)
                If aStr = "Q" Then aStr = "!"
                bog(j, k) = aStr
            Next i
            For i = 1 To 16
                j = Int((i - 1) / 4) + 1
                k = i Mod 4
                If k = 0 Then k = 4
                aStr = bog(j, k)
                If aStr = "!" Then aStr = "Qu"
                If aStr = "Qu" Then
                    If lblG(i - 1).Tag = 0 Then
                        lblG(i - 1).FontSize = 22
                        lblG(i - 1).Left = lblG(i - 1).Left - 105
                        lblG(i - 1).Top = lblG(i - 1).Top + 20
                        lblG(i - 1).Tag = 1
                    End If
                Else
                    If lblG(i - 1).Tag = 1 Then
                        lblG(i - 1).FontSize = 28
                        lblG(i - 1).Left = lblG(i - 1).Left + 105
                        lblG(i - 1).Top = lblG(i - 1).Top - 20
                        lblG(i - 1).Tag = 0
                    End If

                End If
                lblG(i - 1).Caption = aStr
            Next i
            
            ply(0).Score = 0
            lblInterval.Visible = False
            HidePTabs
            For i = 1 To numTopWords
                topWords(i).word = ""
                topWords(i).foundBy = ""
                topWords(i).VaLue = 0
            Next i
            'reset lists
            Do
                If lstAccepted.ListCount = 0 Then Exit Do
                lstAccepted.RemoveItem 0
            Loop
            Do
                If lstFound.ListCount = 0 Then Exit Do
                lstFound.RemoveItem 0
            Loop
            Do
                If lstInv.ListCount = 0 Then Exit Do
                lstInv.RemoveItem 0
            Loop
            Do
                If lstTop.ListCount = 0 Then Exit Do
                lstTop.RemoveItem 0
            Loop
            fraAccepted.Caption = "Accepted (" & lstAccepted.ListCount & ")"
            fraInvalid.Caption = "Invalid (" & lstInv.ListCount & ")"
            fraFound.Caption = "Already Found (" & lstFound.ListCount & ")"
            nAccept = 0
            nAF = 0
            nInv = 0
            winPts = 0
            tnWords = 0
            tnFound = 0
            
            For i = 0 To numP
                ply(i).numAF = 0
                ply(i).numOrig = 0
                For j = 1 To numMedals
                    ply(i).bAward(j) = False
                    ply(i).Prize = 0
                Next j
            Next i
            For i = 1 To numTopWords
                topWords(i).word = ""
                topWords(i).foundBy = ""
                topWords(i).VaLue = 0
            Next i
            
            Do
                If lstFound.ListCount = 0 Then Exit Do
                lstFound.RemoveItem 0
            Loop
            
            tabMain.Tabs(1).Selected = True
            tmrR.Enabled = True
            
            If chkSounds.VaLue = 1 Then
                sndPlaySound App.Path & "\data\1.wav", 1
            End If
            
            txtWord.Locked = False
            txtWord.Text = ""
            txtWord.SetFocus
        Case "endG" 'end Game
            If chkSounds.VaLue = 1 Then
                sndPlaySound App.Path & "\data\2.wav", 1
            End If
            txtWord.Locked = True
        Case "stats" 'stats
        'stats
        ''''''
        'Player Name
        'Position
        'Points
        'Words Found
        'Typing Speed
        'num of awards
        
            k = -1
            For i = 0 To numP
                If ply(i).Con = 1 And ply(i).name = comD(1) Then
                    k = i
                    Exit For
                End If
            Next i
            If k <> -1 Then
                ply(k).Pos = comD(2)
                ply(k).Score = comD(3)
                ply(k).numFound = comD(4)
                ply(k).speed = comD(5)
                ply(k).numAwards = comD(6)
            End If
        Case "stats1"
        'stats1
        '''''''
        'winner
        'winner points
        'total number of words in board
        'number of words found

            lblWinner.Caption = comD(1)
            winPts = comD(2)
            tnWords = comD(3)
            tnFound = comD(4)
            aStr = Str(tnFound)
            If tnWords <> 0 Then
                i = Int(tnFound * 100 / tnWords)
                aStr = aStr & " (" & i & "%)"
            End If
            lblWFound.Caption = aStr
            
        Case "award"
        'award
        ''''''
        'name of player
        'medal number
            j = Val(comD(2))
            For i = 0 To numP
                If ply(i).name = comD(1) And ply(i).Con = 1 Then
                    ply(i).bAward(j) = True
                    Exit For
                End If
            Next i
        Case "prize"
        'prize
        ''''''
        'name of player
        'prize number
            j = Val(comD(2))
            For i = 0 To numP
                If ply(i).name = comD(1) And ply(i).Con = 1 Then
                    ply(i).Prize = j
                    Exit For
                End If
            Next i
        Case "top"
            topWords(comD(1)).VaLue = comD(2)
            topWords(comD(1)).word = comD(3)
            topWords(comD(1)).foundBy = comD(4)
            
        'top50
        ''''''
        'ranking
        'word
        'foundby
    
        Case "topdone"
             k = tnWords
            If k > 50 Then k = 50
            For i = 1 To k
                'strZZ = Str(i)
                'If strZZ < 10 Then strZZ = " " & strZZ
                'strZZ = strZZ & ". "
                lstTop.AddItem Str(topWords(i).VaLue) & vbTab & topWords(i).word & vbTab & topWords(i).foundBy
            Next i
            
            ShowPTabs
        Case "eOK"  'valid word
            k = -1
            For i = 0 To numP
                If ply(i).Con = 1 And ply(i).name = comD(1) Then
                    k = i
                    Exit For
                End If
            Next i
            If k <> -1 Then
                ply(k).numOrig = ply(k).numOrig + 1
                ply(k).lOrig(ply(k).numOrig).VaLue = comD(2)
                ply(k).lOrig(ply(k).numOrig).word = comD(3)
                SortList 4
                LRefresh 4
            End If
        Case "eAF"  'af word
            k = -1
            For i = 0 To numP
                If ply(i).Con = 1 And ply(i).name = comD(1) Then
                    k = i
                    Exit For
                End If
            Next i
            If k <> -1 Then
                ply(k).numAF = ply(k).numAF + 1
                ply(k).lAF(ply(k).numAF).VaLue = comD(2)
                ply(k).lAF(ply(k).numAF).word = comD(3)
                SortList 5
                LRefresh 5
            End If
        Case "wOK"  'valid word
            nAccept = nAccept + 1
            lAccept(nAccept).VaLue = comD(1)
            lAccept(nAccept).word = comD(2)
            SortList 1
            LRefresh 1
            FlashB 0
            fraAccepted.Caption = "Accepted (" & lstAccepted.ListCount & ")"
        Case "wInv" 'invalid word
            lstInv.AddItem comD(1), 0
            fraInvalid.Caption = "Invalid (" & lstInv.ListCount & ")"
            nInv = nInv + 1
            FlashB 2
        Case "wAF" 'already found
            nAF = nAF + 1
            lAF(nAF).VaLue = comD(1)
            lAF(nAF).word = comD(2)
            SortList 2
            LRefresh 2
            FlashB 1
            fraFound.Caption = "Already Found (" & lstFound.ListCount & ")"
        Case "mAF" 'make valid word already found
        
            nAF = nAF + 1
            lAF(nAF).VaLue = comD(1)
            lAF(nAF).word = comD(2)
            WDestroy comD(2)
            SortList 2
            LRefresh 1
            LRefresh 2
        
            FlashB 1
            FlashB 3
            fraFound.Caption = "Already Found ( " & lstFound.ListCount & ")"
            fraAccepted.Caption = "Accepted ( " & lstAccepted.ListCount & ")"
        Case "score" 'update score
            For i = 0 To numP
                If ply(i).name = comD(1) And ply(i).Con = 1 Then
                    ply(i).Score = comD(2)
                    Exit For
                End If
            Next i
            RefreshP
        Case "TStart"   'start
            startT = GetTickCount
            TTS = comD(1)
            If tmrR.Enabled = False Then
                tmrR.Enabled = True
            End If
            gameState = 1
        Case "TLeft"
            startT = GetTickCount
            TTR = comD(1)
            gameState = 2
        Case "TInt"
            startT = GetTickCount
            TTI = comD(1)
            If tmrR.Enabled = False Then
                tmrR.Enabled = True
            End If
            gameState = 3
        Case "state"
            gameState = Val(comD(1))
    End Select
End Sub

Private Sub RefreshP()
    Dim i, j As Integer
    
    nPlay = 0
    For i = 0 To numP
        If ply(i).Con = 1 Or i = 0 Then
            nPlay = nPlay + 1
            lPlay(nPlay).word = ply(i).name
            lPlay(nPlay).VaLue = ply(i).Score
        End If
    Next i
    
    SortList 3
    
    Do
        If lstP.ListCount = 0 Then Exit Do
        lstP.RemoveItem 0
    Loop

    For i = 1 To nPlay
        lstP.AddItem lPlay(i).VaLue & vbTab & lPlay(i).word
    Next i
End Sub

Private Sub WSend(tos As String)
    wsck.SendData tos & vbNullChar
End Sub

Private Sub AddChat(tos As String)
    txtChat.Text = txtChat.Text & tos & vbCrLf
End Sub

Private Sub InitCon()
    ply(0).name = txtPName.Text
    fraCon.Visible = False
    fraScoreBoard.Visible = True
    nPlay = -1
    RefreshP
End Sub

Private Sub SortList(Index As Integer) ' sort list using bubblesort
    'Index =
    '1  accepted
    '2  already found
    '3  scoreboard
    '4  paccepted
    '5  pAF
    
    Dim i, j As Integer
    Dim X, Y, z As Integer
    Dim sto As Boolean
    Dim tmpB As BogWord
    Dim doSwap As Boolean
    
    Select Case Index
        Case 1
            For i = 1 To nAccept
                sto = True
                For j = 1 To (nAccept - 1)
                    If Len(lAccept(j + 1).word) > Len(lAccept(j).word) Then
                        sto = False
                        tmpB = lAccept(j)
                        lAccept(j) = lAccept(j + 1)
                        lAccept(j + 1) = tmpB
                    ElseIf Len(lAccept(j + 1).word) = Len(lAccept(j).word) Then
                        doSwap = False
                        For X = 1 To Len(lAccept(j).word)
                            If Asc(Mid(lAccept(j).word, X, 1)) > Asc(Mid(lAccept(j + 1).word, X, 1)) Then
                                doSwap = True
                                Exit For
                            End If
                            If Asc(Mid(lAccept(j).word, X, 1)) < Asc(Mid(lAccept(j + 1).word, X, 1)) Then Exit For
                        Next X
                        If doSwap Then
                            sto = False
                            tmpB = lAccept(j)
                            lAccept(j) = lAccept(j + 1)
                            lAccept(j + 1) = tmpB
                        End If
                    End If
                Next j
                If sto Then Exit For
            Next i
        Case 2
            For i = 1 To nAF
                sto = True
                For j = 1 To (nAF - 1)
                    If Len(lAF(j + 1).word) > Len(lAF(j).word) Then
                        sto = False
                        tmpB = lAF(j)
                        lAF(j) = lAF(j + 1)
                        lAF(j + 1) = tmpB
                    ElseIf Len(lAF(j + 1).word) = Len(lAF(j).word) Then
                        doSwap = False
                        For X = 1 To Len(lAF(j).word)
                            If Asc(Mid(lAF(j).word, X, 1)) > Asc(Mid(lAF(j + 1).word, X, 1)) Then
                                doSwap = True
                                Exit For
                            End If
                            If Asc(Mid(lAF(j).word, X, 1)) < Asc(Mid(lAF(j + 1).word, X, 1)) Then Exit For
                        Next X
                        If doSwap Then
                            sto = False
                            tmpB = lAF(j)
                            lAF(j) = lAF(j + 1)
                            lAF(j + 1) = tmpB
                        End If
                    End If
                Next j
                If sto Then Exit For
            Next i
        Case 3
            For i = 1 To nPlay
                sto = True
                For j = 1 To (nPlay - 1)
                    If lPlay(j + 1).VaLue > lPlay(j).VaLue Then
                        sto = False
                        tmpB = lPlay(j)
                        lPlay(j) = lPlay(j + 1)
                        lPlay(j + 1) = tmpB
                    ElseIf lPlay(j + 1).VaLue = lPlay(j).VaLue Then
                        doSwap = False
                        For X = 1 To Len(lPlay(j).word)
                            If Asc(Mid(lPlay(j).word, X, 1)) > Asc(Mid(lPlay(j + 1).word, X, 1)) Then
                                doSwap = True
                                Exit For
                            End If
                            If Asc(Mid(lPlay(j).word, X, 1)) < Asc(Mid(lPlay(j + 1).word, X, 1)) Then Exit For
                        Next X
                        If doSwap Then
                            sto = False
                            tmpB = lPlay(j)
                            lPlay(j) = lPlay(j + 1)
                            lPlay(j + 1) = tmpB
                        End If
                    End If
                Next j
                If sto Then Exit For
            Next i
        Case 4
            For i = 1 To ply(curPlay).numOrig
                sto = True
                For j = 1 To (ply(curPlay).numOrig - 1)
                    If Len(ply(curPlay).lOrig(j + 1).word) > Len(ply(curPlay).lOrig(j).word) Then
                        sto = False
                        tmpB = ply(curPlay).lOrig(j)
                        ply(curPlay).lOrig(j) = ply(curPlay).lOrig(j + 1)
                        ply(curPlay).lOrig(j + 1) = tmpB
                    ElseIf Len(ply(curPlay).lOrig(j + 1).word) = Len(ply(curPlay).lOrig(j).word) Then
                        doSwap = False
                        For X = 1 To Len(ply(curPlay).lOrig(j).word)
                            If Asc(Mid(ply(curPlay).lOrig(j).word, X, 1)) > Asc(Mid(ply(curPlay).lOrig(j + 1).word, X, 1)) Then
                                doSwap = True
                                Exit For
                            End If
                            If Asc(Mid(ply(curPlay).lOrig(j).word, X, 1)) < Asc(Mid(ply(curPlay).lOrig(j + 1).word, X, 1)) Then Exit For
                        Next X
                        If doSwap Then
                            sto = False
                            tmpB = ply(curPlay).lOrig(j)
                            ply(curPlay).lOrig(j) = ply(curPlay).lOrig(j + 1)
                            ply(curPlay).lOrig(j + 1) = tmpB
                        End If
                    End If
                Next j
                If sto Then Exit For
            Next i
        Case 5
            For i = 1 To ply(curPlay).numAF
                sto = True
                For j = 1 To (ply(curPlay).numAF - 1)
                    If Len(ply(curPlay).lAF(j + 1).word) > Len(ply(curPlay).lAF(j).word) Then
                        sto = False
                        tmpB = ply(curPlay).lAF(j)
                        ply(curPlay).lAF(j) = ply(curPlay).lAF(j + 1)
                        ply(curPlay).lAF(j + 1) = tmpB
                    ElseIf Len(ply(curPlay).lAF(j + 1).word) = Len(ply(curPlay).lAF(j).word) Then
                        doSwap = False
                        For X = 1 To Len(ply(curPlay).lAF(j).word)
                            If Asc(Mid(ply(curPlay).lAF(j).word, X, 1)) > Asc(Mid(ply(curPlay).lAF(j + 1).word, X, 1)) Then
                                doSwap = True
                                Exit For
                            End If
                            If Asc(Mid(ply(curPlay).lAF(j).word, X, 1)) < Asc(Mid(ply(curPlay).lAF(j + 1).word, X, 1)) Then Exit For
                        Next X
                        If doSwap Then
                            sto = False
                            tmpB = ply(curPlay).lAF(j)
                            ply(curPlay).lAF(j) = ply(curPlay).lAF(j + 1)
                            ply(curPlay).lAF(j + 1) = tmpB
                        End If
                    End If
                Next j
                If sto Then Exit For
            Next i
    End Select
End Sub

Private Sub LRefresh(Index As Integer) 'refresh list
'Index =
'1  accepted
'2  already found
'4  paccepted
'5  paf
'6 awards

    Dim wor As String
    Dim VaLue As Integer
    Dim i, j As Integer
    Dim tmpStr As String
    Select Case Index
        Case 1
            Do
                If lstAccepted.ListCount = 0 Then Exit Do
                lstAccepted.RemoveItem 0
            Loop
            For j = 1 To nAccept
                i = 5 - Len(Str(lAccept(j).VaLue))
                tmpStr = Str(lAccept(j).VaLue) & Space(i) & lAccept(j).word
                lstAccepted.AddItem tmpStr
            Next j
        Case 2
            Do
                If lstFound.ListCount = 0 Then Exit Do
                lstFound.RemoveItem 0
            Loop
            For j = 1 To nAF
                i = 5 - Len(Str(lAF(j).VaLue))
                tmpStr = Str(lAF(j).VaLue) & Space(i) & lAF(j).word
                lstFound.AddItem tmpStr
            Next j
        Case 4
            Do
                If lstPAccepted.ListCount = 0 Then Exit Do
                lstPAccepted.RemoveItem 0
            Loop
            For j = 1 To ply(curPlay).numOrig
                i = 5 - Len(Str(ply(curPlay).lOrig(j).VaLue))
                tmpStr = Str(ply(curPlay).lOrig(j).VaLue) & Space(i) & ply(curPlay).lOrig(j).word
                lstPAccepted.AddItem tmpStr
            Next j
            fraPAccepted.Caption = "Accepted (" & lstPAccepted.ListCount & ")"
        Case 5
            Do
                If lstPFound.ListCount = 0 Then Exit Do
                lstPFound.RemoveItem 0
            Loop
            For j = 1 To ply(curPlay).numAF
                i = 5 - Len(Str(ply(curPlay).lAF(j).VaLue))
                tmpStr = Str(ply(curPlay).lAF(j).VaLue) & Space(i) & ply(curPlay).lAF(j).word
                lstPFound.AddItem tmpStr
            Next j
            fraPFound.Caption = "Already found (" & lstPFound.ListCount & ")"
        Case 6 'awards
            Do
                If clbAwards.LCount = 0 Then Exit Do
                clbAwards.RemoveItem 0
            Loop
            If ply(curPlay).Prize > 0 Then
                clbAwards.AddItem Prizes(ply(curPlay).Prize).name, Prizes(ply(curPlay).Prize).gNum
            End If
            For j = 1 To numMedals
                If ply(curPlay).bAward(j) Then
                    clbAwards.AddItem Medals(j).name, Medals(j).gNum
                End If
            Next j
    End Select
    
End Sub

Private Sub WDestroy(wor As String) 'destroy word in lstAccepted
    Dim ans, i As Integer
    ans = -1
    
    For i = 1 To nAccept
        If lAccept(i).word = wor Then
            nAccept = nAccept - 1
            ans = i
            Exit For
        End If
    Next i
    If ans <> -1 Then
        For i = ans To nAccept
            lAccept(i) = lAccept(i + 1)
        Next i
    End If
End Sub

Private Sub wsck_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If Number = 10061 Then
        MsgBox "Unable to connect." & vbCrLf & "Check that the entered IP and port is correct.", vbCritical
        wsck.Close
        cmdConnect.Enabled = True
        txtPName.Enabled = True
        txtServerIP.Enabled = True
        txtServerPort.Enabled = True
    End If
End Sub
