VERSION 5.00
Begin VB.Form frmLoading 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading Boggle"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrS 
      Interval        =   500
      Left            =   2040
      Top             =   480
   End
   Begin VB.Label lblDEF 
      BackStyle       =   0  'Transparent
      Caption         =   "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblABC 
      BackStyle       =   0  'Transparent
      Caption         =   "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label lblPText 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   760
      Width           =   735
   End
   Begin VB.Label lblP 
      BackColor       =   &H00C00000&
      Height          =   315
      Left            =   150
      TabIndex        =   2
      Top             =   750
      Width           =   15
   End
   Begin VB.Label lblPBack 
      BackColor       =   &H8000000C&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label lblInfo 
      Caption         =   "Loading..."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aProceed As Integer

Private Sub StartLoad()
On Error GoTo errdo
    Dim ff As Integer
    Dim i, j, k As Integer
    Dim LenF As Long
    Dim curF As Long
    Dim tPerc As Integer
    Dim strWor As String
    Dim X, Y, z As Long
    Dim tState As Integer
    Dim curAlp As String
    Dim strIn As String
    Dim strB() As String
    
    i = 0
    j = 0
    X = 0 'word count
    Y = 0 'def count
    tState = 0 '0=words,1=def
    curAlp = "A"
    lblABC.Caption = "A"
    UpdateP 0
    
    ff = FreeFile
    X = 0
    Y = 0
    
    'load words
    lblInfo.Caption = "Loading words.."
    tPerc = numDWor / 300
    Open App.Path & "\data\Def1.dat" For Input As #ff
        Do
            If EOF(ff) Then Exit Do
            Line Input #ff, strIn
            strB = Split(strIn, " ")
            X = X + 1
            If X Mod tPerc = 0 Then
                UpdateP (X * 50 / numDWor)
            End If
            DWor(X).word = strB(0)
            If Asc(Left(UCase(DWor(X).word), 1)) = Asc(curAlp) + 1 Then
                curAlp = Chr(Asc(curAlp) + 1)
                lblABC.Caption = lblABC.Caption & curAlp
            End If
            DWor(X).def = Val(strB(1))
        Loop
    Close #ff
       
    UpdateP 50
    lblDEF.Visible = True
    lblDEF.Caption = "A"
    curAlp = "A"
    
    'load defs
    lblInfo.Caption = "Loading definitions.."
    tPerc = numDDef / 300
    
    ff = FreeFile
    Open App.Path & "\data\Def2.dat" For Input As #ff
        Do
            If EOF(ff) Then Exit Do
            Line Input #ff, strIn
            strB = Split(strIn, " ")
            k = InStr(1, strIn, " ")
            Y = Y + 1
            If Y Mod tPerc = 0 Then
                UpdateP 50 + (Y * 50 / numDDef)
            End If
            DDef(Val(strB(0))) = Right(strIn, Len(strIn) - k)
            If Asc(Left(UCase(strB(1)), 1)) = Asc(curAlp) + 1 Then
                curAlp = Chr(Asc(curAlp) + 1)
                lblDEF.Caption = lblDEF.Caption & curAlp
            End If
        Loop
    Close #ff
    
    UpdateP 100

    DictLoad = True
    Finish
    Exit Sub
errdo:
    DictLoad = False
    MsgBox "Error loading dictionary." & vbCrLf & "Some functions may not work as a result.", vbExclamation, "Error"
    Finish
End Sub

Private Sub Form_Load()
    lblABC.Caption = ""
    lblDEF.Caption = ""
    aProceed = 0
    'frmMain.Show
End Sub

Private Sub Finish()
    aProceed = 1
    frmMain.Show
    Unload Me
End Sub

Private Sub UpdateP(Progress As Double)
    lblP.Width = 4275 / 100 * Progress
    lblPText.Caption = Int(Progress) & "%"
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If aProceed = 0 Then End
End Sub

Private Sub tmrS_Timer()
    tmrS.Enabled = False
    StartLoad
End Sub
