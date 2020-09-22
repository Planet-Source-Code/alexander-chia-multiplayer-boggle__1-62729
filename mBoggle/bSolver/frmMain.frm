VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Boggle Solver"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtIP 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   975
   End
   Begin MSWinsockLib.Winsock sck 
      Left            =   3360
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtPort 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Text            =   "7230"
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblInfo 
      Caption         =   "IP:"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblInfo 
      Caption         =   "Port:"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InProgress As Boolean
Dim op As Integer
Dim minLen As Integer
Dim nAns As Integer
Dim lAns(1 To 5000) As String

Private Sub BSolve(iWord As String)
    Dim BWord As iWord
    Dim i, j As Integer
    Dim ff As Integer
    Dim x, y, z As Integer
    Dim sto As Boolean
    Dim doSwap As Boolean
    Dim aLet As String
    Set BGrid = New IGrid
    
    minLen = 3
    
    BGrid.CreateGrid 4, 4, minLen, False
    
    nAns = 0
    For i = 1 To 4
        For j = 1 To 4
            aLet = Mid(iWord, (i - 1) * 4 + j, 1)
            If aLet = "Q" Then aLet = "!"
            BGrid.AddLetter aLet, i, j
        Next j
    Next i
    BGrid.GetWords
    

    
    For Each BWord In BGrid
        nAns = nAns + 1
        lAns(nAns) = Replace(BWord.Text, "!", "qu")
    Next
    
    'sort
            For i = 1 To nAns
                sto = True
                For j = 1 To (nAns - 1)
                    If Len(lAns(j + 1)) > Len(lAns(j)) Then
                        sto = False
                        tmpB = lAns(j)
                        lAns(j) = lAns(j + 1)
                        lAns(j + 1) = tmpB
                    ElseIf Len(lAns(j + 1)) = Len(lAns(j)) Then
                        doSwap = False
                        For x = 1 To Len(lAns(j))
                            If Asc(Mid(lAns(j), x, 1)) > Asc(Mid(lAns(j + 1), x, 1)) Then
                                doSwap = True
                                Exit For
                            End If
                            If Asc(Mid(lAns(j), x, 1)) < Asc(Mid(lAns(j + 1), x, 1)) Then Exit For
                        Next x
                        If doSwap Then
                            sto = False
                            tmpB = lAns(j)
                            lAns(j) = lAns(j + 1)
                            lAns(j + 1) = tmpB
                        End If
                    End If
                Next j
                If sto Then Exit For
            Next i
        
    ff = FreeFile
    Open "ans.out" For Output As #ff
        For i = 1 To nAns
            Print #ff, lAns(i)
        Next i
    Close #ff
    
    Set BGrid = Nothing
End Sub

Private Sub cmdConnect_Click()

    txtIP.Enabled = False
    txtPort.Enabled = False
    cmdConnect.Enabled = False
    
    sck.RemoteHost = txtIP.Text
    sck.RemotePort = txtPort.Text
    sck.Connect
End Sub

Private Sub Form_Load()
    InProgress = False
    buf = ""
    minLen = 3
End Sub

Private Sub sck_Close()
    MsgBox "Connection with server lost.", vbCritical
    End
End Sub

Private Sub sck_DataArrival(ByVal bytesTotal As Long)
    Dim gStr As String
    Dim a As Integer
    sck.GetData gStr
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
    Dim sStr As String
    
    comD = Split(strC, "=")
    
    Select Case comD(0)
        Case "pass"
            CSend "Pass=" & PW
        Case "solve"
            minLen = Val(comD(1))
            sStr = comD(2)
            If Len(sStr) = 16 And Not InProgress Then
                InProgress = True
                BSolve sStr
                CSend "done"
                InProgress = False
            End If
    End Select
    
End Sub

Private Sub CSend(tos As String)
    sck.SendData tos & vbNullChar
End Sub

Private Sub sck_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If Number = 10061 Then
        MsgBox "Unable to connect." & vbCrLf & "Check that the entered IP and port is correct.", vbCritical
        sck.Close
    txtIP.Enabled = True
    txtPort.Enabled = True
    cmdConnect.Enabled = True
    End If
End Sub
