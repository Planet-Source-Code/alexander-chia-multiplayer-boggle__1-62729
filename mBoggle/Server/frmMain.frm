VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Boggle Server"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7755
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNextRound 
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox txtBSolve 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Text            =   "7230"
      Top             =   600
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock sckBListen 
      Left            =   720
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrLoad 
      Interval        =   200
      Left            =   1920
      Top             =   3600
   End
   Begin VB.Timer tmrR 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   3360
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Log"
      Height          =   255
      Left            =   6720
      TabIndex        =   4
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox txtLog 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   2880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   120
      Width           =   4695
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start server"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtSPort 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Text            =   "7200"
      Top             =   120
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock sckR 
      Left            =   1920
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckMain 
      Index           =   0
      Left            =   2160
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckBSolve 
      Left            =   840
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblNR 
      Caption         =   "Next Round:"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Caption         =   "Solver Port:"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblInfo 
      Caption         =   "Server Port:"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim numP As Integer
Dim gameStart As Boolean
Dim gameState As Integer
Dim minLen As Integer
Dim ply(1 To 1000) As PType
Dim startT As Long
Dim TTS As Long 'time to start
Dim TTR As Long 'length of round
Dim TTI As Long 'interval time
Dim rNum As Integer 'round number
Dim nRank As Integer
Dim lRank(1 To 1000) As Integer
Dim genState As Integer 'generate bogboard state 0=ungen 1=gen,sent 2=solved
Dim numTop As Integer 'number of top words
Dim topWords(1 To 10000) As String
Dim topFound(1 To 10000) As Boolean

Private Sub AddLog(tos As String)
    txtLog.Text = txtLog.Text & Time & vbTab & tos & vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
    txtLog.SelLength = 0
End Sub

Private Sub cmdClear_Click()
    txtLog.Text = ""
End Sub

Private Sub cmdStart_Click()
    txtSPort.Enabled = False
    txtBSolve.Enabled = False
    sckR.LocalPort = txtSPort.Text
    sckR.Listen
    sckBListen.LocalPort = txtBSolve.Text
    sckBListen.Listen
    cmdStart.Enabled = False
    AddLog "Server started at " & Time & "," & Date
End Sub

Private Sub Form_Load()
    numP = 0
    gameStart = False
    nRank = 0
    genState = 0
    AddLog "Loading dictionary from Word.lst..."
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim i As Integer
    For i = 1 To numP
        sckMain(numP).Close
    Next i
    sckR.Close
End Sub

Private Sub sckBListen_ConnectionRequest(ByVal requestID As Long)
    If sckBSolve.State = 2 Then sckBSolve.Close
    sckBSolve.Accept requestID
    SSend "pass"
End Sub

Private Sub sckBSolve_Close()
    sckBSolve.Close
    AddLog "**Boggle Solver Disconnected."
    sckBSolve.Listen
End Sub

Private Sub sckBSolve_DataArrival(ByVal bytesTotal As Long)
    Dim gStr As String
    Dim a As Integer
    
    sckBSolve.GetData gStr
    buf1 = buf1 & gStr
    Do
        a = InStr(1, buf1, vbNullChar)
        If a = 0 Then Exit Do
        gStr = Left$(buf1, a - 1)
        buf1 = Right$(buf1, Len(buf1) - a)
        ProcessCmd1 gStr
        DoEvents
    Loop
End Sub

Private Sub ProcessCmd1(strC As String)
    Dim comD() As String
    Dim ff As Integer
    Dim i As Integer
    Dim inWord As String
    
    comD = Split(strC, "=")
    
    Select Case comD(0)
        Case "Pass"
            If comD(1) = PW Then
                AddLog "**Boggle Solver Connected."
            End If
        Case "done" 'solved!
            
            For i = 1 To 30
                topWords(i) = ""
            Next i
            
            ff = FreeFile
            numTop = 0
            Open "ans.out" For Input As #ff
                Do
                    If EOF(ff) Then Exit Do
                    numTop = numTop + 1
                    Line Input #ff, inWord
                    topWords(numTop) = UCase(inWord)
                    topFound(numTop) = False
                Loop
            Close #ff
            
            AddLog "Round " & rNum & " board solved."
            genState = 2
    End Select
    
End Sub

Private Sub sckMain_Close(Index As Integer)
    If ply(Index).Con = 1 Then
        ply(Index).Con = 0
        AddLog ply(Index).Name & " has left the game."
        WSendAll 0, "leaveP=" & ply(Index).Name
        WSendAll 0, "echo=**" & ply(Index).Name & " has left the game."
    End If
End Sub

Private Sub sckMain_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim gStr As String
    Dim a As Integer
    sckMain(Index).GetData gStr
    buf = buf & gStr
    Do
        a = InStr(1, buf, vbNullChar)
        If a = 0 Then Exit Do
        gStr = Left$(buf, a - 1)
        buf = Right$(buf, Len(buf) - a)
        ProcessCmd Index, gStr
        DoEvents
    Loop
End Sub

Private Function TimeSTR(milli As Long) As String
    Dim hr, min, sec As Integer
    Dim millisecs As Long
    
    millisecs = milli
    TimeSTR = "0 sec"
    hr = 0
    min = 0
    sec = 0
    If millisecs > 0 Then
        millisecs = millisecs / 1000
        sec = millisecs Mod 60
        millisecs = millisecs / 60
        min = millisecs Mod 60
        millisecs = millisecs / 60
        hr = millisecs
        TimeSTR = ""
        If hr <> 0 Then TimeSTR = TimeSTR & hr & " hr "
        If min <> 0 Then TimeSTR = TimeSTR & min & " min "
        If sec <> 0 Then TimeSTR = TimeSTR & sec & " sec "
        
        If Len(TimeSTR) <> 0 Then TimeSTR = Left$(TimeSTR, Len(TimeSTR) - 1)
    End If
End Function

Private Sub ProcessCmd(Index As Integer, strC As String)
    Dim comD() As String
    Dim strChat As String
    Dim i As Integer
    Dim j As Integer
    Dim k, x, y, z As Integer
    Dim sc As Integer
    Dim aFound As Boolean
    Dim aPrev As Boolean
    Dim boo1 As Boolean
    Dim comD1() As String
    Dim aStr As String
    Dim gameType As Integer
    
    If strC = "" Then Exit Sub
    'AddLog "RAWDATA=" & strC
    
    comD = Split(strC, "=")
    Select Case comD(0)
        Case "name"
            ply(Index).Name = comD(1)     'name of player
            ply(Index).Con = 1
            ply(Index).numFound = 0
            ply(Index).numOrig = 0
            ply(Index).numAF = 0
            ply(Index).numInv = 0
            ply(Index).Score = 0
            ply(Index).Speed = 0
            
            AddLog ply(Index).Name & " has entered the game."
            WSend Index, "con"
            WSendAll Index, "addP=" & comD(1)
            WSendAll Index, "echo=**" & ply(Index).Name & " has entered the game."
            For i = 1 To numP
                If sckMain(i).State = sckConnected And i <> Index Then
                    WSend Index, "addP=" & ply(i).Name
                    WSend Index, "echo=**" & ply(i).Name & " has entered the game."
                End If
            Next i
            
            WSend Index, "state=" & gameState
            
            'send gamestate
            Select Case gameState
                Case 1 'nothing happening
                
                Case 2 'game running
                    WSend Index, "game=" & strGame & "=" & minLen & "=" & rNum
                Case 3 'stats
                    WSend Index, "interval"
            End Select
        Case "chat"
            strChat = comD(1)
            boo1 = False
            If Len(strChat) > 1 Then
                If Left$(strChat, 1) = "/" Then 'user command
                    boo1 = True
                    comD1 = Split(comD(1), " ")
                    Select Case LCase(comD1(0))
                        Case "/cstart"
                            If gameStart = False Then
                                
                                TTS = 15000 'default values
                                TTR = 180000
                                TTI = 60000
                                minLen = 3
                                
                                If UBound(comD1) >= 1 Then
                                        If Val(comD1(1)) >= 3 Then minLen = Val(comD1(1))
                                End If
                                If UBound(comD1) >= 2 Then
                                        If Val(comD1(2)) >= 30 Then TTR = Val(comD1(2)) * 1000
                                End If
                                If UBound(comD1) >= 3 Then
                                        If Val(comD1(3)) >= 30 Then TTI = Val(comD1(3)) * 1000
                                End If
                                If UBound(comD1) >= 4 Then
                                        If Val(comD1(4)) >= 15 Then TTS = Val(comD1(4)) * 1000
                                End If

                                
                                'TTS = 20000  'def=15000 'time to start = 10 seconds
                                'TTR = 180000 'def=180000 ' round time = 3 min
                                'TTI = 60000 'def=60000'interval time
                                
                                InitStartG
                            End If
                        Case "/start"
                            If gameStart = False Then
                                
                                TTS = 15000 'default values
                                TTR = 180000
                                TTI = 60000
                                minLen = 3
                                gameType = 1 'normal
                                
                                Select Case UBound(comD1)
                                    Case 1
                                        gameType = Val(comD1(1))
                                End Select

                                Select Case gameType
                                    Case 1  'default
                                        WSendAll 0, "echo=*default game settings"
                                        TTS = 15000 'time to start
                                        TTR = 180000 'round time
                                        TTI = 60000 'interval time
                                    Case 2 'long
                                        WSendAll 0, "echo=*long game settings"
                                        TTS = 20000
                                        TTR = 180000
                                        TTI = 120000
                                    Case 3 'competition
                                        WSendAll 0, "echo=*competition game settings"
                                        TTS = 30000
                                        TTR = 180000
                                        TTI = 180000
                                    Case 4 'short
                                        WSendAll 0, "echo=*short game settings"
                                        TTS = 15000
                                        TTR = 180000
                                        TTI = 30000
                                    Case 5 'lightning
                                        WSendAll 0, "echo=*lightning game settings"
                                        TTS = 15000
                                        TTR = 60000
                                        TTI = 45000
                                    Case 6 'lightning competition
                                        WSendAll 0, "echo=*lightning competition game settings"
                                        TTS = 30000
                                        TTR = 60000
                                        TTI = 90000
                                    Case 7 'debug
                                        TTS = 2000
                                        TTR = 30000
                                        TTI = 30000
                                End Select
                                
                                
                                InitStartG
                            End If
                        Case "/help"
                                Select Case UBound(comD1)
                                    Case 0
                                            WSend Index, "echo=*Help"
                                            WSend Index, "echo=*syntax: /help <a>"
                                            WSend Index, "echo=*" & vbTab & "<a> - 'command name'"
                                            WSend Index, "echo=*" & vbTab & "'start': Start game"
                                            WSend Index, "echo=*" & vbTab & "'cstart': Custom Start game"
                                            WSend Index, "echo=*" & vbTab & "'me': User action"
                                            
                                    Case 1
                                        Select Case comD1(1)
                                            Case "start"
                                                WSend Index, "echo=*Start Game"
                                                WSend Index, "echo=*syntax: /start <a>"
                                                WSend Index, "echo=*" & vbTab & "<a> - Game Setting Type (default:1)"
                                                WSend Index, "echo=*" & vbTab & "Type 1: default [start!%15s round!%3min int!%1min]"
                                                WSend Index, "echo=*" & vbTab & "Type 2: long [start!%20s round!%3min int!%2min]"
                                                WSend Index, "echo=*" & vbTab & "Type 3: competition [start!%30s round!%3min int!%3min]"
                                                WSend Index, "echo=*" & vbTab & "Type 4: short [start!%15s round!%3min int!%30s]"
                                                WSend Index, "echo=*" & vbTab & "Type 5: lightning [start!%15s round!%1min int!%45s]"
                                                WSend Index, "echo=*" & vbTab & "Type 6: lightning competition [start!%30s round!%1min int!%1min 30s]"
                                            Case "cstart"
                                                WSend Index, "echo=*Custom Start Game"
                                                WSend Index, "echo=*syntax: /cstart <a> <b> <c> <d>"
                                                WSend Index, "echo=*" & vbTab & "<a> - Minimum length of words (default:3)"
                                                WSend Index, "echo=*" & vbTab & "<b> - Round time in seconds (default:180)"
                                                WSend Index, "echo=*" & vbTab & "<c> - Interval time in seconds (default:60)"
                                                WSend Index, "echo=*" & vbTab & "<d> - Start time in seconds (default:15)"
                                            Case "me"
                                                WSend Index, "echo=*User action"
                                                WSend Index, "echo=*syntax: /me <a>"
                                                WSend Index, "echo=*" & vbTab & "<a> - Action"
                                        End Select
                                        gameType = Val(comD1(1))
                                End Select
                        Case "/me"
                            If Len(comD(1)) > 4 Then
                                WSendAll 0, "echo=^^" & ply(Index).Name & " " & Right(comD(1), Len(comD(1)) - 4)
                            End If
                    End Select
                End If
            End If
            If boo1 = False Then
                WSendAll 0, "echo=" & ply(Index).Name & " says: " & strChat
            End If
        Case "word"
            If gameStart Then
                'check if entered b4
                aPrev = False
                For i = 1 To ply(Index).numFound
                    If ply(Index).Found(i) = comD(1) Then
                        aPrev = True
                        Exit For
                    End If
                Next i
                If Not aPrev Then
                    For i = 1 To ply(Index).numInv
                        If ply(Index).Inv(i) = comD(1) Then
                            aPrev = True
                            Exit For
                        End If
                    Next i
                End If
                
            If Not aPrev Then ply(Index).Speed = ply(Index).Speed + Len(comD(1))
            
            If ValidateW(comD(1)) Then
                
                If Not aPrev Then
                    
                    sc = ScoreWord(comD(1))
                    
                    aFound = False
                    
                    For i = 1 To numP
                        If sckMain(i).State = sckConnected And i <> Index Then
                            For j = 1 To ply(i).numFound
                                If ply(i).Found(j) = comD(1) Then
                                    aFound = True
                                    boo1 = False
                                    For k = 1 To ply(i).numOrig
                                        If ply(i).Orig(k) = comD(1) Then
                                            x = k
                                            boo1 = True
                                            Exit For
                                        End If
                                    Next k
                                    If boo1 Then
                                        ply(i).numOrig = ply(i).numOrig - 1
                                        ply(i).Score = ply(i).Score - sc
                                        For k = x To ply(i).numOrig
                                            ply(i).Orig(k) = ply(i).Orig(k + 1)
                                        Next k
                                        ply(i).numAF = ply(i).numAF + 1
                                        ply(i).AF(ply(i).numAF) = comD(1)
                                        WSend i, "mAF=" & sc & "=" & comD(1)
                                    End If
                                End If
                                If aFound Then Exit For
                            Next j
                        End If
                    Next i
                    
                    If aFound Then  'already found
                        WSend Index, "wAF=" & sc & "=" & comD(1)
                        ply(Index).numFound = ply(Index).numFound + 1
                        ply(Index).Found(ply(Index).numFound) = comD(1)
                        ply(Index).numAF = ply(Index).numAF + 1
                        ply(Index).AF(ply(Index).numAF) = comD(1)
                    Else    'original word
                        AddLog ply(Index).Name & " found " & comD(1) & "."
                        WSend Index, "wOK=" & sc & "=" & comD(1)
                        ply(Index).Score = ply(Index).Score + sc
                        ply(Index).numFound = ply(Index).numFound + 1
                        ply(Index).Found(ply(Index).numFound) = comD(1)
                        ply(Index).numOrig = ply(Index).numOrig + 1
                        ply(Index).Orig(ply(Index).numOrig) = comD(1)
                    End If
                    UpdateScores
                End If
            Else
                If Not aPrev Then
                    ply(Index).numInv = ply(Index).numInv + 1
                    ply(Index).Inv(ply(Index).numInv) = comD(1)
                    WSend Index, "wInv=" & comD(1)
                End If
            End If
        End If
    End Select
End Sub

Private Sub InitStartG()
                                startT = GetTickCount
                                gameState = 1   'waiting for next round to start
                                rNum = 0
                                genState = 0
                                tmrR.Enabled = True
                                
                                WSendAll 0, "echo=**"
                                WSendAll 0, "echo=**" & "New Game: in " & TimeSTR(TTS)
                                WSendAll 0, "echo=**" & "Settings:"
                                WSendAll 0, "echo=**" & vbTab & "Minimum Length - " & minLen
                                WSendAll 0, "echo=**" & vbTab & "Round Time - " & TimeSTR(TTR)
                                WSendAll 0, "echo=**" & vbTab & "Interval Time - " & TimeSTR(TTI)
                                WSendAll 0, "echo=**"
                                AddLog "New Game: in " & TimeSTR(TTS)
                                AddLog "Settings:"
                                AddLog vbTab & "Minimum Length -" & minLen
                                AddLog vbTab & "Round Time - " & TimeSTR(TTR)
                                AddLog vbTab & "Interval Time - " & TimeSTR(TTI)
                                
                                'gen game baord
                                GenG
End Sub

Private Sub EndG()
On Error Resume Next
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim FoundStr As String
    Dim nSpace As Integer
    
    gameStart = False
    gameState = 3
    startT = GetTickCount
    WSendAll 0, "endG"
    WSendAll 0, "echo=**End of round."
    doRank
    
    k = -1
    If nRank > 0 Then k = lRank(1)
    
    If k <> -1 Then
        WSendAll 0, "echo=**" & ply(k).Name & " won this round with " & ply(k).Score & " points."
    End If
    
    'give awards
    GiveAwards
    For i = 1 To numP
        If CheckOn(i) Then
            For j = 1 To numMedals
                If ply(i).bAward(j) Then
                    WSendAll 0, "award=" & ply(i).Name & "=" & j
                End If
            Next j
        End If
    Next i
    
    For i = 1 To numP
        If CheckOn(i) Then
            WSendAll 0, "prize=" & ply(i).Name & "=" & ply(i).Prize
        End If
    Next i
    
    'Send stats
    For i = 1 To numP
        If sckMain(i).State = sckConnected Then
            For j = 1 To ply(i).numOrig
                WSendAll 0, "eOK=" & ply(i).Name & "=" & ScoreWord(ply(i).Orig(j)) & "=" & ply(i).Orig(j)
            Next j
            For j = 1 To ply(i).numAF
                WSendAll 0, "eAF=" & ply(i).Name & "=" & ScoreWord(ply(i).AF(j)) & "=" & ply(i).AF(j)
            Next j
            
        End If
    Next i
    
    k = 0
    
    'stats
    ''''''
    'Player Name
    'Position
    'Points
    'Words Found
    'Typing Speed
    'num awards
    
    For i = 1 To numP
        If sckMain(i).State = sckConnected Then
            k = k + ply(i).numFound
            ply(i).Speed = ply(i).Speed * 60000 / TTR
            WSendAll 0, "stats=" & ply(i).Name & "=" & ply(i).pos & "=" & ply(i).Score & "=" & ply(i).numFound & "=" & ply(i).Speed & "=" & ply(i).numAwards
        End If
    Next i
        'stats1
        '''''''
        'winner
        'winner points
        'total number of words in board
        'number of words found
    
    k = 0
    For i = 1 To numTop
        If topFound(i) Then k = k + 1
    Next i
        
    WSendAll 0, "stats1=" & ply(lRank(1)).Name & "=" & ply(lRank(1)).Score & "=" & numTop & "=" & k

    'top50
    ''''''
    For i = 1 To 50
        If Len(topWords(i)) > 0 Then
            FoundStr = ""
            For j = 1 To numP
                If sckMain(j).State = sckConnected Then
                    For k = 1 To ply(j).numFound
                        If ply(j).Found(k) = topWords(i) Then
                        FoundStr = FoundStr & ply(j).Name & ", "
                            Exit For
                        End If
                    Next k
                End If
            Next j
            If Len(FoundStr) > 2 Then
                If Right$(FoundStr, 2) = ", " Then FoundStr = Left$(FoundStr, Len(FoundStr) - 2)
            End If
            WSendAll 0, "top=" & i & "=" & ScoreWord(topWords(i)) & "=" & topWords(i) & "=" & FoundStr
        End If
    Next i
    WSendAll 0, "topdone"
    
    'generate next round and send to solver
    genState = 0
    GenG
End Sub

Private Function ScoreWord(wor As String) As Integer
    Dim i As Integer
    
    ScoreWord = 11
    i = Len(wor)
                    
    Select Case i   'scoring system
                        Case 1
                            ScoreWord = 1
                        Case 2
                            ScoreWord = 1
                        Case 3
                            ScoreWord = 1
                        Case 4
                            ScoreWord = 1
                        Case 5
                            ScoreWord = 2
                        Case 6
                            ScoreWord = 3
                        Case 7
                            ScoreWord = 5
                        Case 8
                            ScoreWord = 11
    End Select
End Function

Private Sub UpdateScores()
    For i = 1 To numP
        If sckMain(i).State = sckConnected Then
            WSendAll 0, "score=" & ply(i).Name & "=" & ply(i).Score
        End If
    Next i
End Sub

Private Function ValidateW(strWord As String) As Boolean
    Dim i As Long
    Dim tw As String
    ValidateW = False
    DoEvents
    tw = LCase(strWord)
    If genState = 2 Then 'solved already
        For i = 1 To numTop
            If strWord = topWords(i) Then
                ValidateW = True
                topFound(i) = True
                Exit For
            End If
            DoEvents
        Next i
    Else
        For i = 0 To UBound(Dict)
            If Len(Dict(i)) >= minLen Then
                If tw = Dict(i) Then
                    ValidateW = True
                    Exit For
                End If
            End If
            DoEvents
        Next i
    End If
End Function

Private Sub sckMain_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    sckMain(Index).Close
    If ply(Index).Con = 1 Then
        ply(Index).Con = 0
        AddLog ply(Index).Name & " has left the game."
        WSendAll 0, "leaveP=" & ply(Index).Name
        WSendAll 0, "echo=**" & ply(Index).Name & " has left the game."
    End If
End Sub

Private Sub sckR_ConnectionRequest(ByVal requestID As Long)
    numP = numP + 1
    Load sckMain(numP)
    With sckMain(numP)
        .Accept requestID
        WSend numP, "name?"
    End With
    
End Sub

Private Sub SSend(tos As String)
    If sckBSolve.State = sckConnected Then
        sckBSolve.SendData tos & vbNullChar
    End If
End Sub

Private Sub WSend(Index As Integer, tos As String)
    If sckMain(Index).State = sckConnected Then
        sckMain(Index).SendData tos & vbNullChar
    End If
End Sub

Private Sub WSendAll(Index As Integer, tos As String) ' send to all except Index
    Dim i As Integer
    For i = 1 To numP
        If i <> Index Then
            If sckMain(i).State = sckConnected Then
                sckMain(i).SendData tos & vbNullChar
            End If
        End If
    Next i
End Sub

Private Sub GenG() 'generate game
    Dim diceNum(1 To 16) As Integer
    Dim i As Integer
    Dim j As Integer
    
    genState = 1 'generate,send
    rNum = rNum + 1
    strGame = ""
    
    'generate game
    If Len(txtNextRound.Text) <> 16 Then
        For i = 1 To 16
            Do
                j = ((Rnd * 1000) Mod 16) + 1
                ok = True
                For k = 1 To i - 1
                   If diceNum(k) = j Then ok = False
                   Exit For
                Next k
                If ok Then Exit Do
            Loop
            diceNum(i) = j
        Next i
        
        For i = 1 To 16
        j = ((Rnd * 1000) Mod 6) + 1
        strGame = strGame & Mid(BDice(diceNum(i)), j, 1)
        Next i
    Else
        strGame = UCase(txtNextRound.Text)
        AddLog "Custom game round loaded."
        txtNextRound.Text = ""
    End If
    txtNextRound.Text = ""

    

    SSend "solve=" & minLen & "=" & strGame
    AddLog "Sent round " & rNum & " board to boggle solver."
    
End Sub

Private Sub StartG() 'start game!
    Dim i, j, k As Integer
    Dim ok As Boolean
        
    gameStart = True
    gameState = 2
    startT = GetTickCount
    
    If genState = 0 Then GenG
    
    'reset scores,words
    For i = 1 To numP
        ply(i).numFound = 0
        ply(i).numOrig = 0
        ply(i).numAF = 0
        ply(i).numInv = 0
        ply(i).Score = 0
        ply(i).numAwards = 0
        ply(i).Speed = 0
        ply(i).Prize = 0
        For j = 1 To numMedals
            ply(i).bAward(j) = False
        Next j
    Next i
    
    WSendAll 0, "game=" & strGame & "=" & minLen & "=" & rNum
    DoEvents
    tmrR_Timer
    
    UpdateScores
    WSendAll 0, "echo=**Round " & rNum & " started."
End Sub

Private Sub tmrLoad_Timer()
On Error GoTo errdo
    Dim ff As Integer
    
    tmrLoad.Enabled = False
    
    ff = FreeFile
    Open "Word.lst" For Input As #ff
        Dict = Split(Input(LOF(ff), 1), vbCrLf)
    Close #ff
    
    AddLog UBound(Dict) & " words loaded into Boggle Dictionary."
   Exit Sub
errdo:
    MsgBox Err.Description
    MsgBox "Missing or corrupt Dict.dat"
    End
End Sub

Private Sub tmrR_Timer()
    Dim curT As Long
    Dim dif As Long
    Dim timeLeft As Long
    
    curT = GetTickCount
    
    Select Case gameState
        Case 0  'nothing happening
            tmrR.Enabled = False
        Case 1  'waiting for next round
            
            dif = curT - startT
            If dif >= TTS Then
                StartG
            Else
                timeLeft = TTS - dif
                WSendAll 0, "TStart=" & timeLeft
            End If
        Case 2  'in play
            dif = curT - startT
            If dif >= TTR Then
                EndG
            Else
                timeLeft = TTR - dif
                WSendAll 0, "TLeft=" & timeLeft
            End If
        Case 3  'interval timing
            
            dif = curT - startT
            If dif >= TTI Then
                StartG
            Else
                timeLeft = TTI - dif
                WSendAll 0, "TInt=" & timeLeft
            End If
        Case Else
            tmrR.Enabled = False
        
    End Select
End Sub

Private Sub doRank()
    Dim i, j, k, x As Integer
    Dim sto, doSwap As Boolean
    
    nRank = 0
    For i = 1 To numP
        If sckMain(i).State = sckConnected Then
            nRank = nRank + 1
            lRank(nRank) = i
        End If
    Next i
    
    'bubblesort
    sto = True
    For i = 1 To nRank
        For j = 1 To nRank - 1
            If ply(lRank(j)).Score < ply(lRank(j + 1)).Score Then
                sto = False
                k = lRank(j)
                lRank(j) = lRank(j + 1)
                lRank(j + 1) = k
            ElseIf ply(lRank(j)).Score = ply(lRank(j + 1)).Score Then
                doSwap = False
                For x = 1 To Len(ply(lRank(j)).Name)
                    If Asc(Mid(ply(lRank(j)).Name, x, 1)) > Asc(Mid(ply(lRank(j)).Name, x, 1)) Then
                        doSwap = True
                        Exit For
                    End If
                    If Asc(Mid(ply(lRank(j)).Name, x, 1)) < Asc(Mid(ply(lRank(j)).Name, x, 1)) Then Exit For
                Next x
                If doSwap Then
                    sto = False
                    k = lRank(j)
                    lRank(j) = lRank(j + 1)
                    lRank(j + 1) = k
                End If
            End If
        Next j
        If sto Then Exit For
    Next i
    
    For i = 1 To nRank
        ply(lRank(i)).pos = i
    Next i
End Sub

Private Function CheckOn(pIndex As Integer) As Boolean
    CheckOn = False
    If sckMain(pIndex).State = sckConnected Then CheckOn = True
End Function

Private Sub GiveAwards()
On Error Resume Next
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim z As Integer
    Dim fl As Double
    
    'award 2:champion
    ply(lRank(1)).bAward(2) = True
    
    'award 3:high flyer
    For i = 1 To numP
        If CheckOn(i) Then
            If ply(i).Score >= 50 Then ply(i).bAward(3) = True
            If ply(i).numFound = numTop Then ply(i).bAward(3) = True
        End If
    Next i
    
    'award 4:detective
    For i = 1 To numP
        If CheckOn(i) Then
            z = 0
            For j = 1 To 3
                For k = 1 To ply(i).numFound
                    If ply(i).Found(k) = topWords(j) Then
                        z = z + 1
                        Exit For
                    End If
                Next k
            Next j
            If z = 3 Then ply(i).bAward(4) = True
        End If
    Next i
    
    'award 5:scavenger
    k = 0
    For i = 1 To numP
        If CheckOn(i) Then
            If ply(i).numFound > k Then k = ply(i).numFound
        End If
    Next i
    For i = 1 To numP
        If CheckOn(i) Then
            If ply(i).numFound = k And k <> 0 Then ply(i).bAward(5) = True
        End If
    Next i
    
    'award 6:treasure hunter
    For i = 1 To numP
        If CheckOn(i) Then
            ply(i).tmpInt = 0
            For j = 1 To 5
                For k = 1 To ply(i).numFound
                    If ply(i).Found(k) = topWords(j) Then
                        ply(i).tmpInt = ply(i).tmpInt + 1
                        Exit For
                    End If
                Next k
            Next j
        End If
    Next i
    k = 0
    For i = 1 To numP
        If CheckOn(i) Then
            If ply(i).tmpInt > k Then k = ply(i).tmpInt
        End If
    Next i
    For i = 1 To numP
        If CheckOn(i) Then
            If ply(i).tmpInt = k And k <> 0 Then ply(i).bAward(6) = True
        End If
    Next i
    
    'award 7:top 3
    For i = 1 To 3
        If lRank(i) <> 0 Then
            If CheckOn(lRank(i)) Then ply(lRank(i)).bAward(7) = True
        End If
    Next i
    
    'award 8:caterpillar
    k = 0
    For i = 1 To numP
        If CheckOn(i) Then
            ply(i).tmpInt = 0
            For j = 1 To ply(i).numFound
                If Len(ply(i).Found(j)) > ply(i).tmpInt Then ply(i).tmpInt = Len(ply(i).Found(j))
            Next j
        End If
    Next i
    k = 0
    For i = 1 To numP
        If CheckOn(i) Then
            If ply(i).tmpInt > k Then k = ply(i).tmpInt
        End If
    Next i
    For i = 1 To numP
        If CheckOn(i) Then
            If ply(i).tmpInt = k And k >= 5 Then ply(i).bAward(8) = True
        End If
    Next i
    
    'award 9:fastest typer
    k = 0
    For i = 1 To numP
        If CheckOn(i) Then
            If ply(i).Speed > k Then k = ply(i).Speed
        End If
    Next i
    For i = 1 To numP
        If CheckOn(i) Then
            If ply(i).Speed = k And k <> 0 Then ply(i).bAward(9) = True
        End If
    Next i
    
    'award 10:full of crap
    For i = 1 To numP
        If CheckOn(i) Then
            If (ply(i).numFound + ply(i).numInv) > 0 Then
                If (ply(i).numInv / (ply(i).numFound + ply(i).numInv)) > 0.7 Then ply(i).bAward(10) = True
            End If
        End If
    Next i
    
    'award 11:small change
    For i = 1 To numP
        If CheckOn(i) Then
            k = 0
            For j = 1 To ply(i).numFound
                If Len(ply(i).Found(j)) <= 4 Then k = k + 1
            Next j
            If ply(i).numFound > 0 Then
                If (k / ply(i).numFound) > 0.9 Then ply(i).bAward(11) = True
            End If
        End If
    Next i
    
    'award 12:perfectionist
    For i = 1 To numP
        If CheckOn(i) Then
            If ply(i).numInv = 0 Then
                If ply(i).numFound > 0 Then ply(i).bAward(12) = True
            End If
        End If
    Next i
    
    'award 13:loser
    For i = 1 To numP
        If CheckOn(i) Then
            If ply(i).numFound > 0 Then
                If (ply(i).numAF / ply(i).numFound) > 0.9 Then ply(i).bAward(13) = True
            End If
        End If
    Next i
        
    'number of medals
    For i = 1 To numP
        If CheckOn(i) Then
            ply(i).numAwards = 0
            For j = 2 To 8
                If ply(i).bAward(j) = True Then
                    ply(i).numAwards = ply(i).numAwards + 1
                End If
            Next j
        End If
    Next i
        
    'prizes
    fl = 0
    For i = 1 To numP
        If CheckOn(i) Then
            If numTop >= 1 Then
                fl = (ply(i).numFound * 100 / numTop)
                If fl >= 80 Then
                    ply(i).Prize = 1
                ElseIf fl >= 70 Then
                    ply(i).Prize = 2
                ElseIf fl >= 50 Then
                    ply(i).Prize = 3
                Else
                    ply(i).Prize = 0
                End If
            End If
            'ply(i).Prize = ((Rnd * 1000) Mod 3) + 1
        End If
    Next i
        
    'award 1:solitaire
    For i = 1 To numP
        If CheckOn(i) Then
            If ply(i).numAwards = 7 And ply(i).Prize = 1 Then ply(i).bAward(1) = True
        End If
    Next i
    
End Sub
