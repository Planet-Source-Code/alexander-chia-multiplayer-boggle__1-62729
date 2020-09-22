Attribute VB_Name = "modMain"
Public Const numMedals = 13

Public BDice(1 To 16) As String
Public buf As String
Public buf1 As String
Public Dict() As String
Public strGame As String
Public Const PW = "sdfa;wpe"
Public Type PType
    Name As String
    Con As Integer
    Index As Integer
    TabIndex As Integer
    Score As Integer
    numFound As Integer
    numOrig As Integer
    numInv As Integer
    numAF As Integer
    pos As Integer
    Prize As Integer
    Found(1 To 1000) As String
    Orig(1 To 1000) As String
    AF(1 To 1000) As String
    Inv(1 To 1000) As String
    Speed As Integer    'typing speed
    numAwards As Integer
    bAward(1 To numMedals) As Boolean
    tmpInt As Integer   'for playing around
End Type
Public Declare Function GetTickCount Lib "kernel32" () As Long

Sub Main()
    Dim diceset As Integer
    
    diceset = 1
    
    If diceset = 1 Then
        BDice(1) = "LRYTTE"
        BDice(2) = "VTHRWE"
        BDice(3) = "EGHWNE"
        BDice(4) = "SEOTIS"
        BDice(5) = "ANAEEG"
        BDice(6) = "IDSYTT"
        BDice(7) = "OATTOW"
        BDice(8) = "MTOICU"
        BDice(9) = "AFPKFS"
        BDice(10) = "XLDERI"
        BDice(11) = "HCPOAS"
        BDice(12) = "ENSIEU"
        BDice(13) = "YLDEVR"
        BDice(14) = "ZNRNHL"
        BDice(15) = "NMIQHU"
        BDice(16) = "OBBAOJ"
    ElseIf diceset = 2 Then
        BDice(1) = "ARELSC"
        BDice(2) = "TABIYL"
        BDice(3) = "EDNSWO"
        BDice(4) = "BIOFXR"
        BDice(5) = "MCDPAE"
        BDice(6) = "IHFYEE"
        BDice(7) = "KTDNUO"
        BDice(8) = "MOQAJB"
        BDice(9) = "ESLUPT"
        BDice(10) = "INVTGE"
        BDice(11) = "ZNDVAE"
        BDice(12) = "UKGELY"
        BDice(13) = "OCATAI"
        BDice(14) = "ULGWIR"
        BDice(15) = "SPHEIN"
        BDice(16) = "MSHARO"
    ElseIf diceset = 3 Then
        BDice(1) = "QQQQQQ"
        BDice(2) = "ESTEST"
        BDice(3) = "ESTEST"
        BDice(4) = "ESTEST"
        BDice(5) = "ESTEST"
        BDice(6) = "ESTEST"
        BDice(7) = "ESTEST"
        BDice(8) = "ESTEST"
        BDice(9) = "ESTEST"
        BDice(10) = "ESTEST"
        BDice(11) = "ESTEST"
        BDice(12) = "ESTEST"
        BDice(13) = "ESTEST"
        BDice(14) = "ESTEST"
        BDice(15) = "ESTEST"
        BDice(16) = "ESTEST"
    
    End If
    
    Randomize
    
    buf = ""
    buf1 = ""
    frmMain.Show
End Sub
