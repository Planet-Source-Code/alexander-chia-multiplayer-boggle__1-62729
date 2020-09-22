VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public Enum SearchDirection
    sdUp = 1&
    sdUpLeft = 2&
    sdLeft = 3&
    sdDownLeft = 4&
    sdDown = 5&
    sdDownRight = 6&
    sdRight = 7&
    sdUpRight = 8&
End Enum

Private mblnGridCreated         As Boolean
Private mblnGridPopulated       As Boolean
Private mblnGetWordsOnComplete  As Boolean
Private mintLettersAcross       As Integer
Private mintLettersDown         As Integer
Private mintMinWordSize         As Integer
Private mintIndex               As Integer
Private mstrDictionary()        As String
Private mcolLettersInGrid       As Collection
Private mcolWordsInGrid         As Collection
Private mcolWordsFound          As Collection

Public Function CreateGrid(ByVal intLettersAccross As Integer, ByVal intLettersDown As Integer, Optional ByVal intMinWordSize As Integer = 4, Optional ByVal blnGetWordsOnComplete As Boolean) As Boolean
    mblnGridCreated = intLettersAccross > 0 And intLettersDown > 0 And intMinWordSize > 0
    If Not mblnGridCreated Then Exit Function
    Set mcolLettersInGrid = New Collection
    Set mcolWordsFound = New Collection
    mintLettersAcross = intLettersAccross
    mintLettersDown = intLettersDown
    mintIndex = 0
    mintMinWordSize = intMinWordSize
    mblnGetWordsOnComplete = blnGetWordsOnComplete
    mblnGridPopulated = False
    CreateGrid = True
End Function

Public Function AddLetter(ByVal strLetter As String, ByVal intColumn As Integer, ByVal intRow As Integer) As IChar
    Dim objChar         As IChar
    Set objChar = New IChar
    If intColumn < 0 Then Exit Function
    If intRow < 0 Then Exit Function
    If intColumn > mintLettersAcross Then Exit Function
    If intRow > mintLettersDown Then Exit Function
    If Not (Letter(intColumn, intRow) Is Nothing) Then Exit Function
    If mblnGridPopulated Then Exit Function
    mintIndex = mintIndex + 1
    mblnGridPopulated = mintIndex = mintLettersAcross * mintLettersDown
    With objChar
        .Char = strLetter
        .Index = mintIndex
        .GridX = intColumn
        .GridY = intRow
    End With
    mcolLettersInGrid.Add objChar, intColumn & "," & intRow
    Set AddLetter = objChar
    If mblnGridPopulated And mblnGetWordsOnComplete Then GetWords
    Set objChar = Nothing
End Function

Public Property Get Letter(Optional ByVal intColumn As Integer, Optional ByVal intRow As Integer, Optional ByVal intIndex As Integer) As IChar
    On Error Resume Next
    If intIndex > 0 Then
        Set Letter = mcolLettersInGrid(intIndex)
    Else
        Set Letter = mcolLettersInGrid(intColumn & "," & intRow)
    End If
End Property

Private Function NextLetter(ByVal intCurrentColumn As Integer, ByVal intCurrentRow As Integer, ByVal sdDirection As SearchDirection) As IChar
    Dim intY            As Integer
    Dim intX            As Integer
    Select Case sdDirection
        Case sdUp:          intX = intCurrentColumn:        intY = intCurrentRow + 1
        Case sdUpLeft:      intX = intCurrentColumn - 1:    intY = intCurrentRow + 1
        Case sdLeft:        intX = intCurrentColumn - 1:    intY = intCurrentRow
        Case sdDownLeft:    intX = intCurrentColumn - 1:    intY = intCurrentRow - 1
        Case sdDown:        intX = intCurrentColumn:        intY = intCurrentRow - 1
        Case sdDownRight:   intX = intCurrentColumn + 1:    intY = intCurrentRow - 1
        Case sdRight:       intX = intCurrentColumn + 1:    intY = intCurrentRow
        Case sdUpRight:     intX = intCurrentColumn + 1:    intY = intCurrentRow + 1
    End Select
    If intX < 0 Or intX > mintLettersAcross Then Exit Function
    If intY < 0 Or intY > mintLettersDown Then Exit Function
    Set NextLetter = Letter(intX, intY)
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "440"
    Set NewEnum = mcolWordsFound.[_NewEnum]
End Property

Public Sub GetWords()
    On Error Resume Next
    Dim lngPtr          As Long
    Dim objWord         As iWord
    Static slngIndex    As Long
    Set mcolWordsInGrid = New Collection
    For lngPtr = 0 To UBound(mstrDictionary)
        If Len(mstrDictionary(lngPtr)) >= mintMinWordSize Then
            If IsCompletable(mstrDictionary(lngPtr)) Then
                slngIndex = slngIndex + 1
                Set objWord = New iWord
                With objWord
                    .Index = slngIndex
                    .Text = mstrDictionary(lngPtr)
                End With
                mcolWordsInGrid.Add objWord, objWord.Text
                If Err.Number Then slngIndex = slngIndex - 1
            End If
        End If
        DoEvents
    Next
    For lngPtr = 1 To mcolWordsInGrid.Count
        Set objWord = GetWord(mcolWordsInGrid(lngPtr))
        If Not objWord Is Nothing Then mcolWordsFound.Add objWord
    Next
    Set objWord = Nothing
End Sub

Private Function IsCompletable(ByVal strWord As String) As Boolean
    Dim lngLen          As Long
    Dim intCount        As Integer
    Dim strChar         As String * 1
    strWord = Trim$(strWord)
    lngLen = Len(strWord)
    If lngLen = 0 Then Exit Function
    While lngLen
        strChar = Left$(strWord, 1)
        strWord = Replace(strWord, strChar, vbNullString)
        intCount = lngLen - Len(strWord)
        If GridLetterCount(strChar) < intCount Then Exit Function
        lngLen = Len(strWord)
    Wend
    IsCompletable = True
End Function

Private Function GridLetterCount(ByVal strLetter As String) As Integer
    Dim objChar         As IChar
    Dim intCount        As Integer
    For Each objChar In mcolLettersInGrid
        If objChar.Char = strLetter Then intCount = intCount + 1
    Next
    GridLetterCount = intCount
End Function

Public Function GetWord(Optional ByVal strWord As String) As iWord
    Dim objChar         As IChar
    Dim objLastChar     As IChar
    Dim lngPtr          As Long
    Dim strChar         As String * 1
    Dim sdDirection     As SearchDirection
    Dim objFinalWord    As iWord
    Static sobjWord     As iWord
    
    If Len(strWord) Then
        Set sobjWord = New iWord
        sobjWord.Text = strWord
        strChar = Left$(sobjWord.Text, 1)
        For Each objChar In mcolLettersInGrid
            If objChar.Char = strChar Then
                With objChar
                    sobjWord.AddLetter .Char, .GridX, .GridY
                End With
                GetWord
            End If
            If sobjWord.IsComplete Then
                Set GetWord = sobjWord
                Set sobjWord = Nothing
                Exit For
            Else
                sobjWord.RemoveLetter objChar.GridX, objChar.GridY
            End If
        Next
    Else
        Set objLastChar = sobjWord.Letter(intIndex:=sobjWord.FoundCharacters)
        strChar = Mid$(sobjWord.Text, sobjWord.FoundCharacters + 1, 1)
        For sdDirection = sdUp To sdUpRight
            While objChar Is Nothing
                Set objChar = NextLetter(objLastChar.GridX, objLastChar.GridY, sdDirection)
                If sdDirection = sdUpRight Then
                    If objChar Is Nothing Then
                        sobjWord.RemoveLetter objLastChar.GridX, objLastChar.GridY
                        Exit Function
                    End If
                Else
                     If objChar Is Nothing Then sdDirection = sdDirection + 1
                End If
            Wend
            If objChar.Char = strChar And Not sobjWord.IsLetterUsed(objChar) Then
                With objChar
                    sobjWord.AddLetter .Char, .GridX, .GridY
                End With
                If Not sobjWord.IsComplete Then GetWord
            End If
            If sobjWord.IsComplete Then Exit Function
            If Not sobjWord.IsLetterUsed(objChar) Then
                sobjWord.RemoveLetter objChar.GridX, objChar.GridY
            End If
            If sdDirection = sdUpRight Then
                sobjWord.RemoveLetter objLastChar.GridX, objLastChar.GridY
            End If
            Set objChar = Nothing
        Next
    End If
End Function

Private Sub Class_Initialize()
    Dim ff As Integer
    Dim i As Long
    
    ff = FreeFile
    Open "Word.lst" For Input As #ff
        mstrDictionary = Split(Input(LOF(ff), 1), vbCrLf)
    Close #ff
    For i = 0 To UBound(mstrDictionary)
        mstrDictionary(i) = Replace(mstrDictionary(i), "QU", "!")
    Next i
End Sub

Private Sub Class_Terminate()
    Set mcolLettersInGrid = Nothing
    Set mcolWordsInGrid = Nothing
    Set mcolWordsFound = Nothing
End Sub
