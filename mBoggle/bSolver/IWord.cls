VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mintIndex               As Integer
Private mlngLength              As Long
Private mstrText                As String
Private mcolCharacters          As Collection

Public Property Get Length() As Long
    Length = mlngLength
End Property

Friend Property Let Index(ByVal intIndex As Integer)
    mintIndex = intIndex
End Property

Public Property Get Index() As Integer
    Index = mintIndex
End Property

Friend Property Let Text(ByVal strText As String)
    mstrText = strText
    mlngLength = Len(strText)
End Property

Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
    Text = mstrText
End Property

Public Property Get IsComplete() As Boolean
    IsComplete = Len(mstrText) = mcolCharacters.Count
End Property

Public Property Get FoundCharacters() As Integer
    FoundCharacters = mcolCharacters.Count
End Property

Friend Sub AddLetter(ByVal strLetter As String, ByVal intColumn As Integer, ByVal intRow As Integer)
    Dim objChar         As IChar
    Set objChar = New IChar
    If Not (Letter(intColumn, intRow) Is Nothing) Then Exit Sub
    mintIndex = mcolCharacters.Count + 1
    With objChar
        .Char = strLetter
        .Index = mintIndex
        .GridX = intColumn
        .GridY = intRow
    End With
    mcolCharacters.Add objChar, intColumn & "," & intRow
    Set objChar = Nothing
End Sub

Public Sub RemoveLetter(Optional ByVal intColumn As Integer, Optional ByVal intRow As Integer, Optional ByVal intIndex As Integer)
    On Error Resume Next
    If intIndex > 0 Then
        mcolCharacters.Remove intIndex
    Else
        mcolCharacters.Remove intColumn & "," & intRow
    End If
End Sub

Public Function IsLetterUsed(ByRef objLetter As IChar) As Boolean
    Dim objChar             As IChar
    Set objChar = Letter(objLetter.GridX, objLetter.GridY)
    IsLetterUsed = Not objChar Is Nothing
    Set objChar = Nothing
End Function

Public Property Get Letter(Optional ByVal intColumn As Integer, Optional ByVal intRow As Integer, Optional ByVal intIndex As Integer) As IChar
    On Error Resume Next
    If intIndex > 0 Then
        Set Letter = mcolCharacters(intIndex)
    Else
        Set Letter = mcolCharacters(intColumn & "," & intRow)
    End If
End Property

Public Property Get NewEnum()
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "440"
    Set NewEnum = mcolCharacters.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mcolCharacters = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolCharacters = Nothing
End Sub
