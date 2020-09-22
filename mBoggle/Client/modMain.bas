Attribute VB_Name = "modMain"
Public Const numMedals = 14
Public Const numTopWords = 50
Public Const numDDef = 45501
Public Const numDWor = 103506

Public buf As String
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Type BogWord
    word As String
    VaLue As Integer
End Type
Public Type PType
    name As String  'player name
    Con As Integer  'connection status
    Index As Integer 'player index
    Score As Integer 'score
    numOrig As Integer
    numAF As Integer
    lOrig(1 To 1000) As BogWord
    lAF(1 To 1000) As BogWord
    Pos As Integer
    numFound As Integer
    speed As Integer
    Prize As Integer
    numAwards As Integer
    bAward(1 To numMedals) As Boolean
End Type
Public Type sp_coord
    X As Integer
    Y As Integer
    stopNum As Integer
End Type
Public Type TTopWord
    word As String
    VaLue As Integer
    foundBy As String
End Type
Public Type medalType
    gNum As Integer    'which medal graphic to use
    name As String
    desc As String
End Type
Public Type DWordType
    word As String
    def As Long
End Type
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Medals(1 To numMedals) As medalType
Public Prizes(1 To 3) As medalType
Public DictLoad As Boolean
Public DDef(1 To numDDef) As String
Public DWor(1 To numDWor) As DWordType

Sub Main()
    Randomize
    buf = ""
    DictLoad = False
    
    Medals(1).gNum = 0
    Medals(1).name = "Solitaire"
    Medals(1).desc = "This legendary award is only given to those who are deemed worthy." & vbCrLf & vbCrLf & "Rumour is that collecting all 7 medals and the top star achievement award makes one eligible for this award."
    Medals(2).gNum = 1
    Medals(2).name = "Champion"
    Medals(2).desc = "This gold medal is awarded to the winner of each round."
    Medals(3).gNum = 1
    Medals(3).name = "High Flyer"
    Medals(3).desc = "This highly coveted gold medal is awarded to players who have scored equal or higher than a whopping 50 points." & vbCrLf & vbCrLf & "It is also awarded if a player achieves the feat of finding all possible words in the board."
    Medals(4).gNum = 1
    Medals(4).name = "Detective"
    Medals(4).desc = "This elusive gold medal is awarded to players who found all the top 3 words of the round." & vbCrLf & vbCrLf & "Give yourself a pat on the back if you are awarded this."
    Medals(5).gNum = 2
    Medals(5).name = "Scavenger"
    Medals(5).desc = "This silver medal is awarded to the player(s) who find the highest number of words in each round."
    Medals(6).gNum = 2
    Medals(6).name = "Treasure Hunter"
    Medals(6).desc = "This silver medal is awarded to the player(s) who find the highest number of the top 10 words in each round."
    Medals(7).gNum = 3
    Medals(7).name = "Top 3"
    Medals(7).desc = "This bronze medal is awarded to the top 3 players of each round in terms of points."
    Medals(8).gNum = 3
    Medals(8).name = "Caterpillar"
    Medals(8).desc = "This bronze medal is awarded to the player(s) who find the longest word in each round." & vbCrLf & "(Word length must also be longer than 5)"
    Medals(9).gNum = 4
    Medals(9).name = "Fastest Typer"
    Medals(9).desc = "This award is given to the player who submits the highest number of letters in each round." & vbCrLf & vbCrLf & "The awardee is usually also awarded either the Champion medal or the Full of crap award."
    Medals(10).gNum = 4
    Medals(10).name = "Full of crap"
    Medals(10).desc = "This award is given to players who submit more than 70% of invalid words in a round." & vbCrLf & vbCrLf & "Boo."
    Medals(11).gNum = 4
    Medals(11).name = "Small Change"
    Medals(11).desc = "This booby prize is given to players whose 1 pointer words make up more than 90% of total words found." & vbCrLf & vbCrLf & "Ack, I hope you're not as cheap in real life."
    Medals(12).gNum = 4
    Medals(12).name = "Perfectionist"
    Medals(12).desc = "This prize is given to players who have never submitted a single invalid word!" & vbCrLf & vbCrLf & "Either you're a pro or you're overly cautious.."
    Medals(13).gNum = 5
    Medals(13).name = "Loser"
    Medals(13).desc = "This booby prize is given to players who has submitted 90% or more words which have been already found." & vbCrLf & vbCrLf & "Lacking in originality, eh?"
    Medals(14).gNum = 6
    Medals(14).name = "Star awards"
    Medals(14).desc = "Up to you to find out!"
    
    
    Prizes(1).gNum = 7
    Prizes(1).name = "Gold Star"
    Prizes(1).desc = "This prestigious award is given to players to acknowledge their achievement of finding equal to or more than 80% of total possible words in a board."
    Prizes(2).gNum = 8
    Prizes(2).name = "Silver Star"
    Prizes(2).desc = "This prestigious award is given to players to acknowledge their achievement of finding equal to or more than 70% of total possible words in a board."
    Prizes(3).gNum = 9
    Prizes(3).name = "Bronze Star"
    Prizes(3).desc = "This prestigious award is given to players to acknowledge their achievement of finding equal to or more than 50% of total possible words in a board."
    
    
    frmLoading.Show
End Sub
