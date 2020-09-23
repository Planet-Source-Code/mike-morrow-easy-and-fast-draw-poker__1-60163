Attribute VB_Name = "modCards"
Option Explicit

  Public TenCards(1 To 10) As Long  'Predeal 10 cards here for later use.
  Public TenCardPtr As Long  ' Points to the next card to be dealt (1 through 10)
  
  Public gbWhichHand As Boolean
  Public Const Initial_Hand = True
  Public Const Second_Hand = False
  
  Public glBet As Long
  Public glCredit As Long
  Public gbWin As Boolean
  Public giBack As Integer
  Public gbDeck(1 To 52) As Boolean
  Public Const Is_Dealt = True
  Public Const Not_Dealt = False
  Public giNextCard As Integer  ' Next card to be dealt from shuffled deck
  Public giDealtCards(0 To 4) As Integer
  
  Public Const Add_Bonus = True
  Public Const No_Bonus = False
  
  Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

 'These set the displayed values on the bottom of the
 'form and also are used to call Win with.  Any change
 'made here will change all the right places.
  Public Const R_Flush = 250
  Public Const Str_Flush = 50
  Public Const Four_Kind = 25
  Public Const Full_House = 9
  Public Const Flush = 6
  Public Const Straight = 4
  Public Const Three_Kind = 3
  Public Const Two_Pair = 2
  Public Const One_Pair = 1
  
  Public Const Max_Bet = 50
Function Get_a_Card() As String

 'The process here is to Randomize then do hits until we find a card which
 'has not been used.  The used cards will have their boolean value set to
 'Not_Dealt in Shuffle and when used, the boolean will be set to Is_Dealt.
 
 'Normally, you limit the random hits to something like 3 to 5 and then start
 'a linear pass, beginning at the last random selection that failed,
 'wrapping at the 52nd card back to the 1st card until an
 'available card is found.
 
 'Since the most cards which can be dealt here is 10, we will just let it run
 'until an available card is found.  Less code and the overhead is just
 'not worth it.  It rarely hits the same card again in only 10 tries.

 'Randomize

 'Get_a_Card = (Int(Rnd * 52) + 1)

 'While gbDeck(Get_a_Card) = Is_Dealt
 '  Get_a_Card = (Int(Rnd * 52) + 1)
 'Wend
 'gbDeck(Get_a_Card) = Is_Dealt

  Get_a_Card = TenCards(TenCardPtr)
  TenCardPtr = TenCardPtr + 1
  
End Function

Public Sub Shuffle()

  Dim i As Long

 'Yeah, I really only need to negate the 10 from the previous deal but, what if there
 'was no previous deal.
 
  For i = 1 To 52
    gbDeck(i) = Not_Dealt
  Next
  
  Randomize

  For TenCardPtr = 1 To 10  ' Put ten cards picked in the array
  
    i = (Int(Rnd * 52) + 1)
  
    While gbDeck(i) = Is_Dealt
      i = (Int(Rnd * 52) + 1)
    Wend
    
    TenCards(TenCardPtr) = i
    gbDeck(i) = Is_Dealt
  
  Next

  TenCardPtr = 1  ' Deal the first card first
  
End Sub


Public Function Wait(ByVal TimeToWait As Long) 'Time In seconds

  Dim EndTime As Long
  
  EndTime = GetTickCount + TimeToWait * 50
  
  Do Until GetTickCount > EndTime
    DoEvents
  Loop

End Function

Public Function Suit(Index As Integer) As String
'select case to determine suit

  Select Case Index
    Case Is >= 40
      Suit = "Hearts"
    Case Is >= 27
      Suit = "Diamonds"
    Case Is >= 14
      Suit = "Clubs"
    Case Is >= 1
      Suit = "Spades"
  End Select
   
End Function
Public Function Card2Char(Index As Integer) As String

  Dim iModVal As Integer
  
  iModVal = Index Mod 13
  
  Select Case iModVal
    Case 0
      Card2Char = "K"
    Case 1
      Card2Char = "A"
    Case 11
      Card2Char = "J"
    Case 12
      Card2Char = "Q"
    Case Else
      Card2Char = CStr(iModVal)
  End Select
      
End Function


Public Sub Win(sDesc As String, iWinMultiplier As Long, gAddBonus As Boolean)

  gbWin = True 'this is put here in case you win but have no credits left,
               'otherwise the game would end prematurely.
                
  frmPoker.cmdDraw.Enabled = False 'Don't allow the user to hit "Draw" while the
                                   'credits are ringing up.
  Dim i As Long
  
  frmPoker.lblScore.BackColor = &HFFFFC0
  If gAddBonus Then
    frmPoker.lblScore.Caption = "WIN! " & sDesc & " (" & glBet * iWinMultiplier * 1.4 & ")"
  Else
    frmPoker.lblScore.Caption = "WIN! " & sDesc & " (" & glBet * iWinMultiplier & ")"
  End If
  
  For i = 1 To iWinMultiplier
    Beep
    glCredit = glCredit + glBet
    Wait 1
  Next
  
  If gAddBonus Then glCredit = glCredit + glBet * iWinMultiplier * 0.4
  
  frmPoker.lblCredits = CStr(glCredit)
  
  frmPoker.cmdDraw.Enabled = True

End Sub

Public Sub BubbleSort(intArray() As String)
  
  Dim iOuter As Integer
  Dim iInner As Integer
  Dim iLBound As Integer
  Dim iUBound As Integer
  Dim iTemp As Integer
  Dim i As Integer
  
  iLBound = LBound(intArray)
  iUBound = UBound(intArray)
  
 'Which bubbling pass
  For iOuter = iLBound To iUBound - 1
   'Which comparison
    For iInner = iLBound To iUBound - iOuter - 1
     'Compare this item to the next item
      If CInt(intArray(iInner)) > CInt(intArray(iInner + 1)) Then
       'Swap
        iTemp = CInt(intArray(iInner))
        intArray(iInner) = intArray(iInner + 1)
        intArray(iInner + 1) = iTemp
      End If
    Next
  Next

End Sub


