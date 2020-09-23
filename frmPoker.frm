VERSION 5.00
Begin VB.Form frmPoker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Draw Poker"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7635
   ControlBox      =   0   'False
   Icon            =   "frmPoker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   7635
   Begin VB.CommandButton cmdMinBet 
      Caption         =   "M&in Bet"
      Height          =   375
      Left            =   6840
      TabIndex        =   4
      Top             =   2220
      Width           =   735
   End
   Begin VB.CommandButton cmdBetDown 
      Caption         =   "B&et -1"
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6420
      TabIndex        =   5
      Top             =   2640
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Index           =   4
      Left            =   120
      TabIndex        =   37
      Top             =   5880
      Width           =   4425
      Begin VB.Label Label1 
         Caption         =   "Program code extensively rewritten by Mike Morrow.  See my other VB source submissions on Planet-Source-Code."
         Height          =   405
         Index           =   4
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   4170
      End
      Begin VB.Label lblBonusInfo 
         Caption         =   "*  Max bet pays an additional 40% on the higher hands."
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   4170
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Credits:"
      Height          =   3615
      Index           =   3
      Left            =   4680
      TabIndex        =   18
      Top             =   3360
      Width           =   2895
      Begin VB.Label lblOne_Pair 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2200
         TabIndex        =   36
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label lblTwo_Pair 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2200
         TabIndex        =   35
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label lblThree_Kind 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2200
         TabIndex        =   34
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label lblStraight 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2200
         TabIndex        =   33
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label lblFlush 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2200
         TabIndex        =   32
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblFull_House 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2200
         TabIndex        =   31
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblFour_Kind 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2200
         TabIndex        =   30
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblStr_Flush 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2200
         TabIndex        =   29
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblR_Flush 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2200
         TabIndex        =   28
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "One Pair (Jacks or better)"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Two Pair"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Three of a Kind"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Straight "
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Flush "
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Full House*"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Four of a Kind*"
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Straight Flush*"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Royal Flush* "
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "How to Play:"
      Height          =   2535
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   4425
      Begin VB.Label Label1 
         Caption         =   $"frmPoker.frx":030A
         Height          =   765
         Index           =   7
         Left            =   120
         TabIndex        =   17
         Top             =   1590
         Width           =   4170
      End
      Begin VB.Label Label1 
         Caption         =   $"frmPoker.frx":03C1
         Height          =   825
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   4170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Choose number of credits to play and press ""Draw"""
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   4170
      End
   End
   Begin VB.CommandButton cmdMaxBet 
      Caption         =   "&Max Bet"
      Height          =   375
      Left            =   6060
      TabIndex        =   3
      Top             =   2220
      Width           =   735
   End
   Begin VB.CommandButton cmdBetUp 
      Caption         =   "&Bet +1"
      Height          =   375
      Left            =   6060
      TabIndex        =   1
      Top             =   1800
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Bet"
      Height          =   795
      Index           =   1
      Left            =   6120
      TabIndex        =   12
      Top             =   960
      Width           =   1455
      Begin VB.Label lblBet 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   285
         Left            =   210
         TabIndex        =   40
         Top             =   330
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Credits"
      Height          =   795
      Index           =   0
      Left            =   6120
      TabIndex        =   11
      Top             =   90
      Width           =   1455
      Begin VB.Label lblCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "50"
         Height          =   285
         Left            =   210
         TabIndex        =   39
         Top             =   330
         Width           =   1005
      End
   End
   Begin VB.CheckBox chkHold 
      Caption         =   "Hold"
      Enabled         =   0   'False
      Height          =   255
      Index           =   4
      Left            =   5100
      TabIndex        =   10
      Top             =   1770
      Width           =   650
   End
   Begin VB.CheckBox chkHold 
      Caption         =   "Hold"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   3900
      TabIndex        =   9
      Top             =   1770
      Width           =   650
   End
   Begin VB.CheckBox chkHold 
      Caption         =   "Hold"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   2700
      TabIndex        =   8
      Top             =   1770
      Width           =   650
   End
   Begin VB.CheckBox chkHold 
      Caption         =   "Hold"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   1500
      TabIndex        =   7
      Top             =   1770
      Width           =   650
   End
   Begin VB.CheckBox chkHold 
      Caption         =   "Hold"
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   6
      Top             =   1770
      Width           =   650
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "&Deal"
      Default         =   -1  'True
      Height          =   405
      Left            =   60
      TabIndex        =   0
      Top             =   2100
      Width           =   5925
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "NO PEEKING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   7830
      TabIndex        =   42
      Top             =   0
      Width           =   1275
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   9
      Left            =   7980
      Picture         =   "frmPoker.frx":04AB
      Top             =   2580
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   8
      Left            =   7980
      Picture         =   "frmPoker.frx":55ED
      Top             =   2085
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   7
      Left            =   7980
      Picture         =   "frmPoker.frx":A72F
      Top             =   1575
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   6
      Left            =   7980
      Picture         =   "frmPoker.frx":F871
      Top             =   1080
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   5
      Left            =   7980
      Picture         =   "frmPoker.frx":149B3
      Top             =   585
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   4
      Left            =   4890
      Picture         =   "frmPoker.frx":19AF5
      Top             =   240
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   3
      Left            =   3690
      Picture         =   "frmPoker.frx":1EC37
      Top             =   240
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   2
      Left            =   2490
      Picture         =   "frmPoker.frx":23D79
      Top             =   240
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   1
      Left            =   1260
      Picture         =   "frmPoker.frx":28EBB
      Top             =   240
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   0
      Left            =   60
      Picture         =   "frmPoker.frx":2DFFD
      Top             =   240
      Width           =   1065
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   13
      Top             =   2520
      Width           =   5925
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuBack 
         Caption         =   "Change Back"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnudash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnucredits 
         Caption         =   "Play Help"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuabout 
         Caption         =   "About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmPoker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub InvertCardSelection()

  Dim i As Integer
  
  For i = 0 To 4
    If chkHold(i).Enabled Then
      If chkHold(i).Value = vbChecked Then
        chkHold(i).Value = vbUnchecked
      Else
        chkHold(i).Value = vbChecked
      End If
    End If
  Next
  
End Sub

Sub ProcessHoldByNumber(KeyCode As Integer)

  Dim iImg As Integer
  
  iImg = -1  ' Get out on undefined key pressed
  
  If chkHold(0).Enabled Then
    Select Case KeyCode
      Case 97, 49  ' 1
        iImg = 0
      Case 98, 50  ' 2
        iImg = 1
      Case 99, 51  ' 3
        iImg = 2
      Case 100, 52  ' 4
        iImg = 3
      Case 101, 53  ' 5
        iImg = 4
      Case 102, 54  ' 6
        SelectAllCards
        cmdDraw.SetFocus
      Case 103, 55  ' 7
        InvertCardSelection
        cmdDraw.SetFocus
      Case Else
        Beep
    End Select
    
    If iImg = -1 Then Exit Sub
    
    If chkHold(iImg).Value = vbChecked Then
      chkHold(iImg).Value = vbUnchecked
    Else
      chkHold(iImg).Value = vbChecked
    End If
  End If
  
End Sub


Sub SelectAllCards()

 'Double click the form to select all the hold boxes

  Dim i As Integer
  Dim iInitCondx As Integer
  
  iInitCondx = chkHold(0).Value
  
  For i = 0 To 4
    If chkHold(i).Enabled Then
      If iInitCondx = vbChecked Then
        chkHold(i).Value = vbUnchecked
      Else
        chkHold(i).Value = vbChecked
      End If
    End If
  Next
  
End Sub

Private Sub chkHold_Click(Index As Integer)
  
  If chkHold(Index).Value = vbChecked Then
    imgCard(Index).BorderStyle = 1
  Else
    imgCard(Index).BorderStyle = 0
  End If
  
End Sub

Private Sub chkHold_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  ProcessHoldByNumber (KeyCode)
End Sub


Private Sub cmdBetDown_Click()

  Dim i As Integer
  
  i = CInt(lblBet) - 1
  If i < 1 Then i = 1
  
  lblBet = CStr(i)

End Sub

Private Sub cmdDraw_KeyDown(KeyCode As Integer, Shift As Integer)
  ProcessHoldByNumber (KeyCode)
End Sub
Private Sub cmdExit_Click()

  Unload Me
  End
  
End Sub

Private Sub cmdExit_KeyDown(KeyCode As Integer, Shift As Integer)
  ProcessHoldByNumber (KeyCode)
End Sub


Private Sub cmdminbet_Click()
  lblBet = 1
End Sub

'+--------------------------------------------------------------------------------+
'|                                                                                |
'|                                                                                |
'|  DRAW POKER!                                                                   |
'|  03/15/2004  Version 1.0 By noi_max                                            |
'|  04/22/2005  Version 2.0 Major rewrite by Mike Morrow                          |
'|                                                                                |
'|  This is my first game and is heavily commented.                               |
'|  I wish I could remember where the shuffle algorithm                           |
'|  came from :)                                                                  |
'|  Credit for the Bubble Sort goes to Squirm from                                |
'|  his tutorial.                                                                 |
'|                                                                                |
'|  DEPENDENCY INFORMATION                                                        |
'|   none                                                                         |
'|                                                                                |
'|  www.visualbasicforum.com                                                      |
'|  www.ILikeTheInternet.com                                                      |
'|                                                                                |
'+--------------------------------------------------------------------------------+

Private Sub Form_Load()

  Dim i As Integer
  
  Me.Top = GetSetting(App.EXEName, "Form", "Top", 0)
  If Me.Top < 0 Then Me.Top = 0
  Me.Left = GetSetting(App.EXEName, "Form", "Left", 0)
  If Me.Left < 0 Then Me.Left = 0
  
  Me.Height = 3900
  
  Shuffle  'go ahead and shuffle on load
  lblCredits = GetSetting(App.EXEName, "Parms", "Pot", 500)
  If CInt(lblCredits) < 1 Then lblCredits = 500
  
  lblBet = GetSetting(App.EXEName, "Parms", "Bet", 1)
  If CInt(lblBet) < 1 Then lblBet = 1
  
  giBack = GetSetting(App.EXEName, "Settings", "Back", 0)
  
  For i = 0 To 9
    imgCard(i) = frmCards!imgBack(giBack)
  Next
  
  lblR_Flush = R_Flush
  lblStr_Flush = Str_Flush
  lblFour_Kind = Four_Kind
  lblFull_House = Full_House
  lblFlush = Flush
  lblStraight = Straight
  lblThree_Kind = Three_Kind
  lblTwo_Pair = Two_Pair
  lblOne_Pair = One_Pair
  lblBonusInfo = "* Bet at least " & Max_Bet / 2 & " units for 40% win bonus."
  
End Sub

Private Sub Form_DblClick()
  SelectAllCards
  cmdDraw.SetFocus
End Sub
Private Sub cmdBetUp_Click()

  Dim i As Integer

 'This will add one to the current bet but not allow for
 'a bet more than the max_bet

  i = CInt(lblBet) + 1
  If i > Max_Bet Then i = 1
  
  If glBet > glCredit Then glBet = 1
  
  lblBet = CStr(i)

End Sub

Private Sub cmdMaxbet_Click()
  lblBet = Max_Bet  ' max bet
End Sub

 
Private Sub cmdDraw_Click()

  Dim j As Integer
  
  mnuBack.Enabled = False
  mnuNew.Enabled = False
 
  gbWhichHand = Not gbWhichHand
  
  If gbWhichHand = Second_Hand Then ' if it's the second turn
    For j = 0 To 4
      chkHold(j).Enabled = False  ' turn off the hold checkboxes
    Next
   'gbWhichHand = Initial_Hand
  End If
  
  lblScore.BackColor = &H8000000F
  lblScore = ""
 'If gbWhichHand = Initial_Hand Then
 '  Debug.Print "Dealing Initial Hand"
 'Else
 '  Debug.Print "Dealing Draw Hand"
 'End If
  
  DoEvents
  DealCards  ' Will act differently based on gbWhichHand
  If gbWhichHand = Initial_Hand Then
    lblScore = "Hold 0 to 5 cards then click Draw"
  Else
    mnuBack.Enabled = True
    mnuNew.Enabled = True
  End If
  
  cmdDraw.SetFocus
  
End Sub

Public Sub DealCards()

  Dim i As Integer
  Dim j As Integer
  Dim iCard As Integer
  Dim intResponse As Integer

  cmdBetUp.Enabled = False  'after we make the first draw disable the bets.
  cmdMaxBet.Enabled = False
  cmdMinBet.Enabled = False
  cmdBetDown.Enabled = False
  cmdExit.Enabled = False
  
  glCredit = CLng(lblCredits) 'set module level variables
  glBet = CLng(lblBet)
  
  If gbWhichHand = Initial_Hand Then              'if it's the first turn
    If glBet > glCredit Then  ' Can't bet more than we have
      glBet = glCredit
      lblBet = CStr(glBet)
    End If
  
    Shuffle                       'shuffle the cards
    glCredit = glCredit - glBet         'take away bet amount from credits
    lblCredits = CStr(glCredit)     're-display remaining credits
   
    For j = 0 To 4
      chkHold(j).Value = 0        'turn the hold checkboxes off
    Next
  End If
  
 'make the cards disappear by turning them over so they can later reappear
  For j = 0 To 4
    If chkHold(j).Value = 0 Then imgCard(j) = frmCards!imgBack(giBack)
  Next j
  
  If gbWhichHand = Initial_Hand Then              'if it's the first turn
    For j = 5 To 9
      imgCard(j) = frmCards!imgBack(giBack)
    Next
  End If
  
  For i = 0 To 4  '<===BEGINNING OF MAIN LOOP=====>
    If chkHold(i).Value = 0 Then  'if we're not holding any cards
      cmdDraw.Enabled = False
      Wait 5  'allow time for the cards to appear one by one.
     'iCard = Deck(giNextCard)
      iCard = Get_a_Card()
      giDealtCards(i) = iCard
      giNextCard = giNextCard + 1
      imgCard(i).Picture = frmCards!imgCard(iCard)  'show the card and the value
      If gbWhichHand = Initial_Hand Then chkHold(i).Enabled = True  'if it's our first turn, turn on the check boxes
    End If
  Next   '<====END OF MAIN LOOP====>
  imgCard(5).Picture = frmCards!imgCard(TenCards(6))  ' show the card and the value
  imgCard(6).Picture = frmCards!imgCard(TenCards(7))  'show the card and the value
  imgCard(7).Picture = frmCards!imgCard(TenCards(8))  'show the card and the value
  imgCard(8).Picture = frmCards!imgCard(TenCards(9))  'show the card and the value
  imgCard(9).Picture = frmCards!imgCard(TenCards(10))  'show the card and the value
  
  cmdDraw.Enabled = True
    
  DoEvents  ' Let the fifth card display before starting the scoring.  Looks odd otherwise.

  If gbWhichHand = Second_Hand Then
    lblScore.BackColor = &H8080FF
    lblScore.Caption = "GAME OVER"
    Score  'See if we won anything.
  End If
   
  
 'If we run out of credits it's game over man!!!!!
  If glCredit < 1 And gbWhichHand = Second_Hand And gbWin = False Then
    intResponse = MsgBox("You have run out of Credits!" & vbCrLf & _
     "Would you like to start a new game?", vbOKCancel, "Draw Poker")
    If intResponse = vbOK Then
      lblCredits = 500
      lblBet = GetSetting(App.EXEName, "Parms", "Bet", 1)
    Else
      Unload Me
      End
    End If
  End If
  
 'If the second turn is finished, put the bets back on for the next draw.
  If gbWhichHand = Second_Hand Then
    cmdBetUp.Enabled = True
    cmdMaxBet.Enabled = True
    cmdMinBet.Enabled = True
    cmdBetDown.Enabled = True
    cmdExit.Enabled = True
    cmdDraw.Caption = "&Deal"
  Else
    cmdDraw.Caption = "&Draw"
  End If
  
End Sub

Public Sub Score()
  
  'This procedure is waaaay too long. Maybe someone could optimize it for me?? :)
  '"Win" is a module procedure to beep the computer and display what we've got. It's
  'used a lot below.
  
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  Dim iCards(4) As Integer
  Dim sStraight(4) As String
  Dim sFlush(4) As String
  Dim Multi(12) As Integer
  Dim Pair As Integer
  Dim Sequence As Integer
  Dim iFlush As Integer
  Dim bTriple As Boolean
  Dim bPair As Boolean
  Dim bFour As Boolean
  Dim bRoyal As Boolean
  Dim bFlush As Boolean
  Dim bStraight As Boolean
  
  gbWin = False 'module variable: start this off as false until the Win sub
                 'tells us otherwise.
                 
 'Get the cards from our list
  For i = 0 To 4
    iCards(i) = giDealtCards(i)
  Next
  
 'Call the Card Function to determine suit (it's in the module)
  For i = 0 To 4
    sFlush(i) = Suit(iCards(i))
  Next
  
 'Compare the cards and see how many are the same suit
  For i = 1 To 4
    If sFlush(0) <> sFlush(i) Then Exit For
    iFlush = iFlush + 1
  Next
  
 'If they're all the same suit set the boolean to true
  If iFlush = 4 Then bFlush = True
  
 'Get the alpha values and convert them to numbers for sorting
  For i = 0 To 4
    sStraight(i) = Card2Char(iCards(i))
    If sStraight(i) = "J" Then
      sStraight(i) = "11"
    ElseIf sStraight(i) = "Q" Then
      sStraight(i) = "12"
    ElseIf sStraight(i) = "K" Then
      sStraight(i) = "13"
    ElseIf sStraight(i) = "A" Then
      sStraight(i) = "1"
    End If
  Next
  
 'Sort the numeric values (Bubblesort courtesy of Squirm's tutorial)
  BubbleSort sStraight
  
 'Determine if we have an A-High straight. if so, set a boolean to true
  If sStraight(0) = "1" And sStraight(1) = "10" And _
    sStraight(2) = "11" And sStraight(3) = "12" And _
    sStraight(4) = "13" Then
    bRoyal = True
  End If
     
 'Detect Ace low straights and non-A straights.
  For j = 1 To 4
    If CInt(sStraight(0)) = CInt(sStraight(j)) - j Then
      Sequence = Sequence + 1
    Else
      Exit For  ' Just get out.  The only value of sequence which matters is 4.
    End If
  Next
  
 'This loops through the array to find cards of matching values. When a match is
 'found it will add one to the integer Multi to determine if there's 2 or more
 'of the same card present. This is used for 2, 3 and 4 of a kind and sets
 'up the other conditional loops.
  For i = 0 To 12
    Multi(i) = 0
  Next
  For i = 0 To 4
    iCards(i) = iCards(i) Mod 13
  Next
  For i = 0 To 4  ' Outer loop through all cards
    For j = i + 1 To 4 ' Inner loop for all cards to the right
      If iCards(i) >= 0 Then
       'Debug.Print "Cards: " & iCards(0), iCards(1), iCards(2), iCards(3), iCards(4)
       'Debug.Print i, j, iCards(i), iCards(j)
        If iCards(i) = iCards(j) Then
          Multi(iCards(i)) = Multi(iCards(i)) + 1
          iCards(j) = -j - 1 ' Card has been used already.
        End If
      End If
    Next
  Next
  
 'This loop sets up a variable called Pair to determine if there's more
 'than 1 pair of like values (two pair)
  For i = 0 To 12
   'Debug.Print i, Multi(i), Pair
    If Multi(i) = 1 Then Pair = Pair + 1
  Next
  
 'Declare a win on 2 pair and set a boolean used for a condition to
 'avoid also hitting Jacks or Better.
  If Pair = 2 Then
    Win "Two Pair", Two_Pair, No_Bonus 'module procedure with 2 arguments
    Exit Sub
  End If
  
 'This loop structure will check for 3 or 4 of a kind and set up
 'boolean values for later comparison.
  For i = 0 To 12
   'Debug.Print i, Multi(i)
    If Multi(i) = 2 Then 'three of a kind
      bTriple = True
      Exit For
    End If
    If Multi(i) = 3 Then 'four of a kind
      bFour = True
      Exit For
    End If
  Next
  
 'If we have 4 in the variable "Sequence" then we have a straight
  If Sequence = 4 Then bStraight = True
  
 'If our Royal and Flush booleans are true well....
  If bRoyal = True And bFlush = True Then
    If glBet >= Max_Bet / 2 Then
      Win "Royal Flush", R_Flush, Add_Bonus '40% more for max bet
    Else
      Win "Royal Flush", R_Flush, No_Bonus
    End If
    Exit Sub
  End If
  
 'Here we check for a straight flush
  If bStraight = True And bFlush = True Then
    If glBet >= Max_Bet / 2 Then
      Win "Straight Flush", Str_Flush, Add_Bonus  '40% more for max bet
    Else
      Win "Straight Flush", Str_Flush, No_Bonus
    End If
    Exit Sub
  End If
  
 'Just a plain ol' flush
  If bFlush = True Then
    Win "Flush", Flush, No_Bonus
    Exit Sub
  End If
  
 'Just a plain ol' straight
  If bStraight = True Or bRoyal = True And bFlush = False Then
    Win "Straight", Straight, No_Bonus
    Exit Sub
  End If
  
 'If theres a 3 of a kind and a pair then...
  If bTriple = True And Pair = 1 Then
    If glBet >= Max_Bet / 2 Then
      Win "Full House", Full_House, Add_Bonus ' add 40% more for max bet
    Else
      Win "Full House", Full_House, No_Bonus
    End If
    Exit Sub
  End If
  
 'If there's only a three of a kind...
  If bTriple And Pair <> 1 And Not bFour Then
    Win "Three of a Kind", Three_Kind, No_Bonus
    Exit Sub
  End If
  
 'If there's a four of a kind...
  If bFour = True Then
    If glBet >= Max_Bet / 2 Then
      Win "Four of a Kind", Four_Kind, Add_Bonus '40% more for max bet
    Else
      Win "Four of a Kind", Four_Kind, No_Bonus
    End If
    Exit Sub
  End If
  
 'If it's only Jacks or Better...
  If Not bTriple And Not bPair And Not bFour Then
    If Multi(11) = 1 Or Multi(12) = 1 Or Multi(0) = 1 Or Multi(1) = 1 Then _
     Win "Jacks or Better", One_Pair, No_Bonus
  End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

  SaveSetting App.EXEName, "Form", "Top", Me.Top
  SaveSetting App.EXEName, "Form", "Left", Me.Left
  SaveSetting App.EXEName, "Parms", "Pot", lblCredits
  SaveSetting App.EXEName, "Parms", "Bet", lblBet
  
End Sub

Private Sub imgCard_Click(Index As Integer)

  If chkHold(Index).Enabled = False Then Exit Sub  ' Don't change if not enabled.
  
  If chkHold(Index).Value = vbChecked Then
    chkHold(Index).Value = vbUnchecked
  Else
    chkHold(Index).Value = vbChecked
  End If
  
End Sub

Private Sub mnuabout_Click()

MsgBox "Draw Poker v 1.0" & "by noi_max" & vbCrLf & "Extensively rewritten by Mike Morrow", vbOKOnly, "Draw Poker"
'my shameless self promotion

End Sub

Private Sub mnuBack_Click()
  
  Dim i As Integer
  
  frmBacks.Show vbModal
  
  SaveSetting App.EXEName, "Settings", "Back", giBack
  
  For i = 0 To 4
    imgCard(i) = frmCards!imgBack(giBack)
  Next
  
End Sub

Private Sub mnuCheat_Click()
  Debug.Print Me.Width
End Sub

Private Sub mnucredits_Click()

  If mnucredits.Checked Then
    Me.Height = 3900 'show only the main game
    mnucredits.Checked = False
  Else
    Me.Height = 7870 'show the credits and basic help
    mnucredits.Checked = True
  End If
  
End Sub

Private Sub mnuExit_Click()

  Unload Me  ' bye!
  End

End Sub

Private Sub mnuNew_Click()

Unload Me
Load frmPoker
frmPoker.Show 'reload the form for a new game.

End Sub

