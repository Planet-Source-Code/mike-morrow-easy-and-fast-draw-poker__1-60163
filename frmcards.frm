VERSION 5.00
Begin VB.Form frmCards 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test Cards"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   1
      Left            =   0
      Picture         =   "frmCards.frx":0000
      Top             =   0
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   2
      Left            =   240
      Picture         =   "frmCards.frx":5142
      Top             =   0
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   3
      Left            =   480
      Picture         =   "frmCards.frx":A284
      Top             =   0
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   4
      Left            =   720
      Picture         =   "frmCards.frx":F3C6
      Top             =   0
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   5
      Left            =   960
      Picture         =   "frmCards.frx":14508
      Top             =   0
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   6
      Left            =   1200
      Picture         =   "frmCards.frx":1964A
      Top             =   0
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   7
      Left            =   1440
      Picture         =   "frmCards.frx":1E78C
      Top             =   0
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   8
      Left            =   1680
      Picture         =   "frmCards.frx":238CE
      Top             =   0
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   9
      Left            =   1920
      Picture         =   "frmCards.frx":28A10
      Top             =   0
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   10
      Left            =   2160
      Picture         =   "frmCards.frx":2DB52
      Top             =   0
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   11
      Left            =   2400
      Picture         =   "frmCards.frx":32C94
      Top             =   0
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   12
      Left            =   2640
      Picture         =   "frmCards.frx":37DD6
      Top             =   0
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   13
      Left            =   2880
      Picture         =   "frmCards.frx":3CF18
      Top             =   0
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   14
      Left            =   0
      Picture         =   "frmCards.frx":4205A
      Top             =   2000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   15
      Left            =   240
      Picture         =   "frmCards.frx":4719C
      Top             =   2000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   16
      Left            =   480
      Picture         =   "frmCards.frx":4C2DE
      Top             =   2000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   17
      Left            =   720
      Picture         =   "frmCards.frx":51420
      Top             =   2000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   18
      Left            =   960
      Picture         =   "frmCards.frx":56562
      Top             =   2000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   19
      Left            =   1200
      Picture         =   "frmCards.frx":5B6A4
      Top             =   2000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   20
      Left            =   1440
      Picture         =   "frmCards.frx":607E6
      Top             =   2000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   21
      Left            =   1680
      Picture         =   "frmCards.frx":65928
      Top             =   2000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   22
      Left            =   1920
      Picture         =   "frmCards.frx":6AA6A
      Top             =   2000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   23
      Left            =   2160
      Picture         =   "frmCards.frx":6FBAC
      Top             =   2000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   24
      Left            =   2400
      Picture         =   "frmCards.frx":74CEE
      Top             =   2000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   25
      Left            =   2640
      Picture         =   "frmCards.frx":79E30
      Top             =   2000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   26
      Left            =   2880
      Picture         =   "frmCards.frx":7EF72
      Top             =   2000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   27
      Left            =   0
      Picture         =   "frmCards.frx":840B4
      Top             =   4000
      Width           =   1050
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   28
      Left            =   240
      Picture         =   "frmCards.frx":89076
      Top             =   4000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   29
      Left            =   480
      Picture         =   "frmCards.frx":8E1B8
      Top             =   4000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   30
      Left            =   720
      Picture         =   "frmCards.frx":932FA
      Top             =   4000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   31
      Left            =   960
      Picture         =   "frmCards.frx":9843C
      Top             =   4000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   32
      Left            =   1200
      Picture         =   "frmCards.frx":9D57E
      Top             =   4000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   33
      Left            =   1440
      Picture         =   "frmCards.frx":A26C0
      Top             =   4000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   34
      Left            =   1680
      Picture         =   "frmCards.frx":A7802
      Top             =   4000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   35
      Left            =   1920
      Picture         =   "frmCards.frx":AC944
      Top             =   4000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   36
      Left            =   2160
      Picture         =   "frmCards.frx":B1A86
      Top             =   4000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   37
      Left            =   2400
      Picture         =   "frmCards.frx":B6BC8
      Top             =   4000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   38
      Left            =   2640
      Picture         =   "frmCards.frx":BBD0A
      Top             =   4000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   39
      Left            =   2880
      Picture         =   "frmCards.frx":C0E4C
      Top             =   4000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   40
      Left            =   0
      Picture         =   "frmCards.frx":C5F8E
      Top             =   6000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   41
      Left            =   240
      Picture         =   "frmCards.frx":CB0D0
      Top             =   6000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   42
      Left            =   480
      Picture         =   "frmCards.frx":D0212
      Top             =   6000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   43
      Left            =   720
      Picture         =   "frmCards.frx":D5354
      Top             =   6000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   44
      Left            =   960
      Picture         =   "frmCards.frx":DA496
      Top             =   6000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   45
      Left            =   1200
      Picture         =   "frmCards.frx":DF5D8
      Top             =   6000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   46
      Left            =   1440
      Picture         =   "frmCards.frx":E471A
      Top             =   6000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   47
      Left            =   1680
      Picture         =   "frmCards.frx":E985C
      Top             =   6000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   48
      Left            =   1920
      Picture         =   "frmCards.frx":EE99E
      Top             =   6000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   49
      Left            =   2160
      Picture         =   "frmCards.frx":F3AE0
      Top             =   6000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   50
      Left            =   2400
      Picture         =   "frmCards.frx":F8C22
      Top             =   6000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   51
      Left            =   2640
      Picture         =   "frmCards.frx":FDD64
      Top             =   6000
      Width           =   1065
   End
   Begin VB.Image imgCard 
      Height          =   1440
      Index           =   52
      Left            =   2880
      Picture         =   "frmCards.frx":102EA6
      Top             =   6000
      Width           =   1065
   End
   Begin VB.Image imgBack 
      Height          =   1440
      Index           =   0
      Left            =   4050
      Picture         =   "frmCards.frx":107FE8
      Top             =   0
      Width           =   1065
   End
   Begin VB.Image imgBack 
      Height          =   1440
      Index           =   1
      Left            =   4050
      Picture         =   "frmCards.frx":10D12A
      Top             =   2000
      Width           =   1065
   End
   Begin VB.Image imgBack 
      Height          =   1440
      Index           =   2
      Left            =   4080
      Picture         =   "frmCards.frx":11226C
      Top             =   4000
      Width           =   1065
   End
   Begin VB.Image imgBack 
      Height          =   1440
      Index           =   3
      Left            =   4080
      Picture         =   "frmCards.frx":1173AE
      Top             =   6000
      Width           =   1065
   End
   Begin VB.Image imgBack 
      Height          =   1440
      Index           =   4
      Left            =   6450
      Picture         =   "frmCards.frx":11C4F0
      Top             =   0
      Width           =   1065
   End
   Begin VB.Image imgBack 
      Height          =   1440
      Index           =   5
      Left            =   6450
      Picture         =   "frmCards.frx":121632
      Top             =   2000
      Width           =   1065
   End
   Begin VB.Image imgBack 
      Height          =   1440
      Index           =   6
      Left            =   6480
      Picture         =   "frmCards.frx":126774
      Top             =   4000
      Width           =   1065
   End
   Begin VB.Image imgBack 
      Height          =   1440
      Index           =   7
      Left            =   6450
      Picture         =   "frmCards.frx":12B8B6
      Top             =   6000
      Width           =   1065
   End
   Begin VB.Image imgTop 
      Height          =   45
      Index           =   0
      Left            =   5190
      Picture         =   "frmCards.frx":1309F8
      Top             =   0
      Width           =   1065
   End
   Begin VB.Image imgTop 
      Height          =   45
      Index           =   1
      Left            =   5190
      Picture         =   "frmCards.frx":130CC2
      Top             =   2000
      Width           =   1065
   End
   Begin VB.Image imgTop 
      Height          =   45
      Index           =   2
      Left            =   5220
      Picture         =   "frmCards.frx":130F8C
      Top             =   4000
      Width           =   1065
   End
   Begin VB.Image imgTop 
      Height          =   45
      Index           =   3
      Left            =   5220
      Picture         =   "frmCards.frx":131256
      Top             =   6000
      Width           =   1065
   End
   Begin VB.Image imgTop 
      Height          =   45
      Index           =   4
      Left            =   7590
      Picture         =   "frmCards.frx":131520
      Top             =   0
      Width           =   1065
   End
   Begin VB.Image imgTop 
      Height          =   45
      Index           =   6
      Left            =   7590
      Picture         =   "frmCards.frx":1317EA
      Top             =   4000
      Width           =   1065
   End
   Begin VB.Image imgTop 
      Height          =   45
      Index           =   7
      Left            =   7590
      Picture         =   "frmCards.frx":131AB4
      Top             =   6000
      Width           =   1065
   End
   Begin VB.Image imgSide 
      Height          =   1440
      Index           =   0
      Left            =   6330
      Picture         =   "frmCards.frx":131D7E
      Top             =   0
      Width           =   30
   End
   Begin VB.Image imgSide 
      Height          =   1425
      Index           =   1
      Left            =   6330
      Picture         =   "frmCards.frx":1320C0
      Top             =   2000
      Width           =   30
   End
   Begin VB.Image imgSide 
      Height          =   1425
      Index           =   2
      Left            =   6330
      Picture         =   "frmCards.frx":1323FA
      Top             =   4000
      Width           =   30
   End
   Begin VB.Image imgSide 
      Height          =   1440
      Index           =   3
      Left            =   6330
      Picture         =   "frmCards.frx":132734
      Top             =   6000
      Width           =   30
   End
   Begin VB.Image imgSide 
      Height          =   1440
      Index           =   4
      Left            =   8730
      Picture         =   "frmCards.frx":132A76
      Top             =   0
      Width           =   30
   End
   Begin VB.Image imgSide 
      Height          =   1425
      Index           =   5
      Left            =   8730
      Picture         =   "frmCards.frx":132DB8
      Top             =   2000
      Width           =   30
   End
   Begin VB.Image imgSide 
      Height          =   1440
      Index           =   6
      Left            =   8730
      Picture         =   "frmCards.frx":1330F2
      Top             =   4000
      Width           =   30
   End
   Begin VB.Image imgSide 
      Height          =   1440
      Index           =   7
      Left            =   8730
      Picture         =   "frmCards.frx":133434
      Top             =   6000
      Width           =   30
   End
   Begin VB.Image imgTop 
      Height          =   45
      Index           =   5
      Left            =   7590
      Picture         =   "frmCards.frx":133776
      Top             =   2000
      Width           =   1065
   End
End
Attribute VB_Name = "frmCards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

