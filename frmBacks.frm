VERSION 5.00
Begin VB.Form frmBacks 
   Caption         =   "Card Back Selection"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5325
   ScaleWidth      =   6315
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optBack 
      Caption         =   "Option1"
      Height          =   195
      Index           =   7
      Left            =   5457
      TabIndex        =   9
      Top             =   4100
      Width           =   200
   End
   Begin VB.OptionButton optBack 
      Caption         =   "Option1"
      Height          =   195
      Index           =   6
      Left            =   3857
      TabIndex        =   8
      Top             =   4100
      Width           =   200
   End
   Begin VB.OptionButton optBack 
      Caption         =   "Option1"
      Height          =   195
      Index           =   5
      Left            =   2257
      TabIndex        =   7
      Top             =   4100
      Width           =   200
   End
   Begin VB.OptionButton optBack 
      Caption         =   "Option1"
      Height          =   195
      Index           =   4
      Left            =   657
      TabIndex        =   6
      Top             =   4100
      Width           =   200
   End
   Begin VB.OptionButton optBack 
      Caption         =   "Option1"
      Height          =   195
      Index           =   3
      Left            =   5457
      TabIndex        =   5
      Top             =   1800
      Width           =   200
   End
   Begin VB.OptionButton optBack 
      Caption         =   "Option1"
      Height          =   195
      Index           =   2
      Left            =   3857
      TabIndex        =   4
      Top             =   1800
      Width           =   200
   End
   Begin VB.OptionButton optBack 
      Caption         =   "Option1"
      Height          =   195
      Index           =   1
      Left            =   2257
      TabIndex        =   3
      Top             =   1800
      Width           =   200
   End
   Begin VB.OptionButton optBack 
      Caption         =   "Option1"
      Height          =   195
      Index           =   0
      Left            =   657
      TabIndex        =   2
      Top             =   1800
      Width           =   200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   3465
      TabIndex        =   1
      Top             =   4470
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   525
      Left            =   1605
      TabIndex        =   0
      Top             =   4500
      Width           =   1245
   End
   Begin VB.Image imgBack 
      Height          =   1440
      Index           =   7
      Left            =   5025
      Picture         =   "frmBacks.frx":0000
      Top             =   2450
      Width           =   1065
   End
   Begin VB.Image imgBack 
      Height          =   1440
      Index           =   6
      Left            =   3425
      Picture         =   "frmBacks.frx":5142
      Top             =   2450
      Width           =   1065
   End
   Begin VB.Image imgBack 
      Height          =   1440
      Index           =   5
      Left            =   1825
      Picture         =   "frmBacks.frx":A284
      Top             =   2450
      Width           =   1065
   End
   Begin VB.Image imgBack 
      Height          =   1440
      Index           =   4
      Left            =   225
      Picture         =   "frmBacks.frx":F3C6
      Top             =   2450
      Width           =   1065
   End
   Begin VB.Image imgBack 
      Height          =   1440
      Index           =   3
      Left            =   5025
      Picture         =   "frmBacks.frx":14508
      Top             =   150
      Width           =   1065
   End
   Begin VB.Image imgBack 
      Height          =   1440
      Index           =   2
      Left            =   3425
      Picture         =   "frmBacks.frx":1964A
      Top             =   150
      Width           =   1065
   End
   Begin VB.Image imgBack 
      Height          =   1440
      Index           =   1
      Left            =   1825
      Picture         =   "frmBacks.frx":1E78C
      Top             =   150
      Width           =   1065
   End
   Begin VB.Image imgBack 
      Height          =   1440
      Index           =   0
      Left            =   225
      Picture         =   "frmBacks.frx":238CE
      Top             =   150
      Width           =   1065
   End
End
Attribute VB_Name = "frmBacks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  
  Dim i As Integer
  
  For i = 0 To 7
    If optBack(i) Then
      giBack = i
      Exit For
    End If
  Next
  
  Unload Me  ' Probably not necessary but... one never knows.

End Sub

Private Sub Form_Load()
  optBack(giBack).Value = True
End Sub


Private Sub imgBack_Click(Index As Integer)
  optBack(Index).Value = True
End Sub


