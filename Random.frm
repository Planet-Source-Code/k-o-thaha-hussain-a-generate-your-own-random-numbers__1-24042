VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmRandom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Random Number Generator by K. O. Thaha Hussain"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   Icon            =   "Random.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTRandom 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   8281
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Mixed Congruential"
      TabPicture(0)   =   "Random.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtA1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtB1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtM1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtPrevious1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdMixed"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtRandom1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Multiplicative Congruential"
      TabPicture(1)   =   "Random.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtA2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtM2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtPrevious2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdMultiplicative"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtRandom2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label10"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label8"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label7"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label5"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Additive Congruential"
      TabPicture(2)   =   "Random.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtB3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtM3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "txtPrevious3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdAdditive"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "txtRandom3"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label13"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label12"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label11"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label9"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      Begin VB.Frame Frame3 
         Caption         =   "Formula"
         Height          =   1335
         Left            =   -74520
         TabIndex        =   32
         Top             =   3000
         Width           =   5655
         Begin VB.Image Image3 
            BorderStyle     =   1  'Fixed Single
            Height          =   930
            Left            =   240
            Picture         =   "Random.frx":035E
            Top             =   240
            Width           =   5280
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Formula"
         Height          =   1335
         Left            =   -74520
         TabIndex        =   31
         Top             =   3000
         Width           =   5655
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   930
            Left            =   240
            Picture         =   "Random.frx":F028
            Top             =   240
            Width           =   5280
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Formula"
         Height          =   1335
         Left            =   480
         TabIndex        =   30
         Top             =   3000
         Width           =   5655
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   930
            Left            =   240
            Picture         =   "Random.frx":1DCF2
            Top             =   240
            Width           =   5280
         End
      End
      Begin VB.TextBox txtB3 
         Height          =   285
         Left            =   -73440
         TabIndex        =   25
         Text            =   "18"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtM3 
         Height          =   285
         Left            =   -73440
         TabIndex        =   24
         Text            =   "23"
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtPrevious3 
         Height          =   285
         Left            =   -69480
         TabIndex        =   23
         Text            =   "1"
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmdAdditive 
         Caption         =   "Generate"
         Height          =   375
         Left            =   -70560
         TabIndex        =   22
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtRandom3 
         Height          =   285
         Left            =   -72720
         TabIndex        =   21
         Top             =   2400
         Width           =   3615
      End
      Begin VB.TextBox txtA2 
         Height          =   285
         Left            =   -73440
         TabIndex        =   16
         Text            =   "16"
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtM2 
         Height          =   285
         Left            =   -73440
         TabIndex        =   15
         Text            =   "23"
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtPrevious2 
         Height          =   285
         Left            =   -69480
         TabIndex        =   14
         Text            =   "1"
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmdMultiplicative 
         Caption         =   "Generate"
         Height          =   375
         Left            =   -70560
         TabIndex        =   13
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtRandom2 
         Height          =   285
         Left            =   -72720
         TabIndex        =   12
         Top             =   2400
         Width           =   3615
      End
      Begin VB.TextBox txtRandom1 
         Height          =   285
         Left            =   2280
         TabIndex        =   10
         Top             =   2400
         Width           =   3615
      End
      Begin VB.CommandButton cmdMixed 
         Caption         =   "Generate"
         Height          =   375
         Left            =   4440
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtPrevious1 
         Height          =   285
         Left            =   5520
         TabIndex        =   8
         Text            =   "1"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtM1 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Text            =   "23"
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtB1 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Text            =   "18"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtA1 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Text            =   "16"
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label13 
         Caption         =   "B ="
         Height          =   375
         Left            =   -74160
         TabIndex        =   29
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "M ="
         Height          =   375
         Left            =   -74160
         TabIndex        =   28
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Seed for next Iteration"
         Height          =   495
         Left            =   -70680
         TabIndex        =   27
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Random Number"
         Height          =   255
         Left            =   -74160
         TabIndex        =   26
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "A ="
         Height          =   375
         Left            =   -74160
         TabIndex        =   20
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "M ="
         Height          =   375
         Left            =   -74160
         TabIndex        =   19
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Seed for next Iteration"
         Height          =   495
         Left            =   -70680
         TabIndex        =   18
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Random Number"
         Height          =   255
         Left            =   -74160
         TabIndex        =   17
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Random Number"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Seed for next Iteration"
         Height          =   495
         Left            =   4320
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "M ="
         Height          =   375
         Left            =   840
         TabIndex        =   5
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "B ="
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "A ="
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Top             =   600
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmRandom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'*                Random Number Generator                   *
'*    (C)    K. O. Thaha Hussain, Analyst Programmer        *
'*              Indusware Solutions, India                  *
'*          EMail : thaha@induswareonline.com               *
'*      URL  http://www.bcity.com/thahahussain              *
'*      Company   http://www.induswareonline.com            *
'************************************************************
'Replace seed with the current milli-second to make it an intelligent random.
Option Explicit


Private Sub cmdAdditive_Click()
  txtRandom3 = AdditiveRandom(Int(Val(txtB3)), Int(Val(txtM3)), Int(Val(txtPrevious3)))
  txtPrevious3 = txtRandom3
End Sub

Private Sub cmdMixed_Click()
  txtRandom1 = MixedRandom(Int(Val(txtA1)), Int(Val(txtB1)), Int(Val(txtM1)), Int(Val(txtPrevious1)))
  txtPrevious1 = txtRandom1
End Sub


Private Sub cmdMultiplicative_Click()
   txtRandom2 = MultiplicativeRandom(Int(Val(txtA2)), Int(Val(txtM2)), Int(Val(txtPrevious2)))
   txtPrevious2 = txtRandom2
End Sub

Function MultiplicativeRandom(A As Integer, M As Integer, PreviousRandom As Integer) As Integer
  MultiplicativeRandom = (A * PreviousRandom) Mod M
End Function
Function AdditiveRandom(B As Integer, M As Integer, PreviousRandom As Integer) As Integer
  AdditiveRandom = (PreviousRandom + B) Mod M
End Function
Function MixedRandom(A As Integer, B As Integer, M As Integer, PreviousRandom As Integer) As Integer
  MixedRandom = (A * PreviousRandom + B) Mod M
End Function
