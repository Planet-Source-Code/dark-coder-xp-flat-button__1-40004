VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "XP Flat Button by Dark Coder"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   3030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin prjXPFlatButton.XPFlatButton XPFlatButton6 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Caption         =   "&About"
      Caption         =   "&About"
   End
   Begin prjXPFlatButton.XPFlatButton XPFlatButton1 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   615
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   ""
      Picture         =   "frmTest.frx":0000
   End
   Begin prjXPFlatButton.XPFlatButton XPFlatButton2 
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   0
      Width           =   615
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   ""
      Picture         =   "frmTest.frx":08DA
   End
   Begin prjXPFlatButton.XPFlatButton XPFlatButton3 
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Width           =   615
      _ExtentX        =   873
      _ExtentY        =   873
      Caption         =   ""
      Picture         =   "frmTest.frx":11B4
   End
   Begin prjXPFlatButton.XPFlatButton XPFlatButton4 
      Height          =   615
      Left            =   1800
      TabIndex        =   3
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Caption         =   ""
      Picture         =   "frmTest.frx":3966
   End
   Begin prjXPFlatButton.XPFlatButton XPFlatButton5 
      Height          =   615
      Left            =   2400
      TabIndex        =   4
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
      Caption         =   ""
      Picture         =   "frmTest.frx":6118
   End
   Begin prjXPFlatButton.XPFlatButton XPFlatButton7 
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "&Exit"
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub XPFlatButton1_Click()
MsgBox "A script icon button"
End Sub

Private Sub XPFlatButton2_Click()
MsgBox "a crate icon"
End Sub


Private Sub XPFlatButton3_Click()
"a comm icon"
End Sub


Private Sub XPFlatButton4_Click()
MsgBox "another comm icon"
End Sub


Private Sub XPFlatButton5_Click()
MsgBox "an ink bottle icon"
End Sub


Private Sub XPFlatButton6_Click()
MsgBox "XP Flat Button by Dark Coder"
End Sub

Private Sub XPFlatButton7_Click()
End
End Sub
