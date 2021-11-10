VERSION 5.00
Begin VB.Form frmHelpAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Such About"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Much OK"
      Height          =   375
      Left            =   900
      TabIndex        =   3
      Top             =   3150
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "Visit:"
      Height          =   285
      Left            =   90
      TabIndex        =   2
      Top             =   1290
      Width           =   1065
   End
   Begin VB.Label Label2 
      Caption         =   "Twitter: @windoge_95"
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   1620
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   3600
      Left            =   2730
      Picture         =   "frmHelpAbout.frx":0000
      Top             =   0
      Width           =   3690
   End
   Begin VB.Label Label1 
      Caption         =   "Windoge95 Puzzle Superstar"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2745
   End
End
Attribute VB_Name = "frmHelpAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub
