VERSION 5.00
Begin VB.Form frmWinrar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Congrats!"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmWinner.frx":0000
   ScaleHeight     =   4635
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Wow!"
      Height          =   525
      Left            =   3600
      TabIndex        =   1
      Top             =   3900
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Congrats!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   960
      TabIndex        =   0
      Top             =   330
      Width           =   1455
   End
End
Attribute VB_Name = "frmWinrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub
