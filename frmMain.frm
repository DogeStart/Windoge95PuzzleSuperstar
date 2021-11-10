VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Windoge95 Puzzle Superstar WOW"
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   240
   ClientWidth     =   3870
   DrawWidth       =   10
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCheck 
      Caption         =   "CHK"
      Height          =   345
      Left            =   990
      TabIndex        =   23
      Top             =   600
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.CommandButton Command1 
      Caption         =   "WIN"
      Height          =   345
      Left            =   990
      TabIndex        =   22
      Top             =   240
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.CommandButton btnTEMP 
      Caption         =   "BAK"
      Height          =   315
      Left            =   2280
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.CommandButton CMD 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   0
      Left            =   390
      Picture         =   "frmMain.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "0"
      ToolTipText     =   "Click to Swap number with adjacent location if it is empty."
      Top             =   1275
      Width           =   795
   End
   Begin VB.CommandButton cmdShuffle 
      BackColor       =   &H80000000&
      Height          =   615
      Left            =   1620
      Picture         =   "frmMain.frx":2930
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Re-arrange numbers"
      Top             =   210
      Width           =   615
   End
   Begin VB.CommandButton CMD 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   15
      Left            =   2745
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Click to Swap number with adjacent location if it is empty."
      Top             =   3450
      Width           =   795
   End
   Begin VB.CommandButton CMD 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   14
      Left            =   1970
      Picture         =   "frmMain.frx":558D
      Style           =   1  'Graphical
      TabIndex        =   14
      Tag             =   "14"
      ToolTipText     =   "Click to Swap number with adjacent location if it is empty."
      Top             =   3450
      Width           =   795
   End
   Begin VB.CommandButton CMD 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   13
      Left            =   1170
      Picture         =   "frmMain.frx":62D7
      Style           =   1  'Graphical
      TabIndex        =   13
      Tag             =   "13"
      ToolTipText     =   "Click to Swap number with adjacent location if it is empty."
      Top             =   3450
      Width           =   795
   End
   Begin VB.CommandButton CMD 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   12
      Left            =   390
      Picture         =   "frmMain.frx":6F73
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "12"
      ToolTipText     =   "Click to Swap number with adjacent location if it is empty."
      Top             =   3450
      Width           =   795
   End
   Begin VB.CommandButton CMD 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   11
      Left            =   2745
      Picture         =   "frmMain.frx":7AEC
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "11"
      ToolTipText     =   "Click to Swap number with adjacent location if it is empty."
      Top             =   2730
      Width           =   795
   End
   Begin VB.CommandButton CMD 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   10
      Left            =   1970
      Picture         =   "frmMain.frx":86BE
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "10"
      ToolTipText     =   "Click to Swap number with adjacent location if it is empty."
      Top             =   2730
      Width           =   795
   End
   Begin VB.CommandButton CMD 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   9
      Left            =   1170
      Picture         =   "frmMain.frx":9378
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "9"
      ToolTipText     =   "Click to Swap number with adjacent location if it is empty."
      Top             =   2730
      Width           =   795
   End
   Begin VB.CommandButton CMD 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   8
      Left            =   390
      Picture         =   "frmMain.frx":A253
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "8"
      ToolTipText     =   "Click to Swap number with adjacent location if it is empty."
      Top             =   2730
      Width           =   795
   End
   Begin VB.CommandButton CMD 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   7
      Left            =   2745
      Picture         =   "frmMain.frx":B006
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "7"
      ToolTipText     =   "Click to Swap number with adjacent location if it is empty."
      Top             =   2010
      Width           =   795
   End
   Begin VB.CommandButton CMD 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   6
      Left            =   1970
      Picture         =   "frmMain.frx":BA6D
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "6"
      ToolTipText     =   "Click to Swap number with adjacent location if it is empty."
      Top             =   2010
      Width           =   795
   End
   Begin VB.CommandButton CMD 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   5
      Left            =   1170
      Picture         =   "frmMain.frx":C999
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "5"
      ToolTipText     =   "Click to Swap number with adjacent location if it is empty."
      Top             =   2010
      Width           =   795
   End
   Begin VB.CommandButton CMD 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   4
      Left            =   390
      Picture         =   "frmMain.frx":D996
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "4"
      ToolTipText     =   "Click to Swap number with adjacent location if it is empty."
      Top             =   2010
      Width           =   795
   End
   Begin VB.CommandButton CMD 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   3
      Left            =   2745
      Picture         =   "frmMain.frx":E67A
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "3"
      ToolTipText     =   "Click to Swap number with adjacent location if it is empty."
      Top             =   1275
      Width           =   795
   End
   Begin VB.CommandButton CMD 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   2
      Left            =   1970
      Picture         =   "frmMain.frx":EF44
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "2"
      ToolTipText     =   "Click to Swap number with adjacent location if it is empty."
      Top             =   1275
      Width           =   795
   End
   Begin VB.CommandButton CMD 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   1
      Left            =   1170
      Picture         =   "frmMain.frx":F9D3
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "1"
      ToolTipText     =   "Click to Swap number with adjacent location if it is empty."
      Top             =   1275
      Width           =   795
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000003&
      BorderWidth     =   4
      Height          =   885
      Left            =   30
      Top             =   90
      Width           =   3825
   End
   Begin VB.Label lblTag 
      Height          =   405
      Left            =   1740
      TabIndex        =   21
      Top             =   4740
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Tag"
      Height          =   285
      Left            =   1170
      TabIndex        =   20
      Top             =   4800
      Width           =   435
   End
   Begin VB.Label lblIndex 
      Caption         =   "0"
      Height          =   405
      Left            =   690
      TabIndex        =   19
      Top             =   4710
      Width           =   345
   End
   Begin VB.Label Label1 
      Caption         =   "Index"
      Height          =   255
      Left            =   90
      TabIndex        =   18
      Top             =   4770
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   540
      Left            =   1680
      Picture         =   "frmMain.frx":10703
      Top             =   210
      Width           =   555
   End
   Begin VB.Shape shpBorderGame 
      BorderColor     =   &H80000003&
      BorderWidth     =   4
      Height          =   3330
      Left            =   30
      Top             =   1050
      Width           =   3825
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Windoge95 Puzzle Superstar
'DogeStart

Dim GridArray(16, 4) As Integer
Dim CurrentBlankIndex As Integer

'Don't code like this at home. study algorithms in school kids
Private Sub Grid_Init()
    
    
    GridArray(0, 0) = 1
    GridArray(0, 1) = 4
    
  
    GridArray(1, 0) = 0
    GridArray(1, 1) = 2
    GridArray(1, 2) = 5
    

    GridArray(2, 0) = 1
    GridArray(2, 1) = 3
    GridArray(2, 2) = 6
    

    GridArray(3, 0) = 2
    GridArray(3, 1) = 7

   
    GridArray(4, 0) = 0
    GridArray(4, 1) = 5
    GridArray(4, 2) = 8


    GridArray(5, 0) = 1
    GridArray(5, 1) = 4
    GridArray(5, 2) = 6
    GridArray(5, 3) = 9
    

    GridArray(6, 0) = 2
    GridArray(6, 1) = 5
    GridArray(6, 2) = 7
    GridArray(6, 3) = 10
   

    GridArray(7, 0) = 3
    GridArray(7, 1) = 6
    GridArray(7, 2) = 11
    
  
    GridArray(8, 0) = 4
    GridArray(8, 1) = 9
    GridArray(8, 2) = 12
    
   
    GridArray(9, 0) = 5
    GridArray(9, 1) = 8
    GridArray(9, 2) = 10
    GridArray(9, 3) = 13

    GridArray(10, 0) = 6
    GridArray(10, 1) = 9
    GridArray(10, 2) = 11
    GridArray(10, 3) = 14
    
    
    GridArray(11, 0) = 7
    GridArray(11, 1) = 10
    GridArray(11, 3) = 15
    
    
    GridArray(12, 0) = 8
    GridArray(12, 1) = 13
    

    GridArray(13, 0) = 9
    GridArray(13, 1) = 12
    GridArray(13, 2) = 14
    
    GridArray(14, 0) = 13
    GridArray(14, 1) = 10
    GridArray(14, 2) = 15
    
    GridArray(15, 0) = 11
    GridArray(15, 1) = 14
    

End Sub

Private Function ExistsNeighbor(Index As Integer, TargetIndex As Integer) As Boolean
    Dim i
    For i = 0 To 3
        If GridArray(Index, i) = TargetIndex Then
            ExistsNeighbor = True
            Exit Function
        End If
        
    Next
    ExistsNeighbor = False
End Function


Private Sub CMD_Click(Index As Integer)
    If CMD(Index).Tag = "" Then Exit Sub
    
    If ExistsNeighbor(Index, CurrentBlankIndex) = True Then
        Call SWAP(Index, CurrentBlankIndex)
    End If
    
End Sub


Private Sub cmdCheck_Click()
    CHKWIN
End Sub

Private Sub cmdShuffle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdShuffle.Visible = False
End Sub


Private Sub Command1_Click()
    WIN
End Sub

Private Sub Form_Load()
    Grid_Init
    Shuffle
End Sub

'Swap two tiles
Private Sub SWAP(A As Integer, B As Integer)
    btnTEMP.Tag = CMD(A).Tag
    btnTEMP.Picture = CMD(A).Picture

    CMD(A).Picture = CMD(B).Picture
    CMD(A).Tag = CMD(B).Tag
    
    CMD(B).Picture = btnTEMP.Picture
    CMD(B).Tag = btnTEMP.Tag
    
    If CMD(A).Tag = "" Then
        CurrentBlankIndex = A
    Else
        CurrentBlankIndex = B
    End If
End Sub

'Check for a winning hand
Private Sub CHKWIN()
    Dim x
    For x = 0 To 14
        If CMD(x).Tag <> x Then
            Exit Sub
        End If
        
        If x = 14 Then
           WIN
        End If
    Next x
End Sub

'Shows win
Private Sub WIN()
    Load frmWinrar
    frmWinrar.Show
End Sub

'Shuffle Deck
Private Sub Shuffle()

Dim bug, x, y, c

BACK:
bug = 0

Randomize

For x = 0 To 14
RAM:
    c = Rnd * 15
    If c = 0 Then GoTo RAM
    For y = 0 To x
        If CMD(y).Tag = CStr(c) Then GoTo RAM
    bug = bug + 1
    If bug = 250 Then GoTo BACK
    Next y
    
    'Backup x
    btnTEMP.Picture = CMD(x).Picture
    btnTEMP.Tag = CMD(x).Tag
    
    'new
    CMD(x).Picture = CMD(c).Picture
    CMD(x).Tag = CMD(c).Tag
    CMD(c).Picture = btnTEMP.Picture
    CMD(c).Tag = btnTEMP.Tag
    
Next x

'Find the blank
For x = 0 To 15
    If CMD(x).Tag = "" Then
        CurrentBlankIndex = x
        Exit For
    End If
Next

End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Shuffle
    cmdShuffle.Visible = True
End Sub


Private Sub mnuHelpAbout_Click()
    frmHelpAbout.Show vbModal
End Sub

Private Sub mnuQuit_Click()
    Unload Me
    End
End Sub
