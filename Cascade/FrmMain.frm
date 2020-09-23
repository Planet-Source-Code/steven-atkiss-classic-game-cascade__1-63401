VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   463
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   638
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   75
      Left            =   840
      Top             =   1980
   End
   Begin PicClip.PictureClip Balls 
      Index           =   0
      Left            =   1920
      Top             =   3300
      _ExtentX        =   21167
      _ExtentY        =   1323
      _Version        =   393216
      Cols            =   16
      Picture         =   "FrmMain.frx":0000
   End
   Begin PicClip.PictureClip Balls 
      Index           =   1
      Left            =   1920
      Top             =   4080
      _ExtentX        =   21167
      _ExtentY        =   1323
      _Version        =   393216
      Cols            =   16
      Picture         =   "FrmMain.frx":1D512
   End
   Begin PicClip.PictureClip Balls 
      Index           =   2
      Left            =   1920
      Top             =   4860
      _ExtentX        =   21167
      _ExtentY        =   1323
      _Version        =   393216
      Cols            =   16
      Picture         =   "FrmMain.frx":3AA24
   End
   Begin PicClip.PictureClip Balls 
      Index           =   3
      Left            =   1920
      Top             =   5640
      _ExtentX        =   21167
      _ExtentY        =   1323
      _Version        =   393216
      Cols            =   16
      Picture         =   "FrmMain.frx":57F36
   End
   Begin PicClip.PictureClip Balls 
      Index           =   4
      Left            =   1920
      Top             =   6420
      _ExtentX        =   21167
      _ExtentY        =   1323
      _Version        =   393216
      Cols            =   16
      Picture         =   "FrmMain.frx":75448
   End
   Begin VB.Label Score 
      BackColor       =   &H00000000&
      Caption         =   "Score = "
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1755
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Top             =   300
      Width           =   3915
   End
   Begin VB.Image Cell 
      Height          =   555
      Index           =   0
      Left            =   840
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cell_Click(Index As Integer)
    On Error Resume Next
    If AllowClick = False Then Exit Sub
    Dim Match As Single
    
    AllowClick = False
    
    'Get Arrey Coordinates From Index
    CellY = Int(Index / BoardX)
    GetCellX = Format(((Index / BoardX) - (CellY + 1)), "###.###") + 0.01
    CellX = Int((BoardX / 100) * ((GetCellX + 1) * 100))
    
    If Board(CellX, CellY, 0) = 5 Then
    AllowClick = True
    Exit Sub
    End If
    Dim LP As Single
    
For LP = 0 To Cell.ubound
    
    
    If Board(CellX, CellY, 0) < 5 Then
    
    If CellX < BoardX - 1 Then
        If Board(CellX + 1, CellY, 0) = Board(CellX, CellY, 0) Then Match = Match + 1
    End If
    If CellX > 0 Then
        If Board(CellX - 1, CellY, 0) = Board(CellX, CellY, 0) Then Match = Match + 1
    End If
    If CellY < BoardY - 1 Then
        If Board(CellX, CellY + 1, 0) = Board(CellX, CellY, 0) Then Match = Match + 1
    End If
    If CellY > 0 Then
        If Board(CellX, CellY - 1, 0) = Board(CellX, CellY, 0) Then Match = Match + 1
    End If
    End If
Next LP
    
    If Match = 0 Then
        MsgBox "No More Moves"
        AllowClick = True
        Exit Sub
    End If
    
    Match = 0
    If CellX < BoardX - 1 Then
        If Board(CellX + 1, CellY, 0) = Board(CellX, CellY, 0) Then Match = Match + 1
    End If
    If CellX > 0 Then
        If Board(CellX - 1, CellY, 0) = Board(CellX, CellY, 0) Then Match = Match + 1
    End If
    If CellY < BoardY - 1 Then
        If Board(CellX, CellY + 1, 0) = Board(CellX, CellY, 0) Then Match = Match + 1
    End If
    If CellY > 0 Then
        If Board(CellX, CellY - 1, 0) = Board(CellX, CellY, 0) Then Match = Match + 1
    End If
    
    If Match = 0 Then
        AllowClick = True
        Exit Sub
    End If
    
    BallColor = Board(CellX, CellY, 0)
    Board(CellX, CellY, 0) = 0
    
    FindConnecting CellX, CellY
    
End Sub

Private Sub FindConnecting(BX, BY)
On Error Resume Next

Dim Rep As Single
Dim X, Y As Single
Dim Connections As Single

For Rep = 0 To (BoardX * BoardY)
    For Y = 0 To BoardY - 1
        For X = 0 To BoardX - 1
            If Board(X, Y, 0) = 0 Then
                If Board(X - 1, Y, 0) = BallColor Then Board(X - 1, Y, 0) = 0
                If Board(X + 1, Y, 0) = BallColor Then Board(X + 1, Y, 0) = 0
                If Board(X, Y - 1, 0) = BallColor Then Board(X, Y - 1, 0) = 0
                If Board(X, Y + 1, 0) = BallColor Then Board(X, Y + 1, 0) = 0
            End If
        Next X
    Next Y
Next Rep


DoEvents

dropBallsCount = 5


End Sub


Private Sub RemoveBalls()

Dim BallCount As Single

Dim X, Y As Single

    For Y = 0 To BoardY - 1
        For X = 0 To BoardX - 1
            If Board(X, Y, 0) = 0 Then Board(X, Y, 0) = 5: BallCount = BallCount + 1
        Next X
    Next Y
    
    
        BallCount = BallCount * BallCount
    
    
    Score.Caption = Val(Score.Caption) + BallCount
    DoEvents
    
    DropBalls
    
End Sub

Private Sub DropBalls()
On Error Resume Next

Dim Rep, X, Y As Single

For Rep = 0 To BoardX * BoardY
    For Y = BoardY - 1 To 0 Step -1
        For X = 0 To BoardX - 1
            If Board(X, Y, 0) = 5 Then
               Board(X, Y, 0) = Board(X, Y - 1, 0)
               Board(X, Y - 1, 0) = 5
            End If
        Next X
    Next Y
Next Rep
    
    DoEvents
    AllowClick = True
    MoveLeft
End Sub

Private Sub MoveLeft()

Dim X, Y, BCount, Rep As Single

For Rep = 0 To BoardX * BoardY
For X = 0 To BoardX - 1
    BCount = 0
    For Y = 0 To BoardY - 1
        If Board(X, Y, 0) <> 5 Then BCount = BCount + 1
    Next Y
    If BCount = 0 Then
        For Y = 0 To BoardY - 1
            Board(X, Y, 0) = Board(X + 1, Y, 0)
            Board(X + 1, Y, 0) = 5
        Next Y
    End If
Next X
Next Rep




AllowClick = True
End Sub

Private Sub Cell_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)


    'Get Arrey Coordinates From Index
    CellY = Int(Index / BoardX)
    GetCellX = Format(((Index / BoardX) - (CellY + 1)), "###.###") + 0.01
    CellX = Int((BoardX / 100) * ((GetCellX + 1) * 100))
    
    'Get The Index From X And Y Coordinates
    GetIndex = ((CellY) * BoardX) + (CellX)
    
    Label1.Caption = "X=" & CellX & " Y=" & CellY & " Index=" & Index & " Val=" & Board(CellX, CellY, 0)
End Sub

Private Sub Form_Click()
End
End Sub

Private Sub Form_Load()
Randomize (Timer)
MaxFrames = Balls(0).Cols - 2
AllowClick = True
BuildBoard

End Sub

Private Sub BuildBoard()
On Error Resume Next

Dim GridY As Single
Dim GridX As Single

Dim X As Single
Dim Y As Single

Dim OffX As Single
Dim OffY As Single

'Populate Board Image Arrey
For Y = 0 To BoardY - 1
    For X = 0 To BoardX - 1
        Board(X, Y, 0) = Int(Rnd * 3) + 1
        Select Case Board(X, Y, 0)
            Case 1
                Board(X, Y, 1) = 0
            Case 2
                Board(X, Y, 1) = 4
            Case 3
                Board(X, Y, 1) = 8
            Case 4
                Board(X, Y, 1) = 12
        End Select
    Next X
Next Y

Cell(0).Picture = Balls(0).GraphicCell(0)

'Get Centre Of Screen
OffX = ((Screen.Width / 15) - (BoardX * 50)) / 2
OffY = ((Screen.Height / 15) - (BoardY * 50)) / 2


'Create Board
For GridY = 0 To BoardY - 1
    For GridX = 0 To BoardX - 1
        Load Cell(GridX + (GridY * BoardX))
            Cell(GridX + (GridY * BoardX)).Left = OffX + (GridX * 50)
            Cell((GridX + (GridY * BoardX))).Top = OffY + (GridY * 50)
            Cell((GridX + (GridY * BoardX))).Visible = True
        Next GridX
    Next GridY
End Sub

Private Sub Timer1_Timer()

Dim Y As Single
Dim X As Single


For Y = 0 To BoardY - 1
    For X = 0 To BoardX - 1
        If Board(X, Y, 0) <> 5 Then
            Board(X, Y, 1) = Board(X, Y, 1) + 1
            If Board(X, Y, 1) > MaxFrames Then Board(X, Y, 1) = 0
            Cell((Y * BoardX) + X).Picture = Balls(Board(X, Y, 0)).GraphicCell(Board(X, Y, 1))
        Else
            GetIndex = (Y * BoardX) + X
            Cell(GetIndex).Picture = LoadPicture("")
        End If
    Next X
    DoEvents
Next Y

If dropBallsCount > 0 Then
    dropBallsCount = dropBallsCount - 1
Else
    RemoveBalls
End If
End Sub
