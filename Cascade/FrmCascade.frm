VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form FrmCascade 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11685
   BeginProperty Font 
      Name            =   "Kristen ITC"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   595
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   779
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Animate My Balls"
      ForeColor       =   &H00FF8080&
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   3540
      Width           =   3375
   End
   Begin VB.Timer Timer1 
      Interval        =   75
      Left            =   120
      Top             =   2520
   End
   Begin PicClip.PictureClip Balls 
      Index           =   0
      Left            =   540
      Top             =   4980
      _ExtentX        =   21167
      _ExtentY        =   1323
      _Version        =   393216
      Cols            =   16
      Picture         =   "FrmCascade.frx":0000
   End
   Begin PicClip.PictureClip Balls 
      Index           =   1
      Left            =   660
      Top             =   5880
      _ExtentX        =   21167
      _ExtentY        =   1323
      _Version        =   393216
      Cols            =   16
      Picture         =   "FrmCascade.frx":1D512
   End
   Begin PicClip.PictureClip Balls 
      Index           =   2
      Left            =   840
      Top             =   6660
      _ExtentX        =   21167
      _ExtentY        =   1323
      _Version        =   393216
      Cols            =   16
      Picture         =   "FrmCascade.frx":3AA24
   End
   Begin PicClip.PictureClip Balls 
      Index           =   3
      Left            =   1080
      Top             =   7560
      _ExtentX        =   21167
      _ExtentY        =   1323
      _Version        =   393216
      Cols            =   16
      Picture         =   "FrmCascade.frx":57F36
   End
   Begin VB.Label Score 
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   645
      Left            =   2220
      TabIndex        =   1
      Top             =   120
      Width           =   2955
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Score ="
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   660
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   1890
   End
   Begin VB.Image Cell 
      Height          =   675
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "FrmCascade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    
    'Get Central Position
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

Private Sub Cell_Click(Index As Integer)

    If AllowClick = False Then Exit Sub
    
    'Get Arrey Coordinates From Index
    CellY = Int(Index / BoardX)
    GetCellX = Format(((Index / BoardX) - (CellY + 1)), "###.###") + 0.01
    CellX = Int((BoardX / 100) * ((GetCellX + 1) * 100))
    
    If Board(CellX, CellY, 0) = 5 Then Exit Sub
    
    AllowClick = False
    
    GetConnected CellX, CellY, Int(Board(CellX, CellY, 0))
End Sub




Private Sub Check1_Click()
Animate = Check1.Value
End Sub

Private Sub Form_Activate()
    AllowClick = True
    AllowLeft = True
    AllowMessage = False
    Score.Caption = 0
    BuildBoard
    MaxFrames = 14
End Sub

Private Sub Form_Load()

    AllowClick = True
    BuildBoard
    MaxFrames = 14
    Check1.Left = 20
    Check1.Top = (Screen.Height / 15) - Check1.Height - 20
    Animate = Check1.Value
    
End Sub

Private Sub Timer1_Timer()
    
    Dim Y As Single
    Dim X As Single
    
    'Animate The Balls
    For Y = 0 To BoardY - 1
        For X = 0 To BoardX - 1
        
            If Board(X, Y, 0) < 5 And Y < BoardY - 1 Then
                If Board(X, Y + 1, 0) = 5 Then
                    AllowLeft = False
                    Board(X, Y + 1, 0) = Board(X, Y, 0)
                    Board(X, Y, 0) = 5
                End If
            End If
                      
            If AllowLeft = True Then
                MoveEmLeft
            End If
            
            If Board(X, Y, 0) < 5 Then
                If Animate = True Then
                    Board(X, Y, 1) = Board(X, Y, 1) + 1
                    If Board(X, Y, 1) > MaxFrames Then Board(X, Y, 1) = 0
                Else
                    Board(X, Y, 1) = 4
                End If
                Cell((Y * BoardX) + X).Picture = Balls(Board(X, Y, 0)).GraphicCell(Board(X, Y, 1))
            Else
                GetIndex = (Y * BoardX) + X
                Cell(GetIndex).Picture = LoadPicture("")
            End If
        Next X
    Next Y
    
    If AllowLeft = True Then CheckEnd
    
    AllowLeft = True
    
    DisplayGrey = DisplayGrey - 1
    If DisplayGrey < 0 Then DisplayGrey = 0
    
    If DisplayGrey = 1 Then
        ClearYaWhammy
        ElseIf DisplayGrey > 1 Then AllowClick = False
    End If
    
    
    
    
End Sub
