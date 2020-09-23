VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form FrmMenu 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   14040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   936
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   75
      Left            =   1320
      Top             =   3840
   End
   Begin PicClip.PictureClip Balls 
      Index           =   0
      Left            =   2820
      Top             =   10980
      _ExtentX        =   21167
      _ExtentY        =   1323
      _Version        =   393216
      Cols            =   16
      Picture         =   "FrmMenu.frx":0000
   End
   Begin PicClip.PictureClip Balls 
      Index           =   1
      Left            =   2940
      Top             =   11880
      _ExtentX        =   21167
      _ExtentY        =   1323
      _Version        =   393216
      Cols            =   16
      Picture         =   "FrmMenu.frx":1D512
   End
   Begin PicClip.PictureClip Balls 
      Index           =   2
      Left            =   3120
      Top             =   12660
      _ExtentX        =   21167
      _ExtentY        =   1323
      _Version        =   393216
      Cols            =   16
      Picture         =   "FrmMenu.frx":3AA24
   End
   Begin PicClip.PictureClip Balls 
      Index           =   3
      Left            =   3360
      Top             =   13560
      _ExtentX        =   21167
      _ExtentY        =   1323
      _Version        =   393216
      Cols            =   16
      Picture         =   "FrmMenu.frx":57F36
   End
   Begin VB.Image Cell 
      Height          =   735
      Index           =   0
      Left            =   480
      Top             =   420
      Width           =   735
   End
   Begin VB.Label Menu 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "How Dya Play It?"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   765
      Index           =   0
      Left            =   4560
      TabIndex        =   3
      Top             =   3300
      Width           =   4800
   End
   Begin VB.Label Menu 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "PLAY CASCADE"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   765
      Index           =   3
      Left            =   8100
      TabIndex        =   2
      Top             =   6900
      Width           =   5175
   End
   Begin VB.Label Menu 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Setup How I Wanna PLay"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   765
      Index           =   2
      Left            =   3420
      TabIndex        =   1
      Top             =   5580
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Menu 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Past Gods Of Cascade"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   765
      Index           =   1
      Left            =   5760
      TabIndex        =   0
      Top             =   4560
      Width           =   6525
   End
End
Attribute VB_Name = "FrmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private CircusLights(82) As Single
Private Frame As Single
Private XImage As Single
Private YImage As Single
Private Change As Single
Private ChCount As Single
Private Max As Single



Private Sub Form_Click()
    End
End Sub

Private Sub Form_Load()
    
    Dim FF As Long

    Randomize (Timer)
    
    
    If FileExists(App.Path & "\Scores.Txt") = 0 Then
        FF = FreeFile
        Open App.Path & "\Scores.Txt" For Output As #FF
            Print #FF, "God Like Chap|300000"
            Print #FF, "Knob The Builder|200000"
            Print #FF, "Darkwing Duck|100000"
            Print #FF, "Toilet Duck|50000"
            Print #FF, "A Peanut I Found|45000"
            Print #FF, "An Inanimate Object|30000"
            Print #FF, "College IT Tech|25000"
            Print #FF, "High School IT Teacher|20000"
            Print #FF, "Some  Bloke Off The Street|15000"
            Print #FF, "My Plant|10000"
        Close #FF
    End If
    
    Cell(0).Picture = Balls(0).GraphicCell(4)

    MakeTheFlashyThings
    
    Menu(0).Top = (Screen.Height / 15) / 4
    Menu(1).Top = (Screen.Height / 15) / 3
    Menu(2).Top = (Screen.Height / 15) / 2
    Menu(3).Top = ((Screen.Height / 15) / 4) * 3
    
    For Each Control In Me
        If TypeOf Control Is Label Then
            Control.ForeColor = RGB(28, 155, 28)
        End If
    Next Control
    
    RandomLights
    
End Sub



Private Sub MakeTheFlashyThings()

    Dim MakeNew As Single
    Dim Lp As Single
    
    XImage = Int(((Screen.Width / 15) - 100) / 50)
    YImage = Int(((Screen.Height / 15) - 100) / 50)
    
    For MakeNew = 1 To (XImage * 2) + (YImage * 2)
        Load Cell(MakeNew)
    Next MakeNew
    
    For MakeNew = 0 To XImage - 1
        Cell(MakeNew).Top = 50
        Cell(MakeNew).Left = 50 + (MakeNew * 50)
        Cell(MakeNew).Visible = True
    Next MakeNew
    
    For MakeNew = 0 To YImage - 1
        Cell(MakeNew + XImage).Top = 50 + (MakeNew * 50)
        Cell(MakeNew + XImage).Left = (XImage * 50)
        Cell(MakeNew + XImage).Visible = True
    Next MakeNew
    
    For MakeNew = 0 To XImage - 1
        Cell(MakeNew + XImage + YImage).Top = (YImage * 50)
        Cell(MakeNew + XImage + YImage).Left = (XImage * 50) - (MakeNew * 50)
        Cell(MakeNew + XImage + YImage).Visible = True
    Next MakeNew
    
    For MakeNew = 0 To YImage - 1
        Cell(MakeNew + (XImage * 2) + YImage).Top = (YImage * 50) - (MakeNew * 50)
        Cell(MakeNew + (XImage * 2) + YImage).Left = 50
        Cell(MakeNew + (XImage * 2) + YImage).Visible = True
    Next MakeNew
        
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Menu(0).ForeColor = RGB(28, 155, 28)
    Menu(1).ForeColor = RGB(28, 155, 28)
    Menu(2).ForeColor = RGB(28, 155, 28)
    Menu(3).ForeColor = RGB(28, 155, 28)
    
End Sub

Private Sub Menu_Click(Index As Integer)

Dim FF As Long
Dim SplitString() As String
Dim InString As String
Dim Pos As Single

Select Case Index
Case 0
    frmOtherStuff.Frame(1).ZOrder (0)
    frmOtherStuff.Show
Case 1
    FF = FreeFile
    Open App.Path & "\Scores.Txt" For Input As #FF
        Do
            
            Line Input #FF, InString
            SplitString = Split(InString, "|")
            frmOtherStuff.TxtName(Pos).Caption = SplitString(0)
            frmOtherStuff.Score(Pos).Caption = SplitString(1)
            Pos = Pos + 1
        Loop Until EOF(FF)
    Close #FF
    frmOtherStuff.Frame(0).ZOrder (0)
    frmOtherStuff.Show
Case 3
    Me.Hide
    FrmCascade.Show
End Select


End Sub

Private Sub Menu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Select Case Index
    Case 0
        Menu(Index).ForeColor = RGB(128, 255, 128)
        Menu(1).ForeColor = RGB(28, 155, 28)
        Menu(2).ForeColor = RGB(28, 155, 28)
        Menu(3).ForeColor = RGB(28, 155, 28)
    Case 1
        Menu(Index).ForeColor = RGB(128, 255, 128)
        Menu(0).ForeColor = RGB(28, 155, 28)
        Menu(2).ForeColor = RGB(28, 155, 28)
        Menu(3).ForeColor = RGB(28, 155, 28)
    Case 2
        Menu(Index).ForeColor = RGB(128, 255, 128)
        Menu(0).ForeColor = RGB(28, 155, 28)
        Menu(1).ForeColor = RGB(28, 155, 28)
        Menu(3).ForeColor = RGB(28, 155, 28)
    Case 3
        Menu(Index).ForeColor = RGB(128, 255, 128)
        Menu(0).ForeColor = RGB(28, 155, 28)
        Menu(1).ForeColor = RGB(28, 155, 28)
        Menu(2).ForeColor = RGB(28, 155, 28)
    End Select

End Sub

Private Sub Timer1_Timer()

    For Lp = 0 To Cell.UBound - 1
        CircusLights(Lp) = CircusLights(Lp + 1)
    Next Lp
    
    Frame = Frame + 1
    If Frame > 14 Then Frame = 0
    
    
        For Lp = 0 To Cell.UBound - 1
            CircusLights(Lp) = CircusLights(Lp + 1)
            Cell(Lp).Picture = Balls(CircusLights(Lp)).GraphicCell(Frame)
        Next Lp
        
        CircusLights(Cell.UBound) = CircusLights(0)
        Cell(Cell.UBound).Picture = Balls(CircusLights(Lp)).GraphicCell(Frame)
    
    Dim OChange As Single
    ChCount = ChCount - 1
    If ChCount < 0 Then
        OChange = Change
        Change = Int(Rnd * 3)
        If Change = OChange Then ChCount = 50: Exit Sub
        Select Case Change
        Case 0
            QuatroColour
        Case 1
            DoubleUp
        Case 2
            RandomLights
        End Select
    End If
    
End Sub



Private Sub QuatroColour()

    Dim Lp As Single
    
    For Lp = 0 To XImage
        CircusLights(Lp) = 0
        'Cell(Lp).Picture = Balls(CircusLights(Lp)).GraphicCell(0)
    Next Lp
    
    For Lp = XImage To XImage + YImage
        CircusLights(Lp) = 1
        'Cell(Lp).Picture = Balls(CircusLights(Lp)).GraphicCell(0)
    Next Lp
    
    For Lp = XImage + YImage To (XImage * 2) + YImage
        CircusLights(Lp) = 2
        'Cell(Lp).Picture = Balls(CircusLights(Lp)).GraphicCell(0)
    Next Lp
    
    For Lp = (XImage * 2) + YImage To (XImage * 2) + (YImage * 2)
        CircusLights(Lp) = 3
        'Cell(Lp).Picture = Balls(CircusLights(Lp)).GraphicCell(0)
    Next Lp
    
    ChCount = 50
    
End Sub


Private Sub DoubleUp()

    Dim Lp, Col1, Col2 As Single
    
    Col1 = Int(Rnd * 4)
OtherColour:
    Col2 = Int(Rnd * 4)
    If Col1 = Col2 Then GoTo OtherColour
    
    For Lp = 0 To Cell.UBound
        CircusLights(Lp) = Col1
    Next Lp
    
    For Lp = 0 To Cell.UBound Step 5
        CircusLights(Lp) = Col2
    Next Lp
    
    ChCount = 50
    
End Sub



Private Sub RandomLights()

    For Lp = 0 To Cell.UBound
        CircusLights(Lp) = Int(Rnd * 4)
    Next Lp
    
    ChCount = 50
    
End Sub
