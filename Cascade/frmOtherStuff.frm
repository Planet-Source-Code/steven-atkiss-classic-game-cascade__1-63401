VERSION 5.00
Begin VB.Form frmOtherStuff 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18000
   BeginProperty Font 
      Name            =   "Kristen ITC"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10290
   ScaleWidth      =   18000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   8175
      Index           =   1
      Left            =   1320
      TabIndex        =   32
      Top             =   1140
      Width           =   10635
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FF80&
         Caption         =   "More Help?"
         Height          =   615
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FF80&
         Caption         =   "Done"
         Height          =   615
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   7320
         Width           =   1275
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Oh Yeah... Never Tell Your Partner ""It's OK, I Like Big Boned People!?!!?"
         ForeColor       =   &H0080FF80&
         Height          =   915
         Index           =   5
         Left            =   780
         TabIndex        =   38
         Top             =   6780
         Width           =   7815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "Erm...   Thats It. Get As Higher Score By Eliminating As Many Balls In One Go."
         ForeColor       =   &H0080FF80&
         Height          =   915
         Index           =   4
         Left            =   600
         TabIndex        =   37
         Top             =   5520
         Width           =   6975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   $"frmOtherStuff.frx":0000
         ForeColor       =   &H0080FF80&
         Height          =   1695
         Index           =   3
         Left            =   1680
         TabIndex        =   36
         Top             =   3600
         Width           =   8595
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "3. The Bigger The Number Of Balls Connected In One Whammy, The Higher Your Score Goes."
         ForeColor       =   &H0080FF80&
         Height          =   915
         Index           =   2
         Left            =   360
         TabIndex        =   35
         Top             =   2580
         Width           =   7875
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "2. Clicking On A Ball Will Also Remove Any Ball Sequence Connected To It Horizontally And Vertically."
         ForeColor       =   &H0080FF80&
         Height          =   915
         Index           =   1
         Left            =   900
         TabIndex        =   34
         Top             =   1440
         Width           =   9615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "1. It's Not Quite As Easy As Just Getting Rid Of All Ya Balls."
         ForeColor       =   &H0080FF80&
         Height          =   915
         Index           =   0
         Left            =   300
         TabIndex        =   33
         Top             =   420
         Width           =   6975
      End
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   8175
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   1020
      Width           =   10635
      Begin VB.TextBox TxtEnterName 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H0000FFFF&
         Height          =   435
         Left            =   2760
         TabIndex        =   41
         Top             =   6780
         Visible         =   0   'False
         Width           =   5355
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "Done"
         Height          =   615
         Left            =   9120
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   7320
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Top DOG Players Of CasCade"
         BeginProperty Font 
            Name            =   "Kristen ITC"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   795
         Left            =   780
         TabIndex        =   31
         Top             =   300
         Width           =   9075
      End
      Begin VB.Label TxtName 
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FF80&
         Height          =   435
         Index           =   0
         Left            =   2760
         TabIndex        =   30
         Top             =   1380
         Width           =   5355
      End
      Begin VB.Label TxtName 
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FF80&
         Height          =   435
         Index           =   1
         Left            =   2760
         TabIndex        =   29
         Top             =   1920
         Width           =   5355
      End
      Begin VB.Label TxtName 
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FF80&
         Height          =   435
         Index           =   2
         Left            =   2760
         TabIndex        =   28
         Top             =   2460
         Width           =   5355
      End
      Begin VB.Label TxtName 
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FF80&
         Height          =   435
         Index           =   3
         Left            =   2760
         TabIndex        =   27
         Top             =   3000
         Width           =   5355
      End
      Begin VB.Label TxtName 
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FF80&
         Height          =   435
         Index           =   4
         Left            =   2760
         TabIndex        =   26
         Top             =   3540
         Width           =   5355
      End
      Begin VB.Label TxtName 
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FF80&
         Height          =   435
         Index           =   5
         Left            =   2760
         TabIndex        =   25
         Top             =   4080
         Width           =   5355
      End
      Begin VB.Label TxtName 
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FF80&
         Height          =   435
         Index           =   6
         Left            =   2760
         TabIndex        =   24
         Top             =   4620
         Width           =   5355
      End
      Begin VB.Label TxtName 
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FF80&
         Height          =   435
         Index           =   7
         Left            =   2760
         TabIndex        =   23
         Top             =   5160
         Width           =   5355
      End
      Begin VB.Label TxtName 
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FF80&
         Height          =   435
         Index           =   8
         Left            =   2760
         TabIndex        =   22
         Top             =   5700
         Width           =   5355
      End
      Begin VB.Label TxtName 
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FF80&
         Height          =   435
         Index           =   9
         Left            =   2760
         TabIndex        =   21
         Top             =   6240
         Width           =   5355
      End
      Begin VB.Label Score 
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FF80&
         Height          =   435
         Index           =   0
         Left            =   8220
         TabIndex        =   20
         Top             =   1380
         Width           =   1815
      End
      Begin VB.Label Score 
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FF80&
         Height          =   435
         Index           =   1
         Left            =   8220
         TabIndex        =   19
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Score 
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FF80&
         Height          =   435
         Index           =   2
         Left            =   8220
         TabIndex        =   18
         Top             =   2460
         Width           =   1815
      End
      Begin VB.Label Score 
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FF80&
         Height          =   435
         Index           =   3
         Left            =   8220
         TabIndex        =   17
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Score 
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FF80&
         Height          =   435
         Index           =   4
         Left            =   8220
         TabIndex        =   16
         Top             =   3540
         Width           =   1815
      End
      Begin VB.Label Score 
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FF80&
         Height          =   435
         Index           =   5
         Left            =   8220
         TabIndex        =   15
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Score 
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FF80&
         Height          =   435
         Index           =   6
         Left            =   8220
         TabIndex        =   14
         Top             =   4620
         Width           =   1815
      End
      Begin VB.Label Score 
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FF80&
         Height          =   435
         Index           =   7
         Left            =   8220
         TabIndex        =   13
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Label Score 
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FF80&
         Height          =   435
         Index           =   8
         Left            =   8220
         TabIndex        =   12
         Top             =   5700
         Width           =   1815
      End
      Begin VB.Label Score 
         BackColor       =   &H00000000&
         ForeColor       =   &H0080FF80&
         Height          =   435
         Index           =   9
         Left            =   8220
         TabIndex        =   11
         Top             =   6240
         Width           =   1815
      End
      Begin VB.Label Position 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Top Dog, 1 - "
         ForeColor       =   &H0080FF80&
         Height          =   375
         Index           =   0
         Left            =   540
         TabIndex        =   10
         Top             =   1380
         Width           =   2115
      End
      Begin VB.Label Position 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Under Dog, 2 - "
         ForeColor       =   &H0080FF80&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label Position 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Puppy, 3 -"
         ForeColor       =   &H0080FF80&
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   8
         Top             =   2460
         Width           =   1935
      End
      Begin VB.Label Position 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "4 -"
         ForeColor       =   &H0080FF80&
         Height          =   375
         Index           =   3
         Left            =   720
         TabIndex        =   7
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Position 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Joe Average, 5 - "
         ForeColor       =   &H0080FF80&
         Height          =   375
         Index           =   4
         Left            =   -60
         TabIndex        =   6
         Top             =   3540
         Width           =   2715
      End
      Begin VB.Label Position 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "6 -"
         ForeColor       =   &H0080FF80&
         Height          =   375
         Index           =   5
         Left            =   1920
         TabIndex        =   5
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label Position 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "7 -"
         ForeColor       =   &H0080FF80&
         Height          =   375
         Index           =   6
         Left            =   2160
         TabIndex        =   4
         Top             =   4620
         Width           =   495
      End
      Begin VB.Label Position 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "8 -"
         ForeColor       =   &H0080FF80&
         Height          =   375
         Index           =   7
         Left            =   2160
         TabIndex        =   3
         Top             =   5160
         Width           =   495
      End
      Begin VB.Label Position 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "9 -"
         ForeColor       =   &H0080FF80&
         Height          =   375
         Index           =   8
         Left            =   2040
         TabIndex        =   2
         Top             =   5700
         Width           =   615
      End
      Begin VB.Label Position 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Last, 10 -"
         ForeColor       =   &H0080FF80&
         Height          =   375
         Index           =   9
         Left            =   1080
         TabIndex        =   1
         Top             =   6240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmOtherStuff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
If DoSave = True Then
    If Trim(TxtEnterName.Text) = "" Then
        MsgBox ("Go On, Enter Your Name")
        Exit Sub
    End If
    DoSave = False
    
    HighScores(YourPos, 0) = StrConv(TxtEnterName.Text, vbProperCase)
    TxtEnterName.Text = ""
    TxtEnterName.Visible = False
    SaveScores
End If
Me.Hide
FrmMenu.Show
End Sub

Private Sub Command2_Click()
Me.Hide
FrmMenu.Show
End Sub

Private Sub Command3_Click()
FrmHelp.Show
End Sub

Private Sub Form_Load()
    Me.Height = Screen.Height - 4000
    Me.Width = Screen.Width - 4000
    
    Frame(0).Top = (Me.Height / 2) - (Frame(0).Height / 2)
    Frame(0).Left = (Me.Width / 2) - (Frame(0).Width / 2)
    Frame(1).Top = (Me.Height / 2) - (Frame(1).Height / 2)
    Frame(1).Left = (Me.Width / 2) - (Frame(1).Width / 2)
End Sub

