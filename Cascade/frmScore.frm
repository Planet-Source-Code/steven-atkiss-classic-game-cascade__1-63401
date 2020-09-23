VERSION 5.00
Begin VB.Form frmScore 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13260
   BeginProperty Font 
      Name            =   "Kristen ITC"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11115
   ScaleWidth      =   13260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   360
      Width           =   1275
   End
   Begin VB.Label EndMessage 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FF80&
      Height          =   3075
      Left            =   480
      TabIndex        =   2
      Top             =   4080
      Width           =   11235
   End
   Begin VB.Label Score 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FF80&
      Height          =   735
      Left            =   1320
      TabIndex        =   1
      Top             =   2760
      Width           =   8235
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Well, Ya Got A Grand Total Of:"
      ForeColor       =   &H0080FF80&
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   1980
      Width           =   8235
   End
End
Attribute VB_Name = "frmScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label2_Click()
    Me.Hide
    FrmCascade.Hide
    frmOtherStuff.Frame(0).ZOrder (0)
    
    frmOtherStuff.Show
    If DoSave = True Then
        frmOtherStuff.TxtEnterName.SetFocus
    End If
    
End Sub

Private Sub Form_Click()
    Me.Hide
    FrmCascade.Hide
    frmOtherStuff.Frame(0).ZOrder (0)
    
    frmOtherStuff.Show
    If DoSave = True Then
        frmOtherStuff.TxtEnterName.SetFocus
    End If
    
End Sub

Private Sub Form_Load()

    Label1.Top = Screen.Height / 3
    Score.Top = Label1.Top + Label1.Height
    EndMessage.Top = (Screen.Height / 3) * 2
    
    Label1.Left = (Screen.Width / 2) - (Label1.Width / 2)
    Score.Left = (Screen.Width / 2) - (Score.Width / 2)
    EndMessage.Left = (Screen.Width / 2) - (EndMessage.Width / 2)
    
End Sub
