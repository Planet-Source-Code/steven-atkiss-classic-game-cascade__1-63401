VERSION 5.00
Begin VB.Form FrmHelp 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   12720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Click Screen To Close."
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
      Height          =   315
      Left            =   300
      TabIndex        =   0
      Top             =   180
      Width           =   3435
   End
   Begin VB.Image ImgHelp 
      Enabled         =   0   'False
      Height          =   11520
      Left            =   0
      Picture         =   "FrmHelp.frx":0000
      Top             =   960
      Width           =   15360
   End
End
Attribute VB_Name = "FrmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
Me.Hide
    FrmMenu.Show
    frmOtherStuff.Show
End Sub

Private Sub Form_Load()
    ImgHelp.Top = (Screen.Height / 2) - (ImgHelp.Height / 2)
    ImgHelp.Left = (Screen.Width / 2) - (ImgHelp.Width / 2)
    
End Sub
