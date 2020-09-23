Attribute VB_Name = "Module1"
Option Explicit
Public Const BoardX = 15
Public Const BoardY = 12
Public Board(BoardX - 1, BoardY - 1, 1)
Public AllowClick, AllowLeft As Boolean
Public MaxFrames As Single
Public Frame As Single
Public CellX As Double
Public CellY As Double
Public GetCellX As Double
Public GetIndex As Long
Public Animate As Boolean
Public DisplayGrey As Single
Public AllowMessage As Boolean
Public HighScores(10, 1)
Public YourPos As Single
Public DoSave As Boolean

Public Function FileExists(strPath As String) As Integer
    FileExists = Not (Dir(strPath) = "")
End Function


Public Sub CheckEnd()
    Dim BallX, BallY, Connection, Balls As Single
    
    Connection = 0
    
    For BallY = 0 To BoardY - 1
        For BallX = 0 To BoardX - 1
            If Board(BallX, BallY, 0) < 5 Then
                Balls = Balls + 1
                If BallX < BoardX - 1 Then
                    If Board(BallX + 1, BallY, 0) = Board(BallX, BallY, 0) Then Connection = 1
                End If
        
                If BallX > 0 Then
                    If Board(BallX - 1, BallY, 0) = Board(BallX, BallY, 0) Then Connection = 1
                End If
        
                If BallY > 0 Then
                    If Board(BallX, BallY - 1, 0) = Board(BallX, BallY, 0) Then Connection = 1
                End If
        
                If BallY < BoardY - 1 Then
                    If Board(BallX, BallY + 1, 0) = Board(BallX, BallY, 0) Then Connection = 1
                End If
            End If
            If Connection = 1 Then Exit Sub
            DoEvents
        Next BallX
    Next BallY


    
    If Connection = 0 And AllowMessage = False Then
        AllowMessage = True
        If Balls = 0 Then
            FrmCascade.Score.Caption = Val(FrmCascade.Score.Caption) + 10000
            MsgBox "You Cleared The Board, Bonus 10,000 Points."
            LoadHighScores True, Val(FrmCascade.Score.Caption)
            Exit Sub
        Else
            MsgBox "Your Outa There..."
            LoadHighScores False, Val(FrmCascade.Score.Caption)
        End If
    End If
    
End Sub
Public Sub LoadHighScores(Bonus As Boolean, Score As Double)

    Dim FF As Long
    Dim Pos, rep As Single
    Dim SplitString() As String
    Dim InString As String
    
    Pos = 0
    
    FF = FreeFile
    Open App.Path & "\Scores.Txt" For Input As #FF
        Do
            Line Input #FF, InString
            SplitString = Split(InString, "|")
            HighScores(Pos, 0) = SplitString(0)
            HighScores(Pos, 1) = SplitString(1)
            Pos = Pos + 1
        Loop Until EOF(FF)
    Close #FF
    
    If Score < Val(HighScores(9, 1)) Then
        frmScore.Score = Score
        frmScore.EndMessage.Caption = "Well Done, You Didn't Even Get On The Score Board??!!??"
        For rep = 0 To 9
            frmOtherStuff.TxtName(rep).Caption = HighScores(rep, 0)
            frmOtherStuff.Score(rep).Caption = HighScores(rep, 1)
        Next rep
        FrmMenu.Show
        frmScore.Show
        Exit Sub
    Else
        HighScores(9, 0) = ""
        HighScores(9, 1) = Score
        frmScore.Score = Score
    End If
    
    For rep = 9 To 0 Step -1
        If Score > Val(HighScores(rep, 1)) Then
            YourPos = rep
        End If
    Next rep

    For rep = 1 To 10
        For Pos = 9 To 1 Step -1
            If Val(HighScores(Pos - 1, 1)) < Val(HighScores(Pos, 1)) Then
                
                HighScores(10, 0) = HighScores(Pos - 1, 0)
                HighScores(10, 1) = HighScores(Pos - 1, 1)
                
                HighScores(Pos - 1, 0) = HighScores(Pos, 0)
                HighScores(Pos - 1, 1) = HighScores(Pos, 1)
                
                HighScores(Pos, 0) = HighScores(10, 0)
                HighScores(Pos, 1) = HighScores(10, 1)
            End If
        Next Pos
    Next rep
    
    Select Case YourPos
    Case 0
        frmScore.EndMessage.Caption = "Daaaamn, You Made 1st Place Well Done."
    Case 1
        frmScore.EndMessage.Caption = "Excellant Effort You Made 2nd Place. "
    Case 2
        frmScore.EndMessage.Caption = "Well Done, 3rd Place."
    Case 3
        frmScore.EndMessage.Caption = "Good Effort, 4th Place."
    Case 4
        frmScore.EndMessage.Caption = "Don't Suppose Your An Average Sort Of Person? You Made 5th."
    Case 5
        frmScore.EndMessage.Caption = "HHHhhhhhm 6th Place, At Least Your On The Board."
    Case 6
        frmScore.EndMessage.Caption = "A Bit Below Average, Is That Normal For You? 7th Place."
    Case 7
        frmScore.EndMessage.Caption = "Guess You'd Better Try Again, I Mean 8th Place, C'Mon."
    Case 8
        frmScore.EndMessage.Caption = "Your Not An IT Teacher Are You? 9th Place."
    Case 9
        frmScore.EndMessage.Caption = "Bet You Always get That Hanging On By Your Teeth Feeling, 10th."
    End Select
    
    For rep = 0 To 9
        frmOtherStuff.TxtName(rep).Caption = HighScores(rep, 0)
        frmOtherStuff.Score(rep).Caption = HighScores(rep, 1)
    Next rep
    
    
    
    DoSave = True
    frmOtherStuff.TxtEnterName.Top = frmOtherStuff.TxtName(YourPos).Top
    frmOtherStuff.TxtEnterName.Left = frmOtherStuff.TxtName(YourPos).Left
    frmOtherStuff.TxtEnterName.Visible = True
    FrmMenu.Show
    frmScore.Show

End Sub


Public Sub SaveScores()

    Dim FF As Long
    Dim rep As Single
    
    FF = FreeFile
    Open App.Path & "\Scores.Txt" For Output As #FF
        For rep = 0 To 9
            Print #FF, HighScores(rep, 0) & "|" & HighScores(rep, 1)
        Next rep
    Close #FF
    
End Sub

Public Sub GetConnected(BallX As Double, BallY As Double, BColour As Single)

    Dim Connection, X, Y, rep As Single
    
    Connection = 0
    
    'Check for Lonesome Ball
    If BallX < BoardX - 1 Then
        If Board(BallX + 1, BallY, 0) = Board(BallX, BallY, 0) Then Connection = 1
    End If
    
    If BallX > 0 Then
        If Board(BallX - 1, BallY, 0) = Board(BallX, BallY, 0) Then Connection = 1
    End If
    
    If BallY > 0 Then
    If Board(BallX, BallY - 1, 0) = Board(BallX, BallY, 0) Then Connection = 1
    End If
    
    If BallY < BoardY - 1 Then
    If Board(BallX, BallY + 1, 0) = Board(BallX, BallY, 0) Then Connection = 1
    End If
    
    'It Is A Lonesome Ball Aaaawwwwww
    If Connection = 0 Then
        AllowClick = True
        Exit Sub
    End If
    
    'Change Point Ball To A Grey Ball
    Board(BallX, BallY, 0) = 0
    
    'Find The Number Of Balls In Ya Whammy
    For rep = 0 To (BoardX * BoardY)
        For Y = 0 To BoardY - 1
            For X = 0 To BoardX - 1
                'Convert Connected Balls To Grey Balls
                If Board(X, Y, 0) = 0 Then
                    If X < BoardX - 1 Then
                        If Board(X + 1, Y, 0) = BColour Then Board(X + 1, Y, 0) = 0
                    End If
            
                    If X > 0 Then
                        If Board(X - 1, Y, 0) = BColour Then Board(X - 1, Y, 0) = 0
                    End If
                    
                    If Y > 0 Then
                        If Board(X, Y - 1, 0) = BColour Then Board(X, Y - 1, 0) = 0
                    End If
                    
                    If Y < BoardY - 1 Then
                        If Board(X, Y + 1, 0) = BColour Then Board(X, Y + 1, 0) = 0
                    End If
                End If
                
            Next X
        Next Y
    Next rep

    DisplayGrey = 6
    
End Sub

Public Sub ClearYaWhammy()

    Dim X, Y, Whammy As Single
    
    For Y = 0 To BoardY - 1
        For X = 0 To BoardX - 1
            If Board(X, Y, 0) = 0 Then
                Board(X, Y, 0) = 5
                Whammy = Whammy + 1
            End If
            
        Next X
    Next Y
    
    FrmCascade.Score.Caption = Val(FrmCascade.Score.Caption) + ((Whammy * Whammy) * Whammy)
    
    AllowClick = True
End Sub

Public Sub MoveEmLeft()

    Dim X, Y, X2, Y2, Match, rep As Single


    For X = 0 To BoardX - 2
        Match = 0
        For Y = 0 To BoardY - 1
            If Board(X, Y, 0) < 5 Then Match = 1
        Next Y
        If Match = 0 Then
            AllowClick = False
            For X2 = X To BoardX - 2
                For Y2 = 0 To BoardY - 1
                    Board(X2, Y2, 0) = Board(X2 + 1, Y2, 0)
                    Board(X2 + 1, Y2, 0) = 5
                Next Y2
            Next X2
            AllowClick = True
            Exit Sub
        End If
        
    Next X


End Sub
