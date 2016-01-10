Attribute VB_Name = "main"
Option Explicit


'***************************************************************
'***************************************************************
'*                                                             *
'*   TETRIS XL                                                 *
'*   AUTHOR: OLIVER CHAMBERS                                   *
'*                                                             *
'***************************************************************
'***************************************************************

'******************************
'  DECLARATIONS & DATA TYPES
'******************************

Public Type POINT
    X As Long
    Y As Long
End Type

Public Enum TetrominoeType
    i = 0
    L = 1
    J = 2
    S = 3
    Z = 4
    T = 5
    O = 6
End Enum

Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, _
                                               ByVal nIDEvent As Long, _
                                               ByVal uElapse As Long, _
                                               ByVal lpTimerFunc As Long) As Long

Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, _
                                                ByVal nIDEvent As Long) As Long

'******************************
'  GLOBALS
'******************************

Private GameTimerID     As Long
Private board           As TetrisBoard
Private tet             As Tetrominoe
Private next_tet        As Tetrominoe
Private timeInterval    As Long

'******************************
'  MAIN LOOP
'******************************

Public Sub new_game()
    
    'ensure the timer is not already set (i.e. a game is active)
    endTimer
        
    Application.OnKey "{UP}", "rotate"
    Application.OnKey "{LEFT}", "left"
    Application.OnKey "{RIGHT}", "right"
    Application.OnKey "{DOWN}", "dropOnce"
    Application.OnKey " ", "dropAllTheWay"
    
    Set board = New TetrisBoard
    board.init Sheets("Tetris").Range("tetris_board"), _
               Sheets("Tetris").Range("score"), _
               Sheets("Tetris").Range("preview_window"), _
               Sheets("Tetris").Range("level")
    
    Set tet = New Tetrominoe: tet.init board, randomTetrominoeType
    Set next_tet = New Tetrominoe: next_tet.init board, randomTetrominoeType
    
    refreshLeaderboard
    
    board.updatePreviewWindow next_tet
        
    timeInterval = 200
    startTimer timeInterval
    
    'seed the Rnd() function
    Randomize
    
End Sub

Private Sub dropOnce()
    dropTetrominoe
End Sub

Private Sub dropAllTheWay()
    
    Dim hresult As Long
    Dim bonus As Long: bonus = 0
    Do
        hresult = dropTetrominoe()
        If hresult = 0 Then
            bonus = bonus + 2
        Else
            Exit Do
        End If
    Loop
    
    If hresult = 1 Then
        board.addBonusScore bonus
    End If
    
End Sub

Private Function dropTetrominoe() As Long
    'dropTetrominoe returns:
    ' 0 = dropped
    ' 1 = no drop, new tetrominoe
    ' 2 = gameover
    
    If board.canDrop(tet) Then
        dropTetrominoe = 0
        
        board.eraseTetrominoe tet
        tet.drop
        board.drawTetrominoe tet
    Else
        If board.gameOver(tet) Then
            dropTetrominoe = 2
            gameOver
        Else
            dropTetrominoe = 1
            board.removeFullRows tet
            Set tet = next_tet
            Set next_tet = New Tetrominoe: next_tet.init board, randomTetrominoeType
            board.updatePreviewWindow next_tet
            checkLevel
        End If
    End If
    
End Function

'******************************
'  TIMER
'******************************

Private Function GetAddressOf(a As Long) As Long

    GetAddressOf = a
    
End Function

Private Sub startTimer(ByVal interval As Long)

    GameTimerID = SetTimer(0, 0, interval, GetAddressOf(AddressOf dropOnce))
    
End Sub

Private Sub endTimer()

    If Not (GameTimerID = 0) Then
        KillTimer 0, GameTimerID
        GameTimerID = 0
    End If
    
End Sub

'******************************
'  GAME CONTROL
'******************************

Public Sub pauseGame()
    endTimer
End Sub

Public Sub resumeGame()
    startTimer timeInterval
End Sub

Private Sub checkLevel()
    
    If Int(board.clearedLines / 10) > board.level Then
        board.incrementLevel
        If timeInterval >= 90 Then
            timeInterval = timeInterval - 10
            endTimer
            startTimer timeInterval
        End If
    End If
        
End Sub

Public Sub rotate()
    
    If board.canRotate(tet) Then
        board.eraseTetrominoe tet
        tet.rotate
        board.drawTetrominoe tet
    End If
    
End Sub

Public Sub left()
    
    If board.canMoveLeft(tet) Then
        board.eraseTetrominoe tet
        tet.moveLeft
        board.drawTetrominoe tet
    End If
    
End Sub

Public Sub right()
    
    If board.canMoveRight(tet) Then
        board.eraseTetrominoe tet
        tet.moveRight
        board.drawTetrominoe tet
    End If
    
End Sub


'******************************
'  OTHER FUNCTIONS
'******************************


Private Function randomTetrominoeType() As TetrominoeType

    randomTetrominoeType = Int(Rnd() * 7)
    
End Function

Public Sub clear_board()

    If (board Is Nothing) Then
        Set board = New TetrisBoard
        board.init Sheets("Tetris").Range("tetris_board"), Sheets("Tetris").Range("score"), Sheets("Tetris").Range("preview_window"), Sheets("Tetris").Range("level")
    End If

    board.clearBoard
    deactivate
    
End Sub

Public Sub deactivate()

    Set board = Nothing
    Set tet = Nothing
    Application.OnKey "{UP}"
    Application.OnKey "{LEFT}"
    Application.OnKey "{RIGHT}"
    Application.OnKey "{DOWN}"
    Application.OnKey " "
    
End Sub

Private Function gameOver()
    
    endTimer
    
    MsgBox "GAME OVER!", vbOKOnly, "TETRIS"
    
    Dim place As Long: place = leaderboardPlace(board.score)
    
    If place > 0 Then
        updateLeaderboard board.score, place
    End If
        
    deactivate
    
End Function

Private Function refreshLeaderboard()
    
    Dim wsDB As Worksheet: Set wsDB = Sheets("xTopTen")
    Dim wsT  As Worksheet: Set wsT = Sheets("Tetris")

    wsT.Range("leader_board").Value = wsDB.Range("topTen").Value
    
End Function

Private Function leaderboardPlace(ByVal score As Long) As Integer
    
    Dim topTen As Range: Set topTen = Sheets("xTopTen").Range("topTen")
    
    leaderboardPlace = -1
    Dim iter As Integer
    For iter = 1 To 10
        If score > topTen.Cells(iter, 2).Value Then
            leaderboardPlace = iter
            Exit Function
        End If
    Next iter
    
End Function

Private Function updateLeaderboard(ByVal score As Long, ByVal position As Integer) As Boolean
    
    updateLeaderboard = False

    Dim name As String
    name = Application.InputBox("Name: ", "High Score!", Default:="your name here", Type:=2)
    
    If (name = "your name here") Or (name = vbNullString) Then Exit Function
    
    Dim topTen As Range: Set topTen = Sheets("xTopTen").Range("topTen")
    
    If position <= 9 Then
        'shift other scores down
        topTen.Cells(position + 1, 2).Resize(10 - position, 2).Value = topTen.Cells(position, 2).Resize(10 - position, 2).Value
    End If
    
    topTen.Cells(position, 2) = score
    topTen.Cells(position, 3) = name
    
    refreshLeaderboard
    
    updateLeaderboard = True
    
End Function


Private Sub unhide_sheet()
    If Sheets("xtopten").Visible = xlSheetVeryHidden Then
        Sheets("xtopten").Visible = xlSheetVisible
    Else
        Sheets("xtopten").Visible = xlSheetVeryHidden
    End If
End Sub
