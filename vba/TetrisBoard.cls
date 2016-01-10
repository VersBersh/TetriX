VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TetrisBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_boardColour   As Long
Private m_board         As Range
Private m_memBoard()    As Boolean
Private m_scoreDisplay  As Range
Private m_width         As Long
Private m_height        As Long
Private m_score         As Long
Private m_previewWindow As Range
Private m_levelDisplay  As Range
Private m_level         As Long
Private m_clearedLines  As Long

'******************************
'  BOARD SETUP
'******************************
Public Sub init(ByRef board As Range, ByRef scoreDisplay As Range, ByRef previewWindow As Range, ByRef level As Range)
    Set m_board = board
    Set m_scoreDisplay = scoreDisplay
    Set m_previewWindow = previewWindow
    Set m_levelDisplay = level
    
    m_clearedLines = 0
    m_level = 0
    m_levelDisplay.Value = m_level
    
    m_width = m_board.Columns.Count
    m_height = m_board.Rows.Count
    
    ReDim m_memBoard(1 To m_height, 1 To m_width)
    
    'initialises memory board to FALSE in a loop because VBA sucks
    Dim row As Integer, col As Integer
    For row = 1 To m_height
        For col = 1 To m_width
            m_memBoard(row, col) = False
        Next col
    Next row
    
    m_boardColour = RGB(197, 232, 245)
    m_score = 0
    
    clearBoard
    updateScore
End Sub

'******************************
'  PROPERTIES
'******************************

Property Get width() As Long
    width = m_width
End Property

Property Get height() As Long
    height = m_height
End Property

Property Get score() As Long
    score = m_score
End Property

Property Get clearedLines() As Long
    clearedLines = m_clearedLines
End Property
'******************************
'  NON-MAIN-BOARD UPDATES
'******************************

Public Function addBonusScore(ByVal bonus As Long) As Boolean
    m_score = m_score + bonus
    updateScore
    addBonusScore = True
End Function

Private Function updateScore() As Boolean
    m_scoreDisplay.Value = m_score
End Function

Public Function incrementLevel() As Boolean
    m_level = m_level + 1
    m_levelDisplay.Value = level
End Function

Public Function level() As Long
    level = m_level
End Function

Public Function updatePreviewWindow(ByRef piece As Tetrominoe) As Boolean
    
    Dim colour As Double: colour = piece.colour()
    
    Dim iter  As Integer
    Dim T     As Variant: T = piece.orientation()
    
    m_previewWindow.Interior.Color = m_boardColour
    For iter = 0 To 3
        m_previewWindow.Cells(2 + T(2 * iter), 2 + T(2 * iter + 1)).Interior.Color = colour
    Next iter
    
    updatePreviewWindow = True
End Function

'******************************
'  SCORING AND GAMEOVER
'******************************

Public Function gameOver(ByRef piece As Tetrominoe) As Boolean
    gameOver = (piece.position.X = 1)
End Function

Public Function removeFullRows(ByRef piece As Tetrominoe) As Boolean
       
    Dim row As Long, clines As Long
    For row = piece.position.X To piece.bottom()
        If rowFull(row) Then
            removeRow row
            clines = clines + 1
        End If
    Next row
    
    Dim add_score As Long: add_score = 0
    Select Case clines
        Case 0:
            add_score = 0
        Case 1:
            add_score = 40
        Case 2:
            add_score = 100
        Case 3:
            add_score = 300
        Case 4:
            add_score = 1200
    End Select
    
    m_clearedLines = m_clearedLines + clines
    m_score = m_score + add_score * (m_level + 1)
    updateScore
    removeFullRows = True
    
End Function

Private Function rowFull(ByVal row As Integer) As Boolean
    
    rowFull = True
    
    Dim col As Integer
    For col = 1 To m_width
        If m_memBoard(row, col) = False Then
            rowFull = False
            Exit Function
        End If
    Next col
        
End Function

Private Function removeRow(ByVal row As Integer) As Boolean
    
    Dim r As Integer, c As Integer
    
    For r = row To 2 Step -1
        For c = 1 To m_width
            m_memBoard(r, c) = m_memBoard(r - 1, c)
        Next c
    Next r
    
    m_board.Cells(2, 1).Resize(row - 1, m_width).Value(11) = m_board.Cells(1, 1).Resize(row - 1, m_width).Value(11)
    
    removeRow = True
    
End Function


'******************************
'  DRAWING FUNCTIONS
'******************************

Public Function clearBoard() As Boolean
    m_board.Interior.Color = m_boardColour
    m_previewWindow.Interior.Color = m_boardColour
    clearBoard = True
End Function

Public Function drawTetrominoe(ByRef piece As Tetrominoe) As Boolean
    
    drawTetrominoe = draw(piece, True)
    
End Function

Public Function eraseTetrominoe(ByRef piece As Tetrominoe) As Boolean
   
    eraseTetrominoe = draw(piece, False)
    
End Function

Private Function draw(ByRef piece As Tetrominoe, ByVal paint As Boolean) As Boolean
    'perform drawing operation on memory board and screen
    'if paint = True, the tetrominoe is painted, otherwise it is erased.
    
    Dim colour As Double
    If paint Then
        colour = piece.colour()
    Else
        colour = m_boardColour
    End If
    
    draw = False
    If Not withinBoard(piece) Then Exit Function
    
    Dim iter  As Integer
    Dim T     As Variant: T = piece.orientation()
    
    For iter = 0 To 3
        m_board.Cells(piece.position.X + T(2 * iter), piece.position.Y + T(2 * iter + 1)).Interior.Color = colour
        m_memBoard(piece.position.X + T(2 * iter), piece.position.Y + T(2 * iter + 1)) = paint
    Next iter
    
    draw = True
    
End Function

Private Function memDraw(ByRef piece As Tetrominoe, ByVal paint As Boolean) As Boolean
    'perform drawing operation only on the memory board
    
    Dim colour As Double
    
    memDraw = False
    If Not withinBoard(piece) Then Exit Function
    
    Dim iter  As Integer
    Dim T     As Variant: T = piece.orientation()
    
    For iter = 0 To 3
        m_memBoard(piece.position.X + T(2 * iter), piece.position.Y + T(2 * iter + 1)) = paint
    Next iter
    
    memDraw = True
    
End Function

'******************************
'  BOUNDARY FUNCTIONS
'******************************

Private Function withinBoard(ByRef piece As Tetrominoe)

    withinBoard = (0 <= piece.position.X) And _
                  (piece.bottom <= m_height) And _
                  (0 <= piece.position.Y) And _
                  (piece.right <= m_width)
                  
End Function

Public Function canDrop(ByRef piece As Tetrominoe) As Boolean
    
    'erase tetrominoe from memory board
    memDraw piece, False
    
    If piece.bottom() >= m_height Then
        canDrop = False
    Else
        canDrop = True
        
        Dim T    As Variant: T = piece.orientation()
        Dim iter As Integer
        
        For iter = 0 To 3
            If (m_memBoard(piece.position.X + T(2 * iter) + 1, piece.position.Y + T(2 * iter + 1)) = True) Then
                canDrop = False
                memDraw piece, True
                Exit Function
            End If
        Next iter
    End If
    
    'redraw tetrominoe on memory board
    memDraw piece, True
    
End Function


Public Function canMoveLeft(ByRef piece As Tetrominoe) As Boolean
    
    memDraw piece, False
    
    If piece.position.Y <= 1 Then
        canMoveLeft = False
    Else
        canMoveLeft = True
        
        Dim T    As Variant: T = piece.orientation()
        Dim iter As Integer
        
        For iter = 0 To 3
            If (m_memBoard(piece.position.X + T(2 * iter), piece.position.Y + T(2 * iter + 1) - 1) = True) _
                Or (piece.position.Y + T(2 * iter + 1) - 1 <= 0) Then
                canMoveLeft = False
                memDraw piece, True
                Exit Function
            End If
        Next iter
    End If
    
    memDraw piece, True
    
End Function

Public Function canMoveRight(ByRef piece As Tetrominoe) As Boolean
    
    
    memDraw piece, False
    
    If piece.right >= m_width Then
        canMoveRight = False
    Else
        canMoveRight = True
        
        Dim T    As Variant: T = piece.orientation()
        Dim iter As Integer
        
        For iter = 0 To 3
            If (m_memBoard(piece.position.X + T(2 * iter), piece.position.Y + T(2 * iter + 1) + 1) = True) _
               Or (piece.position.Y + T(2 * iter + 1) + 1 > m_width) Then
                canMoveRight = False
                memDraw piece, True
                Exit Function
            End If
        Next iter
    End If
    
    memDraw piece, True
    
End Function

Public Function canRotate(ByRef piece As Tetrominoe) As Boolean
    
    Dim T    As Variant: T = piece.nextOrientation()
    If Not ((0 <= piece.position.X) And _
            (piece.position.X + T(8) <= m_height) And _
            (0 <= piece.position.Y) And _
            (piece.position.Y + T(9) <= m_width)) Then
    
        canRotate = False
        Exit Function
    End If
    
    memDraw piece, False
    
    canRotate = True
    Dim iter As Integer
    
    For iter = 0 To 3
        If (m_memBoard(piece.position.X + T(2 * iter), piece.position.Y + T(2 * iter + 1)) = True) _
            Or (piece.position.Y + T(2 * iter + 1) <= 0) _
            Or (piece.position.Y + T(2 * iter + 1) > m_width) _
            Or (piece.position.X + T(2 * iter) > m_height) Then
            canRotate = False
            memDraw piece, True
            Exit Function
        End If
    Next iter
    
    memDraw piece, True
    
End Function


