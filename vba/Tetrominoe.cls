VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tetrominoe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_board                As TetrisBoard
Private m_orientation(0 To 3)  As Variant
Private m_currentOrientation   As Integer
Private m_colour               As Long
Private m_position             As POINT

Public Function init(ByRef board As TetrisBoard, ByRef ttype As TetrominoeType) As Boolean
    
    Set m_board = board
    m_currentOrientation = 0
    
    'm_orientation:
    'consider a 4 x 4 range where the top-left is (0,0). A tetrominoe will
    'consist of 4 shaded squares which are encoded in the array m_orientation
    'the elements of the array are: (x0, y0, x1, y1, .., x3, y3, maxX, maxY)
    Select Case ttype
        Case TetrominoeType.i
            m_orientation(0) = Array(0, 0, 0, 1, 0, 2, 0, 3, 0, 3)
            m_orientation(1) = Array(0, 0, 1, 0, 2, 0, 3, 0, 3, 0)
            m_orientation(2) = m_orientation(0)
            m_orientation(3) = m_orientation(1)
            m_colour = RGB(225, 190, 0)
            
        Case TetrominoeType.L
            m_orientation(0) = Array(0, 0, 0, 1, 0, 2, 1, 0, 1, 2)
            m_orientation(1) = Array(0, 0, 0, 1, 1, 1, 2, 1, 2, 1)
            m_orientation(2) = Array(0, 2, 1, 0, 1, 1, 1, 2, 1, 2)
            m_orientation(3) = Array(0, 0, 1, 0, 2, 0, 2, 1, 2, 1)
            m_colour = RGB(0, 0, 153)
            
        Case TetrominoeType.J
            m_orientation(0) = Array(0, 0, 0, 1, 0, 2, 1, 2, 1, 2)
            m_orientation(1) = Array(0, 1, 1, 1, 2, 1, 2, 0, 2, 1)
            m_orientation(2) = Array(0, 0, 1, 0, 1, 1, 1, 2, 1, 2)
            m_orientation(3) = Array(0, 0, 0, 1, 1, 0, 2, 0, 2, 1)
            m_colour = RGB(112, 48, 160)
           
         Case TetrominoeType.S
            m_orientation(0) = Array(0, 1, 0, 2, 1, 0, 1, 1, 1, 2)
            m_orientation(1) = Array(0, 0, 1, 0, 1, 1, 2, 1, 2, 1)
            m_orientation(2) = m_orientation(0)
            m_orientation(3) = m_orientation(1)
            m_colour = RGB(0, 176, 240)
            
         Case TetrominoeType.Z
            m_orientation(0) = Array(0, 0, 0, 1, 1, 1, 1, 2, 1, 2)
            m_orientation(1) = Array(0, 1, 1, 0, 1, 1, 2, 0, 2, 1)
            m_orientation(2) = m_orientation(0)
            m_orientation(3) = m_orientation(1)
            m_colour = RGB(0, 176, 80)
            
         Case TetrominoeType.T
            m_orientation(0) = Array(0, 0, 0, 1, 0, 2, 1, 1, 1, 2)
            m_orientation(1) = Array(0, 1, 1, 0, 1, 1, 2, 1, 2, 1)
            m_orientation(2) = Array(0, 1, 1, 0, 1, 1, 1, 2, 1, 2)
            m_orientation(3) = Array(0, 0, 1, 0, 1, 1, 2, 0, 2, 1)
            m_colour = RGB(255, 255, 0)
            
         Case TetrominoeType.O
            m_orientation(0) = Array(0, 0, 0, 1, 1, 0, 1, 1, 1, 1)
            m_orientation(1) = m_orientation(0)
            m_orientation(2) = m_orientation(0)
            m_orientation(3) = m_orientation(0)
            m_colour = RGB(255, 0, 0)
        
    End Select
    
    m_position.X = 1
    m_position.Y = 5
            
    init = True
    
End Function


'******************************
'  PROPERTIES
'******************************

Property Get orientation() As Variant
    orientation = m_orientation(m_currentOrientation)
End Property

Property Get colour() As Double
    colour = m_colour
End Property

Property Get position() As POINT
    position = m_position
End Property

Property Get bottom() As Long
    'this means the last row of the tetris board occupied by the piece
    bottom = m_position.X + m_orientation(m_currentOrientation)(8)
End Property

Property Get right() As Long
    'this means the right-most column of the tetris board occupied by the piece
    right = m_position.Y + m_orientation(m_currentOrientation)(9)
End Property

'******************************
'  METHODS
'******************************

Public Function rotate() As Boolean
    m_currentOrientation = (m_currentOrientation + 1) Mod 4
    rotate = True
End Function

Public Function nextOrientation() As Variant
    Dim orient As Integer: orient = (m_currentOrientation + 1) Mod 4
    nextOrientation = m_orientation(orient)
End Function

Public Function moveLeft() As Boolean
    m_position.Y = m_position.Y - 1
    moveLeft = True
End Function

Public Function moveRight() As Boolean
    m_position.Y = m_position.Y + 1
    moveRight = True
End Function

Public Function drop() As Boolean
    m_position.X = m_position.X + 1
End Function



